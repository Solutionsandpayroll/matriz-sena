import { useState, useMemo, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import {
  FileSpreadsheet,
  Upload,
  CheckCircle2,
  AlertTriangle,
  HelpCircle,
  ChevronDown,
  ChevronUp,
  FileText,
  Download,
  Sparkles,
  Briefcase,
  Calendar,
  Clock,
  Check,
  Building2,
  Users,
  Search,
  ListChecks,
  Wrench,
  RefreshCw,
  Trash2,
  Save,
  Loader2,
  Circle,
  PlusCircle,
} from "lucide-react";
import listadoCnoLocal from "./data/listado-cno.json";
import {
  NOMBRES_MES,
  MESES_PERIODO_SENA,
  claveCargo,
  formatFecha,
  jornadaLegalSugeridaDetallada,
  calcularProporcionPorFechas,
  calcularProporcionPorHoras,
  calcularCuotaAprendices,
  describirNovedad,
  parsearNomina,
  parsearSeguridadSocial,
  fusionarFechaRetiro,
  buscarCargoEnOtrasNominas,
  indexarListadoCno,
  sugerirCno,
  nombreDeCodigo,
  construirHomologacion,
  homologacionListaParaConfirmar,
  agruparPorCargo,
  TEXTO_REGLA_CUOTA_APRENDICES,
  TOLERANCIA_DIFERENCIA_PROPORCION,
} from "./utils/sena";
import { exportarMatrizExcel } from "./utils/exportarExcel";
import { listarEmpresasApi, cargarEmpresaApi, guardarEmpresaApi, eliminarEmpresaApi } from "./utils/empresasApi";

// El listado se indexa una sola vez (5.000+ filas).
const INDICE_CNO = indexarListadoCno(listadoCnoLocal || []);

const API_URL = import.meta.env?.VITE_API_URL || "http://localhost:3001";
const JORNADAS = [40, 42, 44, 46, 47, 48];

// Cuánto se espera antes de guardar en el servidor después de que la persona
// deja de escribir (para no mandar una petición por cada tecla).
const RETRASO_GUARDADO_MS = 800;

const claveMes = (anio, mesIndex) => `${anio}-${mesIndex}`;

// ---------------------------------------------------------------------
// Traducción de cargos (EN -> ES).
//
// 1) Primero se intenta con el proxy local (server.js), que llama a la API
//    de Anthropic con la key del .env. Es la mejor calidad para cargos
//    técnicos, pero requiere saldo en la cuenta de la API.
// 2) Si eso falla por cualquier razón (sin servidor, sin saldo, etc.), se
//    intenta automáticamente con MyMemory (api.mymemory.translated.net),
//    un servicio de traducción GRATUITO que no necesita API key ni tarjeta.
//    La calidad es más básica para términos técnicos: revisa y corrige a
//    mano en la tabla de Homologación lo que haga falta.
// ---------------------------------------------------------------------
async function traducirConMyMemory(texto) {
  const url = `https://api.mymemory.translated.net/get?q=${encodeURIComponent(texto)}&langpair=en|es`;
  const resp = await fetch(url);
  if (!resp.ok) throw new Error(`MyMemory respondió ${resp.status}`);
  const data = await resp.json();
  const traducido = data?.responseData?.translatedText;
  if (!traducido || String(traducido).toUpperCase().includes("MYMEMORY WARNING")) {
    throw new Error("MyMemory no devolvió una traducción válida");
  }
  return traducido;
}

async function traducirCargosConIA(cargosUnicos) {
  try {
    const response = await fetch(`${API_URL}/api/traducir-cargos`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ cargosUnicos }),
    });
    if (!response.ok) throw new Error(`El servidor respondió ${response.status}`);
    const data = await response.json();
    const normalizadas = {};
    Object.entries(data.traducciones || {}).forEach(([k, v]) => {
      if (v && String(v).trim()) normalizadas[claveCargo(k)] = String(v).trim();
    });
    return normalizadas;
  } catch (errorIA) {
    // Respaldo gratuito: se traduce cargo por cargo con MyMemory. Con una
    // pequeña pausa entre llamadas para no golpear su límite gratuito
    // (~5.000 palabras/día por IP), que para nombres de cargos alcanza bien.
    console.warn(
      `Traducción con IA no disponible (${errorIA.message}). Usando MyMemory (gratis) como respaldo.`
    );
    const normalizadas = {};
    for (const cargo of cargosUnicos) {
      try {
        const traducido = await traducirConMyMemory(cargo);
        if (traducido && traducido.trim()) normalizadas[claveCargo(cargo)] = traducido.trim();
      } catch {
        // Si también falla para este cargo puntual, se deja sin traducir:
        // el aviso existente lo señala y se corrige a mano en Homologación.
      }
      await new Promise((r) => setTimeout(r, 200));
    }
    return normalizadas;
  }
}

async function leerFilas(archivo) {
  const buffer = await archivo.arrayBuffer();
  const libro = XLSX.read(buffer, { type: "array" });
  const hoja = libro.Sheets[libro.SheetNames[0]];
  // Sin cellDates: las fechas llegan como serial de Excel y las interpreta parseFecha().
  return XLSX.utils.sheet_to_json(hoja, { header: 1, defval: null, raw: true });
}

// Sugiere la jornada de un mes respetando primero un ajuste manual guardado
// para ESE mes puntual; si no hay ninguno, usa la jornada fija de la empresa
// o la legal vigente (con aviso si el mes cambia de jornada a mitad de mes).
function resolverJornadaMes(anio, mesIndex, jornadaEmpresa, jornadasManuales) {
  const clave = claveMes(anio, mesIndex);
  if (jornadasManuales?.[clave] !== undefined) {
    return { jornada: Number(jornadasManuales[clave]), manual: true, detalle: null };
  }
  if (jornadaEmpresa === "auto") {
    const detalle = jornadaLegalSugeridaDetallada(anio, mesIndex);
    return { jornada: detalle.jornada, manual: false, detalle };
  }
  return { jornada: Number(jornadaEmpresa), manual: false, detalle: null };
}

// Genera los 6 meses previos al mes de presentación (ej. presentar en septiembre 2026 -> marzo..agosto 2026)
function generarMeses(mesPresentacion, anioPresentacion, jornadaEmpresa, jornadasManuales, previos) {
  return Array.from({ length: 6 }, (_, i) => {
    const d = new Date(anioPresentacion, mesPresentacion - (6 - i), 1);
    const anio = d.getFullYear();
    const mesIndex = d.getMonth();
    const { jornada, detalle } = resolverJornadaMes(anio, mesIndex, jornadaEmpresa, jornadasManuales);
    return {
      id: i,
      anio,
      mesIndex,
      etiqueta: `${NOMBRES_MES[mesIndex]} ${anio}`,
      jornadaSemanal: jornada,
      jornadaDividida: detalle?.dividido ? detalle : null,
      nomina: previos?.[i]?.nomina ?? null,
      segSocial: previos?.[i]?.segSocial ?? null,
    };
  });
}

// ---------------------------------------------------------------------
// Encabezado de sección reutilizable: número de paso + título + descripción.
// Da la misma jerarquía visual a las 4 secciones principales del flujo.
// ---------------------------------------------------------------------
function EncabezadoSeccion({ numero, icono: Icono, titulo, descripcion, extra }) {
  return (
    <div className="flex flex-wrap items-start justify-between gap-3 border-b border-slate-100 pb-4">
      <div className="flex items-start gap-3">
        <span className="flex items-center justify-center w-7 h-7 rounded-full bg-[#003B7A] text-white text-xs font-extrabold shrink-0 mt-0.5">
          {numero}
        </span>
        <div className="space-y-0.5">
          <h2 className="text-sm font-bold text-slate-900 flex items-center gap-2">
            {Icono && <Icono className="w-4 h-4 text-[#003B7A]" />}
            {titulo}
          </h2>
          {descripcion && <p className="text-[11px] text-slate-500 max-w-xl leading-relaxed">{descripcion}</p>}
        </div>
      </div>
      {extra}
    </div>
  );
}

// ---------------------------------------------------------------------
// Panel lateral fijo: muestra los 4 pasos del flujo y su estado, y permite
// saltar directo a cada sección. Es la brújula de la página completa.
// ---------------------------------------------------------------------
function BarraDePasos({ pasos, pasoActivo }) {
  return (
    <nav className="hidden lg:block lg:sticky lg:top-20 h-fit">
      <div className="bg-white border border-slate-200/90 rounded-2xl shadow-xs p-4 space-y-0.5 w-56">
        <p className="text-[10px] font-extrabold text-slate-400 uppercase tracking-wider px-2 pb-2">Flujo del reporte</p>
        {pasos.map((p, i) => {
          const activo = pasoActivo === p.id;
          return (
            <a
              key={p.id}
              href={`#${p.id}`}
              className={`flex items-start gap-2.5 rounded-xl px-2.5 py-2.5 text-xs font-semibold transition-colors ${
                activo ? "bg-blue-50 text-[#003B7A]" : "text-slate-500 hover:bg-slate-50 hover:text-slate-700"
              }`}
            >
              {p.completo ? (
                <CheckCircle2 className="w-4 h-4 text-emerald-500 shrink-0 mt-0.5" />
              ) : (
                <Circle className={`w-4 h-4 shrink-0 mt-0.5 ${activo ? "text-[#003B7A]" : "text-slate-300"}`} />
              )}
              <span className="leading-snug">
                {p.titulo}
                {p.detalle && <span className="block text-[10px] font-medium text-slate-400 mt-0.5">{p.detalle}</span>}
              </span>
            </a>
          );
        })}
      </div>
    </nav>
  );
}

// ---------------------------------------------------------------------
// Fila de la tabla de homologación (un cargo, se aprueba una sola vez)
// ---------------------------------------------------------------------
function FilaHomologacion({ clave, h, cantidad, onCambiar }) {
  const [abierto, setAbierto] = useState(false);
  const [consulta, setConsulta] = useState(h.es);
  const resultados = useMemo(() => (abierto ? sugerirCno(consulta, INDICE_CNO, 25) : []), [abierto, consulta]);

  const elegir = (codigo, ocupacion) => {
    onCambiar(clave, { codigo, ocupacion, confirmado: true, noCalificado: false });
    setAbierto(false);
  };

  const marcarNoCalificado = () => {
    onCambiar(clave, { codigo: "", ocupacion: "", confirmado: true, noCalificado: true });
    setAbierto(false);
  };

  return (
    <>
      <tr className="hover:bg-slate-50/80 transition-colors align-top">
        <td className="py-2.5 px-3 text-slate-400">{h.original}</td>
        <td className="py-2.5 px-3">
          <input
            type="text"
            value={h.es}
            onChange={(e) => onCambiar(clave, { es: e.target.value })}
            className="w-full bg-slate-50 border border-slate-200 rounded-lg px-2 py-1.5 text-xs font-medium text-slate-800 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
          />
        </td>
        <td className="py-2.5 px-3">
          {h.codigo ? (
            <div className="space-y-0.5">
              <span className="font-mono font-bold text-[#003B7A]">{h.codigo}</span>
              <p className="text-[11px] text-slate-500 leading-snug">{h.ocupacion || nombreDeCodigo(h.codigo, INDICE_CNO)}</p>
            </div>
          ) : h.noCalificado ? (
            <span className="text-slate-500 font-semibold">Sin código (no calificado, marcado a mano)</span>
          ) : h.sugerenciaBaja ? (
            <div className="space-y-0.5">
              <span className="text-amber-700 font-semibold">Sin código (sugerencia de baja cobertura, sin confirmar)</span>
              <p className="text-[11px] text-slate-500 leading-snug">
                ¿Será <span className="font-mono font-bold">{h.sugerenciaBaja.codigo}</span> · {h.sugerenciaBaja.ocupacion}? Revísalo con "Cambiar".
              </p>
            </div>
          ) : (
            <span className="text-rose-700 font-semibold">Sin código (no se encontró ninguna coincidencia)</span>
          )}
        </td>
        <td className="py-2.5 px-3 text-center font-semibold text-slate-600">{cantidad}</td>
        <td className="py-2.5 px-3">
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={() => setAbierto(!abierto)}
              className="text-[11px] font-bold text-[#003B7A] hover:underline cursor-pointer"
            >
              {abierto ? "Cerrar" : "Cambiar"}
            </button>
            <label className="flex items-center gap-1.5 text-[11px] font-semibold text-slate-600 cursor-pointer">
              <input
                type="checkbox"
                checked={h.confirmado}
                onChange={(e) => onCambiar(clave, { confirmado: e.target.checked })}
                className="accent-[#003B7A]"
              />
              Confirmado
            </label>
          </div>
        </td>
      </tr>
      {abierto && (
        <tr className="bg-blue-50/40">
          <td colSpan={5} className="px-4 py-3">
            <div className="space-y-2">
              <div className="relative">
                <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" />
                <input
                  type="text"
                  autoFocus
                  value={consulta}
                  onChange={(e) => setConsulta(e.target.value)}
                  placeholder="Busca por cargo en español o por código (ej. 1232)"
                  className="w-full bg-white border border-slate-200 rounded-lg pl-8 pr-3 py-2 text-xs text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                />
              </div>
              <div className="max-h-56 overflow-y-auto rounded-lg border border-slate-200 bg-white divide-y divide-slate-100">
                {resultados.length === 0 && (
                  <p className="px-3 py-2 text-[11px] text-slate-400">Sin coincidencias. Prueba con otra palabra o con el código.</p>
                )}
                {resultados.map((r, i) => (
                  <button
                    type="button"
                    key={`${r.codigo}-${i}`}
                    onClick={() => elegir(r.codigo, r.ocupacion)}
                    className="w-full text-left px-3 py-1.5 text-xs hover:bg-blue-50 cursor-pointer flex gap-2"
                  >
                    <span className="font-mono font-bold text-[#003B7A] shrink-0">{r.codigo}</span>
                    <span className="text-slate-700">{r.ocupacion}</span>
                  </button>
                ))}
              </div>
              <button
                type="button"
                onClick={marcarNoCalificado}
                className="text-[11px] font-bold text-amber-700 hover:underline cursor-pointer"
              >
                Dejar sin código (oficio no calificado)
              </button>
            </div>
          </td>
        </tr>
      )}
    </>
  );
}

// ---------------------------------------------------------------------
// Resumen por cargo por mes (lo que realmente se va a exportar) — punto C.1
// ---------------------------------------------------------------------
function ResumenPorCargo({ grupos }) {
  const calificados = grupos.filter((g) => g.calificado).sort((a, b) => a.codigo.localeCompare(b.codigo));
  const noCalificados = grupos.filter((g) => !g.calificado).sort((a, b) => a.nombreCargo.localeCompare(b.nombreCargo, "es"));

  const Tabla = ({ titulo, primeraCol, lista, campo }) => (
    <div className="space-y-1.5">
      <p className="text-[11px] font-extrabold text-slate-500 uppercase tracking-wider">{titulo}</p>
      {lista.length === 0 ? (
        <p className="text-xs text-slate-400">(ninguno)</p>
      ) : (
        <div className="rounded-lg border border-slate-200 overflow-hidden">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-100/70 text-[10px] uppercase text-slate-500">
              <tr>
                <th className="py-2 px-3">{primeraCol}</th>
                <th className="py-2 px-3 text-center">Personas</th>
                <th className="py-2 px-3 text-center">Horas totales/semana</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {lista.map((g) => (
                <tr key={campo(g)}>
                  <td className="py-2 px-3 font-semibold text-slate-800">{campo(g)}</td>
                  <td className="py-2 px-3 text-center">{g.personas}</td>
                  <td className="py-2 px-3 text-center">{g.totalHoras}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );

  return (
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
      <Tabla titulo="Matriz 1 · Oficios calificados" primeraCol="Código CNO" lista={calificados} campo={(g) => g.codigo} />
      <Tabla titulo="Matriz 2 · Oficios no calificados" primeraCol="Cargo" lista={noCalificados} campo={(g) => g.nombreCargo} />
    </div>
  );
}

// ---------------------------------------------------------------------
// Panel lateral: lista de empresas guardadas en el servidor (SQLite),
// compartida entre todo el equipo. Cargar trae toda la configuración de
// esa empresa; Guardar empuja la configuración actual en pantalla;
// Nueva empresa limpia el formulario para empezar una empresa distinta
// sin arrastrar los datos de la que estaba cargada.
// ---------------------------------------------------------------------
function PanelEmpresas({
  empresas,
  cargandoLista,
  cargandoSeleccion,
  guardando,
  error,
  nombreActual,
  onRecargar,
  onCargar,
  onEliminar,
  onGuardar,
  onNuevaEmpresa,
}) {
  return (
    <aside className="bg-white border border-slate-200/90 rounded-2xl shadow-xs overflow-hidden h-fit">
      <div className="p-4 border-b border-slate-100 bg-slate-50/50 flex items-center justify-between gap-2">
        <h2 className="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center gap-1.5">
          <Building2 className="w-3.5 h-3.5 text-[#003B7A]" /> Empresas guardadas
        </h2>
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={onNuevaEmpresa}
            className="text-slate-400 hover:text-[#003B7A] cursor-pointer"
            title="Empezar una empresa nueva (en blanco)"
          >
            <PlusCircle className="w-4 h-4" />
          </button>
          <button
            type="button"
            onClick={onRecargar}
            className="text-slate-400 hover:text-[#003B7A] cursor-pointer"
            title="Recargar lista"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${cargandoLista ? "animate-spin" : ""}`} />
          </button>
        </div>
      </div>

      <div className="max-h-80 overflow-y-auto">
        {error && <p className="p-3 text-[11px] text-rose-700 font-medium leading-relaxed">{error}</p>}
        {!error && !cargandoLista && empresas.length === 0 && (
          <p className="p-3 text-[11px] text-slate-400 leading-relaxed">
            Todavía no hay empresas guardadas en el servidor. Escribe el nombre arriba, llena los datos y dale a "Guardar empresa actual".
          </p>
        )}
        <div className="p-2 space-y-0.5">
          {empresas.map((e) => {
            const activa = e.nombre.trim().toLowerCase() === nombreActual.trim().toLowerCase();
            return (
              <div
                key={e.nombre}
                className={`flex items-center gap-1.5 rounded-lg px-2 py-1.5 ${activa ? "bg-blue-50/70" : "hover:bg-slate-50"}`}
              >
                <button
                  type="button"
                  onClick={() => onCargar(e.nombre)}
                  disabled={cargandoSeleccion}
                  className="flex-1 text-left text-xs font-semibold text-slate-700 truncate cursor-pointer disabled:opacity-50"
                  title={`Cargar ${e.nombre}`}
                >
                  {e.nombre}
                  {activa && <CheckCircle2 className="inline w-3 h-3 ml-1 text-[#003B7A] align-text-top" />}
                </button>
                <button
                  type="button"
                  onClick={() => onEliminar(e.nombre)}
                  className="text-slate-300 hover:text-rose-600 cursor-pointer shrink-0 p-1"
                  title="Eliminar de la lista"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                </button>
              </div>
            );
          })}
        </div>
      </div>

      <div className="p-3 border-t border-slate-100 bg-slate-50/40 space-y-1.5">
        <button
          type="button"
          onClick={onGuardar}
          disabled={guardando}
          className="w-full bg-[#003B7A] hover:bg-[#002B5B] disabled:bg-slate-200 disabled:text-slate-400 text-white text-xs font-bold py-2.5 rounded-lg flex items-center justify-center gap-1.5 cursor-pointer transition"
        >
          {guardando ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Save className="w-3.5 h-3.5" />}
          {guardando ? "Guardando..." : "Guardar empresa actual"}
        </button>
        <p className="text-[10px] text-slate-400 leading-snug">
          Los campos también se guardan solos unos segundos después de escribir, siempre que hayas puesto un nombre de empresa.
        </p>
      </div>
    </aside>
  );
}

export default function App() {
  const hoy = new Date();

  const [nombreEmpresa, setNombreEmpresa] = useState("");
  const [datosEmpresa, setDatosEmpresa] = useState({
    razonSocial: "", nit: "", representanteLegal: "", cc: "", direccion: "", telefonos: "", email: "",
  });
  const [jornadaEmpresa, setJornadaEmpresa] = useState("auto"); // "auto" = jornada legal según el mes
  const [jornadasManuales, setJornadasManuales] = useState({}); // "anio-mesIndex" -> horas, ajustes mes a mes
  const [baseDias, setBaseDias] = useState("real"); // "real" | "30"
  const [periodo, setPeriodo] = useState({ mes: hoy.getMonth(), anio: hoy.getFullYear() });
  const [anioTexto, setAnioTexto] = useState(String(hoy.getFullYear())); // permite escribir el año dígito por dígito
  const [meses, setMeses] = useState(() => generarMeses(hoy.getMonth(), hoy.getFullYear(), "auto", {}));
  const [aprendicesActivos, setAprendicesActivos] = useState(0);
  const [resultadosPorMes, setResultadosPorMes] = useState({});
  const [homologacion, setHomologacion] = useState({});
  const [asignacionesManuales, setAsignacionesManuales] = useState({}); // documento -> cargo escrito a mano
  const [cargando, setCargando] = useState(false);
  const [exportando, setExportando] = useState(false);
  const [error, setError] = useState(null);
  const [avisos, setAvisos] = useState([]);
  const [mostrarInstrucciones, setMostrarInstrucciones] = useState(false);
  const [mesActivo, setMesActivo] = useState(null);
  const [necesitaReprocesar, setNecesitaReprocesar] = useState(false);

  // ---------------- Empresas guardadas (servidor / SQLite) ----------------
  const [empresasGuardadas, setEmpresasGuardadas] = useState([]);
  const [cargandoListaEmpresas, setCargandoListaEmpresas] = useState(false);
  const [cargandoEmpresaSeleccionada, setCargandoEmpresaSeleccionada] = useState(false);
  const [guardandoEmpresa, setGuardandoEmpresa] = useState(false);
  const [errorEmpresas, setErrorEmpresas] = useState(null);
  const timeoutGuardadoRef = useRef(null);

  const cargarListaEmpresas = async () => {
    setCargandoListaEmpresas(true);
    setErrorEmpresas(null);
    try {
      const lista = await listarEmpresasApi();
      setEmpresasGuardadas(lista);
    } catch (e) {
      setErrorEmpresas(`No se pudo conectar con el servidor (${API_URL}). ¿Está corriendo server.js? ${e.message}`);
    } finally {
      setCargandoListaEmpresas(false);
    }
  };

  // Trae la lista de empresas una vez al abrir la app. No selecciona ninguna
  // automáticamente: el nombre de empresa se escribe o se elige del panel.
  useEffect(() => {
    cargarListaEmpresas();
    return () => {
      if (timeoutGuardadoRef.current) clearTimeout(timeoutGuardadoRef.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Guarda en el servidor una parte de la configuración de la empresa ACTUAL
  // (nombreEmpresa). Por defecto espera un momento por si la persona sigue
  // escribiendo (debounce); pasar {inmediato:true} para acciones puntuales
  // como cambiar un select o presionar un botón.
  const guardarConfigSiHayEmpresa = (cambios, { inmediato = false } = {}) => {
    const nombre = nombreEmpresa.trim();
    if (!nombre) return;
    if (timeoutGuardadoRef.current) clearTimeout(timeoutGuardadoRef.current);

    const ejecutar = async () => {
      setGuardandoEmpresa(true);
      setErrorEmpresas(null);
      try {
        await guardarEmpresaApi(nombre, cambios);
        setEmpresasGuardadas((prev) => {
          const yaEsta = prev.some((e) => e.nombre.trim().toLowerCase() === nombre.toLowerCase());
          return yaEsta ? prev : [...prev, { nombre, actualizadoEn: new Date().toISOString() }].sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));
        });
      } catch (e) {
        setErrorEmpresas("No se pudo guardar en el servidor: " + e.message);
      } finally {
        setGuardandoEmpresa(false);
      }
    };

    if (inmediato) ejecutar();
    else timeoutGuardadoRef.current = setTimeout(ejecutar, RETRASO_GUARDADO_MS);
  };

  // Trae TODA la configuración guardada de una empresa (datos fiscales,
  // jornada, base de días, homologación, asignaciones manuales) y la pone
  // en pantalla. Si la empresa nunca se ha guardado, solo deja el nombre.
  const cargarEmpresaDesdeServidor = async (nombre) => {
    const n = (nombre ?? nombreEmpresa).trim();
    if (!n) return;
    setCargandoEmpresaSeleccionada(true);
    setErrorEmpresas(null);
    try {
      const cfg = await cargarEmpresaApi(n);
      setNombreEmpresa(n);
      if (!cfg) return; // empresa nueva, aún sin nada guardado
      if (cfg.datosEmpresa) setDatosEmpresa((prev) => ({ ...prev, ...cfg.datosEmpresa }));
      if (cfg.baseDias) setBaseDias(cfg.baseDias);
      if (cfg.jornadaEmpresa !== undefined) setJornadaEmpresa(cfg.jornadaEmpresa);
      if (cfg.jornadasPorMes) setJornadasManuales(cfg.jornadasPorMes);
      if (cfg.asignacionesManuales) setAsignacionesManuales(cfg.asignacionesManuales);
      if (cfg.homologacion) setHomologacion(cfg.homologacion);
      setMeses((prev) =>
        generarMeses(periodo.mes, periodo.anio, cfg.jornadaEmpresa ?? jornadaEmpresa, cfg.jornadasPorMes ?? jornadasManuales, prev)
      );
    } catch (e) {
      setErrorEmpresas("No se pudo cargar la empresa: " + e.message);
    } finally {
      setCargandoEmpresaSeleccionada(false);
    }
  };

  // Botón explícito: guarda TODO lo que hay en pantalla de una vez, sin
  // esperar el debounce.
  const guardarEmpresaCompleta = () => {
    if (!nombreEmpresa.trim()) {
      setErrorEmpresas("Escribe un nombre de empresa arriba antes de guardar.");
      return;
    }
    guardarConfigSiHayEmpresa(
      { datosEmpresa, baseDias, jornadaEmpresa, jornadasPorMes: jornadasManuales, asignacionesManuales, homologacion },
      { inmediato: true }
    );
  };

  const eliminarEmpresaGuardada = async (nombre) => {
    if (!window.confirm(`¿Eliminar "${nombre}" y toda su configuración guardada del servidor? Esto no se puede deshacer.`)) return;
    try {
      await eliminarEmpresaApi(nombre);
      setEmpresasGuardadas((prev) => prev.filter((e) => e.nombre.trim().toLowerCase() !== nombre.trim().toLowerCase()));
    } catch (e) {
      setErrorEmpresas("No se pudo eliminar: " + e.message);
    }
  };

  // Limpia todo el formulario para empezar una empresa nueva desde cero,
  // sin arrastrar los datos de la empresa que estaba cargada antes. No
  // guarda nada en el servidor por sí sola: eso pasa cuando se escribe
  // el nombre nuevo y se llenan datos (autoguardado) o se da a "Guardar".
  const nuevaEmpresaEnBlanco = () => {
    if (
      (nombreEmpresa.trim() || Object.keys(resultadosPorMes).length > 0) &&
      !window.confirm("¿Empezar una empresa nueva en blanco? Se perderá lo que no hayas guardado de la empresa actual.")
    ) {
      return;
    }
    setNombreEmpresa("");
    setDatosEmpresa({ razonSocial: "", nit: "", representanteLegal: "", cc: "", direccion: "", telefonos: "", email: "" });
    setJornadaEmpresa("auto");
    setJornadasManuales({});
    setBaseDias("real");
    setAsignacionesManuales({});
    setHomologacion({});
    setAprendicesActivos(0);
    setMeses(generarMeses(periodo.mes, periodo.anio, "auto", {}));
    limpiarResultados();
    setErrorEmpresas(null);
  };

  // ---------------- Periodo / configuración ----------------
  const limpiarResultados = () => {
    setResultadosPorMes({});
    setMesActivo(null);
    setAvisos([]);
    setError(null);
    setNecesitaReprocesar(false);
  };

  const cambiarPeriodo = (mes, anio) => {
    if (!Number.isInteger(anio) || anio < 2000 || anio > 2100) return;
    setPeriodo({ mes, anio });
    setAnioTexto(String(anio));
    // Al cambiar el periodo los meses cambian, así que se descartan los archivos ya cargados
    // para no cruzar por error una nómina con el mes equivocado.
    setMeses(generarMeses(mes, anio, jornadaEmpresa, jornadasManuales));
    limpiarResultados();
  };

  const cambiarJornadaEmpresa = (valor) => {
    setJornadaEmpresa(valor);
    // Los meses que ya tengan un ajuste manual (jornadasManuales) NO se
    // pisan: solo se recalculan los que todavía dependían del valor global.
    setMeses((prev) => generarMeses(periodo.mes, periodo.anio, valor, jornadasManuales, prev));
    guardarConfigSiHayEmpresa({ jornadaEmpresa: valor }, { inmediato: true });
    if (Object.keys(resultadosPorMes).length > 0) setNecesitaReprocesar(true);
  };

  const actualizarMes = (id, campo, valor) => {
    setMeses((prev) => prev.map((m) => (m.id === id ? { ...m, [campo]: valor } : m)));
    if (campo === "jornadaSemanal") {
      const mes = meses.find((m) => m.id === id);
      if (mes) {
        const siguienteManuales = { ...jornadasManuales, [claveMes(mes.anio, mes.mesIndex)]: Number(valor) };
        setJornadasManuales(siguienteManuales);
        guardarConfigSiHayEmpresa({ jornadasPorMes: siguienteManuales }, { inmediato: true });
      }
    }
    if (Object.keys(resultadosPorMes).length > 0) setNecesitaReprocesar(true);
  };

  const actualizarDatosEmpresa = (campo, valor) => {
    const siguiente = { ...datosEmpresa, [campo]: valor };
    setDatosEmpresa(siguiente);
    guardarConfigSiHayEmpresa({ datosEmpresa: siguiente });
  };

  const actualizarBaseDias = (valor) => {
    setBaseDias(valor);
    guardarConfigSiHayEmpresa({ baseDias: valor }, { inmediato: true });
    if (Object.keys(resultadosPorMes).length > 0) setNecesitaReprocesar(true);
  };

  // ---------------- Homologación ----------------
  const aplicarCambioHomologacion = (clave, cambios) => {
    const siguiente = { ...homologacion, [clave]: { ...homologacion[clave], ...cambios } };
    setHomologacion(siguiente);
    guardarConfigSiHayEmpresa({ homologacion: siguiente });
  };

  const confirmarTodas = () => {
    const siguiente = { ...homologacion };
    Object.entries(siguiente).forEach(([k, v]) => {
      // Solo se confirman en bloque los que ya tienen código o quedaron
      // marcados como "no calificado" a mano; los de baja cobertura o sin
      // ninguna pista se dejan pendientes para revisión individual (punto B.2).
      if (homologacionListaParaConfirmar(v)) siguiente[k] = { ...v, confirmado: true };
    });
    setHomologacion(siguiente);
    guardarConfigSiHayEmpresa({ homologacion: siguiente }, { inmediato: true });
  };

  // ---------------- Corrección manual (personas que solo aparecen en SS) ----------------
  const actualizarAsignacionManual = (documento, cargoTexto) => {
    const siguiente = { ...asignacionesManuales, [documento]: cargoTexto };
    setAsignacionesManuales(siguiente);
    guardarConfigSiHayEmpresa({ asignacionesManuales: siguiente });
  };

  // ---------------- Procesamiento ----------------
  const procesarTodosLosMeses = async () => {
    const conArchivos = meses.filter((m) => m.nomina && m.segSocial);
    if (conArchivos.length === 0) {
      setError("Carga al menos un mes con nómina y Seguridad Social.");
      return;
    }

    setCargando(true);
    setError(null);
    setAvisos([]);

    try {
      const nuevosAvisos = [];
      const base = baseDias === "30" ? 30 : "real";

      // 1) Cada archivo se lee UNA sola vez
      const leidos = [];
      for (const mes of conArchivos) {
        let nom;
        let ss;
        try {
          nom = parsearNomina(await leerFilas(mes.nomina));
        } catch (e) {
          throw new Error(`${mes.etiqueta} · Nómina (${mes.nomina.name}): ${e.message}`);
        }
        try {
          ss = parsearSeguridadSocial(await leerFilas(mes.segSocial));
        } catch (e) {
          throw new Error(`${mes.etiqueta} · Seguridad Social (${mes.segSocial.name}): ${e.message}`);
        }
        leidos.push({ mes, nom, ss });
        if (ss.filasSubtotalIgnoradas > 0) {
          nuevosAvisos.push(`${mes.etiqueta}: se ignoraron ${ss.filasSubtotalIgnoradas} fila(s) de subtotal/total por ciudad en Seguridad Social.`);
        }
      }

      // Aviso de jornada dividida a mitad de mes (ej. julio: cambia el día 15)
      conArchivos.forEach((mes) => {
        if (mes.jornadaDividida) nuevosAvisos.push(`${mes.etiqueta}: ${mes.jornadaDividida.mensaje}`);
      });

      // 2) Resolver, por cada mes, quiénes están en SS pero no quedaron activos
      // en la nómina de ESE mes (retirados a mitad de mes, etc. — punto A.2).
      // Se busca su cargo en la nómina de los OTROS meses cargados; si no
      // aparece en ninguna, queda pendiente de asignación manual.
      const resolucionesPorMes = leidos.map(({ mes, nom, ss }) => {
        const docsNominaMes = new Set(nom.empleados.map((e) => e.documento));
        const otrasNominas = leidos
          .filter((l) => l.mes.id !== mes.id)
          .map((l) => ({ etiqueta: l.mes.etiqueta, empleados: l.nom.empleados }));

        const agregados = [];
        const pendientes = [];

        for (const [documento, registroSS] of ss.documentos) {
          if (docsNominaMes.has(documento)) continue; // ya está activo en la nómina de este mes

          const encontrado = buscarCargoEnOtrasNominas(documento, otrasNominas);
          const manual = asignacionesManuales[documento];
          const cargoTexto = encontrado?.cargo || manual || "";

          const proporcion = calcularProporcionPorFechas(registroSS.fechaIngreso, registroSS.fechaRetiro, mes.anio, mes.mesIndex, base);
          if (proporcion <= 0) continue; // sus fechas de SS tampoco caen en este mes

          if (cargoTexto) {
            agregados.push({
              documento,
              documentoOriginal: documento,
              nombre: "",
              cargo: cargoTexto,
              cargoKey: claveCargo(cargoTexto),
              fechaIngreso: registroSS.fechaIngreso,
              fechaRetiro: registroSS.fechaRetiro,
              proporcion,
              trabajoCompleto: proporcion >= 1,
              enSegSocial: true,
              novedad: encontrado ? `Tomado de la nómina de ${encontrado.etiqueta} (no aparece en la nómina de ${mes.etiqueta}).` : "Cargo asignado a mano.",
              agregadoDesdeSS: true,
            });
          } else {
            pendientes.push({ documento, fechaIngreso: registroSS.fechaIngreso, fechaRetiro: registroSS.fechaRetiro, proporcion });
          }
        }
        return { agregados, pendientes };
      });

      // 3) Cargos únicos de todos los meses (nómina original + los resueltos desde SS)
      const cargos = new Map();
      leidos.forEach(({ nom }) =>
        nom.empleados.forEach((e) => {
          if (!cargos.has(e.cargoKey)) cargos.set(e.cargoKey, e.cargo);
        })
      );
      resolucionesPorMes.forEach(({ agregados }) =>
        agregados.forEach((e) => {
          if (!cargos.has(e.cargoKey)) cargos.set(e.cargoKey, e.cargo);
        })
      );

      // 4) Lo ya confirmado (traído del servidor al cargar la empresa, o
      // confirmado en esta misma sesión) se reutiliza y no se vuelve a traducir.
      const previo = { ...homologacion };
      const porTraducir = [...cargos.keys()].filter((k) => !previo[k]?.es && k !== "(SIN CARGO)");

      let traducciones = {};
      if (porTraducir.length > 0) {
        try {
          traducciones = await traducirCargosConIA(porTraducir);
          const sinTraducir = porTraducir.filter((k) => !traducciones[k]).length;
          if (sinTraducir > 0) nuevosAvisos.push(`${sinTraducir} cargo(s) no vinieron traducidos; se dejó el texto original. Puedes editarlos abajo.`);
        } catch (e) {
          nuevosAvisos.push(
            `No se pudo traducir con IA (${e.message}). ¿Está corriendo server.js en ${API_URL}? Se dejó el cargo original; puedes editar cada traducción abajo.`
          );
        }
      }

      const nuevaHomologacion = construirHomologacion(cargos, traducciones, previo, INDICE_CNO);

      // 5) Resultado por mes
      const nuevos = {};
      leidos.forEach(({ mes, nom, ss }, idx) => {
        const { agregados, pendientes } = resolucionesPorMes[idx];
        const vistos = new Set();
        const duplicados = [];
        const fuera = new Map();
        const empleados = [];
        const diferenciasProporcion = [];

        for (const e of nom.empleados) {
          if (vistos.has(e.documento)) {
            duplicados.push(e.documento);
            continue;
          }
          vistos.add(e.documento);

          const registroSS = ss.documentos.get(e.documento);
          // La nómina GTN no siempre trae fecha de retiro fiable: se completa con SS (punto A.3).
          const fechaRetiro = fusionarFechaRetiro(e, registroSS);
          const proporcion = calcularProporcionPorFechas(e.fechaIngreso, fechaRetiro, mes.anio, mes.mesIndex, base);
          if (proporcion <= 0) {
            fuera.set(e.documento, { ...e, fechaRetiro });
            continue;
          }

          // Cruce informativo: proporción por fechas vs. proporción por horas
          // efectivas de SS (punto A.6). Como mínimo se avisa cuando difieren;
          // la decisión de cuál usar para la matriz queda para los abogados (F).
          let proporcionHoras = null;
          if (registroSS && registroSS.horasLaboradas != null) {
            proporcionHoras = calcularProporcionPorHoras(registroSS.horasLaboradas, mes.jornadaSemanal, mes.anio, mes.mesIndex);
            if (proporcionHoras !== null && Math.abs(proporcionHoras - proporcion) > TOLERANCIA_DIFERENCIA_PROPORCION) {
              diferenciasProporcion.push({ ...e, proporcionFechas: proporcion, proporcionHoras });
            }
          }

          empleados.push({
            ...e,
            fechaRetiro,
            proporcion,
            proporcionHoras,
            trabajoCompleto: proporcion >= 1,
            enSegSocial: ss.documentos.has(e.documento),
            novedad: describirNovedad({ ...e, fechaRetiro }, mes.anio, mes.mesIndex),
          });
        }

        // Se agregan quienes solo aparecían en SS y sí se les pudo asignar cargo.
        empleados.push(...agregados);

        const docsNomina = new Set([...vistos, ...agregados.map((a) => a.documento)]);
        const sinMatchSS = empleados.filter((e) => !e.enSegSocial);
        const soloEnSS = [...ss.documentos.keys()]
          .filter((d) => !docsNomina.has(d))
          .map((d) => {
            const e = fuera.get(d);
            const enPendientes = pendientes.some((p) => p.documento === d);
            return {
              documento: d,
              nota: e
                ? `Está en nómina, pero sus fechas lo dejan fuera del mes (ingreso ${formatFecha(e.fechaIngreso) || "s/f"}, retiro ${formatFecha(e.fechaRetiro) || "s/f"}).`
                : enPendientes
                ? "Retirado/activo según SS, sin cargo conocido: pendiente de asignación manual."
                : "",
            };
          });

        if (duplicados.length > 0) nuevosAvisos.push(`${mes.etiqueta}: ${duplicados.length} documento(s) repetido(s) en la nómina; se contó solo la primera fila.`);
        if (empleados.length === 0) nuevosAvisos.push(`${mes.etiqueta}: no quedó ningún trabajador activo. Revisa las fechas de ingreso/retiro y el mes.`);
        if (pendientes.length > 0) nuevosAvisos.push(`${mes.etiqueta}: ${pendientes.length} persona(s) en Seguridad Social sin cargo conocido; usa "Corrección manual" para asignarles uno.`);
        if (agregados.length > 0) nuevosAvisos.push(`${mes.etiqueta}: se agregaron ${agregados.length} persona(s) retirada(s)/no vigente(s) en la nómina de este mes, usando su cargo de otro mes.`);

        nuevos[mes.etiqueta] = {
          etiqueta: mes.etiqueta,
          anio: mes.anio,
          mesIndex: mes.mesIndex,
          jornadaSemanal: mes.jornadaSemanal, // se congela el valor usado al procesar
          totalNomina: empleados.length,
          totalSegSocial: ss.documentos.size,
          cuotaAprendicesRequerida: calcularCuotaAprendices(empleados.length),
          cuadra: empleados.length === ss.documentos.size,
          empleados,
          sinMatchSS,
          soloEnSS,
          duplicados,
          fueraDelMes: fuera.size,
          pendientesAsignacion: pendientes,
          diferenciasProporcion,
        };
      });

      setHomologacion(nuevaHomologacion);
      setResultadosPorMes(nuevos);
      setMesActivo(conArchivos[0].etiqueta);
      setAvisos(nuevosAvisos);
      setNecesitaReprocesar(false);
    } catch (err) {
      setError(err.message || "Error al procesar los archivos");
    } finally {
      setCargando(false);
    }
  };

  // ---------------- Exportación ----------------
  const listaResultados = Object.values(resultadosPorMes); // ya en orden cronológico
  const pendientes = Object.values(homologacion).filter((h) => !h.confirmado).length;

  // Documentos pendientes de asignación manual, únicos en toda la ventana de meses.
  const pendientesAsignacionGlobal = useMemo(() => {
    const mapa = new Map();
    listaResultados.forEach((r) => {
      (r.pendientesAsignacion || []).forEach((p) => {
        if (!mapa.has(p.documento)) mapa.set(p.documento, { ...p, meses: [] });
        mapa.get(p.documento).meses.push(r.etiqueta);
      });
    });
    return [...mapa.values()];
  }, [resultadosPorMes]); // eslint-disable-line react-hooks/exhaustive-deps

  const exportarMatrizCompleta = async () => {
    if (listaResultados.length === 0) return;
    if (necesitaReprocesar && !window.confirm("Cambiaste la configuración después de procesar (jornada, base de días, etc.). ¿Exportar de todas formas sin reprocesar?")) return;
    if (pendientes > 0 && !window.confirm(`Hay ${pendientes} cargo(s) sin confirmar en la homologación. ¿Exportar de todas formas?`)) return;

    setExportando(true);
    setError(null);
    try {
      const { advertenciaLogo } = await exportarMatrizExcel({
        nombreEmpresa: nombreEmpresa.trim(),
        resultados: listaResultados,
        homologacion,
        baseDias: baseDias === "30" ? 30 : "real",
        nombreDeCodigo: (codigo) => nombreDeCodigo(codigo, INDICE_CNO),
        datosEmpresa,
        aprendicesActivos: Number(aprendicesActivos) || 0,
      });
      if (advertenciaLogo) setAvisos((prev) => [...prev, advertenciaLogo]);
    } catch (e) {
      setError("Error al exportar: " + (e.message || e));
    } finally {
      setExportando(false);
    }
  };

  // ---------------- Datos derivados para la vista ----------------
  const resultadoActivo = mesActivo ? resultadosPorMes[mesActivo] : null;
  const resultadoMasReciente = listaResultados.length > 0 ? listaResultados[listaResultados.length - 1] : null;
  const cuotaVigente = resultadoMasReciente ? resultadoMasReciente.cuotaAprendicesRequerida : null;
  const aprendicesNum = Number(aprendicesActivos) || 0;

  const cantidadPorCargo = useMemo(() => {
    const max = {};
    listaResultados.forEach((r) => {
      const conteo = {};
      r.empleados.forEach((e) => {
        conteo[e.cargoKey] = (conteo[e.cargoKey] || 0) + 1;
      });
      Object.entries(conteo).forEach(([k, n]) => {
        max[k] = Math.max(max[k] || 0, n);
      });
    });
    return max;
  }, [resultadosPorMes]); // eslint-disable-line react-hooks/exhaustive-deps

  const homologacionOrdenada = useMemo(
    () => Object.entries(homologacion).sort((a, b) => a[1].es.localeCompare(b[1].es, "es")),
    [homologacion]
  );

  const grupoActivoResumen = useMemo(() => {
    if (!resultadoActivo || resultadoActivo.empleados.length === 0) return [];
    // Misma función que usa exportarExcel.js, para que esto refleje EXACTAMENTE
    // lo que se va a exportar (punto C.1).
    return agruparPorCargo(resultadoActivo.empleados, homologacion, resultadoActivo.jornadaSemanal);
  }, [resultadoActivo, homologacion]);

  const rangoPeriodo = meses.length === 6 ? `${meses[0].etiqueta} a ${meses[5].etiqueta}` : "";
  const mesesListos = meses.filter((m) => m.nomina && m.segSocial).length;

  // ---------------- Estado del flujo, para la barra de pasos ----------------
  const pasos = [
    {
      id: "seccion-empresa",
      titulo: "1. Empresa",
      detalle: nombreEmpresa.trim() || "Sin nombre aún",
      completo: !!nombreEmpresa.trim(),
    },
    {
      id: "seccion-meses",
      titulo: "2. Periodo y archivos",
      detalle: `${mesesListos} de 6 meses listos`,
      completo: mesesListos === 6,
    },
    {
      id: "seccion-homologacion",
      titulo: "3. Homologación",
      detalle: listaResultados.length === 0 ? "Aún sin procesar" : pendientes === 0 ? "Todo confirmado" : `${pendientes} por confirmar`,
      completo: listaResultados.length > 0 && pendientes === 0,
    },
    {
      id: "seccion-resultados",
      titulo: "4. Resultados y exportar",
      detalle: listaResultados.length === 0 ? "Pendiente" : "Listo para exportar",
      completo: listaResultados.length > 0,
    },
  ];

  return (
    <div className="min-h-screen bg-slate-50/50 font-sans text-slate-800 antialiased selection:bg-[#003B7A]/10 [&_section]:scroll-mt-24">
      {/* Header Corporativo */}
      <header className="bg-white border-b border-slate-200/80 sticky top-0 z-50 backdrop-blur-md bg-white/90">
        <div className="max-w-6xl mx-auto px-6 py-3.5 flex items-center justify-between">
          <img src="/logo.jpeg" alt="Solutions & Payroll Logo" className="h-9 w-auto object-contain" />
          <span className="text-xs font-semibold px-3 py-1 rounded-full bg-slate-100 text-slate-600 border border-slate-200/60 flex items-center gap-1.5">
            <Building2 className="w-3.5 h-3.5 text-[#003B7A]" /> Módulo SENA
          </span>
        </div>
      </header>

      {/* Banner principal + guía, a todo lo ancho */}
      <div className="max-w-6xl mx-auto px-6 pt-10 space-y-6">
        <section className="text-center space-y-3">
          <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-blue-50 border border-blue-100 text-[#003B7A] text-xs font-semibold">
            <Sparkles className="w-3.5 h-3.5" /> Normalización inteligente con IA
          </div>
          <h1 className="text-3xl sm:text-4xl font-extrabold text-slate-900 tracking-tight">
            Gestión de Planta y Matriz SENA
          </h1>
          <p className="text-sm text-slate-500 max-w-2xl mx-auto leading-relaxed">
            Elige el periodo de presentación, carga nómina y Seguridad Social de los 6 meses, homologa los cargos con el Listado de Oficios y Ocupaciones y genera la matriz para la determinación de cuota.
          </p>
        </section>

        <div className="bg-blue-50/60 border border-blue-100 rounded-2xl p-5 transition-all">
          <button
            onClick={() => setMostrarInstrucciones(!mostrarInstrucciones)}
            className="w-full flex items-center justify-between text-left font-bold text-[#003B7A] text-sm cursor-pointer"
          >
            <span className="flex items-center gap-2">
              <HelpCircle className="w-4 h-4 text-[#003B7A]" />
              Guía rápida de proceso
            </span>
            {mostrarInstrucciones ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
          </button>
          {mostrarInstrucciones && (
            <div className="mt-4 pt-4 border-t border-blue-100 space-y-3">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs text-slate-600 font-medium">
                <div className="p-3 bg-white/80 rounded-xl border border-blue-50/80 shadow-2xs space-y-1">
                  <span className="font-bold text-[#003B7A] block">1. Periodo y archivos</span>
                  Indica la empresa y el mes en que se presenta la matriz (enero/julio o marzo/septiembre según el SENA). La app calcula los 6 meses anteriores y tú subes la nómina y la Seguridad Social de cada uno.
                </div>
                <div className="p-3 bg-white/80 rounded-xl border border-blue-50/80 shadow-2xs space-y-1">
                  <span className="font-bold text-[#003B7A] block">2. Configuración</span>
                  La jornada semanal se sugiere según la ley vigente en cada mes, pero puedes fijarla por empresa (ej. 40 h) o ajustarla mes a mes; esos ajustes ya no se pierden al cambiar la configuración general.
                </div>
                <div className="p-3 bg-white/80 rounded-xl border border-blue-50/80 shadow-2xs space-y-1">
                  <span className="font-bold text-[#003B7A] block">3. Homologación por cargo</span>
                  Cada cargo se traduce y se le asigna un código una sola vez para todos los meses. Revisa, cambia con el buscador y confirma; la app recuerda lo confirmado para la próxima vez.
                </div>
                <div className="p-3 bg-white/80 rounded-xl border border-blue-50/80 shadow-2xs space-y-1">
                  <span className="font-bold text-[#003B7A] block">4. Cruce, corrección y exportación</span>
                  Revisa las diferencias con Seguridad Social, asigna a mano el cargo de quien no aparezca en ninguna nómina, y exporta el libro Excel con logo y listo para imprimir.
                </div>
              </div>
              <div className="p-3 bg-amber-50 rounded-xl border border-amber-200/80 text-xs text-amber-800 font-medium flex items-start gap-2">
                <AlertTriangle className="w-4 h-4 shrink-0 mt-0.5 text-amber-600" />
                Verifica periódicamente que <code className="font-mono">listado-cno.json</code> corresponda a la versión vigente publicada por el SENA: el listado de homologación cambia con cada nueva resolución.
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Cuerpo principal: barra de pasos fija a la izquierda + contenido */}
      <main className="max-w-6xl mx-auto px-6 py-10 grid grid-cols-1 lg:grid-cols-[224px_1fr] gap-8 items-start">
        <BarraDePasos pasos={pasos} pasoActivo={pasos.find((p) => !p.completo)?.id ?? pasos[pasos.length - 1].id} />

        <div className="space-y-8 min-w-0">
          {/* PASO 1 — Datos de la empresa + panel de empresas guardadas */}
          <section id="seccion-empresa" className="grid grid-cols-1 xl:grid-cols-[1fr_280px] gap-5 items-start">
            <div className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-5">
              <EncabezadoSeccion
                numero={1}
                icono={Building2}
                titulo="Empresa y configuración general"
                descripcion="Identifica la empresa y define la jornada y la base de días que se usarán en todos los cálculos."
                extra={cargandoEmpresaSeleccionada && <Loader2 className="w-3.5 h-3.5 text-slate-400 animate-spin" />}
              />
              <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-700">Nombre corto de la empresa</label>
                  <input
                    type="text"
                    placeholder="Ej. CIPY S.A.S."
                    value={nombreEmpresa}
                    onChange={(e) => setNombreEmpresa(e.target.value)}
                    onBlur={() => cargarEmpresaDesdeServidor()}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs font-semibold text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                  />
                  <p className="text-[11px] text-slate-400">
                    Al salir del campo, si ya existe en el servidor se carga automáticamente. Se usa en el nombre del archivo y como llave para recordar la configuración de esta empresa.
                  </p>
                </div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                    <Clock className="w-3.5 h-3.5 text-slate-500" /> Jornada semanal de la empresa
                  </label>
                  <select
                    value={jornadaEmpresa}
                    onChange={(e) => cambiarJornadaEmpresa(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                  >
                    <option value="auto">Legal vigente según el mes (automática)</option>
                    {JORNADAS.map((h) => (
                      <option key={h} value={h}>
                        {h} horas fijas para todos los meses
                      </option>
                    ))}
                  </select>
                  <p className="text-[11px] text-slate-400">Puedes ajustarla mes a mes más abajo; esos ajustes se conservan aunque cambies este valor.</p>
                </div>
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                    <Calendar className="w-3.5 h-3.5 text-slate-500" /> Base de días para ingresos/retiros
                  </label>
                  <select
                    value={baseDias}
                    onChange={(e) => actualizarBaseDias(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                  >
                    <option value="real">Días reales del mes (28 a 31)</option>
                    <option value="30">30 días (como Seguridad Social)</option>
                  </select>
                  <p className="text-[11px] text-slate-400">Cambia la proporción de quien entra o sale a mitad de mes. Confirma el criterio con tus abogados.</p>
                </div>
              </div>

              <div className="pt-4 border-t border-slate-100 space-y-3">
                <p className="text-[11px] font-extrabold text-slate-500 uppercase tracking-wider">Datos para el encabezado del Excel (opcional pero recomendado)</p>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  {[
                    ["razonSocial", "Razón social"],
                    ["nit", "NIT"],
                    ["representanteLegal", "Representante legal"],
                    ["cc", "CC representante legal"],
                    ["direccion", "Dirección"],
                    ["telefonos", "Teléfonos"],
                    ["email", "E-mail"],
                  ].map(([campo, etiqueta]) => (
                    <div key={campo} className="space-y-1">
                      <label className="text-[11px] font-bold text-slate-600">{etiqueta}</label>
                      <input
                        type="text"
                        value={datosEmpresa[campo]}
                        onChange={(e) => actualizarDatosEmpresa(campo, e.target.value)}
                        className="w-full bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                      />
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <PanelEmpresas
              empresas={empresasGuardadas}
              cargandoLista={cargandoListaEmpresas}
              cargandoSeleccion={cargandoEmpresaSeleccionada}
              guardando={guardandoEmpresa}
              error={errorEmpresas}
              nombreActual={nombreEmpresa}
              onRecargar={cargarListaEmpresas}
              onCargar={cargarEmpresaDesdeServidor}
              onEliminar={eliminarEmpresaGuardada}
              onGuardar={guardarEmpresaCompleta}
              onNuevaEmpresa={nuevaEmpresaEnBlanco}
            />
          </section>

          {/* PASO 2 — Carga de Meses */}
          <section id="seccion-meses" className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-6">
            <EncabezadoSeccion
              numero={2}
              icono={Calendar}
              titulo="Periodo y archivos mensuales"
              descripcion="Elige el mes de presentación y sube la nómina y la Seguridad Social de cada uno de los 6 meses anteriores."
              extra={<span className="text-xs text-slate-400 font-medium">{mesesListos} de 6 meses con archivos completos</span>}
            />

            {/* Periodo de presentación */}
            <div className="grid grid-cols-1 md:grid-cols-12 gap-4 items-end bg-slate-50/70 border border-slate-200/70 rounded-xl p-4">
              <div className="md:col-span-4 space-y-1.5">
                <label className="text-[11px] uppercase font-extrabold text-slate-500 tracking-wider">Mes de presentación</label>
                <select
                  value={periodo.mes}
                  onChange={(e) => cambiarPeriodo(Number(e.target.value), periodo.anio)}
                  className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-xs font-semibold text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                >
                  {NOMBRES_MES.map((nombre, i) => (
                    <option key={nombre} value={i}>
                      {nombre}
                      {MESES_PERIODO_SENA.includes(i) ? " (periodo SENA)" : ""}
                    </option>
                  ))}
                </select>
              </div>
              <div className="md:col-span-2 space-y-1.5">
                <label className="text-[11px] uppercase font-extrabold text-slate-500 tracking-wider">Año</label>
                <input
                  type="number"
                  value={anioTexto}
                  onChange={(e) => {
                    setAnioTexto(e.target.value);
                    cambiarPeriodo(periodo.mes, Number(e.target.value));
                  }}
                  className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-xs font-semibold text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                />
              </div>
              <p className="md:col-span-6 text-xs text-slate-500 leading-relaxed">
                Se reportan los 6 meses anteriores: <span className="font-bold text-slate-700">{rangoPeriodo}</span>. Al cambiar el periodo se limpian los archivos cargados.
              </p>
            </div>

            <div className="space-y-3">
              {meses.map((mes) => (
                <div
                  key={mes.id}
                  className={`p-4 rounded-xl border transition-all ${
                    mes.nomina && mes.segSocial ? "bg-slate-50/50 border-slate-200" : "bg-white border-slate-200/70 hover:border-slate-300"
                  }`}
                >
                  <div className="grid grid-cols-1 md:grid-cols-12 gap-4 items-center">
                    {/* Mes */}
                    <div className="md:col-span-3">
                      <p className="text-sm font-extrabold text-slate-800">{mes.etiqueta}</p>
                      <p className="text-[11px] text-slate-400">{new Date(mes.anio, mes.mesIndex + 1, 0).getDate()} días</p>
                      {mes.jornadaDividida && (
                        <p className="text-[10px] text-amber-700 font-semibold mt-0.5">⚠ Jornada cambia el día {mes.jornadaDividida.diaCambio}</p>
                      )}
                    </div>

                    {/* Carga Archivo Nómina */}
                    <div className="md:col-span-3 space-y-1.5">
                      <label className="text-[11px] font-bold text-slate-600 flex items-center gap-1">
                        <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" /> Nómina
                      </label>
                      <label
                        className={`flex items-center justify-between px-3 py-2 rounded-lg border text-xs cursor-pointer transition ${
                          mes.nomina ? "bg-emerald-50/50 border-emerald-200 text-emerald-800" : "bg-slate-50 border-slate-200 text-slate-400 hover:bg-slate-100/80"
                        }`}
                      >
                        <span className="truncate max-w-[140px]">{mes.nomina ? mes.nomina.name : "Seleccionar..."}</span>
                        <Upload className="w-3.5 h-3.5 shrink-0 opacity-60" />
                        <input
                          type="file"
                          accept=".xlsx,.xls,.csv"
                          onChange={(e) => actualizarMes(mes.id, "nomina", e.target.files?.[0] || null)}
                          className="hidden"
                        />
                      </label>
                    </div>

                    {/* Carga Archivo Seg Social */}
                    <div className="md:col-span-3 space-y-1.5">
                      <label className="text-[11px] font-bold text-slate-600 flex items-center gap-1">
                        <FileText className="w-3.5 h-3.5 text-blue-600" /> Seguridad Social
                      </label>
                      <label
                        className={`flex items-center justify-between px-3 py-2 rounded-lg border text-xs cursor-pointer transition ${
                          mes.segSocial ? "bg-blue-50/50 border-blue-200 text-blue-800" : "bg-slate-50 border-slate-200 text-slate-400 hover:bg-slate-100/80"
                        }`}
                      >
                        <span className="truncate max-w-[140px]">{mes.segSocial ? mes.segSocial.name : "Seleccionar..."}</span>
                        <Upload className="w-3.5 h-3.5 shrink-0 opacity-60" />
                        <input
                          type="file"
                          accept=".xlsx,.xls,.csv"
                          onChange={(e) => actualizarMes(mes.id, "segSocial", e.target.files?.[0] || null)}
                          className="hidden"
                        />
                      </label>
                    </div>

                    {/* Jornada semanal por periodo */}
                    <div className="md:col-span-2 space-y-1.5">
                      <label className="text-[11px] font-bold text-slate-600 flex items-center gap-1">
                        <Clock className="w-3.5 h-3.5 text-slate-500" /> Jornada
                      </label>
                      <select
                        value={mes.jornadaSemanal}
                        onChange={(e) => actualizarMes(mes.id, "jornadaSemanal", Number(e.target.value))}
                        className="w-full bg-slate-50 border border-slate-200 rounded-lg px-2 py-2 text-xs font-semibold text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                      >
                        {[...new Set([...JORNADAS, mes.jornadaSemanal])]
                          .sort((a, b) => a - b)
                          .map((h) => (
                            <option key={h} value={h}>
                              {h}h
                            </option>
                          ))}
                      </select>
                    </div>

                    {/* Estado */}
                    <div className="md:col-span-1 flex md:justify-end items-center">
                      {mes.nomina && mes.segSocial ? (
                        <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-800">
                          <Check className="w-3 h-3 stroke-[3]" />
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-[10px] font-medium bg-slate-100 text-slate-400">—</span>
                      )}
                    </div>
                  </div>
                </div>
              ))}
            </div>

            {/* Configuración de aprendices */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-4 border-t border-slate-100">
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                  <Users className="w-4 h-4 text-slate-500" /> Aprendices activos actualmente:
                </label>
                <input
                  type="number"
                  min="0"
                  value={aprendicesActivos}
                  onChange={(e) => setAprendicesActivos(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs font-semibold text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                />
                <p className="text-[11px] text-slate-400">Se compara con la cuota del último mes procesado: {TEXTO_REGLA_CUOTA_APRENDICES} Confírmalo con tus abogados.</p>
              </div>
              {cuotaVigente !== null && (
                <div
                  className={`rounded-xl border p-4 flex items-center gap-3 text-xs font-bold ${
                    aprendicesNum >= cuotaVigente ? "bg-emerald-50 border-emerald-200 text-emerald-800" : "bg-amber-50 border-amber-200 text-amber-800"
                  }`}
                >
                  {aprendicesNum >= cuotaVigente ? <CheckCircle2 className="w-5 h-5 shrink-0" /> : <AlertTriangle className="w-5 h-5 shrink-0" />}
                  <span>
                    Cuota requerida ({resultadoMasReciente.etiqueta}): {cuotaVigente} aprendiz(es). Actualmente reportas {aprendicesNum}.
                    {aprendicesNum < cuotaVigente ? " Faltan por cubrir." : " Cuota cubierta."}
                  </span>
                </div>
              )}
            </div>

            {necesitaReprocesar && (
              <div className="p-3 bg-amber-50 border border-amber-200 rounded-xl text-xs font-semibold text-amber-800 flex items-center gap-2">
                <RefreshCw className="w-4 h-4 shrink-0" /> Cambiaste algo después de procesar (jornada, base de días, archivos). Vuelve a procesar para que la matriz refleje el cambio.
              </div>
            )}

            <button
              onClick={procesarTodosLosMeses}
              disabled={cargando}
              className="w-full bg-[#003B7A] hover:bg-[#002B5B] disabled:bg-slate-200 disabled:text-slate-400 text-white font-bold py-3.5 rounded-xl transition-all shadow-md hover:shadow-lg disabled:shadow-none text-sm flex items-center justify-center gap-2 cursor-pointer"
            >
              {cargando ? (
                <>
                  <Sparkles className="w-4 h-4 animate-spin text-amber-300" /> Procesando y traduciendo cargos...
                </>
              ) : listaResultados.length > 0 ? (
                <>
                  <RefreshCw className="w-4 h-4 text-blue-200" /> Reprocesar todos los meses
                </>
              ) : (
                <>
                  <Sparkles className="w-4 h-4 text-blue-200" /> Procesar todos los meses
                </>
              )}
            </button>
          </section>

          {/* Notificaciones */}
          {error && (
            <div className="p-4 bg-rose-50 border border-rose-200 text-rose-800 rounded-xl text-xs font-semibold flex items-center gap-3 shadow-2xs">
              <AlertTriangle className="w-5 h-5 shrink-0 text-rose-600" />
              <p>{error}</p>
            </div>
          )}
          {avisos.length > 0 && (
            <div className="p-4 bg-amber-50 border border-amber-200 text-amber-800 rounded-xl text-xs font-medium flex items-start gap-3 shadow-2xs">
              <AlertTriangle className="w-5 h-5 shrink-0 text-amber-600" />
              <ul className="space-y-1 list-disc pl-4">
                {avisos.map((a, i) => (
                  <li key={i}>{a}</li>
                ))}
              </ul>
            </div>
          )}

          {/* Corrección manual: personas de SS sin cargo en ninguna nómina */}
          {pendientesAsignacionGlobal.length > 0 && (
            <section className="bg-white border border-amber-200 rounded-2xl shadow-xs overflow-hidden">
              <div className="p-5 border-b border-amber-100 bg-amber-50/60">
                <h2 className="font-bold text-amber-900 text-base flex items-center gap-2">
                  <Wrench className="w-4 h-4" /> Corrección manual: {pendientesAsignacionGlobal.length} persona(s) sin cargo conocido
                </h2>
                <p className="text-[11px] text-amber-800/80 mt-1">
                  Están en Seguridad Social pero no aparecen en la nómina de ningún mes cargado. Escribe su cargo (o carga la nómina del mes en que sí estaban activos) y vuelve a procesar.
                </p>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-100/70 text-[10px] uppercase text-slate-500">
                    <tr>
                      <th className="py-2 px-3">Documento</th>
                      <th className="py-2 px-3">Meses afectados</th>
                      <th className="py-2 px-3 min-w-[260px]">Cargo (a mano)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {pendientesAsignacionGlobal.map((p) => (
                      <tr key={p.documento}>
                        <td className="py-2 px-3 font-mono text-slate-600">{p.documento}</td>
                        <td className="py-2 px-3 text-slate-500">{p.meses.join(", ")}</td>
                        <td className="py-2 px-3">
                          <input
                            type="text"
                            value={asignacionesManuales[p.documento] || ""}
                            onChange={(e) => actualizarAsignacionManual(p.documento, e.target.value)}
                            placeholder="Escribe el cargo…"
                            className="w-full bg-slate-50 border border-slate-200 rounded-lg px-2 py-1.5 text-xs font-medium text-slate-800 focus:outline-hidden focus:ring-2 focus:ring-amber-300 focus:border-amber-400"
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="p-4 flex justify-end">
                <button
                  onClick={procesarTodosLosMeses}
                  className="text-xs font-bold text-white bg-amber-600 hover:bg-amber-700 px-4 py-2 rounded-lg cursor-pointer flex items-center gap-2"
                >
                  <RefreshCw className="w-3.5 h-3.5" /> Aplicar y reprocesar
                </button>
              </div>
            </section>
          )}

          {/* Resultados */}
          {listaResultados.length > 0 && (
            <>
              {/* PASO 3 — Homologación por cargo */}
              <section id="seccion-homologacion" className="bg-white border border-slate-200 rounded-2xl shadow-xs overflow-hidden">
                <div className="p-5 border-b border-slate-100 bg-slate-50/50">
                  <EncabezadoSeccion
                    numero={3}
                    icono={ListChecks}
                    titulo="Homologación de cargos"
                    descripcion="Se aprueba una sola vez y aplica a todos los meses. Revisa cada sugerencia antes de confirmar."
                    extra={
                      <div className="flex items-center gap-3">
                        <span
                          className={`text-[11px] font-bold px-2.5 py-1 rounded-full border ${
                            pendientes === 0 ? "bg-emerald-50 text-emerald-800 border-emerald-200" : "bg-amber-50 text-amber-800 border-amber-200"
                          }`}
                        >
                          {pendientes === 0 ? "Todos confirmados" : `${pendientes} por confirmar`}
                        </span>
                        {pendientes > 0 && (
                          <button
                            onClick={confirmarTodas}
                            className="text-[11px] font-bold text-white bg-[#003B7A] hover:bg-[#002B5B] px-3 py-1.5 rounded-lg cursor-pointer"
                          >
                            Confirmar todos (excepto baja cobertura)
                          </button>
                        )}
                      </div>
                    }
                  />
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-100/70 text-slate-500 uppercase text-[10px] font-extrabold tracking-wider border-b border-slate-200">
                      <tr>
                        <th className="py-3 px-3">Cargo original</th>
                        <th className="py-3 px-3 min-w-[200px]">Cargo en español</th>
                        <th className="py-3 px-3 min-w-[240px]">Código CNO</th>
                        <th className="py-3 px-3 text-center">Personas</th>
                        <th className="py-3 px-3">Acción</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 text-slate-700">
                      {homologacionOrdenada.map(([clave, h]) => (
                        <FilaHomologacion key={clave} clave={clave} h={h} cantidad={cantidadPorCargo[clave] || 0} onCambiar={aplicarCambioHomologacion} />
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>

              {/* PASO 4 — Detalle por mes y exportación */}
              <section id="seccion-resultados" className="space-y-4">
                <EncabezadoSeccion
                  numero={4}
                  icono={Briefcase}
                  titulo="Resultados por mes y exportación"
                  descripcion="Revisa cada mes, resuelve las diferencias con Seguridad Social y exporta la matriz final."
                />

                <div className="flex gap-2 flex-wrap border-b border-slate-200/80 pb-2">
                  {listaResultados.map((r) => (
                    <button
                      key={r.etiqueta}
                      onClick={() => setMesActivo(r.etiqueta)}
                      className={`px-4 py-2 rounded-xl text-xs font-bold border transition-all flex items-center gap-2 cursor-pointer ${
                        mesActivo === r.etiqueta ? "bg-[#003B7A] text-white border-[#003B7A] shadow-xs" : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50"
                      }`}
                    >
                      {r.etiqueta}
                      {r.cuadra ? <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" /> : <AlertTriangle className="w-3.5 h-3.5 text-amber-400" />}
                    </button>
                  ))}
                </div>

                {resultadoActivo && (
                  <div className="bg-white border border-slate-200 rounded-2xl shadow-xs overflow-hidden">
                    <div className="p-5 border-b border-slate-100 space-y-2 bg-slate-50/50">
                      <h3 className="font-bold text-slate-900 text-base flex items-center gap-2">
                        {nombreEmpresa ? `${nombreEmpresa} — ` : ""}
                        {mesActivo}
                        <span className="text-[11px] font-semibold text-slate-400 bg-slate-100 px-2 py-0.5 rounded-full">
                          Jornada: {resultadoActivo.jornadaSemanal}h
                        </span>
                      </h3>
                      <div className="flex flex-wrap gap-2">
                        {resultadoActivo.cuadra ? (
                          <div className="inline-flex items-center gap-2 text-xs text-emerald-800 font-bold bg-emerald-50 border border-emerald-200/60 px-3 py-1 rounded-md">
                            <CheckCircle2 className="w-4 h-4 text-emerald-600" /> Nómina ({resultadoActivo.totalNomina}) coincide con Seguridad Social ({resultadoActivo.totalSegSocial})
                          </div>
                        ) : (
                          <div className="inline-flex items-center gap-2 text-xs text-rose-800 font-bold bg-rose-50 border border-rose-200/60 px-3 py-1 rounded-md">
                            <AlertTriangle className="w-4 h-4 text-rose-600" /> Discrepancia: Nómina ({resultadoActivo.totalNomina}) vs Seguridad Social ({resultadoActivo.totalSegSocial})
                          </div>
                        )}
                        <div className="inline-flex items-center gap-2 text-xs text-slate-600 font-semibold bg-slate-100 border border-slate-200/60 px-3 py-1 rounded-md">
                          <Briefcase className="w-4 h-4 text-slate-500" /> Cuota de aprendices requerida este mes: {resultadoActivo.cuotaAprendicesRequerida}
                        </div>
                        {resultadoActivo.fueraDelMes > 0 && (
                          <div className="inline-flex items-center text-xs text-slate-600 font-semibold bg-slate-100 border border-slate-200/60 px-3 py-1 rounded-md">
                            {resultadoActivo.fueraDelMes} fila(s) de nómina fuera del mes por fechas
                          </div>
                        )}
                        {resultadoActivo.diferenciasProporcion?.length > 0 && (
                          <div className="inline-flex items-center text-xs text-amber-700 font-semibold bg-amber-50 border border-amber-200/60 px-3 py-1 rounded-md">
                            {resultadoActivo.diferenciasProporcion.length} persona(s) con diferencia entre proporción por fechas y por horas SS
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Resumen por cargo: lo que se va a exportar */}
                    <div className="border-b border-slate-100 px-5 py-4">
                      <p className="text-[11px] font-extrabold text-slate-500 uppercase tracking-wider mb-2">Resumen por cargo (lo que se exportará)</p>
                      <ResumenPorCargo grupos={grupoActivoResumen} />
                    </div>

                    {/* Cruce con Seguridad Social */}
                    {(resultadoActivo.sinMatchSS.length > 0 || resultadoActivo.soloEnSS.length > 0) && (
                      <details open={!resultadoActivo.cuadra} className="border-b border-slate-100 px-5 py-3 text-xs">
                        <summary className="font-bold text-slate-700 cursor-pointer">
                          Diferencias con Seguridad Social ({resultadoActivo.sinMatchSS.length} en nómina sin planilla · {resultadoActivo.soloEnSS.length} en planilla sin nómina)
                        </summary>
                        <div className="mt-3 grid grid-cols-1 lg:grid-cols-2 gap-5">
                          <div className="space-y-2">
                            <p className="font-bold text-slate-600">En nómina, pero no en Seguridad Social</p>
                            {resultadoActivo.sinMatchSS.length === 0 ? (
                              <p className="text-slate-400">Ninguno.</p>
                            ) : (
                              <div className="max-h-56 overflow-y-auto rounded-lg border border-slate-200">
                                <table className="w-full text-left">
                                  <thead className="bg-slate-100/70 text-[10px] uppercase text-slate-500 sticky top-0">
                                    <tr>
                                      <th className="py-2 px-3">Documento</th>
                                      <th className="py-2 px-3">Empleado</th>
                                      <th className="py-2 px-3">Ingreso</th>
                                      <th className="py-2 px-3">Retiro</th>
                                      <th className="py-2 px-3">Posible causa</th>
                                    </tr>
                                  </thead>
                                  <tbody className="divide-y divide-slate-100">
                                    {resultadoActivo.sinMatchSS.map((e) => (
                                      <tr key={e.documento}>
                                        <td className="py-2 px-3 font-mono text-slate-500">{e.documentoOriginal || e.documento}</td>
                                        <td className="py-2 px-3 font-semibold text-slate-800">{e.nombre}</td>
                                        <td className="py-2 px-3">{formatFecha(e.fechaIngreso) || "—"}</td>
                                        <td className="py-2 px-3">{formatFecha(e.fechaRetiro) || "—"}</td>
                                        <td className="py-2 px-3 text-amber-700">{e.novedad || "Revisar"}</td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                            )}
                          </div>
                          <div className="space-y-2">
                            <p className="font-bold text-slate-600">En Seguridad Social, pero no en nómina</p>
                            {resultadoActivo.soloEnSS.length === 0 ? (
                              <p className="text-slate-400">Ninguno.</p>
                            ) : (
                              <div className="max-h-56 overflow-y-auto rounded-lg border border-slate-200">
                                <table className="w-full text-left">
                                  <thead className="bg-slate-100/70 text-[10px] uppercase text-slate-500 sticky top-0">
                                    <tr>
                                      <th className="py-2 px-3">Documento</th>
                                      <th className="py-2 px-3">Nota</th>
                                    </tr>
                                  </thead>
                                  <tbody className="divide-y divide-slate-100">
                                    {resultadoActivo.soloEnSS.map((e) => (
                                      <tr key={e.documento}>
                                        <td className="py-2 px-3 font-mono text-slate-500">{e.documento}</td>
                                        <td className="py-2 px-3 text-amber-700">{e.nota || "No aparece en la nómina de este mes."}</td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                            )}
                          </div>
                        </div>
                      </details>
                    )}

                    <div className="overflow-x-auto max-h-[480px] overflow-y-auto">
                      <table className="w-full text-left text-xs">
                        <thead className="bg-slate-100/70 text-slate-500 uppercase text-[10px] font-extrabold tracking-wider border-b border-slate-200 sticky top-0">
                          <tr>
                            <th className="py-3 px-4">Documento</th>
                            <th className="py-3 px-4">Empleado</th>
                            <th className="py-3 px-4">Cargo (original)</th>
                            <th className="py-3 px-4">Cargo (español)</th>
                            <th className="py-3 px-4">Proporción mes</th>
                            <th className="py-3 px-4">CNO</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100 text-slate-700">
                          {resultadoActivo.empleados.map((emp) => {
                            const h = homologacion[emp.cargoKey];
                            return (
                              <tr key={emp.documento} className="hover:bg-slate-50/80 transition-colors">
                                <td className="py-3 px-4 font-mono text-slate-500">
                                  {emp.documentoOriginal || emp.documento}
                                  {emp.agregadoDesdeSS && <span className="ml-1 text-[9px] text-blue-600 font-bold" title={emp.novedad}>SS</span>}
                                </td>
                                <td className="py-3 px-4 font-bold text-slate-900">{emp.nombre}</td>
                                <td className="py-3 px-4 text-slate-400">{emp.cargo}</td>
                                <td className="py-3 px-4 text-slate-700 font-medium">{h?.es || emp.cargo}</td>
                                <td className="py-3 px-4 font-semibold">
                                  {emp.trabajoCompleto ? (
                                    <span className="inline-flex items-center gap-1 text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-full border border-emerald-200/50">
                                      <CheckCircle2 className="w-3 h-3" /> Completo
                                    </span>
                                  ) : (
                                    <span
                                      title={emp.novedad}
                                      className="inline-flex items-center gap-1 text-amber-700 bg-amber-50 px-2 py-0.5 rounded-full border border-amber-200/50"
                                    >
                                      {emp.proporcion.toFixed(2)}
                                    </span>
                                  )}
                                </td>
                                <td className="py-3 px-4">
                                  {h?.codigo ? (
                                    <span className={`font-mono font-bold ${h.confirmado ? "text-[#003B7A]" : "text-amber-700"}`}>{h.codigo}</span>
                                  ) : (
                                    <span className="text-slate-400">Sin código</span>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                <div className="flex justify-end pt-2">
                  <button
                    onClick={exportarMatrizCompleta}
                    disabled={exportando}
                    className="bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-200 disabled:text-slate-400 text-white text-xs font-bold px-6 py-3.5 rounded-xl transition shadow-md hover:shadow-lg disabled:shadow-none flex items-center gap-2 cursor-pointer"
                  >
                    <Download className="w-4 h-4" /> {exportando ? "Generando Excel..." : "Exportar Matriz SENA completa (.xlsx)"}
                  </button>
                </div>
              </section>
            </>
          )}
        </div>
      </main>
    </div>
  );
}