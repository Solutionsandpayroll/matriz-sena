import { useState, useMemo } from "react";
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
  UserRoundCheck,
} from "lucide-react";

import listadoCnoLocal from "./data/listado-cno.json";

import {
  NOMBRES_MES,
  MESES_PERIODO_SENA,
  claveCargo,
  formatFecha,
  jornadaLegalSugerida,
  calcularProporcionPorFechas,
  calcularCuotaAprendices,
  describirNovedad,
  parsearNomina,
  parsearSeguridadSocial,
  indexarListadoCno,
  sugerirCno,
  nombreDeCodigo,
  construirHomologacion,
  cargarHomologacionGuardada,
  guardarHomologacion,
} from "./utils/sena";

import { exportarMatrizExcel } from "./utils/exportarExcel";

const INDICE_CNO = indexarListadoCno(listadoCnoLocal || []);

const API_URL =
  import.meta.env?.VITE_API_URL || "http://localhost:3001";

const JORNADAS = [40, 42, 44, 46, 47, 48];

/* ============================================================
   UTILIDADES
   ============================================================ */

function normalizarDocumento(valor) {
  if (valor === null || valor === undefined) return "";

  return String(valor)
    .trim()
    .replace(/\s+/g, "")
    .replace(/\.0$/, "");
}

function redondearHoras(valor) {
  const numero = Number(valor);

  if (!Number.isFinite(numero)) return 0;

  return Math.round(numero);
}

function obtenerRegistrosSS(ss) {
  /*
   * La nueva versión de sena.js debe devolver:
   *
   * {
   *   registros: Map([
   *      ["documento", {
   *          documento,
   *          fechaIngreso,
   *          fechaRetiro,
   *          horasLaboradas
   *      }]
   *   ]),
   *   documentos: Set(...)
   * }
   *
   * Dejamos compatibilidad con la estructura anterior.
   */

  if (ss?.registros instanceof Map) {
    return ss.registros;
  }

  if (Array.isArray(ss?.registros)) {
    const mapa = new Map();

    ss.registros.forEach((r) => {
      const documento = normalizarDocumento(
        r.documento ?? r.documentoOriginal
      );

      if (!documento) return;

      mapa.set(documento, {
        ...r,
        documento,
        horasLaboradas: redondearHoras(
          r.horasLaboradas ??
            r.horas ??
            r.horasPila ??
            r.horasTrabajadas ??
            0
        ),
      });
    });

    return mapa;
  }

  /*
   * Compatibilidad con una versión que solamente devolvía Set.
   */
  if (ss?.documentos instanceof Set) {
    return new Map(
      [...ss.documentos].map((documento) => [
        normalizarDocumento(documento),
        {
          documento: normalizarDocumento(documento),
          fechaIngreso: null,
          fechaRetiro: null,
          horasLaboradas: 0,
        },
      ])
    );
  }

  if (ss instanceof Set) {
    return new Map(
      [...ss].map((documento) => [
        normalizarDocumento(documento),
        {
          documento: normalizarDocumento(documento),
          fechaIngreso: null,
          fechaRetiro: null,
          horasLaboradas: 0,
        },
      ])
    );
  }

  return new Map();
}

function obtenerRegistroSS(ssRegistros, documento) {
  if (!ssRegistros) return null;

  return (
    ssRegistros.get(normalizarDocumento(documento)) || null
  );
}

/*
 * Busca la información de una persona en las nóminas de todos
 * los meses procesados.
 */
function buscarEmpleadoEnOtrosMeses(
  documento,
  leidos,
  mesActualId
) {
  const doc = normalizarDocumento(documento);

  for (const item of leidos) {
    if (item.mes.id === mesActualId) continue;

    const empleado = item.nom.empleados.find(
      (e) => normalizarDocumento(e.documento) === doc
    );

    if (empleado) {
      return empleado;
    }
  }

  return null;
}

/*
 * Busca el cargo más reciente disponible para una persona.
 */
function buscarCargoHistorico(documento, leidos, mesActualId) {
  const encontrado = buscarEmpleadoEnOtrosMeses(
    documento,
    leidos,
    mesActualId
  );

  if (!encontrado) return null;

  return {
    ...encontrado,
    cargoKey:
      encontrado.cargoKey ||
      claveCargo(encontrado.cargo || "(SIN CARGO)"),
  };
}

/*
 * Determina la proporción usando primero las fechas.
 *
 * La Seguridad Social también aporta horas laboradas. Por eso
 * guardamos ambas mediciones y luego avisamos cuando difieren.
 */
function calcularProporcionConSS({
  empleado,
  registroSS,
  mes,
  base,
}) {
  const proporcionFechas = calcularProporcionPorFechas(
    empleado.fechaIngreso,
    empleado.fechaRetiro,
    mes.anio,
    mes.mesIndex,
    base
  );

  const horasSS = redondearHoras(
    registroSS?.horasLaboradas ??
      registroSS?.horas ??
      registroSS?.horasPila ??
      0
  );

  const jornada = Number(mes.jornadaSemanal) || 40;

  /*
   * Horas ordinarias aproximadas del mes según jornada.
   * 4.333 semanas promedio.
   */
  const horasOrdinariasMes = Math.round(
    jornada * 52 / 12
  );

  const proporcionHoras =
    horasSS > 0 && horasOrdinariasMes > 0
      ? Math.min(horasSS / horasOrdinariasMes, 1)
      : null;

  const diferencia =
    proporcionHoras !== null
      ? Math.abs(proporcionFechas - proporcionHoras)
      : 0;

  return {
    proporcionFechas,
    proporcionHoras,
    horasSS,
    horasOrdinariasMes,
    diferencia,
    hayDiferenciaHoras:
      proporcionHoras !== null && diferencia > 0.05,
  };
}

/*
 * Genera los seis meses anteriores al periodo de presentación.
 */
function generarMeses(
  mesPresentacion,
  anioPresentacion,
  jornadaEmpresa,
  previos
) {
  return Array.from({ length: 6 }, (_, i) => {
    const d = new Date(
      anioPresentacion,
      mesPresentacion - (6 - i),
      1
    );

    const anio = d.getFullYear();
    const mesIndex = d.getMonth();

    const anterior = previos?.[i];

    return {
      id: i,
      anio,
      mesIndex,
      etiqueta: `${NOMBRES_MES[mesIndex]} ${anio}`,

      /*
       * Si ya existía un ajuste manual para ese mes,
       * NO se pierde al cambiar la jornada general.
       */
      jornadaSemanal:
        anterior?.jornadaManual === true
          ? anterior.jornadaSemanal
          : jornadaEmpresa === "auto"
          ? jornadaLegalSugerida(anio, mesIndex)
          : Number(jornadaEmpresa),

      jornadaManual:
        anterior?.jornadaManual === true,

      nomina: anterior?.nomina ?? null,
      segSocial: anterior?.segSocial ?? null,
    };
  });
}

/* ============================================================
   TRADUCCIÓN DE CARGOS
   ============================================================ */

async function traducirCargosConIA(cargosUnicos) {
  const response = await fetch(
    `${API_URL}/api/traducir-cargos`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        cargosUnicos,
      }),
    }
  );

  if (!response.ok) {
    throw new Error(
      `El servidor respondió ${response.status}`
    );
  }

  const data = await response.json();

  const normalizadas = {};

  Object.entries(data.traducciones || {}).forEach(
    ([k, v]) => {
      if (v && String(v).trim()) {
        normalizadas[claveCargo(k)] = String(v).trim();
      }
    }
  );

  return normalizadas;
}

/* ============================================================
   LECTURA DE EXCEL
   ============================================================ */

async function leerFilas(archivo) {
  const buffer = await archivo.arrayBuffer();

  const libro = XLSX.read(buffer, {
    type: "array",
  });

  const hoja =
    libro.Sheets[libro.SheetNames[0]];

  return XLSX.utils.sheet_to_json(hoja, {
    header: 1,
    defval: null,
    raw: true,
  });
}

/* ============================================================
   HOMOLOGACIÓN
   ============================================================ */

function FilaHomologacion({
  clave,
  h,
  cantidad,
  onCambiar,
}) {
  const [abierto, setAbierto] = useState(false);

  const [consulta, setConsulta] = useState(
    h.es || h.original || ""
  );

  const resultados = useMemo(
    () =>
      abierto
        ? sugerirCno(
            consulta,
            INDICE_CNO,
            25
          )
        : [],
    [abierto, consulta]
  );

  const elegir = (codigo, ocupacion) => {
    onCambiar(clave, {
      codigo,
      ocupacion,
      confirmado: true,
      estadoCodigo: codigo
        ? "encontrado"
        : "no_calificado",
    });

    setAbierto(false);
  };

  return (
    <>
      <tr className="hover:bg-slate-50/80 transition-colors align-top">
        <td className="py-2.5 px-3 text-slate-400">
          {h.original}
        </td>

        <td className="py-2.5 px-3">
          <input
            type="text"
            value={h.es || ""}
            onChange={(e) =>
              onCambiar(clave, {
                es: e.target.value,
                confirmado: false,
              })
            }
            className="w-full bg-slate-50 border border-slate-200 rounded-lg px-2 py-1.5 text-xs font-medium text-slate-800 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
          />
        </td>

        <td className="py-2.5 px-3">
          {h.codigo ? (
            <div className="space-y-0.5">
              <span className="font-mono font-bold text-[#003B7A]">
                {h.codigo}
              </span>

              <p className="text-[11px] text-slate-500 leading-snug">
                {h.ocupacion ||
                  nombreDeCodigo(
                    h.codigo,
                    INDICE_CNO
                  )}
              </p>
            </div>
          ) : h.estadoCodigo ===
            "no_calificado" ? (
            <span className="text-amber-700 font-semibold">
              No calificado
            </span>
          ) : (
            <span className="text-rose-600 font-semibold">
              Sin código encontrado
            </span>
          )}
        </td>

        <td className="py-2.5 px-3 text-center font-semibold text-slate-600">
          {cantidad}
        </td>

        <td className="py-2.5 px-3">
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={() =>
                setAbierto(!abierto)
              }
              className="text-[11px] font-bold text-[#003B7A] hover:underline cursor-pointer"
            >
              {abierto ? "Cerrar" : "Cambiar"}
            </button>

            <label className="flex items-center gap-1.5 text-[11px] font-semibold text-slate-600 cursor-pointer">
              <input
                type="checkbox"
                checked={Boolean(h.confirmado)}
                onChange={(e) =>
                  onCambiar(clave, {
                    confirmado:
                      e.target.checked,
                  })
                }
                className="accent-[#003B7A]"
              />

              Confirmado
            </label>
          </div>
        </td>
      </tr>

      {abierto && (
        <tr className="bg-blue-50/40">
          <td
            colSpan={5}
            className="px-4 py-3"
          >
            <div className="space-y-2">
              <div className="relative">
                <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" />

                <input
                  type="text"
                  autoFocus
                  value={consulta}
                  onChange={(e) =>
                    setConsulta(e.target.value)
                  }
                  placeholder="Busca por cargo en español o por código"
                  className="w-full bg-white border border-slate-200 rounded-lg pl-8 pr-3 py-2 text-xs text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                />
              </div>

              <div className="max-h-56 overflow-y-auto rounded-lg border border-slate-200 bg-white divide-y divide-slate-100">
                {resultados.length === 0 && (
                  <p className="px-3 py-2 text-[11px] text-slate-400">
                    Sin coincidencias. Prueba con otra palabra o código.
                  </p>
                )}

                {resultados.map((r, i) => (
                  <button
                    type="button"
                    key={`${r.codigo}-${i}`}
                    onClick={() =>
                      elegir(
                        r.codigo,
                        r.ocupacion
                      )
                    }
                    className="w-full text-left px-3 py-1.5 text-xs hover:bg-blue-50 cursor-pointer flex gap-2"
                  >
                    <span className="font-mono font-bold text-[#003B7A] shrink-0">
                      {r.codigo}
                    </span>

                    <span className="text-slate-700">
                      {r.ocupacion}
                    </span>
                  </button>
                ))}
              </div>

              <button
                type="button"
                onClick={() =>
                  elegir("", "")
                }
                className="text-[11px] font-bold text-amber-700 hover:underline cursor-pointer"
              >
                Marcar como oficio no calificado
              </button>
            </div>
          </td>
        </tr>
      )}
    </>
  );
}

/* ============================================================
   COMPONENTE PRINCIPAL
   ============================================================ */

export default function App() {
  const hoy = new Date();

  /* ---------------- Empresa ---------------- */

  const [datosEmpresa, setDatosEmpresa] =
    useState({
      nombreEmpresa: "",
      nit: "",
      representanteLegal: "",
      documentoRepresentante: "",
      direccion: "",
      telefonos: "",
      email: "",
    });

  const nombreEmpresa =
    datosEmpresa.nombreEmpresa;

  /* ---------------- Configuración ---------------- */

  const [jornadaEmpresa, setJornadaEmpresa] =
    useState("auto");

  const [baseDias, setBaseDias] =
    useState("real");

  const [periodo, setPeriodo] = useState({
    mes: hoy.getMonth(),
    anio: hoy.getFullYear(),
  });

  const [anioTexto, setAnioTexto] =
    useState(
      String(hoy.getFullYear())
    );

  const [meses, setMeses] = useState(() =>
    generarMeses(
      hoy.getMonth(),
      hoy.getFullYear(),
      "auto"
    )
  );

  const [aprendicesActivos, setAprendicesActivos] =
    useState(0);

  /* ---------------- Resultados ---------------- */

  const [
    resultadosPorMes,
    setResultadosPorMes,
  ] = useState({});

  const [
    homologacion,
    setHomologacion,
  ] = useState({});

  const [
    cargando,
    setCargando,
  ] = useState(false);

  const [
    exportando,
    setExportando,
  ] = useState(false);

  const [
    error,
    setError,
  ] = useState(null);

  const [
    avisos,
    setAvisos,
  ] = useState([]);

  const [
    mostrarInstrucciones,
    setMostrarInstrucciones,
  ] = useState(true);

  const [
    mesActivo,
    setMesActivo,
  ] = useState(null);

  const [
    pendientesManuales,
    setPendientesManuales,
  ] = useState([]);

  /* ==========================================================
     CONFIGURACIÓN DE EMPRESA
     ========================================================== */

  const actualizarEmpresa = (
    campo,
    valor
  ) => {
    setDatosEmpresa((prev) => ({
      ...prev,
      [campo]: valor,
    }));
  };

  /* ==========================================================
     LIMPIAR
     ========================================================== */

  const limpiarResultados = () => {
    setResultadosPorMes({});
    setMesActivo(null);
    setAvisos([]);
    setError(null);
    setPendientesManuales([]);
  };

  /* ==========================================================
     PERIODO
     ========================================================== */

  const cambiarPeriodo = (
    mes,
    anio
  ) => {
    if (
      !Number.isInteger(anio) ||
      anio < 2000 ||
      anio > 2100
    ) {
      return;
    }

    setPeriodo({
      mes,
      anio,
    });

    setAnioTexto(
      String(anio)
    );

    setMeses(
      generarMeses(
        mes,
        anio,
        jornadaEmpresa
      )
    );

    limpiarResultados();
  };

  /* ==========================================================
     JORNADA
     ========================================================== */

  const cambiarJornadaEmpresa = (
    valor
  ) => {
    setJornadaEmpresa(valor);

    setMeses((prev) =>
      generarMeses(
        periodo.mes,
        periodo.anio,
        valor,
        prev
      )
    );

    if (
      Object.keys(resultadosPorMes)
        .length > 0
    ) {
      setAvisos((prev) => [
        ...prev,
        "Cambiaste la jornada de la empresa después de procesar. Debes volver a procesar los meses para que el cambio quede aplicado.",
      ]);
    }
  };

  const actualizarMes = (
    id,
    campo,
    valor
  ) => {
    setMeses((prev) =>
      prev.map((m) =>
        m.id === id
          ? {
              ...m,
              [campo]: valor,
              ...(campo ===
                "jornadaSemanal"
                ? {
                    jornadaManual:
                      true,
                  }
                : {}),
            }
          : m
      )
    );

    if (
      campo === "jornadaSemanal" &&
      Object.keys(resultadosPorMes)
        .length > 0
    ) {
      setAvisos((prev) => [
        ...prev,
        "Modificaste una jornada mensual después de procesar. Vuelve a procesar los meses antes de exportar.",
      ]);
    }
  };

  /* ==========================================================
     HOMOLOGACIÓN
     ========================================================== */

  const aplicarCambioHomologacion = (
    clave,
    cambios
  ) => {
    const actual =
      homologacion[clave] || {};

    const siguiente = {
      ...homologacion,

      [clave]: {
        ...actual,
        ...cambios,
      },
    };

    setHomologacion(
      siguiente
    );

    /*
     * Solamente guardamos si realmente existe
     * una empresa identificada.
     */
    if (
      nombreEmpresa.trim()
    ) {
      guardarHomologacion(
        nombreEmpresa.trim(),
        siguiente
      );
    }
  };

  /*
   * Ya NO confirma automáticamente cargos
   * que no tengan código.
   *
   * Esto evita que "Confirmar todos" termine
   * aprobando sugerencias de baja cobertura.
   */
  const confirmarTodas = () => {
    const siguiente =
      Object.fromEntries(
        Object.entries(
          homologacion
        ).map(([k, v]) => {
          if (
            v.codigo ||
            v.estadoCodigo ===
              "no_calificado"
          ) {
            return [
              k,
              {
                ...v,
                confirmado: true,
              },
            ];
          }

          return [
            k,
            {
              ...v,
              confirmado: false,
            },
          ];
        })
      );

    setHomologacion(
      siguiente
    );

    if (
      nombreEmpresa.trim()
    ) {
      guardarHomologacion(
        nombreEmpresa.trim(),
        siguiente
      );
    }
  };

  /* ==========================================================
     PROCESAMIENTO
     ========================================================== */

  const procesarTodosLosMeses =
    async () => {
      const conArchivos =
        meses.filter(
          (m) =>
            m.nomina &&
            m.segSocial
        );

      if (
        conArchivos.length === 0
      ) {
        setError(
          "Carga al menos un mes con nómina y Seguridad Social."
        );
        return;
      }

      setCargando(true);
      setError(null);
      setAvisos([]);

      try {
        const nuevosAvisos = [];

        const base =
          baseDias === "30"
            ? 30
            : "real";

        /* ====================================================
           1. LEER ARCHIVOS
           ==================================================== */

        const leidos = [];

        for (
          const mes of conArchivos
        ) {
          let nom;
          let ss;

          try {
            nom =
              parsearNomina(
                await leerFilas(
                  mes.nomina
                )
              );
          } catch (e) {
            throw new Error(
              `${mes.etiqueta} · Nómina (${mes.nomina.name}): ${e.message}`
            );
          }

          try {
            ss =
              parsearSeguridadSocial(
                await leerFilas(
                  mes.segSocial
                )
              );
          } catch (e) {
            throw new Error(
              `${mes.etiqueta} · Seguridad Social (${mes.segSocial.name}): ${e.message}`
            );
          }

          const registrosSS =
            obtenerRegistrosSS(
              ss
            );

          leidos.push({
            mes,
            nom,
            ss,
            registrosSS,
          });
        }

        /* ====================================================
           2. CARGOS ÚNICOS
           ==================================================== */

        const cargos =
          new Map();

        leidos.forEach(
          ({ nom }) => {
            nom.empleados.forEach(
              (e) => {
                if (
                  !cargos.has(
                    e.cargoKey
                  )
                ) {
                  cargos.set(
                    e.cargoKey,
                    e.cargo
                  );
                }
              }
            );
          }
        );

        /* ====================================================
           3. HOMOLOGACIÓN GUARDADA
           ==================================================== */

        const previo = {
          ...(nombreEmpresa.trim()
            ? cargarHomologacionGuardada(
                nombreEmpresa.trim()
              )
            : {}),
          ...homologacion,
        };

        const porTraducir =
          [
            ...cargos.keys(),
          ].filter(
            (k) =>
              !previo[k]?.es &&
              k !== "(SIN CARGO)"
          );

        let traducciones =
          {};

        if (
          porTraducir.length >
          0
        ) {
          try {
            traducciones =
              await traducirCargosConIA(
                porTraducir
              );

            const sinTraducir =
              porTraducir.filter(
                (k) =>
                  !traducciones[k]
              ).length;

            if (
              sinTraducir > 0
            ) {
              nuevosAvisos.push(
                `${sinTraducir} cargo(s) no vinieron traducidos; se dejó el texto original.`
              );
            }
          } catch (e) {
            nuevosAvisos.push(
              `No se pudo traducir con IA (${e.message}). ¿Está corriendo server.js en ${API_URL}? Se dejó el cargo original.`
            );
          }
        }

        const nuevaHomologacion =
          construirHomologacion(
            cargos,
            traducciones,
            previo,
            INDICE_CNO
          );

        /*
         * Los cargos que no tienen suficiente evidencia
         * deben quedar pendientes.
         */
        Object.entries(
          nuevaHomologacion
        ).forEach(
          ([k, h]) => {
            if (
              !h.codigo &&
              h.estadoCodigo !==
                "no_calificado"
            ) {
              nuevaHomologacion[
                k
              ] = {
                ...h,
                confirmado: false,
                estadoCodigo:
                  "sin_codigo_encontrado",
              };
            }
          }
        );

        /* ====================================================
           4. RESULTADOS POR MES
           ==================================================== */

        const nuevos = {};

        const pendientes =
          [];

        for (
          const {
            mes,
            nom,
            registrosSS,
          } of leidos
        ) {
          const vistos =
            new Set();

          const duplicados =
            [];

          const fuera =
            new Map();

          const empleados =
            [];

          /*
           * --------------------------------------------------
           * 4A. PERSONAS DE NÓMINA
           * --------------------------------------------------
           */

          for (
            const original of
              nom.empleados
          ) {
            const documento =
              normalizarDocumento(
                original.documento
              );

            if (!documento)
              continue;

            if (
              vistos.has(
                documento
              )
            ) {
              duplicados.push(
                documento
              );
              continue;
            }

            vistos.add(
              documento
            );

            const registroSS =
              obtenerRegistroSS(
                registrosSS,
                documento
              );

            /*
             * Si SS trae fecha de ingreso,
             * se prioriza para el cruce.
             */
            const fechaIngreso =
              registroSS?.fechaIngreso ||
              original.fechaIngreso ||
              null;

            /*
             * La fecha de retiro debe venir
             * de Seguridad Social.
             */
            const fechaRetiro =
              registroSS?.fechaRetiro ||
              original.fechaRetiro ||
              null;

            const empleado = {
              ...original,
              documento,
              fechaIngreso,
              fechaRetiro,
            };

            const calculo =
              calcularProporcionConSS(
                {
                  empleado,
                  registroSS,
                  mes,
                  base,
                }
              );

            if (
              calculo.proporcionFechas <=
              0
            ) {
              fuera.set(
                documento,
                {
                  ...empleado,
                  registroSS,
                }
              );

              continue;
            }

            empleados.push({
              ...empleado,

              proporcion:
                calculo.proporcionFechas,

              proporcionFechas:
                calculo.proporcionFechas,

              proporcionHoras:
                calculo.proporcionHoras,

              horasPila:
                calculo.horasSS,

              horasLaboradas:
                calculo.horasSS,

              horasOrdinariasMes:
                calculo.horasOrdinariasMes,

              diferenciaProporcion:
                calculo.diferencia,

              hayDiferenciaHoras:
                calculo.hayDiferenciaHoras,

              trabajoCompleto:
                calculo.proporcionFechas >=
                1,

              enSegSocial:
                Boolean(
                  registroSS
                ),

              fechaIngresoSS:
                registroSS?.fechaIngreso ||
                null,

              fechaRetiroSS:
                registroSS?.fechaRetiro ||
                null,

              novedad:
                describirNovedad(
                  empleado,
                  mes.anio,
                  mes.mesIndex
                ),
            });
          }

          /*
           * --------------------------------------------------
           * 4B. PERSONAS QUE SOLO ESTÁN EN SEGURIDAD SOCIAL
           *
           * Aquí está uno de los cambios importantes.
           * Ya NO las dejamos simplemente como una diferencia.
           * Intentamos incorporarlas.
           * --------------------------------------------------
           */

          const documentosNomina =
            new Set(
              empleados.map(
                (e) =>
                  normalizarDocumento(
                    e.documento
                  )
              )
            );

          const sinMatchSS =
            empleados.filter(
              (e) =>
                !e.enSegSocial
            );

          const soloEnSS =
            [];

          for (
            const [
              documento,
              registroSS,
            ] of registrosSS
          ) {
            if (
              documentosNomina.has(
                documento
              )
            ) {
              continue;
            }

            /*
             * Buscar el trabajador en
             * otros meses.
             */
            const historico =
              buscarCargoHistorico(
                documento,
                leidos,
                mes.id
              );

            const fechaIngreso =
              registroSS?.fechaIngreso ||
              historico?.fechaIngreso ||
              null;

            const fechaRetiro =
              registroSS?.fechaRetiro ||
              null;

            /*
             * Si no tenemos cargo, queda
             * pendiente para asignarlo manualmente.
             */
            if (!historico) {
              pendientes.push({
                mesId: mes.id,
                etiqueta:
                  mes.etiqueta,
                documento,
                nombre:
                  registroSS?.nombre ||
                  "",
                cargo:
                  "",
                fechaIngreso,
                fechaRetiro,
                horasLaboradas:
                  redondearHoras(
                    registroSS?.horasLaboradas ??
                      registroSS?.horas ??
                      0
                  ),
                registroSS,
              });

              soloEnSS.push({
                documento,
                nombre:
                  registroSS?.nombre ||
                  "",
                fechaIngreso,
                fechaRetiro,
                cargo: "",
                pendienteCargo:
                  true,
                nota:
                  "Está en Seguridad Social pero no aparece en ninguna nómina cargada. Debes asignar el cargo manualmente o cargar la nómina del mes anterior.",
              });

              continue;
            }

            /*
             * Ya tenemos cargo histórico.
             */
            const empleadoHistorico =
              {
                ...historico,
                documento,
                fechaIngreso,
                fechaRetiro,
              };

            const calculo =
              calcularProporcionConSS(
                {
                  empleado:
                    empleadoHistorico,
                  registroSS,
                  mes,
                  base,
                }
              );

            /*
             * Si por fechas no trabajó en el mes,
             * no lo incluimos.
             */
            if (
              calculo.proporcionFechas <=
              0
            ) {
              soloEnSS.push({
                documento,
                nombre:
                  empleadoHistorico.nombre ||
                  registroSS?.nombre ||
                  "",
                fechaIngreso,
                fechaRetiro,
                cargo:
                  empleadoHistorico.cargo,
                pendienteCargo:
                  false,
                nota:
                  `Está en Seguridad Social, pero las fechas calculadas dejan una proporción de 0. Ingreso: ${formatFecha(fechaIngreso) || "s/f"} · Retiro: ${formatFecha(fechaRetiro) || "s/f"}.`,
              });

              continue;
            }

            empleados.push({
              ...empleadoHistorico,

              esRetirado:
                true,

              vieneDeOtroMes:
                true,

              proporcion:
                calculo.proporcionFechas,

              proporcionFechas:
                calculo.proporcionFechas,

              proporcionHoras:
                calculo.proporcionHoras,

              horasPila:
                calculo.horasSS,

              horasLaboradas:
                calculo.horasSS,

              horasOrdinariasMes:
                calculo.horasOrdinariasMes,

              diferenciaProporcion:
                calculo.diferencia,

              hayDiferenciaHoras:
                calculo.hayDiferenciaHoras,

              trabajoCompleto:
                calculo.proporcionFechas >=
                1,

              enSegSocial:
                true,

              fechaIngresoSS:
                registroSS?.fechaIngreso ||
                null,

              fechaRetiroSS:
                registroSS?.fechaRetiro ||
                null,

              novedad:
                "Trabajador encontrado en Seguridad Social y recuperado desde una nómina de otro mes.",
            });
          }

          /*
           * --------------------------------------------------
           * 4C. HORAS Y ALERTAS
           * --------------------------------------------------
           */

          const diferenciasHoras =
            empleados.filter(
              (e) =>
                e.hayDiferenciaHoras
            );

          if (
            diferenciasHoras.length >
            0
          ) {
            nuevosAvisos.push(
              `${mes.etiqueta}: ${diferenciasHoras.length} trabajador(es) presentan diferencia entre la proporción por fechas y la proporción calculada con Horas Laboradas de Seguridad Social.`
            );
          }

          if (
            duplicados.length >
            0
          ) {
            nuevosAvisos.push(
              `${mes.etiqueta}: ${duplicados.length} documento(s) repetido(s) en la nómina; se contó solo la primera fila.`
            );
          }

          if (
            empleados.length ===
            0
          ) {
            nuevosAvisos.push(
              `${mes.etiqueta}: no quedó ningún trabajador después del cruce. Revisa las fechas y los archivos.`
            );
          }

          /*
           * --------------------------------------------------
           * 4D. CUOTA
           * --------------------------------------------------
           */

          const cuota =
            calcularCuotaAprendices(
              empleados.length
            );

          /*
           * --------------------------------------------------
           * 4E. RESULTADO
           * --------------------------------------------------
           */

          nuevos[
            mes.etiqueta
          ] = {
            etiqueta:
              mes.etiqueta,

            anio:
              mes.anio,

            mesIndex:
              mes.mesIndex,

            jornadaSemanal:
              mes.jornadaSemanal,

            jornadaManual:
              mes.jornadaManual,

            baseDias:
              base,

            totalNomina:
              empleados.length,

            totalSegSocial:
              registrosSS.size,

            cuotaAprendicesRequerida:
              cuota,

            aprendicesActivos:
              Number(
                aprendicesActivos
              ) || 0,

            cuadra:
              empleados.length ===
              registrosSS.size,

            empleados,

            sinMatchSS,

            soloEnSS,

            duplicados,

            fueraDelMes:
              fuera.size,

            diferenciasHoras:
              diferenciasHoras.length,

            horasTotalesPila:
              empleados.reduce(
                (suma, e) =>
                  suma +
                  redondearHoras(
                    e.horasPila
                  ),
                0
              ),
          };
        }

        /* ====================================================
           5. GUARDAR PENDIENTES
           ==================================================== */

        setPendientesManuales(
          pendientes
        );

        if (
          pendientes.length >
          0
        ) {
          nuevosAvisos.push(
            `${pendientes.length} trabajador(es) de Seguridad Social no pudieron relacionarse con una nómina cargada. Debes asignar el cargo manualmente o cargar una nómina anterior.`
          );
        }

        /* ====================================================
           6. GUARDAR RESULTADOS
           ==================================================== */

        setHomologacion(
          nuevaHomologacion
        );

        setResultadosPorMes(
          nuevos
        );

        setMesActivo(
          conArchivos[0]
            .etiqueta
        );

        setAvisos(
          nuevosAvisos
        );
      } catch (err) {
        setError(
          err.message ||
            "Error al procesar los archivos."
        );
      } finally {
        setCargando(false);
      }
    };

  /* ==========================================================
     CORRECCIÓN MANUAL DE PENDIENTES
     ========================================================== */

  const asignarCargoManual = (
    pendienteIndex,
    cargo
  ) => {
    setPendientesManuales(
      (prev) =>
        prev.map(
          (p, index) =>
            index ===
            pendienteIndex
              ? {
                  ...p,
                  cargo,
                  cargoKey:
                    claveCargo(
                      cargo
                    ),
                  pendienteCargo:
                    false,
                }
              : p
        )
    );
  };

  /* ==========================================================
     EXPORTACIÓN
     ========================================================== */

  const listaResultados =
    Object.values(
      resultadosPorMes
    );

  const pendientes =
    Object.values(
      homologacion
    ).filter(
      (h) =>
        !h.confirmado
    ).length;

  const exportarMatrizCompleta =
    async () => {
      if (
        listaResultados.length ===
        0
      ) {
        return;
      }

      if (
        pendientes > 0 &&
        !window.confirm(
          `Hay ${pendientes} cargo(s) sin confirmar en la homologación. ¿Exportar de todas formas?`
        )
      ) {
        return;
      }

      if (
        pendientesManuales.some(
          (p) =>
            p.pendienteCargo
        )
      ) {
        if (
          !window.confirm(
            "Hay trabajadores de Seguridad Social sin cargo asignado. Si exportas ahora podrían quedar fuera de la homologación. ¿Continuar?"
          )
        ) {
          return;
        }
      }

      setExportando(true);
      setError(null);

      try {
        const {
          advertenciaLogo,
        } =
          await exportarMatrizExcel({
            nombreEmpresa:
              nombreEmpresa.trim(),

            /*
             * NUEVO: datos completos de empresa.
             */
            datosEmpresa,

            resultados:
              listaResultados,

            homologacion,

            baseDias:
              baseDias === "30"
                ? 30
                : "real",

            aprendicesActivos:
              Number(
                aprendicesActivos
              ) || 0,

            pendientesManuales,

            nombreDeCodigo:
              (codigo) =>
                nombreDeCodigo(
                  codigo,
                  INDICE_CNO
                ),
          });

        if (
          advertenciaLogo
        ) {
          setAvisos(
            (prev) => [
              ...prev,
              advertenciaLogo,
            ]
          );
        }
      } catch (e) {
        setError(
          "Error al exportar: " +
            (e.message || e)
        );
      } finally {
        setExportando(false);
      }
    };

  /* ==========================================================
     DATOS DERIVADOS
     ========================================================== */

  const resultadoActivo =
    mesActivo
      ? resultadosPorMes[
          mesActivo
        ]
      : null;

  const resultadoMasReciente =
    listaResultados.length >
    0
      ? listaResultados[
          listaResultados.length - 1
        ]
      : null;

  const cuotaVigente =
    resultadoMasReciente
      ? resultadoMasReciente.cuotaAprendicesRequerida
      : null;

  const aprendicesNum =
    Number(
      aprendicesActivos
    ) || 0;

  /*
   * Cantidad de personas por cargo.
   */
  const cantidadPorCargo =
    useMemo(() => {
      const max = {};

      listaResultados.forEach(
        (r) => {
          const conteo = {};

          r.empleados.forEach(
            (e) => {
              conteo[
                e.cargoKey
              ] =
                (conteo[
                  e.cargoKey
                ] || 0) + 1;
            }
          );

          Object.entries(
            conteo
          ).forEach(
            ([k, n]) => {
              max[k] = Math.max(
                max[k] || 0,
                n
              );
            }
          );
        }
      );

      return max;
    }, [resultadosPorMes]);

  /*
   * Orden de homologaciones.
   */
  const homologacionOrdenada =
    useMemo(
      () =>
        Object.entries(
          homologacion
        ).sort(
          (a, b) =>
            (
              a[1].es ||
              ""
            ).localeCompare(
              b[1].es ||
                "",
              "es"
            )
        ),
      [homologacion]
    );

  /*
   * Resumen por cargo y mes.
   *
   * Esto es importante porque antes solamente
   * se mostraban las personas individuales.
   */
  const resumenPorCargo =
    useMemo(() => {
      const mapa = {};

      listaResultados.forEach(
        (resultado) => {
          resultado.empleados.forEach(
            (empleado) => {
              const key =
                empleado.cargoKey;

              if (
                !mapa[key]
              ) {
                mapa[key] = {
                  cargoKey:
                    key,
                  cargoOriginal:
                    empleado.cargo,
                  cargoEspanol:
                    homologacion[
                      key
                    ]?.es ||
                    empleado.cargo,
                  codigo:
                    homologacion[
                      key
                    ]?.codigo ||
                    "",
                  meses: {},
                };
              }

              mapa[key].meses[
                resultado.etiqueta
              ] =
                (mapa[key].meses[
                  resultado.etiqueta
                ] || 0) + 1;
            }
          );
        }
      );

      return Object.values(
        mapa
      );
    }, [
      resultadosPorMes,
      homologacion,
    ]);

  const rangoPeriodo =
    meses.length === 6
      ? `${meses[0].etiqueta} a ${meses[5].etiqueta}`
      : "";

  const mesesListos =
    meses.filter(
      (m) =>
        m.nomina &&
        m.segSocial
    ).length;

  /* ==========================================================
     RENDER
     ========================================================== */

  return (
    <div className="min-h-screen bg-slate-50/50 font-sans text-slate-800 antialiased selection:bg-[#003B7A]/10">
      {/* =====================================================
          HEADER
          ===================================================== */}

      <header className="bg-white border-b border-slate-200/80 sticky top-0 z-50 backdrop-blur-md bg-white/90">
        <div className="max-w-6xl mx-auto px-6 py-3.5 flex items-center justify-between">
          <img
            src="/logo.jpeg"
            alt="Solutions & Payroll Logo"
            className="h-9 w-auto object-contain"
          />

          <span className="text-xs font-semibold px-3 py-1 rounded-full bg-slate-100 text-slate-600 border border-slate-200/60 flex items-center gap-1.5">
            <Building2 className="w-3.5 h-3.5 text-[#003B7A]" />
            Módulo SENA
          </span>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-10 space-y-8">
        {/* ===================================================
            TITULO
            =================================================== */}

        <section className="text-center space-y-3">
          <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-blue-50 border border-blue-100 text-[#003B7A] text-xs font-semibold">
            <Sparkles className="w-3.5 h-3.5" />
            Normalización inteligente con IA
          </div>

          <h1 className="text-3xl sm:text-4xl font-extrabold text-slate-900 tracking-tight flex items-center justify-center gap-3">
            Gestión de Planta y Matriz SENA
          </h1>

          <p className="text-sm text-slate-500 max-w-2xl mx-auto leading-relaxed">
            Elige el periodo de presentación,
            carga nómina y Seguridad Social de
            los 6 meses, revisa los cruces,
            homologa los cargos y genera la
            matriz.
          </p>
        </section>

        {/* ===================================================
            INSTRUCCIONES
            =================================================== */}

        <div className="bg-blue-50/60 border border-blue-100 rounded-2xl p-5 transition-all">
          <button
            onClick={() =>
              setMostrarInstrucciones(
                !mostrarInstrucciones
              )
            }
            className="w-full flex items-center justify-between text-left font-bold text-[#003B7A] text-sm cursor-pointer"
          >
            <span className="flex items-center gap-2">
              <HelpCircle className="w-4 h-4" />
              Guía rápida de proceso
            </span>

            {mostrarInstrucciones ? (
              <ChevronUp className="w-4 h-4" />
            ) : (
              <ChevronDown className="w-4 h-4" />
            )}
          </button>

          {mostrarInstrucciones && (
            <div className="mt-4 pt-4 border-t border-blue-100 space-y-3">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs text-slate-600 font-medium">
                <div className="p-3 bg-white/80 rounded-xl border border-blue-50/80 shadow-2xs space-y-1">
                  <span className="font-bold text-[#003B7A] block">
                    1. Periodo y archivos
                  </span>
                  Carga la nómina y Seguridad
                  Social de cada uno de los seis
                  meses.
                </div>

                <div className="p-3 bg-white/80 rounded-xl border border-blue-50/80 shadow-2xs space-y-1">
                  <span className="font-bold text-[#003B7A] block">
                    2. Configuración
                  </span>
                  Define la jornada de la empresa
                  y puedes ajustarla individualmente
                  por mes.
                </div>

                <div className="p-3 bg-white/80 rounded-xl border border-blue-50/80 shadow-2xs space-y-1">
                  <span className="font-bold text-[#003B7A] block">
                    3. Cruce
                  </span>
                  La Seguridad Social aporta
                  ingreso, retiro y horas laboradas.
                  Los retirados se buscan en las
                  nóminas de otros meses.
                </div>

                <div className="p-3 bg-white/80 rounded-xl border border-blue-50/80 shadow-2xs space-y-1">
                  <span className="font-bold text-[#003B7A] block">
                    4. Homologación
                  </span>
                  Revisa manualmente los códigos
                  CNO antes de confirmar.
                </div>
              </div>

              <div className="p-3 bg-amber-50 rounded-xl border border-amber-200/80 text-xs text-amber-800 font-medium flex items-start gap-2">
                <AlertTriangle className="w-4 h-4 shrink-0 mt-0.5 text-amber-600" />

                <span>
                  Los códigos CNO deben revisarse
                  contra el listado vigente que
                  tengas autorizado para el proceso.
                </span>
              </div>
            </div>
          )}
        </div>

        {/* ===================================================
            DATOS DE EMPRESA
            =================================================== */}

        <section className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-5">
          <h2 className="text-sm font-bold text-slate-900 flex items-center gap-2 uppercase tracking-wider">
            <Building2 className="w-4 h-4 text-[#003B7A]" />
            Datos de la empresa
          </h2>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-700">
                Razón social
              </label>

              <input
                type="text"
                value={
                  datosEmpresa.nombreEmpresa
                }
                onChange={(e) =>
                  actualizarEmpresa(
                    "nombreEmpresa",
                    e.target.value
                  )
                }
                placeholder="Ej. SOLUTIONS & PAYROLL S.A.S."
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs font-semibold"
              />
            </div>

            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-700">
                NIT
              </label>

              <input
                type="text"
                value={
                  datosEmpresa.nit
                }
                onChange={(e) =>
                  actualizarEmpresa(
                    "nit",
                    e.target.value
                  )
                }
                placeholder="Ej. 900000000-1"
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs font-semibold"
              />
            </div>

            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-700">
                Representante legal
              </label>

              <input
                type="text"
                value={
                  datosEmpresa.representanteLegal
                }
                onChange={(e) =>
                  actualizarEmpresa(
                    "representanteLegal",
                    e.target.value
                  )
                }
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs font-semibold"
              />
            </div>

            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-700">
                CC representante legal
              </label>

              <input
                type="text"
                value={
                  datosEmpresa.documentoRepresentante
                }
                onChange={(e) =>
                  actualizarEmpresa(
                    "documentoRepresentante",
                    e.target.value
                  )
                }
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs font-semibold"
              />
            </div>

            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-700">
                Dirección
              </label>

              <input
                type="text"
                value={
                  datosEmpresa.direccion
                }
                onChange={(e) =>
                  actualizarEmpresa(
                    "direccion",
                    e.target.value
                  )
                }
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs font-semibold"
              />
            </div>

            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-700">
                Teléfonos
              </label>

              <input
                type="text"
                value={
                  datosEmpresa.telefonos
                }
                onChange={(e) =>
                  actualizarEmpresa(
                    "telefonos",
                    e.target.value
                  )
                }
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs font-semibold"
              />
            </div>

            <div className="space-y-1.5 md:col-span-2">
              <label className="text-xs font-bold text-slate-700">
                E-mail
              </label>

              <input
                type="email"
                value={
                  datosEmpresa.email
                }
                onChange={(e) =>
                  actualizarEmpresa(
                    "email",
                    e.target.value
                  )
                }
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs font-semibold"
              />
            </div>

            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-700">
                Jornada general
              </label>

              <select
                value={
                  jornadaEmpresa
                }
                onChange={(e) =>
                  cambiarJornadaEmpresa(
                    e.target.value
                  )
                }
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold"
              >
                <option value="auto">
                  Legal según el mes
                </option>

                {JORNADAS.map(
                  (h) => (
                    <option
                      key={h}
                      value={h}
                    >
                      {h} horas fijas
                    </option>
                  )
                )}
              </select>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-5 pt-2">
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                <Calendar className="w-3.5 h-3.5" />
                Base de días
              </label>

              <select
                value={baseDias}
                onChange={(e) =>
                  setBaseDias(
                    e.target.value
                  )
                }
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold"
              >
                <option value="real">
                  Días reales del mes
                </option>

                <option value="30">
                  30 días
                </option>
              </select>
            </div>

            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                <Users className="w-3.5 h-3.5" />
                Aprendices activos
              </label>

              <input
                type="number"
                min="0"
                value={
                  aprendicesActivos
                }
                onChange={(e) =>
                  setAprendicesActivos(
                    e.target.value
                  )
                }
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold"
              />
            </div>
          </div>
        </section>

        {/* ===================================================
            MESES
            =================================================== */}

        <section className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-6">
          <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 pb-4">
            <h2 className="text-sm font-bold text-slate-900 flex items-center gap-2 uppercase tracking-wider">
              <Calendar className="w-4 h-4 text-[#003B7A]" />
              Meses a reportar
            </h2>

            <span className="text-xs text-slate-400 font-medium">
              {mesesListos} de 6 meses
            </span>
          </div>

          {/* Periodo */}

          <div className="grid grid-cols-1 md:grid-cols-12 gap-4 items-end bg-slate-50/70 border border-slate-200/70 rounded-xl p-4">
            <div className="md:col-span-4 space-y-1.5">
              <label className="text-[11px] uppercase font-extrabold text-slate-500 tracking-wider">
                Mes de presentación
              </label>

              <select
                value={
                  periodo.mes
                }
                onChange={(e) =>
                  cambiarPeriodo(
                    Number(
                      e.target.value
                    ),
                    periodo.anio
                  )
                }
                className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-xs font-semibold"
              >
                {NOMBRES_MES.map(
                  (
                    nombre,
                    i
                  ) => (
                    <option
                      key={nombre}
                      value={i}
                    >
                      {nombre}
                      {MESES_PERIODO_SENA.includes(
                        i
                      )
                        ? " (periodo SENA)"
                        : ""}
                    </option>
                  )
                )}
              </select>
            </div>

            <div className="md:col-span-2 space-y-1.5">
              <label className="text-[11px] uppercase font-extrabold text-slate-500 tracking-wider">
                Año
              </label>

              <input
                type="number"
                value={
                  anioTexto
                }
                onChange={(e) => {
                  setAnioTexto(
                    e.target.value
                  );

                  cambiarPeriodo(
                    periodo.mes,
                    Number(
                      e.target.value
                    )
                  );
                }}
                className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-xs font-semibold"
              />
            </div>

            <p className="md:col-span-6 text-xs text-slate-500">
              Se reportan:
              <span className="font-bold text-slate-700 ml-1">
                {rangoPeriodo}
              </span>
            </p>
          </div>

          {/* Meses */}

          <div className="space-y-3">
            {meses.map(
              (mes) => (
                <div
                  key={mes.id}
                  className={`p-4 rounded-xl border ${
                    mes.nomina &&
                    mes.segSocial
                      ? "bg-slate-50/50 border-slate-200"
                      : "bg-white border-slate-200/70"
                  }`}
                >
                  <div className="grid grid-cols-1 md:grid-cols-12 gap-4 items-center">
                    <div className="md:col-span-3">
                      <p className="text-sm font-extrabold text-slate-800">
                        {mes.etiqueta}
                      </p>

                      <p className="text-[11px] text-slate-400">
                        {new Date(
                          mes.anio,
                          mes.mesIndex +
                            1,
                          0
                        ).getDate()}{" "}
                        días
                      </p>
                    </div>

                    {/* Nomina */}

                    <div className="md:col-span-3 space-y-1.5">
                      <label className="text-[11px] font-bold text-slate-600 flex items-center gap-1">
                        <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
                        Nómina
                      </label>

                      <label
                        className={`flex items-center justify-between px-3 py-2 rounded-lg border text-xs cursor-pointer ${
                          mes.nomina
                            ? "bg-emerald-50/50 border-emerald-200 text-emerald-800"
                            : "bg-slate-50 border-slate-200 text-slate-400"
                        }`}
                      >
                        <span className="truncate max-w-[140px]">
                          {mes.nomina
                            ? mes.nomina
                                .name
                            : "Seleccionar..."}
                        </span>

                        <Upload className="w-3.5 h-3.5 shrink-0" />

                        <input
                          type="file"
                          accept=".xlsx,.xls,.csv"
                          onChange={(
                            e
                          ) =>
                            actualizarMes(
                              mes.id,
                              "nomina",
                              e.target
                                .files?.[0] ||
                                null
                            )
                          }
                          className="hidden"
                        />
                      </label>
                    </div>

                    {/* Seguridad Social */}

                    <div className="md:col-span-3 space-y-1.5">
                      <label className="text-[11px] font-bold text-slate-600 flex items-center gap-1">
                        <FileText className="w-3.5 h-3.5 text-blue-600" />
                        Seguridad Social
                      </label>

                      <label
                        className={`flex items-center justify-between px-3 py-2 rounded-lg border text-xs cursor-pointer ${
                          mes.segSocial
                            ? "bg-blue-50/50 border-blue-200 text-blue-800"
                            : "bg-slate-50 border-slate-200 text-slate-400"
                        }`}
                      >
                        <span className="truncate max-w-[140px]">
                          {mes.segSocial
                            ? mes
                                .segSocial
                                .name
                            : "Seleccionar..."}
                        </span>

                        <Upload className="w-3.5 h-3.5 shrink-0" />

                        <input
                          type="file"
                          accept=".xlsx,.xls,.csv"
                          onChange={(
                            e
                          ) =>
                            actualizarMes(
                              mes.id,
                              "segSocial",
                              e.target
                                .files?.[0] ||
                                null
                            )
                          }
                          className="hidden"
                        />
                      </label>
                    </div>

                    {/* Jornada */}

                    <div className="md:col-span-2 space-y-1.5">
                      <label className="text-[11px] font-bold text-slate-600 flex items-center gap-1">
                        <Clock className="w-3.5 h-3.5 text-slate-500" />
                        Jornada
                      </label>

                      <select
                        value={
                          mes.jornadaSemanal
                        }
                        onChange={(
                          e
                        ) =>
                          actualizarMes(
                            mes.id,
                            "jornadaSemanal",
                            Number(
                              e.target
                                .value
                            )
                          )
                        }
                        className="w-full bg-slate-50 border border-slate-200 rounded-lg px-2 py-2 text-xs font-semibold"
                      >
                        {[
                          ...new Set([
                            ...JORNADAS,
                            mes.jornadaSemanal,
                          ]),
                        ]
                          .sort(
                            (
                              a,
                              b
                            ) =>
                              a -
                              b
                          )
                          .map(
                            (h) => (
                              <option
                                key={h}
                                value={
                                  h
                                }
                              >
                                {h}h
                              </option>
                            )
                          )}
                      </select>

                      {mes.jornadaManual && (
                        <p className="text-[10px] text-blue-700 font-semibold">
                          Ajuste manual
                        </p>
                      )}
                    </div>

                    <div className="md:col-span-1 flex md:justify-end">
                      {mes.nomina &&
                      mes.segSocial ? (
                        <CheckCircle2 className="w-5 h-5 text-emerald-600" />
                      ) : (
                        <span className="text-slate-300">
                          —
                        </span>
                      )}
                    </div>
                  </div>
                </div>
              )
            )}
          </div>

          <button
            onClick={
              procesarTodosLosMeses
            }
            disabled={cargando}
            className="w-full bg-[#003B7A] hover:bg-[#002B5B] disabled:bg-slate-200 disabled:text-slate-400 text-white font-bold py-3.5 rounded-xl transition-all shadow-md text-sm flex items-center justify-center gap-2 cursor-pointer"
          >
            {cargando ? (
              <>
                <Sparkles className="w-4 h-4 animate-spin" />
                Procesando...
              </>
            ) : (
              <>
                <Sparkles className="w-4 h-4" />
                Procesar todos los meses
              </>
            )}
          </button>
        </section>

        {/* ===================================================
            ERRORES
            =================================================== */}

        {error && (
          <div className="p-4 bg-rose-50 border border-rose-200 text-rose-800 rounded-xl text-xs font-semibold flex items-center gap-3">
            <AlertTriangle className="w-5 h-5 shrink-0" />
            <p>{error}</p>
          </div>
        )}

        {avisos.length > 0 && (
          <div className="p-4 bg-amber-50 border border-amber-200 text-amber-800 rounded-xl text-xs font-medium flex items-start gap-3">
            <AlertTriangle className="w-5 h-5 shrink-0" />

            <ul className="space-y-1 list-disc pl-4">
              {avisos.map(
                (a, i) => (
                  <li key={i}>
                    {a}
                  </li>
                )
              )}
            </ul>
          </div>
        )}

        {/* ===================================================
            PENDIENTES MANUALES
            =================================================== */}

        {pendientesManuales.length >
          0 && (
          <section className="bg-white border border-amber-200 rounded-2xl shadow-xs overflow-hidden">
            <div className="p-5 border-b border-amber-100 bg-amber-50/60">
              <h2 className="font-bold text-amber-900 text-base flex items-center gap-2">
                <UserRoundCheck className="w-5 h-5" />
                Trabajadores pendientes de asignación
              </h2>

              <p className="text-xs text-amber-700 mt-1">
                Estas personas aparecen en Seguridad
                Social pero no se encontraron en
                ninguna de las nóminas cargadas.
                Puedes asignarles el cargo manualmente
                o cargar una nómina anterior.
              </p>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-100 text-slate-500 uppercase text-[10px] font-extrabold">
                  <tr>
                    <th className="py-3 px-4">
                      Documento
                    </th>
                    <th className="py-3 px-4">
                      Nombre
                    </th>
                    <th className="py-3 px-4">
                      Ingreso
                    </th>
                    <th className="py-3 px-4">
                      Retiro
                    </th>
                    <th className="py-3 px-4">
                      Horas
                    </th>
                    <th className="py-3 px-4">
                      Cargo manual
                    </th>
                  </tr>
                </thead>

                <tbody className="divide-y divide-slate-100">
                  {pendientesManuales.map(
                    (
                      p,
                      index
                    ) => (
                      <tr
                        key={`${p.documento}-${index}`}
                      >
                        <td className="py-3 px-4 font-mono">
                          {
                            p.documento
                          }
                        </td>

                        <td className="py-3 px-4 font-semibold">
                          {p.nombre ||
                            "Sin nombre"}
                        </td>

                        <td className="py-3 px-4">
                          {formatFecha(
                            p.fechaIngreso
                          ) ||
                            "—"}
                        </td>

                        <td className="py-3 px-4">
                          {formatFecha(
                            p.fechaRetiro
                          ) ||
                            "—"}
                        </td>

                        <td className="py-3 px-4 font-semibold">
                          {redondearHoras(
                            p.horasLaboradas
                          )}
                        </td>

                        <td className="py-3 px-4">
                          <input
                            type="text"
                            value={
                              p.cargo ||
                              ""
                            }
                            onChange={(
                              e
                            ) =>
                              asignarCargoManual(
                                index,
                                e.target
                                  .value
                              )
                            }
                            placeholder="Escribe el cargo"
                            className="w-full min-w-[220px] bg-white border border-slate-200 rounded-lg px-3 py-2 text-xs"
                          />
                        </td>
                      </tr>
                    )
                  )}
                </tbody>
              </table>
            </div>
          </section>
        )}

        {/* ===================================================
            RESULTADOS
            =================================================== */}

        {listaResultados.length >
          0 && (
          <>
            {/* HOMOLOGACION */}

            <section className="bg-white border border-slate-200 rounded-2xl shadow-xs overflow-hidden">
              <div className="p-5 border-b border-slate-100 bg-slate-50/50 flex flex-wrap items-center justify-between gap-3">
                <div>
                  <h2 className="font-bold text-slate-900 text-base flex items-center gap-2">
                    <ListChecks className="w-4 h-4 text-[#003B7A]" />
                    Homologación de cargos
                  </h2>

                  <p className="text-[11px] text-slate-500">
                    Revisa los códigos antes de
                    confirmarlos.
                  </p>
                </div>

                <div className="flex items-center gap-3">
                  <span
                    className={`text-[11px] font-bold px-2.5 py-1 rounded-full border ${
                      pendientes === 0
                        ? "bg-emerald-50 text-emerald-800 border-emerald-200"
                        : "bg-amber-50 text-amber-800 border-amber-200"
                    }`}
                  >
                    {pendientes ===
                    0
                      ? "Todos confirmados"
                      : `${pendientes} por confirmar`}
                  </span>

                  {pendientes >
                    0 && (
                    <button
                      onClick={
                        confirmarTodas
                      }
                      className="text-[11px] font-bold text-white bg-[#003B7A] px-3 py-1.5 rounded-lg cursor-pointer"
                    >
                      Confirmar válidos
                    </button>
                  )}
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-100/70 text-slate-500 uppercase text-[10px] font-extrabold tracking-wider border-b border-slate-200">
                    <tr>
                      <th className="py-3 px-3">
                        Cargo original
                      </th>

                      <th className="py-3 px-3">
                        Cargo español
                      </th>

                      <th className="py-3 px-3">
                        Código CNO
                      </th>

                      <th className="py-3 px-3 text-center">
                        Personas
                      </th>

                      <th className="py-3 px-3">
                        Acción
                      </th>
                    </tr>
                  </thead>

                  <tbody className="divide-y divide-slate-100">
                    {homologacionOrdenada.map(
                      ([
                        clave,
                        h,
                      ]) => (
                        <FilaHomologacion
                          key={
                            clave
                          }
                          clave={
                            clave
                          }
                          h={h}
                          cantidad={
                            cantidadPorCargo[
                              clave
                            ] ||
                            0
                          }
                          onCambiar={
                            aplicarCambioHomologacion
                          }
                        />
                      )
                    )}
                  </tbody>
                </table>
              </div>
            </section>

            {/* =================================================
                RESUMEN POR CARGO
                ================================================= */}

            <section className="bg-white border border-slate-200 rounded-2xl shadow-xs overflow-hidden">
              <div className="p-5 border-b border-slate-100 bg-slate-50/50">
                <h2 className="font-bold text-slate-900 text-base flex items-center gap-2">
                  <Briefcase className="w-4 h-4 text-[#003B7A]" />
                  Resumen por cargo
                </h2>

                <p className="text-[11px] text-slate-500 mt-1">
                  Esta es la agrupación que se utilizará
                  como base para la matriz.
                </p>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-100/70 text-slate-500 uppercase text-[10px] font-extrabold">
                    <tr>
                      <th className="py-3 px-4">
                        Código CNO
                      </th>

                      <th className="py-3 px-4">
                        Cargo
                      </th>

                      {listaResultados.map(
                        (r) => (
                          <th
                            key={
                              r.etiqueta
                            }
                            className="py-3 px-4 text-center"
                          >
                            {r.etiqueta}
                          </th>
                        )
                      )}

                      <th className="py-3 px-4 text-center">
                        Promedio
                      </th>
                    </tr>
                  </thead>

                  <tbody className="divide-y divide-slate-100">
                    {resumenPorCargo.map(
                      (r) => {
                        const valores =
                          listaResultados.map(
                            (
                              mes
                            ) =>
                              r.meses[
                                mes.etiqueta
                              ] ||
                              0
                          );

                        const suma =
                          valores.reduce(
                            (
                              a,
                              b
                            ) =>
                              a +
                              b,
                            0
                          );

                        const promedio =
                          listaResultados.length >
                          0
                            ? suma /
                              listaResultados.length
                            : 0;

                        return (
                          <tr
                            key={
                              r.cargoKey
                            }
                            className="hover:bg-slate-50"
                          >
                            <td className="py-3 px-4 font-mono font-bold text-[#003B7A]">
                              {r.codigo ||
                                "—"}
                            </td>

                            <td className="py-3 px-4 font-semibold">
                              {r.cargoEspanol}
                            </td>

                            {valores.map(
                              (
                                valor,
                                index
                              ) => (
                                <td
                                  key={
                                    index
                                  }
                                  className="py-3 px-4 text-center font-semibold"
                                >
                                  {valor}
                                </td>
                              )
                            )}

                            <td className="py-3 px-4 text-center font-bold text-[#003B7A]">
                              {promedio.toFixed(
                                2
                              )}
                            </td>
                          </tr>
                        );
                      }
                    )}
                  </tbody>
                </table>
              </div>
            </section>

            {/* =================================================
                DETALLE POR MES
                ================================================= */}

            <section className="space-y-4">
              <div className="flex gap-2 flex-wrap border-b border-slate-200/80 pb-2">
                {listaResultados.map(
                  (r) => (
                    <button
                      key={
                        r.etiqueta
                      }
                      onClick={() =>
                        setMesActivo(
                          r.etiqueta
                        )
                      }
                      className={`px-4 py-2 rounded-xl text-xs font-bold border transition-all flex items-center gap-2 cursor-pointer ${
                        mesActivo ===
                        r.etiqueta
                          ? "bg-[#003B7A] text-white border-[#003B7A]"
                          : "bg-white text-slate-600 border-slate-200"
                      }`}
                    >
                      {r.etiqueta}

                      {r.cuadra ? (
                        <CheckCircle2 className="w-3.5 h-3.5" />
                      ) : (
                        <AlertTriangle className="w-3.5 h-3.5" />
                      )}
                    </button>
                  )
                )}
              </div>

              {resultadoActivo && (
                <div className="bg-white border border-slate-200 rounded-2xl shadow-xs overflow-hidden">
                  <div className="p-5 border-b border-slate-100 space-y-2 bg-slate-50/50">
                    <h3 className="font-bold text-slate-900 text-base flex items-center gap-2">
                      {nombreEmpresa
                        ? `${nombreEmpresa} — `
                        : ""}
                      {mesActivo}

                      <span className="text-[11px] font-semibold text-slate-400 bg-slate-100 px-2 py-0.5 rounded-full">
                        Jornada:{" "}
                        {
                          resultadoActivo.jornadaSemanal
                        }
                        h
                      </span>
                    </h3>

                    <div className="flex flex-wrap gap-2">
                      <div className="inline-flex items-center gap-2 text-xs text-slate-700 font-bold bg-slate-100 border border-slate-200 px-3 py-1 rounded-md">
                        Nómina:
                        {
                          resultadoActivo.totalNomina
                        }
                      </div>

                      <div className="inline-flex items-center gap-2 text-xs text-slate-700 font-bold bg-slate-100 border border-slate-200 px-3 py-1 rounded-md">
                        Seguridad Social:
                        {
                          resultadoActivo.totalSegSocial
                        }
                      </div>

                      <div className="inline-flex items-center gap-2 text-xs text-slate-700 font-bold bg-slate-100 border border-slate-200 px-3 py-1 rounded-md">
                        Horas PILA:
                        {
                          resultadoActivo.horasTotalesPila
                        }
                      </div>

                      <div className="inline-flex items-center gap-2 text-xs text-slate-700 font-bold bg-slate-100 border border-slate-200 px-3 py-1 rounded-md">
                        Cuota:
                        {
                          resultadoActivo.cuotaAprendicesRequerida
                        }
                      </div>
                    </div>
                  </div>

                  {/* CRUCE */}

                  {(resultadoActivo.sinMatchSS.length >
                    0 ||
                    resultadoActivo.soloEnSS.length >
                      0) && (
                    <details
                      open={
                        !resultadoActivo.cuadra
                      }
                      className="border-b border-slate-100 px-5 py-3 text-xs"
                    >
                      <summary className="font-bold text-slate-700 cursor-pointer">
                        Diferencias con Seguridad Social
                      </summary>

                      <div className="mt-3 grid grid-cols-1 lg:grid-cols-2 gap-5">
                        <div className="space-y-2">
                          <p className="font-bold text-slate-600">
                            En nómina, pero no en Seguridad Social
                          </p>

                          {resultadoActivo.sinMatchSS.length ===
                          0 ? (
                            <p className="text-slate-400">
                              Ninguno.
                            </p>
                          ) : (
                            <div className="max-h-56 overflow-y-auto rounded-lg border border-slate-200">
                              <table className="w-full text-left">
                                <thead className="bg-slate-100/70 text-[10px] uppercase text-slate-500">
                                  <tr>
                                    <th className="py-2 px-3">
                                      Documento
                                    </th>

                                    <th className="py-2 px-3">
                                      Empleado
                                    </th>

                                    <th className="py-2 px-3">
                                      Retiro
                                    </th>
                                  </tr>
                                </thead>

                                <tbody className="divide-y divide-slate-100">
                                  {resultadoActivo.sinMatchSS.map(
                                    (
                                      e
                                    ) => (
                                      <tr
                                        key={
                                          e.documento
                                        }
                                      >
                                        <td className="py-2 px-3 font-mono">
                                          {
                                            e.documento
                                          }
                                        </td>

                                        <td className="py-2 px-3 font-semibold">
                                          {
                                            e.nombre
                                          }
                                        </td>

                                        <td className="py-2 px-3">
                                          {formatFecha(
                                            e.fechaRetiro
                                          ) ||
                                            "—"}
                                        </td>
                                      </tr>
                                    )
                                  )}
                                </tbody>
                              </table>
                            </div>
                          )}
                        </div>

                        <div className="space-y-2">
                          <p className="font-bold text-slate-600">
                            Recuperados / pendientes de Seguridad Social
                          </p>

                          <div className="max-h-56 overflow-y-auto rounded-lg border border-slate-200">
                            <table className="w-full text-left">
                              <thead className="bg-slate-100/70 text-[10px] uppercase text-slate-500">
                                <tr>
                                  <th className="py-2 px-3">
                                    Documento
                                  </th>

                                  <th className="py-2 px-3">
                                    Cargo
                                  </th>

                                  <th className="py-2 px-3">
                                    Estado
                                  </th>
                                </tr>
                              </thead>

                              <tbody className="divide-y divide-slate-100">
                                {resultadoActivo.soloEnSS.map(
                                  (
                                    e
                                  ) => (
                                    <tr
                                      key={
                                        e.documento
                                      }
                                    >
                                      <td className="py-2 px-3 font-mono">
                                        {
                                          e.documento
                                        }
                                      </td>

                                      <td className="py-2 px-3">
                                        {e.cargo ||
                                          "Sin cargo"}
                                      </td>

                                      <td className="py-2 px-3 text-amber-700">
                                        {e.nota ||
                                          "Revisar"}
                                      </td>
                                    </tr>
                                  )
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      </div>
                    </details>
                  )}

                  {/* DETALLE */}

                  <div className="overflow-x-auto max-h-[480px] overflow-y-auto">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-slate-100/70 text-slate-500 uppercase text-[10px] font-extrabold tracking-wider border-b border-slate-200 sticky top-0">
                        <tr>
                          <th className="py-3 px-4">
                            Documento
                          </th>

                          <th className="py-3 px-4">
                            Empleado
                          </th>

                          <th className="py-3 px-4">
                            Cargo
                          </th>

                          <th className="py-3 px-4">
                            Proporción
                          </th>

                          <th className="py-3 px-4">
                            Horas PILA
                          </th>

                          <th className="py-3 px-4">
                            CNO
                          </th>
                        </tr>
                      </thead>

                      <tbody className="divide-y divide-slate-100">
                        {resultadoActivo.empleados.map(
                          (emp) => {
                            const h =
                              homologacion[
                                emp.cargoKey
                              ];

                            return (
                              <tr
                                key={
                                  emp.documento
                                }
                                className="hover:bg-slate-50"
                              >
                                <td className="py-3 px-4 font-mono text-slate-500">
                                  {
                                    emp.documento
                                  }
                                </td>

                                <td className="py-3 px-4 font-bold">
                                  {
                                    emp.nombre
                                  }
                                </td>

                                <td className="py-3 px-4">
                                  {
                                    emp.cargo
                                  }
                                </td>

                                <td className="py-3 px-4">
                                  {emp.proporcion.toFixed(
                                    2
                                  )}
                                </td>

                                <td className="py-3 px-4 font-semibold">
                                  {redondearHoras(
                                    emp.horasPila
                                  )}

                                  {emp.hayDiferenciaHoras && (
                                    <span
                                      title="Existe diferencia entre proporción por fechas y por horas PILA"
                                      className="ml-2 text-amber-600"
                                    >
                                      ⚠
                                    </span>
                                  )}
                                </td>

                                <td className="py-3 px-4">
                                  {h?.codigo ? (
                                    <span className="font-mono font-bold text-[#003B7A]">
                                      {
                                        h.codigo
                                      }
                                    </span>
                                  ) : (
                                    <span className="text-slate-400">
                                      Sin código
                                    </span>
                                  )}
                                </td>
                              </tr>
                            );
                          }
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* EXPORTAR */}

              <div className="flex justify-end pt-2">
                <button
                  onClick={
                    exportarMatrizCompleta
                  }
                  disabled={
                    exportando
                  }
                  className="bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-200 disabled:text-slate-400 text-white text-xs font-bold px-6 py-3.5 rounded-xl transition shadow-md flex items-center gap-2 cursor-pointer"
                >
                  <Download className="w-4 h-4" />

                  {exportando
                    ? "Generando Excel..."
                    : "Exportar Matriz SENA completa (.xlsx)"}
                </button>
              </div>
            </section>
          </>
        )}
      </main>
    </div>
  );
}