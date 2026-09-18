import { useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
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
} from "lucide-react";
import listadoCnoLocal from "./data/listado-cno.json";

const listadoCno = listadoCnoLocal || [];

// Mapeo manual como respaldo/override
const mapeoDirectoCno = {
  "LOGISTICS SPECIALIST I": "2233",
  "SALESFORCE JUNIOR DEVELOPER": "4131",
  "LOGISTICS OPERATIONS SPECIALIST AFTER HOURS": "2233",
  "CREDIT AND COLLECTIONS ANALYST II": "1121",
  "SALES & PRICING TEAM LEADER": "1121",
  "BRANCH MANAGER": "0421",
  "ADMINISTRATIVE ANALYST": "9227",
  "GERENTE DE CUENTAS JR": "1121",
  "IT SUPPORT SPECIALIST": "4131",
};

// Jornadas semanales legales vigentes en Colombia por rango de fechas (Ley 2101 de 2021).
// Se usa como sugerencia por defecto según el mes/año del periodo, pero siempre es editable
// manualmente por si la empresa aplica una jornada distinta (ej. reducción anticipada).
function jornadaLegalSugerida(anio, mesIndex) {
  const fecha = new Date(anio, mesIndex, 15); // punto medio del mes
  if (fecha < new Date(2023, 6, 15)) return 48; // hasta jul 2023
  if (fecha < new Date(2024, 6, 15)) return 47; // jul 2023 - jul 2024
  if (fecha < new Date(2025, 6, 15)) return 46; // jul 2024 - jul 2025
  if (fecha < new Date(2026, 6, 15)) return 44; // jul 2025 - jul 2026
  return 42; // jul 2026 en adelante
}

function normalizar(texto) {
  return String(texto || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

// Traduce una lista de cargos EN -> ES usando el modelo, a través de nuestro
// servidor proxy local (server.js en localhost:3001), que es quien realmente
// llama a la API de Anthropic con la API key guardada en su .env.
// Esto evita exponer la API key en el navegador y evita el error de CORS que
// ocurre si se llama a api.anthropic.com directamente desde el frontend.
async function traducirCargosConIA(cargosUnicos) {
  const response = await fetch("http://localhost:3001/api/traducir-cargos", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ cargosUnicos }),
  });

  if (!response.ok) {
    throw new Error("Error al traducir cargos");
  }

  const data = await response.json();
  return data.traducciones || {};
}

function obtenerSugerenciasCNO(cargoIngles, traduccionesIA) {
  const cargoLimpio = String(cargoIngles || "").trim().toUpperCase();

  const codigoFijo = mapeoDirectoCno[cargoLimpio];
  if (codigoFijo) {
    const itemEncontrado = listadoCno.find((item) => String(item.codigo) === String(codigoFijo));
    if (itemEncontrado) {
      return [{ ...itemEncontrado, puntos: 100 }];
    }
  }

  const traduccion = traduccionesIA?.[cargoLimpio] || cargoLimpio;
  const palabrasBusqueda = normalizar(traduccion).split(/[\s,]+/).filter(Boolean);

  return listadoCno
    .map((item) => {
      const ocupacionNorm = normalizar(item.ocupacion);
      let puntos = 0;
      for (const palabra of palabrasBusqueda) {
        if (ocupacionNorm.includes(palabra)) puntos++;
      }
      return { ...item, puntos };
    })
    .filter((item) => item.puntos > 0)
    .sort((a, b) => b.puntos - a.puntos)
    .slice(0, 5);
}

// Devuelve el nombre del cargo en español: usa la traducción de la IA si existe,
// y si no, cae de vuelta al texto original en inglés (para no dejar celdas vacías).
function obtenerCargoTraducido(cargoIngles, traduccionesIA) {
  const cargoLimpio = String(cargoIngles || "").trim().toUpperCase();
  const traduccion = traduccionesIA?.[cargoLimpio];
  return traduccion && traduccion.trim() ? traduccion.trim() : cargoIngles;
}

function diasDelMes(anio, mesIndex) {
  return new Date(anio, mesIndex + 1, 0).getDate();
}

// Calcula proporción trabajada dentro del mes según fecha de ingreso/retiro
function calcularProporcionPorFechas(fechaIngreso, fechaRetiro, anio, mesIndex) {
  const totalDias = diasDelMes(anio, mesIndex);
  const inicioMes = new Date(anio, mesIndex, 1);
  const finMes = new Date(anio, mesIndex, totalDias);

  let inicioReal = inicioMes;
  let finReal = finMes;

  if (fechaIngreso) {
    const ingreso = new Date(fechaIngreso);
    if (!isNaN(ingreso) && ingreso > inicioMes) inicioReal = ingreso;
  }
  if (fechaRetiro) {
    const retiro = new Date(fechaRetiro);
    if (!isNaN(retiro) && retiro < finMes) finReal = retiro;
  }

  if (finReal < inicioReal) return 0;

  const diasTrabajados = Math.round((finReal - inicioReal) / (1000 * 60 * 60 * 24)) + 1;
  return Math.min(diasTrabajados / totalDias, 1);
}

// Cuota legal de aprendices SENA: 1 aprendiz por cada 20 trabajadores (redondeo hacia abajo).
function calcularCuotaAprendices(totalTrabajadores) {
  return Math.floor((totalTrabajadores || 0) / 20);
}

async function obtenerImagenBase64(url) {
  const respuesta = await fetch(url);
  const blob = await respuesta.blob();
  return new Promise((resolve, reject) => {
    const lector = new FileReader();
    lector.onloadend = () => resolve(lector.result);
    lector.onerror = reject;
    lector.readAsDataURL(blob);
  });
}

function crearMesesVacios() {
  const anioActual = new Date().getFullYear();
  return Array.from({ length: 6 }, (_, i) => ({
    id: i,
    etiqueta: "",
    anio: anioActual,
    mesIndex: 0,
    jornadaSemanal: jornadaLegalSugerida(anioActual, 0),
    nomina: null,
    segSocial: null,
  }));
}

export default function App() {
  const [nombreEmpresa, setNombreEmpresa] = useState("");
  const [meses, setMeses] = useState(crearMesesVacios);
  const [aprendicesActivos, setAprendicesActivos] = useState(0);
  const [resultadosPorMes, setResultadosPorMes] = useState({});
  const [cnoSeleccionados, setCnoSeleccionados] = useState({});
  const [cargando, setCargando] = useState(false);
  const [traduciendo, setTraduciendo] = useState(false);
  const [error, setError] = useState(null);
  const [mostrarInstrucciones, setMostrarInstrucciones] = useState(true);
  const [mesActivo, setMesActivo] = useState(null);

  const actualizarMes = (id, campo, valor) => {
    setMeses((prev) =>
      prev.map((m) => {
        if (m.id !== id) return m;
        const actualizado = { ...m, [campo]: valor };
        // Si cambia el año o el mes, sugerimos la jornada legal vigente para esa fecha,
        // salvo que el usuario ya la haya ajustado manualmente en esta misma edición.
        if (campo === "anio" || campo === "mesIndex") {
          actualizado.jornadaSemanal = jornadaLegalSugerida(actualizado.anio, actualizado.mesIndex);
        }
        return actualizado;
      })
    );
  };

  const leerNomina = (buffer) => {
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const filas = XLSX.utils.sheet_to_json(sheet, { range: 1 });

    return filas
      .filter((fila) => fila["DOCUMENTO IDENTIDAD"])
      .map((fila) => ({
        documento: String(fila["DOCUMENTO IDENTIDAD"]).trim(),
        nombre: fila["NOMBRE COMPLETO"],
        cargo: String(fila["DESC. OFICIO"] || "").trim(),
        fechaIngreso: fila["FECHA INGRESO"] || fila["FECHA ANTIGUEDAD"] || null,
        fechaRetiro: fila["FECHA RETIRO"] || null,
      }));
  };

  // Nombres de columna habituales para el número de documento en planillas de
  // Seguridad Social de distintos operadores (PILA, aportes en línea, etc.)
  const ALIAS_COLUMNA_DOCUMENTO = ["no id", "numero de identificacion", "numero identificacion", "documento", "cedula", "identificacion", "n° id", "nro id"];
  const ALIAS_ENCABEZADO_ANCLA = ["no id"];

  const leerSeguridadSocial = (buffer) => {
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const filas = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

    // 1) Ubicar la fila de encabezado buscando cualquiera de los alias conocidos,
    //    en vez de asumir siempre el texto exacto "No id".
    const indiceEncabezado = filas.findIndex(
      (fila) =>
        Array.isArray(fila) &&
        fila.some((celda) => {
          const norm = normalizar(celda);
          return ALIAS_ENCABEZADO_ANCLA.some((alias) => norm.includes(alias)) ||
            ALIAS_COLUMNA_DOCUMENTO.some((alias) => norm === alias);
        })
    );

    if (indiceEncabezado === -1) {
      throw new Error(
        "No se encontró una columna de identificación reconocible (ej. 'No id', 'Documento', 'Cédula') en la planilla de Seguridad Social."
      );
    }

    const filaEncabezado = filas[indiceEncabezado];

    // 2) Determinar dinámicamente en qué columna está el documento, en lugar de
    //    asumir siempre el índice fijo 7 (esto rompía con formatos distintos).
    let indiceColumnaDocumento = filaEncabezado.findIndex((celda) => {
      const norm = normalizar(celda);
      return ALIAS_COLUMNA_DOCUMENTO.some((alias) => norm.includes(alias));
    });

    if (indiceColumnaDocumento === -1) {
      throw new Error(
        "Se encontró el encabezado de la planilla, pero no se pudo identificar la columna del número de documento."
      );
    }

    const filasDatos = filas.slice(indiceEncabezado + 1);
    const documentos = new Set();

    for (const fila of filasDatos) {
      if (!fila) continue;
      const documento = fila[indiceColumnaDocumento];
      if (!documento) continue;
      documentos.add(String(documento).trim());
    }

    return documentos;
  };

  const traducirCargosDeTodosLosMeses = async () => {
    setTraduciendo(true);
    setError(null);
    try {
      const cargosUnicos = new Set();
      for (const mes of meses) {
        if (!mes.nomina) continue;
        const buffer = await mes.nomina.arrayBuffer();
        const empleados = leerNomina(buffer);
        empleados.forEach((e) => e.cargo && cargosUnicos.add(e.cargo.toUpperCase()));
      }
      if (cargosUnicos.size === 0) {
        setError("Carga al menos una nómina antes de traducir los cargos.");
        return null;
      }
      const traducciones = await traducirCargosConIA(Array.from(cargosUnicos));
      setError(null);
      return traducciones;
    } catch (err) {
      setError("Error al traducir cargos con IA: " + (err.message || err));
      return null;
    } finally {
      setTraduciendo(false);
    }
  };

  const procesarTodosLosMeses = async () => {
    const mesesConArchivos = meses.filter((m) => m.nomina && m.segSocial && m.etiqueta);
    if (mesesConArchivos.length === 0) {
      setError("Completa al menos un mes con nombre de mes, nómina y Seguridad Social.");
      return;
    }

    setCargando(true);
    setError(null);

    try {
      const traduccionesIA = await traducirCargosDeTodosLosMeses();

      const nuevosResultados = {};
      const nuevosCno = {};

      for (const mes of mesesConArchivos) {
        const bytesNomina = await mes.nomina.arrayBuffer();
        const bytesSegSocial = await mes.segSocial.arrayBuffer();

        const empleadosNomina = leerNomina(bytesNomina);
        const documentosSegSocial = leerSeguridadSocial(bytesSegSocial);

        const empleadosProcesados = empleadosNomina.map((emp) => {
          const proporcion = calcularProporcionPorFechas(
            emp.fechaIngreso,
            emp.fechaRetiro,
            mes.anio,
            mes.mesIndex
          );
          const sugerencias = obtenerSugerenciasCNO(emp.cargo, traduccionesIA);
          const cargoTraducido = obtenerCargoTraducido(emp.cargo, traduccionesIA);
          return {
            ...emp,
            cargoTraducido,
            proporcion: Number(proporcion.toFixed(2)),
            trabajoCompleto: proporcion >= 0.999,
            enSegSocial: documentosSegSocial.has(emp.documento),
            sugerenciasCno: sugerencias,
            cnoPorDefecto: sugerencias[0]?.codigo || "",
          };
        });

        const totalNomina = empleadosProcesados.length;
        const totalSegSocial = documentosSegSocial.size;
        const noEncontradosEnSegSocial = empleadosProcesados
          .filter((e) => !e.enSegSocial)
          .map((e) => e.documento);

        nuevosCno[mes.etiqueta] = {};
        empleadosProcesados.forEach((emp) => {
          nuevosCno[mes.etiqueta][emp.documento] = emp.cnoPorDefecto;
        });

        nuevosResultados[mes.etiqueta] = {
          anio: mes.anio,
          mesIndex: mes.mesIndex,
          jornadaSemanal: mes.jornadaSemanal, // se congela el valor usado en el procesamiento
          totalNomina,
          totalSegSocial,
          cuotaAprendicesRequerida: calcularCuotaAprendices(totalNomina),
          cuadra: totalNomina === totalSegSocial,
          noEncontradosEnSegSocial,
          empleados: empleadosProcesados,
        };
      }

      setResultadosPorMes(nuevosResultados);
      setCnoSeleccionados(nuevosCno);
      setMesActivo(mesesConArchivos[0].etiqueta);
    } catch (err) {
      setError(err.message || "Error al procesar los archivos");
    } finally {
      setCargando(false);
    }
  };

  // Agrupa por cargo. La clave de agrupación usa el cargo ORIGINAL (inglés) para no
  // fusionar cargos distintos que por azar tradujeran parecido, pero el nombre que se
  // muestra y se exporta siempre es la versión en español.
  const agruparPorCargo = (empleados, cnoDelMes) => {
    const grupos = {};
    empleados.forEach((emp) => {
      const codigo = cnoDelMes[emp.documento] || emp.cnoPorDefecto || "";
      const esCalificado = Boolean(codigo);
      const clave = esCalificado ? codigo : `SIN_CNO::${emp.cargo}`;
      const nombreCargoEs = emp.cargoTraducido || emp.cargo;

      if (!grupos[clave]) {
        grupos[clave] = {
          codigo: esCalificado ? codigo : "",
          nombreCargo: nombreCargoEs,
          calificado: esCalificado,
          completos: 0,
          parciales: [],
        };
      }
      if (emp.trabajoCompleto) {
        grupos[clave].completos += 1;
      } else {
        grupos[clave].parciales.push(emp.proporcion);
      }
    });
    return Object.values(grupos);
  };

  // logoId: id de imagen retornado por wb.addImage(), o null/undefined si no aplica.
  // Cuando se pasa, se inserta en la esquina superior derecha de la hoja.
  const construirHojaMes = (wb, etiquetaMes, resultado, cnoDelMes, incluirEncabezadoEmpresa, logoId) => {
    const GRIS_ENCABEZADO = "FFD9D9D9";
    const bordeNegro = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    const jornadaDelMes = resultado.jornadaSemanal || 40;
    const nombreHoja = `MATRIZ ${etiquetaMes}`.slice(0, 31);
    const ws = wb.addWorksheet(nombreHoja, { views: [{ showGridLines: true }] });
    ws.columns = [{ width: 32 }, { width: 35 }, { width: 35 }, { width: 32 }];

    // Insertar el logo del SENA (si se cargó correctamente) en la esquina superior
    // derecha, sin invadir las columnas de datos. Se ubica "flotando" sobre las
    // celdas usando coordenadas en EMU relativas para no desplazar filas/columnas.
    if (logoId !== null && logoId !== undefined) {
      ws.addImage(logoId, {
        tl: { col: 4.1, row: 0.1 },
        ext: { width: 140, height: 50 },
        editAs: "oneCell",
      });
    }

    // Encabezado de empresa + periodo + jornada aplicada en TODAS las hojas
    // (antes solo aparecía en la primera hoja del libro).
    const tituloEmpresa = nombreEmpresa ? `${nombreEmpresa} — ` : "";
    const fTitulo = ws.addRow([`${tituloEmpresa}Matriz SENA - ${etiquetaMes}`]);
    fTitulo.getCell(1).font = { bold: true, size: 12 };
    ws.mergeCells(fTitulo.number, 1, fTitulo.number, 4);

    const fSubtitulo = ws.addRow([`Jornada laboral semanal aplicada este periodo: ${jornadaDelMes} horas`]);
    fSubtitulo.getCell(1).font = { italic: true, size: 9, color: { argb: "FF666666" } };
    ws.mergeCells(fSubtitulo.number, 1, fSubtitulo.number, 4);
    ws.addRow([]);

    const fAlerta = ws.addRow([
      resultado.cuadra
        ? `OK: Nómina (${resultado.totalNomina}) coincide con Seguridad Social (${resultado.totalSegSocial})`
        : `⚠ ALERTA: Nómina (${resultado.totalNomina}) NO coincide con Seguridad Social (${resultado.totalSegSocial})`,
    ]);
    ws.mergeCells(fAlerta.number, 1, fAlerta.number, 4);
    fAlerta.getCell(1).font = { bold: true, color: { argb: resultado.cuadra ? "FF006100" : "FF9C0006" } };
    fAlerta.getCell(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: resultado.cuadra ? "FFC6EFCE" : "FFFFC7CE" },
    };

    const fCuota = ws.addRow([
      `Cuota de aprendices requerida este periodo (1 por cada 20 trabajadores): ${resultado.cuotaAprendicesRequerida}`,
    ]);
    ws.mergeCells(fCuota.number, 1, fCuota.number, 4);
    fCuota.getCell(1).font = { italic: true, size: 9, color: { argb: "FF444444" } };
    ws.addRow([]);

    const fMatriz1Tit = ws.addRow(["Matriz 1: Oficios Calificados"]);
    fMatriz1Tit.getCell(1).font = { bold: true, size: 10 };

    const fEnc1 = ws.addRow([
      "CÓDIGO DEL OFICIO SEGÚN LISTADO DE OFICIOS Y OCUPACIONES",
      "NÚMERO DE TRABAJADORES",
      "JORNADA LABORAL SEMANAL POR TRABAJADOR",
      "TOTAL JORNADA LABORAL SEMANAL",
    ]);
    fEnc1.height = 28;
    fEnc1.eachCell((c) => {
      c.font = { bold: true, size: 8 };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRIS_ENCABEZADO } };
      c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      c.border = bordeNegro;
    });

    const grupos = agruparPorCargo(resultado.empleados, cnoDelMes);
    const calificados = grupos.filter((g) => g.calificado);
    const noCalificados = grupos.filter((g) => !g.calificado);

    const inicioM1 = ws.lastRow.number + 1;
    calificados.forEach((g) => {
      if (g.completos > 0) {
        const f = ws.addRow([g.codigo, g.completos, jornadaDelMes, g.completos * jornadaDelMes]);
        f.eachCell((c) => {
          c.border = bordeNegro;
          c.alignment = { horizontal: "center", vertical: "middle" };
          c.font = { size: 9 };
        });
      }
      g.parciales.forEach((prop) => {
        const jornadaProp = Number((jornadaDelMes * prop).toFixed(1));
        const f = ws.addRow([g.codigo, 1, jornadaProp, jornadaProp]);
        f.eachCell((c) => {
          c.border = bordeNegro;
          c.alignment = { horizontal: "center", vertical: "middle" };
          c.font = { size: 9, italic: true };
        });
      });
    });
    const finM1 = ws.lastRow.number;

    const fTotal1 = ws.addRow([
      "TOTAL",
      { formula: `SUM(B${inicioM1}:B${finM1})` },
      "",
      { formula: `SUM(D${inicioM1}:D${finM1})` },
    ]);
    fTotal1.eachCell((c) => {
      c.font = { bold: true, size: 9 };
      c.border = bordeNegro;
      c.alignment = { horizontal: "center", vertical: "middle" };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRIS_ENCABEZADO } };
    });
    ws.addRow([]);

    const fMatriz2Tit = ws.addRow(["Matriz 2: Oficios no Calificados"]);
    fMatriz2Tit.getCell(1).font = { bold: true, size: 10 };

    const fEnc2 = ws.addRow([
      "NOMBRE DEL CARGO (ESPAÑOL)",
      "NÚMERO DE TRABAJADORES",
      "JORNADA LABORAL SEMANAL POR TRABAJADOR",
      "TOTAL JORNADA LABORAL SEMANAL",
    ]);
    fEnc2.height = 28;
    fEnc2.eachCell((c) => {
      c.font = { bold: true, size: 8 };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRIS_ENCABEZADO } };
      c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      c.border = bordeNegro;
    });

    const inicioM2 = ws.lastRow.number + 1;
    noCalificados.forEach((g) => {
      if (g.completos > 0) {
        const f = ws.addRow([g.nombreCargo, g.completos, jornadaDelMes, g.completos * jornadaDelMes]);
        f.eachCell((c) => {
          c.border = bordeNegro;
          c.alignment = { horizontal: "center", vertical: "middle" };
          c.font = { size: 9 };
        });
      }
      g.parciales.forEach((prop) => {
        const jornadaProp = Number((jornadaDelMes * prop).toFixed(1));
        const f = ws.addRow([g.nombreCargo, 1, jornadaProp, jornadaProp]);
        f.eachCell((c) => {
          c.border = bordeNegro;
          c.alignment = { horizontal: "center", vertical: "middle" };
          c.font = { size: 9, italic: true };
        });
      });
    });
    if (noCalificados.length === 0) {
      const f = ws.addRow(["(ninguno)", 0, 0, 0]);
      f.eachCell((c) => (c.border = bordeNegro));
    }
    const finM2 = ws.lastRow.number;

    const fTotal2 = ws.addRow([
      "TOTAL",
      { formula: `SUM(B${inicioM2}:B${finM2})` },
      "",
      { formula: `SUM(D${inicioM2}:D${finM2})` },
    ]);
    fTotal2.eachCell((c) => {
      c.font = { bold: true, size: 9 };
      c.border = bordeNegro;
      c.alignment = { horizontal: "center", vertical: "middle" };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRIS_ENCABEZADO } };
    });
  };

  const construirHojaPlantillaPromedio = (wb) => {
    const ws = wb.addWorksheet("PLANTILLA PROMEDIO");
    ws.columns = [{ width: 40 }, { width: 20 }];

    if (nombreEmpresa) {
      const fTitulo = ws.addRow([`${nombreEmpresa} — Plantilla Promedio`]);
      fTitulo.getCell(1).font = { bold: true, size: 12 };
      ws.mergeCells(fTitulo.number, 1, fTitulo.number, 2);
      ws.addRow([]);
    }

    const fEnc = ws.addRow(["CARGO / CÓDIGO CNO", "PROMEDIO TRABAJADORES (6 MESES)"]);
    fEnc.eachCell((c) => {
      c.font = { bold: true };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
    });

    const acumulado = {};
    const etiquetas = Object.keys(resultadosPorMes);
    etiquetas.forEach((etiqueta) => {
      const resultado = resultadosPorMes[etiqueta];
      const cnoDelMes = cnoSeleccionados[etiqueta] || {};
      const grupos = agruparPorCargo(resultado.empleados, cnoDelMes);
      grupos.forEach((g) => {
        const clave = g.calificado ? g.codigo : g.nombreCargo;
        const totalMes = g.completos + g.parciales.reduce((a, b) => a + b, 0);
        acumulado[clave] = (acumulado[clave] || 0) + totalMes;
      });
    });

    Object.entries(acumulado).forEach(([clave, suma]) => {
      ws.addRow([clave, Number((suma / (etiquetas.length || 1)).toFixed(2))]);
    });
  };

  const exportarMatrizCompleta = async () => {
    if (Object.keys(resultadosPorMes).length === 0) return;

    const wb = new ExcelJS.Workbook();

    // Cargamos el logo UNA sola vez al workbook y guardamos su id.
    // Ese id es lo que se debe pasar a ws.addImage() en cada hoja donde
    // se quiera insertar — cargar la imagen no la inserta por sí sola.
    let logoId = null;
    try {
      const base64Logo = await obtenerImagenBase64("/sena-logo.png");
      logoId = wb.addImage({ base64: base64Logo, extension: "png" });
    } catch (e) {
      console.warn("No se pudo cargar el logo del SENA:", e);
    }

    Object.entries(resultadosPorMes).forEach(([etiqueta, resultado], idx) => {
      // Por defecto el logo se inserta solo en la primera hoja (idx === 0).
      // Si prefieres que salga en TODAS las hojas mensuales, cambia
      // `idx === 0 ? logoId : null` por simplemente `logoId`.
      construirHojaMes(wb, etiqueta, resultado, cnoSeleccionados[etiqueta] || {}, idx === 0, idx === 0 ? logoId : null);
    });

    construirHojaPlantillaPromedio(wb);

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const prefijoEmpresa = nombreEmpresa
      ? nombreEmpresa.trim().toUpperCase().replace(/[^A-Z0-9]+/g, "_") + "_"
      : "";
    a.download = `${prefijoEmpresa}MATRIZ_DETERMINACION_CUOTA_SENA.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const resultadoActivo = mesActivo ? resultadosPorMes[mesActivo] : null;

  // Para comparar la cuota de aprendices "actual" usamos el mes más reciente ya procesado.
  const etiquetasProcesadas = Object.keys(resultadosPorMes);
  const resultadoMasReciente =
    etiquetasProcesadas.length > 0 ? resultadosPorMes[etiquetasProcesadas[etiquetasProcesadas.length - 1]] : null;
  const cuotaVigente = resultadoMasReciente ? resultadoMasReciente.cuotaAprendicesRequerida : null;
  const aprendicesNum = Number(aprendicesActivos) || 0;

  return (
    <div className="min-h-screen bg-slate-50/50 font-sans text-slate-800 antialiased selection:bg-[#003B7A]/10">
      {/* Header Corporativo */}
      <header className="bg-white border-b border-slate-200/80 sticky top-0 z-50 backdrop-blur-md bg-white/90">
        <div className="max-w-6xl mx-auto px-6 py-3.5 flex items-center justify-between">
          <img src="/logo.jpeg" alt="Solutions & Payroll Logo" className="h-9 w-auto object-contain" />
          <span className="text-xs font-semibold px-3 py-1 rounded-full bg-slate-100 text-slate-600 border border-slate-200/60 flex items-center gap-1.5">
            <Building2 className="w-3.5 h-3.5 text-[#003B7A]" /> Módulo SENA
          </span>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-10 space-y-8">
        {/* Banner principal */}
        <section className="text-center space-y-3">
          <div className="inline-flex items-center gap-2 px-3 .5 py-1 rounded-full bg-blue-50 border border-blue-100 text-[#003B7A] text-xs font-semibold">
            <Sparkles className="w-3.5 h-3.5" /> Normalización inteligente con IA
          </div>
          <h1 className="text-3xl sm:text-4xl font-extrabold text-slate-900 tracking-tight flex items-center justify-center gap-3">
            Gestión de Planta y Matriz SENA
          </h1>
          <p className="text-sm text-slate-500 max-w-2xl mx-auto leading-relaxed">
            Carga los datos de hasta 6 meses, homologa cargos con la Clasificación Nacional de Ocupaciones y genera la matriz requerida para la determinación de cuota.
          </p>
        </section>

        {/* Acordeón de Instrucciones */}
        <div className="bg-blue-50/60 border border-blue-100 rounded-2xl p-5 transition-all">
          <button
            onClick={() => setMostrarInstrucciones(!mostrarInstrucciones)}
            className="w-full flex items-center justify-between text-left font-bold text-[#003B7A] text-sm"
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
                  <span className="font-bold text-[#003B7A] block">1. Carga de Archivos</span>
                  Indica la empresa, el periodo (ej. "Julio 2026") y sube los archivos de Nómina y Seguridad Social de cada mes.
                </div>
                <div className="p-3 bg-white/80 rounded-xl border border-blue-50/80 shadow-2xs space-y-1">
                  <span className="font-bold text-[#003B7A] block">2. Configuración</span>
                  Revisa la jornada laboral semanal sugerida por periodo (cambia con la ley) e indica los aprendices activos.
                </div>
                <div className="p-3 bg-white/80 rounded-xl border border-blue-50/80 shadow-2xs space-y-1">
                  <span className="font-bold text-[#003B7A] block">3. Procesamiento IA</span>
                  El sistema traduce y homologa automáticamente la ocupación CNO y valida novedades de ingreso/retiro.
                </div>
                <div className="p-3 bg-white/80 rounded-xl border border-blue-50/80 shadow-2xs space-y-1">
                  <span className="font-bold text-[#003B7A] block">4. Exportación</span>
                  Audita las divergencias y genera el libro en Excel listo con el formato de la matriz.
                </div>
              </div>
              <div className="p-3 bg-amber-50 rounded-xl border border-amber-200/80 text-xs text-amber-800 font-medium flex items-start gap-2">
                <AlertTriangle className="w-4 h-4 shrink-0 mt-0.5 text-amber-600" />
                Verifica periódicamente que <code className="font-mono">listado-cno.json</code> corresponda a la versión vigente publicada por el SENA: el listado de homologación cambia con cada nueva resolución.
              </div>
            </div>
          )}
        </div>

        {/* Datos de la empresa */}
        <section className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-2">
          <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
            <Building2 className="w-4 h-4 text-slate-500" /> Nombre de la empresa
          </label>
          <input
            type="text"
            placeholder="Ej. CIPY S.A.S."
            value={nombreEmpresa}
            onChange={(e) => setNombreEmpresa(e.target.value)}
            className="w-full md:w-1/2 bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs font-semibold text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
          />
          <p className="text-[11px] text-slate-400">Aparecerá en cada hoja del Excel exportado y en el nombre del archivo.</p>
        </section>

        {/* Sección de Carga de Meses */}
        <section className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-6">
          <div className="flex items-center justify-between border-b border-slate-100 pb-4">
            <h2 className="text-sm font-bold text-slate-900 flex items-center gap-2 uppercase tracking-wider">
              <Calendar className="w-4 h-4 text-[#003B7A]" />
              Meses a reportar
            </h2>
            <span className="text-xs text-slate-400 font-medium">Máximo 6 periodos</span>
          </div>

          <div className="space-y-3">
            {meses.map((mes) => (
              <div
                key={mes.id}
                className={`p-4 rounded-xl border transition-all ${
                  mes.nomina && mes.segSocial
                    ? "bg-slate-50/50 border-slate-200"
                    : "bg-white border-slate-200/70 hover:border-slate-300"
                }`}
              >
                <div className="grid grid-cols-1 md:grid-cols-12 gap-4 items-center">
                  {/* Etiqueta y fecha */}
                  <div className="md:col-span-4 space-y-2">
                    <label className="text-[11px] uppercase font-extrabold text-slate-500 tracking-wider">Etiqueta del mes</label>
                    <input
                      type="text"
                      placeholder="Ej. Julio 2026"
                      value={mes.etiqueta}
                      onChange={(e) => actualizarMes(mes.id, "etiqueta", e.target.value)}
                      className="w-full bg-white border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-semibold text-slate-800 placeholder:text-slate-300 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                    />
                    <div className="flex gap-2">
                      <input
                        type="number"
                        value={mes.anio}
                        onChange={(e) => actualizarMes(mes.id, "anio", Number(e.target.value))}
                        className="w-1/2 bg-white border border-slate-200 rounded-lg px-2 py-1 text-xs text-slate-600 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                      />
                      <select
                        value={mes.mesIndex}
                        onChange={(e) => actualizarMes(mes.id, "mesIndex", Number(e.target.value))}
                        className="w-1/2 bg-white border border-slate-200 rounded-lg px-2 py-1 text-xs text-slate-600 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                      >
                        {["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"].map((m, i) => (
                          <option key={i} value={i}>
                            {m}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>

                  {/* Carga Archivo Nómina */}
                  <div className="md:col-span-3 space-y-1.5">
                    <label className="text-[11px] font-bold text-slate-600 flex items-center gap-1">
                      <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" /> Nómina (.xlsx)
                    </label>
                    <label
                      className={`flex items-center justify-between px-3 py-2 rounded-lg border text-xs cursor-pointer transition ${
                        mes.nomina
                          ? "bg-emerald-50/50 border-emerald-200 text-emerald-800"
                          : "bg-slate-50 border-slate-200 text-slate-400 hover:bg-slate-100/80"
                      }`}
                    >
                      <span className="truncate max-w-[120px]">{mes.nomina ? mes.nomina.name : "Seleccionar..."}</span>
                      <Upload className="w-3.5 h-3.5 shrink-0 opacity-60" />
                      <input
                        type="file"
                        accept=".xlsx"
                        onChange={(e) => actualizarMes(mes.id, "nomina", e.target.files?.[0] || null)}
                        className="hidden"
                      />
                    </label>
                  </div>

                  {/* Carga Archivo Seg Social */}
                  <div className="md:col-span-3 space-y-1.5">
                    <label className="text-[11px] font-bold text-slate-600 flex items-center gap-1">
                      <FileText className="w-3.5 h-3.5 text-blue-600" /> Seg. Social (.xlsx)
                    </label>
                    <label
                      className={`flex items-center justify-between px-3 py-2 rounded-lg border text-xs cursor-pointer transition ${
                        mes.segSocial
                          ? "bg-blue-50/50 border-blue-200 text-blue-800"
                          : "bg-slate-50 border-slate-200 text-slate-400 hover:bg-slate-100/80"
                      }`}
                    >
                      <span className="truncate max-w-[120px]">{mes.segSocial ? mes.segSocial.name : "Seleccionar..."}</span>
                      <Upload className="w-3.5 h-3.5 shrink-0 opacity-60" />
                      <input
                        type="file"
                        accept=".xlsx"
                        onChange={(e) => actualizarMes(mes.id, "segSocial", e.target.files?.[0] || null)}
                        className="hidden"
                      />
                    </label>
                  </div>

                  {/* Jornada semanal por periodo */}
                  <div className="md:col-span-1.5 space-y-1.5">
                    <label className="text-[11px] font-bold text-slate-600 flex items-center gap-1">
                      <Clock className="w-3.5 h-3.5 text-slate-500" /> Jornada
                    </label>
                    <select
                      value={mes.jornadaSemanal}
                      onChange={(e) => actualizarMes(mes.id, "jornadaSemanal", Number(e.target.value))}
                      className="w-full bg-slate-50 border border-slate-200 rounded-lg px-2 py-2 text-xs font-semibold text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                    >
                      {[40, 42, 44, 46, 47, 48].map((h) => (
                        <option key={h} value={h}>
                          {h}h
                        </option>
                      ))}
                    </select>
                  </div>

                  {/* Indicator Status */}
                  <div className="md:col-span-0.5 flex justify-end items-center pt-2 md:pt-0">
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
              <p className="text-[11px] text-slate-400">
                Se compara contra la cuota calculada (1 aprendiz por cada 20 trabajadores) del último mes procesado.
              </p>
            </div>
            {cuotaVigente !== null && (
              <div
                className={`rounded-xl border p-4 flex items-center gap-3 text-xs font-bold ${
                  aprendicesNum >= cuotaVigente
                    ? "bg-emerald-50 border-emerald-200 text-emerald-800"
                    : "bg-amber-50 border-amber-200 text-amber-800"
                }`}
              >
                {aprendicesNum >= cuotaVigente ? (
                  <CheckCircle2 className="w-5 h-5 shrink-0" />
                ) : (
                  <AlertTriangle className="w-5 h-5 shrink-0" />
                )}
                <span>
                  Cuota requerida ({etiquetasProcesadas[etiquetasProcesadas.length - 1]}): {cuotaVigente} aprendiz(es). Actualmente
                  reportas {aprendicesNum}.
                  {aprendicesNum < cuotaVigente ? " Faltan por cubrir." : " Cuota cubierta."}
                </span>
              </div>
            )}
          </div>

          <button
            onClick={procesarTodosLosMeses}
            disabled={cargando || traduciendo}
            className="w-full bg-[#003B7A] hover:bg-[#002B5B] disabled:bg-slate-200 disabled:text-slate-400 text-white font-bold py-3.5 rounded-xl transition-all shadow-md hover:shadow-lg disabled:shadow-none text-sm flex items-center justify-center gap-2 cursor-pointer"
          >
            {traduciendo ? (
              <>
                <Sparkles className="w-4 h-4 animate-spin text-amber-300" /> Traduciendo cargos con IA...
              </>
            ) : cargando ? (
              <>
                <Upload className="w-4 h-4 animate-bounce" /> Procesando archivos...
              </>
            ) : (
              <>
                <Sparkles className="w-4 h-4 text-blue-200" /> Procesar todos los meses
              </>
            )}
          </button>
        </section>

        {/* Notificación de Error */}
        {error && (
          <div className="p-4 bg-rose-50 border border-rose-200 text-rose-800 rounded-xl text-xs font-semibold flex items-center gap-3 shadow-2xs">
            <AlertTriangle className="w-5 h-5 flex-shrink-0 text-rose-600" />
            <p>{error}</p>
          </div>
        )}

        {/* Resultados */}
        {Object.keys(resultadosPorMes).length > 0 && (
          <section className="space-y-4">
            {/* Nav tabs */}
            <div className="flex gap-2 flex-wrap border-b border-slate-200/80 pb-2">
              {Object.keys(resultadosPorMes).map((etiqueta) => (
                <button
                  key={etiqueta}
                  onClick={() => setMesActivo(etiqueta)}
                  className={`px-4 py-2 rounded-xl text-xs font-bold border transition-all flex items-center gap-2 cursor-pointer ${
                    mesActivo === etiqueta
                      ? "bg-[#003B7A] text-white border-[#003B7A] shadow-xs"
                      : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50"
                  }`}
                >
                  {etiqueta}
                  {resultadosPorMes[etiqueta].cuadra ? (
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                  ) : (
                    <AlertTriangle className="w-3.5 h-3.5 text-amber-400" />
                  )}
                </button>
              ))}
            </div>

            {/* Vista detalle del mes activo */}
            {resultadoActivo && (
              <div className="bg-white border border-slate-200 rounded-2xl shadow-xs overflow-hidden">
                <div className="p-5 border-b border-slate-100 space-y-1.5 bg-slate-50/50">
                  <h3 className="font-bold text-slate-900 text-base flex items-center gap-2">
                    {nombreEmpresa ? `${nombreEmpresa} — ` : ""}
                    {mesActivo}
                    <span className="text-[11px] font-semibold text-slate-400 bg-slate-100 px-2 py-0.5 rounded-full">
                      Jornada: {resultadoActivo.jornadaSemanal}h
                    </span>
                  </h3>
                  {resultadoActivo.cuadra ? (
                    <div className="inline-flex items-center gap-2 text-xs text-emerald-800 font-bold bg-emerald-50 border border-emerald-200/60 px-3 py-1 rounded-md">
                      <CheckCircle2 className="w-4 h-4 text-emerald-600" /> Nómina ({resultadoActivo.totalNomina}) coincide exactamente con Seguridad Social ({resultadoActivo.totalSegSocial})
                    </div>
                  ) : (
                    <div className="inline-flex items-center gap-2 text-xs text-rose-800 font-bold bg-rose-50 border border-rose-200/60 px-3 py-1 rounded-md">
                      <AlertTriangle className="w-4 h-4 text-rose-600" /> Discrepancia: Nómina ({resultadoActivo.totalNomina}) vs Seguridad Social ({resultadoActivo.totalSegSocial}). Sin match: {resultadoActivo.noEncontradosEnSegSocial.join(", ") || "ninguno"}
                    </div>
                  )}
                  <div className="inline-flex items-center gap-2 text-xs text-slate-600 font-semibold bg-slate-100 border border-slate-200/60 px-3 py-1 rounded-md ml-0 mt-1">
                    <Briefcase className="w-4 h-4 text-slate-500" /> Cuota de aprendices requerida este mes: {resultadoActivo.cuotaAprendicesRequerida}
                  </div>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-100/70 text-slate-500 uppercase text-[10px] font-extrabold tracking-wider border-b border-slate-200">
                      <tr>
                        <th className="py-3 px-4">Documento</th>
                        <th className="py-3 px-4">Empleado</th>
                        <th className="py-3 px-4">Cargo (original)</th>
                        <th className="py-3 px-4">Cargo (español)</th>
                        <th className="py-3 px-4">Proporción mes</th>
                        <th className="py-3 px-4">CNO Asignado</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 text-slate-700">
                      {resultadoActivo.empleados.map((emp) => (
                        <tr key={emp.documento} className="hover:bg-slate-50/80 transition-colors">
                          <td className="py-3 px-4 font-mono text-slate-500">{emp.documento}</td>
                          <td className="py-3 px-4 font-bold text-slate-900">{emp.nombre}</td>
                          <td className="py-3 px-4 text-slate-400">{emp.cargo}</td>
                          <td className="py-3 px-4 text-slate-700 font-medium">{emp.cargoTraducido}</td>
                          <td className="py-3 px-4 font-semibold">
                            {emp.trabajoCompleto ? (
                              <span className="inline-flex items-center gap-1 text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-full border border-emerald-200/50">
                                <CheckCircle2 className="w-3 h-3" /> Completo
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 text-amber-700 bg-amber-50 px-2 py-0.5 rounded-full border border-amber-200/50">
                                {emp.proporcion}
                              </span>
                            )}
                          </td>
                          <td className="py-3 px-4">
                            <select
                              value={cnoSeleccionados[mesActivo]?.[emp.documento] || ""}
                              onChange={(e) =>
                                setCnoSeleccionados({
                                  ...cnoSeleccionados,
                                  [mesActivo]: { ...cnoSeleccionados[mesActivo], [emp.documento]: e.target.value },
                                })
                              }
                              className="w-full bg-slate-50 border border-slate-200 rounded-lg p-1.5 text-xs font-medium text-slate-700 focus:outline-hidden focus:ring-2 focus:ring-[#003B7A]/20 focus:border-[#003B7A]"
                            >
                              {emp.sugerenciasCno.map((sug) => (
                                <option key={sug.codigo} value={sug.codigo}>
                                  [{sug.codigo}] {sug.ocupacion}
                                </option>
                              ))}
                              <option value="">(Sin código — No calificado)</option>
                            </select>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            <div className="flex justify-end pt-2">
              <button
                onClick={exportarMatrizCompleta}
                className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold px-6 py-3.5 rounded-xl transition shadow-md hover:shadow-lg flex items-center gap-2 cursor-pointer"
              >
                <Download className="w-4 h-4" /> Exportar Matriz SENA completa (.xlsx)
              </button>
            </div>
          </section>
        )}
      </main>
    </div>
  );
}