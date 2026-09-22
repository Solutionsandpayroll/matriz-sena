// =====================================================================
// Lógica pura de la Matriz SENA (sin React, sin XLSX, sin ExcelJS).
// Todo lo que está aquí se puede probar con Node / Vitest.
// =====================================================================

export const NOMBRES_MES = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
];

export const NOMBRES_MES_CORTO = [
  "Ene", "Feb", "Mar", "Abr", "May", "Jun",
  "Jul", "Ago", "Sep", "Oct", "Nov", "Dic",
];

// Meses en los que el SENA suele pedir la presentación (según Cristian:
// enero/julio para unas empresas, marzo/septiembre para otras).
export const MESES_PERIODO_SENA = [0, 2, 6, 8];

// Umbral de cobertura para auto-asignar un código CNO sin intervención humana.
// Por debajo de esto, el cargo queda pendiente de revisión manual aunque haya
// una sugerencia (evita que "baja cobertura" se confunda con "confirmado").
// DECISIÓN DE NEGOCIO: valor de partida, ajustar con el equipo si hace falta.
export const COBERTURA_MINIMA_AUTO = 0.75;

// Diferencia máxima tolerada entre la proporción por fechas y la proporción
// por horas efectivas (SS) antes de avisar que difieren.
// DECISIÓN DE NEGOCIO / LEGAL: confirmar este umbral con los abogados.
export const TOLERANCIA_DIFERENCIA_PROPORCION = 0.15;

// ---------------------------------------------------------------------
// Texto
// ---------------------------------------------------------------------
export function normalizar(texto) {
  return String(texto ?? "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

// "DESC. OFICIO" -> "desc oficio" ; "N° ID" -> "n id"
export function normalizarEncabezado(texto) {
  return normalizar(texto).replace(/[^a-z0-9]+/g, " ").trim();
}

// Clave estable de un cargo (para agrupar y para guardar la homologación)
export function claveCargo(cargo) {
  const limpio = String(cargo ?? "").trim().replace(/\s+/g, " ").toUpperCase();
  return limpio || "(SIN CARGO)";
}

// Documento comparable entre nómina y Seguridad Social:
// solo dígitos y sin ceros a la izquierda. Si trae letras (pasaporte, etc.)
// se conserva alfanumérico en mayúsculas.
export function normalizarDocumento(valor) {
  if (valor === null || valor === undefined) return "";
  const s = String(valor).trim();
  if (!s) return "";
  if (/[a-z]/i.test(s)) return s.toUpperCase().replace(/[^A-Z0-9]/g, "");
  const digitos = s.replace(/\D/g, "");
  if (!digitos) return "";
  return digitos.replace(/^0+(?=\d)/, "");
}

// Una fila de Seguridad Social es un subtotal/total por ciudad (o cualquier
// fila de resumen) y no una persona, si en cualquiera de sus celdas aparece
// "subtotal" o si la fila arranca con "total". Estas filas NO deben contarse
// como trabajadores.
export function esFilaSubtotal(fila) {
  if (!Array.isArray(fila)) return false;
  const texto = fila.map((c) => normalizar(String(c ?? ""))).join(" ").trim();
  if (!texto) return false;
  if (/\bsubtotal\b/.test(texto)) return true;
  if (/^total\b/.test(texto)) return true;
  return false;
}

// ---------------------------------------------------------------------
// Fechas
// ---------------------------------------------------------------------
function fechaValida(anio, mes, dia) {
  const d = new Date(anio, mes - 1, dia);
  return d.getFullYear() === anio && d.getMonth() === mes - 1 && d.getDate() === dia ? d : null;
}

// Acepta: Date, serial de Excel (45123), "2026-07-15", "15/07/2026",
// "15-07-2026", "15.07.2026". Devuelve un Date en hora LOCAL a las 00:00
// (sin corrimientos de zona horaria) o null si no se puede interpretar.
export function parseFecha(valor) {
  if (valor === null || valor === undefined || valor === "") return null;

  if (valor instanceof Date) {
    return isNaN(valor) ? null : new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
  }

  if (typeof valor === "number") {
    if (!isFinite(valor) || valor < 20000 || valor > 80000) return null;
    const d = new Date(Math.round((valor - 25569) * 86400000)); // 25569 = 1970-01-01 en Excel
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }

  const s = String(valor).trim();
  if (!s) return null;
  if (/^\d{5}(\.\d+)?$/.test(s)) return parseFecha(Number(s));

  let m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) return fechaValida(+m[1], +m[2], +m[3]);

  m = s.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})/); // dd/mm/yyyy (formato Colombia)
  if (m) return fechaValida(+m[3], +m[2], +m[1]);

  return null;
}

export function formatFecha(fecha) {
  const d = parseFecha(fecha);
  if (!d) return "";
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}/${d.getFullYear()}`;
}

export function diasDelMes(anio, mesIndex) {
  return new Date(anio, mesIndex + 1, 0).getDate();
}

function diffDias(a, b) {
  const ua = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
  const ub = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.round((ub - ua) / 86400000);
}

// Proporción trabajada dentro del mes (0 a 1).
// baseDias = "real" -> divide por los días reales del mes (28-31)
// baseDias = 30     -> divide por 30 (como maneja la Seguridad Social)
// Ingreso y retiro se cuentan como días trabajados (inclusivos).
export function calcularProporcionPorFechas(fechaIngreso, fechaRetiro, anio, mesIndex, baseDias = "real") {
  const totalDias = diasDelMes(anio, mesIndex);
  const inicioMes = new Date(anio, mesIndex, 1);
  const finMes = new Date(anio, mesIndex, totalDias);

  const ingreso = parseFecha(fechaIngreso);
  const retiro = parseFecha(fechaRetiro);

  const inicio = ingreso && ingreso > inicioMes ? ingreso : inicioMes;
  const fin = retiro && retiro < finMes ? retiro : finMes;
  if (fin < inicio) return 0;

  const dias = diffDias(inicio, fin) + 1;
  if (dias >= totalDias) return 1;

  const base = Number(baseDias) === 30 ? 30 : totalDias;
  return Math.min(dias / base, 1);
}

// Horas ordinarias esperadas en el mes para una jornada semanal dada
// (aproximación lineal: jornada/7 * días del mes).
export function horasOrdinariasMes(jornadaSemanal, anio, mesIndex) {
  return (Number(jornadaSemanal) || 0) / 7 * diasDelMes(anio, mesIndex);
}

// Proporción trabajada según horas EFECTIVAS reportadas en Seguridad Social,
// tal como lo pide el SENA (horas efectivas / horas ordinarias), en vez de
// fechas de ingreso/retiro. Útil para contrastar contra calcularProporcionPorFechas.
export function calcularProporcionPorHoras(horasLaboradas, jornadaSemanal, anio, mesIndex) {
  const ordinarias = horasOrdinariasMes(jornadaSemanal, anio, mesIndex);
  if (!ordinarias || !isFinite(Number(horasLaboradas))) return null;
  return Math.min(Math.max(Number(horasLaboradas) / ordinarias, 0), 1);
}

// Texto corto que explica por qué un trabajador no aparece completo en el mes
export function describirNovedad(emp, anio, mesIndex) {
  const inicioMes = new Date(anio, mesIndex, 1);
  const finMes = new Date(anio, mesIndex, diasDelMes(anio, mesIndex));
  const ingreso = parseFecha(emp.fechaIngreso);
  const retiro = parseFecha(emp.fechaRetiro);
  const partes = [];
  if (ingreso && ingreso > inicioMes && ingreso <= finMes) partes.push(`Ingresó el ${formatFecha(ingreso)}`);
  if (retiro && retiro >= inicioMes && retiro < finMes) partes.push(`Se retiró el ${formatFecha(retiro)}`);
  return partes.join(" · ");
}

// ---------------------------------------------------------------------
// Jornada y cuota
// ---------------------------------------------------------------------
// Jornada máxima legal semanal en Colombia (Ley 2101 de 2021), como sugerencia.
// Siempre se puede sobrescribir por empresa o por mes.
const CAMBIOS_JORNADA_LEGAL = [
  { desde: new Date(2023, 6, 15), horas: 47 },
  { desde: new Date(2024, 6, 15), horas: 46 },
  { desde: new Date(2025, 6, 15), horas: 44 },
  { desde: new Date(2026, 6, 15), horas: 42 },
];
const JORNADA_ANTES_2023 = 48;

function jornadaLegalEnFecha(fecha) {
  let jornada = JORNADA_ANTES_2023;
  for (const cambio of CAMBIOS_JORNADA_LEGAL) {
    if (fecha >= cambio.desde) jornada = cambio.horas;
  }
  return jornada;
}

// Jornada legal "puntual" para un mes (referencia: día 15). Se mantiene por
// compatibilidad, pero para meses en los que la ley cambia A MITAD DE MES
// (por ejemplo julio de cada año, el 15) este valor es una aproximación:
// usar jornadaLegalSugeridaDetallada() para el detalle día a día.
export function jornadaLegalSugerida(anio, mesIndex) {
  return jornadaLegalEnFecha(new Date(anio, mesIndex, 15));
}

// Devuelve el detalle de jornada legal para un mes, incluyendo el caso en que
// la ley cambia a mitad del mes (ej. julio 2026: 44h hasta el 14, 42h desde
// el 15). "jornadaPromedio" es un promedio ponderado por días, redondeado al
// entero más cercano, pensado como sugerencia para digitar en pantalla; el
// código NUNCA decide solo por su cuenta, por eso también se marca "dividido".
export function jornadaLegalSugeridaDetallada(anio, mesIndex) {
  const totalDias = diasDelMes(anio, mesIndex);
  const cambioEnMes = CAMBIOS_JORNADA_LEGAL.find(
    (c) => c.desde.getFullYear() === anio && c.desde.getMonth() === mesIndex
  );

  if (!cambioEnMes) {
    return { jornada: jornadaLegalEnFecha(new Date(anio, mesIndex, 15)), dividido: false };
  }

  const diaCambio = cambioEnMes.desde.getDate(); // p.ej. 15
  const diasAntes = diaCambio - 1;
  const diasDespues = totalDias - diasAntes;
  const jornadaAntes = jornadaLegalEnFecha(new Date(anio, mesIndex, diaCambio - 1));
  const jornadaDespues = cambioEnMes.horas;
  const promedio = (jornadaAntes * diasAntes + jornadaDespues * diasDespues) / totalDias;

  return {
    jornada: Math.round(promedio),
    dividido: true,
    jornadaAntes,
    jornadaDespues,
    diaCambio,
    fechaCambio: cambioEnMes.desde,
    mensaje: `Este mes cambia de ${jornadaAntes}h a ${jornadaDespues}h el día ${diaCambio}. Se sugiere ${Math.round(promedio)}h como promedio; confírmalo o ajústalo a mano.`,
  };
}

// Cuota de aprendices: 1 por cada 20 trabajadores + 1 adicional por fracción
// de 10 o más; empresas con 15 a 20 trabajadores: 1 aprendiz; menos de 15: 0.
// (Ley 789/2002 y Decreto 933/2003 — confírmalo con el asesor legal.)
export function calcularCuotaAprendices(totalTrabajadores) {
  const n = Math.floor(Number(totalTrabajadores) || 0);
  if (n < 15) return 0;
  return Math.max(1, Math.floor(n / 20) + (n % 20 >= 10 ? 1 : 0));
}

// Texto humano de la regla anterior, para que la pantalla y el Excel nunca
// digan algo distinto de lo que el código realmente calcula.
export const TEXTO_REGLA_CUOTA_APRENDICES =
  "1 aprendiz por cada 20 trabajadores, más 1 adicional cuando el residuo es de 10 o más (mínimo 1 desde 15 trabajadores).";

// ---------------------------------------------------------------------
// Detección de columnas (nómina y Seguridad Social)
// ---------------------------------------------------------------------
function buscarColumna(encabezados, aliases, excluir = []) {
  // 1) coincidencia exacta con algún alias
  for (const alias of aliases) {
    const i = encabezados.findIndex((h) => h === alias);
    if (i !== -1) return i;
  }
  // 2) el encabezado contiene el alias (evitando columnas tipo "tipo documento")
  for (const alias of aliases) {
    const i = encabezados.findIndex(
      (h) => h && h.includes(alias) && !excluir.some((x) => h.includes(x))
    );
    if (i !== -1) return i;
  }
  return -1;
}

const ALIAS_NOMINA = {
  documento: {
    aliases: ["documento identidad", "documento de identidad", "numero de documento", "numero documento", "documento", "cedula", "identificacion", "no id", "nro id"],
    excluir: ["tipo", "expedicion"],
  },
  cargo: {
    aliases: ["desc oficio", "descripcion oficio", "descripcion del oficio", "cargo", "oficio", "puesto", "job title", "position", "ocupacion"],
    excluir: ["cod", "codigo", "centro", "tipo", "nivel"],
  },
  nombre: {
    aliases: ["nombre completo", "nombres y apellidos", "apellidos y nombres", "nombre", "empleado"],
    excluir: ["cargo", "oficio", "banco", "eps"],
  },
  ingreso: {
    aliases: ["fecha ingreso", "fecha de ingreso", "fecha inicio contrato", "ingreso"],
    excluir: ["salario", "base", "valor", "total"],
  },
  antiguedad: { aliases: ["fecha antiguedad"], excluir: [] },
  retiro: {
    aliases: ["fecha retiro", "fecha de retiro", "fecha terminacion", "fecha de terminacion", "fecha fin contrato", "retiro"],
    excluir: ["valor", "aporte", "motivo", "causa", "indemnizacion"],
  },
};

// filas: arreglo de arreglos (XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null }))
// NOTA: la nómina GTN no trae fecha de retiro fiable; por eso el llamador
// (App.jsx) debe completar fechaRetiro con lo que venga de Seguridad Social
// cuando esta columna quede vacía aquí. Ver fusionarFechaRetiro().
export function parsearNomina(filas) {
  const limite = Math.min(filas.length, 40);
  let indiceEncabezado = -1;
  let col = null;

  for (let i = 0; i < limite; i++) {
    const fila = filas[i];
    if (!Array.isArray(fila)) continue;
    const enc = fila.map(normalizarEncabezado);
    const doc = buscarColumna(enc, ALIAS_NOMINA.documento.aliases, ALIAS_NOMINA.documento.excluir);
    const cargo = buscarColumna(enc, ALIAS_NOMINA.cargo.aliases, ALIAS_NOMINA.cargo.excluir);
    if (doc !== -1 && cargo !== -1 && doc !== cargo) {
      indiceEncabezado = i;
      col = {
        documento: doc,
        cargo,
        nombre: buscarColumna(enc, ALIAS_NOMINA.nombre.aliases, ALIAS_NOMINA.nombre.excluir),
        ingreso: buscarColumna(enc, ALIAS_NOMINA.ingreso.aliases, ALIAS_NOMINA.ingreso.excluir),
        antiguedad: buscarColumna(enc, ALIAS_NOMINA.antiguedad.aliases),
        retiro: buscarColumna(enc, ALIAS_NOMINA.retiro.aliases, ALIAS_NOMINA.retiro.excluir),
      };
      break;
    }
  }

  if (indiceEncabezado === -1) {
    throw new Error(
      "No se encontraron las columnas de documento y cargo (ej. 'DOCUMENTO IDENTIDAD' y 'DESC. OFICIO') en las primeras 40 filas."
    );
  }

  const empleados = [];
  for (const fila of filas.slice(indiceEncabezado + 1)) {
    if (!Array.isArray(fila)) continue;
    const documentoOriginal = fila[col.documento];
    const documento = normalizarDocumento(documentoOriginal);
    if (!documento) continue;

    const cargoTexto = String(fila[col.cargo] ?? "").trim();
    const ingreso = col.ingreso !== -1 ? parseFecha(fila[col.ingreso]) : null;
    const antiguedad = col.antiguedad !== -1 ? parseFecha(fila[col.antiguedad]) : null;

    empleados.push({
      documento,
      documentoOriginal: String(documentoOriginal).trim(),
      nombre: col.nombre !== -1 ? String(fila[col.nombre] ?? "").trim() : "",
      cargo: cargoTexto || "(SIN CARGO)",
      cargoKey: claveCargo(cargoTexto),
      fechaIngreso: ingreso || antiguedad,
      // La nómina GTN no trae esta columna de forma confiable; queda en null
      // si no se encontró columna, para que App.jsx la complete desde SS.
      fechaRetiro: col.retiro !== -1 ? parseFecha(fila[col.retiro]) : null,
      fuente: "nomina",
    });
  }

  return { empleados, columnasDetectadas: col };
}

const ALIAS_DOC_SS_EXACTOS = [
  "no id", "n id", "nro id", "numero de identificacion", "numero identificacion",
  "no identificacion", "numero de documento", "numero documento", "identificacion", "documento", "cedula",
];

const ALIAS_SS = {
  ingreso: { aliases: ["fecha ingreso", "fecha de ingreso", "fecha afiliacion", "fecha de afiliacion"], excluir: [] },
  retiro: { aliases: ["fecha retiro", "fecha de retiro", "fecha novedad retiro", "fecha desafiliacion"], excluir: [] },
  horas: { aliases: ["horas laboradas", "horas trabajadas", "total horas laboradas", "no horas laboradas", "horas"], excluir: ["ordinarias", "extras"] },
};

// filas: arreglo de arreglos (igual formato que parsearNomina).
// Devuelve { documentos: Map(documento -> { fechaIngreso, fechaRetiro, horasLaboradas }) }.
// Ignora filas de subtotal/total por ciudad (no son personas).
export function parsearSeguridadSocial(filas) {
  const limite = Math.min(filas.length, 60);
  let indiceEncabezado = -1;
  let col = null;

  for (let i = 0; i < limite; i++) {
    const fila = filas[i];
    if (!Array.isArray(fila)) continue;
    const enc = fila.map(normalizarEncabezado);
    const idx = buscarColumna(enc, ALIAS_DOC_SS_EXACTOS, ["tipo"]);
    if (idx !== -1) {
      indiceEncabezado = i;
      col = {
        documento: idx,
        ingreso: buscarColumna(enc, ALIAS_SS.ingreso.aliases, ALIAS_SS.ingreso.excluir),
        retiro: buscarColumna(enc, ALIAS_SS.retiro.aliases, ALIAS_SS.retiro.excluir),
        horas: buscarColumna(enc, ALIAS_SS.horas.aliases, ALIAS_SS.horas.excluir),
      };
      break;
    }
  }

  if (indiceEncabezado === -1) {
    throw new Error(
      "No se encontró una columna de identificación reconocible (ej. 'No id', 'Documento', 'Cédula') en la planilla de Seguridad Social."
    );
  }

  const documentos = new Map();
  let filasSubtotalIgnoradas = 0;

  for (const fila of filas.slice(indiceEncabezado + 1)) {
    if (!Array.isArray(fila)) continue;
    if (esFilaSubtotal(fila)) {
      filasSubtotalIgnoradas++;
      continue;
    }
    const doc = normalizarDocumento(fila[col.documento]);
    if (!doc || !/\d/.test(doc)) continue;

    const previo = documentos.get(doc) || {};
    const ingreso = col.ingreso !== -1 ? parseFecha(fila[col.ingreso]) : null;
    const retiro = col.retiro !== -1 ? parseFecha(fila[col.retiro]) : null;
    const horasCelda = col.horas !== -1 ? Number(fila[col.horas]) : NaN;

    documentos.set(doc, {
      fechaIngreso: previo.fechaIngreso || ingreso || null,
      fechaRetiro: retiro || previo.fechaRetiro || null,
      horasLaboradas: isFinite(horasCelda) ? horasCelda : (previo.horasLaboradas ?? null),
    });
  }

  return { documentos, columnasDetectadas: col, filasSubtotalIgnoradas };
}

// Completa la fecha de retiro de un empleado de nómina con la de Seguridad
// Social cuando la nómina no la trae (caso GTN). No sobreescribe una fecha
// de nómina que sí exista.
export function fusionarFechaRetiro(empleadoNomina, registroSS) {
  if (empleadoNomina.fechaRetiro) return empleadoNomina.fechaRetiro;
  return registroSS?.fechaRetiro || null;
}

// Busca el cargo de un documento en las nóminas de OTROS meses (recorridas en
// el orden en que se cargaron). Devuelve el primer cargo encontrado o null.
// nominasPorMes: [{ etiqueta, empleados }] (empleados ya parseados)
export function buscarCargoEnOtrasNominas(documento, nominasPorMes) {
  for (const { etiqueta, empleados } of nominasPorMes) {
    const encontrado = empleados.find((e) => e.documento === documento);
    if (encontrado) return { cargo: encontrado.cargo, cargoKey: encontrado.cargoKey, etiqueta };
  }
  return null;
}

// ---------------------------------------------------------------------
// Homologación de cargos -> CNO
// ---------------------------------------------------------------------
const STOPWORDS = new Set([
  "de", "del", "la", "el", "los", "las", "y", "e", "en", "al", "para", "por", "con", "un", "una",
  "jr", "sr", "junior", "senior", "ii", "iii", "iv", "nivel", "especialista",
]);

function tokenizar(texto) {
  return normalizar(texto)
    .split(/[^a-z0-9]+/)
    .filter((t) => t.length > 2 && !STOPWORDS.has(t));
}

// "auxiliares" -> "auxiliar", "gerentes" -> "gerent" (se compara por prefijo)
function raiz(token) {
  return token.length > 4 ? token.replace(/(es|s)$/, "") : token;
}

// Prepara el listado una sola vez (normaliza y separa en palabras)
export function indexarListadoCno(listado) {
  return (listado || []).map((item) => {
    const norm = normalizar(item.ocupacion);
    return { ...item, _norm: norm, _palabras: norm.split(/[^a-z0-9]+/).filter(Boolean) };
  });
}

export function sugerirCno(texto, indice, limite = 8) {
  const consulta = String(texto ?? "").trim();
  if (!consulta) return [];

  // Búsqueda por código (ej. "1232" o "12")
  if (/^\d{2,4}$/.test(consulta)) {
    return indice.filter((it) => it.codigo.startsWith(consulta)).slice(0, limite).map((it) => ({ ...it, puntos: 0, cobertura: 0 }));
  }

  const tokens = [...new Set(tokenizar(consulta).map(raiz))];
  if (tokens.length === 0) return [];
  const normConsulta = normalizar(consulta);

  const resultados = [];
  for (const it of indice) {
    let aciertos = 0;
    for (const t of tokens) {
      if (it._palabras.some((p) => p.startsWith(t))) aciertos++;
    }
    if (aciertos === 0) continue;
    const cobertura = aciertos / tokens.length;
    const precision = aciertos / Math.max(it._palabras.length, 1);
    let puntos = cobertura * 10 + precision * 2;
    if (it._norm === normConsulta) puntos += 100;
    resultados.push({ codigo: it.codigo, ocupacion: it.ocupacion, puntos, cobertura });
  }

  return resultados.sort((a, b) => b.puntos - a.puntos || a.ocupacion.length - b.ocupacion.length).slice(0, limite);
}

// Primer texto que aparece en el listado para un código (suele ser el título del grupo)
export function nombreDeCodigo(codigo, indice) {
  return indice.find((it) => it.codigo === codigo)?.ocupacion || "";
}

// Sugerencias de partida (deben ser CONFIRMADAS por una persona: ver punto B.1
// de la revisión — los códigos de aquí y los del modelo de Excel no coinciden
// y hay que decidir cuáles son los correctos con el equipo/abogados antes de
// confiar en ellos). Solo se incluyen las de las que hay bastante certeza.
export const MAPEO_SUGERIDO = {
  "IT SUPPORT SPECIALIST": "2281",
  "SALESFORCE JUNIOR DEVELOPER": "2173",
  "ADMINISTRATIVE ANALYST": "1122",
  "CREDIT AND COLLECTIONS ANALYST II": "1232",
  "SALES & PRICING TEAM LEADER": "6211",
};

// cargos: Map(cargoKey -> texto original)
// traducciones: { CARGO_KEY: "traducción" }
// previo: { CARGO_KEY: { es, codigo, ocupacion, confirmado, noCalificado } } (guardado o en pantalla)
//
// Estados posibles de un cargo homologado:
//  - codigo con confirmado=true            -> ya resuelto
//  - noCalificado=true (codigo="")         -> alguien decidió a mano que es un oficio no calificado
//  - codigo="" y sugerenciaBaja=null        -> "no encontré ningún código" (sin pista)
//  - codigo="" y sugerenciaBaja={...}       -> "encontré algo pero con baja cobertura", NO se auto-asigna
export function construirHomologacion(cargos, traducciones, previo, indice) {
  const resultado = {};
  for (const [key, original] of cargos) {
    const p = previo?.[key];
    if (p && (p.es || p.codigo !== undefined || p.noCalificado)) {
      resultado[key] = {
        original,
        es: p.es || traducciones?.[key] || original,
        codigo: p.codigo ?? "",
        ocupacion: p.ocupacion || (p.codigo ? nombreDeCodigo(p.codigo, indice) : ""),
        confirmado: Boolean(p.confirmado),
        noCalificado: Boolean(p.noCalificado),
        sugerenciaBaja: p.sugerenciaBaja || null,
      };
      continue;
    }

    const es = (traducciones?.[key] || "").trim() || original;
    const directo = MAPEO_SUGERIDO[key];
    const itemDirecto = directo ? indice.find((it) => it.codigo === directo) : null;

    if (itemDirecto) {
      // Viene del mapeo curado a mano: se preselecciona pero SIGUE sin confirmar,
      // para forzar la revisión humana del punto B.1.
      resultado[key] = {
        original, es, codigo: directo, ocupacion: nombreDeCodigo(directo, indice),
        confirmado: false, noCalificado: false, sugerenciaBaja: null,
      };
      continue;
    }

    const top = sugerirCno(es, indice, 1)[0];
    const coberturaOk = top && top.cobertura >= COBERTURA_MINIMA_AUTO;

    resultado[key] = {
      original,
      es,
      codigo: coberturaOk ? top.codigo : "",
      ocupacion: coberturaOk ? top.ocupacion : "",
      confirmado: false,
      noCalificado: false,
      // Si hubo una sugerencia pero de baja cobertura, se guarda como pista
      // para la persona que revisa, pero NO se auto-asigna (punto B.2).
      sugerenciaBaja: !coberturaOk && top ? { codigo: top.codigo, ocupacion: top.ocupacion, cobertura: top.cobertura } : null,
    };
  }
  return resultado;
}

// Un cargo está "listo para confirmar en bloque" solo si ya tiene código o si
// alguien lo marcó explícitamente como no calificado. Los de baja cobertura
// (sugerenciaBaja) o sin ninguna pista quedan fuera de "Confirmar todos".
export function homologacionListaParaConfirmar(h) {
  return Boolean(h.codigo) || Boolean(h.noCalificado);
}

// ---------------------------------------------------------------------
// Persistencia por empresa (localStorage; falla en silencio si no existe)
// Guarda TODO lo que depende de la empresa en un solo lugar: la homologación
// confirmada, la jornada, la base de días, los datos de la empresa para el
// encabezado del Excel y las asignaciones manuales de cargo.
// ---------------------------------------------------------------------
const PREFIJO_STORAGE = "sena-config:v2:";
const claveStorage = (empresa) => PREFIJO_STORAGE + (normalizar(empresa).trim() || "_sin_empresa");

export function cargarConfiguracionEmpresa(empresa) {
  try {
    const raw = localStorage.getItem(claveStorage(empresa));
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

// config: objeto parcial, se combina con lo ya guardado (merge superficial,
// salvo "homologacion" y "asignacionesManuales" y "jornadasPorMes" que se
// combinan a su vez internamente para no perder lo ya confirmado).
export function guardarConfiguracionEmpresa(empresa, config) {
  try {
    const existente = cargarConfiguracionEmpresa(empresa);
    const siguiente = { ...existente, ...config };

    if (config.homologacion) {
      const confirmadas = {};
      Object.entries(config.homologacion).forEach(([key, h]) => {
        if (h.confirmado || h.noCalificado) {
          confirmadas[key] = { es: h.es, codigo: h.codigo, ocupacion: h.ocupacion, confirmado: h.confirmado, noCalificado: h.noCalificado };
        }
      });
      siguiente.homologacion = { ...(existente.homologacion || {}), ...confirmadas };
    }
    if (config.asignacionesManuales) {
      siguiente.asignacionesManuales = { ...(existente.asignacionesManuales || {}), ...config.asignacionesManuales };
    }
    if (config.jornadasPorMes) {
      siguiente.jornadasPorMes = { ...(existente.jornadasPorMes || {}), ...config.jornadasPorMes };
    }

    localStorage.setItem(claveStorage(empresa), JSON.stringify(siguiente));
  } catch {
    /* sin almacenamiento disponible */
  }
}

// Compatibilidad hacia atrás con el nombre anterior usado en App.jsx.
export function cargarHomologacionGuardada(empresa) {
  return cargarConfiguracionEmpresa(empresa).homologacion || {};
}
export function guardarHomologacion(empresa, homologacion) {
  if (!empresa || !normalizar(empresa).trim()) return; // punto C.6: no guardar con empresa vacía
  guardarConfiguracionEmpresa(empresa, { homologacion });
}

// ---------------------------------------------------------------------
// Agrupación para las matrices
// ---------------------------------------------------------------------
// El modelo del SENA usa horas ENTERAS (20, 28, 26, 24...), no con decimales.
const redondearHoras = (n) => Math.round(n);

// Categorías FIJAS de "oficios no calificados" (Matriz 2) tal como vienen en
// el formulario oficial del SENA. Deben aparecer siempre en este orden, aunque
// tengan 0 personas ese mes; cualquier cargo no calificado que no calce en
// ninguna de estas categorías queda marcado como "categoriaSinAsignar" para
// que se revise a mano (no se inventa una categoría oficial que no exista).
export const CATEGORIAS_NO_CALIFICADAS = [
  "Conductor",
  "Vigilante",
  "Mensajero",
  "Personal de Aseo y Cafetería",
  "Empleados FIC",
  "Contratos de Aprendizaje",
];

// Palabras clave (ya normalizadas, sin tildes) para reconocer automáticamente
// a cuál de las categorías fijas pertenece un cargo no calificado. Es una
// heurística de partida: lo que no calce queda pendiente de revisión manual,
// nunca se fuerza a una categoría que no corresponda.
const PALABRAS_CATEGORIA_NO_CALIFICADA = {
  "Conductor": ["conductor", "chofer", "chofer de"],
  "Vigilante": ["vigilante", "guarda de seguridad", "guarda", "seguridad fisica", "escolta"],
  "Mensajero": ["mensajero", "domiciliario", "repartidor", "courier"],
  "Personal de Aseo y Cafetería": ["aseo", "cafeteria", "limpieza", "servicios generales"],
  "Empleados FIC": ["fic"],
  "Contratos de Aprendizaje": ["aprendiz", "practicante", "pasante"],
};

// Devuelve el nombre de la categoría fija que corresponde a un cargo no
// calificado (según palabras clave), o null si no calza con ninguna.
export function categorizarNoCalificado(nombreCargo) {
  const norm = normalizar(nombreCargo);
  for (const categoria of CATEGORIAS_NO_CALIFICADAS) {
    const palabras = PALABRAS_CATEGORIA_NO_CALIFICADA[categoria] || [];
    if (palabras.some((p) => norm.includes(p))) return categoria;
  }
  return null;
}

// Orden oficial de una categoría no calificada para ordenar la Matriz 2;
// cualquier cargo que no calce en las categorías fijas queda al final.
export function ordenNoCalificado(nombreCargo) {
  const i = CATEGORIAS_NO_CALIFICADAS.indexOf(nombreCargo);
  return i === -1 ? CATEGORIAS_NO_CALIFICADAS.length : i;
}

// Devuelve un grupo por código CNO (calificados) o por nombre de cargo (sin código).
// Cada grupo trae "lineas": completos + parciales agrupados por horas iguales.
//
// opciones.forzarCategoriasFijas: cuando es true (se usa al exportar la Matriz
// 2 oficial), los cargos no calificados se agrupan bajo las categorías fijas
// del SENA (CATEGORIAS_NO_CALIFICADAS) en vez de por el nombre libre del
// cargo, y las 6 categorías siempre aparecen así tengan 0 personas. Un cargo
// no calificado que no calce en ninguna categoría queda con
// "categoriaSinAsignar: true" para que se revise a mano en vez de inventarle
// una categoría oficial que no le corresponde.
export function agruparPorCargo(empleados, homologacion, jornadaSemanal, opciones = {}) {
  const { forzarCategoriasFijas = false } = opciones;
  const grupos = new Map();

  for (const emp of empleados) {
    if (!(emp.proporcion > 0)) continue;
    const h = homologacion?.[emp.cargoKey] || {};
    const codigo = h.codigo || "";
    const nombreTraducido = h.es || emp.cargo;

    const categoriaFija = !codigo && forzarCategoriasFijas ? categorizarNoCalificado(nombreTraducido) : null;
    const nombre = codigo ? nombreTraducido : categoriaFija || nombreTraducido;
    const sinCategoria = !codigo && forzarCategoriasFijas && !categoriaFija;

    const clave = codigo
      ? `CNO::${codigo}`
      : forzarCategoriasFijas
      ? categoriaFija
        ? `SIN::${normalizar(categoriaFija)}`
        : `SIN::OTROS::${normalizar(nombreTraducido)}`
      : `SIN::${normalizar(nombreTraducido)}`;

    if (!grupos.has(clave)) {
      grupos.set(clave, {
        codigo,
        nombreCargo: nombre,
        calificado: Boolean(codigo),
        categoriaSinAsignar: sinCategoria,
        completos: 0,
        parciales: new Map(), // horas -> trabajadores
        personas: 0,
      });
    }
    const g = grupos.get(clave);
    g.personas += 1;
    if (emp.proporcion >= 1) {
      g.completos += 1;
    } else {
      const horas = redondearHoras(jornadaSemanal * emp.proporcion);
      g.parciales.set(horas, (g.parciales.get(horas) || 0) + 1);
    }
  }

  // Con categorías fijas, las 6 deben aparecer siempre en la Matriz 2, así
  // ese mes no haya nadie en alguna de ellas (el formulario del SENA las trae
  // fijas, con o sin personas).
  if (forzarCategoriasFijas) {
    CATEGORIAS_NO_CALIFICADAS.forEach((categoria) => {
      const clave = `SIN::${normalizar(categoria)}`;
      if (!grupos.has(clave)) {
        grupos.set(clave, {
          codigo: "",
          nombreCargo: categoria,
          calificado: false,
          categoriaSinAsignar: false,
          completos: 0,
          parciales: new Map(),
          personas: 0,
        });
      }
    });
  }

  return [...grupos.values()].map((g) => {
    const lineas = [];
    if (g.completos > 0) {
      lineas.push({ trabajadores: g.completos, jornada: jornadaSemanal, total: g.completos * jornadaSemanal, parcial: false });
    }
    [...g.parciales.entries()]
      .sort((a, b) => b[0] - a[0])
      .forEach(([horas, n]) => lineas.push({ trabajadores: n, jornada: horas, total: redondearHoras(n * horas), parcial: true }));
    return {
      codigo: g.codigo,
      nombreCargo: g.nombreCargo,
      calificado: g.calificado,
      categoriaSinAsignar: g.categoriaSinAsignar,
      personas: g.personas,
      lineas,
      totalHoras: redondearHoras(lineas.reduce((a, l) => a + l.total, 0)),
    };
  });
}

// resultados: [{ empleados, jornadaSemanal, etiqueta }]  (uno por mes, en orden cronológico)
// Grilla código-por-mes con SUMA y PROMEDIO (÷6, o ÷meses cargados) de personas,
// tal como la pide el formulario (punto D.4). "columnas" trae un valor de
// personas por cada mes, en el mismo orden que "resultados".
// También usa las categorías fijas del SENA para los no calificados, igual
// que en la exportación mensual, para que la plantilla promedio y las hojas
// de cada mes hablen siempre de los mismos grupos.
export function calcularPlantillaPromedio(resultados, homologacion) {
  const n = resultados.length || 1;
  const acumulado = new Map();

  resultados.forEach((r, idx) => {
    agruparPorCargo(r.empleados, homologacion, r.jornadaSemanal, { forzarCategoriasFijas: true }).forEach((g) => {
      const clave = g.calificado ? `CNO::${g.codigo}` : `SIN::${normalizar(g.nombreCargo)}`;
      if (!acumulado.has(clave)) {
        acumulado.set(clave, {
          codigo: g.codigo,
          nombre: g.nombreCargo,
          calificado: g.calificado,
          personasPorMes: new Array(resultados.length).fill(0),
        });
      }
      acumulado.get(clave).personasPorMes[idx] += g.personas;
    });
  });

  const filas = [...acumulado.values()]
    .map((a) => {
      const suma = a.personasPorMes.reduce((x, y) => x + y, 0);
      return {
        codigo: a.codigo,
        nombre: a.nombre,
        calificado: a.calificado,
        personasPorMes: a.personasPorMes,
        sumaPersonas: suma,
        promedioPersonas: Math.round((suma / n) * 100) / 100,
      };
    })
    .sort((x, y) => {
      if (x.calificado !== y.calificado) return y.calificado - x.calificado;
      if (!x.calificado) return ordenNoCalificado(x.nombre) - ordenNoCalificado(y.nombre) || x.nombre.localeCompare(y.nombre, "es");
      return String(x.codigo).localeCompare(String(y.codigo));
    });

  const totalPorMes = resultados.map((r) => r.empleados.filter((e) => e.proporcion > 0).length);
  const totalPersonas = Math.round((totalPorMes.reduce((a, b) => a + b, 0) / n) * 100) / 100;

  return {
    filas,
    meses: resultados.length,
    etiquetasMes: resultados.map((r) => r.etiqueta),
    promedioTotalPersonas: totalPersonas,
    totalPorMes,
  };
}