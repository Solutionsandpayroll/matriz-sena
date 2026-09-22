import ExcelJS from "exceljs";
import {
  agruparPorCargo,
  calcularPlantillaPromedio,
  normalizar,
  NOMBRES_MES_CORTO,
  TEXTO_REGLA_CUOTA_APRENDICES,
  ordenNoCalificado,
} from "./sena";

// Logo que se imprime en el Excel que se presenta al SENA: es el logo oficial
// del SENA (no el de Solutions & Payroll, que solo se usa en el header web de
// App.jsx). Debe existir el archivo /image.png en la carpeta /public.
export const URL_LOGO = "/image.png";

const GRIS = "FFD9D9D9";
const NEGRO = { argb: "FF000000" };
const BORDE = {
  top: { style: "thin", color: NEGRO },
  bottom: { style: "thin", color: NEGRO },
  left: { style: "thin", color: NEGRO },
  right: { style: "thin", color: NEGRO },
};

// Convierte "solutions and payroll" -> "Solutions And Payroll" para el
// encabezado del Excel. Si una palabra ya viene toda en mayúsculas (ej. una
// sigla como "S.A.S" o "CIPY"), se respeta tal cual y no se toca.
function capitalizarNombreEmpresa(nombre) {
  if (!nombre) return "";
  return nombre
    .trim()
    .split(/\s+/)
    .map((palabra) => {
      const esSigla = palabra === palabra.toUpperCase() && palabra !== palabra.toLowerCase();
      if (esSigla) return palabra;
      return palabra.charAt(0).toUpperCase() + palabra.slice(1).toLowerCase();
    })
    .join(" ");
}

// ---------------------------------------------------------------------
// Logo
// ---------------------------------------------------------------------
export async function cargarLogo(url = URL_LOGO) {
  const resp = await fetch(url);
  if (!resp.ok) throw new Error(`No se pudo leer ${url} (HTTP ${resp.status}).`);
  const blob = await resp.blob();
  // Vite responde con index.html (HTTP 200) cuando el archivo no existe en /public.
  if (!blob.type.startsWith("image/")) {
    throw new Error(`${url} no es una imagen. Verifica que el archivo exista en la carpeta /public.`);
  }
  const extension = /jpe?g/.test(blob.type) ? "jpeg" : /gif/.test(blob.type) ? "gif" : "png";
  if (!/png|jpe?g|gif/.test(blob.type)) throw new Error(`Formato de logo no soportado (${blob.type}). Usa PNG o JPG.`);

  const base64 = await new Promise((resolve, reject) => {
    const lector = new FileReader();
    lector.onloadend = () => resolve(lector.result);
    lector.onerror = reject;
    lector.readAsDataURL(blob);
  });

  const dims = await new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ ancho: img.naturalWidth, alto: img.naturalHeight });
    img.onerror = reject;
    img.src = base64;
  });

  return { base64, extension, ...dims };
}

// ---------------------------------------------------------------------
// Utilidades de hoja
// ---------------------------------------------------------------------
const nombreHojaSeguro = (t) => String(t).replace(/[\\/*?:[\]]/g, "").slice(0, 31);
const escaparFooter = (t) => String(t).replace(/&/g, "&&");

// Nombre corto de pestaña ("Mar") y fecha larga para mostrar ("Mar / 2026").
// Como la ventana siempre son 6 meses consecutivos, la abreviatura del mes
// nunca se repite dentro de un mismo libro.
export function etiquetaCortaMes(anio, mesIndex) {
  return NOMBRES_MES_CORTO[mesIndex];
}
export function etiquetaFechaMes(anio, mesIndex) {
  return `${NOMBRES_MES_CORTO[mesIndex]} / ${anio}`;
}

const estiloEncabezado = (celda) => {
  celda.font = { bold: true, size: 8 };
  celda.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRIS } };
  celda.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  celda.border = BORDE;
};

const estiloDato = (celda, { cursiva = false, izquierda = false } = {}) => {
  celda.font = { size: 9, italic: cursiva };
  celda.alignment = { horizontal: izquierda ? "left" : "center", vertical: "middle", wrapText: true };
  celda.border = BORDE;
};

const estiloTotal = (celda) => {
  celda.font = { bold: true, size: 9 };
  celda.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRIS } };
  celda.alignment = { horizontal: "center", vertical: "middle" };
  celda.border = BORDE;
};

// Ancho de columna (en "caracteres", como lo guarda ws.columns) convertido a
// píxeles, usando la fórmula real de Excel para la fuente por defecto
// (Calibri 11, MDW = ancho en px del carácter "0" = 7). Es la misma cuenta
// que usa Excel para dibujar la columna en pantalla, así que la posición del
// logo coincide con lo que se ve al abrir el archivo.
const MDW = 7;
const anchoColPx = (w) => Math.round(((256 * w + Math.floor(128 / MDW)) / 256) * MDW);

// Calcula la celda (columna + fracción) donde debe ir la esquina superior
// izquierda de una imagen para que quede centrada horizontalmente sobre las
// columnas 1..n de la hoja.
function anchorCentrado(anchosCol, anchoImgPx) {
  const totalPx = anchosCol.reduce((suma, w) => suma + anchoColPx(w), 0);
  let restantePx = Math.max((totalPx - anchoImgPx) / 2, 0);

  let col = 0;
  while (col < anchosCol.length - 1 && restantePx >= anchoColPx(anchosCol[col])) {
    restantePx -= anchoColPx(anchosCol[col]);
    col++;
  }
  const anchoColActualPx = anchoColPx(anchosCol[col]);
  const fraccion = anchoColActualPx > 0 ? Math.min(restantePx / anchoColActualPx, 0.95) : 0;
  return { col: col + fraccion, row: 0.05 };
}

// Bloque con los datos de la empresa en el formato oficial del SENA: título
// "DATOS BASICOS DE LA EMPRESA" y una tabla de 4 filas con pares
// etiqueta/valor (Razón social/Nit, Representante legal/CC,
// Dirección/Teléfonos, E-mail/Fecha), tal como viene en el formulario. Nit,
// CC y Teléfonos se guardan como texto para que Excel no les cambie el
// formato (ceros a la izquierda, guiones, etc.).
// Este bloque SOLO se imprime en las hojas de mes (Matriz 1 / Matriz 2);
// la hoja de Plantilla promedio no lo recibe, así que nunca se dibuja ahí.
function filaDatosEmpresa(ws, n, datosEmpresa, fechaTexto) {
  const d = datosEmpresa || {};
  const hayDatos = d.razonSocial || d.nit || d.representanteLegal || d.cc || d.direccion || d.telefonos || d.email || fechaTexto;
  if (!hayDatos) return;

  const titulo = ws.addRow(["DATOS BASICOS DE LA EMPRESA"]);
  ws.mergeCells(titulo.number, 1, titulo.number, n);
  titulo.getCell(1).font = { bold: true, size: 9 };
  titulo.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRIS } };
  titulo.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
  titulo.getCell(1).border = BORDE;
  titulo.height = 16;

  // Mitad izquierda: etiqueta 1 + valor 1. Mitad derecha: etiqueta 2 + valor 2.
  // Se calcula sobre "n" (el número real de columnas de cada hoja) para que
  // la tabla se vea igual sin importar si la hoja tiene 4, 6 o 10 columnas.
  const mitad = Math.max(2, Math.ceil(n / 2));
  const colEtiqueta1 = 1;
  const colValor1Inicio = 2;
  const colValor1Fin = mitad;
  const colEtiqueta2 = mitad + 1;
  const colValor2Inicio = mitad + 2;
  const colValor2Fin = n;

  const filaPar = (etiqueta1, valor1, etiqueta2, valor2, opciones = {}) => {
    const f = ws.addRow([]);
    for (let c = 1; c <= n; c++) f.getCell(c).border = BORDE;

    f.getCell(colEtiqueta1).value = etiqueta1;
    f.getCell(colValor1Inicio).value = valor1 || "";
    f.getCell(colEtiqueta2).value = etiqueta2 || "";
    f.getCell(colValor2Inicio).value = valor2 || "";

    if (colValor1Fin > colValor1Inicio) ws.mergeCells(f.number, colValor1Inicio, f.number, colValor1Fin);
    if (colValor2Fin > colValor2Inicio) ws.mergeCells(f.number, colValor2Inicio, f.number, colValor2Fin);

    [colEtiqueta1, colEtiqueta2].forEach((c) => {
      const celda = f.getCell(c);
      celda.font = { size: 8, bold: true };
      celda.alignment = { vertical: "middle", horizontal: "right", wrapText: true };
    });
    [colValor1Inicio, colValor2Inicio].forEach((c) => {
      const celda = f.getCell(c);
      celda.font = { size: 8 };
      celda.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    });

    // Nit, CC y Teléfonos son "códigos": se guardan como texto para que
    // Excel nunca les aplique separador de miles, notación científica ni les
    // quite ceros a la izquierda.
    if (opciones.textoValor1) f.getCell(colValor1Inicio).numFmt = "@";
    if (opciones.textoValor2) f.getCell(colValor2Inicio).numFmt = "@";

    if (opciones.correoValor1 && valor1) {
      f.getCell(colValor1Inicio).value = { text: valor1, hyperlink: `mailto:${valor1}` };
      f.getCell(colValor1Inicio).font = { size: 8, color: { argb: "FF0563C1" }, underline: true };
    }
    if (opciones.negrillaValor2) f.getCell(colValor2Inicio).font = { size: 8, bold: true };

    f.height = 15;
    return f;
  };

  filaPar("Razón Social", d.razonSocial, "Nit:", d.nit, { textoValor2: true });
  filaPar("Representante Legal:", d.representanteLegal, "CC:", d.cc, { textoValor2: true });
  filaPar("Dirección (municipio/barrio)", d.direccion, "Telefonos:", d.telefonos, { textoValor2: true });
  filaPar("E_Mail", d.email, "Fecha:", fechaTexto, { correoValor1: true, negrillaValor2: true });
}

// Crea la hoja con: anchos, configuración de impresión, logo del SENA
// CENTRADO en la fila 1, y el bloque de datos de la empresa justo debajo
// (solo si se pasa datosEmpresa). Ya NO imprime título/subtítulo en negrita
// e itálica debajo del logo: eso se quitó a pedido. "pie" solo se usa en el
// pie de página de impresión, no se dibuja en la hoja. Devuelve { ws, n }.
function prepararHoja(wb, { nombre, anchos, logoId, logo, pie, datosEmpresa, fechaTexto }) {
  const ws = wb.addWorksheet(nombreHojaSeguro(nombre), { views: [{ showGridLines: true }] });
  ws.columns = anchos.map((width) => ({ width }));
  const n = anchos.length;

  ws.pageSetup = {
    paperSize: 1, // Carta
    orientation: n > 4 ? "landscape" : "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0, // alto automático: tantas páginas como hagan falta
    horizontalCentered: true,
    margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.7, header: 0.3, footer: 0.3 },
  };
  ws.headerFooter.oddFooter = `&L&8${escaparFooter(pie || "Matriz SENA")}&C&8Página &P de &N&R&8&D`;

  // Fila 1: logo del SENA centrado sobre todo el ancho imprimible.
  let altoFilaLogo = 20;
  let medidasLogo = null;
  if (logoId !== null && logoId !== undefined && logo) {
    const maxAncho = 190;
    const maxAlto = 44;
    const escala = Math.min(maxAncho / logo.ancho, maxAlto / logo.alto, 1);
    medidasLogo = { ancho: Math.round(logo.ancho * escala), alto: Math.round(logo.alto * escala) };
    altoFilaLogo = medidasLogo.alto + 8;
  }
  const filaLogo = ws.addRow([]);
  filaLogo.height = altoFilaLogo;
  if (medidasLogo) {
    ws.addImage(logoId, {
      tl: anchorCentrado(anchos, medidasLogo.ancho),
      ext: { width: medidasLogo.ancho, height: medidasLogo.alto },
      editAs: "oneCell",
    });
  }

  filaDatosEmpresa(ws, n, datosEmpresa, fechaTexto);

  return { ws, n };
}

function cerrarHoja(ws, n) {
  ws.pageSetup.printArea = `A1:${String.fromCharCode(64 + n)}${ws.lastRow.number}`;
}

function filaTexto(ws, texto, n, opciones = {}) {
  const f = ws.addRow([texto]);
  ws.mergeCells(f.number, 1, f.number, n);
  if (opciones.fuente) f.getCell(1).font = opciones.fuente;
  if (opciones.relleno) f.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: opciones.relleno } };
  if (opciones.wrap) {
    f.getCell(1).alignment = { wrapText: true, vertical: "middle" };
    if (opciones.alto) f.height = opciones.alto;
  }
  return f;
}

// ---------------------------------------------------------------------
// Hoja de un mes (Matriz 1: calificados / Matriz 2: no calificados)
// Estas hojas son las que se PRESENTAN al SENA: no llevan alertas internas
// ni la nota de cuota de aprendices. La Matriz 2 usa las categorías fijas
// del SENA (Conductor, Vigilante, Mensajero, Personal de Aseo y Cafetería,
// Empleados FIC, Contratos de Aprendizaje): siempre aparecen las 6, aunque
// tengan 0 personas ese mes.
// ---------------------------------------------------------------------
function construirHojaMes(wb, ctx, resultado) {
  const { nombreEmpresa, homologacion, logo, logoId, datosEmpresa } = ctx;
  const { etiqueta, jornadaSemanal, anio, mesIndex } = resultado;
  const titulo = `${nombreEmpresa ? nombreEmpresa + " — " : ""}Matriz SENA - ${etiquetaFechaMes(anio, mesIndex)}`;

  const grupos = agruparPorCargo(resultado.empleados, homologacion, jornadaSemanal, { forzarCategoriasFijas: true });
  const calificados = grupos.filter((g) => g.calificado).sort((a, b) => a.codigo.localeCompare(b.codigo));
  const noCalificados = grupos
    .filter((g) => !g.calificado)
    .sort((a, b) => ordenNoCalificado(a.nombreCargo) - ordenNoCalificado(b.nombreCargo) || a.nombreCargo.localeCompare(b.nombreCargo, "es"));

  const { ws, n } = prepararHoja(wb, {
    nombre: etiquetaCortaMes(anio, mesIndex),
    anchos: [30, 22, 28, 26],
    logo,
    logoId,
    pie: titulo,
    datosEmpresa,
    fechaTexto: etiquetaFechaMes(anio, mesIndex),
  });
  ws.addRow([]);

  const bloque = (tituloBloque, encabezadoPrimeraColumna, lista, etiquetaGrupo) => {
    const t = ws.addRow([tituloBloque]);
    t.getCell(1).font = { bold: true, size: 10 };

    const enc = ws.addRow([
      encabezadoPrimeraColumna,
      "NÚMERO DE TRABAJADORES",
      "JORNADA LABORAL SEMANAL POR TRABAJADOR",
      "TOTAL JORNADA LABORAL SEMANAL",
    ]);
    enc.height = 30;
    enc.eachCell(estiloEncabezado);

    const inicio = ws.lastRow.number + 1;
    let sumTrab = 0;
    let sumJornada = 0;
    let sumHoras = 0;

    lista.forEach((g) => {
      if (g.lineas.length === 0) {
        // Categoría fija sin personas este mes: se muestra en 0, no se omite.
        const f = ws.addRow([etiquetaGrupo(g), 0, jornadaSemanal, null]);
        const fila = f.number;
        f.getCell(4).value = { formula: `B${fila}*C${fila}`, result: 0 };
        f.eachCell((c, colNumber) => estiloDato(c, { izquierda: colNumber === 1 && !g.calificado }));
        f.getCell(1).numFmt = "@";
        sumJornada += jornadaSemanal;
        return;
      }
      g.lineas.forEach((l) => {
        // D = B × C como FÓRMULA de Excel, no como valor calculado (punto D.3),
        // para que quien revise pueda ver y auditar la cuenta en la hoja misma.
        const f = ws.addRow([etiquetaGrupo(g), l.trabajadores, l.jornada, null]);
        const fila = f.number;
        f.getCell(4).value = { formula: `B${fila}*C${fila}`, result: l.total };
        f.eachCell((c, colNumber) => estiloDato(c, { cursiva: l.parcial, izquierda: colNumber === 1 && !g.calificado }));
        f.getCell(1).numFmt = "@";
        sumTrab += l.trabajadores;
        sumJornada += l.jornada;
        sumHoras += l.total;
      });
    });

    if (lista.length === 0) {
      const f = ws.addRow(["(ninguno)", 0, 0, 0]);
      f.eachCell((c) => estiloDato(c));
    }
    const fin = ws.lastRow.number;

    const total = ws.addRow([
      "TOTAL",
      { formula: `SUM(B${inicio}:B${fin})`, result: sumTrab },
      { formula: `SUM(C${inicio}:C${fin})`, result: sumJornada },
      { formula: `SUM(D${inicio}:D${fin})`, result: Math.round(sumHoras) },
    ]);
    total.eachCell({ includeEmpty: true }, estiloTotal);
    ws.addRow([]);
  };

  bloque("Matriz 1: Oficios Calificados", "CÓDIGO DEL OFICIO SEGÚN LISTADO DE OFICIOS Y OCUPACIONES", calificados, (g) => g.codigo);
  bloque("Matriz 2: Oficios no Calificados", "NOMBRE DEL CARGO (ESPAÑOL)", noCalificados, (g) => g.nombreCargo);

  cerrarHoja(ws, n);
}

// ---------------------------------------------------------------------
// Plantilla promedio: grilla CÓDIGO × 6 MESES, con SUMA y PROMEDIO (÷ meses)
// de personas por fila (punto D.4). Si el formulario del SENA pide el
// promedio calculado sobre las planillas PILA bimestrales en vez de sobre
// estos 6 meses mensuales, hay que confirmarlo y ajustar aquí (ver punto F).
// No lleva el bloque de datos de la empresa: eso solo va en las hojas de mes.
// ---------------------------------------------------------------------
function construirHojaPlantillaPromedio(wb, ctx, resultados) {
  const { nombreEmpresa, homologacion, nombreDeCodigo } = ctx;
  const plantilla = calcularPlantillaPromedio(resultados, homologacion);
  const titulo = `${nombreEmpresa ? nombreEmpresa + " — " : ""}Plantilla promedio`;

  const anchosMeses = plantilla.etiquetasMes.map(() => 10);
  const anchos = [16, 40, ...anchosMeses, 12, 14];

  const { ws, n } = prepararHoja(wb, {
    nombre: "PLANTILLA PROM",
    anchos,
    pie: titulo,
  });
  ws.addRow([]);

  const encabezados = [
    "CÓDIGO CNO", "OCUPACIÓN / CARGO",
    ...plantilla.etiquetasMes,
    "SUMA (6 MESES)", "PROMEDIO (÷6)",
  ];
  const enc = ws.addRow(encabezados);
  enc.height = 30;
  enc.eachCell(estiloEncabezado);

  const inicio = ws.lastRow.number + 1;
  const colSuma = 3 + plantilla.etiquetasMes.length;
  const colProm = colSuma + 1;
  let sumaTotalPersonas = 0;

  plantilla.filas.forEach((f) => {
    const fila = ws.addRow([
      f.calificado ? f.codigo : "(sin código)",
      f.calificado ? nombreDeCodigo(f.codigo) || f.nombre : f.nombre,
      ...f.personasPorMes,
      null,
      null,
    ]);
    const filaNum = fila.number;
    const colInicioMeses = 3;
    const colFinMeses = 2 + plantilla.etiquetasMes.length;
    const letra = (col) => String.fromCharCode(64 + col);
    fila.getCell(colSuma).value = {
      formula: `SUM(${letra(colInicioMeses)}${filaNum}:${letra(colFinMeses)}${filaNum})`,
      result: f.sumaPersonas,
    };
    fila.getCell(colProm).value = {
      formula: `${letra(colSuma)}${filaNum}/${plantilla.meses}`,
      result: f.promedioPersonas,
    };
    fila.eachCell((c) => estiloDato(c));
    fila.getCell(1).numFmt = "@";
    fila.getCell(2).alignment = { horizontal: "left", vertical: "middle", wrapText: true };
    sumaTotalPersonas += f.sumaPersonas;
  });
  const fin = ws.lastRow.number;

  if (plantilla.filas.length > 0) {
    const letra = (col) => String.fromCharCode(64 + col);
    const filaTotal = ws.addRow([
      "TOTAL", "",
      ...plantilla.totalPorMes,
      { formula: `SUM(${letra(colSuma)}${inicio}:${letra(colSuma)}${fin})`, result: sumaTotalPersonas },
      { formula: `SUM(${letra(colProm)}${inicio}:${letra(colProm)}${fin})`, result: Math.round(plantilla.filas.reduce((a, f) => a + f.promedioPersonas, 0) * 100) / 100 },
    ]);
    filaTotal.eachCell({ includeEmpty: true }, estiloTotal);
  }

  cerrarHoja(ws, n);
}

// ---------------------------------------------------------------------
// API pública
// ---------------------------------------------------------------------
// resultados: array de resultados mensuales en orden cronológico
export async function exportarMatrizExcel({ nombreEmpresa, resultados, homologacion, baseDias, nombreDeCodigo, datosEmpresa }) {
  if (!resultados || resultados.length === 0) throw new Error("No hay resultados para exportar.");

  const wb = new ExcelJS.Workbook();
  wb.creator = "Solutions & Payroll";
  wb.created = new Date();

  // El logo se registra UNA vez en el libro; luego se inserta en cada hoja.
  let logo = null;
  let logoId = null;
  let advertenciaLogo = null;
  try {
    logo = await cargarLogo(URL_LOGO);
    logoId = wb.addImage({ base64: logo.base64, extension: logo.extension });
  } catch (e) {
    advertenciaLogo = `No se pudo insertar el logo: ${e.message}`;
  }

  const ctx = {
    nombreEmpresa: capitalizarNombreEmpresa(nombreEmpresa),
    homologacion,
    logo,
    logoId,
    baseDias,
    nombreDeCodigo: nombreDeCodigo || (() => ""),
    datosEmpresa,
  };

  resultados.forEach((r) => construirHojaMes(wb, ctx, r));
  construirHojaPlantillaPromedio(wb, ctx, resultados);

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const limpio = (t) => normalizar(t).toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_|_$/g, "");
  const prefijo = nombreEmpresa ? `${limpio(nombreEmpresa)}_` : "";
  const periodo = `${limpio(resultados[0].etiqueta)}_A_${limpio(resultados[resultados.length - 1].etiqueta)}`;

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${prefijo}MATRIZ_DETERMINACION_CUOTA_SENA_${periodo}.xlsx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);

  return { advertenciaLogo };
}