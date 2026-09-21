import ExcelJS from "exceljs";

import {
  agruparPorCargo,
  calcularPlantillaPromedio,
  normalizar,
} from "./sena";

// ============================================================
// CONFIGURACIÓN
// ============================================================

export const URL_LOGO = "/image.png";

const GRIS = "FFD9D9D9";
const NEGRO = { argb: "FF000000" };

const BORDE = {
  top: { style: "thin", color: NEGRO },
  bottom: { style: "thin", color: NEGRO },
  left: { style: "thin", color: NEGRO },
  right: { style: "thin", color: NEGRO },
};

// ============================================================
// LOGO
// ============================================================

export async function cargarLogo(url = URL_LOGO) {
  const resp = await fetch(url);

  if (!resp.ok) {
    throw new Error(
      "No se pudo leer " + url + " (HTTP " + resp.status + ")."
    );
  }

  const blob = await resp.blob();

  if (!blob.type.startsWith("image/")) {
    throw new Error(
      url + " no es una imagen. Verifica que el archivo exista en la carpeta /public."
    );
  }

  let extension = "png";

  if (/jpe?g/i.test(blob.type)) {
    extension = "jpeg";
  } else if (/gif/i.test(blob.type)) {
    extension = "gif";
  }

  if (!/png|jpe?g|gif/i.test(blob.type)) {
    throw new Error(
      "Formato de logo no soportado (" + blob.type + "). Usa PNG o JPG."
    );
  }

  const base64 = await new Promise((resolve, reject) => {
    const lector = new FileReader();

    lector.onloadend = () => resolve(lector.result);
    lector.onerror = reject;

    lector.readAsDataURL(blob);
  });

  const dims = await new Promise((resolve, reject) => {
    const img = new Image();

    img.onload = () => {
      resolve({
        ancho: img.naturalWidth,
        alto: img.naturalHeight,
      });
    };

    img.onerror = reject;
    img.src = base64;
  });

  return {
    base64,
    extension,
    ...dims,
  };
}

// ============================================================
// UTILIDADES
// ============================================================

const nombreHojaSeguro = (texto) => {
  return String(texto)
    .replace(/[\\/*?:[\]]/g, "")
    .slice(0, 31);
};

const escaparFooter = (texto) => {
  return String(texto).replace(/&/g, "&&");
};

// ============================================================
// ESTILOS
// ============================================================

const estiloEncabezado = (celda) => {
  celda.font = {
    bold: true,
    size: 8,
    color: NEGRO,
  };

  celda.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: {
      argb: GRIS,
    },
  };

  celda.alignment = {
    horizontal: "center",
    vertical: "middle",
    wrapText: true,
  };

  celda.border = BORDE;
};

const estiloDato = (
  celda,
  { cursiva = false, izquierda = false } = {}
) => {
  celda.font = {
    size: 9,
    italic: cursiva,
    color: NEGRO,
  };

  celda.alignment = {
    horizontal: izquierda ? "left" : "center",
    vertical: "middle",
    wrapText: true,
  };

  celda.border = BORDE;
};

const estiloTotal = (celda) => {
  celda.font = {
    bold: true,
    size: 9,
    color: NEGRO,
  };

  celda.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: {
      argb: GRIS,
    },
  };

  celda.alignment = {
    horizontal: "center",
    vertical: "middle",
    wrapText: true,
  };

  celda.border = BORDE;
};

// ============================================================
// ESTILO TIPO PLANTILLA SENA
// ============================================================

const crearEncabezadoSena = (
  ws,
  numeroFila,
  textos,
  { alto = 35 } = {}
) => {
  const fila = ws.getRow(numeroFila);

  textos.forEach((texto, indice) => {
    const celda = fila.getCell(indice + 1);

    celda.value = texto;

    estiloEncabezado(celda);
  });

  fila.height = alto;
};

// ============================================================
// PREPARAR HOJA
// ============================================================

function prepararHoja(
  wb,
  {
    nombre,
    anchos,
    logoId,
    logo,
    titulo,
    subtitulo,
    nota,
    pie,
  }
) {
  const ws = wb.addWorksheet(nombreHojaSeguro(nombre), {
    views: [
      {
        showGridLines: false,
      },
    ],
  });

  ws.columns = anchos.map((width) => ({
    width,
  }));

  const n = anchos.length;

  ws.pageSetup = {
    paperSize: 1,
    orientation: n > 4 ? "landscape" : "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    horizontalCentered: true,

    margins: {
      left: 0.35,
      right: 0.35,
      top: 0.45,
      bottom: 0.55,
      header: 0.2,
      footer: 0.25,
    },
  };

  ws.headerFooter.oddFooter =
    "&L&8" + escaparFooter(pie || "Matriz SENA") +
    "&C&8Página &P de &N" +
    "&R&8&D";

  // ==========================================================
  // ENCABEZADO
  // ==========================================================

  const fila1 = ws.addRow([
    "SERVICIO NACIONAL DE APRENDIZAJE - SENA",
  ]);

  const fila2 = ws.addRow([
    "SISTEMA DE GESTIÓN VIRTUAL DE APRENDICES - SGVA",
  ]);

  const fila3 = ws.addRow([
    "MODELO DE PRESENTACIÓN DE INFORMACIÓN PARA CUOTA DE APRENDICES",
  ]);

  const fila4 = ws.addRow([
    titulo,
  ]);

  const fila5 = ws.addRow([
    subtitulo || "",
  ]);

  const fila6 = ws.addRow([
    nota || "",
  ]);

  [fila1, fila2, fila3, fila4, fila5, fila6].forEach(
    (fila) => {
      ws.mergeCells(
        fila.number,
        1,
        fila.number,
        n
      );

      const celda = fila.getCell(1);

      celda.alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };

      celda.border = BORDE;
    }
  );

  fila1.height = 24;
  fila2.height = 20;
  fila3.height = 24;
  fila4.height = 24;
  fila5.height = 20;
  fila6.height = 20;

  fila1.getCell(1).font = {
    bold: true,
    size: 13,
  };

  fila2.getCell(1).font = {
    bold: true,
    size: 10,
  };

  fila3.getCell(1).font = {
    bold: true,
    size: 10,
  };

  fila4.getCell(1).font = {
    bold: true,
    size: 12,
  };

  fila5.getCell(1).font = {
    italic: true,
    size: 9,
  };

  fila6.getCell(1).font = {
    italic: true,
    size: 8,
    color: {
      argb: "FF666666",
    },
  };

  // ==========================================================
  // LOGO
  // ==========================================================

  if (
    logoId !== null &&
    logoId !== undefined &&
    logo
  ) {
    const anchoColPx =
      anchos[n - 1] * 7 + 5;

    const maxAncho = Math.min(
      Math.max(anchoColPx - 12, 40),
      180
    );

    const maxAlto = 48;

    const escala = Math.min(
      maxAncho / logo.ancho,
      maxAlto / logo.alto
    );

    ws.addImage(logoId, {
      tl: {
        col: Math.max(0, n - 2),
        row: 0.15,
      },

      ext: {
        width: Math.round(
          logo.ancho * escala
        ),
        height: Math.round(
          logo.alto * escala
        ),
      },

      editAs: "oneCell",
    });
  }

  return {
    ws,
    n,
  };
}

// ============================================================
// CERRAR HOJA
// ============================================================

function obtenerLetraColumna(numero) {
  let resultado = "";
  let n = numero;

  while (n > 0) {
    const residuo = (n - 1) % 26;

    resultado =
      String.fromCharCode(65 + residuo) +
      resultado;

    n = Math.floor((n - 1) / 26);
  }

  return resultado;
}

function cerrarHoja(ws, n) {
  const ultimaFila = ws.lastRow
    ? ws.lastRow.number
    : 1;

  const letraFinal =
    obtenerLetraColumna(n);

  ws.pageSetup.printArea =
    "A1:" + letraFinal + ultimaFila;
}

// ============================================================
// FILA DE INFORMACIÓN
// ============================================================

function filaTexto(
  ws,
  texto,
  n,
  opciones = {}
) {
  const fila = ws.addRow([]);

  ws.mergeCells(
    fila.number,
    1,
    fila.number,
    n
  );

  const celda = fila.getCell(1);

  celda.value = texto;

  if (opciones.fuente) {
    celda.font = opciones.fuente;
  }

  if (opciones.relleno) {
    celda.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: {
        argb: opciones.relleno,
      },
    };
  }

  celda.alignment = {
    vertical: "middle",
    horizontal:
      opciones.horizontal || "left",
    wrapText: true,
  };

  celda.border = BORDE;

  if (opciones.alto) {
    fila.height = opciones.alto;
  }

  return fila;
}

// ============================================================
// MATRIZ MENSUAL
// ============================================================

function construirHojaMes(
  wb,
  ctx,
  resultado
) {
  const {
    nombreEmpresa,
    homologacion,
    logo,
    logoId,
    baseDias,
  } = ctx;

  const {
    etiqueta,
    jornadaSemanal,
  } = resultado;

  const titulo =
    (nombreEmpresa ? nombreEmpresa + " — " : "") +
    "MATRIZ DE DETERMINACIÓN DE CUOTA SENA - " + etiqueta;

  const { ws, n } = prepararHoja(
    wb,
    {
      nombre: "MATRIZ " + etiqueta,

      anchos: [
        34,
        18,
        25,
        25,
      ],

      logo,
      logoId,

      titulo,

      subtitulo:
        "Jornada laboral semanal aplicada: " + jornadaSemanal + " horas",

      nota:
        "Proporción calculada sobre " +
        (baseDias === 30 ? "30 días" : "los días reales del mes") +
        ".",

      pie: titulo,
    }
  );

  // ==========================================================
  // EMPRESA
  // ==========================================================

  filaTexto(
    ws,
    "EMPRESA: " + (nombreEmpresa || "No especificada"),
    n,
    {
      fuente: {
        bold: true,
        size: 9,
      },

      relleno: GRIS,
      horizontal: "left",
      alto: 20,
    }
  );

  // ==========================================================
  // CRUCE
  // ==========================================================

  filaTexto(
    ws,
    resultado.cuadra
      ? "VERIFICACIÓN: Nómina (" + resultado.totalNomina + ") coincide con Seguridad Social (" + resultado.totalSegSocial + ")"
      : "ALERTA: Nómina (" + resultado.totalNomina + ") NO coincide con Seguridad Social (" + resultado.totalSegSocial + ")",
    n,
    {
      fuente: {
        bold: true,
        size: 9,
        color: {
          argb: resultado.cuadra
            ? "FF006100"
            : "FF9C0006",
        },
      },

      relleno: resultado.cuadra
        ? "FFC6EFCE"
        : "FFFFC7CE",

      alto: 24,
    }
  );

  // ==========================================================
  // CUOTA
  // ==========================================================

  filaTexto(
    ws,
    "CUOTA DE APRENDICES REQUERIDA: " + resultado.cuotaAprendicesRequerida,
    n,
    {
      fuente: {
        bold: true,
        size: 9,
      },

      relleno: GRIS,
      alto: 22,
    }
  );

  ws.addRow([]);

  // ==========================================================
  // AGRUPACIÓN
  // ==========================================================

  const grupos = agruparPorCargo(
    resultado.empleados,
    homologacion,
    jornadaSemanal
  );

  const calificados = grupos
    .filter((g) => g.calificado)
    .sort((a, b) =>
      String(a.codigo).localeCompare(
        String(b.codigo)
      )
    );

  const noCalificados = grupos
    .filter((g) => !g.calificado)
    .sort((a, b) =>
      String(a.nombreCargo).localeCompare(
        String(b.nombreCargo),
        "es"
      )
    );

  // ==========================================================
  // BLOQUES
  // ==========================================================

  const bloque = (
    tituloBloque,
    encabezadoPrimeraColumna,
    lista,
    etiquetaGrupo
  ) => {
    const tituloFila = ws.addRow([]);

    ws.mergeCells(
      tituloFila.number,
      1,
      tituloFila.number,
      n
    );

    tituloFila.getCell(1).value =
      tituloBloque;

    tituloFila.getCell(1).font = {
      bold: true,
      size: 11,
    };

    tituloFila.getCell(1).alignment = {
      horizontal: "left",
      vertical: "middle",
      wrapText: true,
    };

    tituloFila.getCell(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: {
        argb: GRIS,
      },
    };

    tituloFila.getCell(1).border =
      BORDE;

    tituloFila.height = 22;

    // ========================================================
    // ENCABEZADOS
    // ========================================================

    const numeroEncabezado =
      ws.lastRow.number + 1;

    crearEncabezadoSena(
      ws,
      numeroEncabezado,
      [
        encabezadoPrimeraColumna,
        "NÚMERO DE\nTRABAJADORES",
        "JORNADA LABORAL\nSEMANAL POR TRABAJADOR",
        "TOTAL JORNADA\nLABORAL SEMANAL",
      ],
      {
        alto: 42,
      }
    );

    const inicio =
      ws.lastRow.number + 1;

    let sumTrab = 0;
    let sumHoras = 0;

    // ========================================================
    // DATOS
    // ========================================================

    lista.forEach((g) => {
      g.lineas.forEach((l) => {
        const fila = ws.addRow([
          etiquetaGrupo(g),
          l.trabajadores,
          l.jornada,
          l.total,
        ]);

        fila.eachCell(
          (celda, numeroColumna) => {
            estiloDato(celda, {
              cursiva: l.parcial,
              izquierda:
                numeroColumna === 1 &&
                !g.calificado,
            });
          }
        );

        fila.getCell(1).numFmt = "@";

        sumTrab += Number(
          l.trabajadores || 0
        );

        sumHoras += Number(
          l.total || 0
        );
      });
    });

    if (lista.length === 0) {
      const fila = ws.addRow([
        "(ninguno)",
        0,
        0,
        0,
      ]);

      fila.eachCell((celda) =>
        estiloDato(celda)
      );
    }

    const fin =
      ws.lastRow.number;

    // ========================================================
    // TOTAL
    // ========================================================

    const total = ws.addRow([
      "TOTAL",
      {
        formula: "SUM(B" + inicio + ":B" + fin + ")",
        result: sumTrab,
      },
      "",
      {
        formula: "SUM(D" + inicio + ":D" + fin + ")",
        result:
          Math.round(
            sumHoras * 10
          ) / 10,
      },
    ]);

    total.eachCell(
      {
        includeEmpty: true,
      },
      estiloTotal
    );

    total.height = 22;

    ws.addRow([]);
  };

  // ==========================================================
  // MATRIZ 1
  // ==========================================================

  bloque(
    "MATRIZ 1 - OFICIOS CALIFICADOS",
    "CÓDIGO DEL OFICIO SEGÚN LISTADO DE OFICIOS Y OCUPACIONES",
    calificados,
    (g) => g.codigo
  );

  // ==========================================================
  // MATRIZ 2
  // ==========================================================

  bloque(
    "MATRIZ 2 - OFICIOS NO CALIFICADOS",
    "NOMBRE DEL CARGO",
    noCalificados,
    (g) => g.nombreCargo
  );

  // ==========================================================
  // OBSERVACIÓN
  // ==========================================================

  filaTexto(
    ws,
    "OBSERVACIÓN: La información presentada corresponde a los trabajadores identificados en los archivos de nómina y Seguridad Social suministrados para el periodo.",
    n,
    {
      fuente: {
        italic: true,
        size: 8,
      },

      alto: 30,
    }
  );

  cerrarHoja(ws, n);
}

// ============================================================
// PLANTILLA PROMEDIO
// ============================================================

function construirHojaPlantillaPromedio(
  wb,
  ctx,
  resultados
) {
  const {
    nombreEmpresa,
    homologacion,
    logo,
    logoId,
    nombreDeCodigo,
  } = ctx;

  const plantilla =
    calcularPlantillaPromedio(
      resultados,
      homologacion
    );

  const titulo =
    (nombreEmpresa ? nombreEmpresa + " — " : "") +
    "PLANTILLA PROMEDIO";

  const { ws, n } =
    prepararHoja(wb, {
      nombre:
        "PLANTILLA PROMEDIO",

      anchos: [
        22,
        40,
        24,
        28,
      ],

      logo,
      logoId,

      titulo,

      subtitulo:
        "Promedio calculado sobre " + plantilla.meses + " mes(es) cargado(s)",

      nota:
        plantilla.meses < 6
          ? "ATENCIÓN: solo se cargaron " + plantilla.meses + " de 6 meses."
          : "Información consolidada de los periodos analizados.",

      pie: titulo,
    });

  filaTexto(
    ws,
    "EMPRESA: " + (nombreEmpresa || "No especificada"),
    n,
    {
      fuente: {
        bold: true,
        size: 9,
      },

      relleno: GRIS,
      alto: 20,
    }
  );

  ws.addRow([]);

  crearEncabezadoSena(
    ws,
    ws.lastRow.number + 1,
    [
      "CÓDIGO CNO",
      "OCUPACIÓN / CARGO",
      "PROMEDIO TRABAJADORES (" + plantilla.meses + " MESES)",
      "PROMEDIO JORNADA SEMANAL TOTAL (H)",
    ],
    {
      alto: 42,
    }
  );

  const inicio =
    ws.lastRow.number + 1;

  let sumP = 0;
  let sumH = 0;

  plantilla.filas.forEach((f) => {
    const fila = ws.addRow([
      f.calificado
        ? f.codigo
        : "(sin código)",

      f.calificado
        ? nombreDeCodigo(f.codigo) ||
          f.nombre
        : f.nombre,

      f.promedioPersonas,

      f.promedioHoras,
    ]);

    fila.eachCell((celda) =>
      estiloDato(celda)
    );

    fila.getCell(1).numFmt = "@";

    fila.getCell(2).alignment = {
      horizontal: "left",
      vertical: "middle",
      wrapText: true,
    };

    sumP += Number(
      f.promedioPersonas || 0
    );

    sumH += Number(
      f.promedioHoras || 0
    );
  });

  const fin =
    ws.lastRow.number;

  if (plantilla.filas.length > 0) {
    const total = ws.addRow([
      "TOTAL",
      "",
      {
        formula: "SUM(C" + inicio + ":C" + fin + ")",
        result:
          Math.round(
            sumP * 100
          ) / 100,
      },
      {
        formula: "SUM(D" + inicio + ":D" + fin + ")",
        result:
          Math.round(
            sumH * 100
          ) / 100,
      },
    ]);

    total.eachCell(
      {
        includeEmpty: true,
      },
      estiloTotal
    );
  }

  cerrarHoja(ws, n);
}

// ============================================================
// CRUCE SEGURIDAD SOCIAL
// ============================================================

function construirHojaCruce(
  wb,
  ctx,
  resultados
) {
  const {
    nombreEmpresa,
    logo,
    logoId,
  } = ctx;

  const titulo =
    (nombreEmpresa ? nombreEmpresa + " — " : "") +
    "CRUCE CON SEGURIDAD SOCIAL";

  const { ws, n } =
    prepararHoja(wb, {
      nombre:
        "CRUCE SEG. SOCIAL",

      anchos: [
        16,
        12,
        14,
        12,
        16,
        70,
      ],

      logo,
      logoId,

      titulo,

      subtitulo:
        "Trabajadores en nómina frente a trabajadores identificados en Seguridad Social.",

      pie: titulo,
    });

  ws.addRow([]);

  crearEncabezadoSena(
    ws,
    ws.lastRow.number + 1,
    [
      "MES",
      "NÓMINA",
      "SEG. SOCIAL",
      "DIFERENCIA",
      "ESTADO",
      "DETALLE",
    ],
    {
      alto: 35,
    }
  );

  const lista = (
    arr,
    max = 12
  ) => {
    if (!arr || arr.length === 0) {
      return "";
    }

    return arr.length > max
      ? arr.slice(0, max).join(", ") + " … (+" + (arr.length - max) + ")"
      : arr.join(", ");
  };

  resultados.forEach((r) => {
    const detalle = [];

    if (
      r.sinMatchSS &&
      r.sinMatchSS.length
    ) {
      detalle.push(
        "En nómina sin Seg. Social: " + lista(
          r.sinMatchSS.map(
            (e) =>
              e.documentoOriginal ||
              e.documento
          )
        )
      );
    }

    if (
      r.soloEnSS &&
      r.soloEnSS.length
    ) {
      detalle.push(
        "En Seg. Social sin nómina: " + lista(
          r.soloEnSS.map(
            (e) => e.documento
          )
        )
      );
    }

    if (
      r.duplicados &&
      r.duplicados.length
    ) {
      detalle.push(
        "Documentos repetidos en nómina (se contaron una vez): " + lista(
          r.duplicados
        )
      );
    }

    const fila = ws.addRow([
      r.etiqueta,
      r.totalNomina,
      r.totalSegSocial,
      r.totalNomina -
        r.totalSegSocial,
      r.cuadra
        ? "CUADRA"
        : "REVISAR",
      detalle.join("\n") || "—",
    ]);

    fila.eachCell((celda) =>
      estiloDato(celda)
    );

    fila.getCell(6).alignment = {
      horizontal: "left",
      vertical: "top",
      wrapText: true,
    };

    fila.getCell(5).font = {
      bold: true,
      size: 9,
      color: {
        argb: r.cuadra
          ? "FF006100"
          : "FF9C0006",
      },
    };
  });

  cerrarHoja(ws, n);
}

// ============================================================
// HOMOLOGACIÓN
// ============================================================

function construirHojaHomologacion(
  wb,
  ctx
) {
  const {
    nombreEmpresa,
    homologacion,
    logo,
    logoId,
  } = ctx;

  const titulo =
    (nombreEmpresa ? nombreEmpresa + " — " : "") +
    "HOMOLOGACIÓN DE CARGOS";

  const { ws, n } =
    prepararHoja(wb, {
      nombre:
        "HOMOLOGACIÓN",

      anchos: [
        34,
        34,
        12,
        44,
        14,
      ],

      logo,
      logoId,

      titulo,

      subtitulo:
        "Cargo original → cargo en español → código del Listado de Oficios y Ocupaciones.",

      pie: titulo,
    });

  ws.addRow([]);

  crearEncabezadoSena(
    ws,
    ws.lastRow.number + 1,
    [
      "CARGO ORIGINAL",
      "CARGO (ESPAÑOL)",
      "CÓDIGO CNO",
      "OCUPACIÓN EN EL LISTADO",
      "CONFIRMADO",
    ],
    {
      alto: 32,
    }
  );

  const datosHomologacion =
    homologacion || {};

  Object.values(datosHomologacion)
    .sort((a, b) =>
      String(a.es || "").localeCompare(
        String(b.es || ""),
        "es"
      )
    )
    .forEach((h) => {
      const fila = ws.addRow([
        h.original || "",
        h.es || "",
        h.codigo || "(sin código)",
        h.ocupacion || "",
        h.confirmado
          ? "Sí"
          : "No",
      ]);

      fila.eachCell((celda) =>
        estiloDato(celda, {
          izquierda: true,
        })
      );

      fila.getCell(3).alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      fila.getCell(3).numFmt = "@";

      fila.getCell(5).alignment = {
        horizontal: "center",
        vertical: "middle",
      };
    });

  cerrarHoja(ws, n);
}

// ============================================================
// API PÚBLICA
// ============================================================

export async function exportarMatrizExcel({
  nombreEmpresa,
  resultados,
  homologacion,
  baseDias,
  nombreDeCodigo,
}) {
  if (
    !resultados ||
    resultados.length === 0
  ) {
    throw new Error(
      "No hay resultados para exportar."
    );
  }

  const wb =
    new ExcelJS.Workbook();

  wb.creator =
    "Solutions & Payroll";

  wb.created = new Date();

  // ==========================================================
  // LOGO
  // ==========================================================

  let logo = null;
  let logoId = null;
  let advertenciaLogo = null;

  try {
    logo = await cargarLogo(
      URL_LOGO
    );

    logoId = wb.addImage({
      base64: logo.base64,
      extension: logo.extension,
    });
  } catch (e) {
    advertenciaLogo =
      "No se pudo insertar el logo: " + (
        e instanceof Error
          ? e.message
          : String(e)
      );
  }

  // ==========================================================
  // CONTEXTO
  // ==========================================================

  const ctx = {
    nombreEmpresa,
    homologacion:
      homologacion || {},
    logo,
    logoId,
    baseDias,

    nombreDeCodigo:
      nombreDeCodigo ||
      (() => ""),
  };

  // ==========================================================
  // HOJAS MENSUALES
  // ==========================================================

  resultados.forEach(
    (resultado) => {
      construirHojaMes(
        wb,
        ctx,
        resultado
      );
    }
  );

  // ==========================================================
  // PLANTILLA PROMEDIO
  // ==========================================================

  construirHojaPlantillaPromedio(
    wb,
    ctx,
    resultados
  );

  // ==========================================================
  // CRUCE SEGURIDAD SOCIAL
  // ==========================================================

  construirHojaCruce(
    wb,
    ctx,
    resultados
  );

  // ==========================================================
  // HOMOLOGACIÓN
  // ==========================================================

  construirHojaHomologacion(
    wb,
    ctx
  );

  // ==========================================================
  // GENERAR ARCHIVO
  // ==========================================================

  const buffer =
    await wb.xlsx.writeBuffer();

  const blob = new Blob(
    [buffer],
    {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }
  );

  // ==========================================================
  // NOMBRE DEL ARCHIVO
  // ==========================================================

  const limpio = (texto) => {
    return normalizar(texto)
      .toUpperCase()
      .replace(/[^A-Z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
  };

  const prefijo =
    nombreEmpresa
      ? limpio(nombreEmpresa) + "_"
      : "";

  const primerPeriodo =
    resultados[0]?.etiqueta ||
    "INICIO";

  const ultimoPeriodo =
    resultados[
      resultados.length - 1
    ]?.etiqueta ||
    "FIN";

  const periodo =
    limpio(primerPeriodo) + "_A_" + limpio(ultimoPeriodo);

  const nombreArchivo =
    prefijo + "MATRIZ_DETERMINACION_CUOTA_SENA_" + periodo + ".xlsx";

  // ==========================================================
  // DESCARGA
  // ==========================================================

  const url =
    URL.createObjectURL(blob);

  const enlace =
    document.createElement("a");

  enlace.href = url;
  enlace.download =
    nombreArchivo;

  document.body.appendChild(
    enlace
  );

  enlace.click();

  enlace.remove();

  setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 1000);

  return {
    advertenciaLogo,
  };
}