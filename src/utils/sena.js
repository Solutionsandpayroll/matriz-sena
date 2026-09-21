// =====================================================================
// Lógica pura de la Matriz SENA
// Sin React, sin XLSX, sin ExcelJS.
// Todo lo que está aquí se puede probar con Node / Vitest.
// =====================================================================

export const NOMBRES_MES = [
  "Enero",
  "Febrero",
  "Marzo",
  "Abril",
  "Mayo",
  "Junio",
  "Julio",
  "Agosto",
  "Septiembre",
  "Octubre",
  "Noviembre",
  "Diciembre",
];

// Meses en los que el SENA suele pedir la presentación.
export const MESES_PERIODO_SENA = [0, 2, 6, 8];

// ---------------------------------------------------------------------
// Texto
// ---------------------------------------------------------------------

export function normalizar(texto) {
  return String(texto ?? "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

// "DESC. OFICIO" -> "desc oficio"
// "N° ID" -> "n id"
export function normalizarEncabezado(texto) {
  return normalizar(texto)
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

// Clave estable de un cargo
export function claveCargo(cargo) {
  const limpio = String(cargo ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();

  return limpio || "(SIN CARGO)";
}

// ---------------------------------------------------------------------
// Documentos
// ---------------------------------------------------------------------

export function normalizarDocumento(valor) {
  if (valor === null || valor === undefined) return "";

  const s = String(valor).trim();

  if (!s) return "";

  // Pasaportes / documentos alfanuméricos
  if (/[a-z]/i.test(s)) {
    return s.toUpperCase().replace(/[^A-Z0-9]/g, "");
  }

  const digitos = s.replace(/\D/g, "");

  if (!digitos) return "";

  return digitos.replace(/^0+(?=\d)/, "");
}

// ---------------------------------------------------------------------
// Fechas
// ---------------------------------------------------------------------

function fechaValida(anio, mes, dia) {
  const d = new Date(anio, mes - 1, dia);

  return d.getFullYear() === anio &&
    d.getMonth() === mes - 1 &&
    d.getDate() === dia
    ? d
    : null;
}

// Acepta:
// Date
// Excel serial
// YYYY-MM-DD
// DD/MM/YYYY
// DD-MM-YYYY
// DD.MM.YYYY
export function parseFecha(valor) {
  if (valor === null || valor === undefined || valor === "") {
    return null;
  }

  if (valor instanceof Date) {
    return isNaN(valor)
      ? null
      : new Date(
          valor.getFullYear(),
          valor.getMonth(),
          valor.getDate()
        );
  }

  if (typeof valor === "number") {
    if (!isFinite(valor) || valor < 20000 || valor > 80000) {
      return null;
    }

    const d = new Date(
      Math.round((valor - 25569) * 86400000)
    );

    return new Date(
      d.getUTCFullYear(),
      d.getUTCMonth(),
      d.getUTCDate()
    );
  }

  const s = String(valor).trim();

  if (!s) return null;

  if (/^\d{5}(\.\d+)?$/.test(s)) {
    return parseFecha(Number(s));
  }

  let m = s.match(
    /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/
  );

  if (m) {
    return fechaValida(
      +m[1],
      +m[2],
      +m[3]
    );
  }

  m = s.match(
    /^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})/
  );

  if (m) {
    return fechaValida(
      +m[3],
      +m[2],
      +m[1]
    );
  }

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
  return new Date(
    anio,
    mesIndex + 1,
    0
  ).getDate();
}

function diffDias(a, b) {
  const ua = Date.UTC(
    a.getFullYear(),
    a.getMonth(),
    a.getDate()
  );

  const ub = Date.UTC(
    b.getFullYear(),
    b.getMonth(),
    b.getDate()
  );

  return Math.round(
    (ub - ua) / 86400000
  );
}

// ---------------------------------------------------------------------
// Proporción por fechas
// ---------------------------------------------------------------------

export function calcularProporcionPorFechas(
  fechaIngreso,
  fechaRetiro,
  anio,
  mesIndex,
  baseDias = "real"
) {
  const totalDias = diasDelMes(
    anio,
    mesIndex
  );

  const inicioMes = new Date(
    anio,
    mesIndex,
    1
  );

  const finMes = new Date(
    anio,
    mesIndex,
    totalDias
  );

  const ingreso = parseFecha(fechaIngreso);
  const retiro = parseFecha(fechaRetiro);

  const inicio =
    ingreso && ingreso > inicioMes
      ? ingreso
      : inicioMes;

  const fin =
    retiro && retiro < finMes
      ? retiro
      : finMes;

  if (fin < inicio) return 0;

  const dias =
    diffDias(inicio, fin) + 1;

  if (dias >= totalDias) return 1;

  const base =
    Number(baseDias) === 30
      ? 30
      : totalDias;

  return Math.min(
    dias / base,
    1
  );
}

// ---------------------------------------------------------------------
// Proporción utilizando horas de Seguridad Social
// ---------------------------------------------------------------------

export function calcularProporcionPorHoras(
  horasLaboradas,
  jornadaSemanal,
  diasMes = 30
) {
  const horas = Number(horasLaboradas);
  const jornada = Number(jornadaSemanal);

  if (
    !Number.isFinite(horas) ||
    !Number.isFinite(jornada) ||
    horas < 0 ||
    jornada <= 0
  ) {
    return null;
  }

  /*
   * Una jornada semanal se convierte a una referencia mensual
   * aproximada usando 4.333 semanas por mes.
   *
   * Esto NO reemplaza el criterio que defina finalmente el SENA.
   * Sirve para cruzar la información de horas de SS con la
   * proporción calculada por fechas.
   */
  const semanasMes =
    Number(diasMes) / 7;

  const horasOrdinariasMes =
    jornada * semanasMes;

  if (horasOrdinariasMes <= 0) {
    return null;
  }

  return Math.min(
    horas / horasOrdinariasMes,
    1
  );
}

// ---------------------------------------------------------------------
// Comparación de proporciones
// ---------------------------------------------------------------------

export function compararProporciones(
  proporcionFechas,
  proporcionHoras,
  tolerancia = 0.02
) {
  if (
    proporcionHoras === null ||
    proporcionHoras === undefined
  ) {
    return {
      disponible: false,
      difiere: false,
      diferencia: null,
    };
  }

  const fechas =
    Number(proporcionFechas) || 0;

  const horas =
    Number(proporcionHoras) || 0;

  const diferencia =
    Math.abs(fechas - horas);

  return {
    disponible: true,
    difiere: diferencia > tolerancia,
    diferencia,
    proporcionFechas: fechas,
    proporcionHoras: horas,
  };
}

// ---------------------------------------------------------------------
// Novedades
// ---------------------------------------------------------------------

export function describirNovedad(
  emp,
  anio,
  mesIndex
) {
  const inicioMes = new Date(
    anio,
    mesIndex,
    1
  );

  const finMes = new Date(
    anio,
    mesIndex,
    diasDelMes(
      anio,
      mesIndex
    )
  );

  const ingreso = parseFecha(
    emp.fechaIngreso
  );

  const retiro = parseFecha(
    emp.fechaRetiro
  );

  const partes = [];

  if (
    ingreso &&
    ingreso > inicioMes &&
    ingreso <= finMes
  ) {
    partes.push(
      `Ingresó el ${formatFecha(ingreso)}`
    );
  }

  if (
    retiro &&
    retiro >= inicioMes &&
    retiro < finMes
  ) {
    partes.push(
      `Se retiró el ${formatFecha(retiro)}`
    );
  }

  return partes.join(" · ");
}

// ---------------------------------------------------------------------
// Jornada
// ---------------------------------------------------------------------

/*
 * Jornada máxima legal semanal como sugerencia.
 *
 * IMPORTANTE:
 * Para julio de 2026 existe un cambio dentro del mismo mes:
 *
 * 01/07/2026 - 14/07/2026 -> 44 horas
 * Desde 15/07/2026       -> 42 horas
 *
 * Por eso NO debemos considerar julio completo como 42 horas.
 */

export function jornadaLegalSugerida(
  anio,
  mesIndex
) {
  const fecha = new Date(
    anio,
    mesIndex,
    15
  );

  if (
    fecha < new Date(2023, 6, 15)
  ) {
    return 48;
  }

  if (
    fecha < new Date(2024, 6, 15)
  ) {
    return 47;
  }

  if (
    fecha < new Date(2025, 6, 15)
  ) {
    return 46;
  }

  if (
    fecha < new Date(2026, 6, 15)
  ) {
    return 44;
  }

  return 42;
}

// Jornada aplicable a un día concreto.
export function jornadaLegalParaFecha(
  fecha
) {
  const d = parseFecha(fecha);

  if (!d) return 42;

  if (
    d < new Date(2023, 6, 15)
  ) {
    return 48;
  }

  if (
    d < new Date(2024, 6, 15)
  ) {
    return 47;
  }

  if (
    d < new Date(2025, 6, 15)
  ) {
    return 46;
  }

  if (
    d < new Date(2026, 6, 15)
  ) {
    return 44;
  }

  return 42;
}

// Información detallada de jornada del mes.
export function obtenerJornadaMes(
  anio,
  mesIndex,
  jornadaEmpresa = "auto"
) {
  const dias = diasDelMes(
    anio,
    mesIndex
  );

  if (
    jornadaEmpresa !== "auto" &&
    Number(jornadaEmpresa) > 0
  ) {
    const jornada = Number(
      jornadaEmpresa
    );

    return {
      jornadaSemanal: jornada,
      jornadaInicial: jornada,
      jornadaFinal: jornada,
      diasPrimeraJornada: dias,
      diasSegundaJornada: 0,
      tieneCambio: false,
      fechaCambio: null,
      descripcion: `${jornada} horas semanales`,
    };
  }

  // Julio 2026: cambio de 44 a 42 horas el día 15.
  if (
    anio === 2026 &&
    mesIndex === 6
  ) {
    return {
      jornadaSemanal: 42,
      jornadaInicial: 44,
      jornadaFinal: 42,
      diasPrimeraJornada: 14,
      diasSegundaJornada: dias - 14,
      tieneCambio: true,
      fechaCambio: new Date(
        2026,
        6,
        15
      ),
      descripcion:
        "44 horas del 01/07/2026 al 14/07/2026 y 42 horas desde el 15/07/2026",
    };
  }

  const jornada =
    jornadaLegalSugerida(
      anio,
      mesIndex
    );

  return {
    jornadaSemanal: jornada,
    jornadaInicial: jornada,
    jornadaFinal: jornada,
    diasPrimeraJornada: dias,
    diasSegundaJornada: 0,
    tieneCambio: false,
    fechaCambio: null,
    descripcion: `${jornada} horas semanales`,
  };
}

// ---------------------------------------------------------------------
// Cuota de aprendices
// ---------------------------------------------------------------------

export function calcularCuotaAprendices(
  totalTrabajadores
) {
  const n = Math.floor(
    Number(totalTrabajadores) || 0
  );

  if (n < 15) return 0;

  return Math.max(
    1,
    Math.floor(n / 20) +
      (n % 20 >= 10 ? 1 : 0)
  );
}

// ---------------------------------------------------------------------
// Detección de columnas
// ---------------------------------------------------------------------

function buscarColumna(
  encabezados,
  aliases,
  excluir = []
) {
  // 1. Coincidencia exacta.
  for (const alias of aliases) {
    const i = encabezados.findIndex(
      (h) => h === alias
    );

    if (i !== -1) return i;
  }

  // 2. Contiene el alias.
  for (const alias of aliases) {
    const i = encabezados.findIndex(
      (h) =>
        h &&
        h.includes(alias) &&
        !excluir.some(
          (x) => h.includes(x)
        )
    );

    if (i !== -1) return i;
  }

  return -1;
}

// ---------------------------------------------------------------------
// Alias Nómina
// ---------------------------------------------------------------------

const ALIAS_NOMINA = {
  documento: {
    aliases: [
      "documento identidad",
      "documento de identidad",
      "numero de documento",
      "numero documento",
      "documento",
      "cedula",
      "identificacion",
      "no id",
      "nro id",
    ],
    excluir: [
      "tipo",
      "expedicion",
    ],
  },

  cargo: {
    aliases: [
      "desc oficio",
      "descripcion oficio",
      "descripcion del oficio",
      "cargo",
      "oficio",
      "puesto",
      "job title",
      "position",
      "ocupacion",
    ],
    excluir: [
      "cod",
      "codigo",
      "centro",
      "tipo",
      "nivel",
    ],
  },

  nombre: {
    aliases: [
      "nombre completo",
      "nombres y apellidos",
      "apellidos y nombres",
      "nombre",
      "empleado",
    ],
    excluir: [
      "cargo",
      "oficio",
      "banco",
      "eps",
    ],
  },

  ingreso: {
    aliases: [
      "fecha ingreso",
      "fecha de ingreso",
      "fecha inicio contrato",
      "ingreso",
    ],
    excluir: [
      "salario",
      "base",
      "valor",
      "total",
    ],
  },

  antiguedad: {
    aliases: [
      "fecha antiguedad",
    ],
    excluir: [],
  },

  retiro: {
    aliases: [
      "fecha retiro",
      "fecha de retiro",
      "fecha terminacion",
      "fecha de terminacion",
      "fecha fin contrato",
      "retiro",
    ],
    excluir: [
      "valor",
      "aporte",
      "motivo",
      "causa",
      "indemnizacion",
    ],
  },
};

// ---------------------------------------------------------------------
// Parsear Nómina
// ---------------------------------------------------------------------

export function parsearNomina(
  filas
) {
  const limite = Math.min(
    filas.length,
    40
  );

  let indiceEncabezado = -1;
  let col = null;

  for (
    let i = 0;
    i < limite;
    i++
  ) {
    const fila = filas[i];

    if (!Array.isArray(fila)) {
      continue;
    }

    const enc = fila.map(
      normalizarEncabezado
    );

    const doc = buscarColumna(
      enc,
      ALIAS_NOMINA.documento.aliases,
      ALIAS_NOMINA.documento.excluir
    );

    const cargo = buscarColumna(
      enc,
      ALIAS_NOMINA.cargo.aliases,
      ALIAS_NOMINA.cargo.excluir
    );

    if (
      doc !== -1 &&
      cargo !== -1 &&
      doc !== cargo
    ) {
      indiceEncabezado = i;

      col = {
        documento: doc,
        cargo,
        nombre: buscarColumna(
          enc,
          ALIAS_NOMINA.nombre.aliases,
          ALIAS_NOMINA.nombre.excluir
        ),
        ingreso: buscarColumna(
          enc,
          ALIAS_NOMINA.ingreso.aliases,
          ALIAS_NOMINA.ingreso.excluir
        ),
        antiguedad: buscarColumna(
          enc,
          ALIAS_NOMINA.antiguedad.aliases
        ),
        retiro: buscarColumna(
          enc,
          ALIAS_NOMINA.retiro.aliases,
          ALIAS_NOMINA.retiro.excluir
        ),
      };

      break;
    }
  }

  if (
    indiceEncabezado === -1
  ) {
    throw new Error(
      "No se encontraron las columnas de documento y cargo (ej. 'DOCUMENTO IDENTIDAD' y 'DESC. OFICIO') en las primeras 40 filas."
    );
  }

  const empleados = [];

  for (
    const fila of filas.slice(
      indiceEncabezado + 1
    )
  ) {
    if (!Array.isArray(fila)) {
      continue;
    }

    const documentoOriginal =
      fila[col.documento];

    const documento =
      normalizarDocumento(
        documentoOriginal
      );

    if (!documento) continue;

    const cargoTexto =
      String(
        fila[col.cargo] ?? ""
      ).trim();

    const ingreso =
      col.ingreso !== -1
        ? parseFecha(
            fila[col.ingreso]
          )
        : null;

    const antiguedad =
      col.antiguedad !== -1
        ? parseFecha(
            fila[col.antiguedad]
          )
        : null;

    const retiro =
      col.retiro !== -1
        ? parseFecha(
            fila[col.retiro]
          )
        : null;

    empleados.push({
      documento,

      documentoOriginal:
        String(
          documentoOriginal
        ).trim(),

      nombre:
        col.nombre !== -1
          ? String(
              fila[col.nombre] ?? ""
            ).trim()
          : "",

      cargo:
        cargoTexto ||
        "(SIN CARGO)",

      cargoKey:
        claveCargo(cargoTexto),

      fechaIngreso:
        ingreso || antiguedad,

      fechaRetiro:
        retiro,
    });
  }

  return {
    empleados,
    columnasDetectadas: col,
  };
}

// =====================================================================
// SEGURIDAD SOCIAL
// =====================================================================

const ALIAS_DOC_SS_EXACTOS = [
  "no id",
  "n id",
  "nro id",
  "numero de identificacion",
  "numero identificacion",
  "no identificacion",
  "numero de documento",
  "numero documento",
  "identificacion",
  "documento",
  "cedula",
];

const ALIAS_SS = {
  documento: ALIAS_DOC_SS_EXACTOS,

  ingreso: [
    "fecha ingreso",
    "fecha de ingreso",
    "fecha inicio",
    "fecha inicio contrato",
    "inicio contrato",
    "ingreso",
  ],

  retiro: [
    "fecha retiro",
    "fecha de retiro",
    "fecha terminacion",
    "fecha de terminacion",
    "fecha fin",
    "fecha fin contrato",
    "retiro",
    "terminacion",
  ],

  horas: [
    "horas laboradas",
    "horas trabajadas",
    "horas laborales",
    "horas",
    "hrs laboradas",
    "hrs trabajadas",
    "numero horas",
    "total horas",
  ],

  nombre: [
    "nombre",
    "nombre completo",
    "nombres y apellidos",
    "apellidos y nombres",
    "trabajador",
    "empleado",
  ],

  cargo: [
    "cargo",
    "oficio",
    "ocupacion",
    "descripcion oficio",
    "desc oficio",
  ],
};

// ---------------------------------------------------------------------
// Detectar si una fila es subtotal/resumen
// ---------------------------------------------------------------------

function esFilaSubtotalSeguridadSocial(
  fila
) {
  if (!Array.isArray(fila)) {
    return true;
  }

  const texto = fila
    .map((v) =>
      normalizar(v)
    )
    .join(" ");

  if (!texto.trim()) {
    return true;
  }

  const palabrasBloqueantes = [
    "subtotal",
    "total ciudad",
    "total por ciudad",
    "total municipio",
    "total departamento",
    "total general",
    "gran total",
    "totales",
    "resumen",
    "consolidado",
  ];

  return palabrasBloqueantes.some(
    (palabra) =>
      texto.includes(palabra)
  );
}

// ---------------------------------------------------------------------
// Número seguro
// ---------------------------------------------------------------------

function parseNumero(valor) {
  if (
    valor === null ||
    valor === undefined ||
    valor === ""
  ) {
    return null;
  }

  if (typeof valor === "number") {
    return Number.isFinite(valor)
      ? valor
      : null;
  }

  const limpio = String(valor)
    .trim()
    .replace(/\s/g, "")
    .replace(",", ".");

  if (!limpio) return null;

  const numero = Number(limpio);

  return Number.isFinite(numero)
    ? numero
    : null;
}

// ---------------------------------------------------------------------
// Parsear Seguridad Social
//
// IMPORTANTE:
// Ahora NO devuelve únicamente un Set.
// Devuelve:
//
// {
//   documentos: Set,
//   registros: Map,
//   empleados: [...],
//   columnasDetectadas: {...},
//   filasIgnoradas: [...]
// }
//
// Se conserva "documentos" para no romper temporalmente el App actual.
// ---------------------------------------------------------------------

export function parsearSeguridadSocial(
  filas
) {
  const limite = Math.min(
    filas.length,
    60
  );

  let indiceEncabezado = -1;

  let col = {
    documento: -1,
    ingreso: -1,
    retiro: -1,
    horas: -1,
    nombre: -1,
    cargo: -1,
  };

  // ---------------------------------------------------------------
  // Buscar encabezado
  // ---------------------------------------------------------------

  for (
    let i = 0;
    i < limite;
    i++
  ) {
    const fila = filas[i];

    if (!Array.isArray(fila)) {
      continue;
    }

    const enc = fila.map(
      normalizarEncabezado
    );

    const documento =
      buscarColumna(
        enc,
        ALIAS_SS.documento,
        ["tipo"]
      );

    if (documento === -1) {
      continue;
    }

    indiceEncabezado = i;

    col = {
      documento,

      ingreso:
        buscarColumna(
          enc,
          ALIAS_SS.ingreso,
          []
        ),

      retiro:
        buscarColumna(
          enc,
          ALIAS_SS.retiro,
          []
        ),

      horas:
        buscarColumna(
          enc,
          ALIAS_SS.horas,
          []
        ),

      nombre:
        buscarColumna(
          enc,
          ALIAS_SS.nombre,
          []
        ),

      cargo:
        buscarColumna(
          enc,
          ALIAS_SS.cargo,
          []
        ),
    };

    break;
  }

  if (
    indiceEncabezado === -1
  ) {
    throw new Error(
      "No se encontró una columna de identificación reconocible en la planilla de Seguridad Social."
    );
  }

  // ---------------------------------------------------------------
  // Estructuras de resultado
  // ---------------------------------------------------------------

  const documentos =
    new Set();

  const registros =
    new Map();

  const empleados = [];

  const filasIgnoradas = [];

  // ---------------------------------------------------------------
  // Recorrer filas
  // ---------------------------------------------------------------

  for (
    let numeroFila = indiceEncabezado + 1;
    numeroFila < filas.length;
    numeroFila++
  ) {
    const fila =
      filas[numeroFila];

    if (
      esFilaSubtotalSeguridadSocial(
        fila
      )
    ) {
      filasIgnoradas.push({
        numeroFila:
          numeroFila + 1,
        motivo:
          "Subtotal / resumen / fila vacía",
        fila,
      });

      continue;
    }

    if (!Array.isArray(fila)) {
      continue;
    }

    const documento =
      normalizarDocumento(
        fila[col.documento]
      );

    // Una fila sin documento no representa una persona.
    if (!documento) {
      filasIgnoradas.push({
        numeroFila:
          numeroFila + 1,
        motivo:
          "Sin documento válido",
        fila,
      });

      continue;
    }

    const fechaIngreso =
      col.ingreso !== -1
        ? parseFecha(
            fila[col.ingreso]
          )
        : null;

    const fechaRetiro =
      col.retiro !== -1
        ? parseFecha(
            fila[col.retiro]
          )
        : null;

    const horasLaboradas =
      col.horas !== -1
        ? parseNumero(
            fila[col.horas]
          )
        : null;

    const nombre =
      col.nombre !== -1
        ? String(
            fila[col.nombre] ?? ""
          ).trim()
        : "";

    const cargo =
      col.cargo !== -1
        ? String(
            fila[col.cargo] ?? ""
          ).trim()
        : "";

    const registro = {
      documento,

      nombre,

      cargo:
        cargo || "(SIN CARGO)",

      cargoKey:
        claveCargo(cargo),

      fechaIngreso,

      fechaRetiro,

      horasLaboradas,

      filaOrigen:
        numeroFila + 1,

      documentoOriginal:
        String(
          fila[col.documento] ?? ""
        ).trim(),
    };

    documentos.add(
      documento
    );

    // -------------------------------------------------------------
    // Si aparece más de una vez el mismo documento,
    // consolidamos la información.
    // -------------------------------------------------------------

    const existente =
      registros.get(
        documento
      );

    if (!existente) {
      registros.set(
        documento,
        registro
      );
    } else {
      const combinado = {
        ...existente,

        nombre:
          existente.nombre ||
          registro.nombre,

        cargo:
          existente.cargo !==
            "(SIN CARGO)"
            ? existente.cargo
            : registro.cargo,

        cargoKey:
          existente.cargo !==
            "(SIN CARGO)"
            ? existente.cargoKey
            : registro.cargoKey,

        fechaIngreso:
          existente.fechaIngreso ||
          registro.fechaIngreso,

        fechaRetiro:
          registro.fechaRetiro ||
          existente.fechaRetiro,

        horasLaboradas:
          registro.horasLaboradas ??
          existente.horasLaboradas,
      };

      registros.set(
        documento,
        combinado
      );
    }
  }

  // ---------------------------------------------------------------
  // Convertir Map a arreglo
  // ---------------------------------------------------------------

  for (const registro of registros.values()) {
    empleados.push({
      documento:
        registro.documento,

      documentoOriginal:
        registro.documentoOriginal,

      nombre:
        registro.nombre,

      cargo:
        registro.cargo,

      cargoKey:
        registro.cargoKey,

      fechaIngreso:
        registro.fechaIngreso,

      fechaRetiro:
        registro.fechaRetiro,

      horasLaboradas:
        registro.horasLaboradas,

      filaOrigen:
        registro.filaOrigen,

      // Alias útiles para el resto de la aplicación.
      ingreso:
        registro.fechaIngreso,

      retiro:
        registro.fechaRetiro,

      horas:
        registro.horasLaboradas,
    });
  }

  return {
    documentos,

    registros,

    empleados,

    filasIgnoradas,

    columnasDetectadas: col,

    indiceEncabezado,
  };
}

// =====================================================================
// CRUCE NÓMINA / SEGURIDAD SOCIAL
// =====================================================================

export function cruzarNominaSeguridadSocial(
  empleadosNomina = [],
  seguridadSocial = {}
) {
  const registrosSS =
    seguridadSocial.registros instanceof Map
      ? seguridadSocial.registros
      : new Map();

  const documentosSS =
    seguridadSocial.documentos instanceof Set
      ? seguridadSocial.documentos
      : new Set();

  const nominaPorDocumento =
    new Map();

  for (const empleado of empleadosNomina) {
    const documento =
      normalizarDocumento(
        empleado.documento
      );

    if (!documento) continue;

    nominaPorDocumento.set(
      documento,
      empleado
    );
  }

  const presentesEnAmbas = [];
  const soloNomina = [];
  const soloSeguridadSocial = [];

  // ---------------------------------------------------------------
  // Nómina -> SS
  // ---------------------------------------------------------------

  for (const empleado of empleadosNomina) {
    const documento =
      normalizarDocumento(
        empleado.documento
      );

    if (!documento) continue;

    const ss =
      registrosSS.get(
        documento
      );

    if (ss) {
      presentesEnAmbas.push({
        ...empleado,
        seguridadSocial: ss,
        enSegSocial: true,

        // La fecha de retiro debe venir de SS.
        fechaRetiro:
          ss.fechaRetiro ||
          empleado.fechaRetiro ||
          null,

        horasLaboradasSS:
          ss.horasLaboradas ??
          null,
      });
    } else {
      soloNomina.push({
        ...empleado,
        enSegSocial: false,
        fechaRetiro:
          empleado.fechaRetiro ||
          null,
      });
    }
  }

  // ---------------------------------------------------------------
  // SS -> Nómina
  // ---------------------------------------------------------------

  for (
    const documento of documentosSS
  ) {
    if (
      !nominaPorDocumento.has(
        documento
      )
    ) {
      const ss =
        registrosSS.get(
          documento
        );

      if (!ss) continue;

      soloSeguridadSocial.push({
        ...ss,

        enNomina: false,

        // Esta persona debe ser buscada
        // en las otras nóminas.
        requiereBusquedaCargo:
          !ss.cargo ||
          ss.cargo ===
            "(SIN CARGO)",
      });
    }
  }

  return {
    presentesEnAmbas,

    soloNomina,

    soloSeguridadSocial,

    totalNomina:
      empleadosNomina.length,

    totalSegSocial:
      documentosSS.size,

    cuadra:
      empleadosNomina.length ===
      documentosSS.size,
  };
}

// =====================================================================
// BUSCAR CARGO DE RETIRADOS EN OTROS MESES
// =====================================================================

export function buscarCargoEnOtrosMeses(
  documento,
  resultadosMeses = []
) {
  const doc =
    normalizarDocumento(
      documento
    );

  if (!doc) return null;

  for (
    const resultado of resultadosMeses
  ) {
    const empleados =
      resultado?.empleados || [];

    const encontrado =
      empleados.find(
        (empleado) =>
          normalizarDocumento(
            empleado.documento
          ) === doc &&
          empleado.cargo &&
          empleado.cargo !==
            "(SIN CARGO)"
      );

    if (encontrado) {
      return {
        documento: doc,

        cargo:
          encontrado.cargo,

        cargoKey:
          encontrado.cargoKey ||
          claveCargo(
            encontrado.cargo
          ),

        nombre:
          encontrado.nombre || "",

        fechaIngreso:
          encontrado.fechaIngreso ||
          null,

        fechaRetiro:
          encontrado.fechaRetiro ||
          null,

        encontradoEnMes:
          resultado.mes ??
          resultado.nombreMes ??
          null,
      };
    }
  }

  return null;
}

// ---------------------------------------------------------------------
// Completar retirados usando otras nóminas
// ---------------------------------------------------------------------

export function completarRetiradosConOtrosMeses(
  empleados,
  resultadosMeses = []
) {
  const completados = [];
  const pendientes = [];

  for (
    const empleado of empleados || []
  ) {
    if (
      empleado.cargo &&
      empleado.cargo !==
        "(SIN CARGO)"
    ) {
      completados.push({
        ...empleado,
        cargoResuelto:
          "nomina_actual",
      });

      continue;
    }

    const encontrado =
      buscarCargoEnOtrosMeses(
        empleado.documento,
        resultadosMeses
      );

    if (encontrado) {
      completados.push({
        ...empleado,

        nombre:
          empleado.nombre ||
          encontrado.nombre,

        cargo:
          encontrado.cargo,

        cargoKey:
          encontrado.cargoKey,

        cargoResuelto:
          "otra_nomina",

        mesCargoOrigen:
          encontrado.encontradoEnMes,
      });
    } else {
      const pendiente = {
        ...empleado,

        cargo:
          empleado.cargo ||
          "(SIN CARGO)",

        cargoKey:
          empleado.cargoKey ||
          claveCargo(
            empleado.cargo
          ),

        cargoResuelto:
          "pendiente_manual",

        requiereCorreccionManual:
          true,
      };

      completados.push(
        pendiente
      );

      pendientes.push(
        pendiente
      );
    }
  }

  return {
    empleados: completados,
    pendientes,
  };
}

// =====================================================================
// CREAR EMPLEADOS DEL MES A PARTIR DE NÓMINA + SS
// =====================================================================

export function integrarNominaConSeguridadSocial({
  nomina,
  seguridadSocial,
  anio,
  mesIndex,
  baseDias = "real",
  jornadaSemanal = 42,
  resultadosMesesAnteriores = [],
} = {}) {
  const empleadosNomina =
    nomina?.empleados || [];

  const ss =
    seguridadSocial || {};

  const registrosSS =
    ss.registros instanceof Map
      ? ss.registros
      : new Map();

  const empleados = [];

  // ---------------------------------------------------------------
  // 1. Personas de nómina
  // ---------------------------------------------------------------

  for (
    const empleadoNomina of empleadosNomina
  ) {
    const documento =
      normalizarDocumento(
        empleadoNomina.documento
      );

    if (!documento) continue;

    const registroSS =
      registrosSS.get(
        documento
      );

    // Fecha de retiro viene preferiblemente de SS.
    const fechaRetiro =
      registroSS?.fechaRetiro ||
      empleadoNomina.fechaRetiro ||
      null;

    const fechaIngreso =
      empleadoNomina.fechaIngreso ||
      registroSS?.fechaIngreso ||
      null;

    const proporcionFechas =
      calcularProporcionPorFechas(
        fechaIngreso,
        fechaRetiro,
        anio,
        mesIndex,
        baseDias
      );

    const horasLaboradas =
      registroSS?.horasLaboradas ??
      null;

    const proporcionHoras =
      horasLaboradas !== null
        ? calcularProporcionPorHoras(
            horasLaboradas,
            jornadaSemanal,
            diasDelMes(
              anio,
              mesIndex
            )
          )
        : null;

    const comparacion =
      compararProporciones(
        proporcionFechas,
        proporcionHoras
      );

    empleados.push({
      ...empleadoNomina,

      documento,

      fechaIngreso,

      // SIEMPRE preferir SS para retiro.
      fechaRetiro,

      horasLaboradasSS:
        horasLaboradas,

      proporcion:
        proporcionFechas,

      proporcionFechas,

      proporcionHoras,

      diferenciaProporciones:
        comparacion.diferencia,

      alertaHoras:
        comparacion.difiere,

      enSegSocial:
        Boolean(
          registroSS
        ),

      novedad:
        describirNovedad(
          {
            fechaIngreso,
            fechaRetiro,
          },
          anio,
          mesIndex
        ),
    });
  }

  // ---------------------------------------------------------------
  // 2. Personas que están en SS pero no en nómina.
  //
  // Estas son especialmente importantes porque pueden ser
  // personas retiradas durante el mes.
  // ---------------------------------------------------------------

  const documentosNomina =
    new Set(
      empleadosNomina.map(
        (e) =>
          normalizarDocumento(
            e.documento
          )
      )
    );

  const retiradosNoEnNomina = [];

  for (
    const [documento, registroSS]
      of registrosSS.entries()
  ) {
    if (
      documentosNomina.has(
        documento
      )
    ) {
      continue;
    }

    const fechaIngreso =
      registroSS.fechaIngreso ||
      null;

    const fechaRetiro =
      registroSS.fechaRetiro ||
      null;

    const proporcionFechas =
      calcularProporcionPorFechas(
        fechaIngreso,
        fechaRetiro,
        anio,
        mesIndex,
        baseDias
      );

    const horasLaboradas =
      registroSS.horasLaboradas ??
      null;

    const proporcionHoras =
      horasLaboradas !== null
        ? calcularProporcionPorHoras(
            horasLaboradas,
            jornadaSemanal,
            diasDelMes(
              anio,
              mesIndex
            )
          )
        : null;

    const comparacion =
      compararProporciones(
        proporcionFechas,
        proporcionHoras
      );

    // Buscar cargo en las nóminas anteriores.
    const cargoAnterior =
      buscarCargoEnOtrosMeses(
        documento,
        resultadosMesesAnteriores
      );

    const cargo =
      cargoAnterior?.cargo ||
      registroSS.cargo ||
      "(SIN CARGO)";

    const cargoKey =
      cargoAnterior?.cargoKey ||
      registroSS.cargoKey ||
      claveCargo(cargo);

    const empleadoRetirado = {
      documento,

      documentoOriginal:
        registroSS.documentoOriginal ||
        documento,

      nombre:
        registroSS.nombre ||
        cargoAnterior?.nombre ||
        "",

      cargo,

      cargoKey,

      fechaIngreso,

      fechaRetiro,

      horasLaboradasSS:
        horasLaboradas,

      proporcion:
        proporcionFechas,

      proporcionFechas,

      proporcionHoras,

      diferenciaProporciones:
        comparacion.diferencia,

      alertaHoras:
        comparacion.difiere,

      enSegSocial: true,

      enNomina: false,

      esRetiradoNoEnNomina:
        true,

      cargoResuelto:
        cargoAnterior
          ? "otra_nomina"
          : registroSS.cargo &&
            registroSS.cargo !==
              "(SIN CARGO)"
          ? "seguridad_social"
          : "pendiente_manual",

      requiereCorreccionManual:
        !cargoAnterior &&
        (!registroSS.cargo ||
          registroSS.cargo ===
            "(SIN CARGO)"),

      novedad:
        describirNovedad(
          {
            fechaIngreso,
            fechaRetiro,
          },
          anio,
          mesIndex
        ),
    };

    // Solo lo incorporamos si realmente tiene proporción.
    if (
      empleadoRetirado.proporcion >
      0
    ) {
      empleados.push(
        empleadoRetirado
      );
    }

    retiradosNoEnNomina.push(
      empleadoRetirado
    );
  }

  // ---------------------------------------------------------------
  // 3. Pendientes manuales
  // ---------------------------------------------------------------

  const pendientesManual =
    empleados.filter(
      (e) =>
        e.requiereCorreccionManual
    );

  // ---------------------------------------------------------------
  // 4. Totales
  // ---------------------------------------------------------------

  const totalNomina =
    empleadosNomina.length;

  const totalSegSocial =
    registrosSS.size;

  const documentosNominaSet =
    new Set(
      empleadosNomina.map(
        (e) =>
          normalizarDocumento(
            e.documento
          )
      )
    );

  const soloEnNomina =
    empleadosNomina.filter(
      (e) =>
        !registrosSS.has(
          normalizarDocumento(
            e.documento
          )
        )
    );

  const soloEnSS =
    retiradosNoEnNomina;

  return {
    empleados,

    retiradosNoEnNomina,

    pendientesManual,

    soloEnNomina,

    soloEnSS,

    totalNomina,

    totalSegSocial,

    cuadra:
      totalNomina ===
      totalSegSocial,

    diferenciaConteo:
      totalSegSocial -
      totalNomina,

    documentosNomina:
      documentosNominaSet,

    documentosSegSocial:
      new Set(
        registrosSS.keys()
      ),
  };
}

// =====================================================================
// HOMOLOGACIÓN DE CARGOS -> CNO
// =====================================================================

const STOPWORDS = new Set([
  "de",
  "del",
  "la",
  "el",
  "los",
  "las",
  "y",
  "e",
  "en",
  "al",
  "para",
  "por",
  "con",
  "un",
  "una",
  "jr",
  "sr",
  "junior",
  "senior",
  "ii",
  "iii",
  "iv",
  "nivel",
  "especialista",
]);

function tokenizar(texto) {
  return normalizar(texto)
    .split(/[^a-z0-9]+/)
    .filter(
      (t) =>
        t.length > 2 &&
        !STOPWORDS.has(t)
    );
}

// "auxiliares" -> "auxiliar"
function raiz(token) {
  return token.length > 4
    ? token.replace(
        /(es|s)$/,
        ""
      )
    : token;
}

// ---------------------------------------------------------------------
// Índice CNO
// ---------------------------------------------------------------------

export function indexarListadoCno(
  listado
) {
  return (listado || []).map(
    (item) => {
      const norm =
        normalizar(
          item.ocupacion
        );

      return {
        ...item,

        _norm: norm,

        _palabras:
          norm
            .split(
              /[^a-z0-9]+/
            )
            .filter(Boolean),
      };
    }
  );
}

// ---------------------------------------------------------------------
// Sugerencias CNO
// ---------------------------------------------------------------------

export function sugerirCno(
  texto,
  indice,
  limite = 8
) {
  const consulta =
    String(texto ?? "").trim();

  if (!consulta) return [];

  // Búsqueda por código.
  if (
    /^\d{2,4}$/.test(
      consulta
    )
  ) {
    return indice
      .filter((it) =>
        String(
          it.codigo
        ).startsWith(
          consulta
        )
      )
      .slice(0, limite)
      .map((it) => ({
        ...it,
        puntos: 0,
        cobertura: 0,
      }));
  }

  const tokens = [
    ...new Set(
      tokenizar(
        consulta
      ).map(raiz)
    ),
  ];

  if (
    tokens.length === 0
  ) {
    return [];
  }

  const normConsulta =
    normalizar(
      consulta
    );

  const resultados = [];

  for (
    const it of indice
  ) {
    let aciertos = 0;

    for (
      const t of tokens
    ) {
      if (
        it._palabras.some(
          (p) =>
            p.startsWith(t)
        )
      ) {
        aciertos++;
      }
    }

    if (aciertos === 0) {
      continue;
    }

    const cobertura =
      aciertos /
      tokens.length;

    const precision =
      aciertos /
      Math.max(
        it._palabras.length,
        1
      );

    let puntos =
      cobertura * 10 +
      precision * 2;

    if (
      it._norm ===
      normConsulta
    ) {
      puntos += 100;
    }

    resultados.push({
      codigo:
        it.codigo,

      ocupacion:
        it.ocupacion,

      puntos,

      cobertura,
    });
  }

  return resultados
    .sort(
      (a, b) =>
        b.puntos -
          a.puntos ||
        a.ocupacion.length -
          b.ocupacion.length
    )
    .slice(0, limite);
}

// ---------------------------------------------------------------------
// Nombre del código
// ---------------------------------------------------------------------

export function nombreDeCodigo(
  codigo,
  indice
) {
  return (
    indice.find(
      (it) =>
        String(it.codigo) ===
        String(codigo)
    )?.ocupacion || ""
  );
}

// =====================================================================
// MAPEO SUGERIDO
// =====================================================================
//
// IMPORTANTE:
// Estos códigos siguen siendo sugerencias.
// NO se consideran confirmados automáticamente.
//
// No se debe usar esta tabla como verdad definitiva hasta comparar
// contra el listado CNO vigente y el modelo oficial.
// =====================================================================

export const MAPEO_SUGERIDO = {
  "IT SUPPORT SPECIALIST":
    "2281",

  "SALESFORCE JUNIOR DEVELOPER":
    "2173",

  "ADMINISTRATIVE ANALYST":
    "1122",

  "CREDIT AND COLLECTIONS ANALYST II":
    "1232",

  "SALES & PRICING TEAM LEADER":
    "6211",
};

// ---------------------------------------------------------------------
// Estado de homologación
// ---------------------------------------------------------------------

export const ESTADOS_HOMOLOGACION = {
  CONFIRMADO: "confirmado",
  SUGERIDO: "sugerido",
  BAJA_COBERTURA: "baja_cobertura",
  NO_ENCONTRADO: "no_encontrado",
  NO_CALIFICADO: "no_calificado",
};

// Cobertura mínima para considerar una sugerencia
// suficientemente buena para mostrarla como sugerencia normal.
export const COBERTURA_MINIMA_CNO = 0.5;

// Cobertura alta para poder sugerir automáticamente.
// Aun así NO queda confirmada.
export const COBERTURA_ALTA_CNO = 0.75;

// ---------------------------------------------------------------------
// Determinar estado de una sugerencia
// ---------------------------------------------------------------------

export function determinarEstadoHomologacion(
  resultado
) {
  if (
    !resultado ||
    !resultado.codigo
  ) {
    return {
      estado:
        ESTADOS_HOMOLOGACION.NO_ENCONTRADO,

      confirmado: false,
    };
  }

  const cobertura =
    Number(
      resultado.cobertura
    ) || 0;

  if (
    cobertura <
    COBERTURA_MINIMA_CNO
  ) {
    return {
      estado:
        ESTADOS_HOMOLOGACION.BAJA_COBERTURA,

      confirmado: false,
    };
  }

  return {
    estado:
      ESTADOS_HOMOLOGACION.SUGERIDO,

    confirmado: false,
  };
}

// ---------------------------------------------------------------------
// Construir homologación
// ---------------------------------------------------------------------

export function construirHomologacion(
  cargos,
  traducciones,
  previo,
  indice
) {
  const resultado = {};

  for (
    const [key, original]
      of cargos
  ) {
    const p =
      previo?.[key];

    // ---------------------------------------------------------------
    // Homologación previamente guardada
    // ---------------------------------------------------------------

    if (
      p &&
      (
        p.es ||
        p.codigo !==
          undefined
      )
    ) {
      let estado =
        p.estado;

      if (
        p.confirmado
      ) {
        estado =
          ESTADOS_HOMOLOGACION.CONFIRMADO;
      }

      resultado[key] = {
        original,

        es:
          p.es ||
          traducciones?.[key] ||
          original,

        codigo:
          p.codigo ?? "",

        ocupacion:
          p.ocupacion ||
          (
            p.codigo
              ? nombreDeCodigo(
                  p.codigo,
                  indice
                )
              : ""
          ),

        confirmado:
          Boolean(
            p.confirmado
          ),

        estado:
          estado ||
          (
            p.codigo
              ? ESTADOS_HOMOLOGACION.SUGERIDO
              : ESTADOS_HOMOLOGACION.NO_ENCONTRADO
          ),

        cobertura:
          p.cobertura ??
          null,
      };

      continue;
    }

    const es =
      (
        traducciones?.[key] ||
        ""
      ).trim() ||
      original;

    // ---------------------------------------------------------------
    // Mapeo directo
    //
    // Antes se trataba casi como una asignación automática.
    // Ahora sigue siendo únicamente una SUGERENCIA.
    // ---------------------------------------------------------------

    const directo =
      MAPEO_SUGERIDO[key];

    const itemDirecto =
      directo
        ? indice.find(
            (it) =>
              String(
                it.codigo
              ) ===
              String(
                directo
              )
          )
        : null;

    if (itemDirecto) {
      resultado[key] = {
        original,

        es,

        codigo:
          String(
            directo
          ),

        ocupacion:
          nombreDeCodigo(
            directo,
            indice
          ),

        confirmado: false,

        estado:
          ESTADOS_HOMOLOGACION.SUGERIDO,

        cobertura: null,

        origenSugerencia:
          "MAPEO_SUGERIDO",
      };

      continue;
    }

    // ---------------------------------------------------------------
    // Buscar coincidencias
    // ---------------------------------------------------------------

    const sugerencias =
      sugerirCno(
        es,
        indice,
        8
      );

    const top =
      sugerencias[0];

    const estado =
      determinarEstadoHomologacion(
        top
      );

    // ---------------------------------------------------------------
    // IMPORTANTE:
    // Si la cobertura es baja, NO asignamos código.
    // ---------------------------------------------------------------

    const codigoSeguro =
      estado.estado ===
      ESTADOS_HOMOLOGACION.BAJA_COBERTURA
        ? ""
        : top?.codigo || "";

    const ocupacionSegura =
      codigoSeguro
        ? top?.ocupacion ||
          ""
        : "";

    resultado[key] = {
      original,

      es,

      codigo:
        codigoSeguro,

      ocupacion:
        ocupacionSegura,

      confirmado: false,

      estado:
        estado.estado,

      cobertura:
        top?.cobertura ??
        null,

      puntos:
        top?.puntos ??
        0,

      sugerencias:
        sugerencias.map(
          (s) => ({
            codigo:
              s.codigo,

            ocupacion:
              s.ocupacion,

            cobertura:
              s.cobertura,

            puntos:
              s.puntos,
          })
        ),
    };
  }

  return resultado;
}

// =====================================================================
// Persistencia de homologación
// =====================================================================

const PREFIJO_STORAGE =
  "sena-homologacion:v2:";

const claveStorage = (
  empresa
) =>
  PREFIJO_STORAGE +
  (
    normalizar(
      empresa
    ).trim() ||
    "_sin_empresa"
  );

// ---------------------------------------------------------------------
// Cargar
// ---------------------------------------------------------------------

export function cargarHomologacionGuardada(
  empresa
) {
  try {
    const raw =
      localStorage.getItem(
        claveStorage(
          empresa
        )
      );

    return raw
      ? JSON.parse(raw)
      : {};
  } catch {
    return {};
  }
}

// ---------------------------------------------------------------------
// Guardar
//
// Solo guarda homologaciones confirmadas.
// ---------------------------------------------------------------------

export function guardarHomologacion(
  empresa,
  homologacion
) {
  try {
    if (
      !String(
        empresa ?? ""
      ).trim()
    ) {
      return {
        guardado: false,
        motivo:
          "Debe indicar la empresa antes de guardar la homologación.",
      };
    }

    const existente =
      cargarHomologacionGuardada(
        empresa
      );

    const confirmadas = {};

    Object.entries(
      homologacion || {}
    ).forEach(
      ([key, h]) => {
        if (
          h?.confirmado &&
          h?.codigo
        ) {
          confirmadas[key] = {
            es:
              h.es || "",

            codigo:
              h.codigo,

            ocupacion:
              h.ocupacion ||
              "",

            confirmado:
              true,

            estado:
              ESTADOS_HOMOLOGACION.CONFIRMADO,
          };
        }
      }
    );

    localStorage.setItem(
      claveStorage(
        empresa
      ),
      JSON.stringify({
        ...existente,
        ...confirmadas,
      })
    );

    return {
      guardado: true,
      cantidad:
        Object.keys(
          confirmadas
        ).length,
    };
  } catch (error) {
    return {
      guardado: false,
      motivo:
        error?.message ||
        "No se pudo guardar la homologación.",
    };
  }
}

// =====================================================================
// Agrupación para las matrices
// =====================================================================

function redondearEntero(
  numero
) {
  return Math.round(
    Number(numero) || 0
  );
}

// ---------------------------------------------------------------------
// Agrupar por cargo
//
// Las horas resultantes son ENTERAS.
// ---------------------------------------------------------------------

export function agruparPorCargo(
  empleados,
  homologacion,
  jornadaSemanal
) {
  const grupos =
    new Map();

  const jornada =
    Number(
      jornadaSemanal
    ) || 0;

  for (
    const emp of empleados || []
  ) {
    if (
      !(emp.proporcion > 0)
    ) {
      continue;
    }

    const h =
      homologacion?.[
        emp.cargoKey
      ] || {};

    const codigo =
      h.codigo || "";

    const nombre =
      h.es ||
      emp.cargo ||
      "(SIN CARGO)";

    /*
     * Si hay código pero NO está confirmado,
     * seguimos mostrando la sugerencia en la
     * revisión, pero el exportador podrá decidir
     * no enviarla hasta confirmación.
     */
    const clave =
      codigo
        ? `CNO::${codigo}`
        : `SIN::${normalizar(
            nombre
          )}`;

    if (
      !grupos.has(clave)
    ) {
      grupos.set(
        clave,
        {
          codigo,

          nombreCargo:
            nombre,

          calificado:
            Boolean(codigo),

          confirmado:
            Boolean(
              h.confirmado
            ),

          estado:
            h.estado ||
            (
              codigo
                ? ESTADOS_HOMOLOGACION.SUGERIDO
                : ESTADOS_HOMOLOGACION.NO_ENCONTRADO
            ),

          completos: 0,

          parciales:
            new Map(),

          personas: 0,
        }
      );
    }

    const g =
      grupos.get(
        clave
      );

    g.personas += 1;

    if (
      emp.proporcion >= 1
    ) {
      g.completos += 1;
    } else {
      const horas =
        redondearEntero(
          jornada *
            emp.proporcion
        );

      g.parciales.set(
        horas,
        (
          g.parciales.get(
            horas
          ) || 0
        ) + 1
      );
    }
  }

  return [
    ...grupos.values(),
  ].map(
    (g) => {
      const lineas = [];

      if (
        g.completos > 0
      ) {
        lineas.push({
          trabajadores:
            g.completos,

          jornada:
            redondearEntero(
              jornada
            ),

          total:
            redondearEntero(
              g.completos *
                jornada
            ),

          parcial:
            false,
        });
      }

      [
        ...g.parciales.entries(),
      ]
        .sort(
          (a, b) =>
            b[0] - a[0]
        )
        .forEach(
          ([horas, n]) => {
            lineas.push({
              trabajadores:
                n,

              jornada:
                redondearEntero(
                  horas
                ),

              total:
                redondearEntero(
                  n * horas
                ),

              parcial:
                true,
            });
          }
        );

      return {
        codigo:
          g.codigo,

        nombreCargo:
          g.nombreCargo,

        calificado:
          g.calificado,

        confirmado:
          g.confirmado,

        estado:
          g.estado,

        personas:
          g.personas,

        lineas,

        totalHoras:
          redondearEntero(
            lineas.reduce(
              (a, l) =>
                a +
                l.total,
              0
            )
          ),
      };
    }
  );
}

// =====================================================================
// Plantilla promedio
// =====================================================================

export function calcularPlantillaPromedio(
  resultados,
  homologacion
) {
  const meses =
    resultados?.length ||
    0;

  const n =
    meses || 1;

  const acumulado =
    new Map();

  resultados.forEach(
    (r, indiceMes) => {
      const grupos =
        agruparPorCargo(
          r.empleados,
          homologacion,
          r.jornadaSemanal
        );

      grupos.forEach(
        (g) => {
          const clave =
            g.calificado
              ? `CNO::${g.codigo}`
              : `SIN::${normalizar(
                  g.nombreCargo
                )}`;

          if (
            !acumulado.has(
              clave
            )
          ) {
            acumulado.set(
              clave,
              {
                codigo:
                  g.codigo,

                nombre:
                  g.nombreCargo,

                calificado:
                  g.calificado,

                confirmado:
                  g.confirmado,

                estado:
                  g.estado,

                personas:
                  0,

                horas:
                  0,

                porMes:
                  Array(
                    n
                  ).fill(0),
              }
            );
          }

          const a =
            acumulado.get(
              clave
            );

          a.personas +=
            g.personas;

          a.horas +=
            g.totalHoras;

          /*
           * Guardamos el dato de cada mes.
           * Esto permitirá construir después la cuadrícula:
           *
           * Código | Mes 1 | Mes 2 | ... | Mes 6 | Suma | Promedio
           */
          a.porMes[
            indiceMes
          ] += g.personas;
        }
      );
    }
  );

  const filas = [
    ...acumulado.values(),
  ]
    .map(
      (a) => {
        const suma =
          a.porMes.reduce(
            (total, valor) =>
              total + valor,
            0
          );

        return {
          codigo:
            a.codigo,

          nombre:
            a.nombre,

          calificado:
            a.calificado,

          confirmado:
            a.confirmado,

          estado:
            a.estado,

          porMes:
            a.porMes.map(
              (valor) =>
                redondearEntero(
                  valor
                )
            ),

          suma:
            redondearEntero(
              suma
            ),

          promedioPersonas:
            Math.round(
              (suma / n) *
                100
            ) / 100,

          promedioHoras:
            Math.round(
              (a.horas / n) *
                100
            ) / 100,
        };
      }
    )
    .sort(
      (x, y) =>
        (y.calificado -
          x.calificado) ||
        String(
          x.codigo ||
            x.nombre
        ).localeCompare(
          String(
            y.codigo ||
              y.nombre
          )
        )
    );

  const totalPersonasPorMes =
    resultados.map(
      (r) =>
        (
          r.empleados ||
          []
        ).filter(
          (e) =>
            e.proporcion > 0
        ).length
    );

  const promedioTotalPersonas =
    Math.round(
      (
        totalPersonasPorMes.reduce(
          (a, b) =>
            a + b,
          0
        ) / n
      ) * 100
    ) / 100;

  return {
    filas,

    meses,

    promedioTotalPersonas,

    totalPersonasPorMes,
  };
}

// =====================================================================
// Utilidades para revisión
// =====================================================================

export function obtenerPendientesHomologacion(
  empleados,
  homologacion
) {
  const pendientes = [];

  const vistos =
    new Set();

  for (
    const empleado of empleados || []
  ) {
    const key =
      empleado.cargoKey ||
      claveCargo(
        empleado.cargo
      );

    if (
      vistos.has(key)
    ) {
      continue;
    }

    vistos.add(key);

    const h =
      homologacion?.[
        key
      ];

    if (!h) {
      pendientes.push({
        cargoKey: key,

        cargo:
          empleado.cargo ||
          "(SIN CARGO)",

        estado:
          ESTADOS_HOMOLOGACION.NO_ENCONTRADO,

        motivo:
          "No existe homologación para este cargo.",
      });

      continue;
    }

    if (
      h.estado ===
        ESTADOS_HOMOLOGACION.NO_ENCONTRADO ||
      h.estado ===
        ESTADOS_HOMOLOGACION.BAJA_COBERTURA ||
      !h.codigo
    ) {
      pendientes.push({
        cargoKey: key,

        cargo:
          empleado.cargo ||
          "(SIN CARGO)",

        estado:
          h.estado ||
          ESTADOS_HOMOLOGACION.NO_ENCONTRADO,

        motivo:
          h.estado ===
          ESTADOS_HOMOLOGACION.BAJA_COBERTURA
            ? "La coincidencia CNO tiene baja cobertura."
            : "No existe un código CNO confirmado.",
      });
    }
  }

  return pendientes;
}

// ---------------------------------------------------------------------
// Validación de un resultado mensual
// ---------------------------------------------------------------------

export function validarResultadoMensual(
  resultado
) {
  const errores = [];
  const advertencias = [];

  if (
    !resultado
  ) {
    errores.push(
      "No existe resultado mensual."
    );

    return {
      valido: false,
      errores,
      advertencias,
    };
  }

  if (
    !Array.isArray(
      resultado.empleados
    )
  ) {
    errores.push(
      "El resultado no contiene la lista de empleados."
    );
  }

  if (
    resultado.jornadaSemanal ===
      undefined ||
    resultado.jornadaSemanal ===
      null
  ) {
    advertencias.push(
      "No se indicó la jornada semanal."
    );
  }

  const empleados =
    resultado.empleados ||
    [];

  const documentos =
    new Set();

  for (
    const empleado of empleados
  ) {
    const documento =
      normalizarDocumento(
        empleado.documento
      );

    if (!documento) {
      errores.push(
        "Existe un empleado sin documento válido."
      );

      continue;
    }

    if (
      documentos.has(
        documento
      )
    ) {
      advertencias.push(
        `Documento duplicado en el resultado: ${documento}.`
      );
    }

    documentos.add(
      documento
    );

    if (
      !empleado.cargo ||
      empleado.cargo ===
        "(SIN CARGO)"
    ) {
      advertencias.push(
        `El documento ${documento} no tiene cargo resuelto.`
      );
    }

    if (
      empleado.alertaHoras
    ) {
      advertencias.push(
        `Las horas de Seguridad Social difieren de la proporción por fechas para ${documento}.`
      );
    }
  }

  return {
    valido:
      errores.length === 0,

    errores,

    advertencias,
  };
}