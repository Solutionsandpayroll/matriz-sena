import { useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import listadoCnoLocal from "./data/listado-cno.json";

const listadoCno = listadoCnoLocal || [];

// 1. Mapeo Directo Determinista (Prioridad Alta)
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

// 2. Equivalencias TPL -> Español para búsqueda semántica de respaldo
const traducciones = {
  "LOGISTICS SPECIALIST I": "tecnico estudio del trabajo",
  "LOGISTICS OPERATIONS SPECIALIST AFTER HOURS": "especialista en operaciones logisticas",
  "CREDIT AND COLLECTIONS ANALYST II": "analista de credito y cobranzas",
  "SALES & PRICING TEAM LEADER": "lider de ventas y precios",
  "BRANCH MANAGER": "gerente de sucursal",
  "ADMINISTRATIVE ANALYST": "analista administrativo",
  "IT SUPPORT SPECIALIST": "soporte tecnico",
  "SALESFORCE JUNIOR DEVELOPER": "desarrollador junior",
  "GERENTE DE CUENTAS JR": "gerente de cuentas",
};

function normalizar(texto) {
  return String(texto || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function obtenerSugerenciasCNO(cargoIngles) {
  const cargoLimpio = String(cargoIngles || "").trim().toUpperCase();

  const codigoFijo = mapeoDirectoCno[cargoLimpio];
  if (codigoFijo) {
    const itemEncontrado = listadoCno.find((item) => String(item.codigo) === String(codigoFijo));
    if (itemEncontrado) {
      return [{ ...itemEncontrado, puntos: 100 }];
    }
  }

  const traduccion = traducciones[cargoLimpio] ?? cargoLimpio;
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

function calcularCuotaSena(plantaPromedio) {
  if (plantaPromedio < 15) return 0;

  let cuota = 1;
  let sobrante = plantaPromedio - 20;

  if (sobrante > 0) {
    cuota += Math.floor(sobrante / 20);
    if (sobrante % 20 >= 10) {
      cuota += 1;
    }
  }
  return cuota;
}

// Función para convertir imagen a Base64 para ExcelJS
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

export default function App() {
  const [nomina, setNomina] = useState(null);
  const [segSocial, setSegSocial] = useState(null);
  const [horasBase, setHorasBase] = useState(210);
  const [aprendicesActivos, setAprendicesActivos] = useState(0);
  const [resultado, setResultado] = useState(null);
  const [cargando, setCargando] = useState(false);
  const [error, setError] = useState(null);
  const [cnoSeleccionados, setCnoSeleccionados] = useState({});
  const [mostrarInstrucciones, setMostrarInstrucciones] = useState(true);

  const leerNomina = (buffer) => {
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const filas = XLSX.utils.sheet_to_json(sheet, { range: 1 });

    return filas
      .filter((fila) => fila["DOCUMENTO IDENTIDAD"])
      .map((fila) => {
        const cargoRaw = String(fila["DESC. OFICIO"] || "").trim();
        const sugerencias = obtenerSugerenciasCNO(cargoRaw);
        return {
          documento: String(fila["DOCUMENTO IDENTIDAD"]).trim(),
          nombre: fila["NOMBRE COMPLETO"],
          cargo: cargoRaw,
          sugerenciasCno: sugerencias,
          cnoPorDefecto: sugerencias[0]?.codigo || "",
          fechaAntiguedad: fila["FECHA ANTIGUEDAD"] || "",
        };
      });
  };

  const leerSeguridadSocial = (buffer) => {
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const filas = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

    const indiceEncabezado = filas.findIndex((fila) => fila && fila.includes("No id"));
    if (indiceEncabezado === -1) {
      throw new Error("No se encontró la cabecera 'No id' en la planilla de Seguridad Social.");
    }

    const filasDatos = filas.slice(indiceEncabezado + 1);
    const horasPorDocumento = {};
    const nombresPorDocumento = {};

    for (const fila of filasDatos) {
      if (!fila) continue;
      const documento = fila[7];
      if (!documento) continue;

      const doc = String(documento).trim();
      const horas = Number(fila[15]) || 0;
      horasPorDocumento[doc] = (horasPorDocumento[doc] || 0) + horas;
      if (fila[9]) nombresPorDocumento[doc] = String(fila[9]).trim();
    }

    return { horasPorDocumento, nombresPorDocumento };
  };

  const procesarArchivos = async () => {
    if (!nomina || !segSocial) return;
    setCargando(true);
    setError(null);

    try {
      const bytesNomina = await nomina.arrayBuffer();
      const bytesSegSocial = await segSocial.arrayBuffer();

      const empleadosNomina = leerNomina(bytesNomina);
      const { horasPorDocumento } = leerSeguridadSocial(bytesSegSocial);

      let sumaProporciones = 0;
      const empleadosProcesados = empleadosNomina.map((emp) => {
        const horasPila = horasPorDocumento[emp.documento] ?? horasBase;
        const proporcionTrabajador = Math.min(horasPila / horasBase, 1);
        sumaProporciones += proporcionTrabajador;

        return {
          ...emp,
          horasPila,
          proporcionTrabajador: proporcionTrabajador.toFixed(2),
        };
      });

      const plantaEquivalente = Number(sumaProporciones.toFixed(2));
      const cuotaSenaCalculada = calcularCuotaSena(plantaEquivalente);
      const deficitAprendices = Math.max(0, cuotaSenaCalculada - Number(aprendicesActivos || 0));

      const iniciales = {};
      empleadosProcesados.forEach((emp) => {
        iniciales[emp.documento] = emp.cnoPorDefecto;
      });
      setCnoSeleccionados(iniciales);

      setResultado({
        totalNomina: empleadosNomina.length,
        totalSeguridadSocial: Object.keys(horasPorDocumento).length,
        plantaEquivalente,
        cuotaSenaCalculada,
        deficitAprendices,
        empleados: empleadosProcesados,
      });
    } catch (err) {
      setError(err.message || "Error al procesar los archivos");
    } finally {
      setCargando(false);
    }
  };

  const exportarPlantillaSena = async () => {
    if (!resultado) return;

    const GRIS_ENCABEZADO = "FFD9D9D9";
    const GRIS_BORDE = "FF000000";

    const bordeNegro = {
      top: { style: "thin", color: { argb: GRIS_BORDE } },
      bottom: { style: "thin", color: { argb: GRIS_BORDE } },
      left: { style: "thin", color: { argb: GRIS_BORDE } },
      right: { style: "thin", color: { argb: GRIS_BORDE } },
    };

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("MATRIZ SENA", {
      views: [{ showGridLines: true }],
    });

    ws.columns = [
      { width: 32 },
      { width: 35 },
      { width: 35 },
      { width: 32 },
    ];

    // 1. INSERTAR LOGO DEL SENA
    try {
      // Asegúrate de colocar el archivo sena-logo.png en tu carpeta /public/
      const base64Logo = await obtenerImagenBase64("/sena-logo.png");
      const logoId = wb.addImage({
        base64: base64Logo,
        extension: "png",
      });

      // Ubicación de la imagen sobre las columnas B y C (Filas 1 a 3)
      ws.addImage(logoId, {
        tl: { col: 1.6, row: 0.2 },
        ext: { width: 130, height: 60 },
      });
    } catch (e) {
      console.warn("No se pudo cargar la imagen del logo SENA:", e);
    }

    // Espacio para las filas 1 a 3 del Logo
    for (let i = 1; i <= 3; i++) ws.addRow([]);

    // 2. DATOS BÁSICOS DE LA EMPRESA
    const fHeaderEmpresa = ws.addRow(["DATOS BASICOS DE LA EMPRESA", "", "", ""]);
    ws.mergeCells(fHeaderEmpresa.number, 1, fHeaderEmpresa.number, 4);

    fHeaderEmpresa.getCell(1).font = { bold: true, size: 10 };
    fHeaderEmpresa.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
    fHeaderEmpresa.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRIS_ENCABEZADO } };

    const datosEmpresa = [
      ["Razón Social", "TPL LOGISTICS SUPPORT SAS", "Nit:", "901.593.999~4"],
      ["Representante Legal:", "JEAN PIERRE ASSELIN", "CC:", "AK678743"],
      ["Dirección (municipio/barrio)", "Carrera 45 # 108 ~ 27 Oficina 1601 de la Torre 2", "Telefonos:", "3003689121"],
      ["E_Mail", "JCERVANTES@SOLUTIONSANDPAYROLL.COM", "Fecha:", "Ago / 2026"],
    ];

    datosEmpresa.forEach((row) => {
      const f = ws.addRow(row);
      f.height = 18;
      f.eachCell((celda, colIndex) => {
        celda.border = bordeNegro;
        celda.font = { size: 9 };
        if (colIndex === 1 || colIndex === 3) {
          celda.font = { bold: true, size: 9 };
        }
        if (colIndex === 2 && row[0] === "E_Mail") {
          celda.font = { color: { argb: "FF0000FF" }, underline: true, size: 9 };
        }
      });
    });

    ws.addRow([]);

    // 3. MATRIZ 1: OFICIOS CALIFICADOS
    const fMatriz1Tit = ws.addRow(["Matriz 1: Oficios Calificados"]);
    fMatriz1Tit.getCell(1).font = { bold: true, size: 10 };

    const fEncMatriz1 = ws.addRow([
      "CÓDIGO DEL OFICIO SEGÚN LISTADO DE OFICIOS Y OCUPACIONES",
      "NÚMERO DE TRABAJADORES",
      "JORNADA LABORAL SEMANAL POR TRABAJADOR",
      "TOTAL~JORNADA LABORAL SEMANAL",
    ]);
    fEncMatriz1.height = 28;
    fEncMatriz1.eachCell((celda) => {
      celda.font = { bold: true, size: 8 };
      celda.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRIS_ENCABEZADO } };
      celda.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      celda.border = bordeNegro;
    });

    const conteoCno = {};
    resultado.empleados.forEach((emp) => {
      const cod = cnoSeleccionados[emp.documento] || emp.cnoPorDefecto || "0000";
      conteoCno[cod] = (conteoCno[cod] || 0) + 1;
    });

    let filaInicioM1 = ws.lastRow.number + 1;

    Object.entries(conteoCno).forEach(([codigo, cantidad]) => {
      const f = ws.addRow([codigo, cantidad, 40, cantidad * 40]);
      f.height = 18;
      f.eachCell((celda) => {
        celda.border = bordeNegro;
        celda.alignment = { horizontal: "center", vertical: "middle" };
        celda.font = { size: 9 };
      });
    });

    let filaFinM1 = ws.lastRow.number;

    const fTotalM1 = ws.addRow([
      "TOTAL",
      { formula: `SUM(B${filaInicioM1}:B${filaFinM1})` },
      { formula: `SUM(C${filaInicioM1}:C${filaFinM1})` },
      { formula: `SUM(D${filaInicioM1}:D${filaFinM1})` },
    ]);
    fTotalM1.height = 20;
    fTotalM1.eachCell((celda) => {
      celda.font = { bold: true, size: 9 };
      celda.border = bordeNegro;
      celda.alignment = { horizontal: "center", vertical: "middle" };
      celda.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRIS_ENCABEZADO } };
    });

    ws.addRow([]);

    // 4. MATRIZ 2: OFICIOS NO CALIFICADOS
    const fMatriz2Tit = ws.addRow(["Matriz 2 Oficios no Calificados"]);
    fMatriz2Tit.getCell(1).font = { bold: true, size: 10 };

    const fEncMatriz2 = ws.addRow([
      "NOBRE DEL CARGO",
      "JORNADA LABORAL SEMANAL POR TRABAJADOR",
      "TOTAL~JORNADA LABORAL SEMANAL",
    ]);
    fEncMatriz2.height = 25;
    fEncMatriz2.eachCell((celda) => {
      celda.font = { bold: true, size: 8 };
      celda.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRIS_ENCABEZADO } };
      celda.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      celda.border = bordeNegro;
    });

    const fFilaVaciaM2 = ws.addRow([0, 0, 0]);
    fFilaVaciaM2.eachCell((c) => {
      c.border = bordeNegro;
      c.alignment = { horizontal: "center" };
    });

    const fTotalM2 = ws.addRow(["TOTAL", 0, 0]);
    fTotalM2.eachCell((celda) => {
      celda.font = { bold: true, size: 9 };
      celda.border = bordeNegro;
      celda.alignment = { horizontal: "center" };
      celda.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRIS_ENCABEZADO } };
    });

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "MATRIZ_DETERMINACION_CUOTA_SENA_TPL.xlsx";
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-slate-50/60 font-sans text-slate-800 antialiased">
      <header className="bg-white border-b border-slate-200/80 sticky top-0 z-50">
        <div className="max-w-6xl mx-auto px-6 py-3 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <img
              src="/logo.jpeg"
              alt="Solutions & Payroll Logo"
              className="h-9 w-auto object-contain"
            />
          </div>

          <div className="flex items-center gap-2 bg-white border border-gray-200 rounded-full px-4 py-1.5 shadow-2xs">
            <svg
              xmlns="http://www.w3.org/2000/svg"
              fill="none"
              viewBox="0 0 24 24"
              strokeWidth="1.8"
              stroke="#2563eb"
              className="w-4 h-4"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                d="M15.75 6a3.75 3.75 0 1 1-7.5 0 3.75 3.75 0 0 1 7.5 0ZM4.501 20.118a7.5 7.5 0 0 1 14.998 0A17.933 17.933 0 0 1 12 21.75c-2.676 0-5.216-.584-7.499-1.632Z"
              />
            </svg>
            <span className="text-xs font-semibold text-slate-700">Bienvenida, Greylin</span>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-10 space-y-8">
        <section className="text-center space-y-2">
          <h1 className="text-3xl font-extrabold text-slate-900 tracking-tight">
            Gestión de Planta y Cuota Aprendices SENA
          </h1>
          <p className="text-sm text-slate-500 max-w-2xl mx-auto">
            Calcula la planta equivalente según PILA, verifica la cuota requerida y exporta la Matriz de Determinación de Cuota SENA.
          </p>
        </section>

        <div className="bg-[#EFF6FF] border border-[#BFDBFE] rounded-2xl p-5 shadow-sm transition-all">
          <button
            onClick={() => setMostrarInstrucciones(!mostrarInstrucciones)}
            className="w-full flex items-center justify-between text-left font-bold text-[#1E40AF] text-base"
          >
            <div className="flex items-center gap-2.5">
              <div className="w-6 h-6 rounded-full border-2 border-[#1E40AF] flex items-center justify-center text-xs font-bold">
                i
              </div>
              <span>¿Cómo usar esta aplicación?</span>
            </div>
            <svg
              className={`w-5 h-5 transition-transform duration-200 ${mostrarInstrucciones ? "rotate-180" : ""}`}
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
            >
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 9l-7 7-7-7" />
            </svg>
          </button>

          {mostrarInstrucciones && (
            <div className="mt-4 pt-4 border-t border-[#DBEAFE] space-y-2 text-xs font-medium text-[#1E3A8A] leading-relaxed">
              <p><strong className="font-bold">1. Cargar Archivos:</strong> Adjunta la Nómina mensual (.xlsx) y la Planilla de Seguridad Social PILA (.xlsx).</p>
              <p><strong className="font-bold">2. Ajustar Parámetros:</strong> Selecciona la jornada base mensual (210h según Ley 2101 vigente) e indica los aprendices contratados actualmente.</p>
              <p><strong className="font-bold">3. Procesar Planta:</strong> Haz clic en "Calcular Planta Equivalente" para procesar el cálculo de horas y cuota legal.</p>
              <p><strong className="font-bold">4. Homologación CNO:</strong> Revisa y ajusta los códigos CNO asignados automáticamente a cada cargo.</p>
              <p><strong className="font-bold">5. Exportar Reporte SENA:</strong> Descarga la Matriz de Determinación de Cuota SENA (.xlsx).</p>
            </div>
          )}
        </div>

        <section className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-sm space-y-6">
          <div className="border-b border-slate-100 pb-3">
            <h2 className="text-base font-bold text-slate-800">Carga de Fuentes de Datos</h2>
            <p className="text-xs text-slate-500">Sube los archivos en formato Excel (.xlsx) para iniciar el análisis.</p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <FileInput
              label="1. Archivo de Nómina Mensual"
              file={nomina}
              onChange={setNomina}
            />
            <FileInput
              label="2. Planilla Seguridad Social (PILA)"
              file={segSocial}
              onChange={setSegSocial}
            />
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-2">
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-700">
                Jornada Laboral Ordinaria Base:
              </label>
              <select
                value={horasBase}
                onChange={(e) => setHorasBase(Number(e.target.value))}
                className="w-full bg-slate-50 border border-slate-200 text-slate-800 rounded-xl px-4 py-2.5 text-xs font-medium focus:ring-2 focus:ring-[#003B7A] focus:outline-none transition"
              >
                <option value={210}>210 hrs/mes (42h semanales - Vigente Ley 2101 de 2021)</option>
                <option value={220}>220 hrs/mes (44h semanales - Periodos 2024-2025)</option>
                <option value={240}>240 hrs/mes (48h semanales - Histórico)</option>
              </select>
            </div>

            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-700">
                Aprendices Activos Actualmente:
              </label>
              <input
                type="number"
                min="0"
                value={aprendicesActivos}
                onChange={(e) => setAprendicesActivos(e.target.value)}
                placeholder="Ej. 1"
                className="w-full bg-slate-50 border border-slate-200 text-slate-800 rounded-xl px-4 py-2.5 text-xs font-medium focus:ring-2 focus:ring-[#003B7A] focus:outline-none transition"
              />
            </div>
          </div>

          <button
            onClick={procesarArchivos}
            disabled={!nomina || !segSocial || cargando}
            className="w-full bg-[#003B7A] hover:bg-[#002D5E] disabled:bg-slate-200 disabled:text-slate-400 text-white font-semibold py-3 rounded-xl transition shadow-md text-sm flex items-center justify-center gap-2"
          >
            {cargando ? "Procesando Matriz SENA..." : "Calcular Planta Equivalente y Generar Matriz"}
          </button>
        </section>

        {error && (
          <div className="p-4 bg-red-50 border border-red-200 text-red-700 rounded-xl text-sm font-medium">
            {error}
          </div>
        )}

        {resultado && (
          <section className="space-y-6">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <StatCard label="Headcount Nómina" value={resultado.totalNomina} />
              <StatCard label="Planta Real (Horas PILA)" value={resultado.plantaEquivalente} />
              <StatCard label="Cuota SENA Requerida" value={`${resultado.cuotaSenaCalculada}`} />
              <StatCard
                label="Déficit / Pendientes"
                value={`${resultado.deficitAprendices}`}
                highlight={resultado.deficitAprendices > 0}
              />
            </div>

            <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
              <div className="p-5 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                <div>
                  <h3 className="font-bold text-slate-900 text-base">Matriz Homologada Planta SENA</h3>
                  <p className="text-xs text-slate-500 mt-0.5">
                    Verifica la asignación de ocupaciones CNO antes de descargar la Matriz SENA.
                  </p>
                </div>
                <button
                  onClick={exportarPlantillaSena}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold px-4 py-2.5 rounded-xl transition shadow-sm flex items-center justify-center gap-2"
                >
                  📥 Exportar Matriz SENA (.xlsx)
                </button>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-sm">
                  <thead className="bg-slate-50 text-slate-600 uppercase text-[10px] font-bold tracking-wider border-b border-slate-200/80">
                    <tr>
                      <th className="py-3.5 px-4">Documento</th>
                      <th className="py-3.5 px-4">Empleado</th>
                      <th className="py-3.5 px-4">Cargo Empresa</th>
                      <th className="py-3.5 px-4">Horas / Fracción</th>
                      <th className="py-3.5 px-4">Ocupación CNO Asignada</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 text-slate-700">
                    {resultado.empleados.map((emp) => (
                      <tr key={emp.documento} className="hover:bg-slate-50/80 transition">
                        <td className="py-3.5 px-4 font-mono text-xs text-slate-500">{emp.documento}</td>
                        <td className="py-3.5 px-4 font-bold text-slate-900">{emp.nombre}</td>
                        <td className="py-3.5 px-4 text-xs text-slate-600">{emp.cargo}</td>
                        <td className="py-3.5 px-4">
                          <span className="text-xs font-semibold text-slate-800">{emp.horasPila} hrs</span>
                          <span className="ml-2 bg-blue-50 text-[#003B7A] border border-blue-200 px-2 py-0.5 rounded text-[10px] font-bold">
                            {emp.proporcionTrabajador}
                          </span>
                        </td>
                        <td className="py-3.5 px-4">
                          <select
                            value={cnoSeleccionados[emp.documento] || ""}
                            onChange={(e) => setCnoSeleccionados({ ...cnoSeleccionados, [emp.documento]: e.target.value })}
                            className="w-full bg-slate-50 border border-slate-200 text-slate-800 rounded-lg p-2 text-xs focus:ring-2 focus:ring-[#003B7A] focus:outline-none"
                          >
                            {emp.sugerenciasCno.map((sug) => (
                              <option key={sug.codigo} value={sug.codigo}>
                                [{sug.codigo}] {sug.ocupacion}
                              </option>
                            ))}
                          </select>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        )}
      </main>
    </div>
  );
}

function FileInput({ label, file, onChange }) {
  return (
    <div className="space-y-1.5">
      <label className="text-xs font-bold text-slate-700">{label}</label>
      <label className="flex items-center justify-center border-2 border-dashed border-slate-200 hover:border-[#003B7A]/50 hover:bg-blue-50/30 rounded-xl py-6 px-4 cursor-pointer transition text-center group">
        <span className="text-xs text-slate-500 group-hover:text-slate-700 font-medium transition">
          {file ? `📄 ${file.name}` : "Seleccionar archivo Excel (.xlsx)"}
        </span>
        <input type="file" accept=".xlsx" className="hidden" onChange={(e) => onChange(e.target.files?.[0] || null)} />
      </label>
    </div>
  );
}

function StatCard({ label, value, highlight }) {
  return (
    <div className={`p-5 rounded-2xl border ${highlight ? "bg-amber-50/60 border-amber-200 text-amber-900" : "bg-white border-slate-200 text-slate-800"}`}>
      <p className="text-2xl font-black tracking-tight">{value}</p>
      <p className="text-xs text-slate-500 mt-1 font-semibold">{label}</p>
    </div>
  );
}