import express from "express";
import cors from "cors";
import dotenv from "dotenv";

dotenv.config();

const app = express();
app.use(cors({ origin: ["http://localhost:5173", "http://127.0.0.1:5173"] }));
app.use(express.json({ limit: "1mb" }));

const API_KEY = process.env.ANTHROPIC_API_KEY;
const MODELO = process.env.MODELO || "claude-sonnet-5";
const TAMANO_LOTE = 40;

if (!API_KEY) {
  console.error("Falta ANTHROPIC_API_KEY en el archivo .env");
  process.exit(1);
}

async function traducirLote(cargos) {
  const prompt = `Traduce estos nombres de cargos de trabajo del inglés al español, de forma corta y natural (como se usaría en una nómina colombiana). Si un cargo ya está en español, déjalo igual. Responde SOLO un JSON válido, sin texto adicional. Cada llave debe ser el cargo original EXACTO y el valor su traducción:
{"CARGO EN INGLES": "cargo en español", ...}

Cargos:
${cargos.join("\n")}`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: MODELO,
      max_tokens: 4000,
      messages: [{ role: "user", content: prompt }],
    }),
  });

  if (!response.ok) {
    const detalle = await response.text();
    throw new Error(`Anthropic respondió ${response.status}: ${detalle}`);
  }

  const data = await response.json();
  const texto = (data.content || []).map((b) => b.text || "").join("");

  // Extrae el objeto JSON aunque venga con texto o ```json alrededor
  const inicio = texto.indexOf("{");
  const fin = texto.lastIndexOf("}");
  if (inicio === -1 || fin === -1) {
    throw new Error("La respuesta no contenía JSON: " + texto.slice(0, 200));
  }
  const json = JSON.parse(texto.slice(inicio, fin + 1));

  // Empareja por llave sin importar mayúsculas/espacios
  const porClave = new Map(
    Object.entries(json).map(([k, v]) => [k.trim().toUpperCase(), String(v).trim()])
  );
  const salida = {};
  for (const cargo of cargos) {
    const t = porClave.get(cargo.trim().toUpperCase());
    if (t) salida[cargo] = t;
  }
  return salida;
}

app.post("/api/traducir-cargos", async (req, res) => {
  const { cargosUnicos } = req.body;
  if (!Array.isArray(cargosUnicos) || cargosUnicos.length === 0) {
    return res.status(400).json({ error: "cargosUnicos debe ser un arreglo no vacío" });
  }

  const traducciones = {};
  const errores = [];

  for (let i = 0; i < cargosUnicos.length; i += TAMANO_LOTE) {
    const lote = cargosUnicos.slice(i, i + TAMANO_LOTE);
    try {
      Object.assign(traducciones, await traducirLote(lote));
    } catch (e) {
      console.error(`Error en el lote ${i / TAMANO_LOTE + 1}:`, e.message);
      errores.push(e.message);
    }
  }

  // Si no se tradujo nada, que la app lo vea como error (no como éxito vacío)
  if (Object.keys(traducciones).length === 0 && errores.length > 0) {
    return res.status(502).json({ error: errores[0] });
  }
  res.json({ traducciones, advertencia: errores.length ? errores : undefined });
});

const PORT = 3001;
app.listen(PORT, () => console.log(`Proxy de Anthropic escuchando en http://localhost:${PORT}`));