// server.js
// Servidor proxy mínimo: recibe la petición de tu app React (frontend)
// y llama él mismo a la API de Anthropic (servidor-a-servidor, sin CORS).
//
// Instalación:
//   npm install express cors dotenv
//
// Crea un archivo .env en la raíz del proyecto (NO lo subas a git) con:
//   ANTHROPIC_API_KEY=sk-ant-xxxxxxxxxxxxxxxxxxxx
//
// Ejecución:
//   node server.js
//
// Esto levanta el proxy en http://localhost:3001

import express from "express";
import cors from "cors";
import dotenv from "dotenv";

dotenv.config();

const app = express();
app.use(cors()); // permite que tu frontend en :5173 le hable a este servidor
app.use(express.json());

const API_KEY = process.env.ANTHROPIC_API_KEY;

if (!API_KEY) {
  console.error("Falta ANTHROPIC_API_KEY en el archivo .env");
  process.exit(1);
}

app.post("/api/traducir-cargos", async (req, res) => {
  try {
    const { cargosUnicos } = req.body;

    if (!Array.isArray(cargosUnicos) || cargosUnicos.length === 0) {
      return res.status(400).json({ error: "cargosUnicos debe ser un arreglo no vacío" });
    }

    const prompt = `Traduce estos nombres de cargos de trabajo del inglés al español, de forma corta y natural (como se usaría en una nómina colombiana). Responde SOLO un JSON válido, sin texto adicional, con este formato exacto:
{"CARGO EN INGLES": "cargo en español", ...}

Cargos:
${cargosUnicos.join("\n")}`;

    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": API_KEY,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 1000,
        messages: [{ role: "user", content: prompt }],
      }),
    });

    if (!response.ok) {
      const detalle = await response.text();
      console.error("Error de la API de Anthropic:", detalle);
      return res.status(502).json({ error: "Error al llamar a la API de Anthropic", detalle });
    }

    const data = await response.json();
    const texto = (data.content || [])
      .map((b) => b.text || "")
      .join("")
      .replace(/```json|```/g, "")
      .trim();

    let traducciones = {};
    try {
      traducciones = JSON.parse(texto);
    } catch (e) {
      console.warn("No se pudo parsear la respuesta como JSON:", texto);
    }

    res.json({ traducciones });
  } catch (err) {
    console.error("Error en /api/traducir-cargos:", err);
    res.status(500).json({ error: err.message || "Error interno del servidor" });
  }
});

const PORT = 3001;
app.listen(PORT, () => {
  console.log(`Proxy de Anthropic escuchando en http://localhost:${PORT}`);
});