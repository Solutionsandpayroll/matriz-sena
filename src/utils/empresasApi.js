const API_URL = import.meta.env?.VITE_API_URL || "http://localhost:3001";

async function manejarRespuesta(response, accion) {
  if (!response.ok) {
    let detalle = "";
    try {
      const data = await response.json();
      detalle = data?.error || "";
    } catch {
      // sin cuerpo JSON, se ignora
    }
    throw new Error(`Error al ${accion} (${response.status})${detalle ? ": " + detalle : ""}`);
  }
  return response.json();
}

// Devuelve la lista de empresas guardadas: [{ nombre, actualizadoEn }, ...]
export async function listarEmpresasApi() {
  const res = await fetch(`${API_URL}/api/empresas`);
  return manejarRespuesta(res, "listar empresas");
}

// Devuelve la configuración completa de una empresa, o null si nunca se ha guardado.
export async function cargarEmpresaApi(nombre) {
  const res = await fetch(`${API_URL}/api/empresas/${encodeURIComponent(nombre)}`);
  if (res.status === 404) return null;
  return manejarRespuesta(res, "cargar la empresa");
}

// Guarda/actualiza (merge) una parte de la configuración de una empresa.
export async function guardarEmpresaApi(nombre, cambios) {
  const res = await fetch(`${API_URL}/api/empresas/${encodeURIComponent(nombre)}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(cambios),
  });
  return manejarRespuesta(res, "guardar la empresa");
}

// Elimina una empresa y toda su configuración guardada.
export async function eliminarEmpresaApi(nombre) {
  const res = await fetch(`${API_URL}/api/empresas/${encodeURIComponent(nombre)}`, {
    method: "DELETE",
  });
  return manejarRespuesta(res, "eliminar la empresa");
}