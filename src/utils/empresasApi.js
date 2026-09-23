import { collection, doc, getDoc, getDocs, setDoc, deleteDoc } from "firebase/firestore";
import { db } from "../firebase";

const COLECCION = "empresasSena";
// El id del documento no puede tener "/", y así "CIPY" y "cipy " son la misma empresa
const idDeEmpresa = (nombre) => nombre.trim().toLowerCase().replace(/\//g, "-");

export async function listarEmpresasApi() {
  const snap = await getDocs(collection(db, COLECCION));
  return snap.docs
    .map((d) => ({ nombre: d.data().nombre, actualizadoEn: d.data().actualizadoEn }))
    .sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));
}

export async function cargarEmpresaApi(nombre) {
  const snap = await getDoc(doc(db, COLECCION, idDeEmpresa(nombre)));
  if (!snap.exists()) return null; // empresa nueva
  const { nombre: _n, actualizadoEn: _a, ...config } = snap.data();
  return config;
}

export async function guardarEmpresaApi(nombre, cambios) {
  // Firestore no acepta "undefined"; esto lo limpia
  const limpio = JSON.parse(JSON.stringify(cambios));
  const ahora = new Date().toISOString();
  await setDoc(
    doc(db, COLECCION, idDeEmpresa(nombre)),
    { ...limpio, nombre: nombre.trim(), actualizadoEn: ahora },
    // mergeFields reemplaza cada campo completo (igual que tu server.js), no lo mezcla por dentro
    { mergeFields: [...Object.keys(limpio), "nombre", "actualizadoEn"] }
  );
  return { ok: true, actualizadoEn: ahora };
}

export async function eliminarEmpresaApi(nombre) {
  await deleteDoc(doc(db, COLECCION, idDeEmpresa(nombre)));
  return { ok: true };
}