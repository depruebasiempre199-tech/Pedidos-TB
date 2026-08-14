import { initializeApp } from "firebase/app";
import {
  getDatabase,
  ref,
  set,
  get,
  update,
} from "firebase/database";
import { firebaseConfig } from "./firebase-config";

const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

// Nodo "historial": un hijo por semana guardada, bajo su propio id.
// Estructura en la base de datos:
// historial/
//   {snapshotId}/
//     fecha: string (fecha del pedido, ej. "2026-08-10")
//     items: [
//       { producto, proveedor, dia, unidad, pedir, selected }
//     ]

export async function fetchHistorial() {
  const snapshot = await get(ref(db, "historial"));
  if (!snapshot.exists()) return [];
  const data = snapshot.val();
  const list = Object.entries(data).map(([id, value]) => ({ id, ...value }));
  list.sort((a, b) => (a.fecha < b.fecha ? 1 : -1));
  return list;
}

export async function saveSnapshot(snapshot) {
  await set(ref(db, `historial/${snapshot.id}`), {
    fecha: snapshot.fecha,
    items: snapshot.items,
  });
  return fetchHistorial();
}

export async function updateSnapshot(snapshotId, items) {
  await update(ref(db, `historial/${snapshotId}`), { items });
  return fetchHistorial();
}

// Nodo "estado/actual": guarda el estado editable completo de la app
// (productos, proveedores, configuración de días/calendario) para que
// no se pierda al cerrar o recargar la página, ni cambiar de dispositivo.
export async function fetchEstado() {
  const snapshot = await get(ref(db, "estado/actual"));
  if (!snapshot.exists()) return null;
  return snapshot.val();
}

export async function saveEstado(estado) {
  await set(ref(db, "estado/actual"), estado);
}
