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

                                        // --- Próxima fase (no conectado todavía) ---
                                        // Cuando el historial esté funcionando bien, el mismo patrón se puede
                                        // repetir para persistir el resto del estado en sus propios nodos:
                                        //   - "productos"      (lo que hoy es el estado `rows` en el prototipo)
                                        //   - "proveedores"     (alta/baja de proveedores)
                                        //   - "diasInventario"  (objetivo de días de inventario por proveedor+día)
                                        //   - "calendario"      (día de despacho por proveedor)
                                        // Recordá agregar cada nodo nuevo a las reglas de seguridad
                                        // (Realtime Database > Reglas) antes de usarlo, o quedará bloqueado
                                        // por defecto.
                                        // Avisame cuando quieras avanzar con esa parte y armo esos nodos
                                        // siguiendo la misma estructura.
                                        