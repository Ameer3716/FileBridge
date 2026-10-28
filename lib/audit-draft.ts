const DB = "filebridge-audit",
  STORE = "draft";
export type AuditDraft = {
  bytes: ArrayBuffer;
  filename: string;
  sheetName: string;
};
async function database(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = () => r.result.createObjectStore(STORE);
    r.onsuccess = () => resolve(r.result);
    r.onerror = () => reject(r.error);
  });
}
export async function draft(
  action: "read" | "write",
  value?: AuditDraft,
): Promise<AuditDraft | undefined> {
  const db = await database();
  try {
    return await new Promise((resolve, reject) => {
      const t = db.transaction(
        STORE,
        action === "read" ? "readonly" : "readwrite",
      );
      const s = t.objectStore(STORE);
      const r = action === "read" ? s.get("current") : s.put(value, "current");
      t.oncomplete = () => resolve(action === "read" ? r.result : undefined);
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error);
    });
  } finally {
    db.close();
  }
}
