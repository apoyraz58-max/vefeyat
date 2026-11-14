/**
 * Catalog storage in the add-in's own browser storage (IndexedDB).
 * The catalog published with the add-in (catalog.csv) loads automatically;
 * a CSV chosen by the user replaces it on this computer until reset.
 * Nothing is sent anywhere.
 */

const DB_NAME = "vefeyat";
const STORE = "catalog";
const KEY = "current";

export interface StoredCatalog {
  readonly csv: string;
  readonly fileName: string;
  readonly savedAt: string;
  /** "user": chosen by the user; "embedded": copy of the published catalog. */
  readonly source?: "user" | "embedded";
}

function open(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("IndexedDB unavailable"));
  });
}

async function withStore<T>(mode: IDBTransactionMode, run: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await open();
  try {
    return await new Promise<T>((resolve, reject) => {
      const request = run(db.transaction(STORE, mode).objectStore(STORE));
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error ?? new Error("IndexedDB request failed"));
    });
  } finally {
    db.close();
  }
}

export async function loadCatalog(): Promise<StoredCatalog | null> {
  try {
    const value = await withStore<StoredCatalog | undefined>("readonly", (s) => s.get(KEY));
    return value ?? null;
  } catch {
    return null;
  }
}

/** Returns false when the catalog could only be kept for this session. */
export async function saveCatalog(value: StoredCatalog): Promise<boolean> {
  try {
    await withStore("readwrite", (s) => s.put(value, KEY));
    return true;
  } catch {
    return false;
  }
}

export async function clearCatalog(): Promise<void> {
  try {
    await withStore("readwrite", (s) => s.delete(KEY));
  } catch {
    // Nothing stored or storage unavailable.
  }
}

/** The catalog published next to the task pane, or null when there is none. */
export async function fetchEmbeddedCatalog(): Promise<string | null> {
  try {
    const response = await fetch("catalog.csv", { cache: "no-cache" });
    if (!response.ok) return null;
    return await response.text();
  } catch {
    return null;
  }
}
