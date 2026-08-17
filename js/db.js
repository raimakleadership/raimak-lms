const LocalDB = {
  dbName: "RaimakDB",
  version: 1, // We can increment this later if we need to add more tables!
  db: null,

  // 1. Open the connection and build the tables
  init: async function () {
    return new Promise((resolve, reject) => {
      // 🚀 MOBILE FIX 1: The Private Browsing / Missing IDB Guard
      if (!window.indexedDB) {
        console.warn(
          "IndexedDB is not supported or is blocked by Private Browsing. App will run in memory-only mode.",
        );
        return resolve(false);
      }

      try {
        const request = indexedDB.open(this.dbName, this.version);

        // This only runs the very first time the app loads, or if we change the version number
        request.onupgradeneeded = (event) => {
          const db = event.target.result;

          // Create our two heavy-duty tables. 'id' is the primary key.
          if (!db.objectStoreNames.contains("activity_logs")) {
            db.createObjectStore("activity_logs", { keyPath: "id" });
          }
          if (!db.objectStoreNames.contains("leads")) {
            db.createObjectStore("leads", { keyPath: "id" });
          }
        };

        request.onsuccess = (event) => {
          this.db = event.target.result;
          resolve(true);
        };

        request.onerror = (event) => {
          console.error("IndexedDB Error:", event.target.error);
          reject(event.target.error);
        };
      } catch (err) {
        console.error("Critical failure opening IndexedDB:", err);
        resolve(false);
      }
    });
  },

  // 2. Save a massive array of items instantly
  saveItems: async function (storeName, items) {
    if (!this.db) return false; // Safety fallback if DB failed to initialize

    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([storeName], "readwrite");
      const store = transaction.objectStore(storeName);

      items.forEach((item) => store.put(item));

      transaction.oncomplete = () => resolve(true);
      transaction.onerror = (event) => reject(event.target.error);

      // 🚀 MOBILE FIX 2: Catch iOS background tab assassinations
      transaction.onabort = (event) => {
        console.warn(
          `Transaction aborted in ${storeName} (likely iOS backgrounding):`,
          event,
        );
        reject(new Error("Transaction aborted by browser"));
      };
    });
  },

  // 3. Load the entire iceberg out of the hard drive
  getAllItems: async function (storeName) {
    if (!this.db) return [];

    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([storeName], "readonly");
      const store = transaction.objectStore(storeName);
      const request = store.getAll();

      request.onsuccess = (event) => resolve(event.target.result);
      request.onerror = (event) => reject(event.target.error);
      transaction.onabort = () =>
        reject(new Error("Transaction aborted by browser"));
    });
  },

  // 4. Delete a specific item by ID
  deleteItem: async function (storeName, id) {
    if (!this.db) return false;

    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([storeName], "readwrite");
      const store = transaction.objectStore(storeName);

      const request = store.delete(id);

      request.onsuccess = () => resolve(true);
      request.onerror = (event) => reject(event.target.error);
      transaction.onabort = () =>
        reject(new Error("Transaction aborted by browser"));
    });
  },
};
