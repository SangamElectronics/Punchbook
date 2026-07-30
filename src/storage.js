/*
  Standalone storage layer for Punchbook.

  The app was originally built for Claude's artifact sandbox, which provides
  a `window.storage` key/value API (get/set/delete/list) backed by a remote
  service. Outside that sandbox there is no such API, so this file installs
  a drop-in replacement with the same method signatures, backed by the
  browser's localStorage instead. Every call in App.jsx (`window.storage.get`,
  `window.storage.set`, ...) keeps working unchanged.

  Signatures match the original:
    get(key, shared?)    -> { key, value, shared } | null
    set(key, value, shared?) -> { key, value, shared } | null
    delete(key, shared?)  -> { key, deleted, shared } | null
    list(prefix?, shared?) -> { keys, prefix?, shared } | null

  "shared" is accepted for API compatibility but ignored — in a plain
  localStorage-backed single-browser deployment there is only one storage
  scope. All data lives in this browser's localStorage under the
  "punchbook:" prefix.
*/

const NAMESPACE = "punchbook:";

function fullKey(key) {
  return `${NAMESPACE}${key}`;
}

function safeGetLocalStorage(key) {
  try {
    return window.localStorage.getItem(key);
  } catch (e) {
    console.error("Punchbook storage read failed:", e);
    return null;
  }
}

function safeSetLocalStorage(key, value) {
  try {
    window.localStorage.setItem(key, value);
    return true;
  } catch (e) {
    console.error("Punchbook storage write failed:", e);
    return false;
  }
}

const storage = {
  async get(key /*, shared */) {
    const raw = safeGetLocalStorage(fullKey(key));
    if (raw === null) return null;
    return { key, value: raw, shared: false };
  },

  async set(key, value /*, shared */) {
    const ok = safeSetLocalStorage(fullKey(key), value);
    if (!ok) return null;
    return { key, value, shared: false };
  },

  async delete(key /*, shared */) {
    try {
      window.localStorage.removeItem(fullKey(key));
      return { key, deleted: true, shared: false };
    } catch (e) {
      console.error("Punchbook storage delete failed:", e);
      return null;
    }
  },

  async list(prefix = "" /*, shared */) {
    try {
      const keys = [];
      const searchPrefix = fullKey(prefix);
      for (let i = 0; i < window.localStorage.length; i++) {
        const k = window.localStorage.key(i);
        if (k && k.startsWith(searchPrefix)) {
          keys.push(k.slice(NAMESPACE.length));
        }
      }
      return { keys, prefix, shared: false };
    } catch (e) {
      console.error("Punchbook storage list failed:", e);
      return null;
    }
  },
};

// Install on window so the untouched App.jsx code (window.storage.get/set/...)
// works exactly as it did in the artifact sandbox.
if (typeof window !== "undefined") {
  window.storage = storage;
}

export default storage;
