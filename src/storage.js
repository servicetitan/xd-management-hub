import { ref, get, set, remove, onValue } from "firebase/database";
import { db } from "./firebase";

// Firebase keys can't contain dots/hyphens in some contexts — normalize them
const toDbKey = (key) => key.replace(/-/g, "_");

/**
 * Unified storage adapter.
 * - shared=false  →  localStorage  (per-browser session data, e.g. current user)
 * - shared=true   →  Firebase RTDB (team-shared data, e.g. projects & settings)
 */
export const storage = {
  get: async (key, shared = false) => {
    if (shared && db) {
      const snapshot = await get(ref(db, toDbKey(key)));
      return snapshot.exists() ? { value: snapshot.val() } : null;
    }
    const val = localStorage.getItem(key);
    return val !== null ? { value: val } : null;
  },

  set: async (key, value, shared = false) => {
    if (shared && db) {
      await set(ref(db, toDbKey(key)), value);
    } else {
      localStorage.setItem(key, value);
    }
  },

  delete: async (key) => {
    localStorage.removeItem(key);
  },
};

/**
 * Subscribe to a shared (Firebase) key with a callback.
 * Returns an unsubscribe function. Falls back to no-op if Firebase is unavailable.
 */
export function subscribeToKey(key, callback) {
  if (!db) return () => {};
  return onValue(ref(db, toDbKey(key)), (snapshot) => {
    if (snapshot.exists()) callback(snapshot.val());
  });
}
