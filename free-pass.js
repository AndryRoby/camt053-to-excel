// free-pass.js: one free MT940 export and one free DATEV Buchungsstapel
// export per browser, no account, no licence, so a visitor can prove to
// themselves that the file really imports into DATEV before paying for
// Pro. See index.html (the MT940/DATEV download buttons) for how this
// gates ahead of the licence check.
//
// Pure, dependency-free, testable under Node (tests.mjs imports it
// directly): every localStorage access goes through the optional `store`
// parameter, so a test can pass a fake store, including one whose
// getItem/setItem throw, without touching a real browser. Same
// "wrap every localStorage call in try/catch" convention as licence.js
// and pro.js in this repo.
//
// Storage: localStorage key arling_free_export holds a small object
// {"mt940": "<ISO date>"|null, "datev": "<ISO date>"|null}. A missing
// key, malformed JSON, a missing localStorage, or a store that throws on
// read all resolve to "no export used yet" -- we would rather hand out a
// second free export (a private window included) than wrongly block a
// real buyer on their one chance to try the tool.

export const FREE_EXPORT_KEY = 'arling_free_export';
export const FREE_EXPORT_KINDS = ['mt940', 'datev'];

function hasRealLocalStorage() {
  try {
    return typeof localStorage !== 'undefined' && localStorage !== null;
  } catch (e) {
    return false;
  }
}

/** Resolves the effective store: the explicit `store` argument when the
 * caller passed one (including `null`, meaning "no storage"), otherwise
 * the real browser localStorage when available. */
function resolveStore(store) {
  if (store !== undefined) return store;
  return hasRealLocalStorage() ? localStorage : null;
}

/** @returns {Object} the parsed state object, or {} for anything that
 *   cannot be read (no store, throwing store, malformed JSON). */
function readState(store) {
  try {
    const s = resolveStore(store);
    if (!s) return {};
    const raw = s.getItem(FREE_EXPORT_KEY);
    if (!raw) return {};
    const obj = JSON.parse(raw);
    return obj && typeof obj === 'object' ? obj : {};
  } catch (e) {
    return {};
  }
}

/**
 * @param {'mt940'|'datev'} kind
 * @param {Storage|null} [store] Defaults to the real localStorage; pass a
 *   fake (or `null`, or a store whose methods throw) to test without one.
 * @returns {boolean} true only when this kind's free export is on record
 *   as already used. Any doubt (no storage, a throwing store, malformed
 *   state) resolves to false, i.e. "free still available".
 */
export function hasUsedFree(kind, store) {
  const state = readState(store);
  return typeof state[kind] === 'string' && state[kind].length > 0;
}

/**
 * Records that `kind`'s free export was just used. Never throws: when the
 * store cannot persist the flag the download still went ahead (the caller
 * downloads before or regardless of this call's result), so the visitor
 * may simply get more than one free export -- the safe side to fail on.
 * @param {'mt940'|'datev'} kind
 * @param {Storage|null} [store] Same convention as hasUsedFree.
 * @param {Date} [now] Defaults to `new Date()`; a test can pin it.
 * @returns {boolean} true if the flag was actually persisted.
 */
export function markFreeUsed(kind, store, now) {
  try {
    const s = resolveStore(store);
    if (!s) return false;
    const state = readState(store);
    state[kind] = (now instanceof Date ? now : new Date()).toISOString();
    s.setItem(FREE_EXPORT_KEY, JSON.stringify(state));
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * @param {Storage|null} [store] Same convention as hasUsedFree.
 * @returns {{mt940:boolean, datev:boolean}} which kinds still have a free
 *   export available (the mirror image of hasUsedFree per kind).
 */
export function freeRemaining(store) {
  const state = readState(store);
  const out = {};
  FREE_EXPORT_KINDS.forEach((kind) => {
    out[kind] = !(typeof state[kind] === 'string' && state[kind].length > 0);
  });
  return out;
}

// Also expose as a plain browser global when loaded via <script type="module">.
if (typeof window !== 'undefined') {
  window.ArlingFreePass = { FREE_EXPORT_KEY, FREE_EXPORT_KINDS, hasUsedFree, markFreeUsed, freeRemaining };
}
