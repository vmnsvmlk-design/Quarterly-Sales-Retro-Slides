"use strict";

// One in-memory copy per key; concurrent callers share the same in-flight promise.
const store = new Map();

async function cached(key, ttlMs, fn, { force = false } = {}) {
  const hit = store.get(key);
  const now = Date.now();
  if (!force && hit && hit.value && now - hit.at < ttlMs) return hit.value;
  if (hit && hit.pending) return hit.pending;
  const pending = fn().then(
    (value) => { store.set(key, { value, at: Date.now() }); return value; },
    (err) => { store.set(key, hit && hit.value ? { value: hit.value, at: hit.at } : {}); throw err; }
  );
  store.set(key, { ...(hit || {}), pending });
  return pending;
}

function peek(key) {
  const hit = store.get(key);
  return hit && hit.value ? { value: hit.value, at: hit.at } : null;
}

module.exports = { cached, peek };
