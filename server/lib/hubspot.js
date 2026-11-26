"use strict";

/**
 * Minimal HubSpot REST client (Node 18+ fetch). Every call goes through a
 * limiter that stays under HubSpot's private-app limits (search: ~4 req/s,
 * everything else: ~9 req/s); 429s and 5xx are retried with backoff.
 */
const BASE = "https://api.hubapi.com";

class HubSpotConfigError extends Error {}

function token() {
  const t = process.env.HUBSPOT_ACCESS_TOKEN;
  if (!t) throw new HubSpotConfigError("HUBSPOT_ACCESS_TOKEN is not set.");
  return t;
}

function limiter(minGapMs) {
  let chain = Promise.resolve();
  let last = 0;
  return (fn) => {
    const run = chain.then(async () => {
      const wait = last + minGapMs - Date.now();
      if (wait > 0) await new Promise((r) => setTimeout(r, wait));
      last = Date.now();
    });
    chain = run.catch(() => {});
    return run.then(fn);
  };
}
const searchSlot = limiter(270);
const otherSlot = limiter(115);

async function request(method, path, body, { search = false } = {}) {
  const slot = search ? searchSlot : otherSlot;
  for (let attempt = 0; ; attempt++) {
    const res = await slot(() => fetch(BASE + path, {
      method,
      headers: { Authorization: `Bearer ${token()}`, "Content-Type": "application/json" },
      body: body ? JSON.stringify(body) : undefined,
    }));
    if (res.ok) return res.json();
    if ((res.status === 429 || res.status >= 500) && attempt < 6) {
      const retryAfter = Number(res.headers.get("retry-after")) || 0;
      await new Promise((r) => setTimeout(r, Math.max(retryAfter * 1000, 500 * 2 ** attempt)));
      continue;
    }
    const text = await res.text().catch(() => "");
    throw new Error(`HubSpot API error (${res.status}) on ${method} ${path}: ${text.slice(0, 300)}`);
  }
}

/**
 * Every record matching `filterGroups`. HubSpot search stops at 10,000
 * results, so when `splitProp` is given every group is bounded to
 * [from, to) on it and the window is halved until each half fits.
 */
async function searchAll(objectType, { filterGroups, properties, splitProp, from, to }) {
  const groups = splitProp
    ? filterGroups.map((g) => ({ filters: [...g.filters,
        { propertyName: splitProp, operator: "GTE", value: String(from) },
        { propertyName: splitProp, operator: "LT", value: String(to) }] }))
    : filterGroups;
  const out = [];
  let after;
  for (;;) {
    const page = await request("POST", `/crm/v3/objects/${objectType}/search`, {
      filterGroups: groups, properties, limit: 200, after,
      sorts: [{ propertyName: "hs_object_id", direction: "ASCENDING" }],
    }, { search: true });
    if (!after && page.total >= 9900 && splitProp && to - from > 86400000) {
      const mid = Math.floor((from + to) / 2);
      const a = await searchAll(objectType, { filterGroups, properties, splitProp, from, to: mid });
      const b = await searchAll(objectType, { filterGroups, properties, splitProp, from: mid, to });
      return a.concat(b);
    }
    out.push(...page.results);
    if (!page.paging || !page.paging.next) break;
    after = page.paging.next.after;
    if (out.length >= 9999) break; // hard cap; only reachable without splitProp
  }
  return out;
}

async function batchRead(objectType, ids, properties) {
  const out = {};
  const unique = [...new Set(ids.filter(Boolean).map(String))];
  for (let i = 0; i < unique.length; i += 100) {
    const res = await request("POST", `/crm/v3/objects/${objectType}/batch/read`, {
      inputs: unique.slice(i, i + 100).map((id) => ({ id })), properties,
    });
    for (const r of res.results || []) out[r.id] = r;
  }
  return out;
}

/** dealId -> primary company id (falls back to the first associated company). */
async function dealPrimaryCompanies(dealIds) {
  const out = {};
  const unique = [...new Set(dealIds.map(String))];
  for (let i = 0; i < unique.length; i += 1000) {
    const res = await request("POST", "/crm/v4/associations/deals/companies/batch/read", {
      inputs: unique.slice(i, i + 1000).map((id) => ({ id })),
    });
    for (const r of res.results || []) {
      const tos = r.to || [];
      const primary = tos.find((t) => (t.associationTypes || []).some((a) => a.typeId === 5)) || tos[0];
      if (primary) out[String(r.from.id)] = String(primary.toObjectId);
    }
  }
  return out;
}

async function pipelines() {
  const res = await request("GET", "/crm/v3/pipelines/deals");
  return res.results || [];
}

async function propertyOptions(objectType, name) {
  const res = await request("GET", `/crm/v3/properties/${objectType}/${name}`);
  const map = {};
  for (const o of res.options || []) map[o.value] = o.label;
  return map;
}

async function owners() {
  const map = {};
  for (const archived of [false, true]) {
    let after;
    for (;;) {
      const res = await request("GET", `/crm/v3/owners?limit=500&archived=${archived}${after ? `&after=${after}` : ""}`);
      for (const o of res.results || []) {
        map[String(o.id)] = `${o.firstName || ""} ${o.lastName || ""}`.trim() || o.email || String(o.id);
      }
      if (!res.paging || !res.paging.next) break;
      after = res.paging.next.after;
    }
  }
  return map;
}

module.exports = { searchAll, batchRead, dealPrimaryCompanies, pipelines, propertyOptions, owners, HubSpotConfigError };
