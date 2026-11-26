"use strict";

/**
 * Click-through lists. A metric wraps the deals behind a number with
 * deals(ids, tags); finalize() swaps each wrapper for a short key, and
 * GET /api/drill/:key returns the rows from the same in-memory pull the
 * number was computed from (so the list always adds up to the number).
 * tags = { dealId: "how this deal counts in the metric" }.
 */
const crypto = require("crypto");

const deals = (ids, tags) => {
  const list = [...new Set((ids || []).filter(Boolean).map(String))];
  const t = {};
  if (tags) for (const id of list) if (tags[id]) t[id] = tags[id];
  return { __drill: "deals", ids: list, tags: t };
};

let registry = new Map();
let context = { ownerNames: {}, stageLabel: {}, dealsById: new Map(), teamOfOwner: {} };

function finalize(payload, ctx) {
  const next = new Map();
  const walk = (v) => {
    if (Array.isArray(v)) return v.map(walk);
    if (v && typeof v === "object") {
      if (v.__drill) {
        const sig = v.ids.slice().sort().map((id) => id + "=" + (v.tags[id] || "")).join(",");
        const key = "d" + crypto.createHash("sha1").update(sig).digest("hex").slice(0, 16);
        next.set(key, { ids: v.ids, tags: v.tags });
        return key;
      }
      const out = {};
      for (const [k, x] of Object.entries(v)) out[k] = walk(x);
      return out;
    }
    return v;
  };
  const result = walk(payload);
  registry = next;
  if (ctx) context = ctx;
  return result;
}

const day = (t) => (t == null ? "" : new Date(t).toISOString().slice(0, 10));
const owner = (id) => (id ? context.ownerNames[String(id)] || "Owner " + id : "");

function details(key) {
  const entry = registry.get(key);
  if (!entry) return null;
  const hasTags = Object.keys(entry.tags).length > 0;
  const portal = process.env.HUBSPOT_PORTAL_ID || "7712708";
  const rows = entry.ids.map((id) => context.dealsById.get(id)).filter(Boolean).map((d) => {
    const sql = d.sqlAt != null ? d.sqlAt : d.sqlAtAssumed;
    const cycle = d.won && d.sqlAt != null && d.closeAt != null && d.closeAt >= d.sqlAt ? Math.round((d.closeAt - d.sqlAt) / 86400000) : null;
    const cells = [d.name, owner(d.ownerId), context.teamOfOwner[d.ownerId] || "", d.regionRaw || "", day(d.mqlAt),
      day(sql) + (d.sqlAt == null && d.sqlAtAssumed != null ? " (assumed)" : ""), day(d.closeAt),
      context.stageLabel[d.stage] || d.stage || "", d.revenue || null, cycle, d.leadSource || "", d.finalLeadSource || ""];
    if (hasTags) cells.push(entry.tags[d.id] || "");
    return { id: d.id, url: portal ? `https://app.hubspot.com/contacts/${portal}/record/0-3/${d.id}` : null, cells };
  });
  rows.sort((a, b) => (hasTags ? String(a.cells[12]).localeCompare(String(b.cells[12])) : 0) || (b.cells[8] || 0) - (a.cells[8] || 0) || String(a.cells[0]).localeCompare(String(b.cells[0])));
  const columns = ["Deal name", "Deal owner", "Owner's team", "Deal region", "MQL date (entered Scheduled)", "SQL date", "Close date", "Deal stage", "Annual Subscription Fee $", "Days SQL → close", "Lead source", "Final lead source"];
  if (hasTags) columns.push("Counts in this metric as");
  return { kind: "deals", total: rows.length, columns, moneyCol: 8, rows };
}

module.exports = { deals, finalize, details };
