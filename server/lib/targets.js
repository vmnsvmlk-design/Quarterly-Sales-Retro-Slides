"use strict";

/**
 * AE roster, team and New Business ARR targets.
 *
 * Live source: the Regional Interlocks dashboard's Inputs tab (GET /api/config,
 * server to server, its login never reaches the browser). Its shape:
 *   roster:         [{ ownerId, name, role: "AE" | "SDR", region: "NAM" | "EMEA" | "APAC" }]
 *   revenueTargets: { "<AE name>": { Q1, Q2, Q3, Q4 } }  (number | null = "TBD" | "N/A")
 * Quarters there are FY 26-27 Q1-Q4.
 *
 * If that dashboard can't be reached, the snapshot below is used (the Inputs tab
 * as of 26 Nov 2026) and the deck says so.
 */
const SNAPSHOT = {
  asOf: "2026-11-26",
  fy: "FY 26-27",
  roster: [
    { ownerId: 633725813, name: "Aadarsh Goyal", region: "APAC", t: [43273, 57000, 37350, 45000] },
    { ownerId: 95241225, name: "Aditya Deepak", region: "NAM", t: [174700, 182300, 260245, 202754] },
    { ownerId: 2038929718, name: "Alen Alex Joseph", region: "APAC", t: [39273, 53500, 37350, 55000] },
    { ownerId: 79976078, name: "Allison Green", region: "NAM", t: [318800, 332700, 500000, 325000] },
    { ownerId: 95421364, name: "Andrew Ketler", region: "NAM", t: [null, null, 360000, 360000] },
    { ownerId: 97858519, name: "Andrew Lomanto", region: "NAM", t: [null, null, 250000, 250000] },
    { ownerId: 80103683, name: "Elizabeth Doolittle", region: "NAM", t: [260900, 272200, null, null] },
    { ownerId: 327954719, name: "Gaurav Gururaj", region: "APAC", t: [47727, 70000, 55000, 100000] },
    { ownerId: 721568845, name: "Harish Murali", region: "APAC", t: [34036, null, null, null] },
    { ownerId: 1936367598, name: "John Manackal", region: "NAM", t: [133300, 139100, 222611, 156404] },
    { ownerId: 598847414, name: "Marshall Deutz", region: "NAM", t: [274500, null, null, null] },
    { ownerId: 101025397, name: "Myla Cox", region: "NAM", t: [null, null, 90000, 90000] },
    { ownerId: 102223703, name: "Patrick Conley", region: "NAM", t: [null, null, 100000, 100000] },
    { ownerId: 90053101, name: "Pranav Chandna", region: "APAC", t: [46080, 65000, 47500, 55000] },
    { ownerId: 81398225, name: "Sabrina Pervez", region: "EMEA", t: [134000, 329174, 165000, 302785] },
    { ownerId: 609609953, name: "Sanjiv Nidamboor", region: "EMEA", t: [null, null, null, null] },
    { ownerId: 233546080, name: "Viraj Rajeev", region: "EMEA", t: [102000, 170826, 165000, 315215] },
    { ownerId: 415418450, name: "Yash Khanna", region: "APAC", t: [49077, 65000, 55000, 100000] },
  ],
};

function basicAuth(user, pass) {
  return user && pass ? { Authorization: "Basic " + Buffer.from(`${user}:${pass}`).toString("base64") } : {};
}

function fromSnapshot() {
  const roster = SNAPSHOT.roster.map((r) => ({ ownerId: String(r.ownerId), name: r.name, region: r.region }));
  const revenueTargets = {};
  for (const r of SNAPSHOT.roster) revenueTargets[r.name] = { Q1: r.t[0], Q2: r.t[1], Q3: r.t[2], Q4: r.t[3] };
  return { source: "snapshot", asOf: SNAPSHOT.asOf, fy: SNAPSHOT.fy, roster, revenueTargets };
}

async function fetchLive() {
  const base = process.env.REGIONAL_INTERLOCKS_URL;
  if (!base) throw new Error("REGIONAL_INTERLOCKS_URL is not set.");
  const res = await fetch(base.replace(/\/+$/, "") + "/api/config", {
    headers: { Accept: "application/json", ...basicAuth(process.env.REGIONAL_INTERLOCKS_USERNAME, process.env.REGIONAL_INTERLOCKS_PASSWORD) },
  });
  if (res.status === 401) throw new Error("Regional Interlocks rejected the login (401) - check REGIONAL_INTERLOCKS_USERNAME / PASSWORD.");
  if (!res.ok) throw new Error(`Regional Interlocks returned ${res.status} for /api/config.`);
  const json = await res.json();
  const roster = (json.roster || []).filter((r) => r && r.name && r.role === "AE")
    .map((r) => ({ ownerId: r.ownerId ? String(r.ownerId) : null, name: r.name, region: r.region || null }));
  if (!roster.length) throw new Error("Regional Interlocks returned an empty AE roster.");
  return { source: "live", asOf: new Date().toISOString().slice(0, 10), fy: "FY 26-27", roster, revenueTargets: json.revenueTargets || {} };
}

let lastGood = null;
/** { source, asOf, fy, roster, revenueTargets, error? } - never throws. */
async function loadTargets() {
  if (process.env.MOCK_DATA === "1") return fromSnapshot();
  try {
    lastGood = await fetchLive();
    return lastGood;
  } catch (err) {
    console.warn("[targets]", err.message, lastGood ? "Using the last good copy." : "Using the built-in snapshot.");
    return { ...(lastGood || fromSnapshot()), error: err.message };
  }
}

/**
 * Target for one AE and quarter. Returns { value, status }:
 *  status "set" (a number), "tbd" (blank in Inputs), "na" (not in the role that quarter),
 *  "other-fy" (the Inputs tab only holds one fiscal year).
 */
function targetFor(targets, name, quarter) {
  if (quarter.fyLabel !== targets.fy) return { value: null, status: "other-fy" };
  const row = targets.revenueTargets[name];
  const v = row ? row[quarter.short] : null;
  if (v === "N/A") return { value: null, status: "na" };
  if (v === null || v === undefined || v === "") return { value: null, status: "tbd" };
  const n = Number(v);
  return Number.isFinite(n) ? { value: n, status: "set" } : { value: null, status: "tbd" };
}

module.exports = { loadTargets, targetFor, SNAPSHOT };
