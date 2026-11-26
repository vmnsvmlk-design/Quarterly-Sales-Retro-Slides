"use strict";

/**
 * Insights typed on the slides, shared by everyone who opens the deck, kept
 * per retro (e.g. "FY 26-27 H1") in one JSON file. On Railway this only
 * survives redeploys if a Volume is attached and DATA_DIR points at it.
 */
const fs = require("fs");
const path = require("path");

const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, "..", "..", "data");
const FILE = path.join(DATA_DIR, "retro-insights.json");
let cache = null;

function load() {
  if (cache) return cache;
  try { cache = JSON.parse(fs.readFileSync(FILE, "utf8")); } catch { cache = {}; }
  return cache;
}
function save() {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  const tmp = FILE + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(cache, null, 2));
  fs.renameSync(tmp, FILE);
}

function getInsights(retro) { return (load()[retro] || {}); }

function setInsight(retro, key, text, author) {
  const s = load();
  s[retro] = s[retro] || {};
  const clean = String(text || "").slice(0, 6000);
  if (!clean.trim()) delete s[retro][key];
  else s[retro][key] = { text: clean, author: String(author || "").slice(0, 80) || null, updatedAt: new Date().toISOString() };
  save();
  return s[retro][key] || { text: "", updatedAt: new Date().toISOString() };
}

module.exports = { getInsights, setInsight, DATA_DIR };
