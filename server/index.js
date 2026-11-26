"use strict";

require("dotenv").config();
const path = require("path");
const crypto = require("crypto");
const express = require("express");
const { cached, peek } = require("./lib/cache");
const { buildRetro } = require("./retro");
const store = require("./lib/store");
const drill = require("./lib/drill");

const app = express();
const PORT = Number(process.env.PORT) || 4000;
const TTL_MS = 20 * 60 * 1000;

// Login (HTTP Basic Auth) - on only when both RETRO_USERNAME and RETRO_PASSWORD are set.
const USER = process.env.RETRO_USERNAME || "";
const PASS = process.env.RETRO_PASSWORD || "";
const digest = (s) => crypto.createHash("sha256").update(String(s)).digest();
if (USER && PASS) {
  app.use((req, res, next) => {
    const [scheme, encoded] = (req.headers.authorization || "").split(" ");
    if (scheme === "Basic" && encoded) {
      const raw = Buffer.from(encoded, "base64").toString("utf8");
      const i = raw.indexOf(":");
      if (i > -1 && crypto.timingSafeEqual(digest(raw.slice(0, i)), digest(USER)) && crypto.timingSafeEqual(digest(raw.slice(i + 1)), digest(PASS))) return next();
    }
    res.set("WWW-Authenticate", 'Basic realm="SpotDraft Sales Retro"');
    res.status(401).send("Authentication required.");
  });
} else if (process.env.RAILWAY_ENVIRONMENT) {
  console.warn("RETRO_USERNAME / RETRO_PASSWORD are not set - this deck is open to anyone with the URL.");
}

app.use(express.json({ limit: "100kb" }));
app.use(express.static(path.join(__dirname, "..", "public")));
app.get("/vendor/pptxgen.bundle.js", (req, res) => res.sendFile(require.resolve("pptxgenjs/dist/pptxgen.bundle.js")));

app.get("/api/health", (req, res) => {
  const last = peek("retro");
  res.json({
    ok: true,
    hubspotConfigured: !!process.env.HUBSPOT_ACCESS_TOKEN,
    regionalInterlocksConfigured: !!process.env.REGIONAL_INTERLOCKS_URL,
    mock: process.env.MOCK_DATA === "1",
    dataAsOf: last ? new Date(last.at).toISOString() : null,
  });
});

// Served from a warmed in-memory copy (refreshed every 20 minutes); ?refresh=1 forces a fresh pull.
app.get("/api/retro", async (req, res) => {
  try {
    res.json(await cached("retro", TTL_MS, buildRetro, { force: req.query.refresh === "1" }));
  } catch (err) {
    console.error("[retro] build failed:", err);
    res.status(502).json({ error: String(err.message || err) });
  }
});

// Click-through: the deals behind any number on the deck.
app.get("/api/drill/:key", (req, res) => {
  const out = drill.details(req.params.key);
  if (!out) return res.status(404).json({ error: "This list is no longer available - reload the page to get the latest data." });
  res.json(out);
});

// Insights (editable on every slide), per retro.
const retroOf = (req) => String(req.query.retro || "").slice(0, 60);
app.get("/api/insights", (req, res) => res.json(store.getInsights(retroOf(req))));
app.put("/api/insights/:key", (req, res) => {
  const b = req.body || {};
  res.json(store.setInsight(retroOf(req), String(req.params.key).slice(0, 120), b.text, b.author));
});

app.listen(PORT, () => {
  console.log(`SpotDraft Sales Retro running on http://localhost:${PORT}`);
  const warm = () => cached("retro", TTL_MS, buildRetro, { force: true }).then(() => console.log("[retro] data warmed")).catch((e) => console.error("[retro] warm-up failed:", e.message));
  warm();
  setInterval(warm, TTL_MS).unref();
});
