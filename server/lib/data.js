"use strict";

/**
 * One HubSpot pull for the whole deck. Every deal comes back with the same
 * properties and is normalised into one record, so every metric (and every
 * click-through list) reads from the same set:
 *
 *  1. SQLs: Sales Pipeline deals with an SQL date that were still open at, or
 *     closed after, the start of the retro (BOFU, SQLs produced).
 *  2. Closed won: Sales Pipeline deals closed won from the start of the retro
 *     (revenue, ACV, sales cycle).
 *  3. MQLs: Sales Pipeline deals that entered Scheduled from two years before
 *     the retro on (MOFU, incl. the opening balance).
 */
const hs = require("./hubspot");
const C = require("./config");

const D = C.deal, S = C.stages;
const DAY = 86400000;
const ms = (v) => { if (v == null || v === "") return null; const t = new Date(isNaN(v) ? v : Number(v)).getTime(); return Number.isFinite(t) ? t : null; };
const SQL_STAGES = new Set(C.sqlStageIds);

/**
 * Normalises a HubSpot deal. exit (for deals with an SQL date):
 *  won  = closed won (incl. Deal Desk Verified), at the close date;
 *  lost = closed lost, debooked, or slipped back to a pre-SQL stage (at the
 *         date it moved; a deal closed lost and later moved back keeps its loss date).
 */
function normalise(d) {
  const p = d.properties || {};
  const stage = p[D.stage];
  const sqlAt = ms(p[D.sqlDate]);
  const closeAt = ms(p[D.closeDate]);
  const won = p[D.isClosedWon] === "true";
  const isClosed = p[D.isClosed] === "true";
  let exit = null;
  if (sqlAt != null) {
    if (won) exit = { type: "won", at: closeAt };
    else if (stage === S.closedLost) exit = { type: "lost", at: closeAt };
    else if (!SQL_STAGES.has(stage) || stage === S.debook) {
      const lostAt = ms(p[C.dateEntered(S.closedLost)]);
      const movedAt = ms(p[C.dateEntered(stage)]) || ms(p.hs_lastmodifieddate);
      const priorLoss = lostAt != null && lostAt >= sqlAt && (movedAt == null || lostAt <= movedAt);
      exit = { type: "lost", slipBack: stage !== S.debook, debook: stage === S.debook, at: priorLoss ? lostAt : movedAt };
    }
    if (exit && (exit.at == null || exit.at < sqlAt)) exit.at = exit.at == null ? ms(p.hs_lastmodifieddate) : sqlAt;
  }
  return {
    id: String(d.id), name: p[D.name] || "(unnamed)", ownerId: p[D.owner] ? String(p[D.owner]) : null,
    stage, regionRaw: p[D.region] || null, leadSource: p[D.leadSource] || null, finalLeadSource: p[D.finalLeadSource] || null,
    sqlAt, closeAt, won, isClosed, revenue: Number(p[D.revenue]) || 0,
    mqlAt: ms(p[C.dateEntered(S.scheduled)]), nqAt: ms(p[C.dateEntered(S.notQualified)]),
    exit, sqlStage: SQL_STAGES.has(stage),
  };
}

/**
 * Assumption (same as the QBR deck): a closed-won deal with no SQL date became
 * an SQL in the quarter it closed. It counts as an SQL produced then (and in
 * BOFU), but is left out of the sales-cycle average.
 */
function addAssumedSqls(deals) {
  for (const d of deals) {
    if (d.won && d.sqlAt == null && d.closeAt != null) {
      d.sqlAtAssumed = d.closeAt;
      d.exit = { type: "won", at: d.closeAt };
    }
  }
}

async function loadHubSpot(periods) {
  const from = periods.combined.start.getTime();
  const now = Date.now();
  const pipes = await hs.pipelines();
  const stageLabel = {};
  for (const p of pipes) for (const s of p.stages || []) stageLabel[s.id] = s.label;
  const sales = pipes.find((p) => p.id === C.salesPipelineId) || { stages: [] };
  const stageIds = [...new Set((sales.stages || []).map((s) => s.id).concat(Object.values(S)))];
  const properties = [D.name, D.stage, D.owner, D.region, D.leadSource, D.finalLeadSource, D.sqlDate, D.closeDate,
    D.isClosedWon, D.isClosed, D.revenue, "hs_lastmodifieddate", ...stageIds.map(C.dateEntered)];
  const inSales = { propertyName: D.pipeline, operator: "EQ", value: C.salesPipelineId };
  const far = now + 5 * 365 * DAY;

  const [ownerNames, closedAfter, openSqls, won, mqls] = await Promise.all([
    hs.owners(),
    hs.searchAll("deals", { properties, filterGroups: [{ filters: [inSales, { propertyName: D.sqlDate, operator: "HAS_PROPERTY" }] }], splitProp: D.closeDate, from, to: far }),
    hs.searchAll("deals", { properties, filterGroups: [{ filters: [inSales, { propertyName: D.sqlDate, operator: "HAS_PROPERTY" }, { propertyName: D.isClosed, operator: "EQ", value: "false" }] }], splitProp: D.sqlDate, from: 0, to: now + DAY }),
    hs.searchAll("deals", { properties, filterGroups: [{ filters: [inSales, { propertyName: D.isClosedWon, operator: "EQ", value: "true" }] }], splitProp: D.closeDate, from, to: now + DAY }),
    hs.searchAll("deals", { properties, filterGroups: [{ filters: [inSales] }], splitProp: C.dateEntered(S.scheduled), from: from - C.mqlLookbackDays * DAY, to: now + DAY }),
  ]);
  const byId = new Map();
  for (const d of [...closedAfter, ...openSqls, ...won, ...mqls]) byId.set(String(d.id), d);
  const deals = [...byId.values()].map(normalise);
  addAssumedSqls(deals);
  return { ownerNames, stageLabel, deals };
}

module.exports = { loadHubSpot, normalise, addAssumedSqls };
