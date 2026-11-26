"use strict";

/**
 * MOCK_DATA=1: synthetic deals for the roster's AEs, so the layout, click-through
 * lists and .pptx export can be tried without a HubSpot token. Never used otherwise.
 */
const C = require("./config");
const S = C.stages;
const DAY = 86400000;

function rng(seed) { let s = seed >>> 0; return () => ((s = (s * 1664525 + 1013904223) >>> 0) / 4294967296); }

function mockHubSpot(periods, targets) {
  const r = rng(42);
  const ownerNames = {};
  const deals = [];
  const start = periods.combined.start.getTime() - 120 * DAY, end = periods.combined.end.getTime();
  const acv = { NAM: 50000, EMEA: 28000, APAC: 15000 };
  const regionRaw = { NAM: "Americas", EMEA: "EMEA", APAC: "APAC" };
  const sources = [["Inbound", "Pure Inbound"], ["Outbound", "Pure Outbound"], ["Partnership-Event", "Allbound"], ["Events", "Allbound"]];
  let n = 1;
  for (const a of targets.roster) {
    ownerNames[a.ownerId] = a.name;
    const count = 25 + Math.floor(r() * 35);
    for (let i = 0; i < count; i++) {
      const mqlAt = start + r() * (end - start);
      const src = sources[Math.floor(r() * sources.length)];
      const deal = { id: String(9000000 + n++), name: `Mock Co ${n} - ${a.name.split(" ")[0]}`, ownerId: a.ownerId, regionRaw: r() < 0.85 ? regionRaw[a.region] : "EMEA",
        leadSource: src[0], finalLeadSource: src[1], mqlAt, nqAt: null, sqlAt: null, closeAt: null, won: false, isClosed: false, revenue: 0, stage: S.scheduled, exit: null, sqlStage: false };
      const x = r();
      if (x < 0.25) { deal.stage = S.notQualified; deal.nqAt = mqlAt + r() * 20 * DAY; deal.isClosed = false; }
      else if (x < 0.62) {
        deal.sqlAt = mqlAt + (3 + r() * 25) * DAY;
        const y = r();
        if (y < 0.28) { deal.won = true; deal.isClosed = true; deal.stage = S.closedWon; deal.closeAt = deal.sqlAt + (20 + r() * 140) * DAY; deal.revenue = Math.round(acv[a.region] * (0.4 + r() * 1.4) / 100) * 100; }
        else if (y < 0.65) { deal.isClosed = true; deal.stage = S.closedLost; deal.closeAt = deal.sqlAt + (10 + r() * 120) * DAY; }
        else { deal.stage = [S.qualifiedToBuy, S.presentationDone, S.proposalShared, S.contracting][Math.floor(r() * 4)]; deal.closeAt = end + r() * 90 * DAY; }
        if (deal.closeAt && deal.closeAt > Date.now() && deal.isClosed) { deal.isClosed = false; deal.won = false; deal.revenue = 0; deal.stage = S.proposalShared; }
        deal.sqlStage = C.sqlStageIds.includes(deal.stage);
        if (deal.won) deal.exit = { type: "won", at: deal.closeAt };
        else if (deal.stage === S.closedLost) deal.exit = { type: "lost", at: deal.closeAt };
      } else if (x < 0.7) { deal.stage = S.salesAccepted; }
      if (deal.sqlAt && deal.sqlAt > Date.now()) continue;
      deals.push(deal);
    }
  }
  const stageLabel = { [S.scheduled]: "Scheduled", [S.notQualified]: "Not Qualified", [S.salesAccepted]: "Sales Accepted Account", [S.qualifiedToBuy]: "Qualified to buy",
    [S.presentationDone]: "Presentation done", [S.proposalShared]: "Proposal Shared / Negotiations", [S.contracting]: "Contracting", [S.closedWon]: "Closed won", [S.closedLost]: "Closed lost" };
  return { ownerNames, stageLabel, deals };
}

module.exports = { mockHubSpot };
