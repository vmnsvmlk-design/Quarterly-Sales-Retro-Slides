"use strict";
// Hand-built deals with known answers: run with `npm test`.
const assert = require("assert");
const { normalise, addAssumedSqls } = require("../server/lib/data");
const { metrics } = require("../server/metrics/rep");
const { parseLabel } = require("../server/lib/fiscal");
const C = require("../server/lib/config");
const S = C.stages, D = C.deal;

const deal = (id, owner, p) => ({ id, properties: Object.assign({ [D.owner]: owner, [D.name]: "Deal " + id }, p) });
const t = (s) => new Date(s + "T12:00:00Z").toISOString();
const raw = [
  // MQL in Q2 -> SQL in Q2 -> won in Q2 ($10K, 41-day cycle)
  deal("1", "A", { [D.stage]: S.closedWon, [D.isClosedWon]: "true", [D.isClosed]: "true", [C.dateEntered(S.scheduled)]: t("2026-08-05"), [D.sqlDate]: t("2026-08-20"), [D.closeDate]: t("2026-09-30"), [D.revenue]: "10000" }),
  // MQL in Q1, still an MQL at the start of Q2, SQL in Q2, open
  deal("2", "A", { [D.stage]: S.qualifiedToBuy, [D.isClosedWon]: "false", [D.isClosed]: "false", [C.dateEntered(S.scheduled)]: t("2026-07-10"), [D.sqlDate]: t("2026-08-10"), [D.closeDate]: t("2027-01-30") }),
  // MQL in Q2, not qualified
  deal("3", "A", { [D.stage]: S.notQualified, [D.isClosedWon]: "false", [D.isClosed]: "false", [C.dateEntered(S.scheduled)]: t("2026-08-10"), [C.dateEntered(S.notQualified)]: t("2026-08-15") }),
  // SQL in Q1, open at the start of Q2, closed lost in Q2
  deal("4", "A", { [D.stage]: S.closedLost, [D.isClosedWon]: "false", [D.isClosed]: "true", [D.sqlDate]: t("2026-06-01"), [D.closeDate]: t("2026-08-15") }),
  // Closed won in Q2 with no SQL date ($5K): assumed SQL in Q2, left out of the cycle
  deal("5", "A", { [D.stage]: S.closedWon, [D.isClosedWon]: "true", [D.isClosed]: "true", [D.closeDate]: t("2026-10-01"), [D.revenue]: "5000" }),
  // Another owner: ignored
  deal("6", "B", { [D.stage]: S.closedWon, [D.isClosedWon]: "true", [D.isClosed]: "true", [D.sqlDate]: t("2026-08-02"), [D.closeDate]: t("2026-08-20"), [D.revenue]: "99000" }),
  // SQL date in Q2 but slipped back to Sales Accepted: not an SQL (current stage rule)
  deal("7", "A", { [D.stage]: S.salesAccepted, [D.isClosedWon]: "false", [D.isClosed]: "false", [D.sqlDate]: t("2026-08-02"), [C.dateEntered(S.salesAccepted)]: t("2026-09-01") }),
];
const deals = raw.map(normalise);
addAssumedSqls(deals);
const Q2 = parseLabel("FY 26-27 Q2");
const m = metrics(deals, ["A"], Q2);

assert.strictEqual(m.revenue.won, 2);
assert.strictEqual(m.revenue.achieved, 15000);
assert.strictEqual(m.revenue.acv, 7500);
assert.strictEqual(m.revenue.cycle, 41);
assert.strictEqual(m.revenue.noSqlDate, 1);
assert.strictEqual(m.mofu.mqlNew, 2);
assert.strictEqual(m.mofu.mqlNewToSql, 1);
assert.strictEqual(m.mofu.funnel, 0.5);
assert.strictEqual(m.mofu.mqlOpen, 1);
assert.strictEqual(m.mofu.sqlNew, 3);
assert.strictEqual(m.mofu.production, 1);
assert.strictEqual(m.bofu.sqlOpen, 1);
assert.strictEqual(m.bofu.inPlay, 4);
assert.strictEqual(m.bofu.won, 2);
assert.strictEqual(m.bofu.lost, 1);
assert.strictEqual(m.bofu.open, 1);
assert.strictEqual(m.bofu.winRate, 0.5);
assert.strictEqual(Math.round(m.bofu.funnel * 100), 67);
// Q1: deal 2 was an MQL created in Q1, deal 4 an SQL produced in Q1 (still open at the end of Q1).
const q1 = metrics(deals, ["A"], parseLabel("FY 26-27 Q1"));
assert.strictEqual(q1.mofu.mqlNew, 1);
assert.strictEqual(q1.mofu.sqlNew, 1);
assert.strictEqual(q1.bofu.open, 1);
console.log("metrics tests passed");
