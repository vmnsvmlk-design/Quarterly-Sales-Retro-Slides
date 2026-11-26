"use strict";

const { retroPeriods } = require("./lib/fiscal");
const { loadHubSpot } = require("./lib/data");
const { loadTargets, targetFor } = require("./lib/targets");
const drill = require("./lib/drill");
const C = require("./lib/config");
const { metrics } = require("./metrics/rep");

const norm = (n) => String(n || "").trim().toLowerCase().replace(/\s+/g, " ");
const PKEYS = ["compare", "review", "combined"];

function targetsBlock(targets, name, periods) {
  const q = (quarter) => targetFor(targets, name, quarter);
  const compare = q(periods.compare), review = q(periods.review);
  const parts = [compare, review].filter((t) => t.status === "set");
  return {
    compare, review,
    combined: { value: parts.length ? parts.reduce((a, t) => a + t.value, 0) : null, status: parts.length === 2 ? "set" : parts.length ? "partial" : "tbd" },
  };
}

const attain = (achieved, target) => (target && target.value ? achieved / target.value : null);

async function buildRetro() {
  const periods = retroPeriods();
  const errors = {};
  const targets = await loadTargets();
  if (targets.error) errors.targets = targets.error;

  let hub;
  if (process.env.MOCK_DATA === "1") hub = require("./lib/mock").mockHubSpot(periods, targets);
  else hub = await loadHubSpot(periods);

  const idByName = {};
  for (const [id, name] of Object.entries(hub.ownerNames)) idByName[norm(name)] = id;

  const excluded = new Set((C.excludedAEs || []).map(norm));
  const teamOfOwner = {};
  const teams = C.teams.map((t) => {
    const reps = targets.roster.filter((r) => r.region === t.key && !excluded.has(norm(r.name))).map((r) => {
      const ownerId = r.ownerId && hub.ownerNames[r.ownerId] ? r.ownerId : idByName[norm(r.name)] || r.ownerId || null;
      const tg = targetsBlock(targets, r.name, periods);
      const per = {};
      for (const k of PKEYS) {
        const m = ownerId ? metrics(hub.deals, [ownerId], periods[k]) : metrics([], [], periods[k]);
        m.target = tg[k];
        m.attainment = attain(m.revenue.achieved, tg[k]);
        per[k] = m;
      }
      const c = per.combined;
      const active = c.revenue.won + c.mofu.mqlNew + c.mofu.sqlNew + c.bofu.inPlay > 0;
      const inRole = [tg.compare, tg.review].some((x) => x.status === "set");
      return { name: r.name, ownerId, matched: !!(ownerId && hub.ownerNames[ownerId]), periods: per, show: inRole || active, inRole, active };
    }).filter((r) => r.show).sort((a, b) => (b.periods.combined.target.value || 0) - (a.periods.combined.target.value || 0) || b.periods.combined.revenue.achieved - a.periods.combined.revenue.achieved);

    for (const r of reps) if (r.ownerId) teamOfOwner[r.ownerId] = t.key;
    const ownerIds = reps.map((r) => r.ownerId).filter(Boolean);
    const total = {};
    for (const k of PKEYS) {
      const m = metrics(hub.deals, ownerIds, periods[k]);
      const set = reps.map((r) => r.periods[k].target).filter((x) => x.value != null);
      m.target = { value: set.length ? set.reduce((a, x) => a + x.value, 0) : null, status: reps.every((r) => r.periods[k].target.status === "set" || r.periods[k].target.status === "na") ? "set" : "partial" };
      m.attainment = attain(m.revenue.achieved, m.target);
      total[k] = m;
    }
    return { key: t.key, title: t.title, reps: reps.map(({ show, ...r }) => r), total };
  });

  const allOwners = teams.flatMap((t) => t.reps.map((r) => r.ownerId)).filter(Boolean);
  const company = {};
  for (const k of PKEYS) {
    const m = metrics(hub.deals, allOwners, periods[k]);
    const tv = teams.map((t) => t.total[k].target.value).filter((v) => v != null);
    m.target = { value: tv.length ? tv.reduce((a, b) => a + b, 0) : null, status: teams.every((t) => t.total[k].target.status === "set") ? "set" : "partial" };
    m.attainment = attain(m.revenue.achieved, m.target);
    company[k] = m;
  }

  const dealsById = new Map(hub.deals.map((d) => [d.id, d]));
  const label = (q) => ({ label: q.label, short: q.short });
  return drill.finalize({
    generatedAt: new Date().toISOString(),
    mock: process.env.MOCK_DATA === "1",
    periods: { compare: label(periods.compare), review: label(periods.review), combined: label(periods.combined) },
    targets: { source: targets.source, asOf: targets.asOf, error: targets.error || null },
    teams, company, errors,
  }, { ownerNames: hub.ownerNames, stageLabel: hub.stageLabel, dealsById, teamOfOwner });
}

module.exports = { buildRetro };
