"use strict";

/**
 * Every AE metric in the deck, for one set of owners (one AE, or a whole team)
 * and one period (a quarter, or both quarters together). Deals are attributed
 * by DEAL OWNER only: an AE's deals all count for the AE's team, whatever the
 * deal's own region.
 *
 *  Revenue achieved  = Annual Subscription Fee $ of deals closed won in the period.
 *  ACV               = revenue achieved ÷ deals won.
 *  Sales cycle       = average days from SQL date to close date, deals won in the
 *                      period that have an SQL date.
 *  MOFU (MQL → SQL)
 *    MQL             = deal entered Scheduled.
 *    SQL             = SQL date in the period and current stage Qualified to buy or later
 *                      (incl. closed won / lost, Deal Desk Verified, Debook); a closed-won
 *                      deal with no SQL date counts as an SQL in the period it closed.
 *    Funnel %        = MQLs created in the period that became SQLs in it ÷ MQLs created in it.
 *    Production %    = SQLs produced in the period ÷ (MQLs open at the start + MQLs created).
 *  BOFU (SQL → Closed won)
 *    In play         = SQLs open at the start + SQLs produced in the period.
 *    Win rate        = in-play SQLs won in the period ÷ SQLs in play.
 *    Funnel win %    = SQLs produced in the period and won in it ÷ SQLs produced.
 */
const { inRange } = require("../lib/fiscal");
const { deals: drill } = require("../lib/drill");
const C = require("../lib/config");

const DAY = 86400000;
const ratio = (n, d) => (d ? n / d : null);
const sqlOf = (d) => (d.sqlAt != null ? d.sqlAt : d.sqlAtAssumed);
const isSqlIn = (d, P) => (d.sqlAt != null ? inRange(d.sqlAt, P) && d.sqlStage : d.sqlAtAssumed != null && inRange(d.sqlAtAssumed, P));
const tagMap = (pairs) => { const t = {}; for (const [list, label] of pairs) for (const d of list) t[d.id] = t[d.id] ? t[d.id] + " + " + label : label; return t; };
const ids = (list) => list.map((d) => d.id);
const union = (...lists) => { const m = new Map(); for (const l of lists) for (const d of l) m.set(d.id, d); return [...m.values()]; };

function metrics(allDeals, ownerIds, P) {
  const s = P.start.getTime();
  const owners = new Set([...ownerIds].map(String));
  const mine = allDeals.filter((d) => d.ownerId && owners.has(d.ownerId));

  // ---- revenue, ACV, sales cycle
  const won = mine.filter((d) => d.won && inRange(d.closeAt, P));
  const arr = won.reduce((a, d) => a + d.revenue, 0);
  const cycleDeals = won.filter((d) => d.sqlAt != null && d.closeAt != null && d.closeAt >= d.sqlAt);
  const noSqlDate = won.filter((d) => d.sqlAt == null);
  const cycle = cycleDeals.length ? cycleDeals.reduce((a, d) => a + (d.closeAt - d.sqlAt) / DAY, 0) / cycleDeals.length : null;

  // ---- MOFU
  const lookback = s - C.mqlLookbackDays * DAY;
  const mqlNew = mine.filter((d) => inRange(d.mqlAt, P));
  const mqlNewToSql = mqlNew.filter((d) => isSqlIn(d, P));
  const mqlOpen = mine.filter((d) => d.mqlAt != null && d.mqlAt < s && d.mqlAt >= lookback
    && !(sqlOf(d) != null && sqlOf(d) < s) && !(d.nqAt != null && d.nqAt < s) && !(d.isClosed && d.closeAt != null && d.closeAt < s));
  const sqlNew = mine.filter((d) => isSqlIn(d, P));

  // ---- BOFU
  const sqlOpen = mine.filter((d) => d.sqlAt != null && d.sqlAt < s && (!d.exit || d.exit.at == null || d.exit.at >= s));
  const inPlay = union(sqlOpen, sqlNew);
  const playWon = inPlay.filter((d) => d.exit && d.exit.type === "won" && inRange(d.exit.at, P));
  const playLost = inPlay.filter((d) => d.exit && d.exit.type === "lost" && inRange(d.exit.at, P));
  const playOpen = inPlay.filter((d) => !d.exit || d.exit.at == null || d.exit.at >= P.end.getTime());
  const newWon = sqlNew.filter((d) => d.exit && d.exit.type === "won" && inRange(d.exit.at, P));

  const mqlBase = union(mqlOpen, mqlNew);
  return {
    revenue: {
      achieved: Math.round(arr), won: won.length, acv: won.length ? Math.round(arr / won.length) : null,
      cycle: cycle == null ? null : Math.round(cycle), cycleDeals: cycleDeals.length, noSqlDate: noSqlDate.length,
      drill: {
        won: drill(ids(won)),
        acv: drill(ids(won)),
        cycle: drill(ids(won), tagMap([[cycleDeals, "In the average"], [noSqlDate, "No SQL date: left out"]])),
      },
    },
    mofu: {
      mqlNew: mqlNew.length, mqlNewToSql: mqlNewToSql.length, funnel: ratio(mqlNewToSql.length, mqlNew.length),
      mqlOpen: mqlOpen.length, sqlNew: sqlNew.length, production: ratio(sqlNew.length, mqlOpen.length + mqlNew.length),
      drill: {
        mqlNew: drill(ids(mqlNew), tagMap([[mqlNewToSql, "Became an SQL in the period"], [mqlNew.filter((d) => !mqlNewToSql.includes(d)), "Not an SQL in the period"]])),
        mqlNewToSql: drill(ids(mqlNewToSql)),
        funnel: drill(ids(mqlNew), tagMap([[mqlNewToSql, "Converted to SQL (numerator and denominator)"], [mqlNew.filter((d) => !mqlNewToSql.includes(d)), "Not converted (denominator only)"]])),
        mqlOpen: drill(ids(mqlOpen)),
        sqlNew: drill(ids(sqlNew)),
        production: drill(ids(union(mqlBase, sqlNew)), tagMap([[mqlOpen, "Opening MQL (denominator)"], [mqlNew, "New MQL (denominator)"], [sqlNew, "SQL produced (numerator)"]])),
      },
    },
    bofu: {
      sqlOpen: sqlOpen.length, sqlNew: sqlNew.length, inPlay: inPlay.length, won: playWon.length, lost: playLost.length, open: playOpen.length,
      winRate: ratio(playWon.length, inPlay.length), lossRate: ratio(playLost.length, inPlay.length),
      newWon: newWon.length, funnel: ratio(newWon.length, sqlNew.length),
      drill: {
        inPlay: drill(ids(inPlay), tagMap([[sqlOpen, "Open at the start"], [sqlNew, "Produced in the period"]])),
        won: drill(ids(playWon)),
        lost: drill(ids(playLost), tagMap([[playLost.filter((d) => d.exit.slipBack), "Slipped back to a pre-SQL stage"], [playLost.filter((d) => d.exit.debook), "Debooked"], [playLost.filter((d) => !d.exit.slipBack && !d.exit.debook), "Closed lost"]])),
        winRate: drill(ids(inPlay), tagMap([[playWon, "Won (numerator and denominator)"], [playLost, "Lost (denominator only)"], [playOpen, "Still open at the end (denominator only)"]])),
        funnel: drill(ids(sqlNew), tagMap([[newWon, "Won (numerator and denominator)"], [sqlNew.filter((d) => !newWon.includes(d)), "Not won in the period (denominator only)"]])),
      },
    },
  };
}

module.exports = { metrics };
