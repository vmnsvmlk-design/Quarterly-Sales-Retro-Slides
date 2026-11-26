"use strict";

// Fiscal year starts 1 May (override with FISCAL_YEAR_START_MONTH, 1-12).
// "FY 26-27" = 1 May 2026 - 30 Apr 2027; Q1 = May-Jul, Q2 = Aug-Oct, Q3 = Nov-Jan, Q4 = Feb-Apr.
const START_MONTH = (Number(process.env.FISCAL_YEAR_START_MONTH) || 5) - 1; // 0-indexed

function build(fyStartYear, q) {
  const start = new Date(Date.UTC(fyStartYear, START_MONTH + (q - 1) * 3, 1));
  const end = new Date(Date.UTC(fyStartYear, START_MONTH + q * 3, 1));
  const fy = `FY ${String(fyStartYear).slice(2)}-${String(fyStartYear + 1).slice(2)}`;
  return { fyStartYear, q, start, end, label: `${fy} Q${q}`, short: `Q${q}`, fyLabel: fy };
}

function quarterOf(date) {
  const y = date.getUTCFullYear();
  const m = date.getUTCMonth();
  const fyStartYear = m >= START_MONTH ? y : y - 1;
  const monthsIn = (m - START_MONTH + 12) % 12;
  return build(fyStartYear, Math.floor(monthsIn / 3) + 1);
}

function shift(quarter, n) {
  const idx = quarter.fyStartYear * 4 + (quarter.q - 1) + n;
  return build(Math.floor(idx / 4), (idx % 4) + 1);
}

function parseLabel(label) {
  const m = /^FY\s*(\d{2})-(\d{2})\s*Q([1-4])$/i.exec(String(label || "").trim());
  return m ? build(2000 + Number(m[1]), Number(m[3])) : null;
}

/**
 * The retro compares two consecutive quarters and the two together:
 *  - review  = last completed quarter (or RETRO_REVIEW_QUARTER),
 *  - compare = the quarter before it,
 *  - combined = both ("H1" / "H2" when they make up a fiscal half).
 */
function retroPeriods(now = new Date()) {
  const review = parseLabel(process.env.RETRO_REVIEW_QUARTER) || shift(quarterOf(now), -1);
  const compare = shift(review, -1);
  const isHalf = review.q === 2 || review.q === 4;
  const short = isHalf ? (review.q === 2 ? "H1" : "H2") : `${compare.short}+${review.short}`;
  const combined = {
    start: compare.start, end: review.end, short,
    label: isHalf ? `${review.fyLabel} ${short}` : `${compare.label} + ${review.short}`,
    quarters: [compare, review],
  };
  return { compare, review, combined, list: [compare, review, combined] };
}

const inRange = (ms, r) => ms != null && ms >= r.start.getTime() && ms < r.end.getTime();

module.exports = { quarterOf, shift, parseLabel, retroPeriods, inRange };
