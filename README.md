# SpotDraft Quarterly Sales Retro

A live retro deck (Node.js + Express) of **AE performance by team**, in the same format as the GTM QBR deck.
Numbers come from HubSpot at page load; AE roster, team and New Business ARR targets come from the
**Regional Interlocks** dashboard's Inputs tab. Every number is clickable and opens the deals behind it,
every slide has an editable Insights box, and **Download as .pptx** lets you pick the slides to export.

## What's in the deck

It compares the last two completed fiscal quarters and both together (fiscal year starts 1 May):
today that is **FY 26-27 Q1 vs Q2, and H1**.

| Section | Slides |
|---|---|
| Opening | Title · Teams at a glance (NAM / EMEA / APAC / all teams, Q1, Q2, H1) |
| NAM, EMEA, APAC (each) | Section title (H1 tiles) · AE scorecard (H1 with Q1 → Q2 trend) · Revenue: target vs achieved · MOFU conversion · BOFU conversion · ACV and sales cycle · Looking back (What went right / could be better / what's next, editable) |
| Appendix | How the numbers are built |

**Team view.** Each AE belongs to the team (region) on the Regional Interlocks roster, and every deal counts
for its **deal owner**, whatever the deal's own region: an EMEA deal closed by Allison counts for NAM.
Team rows are all the team's AEs together.

**Click-through.** Click any number to see its deals (with a "Counts in this metric as" column for
percentages: numerator / denominator, won / lost / still open), filter them, download as CSV, or open a
deal in HubSpot. Targets are static inputs and say so.

**Insights.** Click an Insights box, type, click away: it saves for everyone and goes into the `.pptx`.
Stored per retro (e.g. "FY 26-27 H1") in `retro-insights.json` under `DATA_DIR`.

**Download.** "Download as .pptx" (top right) opens a picker: tick slides or whole sections, then download.
Page numbers follow the slides picked.

## Metric definitions

- **Achieved**: Annual Subscription Fee $ of Sales Pipeline deals closed won (incl. Deal Desk Verified) in the period.
- **Target**: Regional Interlocks New Business ARR target. H1 = sum of both quarters; `*` = one quarter is TBD.
- **MQL**: Sales Pipeline deal that entered Scheduled.
- **SQL**: SQL date in the period and current stage Qualified to buy, Presentation done, Proposal shared,
  Contracting, Closed won, Closed lost, Debook or Deal Desk Verified Closed Won. A closed-won deal with no SQL
  date counts as an SQL in the quarter it closed (left out of the sales-cycle average).
- **MOFU funnel %** = MQLs created in the period that became SQLs in it ÷ MQLs created.
  **Production %** = SQLs produced ÷ (MQLs open at the start + MQLs created).
- **BOFU win rate** = in-play SQLs won in the period ÷ SQLs in play (open at the start + produced).
  **New-SQL win %** = SQLs produced and won in the period ÷ SQLs produced. Slip-backs to pre-SQL stages count as lost.
- **ACV** = ARR won ÷ deals won. **Sales cycle** = average days from SQL date to close date of deals won.
- **H1** is computed on the combined deal set (not an average of the two quarters).
- **AEs shown**: on the roster with a target in either quarter, or with any MQL, SQL or win in the period.

Code: `server/metrics/rep.js` (every metric), `server/lib/data.js` (HubSpot pull), `server/lib/targets.js`
(roster + targets, with a built-in snapshot fallback), `public/slides.js` (slide layouts), `public/render.js`
(screen + `.pptx` drawing). `npm test` checks the metrics on hand-built deals.

## Railway

Create a service from this repo (it runs `npm install` and `npm start`), then under **Variables**:

| Variable | Value |
|---|---|
| `HUBSPOT_ACCESS_TOKEN` | HubSpot private app token (read deals, owners, deal schemas). The Regional Interlocks token works. |
| `REGIONAL_INTERLOCKS_URL` | `https://regional-interlocks-dashboard-production.up.railway.app` |
| `REGIONAL_INTERLOCKS_USERNAME` / `REGIONAL_INTERLOCKS_PASSWORD` | same as that service's `DASHBOARD_USERNAME` / `DASHBOARD_PASSWORD` |
| `RETRO_USERNAME` / `RETRO_PASSWORD` | login for this deck (recommended: the repo is public) |
| `DATA_DIR` | a mounted Volume path (e.g. `/data`) so insights survive redeploys |

Optional: `RETRO_REVIEW_QUARTER` (e.g. `FY 26-27 Q2`) to pin the quarters; `MOCK_DATA=1` for synthetic data.
Leave `PORT` unset. If Regional Interlocks can't be reached, the deck uses the targets snapshot of
26 Nov 2026 and says so in the bottom bar.

Data refreshes every 20 minutes; **Refresh data** (top bar) or `/api/retro?refresh=1` forces a pull.

## Local run

```bash
npm install
cp .env.example .env   # fill in the values above (or MOCK_DATA=1)
npm start              # http://localhost:4000
```
