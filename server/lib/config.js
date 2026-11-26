"use strict";

/**
 * HubSpot names and stage ids. Same values as the QBR deck and the Regional
 * Interlocks dashboard (propertyMap.js / stageCriteria.js), confirmed there
 * against the live portal.
 */
const stages = {
  scheduled: "144930627",          // entering it = an MQL ("meeting")
  noShow: "198705765",
  salesAccepted: "144930628",      // Sales Accepted Account
  notQualified: "144930629",
  qualifiedToBuy: "qualifiedtobuy",
  presentationDone: "presentationscheduled",
  proposalShared: "decisionmakerboughtin",
  contracting: "10679404",
  closedWon: "closedwon",
  ddvClosedWon: "1365874062",      // Deal Desk Verified Closed Won
  closedLost: "closedlost",
  debook: "1422037570",
};

module.exports = {
  salesPipelineId: "default",
  deal: {
    name: "dealname",
    stage: "dealstage",
    pipeline: "pipeline",
    owner: "hubspot_owner_id",
    sdrOwner: "sales_owner",                                    // "SDR Owner (Deal)"
    region: "region",
    leadSource: "lead_source_deal",                             // "Lead Source"
    finalLeadSource: "final_lead_source__for_org_reporting_",   // "Final Lead Source"
    sqlDate: "became_qualified_to_buy__sales_pipeline_",        // "Became Qualified to Buy Date"
    closeDate: "closedate",
    isClosedWon: "hs_is_closed_won",                            // true for Closed Won and Deal Desk Verified Closed Won
    isClosed: "hs_is_closed",
    revenue: "annual_license_fee__",                            // "Annual Subscription Fee $"
  },
  stages,
  // An SQL is a deal whose SQL date falls in the period AND whose current stage is one of these.
  sqlStageIds: [stages.qualifiedToBuy, stages.presentationDone, stages.proposalShared, stages.contracting,
    stages.closedWon, stages.ddvClosedWon, stages.closedLost, stages.debook],
  dateEntered: (stageId) => `hs_v2_date_entered_${stageId}`,
  teams: [
    { key: "NAM", title: "NAM", long: "North America" },
    { key: "EMEA", title: "EMEA", long: "Europe, Middle East & Africa" },
    { key: "APAC", title: "APAC", long: "Asia Pacific" },
  ],
  // How far back an MQL can have been created and still be in the opening balance.
  mqlLookbackDays: 730,
};
