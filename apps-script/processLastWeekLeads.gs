/**
 * Processes 7 days of lead files from email, applies filters,
 * removes duplicates, and generates a unified weekly report.
 *
 * Key operations:
 * 1. Identify last week's Monday.
 * 2. For each day (Mon–Sun), find the corresponding CSV email.
 * 3. Apply filters: valid statuses, valid cities, correct date window.
 * 4. De-duplicate buyers across the entire week.
 * 5. Compute multiple derived fields for MIS and CRM.
 * 6. Append all cleaned rows into a final output sheet.
 */

function processLastWeekLeads() {
  const SUBJECT = "You can mention your mail subject";
  const OUT_SHEET = "LastWeekProcessedLeads";
  const TZ = "GMT+5:30";

  const VALID_STATUSES = ["AS-N", "AS-Y", "Online with Date", "Online Without Date"];
  const VALID_CITIES = ["Bangalore", "Hyderabad", "Chennai", "Kolkata"];

  try {
    const today = new Date();

    // Determine last week's Monday
    const lastMonday = getLastWeekMonday(today);

    // Get all 7 days of last week
    const days = [];
    for (let i = 0; i < 7; i++) {
      const d = new Date(lastMonday);
      d.setDate(lastMonday.getDate() + i);
      days.push(d);
    }

    const combined = [];
    const globalSeenBuyers = new Set();

    for (const day of days) {
      const prev = new Date(day);
      prev.setDate(day.getDate() - 1);

      // Daily window: previous day 6 PM → current day 6 PM
      const winStart = new Date(prev); winStart.setHours(18,0,0,0);
      const winEnd = new Date(day);  winEnd.setHours(18,0,0,0);

      // Get latest email for this calendar day
      const message = getLastMessageForCalendarDate(SUBJECT, day);
      if (!message) continue;

      const attachments = message.getAttachments();
      const csvBlob = attachments.reverse().find(a => a.getName().toLowerCase().endsWith(".csv"));
      if (!csvBlob) continue;

      const csvText = csvBlob.getDataAsString("UTF-8");
      const rows = Utilities.parseCsv(csvText);
      if (!rows || rows.length <= 1) continue;

      const header = rows[0];
      const dataRows = rows.slice(1);

      // Filter rows matching city, status, and the correct 6 PM window
      const filtered = dataRows.filter(r => {
        const buyer = String(r[1] || "").trim();
        const city = String(r[9] || "").trim();
        const status = String(r[18] || "").trim();
        const created = tryParseDate(r[11]);

        if (!buyer || !city || !status || !created) return false;
        if (VALID_STATUSES.indexOf(status) === -1) return false;
        if (VALID_CITIES.indexOf(city) === -1) return false;
        return created >= winStart && created < winEnd;
      });

      const seenToday = new Set();
      const uniqueForDay = [];

      // Deduplicate at both daily and weekly level
      for (const r of filtered) {
        const buyer = String(r[1]).trim();
        if (!buyer) continue;
        if (seenToday.has(buyer)) continue;
        if (globalSeenBuyers.has(buyer)) continue;

        seenToday.add(buyer);
        globalSeenBuyers.add(buyer);
        uniqueForDay.push(r);
      }

      // Compute derived fields and add to final combined array
      for (const r of uniqueForDay) {
        combined.push(computeCalculatedColumns(r, TZ));
      }
    }

    // Write final output
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(OUT_SHEET);
    if (!sh) sh = ss.insertSheet(OUT_SHEET);
    sh.clear();

    const outHeaders = [
      "BuyerID","BuyerPhoneNumber","ProjectID","ProjectName","City",
      "LeadCreatedDate&Time","AccountManagerScheduleID","AccountManagerCRMID",
      "TypeOfLead","CRM-LeadCreatedTime","CRMLeadAssignedTime","LeadCreatedTime",
      "LeadAssignedTime","CreatedDateAsPerMIS","LeadCreatedDatexBuyerPhoneNumber(CRM)",
      "LeadCreatedDatexBuyerPhoneNumber(MIS)","LeadAssignedDatexBuyerPhoneNumber(CRM)",
      "LastAssignedDateAsPerMIS","LeadAssignedDatexBuyerPhoneNumber(MIS)"
    ];

    sh.getRange(1,1,1,outHeaders.length).setValues([outHeaders]);

    if (combined.length) {
      sh.getRange(2,1,combined.length,outHeaders.length).setValues(combined);
    }

  } catch (e) {
    Logger.log("ERROR: " + e.message);
  }
}

function computeCalculatedColumns(r, TZ) {
  const buyer  = String(r[1] || "").trim();
  const encId  = String(r[2] || "").trim();
  const project = r[7] || "";
  const source  = r[8] || "";
  const city    = r[9] || "";

  const createdRaw = tryParseDate(r[11]);
  const subStatus = r[14] || "";
  const assignedTo = r[15] || "";
  const status = r[18] || "";
  const lastAssignedRaw = tryParseDate(r[21]);
  const lastAssignedBy = r[22] || "";

  const leadCreatedTimeOnly  = createdRaw ? Utilities.formatDate(createdRaw, TZ, "hh:mm a") : "";
  const leadAssignedTimeOnly = lastAssignedRaw ? Utilities.formatDate(lastAssignedRaw, TZ, "hh:mm a") : "";

  // Produce MIS-sorted dates for created and assigned
  const createdMIS = normalizeAfter6(createdRaw, TZ);
  const createdMISFormatted = createdMIS ? Utilities.formatDate(createdMIS, TZ, "dd-MMM-yyyy") : "";

  const assignedMIS = normalizeAfter6(lastAssignedRaw, TZ);
  const assignedMISFormatted = assignedMIS ? Utilities.formatDate(assignedMIS, TZ, "dd-MMM-yyyy hh:mm a") : "";

  // Numeric encodings
  const leadInt = createdRaw ? excelDateInt(createdRaw) : "";
  const leadCreatedXEnc = (leadInt && encId) ? `${leadInt}${encId}` : "";

  const createdSortedInt = createdMIS ? excelDateInt(createdMIS) : "";
  const leadFlagAfter6Int = (createdSortedInt && encId) ? `${createdSortedInt}${encId}` : "";

  const lastAssignInt = lastAssignedRaw ? excelDateInt(lastAssignedRaw) : "";
  const lastAssignedXEnc = (lastAssignInt && encId) ? `${lastAssignInt}${encId}` : "";

  const lastAssignMISInt = assignedMIS ? excelDateInt(assignedMIS) : "";
  const lastAssignedSortedIntXEnc = (lastAssignMISInt && encId) ? `${lastAssignMISInt}${encId}` : "";

  return [
    buyer, encId, project, source, city,
    createdRaw ? Utilities.formatDate(createdRaw, TZ, "dd-MMM-yyyy HH:mm") : "",
    subStatus, assignedTo, status,
    lastAssignedRaw ? Utilities.formatDate(lastAssignedRaw, TZ, "dd-MMM-yyyy HH:mm") : "",
    lastAssignedBy, leadCreatedTimeOnly, leadAssignedTimeOnly,
    createdMISFormatted,
    leadCreatedXEnc, leadFlagAfter6Int, lastAssignedXEnc,
    assignedMISFormatted, lastAssignedSortedIntXEnc
  ];
}

function normalizeAfter6(dt, TZ) {
  if (!dt) return null;
  const cutoff = new Date(dt); cutoff.setHours(18,0,0,0);
  if (dt >= cutoff) {
    const next = new Date(dt);
    next.setDate(dt.getDate() + 1);
    next.setHours(9,30,0,0);
    return next;
  }
  return dt;
}
