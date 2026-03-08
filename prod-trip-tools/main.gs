/**
 * ENTRY POINT 1: DISTANCE
 */
function GET_TRIP_DISTANCE() {
  const data = processTrip_("Distance");
  if (!data) return;
  const miles = Math.round(data.meters * 0.000621371);
  const commaMiles = miles.toLocaleString('en-US');
  const stringValue = "(" + commaMiles + " mi)";
  updateCell(data.targetCell, stringValue);
}

/**
 * ENTRY POINT 2: TIME
 */
function GET_TRIP_TIME() {
  const data = processTrip_("Time (50 MPH)");
  if (!data) return;
  const miles = data.meters * 0.000621371;
  const totalHours = miles / 50;
  const hours = Math.floor(totalHours);
  const minutes = Math.round((totalHours - hours) * 60);
  const formattedTime = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
  const stringValue = "'" + formattedTime;
  updateCell(data.targetCell, stringValue);
}

/**
 * Called from the dialog. Returns true on success, false on empty input
 * so the dialog can decide whether to close.
 */
function parseDispatchDetails(fullText) {
  if (!fullText) return false;

  const ui         = SpreadsheetApp.getUi();
  const activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveCell();
  const sheet      = activeCell.getSheet();

  const sections    = fullText.split("PICKUP & DELIVERY DETAILS");
  const loadSection = sections[0] || "";
  const stopsSection = sections[1] || "";

  updateLoadInfo(loadSection, activeCell, sheet, ui);
  updateStopsInfo(stopsSection, activeCell, sheet, ui);
  SpreadsheetApp.flush();
  return true;
}

function updateLoadInfo(loadText, startCell, sheet, ui) {
  const loadMatch    = loadText.match(/Load #\s*(\d+)/);
  const freightMatch = loadText.match(/Freight Type:\s*([^\n\r]*)/);
  const tempMatch    = loadText.match(/Temp:\s*([^\n\r]*)/);

  const loadNum  = loadMatch ? "'" + loadMatch[1].trim() : "N/A";
  const temp     = tempMatch ? tempMatch[1].trim() : "N/A";

  const rawFreightType = freightMatch ? freightMatch[1].trim() : "N/A";
  const freightType    = (rawFreightType === "Reefer") ? "Reefer" : "Van";

  updateCell(startCell.getA1Notation(),                  loadNum,     null, sheet, ui);
  updateCell(startCell.offset(1, 0).getA1Notation(),     freightType, null, sheet, ui);
  updateCell(startCell.offset(2, 0).getA1Notation(),     temp,        null, sheet, ui);
}

function updateStopsInfo(stopsText, startCell, sheet, ui) {
  const startRow = startCell.getRow();
  const startCol = startCell.getColumn();

  // Split using capturing parentheses to keep the delimiter
  const rawParts = stopsText.split(/(Stop #\d+:)/).filter(s => s.trim() !== "");

  // Reconstruct stops into full strings: ["Stop #1: ...content...", "Stop #2: ...content..."]
  const fullStops = [];
  for (let i = 0; i < rawParts.length; i += 2) {
    fullStops.push(rawParts[i] + (rawParts[i + 1] || ""));
  }

  const pickups  = fullStops.filter(s => /PICKUP/i.test(s));
  const drops    = fullStops.filter(s => /DROP/i.test(s));
  const maxStops = Math.max(pickups.length, drops.length);

  // If max stops > 2, insert extra rows
  if (maxStops > 2) {
    const rowsToAdd    = maxStops - 2;
    const lastCol      = sheet.getLastColumn();
    const bottomRow    = sheet.getRange(startRow + 2, 1, 1, lastCol);
    bottomRow.setBorder(null, null, false, null, null, null);
    sheet.insertRowsAfter(startRow + 2, rowsToAdd);
    sheet.getRange(startRow + 2 + rowsToAdd, 1, 1, lastCol).setBorder(
      null, null, true, null, null, null,
      "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
  }

  const dateRegex = /[A-Z][a-z]{2}\s\d{1,2},\s\d{4}/;

  // Process PICKUPS (Col offset 3)
  pickups.forEach((stop, index) => {
    const location  = extractCityState(stop);
    const stopLabel = stop.trim();
    if (index === 0) {
      const pMatch = stop.match(dateRegex);
      updateCell(startCell.offset(0, 3).getA1Notation(), pMatch ? pMatch[0] : "N/A", null, sheet, ui);
    }
    updateCell(startCell.offset(index + 1, 3).getA1Notation(), location, stopLabel, sheet, ui);
  });

  // Process DROPS (Col offset 4)
  drops.forEach((stop, index) => {
    const location  = extractCityState(stop);
    const stopLabel = stop.trim();
    if (index === 0) {
      const dMatch = stop.match(dateRegex);
      updateCell(startCell.offset(0, 4).getA1Notation(), dMatch ? dMatch[0] : "N/A", null, sheet, ui);
    }
    updateCell(startCell.offset(index + 1, 4).getA1Notation(), location, stopLabel, sheet, ui);
  });
}

function extractCityState(stopContent) {
  const addressLineMatch = stopContent.match(/Address:[\s\S]*?([^\n\r]+,\s[A-Z]{2}\s\d{5})/i);
  if (!addressLineMatch) return "Location N/A";
  const addressLine = addressLineMatch[1].trim();
  const match = addressLine.match(/([^,]+),\s*([A-Z]{2})\s+\d{5}/i);
  if (match && match[1] && match[2]) {
    return match[1].trim() + ", " + match[2].toUpperCase();
  }
  return "City/ST N/A";
}
