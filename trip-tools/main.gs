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

// 2. This function ONLY processes data. The HTML calls THIS.
function parseDispatchDetails(fullText) {
  if (!fullText) return;
  const ui = SpreadsheetApp.getUi();
  const activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveCell();
  const sections = fullText.split("PICKUP & DELIVERY DETAILS");
  const loadSection = sections[0] || "";
  const stopsSection = sections[1] || "";
  updateLoadInfo(loadSection, activeCell);
  updateStopsInfo(stopsSection, activeCell);
  SpreadsheetApp.flush(); 
  ui.alert("Dispatch Details Updated Successfully!");
}

function updateLoadInfo(loadText, startCell) {
  const loadMatch    = loadText.match(/Load #\s*(\d+)/);
  const freightMatch = loadText.match(/Freight Type:\s*([^\n\r]*)/);
  const tempMatch    = loadText.match(/Temp:\s*([^\n\r]*)/);

  const loadNum     = loadMatch ? "'" + loadMatch[1].trim() : "N/A";
  const temp        = tempMatch ? tempMatch[1].trim() : "N/A";

  // Extract raw type first, then check if it equals "Reefer"
  const rawFreightType = freightMatch ? freightMatch[1].trim() : "N/A";
  const freightType    = (rawFreightType === "Reefer") ? "Reefer" : "Van";

  updateCell(startCell.getA1Notation(), loadNum);
  updateCell(startCell.offset(1, 0).getA1Notation(), freightType);
  updateCell(startCell.offset(2, 0).getA1Notation(), temp);
}

function updateStopsInfo(stopsText, startCell) {
  const sheet = startCell.getSheet();
  const startRow = startCell.getRow();
  const startCol = startCell.getColumn();

  // 1. Split using capturing parentheses to keep the delimiter
  const rawParts = stopsText.split(/(Stop #\d+:)/).filter(s => s.trim() !== "");

  // 2. Reconstruct the stops into full strings [ "Stop #1: ...content...", "Stop #2: ...content..." ]
  const fullStops = [];
  for (let i = 0; i < rawParts.length; i += 2) {
    const label = rawParts[i];
    const content = rawParts[i + 1] || "";
    fullStops.push(label + content);
  }

  const pickups = fullStops.filter(s => /PICKUP/i.test(s));
  const drops = fullStops.filter(s => /DROP/i.test(s));
  const maxStops = Math.max(pickups.length, drops.length);

  // if max stops > 2 add stop rows
  if (maxStops > 2) {
    const rowsToAdd = maxStops - 2;
    const lastCol = sheet.getLastColumn();
  
    const bottomRowRange = sheet.getRange(startRow + 2, 1, 1, lastCol);
    bottomRowRange.setBorder(null, null, false, null, null, null)

    sheet.insertRowsAfter(startRow + 2, rowsToAdd);

    const newBottomRowRange = sheet.getRange(startRow + 2 + rowsToAdd, 1, 1, lastCol);
    newBottomRowRange.setBorder(
      null,
      null,
      true,
      null,
      null,
      null,
      "black",
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
  }

  const dateRegex = /[A-Z][a-z]{2}\s\d{1,2},\s\d{4}/;
  // Process PICKUPS (Col 3)
  pickups.forEach((stop, index) => {
    const location = extractCityState(stop);
    const stopLabel = stop.trim();
    if (index === 0) {
      const pMatch = stop.match(dateRegex);
      updateCell(startCell.offset(0, 3).getA1Notation(), pMatch ? pMatch[0] : "N/A");
    }
    // 'stop' now includes "Stop #1:" at the beginning
    updateCell(startCell.offset(index + 1, 3).getA1Notation(), location, stopLabel);
  });

  // Process DROPS (Col 4)
  drops.forEach((stop, index) => {
    const location = extractCityState(stop);
    const stopLabel = stop.trim();
    if (index === 0) {
      const dMatch = stop.match(dateRegex);
      updateCell(startCell.offset(0, 4).getA1Notation(), dMatch ? dMatch[0] : "N/A");
    }
    updateCell(startCell.offset(index + 1, 4).getA1Notation(), location, stopLabel);
  });
}

function extractCityState(stopContent) {
  // 1. Isolate the Address line specifically (ends at the Zip code)
  // This helps avoid matching words in the Name: or Stop #: lines
  const addressLineMatch = stopContent.match(/Address:[\s\S]*?([^\n\r]+,\s[A-Z]{2}\s\d{5})/i);
  if (!addressLineMatch) return "Location N/A";
  const addressLine = addressLineMatch[1].trim();
  /**
   * 2. Extract City and State from the isolated line
   * ([^,]+) -> Captures everything before the comma (City)
   * ,\s*([A-Z]{2}) -> Captures the 2-letter State after the comma
   */
  const match = addressLine.match(/([^,]+),\s*([A-Z]{2})\s+\d{5}/i);
  if (match && match[1] && match[2]) {
    const city = match[1].trim();
    const state = match[2].toUpperCase();
    return city + ", " + state;
  }
  return "City/ST N/A";
}