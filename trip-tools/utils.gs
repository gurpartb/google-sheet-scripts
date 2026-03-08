/**
 * Updates a cell with newValue and an optional hidden Note.
 * Accepts optional sheet and ui to avoid repeated service lookups when called in loops.
 */
function updateCell(cellRef, newValue, noteValue = null, sheet = null, ui = null) {
  ui    = ui    || SpreadsheetApp.getUi();
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const range = sheet.getRange(cellRef);
  range.setHorizontalAlignment("left");

  const currentValue = range.getValue().toString();
  const cleanNewValue = newValue.toString().replace(/^'/, "");

  if (currentValue === "" || currentValue === cleanNewValue) {
    range.setValue(newValue);
    if (noteValue) range.setNote(noteValue);
    return;
  }

  const response = ui.alert('Conflict at ' + cellRef, `Overwrite "${currentValue}" with "${newValue}"?`, ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    range.setValue(newValue);
    if (noteValue) range.setNote(noteValue);
  }
}

/**
 * THE SHARED ENGINE
 * Reads all address cell values once upfront to avoid re-reading in the calculate loop.
 */
function processTrip_(modeName) {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const response = ui.prompt(
    'Manual ' + modeName + ' Calculation',
    'Enter cells (Origin, Stop 1... Output Cell):\nExample: A2, B2, C2',
    ui.ButtonSet.OK
  );

  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const cells = response.getResponseText().split(',').map(s => s.trim().toUpperCase());

  if (cells.length < 3) {
    ui.alert("Error: Please enter at least 3 cells.");
    return null;
  }

  if (new Set(cells).size !== cells.length) {
    ui.alert("Safety Error: Duplicate cell references detected.");
    return null;
  }

  const targetCell   = cells[cells.length - 1];
  const addressCells = cells.slice(0, -1);
  let totalMeters    = 0;

  try {
    // Batch read all address values once — reused in both the validate and calculate loops
    const addressValues = addressCells.map(c => sheet.getRange(c).getValue());

    // --- STEP 1: STRICT VALIDATION & USA FILTERING ---
    for (let i = 0; i < addressCells.length; i++) {
      const val = addressValues[i];

      if (!val || val.toString().trim() === "") {
        ui.alert("Error: Cell " + addressCells[i] + " is empty.");
        return null;
      }

      const geocoder = Maps.newGeocoder().setRegion('US').geocode(val.toString());

      if (geocoder.status !== 'OK') {
        ui.alert("Bad Address Error: Google cannot find '" + val + "' in cell " + addressCells[i]);
        return null;
      }

      const result = geocoder.results[0];

      if (result.partial_match) {
        const suggested = result.formatted_address.replace(/, USA$|, United States$/gi, "");
        ui.alert("Invalid City: '" + val + "' in cell " + addressCells[i] +
                 " is not a precise match.\n\nDid you mean: " + suggested + "?");
        return null;
      }
    }

    // --- STEP 2: CALCULATE (reuse cached addressValues — no second sheet reads) ---
    for (let j = 0; j < addressCells.length - 1; j++) {
      const directions = Maps.newDirectionFinder()
        .setOrigin(addressValues[j].toString())
        .setDestination(addressValues[j + 1].toString())
        .setMode(Maps.DirectionFinder.Mode.DRIVING)
        .getDirections();

      if (directions && directions.routes && directions.routes.length > 0) {
        totalMeters += directions.routes[0].legs[0].distance.value;
      } else {
        ui.alert("Route Error: No driving path found between " + addressCells[j] + " and " + addressCells[j + 1]);
        return null;
      }
    }

    return { meters: totalMeters, targetCell: targetCell };
  } catch (e) {
    ui.alert("System Error: " + e.message);
    return null;
  }
}
