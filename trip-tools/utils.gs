/**
 * Updates a cell with newValue and an optional hidden Note.
 */
function updateCell(cellRef, newValue, noteValue = null) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(cellRef);
  range.setHorizontalAlignment("left");

  const currentValue = range.getValue().toString();
  const cleanNewValue = newValue.toString().replace(/^'/, "");

  // Update logic
  if (currentValue === "" || currentValue === cleanNewValue) {
    range.setValue(newValue);
    if (noteValue) range.setNote(noteValue); // Add the label/note here
    return;
  }

  // Conflict prompt
  const response = ui.alert('Conflict at ' + cellRef, `Overwrite "${currentValue}" with "${newValue}"?`, ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    range.setValue(newValue);
    if (noteValue) range.setNote(noteValue);
  }
}

/**
 * THE SHARED ENGINE
 * Fixed: Variable references, JSON pathing, and strict USA city validation.
 */
function processTrip_(modeName) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const response = ui.prompt('Manual ' + modeName + ' Calculation', 
                           'Enter cells (Origin, Stop 1... Output Cell):\nExample: A2, B2, C2', 
                           ui.ButtonSet.OK);

  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const cells = response.getResponseText().split(',').map(function(s) { return s.trim().toUpperCase(); });
  
  if (cells.length < 3) {
    ui.alert("Error: Please enter at least 3 cells.");
    return null;
  }

  if (new Set(cells).size !== cells.length) {
    ui.alert("Safety Error: Duplicate cell references detected.");
    return null;
  }

  const targetCell = cells[cells.length - 1];
  const addressCells = cells.slice(0, -1);
  var totalMeters = 0;

  try {
    // --- STEP 1: STRICT VALIDATION & USA FILTERING ---
    for (var i = 0; i < addressCells.length; i++) {
      var val = sheet.getRange(addressCells[i]).getValue();
      
      if (!val || val.toString().trim() === "") {
        ui.alert("Error: Cell " + addressCells[i] + " is empty.");
        return null;
      }
      
      // Fixed: Variable name consistently used as 'geocoder'
      var geocoder = Maps.newGeocoder().setRegion('US').geocode(val.toString());
      
      if (geocoder.status !== 'OK') {
        ui.alert("Bad Address Error: Google cannot find '" + val + "' in cell " + addressCells[i]);
        return null;
      }
      
      var result = geocoder.results[0]; // Access the first result
      
      // STRICT CHECK: Reject if Google had to "guess" (partial_match)
      if (result.partial_match) {
        var suggested = result.formatted_address.replace(/, USA$|, United States$/gi, "");
        ui.alert("Invalid City: '" + val + "' in cell " + addressCells[i] + 
                 " is not a precise match.\n\nDid you mean: " + suggested + "?");
        return null;
      }
    }

    // --- STEP 2: CALCULATE ---
    for (var j = 0; j < addressCells.length - 1; j++) {
      var origin = sheet.getRange(addressCells[j]).getValue();
      var destination = sheet.getRange(addressCells[j+1]).getValue();

      var directions = Maps.newDirectionFinder()
        .setOrigin(origin).setDestination(destination)
        .setMode(Maps.DirectionFinder.Mode.DRIVING).getDirections();

      // Fixed: Precise pathing for routes[0].legs[0]
      if (directions && directions.routes && directions.routes.length > 0) {
        totalMeters += directions.routes[0].legs[0].distance.value;
      } else {
        ui.alert("Route Error: No driving path found between " + addressCells[j] + " and " + addressCells[j+1]);
        return null;
      }
    }
    return { meters: totalMeters, targetCell: targetCell };
  } catch (e) {
    ui.alert("System Error: " + e.message);
    return null;
  }
}