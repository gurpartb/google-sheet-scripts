// ─── MENU ────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Trip Tools')
    .addItem('Add Dispatch Details', 'openDispatchDialog')
    .addToUi();
}

function openDispatchDialog() {
  const html = HtmlService.createHtmlOutputFromFile('dialog')
    .setWidth(500)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Paste Dispatch Details');
}

// ─── DIALOG HANDLERS (called via google.script.run) ──────────────────────────

/**
 * Dry-run: returns [{cell, current, incoming}] for any cell that would be overwritten.
 * No sheet writes — safe to call from a modal dialog.
 */
function checkConflicts(text) {
  if (!text) return [];

  const row     = SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getRow();
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const planned = buildWrites_(text, row);

  return planned.reduce(function(acc, w) {
    if (w.newRow) return acc;
    const current    = sheet.getRange(w.cellRef).getValue().toString();
    const cleanValue = w.value.toString().replace(/^'/, '');
    if (current !== '' && current !== cleanValue) {
      acc.push({ cell: w.cellRef, current: current, incoming: cleanValue });
    }
    return acc;
  }, []);
}

/**
 * Writes dispatch details to the sheet. Always overwrites — conflicts were
 * already confirmed by the user in the dialog before this is called.
 * Returns true on success, false on empty input.
 */
function parseDispatchDetails(text) {
  if (!text) return false;

  const sheet     = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow  = SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getRow();
  const totalCols = sheet.getMaxColumns();
  const planned   = buildWrites_(text, startRow);

  // Insert extra rows when there are more than 2 pickup OR drop stops
  const pickupCount = planned.filter(w => w.stopType === 'pickup').length;
  const dropCount   = planned.filter(w => w.stopType === 'drop').length;
  const extraRows   = Math.max(0, Math.max(pickupCount, dropCount) - 2);
  if (extraRows > 0) {
    sheet.getRange(startRow + 2, 1, 1, totalCols)
         .setBorder(null, null, false, null, null, null);
    sheet.insertRowsAfter(startRow + 2, extraRows);
    sheet.getRange(startRow + 2 + extraRows, 1, 1, totalCols)
         .setBorder(null, null, true, null, null, null,
                    'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  planned.forEach(function(w) {
    const range = sheet.getRange(w.cellRef);
    range.setHorizontalAlignment('left');
    range.setValue(w.value);
    if (w.note) range.setNote(w.note);
  });

  SpreadsheetApp.flush();
  return true;
}

// ─── INTERNAL HELPERS ─────────────────────────────────────────────────────────

/**
 * Parses dispatch text and returns the full list of planned cell writes as
 * [{cellRef, value, note?, stopRow?}]. Used by both checkConflicts and
 * parseDispatchDetails so parsing logic lives in one place.
 */
function buildWrites_(text, row) {
  const writes = [];

  const sections     = text.split('PICKUP & DELIVERY DETAILS');
  const loadSection  = sections[0] || '';
  const stopsSection = sections[1] || '';

  // ── Load header — Column A ───────────────────────────────────────────────────
  const loadMatch    = loadSection.match(/Load #\s*(\d+)/);
  const freightMatch = loadSection.match(/Freight Type:\s*([^\n\r]*)/);
  const tempMatch    = loadSection.match(/Temp:\s*([^\n\r]*)/);

  const loadNum     = loadMatch    ? "'" + loadMatch[1].trim()  : 'N/A';
  const temp        = tempMatch    ? tempMatch[1].trim()        : 'N/A';
  const rawFreight  = freightMatch ? freightMatch[1].trim()     : 'N/A';
  const freightType = rawFreight === 'Reefer' ? 'Reefer' : 'Van';

  writes.push({ cellRef: `A${row}`,     value: loadNum });
  writes.push({ cellRef: `A${row + 1}`, value: freightType });
  writes.push({ cellRef: `A${row + 2}`, value: temp });

  // ── Stops ───────────────────────────────────────────────────────────────────
  const rawParts  = stopsSection.split(/(Stop #\d+:)/).filter(s => s.trim() !== '');
  const fullStops = [];
  for (let i = 0; i < rawParts.length; i += 2) {
    fullStops.push(rawParts[i] + (rawParts[i + 1] || ''));
  }

  const pickups   = fullStops.filter(s => /PICKUP/i.test(s));
  const drops     = fullStops.filter(s => /DROP/i.test(s));
  const dateRegex = /[A-Z][a-z]{2}\s\d{1,2},\s\d{4}/;

  // Pickup dates and cities — Column D
  pickups.forEach(function(stop, i) {
    if (i === 0) {
      const m = stop.match(dateRegex);
      writes.push({ cellRef: `D${row}`, value: m ? m[0] : 'N/A' });
    }
    writes.push({
      cellRef:  `D${row + i + 1}`,
      value:    cityState_(stop),
      note:     stop.trim(),
      stopType: 'pickup',
      newRow:   i >= 2
    });
  });

  // Drop dates and cities — Column E
  drops.forEach(function(stop, i) {
    if (i === 0) {
      const m = stop.match(dateRegex);
      writes.push({ cellRef: `E${row}`, value: m ? m[0] : 'N/A' });
    }
    writes.push({
      cellRef:  `E${row + i + 1}`,
      value:    cityState_(stop),
      note:     stop.trim(),
      stopType: 'drop',
      newRow:   i >= 2
    });
  });

  return writes;
}

/** Extracts "City, ST" from a stop block's Address field. */
function cityState_(stopText) {
  const addrMatch = stopText.match(/Address:[\s\S]*?([^\n\r]+,\s[A-Z]{2}\s\d{5})/i);
  if (!addrMatch) return 'Location N/A';
  const m = addrMatch[1].trim().match(/([^,]+),\s*([A-Z]{2})\s+\d{5}/i);
  return (m && m[1] && m[2]) ? m[1].trim() + ', ' + m[2].toUpperCase() : 'City/ST N/A';
}
