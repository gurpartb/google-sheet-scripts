// ─── MENU ────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Trip Tools')
    .addItem('Add Dispatch Details', 'openDispatchDialog')
    .addSeparator()
    .addItem('Update Trip Status', 'updateTripStatus')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Help')
        .addItem('Add Dispatch Details', 'helpDispatch')
        .addItem('Update Trip Status',   'helpTripStatus')
    )
    .addToUi();
}

function helpDispatch() {
  showHelpSidebar_('dispatch');
}

function helpTripStatus() {
  showHelpSidebar_('tripstatus');
}

function showHelpSidebar_(topic) {
  const tmpl = HtmlService.createTemplateFromFile('help');
  tmpl.topic  = topic;
  SpreadsheetApp.getUi().showSidebar(
    tmpl.evaluate().setTitle('Trip Tools — Help')
  );
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
    if (w.format) range.setNumberFormat(w.format);
    range.setValue(w.value);
    if (w.note) range.setNote(w.note);
  });

  SpreadsheetApp.flush();
  return true;
}

// ─── INTERNAL HELPERS ─────────────────────────────────────────────────────────

/**
 * @typedef {{ cellRef: string, value: any, note?: string, format?: string, stopType?: string, newRow?: boolean }} WriteEntry
 */

/**
 * Parses dispatch text and returns the full list of planned cell writes as
 * WriteEntry[]. Used by both checkConflicts and parseDispatchDetails so
 * parsing logic lives in one place.
 * @returns {WriteEntry[]}
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

  // Anchor the J–W window to the Sunday of the first pickup's week
  const firstPickupDateMatch = pickups.length ? pickups[0].match(dateRegex) : null;
  const windowSunday = firstPickupDateMatch ? weekSunday_(new Date(firstPickupDateMatch[0])) : null;

  // Pickup dates and cities — Column D; time range — Column J–W
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
    const schedule = stopSchedule_(stop, windowSunday);
    if (schedule) {
      writes.push({ cellRef: `${schedule.col}${row + i + 1}`, value: schedule.timeRange, format: '@', newRow: i >= 2 });
    }
  });

  // Drop dates and cities — Column E; time range — Column J–W
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
    const schedule = stopSchedule_(stop, windowSunday);
    if (schedule) {
      writes.push({ cellRef: `${schedule.col}${row + i + 1}`, value: schedule.timeRange, format: '@', newRow: i >= 2 });
    }
  });

  const miles      = fetchRouteMiles_([...pickups, drops[0]]);
  const totalMiles = fetchRouteMiles_([...pickups, ...drops]);
  if (miles !== null || totalMiles !== null) {
    const lastDropSched = stopSchedule_(drops[drops.length - 1], windowSunday);
    const milesCol = lastDropSched
      ? String.fromCharCode(lastDropSched.col.charCodeAt(0) + 2)
      : 'Y'; // fallback: 2 past W (last possible schedule col)
    if (miles !== null)
      writes.push({ cellRef: `${milesCol}${row}`,     value: `${toHHMM_(miles / 50)} to del #1` });
    if (totalMiles !== null)
      writes.push({ cellRef: `${milesCol}${row + 1}`, value: `(${Math.round(totalMiles).toLocaleString()} mi)` });
  }

  return writes;
}

/**
 * Parses a stop block's datetime line and returns {col, timeRange}, or null
 * if the date falls outside the 2-week window anchored to windowSunday.
 * Window: J = windowSunday (day 0) … W = windowSunday + 13 (day 13).
 */
function stopSchedule_(stopText, windowSunday) {
  if (!windowSunday) return null;
  const m = stopText.match(
    /([A-Z][a-z]{2}\s+\d{1,2},\s+\d{4})\s+(\d{1,2}:\d{2}\s+(?:AM|PM))\s*-\s*(?:[A-Z][a-z]{2}\s+\d{1,2},\s+\d{4}\s+)?(\d{1,2}:\d{2}\s+(?:AM|PM))/
  );
  if (!m) return null;

  const dayOffset = Math.round((new Date(m[1]) - windowSunday) / 864e5);
  if (dayOffset < 0 || dayOffset > 13) return null;

  return {
    col:       String.fromCharCode(74 + dayOffset), // J(0)…W(13)
    timeRange: m[2] === m[3] ? to24Hour_(m[2]) : to24Hour_(m[2]) + ' - ' + to24Hour_(m[3])
  };
}

/** Returns midnight of the Sunday of the week containing the given date. */
function weekSunday_(date) {
  const d = new Date(date);
  d.setDate(d.getDate() - d.getDay());
  d.setHours(0, 0, 0, 0);
  return d;
}

/** Converts "h:mm AM/PM" to 24-hour HH:mm string e.g. "08:00", "15:30".
 * @param {string} timeStr */
function to24Hour_(timeStr) {
  const m = timeStr.trim().match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!m) return timeStr;
  let h = parseInt(m[1], 10);
  const min = m[2];
  const period = m[3].toUpperCase();
  if (period === 'AM' && h === 12) h = 0;
  if (period === 'PM' && h !== 12) h += 12;
  return `${String(h).padStart(2, '0')}:${min}`;
}

/**
 * Uses the Apps Script Maps service to drive pick0 → pick1 → … → pickN → drop0.
 * Returns total miles as a number, or null on failure.
 */
function fetchRouteMiles_(stops) {
  if (!stops || stops.length < 2) return null;
  const locs = stops.map(cityState_).filter(Boolean);
  if (locs.length < 2) return null;

  try {
    const finder = Maps.newDirectionFinder()
      .setOrigin(locs[0])
      .setDestination(locs[locs.length - 1])
      .setMode(Maps.DirectionFinder.Mode.DRIVING);
    for (let i = 1; i < locs.length - 1; i++) finder.addWaypoint(locs[i]);

    const result = finder.getDirections();
    if (!result.routes || !result.routes.length) return null;

    let totalMeters = 0;
    result.routes[0].legs.forEach(function(leg) { totalMeters += leg.distance.value; });
    return totalMeters / 1609.344;
  } catch (e) {
    return null;
  }
}

/** Converts a decimal hours value to "h:mm" string e.g. 12.5 → "12:30". */
function toHHMM_(hours) {
  const h = Math.floor(hours);
  const m = Math.round((hours - h) * 60);
  return `${h}:${String(m).padStart(2, '0')}`;
}

/** Extracts "City, ST" from a stop block's Address field. */
function cityState_(stopText) {
  const addrMatch = stopText.match(/Address:[\s\S]*?([^\n\r]+,\s[A-Z]{2}\s\d{5})/i);
  if (!addrMatch) return 'Location N/A';
  const m = addrMatch[1].trim().match(/([^,]+),\s*([A-Z]{2})\s+\d{5}/i);
  return (m && m[1] && m[2]) ? m[1].trim() + ', ' + m[2].toUpperCase() : 'City/ST N/A';
}
