function updateTripStatus() {
  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell     = sheet.getActiveCell();
  const row      = cell.getRow();
  const col      = cell.getColumn();
  const colLetter = String.fromCharCode(64 + col);
  const cellRef  = colLetter + row;
  const belowRef = colLetter + (row + 1);
  const existing = {
    time: sheet.getRange(row,     col).getValue().toString(),
    dist: sheet.getRange(row + 1, col).getValue().toString()
  };
  const load    = findLoadNumber_(sheet, row);
  const pickups = load ? stopCities_(sheet, load.row, 4) : [];
  const drops   = load ? stopCities_(sheet, load.row, 5) : [];

  const html = HtmlService.createHtmlOutput(
    buildCellInfoHtml_(load, pickups, drops, cellRef, belowRef, existing))
    .setWidth(340)
    .setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, 'Update Trip Status');
}

function writeStatusToSheet(timeStr, distStr, cellRef, belowRef, origin) {
  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const now      = new Date().toUTCString();
  const user     = Session.getActiveUser().getEmail() || 'unknown';
  const noteHead = `${now}\n${origin}\n${user}\n-----------\n`;
  const timeCell = sheet.getRange(cellRef);
  const existing = timeCell.getNote() || '';
  timeCell.setValue(timeStr);
  timeCell.setNote(noteHead + existing);
  sheet.getRange(belowRef).setValue(distStr);
}

function buildCellInfoHtml_(load, pickups, drops, cellRef, belowRef, existing) {
  const loadNum   = load ? load.loadNumber : 'not found';
  const del1      = drops.length ? drops[0] : 'unknown';
  const pickStr   = pickups.length ? pickups.join(', ') : 'none';
  const dropStr   = drops.length  ? drops.join(', ')   : 'none';
  const warnTime  = existing.time ? `⚠ ${cellRef} has: "${existing.time}"` : '';
  const warnDist  = existing.dist ? `⚠ ${belowRef} has: "${existing.dist}"` : '';

  return `<!DOCTYPE html><html><body style="font-family:sans-serif;font-size:13px;padding:12px">
  <table style="width:100%;border-collapse:collapse;margin-bottom:10px">
    <tr><td style="color:#888;width:80px">Load #</td>
        <td><strong>${loadNum}</strong></td></tr>
    <tr><td style="color:#888">Pickups</td>
        <td>${pickStr}</td></tr>
    <tr><td style="color:#888">Drops</td>
        <td>${dropStr}</td></tr>
  </table>
  <hr style="margin-bottom:10px">
  <table style="width:100%;border-collapse:collapse">
    <tr><td style="color:#888;width:90px">Current city</td>
        <td><input id="origin" type="text" placeholder="e.g. Fresno, CA"
                   style="width:100%;box-sizing:border-box"
                   onkeydown="if(event.key==='Enter')calc()"></td></tr>
    <tr><td style="color:#888;padding-top:6px">Destination</td>
        <td style="padding-top:6px">
          <input id="dest" type="text" value="${del1}"
                 style="width:100%;box-sizing:border-box">
        </td></tr>
    <tr><td style="color:#888;padding-top:6px">Speed</td>
        <td style="padding-top:6px">
          <input id="speed" type="number" value="50" min="1"
                 style="width:70px;box-sizing:border-box"> mph
        </td></tr>
  </table>
  <div style="margin-top:10px">
    <div style="display:flex;align-items:center;gap:6px;min-height:16px">
      <span id="time" style="color:#333"></span>
      <input id="time-suffix" type="text" value=" to del #1"
             style="display:none;width:90px;font-size:12px;box-sizing:border-box">
    </div>
    <div id="result" style="color:#333;min-height:16px"></div>
    <div id="hint"   style="margin-top:2px;font-size:11px;color:#b06000;min-height:14px"></div>
    <button style="margin-top:8px" onclick="calc()">Calculate</button>
  </div>
  <div id="save-section" style="display:none;margin-top:10px;border-top:1px solid #ddd;padding-top:8px;font-size:12px">
    <div>Save to: <strong>${cellRef}</strong> (Time) · <strong>${belowRef}</strong> (Distance)</div>
    <div id="warn-time" style="color:#b06000">${warnTime}</div>
    <div id="warn-dist" style="color:#b06000">${warnDist}</div>
    <button style="margin-top:6px" onclick="save()">Update Sheet</button>
  </div>
  <script>
    var _timeStr = '', _distStr = '', _resolvedOrigin = '';
    function toHHMM(hours) {
      const h = Math.floor(hours);
      const m = Math.round((hours - h) * 60);
      return h + ':' + (m < 10 ? '0' : '') + m;
    }
    function calc() {
      const origin = document.getElementById('origin').value.trim();
      const dest   = document.getElementById('dest').value.trim();
      const speed  = parseFloat(document.getElementById('speed').value) || 50;
      if (!origin || !dest) return;
      document.getElementById('time').textContent   = 'Calculating…';
      document.getElementById('result').textContent = '';
      document.getElementById('hint').textContent   = '';
      document.getElementById('save-section').style.display = 'none';
      google.script.run
        .withSuccessHandler(function(r) {
          if (r.miles !== null) {
            _timeStr = toHHMM(r.miles / speed);
            _distStr = Math.round(r.miles).toLocaleString() + ' mi';
            document.getElementById('time').textContent          = 'Time: ' + _timeStr;
            document.getElementById('time-suffix').style.display = 'inline';
            document.getElementById('result').textContent        = 'Distance: ' + _distStr;
            document.getElementById('save-section').style.display = 'block';
          } else {
            document.getElementById('time').textContent   = 'Time: unavailable';
            document.getElementById('result').textContent = 'Distance: unavailable';
          }
          _resolvedOrigin = r.hint || origin;
          document.getElementById('hint').textContent = r.hint ? 'Did you mean: ' + r.hint : '';
        })
        .drivingMilesPublic(origin, dest);
    }
    function save() {
      const suffix = document.getElementById('time-suffix').value;
      google.script.run
        .withSuccessHandler(function() { google.script.host.close(); })
        .writeStatusToSheet(_timeStr + suffix, _distStr, '${cellRef}', '${belowRef}', _resolvedOrigin);
    }
  </script>
</body></html>`;
}

/** Called from HTML via google.script.run — returns {miles, hint} */
function drivingMilesPublic(origin, destination) {
  const geo     = Maps.newGeocoder().geocode(origin);
  const top     = geo.results && geo.results[0];
  const hint    = top && (top.partial_match || top.formatted_address.toLowerCase().indexOf(origin.toLowerCase()) === -1)
                  ? top.formatted_address
                  : null;
  const miles   = drivingMiles_(origin, destination);
  return { miles: miles, hint: hint };
}

function drivingMiles_(origin, destination) {
  try {
    const result = Maps.newDirectionFinder()
      .setOrigin(origin)
      .setDestination(destination)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();
    if (!result.routes || !result.routes.length) return null;
    let meters = 0;
    result.routes[0].legs.forEach(function(leg) { meters += leg.distance.value; });
    return meters / 1609.344;
  } catch (e) {
    return null;
  }
}

function stopCities_(sheet, loadRow, col) {
  const cities = [];
  const values = sheet.getRange(loadRow + 1, col, 6, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const val = values[i][0];
    if (!val || val instanceof Date) break;
    cities.push(String(val));
  }
  return cities;
}

function findLoadNumber_(sheet, fromRow) {
  const startRow = Math.max(1, fromRow - 6);
  const values = sheet.getRange(startRow, 1, fromRow - startRow + 1, 1).getValues();
  for (let r = fromRow - startRow; r >= 0; r--) {
    const val = String(values[r][0]).replace(/^`/, '');
    if (/^\d{5,}$/.test(val)) return { row: startRow + r, loadNumber: val };
  }
  return null;
}
