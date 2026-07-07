// Your Google Sheet ID
const SHEET_ID = '1HCHfqYJIvrIH9OOeNchGVQAQrHXfytFl1iegz22qS4Y';

// This Apps Script deployment is now purely a JSON data API. The actual app
// page (HTML/CSS/JS, icon, manifest) is hosted separately on GitHub Pages,
// which fetch()es data from here. Apps Script's web app hosting can't
// support a real custom home-screen icon due to how it wraps served pages,
// so the front end moved elsewhere; this file just serves data now.
function doGet(e) {
  const page = e && e.parameter && e.parameter.page;

  if (page === 'data') {
    return ContentService.createTextOutput(JSON.stringify(getBuildingData()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (page === 'refresh') {
    return ContentService.createTextOutput(JSON.stringify(refreshBuildingData()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput('Building Locator data API is running. Use ?page=data');
}

// Follows redirects on shortened Google Maps links (goo.gl / maps.app.goo.gl)
// to get the real URL, which usually contains @lat,lng we can extract
function resolveShortUrl(url) {
  let current = url;
  for (let hop = 0; hop < 5; hop++) {
    try {
      const response = UrlFetchApp.fetch(current, {
        followRedirects: false,
        muteHttpExceptions: true
      });
      const headers = response.getHeaders();
      const location = headers['Location'] || headers['location'];
      if (!location) break;
      current = location;
    } catch (e) {
      break;
    }
  }
  return current;
}

// Builds Google + Apple Maps links that launch turn-by-turn directions directly
function buildMapLinks(raw) {
  if (!raw) return { google: '', apple: '' };

  let text = String(raw);

  const urlMatch = text.match(/https?:\/\/[^\s]+/);
  let existingUrl = urlMatch ? urlMatch[0] : '';

  // Resolve short links (goo.gl / maps.app.goo.gl) to get the real URL with coordinates
  if (existingUrl && (existingUrl.includes('goo.gl') || existingUrl.includes('maps.app'))) {
    const resolved = resolveShortUrl(existingUrl);
    if (resolved && resolved !== existingUrl) {
      text = resolved;
      existingUrl = resolved;
    }
  }

  // Try to pull a lat,lng pair out of the text (handles plain "lat,lng",
  // Google's "@lat,lng" share-link format, and "lat,+lng" from resolved short links)
  const coordMatch = text.match(/(-?\d{1,3}\.\d+)[,+\s]+(-?\d{1,3}\.\d+)/);
  const lat = coordMatch ? coordMatch[1] : null;
  const lng = coordMatch ? coordMatch[2] : null;

  let destination = '';
  if (lat && lng) {
    destination = `${lat},${lng}`;
  } else if (existingUrl) {
    const qMatch = existingUrl.match(/[?&]q=([^&]+)/);
    if (qMatch) {
      destination = decodeURIComponent(qMatch[1]);
    } else {
      const placeMatch = existingUrl.match(/\/place\/([^\/@]+)/);
      destination = placeMatch ? decodeURIComponent(placeMatch[1].replace(/\+/g, ' ')) : '';
    }
  } else {
    destination = text;
  }

  if (!destination) {
    return { google: existingUrl || text, apple: existingUrl || text };
  }

  const google = `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(destination)}&travelmode=driving`;
  const apple = `https://maps.apple.com/?daddr=${encodeURIComponent(destination)}&dirflg=d`;

  return { google, apple };
}

function getBuildingData() {
  const cache = CacheService.getScriptCache();
  const metaStr = cache.get('buildingData_meta');

  if (metaStr) {
    try {
      const meta = JSON.parse(metaStr);
      let combined = [];
      let allFound = true;
      for (let c = 0; c < meta.chunks; c++) {
        const chunk = cache.get('buildingData_chunk_' + c);
        if (!chunk) { allFound = false; break; }
        combined = combined.concat(JSON.parse(chunk));
      }
      if (allFound) return combined;
    } catch (e) {
      // fall through and rebuild
    }
  }

  const data = buildAllData();
  cacheBuildingData(data);
  return data;
}

// Splits data into small chunks (CacheService has a 100KB-per-key limit)
// and stores them for 1 hour so repeat loads are instant
function cacheBuildingData(data) {
  const cache = CacheService.getScriptCache();
  const chunkSize = 50;
  const chunks = [];
  for (let i = 0; i < data.length; i += chunkSize) {
    chunks.push(data.slice(i, i + chunkSize));
  }
  chunks.forEach((chunk, idx) => {
    cache.put('buildingData_chunk_' + idx, JSON.stringify(chunk), 3600);
  });
  cache.put('buildingData_meta', JSON.stringify({ chunks: chunks.length }), 3600);
}

function clearBuildingCache() {
  const cache = CacheService.getScriptCache();
  const metaStr = cache.get('buildingData_meta');
  if (metaStr) {
    try {
      const meta = JSON.parse(metaStr);
      for (let c = 0; c < meta.chunks; c++) cache.remove('buildingData_chunk_' + c);
    } catch (e) {
      // ignore
    }
    cache.remove('buildingData_meta');
  }
}

// Called from the app's refresh button to force fresh data from the sheet
function refreshBuildingData() {
  clearBuildingCache();
  const data = buildAllData();
  cacheBuildingData(data);
  return data;
}

function buildAllData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheets = ss.getSheets();
  let allData = [];

  sheets.forEach(sheet => {
    try {
      const sheetName = sheet.getName();
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;

      let headerRowIndex = 0;
      let headers = data[0];

      for (let r = 0; r < Math.min(6, data.length); r++) {
        const potentialHeaders = data[r];
        const hasTextHeaders = potentialHeaders.some(cell => {
          const c = String(cell).toLowerCase().trim();
          if (c.includes('http')) return false;
          return c.includes('building') || c.includes('name') || c.includes('map') ||
                 c.includes('location') || c.includes('facility') || c.includes('number') ||
                 c.includes('address') || c.includes('place');
        });
        if (hasTextHeaders) {
          headerRowIndex = r;
          headers = potentialHeaders;
          break;
        }
      }

      let buildingNumCol = -1;
      let nameCol = -1;
      let mapLinkCol = -1;

      headers.forEach((header, index) => {
        const h = String(header).toLowerCase().trim();

        if (buildingNumCol === -1) {
          if (h.includes('building') || h.includes('bldg') || h.includes('number') || h.includes('num') || h === '#') {
            buildingNumCol = index;
          }
        }
        if (nameCol === -1) {
          if (h.includes('name') || h.includes('title') || h.includes('description') || h.includes('facility')) {
            nameCol = index;
          }
        }
        if (mapLinkCol === -1) {
          if (h.includes('map') && (h.includes('link') || h.includes('it'))) {
            mapLinkCol = index;
          } else if (h === 'map' || h === 'maps' || h === 'directions' || h === 'gmap' ||
                     h.includes('address') || h.includes('location') || h.includes('place')) {
            mapLinkCol = index;
          }
        }
      });

      if (buildingNumCol === -1 && headers.length > 0) buildingNumCol = 0;

      if (mapLinkCol === -1) {
        for (let sampleRow = headerRowIndex + 1; sampleRow < Math.min(headerRowIndex + 5, data.length); sampleRow++) {
          const sample = data[sampleRow];
          if (!sample) continue;
          for (let col = 0; col < sample.length; col++) {
            const val = String(sample[col] || '').trim().toLowerCase();
            if (val.startsWith('http')) {
              mapLinkCol = col;
              break;
            }
          }
          if (mapLinkCol !== -1) break;
        }
      }

      if (nameCol === -1) {
        for (let idx = 0; idx < headers.length; idx++) {
          if (idx !== buildingNumCol && idx !== mapLinkCol) {
            nameCol = idx;
            break;
          }
        }
      }

      const dataRange = sheet.getDataRange();
      const formulas = dataRange.getFormulas();
      const richTextValues = dataRange.getRichTextValues();

      let rowsAdded = 0;
      for (let i = headerRowIndex + 1; i < data.length; i++) {
        let row = data[i];
        if (!row || row.every(cell => !cell || String(cell).trim() === '')) continue;

        let buildingValue = buildingNumCol >= 0 ? String(row[buildingNumCol] || '').trim() : '';
        if (!buildingValue) continue;

        // HOSPITALS-specific fix: names are split across two rows
        if (sheetName.toUpperCase() === 'HOSPITALS') {
          const restOfRowEmpty = row.every((cell, idx) =>
            idx === buildingNumCol || !cell || String(cell).trim() === ''
          );
          if (restOfRowEmpty && i + 1 < data.length) {
            const nextRow = data[i + 1];
            const nextBuildingValue = buildingNumCol >= 0 ? String(nextRow[buildingNumCol] || '').trim() : '';
            if (nextBuildingValue) {
              buildingValue = buildingValue + ' ' + nextBuildingValue;
              row = nextRow;
              i++;
            }
          }
        }

        let rawLocation = '';

        // Special handling for NBSD tab coordinate format
        if (sheetName.toUpperCase() === 'NBSD') {
          for (let col = 0; col < row.length; col++) {
            const cellValue = String(row[col] || '').trim();
            if (!cellValue) continue;

            // Format A: degree-symbol with N/S/E/W letters, e.g. 32.69°N 117.15°W
            if (cellValue.match(/\d+\.\d+.*[NS].*\d+\.\d+.*[EW]/i)) {
              const latMatch = cellValue.match(/(\d+\.\d+)\s*[NS]/i);
              const lonMatch = cellValue.match(/(\d+\.\d+)\s*[EW]/i);
              if (latMatch && lonMatch) {
                let lat = parseFloat(latMatch[1]);
                let lon = parseFloat(lonMatch[1]);
                if (cellValue.match(/S/i)) lat = -lat;
                if (cellValue.match(/W/i)) lon = -lon;
                rawLocation = `${lat},${lon}`;
              }
              break;
            }

            // Format B: plain decimal "lat, lon" pair, no letters/symbols
            const plainMatch = cellValue.match(/^(-?\d{1,3}\.\d+)\s*,\s*(-?\d{1,3}\.\d+)$/);
            if (plainMatch) {
              rawLocation = `${plainMatch[1]},${plainMatch[2]}`;
              break;
            }
          }
        }

        // Standard extraction for all other tabs
        if (!rawLocation && mapLinkCol >= 0) {
          const formula = formulas[i][mapLinkCol];
          if (formula && formula.includes('HYPERLINK')) {
            const match = formula.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
            if (match && match[1]) rawLocation = match[1];
          }
          if (!rawLocation) {
            const richText = richTextValues[i][mapLinkCol];
            if (richText && richText.getLinkUrl()) rawLocation = richText.getLinkUrl();
          }
          if (!rawLocation) {
            const cellValue = String(row[mapLinkCol] || '').trim();
            if (cellValue) rawLocation = cellValue;
          }
        }

        const { google, apple } = buildMapLinks(rawLocation);

        allData.push({
          tab: sheetName,
          buildingNumber: buildingValue,
          name: nameCol >= 0 ? String(row[nameCol] || '').trim() : '',
          raw: rawLocation,
          google: rawLocation ? google : '',
          apple: rawLocation ? apple : ''
        });
        rowsAdded++;
      }

      Logger.log('Added ' + rowsAdded + ' rows from ' + sheetName);

    } catch (error) {
      Logger.log('Error processing sheet ' + sheet.getName() + ': ' + error);
    }
  });

  return allData;
}

// Debug: dump raw row data for a given tab so we can see exactly what's in each cell
function inspectTab(tabName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    Logger.log(tabName + ' sheet not found');
    return;
  }
  const data = sheet.getDataRange().getValues();
  Logger.log('Header row: ' + JSON.stringify(data[0]));
  for (let i = 1; i <= 5 && i < data.length; i++) {
    Logger.log('Row ' + (i + 1) + ': ' + JSON.stringify(data[i]));
  }
}

function inspectNBSD() {
  inspectTab('NBSD');
}

// Test function - run this manually to check the logs
function testGetData() {
  const data = getBuildingData();

  const byTab = {};
  data.forEach(item => {
    if (!byTab[item.tab]) byTab[item.tab] = { total: 0, withLink: 0 };
    byTab[item.tab].total++;
    if (item.google) byTab[item.tab].withLink++;
  });

  const summaryLines = Object.keys(byTab).map(tab =>
    tab + ': ' + byTab[tab].withLink + '/' + byTab[tab].total + ' have links'
  );

  const withLinks = data.filter(item => item.google).length;

  Logger.log(
    'SUMMARY | Total buildings: ' + data.length +
    ' | Total with map links: ' + withLinks + '/' + data.length +
    ' || ' + summaryLines.join(' | ')
  );

  ['NBSD', 'HOSPITALS'].forEach(tabName => {
    const samples = data.filter(item => item.tab === tabName).slice(0, 3);
    samples.forEach(s => {
      Logger.log(tabName + ' SAMPLE | building=' + s.buildingNumber +
        ' | raw="' + s.raw + '" | google="' + s.google + '" | apple="' + s.apple + '"');
    });
  });
}
