function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('PCN-Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function normalizeStr(str) {
  if (!str) return "";
  return str.toString().toLowerCase()
            .replace(/ö/g, "oe")
            .replace(/ä/g, "ae")
            .replace(/ü/g, "ue")
            .replace(/ß/g, "ss")
            .replace(/straße/g, "str")
            .replace(/strasse/g, "str")
            .replace(/str\./g, "str")
            .replace(/[^a-z0-9]/g, ""); 
}

function parseISODate(ds) {
  if (!ds) return null;
  var p = ds.split("-");
  if(p.length === 3) return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
  return new Date(ds);
}

function aggregatePeriod(startObj, endObj, sheets, monthNames) {
  var stats = { total: 0, revenue: 0, staff: {}, locs: {} };
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var nameParts = sheet.getName().split(" ");
    if (nameParts.length < 2) continue; 
    
    var sheetMonth = monthNames.indexOf(nameParts[0]);
    var sheetYear = parseInt(nameParts[1], 10);
    if (sheetMonth === -1 || isNaN(sheetYear)) continue; 
    
    var sheetStart = new Date(sheetYear, sheetMonth, 1);
    var sheetEnd = new Date(sheetYear, sheetMonth + 1, 0);
    if (sheetEnd < startObj || sheetStart > endObj) continue; 
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 0; i < data.length; i++) {
      var standort = data[i][1];
      var strasse = data[i][2];
      var hausnummer = data[i][3];
      var stadt = data[i][4];
      var staff = data[i][5] || "Ohne Zuweisung";

      var gebuehrRaw = data[i][8];
      var gebuehr = parseFloat((gebuehrRaw||"").toString().replace(',', '.')) || 0;

      if (!standort) continue;

      var locationName = standort + " (" + stadt + ")";
      var address = ((strasse||"").toString().trim() + " " + (hausnummer||"").toString().trim()).trim();

      if (!stats.locs[locationName]) stats.locs[locationName] = { count: 0, revenue: 0, staffSet: {}, address: address };
      stats.locs[locationName].staffSet[staff] = true;

      for (var col = 10; col < data[i].length; col++) {
        var val = parseInt(data[i][col], 10);
        if (!isNaN(val) && val > 0) {
          var dayOfMonth = col - 9; 
          var colDate = new Date(sheetYear, sheetMonth, dayOfMonth, 12, 0, 0, 0);
          
          if (colDate >= startObj && colDate <= endObj) {
             var netto = (val * gebuehr) / 1.19;
             stats.total += val;
             stats.revenue += netto; 
             
             if (!stats.staff[staff]) stats.staff[staff] = { count: 0, revenue: 0 };
             stats.staff[staff].count += val;
             stats.staff[staff].revenue += netto;

             stats.locs[locationName].count += val;
             stats.locs[locationName].revenue += netto;
          }
        }
      }
    }
  }
  return stats;
}

function getReportData(startDateStr, endDateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets(); 
    var monthNames = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];
    
    var startP1 = parseISODate(startDateStr); if(startP1) startP1.setHours(0,0,0,0);
    var endP1 = parseISODate(endDateStr); if(endP1) endP1.setHours(23,59,59,999);
    
    var diffDays = Math.round((endP1.getTime() - startP1.getTime()) / 86400000) + 1;
    
    var endP2 = new Date(startP1.getFullYear(), startP1.getMonth(), startP1.getDate() - 1, 23, 59, 59, 999);
    var startP2 = new Date(endP2.getFullYear(), endP2.getMonth(), endP2.getDate() - diffDays + 1, 0, 0, 0, 0);

    var statsP1 = aggregatePeriod(startP1, endP1, sheets, monthNames);
    var statsP2 = aggregatePeriod(startP2, endP2, sheets, monthNames);

    var locArray = Object.keys(statsP1.locs).map(function(k) { return {name: k, count: statsP1.locs[k].count, revenue: statsP1.locs[k].revenue, staff: Object.keys(statsP1.locs[k].staffSet).join(", "), address: statsP1.locs[k].address}; });
    locArray.sort(function(a, b) { return b.count - a.count; });
    
    var staffArray = Object.keys(statsP1.staff).map(function(k) { return {name: k, count: statsP1.staff[k].count, revenue: statsP1.staff[k].revenue}; });
    staffArray.sort(function(a, b) { return b.count - a.count; });

    var trendVal = 0;
    var trendText = "0%";
    if (statsP2.total > 0) {
       trendVal = ((statsP1.total - statsP2.total) / statsP2.total) * 100;
       trendText = (trendVal > 0 ? "+" : "") + trendVal.toFixed(1) + "%";
    } else if (statsP1.total > 0) {
       trendText = "+100%";
    } else {
       trendText = "-";
    }
    
    var revTrendVal = 0;
    var revTrendText = "0%";
    if (statsP2.revenue > 0) {
       revTrendVal = ((statsP1.revenue - statsP2.revenue) / statsP2.revenue) * 100;
       revTrendText = (revTrendVal > 0 ? "+" : "") + revTrendVal.toFixed(1) + "%";
    } else if (statsP1.revenue > 0) {
       revTrendText = "+100%";
    } else {
       revTrendText = "-";
    }

    var formatObj = function(d) { return ("0"+d.getDate()).slice(-2)+"."+("0"+(d.getMonth()+1)).slice(-2)+"."+d.getFullYear(); };

    return {
      success: true,
      dateRange: formatObj(startP1) + " - " + formatObj(endP1),
      prevRange: formatObj(startP2) + " - " + formatObj(endP2),
      currentTotal: statsP1.total,
      previousTotal: statsP2.total,
      currentRevenue: statsP1.revenue,
      trend: trendText,
      trendIsPositive: trendVal >= 0,
      revTrend: revTrendText,
      revTrendIsPositive: revTrendVal >= 0,
      topStaff: staffArray,
      topLocs: locArray.slice(0, 15) 
    };

  } catch (e) {
    return {success: false, message: e.message};
  }
}

function getDashboardStats(startDateStr, endDateStr, baseDateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets(); 
    var monthNames = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];
    
    var startObj = startDateStr ? parseISODate(startDateStr) : null;
    var endObj = endDateStr ? parseISODate(endDateStr) : null;
    if (startObj) startObj.setHours(0,0,0,0);
    if (endObj) endObj.setHours(23,59,59,999);
    
    var baseDateObj = baseDateStr ? parseISODate(baseDateStr) : new Date();
    var targetActiveSheetName = monthNames[baseDateObj.getMonth()] + " " + baseDateObj.getFullYear();
    var activeSheet = ss.getSheetByName(targetActiveSheetName) || ss.getSheets()[0];
    
    var defaultToActiveSheet = (!startObj || !endObj);
    
    if (defaultToActiveSheet) {
       var sM = monthNames.indexOf(activeSheet.getName().split(" ")[0]);
       var sY = parseInt(activeSheet.getName().split(" ")[1], 10);
       startObj = new Date(sY, sM, 1);
       endObj = new Date(sY, sM + 1, 0);
       endObj.setHours(23,59,59,999);
    }
    
    var stats = aggregatePeriod(startObj, endObj, sheets, monthNames);
    
    var locationArray = Object.keys(stats.locs).map(function(key) { return {name: key, count: stats.locs[key].count, revenue: stats.locs[key].revenue, staff: Object.keys(stats.locs[key].staffSet).join(", "), address: stats.locs[key].address}; });
    locationArray.sort(function(a, b) { return b.count - a.count; });
    
    var staffArray = Object.keys(stats.staff).map(function(key) { return {name: key, count: stats.staff[key].count, revenue: stats.staff[key].revenue}; });
    staffArray.sort(function(a, b) { return b.count - a.count; });
    
    var formatObj = function(d) { return ("0"+d.getDate()).slice(-2)+"."+("0"+(d.getMonth()+1)).slice(-2)+"."+d.getFullYear(); };
    var contextLabel = formatObj(startObj) + " - " + formatObj(endObj);
    
    return {
      success: true,
      currentTotal: stats.total,
      totalRevenue: stats.revenue,
      staff: staffArray, 
      locations: locationArray,
      sheetName: contextLabel 
    };
    
  } catch (e) {
    return {success: false, message: e.message};
  }
}

function computeHalfTrend(series) {
  var len = series.length;
  var half = Math.floor(len / 2);
  if (half < 1) return { firstAvg: 0, secondAvg: 0, percent: null };

  var firstHalf = series.slice(0, half);
  var secondHalf = series.slice(len - half);

  var firstAvg = firstHalf.reduce(function(a, b) { return a + b; }, 0) / firstHalf.length;
  var secondAvg = secondHalf.reduce(function(a, b) { return a + b; }, 0) / secondHalf.length;

  var percent;
  if (firstAvg > 0) {
    percent = ((secondAvg - firstAvg) / firstAvg) * 100;
  } else if (secondAvg > 0) {
    percent = 100;
  } else {
    percent = null;
  }

  return { firstAvg: firstAvg, secondAvg: secondAvg, percent: percent };
}

function getTrendData(monthsCountStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var monthNames = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];

    var monthSheets = [];
    for (var s = 0; s < sheets.length; s++) {
      var nameParts = sheets[s].getName().split(" ");
      if (nameParts.length < 2) continue;
      var mIdx = monthNames.indexOf(nameParts[0]);
      var y = parseInt(nameParts[1], 10);
      if (mIdx === -1 || isNaN(y)) continue;
      monthSheets.push({ sheet: sheets[s], year: y, month: mIdx, sortKey: y * 12 + mIdx });
    }
    monthSheets.sort(function(a, b) { return a.sortKey - b.sortKey; });

    var n = parseInt(monthsCountStr, 10) || 6;
    var selected = monthSheets.slice(Math.max(0, monthSheets.length - n));

    if (selected.length < 2) {
      return { success: false, message: "Es werden mindestens 2 Monatsblätter für eine Trendanalyse benötigt." };
    }

    var monthLabels = selected.map(function(m) { return monthNames[m.month] + " " + m.year; });

    var locMap = {};
    var staffMap = {};

    for (var mi = 0; mi < selected.length; mi++) {
      var startObj = new Date(selected[mi].year, selected[mi].month, 1);
      var endObj = new Date(selected[mi].year, selected[mi].month + 1, 0);
      endObj.setHours(23, 59, 59, 999);

      var monthStats = aggregatePeriod(startObj, endObj, [selected[mi].sheet], monthNames);

      Object.keys(monthStats.locs).forEach(function(key) {
        if (!locMap[key]) {
          locMap[key] = {
            address: monthStats.locs[key].address,
            staffSet: {},
            ticketSeries: new Array(selected.length).fill(0),
            revenueSeries: new Array(selected.length).fill(0)
          };
        }
        locMap[key].ticketSeries[mi] = monthStats.locs[key].count;
        locMap[key].revenueSeries[mi] = monthStats.locs[key].revenue;
        Object.keys(monthStats.locs[key].staffSet).forEach(function(st) { locMap[key].staffSet[st] = true; });
      });

      Object.keys(monthStats.staff).forEach(function(key) {
        if (!staffMap[key]) {
          staffMap[key] = {
            ticketSeries: new Array(selected.length).fill(0),
            revenueSeries: new Array(selected.length).fill(0)
          };
        }
        staffMap[key].ticketSeries[mi] = monthStats.staff[key].count;
        staffMap[key].revenueSeries[mi] = monthStats.staff[key].revenue;
      });
    }

    var locations = Object.keys(locMap).map(function(key) {
      var l = locMap[key];
      var ticketTrend = computeHalfTrend(l.ticketSeries);
      var revenueTrend = computeHalfTrend(l.revenueSeries);
      return {
        name: key,
        address: l.address,
        staff: Object.keys(l.staffSet).join(", "),
        ticketSeries: l.ticketSeries,
        revenueSeries: l.revenueSeries,
        ticketFirstAvg: ticketTrend.firstAvg,
        ticketSecondAvg: ticketTrend.secondAvg,
        ticketTrendPercent: ticketTrend.percent,
        revenueTrendPercent: revenueTrend.percent
      };
    });

    var staffList = Object.keys(staffMap).map(function(key) {
      var st = staffMap[key];
      var ticketTrend = computeHalfTrend(st.ticketSeries);
      var revenueTrend = computeHalfTrend(st.revenueSeries);
      return {
        name: key,
        ticketSeries: st.ticketSeries,
        revenueSeries: st.revenueSeries,
        ticketFirstAvg: ticketTrend.firstAvg,
        ticketSecondAvg: ticketTrend.secondAvg,
        ticketTrendPercent: ticketTrend.percent,
        revenueTrendPercent: revenueTrend.percent
      };
    });

    return { success: true, months: monthLabels, locations: locations, staff: staffList };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Finds the chronologically most recent "MonthName Year" sheet. Used as the
// template for a brand-new month instead of ss.getActiveSheet(), which
// depends on whatever tab a human happened to have open and can silently
// clone a stale/incomplete location list.
function findLatestMonthSheet(ss, monthNames) {
  var sheets = ss.getSheets();
  var best = null;
  var bestKey = -1;
  for (var i = 0; i < sheets.length; i++) {
    var parts = sheets[i].getName().split(" ");
    if (parts.length < 2) continue;
    var mIdx = monthNames.indexOf(parts[0]);
    var y = parseInt(parts[1], 10);
    if (mIdx === -1 || isNaN(y)) continue;
    var key = y * 12 + mIdx;
    if (key > bestKey) { bestKey = key; best = sheets[i]; }
  }
  return best;
}

// Shared by previewPastedData and commitPastedData so the two can never
// disagree about what counts as a match. Returns the 0-indexed row in
// `data`, or -1 if nothing matches.
function findMatchingRowIndex(data, strasseInput, ortInput) {
  var inputStrasseNorm = normalizeStr(strasseInput);
  var inputOrtNorm = normalizeStr(ortInput);

  for (var r = 0; r < data.length; r++) {
    var strasseSheet = data[r][2] || "";
    var hnrSheet = data[r][3] || "";
    var stadtSheet = data[r][4] || "";

    var sheetStrasseFull = normalizeStr(strasseSheet.toString() + " " + hnrSheet.toString());
    var sheetStrasseJust = normalizeStr(strasseSheet.toString());
    var sheetOrtNorm = normalizeStr(stadtSheet.toString());

    var streetMatches = (sheetStrasseFull === inputStrasseNorm ||
                        (sheetStrasseJust !== "" && inputStrasseNorm.includes(sheetStrasseJust)) ||
                        (inputStrasseNorm !== "" && sheetStrasseJust.includes(inputStrasseNorm)));

    var cityMatches = false;
    if (sheetOrtNorm === "" || inputOrtNorm === "") {
      cityMatches = true;
    } else if (inputOrtNorm.includes(sheetOrtNorm) || sheetOrtNorm.includes(inputOrtNorm)) {
      cityMatches = true;
    }

    if (streetMatches && cityMatches) return r;
  }
  return -1;
}

function splitStrasseOrt(strasseInput, ortInput) {
  var sName = strasseInput;
  var sHnr = "";
  var matchStrasse = strasseInput.match(/(.*?)\s+(\d.*)/);
  if (matchStrasse) { sName = matchStrasse[1].trim(); sHnr = matchStrasse[2].trim(); }
  var sStadt = ortInput;
  var matchOrt = ortInput.match(/^(\d{5})\s+(.*)/);
  if (matchOrt) { sStadt = matchOrt[2].trim(); }
  return { name: sName, hnr: sHnr, stadt: sStadt };
}

// Read-only: reports which pasted rows would NOT match an existing Standort
// (and would therefore trigger a brand-new row) without writing anything, so
// the UI can ask for confirmation before any new Standort gets created.
function previewPastedData(dateString, parsedData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dateObj = parseISODate(dateString);
    if (!dateObj) return { success: false, message: "Ungültiges Datum." };

    var monthNames = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];
    var targetSheetName = monthNames[dateObj.getMonth()] + " " + dateObj.getFullYear();

    var sheet = ss.getSheetByName(targetSheetName);
    var matchSheet = sheet || findLatestMonthSheet(ss, monthNames) || ss.getSheets()[0];
    var data = matchSheet ? matchSheet.getDataRange().getValues() : [];

    var unmatched = [];
    var matchedCount = 0;

    for (var i = 0; i < parsedData.length; i++) {
      var ortInput = parsedData[i].ort || "";
      var strasseInput = parsedData[i].strasse || "";
      var anzahl = parsedData[i].anzahl;

      var rowIndex = findMatchingRowIndex(data, strasseInput, ortInput);
      if (rowIndex === -1) {
        var split = splitStrasseOrt(strasseInput, ortInput);
        unmatched.push({ ort: ortInput, strasse: strasseInput, anzahl: anzahl, proposedName: split.name, proposedHnr: split.hnr, proposedStadt: split.stadt });
      } else {
        matchedCount++;
      }
    }

    return {
      success: true,
      sheetExists: !!sheet,
      targetSheetName: targetSheetName,
      totalCount: parsedData.length,
      matchedCount: matchedCount,
      unmatched: unmatched
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Actually writes the pasted data. Re-runs the same matching logic against
// the sheet's current state at call time (never trusts a stale preview), so
// it stays correct even if the sheet changed between preview and commit.
function commitPastedData(dateString, parsedData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dateObj = parseISODate(dateString);
    if (!dateObj) return { success: false, message: "Ungültiges Datum." };

    var monthNames = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];
    var targetSheetName = monthNames[dateObj.getMonth()] + " " + dateObj.getFullYear();

    var sheet = ss.getSheetByName(targetSheetName);
    var wasCreated = false;

    if (!sheet) {
      var templateSheet = findLatestMonthSheet(ss, monthNames) || ss.getSheets()[0];
      sheet = templateSheet.copyTo(ss);
      sheet.setName(targetSheetName);
      var lastCol = sheet.getLastColumn();
      var lastRow = sheet.getLastRow();
      if (lastCol >= 11 && lastRow >= 1) { sheet.getRange(1, 11, lastRow, lastCol - 10).clearContent(); }
      wasCreated = true;
    }

    ss.setActiveSheet(sheet);
    var data = sheet.getDataRange().getValues();
    var dayOfMonth = dateObj.getDate();
    var targetColIndex = 10 + dayOfMonth;

    var matchCount = 0;
    var newLocCount = 0;

    for (var i = 0; i < parsedData.length; i++) {
      var ortInput = parsedData[i].ort || "";
      var strasseInput = parsedData[i].strasse || "";
      var anzahl = parsedData[i].anzahl;

      var rowIndex = findMatchingRowIndex(data, strasseInput, ortInput);

      if (rowIndex !== -1) {
        sheet.getRange(rowIndex + 1, targetColIndex).setValue(anzahl);
        matchCount++;
      } else {
        var currentLastRow = sheet.getLastRow();
        var newRowIndex = currentLastRow + 1;
        var split = splitStrasseOrt(strasseInput, ortInput);

        sheet.getRange(newRowIndex, 1).setValue("NEU");
        sheet.getRange(newRowIndex, 3).setValue(split.name);
        sheet.getRange(newRowIndex, 4).setValue(split.hnr);
        sheet.getRange(newRowIndex, 5).setValue(split.stadt);
        sheet.getRange(newRowIndex, 8).setValue("aktive");
        sheet.getRange(newRowIndex, targetColIndex).setValue(anzahl);

        if (currentLastRow >= 1) {
            var templateRange = sheet.getRange(currentLastRow, 1, 1, sheet.getLastColumn());
            templateRange.copyTo(sheet.getRange(newRowIndex, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
            templateRange.copyTo(sheet.getRange(newRowIndex, 1), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
        }

        var emptyRow = new Array(Math.max(12, targetColIndex)).fill("");
        emptyRow[0] = "NEU"; emptyRow[2] = split.name; emptyRow[3] = split.hnr; emptyRow[4] = split.stadt;
        data.push(emptyRow);

        matchCount++;
        newLocCount++;
      }
    }

    var msg = "🎉 " + matchCount + " Standorte am verarbeitet!";
    if (wasCreated) { msg = "✨ Neues Blatt '" + targetSheetName + "' wurde angelegt! " + msg; }

    return {success: true, message: msg, newLocCount: newLocCount};

  } catch (e) {
    return {success: false, message: "Fehler: " + e.message};
  }
}

function getLocations(baseDateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var baseDateObj = baseDateStr ? parseISODate(baseDateStr) : new Date();
  var monthNames = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];
  var targetSheetName = monthNames[baseDateObj.getMonth()] + " " + baseDateObj.getFullYear();
  var sheet = ss.getSheetByName(targetSheetName) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return { data: [], sheetName: sheet.getName() };
  
  var data = sheet.getRange(1, 1, lastRow, 10).getDisplayValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
     var row = data[i]; 
     row.push(i + 1); 
     result.push(row);
  }
  return { data: result, sheetName: sheet.getName() };
}

function getDailyData(monthStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var monthNames = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];

    var parts = monthStr.split("-");
    var y = parseInt(parts[0], 10);
    var m = parseInt(parts[1], 10) - 1;
    var targetSheetName = monthNames[m] + " " + y;
    var daysInMonth = new Date(y, m + 1, 0).getDate();

    var sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) return { success: true, rows: [], daysInMonth: daysInMonth, sheetName: targetSheetName, sheetExists: false };

    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return { success: true, rows: [], daysInMonth: daysInMonth, sheetName: sheet.getName(), sheetExists: true };

    var lastCol = Math.max(sheet.getLastColumn(), 10 + daysInMonth);
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    var rows = [];
    for (var i = 0; i < data.length; i++) {
      var standort = data[i][1];
      if (!standort) continue;

      var days = [];
      for (var d = 1; d <= daysInMonth; d++) {
        var val = data[i][9 + d];
        days.push((val === "" || val === null || val === undefined) ? "" : val);
      }

      rows.push({
        rowIndex: i + 1,
        standort: standort,
        stadt: data[i][4] || "",
        staff: data[i][5] || "",
        days: days
      });
    }

    return { success: true, rows: rows, daysInMonth: daysInMonth, sheetName: sheet.getName(), sheetExists: true };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateDailyValue(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(payload.sheetName);
    if (!sheet) return { success: false, message: "Tabellenblatt nicht gefunden." };

    var rowIndex = parseInt(payload.rowIndex, 10);
    var day = parseInt(payload.day, 10);
    if (isNaN(rowIndex) || rowIndex < 1) return { success: false, message: "Ungültige Zeile." };
    if (isNaN(day) || day < 1 || day > 31) return { success: false, message: "Ungültiger Tag." };
    var col = 10 + day;

    var rawValue = payload.value === null || payload.value === undefined ? "" : payload.value.toString().trim();
    var value = rawValue;
    if (rawValue !== "") {
      value = parseInt(rawValue, 10);
      if (isNaN(value) || value < 0) return { success: false, message: "Bitte eine gültige, positive Zahl eingeben." };
    }

    sheet.getRange(rowIndex, col).setValue(value);
    return { success: true };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveLocation(locData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(locData.sheetName) || ss.getSheets()[0];
    
    var gFloat = parseFloat(locData.gebuehr.toString().replace(',','.')) || "";
    var mFloat = parseFloat(locData.mahngebuehr.toString().replace(',','.')) || "";
    
    var rowValues = [ 
      locData.region, locData.standort, locData.strasse, locData.hausnummer, 
      locData.stadt, locData.staff, locData.start, locData.status,
      gFloat, mFloat
    ];
    
    if (locData.rowIndex) {
       sheet.getRange(locData.rowIndex, 1, 1, 10).setValues([rowValues]);
       return {success: true, message: "✏️ Standort in Zeile " + locData.rowIndex + " erfolgreich aktualisiert!"};
    } else {
       var lastRow = sheet.getLastRow();
       var nextRow = lastRow + 1;
       sheet.getRange(nextRow, 1, 1, 10).setValues([rowValues]);
       if (lastRow >= 1) {
         var templateRange = sheet.getRange(lastRow, 1, 1, 10);
         var targetRange = sheet.getRange(nextRow, 1, 1, 10);
         templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
         templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
       }
       return {success: true, message: "✅ Neuer Standort erfolgreich hinzugefügt!"};
    }
  } catch(e) {
    return {success: false, message: e.message};
  }
}

function deleteLocation(locData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(locData.sheetName);
    if (!sheet) return {success: false, message: "Fehler: Tabellenblatt nicht gefunden."};
    
    if (locData.rowIndex) {
       sheet.deleteRow(locData.rowIndex);
       return {success: true, message: "🗑️ Standort in Zeile " + locData.rowIndex + " wurde restlos entfernt!"};
    } else {
       return {success: false, message: "Fehler: Kein Zeilenindex angegeben."};
    }
  } catch(e) {
    return {success: false, message: e.message};
  }
}
