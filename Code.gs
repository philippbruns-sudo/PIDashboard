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
      var stadt = data[i][4];    
      var staff = data[i][5] || "Ohne Zuweisung";
      
      var gebuehrRaw = data[i][8];
      var gebuehr = parseFloat((gebuehrRaw||"").toString().replace(',', '.')) || 0;
      
      if (!standort) continue;

      var locationName = standort + " (" + stadt + ")";

      if (!stats.locs[locationName]) stats.locs[locationName] = { count: 0, revenue: 0, staffSet: {} };
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

    var locArray = Object.keys(statsP1.locs).map(function(k) { return {name: k, count: statsP1.locs[k].count, revenue: statsP1.locs[k].revenue, staff: Object.keys(statsP1.locs[k].staffSet).join(", ")}; });
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
    
    var locationArray = Object.keys(stats.locs).map(function(key) { return {name: key, count: stats.locs[key].count, revenue: stats.locs[key].revenue, staff: Object.keys(stats.locs[key].staffSet).join(", ")}; });
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

function processPastedData(dateString, parsedData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dateObj = parseISODate(dateString);
    var monthNames = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];
    var targetSheetName = monthNames[dateObj.getMonth()] + " " + dateObj.getFullYear();
    
    var sheet = ss.getSheetByName(targetSheetName);
    var wasCreated = false;
    
    if (!sheet) {
      var templateSheet = ss.getActiveSheet();
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
      
      var inputStrasseNorm = normalizeStr(strasseInput);
      var inputOrtNorm = normalizeStr(ortInput);
      var matchFound = false;
      
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
        
        if (streetMatches && cityMatches) {
          sheet.getRange(r + 1, targetColIndex).setValue(anzahl);
          matchCount++;
          matchFound = true;
          break; 
        }
      }
      
      if (!matchFound) {
        var currentLastRow = sheet.getLastRow();
        var newRowIndex = currentLastRow + 1;
        var sName = strasseInput;
        var sHnr = "";
        var matchStrasse = strasseInput.match(/(.*?)\s+(\d.*)/);
        if (matchStrasse) { sName = matchStrasse[1].trim(); sHnr = matchStrasse[2].trim(); }
        var sStadt = ortInput;
        var matchOrt = ortInput.match(/^(\d{5})\s+(.*)/);
        if (matchOrt) { sStadt = matchOrt[2].trim(); }
        
        sheet.getRange(newRowIndex, 1).setValue("NEU");      
        sheet.getRange(newRowIndex, 3).setValue(sName);      
        sheet.getRange(newRowIndex, 4).setValue(sHnr);       
        sheet.getRange(newRowIndex, 5).setValue(sStadt);     
        sheet.getRange(newRowIndex, 8).setValue("aktive");   
        sheet.getRange(newRowIndex, targetColIndex).setValue(anzahl);
        
        if (currentLastRow >= 1) {
            var templateRange = sheet.getRange(currentLastRow, 1, 1, sheet.getLastColumn());
            templateRange.copyTo(sheet.getRange(newRowIndex, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
            templateRange.copyTo(sheet.getRange(newRowIndex, 1), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
        }
        
        var emptyRow = new Array(Math.max(12, targetColIndex)).fill("");
        emptyRow[0] = "NEU"; emptyRow[2] = sName; emptyRow[3] = sHnr; emptyRow[4] = sStadt;
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
