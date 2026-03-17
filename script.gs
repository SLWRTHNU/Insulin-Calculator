/** ======================================================
 *  MEAL CALCULATOR (LEGACY SHEET) — CONSOLIDATED SCRIPT
 *  Single .gs file. All functions share the same scope.
 *
 *  SECTIONS:
 *    1.  CONFIG & CONSTANTS
 *    2.  MENU  (onOpen / onEdit router)
 *    3.  MEAL VALUES  (weight↔carbs, F7/F8 totals)
 *    4.  BOLUS CALCULATION  (F13 formula-driven, COB-aware)
 *    5.  BG SYNC  (Nightscout → meal sheets)
 *    6.  POST-BOLUS TRACKING
 *    7.  DAILY EXPORT
 *    8.  FOOD MANAGER  (Add to Food Chart sidebar backend)
 *    9.  FOOD SEARCH
 *   10.  LEGACY LOG IMPORT
 *   11.  MEAL BUILDER
 *   12.  COPY NOTE
 *   13.  CLEAR / RESTORE HELPERS
 *   14.  SHARED HELPERS
 * ====================================================== */


/* ======================================================
   1.  CONFIG & CONSTANTS
   ====================================================== */

var NS_BASE_URL            = 'https://sennaloop-673ad2782247.herokuapp.com';
var NS_TOKEN               = '';
var EXPORT_FOLDER_ID       = '1vSuFetWRZBd3yeJWfA5MSKhn_yXiXJqJ';
var ROOT_FOLDER_ID         = '1MUasRBNZeNo_EWpMja6iQmJjF9uvtSiz';
var INDEX_SHEET_NAME       = 'Food Search Index';
var LEGACY_LOGS_FOLDER_ID  = '1e35dgX4rzN_R38Fm8q5n2HcW7oxYX_M5';

var MEAL_SHEET_NAMES = [
  'Breakfast', 'Morning Snack', 'Lunch',
  'Afternoon Snack', 'Dinner', 'Evening Snack'
];

var PB_OFFSETS_MIN = [30, 60, 90, 120, 150, 180];
var PB_START_ROW   = 5;


/* ======================================================
   2.  MENU — onOpen / onEdit router
   ====================================================== */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Data Management')
    .addItem('Update BG & IOB Now',     'updateMealSheetsBackgroundSync')
    .addItem('Export Today Now',        'DailyExport_exportNow')
    .addItem('Clear Active Meal Sheet', 'ClearMealChart')
    .addItem('Rebuild Search Index',    'FoodSearch_rebuildIndex')
    .addToUi();
  ui.createMenu('Tools & Sidebars')
    .addItem('Food Search',       'FoodSearch_showSidebar')
    .addItem('Add to Food Chart', 'showFoodManagerSidebar')
    .addItem('Meal Builder',      'MealBuilder_showSidebar')
    .addToUi();
}

function onEdit(e) {
  if (!e) return;
  try { onEdit_UpdateMealValues(e);  } catch(err) { console.warn('onEdit_UpdateMealValues:',  err.message); }
  try { onEdit_PostBolusTracking(e); } catch(err) { console.warn('onEdit_PostBolusTracking:', err.message); }
}


/* ======================================================
   3.  MEAL VALUES  (weight↔carbs, F7/F8 totals)
   ====================================================== */

function onEdit_UpdateMealValues(e) {
  if (!e) return;
  var sh   = e.range.getSheet();
  var name = sh.getName();
  if (MEAL_SHEET_NAMES.indexOf(name) === -1) return;

  var r = e.range.getRow();
  var c = e.range.getColumn();
  if (c !== 3) return;

  var BLOCK          = 6;
  var FIRST_FOOD_ROW = 2;
  var MAX_ITEMS      = 15;
  var mod            = (r - FIRST_FOOD_ROW) % BLOCK;

  // Bi-directional weight ↔ carbs math
  if (mod === 3 || mod === 4) {
    var startOfBlock = r - mod;
    var factor = numFromDisplay_(sh.getRange(startOfBlock + 1, 3).getDisplayValue());
    if (mod === 3) {
      sh.getRange(r + 1, 3).setValue(round1_(numFromDisplay_(e.range.getDisplayValue()) * factor));
    } else {
      var carbs = numFromDisplay_(e.range.getDisplayValue());
      sh.getRange(r - 1, 3).setValue(factor > 0 ? Math.round(carbs / factor) : 0);
    }
  }

  // Recalculate F7 (total net carbs) and F8 (avg absorption)
  var totalNetCarbs = 0;
  var sumProduct    = 0;
  var itemsFound    = 0;

  for (var i = 0; i < MAX_ITEMS; i++) {
    var base     = FIRST_FOOD_ROW + (i * BLOCK);
    var foodName = sh.getRange(base,     3).getDisplayValue().trim();
    var weight   = numFromDisplay_(sh.getRange(base + 3, 3).getDisplayValue());
    var netCarbs = numFromDisplay_(sh.getRange(base + 4, 3).getDisplayValue());
    var absVal   = numFromDisplay_(sh.getRange(base + 2, 3).getDisplayValue());

    if (foodName !== '' || weight > 0) {
      itemsFound++;
      totalNetCarbs += netCarbs;
      if (absVal >= 0.5) sumProduct += absVal * netCarbs;
    }
  }

  if (itemsFound === 0) {
    sh.getRange(7, 6).setValue('');
    sh.getRange(8, 6).setValue('');
  } else {
    sh.getRange(7, 6).setValue(round1_(totalNetCarbs));
    var avgAbs = totalNetCarbs > 0 ? sumProduct / totalNetCarbs : 0;
    sh.getRange(8, 6).setValue(round05_(avgAbs));
  }

  // F13 is formula-driven — always restore after any edit to ensure it's present
  restoreF13Formula_(sh);
}


/* ======================================================
   4.  BOLUS CALCULATION  (F13 formula-driven, COB-aware)
   ====================================================== */

/**
 * Restores the F13 bolus formula.
 * Accounts for IOB (F10) and COB converted to insulin units (F11 / F2).
 * Called after any recalculation and after clear operations.
 */
function restoreF13Formula_(sh) {
  sh.getRange('F13').setFormula(
    '=IF(OR(F2<=0,F4<=0),"",IF(AND(F6=0,F7=0),"",FLOOR(MAX(0,' +
    'IF(F2>0,F7/F2,0)+IF(AND(F4>0,F3>0,F6>0),(F6-F3)/F4,0)' +
    '-IF(F10>0,F10,0)-IF(F11>0,F11/F2,0)),0.05)))');
}


/* ======================================================
   5.  BG SYNC  (Nightscout → meal sheets)
   ====================================================== */

/** Creates the 1-minute background sync trigger. Safe to call multiple times. */
function initializeSync() {
  deleteAllTimeTriggers_();
  ScriptApp.newTrigger('updateMealSheetsBackgroundSync').timeBased().everyMinutes(1).create();
  SpreadsheetApp.getActive().toast('Sync initialized — running every 60 seconds (6am–8pm).');
}

function deleteAllTimeTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'updateMealSheetsBackgroundSync') ScriptApp.deleteTrigger(t);
  });
}

function updateMealSheetsBackgroundSync() {
  var now  = new Date();
  var hour = now.getHours();
  if (hour < 6 || hour >= 20) { clearAgeLabelsOnly_(); return; }

  var ss          = SpreadsheetApp.getActive();
  var bgResult    = ns_fetchLatestBG_();
  var profileData = ns_fetchProfileData_();
  var boardData   = ns_fetchBoardData_();

  MEAL_SHEET_NAMES.forEach(function(sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    // Skip locked sheets (bolus already taken)
    var lockTime = String(sh.getRange('F16').getDisplayValue() || '').trim();
    if (lockTime) { sh.getRange('G6').setValue(''); return; }

    if (profileData) {
      var schedule = profileData.store[profileData.defaultProfile];
      sh.getRange('F2').setValue(ns_getValueForCurrentTime_(schedule.carbratio));
      sh.getRange('F4').setValue(ns_getValueForCurrentTime_(schedule.sens));
      var low  = ns_getValueForCurrentTime_(schedule.target_low);
      var high = ns_getValueForCurrentTime_(schedule.target_high);
      if (low !== null && high !== null)
        sh.getRange('F3').setValue(((low + high) / 2).toFixed(1));
    }

    if (boardData) {
      sh.getRange('F10').setValue(Number(boardData.iob).toFixed(2));
      sh.getRange('F11').setValue(Number(boardData.cob).toFixed(1));
    }

    if (bgResult) {
      sh.getRange('F6').setValue(bgResult.mmol.toFixed(1));
      var minutesAgo = Math.round(bgResult.msAgo / 60000);
      sh.getRange('G6').setValue(minutesAgo < 1 ? 'now' : minutesAgo + ' mins ago');
    }

    // Always show current time in F17 (freezes when F16 is filled)
    // F17 = current time + BG value (mmol) as minutes
    var bgMinutes = bgResult ? Math.round(bgResult.mmol) : 0;
    sh.getRange('F17').setValue(new Date(now.getTime() + (bgMinutes * 60000)));

  });
}

function clearAgeLabelsOnly_() {
  var ss = SpreadsheetApp.getActive();
  MEAL_SHEET_NAMES.forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (s) s.getRange('G6').setValue('');
  });
}

function ns_fetchBoardData_() {
  var url = NS_BASE_URL + '/api/v1/devicestatus.json?count=1' + (NS_TOKEN ? '&token=' + NS_TOKEN : '');
  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      var json = JSON.parse(res.getContentText());
      if (json && json.length > 0 && json[0].loop)
        return { iob: json[0].loop.iob.iob || 0, cob: json[0].loop.cob.cob || 0 };
    }
  } catch(e) { Logger.log('Board Fetch Error: ' + e); }
  return { iob: 0, cob: 0 };
}

function ns_fetchProfileData_() {
  var url = NS_BASE_URL + '/api/v1/profile.json' + (NS_TOKEN ? '?token=' + NS_TOKEN : '');
  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      var json = JSON.parse(res.getContentText());
      return json && json.length > 0 ? json[0] : null;
    }
  } catch(e) { Logger.log('Profile Fetch Error: ' + e); }
  return null;
}

function ns_fetchLatestBG_() {
  var url = NS_BASE_URL + '/api/v1/entries/sgv.json?count=1' + (NS_TOKEN ? '&token=' + NS_TOKEN : '');
  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      var arr = JSON.parse(res.getContentText());
      if (arr && arr.length) {
        var entry = arr[0];
        var mgdl  = entry.sgv || entry.mgdl || entry.sg;
        return { mmol: mgdl / 18, msAgo: Date.now() - new Date(entry.date || entry.timestamp).getTime() };
      }
    }
  } catch(e) { Logger.log('BG Fetch Error: ' + e); }
  return null;
}

function ns_getValueForCurrentTime_(scheduleArray) {
  if (!scheduleArray || scheduleArray.length === 0) return null;
  if (typeof scheduleArray === 'number') return scheduleArray;
  var now            = new Date();
  var currentMinutes = (now.getHours() * 60) + now.getMinutes();
  var activeValue    = scheduleArray[0].value;
  for (var i = 0; i < scheduleArray.length; i++) {
    var parts       = scheduleArray[i].time.split(':');
    var itemMinutes = (parseInt(parts[0], 10) * 60) + parseInt(parts[1], 10);
    if (currentMinutes >= itemMinutes) activeValue = scheduleArray[i].value;
    else break;
  }
  return activeValue;
}


/* ======================================================
   6.  POST-BOLUS TRACKING
   ====================================================== */

function onEdit_PostBolusTracking(e) {
  if (!e) return;
  var rng  = e.range;
  var sh   = rng.getSheet();
  if (MEAL_SHEET_NAMES.indexOf(sh.getName()) === -1) return;
  if (rng.getRow() !== 16 || rng.getColumn() !== 6) return;

  var val = rng.getDisplayValue();
  if (!val) { sh.getRange('I5:M10').clearContent(); return; }

  var tz        = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var bolusDate = pb_buildDateFromTimeString_(val, tz);
  var now       = new Date();

  if (bolusDate && bolusDate.getTime() > now.getTime()) {
    sh.getRange('I5:M10').clearContent();
    return;
  }

  // Freeze bolus-time snapshot — replace live values with static values
  var cellsToFreeze = ['F2', 'F3', 'F4', 'F6', 'F10', 'F11', 'F13', 'F17'];
  cellsToFreeze.forEach(function(cell) {
    var v = sh.getRange(cell).getValue();
    if (v !== '' && v !== null) sh.getRange(cell).setValue(v);
  });

  pb_ensureTimer_();
  pb_tick_();
}

function pb_tick_() {
  var ss     = SpreadsheetApp.getActive();
  var tz     = ss.getSpreadsheetTimeZone();
  var latest = pb_fetchLatestTwo_();
  if (!latest) return;
  var now = new Date();

  MEAL_SHEET_NAMES.forEach(function(sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    var f16Text = (sh.getRange('F16').getDisplayValue() || '').trim();
    if (!f16Text) return;
    var bolusDate = pb_buildDateFromTimeString_(f16Text, tz);
    if (!bolusDate) return;

    for (var i = 0; i < PB_OFFSETS_MIN.length; i++) {
      var row      = PB_START_ROW + i;
      var dueTime  = new Date(bolusDate.getTime() + PB_OFFSETS_MIN[i] * 60000);
      var timeCell = sh.getRange(row, 9);
      if (timeCell.getDisplayValue() === '' && now.getTime() >= dueTime.getTime()) {
        sh.getRange(row, 9).setValue(Utilities.formatDate(now, tz, 'h:mm a'));
        sh.getRange(row, 10).setValue(PB_OFFSETS_MIN[i]);
        sh.getRange(row, 11).setValue(latest.mmol);
        sh.getRange(row, 12).setValue(pb_arrow_(latest.direction));
        sh.getRange(row, 13).setValue(pb_signed1dp_(latest.delta));
      }
    }
  });
}

function pb_ensureTimer_() {
  var exists = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'pb_tick_';
  });
  if (!exists) ScriptApp.newTrigger('pb_tick_').timeBased().everyMinutes(1).create();
}

function pb_fetchLatestTwo_() {
  var url = NS_BASE_URL + '/api/v1/entries/sgv.json?count=2' + (NS_TOKEN ? '&token=' + NS_TOKEN : '');
  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    var arr = JSON.parse(res.getContentText());
    if (!arr || !arr.length) return null;
    var a0   = arr[0];
    var a1   = arr.length > 1 ? arr[1] : null;
    var mg0  = a0.sgv || a0.mgdl || a0.sg;
    var mg1  = a1 ? (a1.sgv || a1.mgdl || a1.sg) : null;
    var mm0  = Math.round((mg0 / 18.018) * 10) / 10;
    var dmm  = mg1 ? (mg0 - mg1) / 18.018 : 0;
    return { mmol: mm0, delta: Math.round(dmm * 10) / 10, direction: a0.direction || '' };
  } catch(e) { return null; }
}

function pb_arrow_(direction) {
  var arrows = {
    flat: '→', fortyfiveup: '↗', singleup: '↑', doubleup: '↑↑',
    fortyfivedown: '↘', singledown: '↓', doubledown: '↓↓'
  };
  return arrows[String(direction || '').toLowerCase()] || '';
}

function pb_signed1dp_(x) {
  return isFinite(x) ? (x >= 0 ? '+' : '') + Number(x).toFixed(1) : '';
}

function pb_buildDateFromTimeString_(timeText, tz) {
  try {
    var parsed = new Date(Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd') + ' ' + timeText);
    return isNaN(parsed.getTime()) ? null : parsed;
  } catch(e) { return null; }
}


/* ======================================================
   7.  DAILY EXPORT
   ====================================================== */

function DailyExport_installDaily10pm() {
  DailyExport_removeTriggers();
  ScriptApp.newTrigger('DailyExport_exportNow').timeBased().atHour(23).everyDays(1).create();
}

function DailyExport_removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'DailyExport_exportNow') ScriptApp.deleteTrigger(t);
  });
}

function DailyExport_exportNow() {
  var ss      = SpreadsheetApp.getActive();
  var tz      = ss.getSpreadsheetTimeZone() || 'Etc/UTC';
  var docName = Utilities.formatDate(new Date(), tz, 'MMMM d, yyyy');
  var folder  = DriveApp.getFolderById(EXPORT_FOLDER_ID);

  var existing = folder.getFilesByName(docName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  var sections  = [];
  var indexData = [];

  MEAL_SHEET_NAMES.forEach(function(mealName) {
    var sh = ss.getSheetByName(mealName);
    if (!sh) return;
    var mealHasFood = false;

    for (var row = 2; row <= 86; row += 6) {
      var foodTitle = String(sh.getRange('C' + row).getValue()).trim();
      if (foodTitle) {
        mealHasFood = true;
        indexData.push([
          docName, mealName, foodTitle,
          sh.getRange('C' + (row + 1)).getValue(),  // Carb Factor
          sh.getRange('C' + (row + 3)).getValue(),  // Weight Given
          sh.getRange('C' + (row + 4)).getValue()   // Net Carbs
        ]);
      }
    }

    if (mealHasFood) {
      sections.push({ title: mealName, text: DailyExport_buildMealNote_(sh, tz) });
    }
  });

  if (!sections.length) return;

  var doc  = DocumentApp.create(docName);
  var body = doc.getBody();
  body.clear();
  body.appendParagraph(docName).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  sections.forEach(function(s) {
    body.appendParagraph(s.title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(s.text).setFontFamily('Arial').setFontSize(11);
    body.appendParagraph('');
  });
  doc.saveAndClose();

  var file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch(e) {}

  var docUrl = doc.getUrl();

  if (indexData.length > 0) {
    var indexSheet = ss.getSheetByName(INDEX_SHEET_NAME);
    if (indexSheet) {
      indexData.forEach(function(r) { r.push(docUrl); });
      indexSheet.insertRowsBefore(2, indexData.length);
      indexSheet.getRange(2, 1, indexData.length, indexData[0].length).setValues(indexData);
    }
  }

  MEAL_SHEET_NAMES.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh) ClearMealChartForSheet_(sh);
  });
}

function DailyExport_buildMealNote_(sh, tz) {
  var cv     = function(a1) { return String(sh.getRange(a1).getDisplayValue() || '').trim(); };
  var fmtNum = function(v, d) {
    var n = parseFloat(v);
    return isNaN(n) ? v : n.toFixed(d).replace(/\.0+$/, '').replace(/(\.\d*[1-9])0+$/, '$1');
  };
  var pad = function(str, w) {
    var s = String(str || '');
    return s.length > w ? s.substring(0, w) : s + Array(w - s.length + 1).join(' ');
  };
  var lines = [];

  for (var start = 2; start <= 86; start += 6) {
    var foodName = cv('C' + start);
    if (!foodName) continue;
    lines.push('Food:\t' + foodName + (cv('D' + start) ? ' (' + cv('D' + start) + ')' : ''));
    lines.push('Carb factor:\t'     + cv('C' + (start + 1)));
    lines.push('Absorption Rate:\t' + cv('C' + (start + 2)));
    lines.push('Weight given:\t'    + fmtNum(cv('C' + (start + 3)), 2));
    lines.push('Net carbs:\t'       + fmtNum(cv('C' + (start + 4)), 2));
    lines.push('');

    if (foodName.indexOf('Custom') !== -1 || start === 2) {
      if (cv('P5') !== '') {
        lines.push('--- Meal Builder Breakdown ---');
        for (var rMB = 5; rMB <= 24; rMB++) {
          var ingName = cv('P' + rMB);
          if (!ingName) continue;
          lines.push(pad(ingName, 20) + ' CF: ' + cv('Q' + rMB) +
                     ' | Abs: ' + cv('R' + rMB) + ' | Wt: ' + cv('S' + rMB) +
                     ' | Carb: ' + cv('T' + rMB));
        }
        lines.push('Recipe Net Carbs:\t' + cv('W4'));
        lines.push('Recipe Total Wt:\t'  + cv('W5'));
        lines.push('Recipe Final CF:\t'  + cv('W6'));
        lines.push('Recipe Avg Abs: \t'  + cv('W7'));
        lines.push('------------------------------');
        lines.push('');
      }
    }
  }

  lines.push('Carb Ratio:\t' + cv('F2') + ' | Target: ' + cv('F3') + ' | ISF: ' + cv('F4'));
  lines.push('Current BG:\t' + cv('F6') + ' | Total Net Carbs: ' + cv('F7'));
  lines.push('IOB:\t'        + cv('F10') + ' | COB: ' + cv('F11'));
  lines.push('Total Bolus:\t' + cv('F13') + ' U');
  lines.push('');

  if (cv('F16')) lines.push('Bolus Time: ' + cv('F16'));
  if (cv('F18')) lines.push('Eat Time:   ' + cv('F18'));
  lines.push('');

  var bgData = sh.getRange('I5:M10').getValues();
  if (bgData.some(function(r) { return r[0] !== '' || r[2] !== ''; })) {
    lines.push('----- Post Meal BG Tracking -----');
    lines.push(pad('Time', 10) + pad('Mins', 8) + pad('BG', 8) + pad('Trend', 10) + pad('Delta', 8));
    bgData.forEach(function(row) {
      if (row[0] || row[2]) {
        var tStr = (row[0] instanceof Date)
          ? Utilities.formatDate(row[0], tz, 'h:mm a') : String(row[0]);
        lines.push(pad(tStr, 10) + pad(row[1], 8) + pad(row[2], 8) +
                   pad(row[3], 10) + pad(row[4], 8));
      }
    });
    lines.push('');
  }

  var notes = [cv('E23'), cv('E24'), cv('E25')].filter(Boolean);
  if (notes.length) {
    lines.push('-- Notes --');
    notes.forEach(function(n) { lines.push(n); });
  }

  return lines.join('\n');
}


/* ======================================================
   8.  FOOD MANAGER  (Add to Food Chart sidebar backend)
   ====================================================== */

function showFoodManagerSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createTemplateFromFile('Add_Food_to_Chart')
      .evaluate().setTitle('Add Entry to Food Chart').setWidth(300)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  );
}

function getFoodLibrary() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Food Chart');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  data.shift();
  return data.filter(function(r) { return String(r[0] || '').trim() !== ''; })
    .map(function(r) {
      return { name: String(r[0] || ''), factor: typeof r[1] === 'number' ? r[1] : 0, absorption: r[2] || 3.0 };
    });
}

function addFoodToChart(foodObj) {
  if (!foodObj || typeof foodObj !== 'object')
    throw new Error('No food data received. Please fill in the form and try again.');
  if (typeof foodObj.factor !== 'number' || isNaN(foodObj.factor))
    throw new Error('Invalid carb factor — enter a valid portion and carb value before submitting.');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Food Chart');
  if (!sheet) throw new Error('Food Chart sheet not found!');
  var truncatedFactor = Math.floor(foodObj.factor * 100) / 100;
  var data = sheet.getRange('A1:A' + sheet.getLastRow()).getValues();
  var nextRow = 1;
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== '') { nextRow = i + 2; break; }
  }
  sheet.getRange(nextRow, 1).setValue(foodObj.name);
  sheet.getRange(nextRow, 2).setValue(truncatedFactor);
  sheet.getRange(nextRow, 3).setValue(foodObj.absorption);
  return foodObj.name + ' successfully added';
}

function estimateAbsorption(meta) {
  var netCarbs  = Number(meta.netCarbs  || 0);
  var fat       = Number(meta.fat       || 0);
  var protein   = Number(meta.protein   || 0);
  var method    = String(meta.method    || '').toLowerCase();
  var texture   = String(meta.texture   || '').toLowerCase();
  var mixedMeal = !!meta.mixedMeal;
  var t = 3.0;
  if      (netCarbs >= 40) t += 0.5;
  else if (netCarbs >= 25) t += 0.25;
  if      (fat >= 15) t += 1.0;
  else if (fat >= 10) t += 0.75;
  else if (fat >= 7)  t += 0.5;
  else if (fat >= 4)  t += 0.25;
  if      (protein >= 25) t += 0.5;
  else if (protein >= 15) t += 0.25;
  switch (method) {
    case 'fried':                                           t += 0.5;  break;
    case 'baked': case 'toasted': case 'air_fryer':
    case 'pan_fried_oil':                                   t += 0.25; break;
    case 'boiled_steamed':                                  t -= 0.25; break;
  }
  switch (texture) {
    case 'liquid': case 'blended': t -= 0.5;  break;
    case 'mashed_soft':            t -= 0.25; break;
  }
  if (mixedMeal) t += 0.25;
  return Math.round(Math.max(2.0, Math.min(6.0, t)) * 2) / 2;
}

function getEstimatedAbsorption(meta) { return estimateAbsorption(meta); }


/* ======================================================
   9.  FOOD SEARCH
   ====================================================== */

function FoodSearch_showSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createTemplateFromFile('FoodSearch_Sidebar')
      .evaluate().setTitle('Food Search').setWidth(300)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  );
}

function FoodSearch_getIndexCounts() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(INDEX_SHEET_NAME);
  return { count: sheet ? Math.max(0, sheet.getLastRow() - 1) : 0 };
}

function FoodSearch_query(q) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(INDEX_SHEET_NAME);
  if (!sheet) return { rows: [] };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [] };

  var data  = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var lower = q.toLowerCase().trim();
  var tz    = Session.getScriptTimeZone();
  var rows  = [];

  for (var i = 0; i < data.length; i++) {
    var food = String(data[i][2] || '').replace(/\s+/g, ' ').trim();
    if (food.toLowerCase().indexOf(lower) === -1) continue;
    var dateObj = data[i][0];
    if (!(dateObj instanceof Date)) dateObj = new Date(dateObj);
    if (isNaN(dateObj.getTime())) continue;
    rows.push({
      timestamp: dateObj.getTime(),
      date:      Utilities.formatDate(dateObj, tz, 'EEEE, MMMM d'),
      meal:      String(data[i][1] || ''),
      food:      food,
      cf:        data[i][3] !== '' ? data[i][3] : '—',
      wt:        data[i][4] !== '' ? data[i][4] : '—',
      carbs:     data[i][5] !== '' ? data[i][5] : '—',
      docUrl:    String(data[i][6] || '')
    });
  }

  rows.sort(function(a, b) { return b.timestamp - a.timestamp; });
  return { rows: rows.slice(0, 50) };
}

/**
 * Rebuilds the Food Search Index by scanning all Google Docs
 * in ROOT_FOLDER_ID. Parses meal names from document content
 * rather than folder name.
 */
function FoodSearch_rebuildIndex() {
  var sheet = FoodSearch_ensureIndexSheet_();
  sheet.clear();
  sheet.appendRow(['Date', 'Meal', 'Food', 'Carb Factor', 'Weight (g)', 'Net Carbs', 'Doc URL']);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#f3f3f3');

  var folder = DriveApp.getFolderById(ROOT_FOLDER_ID);
  FoodSearch_walkFolder_(folder, sheet);
  return { count: Math.max(0, sheet.getLastRow() - 1) };
}

function FoodSearch_walkFolder_(folder, sheet) {
  var files     = folder.getFilesByType(MimeType.GOOGLE_DOCS);
  var batchData = [];

  while (files.hasNext()) {
    var file = files.next();
    try {
      var text  = DocumentApp.openById(file.getId()).getBody().getText();
      var clean = text.replace(/[\s\xa0]+/g, ' ').trim();

      // Split doc into meal sections by known meal header names
      var mealPattern = /\b(Breakfast|Morning Snack|Lunch|Afternoon Snack|Dinner|Evening Snack)\b/g;
      var mealMatches = [];
      var m;
      while ((m = mealPattern.exec(clean)) !== null) {
        mealMatches.push({ name: m[1], index: m.index });
      }

      for (var mi = 0; mi < mealMatches.length; mi++) {
        var mealName  = mealMatches[mi].name;
        var start     = mealMatches[mi].index;
        var end       = mi + 1 < mealMatches.length ? mealMatches[mi + 1].index : clean.length;
        var mealChunk = clean.substring(start, end);

        // Support both legacy ("Total weight given") and current ("Weight given") formats
        var hasTotalLabels = /Total\s*weight\s*given:/i.test(mealChunk);
        var foodPattern = hasTotalLabels
          ? /Food:\s*(.*?)\s*Carb\s*factor:\s*([\d.]+).*?Total\s*weight\s*given:\s*([\d.]+)\s*Total\s*net\s*carbs:\s*([\d.]+)/gi
          : /Food:\s*(.*?)\s*Carb\s*factor:\s*([\d.]+).*?Weight\s*given:\s*([\d.]+)\s*Net\s*carbs:\s*([\d.]+)/gi;

        var match;
        while ((match = foodPattern.exec(mealChunk)) !== null) {
          var foodName = match[1].replace(/\(.*?\)/g, '').trim();
          if (foodName && foodName.toLowerCase() !== 'custom (1:1)') {
            batchData.push([
              file.getDateCreated(), mealName, foodName,
              match[2], match[3], match[4], file.getUrl()
            ]);
          }
        }
      }
    } catch(e) {
      Logger.log('FoodSearch error on file: ' + file.getName() + ' | ' + e.toString());
    }
  }

  if (batchData.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, batchData.length, 7).setValues(batchData);
  }

  var subs = folder.getFolders();
  while (subs.hasNext()) FoodSearch_walkFolder_(subs.next(), sheet);
}

function FoodSearch_ensureIndexSheet_() {
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(INDEX_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(INDEX_SHEET_NAME);
  if (!sheet.getRange('A1').getValue()) {
    sheet.getRange('A1:G1').setValues([['Date', 'Meal', 'Food', 'Carb Factor', 'Weight (g)', 'Net Carbs', 'Doc URL']]);
    sheet.getRange('A1:G1').setFontWeight('bold');
  }
  return sheet;
}

function FoodSearch_appendToIndex_(rows) {
  if (!rows || !rows.length) return;
  var sheet   = FoodSearch_ensureIndexSheet_();
  var trimmed = rows.map(function(r) { return r.slice(0, 7); });
  sheet.insertRowsBefore(2, trimmed.length);
  sheet.getRange(2, 1, trimmed.length, 7).setValues(trimmed);
}


/* ======================================================
   10.  LEGACY LOG IMPORT
   ====================================================== */

function LegacyImport_indexLegacyDocs() {
  if (!LEGACY_LOGS_FOLDER_ID) throw new Error('Set LEGACY_LOGS_FOLDER_ID at the top of the script.');
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'America/Toronto';
  FoodSearch_ensureIndexSheet_();
  var allRows = [];
  LegacyImport_processFolder_(DriveApp.getFolderById(LEGACY_LOGS_FOLDER_ID), allRows, tz);
  if (!allRows.length) { Logger.log('No legacy entries found.'); return; }
  FoodSearch_appendToIndex_(allRows);
  Logger.log('LegacyImport: added ' + allRows.length + ' rows.');
}

function LegacyImport_processFolder_(folder, allRows, tz) {
  var files = folder.getFiles();
  while (files.hasNext()) LegacyImport_processDoc_(files.next(), allRows, tz);
  var subs = folder.getFolders();
  while (subs.hasNext()) LegacyImport_processFolder_(subs.next(), allRows, tz);
}

function LegacyImport_processDoc_(file, allRows, tz) {
  var mime = file.getMimeType();
  var text = '';
  if      (mime === MimeType.GOOGLE_DOCS) text = DocumentApp.openById(file.getId()).getBody().getText();
  else if (mime === MimeType.PLAIN_TEXT)  text = file.getBlob().getDataAsString();
  else return;

  var docName     = file.getName().replace(/\.txt$/i, '');
  var docUrl      = file.getUrl();
  var lines       = String(text).split('\n');
  var currentMeal = '', food = '', cf = '', wt = '', carbs = '';

  function flush() {
    if (!food && !cf && !wt && !carbs) return;
    // 7-column row to match index structure: date, meal, food, cf, wt, carbs, docUrl
    allRows.push([docName, currentMeal, food || '', cf || '', wt || '', carbs || '', docUrl]);
    food = ''; cf = ''; wt = ''; carbs = '';
  }

  for (var i = 0; i < lines.length; i++) {
    var line = String(lines[i] || '').trim();
    if (!line) continue;
    if (/morning snack|afternoon snack|evening snack|breakfast|lunch|dinner/i.test(line)) {
      flush(); currentMeal = line; continue;
    }
    var m;
    if      ((m = /Food:\s*(.*)/i.exec(line)))               { flush(); food  = m[1].trim(); }
    else if ((m = /Carb factor:\s*(.*)/i.exec(line)))         { cf    = m[1].trim(); }
    else if ((m = /Total weight given:\s*(.*)/i.exec(line)))  { wt    = m[1].trim(); }
    else if ((m = /Total net carbs:\s*(.*)/i.exec(line)))     { carbs = m[1].trim(); }
  }
  flush();
}


/* ======================================================
   11.  MEAL BUILDER
   ====================================================== */

function MealBuilder_showSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createTemplateFromFile('MealBuilder_Sidebar')
      .evaluate().setTitle('Meal Builder')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  );
}

function MealBuilder_getFoodChart() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Food Chart');
  if (!sh) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(function(row) { return String(row[0]).trim() !== ''; })
    .map(function(row) {
      return { name: String(row[0]).trim(), cf: parseFloat(row[1]) || 0, abs: parseFloat(row[2]) || 3.0 };
    });
}

function MealBuilder_addToMeal(payload) {
  var sheet = SpreadsheetApp.getActive().getActiveSheet();
  if (MEAL_SHEET_NAMES.indexOf(sheet.getName()) === -1)
    throw new Error('Please switch to a Meal Sheet first.');

  var targetRow = null;
  for (var r = 2; r <= 86; r += 6) {
    if (!sheet.getRange(r, 3).getValue()) { targetRow = r; break; }
  }
  if (!targetRow) throw new Error('No empty slots left on this sheet.');

  sheet.getRange(targetRow,     3).setValue(payload.title);
  sheet.getRange(targetRow + 1, 3).setValue(payload.totalCarbFactor);
  sheet.getRange(targetRow + 2, 3).setValue(payload.mealAbs);
  sheet.getRange(targetRow + 3, 3).setValue(payload.weightGiven);
  sheet.getRange(targetRow + 4, 3).setValue(payload.totalCarbs);

  var mbRows = (payload.ingredients || []).map(function(ing) {
    return [ing.food, ing.cf, ing.abs, ing.weight, ing.carbs];
  });
  sheet.getRange('P5:T24').clearContent();
  if (mbRows.length) sheet.getRange(5, 16, mbRows.length, 5).setValues(mbRows);

  sheet.getRange('W4').setValue(payload.recipeNetCarbs);
  sheet.getRange('W5').setValue(payload.recipeWeight);
  sheet.getRange('W6').setValue(Math.floor(payload.totalCarbFactor * 100) / 100);
  sheet.getRange('W7').setValue(payload.mealAbs);
  sheet.getRange('W9').setValue(payload.weightGiven);
  sheet.getRange('W10').setValue(payload.totalCarbs);
  return { success: true };
}


/* ======================================================
   12.  COPY NOTE
   ====================================================== */

function runCopyNoteQuick() {
  var sh   = SpreadsheetApp.getActiveSheet();
  var name = sh.getName();
  var note;

  if (MEAL_SHEET_NAMES.indexOf(name) !== -1) {
    note = copyNote_buildMealNote_();
  } else if (name === 'Meal Carb Factor') {
    note = copyNote_buildCarbFactor_();
  } else {
    SpreadsheetApp.getActive().toast(
      'Copy Note works only on meal sheets or the Meal Carb Factor sheet.', 'Action Blocked', -1);
    return;
  }
  copyNote_viaDialog_(name === 'Meal Carb Factor' ? 'Meal Carb Factor' : name + ' Note', note);
}

function copyNote_buildMealNote_() {
  var sh          = SpreadsheetApp.getActiveSheet();
  var lines       = [];
  var totalWeight = 0;
  var totalCarbs  = 0;

  var getVal = function(a1) { return sh.getRange(a1).getValue(); };
  var joinCD = function(r) {
    var c = String(sh.getRange('C' + r).getValue() || '');
    var d = String(sh.getRange('D' + r).getValue() || '');
    return (c && d) ? c + ' | ' + d : (c || d);
  };
  var toNum = function(v) {
    var n = parseFloat(String(v).replace(/[^0-9.-]/g, ''));
    return isFinite(n) ? n : NaN;
  };

  for (var start = 3; start <= 58; start += 5) {
    var foodRow  = start - 1;
    var foodDesc = getVal('B' + foodRow);
    if (!foodDesc) continue;
    lines.push(foodDesc + '\t' + joinCD(foodRow));

    var mealRows = [
      { fallback: 'Carb factor:',        r: start     },
      { fallback: 'Total weight given:', r: start + 1 },
      { fallback: 'Total net carbs:',    r: start + 2 }
    ];
    var labels = mealRows.map(function(x) { return getVal('B' + x.r) || x.fallback; });
    var vals   = mealRows.map(function(x) { return joinCD(x.r); });
    if (!vals.some(function(s) { return s.length; })) continue;
    mealRows.forEach(function(x, i) { lines.push(labels[i] + '\t' + vals[i]); });
    lines.push('');

    var w = toNum(vals[1]); var c = toNum(vals[2]);
    if (!isNaN(w)) totalWeight += w;
    if (!isNaN(c)) totalCarbs  += c;
  }

  var weightGivenStr   = getVal('F6');
  var totalCarbsOutStr = getVal('F7');
  var totalFactor      = totalWeight > 0 ? totalCarbs / totalWeight : NaN;
  if (!totalCarbsOutStr && weightGivenStr && !isNaN(totalFactor)) {
    var wg = toNum(weightGivenStr);
    if (!isNaN(wg)) totalCarbsOutStr = (Math.round(wg * totalFactor * 100) / 100).toString();
  }

  lines.push('Total net carbs:\t'    + (totalCarbsOutStr || ''));
  lines.push('Total weight given:\t' + (weightGivenStr   || ''));
  lines.push('');
  lines.push('IOB:\t'         + (getVal('F10') || ''));
  lines.push('COB:\t'         + (getVal('F11') || ''));
  lines.push('Total Bolus:\t' + (getVal('F13') || ''));
  return lines.join('\n');
}

function copyNote_buildCarbFactor_() {
  var sh    = SpreadsheetApp.getActiveSheet();
  var lines = [];
  var getVal = function(a1) { return sh.getRange(a1).getValue(); };
  var joinCD = function(r) {
    var c = String(sh.getRange('C' + r).getValue() || '');
    var d = String(sh.getRange('D' + r).getValue() || '');
    return (c && d) ? c + ' | ' + d : (c || d);
  };

  for (var start = 2; start <= 57; start += 5) {
    var foodDesc = getVal('B' + start);
    if (!foodDesc) continue;
    lines.push(foodDesc);
    var hasContent = false;
    [{ r: start }, { r: start + 1 }, { r: start + 2 }].forEach(function(x, idx) {
      var fallbacks = ['Carb factor:', 'Total weight given:', 'Total net carbs:'];
      var label = getVal('B' + x.r) || fallbacks[idx];
      var v     = joinCD(x.r);
      if (v) { lines.push(label + '\t' + v); hasContent = true; }
    });
    var noteVal = getVal('C' + (start + 3));
    if (noteVal) { lines.push((getVal('B' + (start + 3)) || 'Note:') + '\t' + noteVal); hasContent = true; }
    if (hasContent) lines.push('');
    else lines.pop();
  }
  return lines.join('\n');
}

function copyNote_viaDialog_(title, note) {
  var payload = JSON.stringify(String(note || ''));
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">' +
    '<style>body{margin:0;padding:10px;font:13px system-ui,Arial}' +
    '#wrap{display:flex;align-items:center;gap:8px}#msg{color:#0a7f00}' +
    '#btn{padding:6px 10px;background:#f8f8f8;border:1px solid #ccc;border-radius:4px;cursor:pointer}' +
    '</style></head><body>' +
    '<div id="wrap"><span id="msg">...</span>' +
    '<button id="btn" style="display:none">Copy Manually</button></div>' +
    '<textarea id="txt" style="position:absolute;left:-9999px;top:-9999px"></textarea>' +
    '<script>' +
    'var text=' + payload + ',' +
    'txt=document.getElementById("txt"),' +
    'msg=document.getElementById("msg"),' +
    'btn=document.getElementById("btn");' +
    'txt.value=text;' +
    'function doCopy(){' +
    'try{if(navigator.clipboard&&navigator.clipboard.writeText){' +
    'navigator.clipboard.writeText(text).then(function(){msg.textContent="Copied!";setTimeout(function(){google.script.host.close();},120);});' +
    '}else{txt.select();document.execCommand("copy");msg.textContent="Copied!";setTimeout(function(){google.script.host.close();},120);}' +
    'return true;}catch(e){return false;}}' +
    'if(!doCopy()){msg.textContent="Click Copy Manually";msg.style.color="#b00";btn.style.display="";' +
    'btn.addEventListener("click",function(){try{txt.select();document.execCommand("copy");' +
    'msg.textContent="Copied!";msg.style.color="#0a7f00";btn.disabled=true;' +
    'setTimeout(function(){google.script.host.close();},120);}' +
    'catch(e){msg.textContent="Copy failed.";msg.style.color="#b00";}});}' +
    '</script></body></html>'
  );
  SpreadsheetApp.getUi().showModalDialog(html, title);
}


/* ======================================================
   13.  CLEAR / RESTORE HELPERS
   ====================================================== */

/**
 * Clears the active meal sheet and restores all formulas.
 * Assigned to the Clear button image on each meal sheet.
 */
function ClearMealChart() {
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  if (MEAL_SHEET_NAMES.indexOf(sheet.getName()) === -1) {
    ss.toast('Please run this on a Meal Sheet.');
    return;
  }

  var rangesToClear = [];
  for (var i = 0; i < 15; i++) {
    var startRow = (i * 6) + 1;
    rangesToClear.push('C' + (startRow + 1)); // Food name
    rangesToClear.push('C' + (startRow + 4)); // Weight given
    rangesToClear.push('C' + (startRow + 5)); // Total carbs
  }
  rangesToClear.push(
    'D1:D90', 'E23:E25',
    'F7', 'F8', 'F13', 'F14',
    'F16', 'F17', 'F18',
    'I5:M10', 'P5:T24', 'W4:W10'
  );
  sheet.getRangeList(rangesToClear).clearContent();

  for (var j = 0; j < 15; j++) {
    var base = (j * 6) + 2;
    sheet.getRange(base + 1, 3).setFormula(
      '=IF(C' + base + '="", "", IFERROR(VLOOKUP(C' + base + ', \'Food Chart\'!A:B, 2, FALSE), ""))');
    sheet.getRange(base + 2, 3).setFormula(
      '=IF(C' + base + '="", "", IFERROR(VLOOKUP(C' + base + ', \'Food Chart\'!A:C, 3, FALSE), ""))');
    sheet.getRange(base + 2, 4).setFormula(
      '=IF(C' + base + '="", "", IFERROR(VLOOKUP(C' + base + ', \'Food Chart\'!A:D, 4, FALSE), ""))');
  }

  sheet.getRange('F19').setFormula('=IF(OR(F18="",F16=""),"",ROUND((F18-F16)*24*60,0)&" minutes")');
  restoreF13Formula_(sheet);

  sheet.getRange('A1').activate();
  ss.toast('Sheet cleared');
}

/**
 * Used by DailyExport after nightly export.
 * Same logic as ClearMealChart but operates on a passed sheet reference.
 */
function ClearMealChartForSheet_(sh) {
  var rangesToClear = [];
  for (var i = 0; i < 15; i++) {
    var startRow = (i * 6) + 1;
    rangesToClear.push('C' + (startRow + 1));
    rangesToClear.push('C' + (startRow + 4));
    rangesToClear.push('C' + (startRow + 5));
  }
  rangesToClear.push(
    'D1:D90', 'E23:E25',
    'F7', 'F8', 'F13', 'F14',
    'F16', 'F17', 'F18',
    'I5:M10', 'P5:T24', 'W4:W10'
  );
  sh.getRangeList(rangesToClear).clearContent();

  for (var j = 0; j < 15; j++) {
    var base = (j * 6) + 2;
    sh.getRange(base + 1, 3).setFormula(
      '=IF(C' + base + '="", "", IFERROR(VLOOKUP(C' + base + ', \'Food Chart\'!A:B, 2, FALSE), ""))');
    sh.getRange(base + 2, 3).setFormula(
      '=IF(C' + base + '="", "", IFERROR(VLOOKUP(C' + base + ', \'Food Chart\'!A:C, 3, FALSE), ""))');
    sh.getRange(base + 2, 4).setFormula(
      '=IF(C' + base + '="", "", IFERROR(VLOOKUP(C' + base + ', \'Food Chart\'!A:D, 4, FALSE), ""))');
  }

  sh.getRange('F19').setFormula('=IF(OR(F18="",F16=""),"",ROUND((F18-F16)*24*60,0)&" minutes")');
  restoreF13Formula_(sh);
}


/* ======================================================
   14.  SHARED HELPERS
   ====================================================== */

function numFromDisplay_(val) {
  var s = String(val === null || val === undefined ? '' : val).replace(/[^0-9.\-]/g, '');
  var n = Number(s);
  return isFinite(n) ? n : 0;
}

function round1_(n)  { return Math.round(n * 10)  / 10;  }
function round05_(n) { return Math.round(n * 2)   / 2;   }
function round2_(n)  { return Math.round(n * 100) / 100; }
