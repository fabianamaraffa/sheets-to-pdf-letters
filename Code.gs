// ============================================================
//  Sheets to PDF Letters  |  Code.gs
//
//  Paste this file into Extensions > Apps Script of your Google Sheet.
//  It needs a second file called Sidebar.html (see README.md).
//
//  What it does: for each row of your sheet it copies a Google Doc
//  template, replaces {{placeholders}} with the values of that row,
//  and saves the result as a PDF in a Drive folder.
// ============================================================


// ---------- SETTINGS (safe to edit) ----------

var SETTINGS = {
  MENU_NAME: 'PDF Generator',              // name of the menu in the Sheet
  SIDEBAR_TITLE: 'PDF Letter Generator',   // title shown at the top of the sidebar
  TEMP_PREFIX: '_temp_pdf_gen_',           // name of the short-lived working copies
  MAX_SECONDS_PER_CALL: 240,               // safety limit: Google stops scripts after about 360 s
  FILE_NAME_MAX_LENGTH: 120                // longest PDF file name, in characters
};


// ---------- MENU ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(SETTINGS.MENU_NAME)
    .addItem('Open configurator', 'openSidebar')
    .addItem('Quick instructions', 'showInstructions')
    .addSeparator()
    .addItem('Inspect template structure (advanced)', 'debugTemplateStructure')
    .addToUi();
}

function openSidebar() {
  // The HTML file must be called "Sidebar" (it is saved as Sidebar.html)
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle(SETTINGS.SIDEBAR_TITLE);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showInstructions() {
  var msg = 'HOW TO USE ' + SETTINGS.SIDEBAR_TITLE.toUpperCase() + '\n\n' +
    '1. TEMPLATE DOC\n' +
    '   Create a Google Doc and write placeholders in double curly braces.\n' +
    '   Example: {{first_name}}, {{country}}, {{date}}\n' +
    '   You only need to be able to open the Doc yourself.\n' +
    '   Use a real Google Doc. If you have a Word file, upload it to Drive,\n' +
    '   open it and choose File > Save as Google Docs.\n\n' +
    '2. GOOGLE SHEET\n' +
    '   The first row holds the column headers.\n' +
    '   Every row after it becomes one PDF.\n\n' +
    '3. SIDEBAR\n' +
    '   - Paste the link of the Doc template\n' +
    '   - Click "Detect placeholders"\n' +
    '   - Paste the link of the Drive folder where the PDFs will go\n' +
    '   - Choose the file name pattern (for example Letter_{{last_name}})\n' +
    '   - Match each placeholder to a column\n' +
    '   - Decide what happens when a cell is empty\n\n' +
    '4. PRE-FLIGHT CHECK\n' +
    '   Always run it before the full batch. It shows what the first selected\n' +
    '   row will look like and warns you about problems.\n\n' +
    '5. GENERATE\n' +
    '   Click "Generate PDFs". Progress is shown in the sidebar and the files\n' +
    '   appear in your Drive folder. You can stop at any time.';
  SpreadsheetApp.getUi().alert('Instructions', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}


// ---------- CALLS FROM THE SIDEBAR ----------

function getSheetData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { headers: [], totalRows: 0, sheetName: sheet.getName() };
  }
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var headers = headerRow.map(function (h, i) {
    return { label: String(h).trim() || 'Column ' + (i + 1), index: i };
  });
  return {
    headers: headers,
    totalRows: lastRow - 1,
    sheetName: sheet.getName()
  };
}

// Returns the sheet the sidebar was showing, falling back to the active one
function getSheet_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (sheetName) {
    var named = ss.getSheetByName(sheetName);
    if (named) return named;
  }
  return ss.getActiveSheet();
}


// ---------- SANITISING VALUES ----------

// Error values that Google Sheets can show in a cell
var SHEET_ERRORS = ['#REF!', '#N/A', '#VALUE!', '#DIV/0!', '#NAME?', '#NUM!', '#NULL!', '#ERROR!'];

// Escape characters that have a special meaning in a regular expression
function escapeForRegex_(str) {
  return String(str).replace(/[\\.+*?^${}()|[\]\/]/g, '\\$&');
}

// Clean a raw cell value. Sheet error values count as empty.
function sanitiseValue(raw) {
  var str = String(raw === null || raw === undefined ? '' : raw).trim();
  for (var i = 0; i < SHEET_ERRORS.length; i++) {
    if (str === SHEET_ERRORS[i]) return '';
  }
  return str;
}

// Final value for a placeholder, given its "if empty" setting.
// fieldConfig = { behaviour: 'empty' | 'default' | 'skip', defaultValue: string }
// Returns { value: string, skip: boolean }
function resolveValue(raw, fieldConfig) {
  var value = sanitiseValue(raw);
  if (value !== '') return { value: value, skip: false };

  var behaviour = (fieldConfig && fieldConfig.behaviour) || 'empty';
  if (behaviour === 'skip') return { value: '', skip: true };
  if (behaviour === 'default') {
    var def = (fieldConfig && fieldConfig.defaultValue) ? String(fieldConfig.defaultValue) : '';
    return { value: def, skip: false };
  }
  return { value: '', skip: false };
}

// Resolve every mapped placeholder for one row.
// Returns { values: {placeholder: text}, skipReason: string }
function resolveRow_(config, rowData) {
  var values = {};
  var skipReason = '';
  for (var placeholder in config.mapping) {
    var fieldConfig = config.fieldConfigs ? config.fieldConfigs[placeholder] : null;
    var resolved = resolveValue(rowData[config.mapping[placeholder]], fieldConfig);
    values[placeholder] = resolved.value;
    if (resolved.skip && !skipReason) {
      skipReason = '{{' + placeholder + '}} is empty (set to skip the row)';
    }
  }
  return { values: values, skipReason: skipReason };
}


// ---------- WALKING THROUGH A DOC ----------

// Calls callback(textElement) for every piece of text inside an element.
// Works for the body, header, footer, tables, lists and so on.
function forEachText_(element, callback) {
  if (element.getType() === DocumentApp.ElementType.TEXT) {
    callback(element.asText());
    return;
  }
  if (typeof element.getNumChildren === 'function') {
    var n = element.getNumChildren();
    for (var i = 0; i < n; i++) {
      forEachText_(element.getChild(i), callback);
    }
  }
}

// Body, header and footer of a Doc (the ones that exist)
function getDocSections_(doc) {
  var sections = [doc.getBody()];
  try { var header = doc.getHeader(); if (header) sections.push(header); } catch (e) {}
  try { var footer = doc.getFooter(); if (footer) sections.push(footer); } catch (e) {}
  return sections;
}

var PLACEHOLDER_SOURCE = '\\{\\{([^}]+)\\}\\}';


// ---------- FINDING PLACEHOLDERS ----------

function getPlaceholdersFromDoc(docUrl) {
  try {
    var doc = DocumentApp.openById(extractIdFromUrl(docUrl));
    var unique = [];
    var seen = {};

    getDocSections_(doc).forEach(function (section) {
      forEachText_(section, function (textEl) {
        var re = new RegExp(PLACEHOLDER_SOURCE, 'g');
        var text = textEl.getText();
        var m;
        while ((m = re.exec(text)) !== null) {
          var key = m[1].trim();
          if (key && !seen[key]) { seen[key] = true; unique.push(key); }
        }
      });
    });

    return { placeholders: unique, error: null };
  } catch (e) {
    return { placeholders: [], error: 'Cannot open the Doc. Check the link and that you can open it: ' + e.message };
  }
}


// ---------- VALIDATION ----------

// Returns the list of placeholders that have no column assigned
function validateMapping(placeholders, mapping) {
  var unmapped = [];
  placeholders.forEach(function (p) {
    if (mapping[p] === undefined || mapping[p] === null || mapping[p] === '') {
      unmapped.push('{{' + p + '}}');
    }
  });
  return unmapped;
}


// ---------- REPLACING PLACEHOLDERS ----------

// Replaces the placeholders found in one text element.
// We insert and delete text ourselves instead of using replaceText(), because
// replaceText() reads characters such as $ and \ in the value as instructions,
// which would corrupt things like "US$ 500" or "C:\Files".
function replacePlaceholdersInText_(textEl, values) {
  var text = textEl.getText();
  if (text.indexOf('{{') === -1) return 0;

  var re = new RegExp(PLACEHOLDER_SOURCE, 'g');
  var found = [];
  var m;
  while ((m = re.exec(text)) !== null) {
    var key = m[1].trim();
    if (Object.prototype.hasOwnProperty.call(values, key)) {
      found.push({ start: m.index, end: m.index + m[0].length - 1, raw: m[0], value: String(values[key]) });
    }
  }

  // Work from the end, so the positions of earlier matches stay valid
  for (var i = found.length - 1; i >= 0; i--) {
    var f = found[i];
    try {
      if (f.value !== '') {
        // Insert right after the placeholder, then delete the placeholder
        if (f.end + 1 >= textEl.getText().length) textEl.appendText(f.value);
        else textEl.insertText(f.end + 1, f.value);
      }
      textEl.deleteText(f.start, f.end);
    } catch (err) {
      if (f.value === '') {
        // Some Docs refuse to delete a whole text run. Removing with an empty value is safe here.
        textEl.replaceText(escapeForRegex_(f.raw), '');
      } else {
        throw err;
      }
    }
  }
  return found.length;
}

// Replaces placeholders in the body, header and footer of a Doc
function replacePlaceholdersInDoc_(doc, values) {
  getDocSections_(doc).forEach(function (section) {
    forEachText_(section, function (textEl) {
      replacePlaceholdersInText_(textEl, values);
    });
  });
}


// ---------- FILE NAMES ----------

// Builds a safe PDF file name (without the .pdf part)
function buildFileName_(pattern, values, rowNumber) {
  var name = (pattern && String(pattern).trim()) || 'Letter_{{row}}';
  name = name.replace(/\{\{([^}]+)\}\}/g, function (whole, rawKey) {
    var key = rawKey.trim();
    if (key === 'row') return String(rowNumber);
    return Object.prototype.hasOwnProperty.call(values, key) ? values[key] : '';
  });
  name = name.replace(/[\\\/:*?"<>|\r\n\t]/g, '_').replace(/\s+/g, ' ').trim();
  name = name.replace(/^\.+|\.+$/g, '').trim();
  if (name.length > SETTINGS.FILE_NAME_MAX_LENGTH) {
    name = name.substring(0, SETTINGS.FILE_NAME_MAX_LENGTH).trim();
  }
  return name || ('Letter_row_' + rowNumber);
}

// Placeholders used in the file name pattern that do not exist in the mapping
function unknownNamePlaceholders_(pattern, mapping) {
  var unknown = [];
  var re = new RegExp(PLACEHOLDER_SOURCE, 'g');
  var m;
  while ((m = re.exec(pattern || '')) !== null) {
    var key = m[1].trim();
    if (key !== 'row' && !Object.prototype.hasOwnProperty.call(mapping, key) && unknown.indexOf(key) === -1) {
      unknown.push(key);
    }
  }
  return unknown;
}


// ---------- PRE-FLIGHT CHECK ----------

// Simulates the substitution on ONE row and reports what would happen.
// Nothing is created in Drive.
function preflightCheck(config) {
  try {
    var sheet = getSheet_(config.sheetName);
    var allData = sheet.getDataRange().getDisplayValues();
    var lastDataRow = allData.length - 1;
    if (lastDataRow < 1) return { success: false, error: 'No data rows found in the sheet.' };

    var rowNumber = Math.min(Math.max(parseInt(config.previewRow, 10) || 1, 1), lastDataRow);
    var rowData = allData[rowNumber];
    var warnings = [];

    var unmapped = validateMapping(config.placeholders || [], config.mapping);
    if (unmapped.length > 0) {
      warnings.push('Unmapped placeholders: ' + unmapped.join(', ') + '. They will stay in the PDF as they are.');
    }

    var unknown = unknownNamePlaceholders_(config.fileNamePattern, config.mapping);
    if (unknown.length > 0) {
      warnings.push('The file name pattern uses placeholders that are not mapped: {{' + unknown.join('}}, {{') + '}}. They will be left out of the name.');
    }

    var resolvedRow = resolveRow_(config, rowData);

    for (var placeholder in config.mapping) {
      var raw = rowData[config.mapping[placeholder]];
      var rawStr = String(raw === null || raw === undefined ? '' : raw).trim();
      var fieldConfig = config.fieldConfigs ? config.fieldConfigs[placeholder] : null;

      if (SHEET_ERRORS.indexOf(rawStr) !== -1) {
        warnings.push('{{' + placeholder + '}} contains a sheet error (' + rawStr + '). It is treated as empty.');
      }
      if (sanitiseValue(raw) === '') {
        var behaviour = (fieldConfig && fieldConfig.behaviour) || 'empty';
        if (behaviour === 'skip') {
          warnings.push('{{' + placeholder + '}} is empty, so this row would be SKIPPED.');
        } else if (behaviour === 'default') {
          warnings.push('{{' + placeholder + '}} is empty, so the default "' + ((fieldConfig && fieldConfig.defaultValue) || '') + '" is used.');
        } else {
          warnings.push('{{' + placeholder + '}} is empty and will be left blank in the PDF.');
        }
      }
    }

    return {
      success: true,
      rowNumber: rowNumber,
      preview: resolvedRow.values,
      fileName: buildFileName_(config.fileNamePattern, resolvedRow.values, rowNumber) + '.pdf',
      warnings: warnings
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// ---------- MAIN GENERATION ----------

// Called by the sidebar several times, a few rows at a time, so that
// progress can be shown and Google's time limit is never reached.
// config.rowFrom and config.rowTo are data row numbers (1 = first row under the headers).
// If time runs short the result contains nextRow: the row to continue from.
function generatePDFs(config) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    return { success: false, error: 'Another generation is still running. Wait for it to finish and try again.' };
  }
  try {
    return generateRows_(config);
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function generateRows_(config) {
  var started = Date.now();
  var budgetMs = SETTINGS.MAX_SECONDS_PER_CALL * 1000;

  var unmapped = validateMapping(config.placeholders || [], config.mapping);
  if (unmapped.length > 0 && !config.forceGenerate) {
    return {
      success: false,
      validationError: true,
      unmapped: unmapped,
      error: 'Some placeholders are not mapped: ' + unmapped.join(', ')
    };
  }

  var sheet = getSheet_(config.sheetName);
  var allData = sheet.getDataRange().getDisplayValues();
  var lastDataRow = allData.length - 1;
  if (lastDataRow < 1) return { success: false, error: 'No data rows found in the sheet.' };

  var templateFile = DriveApp.getFileById(extractIdFromUrl(config.docUrl));
  var folder = DriveApp.getFolderById(extractIdFromUrl(config.folderUrl));

  var fromRow = Math.min(Math.max(parseInt(config.rowFrom, 10) || 1, 1), lastDataRow);
  var toRow = Math.min(Math.max(parseInt(config.rowTo, 10) || lastDataRow, 1), lastDataRow);

  var results = [];
  var successCount = 0;
  var errorCount = 0;
  var skippedCount = 0;
  var nextRow = null;

  for (var i = fromRow; i <= toRow; i++) {
    if (Date.now() - started > budgetMs) { nextRow = i; break; }

    var rowData = allData[i];
    var copyFile = null;
    try {
      if (config.skipIfEmpty && rowData.every(function (cell) { return String(cell).trim() === ''; })) {
        results.push({ row: i, status: 'skipped', name: 'Row ' + (i + 1), reason: 'Empty row' });
        skippedCount++;
        continue;
      }

      var resolvedRow = resolveRow_(config, rowData);
      if (resolvedRow.skipReason) {
        results.push({ row: i, status: 'skipped', name: 'Row ' + (i + 1), reason: resolvedRow.skipReason });
        skippedCount++;
        continue;
      }

      // Work on a copy of the template
      copyFile = templateFile.makeCopy(SETTINGS.TEMP_PREFIX + i, folder);
      var copyDoc = DocumentApp.openById(copyFile.getId());
      replacePlaceholdersInDoc_(copyDoc, resolvedRow.values);
      copyDoc.saveAndClose();

      // Pick a file name that is not taken yet
      var baseName = buildFileName_(config.fileNamePattern, resolvedRow.values, i);
      var finalName = baseName;
      var counter = 2;
      while (folder.getFilesByName(finalName + '.pdf').hasNext()) {
        finalName = baseName + '_' + counter;
        counter++;
      }

      var pdfBlob = copyFile.getAs('application/pdf').setName(finalName + '.pdf');
      folder.createFile(pdfBlob);

      successCount++;
      results.push({ row: i, status: 'ok', name: finalName + '.pdf' });

    } catch (rowErr) {
      errorCount++;
      results.push({ row: i, status: 'error', name: 'Row ' + (i + 1), error: rowErr.message });
    } finally {
      // The working copy is always removed, also when something went wrong
      if (copyFile) {
        try { copyFile.setTrashed(true); } catch (trashErr) {}
      }
    }
  }

  return {
    success: true,
    successCount: successCount,
    errorCount: errorCount,
    skippedCount: skippedCount,
    nextRow: nextRow,
    results: results
  };
}


// ---------- ADVANCED: INSPECT A TEMPLATE ----------

// Menu item for troubleshooting: shows how a Doc is built and where each
// placeholder sits. Useful when a placeholder is not found.
function debugTemplateStructure() {
  var ui = SpreadsheetApp.getUi();
  var answer = ui.prompt('Inspect template structure', 'Paste the link of the Doc template:', ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;

  try {
    var doc = DocumentApp.openById(extractIdFromUrl(answer.getResponseText().trim()));
    var output = '';
    var header = null;
    var footer = null;
    try { header = doc.getHeader(); } catch (e) {}
    try { footer = doc.getFooter(); } catch (e) {}
    var labelled = [
      { label: 'BODY', section: doc.getBody() },
      { label: 'HEADER', section: header },
      { label: 'FOOTER', section: footer }
    ];

    labelled.forEach(function (item) {
      output += '=== ' + item.label + ' ===\n';
      if (!item.section) { output += '(none)\n\n'; return; }
      var count = 0;
      forEachText_(item.section, function (textEl) {
        var text = textEl.getText();
        if (text.trim() === '') return;
        count++;
        if (count <= 40) output += '- ' + text.substring(0, 80).replace(/\n/g, ' ') + '\n';
      });
      if (count > 40) output += '... and ' + (count - 40) + ' more text blocks\n';
      output += '\n';
    });

    var found = getPlaceholdersFromDoc(answer.getResponseText().trim());
    output += 'Placeholders found: ' + (found.placeholders.length ? found.placeholders.join(', ') : 'none');
    ui.alert('Template structure', output, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}


// ---------- UTILITY ----------

// Extracts the id from a Google Docs or Drive link (or accepts a bare id)
function extractIdFromUrl(url) {
  if (!url) throw new Error('Missing link');
  var patterns = [
    /\/d\/([a-zA-Z0-9_-]+)/,
    /\/folders\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/
  ];
  for (var i = 0; i < patterns.length; i++) {
    var match = url.match(patterns[i]);
    if (match) return match[1];
  }
  if (/^[a-zA-Z0-9_-]{20,}$/.test(url.trim())) return url.trim();
  throw new Error('Unrecognised link: ' + url);
}
