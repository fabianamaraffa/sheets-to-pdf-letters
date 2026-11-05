// Tests for Code.gs, run with plain Node (no installs needed):
//
//     node tests/code.test.js
//
// Google's services (Sheets, Docs, Drive) are replaced by small fakes, so
// these tests check OUR logic: placeholder replacement, skipping, file names,
// batching and clean-up. They do not talk to Google.

var fs = require('fs');
var path = require('path');
var vm = require('vm');
var assert = require('assert');

// ---------- fakes for the Google services ----------

function Text(str) { this.str = str; }
Text.prototype.getType = function () { return 'TEXT'; };
Text.prototype.asText = function () { return this; };
Text.prototype.getText = function () { return this.str; };
Text.prototype.insertText = function (offset, s) {
  if (offset < 0 || offset > this.str.length) throw new Error('bad offset ' + offset);
  this.str = this.str.slice(0, offset) + s + this.str.slice(offset);
  return this;
};
Text.prototype.appendText = function (s) { this.str += s; return this; };
Text.prototype.deleteText = function (a, b) {
  if (a < 0 || b >= this.str.length || a > b) throw new Error('bad range ' + a + '-' + b);
  this.str = this.str.slice(0, a) + this.str.slice(b + 1);
  return this;
};
Text.prototype.replaceText = function (pattern, replacement) {
  this.str = this.str.replace(new RegExp(pattern, 'g'), replacement);
  return this;
};
Text.prototype.clone = function () { return new Text(this.str); };

function Container(type, children) { this.type = type; this.children = children || []; }
Container.prototype.getType = function () { return this.type; };
Container.prototype.getNumChildren = function () { return this.children.length; };
Container.prototype.getChild = function (i) { return this.children[i]; };
Container.prototype.clone = function () {
  return new Container(this.type, this.children.map(function (c) { return c.clone(); }));
};
Container.prototype.allText = function () {
  var out = [];
  (function walk(n) {
    if (n instanceof Text) out.push(n.str); else n.children.forEach(walk);
  })(this);
  return out.join('\n');
};

function para(s) { return new Container('PARAGRAPH', [new Text(s)]); }
function table(cells) {
  return new Container('TABLE', [new Container('TABLE_ROW', cells.map(function (c) {
    return new Container('TABLE_CELL', [para(c)]);
  }))]);
}

function Doc(body, header, footer) {
  this.body = body; this.header = header || null; this.footer = footer || null;
}
Doc.prototype.getBody = function () { return this.body; };
Doc.prototype.getHeader = function () { return this.header; };
Doc.prototype.getFooter = function () { return this.footer; };
Doc.prototype.saveAndClose = function () {};
Doc.prototype.clone = function () {
  return new Doc(this.body.clone(), this.header && this.header.clone(), this.footer && this.footer.clone());
};
Doc.prototype.allText = function () {
  return [this.body, this.header, this.footer].filter(Boolean).map(function (s) { return s.allText(); }).join('\n');
};

function makeEnvironment() {
  var env = { docs: {}, files: {}, folderFiles: [], trashed: [], created: [], nextId: 1, failPdfFor: null };

  function newId() { return 'id' + 'x'.repeat(20) + (env.nextId++); }

  function DriveFile(id, doc, name) {
    this.id = id; this.doc = doc; this.name = name;
  }
  DriveFile.prototype.getId = function () { return this.id; };
  DriveFile.prototype.makeCopy = function (name) {
    var id = newId();
    var copy = new DriveFile(id, this.doc.clone(), name);
    env.docs[id] = copy.doc;
    env.files[id] = copy;
    return copy;
  };
  DriveFile.prototype.getAs = function () {
    var self = this;
    if (env.failPdfFor && this.doc.allText().indexOf(env.failPdfFor) !== -1) throw new Error('export failed');
    return { text: this.doc.allText(), setName: function (n) { this.name = n; return this; }, name: null };
  };
  DriveFile.prototype.setTrashed = function (v) { if (v) env.trashed.push(this.name); };

  env.addTemplate = function (doc) {
    var id = newId();
    env.docs[id] = doc;
    env.files[id] = new DriveFile(id, doc, 'template');
    return id;
  };

  env.folder = {
    getFilesByName: function (n) {
      var hit = env.folderFiles.indexOf(n) !== -1;
      return { hasNext: function () { return hit; } };
    },
    createFile: function (blob) { env.folderFiles.push(blob.name); env.created.push(blob); }
  };

  env.setSheet = function (rows, name) {
    env.sheet = {
      getName: function () { return name || 'Sheet1'; },
      getLastRow: function () { return rows.length; },
      getLastColumn: function () { return rows.length ? rows[0].length : 0; },
      getRange: function (r, c, nr, nc) {
        return { getDisplayValues: function () {
          return rows.slice(r - 1, r - 1 + nr).map(function (row) { return row.slice(c - 1, c - 1 + nc); });
        } };
      },
      getDataRange: function () { return { getDisplayValues: function () { return rows; } }; }
    };
  };

  env.sandbox = {
    console: console,
    Date: Date,
    DocumentApp: {
      ElementType: { TEXT: 'TEXT' },
      openById: function (id) { if (!env.docs[id]) throw new Error('No Doc with id ' + id); return env.docs[id]; }
    },
    DriveApp: {
      getFileById: function (id) { if (!env.files[id]) throw new Error('No file ' + id); return env.files[id]; },
      getFolderById: function () { return env.folder; }
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: function () {
        return { getSheetByName: function () { return null; }, getActiveSheet: function () { return env.sheet; } };
      }
    },
    LockService: { getDocumentLock: function () { return { tryLock: function () { return true; }, releaseLock: function () {} }; } }
  };
  vm.createContext(env.sandbox);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), env.sandbox, { filename: 'Code.gs' });
  env.gs = env.sandbox;
  return env;
}

// ---------- tiny test runner ----------

var passed = 0, failed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok   ' + name); }
  catch (e) { failed++; console.log('  FAIL ' + name + '\n       ' + (e && e.message)); }
}

// ---------- tests ----------

console.log('Replacing placeholders');

test('replaces in body, table, header and footer', function () {
  var env = makeEnvironment();
  var doc = new Doc(
    new Container('BODY', [para('Dear {{first_name}} {{last_name}},'), table(['{{country}}', 'x'])]),
    new Container('HEADER', [table(['{{institution}}'])]),
    new Container('FOOTER', [para('Page {{ref}}')])
  );
  env.gs.replacePlaceholdersInDoc_(doc, { first_name: 'Ada', last_name: 'Lovelace', country: 'UK', institution: 'Royal Society', ref: 'A1' });
  assert.strictEqual(doc.allText(), 'Dear Ada Lovelace,\nUK\nx\nRoyal Society\nPage A1');
});

test('special characters in values are inserted exactly as they are', function () {
  var env = makeEnvironment();
  var values = ['US$ 500', 'costs $1 and $&', 'C:\\Files\\x', 'a (b) [c] {d} ^e', '$$'];
  values.forEach(function (v) {
    var doc = new Doc(new Container('BODY', [para('Value: {{v}}.')]));
    env.gs.replacePlaceholdersInDoc_(doc, { v: v });
    assert.strictEqual(doc.allText(), 'Value: ' + v + '.');
  });
});

test('same placeholder several times, neighbours and spaces in braces', function () {
  var env = makeEnvironment();
  var doc = new Doc(new Container('BODY', [para('{{a}}{{b}} and {{ a }} then {{a}}')]));
  env.gs.replacePlaceholdersInDoc_(doc, { a: 'X', b: 'YY' });
  assert.strictEqual(doc.allText(), 'XYY and X then X');
});

test('unmapped placeholders are left untouched', function () {
  var env = makeEnvironment();
  var doc = new Doc(new Container('BODY', [para('Hi {{name}}, ref {{other}}')]));
  env.gs.replacePlaceholdersInDoc_(doc, { name: 'Bo' });
  assert.strictEqual(doc.allText(), 'Hi Bo, ref {{other}}');
});

test('empty value removes the placeholder, also when it is the whole paragraph', function () {
  var env = makeEnvironment();
  var doc = new Doc(new Container('BODY', [para('Dear {{title}} Bo'), para('{{note}}')]));
  env.gs.replacePlaceholdersInDoc_(doc, { title: '', note: '' });
  assert.strictEqual(doc.allText(), 'Dear  Bo\n');
});

test('a value that looks like a placeholder is not replaced again', function () {
  var env = makeEnvironment();
  var doc = new Doc(new Container('BODY', [para('{{a}} / {{b}}')]));
  env.gs.replacePlaceholdersInDoc_(doc, { a: '{{b}}', b: 'B' });
  assert.strictEqual(doc.allText(), '{{b}} / B');
});

console.log('Finding placeholders');

test('finds unique, trimmed placeholders in body, tables, header and footer', function () {
  var env = makeEnvironment();
  var id = env.addTemplate(new Doc(
    new Container('BODY', [para('{{a}} {{ b }} {{a}}'), table(['{{c}}'])]),
    new Container('HEADER', [para('{{d}}')]),
    new Container('FOOTER', [para('no fields')])
  ));
  var res = env.gs.getPlaceholdersFromDoc('https://docs.google.com/document/d/' + id + '/edit');
  assert.strictEqual(res.error, null);
  assert.deepStrictEqual(Array.from(res.placeholders), ['a', 'b', 'c', 'd']);
});

test('reports a clear error for a bad link', function () {
  var env = makeEnvironment();
  var res = env.gs.getPlaceholdersFromDoc('not a link');
  assert.ok(/Cannot open the Doc/.test(res.error));
});

console.log('File names');

test('builds names, row numbers and removes illegal characters', function () {
  var env = makeEnvironment();
  assert.strictEqual(env.gs.buildFileName_('Letter_{{country}}_{{row}}', { country: 'Italy' }, 7), 'Letter_Italy_7');
  assert.strictEqual(env.gs.buildFileName_('{{n}}', { n: 'a/b:c*d?e"f<g>h|i\\j' }, 1), 'a_b_c_d_e_f_g_h_i_j');
  assert.strictEqual(env.gs.buildFileName_('', {}, 3), 'Letter_3');
  assert.strictEqual(env.gs.buildFileName_('{{missing}}', {}, 4), 'Letter_row_4');
  assert.strictEqual(env.gs.buildFileName_('x{{n}}', { n: 'y'.repeat(300) }, 1).length, 120);
});

test('flags placeholders in the pattern that have no column', function () {
  var env = makeEnvironment();
  var unknown = env.gs.unknownNamePlaceholders_('A_{{x}}_{{row}}_{{y}}', { x: 0 });
  assert.deepStrictEqual(Array.from(unknown), ['y']);
});

test('reads ids from different kinds of Google links', function () {
  var env = makeEnvironment();
  var id = 'abcDEF123_-abcDEF123_-';
  assert.strictEqual(env.gs.extractIdFromUrl('https://docs.google.com/document/d/' + id + '/edit'), id);
  assert.strictEqual(env.gs.extractIdFromUrl('https://drive.google.com/drive/u/0/folders/' + id + '?usp=sharing'), id);
  assert.strictEqual(env.gs.extractIdFromUrl('https://drive.google.com/open?id=' + id), id);
  assert.strictEqual(env.gs.extractIdFromUrl(id), id);
  assert.throws(function () { env.gs.extractIdFromUrl('hello'); });
});

console.log('Sheet reading');

test('getSheetData uses displayed values and counts data rows', function () {
  var env = makeEnvironment();
  env.setSheet([['First name', '', 'Date'], ['Ada', 'x', '05/11/2026']]);
  var data = env.gs.getSheetData();
  assert.strictEqual(data.totalRows, 1);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(data.headers)), [
    { label: 'First name', index: 0 }, { label: 'Column 2', index: 1 }, { label: 'Date', index: 2 }
  ]);
});

console.log('Pre-flight');

test('previews the chosen row and warns about empty cells and errors', function () {
  var env = makeEnvironment();
  env.setSheet([['first', 'last', 'city'], ['Ada', 'Lovelace', 'London'], ['Bo', '', '#N/A']]);
  var res = env.gs.preflightCheck({
    mapping: { first: 0, last: 1, city: 2 },
    fieldConfigs: { last: { behaviour: 'skip' }, city: { behaviour: 'default', defaultValue: 'n/a' } },
    placeholders: ['first', 'last', 'city', 'extra'],
    fileNamePattern: '{{first}}_{{typo}}',
    previewRow: 2
  });
  assert.ok(res.success);
  assert.strictEqual(res.rowNumber, 2);
  assert.strictEqual(res.preview.city, 'n/a');
  assert.strictEqual(res.fileName, 'Bo_.pdf');
  var text = res.warnings.join('|');
  assert.ok(/Unmapped placeholders: \{\{extra\}\}/.test(text));
  assert.ok(/SKIPPED/.test(text));
  assert.ok(/sheet error/.test(text));
  assert.ok(/\{\{typo\}\}/.test(text));
});

console.log('Generating PDFs');

function setupBatch() {
  var env = makeEnvironment();
  var templateId = env.addTemplate(new Doc(
    new Container('BODY', [para('Dear {{name}}, you owe {{amount}}.')]),
    new Container('HEADER', [table(['{{city}}'])])
  ));
  env.setSheet([
    ['name', 'amount', 'city'],
    ['Ada', 'US$ 5', 'London'],     // 1
    ['', '', ''],                   // 2 empty row
    ['', '$1', 'Rome'],             // 3 name empty -> skip
    ['Eve', '7', 'Paris'],          // 4
    ['Ada', '9', 'Oslo'],           // 5 same file name as row 1
    ['Zed', '1', 'FAIL']            // 6 export fails
  ]);
  env.failPdfFor = 'FAIL';
  env.config = {
    docUrl: 'https://docs.google.com/document/d/' + templateId + '/edit',
    folderUrl: 'https://drive.google.com/drive/folders/' + 'f'.repeat(25),
    fileNamePattern: 'Letter_{{name}}',
    mapping: { name: 0, amount: 1, city: 2 },
    fieldConfigs: { name: { behaviour: 'skip' } },
    placeholders: ['name', 'amount', 'city'],
    skipIfEmpty: true,
    rowFrom: 1, rowTo: 6
  };
  return env;
}

test('creates PDFs, skips rows, keeps names unique and cleans up the working copies', function () {
  var env = setupBatch();
  var res = env.gs.generatePDFs(env.config);
  assert.ok(res.success, res.error);
  assert.strictEqual(res.successCount, 3);
  assert.strictEqual(res.skippedCount, 2);
  assert.strictEqual(res.errorCount, 1);
  assert.strictEqual(res.nextRow, null);
  assert.deepStrictEqual(env.folderFiles, ['Letter_Ada.pdf', 'Letter_Eve.pdf', 'Letter_Ada_2.pdf']);
  assert.ok(env.created[0].text.indexOf('Dear Ada, you owe US$ 5.') !== -1);
  assert.ok(env.created[0].text.indexOf('London') !== -1);
  // every working copy is trashed, including the one whose export failed
  var tempTrashed = env.trashed.filter(function (n) { return n.indexOf('_temp_pdf_gen_') === 0; });
  assert.strictEqual(tempTrashed.length, 4);
  var errRow = res.results.filter(function (r) { return r.status === 'error'; })[0];
  assert.strictEqual(errRow.row, 6);
  assert.ok(/export failed/.test(errRow.error));
});

test('respects the row range', function () {
  var env = setupBatch();
  env.config.rowFrom = 4; env.config.rowTo = 4;
  var res = env.gs.generatePDFs(env.config);
  assert.strictEqual(res.successCount, 1);
  assert.deepStrictEqual(env.folderFiles, ['Letter_Eve.pdf']);
});

test('refuses to start with unmapped placeholders unless forced', function () {
  var env = setupBatch();
  env.config.placeholders = ['name', 'amount', 'city', 'extra'];
  var res = env.gs.generatePDFs(env.config);
  assert.strictEqual(res.success, false);
  assert.strictEqual(res.validationError, true);
  assert.deepStrictEqual(env.folderFiles, []);
  env.config.forceGenerate = true;
  env.config.rowFrom = 4; env.config.rowTo = 4;
  assert.strictEqual(env.gs.generatePDFs(env.config).successCount, 1);
});

test('stops cleanly when time runs short and says where to continue', function () {
  var env = setupBatch();
  var realNow = Date.now;
  var calls = 0;
  env.sandbox.Date = { now: function () { calls++; return calls <= 3 ? 0 : 10 * 60 * 1000; } };
  var res = env.gs.generatePDFs(env.config);
  env.sandbox.Date = Date;
  Date.now = realNow;
  assert.ok(res.success);
  assert.ok(res.nextRow > 1 && res.nextRow <= 6, 'nextRow was ' + res.nextRow);
  assert.strictEqual(res.results.length, res.nextRow - 1);
});

test('reports an error when the sheet has no data rows', function () {
  var env = setupBatch();
  env.setSheet([['name']]);
  var res = env.gs.generatePDFs(env.config);
  assert.strictEqual(res.success, false);
  assert.ok(/No data rows/.test(res.error));
});

console.log('\n' + passed + ' passed, ' + failed + ' failed');
process.exit(failed ? 1 : 0);
