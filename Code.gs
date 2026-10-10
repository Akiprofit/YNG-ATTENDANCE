const CONFIG = {
  ATTENDANCE_SHEET: 'Attendance',
  REGISTRY_SHEET: 'People Registry',
  HEADER_MONTH_ROW: 3,
  HEADER_DAY_ROW: 4,
  FIRST_PERSON_ROW: 5,
  NAME_COLUMN: 1,
  FIRST_DATE_COLUMN: 2,
  LAST_DATE_COLUMN: 51, // AY
  MIRROR_NAME_COLUMN: 52, // AZ
  PRESENT_COUNT_COLUMN: 53, // BA
  LAST_PERSON_COLUMN: 53, // BA
  YEAR: 2026
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Church Attendance POS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Run once manually from Apps Script to repair existing rows and formulas. */
function fixAttendanceRows() {
  const sheet = getAttendanceSheet_();
  repairAttendanceLayout_(sheet);
  ensureTotalsRow_(sheet);
  syncPeopleMirrorAndCountFormulas_(sheet);
  SpreadsheetApp.flush();
  return 'Attendance rows fixed: names in A/AZ, per-person P counts in BA, totals below attendees.';
}

function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getAttendanceSheet_();
  const registry = getRegistryMap_();
  const people = getPeopleRows_(sheet).map(person => ({
    id: person.row,
    name: person.name,
    active: registry[person.row] ? registry[person.row].active : true
  })).filter(person => person.active);

  return {
    people: people,
    dates: getDateOptions_(sheet),
    spreadsheetName: ss.getName()
  };
}

function getDateOptions_(sheet) {
  const width = CONFIG.LAST_DATE_COLUMN - CONFIG.FIRST_DATE_COLUMN + 1;
  const months = sheet.getRange(CONFIG.HEADER_MONTH_ROW, CONFIG.FIRST_DATE_COLUMN, 1, width).getDisplayValues()[0];
  const days = sheet.getRange(CONFIG.HEADER_DAY_ROW, CONFIG.FIRST_DATE_COLUMN, 1, width).getValues()[0];
  let currentMonth = '';
  const monthNumbers = {
    JANUARY: 1, FEBRUARY: 2, MARCH: 3, APRIL: 4, MAY: 5, JUNE: 6,
    JULY: 7, AUGUST: 8, SEPTEMBER: 9, OCTOBER: 10, NOVEMBER: 11, DECEMBER: 12
  };
  const dates = [];

  for (let i = 0; i < width; i++) {
    const heading = String(months[i] || '').trim().toUpperCase();
    if (heading && monthNumbers[heading]) currentMonth = heading;
    const day = Number(days[i]);
    if (!currentMonth || !Number.isInteger(day) || day < 1 || day > 31) continue;

    const month = monthNumbers[currentMonth];
    const date = new Date(CONFIG.YEAR, month - 1, day, 12, 0, 0);
    if (date.getFullYear() !== CONFIG.YEAR || date.getMonth() !== month - 1 || date.getDate() !== day) continue;

    const value = CONFIG.YEAR + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
    dates.push({
      value: value,
      label: Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM d, yyyy'),
      column: CONFIG.FIRST_DATE_COLUMN + i
    });
  }
  return dates;
}

function getAttendanceForDate(dateISO) {
  const sheet = getAttendanceSheet_();
  const dateInfo = findDateColumn_(sheet, dateISO);
  const people = getPeopleRows_(sheet);
  const registry = getRegistryMap_();
  const result = [];

  people.forEach(person => {
    const entry = registry[person.row];
    const active = entry ? entry.active : true;
    if (!active) return;
    const value = String(sheet.getRange(person.row, dateInfo.column).getDisplayValue()).trim().toUpperCase();
    if (value === 'P') result.push(person.row);
  });
  return result;
}

function saveAttendance(dateISO, presentIds) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const sheet = getAttendanceSheet_();
    const dateInfo = findDateColumn_(sheet, dateISO);
    const registry = getRegistryMap_();
    const presentSet = new Set((presentIds || []).map(value => Number(value)));
    const people = getPeopleRows_(sheet);
    if (!people.length) throw new Error('No church attendees were found in column A.');

    let written = 0;
    people.forEach(person => {
      const entry = registry[person.row];
      const active = entry ? entry.active : true;
      if (!active) return;
      sheet.getRange(person.row, dateInfo.column).setValue(presentSet.has(person.row) ? 'P' : 'A');
      written++;
    });

    syncPeopleMirrorAndCountFormulas_(sheet);
    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Attendance saved for ' + dateInfo.label + '. ' + written + ' people updated.',
      date: dateISO,
      updated: written
    };
  } finally {
    lock.releaseLock();
  }
}

function addPerson(name) {
  name = String(name || '').trim().replace(/\s+/g, ' ');
  if (!name) throw new Error('Enter a name first.');

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const sheet = getAttendanceSheet_();
    const registrySheet = getOrCreateRegistrySheet_();

    // Repair any names accidentally placed below the totals row before adding another person.
    repairAttendanceLayout_(sheet);

    const existing = getPeopleRows_(sheet);
    if (existing.some(person => person.name.toLocaleLowerCase() === name.toLocaleLowerCase())) {
      throw new Error('That name already exists in the attendance sheet.');
    }

    const totalRow = findTotalRow_(sheet);
    const newRow = totalRow || (existing.length ? existing[existing.length - 1].row + 1 : CONFIG.FIRST_PERSON_ROW);
    if (newRow > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), newRow - sheet.getMaxRows());

    // Insert before totals so the totals row always remains beneath the full people list.
    if (totalRow) sheet.insertRowsBefore(totalRow, 1);

    const templateRow = findAttendanceTemplateRow_(sheet);
    const dateWidth = CONFIG.LAST_DATE_COLUMN - CONFIG.FIRST_DATE_COLUMN + 1;
    sheet.getRange(templateRow, 1, 1, CONFIG.LAST_PERSON_COLUMN).copyTo(
      sheet.getRange(newRow, 1, 1, CONFIG.LAST_PERSON_COLUMN),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );
    sheet.getRange(templateRow, CONFIG.FIRST_DATE_COLUMN, 1, dateWidth).copyTo(
      sheet.getRange(newRow, CONFIG.FIRST_DATE_COLUMN, 1, dateWidth),
      SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
      false
    );

    // Blank all attendance values but keep the copied P/A dropdowns and formatting.
    sheet.getRange(newRow, 1, 1, CONFIG.LAST_PERSON_COLUMN).clearContent();
    sheet.getRange(newRow, CONFIG.NAME_COLUMN).setValue(name);
    sheet.getRange(newRow, CONFIG.MIRROR_NAME_COLUMN).setValue(name);
    sheet.getRange(newRow, CONFIG.PRESENT_COUNT_COLUMN).setFormula(
      '=COUNTIF(B' + newRow + ':AY' + newRow + ', \"P\")'
    );

    ensureTotalsRow_(sheet);
    syncPeopleMirrorAndCountFormulas_(sheet);
    SpreadsheetApp.flush();

    ensureRegistryHeader_(registrySheet);
    registrySheet.appendRow([newRow, name, true, new Date()]);
    return { id: newRow, name: name, active: true };
  } finally {
    lock.releaseLock();
  }
}

/** Finds the totals row by label or COUNTIF formulas in the date columns. */
function findTotalRow_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.FIRST_PERSON_ROW) return null;
  const rowCount = lastRow - CONFIG.FIRST_PERSON_ROW + 1;
  const names = sheet.getRange(CONFIG.FIRST_PERSON_ROW, CONFIG.NAME_COLUMN, rowCount, 1).getDisplayValues();
  const formulas = sheet.getRange(CONFIG.FIRST_PERSON_ROW, CONFIG.FIRST_DATE_COLUMN, rowCount,
    CONFIG.LAST_DATE_COLUMN - CONFIG.FIRST_DATE_COLUMN + 1).getFormulas();

  for (let i = 0; i < rowCount; i++) {
    const row = CONFIG.FIRST_PERSON_ROW + i;
    const label = String(names[i][0] || '').toLowerCase();
    if (label.includes('total') && label.includes('attendee')) return row;
    if (formulas[i].some(formula => /COUNTIF\s*\(/i.test(formula))) return row;
  }
  return null;
}

/**
 * Moves named attendee rows found below the totals row back above it.
 * This repairs names previously added beneath the totals (for example, row 284).
 * Existing attendance values are preserved when moving a misplaced row.
 */
function repairAttendanceLayout_(sheet) {
  let totalRow = findTotalRow_(sheet);
  if (!totalRow) return;

  // Scan bottom-up so each moved row is handled once and formulas/totals remain below people.
  for (let row = sheet.getLastRow(); row > totalRow; row--) {
    const name = String(sheet.getRange(row, CONFIG.NAME_COLUMN).getDisplayValue() || '').trim();
    if (!name || /^total\s*#?\s*of\s*attendees:?$/i.test(name) || /^(NAMES:|PRESENT|ABSENT)$/i.test(name)) continue;

    const rowValues = sheet.getRange(row, 1, 1, CONFIG.LAST_PERSON_COLUMN).getValues()[0];
    const hasAttendance = rowValues.slice(CONFIG.FIRST_DATE_COLUMN - 1).some(value => value !== '' && value !== null);
    const hasDropdowns = sheet.getRange(row, CONFIG.FIRST_DATE_COLUMN, 1,
      CONFIG.LAST_DATE_COLUMN - CONFIG.FIRST_DATE_COLUMN + 1).getDataValidations()[0].some(rule => rule !== null);
    const registry = getRegistryMap_();
    const registered = registry[row] && String(registry[row].name || '').trim().toLowerCase() === name.toLowerCase();
    // Don't move unrelated notes below the table; only move rows that resemble attendee rows.
    if (!hasAttendance && !hasDropdowns && !registered) continue;

    const templateRow = findAttendanceTemplateRow_(sheet);
    const valuesToMove = rowValues.slice(0, CONFIG.LAST_PERSON_COLUMN);
    sheet.insertRowsBefore(totalRow, 1);
    const destinationRow = totalRow;
    const shiftedSourceRow = row + 1;

    sheet.getRange(templateRow, 1, 1, CONFIG.LAST_PERSON_COLUMN).copyTo(
      sheet.getRange(destinationRow, 1, 1, CONFIG.LAST_PERSON_COLUMN),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false
    );
    sheet.getRange(templateRow, CONFIG.FIRST_DATE_COLUMN, 1,
      CONFIG.LAST_DATE_COLUMN - CONFIG.FIRST_DATE_COLUMN + 1).copyTo(
      sheet.getRange(destinationRow, CONFIG.FIRST_DATE_COLUMN, 1,
        CONFIG.LAST_DATE_COLUMN - CONFIG.FIRST_DATE_COLUMN + 1),
      SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false
    );
    sheet.getRange(destinationRow, 1, 1, CONFIG.LAST_PERSON_COLUMN).setValues([valuesToMove]);
    sheet.getRange(shiftedSourceRow, 1, 1, CONFIG.LAST_PERSON_COLUMN).clearContent();

    // Keep an existing registry record attached to the moved person.
    const registrySheet = getOrCreateRegistrySheet_();
    ensureRegistryHeader_(registrySheet);
    const registryLastRow = registrySheet.getLastRow();
    if (registryLastRow >= 2) {
      const registryRows = registrySheet.getRange(2, 1, registryLastRow - 1, 4).getValues();
      for (let i = 0; i < registryRows.length; i++) {
        if (Number(registryRows[i][0]) === row && String(registryRows[i][1] || '').trim().toLowerCase() === name.toLowerCase()) {
          registrySheet.getRange(i + 2, 1).setValue(destinationRow);
          break;
        }
      }
    }
    totalRow++;
  }
  ensureTotalsRow_(sheet);
  syncPeopleMirrorAndCountFormulas_(sheet);
}

/** Rebuilds totals formulas so every date counts all attendee rows above totals. */
function ensureTotalsRow_(sheet) {
  let totalRow = findTotalRow_(sheet);
  if (!totalRow) {
    totalRow = Math.max(CONFIG.FIRST_PERSON_ROW, sheet.getLastRow() + 1);
    sheet.getRange(totalRow, CONFIG.NAME_COLUMN).setValue('Total # of attendees:');
  }
  const lastPersonRow = totalRow - 1;
  const width = CONFIG.LAST_DATE_COLUMN - CONFIG.FIRST_DATE_COLUMN + 1;
  const formulas = [];
  for (let col = CONFIG.FIRST_DATE_COLUMN; col <= CONFIG.LAST_DATE_COLUMN; col++) {
    const letter = columnToLetter_(col);
    formulas.push('=COUNTIF(' + letter + CONFIG.FIRST_PERSON_ROW + ':' + letter + Math.max(CONFIG.FIRST_PERSON_ROW, lastPersonRow) + ',"P")');
  }
  sheet.getRange(totalRow, CONFIG.FIRST_DATE_COLUMN, 1, width).setFormulas([formulas]);
  sheet.getRange(totalRow, CONFIG.NAME_COLUMN).setValue('Total # of attendees:');
}

/** Keeps each attendee's name mirrored in AZ and present-count formula in BA. */
function syncPeopleMirrorAndCountFormulas_(sheet) {
  const people = getPeopleRows_(sheet);
  people.forEach(person => {
    sheet.getRange(person.row, CONFIG.MIRROR_NAME_COLUMN).setValue(person.name);
    sheet.getRange(person.row, CONFIG.PRESENT_COUNT_COLUMN).setFormula(
      '=COUNTIF(B' + person.row + ':AY' + person.row + ', "P")'
    );
  });
}

function columnToLetter_(column) {
  let letter = '';
  while (column > 0) {
    const remainder = (column - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    column = Math.floor((column - 1) / 26);
  }
  return letter;
}

function findAttendanceTemplateRow_(sheet) {
  const first = CONFIG.FIRST_PERSON_ROW;
  const width = CONFIG.LAST_DATE_COLUMN - CONFIG.FIRST_DATE_COLUMN + 1;

  // Prefer row 5 if it contains the original attendance dropdown validation.
  const firstValidation = sheet.getRange(first, CONFIG.FIRST_DATE_COLUMN, 1, width).getDataValidations()[0];
  if (firstValidation.some(rule => rule !== null)) return first;

  // Otherwise find the first row with at least one dropdown validation in B:AY.
  const scanEnd = Math.min(Math.max(sheet.getLastRow(), first + 20), first + 100);
  for (let row = first + 1; row <= scanEnd; row++) {
    const validations = sheet.getRange(row, CONFIG.FIRST_DATE_COLUMN, 1, width).getDataValidations()[0];
    if (validations.some(rule => rule !== null)) return row;
  }

  throw new Error('Could not find an existing P/A dropdown to copy. Check that an original attendance cell in B:AY still has its dropdown rule.');
}

function archivePeople(ids) {
  const idSet = new Set((ids || []).map(value => Number(value)));
  if (!idSet.size) return { success: true, archived: 0 };

  const sheet = getAttendanceSheet_();
  const registrySheet = getOrCreateRegistrySheet_();
  ensureRegistryHeader_(registrySheet);
  const registry = getRegistryMap_();
  const people = getPeopleRows_(sheet);
  let archived = 0;

  people.forEach(person => {
    if (!idSet.has(person.row)) return;
    const existing = registry[person.row];
    if (existing) {
      const row = existing.registryRow;
      registrySheet.getRange(row, 3).setValue(false);
      registrySheet.getRange(row, 4).setValue(new Date());
      registrySheet.getRange(row, 2).setValue(person.name);
    } else {
      registrySheet.appendRow([person.row, person.name, false, new Date()]);
    }
    archived++;
  });
  return { success: true, archived: archived };
}

function restorePerson(rowId) {
  const id = Number(rowId);
  const sheet = getAttendanceSheet_();
  const registrySheet = getOrCreateRegistrySheet_();
  ensureRegistryHeader_(registrySheet);
  const person = getPeopleRows_(sheet).find(item => item.row === id);
  if (!person) throw new Error('Person was not found.');
  const registry = getRegistryMap_();
  if (registry[id]) {
    registrySheet.getRange(registry[id].registryRow, 3).setValue(true);
    registrySheet.getRange(registry[id].registryRow, 4).setValue(new Date());
  } else {
    registrySheet.appendRow([id, person.name, true, new Date()]);
  }
  return { success: true, id: id, name: person.name };
}

function getArchivedPeople() {
  const sheet = getAttendanceSheet_();
  const registry = getRegistryMap_();
  return getPeopleRows_(sheet)
    .filter(person => registry[person.row] && !registry[person.row].active)
    .map(person => ({ id: person.row, name: person.name }));
}

function findDateColumn_(sheet, dateISO) {
  const requested = String(dateISO || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!requested) throw new Error('Choose a valid service date.');
  const options = getDateOptions_(sheet);
  const found = options.find(item => item.value === dateISO);
  if (!found) throw new Error('That date does not exist in the attendance table. Choose one of the listed dates.');
  return found;
}

function getPeopleRows_(sheet) {
  const totalRow = findTotalRow_(sheet);
  const lastPersonRow = totalRow ? totalRow - 1 : sheet.getLastRow();
  if (lastPersonRow < CONFIG.FIRST_PERSON_ROW) return [];
  const values = sheet.getRange(CONFIG.FIRST_PERSON_ROW, CONFIG.NAME_COLUMN,
    lastPersonRow - CONFIG.FIRST_PERSON_ROW + 1, 1).getDisplayValues();
  const people = [];
  values.forEach((row, index) => {
    const name = String(row[0] || '').trim();
    if (name && name.toUpperCase() !== 'NAMES:' && name.toUpperCase() !== 'PRESENT' && name.toUpperCase() !== 'ABSENT') {
      people.push({ row: CONFIG.FIRST_PERSON_ROW + index, name: name });
    }
  });
  return people;
}

function getAttendanceSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  if (!sheet) throw new Error('Could not find the "' + CONFIG.ATTENDANCE_SHEET + '" sheet.');
  return sheet;
}

function getOrCreateRegistrySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.REGISTRY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.REGISTRY_SHEET);
    sheet.getRange(1, 1, 1, 4).setValues([['Attendance Row ID', 'Name', 'Active', 'Last Updated']]);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

function ensureRegistryHeader_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([['Attendance Row ID', 'Name', 'Active', 'Last Updated']]);
    sheet.setFrozenRows(1);
  }
}

function getRegistryMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.REGISTRY_SHEET);
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const map = {};
  rows.forEach((row, index) => {
    const id = Number(row[0]);
    if (!id) return;
    map[id] = {
      id: id,
      name: String(row[1] || ''),
      active: row[2] !== false && String(row[2]).toUpperCase() !== 'FALSE',
      updated: row[3],
      registryRow: index + 2
    };
  });
  return map;
}