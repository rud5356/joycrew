/**
 * Joycrew 신규입장자 입력 도구
 *
 * 사용 방법
 * 1. 구글시트에서 확장 프로그램 > Apps Script를 엽니다.
 * 2. 이 파일 전체를 붙여넣고 저장합니다.
 * 3. 시트를 새로고침한 뒤 상단 메뉴의 "쪼이크루"를 사용합니다.
 * 4. "입력/결과 시트 만들기"를 먼저 누르고, "신규입장자_입력" 시트 A2 아래에 원문을 붙여넣습니다.
 * 5. "신규입장자 반영"을 누르면 비어 있는 A열만 채우고, 이미 값이 있는 번호는 결과 시트에 표시합니다.
 */

const JOYCREW_CONFIG = {
  // 실제 명단이 있는 시트 이름입니다. 비워두면 입력/결과 시트를 제외한 첫 번째 시트를 사용합니다.
  targetSheetName: '',
  inputSheetName: '신규입장자_입력',
  resultSheetName: '신규입장자_결과',
  actionInputSheetName: '언팔차단_입력',
  actionResultSheetName: '언팔차단_결과',
  targetColumn: 1,
  unfollowColumn: 4,
  blockColumn: 5,
  inputCell: 'A2',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('쪼이크루')
    .addItem('입력/결과 시트 만들기', 'setupJoycrewNewEntrantSheets')
    .addItem('신규입장자 반영', 'applyJoycrewNewEntrants')
    .addSeparator()
    .addItem('언팔/차단 반영', 'applyJoycrewActionList')
    .addToUi();
}

function setupJoycrewNewEntrantSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = getOrCreateSheet_(spreadsheet, JOYCREW_CONFIG.inputSheetName);
  const resultSheet = getOrCreateSheet_(spreadsheet, JOYCREW_CONFIG.resultSheetName);
  const actionInputSheet = getOrCreateSheet_(spreadsheet, JOYCREW_CONFIG.actionInputSheetName);
  const actionResultSheet = getOrCreateSheet_(spreadsheet, JOYCREW_CONFIG.actionResultSheetName);

  inputSheet.clear();
  inputSheet.getRange('A1').setValue('아래 A2 칸에 신규입장자 원문을 그대로 붙여넣고, 메뉴에서 "쪼이크루 > 신규입장자 반영"을 누르세요.');
  inputSheet.getRange(JOYCREW_CONFIG.inputCell).setValue('');
  inputSheet.getRange(JOYCREW_CONFIG.inputCell).setWrap(true);
  inputSheet.setColumnWidth(1, 760);
  inputSheet.setRowHeight(2, 360);

  resultSheet.clear();
  writeResultHeader_(resultSheet);
  resultSheet.setFrozenRows(1);
  resultSheet.setColumnWidths(1, 5, 160);

  actionInputSheet.clear();
  actionInputSheet.getRange('A1').setValue('아래 A2 칸에 "언팔" 또는 "차단"을 첫 줄에 쓰고, 그 아래에 @아이디나 인스타그램 주소를 붙여넣은 뒤 "쪼이크루 > 언팔/차단 반영"을 누르세요.');
  actionInputSheet.getRange(JOYCREW_CONFIG.inputCell).setValue('');
  actionInputSheet.getRange(JOYCREW_CONFIG.inputCell).setWrap(true);
  actionInputSheet.setColumnWidth(1, 760);
  actionInputSheet.setRowHeight(2, 360);

  actionResultSheet.clear();
  writeActionResultHeader_(actionResultSheet);
  actionResultSheet.setFrozenRows(1);
  actionResultSheet.setColumnWidths(1, 4, 160);

  SpreadsheetApp.getUi().alert('신규입장자/언팔차단 입력 시트와 결과 시트를 만들었어요.');
}

function applyJoycrewNewEntrants() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = findTargetSheet_(spreadsheet);
  const inputSheet = spreadsheet.getSheetByName(JOYCREW_CONFIG.inputSheetName);

  if (!inputSheet) {
    SpreadsheetApp.getUi().alert('먼저 "쪼이크루 > 입력/결과 시트 만들기"를 실행해주세요.');
    return;
  }

  const rawText = getJoycrewRawInputText_(spreadsheet, inputSheet);
  const parsed = parseJoycrewEntrants_(rawText);

  if (parsed.entries.length === 0) {
    SpreadsheetApp.getUi().alert('번호와 @아이디 형식의 신규입장자를 찾지 못했어요. A2 아래에 원문을 붙여넣었는지 확인해주세요.');
    return;
  }

  const maxNumber = Math.max(...parsed.entries.map(entry => entry.number));
  ensureRows_(targetSheet, maxNumber);

  const existingValues = targetSheet
    .getRange(1, JOYCREW_CONFIG.targetColumn, maxNumber, 1)
    .getDisplayValues()
    .map(row => normalizeJoycrewId_(row[0]));

  const results = [];
  let updatedCount = 0;
  let occupiedCount = 0;
  let sameCount = 0;

  parsed.entries.forEach(entry => {
    const existingId = existingValues[entry.number - 1];

    if (!existingId) {
      targetSheet.getRange(entry.number, JOYCREW_CONFIG.targetColumn).setValue(entry.id);
      existingValues[entry.number - 1] = entry.id;
      updatedCount += 1;
      results.push([entry.number, entry.id, '입력 완료', '']);
      return;
    }

    if (existingId === entry.id) {
      sameCount += 1;
      results.push([entry.number, entry.id, '이미 같은 아이디 있음', existingId]);
      return;
    }

    occupiedCount += 1;
    results.push([entry.number, entry.id, '이미 값 있음 - 확인 필요', existingId]);
  });

  parsed.duplicates.forEach(item => {
    results.push([item.number, item.id, '입력문 안에서 중복 번호 - 건너뜀', item.firstId]);
  });

  const resultSheet = getOrCreateSheet_(spreadsheet, JOYCREW_CONFIG.resultSheetName);
  resultSheet.clear();
  writeResultHeader_(resultSheet);
  if (results.length > 0) {
    resultSheet.getRange(2, 1, results.length, 4).setValues(results);
  }
  resultSheet.autoResizeColumns(1, 4);

  SpreadsheetApp.getUi().alert(
    [
      `대상 시트: ${targetSheet.getName()}`,
      `파싱: ${parsed.entries.length}건`,
      `입력 완료: ${updatedCount}건`,
      `이미 값 있음: ${occupiedCount}건`,
      `이미 같은 아이디: ${sameCount}건`,
      `입력문 중복 번호: ${parsed.duplicates.length}건`,
      '자세한 내용은 신규입장자_결과 시트를 확인해주세요.',
    ].join('\n')
  );
}

function applyJoycrewActionList() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = findTargetSheet_(spreadsheet);
  const inputSheet = spreadsheet.getSheetByName(JOYCREW_CONFIG.actionInputSheetName);

  if (!inputSheet) {
    SpreadsheetApp.getUi().alert('먼저 "쪼이크루 > 입력/결과 시트 만들기"를 실행해주세요.');
    return;
  }

  const rawText = getJoycrewRawInputText_(spreadsheet, inputSheet);
  const parsed = parseJoycrewActionList_(rawText);

  if (!parsed.action) {
    SpreadsheetApp.getUi().alert('첫 줄이나 내용에 "언팔" 또는 "차단"을 넣어주세요.');
    return;
  }

  if (parsed.ids.length === 0) {
    SpreadsheetApp.getUi().alert('@아이디 또는 instagram.com/아이디 형식의 계정을 찾지 못했어요.');
    return;
  }

  const actionLabel = parsed.action === 'block' ? '차단' : '언팔';
  const targetColumn = parsed.action === 'block'
    ? JOYCREW_CONFIG.blockColumn
    : JOYCREW_CONFIG.unfollowColumn;

  ensureColumns_(targetSheet, targetColumn);

  const existingValues = targetSheet
    .getRange(1, targetColumn, targetSheet.getMaxRows(), 1)
    .getDisplayValues()
    .map(row => normalizeJoycrewId_(row[0]));

  const existingRowsById = {};
  const blankRows = [];
  existingValues.forEach((id, index) => {
    const rowNumber = index + 1;
    if (id) {
      existingRowsById[id] = rowNumber;
    } else {
      blankRows.push(rowNumber);
    }
  });

  const results = [];
  let updatedCount = 0;
  let alreadyExistsCount = 0;

  parsed.ids.forEach(id => {
    if (existingRowsById[id]) {
      alreadyExistsCount += 1;
      results.push([existingRowsById[id], id, `${actionLabel} 목록에 이미 있음`, `${existingRowsById[id]}행`]);
      return;
    }

    let targetRow = blankRows.shift();
    if (!targetRow) {
      const currentRows = targetSheet.getMaxRows();
      targetSheet.insertRowsAfter(currentRows, 1);
      targetRow = currentRows + 1;
    }

    targetSheet.getRange(targetRow, targetColumn).setValue(id);
    existingRowsById[id] = targetRow;
    updatedCount += 1;
    results.push([targetRow, id, `${actionLabel} 입력 완료`, '']);
  });

  const resultSheet = getOrCreateSheet_(spreadsheet, JOYCREW_CONFIG.actionResultSheetName);
  resultSheet.clear();
  writeActionResultHeader_(resultSheet);
  if (results.length > 0) {
    resultSheet.getRange(2, 1, results.length, 4).setValues(results);
  }
  resultSheet.autoResizeColumns(1, 4);

  SpreadsheetApp.getUi().alert(
    [
      `대상 시트: ${targetSheet.getName()}`,
      `작업: ${actionLabel}`,
      `파싱: ${parsed.ids.length}건`,
      `입력 완료: ${updatedCount}건`,
      `이미 있음: ${alreadyExistsCount}건`,
      '자세한 내용은 언팔차단_결과 시트를 확인해주세요.',
    ].join('\n')
  );
}

function parseJoycrewEntrants_(text) {
  const entries = [];
  const duplicates = [];
  const seenByNumber = {};

  String(text || '').split(/\r?\n/).forEach(line => {
    const match = line.match(/^\s*(\d+)\s*[.)]\s*@?([A-Za-z0-9._]+)\b/);
    if (!match) return;

    const number = Number(match[1]);
    const id = normalizeJoycrewId_(match[2]);
    if (!number || !id) return;

    if (seenByNumber[number]) {
      duplicates.push({
        number,
        id,
        firstId: seenByNumber[number],
      });
      return;
    }

    seenByNumber[number] = id;
    entries.push({ number, id });
  });

  return { entries, duplicates };
}

function parseJoycrewActionList_(text) {
  const rawText = String(text || '');
  const action = detectJoycrewAction_(rawText);
  const ids = [];
  const seenIds = {};
  let duplicateCount = 0;

  rawText.split(/\r?\n/).forEach(line => {
    extractJoycrewIdsFromLine_(line).forEach(id => {
      if (seenIds[id]) {
        duplicateCount += 1;
        return;
      }

      seenIds[id] = true;
      ids.push(id);
    });
  });

  return { action, ids, duplicateCount };
}

function detectJoycrewAction_(text) {
  const normalized = String(text || '').trim();
  if (/(^|\s)(차단필수|차단)(\s|$)/m.test(normalized)) {
    return 'block';
  }
  if (/(^|\s)(언팔필수|언팔)(\s|$)/m.test(normalized)) {
    return 'unfollow';
  }
  return '';
}

function extractJoycrewIdsFromLine_(line) {
  const ids = [];
  const seenInLine = {};
  const text = String(line || '');

  addJoycrewIdMatches_(ids, seenInLine, text, /@([A-Za-z0-9._]+)/g);
  addJoycrewIdMatches_(ids, seenInLine, text, /(?:https?:\/\/)?(?:www\.)?instagram\.com\/([A-Za-z0-9._]+)/gi);

  const commandRemoved = text
    .replace(/^(언팔필수|언팔|차단필수|차단)\s*/g, '')
    .trim();

  if (!/@|instagram\.com/i.test(commandRemoved)) {
    const bareMatch = commandRemoved.match(/^([A-Za-z0-9._]+)$/);
    if (bareMatch) {
      addJoycrewId_(ids, seenInLine, bareMatch[1]);
    }
  }

  return ids;
}

function addJoycrewIdMatches_(ids, seenInLine, text, regex) {
  let match;
  while ((match = regex.exec(text)) !== null) {
    addJoycrewId_(ids, seenInLine, match[1]);
  }
}

function addJoycrewId_(ids, seenInLine, value) {
  const id = normalizeJoycrewId_(String(value || '').replace(/[),;:]+$/g, ''));
  if (!id || seenInLine[id]) return;

  seenInLine[id] = true;
  ids.push(id);
}

function getJoycrewRawInputText_(spreadsheet, inputSheet) {
  const inputText = collectColumnText_(inputSheet);
  if (inputText) return inputText;

  // 실수로 결과 시트나 현재 보고 있는 시트에 붙여넣어도 처리할 수 있게 한 번 더 확인합니다.
  const activeSheet = spreadsheet.getActiveSheet();
  if (activeSheet && activeSheet.getName() !== inputSheet.getName()) {
    return collectColumnText_(activeSheet);
  }

  return '';
}

function collectColumnText_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';

  return sheet
    .getRange(2, JOYCREW_CONFIG.targetColumn, lastRow - 1, 1)
    .getDisplayValues()
    .map(row => row[0])
    .join('\n')
    .trim();
}

function normalizeJoycrewId_(value) {
  const id = String(value || '')
    .trim()
    .replace(/^@+/, '')
    .toLowerCase();

  return /^[a-z0-9._]+$/.test(id) && id !== 'undefined' ? id : '';
}

function findTargetSheet_(spreadsheet) {
  if (JOYCREW_CONFIG.targetSheetName) {
    const namedSheet = spreadsheet.getSheetByName(JOYCREW_CONFIG.targetSheetName);
    if (!namedSheet) {
      throw new Error(`대상 시트를 찾을 수 없습니다: ${JOYCREW_CONFIG.targetSheetName}`);
    }
    return namedSheet;
  }

  const excludedNames = new Set([
    JOYCREW_CONFIG.inputSheetName,
    JOYCREW_CONFIG.resultSheetName,
    JOYCREW_CONFIG.actionInputSheetName,
    JOYCREW_CONFIG.actionResultSheetName,
  ]);

  const targetSheet = spreadsheet.getSheets().find(sheet => !excludedNames.has(sheet.getName()));
  if (!targetSheet) {
    throw new Error('명단 대상 시트를 찾을 수 없습니다.');
  }
  return targetSheet;
}

function ensureRows_(sheet, requiredRows) {
  const currentRows = sheet.getMaxRows();
  if (requiredRows > currentRows) {
    sheet.insertRowsAfter(currentRows, requiredRows - currentRows);
  }
}

function ensureColumns_(sheet, requiredColumns) {
  const currentColumns = sheet.getMaxColumns();
  if (requiredColumns > currentColumns) {
    sheet.insertColumnsAfter(currentColumns, requiredColumns - currentColumns);
  }
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function writeResultHeader_(sheet) {
  sheet.getRange(1, 1, 1, 4).setValues([[
    '번호',
    '입력 아이디',
    '결과',
    '기존 아이디',
  ]]);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
}

function writeActionResultHeader_(sheet) {
  sheet.getRange(1, 1, 1, 4).setValues([[
    '행',
    '입력 아이디',
    '결과',
    '기존 위치',
  ]]);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
}
