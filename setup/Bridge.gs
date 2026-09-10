const BRIDGE_VERSION = '1.0-template';

// Replace only these two values for each student.
const STUDENT_SPREADSHEET_ID = 'PASTE_SPREADSHEET_ID_HERE';
const STUDENT_APP_URL = 'PASTE_GITHUB_PAGES_APP_URL_HERE';

function setupBridge() {
  if (STUDENT_SPREADSHEET_ID === 'PASTE_SPREADSHEET_ID_HERE') {
    throw new Error('Enter the student spreadsheet ID first.');
  }

  const properties = PropertiesService.getScriptProperties();
  let token = properties.getProperty('ACCESS_TOKEN');

  if (!token) {
    token =
      Utilities.getUuid().replace(/-/g, '') +
      Utilities.getUuid().replace(/-/g, '');
  }

  properties.setProperties({
    SPREADSHEET_ID: STUDENT_SPREADSHEET_ID,
    ACCESS_TOKEN: token
  });

  properties.deleteProperty('MAX_ACCESSIBLE_DAY');

  console.log('Bridge setup complete.');
  console.log('PERSONAL ACCESS LINK: ' + buildPersonalAccessLink_(token));
}

function getPersonalAccessLink() {
  const token = PropertiesService
    .getScriptProperties()
    .getProperty('ACCESS_TOKEN');

  if (!token) {
    throw new Error('Run setupBridge first.');
  }

  const link = buildPersonalAccessLink_(token);
  console.log('PERSONAL ACCESS LINK: ' + link);
  return link;
}

function buildPersonalAccessLink_(token) {
  const appUrl = cleanText_(STUDENT_APP_URL);

  if (
    !appUrl ||
    appUrl === 'PASTE_GITHUB_PAGES_APP_URL_HERE' ||
    !/^https:\/\//i.test(appUrl)
  ) {
    throw new Error('Enter the student GitHub Pages App URL first.');
  }

  return appUrl.replace(/#.*$/, '') +
    '#token=' +
    encodeURIComponent(token);
}

function doGet(e) {
  try {
    const settings = getSettings_();

    return jsonResponse_({
      ok: true,
      service: 'A-Z Word Bank Bridge',
      version: BRIDGE_VERSION,
      status: 'ready',
      maxAccessibleDay: settings.maxAccessibleDay
    });
  } catch (error) {
    return errorResponse_(error);
  }
}

function doPost(e) {
  try {
    const settings = getSettings_();
    const request = parseRequest_(e);

    if (!safeEquals_(request.token, settings.accessToken)) {
      throw new Error('Unauthorized request.');
    }

    const action = cleanText_(request.action).toLowerCase();
    const spreadsheet = SpreadsheetApp.openById(settings.spreadsheetId);

    if (action === 'days') {
      return jsonResponse_({
        ok: true,
        action: 'days',
        days: getAccessibleDayNames_(
          spreadsheet,
          settings.maxAccessibleDay
        )
      });
    }

    const sheetName = cleanText_(request.sheet);
    const sheet = getAccessibleDaySheet_(
      spreadsheet,
      sheetName,
      settings.maxAccessibleDay
    );

    if (action === 'list') {
      return jsonResponse_({
        ok: true,
        action: 'list',
        sheet: sheetName,
        entries: getEntries_(sheet)
      });
    }

    if (
      action !== 'add' &&
      action !== 'update' &&
      action !== 'delete'
    ) {
      throw new Error('Unsupported action.');
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      let result;

      if (action === 'add') {
        const word = cleanText_(request.word);
        const definition = cleanText_(request.definition);

        validateEntry_(word, definition);

        const row = findNextEntryRow_(sheet);
        sheet.getRange(row, 1, 1, 2).setValues([[word, definition]]);

        result = {
          ok: true,
          action: 'add',
          sheet: sheetName,
          row: row,
          word: word
        };
      }

      if (action === 'update') {
        const row = validateRow_(request.row, sheet);
        const word = cleanText_(request.word);
        const definition = cleanText_(request.definition);

        validateEntry_(word, definition);
        ensureEntryExists_(sheet, row);
        sheet.getRange(row, 1, 1, 2).setValues([[word, definition]]);

        result = {
          ok: true,
          action: 'update',
          sheet: sheetName,
          row: row,
          word: word
        };
      }

      if (action === 'delete') {
        const row = validateRow_(request.row, sheet);

        ensureEntryExists_(sheet, row);
        sheet.getRange(row, 1, 1, 2).clearContent();

        result = {
          ok: true,
          action: 'delete',
          sheet: sheetName,
          row: row
        };
      }

      SpreadsheetApp.flush();
      rebuildWordBank_(spreadsheet, settings.maxAccessibleDay);

      return jsonResponse_(result);
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return errorResponse_(error);
  }
}

function getSettings_() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
  const accessToken = properties.getProperty('ACCESS_TOKEN');

  if (!spreadsheetId || !accessToken) {
    throw new Error(
      'Bridge setup is incomplete. Run setupBridge first.'
    );
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const adminSheet = spreadsheet.getSheetByName('Admin Settings');

  if (!adminSheet) {
    throw new Error('The Admin Settings sheet was not found.');
  }

  const maxAccessibleDay = Number(
    adminSheet.getRange('B1').getValue()
  );

  if (
    !Number.isInteger(maxAccessibleDay) ||
    maxAccessibleDay < 1 ||
    maxAccessibleDay > 100
  ) {
    throw new Error(
      'Admin Settings!B1 must contain a whole number from 1 to 100.'
    );
  }

  return {
    spreadsheetId: spreadsheetId,
    accessToken: accessToken,
    maxAccessibleDay: maxAccessibleDay
  };
}

function parseRequest_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('No request data was received.');
  }

  try {
    return JSON.parse(e.postData.contents);
  } catch (error) {
    throw new Error('The request data was not valid JSON.');
  }
}

function getAccessibleDayNames_(spreadsheet, maxDay) {
  return spreadsheet
    .getSheets()
    .map(function(sheet) {
      return sheet.getName();
    })
    .filter(function(name) {
      const match = name.match(/^Day\s+([1-9]\d*)$/i);
      return match && Number(match[1]) <= maxDay;
    })
    .sort(function(a, b) {
      return getDayNumber_(a) - getDayNumber_(b);
    });
}

function getAccessibleDaySheet_(spreadsheet, sheetName, maxDay) {
  const match = sheetName.match(/^Day\s+([1-9]\d*)$/i);

  if (!match) {
    throw new Error('Entries may be managed only in Day sheets.');
  }

  if (Number(match[1]) > maxDay) {
    throw new Error(
      'This Day has not been released by the administrator.'
    );
  }

  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('The selected Day sheet was not found.');
  }

  return sheet;
}

function getDayNumber_(sheetName) {
  const match = sheetName.match(/^Day\s+([1-9]\d*)$/i);
  return match ? Number(match[1]) : 0;
}

function getEntries_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 2)
    .getDisplayValues();

  const entries = [];

  values.forEach(function(row, index) {
    const word = cleanText_(row[0]);
    const definition = cleanText_(row[1]);

    if (word || definition) {
      entries.push({
        row: index + 2,
        word: word,
        definition: definition
      });
    }
  });

  return entries;
}

function findNextEntryRow_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);

  if (lastRow >= 2) {
    const values = sheet
      .getRange(2, 1, lastRow - 1, 2)
      .getDisplayValues();

    for (let i = 0; i < values.length; i++) {
      const word = cleanText_(values[i][0]);
      const definition = cleanText_(values[i][1]);

      if (!word && !definition) {
        return i + 2;
      }
    }
  }

  return Math.max(lastRow + 1, 2);
}

function validateRow_(value, sheet) {
  const row = Number(value);

  if (
    !Number.isInteger(row) ||
    row < 2 ||
    row > sheet.getLastRow()
  ) {
    throw new Error('The selected vocabulary entry is invalid.');
  }

  return row;
}

function ensureEntryExists_(sheet, row) {
  const values = sheet
    .getRange(row, 1, 1, 2)
    .getDisplayValues()[0];

  if (!cleanText_(values[0]) && !cleanText_(values[1])) {
    throw new Error(
      'That vocabulary entry no longer exists. Refresh the list.'
    );
  }
}

function validateEntry_(word, definition) {
  if (!word) {
    throw new Error('Please enter a word or phrase.');
  }

  if (!definition) {
    throw new Error('Please enter a definition.');
  }

  if (word.length > 250) {
    throw new Error('The word or phrase is too long.');
  }

  if (definition.length > 2000) {
    throw new Error('The definition is too long.');
  }
}

function rebuildWordBank_(spreadsheet, maxDay) {
  const target = spreadsheet.getSheetByName('A-Z Word Bank');

  if (!target) {
    throw new Error('The A-Z Word Bank sheet was not found.');
  }

  if (target.getMaxRows() > 1) {
    target
      .getRange(2, 1, target.getMaxRows() - 1, 3)
      .clearContent();
  }

  const data = [];

  getAccessibleDayNames_(spreadsheet, maxDay)
    .forEach(function(sheetName) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      const lastRow = sheet.getLastRow();

      if (lastRow < 2) {
        return;
      }

      const values = sheet
        .getRange(2, 1, lastRow - 1, 2)
        .getValues();

      values.forEach(function(row) {
        const word = cleanText_(row[0]);

        if (word) {
          data.push([word, row[1] || '', sheetName]);
        }
      });
    });

  data.sort(function(a, b) {
    return a[0].localeCompare(b[0]);
  });

  if (data.length) {
    target.getRange(2, 1, data.length, 3).setValues(data);
  }

  SpreadsheetApp.flush();
}

function cleanText_(value) {
  return value == null ? '' : String(value).trim();
}

function safeEquals_(left, right) {
  const a = String(left || '');
  const b = String(right || '');
  let difference = a.length ^ b.length;
  const length = Math.max(a.length, b.length);

  for (let i = 0; i < length; i++) {
    difference |=
      (a.charCodeAt(i) || 0) ^
      (b.charCodeAt(i) || 0);
  }

  return difference === 0;
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorResponse_(error) {
  return jsonResponse_({
    ok: false,
    error: error && error.message
      ? error.message
      : String(error)
  });
}

/** Read-only test: lists the Days currently released. */
function testListAccessibleDays() {
  const settings = getSettings_();
  const testEvent = {
    postData: {
      contents: JSON.stringify({
        action: 'days',
        token: settings.accessToken
      })
    }
  };

  console.log(doPost(testEvent).getContent());
}
function rotateAccessToken() {
  const newToken =
    Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '');

  PropertiesService
    .getScriptProperties()
    .setProperty('ACCESS_TOKEN', newToken);

  console.log('Access token rotated successfully.');
  console.log('NEW PERSONAL ACCESS LINK: ' + buildPersonalAccessLink_(newToken));
  return newToken;
}
