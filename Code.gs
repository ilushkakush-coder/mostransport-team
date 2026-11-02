const QUESTION_ROWS = [
  ['Q01', 'Остаться в офисе после рабочего дня?'],
  ['Q02', 'Работать рядом с домом?'],
  ['Q03', 'Начинать смену в семь утра?'],
  ['Q04', 'Получать зарплату два раза в месяц?'],
  ['Q05', 'Выйти на смену в субботу?'],
  ['Q06', 'Пройти собеседование по видеосвязи?'],
  ['Q07', 'Носить форму на работе?'],
  ['Q08', 'Учиться новому прямо на работе?'],
  ['Q09', 'Общаться с людьми весь день?'],
  ['Q10', 'Выбрать гибкий график?'],
  ['Q11', 'Работать в команде друзей?'],
  ['Q12', 'Получить первую работу без опыта?'],
  ['Q13', 'Отвечать в рабочих чатах вечером?'],
  ['Q14', 'Поработать на городском событии?'],
  ['Q15', 'Работать на свежем воздухе?'],
  ['Q16', 'Совмещать учёбу и работу?'],
  ['Q17', 'Снимать короткие видео для бренда?'],
  ['Q18', 'Пройти оплачиваемую стажировку?'],
  ['Q19', 'Попробовать себя в профессии водителя?'],
  ['Q20', 'Это новая реклама профессий Московского транспорта?']
];

const PROP_SHEET_ID = 'RESULTS_SPREADSHEET_ID';
const PROP_BOT_TOKEN = 'TELEGRAM_BOT_TOKEN';

/** Однократно создаёт таблицу: вкладки Sessions, Responses и Summary. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty(PROP_SHEET_ID);
  let spreadsheet;
  if (spreadsheetId) {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } else {
    spreadsheet = SpreadsheetApp.create('Работа для зумеров — результаты свайпов');
    spreadsheetId = spreadsheet.getId();
    props.setProperty(PROP_SHEET_ID, spreadsheetId);
  }

  const sessions = getOrCreateSheet_(spreadsheet, 'Sessions', ['Открыто', 'Telegram ID', 'ID сессии']);
  const responses = getOrCreateSheet_(spreadsheet, 'Responses', ['Время ответа', 'Telegram ID', 'ID сессии', 'ID вопроса', 'Вопрос', 'Ответ']);
  const summary = getOrCreateSheet_(spreadsheet, 'Summary', ['ID вопроса', 'Вопрос', 'Норм', 'Стрём']);

  sessions.setFrozenRows(1);
  sessions.getRange('B:B').setNumberFormat('@');
  responses.setFrozenRows(1);
  responses.getRange('B:B').setNumberFormat('@');

  summary.getRange(2, 1, QUESTION_ROWS.length, 2).setValues(QUESTION_ROWS);
  const countFormulas = QUESTION_ROWS.map((_, index) => {
    const row = index + 2;
    return [
      `=COUNTIFS(Responses!$D:$D,$A${row},Responses!$F:$F,"норм")`,
      `=COUNTIFS(Responses!$D:$D,$A${row},Responses!$F:$F,"стрём")`
    ];
  });
  summary.getRange(2, 3, countFormulas.length, 2).setFormulas(countFormulas);
  summary.setFrozenRows(1);
  [sessions, responses, summary].forEach(sheet => {
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold').setBackground('#ffd52e');
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  });
  Logger.log('Таблица создана: ' + spreadsheet.getUrl());
  return spreadsheet.getUrl();
}

/** Веб-приложение принимает события открытия Mini App и свайпы. */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const scriptProperties = PropertiesService.getScriptProperties();
    const botToken = scriptProperties.getProperty(PROP_BOT_TOKEN) || scriptProperties.getProperty('`TELEGRAM_BOT_TOKEN`');
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID);
    if (!botToken) throw new Error('Не задан Script Property TELEGRAM_BOT_TOKEN');
    if (!spreadsheetId) throw new Error('Сначала запустите setup()');

    const telegramId = verifyTelegramInitData_(body.initData, botToken);
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sessionId = String(body.sessionId || '');
    if (!sessionId) throw new Error('Нет ID сессии');

    if (body.event === 'open') {
      appendLocked_(spreadsheet.getSheetByName('Sessions'), [new Date(), telegramId, sessionId]);
      return json_({ ok: true });
    }

    if (body.event === 'swipe') {
      const question = QUESTION_ROWS.find(row => row[0] === body.questionId);
      const answer = String(body.answer || '').toLowerCase();
      if (!question) throw new Error('Неизвестный вопрос');
      if (answer !== 'норм' && answer !== 'стрём') throw new Error('Ответ должен быть «норм» или «стрём»');
      appendLocked_(spreadsheet.getSheetByName('Responses'), [new Date(), telegramId, sessionId, question[0], question[1], answer]);
      return json_({ ok: true });
    }

    throw new Error('Неизвестный тип события');
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    return json_({ ok: false, error: String(error.message || error) });
  }
}

function doGet() {
  return json_({ ok: true, service: 'HR Swipe Apps Script endpoint' });
}

function getOrCreateSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function appendLocked_(sheet, values) {
  if (!sheet) throw new Error('Не найдена вкладка таблицы');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet.appendRow(values);
  } finally {
    lock.releaseLock();
  }
}

/** Проверяет подпись Telegram и извлекает Telegram ID только из initData. */
function verifyTelegramInitData_(initData, botToken) {
  if (!initData) throw new Error('Откройте приложение через Telegram Mini App');
  const params = parseInitData_(String(initData));
  const suppliedHash = params.hash;
  if (!suppliedHash) throw new Error('В initData нет подписи Telegram');

  const secretKey = Utilities.computeHmacSha256Signature('WebAppData', botToken);
  const entries = Object.keys(params).filter(key => key !== 'hash').sort();
  const fullDataCheck = entries.map(key => `${key}=${params[key]}`).join('\n');
  // Поддерживаем initData как без signature, так и с отдельной подписью для third-party validation.
  const candidates = [fullDataCheck];
  if (Object.prototype.hasOwnProperty.call(params, 'signature')) {
    candidates.push(entries.filter(key => key !== 'signature').map(key => `${key}=${params[key]}`).join('\n'));
  }
  const valid = candidates.some(dataCheck => {
    const calculatedHash = toHex_(Utilities.computeHmacSha256Signature(dataCheck, secretKey));
    return constantTimeEquals_(calculatedHash, suppliedHash.toLowerCase());
  });
  if (!valid) throw new Error('Подпись Telegram не прошла проверку');

  const authDate = Number(params.auth_date);
  const now = Math.floor(Date.now() / 1000);
  if (!authDate || now - authDate > 86400 || authDate - now > 60) {
    throw new Error('Данные Telegram устарели; переоткройте Mini App');
  }
  const user = JSON.parse(params.user || '{}');
  if (!user.id) throw new Error('В initData не найден Telegram ID');
  return String(user.id);
}

function parseInitData_(raw) {
  const result = {};
  raw.split('&').forEach(part => {
    if (!part) return;
    const separator = part.indexOf('=');
    const rawKey = separator < 0 ? part : part.slice(0, separator);
    const rawValue = separator < 0 ? '' : part.slice(separator + 1);
    const key = decodeURIComponent(rawKey.replace(/\+/g, ' '));
    const value = decodeURIComponent(rawValue.replace(/\+/g, ' '));
    result[key] = value;
  });
  return result;
}

function toHex_(bytes) {
  return bytes.map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2)).join('');
}

function constantTimeEquals_(left, right) {
  if (left.length !== right.length) return false;
  let mismatch = 0;
  for (let i = 0; i < left.length; i++) mismatch |= left.charCodeAt(i) ^ right.charCodeAt(i);
  return mismatch === 0;
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
