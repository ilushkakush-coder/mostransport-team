const QUESTION_ROWS = [
  ['V2-01', 'Получать 600 рублей каждый день на завтраки и обеды'],
  ['V2-02', 'Быстро расти по карьере, даже если это временно мешает личной жизни'],
  ['V2-03', 'Присоединиться к корпоративным творческим комьюнити: петь в хоре, служить в театре'],
  ['V2-04', 'Работать по ночам'],
  ['V2-05', 'Понимать задачи и условия для повышения — то есть строить прозрачный карьерный трек'],
  ['V2-06', 'Бесплатно ездить на общественном транспорте'],
  ['V2-07', 'Участвовать в профессиональных конкурсах среди коллег из компании и со всей страны'],
  ['V2-08', 'Работать по фиксированному графику без переработок'],
  ['V2-09', 'Получать юридическую и финансовую помощь, путевки и подарки от профсоюза'],
  ['V2-10', 'Сменить специальность внутри компании без карьерных потерь'],
  ['V2-11', 'Работать у дома'],
  ['V2-12', 'Возможность стать амбассадором компании: представлять бренд работодателя на событиях и в соцсетях'],
  ['V2-13', 'Работать сменами 1/1, 2/2 или 3/3 по 12–24 часа'],
  ['V2-14', 'Устроиться на работу, где не нужно образование и опыт'],
  ['V2-15', 'Заниматься спортом и участвовать в соревнованиях за счет работодателя'],
  ['V2-16', 'Собираться на планерках и писать начальнику напрямую по рабочим и нерабочим вопросам'],
  ['V2-17', 'Работать в стабильной компании много лет, иногда выполняя скучные задачи'],
  ['V2-18', 'Отдыхать в официальном отпуске до 48 дней в год'],
  ['V2-19', 'Работать в тоннеле метро'],
  ['V2-20', 'Принимать решение об устройстве в компанию по отзывам в интернете'],
  ['V2-21', 'Работать на физически тяжелой грязной должности с высокой зарплатой и быстрым карьерным ростом'],
  ['V2-22', 'Присоединиться к корпоративному комьюнити молодых сотрудников до 35 лет'],
  ['V2-23', 'Стабильно зарабатывать каждый год на 1–5% больше за счет индексации'],
  ['V2-24', 'Подчиняться руководителю, который поднялся с рабочей специальности'],
  ['V2-25', 'Получить рабочую специальность за счет компании'],
  ['V2-26', 'Совмещать работу и учебу/хобби/бизнес'],
  ['V2-27', 'Понимать ценность своей работы и чувствовать командность'],
  ['V2-28', 'Работать по трудовому договору'],
  ['V2-29', 'Получать бонусы типа мерча, билетов, путевок за участие во внерабочих корпоративных событиях'],
  ['V2-30', 'Выйти на пенсию после 20–25 лет стажа']
];

const PROP_SHEET_ID = 'RESULTS_SPREADSHEET_ID';
const PROP_BOT_TOKEN = 'TELEGRAM_BOT_TOKEN';
const PROP_BOT_WEBHOOK_SECRET = 'TELEGRAM_WEBHOOK_SECRET';
const MINI_APP_URL = 'https://mostransport-team.vercel.app';

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
  const summary = getOrCreateSheet_(spreadsheet, 'Summary', ['ID вопроса', 'Вопрос', 'Нрав', 'Не нрав']);

  sessions.setFrozenRows(1);
  sessions.getRange('B:B').setNumberFormat('@');
  responses.setFrozenRows(1);
  responses.getRange('B:B').setNumberFormat('@');
  summary.getRange(1, 1, 1, 4).setValues([['ID вопроса', 'Вопрос', 'Нрав', 'Не нрав']]);

  const oldSummaryRows = Math.max(summary.getLastRow() - 1, QUESTION_ROWS.length);
  if (oldSummaryRows > 0) summary.getRange(2, 1, oldSummaryRows, 4).clearContent();
  summary.getRange(2, 1, QUESTION_ROWS.length, 2).setValues(QUESTION_ROWS);
  const countFormulas = QUESTION_ROWS.map((_, index) => {
    const row = index + 2;
    return [
      `=COUNTIFS(Responses!D:D;A${row};Responses!F:F;"нрав")+COUNTIFS(Responses!D:D;A${row};Responses!F:F;"норм")`,
      `=COUNTIFS(Responses!D:D;A${row};Responses!F:F;"не нрав")+COUNTIFS(Responses!D:D;A${row};Responses!F:F;"стрём")`
    ];
  });
  summary.getRange(2, 3, countFormulas.length, 2).setFormulas(countFormulas);
  summary.setFrozenRows(1);
  [sessions, responses, summary].forEach(sheet => {
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold').setBackground('#222222').setFontColor('#ffffff');
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  });
  Logger.log('Таблица создана: ' + spreadsheet.getUrl());
  return spreadsheet.getUrl();
}

/** Веб-приложение принимает события открытия Mini App и свайпы. */
function doPost(e) {
  // Telegram bot updates use a private webhook URL parameter; swipe events use the JSON body.
  if (e && e.parameter && e.parameter.hook) {
    return handleTelegramUpdate_(e);
  }

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
      if (answer !== 'нрав' && answer !== 'не нрав') throw new Error('Ответ должен быть «нрав» или «не нрав»');
      appendLocked_(spreadsheet.getSheetByName('Responses'), [new Date(), telegramId, sessionId, question[0], question[1], answer]);
      return json_({ ok: true });
    }

    throw new Error('Неизвестный тип события');
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    return json_({ ok: false, error: String(error.message || error) });
  }
}

/**
 * Run once from the Apps Script editor to register this Apps Script web app as
 * the Telegram bot webhook. It reuses the token already stored in Script Properties.
 */
function setupTelegramWebhook() {
  const props = PropertiesService.getScriptProperties();
  const botToken = props.getProperty(PROP_BOT_TOKEN) || props.getProperty('`TELEGRAM_BOT_TOKEN`');
  if (!botToken) throw new Error('Не задан Script Property TELEGRAM_BOT_TOKEN');

  let secret = props.getProperty(PROP_BOT_WEBHOOK_SECRET);
  if (!secret) {
    secret = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty(PROP_BOT_WEBHOOK_SECRET, secret);
  }

  const webAppUrl = ScriptApp.getService().getUrl();
  if (!webAppUrl) throw new Error('Сначала разверните Apps Script как веб-приложение');
  const response = UrlFetchApp.fetch(`https://api.telegram.org/bot${botToken}/setWebhook`, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ url: `${webAppUrl}?hook=${secret}`, allowed_updates: ['message'] }),
    muteHttpExceptions: true
  });
  const result = JSON.parse(response.getContentText());
  if (!result.ok) throw new Error(`Telegram не принял webhook: ${result.description || response.getResponseCode()}`);
  Logger.log('Webhook зарегистрирован для бота. URL и секрет не выводятся.');
  return 'Webhook зарегистрирован';
}

function handleTelegramUpdate_(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const expectedSecret = props.getProperty(PROP_BOT_WEBHOOK_SECRET);
    if (!expectedSecret || !constantTimeEquals_(String(e.parameter.hook), expectedSecret)) {
      return json_({ ok: false, error: 'Unauthorized' });
    }

    const update = JSON.parse((e.postData && e.postData.contents) || '{}');
    const message = update.message;
    if (message && message.chat && message.text && /^\/start(?:@\w+)?(?:\s|$)/i.test(message.text)) {
      const botToken = props.getProperty(PROP_BOT_TOKEN) || props.getProperty('`TELEGRAM_BOT_TOKEN`');
      if (!botToken) throw new Error('Не задан Script Property TELEGRAM_BOT_TOKEN');
      const text = [
        'хай 👋🏻 на связи админ команды транспорта',
        '',
        '🙏🏻 помоги написать ✨ <b>вакансию мечты</b> ✨ для зумеров:',
        'выбери предложения по душе',
        '',
        'нрав — свайп вправо ➡',
        'не нрав — свайп влево ⬅'
      ].join('\n');
      const payload = {
        chat_id: message.chat.id,
        text,
        parse_mode: 'HTML',
        reply_markup: {
          inline_keyboard: [[{
            text: 'Открыть приложение',
            web_app: { url: MINI_APP_URL }
          }]]
        }
      };
      const response = UrlFetchApp.fetch(`https://api.telegram.org/bot${botToken}/sendMessage`, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      const result = JSON.parse(response.getContentText());
      if (!result.ok) throw new Error(`Telegram sendMessage failed: ${result.description || response.getResponseCode()}`);
    }
    return json_({ ok: true });
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

  const secretKey = Utilities.computeHmacSha256Signature(botToken, 'WebAppData');
  const entries = Object.keys(params).filter(key => key !== 'hash').sort();
  const fullDataCheck = entries.map(key => `${key}=${params[key]}`).join('\n');
  // Поддерживаем initData как без signature, так и с отдельной подписью для third-party validation.
  const candidates = [fullDataCheck];
  if (Object.prototype.hasOwnProperty.call(params, 'signature')) {
    candidates.push(entries.filter(key => key !== 'signature').map(key => `${key}=${params[key]}`).join('\n'));
  }
  const valid = candidates.some(dataCheck => {
    const dataBytes = Utilities.newBlob(dataCheck, 'text/plain').getBytes();
    const calculatedHash = toHex_(Utilities.computeHmacSha256Signature(dataBytes, secretKey));
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
