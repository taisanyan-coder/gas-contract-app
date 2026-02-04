// ====== 設定 ======
const SPREADSHEET_ID = '1kL0QaO2ADETdbWvDOMj0xEZROiwq9evs23Rh7mXEcaw';
const SHEET_CONTRACTS = 'Contracts';
const SHEET_LINKS = 'Links';
const SHEET_TEMPLATES = 'Templates';

// ====== Webアプリ入口 ======
function doGet(e) {
  const format = (e && e.parameter && e.parameter.format) ? String(e.parameter.format) : '';
  const mode = format.toLowerCase();

  const page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page) : '';

  // JSONモード（デバッグ用）
  // ?format=json        -> Contracts 全件（headers+rows）
  // ?format=json1       -> Contracts 先頭1件
  // ?format=links       -> Links（enabledのみ）
  // ?format=templates   -> Templates（enabledのみ）
  if (mode === 'json' || mode === 'json1' || mode === 'links' || mode === 'templates') {

    if (mode === 'links') {
      const linksData = getLinksData();
      return ContentService
        .createTextOutput(JSON.stringify(linksData, null, 2))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (mode === 'templates') {
      const tmplData = getTemplatesData();
      return ContentService
        .createTextOutput(JSON.stringify(tmplData, null, 2))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = getContractsData();
    const payload = (mode === 'json1')
      ? {
          headers: data.headers,
          row0: (data.rows && data.rows.length > 0) ? data.rows[0] : [],
          error: data.error || ''
        }
      : data;

    return ContentService
      .createTextOutput(JSON.stringify(payload, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // HTMLページ切替（一覧 / 詳細）
  if (page === 'detail') {
    const t = HtmlService.createTemplateFromFile('detail');
    t.type = (e && e.parameter && e.parameter.type) ? String(e.parameter.type) : 'contract';
    return t.evaluate().setTitle('契約詳細');
  }

  // 通常は一覧
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Contracts 一覧');
}


// ====== Contracts データ取得（Dateを必ず文字列に変換して返す） ======
function getContractsData() {
  try {
    if (!SPREADSHEET_ID) {
      return { headers: [], rows: [], error: 'SPREADSHEET_ID が未設定です' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_CONTRACTS);
    if (!sheet) {
      return { headers: [], rows: [], error: `Sheet not found: ${SHEET_CONTRACTS}` };
    }

    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 1) {
      return { headers: [], rows: [], error: 'No data' };
    }

    const tz = Session.getScriptTimeZone();
    const toSafe = (v) => {
      if (v === null || v === undefined) return '';
      if (v instanceof Date) return Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ss");
      if (typeof v === 'object') return String(v);
      return v; // string/number/boolean
    };

    const headers = (values[0] || []).map(v => String(v ?? '').trim());
    const rows = values.slice(1).map(r => (r || []).map(toSafe));

    return { headers, rows, error: '' };
  } catch (e) {
    return { headers: [], rows: [], error: (e && e.message) ? e.message : String(e) };
  }
}


/**
 * 新規契約を1行追加する
 * @param {string} staffName
 * @param {string} contractEnd  // yyyy-MM-dd
 * @param {string} templateType
 */
function addContract(staffName, contractEnd, templateType) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONTRACTS);
  if (!sheet) throw new Error('Contracts シートが見つかりません');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

  const row = Array(headers.length).fill('');

  row[idx['契約ID']] = Utilities.getUuid();
  row[idx['契約ステータス']] = 'ACTIVE';
  row[idx['単価UP実現']] = '未';

  row[idx['スタッフ名']] = staffName || '';
  row[idx['契約終了日']] = contractEnd || '';
  row[idx['テンプレ種別']] = templateType || '';

  sheet.appendRow(row);
  return { ok: true };
}


// ====== Links 取得（よく使うリンク用） ======
// 列想定：label / url / order / enabled
function getLinksData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_LINKS);
    if (!sheet) return { links: [], error: `Sheet not found: ${SHEET_LINKS}` };

    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return { links: [], error: '' };

    const headers = values[0].map(v => String(v ?? '').trim());
    const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

    const links = values.slice(1)
      .filter(r => {
        const enabled = r[idx['enabled']];
        return enabled === true || String(enabled).toLowerCase() === 'true' || String(enabled) === '1';
      })
      .map(r => ({
        label: String(r[idx['label']] ?? '').trim(),
        url: String(r[idx['url']] ?? '').trim(),
        order: Number(r[idx['order']] ?? 9999)
      }))
      .filter(x => x.label && x.url)
      .sort((a, b) => (a.order - b.order));

    return { links, error: '' };
  } catch (e) {
    return { links: [], error: (e && e.message) ? e.message : String(e) };
  }
}


// ====== Templates 取得（テンプレ本文マスタ用） ======
// 列想定：template_key / template_type / label / body / enabled / order
function getTemplatesData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_TEMPLATES);
    if (!sheet) return { templates: [], error: `Sheet not found: ${SHEET_TEMPLATES}` };

    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return { templates: [], error: '' };

    const headers = values[0].map(v => String(v ?? '').trim());
    const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

    const templates = values.slice(1)
      .filter(r => {
        const enabled = r[idx['enabled']];
        return enabled === true || String(enabled).toLowerCase() === 'true' || String(enabled) === '1';
      })
      .map(r => ({
        template_key: String(r[idx['template_key']] ?? '').trim(),
        template_type: String(r[idx['template_type']] ?? '').trim(), // 契約更新/単価変更/情報変更
        label: String(r[idx['label']] ?? '').trim(),
        body: String(r[idx['body']] ?? ''),
        order: Number(r[idx['order']] ?? 9999)
      }))
      .filter(x => x.template_key && x.template_type && x.label)
      .sort((a, b) => {
        // type → order → label
        if (a.template_type !== b.template_type) return a.template_type.localeCompare(b.template_type, 'ja');
        if (a.order !== b.order) return a.order - b.order;
        return a.label.localeCompare(b.label, 'ja');
      });

    return { templates, error: '' };
  } catch (e) {
    return { templates: [], error: (e && e.message) ? e.message : String(e) };
  }
}
