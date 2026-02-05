// ====== 設定 ======
const SPREADSHEET_ID = '1kL0QaO2ADETdbWvDOMj0xEZROiwq9evs23Rh7mXEcaw';
const SHEET_CONTRACTS = 'Contracts';
const SHEET_LINKS = 'Links';
const SHEET_TEMPLATES = 'Templates';
const TZ = 'Asia/Tokyo';

// 単価UP実現の許容値
const UNIT_UP_ALLOWED = new Set(['未', '〇', '×']);

// ====== Webアプリ入口 ======
function doGet(e) {
  const format = (e && e.parameter && e.parameter.format) ? String(e.parameter.format) : '';
  const mode = format.toLowerCase();
  const page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page) : '';

  // JSON系（リンク/テンプレ含む）
  if (mode === 'json' || mode === 'json1' || mode === 'links' || mode === 'templates') {
    if (mode === 'links') {
      return ContentService.createTextOutput(JSON.stringify(getLinksData(), null, 2))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (mode === 'templates') {
      return ContentService.createTextOutput(JSON.stringify(getTemplatesData(), null, 2))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = getContractsData();
    const payload = (mode === 'json1')
      ? { headers: data.headers, row0: data.rows[0] || [], error: data.error || '' }
      : data;

    return ContentService.createTextOutput(JSON.stringify(payload, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 詳細ページ（既存運用を壊さない）
  if (page === 'detail') {
    const t = HtmlService.createTemplateFromFile('detail');
    t.type = (e && e.parameter && e.parameter.type) ? String(e.parameter.type) : 'contract';
    return t.evaluate().setTitle('契約詳細');
  }

  // 一覧
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Contracts 一覧');
}

// ====== Contracts 取得（完了フラグOFFのみ表示） ======
function getContractsData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_CONTRACTS);
    if (!sheet) return { headers: [], rows: [], rowNumbers: [], error: 'Contracts not found' };

    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return { headers: [], rows: [], rowNumbers: [], error: '' };

    const headers = values[0].map(v => String(v ?? '').trim());
    const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

    // 必須列（手動整理方式）
    const required = ['スタッフ名', '契約終了日', '契約ステータス', 'テンプレ種別', '単価UP実現', '登録日時', '最終更新日', '完了フラグ'];
    const missing = required.filter(h => idx[h] === undefined);
    if (missing.length) {
      return { headers, rows: [], rowNumbers: [], error: '必須列がありません: ' + missing.join(', ') };
    }

    // 日付表示：契約開始日/終了日は「日付だけ」、登録日時/最終更新日は「日時」
    const DATE_ONLY_HEADERS = new Set(['契約開始日', '契約終了日']);
    const DATETIME_HEADERS = new Set(['登録日時', '最終更新日']);

    const toSafeByHeader = (header, v) => {
      if (v === null || v === undefined) return '';
      if (v instanceof Date) {
        if (DATE_ONLY_HEADERS.has(header)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
        if (DATETIME_HEADERS.has(header)) return Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm:ss");
        return Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm:ss");
      }
      return v;
    };

    const rows = [];
    const rowNumbers = [];

    for (let i = 1; i < values.length; i++) {
      const r = values[i];

      // 完了フラグが TRUE の行は表示しない
      const doneVal = r[idx['完了フラグ']];
      const isDone = (doneVal === true) || (String(doneVal).toLowerCase() === 'true');
      if (isDone) continue;

      // 単価UP実現が空なら「未」に寄せる（表示・運用の安定化）
      const uIdx = idx['単価UP実現'];
      const cur = String(r[uIdx] ?? '').trim();
      if (!cur) r[uIdx] = '未';

      rows.push(headers.map((h, j) => toSafeByHeader(h, r[j])));
      rowNumbers.push(i + 1); // シート上の行番号（1行目ヘッダなので +1）
    }

    return { headers, rows, rowNumbers, error: '' };
  } catch (e) {
    return { headers: [], rows: [], rowNumbers: [], error: e.message };
  }
}

// ====== 新規登録（登録日時・最終更新日・完了フラグを自動セット） ======
// staffName: スタッフ名
// contractEnd: 'YYYY-MM-DD'
// templateType: テンプレ種別
function addContract(staffName, contractEnd, templateType) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONTRACTS);
  if (!sheet) throw new Error('Contracts シートが見つかりません');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(v => String(v ?? '').trim());
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));
  const row = Array(headers.length).fill('');

  const required = ['契約ID', '契約ステータス', '単価UP実現', 'スタッフ名', '契約終了日', 'テンプレ種別', '登録日時', '最終更新日', '完了フラグ'];
  const missing = required.filter(h => idx[h] === undefined);
  if (missing.length) throw new Error('Contracts に必須列がありません: ' + missing.join(', '));

  if (!String(staffName || '').trim()) throw new Error('スタッフ名が空です');
  if (!String(contractEnd || '').trim()) throw new Error('契約終了日が空です');

  const now = new Date();

  row[idx['契約ID']] = Utilities.getUuid();
  row[idx['契約ステータス']] = 'ACTIVE';
  row[idx['単価UP実現']] = '未';
  row[idx['スタッフ名']] = String(staffName || '').trim();

  // UTCズレ防止：ローカル日付(0:00)
  row[idx['契約終了日']] = parseDateYMD_(contractEnd);

  row[idx['テンプレ種別']] = String(templateType || '');

  // 管理列
  row[idx['登録日時']] = now;
  row[idx['最終更新日']] = now;
  row[idx['完了フラグ']] = false;

  sheet.appendRow(row);
  return { ok: true };
}

// ====== 単価UP実現を更新（プルダウン保存用） ======
// rowNumber: シートの行番号（2以上）
// value: '未' | '〇' | '×'
function setUnitUpByRow(rowNumber, value) {
  const rn = Number(rowNumber);
  if (!rn || rn < 2) throw new Error('rowNumber が不正です');

  const v = String(value ?? '').trim();
  if (!UNIT_UP_ALLOWED.has(v)) throw new Error('単価UP実現の値が不正です（未/〇/×）: ' + v);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONTRACTS);
  if (!sheet) throw new Error('Contracts シートが見つかりません');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(x => String(x ?? '').trim());
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

  const need = ['単価UP実現', '最終更新日'];
  const missing = need.filter(h => idx[h] === undefined);
  if (missing.length) throw new Error('Contracts に必須列がありません: ' + missing.join(', '));

  // 完了済みは更新しない（事故防止）
  if (idx['完了フラグ'] !== undefined) {
    const doneVal = sheet.getRange(rn, idx['完了フラグ'] + 1).getValue();
    const isDone = (doneVal === true) || (String(doneVal).toLowerCase() === 'true');
    if (isDone) throw new Error('この行は完了済みのため更新できません');
  }

  sheet.getRange(rn, idx['単価UP実現'] + 1).setValue(v);
  sheet.getRange(rn, idx['最終更新日'] + 1).setValue(new Date());

  return { ok: true };
}

// ====== 完了（= 非表示）処理：完了フラグをTRUEにする（スプシは保持） ======
function completeContractByRow(rowNumber) {
  const rn = Number(rowNumber);
  if (!rn || rn < 2) throw new Error('rowNumber が不正です');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONTRACTS);
  if (!sheet) throw new Error('Contracts シートが見つかりません');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(v => String(v ?? '').trim());
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

  const need = ['完了フラグ', '最終更新日'];
  const missing = need.filter(h => idx[h] === undefined);
  if (missing.length) throw new Error('Contracts に必須列がありません: ' + missing.join(', '));

  sheet.getRange(rn, idx['完了フラグ'] + 1).setValue(true);
  sheet.getRange(rn, idx['最終更新日'] + 1).setValue(new Date());

  return { ok: true };
}

// 複数行まとめて完了（rowNumbers: number[]）
function completeContractsByRows(rowNumbers) {
  if (!Array.isArray(rowNumbers) || rowNumbers.length === 0) return { ok: true, done: 0 };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONTRACTS);
  if (!sheet) throw new Error('Contracts シートが見つかりません');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(v => String(v ?? '').trim());
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

  const need = ['完了フラグ', '最終更新日'];
  const missing = need.filter(h => idx[h] === undefined);
  if (missing.length) throw new Error('Contracts に必須列がありません: ' + missing.join(', '));

  const doneCol = idx['完了フラグ'] + 1;
  const updatedCol = idx['最終更新日'] + 1;
  const now = new Date();

  rowNumbers
    .map(n => Number(n))
    .filter(n => n && n >= 2)
    .forEach(rn => {
      sheet.getRange(rn, doneCol).setValue(true);
      sheet.getRange(rn, updatedCol).setValue(now);
    });

  return { ok: true, done: rowNumbers.length };
}

// ====== Links ======
function getLinksData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_LINKS);
  if (!sheet) return { links: [] };

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return { links: [] };

  const headers = values[0].map(v => String(v ?? '').trim());
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

  const need = ['enabled', 'label', 'url', 'order'];
  const missing = need.filter(h => idx[h] === undefined);
  if (missing.length) return { links: [], error: 'Links に列がありません: ' + missing.join(', ') };

  const links = values.slice(1)
    .filter(r => String(r[idx['enabled']] ?? '').toLowerCase() === 'true')
    .map(r => ({
      label: r[idx['label']],
      url: r[idx['url']],
      order: Number(r[idx['order']] || 999)
    }))
    .sort((a, b) => a.order - b.order);

  return { links };
}

// ====== Templates ======
function getTemplatesData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TEMPLATES);
  if (!sheet) return { templates: [] };

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return { templates: [] };

  const headers = values[0].map(v => String(v ?? '').trim());
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

  const need = ['enabled', 'template_key', 'template_type', 'label', 'body', 'order'];
  const missing = need.filter(h => idx[h] === undefined);
  if (missing.length) return { templates: [], error: 'Templates に列がありません: ' + missing.join(', ') };

  const templates = values.slice(1)
    .filter(r => String(r[idx['enabled']] ?? '').toLowerCase() === 'true')
    .map(r => ({
      template_key: r[idx['template_key']],
      template_type: r[idx['template_type']],
      label: r[idx['label']],
      body: r[idx['body']],
      order: Number(r[idx['order']] || 999)
    }))
    .sort((a, b) => a.order - b.order);

  return { templates };
}

// ====== 日付ユーティリティ ======

// 'YYYY-MM-DD' をローカル日付(0:00)の Date にする（UTCズレ防止）
function parseDateYMD_(ymd) {
  const s = String(ymd || '').trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('契約終了日の形式が不正です（YYYY-MM-DD）: ' + s);
  const y = Number(m[1]);
  const mo = Number(m[2]) - 1;
  const d = Number(m[3]);
  return new Date(y, mo, d, 0, 0, 0);
}

// 既存値（Date or 文字列）を可能なら Date にする
function parseDateAny_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const s = String(v).trim();
  if (!s) return null;

  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return parseDateYMD_(s);

  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

// ====== 単価UP反映日を更新（カレンダー選択用） ======
// rowNumber: シートの行番号（2以上）
// ymd: 'YYYY-MM-DD'
function setUnitUpReflectDateByRow(rowNumber, ymd) {
  const rn = Number(rowNumber);
  if (!rn || rn < 2) throw new Error('rowNumber が不正です');

  const s = String(ymd ?? '').trim();
  if (!s) throw new Error('単価UP反映日が空です');
  const reflectDate = parseDateYMD_(s); // 0:00固定（UTCズレ防止）

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONTRACTS);
  if (!sheet) throw new Error('Contracts シートが見つかりません');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(v => String(v ?? '').trim());

  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

  const need = ['単価UP実現', '単価UP反映日', '最終更新日'];
  const missing = need.filter(h => idx[h] === undefined);
  if (missing.length) throw new Error('Contracts に必須列がありません: ' + missing.join(', '));

  // 完了済みは更新しない（事故防止）
  if (idx['完了フラグ'] !== undefined) {
    const doneVal = sheet.getRange(rn, idx['完了フラグ'] + 1).getValue();
    const isDone = (doneVal === true) || (String(doneVal).toLowerCase() === 'true');
    if (isDone) throw new Error('この行は完了済みのため更新できません');
  }

  // 単価UP実現が「〇」のときだけ反映日を保存（事故防止）
  const unitUp = String(sheet.getRange(rn, idx['単価UP実現'] + 1).getValue() ?? '').trim();
  if (unitUp !== '〇') throw new Error('単価UP反映日は、単価UP実現が「〇」の行だけ設定できます');

  sheet.getRange(rn, idx['単価UP反映日'] + 1).setValue(reflectDate);
  sheet.getRange(rn, idx['最終更新日'] + 1).setValue(new Date());

  return { ok: true };
}

// ====== 単価UP反映日＋単価を更新（3引数版） ======
// rowNumber: シートの行番号（2以上）
// reflectYmd: 'YYYY-MM-DD'
// unitValue: number（文字列でも可）
function setUnitUpReflectAndUnitByRow(rowNumber, reflectYmd, unitValue) {
  const rn = Number(rowNumber);
  if (!rn || rn < 2) throw new Error('rowNumber が不正です');

  const ymd = String(reflectYmd ?? '').trim();
  if (!ymd) throw new Error('反映日が空です');
  const reflectDate = parseDateYMD_(ymd); // 0:00固定（UTCズレ防止）

  // 数値化（カンマ・円マーク・空白は除去）
  const raw = String(unitValue ?? '').trim();
  const n = Number(raw.replace(/[¥,，\s]/g, ''));
  if (!isFinite(n) || n <= 0) throw new Error('単価が不正です（正の数）: ' + raw);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONTRACTS);
  if (!sheet) throw new Error('Contracts シートが見つかりません');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(v => String(v ?? '').trim());
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

  // 必須列（※単価区分は使わない）
  const need = ['単価UP実現', '最終更新日', '完了フラグ', '単価UP反映日', '単価'];
  const missing = need.filter(h => idx[h] === undefined);
  if (missing.length) {
    throw new Error('Contracts に必須列がありません: ' + missing.join(', ') + '（1行目ヘッダを確認してね）');
  }

  // 完了済みは更新しない（事故防止）
  const doneVal = sheet.getRange(rn, idx['完了フラグ'] + 1).getValue();
  const isDone = (doneVal === true) || (String(doneVal).toLowerCase() === 'true');
  if (isDone) throw new Error('この行は完了済みのため更新できません');

  // 単価UP実現が〇以外なら保存させない
  const curUnitUp = String(sheet.getRange(rn, idx['単価UP実現'] + 1).getValue() ?? '').trim();
  if (curUnitUp !== '〇') throw new Error('単価UP実現が「〇」の行だけ保存できます（現在: ' + (curUnitUp || '未') + '）');

  // 保存
  sheet.getRange(rn, idx['単価UP反映日'] + 1).setValue(reflectDate);

  const unitCell = sheet.getRange(rn, idx['単価'] + 1);
  unitCell.setNumberFormat('#,##0');  // ★ここが肝：日付化を潰して数値表示に固定
  unitCell.setValue(n);

  sheet.getRange(rn, idx['最終更新日'] + 1).setValue(new Date());

  return { ok: true };
}

