/**
 * セガサミー チャリティープロアマ｜スコア速報ランキング API
 *
 * スプレッドシート内の配信シート（配信1〜配信4）のうち、
 * B1に何かしら入力されている（空欄でない）シートをJSONで返す。
 * チェックボックスのTRUEでも、"OK"「公開」などの文字列でもOK。空欄なら非公開扱い。
 * 複数入力されている場合は一番右（＝最新の配信シート）を優先する。
 *
 * ── シートのレイアウト（固定位置）──────────────
 *   A1: 公開        B1: TRUE または OK・公開など何か入力（空欄=非公開）
 *   A2: 見出し      B2: 前半スコア速報
 *   A3: 説明文      B3: 10:00時点のチーム戦最新順位をお届けします。
 *   A4: 更新時刻    B4: 7/23 10:40
 *   A5: 表示数      B5: 表示したい順位まで数値で入力（例：10）。空欄なら全組表示
 *   6行目: ヘッダー（順位｜チーム名｜ポイント）
 *   7行目〜: データ（A:順位 B:チーム名 C:ポイント）
 *     ・チーム名が空になった行で読み込み終了
 *     ・順位が空欄なら上から自動連番。同点タイは順位を手入力すればそのまま出る
 *     ・データは毎回全組入力してOK。表示数（B5）で何位まで見せるかを別途絞れる
 * ──────────────────────────────────
 */

const SPREADSHEET_ID = '10FbVGqe5_QQE-ynHQRHjzFz4yPPfIdeMpeJEhGeB-Gw';

// データ開始行（1始まり）。6行目がヘッダー、7行目からデータ
const DATA_START_ROW = 7;

function doGet() {
  let payload;
  try {
    payload = buildResponse_();
  } catch (e) {
    payload = { status: 'error', message: String(e) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildResponse_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();

  // B1に何か入力されているシートを探す（後勝ち＝一番右を採用）
  let target = null;
  for (const sheet of sheets) {
    if (isFlagOn_(sheet.getRange('B1').getValue())) {
      target = sheet;
    }
  }

  if (!target) {
    return { status: 'preparing' };
  }

  const values = target.getDataRange().getValues();

  const meta = {
    label:        toText_(values[1] && values[1][1]), // B2 見出し
    desc:         toText_(values[2] && values[2][1]), // B3 説明文
    updated:      toTimeText_(values[3] && values[3][1]), // B4 更新時刻
    displayCount: toDisplayCount_(values[4] && values[4][1]) // B5 表示数
  };

  const allRows = [];
  let autoRank = 0;
  for (let i = DATA_START_ROW - 1; i < values.length; i++) {
    const team = toText_(values[i][1]); // B列
    if (!team) break; // チーム名が空なら終了

    autoRank++;
    const rankCell = toText_(values[i][0]); // A列
    const rank = rankCell !== '' ? rankCell : String(autoRank);

    allRows.push({
      rank: rank,
      team: team,
      point: toPoint_(values[i][2]) // C列
    });
  }

  // 表示数（B5）が指定されていればそこまでに絞る。空欄なら全組表示
  const rows = meta.displayCount ? allRows.slice(0, meta.displayCount) : allRows;

  return {
    status: 'ok',
    sheet: target.getName(),
    label: meta.label,
    desc: meta.desc,
    updated: meta.updated,
    rows: rows
  };
}

/**
 * B1が「公開中」を意味する値かどうかを判定。
 * チェックボックスのTRUE、"OK"「公開」「ON」などの文字列、1 などの数値はON。
 * 空欄・FALSE・0・空文字はOFF。
 */
function isFlagOn_(v) {
  if (v === true) return true;
  if (v === false || v === null || v === undefined) return false;
  if (typeof v === 'number') return v !== 0;
  return String(v).trim() !== '';
}

/** セル値を文字列化（前後の空白を除去） */
function toText_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

/** 更新時刻：Date型なら「7/23 10:40」形式に、文字列ならそのまま */
function toTimeText_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'M/d HH:mm');
  }
  return toText_(v);
}

/**
 * 表示数を正の整数に正規化。
 * 空欄・0・数値でない・マイナスは null（=全組表示）を返す
 */
function toDisplayCount_(v) {
  if (typeof v === 'number') return v > 0 ? Math.floor(v) : null;
  const s = toText_(v);
  if (s === '') return null;
  const n = Number(s);
  return (!isNaN(n) && n > 0) ? Math.floor(n) : null;
}

/**
 * ポイントを数値に正規化。
 * "+24" → 24、"−3"（全角マイナス）→ -3、24 → 24、空欄 → 0
 */
function toPoint_(v) {
  if (typeof v === 'number') return v;
  const s = toText_(v)
    .replace(/[＋+]/g, '')
    .replace(/[−ー－‐]/g, '-')
    .replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}
