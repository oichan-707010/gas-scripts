/**
 * 売上データの月次集計ダッシュボード
 *
 * 「売上データ」シートを月ごとに集計し、「月次サマリー」シートに
 * 合計売上と件数を書き込んで、月次推移の棒グラフを作成する。
 */

// シート名の定義
const SALES_SHEET_NAME = '売上データ';
const SUMMARY_SHEET_NAME = '月次サマリー';

// 自動実行する時刻（時）
const TRIGGER_HOUR = 9;

/**
 * メイン処理：売上データを月ごとに集計し、サマリーとグラフを作成する
 */
function runDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName(SALES_SHEET_NAME);
  if (!salesSheet) {
    throw new Error(`「${SALES_SHEET_NAME}」シートが見つかりません`);
  }

  // 出力先シートがなければ作成する
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME) || ss.insertSheet(SUMMARY_SHEET_NAME);

  const summary = aggregateByMonth_(salesSheet, ss.getSpreadsheetTimeZone());
  writeSummary_(summarySheet, summary);
  createChart_(summarySheet, summary.length);
}

/**
 * 売上データを月ごとに集計する
 * @param {Sheet} sheet 売上データのシート
 * @param {string} timeZone スプレッドシートのタイムゾーン
 * @return {Array<Array>} [月の表示名, 合計売上, 件数] の配列（月の古い順）
 */
function aggregateByMonth_(sheet, timeZone) {
  const lastRow = sheet.getLastRow();
  // 1行目は見出しとして扱い、データがなければ空の結果を返す
  if (lastRow < 2) {
    return [];
  }

  // A〜D列の全データを一度に読み込む
  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

  // キーは並べ替え用に「yyyy-MM」形式にする
  const totals = {};
  values.forEach(([dateValue, , , amount]) => {
    const date = toDate_(dateValue);
    // 日付が読み取れない行や金額が数値でない行は集計しない
    if (!date || typeof amount !== 'number') {
      return;
    }
    const key = Utilities.formatDate(date, timeZone, 'yyyy-MM');
    if (!totals[key]) {
      totals[key] = { amount: 0, count: 0 };
    }
    totals[key].amount += amount;
    totals[key].count += 1;
  });

  return Object.keys(totals)
    .sort()
    .map((key) => {
      const [year, month] = key.split('-');
      return [`${year}年${Number(month)}月`, totals[key].amount, totals[key].count];
    });
}

/**
 * セルの値を Date に変換する（日付型と「2026/01/05」形式の文字列に対応）
 * @param {*} value セルの値
 * @return {Date|null} 変換できない場合は null
 */
function toDate_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }
  if (typeof value === 'string' && value.trim() !== '') {
    const date = new Date(value.trim().replace(/-/g, '/'));
    return isNaN(date.getTime()) ? null : date;
  }
  return null;
}

/**
 * 「月次サマリー」シートをクリアして集計結果を書き込む
 * @param {Sheet} sheet 出力先のシート
 * @param {Array<Array>} summary 集計結果
 */
function writeSummary_(sheet, summary) {
  // 前回の内容とグラフをすべて消す
  sheet.clear();
  sheet.getCharts().forEach((chart) => sheet.removeChart(chart));

  const rows = [['月', '合計売上', '件数'], ...summary];
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);

  // 見出しと数値の書式を整える
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  if (summary.length > 0) {
    sheet.getRange(2, 2, summary.length, 1).setNumberFormat('#,##0');
    sheet.getRange(2, 3, summary.length, 1).setNumberFormat('0');
  }
  sheet.autoResizeColumns(1, 3);
}

/**
 * 月次推移の棒グラフを「月次サマリー」シートに作成する
 * @param {Sheet} sheet 出力先のシート
 * @param {number} monthCount 集計した月の数
 */
function createChart_(sheet, monthCount) {
  // 集計結果がなければグラフは作らない
  if (monthCount === 0) {
    return;
  }

  // 月（A列）と合計売上（B列）を見出し込みで使う
  const range = sheet.getRange(1, 1, monthCount + 1, 2);
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(range)
    .setNumHeaders(1)
    .setPosition(1, 5, 0, 0)
    .setOption('title', '月次売上推移')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: '月' })
    .setOption('vAxis', { title: '合計売上', format: '#,##0' })
    .build();
  sheet.insertChart(chart);
}

/**
 * 毎朝9時に runDashboard を実行するトリガーを登録する
 * 重複を防ぐため、既存の runDashboard のトリガーは削除してから登録する
 */
function setDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === 'runDashboard')
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('runDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(TRIGGER_HOUR)
    .create();
}
