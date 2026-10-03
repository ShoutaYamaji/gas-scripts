/**
 * 売上サマリーダッシュボード
 *
 * 「売上データ」シートの売上を月ごとに集計して「月次サマリー」シートに書き込み、
 * 月次推移を棒グラフで表示する。
 * setDailyTrigger を一度実行すると、毎朝9時に runDashboard が自動実行される。
 */

// シート名の定義
const SOURCE_SHEET_NAME = '売上データ';
const SUMMARY_SHEET_NAME = '月次サマリー';

// 自動実行の時刻（時）
const TRIGGER_HOUR = 9;

/**
 * メイン処理：売上データを月次集計し、サマリーとグラフを作成する
 * トリガー：時間主導型（毎日 9時台）※ setDailyTrigger で登録
 */
function runDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    throw new Error('「' + SOURCE_SHEET_NAME + '」シートが見つかりません。');
  }

  // 出力先シートがなければ作成する
  let summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SUMMARY_SHEET_NAME);
  }

  const monthlyTotals = aggregateByMonth_(sourceSheet);
  writeSummary_(summarySheet, monthlyTotals);
  drawChart_(summarySheet, monthlyTotals.length);
}

/**
 * 「売上データ」シートの全行を読み込み、月ごとの合計売上と件数を集計する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 売上データシート
 * @return {Array<{year: number, month: number, total: number, count: number}>} 古い月順の集計結果
 */
function aggregateByMonth_(sheet) {
  const lastRow = sheet.getLastRow();
  // 1行目は見出しとして読み飛ばす
  if (lastRow < 2) {
    return [];
  }

  // A〜D列をまとめて読み込む（1行ずつ読むと遅いため）
  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const totals = {};

  values.forEach(function (row) {
    const date = toDate_(row[0]);
    const amount = Number(row[3]);

    // 日付が読めない行・金額が数値でない行（空行など）は集計対象外
    if (!date || row[3] === '' || isNaN(amount)) {
      return;
    }

    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    const key = year * 100 + month; // 例：202601（並べ替え用のキー）

    if (!totals[key]) {
      totals[key] = { year: year, month: month, total: 0, count: 0 };
    }
    totals[key].total += amount;
    totals[key].count += 1;
  });

  // 月の古い順に並べる
  return Object.keys(totals)
    .map(Number)
    .sort(function (a, b) { return a - b; })
    .map(function (key) { return totals[key]; });
}

/**
 * セルの値を Date に変換する（日付型・"2026/01/05" 形式の文字列の両方に対応）
 * @param {*} value セルの値
 * @return {Date|null} 変換できない場合は null
 */
function toDate_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }
  if (typeof value === 'string' && value.trim() !== '') {
    const parsed = new Date(value.trim().replace(/-/g, '/'));
    return isNaN(parsed.getTime()) ? null : parsed;
  }
  return null;
}

/**
 * 「月次サマリー」シートをクリアし、集計結果を書き込む
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 月次サマリーシート
 * @param {Array<Object>} monthlyTotals aggregateByMonth_ の結果
 */
function writeSummary_(sheet, monthlyTotals) {
  // 値・書式を毎回クリアする（グラフは clear() で消えないため drawChart_ 側で削除）
  sheet.clear();

  const rows = [['月', '合計売上', '件数']];
  monthlyTotals.forEach(function (m) {
    rows.push([m.year + '年' + m.month + '月', m.total, m.count]);
  });

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);

  // 見出しの書式と数値の表示形式
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#f1f3f4');
  if (monthlyTotals.length > 0) {
    // 月の列は「2026年1月」の文字列のまま表示する（日付に自動変換させない）
    sheet.getRange(2, 1, monthlyTotals.length, 1).setNumberFormat('@');
    sheet.getRange(2, 2, monthlyTotals.length, 1).setNumberFormat('#,##0');
    sheet.getRange(2, 3, monthlyTotals.length, 1).setNumberFormat('#,##0');
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
}

/**
 * 「月次サマリー」シートに月次推移の棒グラフを作成する（既存のグラフは作り直す）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 月次サマリーシート
 * @param {number} monthCount 集計した月の数
 */
function drawChart_(sheet, monthCount) {
  // 前回作成したグラフを削除する
  sheet.getCharts().forEach(function (chart) {
    sheet.removeChart(chart);
  });

  if (monthCount === 0) {
    return;
  }

  // A列（月）と B列（合計売上）をグラフの範囲にする
  const range = sheet.getRange(1, 1, monthCount + 1, 2);

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN) // 縦棒グラフ
    .addRange(range)
    .setNumHeaders(1)
    .setPosition(1, 5, 0, 0) // E1 セルの位置に配置
    .setOption('title', '月次売上推移')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: '月' })
    .setOption('vAxis', { title: '合計売上', format: '#,##0' })
    .setOption('width', 640)
    .setOption('height', 360)
    .build();

  sheet.insertChart(chart);
}

/**
 * runDashboard を毎朝9時に自動実行するトリガーを登録する
 * ※ 手動で一度だけ実行する。既存の runDashboard トリガーは削除してから登録し直すため、何度実行しても重複しない
 * ※ 実行時刻はスクリプトのタイムゾーン（appsscript.json の timeZone）基準で、9時〜10時の間に実行される
 */
function setDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'runDashboard') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('runDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(TRIGGER_HOUR)
    .create();
}
