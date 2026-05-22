/**
 * =================================================================
 * 【ファイル 1/3: メイン司令塔 (A_Main.gs)】
 * =================================================================
 */

// 共通UI（カスタムメニュー）
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カスタムメニュー')
    .addItem('月次集計ツールを開く', 'showSidebar')
    .addSeparator()
    .addItem('グラフ生成ツールを開く', 'gt_showSidebar')
    .addSeparator()
    .addItem('★ 平均時給シートを更新', 'calculateAverageWages_v2')
    .addSeparator()
    .addItem('お財布くんを開く', 'showOsaifuKunSidebar')
    .addToUi();
}

// 各ツールのUI（サイドバー）を開く関数群
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('日次集計ツール');
  SpreadsheetApp.getUi().showSidebar(html);
}

function gt_showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('GraphTool_UI').setTitle('グラフ生成ツール');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 「お財布くん」のポップアップを表示する関数
 */
function showOsaifuKunSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('OsaifuKun_UI')
      .setWidth(550)
      .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'お財布くん');
}