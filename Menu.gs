// ==========================================
// 1. MENU & UI TOOLS
// ==========================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Convention Tools')
      .addItem('1. Generate Clean Lists & Schedule', 'generateScheduleSummary')
      .addItem('2. Generate Google Docs Table Tents', 'generateDocsTableTents')
      .addItem('3. Generate Daily table Schedule', 'generatePrintableSchedules')

      .addSeparator()
      .addItem('4. Clear Entry Tab', 'clearEntryTab')
      .addToUi();
}