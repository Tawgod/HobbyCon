// ==========================================
// 1. MENU & UI TOOLS
// ==========================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Convention Tools')
      .addItem('1. Generate Master Schedule (For Web)', 'generateScheduleSummary')
      .addItem('2. Generate "All Events" Catalog', 'generateAllEventsTab')
      .addItem('3. Generate Daily Schedule Tabs', 'generateDailyTabs')
      .addSeparator()
      .addItem('4. Clear Entry Tab', 'clearEntryTab')
      .addToUi();
}