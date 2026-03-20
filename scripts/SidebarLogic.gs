function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Scripts')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Admin Panel');
  SpreadsheetApp.getUi() 
      .showSidebar(html);
}
