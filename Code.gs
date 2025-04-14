/**
 * @OnlyCurrentDoc
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  DocumentApp.getUi()
      .createMenu('My Add-on')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
}

/**
 * Opens a sidebar in the document.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('My Custom Sidebar');
  DocumentApp.getUi().showSidebar(html);
} 