/**
 * Entry point for the web app.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.mode === 'bootstrap') {
    try {
      const data = getBootstrapData();
      return ContentService.createTextOutput(JSON.stringify({
        ok: true,
        data
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
      return ContentService.createTextOutput(JSON.stringify({
        ok: false,
        error: error && error.message ? error.message : String(error)
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return HtmlService.createTemplateFromFile('app/client/index')
    .evaluate()
    .setTitle('OKRsMatter')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
