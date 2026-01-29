/**
 * Entry point for the web app.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.mode === 'debug') {
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty('OKR_SPREADSHEET_ID') || '';
    let url = '';
    try {
      url = id ? SpreadsheetApp.openById(id).getUrl() : '';
    } catch (error) {
      url = `Error: ${error.message}`;
    }
    return ContentService.createTextOutput(JSON.stringify({
      scriptId: ScriptApp.getScriptId(),
      spreadsheetId: id,
      spreadsheetUrl: url
    })).setMimeType(ContentService.MimeType.JSON);
  }

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
