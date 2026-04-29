/**
 * HPP Seva Connect — web app entry and HTML includes.
 */

/**
 * Serves the main shell (Index.html as evaluated template).
 * @param {Object} e Request parameters (unused).
 * @return {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('HPP - Akshar Setu')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0');
}

/**
 * @param {string} filename File in project without extension, e.g. 'Stylesheet'.
 * @return {string} Raw file contents for template inclusion.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Optional HTTP POST handler for future structured routing (not used by google.script.run).
 * @param {Object} e Post payload / parameters.
 * @return {GoogleAppsScript.Content.TextOutput}
 */
function doPost(e) {
  return ContentService.createTextOutput(
    JSON.stringify({
      success: false,
      error:
        'POST endpoint is reserved for future use. Use the web app UI with google.script.run.',
    })
  ).setMimeType(ContentService.MimeType.JSON);
}
