/** Classroom Jeopardy Studio — Google Sheets storage API
 * 1. Create a blank Google Sheet.
 * 2. Open Extensions > Apps Script and replace the editor contents with this file.
 * 3. Deploy > New deployment > Web app.
 * 4. Execute as: Me. Who has access: Anyone.
 * 5. Put the deployment's /exec URL in SHEET_WEB_APP_URL in the site code.
 * When updating this script: Deploy > Manage deployments > Edit > New version > Deploy.
 */
const SHEET_NAME = 'Games';
const HEADERS = ['id', 'title', 'subtitle', 'createdAt', 'updatedAt', 'gameJson'];

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'list';
    if (action !== 'list') return json_({ ok: false, error: 'Unknown action.' });
    return json_({ ok: true, games: readGames_() });
  } catch (error) { return json_({ ok: false, error: String(error.message || error) }); }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'save' && body.game) {
      const game = validateGame_(body.game); upsertGame_(game);
      return json_({ ok: true, id: game.id, updatedAt: game.updatedAt });
    }
    if (body.action === 'delete' && body.id) {
      return json_({ ok: true, id: body.id, deleted: deleteGame_(String(body.id)) });
    }
    return json_({ ok: false, error: 'Expected a save or delete action.' });
  } catch (error) { return json_({ ok: false, error: String(error.message || error) }); }
}

function readGames_() {
  const sheet = getSheet_();
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues().map(row => JSON.parse(row[5])).sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
}

function upsertGame_(game) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const ids = sheet.getLastRow() < 2 ? [] : sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const existingIndex = ids.indexOf(game.id);
    const row = [game.id, game.title, game.subtitle || '', game.createdAt, game.updatedAt, JSON.stringify(game)];
    if (existingIndex === -1) sheet.appendRow(row); else sheet.getRange(existingIndex + 2, 1, 1, HEADERS.length).setValues([row]);
  } finally { lock.releaseLock(); }
}

function deleteGame_(id) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    if (sheet.getLastRow() < 2) return false;
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const existingIndex = ids.indexOf(id);
    if (existingIndex === -1) return false;
    sheet.deleteRow(existingIndex + 2); return true;
  } finally { lock.releaseLock(); }
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS); sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#10267a').setFontColor('#ffffff');
    sheet.autoResizeColumns(1, 5); sheet.setColumnWidth(6, 520);
  }
  return sheet;
}

function validateGame_(game) {
  if (!game.id || !game.title) throw new Error('The game needs an id and title.');
  if (!Array.isArray(game.categories) || game.categories.length !== 5) throw new Error('The game must have five categories.');
  game.categories.forEach(category => { if (!category.name || !Array.isArray(category.clues) || category.clues.length !== 5) throw new Error('Each category needs a name and five clues.'); });
  game.createdAt = game.createdAt || new Date().toISOString(); game.updatedAt = new Date().toISOString(); return game;
}

function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
