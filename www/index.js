// OAuth client ID from the Google Cloud console. It's public, and Google only
// accepts it from its authorized origins (localhost:8000, andrewmacheret.com)
const CLIENT_ID = '248237455343-92pm91eir0o4mh86obkq2v60u8a30dkp.apps.googleusercontent.com';
// The file picker needs the Google Cloud project number and an API key. The
// key is public too, and only works from andrewmacheret.com for the picker
const APP_ID = '248237455343';
const PICKER_API_KEY = 'AIzaSyAxkzeJQJKcw5kZ_P-NXqG15scFqMIrf3A';
// drive.file only reaches spreadsheets the user picks or Goalie creates
const SCOPES = 'https://www.googleapis.com/auth/drive.file';
const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';

const formName = 'goalie-form';
const form = document.forms[formName];

const $id = document.getElementById.bind(document);

const dataSheet = 'Input';
const settingsSheet = 'Settings';
const startingRow = 4;

// The spreadsheets this browser has used, newest first
const sheetsKey = 'goalie-sheets';
const maxRecentSheets = 10;
let sheetList = loadSheetList();

const settingsVersion = 3;
let settings = {
  settingsVersion,
  lastRowNum: 0,
  questions: []
}
let submittingForm = false;
let tokenClient;
// Goalie and Miler Lite share one Google sign-in, so they share the saved token
const tokenKey = 'daily-tracking-token';
let accessToken = loadSavedToken();
let signedOut = false;

// Older versions kept settings for every sheet under one key
window.localStorage.removeItem('settings');

function removeTime(date) {
  return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
}
$id('date').valueAsDate = removeTime(new Date());

setMessage('info', 'Loading Google sign-in...');
$id('remove-button').onclick = handleRemoveClick;
renderSheetControls();
loadSheetsCached();

// Escape a string for HTML interpolation.
function escapeHTML(string) {
  return ('' + string).replace(/[&<>"'\/]/g, match => {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#x27;',
      '/': '&#x2F;'
    }[match];
  });
}

function escapeClassName(string) {
  return ('' + string).toLowerCase().replace(/[^a-z0-9]/g, '-')
}

function loadSheetList() {
  try {
    const list = JSON.parse(window.localStorage.getItem(sheetsKey));
    if (list && Array.isArray(list.recent)) return list;
  } catch (error) {
    console.error(error);
  }
  return {current: null, recent: []};
}

function currentSheetId() {
  return sheetList.current;
}

function spreadsheetLink() {
  return `https://docs.google.com/spreadsheets/d/${currentSheetId()}/edit`;
}

function rememberSheet(id, name) {
  sheetList.recent = [{id, name}, ...sheetList.recent.filter(sheet => sheet.id !== id)].slice(0, maxRecentSheets);
  sheetList.current = id;
  window.localStorage.setItem(sheetsKey, JSON.stringify(sheetList));
  renderSheetControls();
}

function renderSheetControls() {
  const select = $id('sheet-select');
  select.innerHTML = sheetList.recent.map(sheet =>
    `<option value="${escapeHTML(sheet.id)}">${escapeHTML(sheet.name)}</option>`).join('');
  select.value = currentSheetId() || '';
  select.style.display = sheetList.recent.length ? '' : 'none';
  $id('remove-button').style.display = currentSheetId() ? '' : 'none';

  const link = $id('spreadsheet-link');
  link.style.display = currentSheetId() ? 'inline-block' : 'none';
  link.href = spreadsheetLink();
}

function handleSheetSelect(event) {
  const sheet = sheetList.recent.find(sheet => sheet.id === event.target.value);
  if (!sheet) return;
  rememberSheet(sheet.id, sheet.name);
  form.style.display = 'none';
  loadSheetsCached();
  if (accessToken) loadSheets();
}

// Takes the spreadsheet off this browser's list. It stays in Google Drive
function handleRemoveClick() {
  const id = currentSheetId();
  if (!id) return;
  window.localStorage.removeItem(settingsKey());
  sheetList.recent = sheetList.recent.filter(sheet => sheet.id !== id);
  sheetList.current = sheetList.recent.length ? sheetList.recent[0].id : null;
  window.localStorage.setItem(sheetsKey, JSON.stringify(sheetList));
  renderSheetControls();
  form.style.display = 'none';
  loadSheetsCached();
  updateSigninStatus(!!accessToken);
}


/**
 *  Called when the Google Identity Services script loads.
 */
function handleClientLoad() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: handleTokenResponse,
    error_callback: error => setMessage('danger', 'Sign-in failed: ' + escapeHTML(error.message || error.type)),
  });

  $id('authorize-button').onclick = handleAuthClick;
  $id('signout-button').onclick = handleSignoutClick;
  $id('sheet-select').onchange = handleSheetSelect;
  $id('pick-button').onclick = handlePickClick;
  $id('create-button').onclick = handleCreateClick;
  updateSigninStatus(!!accessToken);
}

// Access tokens last an hour. Keeping it until then means a reload doesn't
// ask the user to sign in again
function loadSavedToken() {
  try {
    const saved = JSON.parse(window.localStorage.getItem(tokenKey));
    if (saved && saved.expiresAt > Date.now()) return saved.token;
  } catch (error) {
    console.error(error);
  }
  return null;
}

function saveToken(token, expiresInSeconds) {
  // Stop a minute early so a request doesn't start with an almost-expired token
  const expiresAt = Date.now() + (expiresInSeconds - 60) * 1000;
  window.localStorage.setItem(tokenKey, JSON.stringify({token, expiresAt}));
}

function forgetToken() {
  accessToken = null;
  window.localStorage.removeItem(tokenKey);
}

function handleClientLoadError() {
  setMessage('danger', 'Could not load Google sign-in. Check your connection and reload.');
}

function handleTokenResponse(response) {
  if (response.error) {
    setMessage('danger', 'Sign-in failed: ' + escapeHTML(response.error_description || response.error));
    return;
  }
  if (!google.accounts.oauth2.hasGrantedAllScopes(response, SCOPES)) {
    setMessage('warning', 'Goalie needs access to its spreadsheets to work.');
    return;
  }
  accessToken = response.access_token;
  saveToken(accessToken, Number(response.expires_in) || 3600);
  updateSigninStatus(true);
}

/**
 *  Called when the signed in status changes, to update the UI
 *  appropriately. After a sign-in, the API is called.
 */
function updateSigninStatus(isSignedIn) {
  $id('authorize-button').style.display = isSignedIn ? 'none' : '';
  $id('signout-button').style.display = isSignedIn ? '' : 'none';
  $id('sheet-buttons').style.display = isSignedIn ? '' : 'none';
  if (!isSignedIn) {
    setMessage('warning', 'Click Authorize to sign in with Google.');
  } else if (currentSheetId()) {
    loadSheets();
  } else {
    setMessage('info', 'Click <b>Open</b> to choose your Goalie spreadsheet, or <b>New</b> to create one.');
  }
}

/**
 *  Sign in the user upon button click.
 */
function handleAuthClick(event) {
  // An empty prompt skips the consent screen once the user has agreed to it
  tokenClient.requestAccessToken({prompt: signedOut ? 'select_account' : ''});
}

/**
 *  Sign out the user upon button click.
 */
function handleSignoutClick(event) {
  // Forget the token without revoking it. Revoking would also drop access to
  // the spreadsheets the user picked, and they'd have to pick them again
  forgetToken();
  signedOut = true;
  updateSigninStatus(false);
}

function handlePickClick() {
  if (typeof gapi === 'undefined') {
    setMessage('warning', 'The file picker is still loading. Try again in a moment.');
    return;
  }
  gapi.load('picker', () => {
    const view = new google.picker.DocsView(google.picker.ViewId.SPREADSHEETS)
      .setMode(google.picker.DocsViewMode.LIST);
    new google.picker.PickerBuilder()
      .setTitle('Choose your Goalie spreadsheet')
      .addView(view)
      .setOAuthToken(accessToken)
      .setDeveloperKey(PICKER_API_KEY)
      .setAppId(APP_ID)
      .setCallback(handlePicked)
      .build()
      .setVisible(true);
  });
}

function handlePicked(data) {
  if (data[google.picker.Response.ACTION] !== google.picker.Action.PICKED) return;
  const doc = data[google.picker.Response.DOCUMENTS][0];
  rememberSheet(doc[google.picker.Document.ID], doc[google.picker.Document.NAME]);
  form.style.display = 'none';
  loadSheetsCached();
  loadSheets();
}

function handleCreateClick() {
  setMessage('info', 'Creating a new Goalie spreadsheet...');
  sheetsFetch('', {method: 'POST', body: JSON.stringify(starterSpreadsheet())})
  .then(result => {
    rememberSheet(result.spreadsheetId, result.properties.title);
    loadSheets();
  }).catch(error => {
    console.error(error);
    setMessage('danger', 'Error: ' + error);
  });
}

// One spreadsheet cell. A [value, color] pair sets the font color too
function starterCell(value) {
  if (value === null) return {};
  const [content, color] = Array.isArray(value) ? value : [value];
  const cell = {userEnteredValue: typeof content === 'number' ? {numberValue: content}
    : content.startsWith('=') ? {formulaValue: content} : {stringValue: content}};
  if (color) cell.userEnteredFormat = {textFormat: {foregroundColor: color}};
  return cell;
}

function starterTab(title, rows) {
  return {
    properties: {title},
    data: [{rowData: rows.map(row => ({values: row.map(starterCell)}))}],
  };
}

// A small example tracker in the layout Goalie reads
function starterSpreadsheet() {
  const green = {red: 0.2, green: 0.6, blue: 0.2};
  const red = {red: 0.8, green: 0.2, blue: 0.2};
  return {
    properties: {title: `Goalie ${new Date().getFullYear()}`},
    sheets: [
      starterTab(dataSheet, [
        ['Goalie adds a row here each time you submit.'],
        [],
        ['Date', `=ARRAYFORMULA('${settingsSheet}'!C2:2)`],
      ]),
      starterTab(settingsSheet, [
        ['Questions go in row 2 from column C, their type in row 3 (radio, checkbox, text, decimal, or decimal-1 for one decimal place), and their choices from row 4 down. The font color of a choice shows in Goalie.'],
        ['Next row', null, 'Exercised', 'Meals', 'Glasses of water', 'Notes'],
        [`=COUNTA('${dataSheet}'!A${startingRow}:A)+${startingRow}`, null, 'radio', 'checkbox', 'decimal-0', 'text'],
        [null, null, ['Yes', green], 'Breakfast'],
        [null, null, ['No', red], 'Lunch'],
        [null, null, null, 'Dinner'],
      ]),
    ],
  };
}

/**
 * Append a pre element to the body containing the given message
 * as its text node. Used to display the results of the API call.
 *
 * @param {string} message Text to be placed in pre element.
 */
function setMessage(level, message) {
  const messageElement = $id('message');
  messageElement.innerHTML = message;
  messageElement.className = 'alert alert-' + level;

  const messageElement2 = $id('message-below');
  messageElement2.innerHTML = message;
  messageElement2.className = 'alert alert-' + level;

  console.log(level, message);
}

// Calls the Sheets API and rejects with the error message
async function sheetsFetch(path, options = {}) {
  const headers = {...options.headers, Authorization: `Bearer ${accessToken}`};
  if (options.body) headers['Content-Type'] = 'application/json';
  const response = await fetch(`${SHEETS_API}${path}`, {...options, headers});
  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    // Access tokens expire after an hour, so ask to sign in again
    if (response.status === 401) {
      forgetToken();
      updateSigninStatus(false);
      throw 'Sign-in expired. Click Authorize to sign in again.';
    }
    throw escapeHTML(body.error ? body.error.message : response.statusText);
  }
  return body;
}

function getSpreadsheet(range) {
  const params = new URLSearchParams({ranges: range, includeGridData: 'true'});
  return sheetsFetch(`/${currentSheetId()}?${params}`);
}

function appendSpreadsheetRow(range, row) {
  const params = new URLSearchParams({
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    includeValuesInResponse: 'false'
  });
  return sheetsFetch(`/${currentSheetId()}/values/${encodeURIComponent(range)}:append?${params}`, {
    method: 'POST',
    body: JSON.stringify({values: [row]}),
  }).then(result => result.updates);
}

function settingsKey() {
  return `goalie-settings-${currentSheetId()}`;
}

function loadSheetsCached() {
  if (!currentSheetId()) return;
  try {
    const cached = JSON.parse(window.localStorage.getItem(settingsKey()));
    if (cached && cached.settingsVersion === settingsVersion) {
      settings = cached;
      displaySettings();
    }
  } catch (error) {
    console.error(error);
  }
}

function loadSheets() {
  setMessage('info', 'Loading spreadsheet...');

  getSpreadsheet(`'${settingsSheet}'!A1:AZ100`)
  .then(result => {
    // Keep the saved name up to date if the spreadsheet was renamed
    rememberSheet(currentSheetId(), result.properties.title);

    const data = result.sheets[0].data[0];
    if (!data || !data.rowData) {
      setMessage('danger', 'The Settings tab is empty.');
      return;
    }

    loadSettings(data);

    setMessage('success', 'Loaded!');
  }).catch(error => {
    console.error(error);
    if (String(error).includes('Unable to parse range')) {
      setMessage('danger', `This spreadsheet has no ${settingsSheet} tab. Click <b>Open</b> to choose a Goalie spreadsheet, or <b>New</b> to create one.`);
    } else if (/permission|not found/i.test(error)) {
      // drive.file only reaches spreadsheets this Google account picked in Goalie
      setMessage('danger', 'Goalie can\'t open this spreadsheet with this Google account. Click <b>Open</b> to choose it again.');
    } else {
      setMessage('danger', 'Error: ' + error);
    }
  })
}

function dontSubmit(event) {
  if (event.keyCode == 13) {
    const focusable = Array.from(form.querySelectorAll('input:not([readonly]):not([type="radio"]):not([type="checkbox"]),button[type="submit"]'));
    const next = focusable[focusable.indexOf(event.target) + 1];
    if (next) {
      next.focus();
      return false;
    }
  }
  return true;
}

function loadField(name, type, values, colors) {
  const className = escapeClassName(name);
  //console.log(className, name, type, values);
  // decimal-N types share the decimal template
  const template = $id(`template-${type.startsWith('decimal') ? 'decimal' : type}`);
  if (!template) throw `Unknown question type "${escapeHTML(type)}" for ${escapeHTML(name)}`;
  let mainHtml = template.innerHTML
      .replace(/\{\{CLASS\}\}/g, className)
      .replace(/\{\{FIELD_LABEL\}\}/g, escapeHTML(name));
  if (type.startsWith('decimal')) {
    mainHtml = mainHtml.replace(/\{\{DECIMAL_STEP\}\}/, Math.pow(10, -decimalPrecision(type)));
  }
  const html = { main: mainHtml };

  if (type === 'radio' || type === 'checkbox') {
    let choicesHtml = '';
    for (let i=0; i<values.length; i++) {
      const value = escapeHTML(values[i]);
      const color = colors[i];
      choicesHtml += $id(`template-${type}-choice`).innerHTML
        .replace(/\{\{CLASS\}\}/g, className)
        .replace(/\{\{INDEX\}\}/g, ''+i)
        .replace(/\{\{VALUE\}\}/g, value)
        .replace(/\{\{LABEL\}\}/g, value)
        .replace(/\{\{COLOR\}\}/g, color)
        .replace(/\{\{BUTTON_ACTIVE\}\}/g, i === 0 && type === 'radio' ? 'active' : ''); // TODO
    }
    html.choices = choicesHtml;
  }

  return {
    name,
    className,
    type,
    values,
    colors,
    html
  };
}

function loadSettings(data) {
  
  const values = data.rowData.map(y => (y.values || []).map(z => z.formattedValue));
  const colors = data.rowData.map(y => (y.values || []).map(z =>
                 z && z.effectiveFormat && z.effectiveFormat.textFormat && z.effectiveFormat.textFormat.foregroundColor || {}));
  while (values.length < 3) values.push([]);

  // values
  settings.lastRowNum = parseInt(values[2][0], 10) || 4;

  // custom fields
  settings.questions = []
  for (let c = 2; c < values[1].length; c++) {
    const fieldName = (values[1][c] || '').trim();
    const fieldType = (values[2][c] || '').trim().toLowerCase();
    if (fieldName !== '' && fieldType !== '') {
      const fieldValues = [];
      const fieldColors = [];
      for (let r = 3; r < values.length; r++) {
        const fieldValue = (values[r][c] || '').trim();
        if (fieldValue !== '') {
          fieldValues.push(fieldValue);
          const color = colors[r][c];
          const colorString = '#' + ['red', 'green', 'blue'].map(k =>
                              Math.round((color[k] || 0) * 255).toString(16).padStart(2, '0')).join('');
          fieldColors.push(colorString);
        }
      }
      settings.questions.push(loadField(fieldName, fieldType, fieldValues, fieldColors));
    }
  }

  window.localStorage.setItem(settingsKey(), JSON.stringify(settings));

  displaySettings();
}

function displaySettings() {
  let html = '';
  for (const question of settings.questions) {
    html += question.html.main;
  }
  $id('dynamic-fields').innerHTML = html;

  for (const question of settings.questions) {
    if (question.html.choices) {
      $id(`${question.className}-choices`).innerHTML = question.html.choices;
      if (question.type === 'radio') {
        form[`${question.className}`][0].checked = true; // TODO
      }
    }
  }

  form.style.display = '';
}

// TODO - validate customValues?
function validate({date, customValues}) {
  $id('date').classList.remove('is-invalid');

  if (!date.match(/^\d{4}-\d{2}-\d{2}$/)) {
    setMessage('warning', 'Date is not valid.');
    $id('date').focus();
    $id('date').classList.add('is-invalid');
    return false;
  }

  return true;
}

function decimalPrecision(decimalType) {
  const typeParts = decimalType.split('-', 2);
  return typeParts.length > 1 ? typeParts[1] : 2;
}

function submitForm() {
  try {
    setMessage('info', `Saving...`);

    const date = form['date'].valueAsDate.toISOString().substring(0, 10);

    const customValues = [];
    for (const question of settings.questions) {
      let customValue;
      if (question.type.startsWith('decimal')) {
        const precision = decimalPrecision(question.type);
        customValue = parseFloat(form[question.className].value || 0).toFixed(precision);
      } else if (question.type === 'checkbox') {
        customValue = [].slice.call(form[question.className]).filter(e => e.checked).map(e => e.value).join(', ');
      } else {
        customValue = form[question.className].value.trim();
      }
      customValues.push(customValue);
    }

    if (!validate({date, customValues})) {
      return;
    }

    const row = [date, ...customValues];

    setSubmitEnabled(false);

    const endingColumn = String.fromCharCode('A'.charCodeAt(0) + customValues.length);

    appendSpreadsheetRow(`'${dataSheet}'!A${startingRow}:${endingColumn}`, row)
    .then(updates => {
      //console.log(updates);
      setSubmitEnabled(true);
      setMessage('success', `Saved <a href="${spreadsheetLink()}" target="_blank" class="alert-link">${updates.updatedRange}</a>`);

      settings.lastRowNum += 1;

      window.localStorage.setItem(settingsKey(), JSON.stringify(settings));
    }).catch(error => {
      console.error(error);
      setSubmitEnabled(true);
      setMessage('danger', 'Error: ' + error);
    });
  } catch(error) {
    console.error(error);
    setSubmitEnabled(true);
    setMessage('danger', 'Error: ' + error);
  }
}

function setSubmitEnabled(shouldBeEnabled) {
  if (shouldBeEnabled) {
    $id('submit').removeAttribute('disabled');
    submittingForm = false;
  } else {
    $id('submit').setAttribute('disabled', 'disabled');
    submittingForm = true;
  }
}

