// Client ID and API key from the Developer Console
//localhost:8000, andrewmacheret.com
const CLIENT_ID = '303548077940-ov8iafec5pqhrd457fhe8sb2q5ak6o8s.apps.googleusercontent.com';
const API_KEY = 'AIzaSyAF3bxFGxELOH7aK3imyc__dpRuo9M2j5M';

// Array of API discovery doc URLs for APIs used by the quickstart
const DISCOVERY_DOCS = ["https://sheets.googleapis.com/$discovery/rest?version=v4"];

// Authorization scopes required by the API; multiple scopes can be
// included, separated by spaces.
const SCOPES = "https://www.googleapis.com/auth/spreadsheets";

const formName = 'goalie-form';
const form = document.forms[formName];

const $id = document.getElementById.bind(document);

const spreadsheetId = '1lKPFQXbZsQDzUD5eVrmzDee50dBdB9c0WuBKXLZ6Rfk';
const spreadsheetLink = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
const year = Math.max(2019, new Date().getFullYear());
const dataSheet = `${year} Goals`;
const settingsSheet = `${year} Settings`;
const startingRow = 4;

const settingsVersion = 1;
let settings = {
  settingsVersion,
  lastRowNum: 0,
  questions: []
}
let submittingForm = false;

function removeTime(date) {
  return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
}
$id('date').valueAsDate = removeTime(new Date());
$id('spreadsheet-link').href = spreadsheetLink;

setMessage('info', 'Loading Google APIs...');
let autoreload = setTimeout(function() {
  setMessage('info', 'Google APIs timed out, reloading...');
  window.location.reload();
}, 5000);

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


/**
 *  On load, called to load the auth2 library and API client library.
 */
function handleClientLoad() {
  gapi.load('client:auth2', initClient);
}

/**
 *  Initializes the API client library and sets up sign-in state
 *  listeners.
 */
function initClient() {
  gapi.client.init({
    apiKey: API_KEY,
    clientId: CLIENT_ID,
    discoveryDocs: DISCOVERY_DOCS,
    scope: SCOPES
  }).then(() => {
    window.clearTimeout(autoreload);

    // Listen for sign-in state changes.
    gapi.auth2.getAuthInstance().isSignedIn.listen(updateSigninStatus);

    // Handle the initial sign-in state.
    updateSigninStatus(gapi.auth2.getAuthInstance().isSignedIn.get());
    $id('authorize-button').onclick = handleAuthClick;
    $id('signout-button').onclick = handleSignoutClick;
  });
}

/**
 *  Called when the signed in status changes, to update the UI
 *  appropriately. After a sign-in, the API is called.
 */
function updateSigninStatus(isSignedIn) {
  if (isSignedIn) {
    $id('authorize-button').style.display = 'none';
    $id('signout-button').style.display = '';
    loadSheets();
  } else {
    $id('authorize-button').style.display = '';
    $id('signout-button').style.display = 'none';
    setMessage('warning', 'Need authorization.');
  }
}

/**
 *  Sign in the user upon button click.
 */
function handleAuthClick(event) {
  gapi.auth2.getAuthInstance().signIn();
}

/**
 *  Sign out the user upon button click.
 */
function handleSignoutClick(event) {
  gapi.auth2.getAuthInstance().signOut();
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

function getSpreadsheetValues(range) {
  return new Promise((resolve, reject) => {
    gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId,
      range
    }).then(response => {
      resolve(response.result.values);
    }, response => {
      reject(response.result.error.message);
    });
  });
}

function appendSpreadsheetRow(range, row) {
  return new Promise((resolve, reject) => {
    const params = {
      spreadsheetId,
      range,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      includeValuesInResponse: false
    };

    const valueRange = {
      'values': [
        row
      ],
    };

    gapi.client.sheets.spreadsheets.values.append(params, valueRange).then(response => {
      resolve(response.result.updates);
    }, response => {
      reject(response.result.error.message);
    });
  });
}

function loadSheetsCached() {
  let cached = window.localStorage.getItem('settings');
  if (!cached) return;
  
  newSettings = JSON.parse(window.localStorage.getItem('settings'));
  if (newSettings.version === settingsVersion) {
    settings = newSettings
  }

  displaySettings();
}

function loadSheets() {
  setMessage('info', 'Loading sheets...');

  getSpreadsheetValues(`'${settingsSheet}'!A1:Z`) // TODO - confirm
  .then(data => {
    if (data.length === 0) {
      setMessage('danger', 'No data found.');
      return;
    }

    loadSettings(data);

    setMessage('success', 'Loaded!');
  }).catch(error => {
    console.error(error);
    setMessage('danger', 'Error: ' + error);
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

function loadField(name, type, values) {
  const className = escapeClassName(name);
  console.log(className, name, type, values);
  const html = {
    main: $id(`template-${type}`).innerHTML
            .replace(/\{\{CLASS\}\}/g, className)
            .replace(/\{\{FIELD_LABEL\}\}/g, escapeHTML(name))
  };

  if (type === 'radio' || type === 'checkbox') {
    let choicesHtml = '';
    for (let i=0; i<values.length; i++) {
      const value = escapeHTML(values[i]);
      choicesHtml += $id(`template-${type}-choice`).innerHTML
        .replace(/\{\{CLASS\}\}/g, className)
        .replace(/\{\{INDEX\}\}/g, ''+i)
        .replace(/\{\{VALUE\}\}/g, value)
        .replace(/\{\{LABEL\}\}/g, value)
        .replace(/\{\{BUTTON_ACTIVE\}\}/g, i === 0 && type === 'radio' ? 'active' : ''); // TODO
    }
    html.choices = choicesHtml;
  }

  return {
    name,
    className,
    type,
    values,
    html
  };
}

function loadSettings(values) {

  // values
  settings.lastRowNum = parseInt(values[2][0], 10) || 4;

  // custom fields
  settings.questions = []
  for (let c = 2; c < values[1].length; c++) {
    const fieldName = (values[1][c] || '').trim();
    const fieldType = (values[2][c] || '').trim().toLowerCase();
    if (fieldName !== '' && fieldType !== '') {
      const fieldValues = [];
      for (let r = 3; r < values.length; r++) {
        const fieldValue = (values[r][c] || '').trim();
        if (fieldValue !== '') {
          fieldValues.push(fieldValue);
        }
      }
      settings.questions.push(loadField(fieldName, fieldType, fieldValues));
    }
  }

  window.localStorage.setItem('settings', JSON.stringify(settings));

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

  if (date.substring(0, 4) !== (''+year)) {
    setMessage('warning', `Date is not in ${year}. ${date}`);
    $id('date').focus();
    $id('date').classList.add('is-invalid');
    return false;
  }

  return true;
}

function submitForm() {
  try {
    setMessage('info', `Saving...`);

    const date = form['date'].valueAsDate.toISOString().substring(0, 10);

    const customValues = [];
    for (const question of settings.questions) {
      let customValue;
      if (question.type === 'decimal') {
        customValue = parseInt(form[question.className].value, 10) || 0;
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
      console.log(updates);
      setSubmitEnabled(true);
      setMessage('success', `Saved <a href="${spreadsheetLink}" target="_blank" class="alert-link">${updates.updatedRange}</a>`);

      settings.lastRowNum += 1;

      window.localStorage.setItem('settings', JSON.stringify(settings));
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

