// === CONFIGURATION ===

/** Spreadsheet & sheet identifiers */
const SPREADSHEET_ID = "1BXz46rvcFocg3F4ArDtq68MTqeslu-pX5icL7Mcg7ms";
const SHEET_NAME = "Form Responses 1";

/** Fixed stakeholder emails */
const STAKEHOLDER_EMAILS = [
  "sam@claimclimbers.com",
  "matt@claimclimbers.com",
  "james@claimclimbers.com"
];

/** Dynamically resolved URL for this deployed script */
const SCRIPT_URL = ScriptApp.getService().getUrl();

/** Cached spreadsheet objects */
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
const sheet = ss.getSheetByName(SHEET_NAME);

/** Header name to column index map */
const HEADER_MAP = (() => {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.reduce((m, h, i) => {
    m[h] = i + 1;
    return m;
  }, {});
})();
/**
 * Retrieve the column index for a given header name.
 * @param {string} name The header to look up.
 * @returns {number} 1-based column index.
 * @throws {Error} If the header is missing.
 */
function getColumnIndex(name) {
  const idx = HEADER_MAP[name];
  if (!idx) throw new Error('Header "' + name + '" not found');
  return idx;
}

/**
 * Creates an installable trigger for onFormSubmit when the script is installed.
 */
function onInstall() {
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
}

/**
 * Form submit handler.
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e Event object.
 */
function onFormSubmit(e) {
  try {
    const row = e.range.getRow();
    const data = buildNamedValueMap(e.namedValues);
    const dept = (data['Department'] || '').trim() || 'Other';

    const urlsByEmail = {};
    STAKEHOLDER_EMAILS.forEach(email => {
      urlsByEmail[email] = buildActionUrls(row, 1, email);
    });

    notifyStakeholders(1, row, data, urlsByEmail);
    Logger.log('Sent approval request to stakeholders for row ' + row);
  } catch (err) {
    Logger.log('Error in onFormSubmit: ' + err);
  }
}

/**
 * GET handler used by approval links.
 * @param {Object} e Query parameters.
 * @returns {HtmlOutput}
 */
function doGet(e) {
  try {
    const params = e.parameter;
    const row = parseInt(params.row, 10);
    const stage = parseInt(params.stage, 10);
    const decision = params.decision;
    const approver = params.approver;
    const token = params.token;

    if (isNaN(row) || isNaN(stage) || !decision || !approver || !token) {
      return errorPage('Invalid or missing parameters.');
    }

    if (!verifyToken(row, stage, approver, token)) {
      return errorPage('This approval link is no longer valid.');
    }

    if (Session.getActiveUser().getEmail() !== approver) {
      return errorPage('You are not authorized to act on this request.');
    }

    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    const data = getRowData(row);

    recordDecision(row, stage, decision, approver);

    if (stage === 1 && decision === 'Approved') {
      handleFirstApproval(row, data);
    } else if (stage === 1 && decision === 'Denied') {
      sendRequestorNotification(data, 'Denied');
    } else if (stage === 2) {
      handleSecondApproval(row, data, decision);
    }
    lock.releaseLock();
    deleteToken(row, stage, approver);

    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial,sans-serif;padding:40px;text-align:center;">' +
      '<h2 style="color:#2e6c80;">✅ Response Recorded</h2>' +
      '<p>Your decision of <strong>' + decision + '</strong> was logged.</p>' +
      '<p>You may now close this window.</p></div>'
    );
  } catch (err) {
    Logger.log('Error in doGet: ' + err);
    return errorPage('An unexpected error occurred.');
  }
}

/** Utility Functions **/

/**
 * Build a simple object from namedValues.
 * @param {Object} namedValues e.namedValues from form submit.
 * @returns {Object<string,string>}
 */
function buildNamedValueMap(namedValues) {
  const obj = {};
  Object.keys(namedValues).forEach(k => {
    obj[k] = Array.isArray(namedValues[k]) ? namedValues[k][0] : namedValues[k];
  });
  return obj;
}

/**
 * Builds approve/deny URLs for a given row, stage and approver.
 * @param {number} row
 * @param {number} stage
 * @param {string} approver
 * @returns {{approve:string,deny:string}}
 */
function buildActionUrls(row, stage, approver) {
  const token = generateToken(row, stage, approver);
  const base = SCRIPT_URL;
  return {
    approve: `${base}?row=${row}&stage=${stage}&decision=Approved&approver=${encodeURIComponent(approver)}&token=${token}`,
    deny: `${base}?row=${row}&stage=${stage}&decision=Denied&approver=${encodeURIComponent(approver)}&token=${token}`
  };
}

/**
 * Generates and stores a one-time token for the approval URL.
 * @private
 */
function generateToken(row, stage, approver) {
  const token = Utilities.getUuid();
  PropertiesService.getDocumentProperties().setProperty(`${row}_${stage}_${approver}`, token);
  return token;
}

/**
 * Checks whether the provided token matches the stored token.
 */
function verifyToken(row, stage, approver, token) {
  const stored = PropertiesService.getDocumentProperties().getProperty(`${row}_${stage}_${approver}`);
  return stored && stored === token;
}

function deleteToken(row, stage, approver) {
  PropertiesService.getDocumentProperties().deleteProperty(`${row}_${stage}_${approver}`);
}

/**
 * Retrieves data for a specific row as an object.
 * @param {number} row Row number
 * @returns {Object<string,string>}
 */
function getRowData(row) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const obj = {};
  headers.forEach((h, i) => obj[h] = values[i]);
  return obj;
}

/**
 * Records an approval decision into the sheet.
 */
function recordDecision(row, stage, decision, approver) {
  const timestamp = new Date();
  if (stage === 1) {
    sheet.getRange(row, getColumnIndex('1st Approval Status'), 1, 3)
         .setValues([[decision, timestamp, approver]]);
  } else {
    sheet.getRange(row, getColumnIndex('2nd Approval Status'), 1, 3)
         .setValues([[decision, timestamp, approver]]);
  }
}

/**
 * Sends second-level approval request emails on first approval.
 */
function handleFirstApproval(row, data) {
  const urlsByEmail = {};
  STAKEHOLDER_EMAILS.forEach(email => {
    urlsByEmail[email] = buildActionUrls(row, 2, email);
  });
  notifyStakeholders(2, row, data, urlsByEmail);
}

/**
 * Handles completion of the second approval and notifies the requestor.
 */
function handleSecondApproval(row, data, decision) {
  sendRequestorNotification(data, decision);
}

/**
 * Sends approval request emails to stakeholders.
 * @param {number} stage Approval stage (1 or 2)
 * @param {number} row Sheet row number
 * @param {Object<string,string>} data Row data
 * @param {Object<string,{approve:string,deny:string}>} urls Map of approver email to action URLs
 */
function notifyStakeholders(stage, row, data, urls) {
  const title = stage === 1 ? '📋 Requisition Approval Needed' : '📋 Second-Level Requisition Approval';
  const subjectPrefix = stage === 1 ? '' : '2nd ';
  const subject = `${subjectPrefix}Approval Needed: ${data['Requisition Title']}`;
  const name = data['Requestor Name'] || 'Requestor';

  Object.keys(urls).forEach(email => {
    try {
      const body = generateStyledEmail(data, urls[email].approve, urls[email].deny, title, name);
      MailApp.sendEmail({ to: email, subject, htmlBody: body });
    } catch (err) {
      Logger.log('Error sending notification to ' + email + ': ' + err);
    }
  });
}

/**
 * Notifies the requestor of the final decision.
 */
function sendRequestorNotification(data, finalDecision) {
  const email = data['Email Address'];
  if (!email) return;
  const title = finalDecision === 'Approved'
    ? '✅ Your Requisition Was Approved'
    : '❌ Your Requisition Was Denied';
  const color = finalDecision === 'Approved' ? '#2e6c80' : '#c0392b';
  const body = `
    <div style="font-family:Arial,sans-serif;line-height:1.6;">
      <h2 style="color:${color};">${title}</h2>
      <p>Your requisition titled <strong>${data['Requisition Title']}</strong> has been ${finalDecision.toLowerCase()}.</p>
      <hr/>
      <p style="font-size:0.9em;color:#888;">This message was generated automatically.</p>
    </div>`;
  try {
    MailApp.sendEmail({ to: email, subject: title, htmlBody: body });
  } catch (err) {
    Logger.log('Error sending requestor notification: ' + err);
  }
}

/**
 * Creates an error HtmlOutput.
 */
function errorPage(msg) {
  return HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;padding:40px;text-align:center;color:red;">' +
    '<h2>Error</h2><p>' + msg + '</p></div>'
  );
}

/**
 * Generates the styled email body used for approval requests.
 * @param {Object<string,string>} data Requisition data
 * @param {string} approveUrl URL for approval action
 * @param {string} denyUrl URL for denial action
 * @param {string} title Email title
 * @param {string} requestorName Name of the requestor
 * @returns {string} HTML string
 */
function generateStyledEmail(data, approveUrl, denyUrl, title, requestorName) {
  let html = `<div style="font-family:Arial,sans-serif;line-height:1.6;">`+
             `<p><strong>${requestorName}</strong> submitted a requisition that requires your review.</p>`+
             `<h2 style="color:#2e6c80;">${title}</h2>`+
             `<table cellpadding="6" cellspacing="0" style="border-collapse:collapse;">`;
  for (const [key, value] of Object.entries(data)) {
    if (!key.toLowerCase().includes('approval') && value) {
      html += `<tr><td style="font-weight:bold;padding:4px 8px;vertical-align:top;">${key}:</td>`+
              `<td style="padding:4px 8px;">${value}</td></tr>`;
    }
  }
  html += `</table><br/><div><strong>Take Action:</strong><br/><br/>`+
          `<a href="${approveUrl}" style="padding:10px 20px;background-color:#4CAF50;color:white;text-decoration:none;border-radius:4px;margin-right:10px;">✅ Approve</a>`+
          `<a href="${denyUrl}" style="padding:10px 20px;background-color:#f44336;color:white;text-decoration:none;border-radius:4px;">❌ Deny</a>`+
          `<p style="font-size:0.8em;color:#999;">Do not bookmark these links.</p>`+
          `</div></div>`;
  return html;
}
