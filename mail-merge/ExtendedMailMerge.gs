/**
 * @OnlyCurrentDoc
 *
 * This script provides tools for modular mail merge automation in Google Sheets.
 * All settings (sheet names, column headers, status mappings) are dynamically configured
 * via the custom menu UI and stored in the document's properties.
 * * Original script inspired by Martin Hawksey, 2022
 * Extended functionalities developed by Morell Alargha, 2025
 * Modified for fully modular UI configuration by Gemini.
 */

const CONFIG_KEY = "MAIL_MERGE_CONFIG";
const LOG_SHEET = "Send Log";

// =================================================================
// === CONFIGURATION MANAGEMENT ====================================
// =================================================================

function getSavedConfig_() {
  const props = PropertiesService.getDocumentProperties();
  const configStr = props.getProperty(CONFIG_KEY);
  return configStr ? JSON.parse(configStr) : null;
}

function saveConfig_(config) {
  PropertiesService.getDocumentProperties().setProperty(
    CONFIG_KEY,
    JSON.stringify(config),
  );
}

// =================================================================
// === SCRIPT SETUP & MENU =========================================
// =================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Mail Merge Automation")
    .addItem("1. Master Setup (Run First)", "masterSetup")
    .addItem("2. Add Auto-Trigger Mapping", "addTriggerMapping")
    .addItem("3. Set Sender Name & Reply-To", "setSenderIdentity")
    .addItem("4. View Current Config", "viewConfig")
    .addSeparator()
    .addItem("Enable Automation Trigger", "createSpreadsheetEditTrigger")
    .addSeparator()
    .addItem("Send Bulk Emails for a Status", "sendBulkEmailForStatus")
    .addItem("Send Bulk Email to Individuals", "sendBulkEmailToIndividuals")
    .addSeparator()
    .addItem("Check Email Quota", "checkQuota")
    .addItem("Check Sent Status vs. Gmail", "checkGmailSentStatus")
    .addItem("Scan for Bounced Emails", "scanForBounces")
    .addSeparator()
    .addItem("Debug Last Edit", "debugLastEdit")
    .addItem("Clear Cache (Global Params)", "clearCache")
    .addItem("Reset All Settings", "clearAllConfig")
    .addToUi();
}

function masterSetup() {
  const ui = SpreadsheetApp.getUi();
  let config = getSavedConfig_() || { STATUS_MAPPINGS: {} };

  let res = ui.prompt(
    "Data Sheet Name",
    "Enter the EXACT name of the sheet containing your mailing list:",
    ui.ButtonSet.OK_CANCEL,
  );
  if (res.getSelectedButton() !== ui.Button.OK || !res.getResponseText().trim())
    return;
  config.MAILER_SHEET = res.getResponseText().trim();

  res = ui.prompt(
    "Email Column",
    "Enter the EXACT header name of the column containing recipient email addresses:",
    ui.ButtonSet.OK_CANCEL,
  );
  if (res.getSelectedButton() !== ui.Button.OK || !res.getResponseText().trim())
    return;
  config.RECIPIENT_COLUMN = res.getResponseText().trim();

  res = ui.prompt(
    "Status Column",
    "Enter the EXACT header name of the column containing the status dropdown:",
    ui.ButtonSet.OK_CANCEL,
  );
  if (res.getSelectedButton() !== ui.Button.OK || !res.getResponseText().trim())
    return;
  config.STATUS_COLUMN = res.getResponseText().trim();

  res = ui.prompt(
    "Global Parameters Sheet",
    "Optional: Enter the name of a sheet containing global variables (leave blank to skip):",
    ui.ButtonSet.OK_CANCEL,
  );
  if (res.getSelectedButton() === ui.Button.OK) {
    config.GENERAL_SHEET = res.getResponseText().trim();
  }

  saveConfig_(config);
  ui.alert(
    "Success!",
    "Basic sheet settings have been saved. You can now set up auto-trigger mappings or use the bulk sending tools.",
    ui.ButtonSet.OK,
  );
}

function addTriggerMapping() {
  const ui = SpreadsheetApp.getUi();
  let config = getSavedConfig_();

  if (!config || !config.MAILER_SHEET) {
    ui.alert(
      "Configuration Missing",
      "Please run '1. Master Setup' first.",
      ui.ButtonSet.OK,
    );
    return;
  }

  let res = ui.prompt(
    "Trigger Status",
    "Enter the exact status value that should trigger an email (e.g., 'Confirmed'):",
    ui.ButtonSet.OK_CANCEL,
  );
  if (res.getSelectedButton() !== ui.Button.OK || !res.getResponseText().trim())
    return;
  const status = res.getResponseText().trim();

  res = ui.prompt(
    "Draft Subject",
    "Enter the exact subject line of the Gmail draft to send for this status:",
    ui.ButtonSet.OK_CANCEL,
  );
  if (res.getSelectedButton() !== ui.Button.OK || !res.getResponseText().trim())
    return;
  const draftSubject = res.getResponseText().trim();

  res = ui.prompt(
    "Timestamp Column",
    "Optional: Enter the column header to record the exact timestamp when sent (leave blank to skip):",
    ui.ButtonSet.OK_CANCEL,
  );
  const timestampColumn =
    res.getSelectedButton() === ui.Button.OK
      ? res.getResponseText().trim()
      : "";

  config.STATUS_MAPPINGS = config.STATUS_MAPPINGS || {};
  config.STATUS_MAPPINGS[status] = { draftSubject, timestampColumn };
  saveConfig_(config);

  ui.alert(
    "Success",
    `Mapping added!\n\nWhen status becomes: "${status}"\nSend Draft: "${draftSubject}"\nTimestamp in: "${timestampColumn || "None"}"`,
    ui.ButtonSet.OK,
  );
}

function setSenderIdentity() {
  const ui = SpreadsheetApp.getUi();
  let config = getSavedConfig_();
  if (!config || !config.MAILER_SHEET) {
    ui.alert("Please run '1. Master Setup' first.");
    return;
  }

  let res = ui.prompt(
    "Sender Name",
    "Enter the name recipients should see in the 'From' field (e.g., 'Forwards Team Europe Hikes'). Leave blank to keep the default account name:",
    ui.ButtonSet.OK_CANCEL,
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  config.SENDER_NAME = res.getResponseText().trim();

  res = ui.prompt(
    "Reply-To Address",
    "Enter the email address replies should go to (e.g., 'hikes@forwards-team.org'). Leave blank to use the sending account:",
    ui.ButtonSet.OK_CANCEL,
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  config.REPLY_TO = res.getResponseText().trim();

  saveConfig_(config);
  ui.alert(
    "Saved",
    `Sender Name: "${config.SENDER_NAME || "(default)"}"\nReply-To: "${config.REPLY_TO || "(default)"}"\n\nThese will be applied to every email the script sends.`,
    ui.ButtonSet.OK,
  );
}

function viewConfig() {
  const ui = SpreadsheetApp.getUi();
  const config = getSavedConfig_();
  if (!config) {
    ui.alert("No configuration found. Run Master Setup.");
    return;
  }
  const htmlOutput = HtmlService.createHtmlOutput(
    `<p><b>Current Settings:</b></p><textarea rows="20" cols="60" readonly>${JSON.stringify(config, null, 2)}</textarea>`,
  )
    .setWidth(500)
    .setHeight(400);
  ui.showModalDialog(htmlOutput, "System Configuration");
}

function clearAllConfig() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    "Warning",
    "This will delete all sheet mappings and trigger settings. Are you sure?",
    ui.ButtonSet.YES_NO,
  );
  if (res === ui.Button.YES) {
    PropertiesService.getDocumentProperties().deleteProperty(CONFIG_KEY);
    ui.alert("Configuration cleared. Please run Master Setup again.");
  }
}

function createSpreadsheetEditTrigger() {
  const sheet = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === "handleEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger("handleEdit").forSpreadsheet(sheet).onEdit().create();
  SpreadsheetApp.getUi().alert(
    "Success",
    "The automated email trigger has been enabled. Editing the configured status column will now send mapped emails.",
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

// =================================================================
// === CORE: AUTOMATED EMAIL TRIGGER LOGIC =========================
// =================================================================

function handleEdit(e) {
  if (!e) return;
  const config = getSavedConfig_();
  if (
    !config ||
    !config.MAILER_SHEET ||
    !config.STATUS_COLUMN ||
    !config.RECIPIENT_COLUMN
  )
    return;

  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== config.MAILER_SHEET) return;

  const range = e.range;
  if (range.getRow() <= 1) return;

  const rawHeaders = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  const headers = rawHeaders.map((h) => (typeof h === "string" ? h.trim() : h));
  const statusColIdx = headers.indexOf(config.STATUS_COLUMN);

  if (statusColIdx === -1) return; // Status column not found

  if (range.getColumn() === statusColIdx + 1) {
    const status = String(e.value).trim();
    const rowDataArray = sheet
      .getRange(range.getRow(), 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const rowObject = headers.reduce((obj, header, i) => {
      let cellValue = rowDataArray[i] || "";
      if (typeof cellValue === "string") cellValue = cellValue.trim();
      obj[header] = cellValue;
      return obj;
    }, {});

    CacheService.getScriptCache().put(
      "last_edited_row",
      JSON.stringify(rowObject),
      3600,
    );
    processEmailTrigger(
      rowObject,
      status,
      sheet,
      range.getRow(),
      headers,
      config,
    );
  }
}

function processEmailTrigger(
  rowObject,
  status,
  sheet,
  rowNum,
  headers,
  config,
) {
  if (!config.STATUS_MAPPINGS || !config.STATUS_MAPPINGS[status]) return;

  const mapping = config.STATUS_MAPPINGS[status];
  if (mapping.timestampColumn && rowObject[mapping.timestampColumn]) return; // Already sent

  const recipient = rowObject[config.RECIPIENT_COLUMN];
  try {
    if (!recipient) throw new Error("Recipient email address is missing.");
    if (!isValidEmail_(recipient))
      throw new Error(`Recipient "${recipient}" is not a valid email address.`);
    if (MailApp.getRemainingDailyQuota() <= 0)
      throw new Error("Daily email quota exhausted. Try again tomorrow.");

    const globalParameters = getGlobalParameters_(config);
    const emailTemplate = getGmailTemplateFromDrafts_(mapping.draftSubject);
    const messageObject = fillInTemplateFromObject_(
      emailTemplate.message,
      rowObject,
      globalParameters,
    );

    sendMergedEmail_(recipient, messageObject, emailTemplate, config);
    logSend_(recipient, messageObject.subject, "SENT", `Row ${rowNum} (auto)`);

    if (mapping.timestampColumn) {
      const timestampColIdx = headers.indexOf(mapping.timestampColumn);
      if (timestampColIdx !== -1) {
        sheet.getRange(rowNum, timestampColIdx + 1).setValue(new Date());
      }
    }
  } catch (e) {
    Logger.log(`Error sending email for row ${rowNum}: ${e.message}`);
    logSend_(recipient || "(missing)", mapping.draftSubject, "FAILED", e.message);
    try {
      SpreadsheetApp.getUi().alert(
        `An error occurred while sending an email: ${e.message}`,
      );
    } catch (uiError) {
      // UI is not available in automatic trigger context; the Send Log captured it.
    }
  }
}

// =================================================================
// === BULK EMAIL SENDER FUNCTIONS =================================
// =================================================================

function sendBulkEmailForStatus() {
  const ui = SpreadsheetApp.getUi();
  const config = getSavedConfig_();
  if (!config || !config.MAILER_SHEET) {
    ui.alert("Please run 'Master Setup' first.");
    return;
  }

  const statusPrompt = ui.prompt(
    "Step 1: Select Status",
    "Enter the exact status you want to send emails to:",
    ui.ButtonSet.OK_CANCEL,
  );
  const targetStatus = statusPrompt.getResponseText().trim();
  if (statusPrompt.getSelectedButton() !== ui.Button.OK || !targetStatus)
    return;

  const draftPrompt = ui.prompt(
    "Step 2: Select Draft",
    "Enter the exact subject of the Gmail draft you want to send:",
    ui.ButtonSet.OK_CANCEL,
  );
  const draftSubject = draftPrompt.getResponseText().trim();
  if (draftPrompt.getSelectedButton() !== ui.Button.OK || !draftSubject) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    config.MAILER_SHEET,
  );
  if (!sheet) {
    ui.alert(
      "Error",
      `Sheet "${config.MAILER_SHEET}" not found.`,
      ui.ButtonSet.OK,
    );
    return;
  }

  try {
    const allData = sheet.getDataRange().getValues();
    const headers = allData
      .shift()
      .map((h) => (typeof h === "string" ? h.trim() : h));
    const statusColIdx = headers.indexOf(config.STATUS_COLUMN);

    if (statusColIdx === -1) {
      ui.alert(
        "Error",
        `Status column "${config.STATUS_COLUMN}" not found.`,
        ui.ButtonSet.OK,
      );
      return;
    }

    const recipientsToEmail = allData
      .filter((row) => String(row[statusColIdx]).trim() === targetStatus)
      .map((rowDataArray) =>
        headers.reduce((obj, header, i) => {
          let cellValue = rowDataArray[i] || "";
          if (typeof cellValue === "string") cellValue = cellValue.trim();
          obj[header] = cellValue;
          return obj;
        }, {}),
      );

    if (recipientsToEmail.length === 0) {
      ui.alert(
        "No Recipients Found",
        `No users with the status "${targetStatus}" were found.`,
        ui.ButtonSet.OK,
      );
      return;
    }

    let remainingQuota = MailApp.getRemainingDailyQuota();
    const startAlert = ui.alert(
      "Starting Bulk Send",
      `Found ${recipientsToEmail.length} recipient(s). You have ${remainingQuota} email(s) left in today's quota.\n\n${
        remainingQuota < recipientsToEmail.length
          ? `WARNING: quota is lower than the number of recipients. Only the first ${remainingQuota} will be sent; the rest will be marked SKIPPED in the Send Log.\n\n`
          : ""
      }Continue?`,
      ui.ButtonSet.OK_CANCEL,
    );
    if (startAlert !== ui.Button.OK) return;

    const globalParameters = getGlobalParameters_(config);
    const emailTemplate = getGmailTemplateFromDrafts_(draftSubject);
    const sentLog = [];
    let failed = 0;
    let skipped = 0;

    recipientsToEmail.forEach((rowObject) => {
      const recipient = rowObject[config.RECIPIENT_COLUMN];
      try {
        if (!recipient || !isValidEmail_(recipient)) {
          failed++;
          logSend_(recipient || "(missing)", draftSubject, "FAILED", "Invalid or missing email address.");
          return;
        }
        if (remainingQuota <= 0) {
          skipped++;
          logSend_(recipient, draftSubject, "SKIPPED", "Daily quota exhausted.");
          return;
        }

        const messageObject = fillInTemplateFromObject_(
          emailTemplate.message,
          rowObject,
          globalParameters,
        );
        sendMergedEmail_(recipient, messageObject, emailTemplate, config);
        remainingQuota--;
        sentLog.push(recipient);
        logSend_(recipient, messageObject.subject, "SENT", "Bulk by status");
      } catch (e) {
        failed++;
        Logger.log(`Bulk send failed for ${recipient}: ${e.message}`);
        logSend_(recipient || "(missing)", draftSubject, "FAILED", e.message);
      }
    });

    const htmlOutput = HtmlService.createHtmlOutput(
      `<p>Bulk send complete. <b>Sent: ${sentLog.length}</b> &nbsp; Failed: ${failed} &nbsp; Skipped: ${skipped}.<br>Full details are in the "${LOG_SHEET}" sheet.</p><textarea rows="15" cols="80" readonly>${sentLog.join("\n")}</textarea>`,
    )
      .setWidth(600)
      .setHeight(350);
    ui.showModalDialog(htmlOutput, "Bulk Send Log");
  } catch (e) {
    ui.alert("An Error Occurred", e.message, ui.ButtonSet.OK);
  }
}

function sendBulkEmailToIndividuals() {
  const ui = SpreadsheetApp.getUi();
  const config = getSavedConfig_();
  if (!config || !config.MAILER_SHEET) {
    ui.alert("Please run 'Master Setup' first.");
    return;
  }

  const emailPrompt = ui.prompt(
    "Step 1: Enter Recipients",
    "Enter email addresses separated by commas or spaces:",
    ui.ButtonSet.OK_CANCEL,
  );
  const emailListStr = emailPrompt.getResponseText();
  if (emailPrompt.getSelectedButton() !== ui.Button.OK || !emailListStr) return;

  const targetEmails = emailListStr
    .split(/[,\s]+/)
    .map((email) => email.trim())
    .filter((email) => email);
  if (targetEmails.length === 0) return;

  const draftPrompt = ui.prompt(
    "Step 2: Select Draft",
    "Enter the exact subject of the Gmail draft you want to send:",
    ui.ButtonSet.OK_CANCEL,
  );
  const draftSubject = draftPrompt.getResponseText().trim();
  if (draftPrompt.getSelectedButton() !== ui.Button.OK || !draftSubject) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    config.MAILER_SHEET,
  );
  if (!sheet) return;

  try {
    const allData = sheet.getDataRange().getValues();
    const headers = allData
      .shift()
      .map((h) => (typeof h === "string" ? h.trim() : h));
    const emailColIdx = headers.indexOf(config.RECIPIENT_COLUMN);

    if (emailColIdx === -1) {
      ui.alert(
        "Error",
        `Email column "${config.RECIPIENT_COLUMN}" not found.`,
        ui.ButtonSet.OK,
      );
      return;
    }

    const dataMap = new Map();
    allData.forEach((rowDataArray) => {
      const email = rowDataArray[emailColIdx];
      if (email) {
        const rowObject = headers.reduce((obj, header, i) => {
          let cellValue = rowDataArray[i] || "";
          if (typeof cellValue === "string") cellValue = cellValue.trim();
          obj[header] = cellValue;
          return obj;
        }, {});
        dataMap.set(String(email).trim(), rowObject);
      }
    });

    const globalParameters = getGlobalParameters_(config);
    const emailTemplate = getGmailTemplateFromDrafts_(draftSubject);
    const sentLog = [];
    let failed = 0;
    let skipped = 0;
    let remainingQuota = MailApp.getRemainingDailyQuota();

    targetEmails.forEach((email) => {
      try {
        if (!isValidEmail_(email)) {
          failed++;
          logSend_(email, draftSubject, "FAILED", "Invalid email address.");
          return;
        }
        if (remainingQuota <= 0) {
          skipped++;
          logSend_(email, draftSubject, "SKIPPED", "Daily quota exhausted.");
          return;
        }
        const rowObject = dataMap.get(email) || {}; // Will use global params if not found in sheet
        const messageObject = fillInTemplateFromObject_(
          emailTemplate.message,
          rowObject,
          globalParameters,
        );
        sendMergedEmail_(email, messageObject, emailTemplate, config);
        remainingQuota--;
        sentLog.push(email);
        logSend_(email, messageObject.subject, "SENT", "Bulk to individuals");
      } catch (e) {
        failed++;
        Logger.log(`Bulk send failed for ${email}: ${e.message}`);
        logSend_(email, draftSubject, "FAILED", e.message);
      }
    });

    const htmlOutput = HtmlService.createHtmlOutput(
      `<p>Sent successfully to ${sentLog.length} user(s). Failed: ${failed} &nbsp; Skipped: ${skipped}.<br>Full details are in the "${LOG_SHEET}" sheet.</p><textarea rows="15" cols="80" readonly>${sentLog.join("\n")}</textarea>`,
    )
      .setWidth(600)
      .setHeight(350);
    ui.showModalDialog(htmlOutput, "Bulk Send Log");
  } catch (e) {
    ui.alert("An Error Occurred", e.message, ui.ButtonSet.OK);
  }
}

// =================================================================
// === HELPER & DEBUGGING FUNCTIONS ================================
// =================================================================

function clearCache() {
  CacheService.getScriptCache().remove("global_parameters");
  SpreadsheetApp.getUi().alert(
    "Cache Cleared",
    "Global sheet cache has been cleared.",
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

function debugLastEdit() {
  const ui = SpreadsheetApp.getUi();
  const config = getSavedConfig_();
  const lastRowJSON = CacheService.getScriptCache().get("last_edited_row");
  if (!lastRowJSON) {
    ui.alert(
      "No Edit Data Found",
      "Please edit a status first, then run this command again.",
      ui.ButtonSet.OK,
    );
    return;
  }
  const rowData = JSON.parse(lastRowJSON);
  const globalData = config ? getGlobalParameters_(config) : {};
  let debugText = "--- GLOBAL DATA ---\n" + JSON.stringify(globalData, null, 2);
  debugText += "\n\n--- ROW DATA ---\n" + JSON.stringify(rowData, null, 2);
  const htmlOutput = HtmlService.createHtmlOutput(
    `<p>Data for the last edited row.</p><textarea rows="20" cols="80" readonly>${debugText}</textarea>`,
  )
    .setWidth(600)
    .setHeight(400);
  ui.showModalDialog(htmlOutput, "Last Edit Debug Data");
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email).replace(/\s/g, ""));
}

function sendMergedEmail_(recipient, messageObject, emailTemplate, config) {
  const options = {
    to: String(recipient).replace(/\s/g, ""),
    subject: messageObject.subject,
    body: messageObject.text,
    htmlBody: messageObject.html,
    attachments: emailTemplate.attachments,
    inlineImages: emailTemplate.inlineImages,
  };
  if (config.SENDER_NAME) options.name = config.SENDER_NAME;
  if (config.REPLY_TO) options.replyTo = config.REPLY_TO;
  MailApp.sendEmail(options);
}

function getOrCreateLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet
      .getRange(1, 1, 1, 5)
      .setValues([["Timestamp", "Recipient", "Subject", "Status", "Detail"]])
      .setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function logSend_(recipient, subject, status, detail) {
  try {
    getOrCreateLogSheet_().appendRow([
      new Date(),
      recipient,
      subject,
      status,
      detail || "",
    ]);
  } catch (e) {
    Logger.log(`Could not write to log sheet: ${e.message}`);
  }
}

function getGlobalParameters_(config) {
  if (!config || !config.GENERAL_SHEET) return {};
  const cache = CacheService.getScriptCache();
  const cached = cache.get("global_parameters");
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const generalSheet = ss.getSheetByName(config.GENERAL_SHEET);
  if (!generalSheet) return {};

  const data = generalSheet
    .getRange(2, 2, generalSheet.getLastRow() - 1, 3)
    .getValues();
  const globalParams = {};
  data.forEach((row) => {
    let value = row[1];
    const alias = row[2];
    if (alias && typeof alias === "string" && value !== "") {
      const cleanedAlias = alias.replace(/[{}]+/g, "").trim();
      if (cleanedAlias) {
        if (typeof value === "string")
          value = value.replace(/\u00A0/g, " ").trim();
        globalParams[cleanedAlias] = value;
      }
    }
  });
  cache.put("global_parameters", JSON.stringify(globalParams), 21600);
  return globalParams;
}

function fillInTemplateFromObject_(template, rowData, globalData) {
  // Aliases removed: Map directly from column headers. (e.g., {{First Name}} looks for "First Name" column)
  const replacementData = { ...globalData, ...rowData };

  let subject = template.subject;
  let htmlBody = template.html;
  let textBody = template.text;

  for (const key in replacementData) {
    let value = replacementData[key] || "";
    if (key.toLowerCase().includes("email")) {
      value = String(value).replace(/\s/g, "");
    }

    // Safely escape the key to prevent regex failures with special characters
    const escapedKey = key.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, "\\$&");
    const placeholderRegex = new RegExp("{{" + escapedKey + "}}", "gi"); // Make it case insensitive to help avoid errors

    subject = subject.replace(placeholderRegex, value);
    htmlBody = htmlBody.replace(placeholderRegex, value);
    textBody = textBody.replace(placeholderRegex, value);
  }
  return { subject: subject, html: htmlBody, text: textBody };
}

function getGmailTemplateFromDrafts_(subject_line) {
  try {
    const drafts = GmailApp.getDrafts();
    const draft = drafts.filter(
      (element) => element.getMessage().getSubject() === subject_line,
    )[0];
    if (!draft)
      throw new Error(`Draft with subject "${subject_line}" not found.`);

    const msg = draft.getMessage();
    const allInlineImages = msg.getAttachments({
      includeInlineImages: true,
      includeAttachments: false,
    });
    const attachments = msg.getAttachments({ includeInlineImages: false });
    const htmlBody = msg.getBody();
    const img_obj = allInlineImages.reduce(
      (obj, i) => ((obj[i.getName()] = i), obj),
      {},
    );
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^>]+>', "g");
    const matches = [...htmlBody.matchAll(imgexp)];
    const inlineImagesObj = {};
    matches.forEach((match) => (inlineImagesObj[match[1]] = img_obj[match[2]]));

    return {
      message: {
        subject: msg.getSubject(),
        text: msg.getPlainBody(),
        html: htmlBody,
      },
      attachments: attachments,
      inlineImages: inlineImagesObj,
    };
  } catch (e) {
    throw new Error(
      `Oops - can't find or process Gmail draft with subject "${subject_line}". Error: ${e.message}`,
    );
  }
}

function checkQuota() {
  const remaining = MailApp.getRemainingDailyQuota();
  SpreadsheetApp.getUi().alert(
    "Email Quota",
    `You can still send ${remaining} more email(s) today.\n\nLimits are set by Google and reset every 24 hours:\n  - Consumer Gmail: 500 / day\n  - Google Workspace: 1,500 / day\n\nThese cannot be raised by the script. For higher volume or better deliverability, use a dedicated email service (Brevo, SendGrid, Amazon SES).`,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

function scanForBounces() {
  const ui = SpreadsheetApp.getUi();
  const config = getSavedConfig_();
  if (!config || !config.MAILER_SHEET || !config.RECIPIENT_COLUMN) {
    ui.alert("Please run 'Master Setup' first.");
    return;
  }

  const daysPrompt = ui.prompt(
    "Scan for Bounced Emails",
    "How many days back should I scan your inbox for delivery failures? (default: 14)",
    ui.ButtonSet.OK_CANCEL,
  );
  if (daysPrompt.getSelectedButton() !== ui.Button.OK) return;
  let days = parseInt(daysPrompt.getResponseText().trim(), 10);
  if (isNaN(days) || days <= 0) days = 14;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    config.MAILER_SHEET,
  );
  if (!sheet) return;

  // Build the set of addresses we actually mailed, so we only flag real participants.
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map((h) => (typeof h === "string" ? h.trim() : h));
  const emailColIdx = headers.indexOf(config.RECIPIENT_COLUMN);
  if (emailColIdx === -1) {
    ui.alert(
      "Error",
      `Email column "${config.RECIPIENT_COLUMN}" not found.`,
      ui.ButtonSet.OK,
    );
    return;
  }
  const recipientSet = new Set(
    sheet
      .getRange(2, emailColIdx + 1, Math.max(sheet.getLastRow() - 1, 1), 1)
      .getValues()
      .flat()
      .map((e) => String(e).replace(/\s/g, "").toLowerCase())
      .filter((e) => e.includes("@")),
  );

  // Delivery-failure notices come back from the mail delivery subsystem.
  const query =
    'from:mailer-daemon OR subject:("Delivery Status Notification") OR subject:("Undelivered Mail Returned to Sender") OR subject:("Address not found") newer_than:' +
    days +
    "d";
  const threads = GmailApp.search(query, 0, 300);

  const bounced = new Set();
  threads.forEach((thread) => {
    thread.getMessages().forEach((msg) => {
      const body = (msg.getPlainBody() || "").toLowerCase();
      recipientSet.forEach((addr) => {
        if (body.indexOf(addr) !== -1) bounced.add(addr);
      });
    });
  });

  if (bounced.size === 0) {
    ui.alert(
      "No Bounces Found",
      `No delivery-failure messages matching your participant list were found in the last ${days} day(s).\n\nNote: this only catches HARD bounces that generate a reply. Silent spam-foldering and DKIM/DMARC drops do NOT bounce, so they will not appear here.`,
      ui.ButtonSet.OK,
    );
    return;
  }

  const list = [...bounced].sort();
  list.forEach((addr) =>
    logSend_(
      addr,
      "(bounce scan)",
      "BOUNCED",
      `Delivery failure found within ${days} days`,
    ),
  );

  const htmlOutput = HtmlService.createHtmlOutput(
    `<p><b>${list.length}</b> participant address(es) had delivery failures in the last ${days} day(s). These did NOT arrive and have been logged as BOUNCED in the "${LOG_SHEET}" sheet.</p><textarea rows="15" cols="60" readonly>${list.join("\n")}</textarea>`,
  )
    .setWidth(550)
    .setHeight(380);
  ui.showModalDialog(htmlOutput, "Bounced Emails");
}

function checkGmailSentStatus() {
  const ui = SpreadsheetApp.getUi();
  const config = getSavedConfig_();
  if (!config || !config.MAILER_SHEET) {
    ui.alert("Please run 'Master Setup' first.");
    return;
  }

  const columnPrompt = ui.prompt(
    "Step 1: Select Column",
    "Enter the exact name of the column containing the email addresses:",
    ui.ButtonSet.OK_CANCEL,
  );
  const columnToCheck = columnPrompt.getResponseText();
  if (columnPrompt.getSelectedButton() !== ui.Button.OK || !columnToCheck)
    return;

  const subjectPrompt = ui.prompt(
    "Step 2: Specify Subject",
    "Enter the exact subject line to search for:",
    ui.ButtonSet.OK_CANCEL,
  );
  const subjectToCheck = subjectPrompt.getResponseText();
  if (subjectPrompt.getSelectedButton() !== ui.Button.OK || !subjectToCheck)
    return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    config.MAILER_SHEET,
  );
  if (!sheet) return;

  try {
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const emailColIdx = headers.indexOf(columnToCheck);
    if (emailColIdx === -1) {
      ui.alert(
        "Error",
        `Column "${columnToCheck}" was not found.`,
        ui.ButtonSet.OK,
      );
      return;
    }

    const emailData = sheet
      .getRange(2, emailColIdx + 1, sheet.getLastRow() - 1, 1)
      .getValues();
    const emailsToCheck = [
      ...new Set(
        emailData
          .flat()
          .filter((email) => typeof email === "string" && email.includes("@")),
      ),
    ];

    if (emailsToCheck.length === 0) return;

    ui.alert(
      "Starting Check",
      `Checking ${emailsToCheck.length} unique email addresses. This may take some time...`,
      ui.ButtonSet.OK,
    );

    const twoMonthsAgo = new Date();
    twoMonthsAgo.setMonth(twoMonthsAgo.getMonth() - 2);
    const dateQuery = `${twoMonthsAgo.getFullYear()}/${twoMonthsAgo.getMonth() + 1}/${twoMonthsAgo.getDate()}`;
    const notSentList = [];

    emailsToCheck.forEach((email) => {
      const query = `to:(${email}) subject:("${subjectToCheck}") after:${dateQuery}`;
      const threads = GmailApp.search(query, 0, 1);
      if (threads.length === 0) notSentList.push(email);
    });

    if (notSentList.length === 0) {
      ui.alert(
        "All Clear!",
        `An email with the subject "${subjectToCheck}" has been sent to every address.`,
        ui.ButtonSet.OK,
      );
    } else {
      const htmlOutput = HtmlService.createHtmlOutput(
        `<p>Found ${notSentList.length} emails that were not sent:</p><textarea rows="15" cols="60" readonly>${notSentList.join("\n")}</textarea>`,
      )
        .setWidth(500)
        .setHeight(350);
      ui.showModalDialog(htmlOutput, "Emails Not Sent Recently");
    }
  } catch (e) {
    ui.alert("An Error Occurred", e.message, ui.ButtonSet.OK);
  }
}
