/**
 * @OnlyCurrentDoc
 *
 * This script provides tools for mail merge automation in Google Sheets.
 * 1. Automatically sends template emails based on a status change in the "Mailer" sheet.
 * 2. Pulls global parameters (e.g., event name, dates) from a "General" sheet.
 * 3. Includes a debugging tool to inspect data for the last edited row.
 * 4. Provides a utility to send bulk emails to all users with a specific status.
 *
 * Original script inspired by Martin Hawksey, 2022
 * Extended functionalities developed by Morell Alargha, 2025
 * Licensed under the Apache License, Version 2.0 (the "License");
 * https://www.apache.org/licenses/LICENSE-2.0
 *
 * Modified and enhanced for dynamic configuration and global parameters.
 */

// =================================================================
// === CONFIGURATION ===============================================
// =================================================================
// --- User-defined settings. These will be preserved. ---
// =================================================================

const CONFIG = {
  // The name of the sheet where the mail merge is triggered.
  MAILER_SHEET: "Mailer",

  // The name of the sheet containing global parameters.
  GENERAL_SHEET: "General",

  // The name of the column that contains the status dropdown.
  STATUS_COLUMN: "Status",

  // The name of the column with the recipient's email address.
  RECIPIENT_COLUMN: "Email",

  /**
   * Define short aliases for your column names in the MAILER_SHEET.
   * You can use {{fname}} in your email draft instead of {{First name}}.
   */
  PLACEHOLDER_ALIASES: {
    fname: "First name",
    lname: "Last name",
    email: "Email",
    fees: "Fees",
  },

  // This section maps a status value to a specific Gmail draft subject template
  // and the column where the 'sent' timestamp should be recorded.
  STATUS_MAPPINGS: {
    Registered: {
      draftSubject: "Registration Confirmation - {{eventName}}",
      timestampColumn: "Registration Email",
    },
    Reminded: {
      draftSubject: "Payment Reminder - {{eventName}}",
      timestampColumn: "Reminder Email",
    },
    Confirmed: {
      draftSubject: "Payment Confirmation - {{eventName}}",
      timestampColumn: "Payment Email",
    },
    Cancelled: {
      draftSubject: "Cancellation Confirmation - {{eventName}}",
      timestampColumn: "Cancellation Email",
    },
    Revoked: {
      draftSubject: "Cancellation - {{eventName}}",
      timestampColumn: "Revocation Email",
    },
  },
};

// =================================================================
// === SCRIPT SETUP & MENU =========================================
// =================================================================

/**
 * Creates a custom menu in the spreadsheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Mail Merge Automation")
    .addItem("Setup Edit Trigger", "createSpreadsheetEditTrigger")
    .addSeparator()
    .addItem("Send Bulk Emails for a Status", "sendBulkEmailForStatus") // New Feature
    .addSeparator()
    .addItem("Check Sent Status vs. Gmail", "checkGmailSentStatus")
    .addSeparator()
    .addItem("Debug Last Edit", "debugLastEdit")
    .addItem("Clear Cache & Re-read General Sheet", "clearCache")
    .addToUi();
}

/**
 * Creates an installable trigger for the handleEdit function.
 */
function createSpreadsheetEditTrigger() {
  const sheet = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === "handleEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger("handleEdit").forSpreadsheet(sheet).onEdit().create();
  SpreadsheetApp.getUi().alert(
    "Success!",
    "The automated email trigger has been set up. Emails will now be sent automatically when you change a status.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// =================================================================
// === NEW: BULK EMAIL SENDER ======================================
// =================================================================

/**
 * Sends a specified draft email to all users with a specified status.
 */
function sendBulkEmailForStatus() {
  const ui = SpreadsheetApp.getUi();

  // Prompt 1: Get the target status
  const statusPrompt = ui.prompt(
    "Step 1: Select Status",
    'Enter the exact status you want to send emails to (e.g., "Confirmed"):',
    ui.ButtonSet.OK_CANCEL
  );
  const statusButton = statusPrompt.getSelectedButton();
  const targetStatus = statusPrompt.getResponseText().trim();
  if (statusButton !== ui.Button.OK || !targetStatus) return;

  // Prompt 2: Get the draft subject
  const draftPrompt = ui.prompt(
    "Step 2: Select Draft",
    "Enter the exact subject of the Gmail draft you want to send:",
    ui.ButtonSet.OK_CANCEL
  );
  const draftButton = draftPrompt.getSelectedButton();
  const draftSubject = draftPrompt.getResponseText().trim();
  if (draftButton !== ui.Button.OK || !draftSubject) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.MAILER_SHEET
  );
  if (!sheet) {
    ui.alert(
      "Error",
      `Sheet with name "${CONFIG.MAILER_SHEET}" was not found. Please check the CONFIG.`,
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    const dataRange = sheet.getDataRange();
    const allData = dataRange.getValues();
    const rawHeaders = allData.shift(); // Get and remove header row
    const headers = rawHeaders.map((h) =>
      typeof h === "string" ? h.trim() : h
    );

    const statusColIdx = headers.indexOf(CONFIG.STATUS_COLUMN);
    if (statusColIdx === -1) {
      ui.alert(
        "Error",
        `Status column "${CONFIG.STATUS_COLUMN}" not found.`,
        ui.ButtonSet.OK
      );
      return;
    }

    const recipientsToEmail = [];
    allData.forEach((rowDataArray, index) => {
      if (rowDataArray[statusColIdx] === targetStatus) {
        const rowObject = headers.reduce((obj, header, i) => {
          let cellValue = rowDataArray[i] || "";
          if (typeof cellValue === "string") cellValue = cellValue.trim();
          obj[header] = cellValue;
          return obj;
        }, {});
        recipientsToEmail.push(rowObject);
      }
    });

    if (recipientsToEmail.length === 0) {
      ui.alert(
        "No Recipients Found",
        `No users with the status "${targetStatus}" were found.`,
        ui.ButtonSet.OK
      );
      return;
    }

    ui.alert(
      "Starting Bulk Send",
      `Found ${recipientsToEmail.length} recipient(s) with status "${targetStatus}". The script will now send the emails. This may take a moment.`,
      ui.ButtonSet.OK
    );

    const globalParameters = getGlobalParameters_();
    const emailTemplate = getGmailTemplateFromDrafts_(draftSubject);
    const sentLog = [];

    recipientsToEmail.forEach((rowObject) => {
      try {
        const recipient = rowObject[CONFIG.RECIPIENT_COLUMN];
        if (!recipient) return; // Skip if no email address

        const messageObject = fillInTemplateFromObject_(
          emailTemplate.message,
          rowObject,
          globalParameters,
          CONFIG.PLACEHOLDER_ALIASES
        );
        const cleanRecipient = String(recipient).replace(/\s/g, "");

        MailApp.sendEmail({
          to: cleanRecipient,
          subject: messageObject.subject,
          body: messageObject.text,
          htmlBody: messageObject.html,
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages,
        });
        sentLog.push(cleanRecipient);
      } catch (e) {
        Logger.log(
          `Bulk send failed for ${rowObject[CONFIG.RECIPIENT_COLUMN]}: ${
            e.message
          }`
        );
      }
    });

    // Display confirmation log
    const logMessage = `Bulk email send complete.\n\nDraft Sent: "${draftSubject}"\nTimestamp: ${new Date().toLocaleString()}\n\nSuccessfully sent to ${
      sentLog.length
    } recipient(s):\n\n${sentLog.join("\n")}`;
    const htmlOutput = HtmlService.createHtmlOutput(
      `<p>Bulk send complete for status "${targetStatus}".</p><textarea rows="15" cols="80" readonly>${logMessage}</textarea>`
    )
      .setWidth(600)
      .setHeight(350);
    ui.showModalDialog(htmlOutput, "Bulk Send Log");
  } catch (e) {
    ui.alert("An Error Occurred", e.message, ui.ButtonSet.OK);
  }
}

// =================================================================
// === CORE: AUTOMATED EMAIL TRIGGER LOGIC =========================
// =================================================================

function handleEdit(e) {
  if (!e) return;

  const sheet = e.source.getActiveSheet();
  const range = e.range;

  if (sheet.getName() !== CONFIG.MAILER_SHEET || range.getRow() <= 1) {
    return;
  }

  const rawHeaders = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  const headers = rawHeaders.map((h) => (typeof h === "string" ? h.trim() : h));

  const statusColIdx = headers.indexOf(CONFIG.STATUS_COLUMN);

  if (range.getColumn() === statusColIdx + 1) {
    const status = e.value;
    const rowDataArray = sheet
      .getRange(range.getRow(), 1, 1, sheet.getLastColumn())
      .getValues()[0];

    const rowObject = headers.reduce((obj, header, i) => {
      let cellValue = rowDataArray[i] || "";
      if (typeof cellValue === "string") {
        cellValue = cellValue.trim();
      }
      obj[header] = cellValue;
      return obj;
    }, {});

    CacheService.getScriptCache().put(
      "last_edited_row",
      JSON.stringify(rowObject),
      3600
    );
    processEmailTrigger(rowObject, status, sheet, range.getRow(), headers);
  }
}

function processEmailTrigger(rowObject, status, sheet, rowNum, headers) {
  const mapping = CONFIG.STATUS_MAPPINGS[status];
  if (!mapping) return;
  if (rowObject[mapping.timestampColumn]) return;

  try {
    const recipient = rowObject[CONFIG.RECIPIENT_COLUMN];
    if (!recipient) throw new Error("Recipient email address is missing.");

    const globalParameters = getGlobalParameters_();
    const emailTemplate = getGmailTemplateFromDrafts_(mapping.draftSubject);
    const messageObject = fillInTemplateFromObject_(
      emailTemplate.message,
      rowObject,
      globalParameters,
      CONFIG.PLACEHOLDER_ALIASES
    );

    const cleanRecipient = String(recipient).replace(/\s/g, "");

    MailApp.sendEmail({
      to: cleanRecipient,
      subject: messageObject.subject,
      body: messageObject.text,
      htmlBody: messageObject.html,
      attachments: emailTemplate.attachments,
      inlineImages: emailTemplate.inlineImages,
    });

    const timestampColIdx = headers.indexOf(mapping.timestampColumn);
    if (timestampColIdx !== -1) {
      sheet.getRange(rowNum, timestampColIdx + 1).setValue(new Date());
    }
  } catch (e) {
    Logger.log(`Error sending email for row ${rowNum}: ${e.message}`);
    SpreadsheetApp.getUi().alert(
      `An error occurred while sending an email: ${e.message}`
    );
  }
}

// =================================================================
// === HELPER & DEBUGGING FUNCTIONS ================================
// =================================================================

function clearCache() {
  CacheService.getScriptCache().remove("global_parameters");
  SpreadsheetApp.getUi().alert(
    "Cache Cleared",
    'The cache for the "General" sheet has been cleared. The script will now read the latest data on the next edit.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function debugLastEdit() {
  const ui = SpreadsheetApp.getUi();
  const lastRowJSON = CacheService.getScriptCache().get("last_edited_row");

  if (!lastRowJSON) {
    ui.alert(
      "No Edit Data Found",
      'Please edit a status in a row in the "Mailer" sheet first, then run this command again.',
      ui.ButtonSet.OK
    );
    return;
  }

  const rowData = JSON.parse(lastRowJSON);
  const globalData = getGlobalParameters_();

  let debugText = "--- GLOBAL DATA (from General sheet) ---\n";
  debugText += JSON.stringify(globalData, null, 2);
  debugText += "\n\n--- ROW DATA (from last edited row) ---\n";
  debugText += JSON.stringify(rowData, null, 2);

  const htmlOutput = HtmlService.createHtmlOutput(
    `<p>Data for the last edited row. Check that "Fees" exists in the ROW DATA below with the correct value.</p><textarea rows="20" cols="80" readonly>${debugText}</textarea>`
  )
    .setWidth(600)
    .setHeight(400);
  ui.showModalDialog(htmlOutput, "Last Edit Debug Data");
}

/**
 * Fetches global parameters from the "General" sheet and caches them.
 */

function getGlobalParameters_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("global_parameters");
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const generalSheet = ss.getSheetByName(CONFIG.GENERAL_SHEET);
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
        if (typeof value === "string") {
          value = value.replace(/\u00A0/g, " ").trim();
        }
        globalParams[cleanedAlias] = value;
      }
    }
  });

  cache.put("global_parameters", JSON.stringify(globalParams), 21600);
  return globalParams;
}

/**
 * Get a Gmail draft message by matching the subject line.
 */
function getGmailTemplateFromDrafts_(subject_line) {
  try {
    const drafts = GmailApp.getDrafts();
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
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
      {}
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
      `Oops - can't find or process Gmail draft with subject "${subject_line}". Error: ${e.message}`
    );
  }

  function subjectFilter_(subject_line) {
    return function (element) {
      if (element.getMessage().getSubject() === subject_line) {
        return element;
      }
    };
  }
}

/**
 * **REWRITTEN FUNCTION**
 * Fills a template by directly replacing placeholders in the subject and body strings.
 * This method is robust and preserves emojis, HTML formatting, and special characters.
 *
 * @param {object} template The message object {subject, text, html}.
 * @param {object} rowData The data object for the current row.
 * @param {object} globalData The global parameters from the "General" sheet.
 * @param {object} aliases The mapping of short names to full column names.
 * @return {object} A new message object with all placeholders replaced with data.
 */
function fillInTemplateFromObject_(template, rowData, globalData, aliases) {
  // Create a unified data object for replacement.
  const replacementData = {};
  // 1. Add aliased row data (e.g., "fees" -> value from "Fees" column)
  for (const alias in aliases) {
    const realColumnName = aliases[alias];
    if (rowData.hasOwnProperty(realColumnName)) {
      replacementData[alias] = rowData[realColumnName];
    }
  }
  // 2. Add all raw row data (for direct {{Column Name}} access)
  for (const columnName in rowData) {
    replacementData[columnName] = rowData[columnName];
  }
  // 3. Add global data, overwriting any previous keys. This gives globals the highest priority.
  for (const globalKey in globalData) {
    replacementData[globalKey] = globalData[globalKey];
  }
  // Get the raw subject and body from the template
  let subject = template.subject;
  let htmlBody = template.html;
  let textBody = template.text;
  // Iterate through the unified data and perform replacements
  for (const key in replacementData) {
    let value = replacementData[key] || "";
    // Special handling for email fields to remove all spaces
    if (key.toLowerCase().includes("email")) {
      value = String(value).replace(/\s/g, "");
    }

    // Create a regular expression to replace all instances of the placeholder
    // The 'g' flag ensures all occurrences are replaced.
    const placeholderRegex = new RegExp("{{" + key + "}}", "g");

    subject = subject.replace(placeholderRegex, value);
    htmlBody = htmlBody.replace(placeholderRegex, value);
    textBody = textBody.replace(placeholderRegex, value);
  }
  // Return the new message object with the replaced content
  return {
    subject: subject,
    html: htmlBody,
    text: textBody,
  };
}
// =================================================================
// === GMAIL SENT STATUS CHECKER (No changes needed here) ==========
// =================================================================

function checkGmailSentStatus() {
  const ui = SpreadsheetApp.getUi();
  const columnPrompt = ui.prompt(
    "Step 1: Select Column",
    "Enter the exact name of the column containing the email addresses you want to check:",
    ui.ButtonSet.OK_CANCEL
  );
  const columnButton = columnPrompt.getSelectedButton();
  const columnToCheck = columnPrompt.getResponseText();
  if (columnButton !== ui.Button.OK || !columnToCheck) return;

  const subjectPrompt = ui.prompt(
    "Step 2: Specify Subject",
    'Enter the exact subject line of the sent email you want to search for (e.g., "Payment Confirmation - Wieda Summer Hike 2025"):',
    ui.ButtonSet.OK_CANCEL
  );
  const subjectButton = subjectPrompt.getSelectedButton();
  const subjectToCheck = subjectPrompt.getResponseText();
  if (subjectButton !== ui.Button.OK || !subjectToCheck) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.MAILER_SHEET
  );
  if (!sheet) {
    ui.alert(
      "Error",
      `Sheet with name "${CONFIG.MAILER_SHEET}" was not found. Please check the CONFIG.`,
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const emailColIdx = headers.indexOf(columnToCheck);
    if (emailColIdx === -1) {
      ui.alert(
        "Error",
        `Column "${columnToCheck}" was not found.`,
        ui.ButtonSet.OK
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
          .filter((email) => typeof email === "string" && email.includes("@"))
      ),
    ];
    if (emailsToCheck.length === 0) {
      ui.alert(
        "No Emails Found",
        "Could not find any valid email addresses in the specified column.",
        ui.ButtonSet.OK
      );
      return;
    }

    ui.alert(
      "Starting Check",
      `Checking ${emailsToCheck.length} unique email addresses for the subject "${subjectToCheck}" from the last two months. This may take some time...`,
      ui.ButtonSet.OK
    );

    const twoMonthsAgo = new Date();
    twoMonthsAgo.setMonth(twoMonthsAgo.getMonth() - 2);
    const dateQuery = `${twoMonthsAgo.getFullYear()}/${
      twoMonthsAgo.getMonth() + 1
    }/${twoMonthsAgo.getDate()}`;

    const notSentList = [];
    emailsToCheck.forEach((email) => {
      const query = `to:(${email}) subject:("${subjectToCheck}") after:${dateQuery}`;
      const threads = GmailApp.search(query, 0, 1);
      if (threads.length === 0) {
        notSentList.push(email);
      }
    });

    if (notSentList.length === 0) {
      ui.alert(
        "All Clear!",
        `It appears an email with the subject "${subjectToCheck}" has been sent to every address on the list within the last two months.`,
        ui.ButtonSet.OK
      );
    } else {
      const htmlOutput = HtmlService.createHtmlOutput(
        `<p>Found ${
          notSentList.length
        } emails that were not sent an email with the subject "${subjectToCheck}" in the last two months:</p><textarea rows="15" cols="60" readonly>${notSentList.join(
          "\n"
        )}</textarea>`
      )
        .setWidth(500)
        .setHeight(350);
      ui.showModalDialog(htmlOutput, "Emails Not Sent Recently");
    }
  } catch (e) {
    ui.alert("An Error Occurred", e.message, ui.ButtonSet.OK);
  }
}
