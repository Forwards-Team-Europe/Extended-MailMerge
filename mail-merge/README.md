# 📧 Mail Merge Automation for Google Sheets & Gmail

A Google Apps Script project for managing event registrations in Google Sheets. This script automates sending personalized bulk emails by using Gmail drafts as templates, triggered by status changes made directly in the sheet. It's designed to be powerful, flexible, and easy to manage for any event.

---

## 🎯 Who is this for?

This platform is designed for event organizers who are currently managing their operations manually and are looking for a more efficient, automated solution. It is ideal for you if:

- **You rely on Google Forms:** You currently use Google Forms to collect event registrations.
- **You manage emails manually:** You use Gmail to send emails and rely on standardized drafts for communication.
- **You have complex pricing:** You offer different price tiers for your participants (e.g., discounts for families, students, or early birds).
- **You need personalized bulk messaging:** You want to send bulk emails to your participants but need them to include personalized details.
- **You track status manually:** You are currently tracking participant statuses (e.g., cancellations, delayed payments) manually and need a better way to manage these changes.

---

## ✨ Key Features

- **✍️ Status-Driven Automation:** Automatically sends a specific email when a participant's status is changed (e.g., from "Registered" to "Confirmed").
- **📝 Gmail Draft Templates:** Uses your pre-written Gmail drafts as rich HTML email templates, complete with images and attachments.
- **🌍 Centralized Configuration:** Manages all event details (names, dates, fees, email subjects) from a central "General" sheet for easy updates without touching the code.
- **🔗 Flexible Placeholders:** Use placeholders for both global data (like `{{eventName}}`) and row-specific data (like `{{fname}}` or `{{fees}}`) in any email.
- **🛠️ Built-in Utilities:** Includes tools to check who hasn't received an email, debug data issues, and clear the script's cache.

---

## ⚙️ How It Works

This system uses two main Google Sheets and a Google Apps Script to create a powerful mail merge workflow:

1.  **`General` Sheet:** This sheet acts as your master control panel. You define all your event-wide parameters here, such as the event name, dates, fees, and other details you want to use in your emails.
2.  **`Mailer` Sheet:** This is your operational sheet. It contains the list of participants, their details, and a special `Status` column. When you change the value in the `Status` column for any participant, the script triggers.
3.  **Gmail Drafts:** You create email templates as simple drafts in your Gmail account. You can use placeholders in both the subject and the body.
4.  **The Script:** The Apps Script is the engine that connects everything. When a status is changed, it finds the correct email draft, combines the global data from the `General` sheet with the specific participant's data from the `Mailer` sheet, and sends a personalized email.

---

## 🚀 Setup Instructions

Follow these steps carefully to get the system up and running.

### 1. 📝 Set Up Your Google Sheets

You need two sheets in the same Google Sheets document:

**A. The `General` Sheet:**

- Create a new sheet and name it **`General`**.
- Set up three columns: `Parameter Name`, `Value`, and `Email Alias`.
- Populate this sheet with all your event-wide data. The script will read the `Value` (Column B) and make it available using the `Email Alias` (Column C).

*Example:*
| Parameter Name | Value | Email Alias |
| :--- | :--- | :--- |
| Event Name | Photography Workshop | `{{eventName}}` |
| Registration Deadline | 01.01.2021 | `{{regDeadline}}` |
| Organizer Email | your@email.com | `{{orgEmail}}` |
| Participation Fees | 120$ | `{{fees}}` |

**B. The `Mailer` Sheet:**

- Create a second sheet and name it **`Mailer`**. This is where your participant data will live (you can use an `ARRAYFORMULA` to pull data from another sheet, like a Google Form response sheet).
- It must contain a **Status** column (the name can be configured in the script). This column should use Data Validation to create a dropdown menu with your defined statuses (e.g., "Registered", "Confirmed", "Cancelled", etc.).
- It must also contain columns for any data you want to use in your emails, like `First name`, `Last name`, `Email`, and `Fees`.

### 2. ✍️ Create Your Gmail Drafts

- In your Gmail account, compose a new email for each status.
- **The Subject Line is Critical:** The subject of the draft must **exactly match** the `draftSubject` defined in the script's configuration. You can use global placeholders here.
    - *Example Subject:* `Registration Confirmation - {{eventName}}`
- **Use Placeholders:** Write your email body using both global placeholders (from the `General` sheet) and row-specific placeholders (from the `Mailer` sheet).
    - *Example Body:* `Hello {{fname}}, thank you for registering for the {{eventName}}! Your total fee is {{fees}}.`
- Save each email as a draft. Do not send it!

### 3. 💻 Install and Configure the Script

1.  **Open the Script Editor:** In your Google Sheet, go to `Extensions` > `Apps Script`.
2.  **Paste the Code:** Delete any existing code and paste the entire script from [`ExtendedMailMerge.gs`](https://github.com/Forwards-Team-Europe/Extended-MailMerge/blob/main/ExtendedMailMerge.gs).
3.  **Configure the `CONFIG` Block:** Carefully review the `CONFIG` object at the top of the script and ensure all the sheet names, column names, and status mappings match your setup exactly.
4.  **Save the Project:** Click the "Save project" icon.

### 4. 🔐 Authorize the Script

1.  **Run the Setup Trigger:** In the script editor, select the `createSpreadsheetEditTrigger` function from the dropdown menu and click **Run**.
2.  **Authorization Required:** A dialog box will appear. Click "Review permissions".
3.  **Choose Your Account:** Select the Google account you want to run the script from.
4.  **Advanced Warning:** You will likely see a "Google hasn’t verified this app" screen. This is normal. Click **"Advanced"**, then click **"Go to [Your Project Name] (unsafe)"**.
5.  **Grant Permissions:** Review the permissions the script needs (to manage your sheets and send email on your behalf) and click **"Allow"**.
6.  You should see a "Success!" message in your spreadsheet. The automation is now active!

---

## 🕹️ How to Use the Automation Menu

After reloading your spreadsheet, you will see a new **"Mail Merge Automation"** menu.

- **Setup Edit Trigger:** Run this once during setup to activate the automation.
- **Check Sent Status vs. Gmail:** A utility to check who on your list has not received a specific email. It will ask for the email column and the subject line to check against your Gmail Sent folder.
- **Debug Last Edit:** An essential troubleshooting tool. After editing a status, run this to see the exact data the script is trying to use for its placeholders. This helps you find typos in column names or other data issues.
- **Clear Cache & Re-read General Sheet:** The script saves a copy of your `General` sheet data for performance. If you make changes to the `General` sheet, run this to force the script to read the new data immediately.

---

## 💡 Placeholders Guide

You can use two types of placeholders in your Gmail draft subjects and bodies:

1.  **Global Placeholders:** These pull data from your `General` sheet.
    - **Format:** `{{alias}}` (e.g., `{{eventName}}`, `{{regDeadline}}`)
    - Defined by the `Email Alias` column in your `General` sheet.
2.  **Row-Specific Placeholders:** These pull data from the participant's row in the `Mailer` sheet.
    - **Format:** `{{alias}}` or `{{Column Name}}`
    - You can use short aliases defined in the `PLACEHOLDER_ALIASES` section of the script (e.g., `{{fname}}` for the "First name" column) or use the full column name directly (e.g., `{{First name}}`).

The script is smart enough to use data from both sheets to compose the final email!

---

## ⚠️ Troubleshooting

- **Emails Not Sending:**
    1.  Ensure you have run the `Setup Edit Trigger` and authorized the script.
    2.  Check that the status you are selecting in the dropdown exactly matches a status defined in the `STATUS_MAPPINGS` in the script.
    3.  Check that the subject of your Gmail draft exactly matches the `draftSubject` for that status in the script.
- **Placeholders Not Working (`{{fees}}` shows up as `{{fees}}`):**
    1.  Run the `Debug Last Edit` tool immediately after an edit.
    2.  Examine the `ROW DATA` section. Check if the column header (e.g., "Fees") is spelled correctly and has no extra spaces.
    3.  Examine the `GLOBAL DATA` section. Make sure you haven't accidentally defined a row-specific placeholder (like `{{fees}}`) in your `General` sheet.
- **Changes to `General` Sheet Not Appearing:**
    1.  Run the `Clear Cache & Re-read General Sheet` command from the menu.

---

## 📬 Emails Show in "Sent" but Never Arrive (Deliverability)

If your messages appear in Gmail's **Sent** folder but recipients never receive them (not even in Spam), the script did its job — Google **accepted and sent** the message, and the **receiving server then dropped or rejected it.** This is almost always a **sender-authentication** problem with your domain, *not* a script bug. No code change can fix it; it must be fixed in DNS / Google Admin.

**Fix it in this order:**

1.  **Verify authentication.** Send one merge email to a Gmail address you own → open it → **⋮ → Show original**. You must see **SPF: PASS**, **DKIM: PASS**, and **DMARC: PASS**. (Or send to `check-auth@verifier.port25.com` for a full report.) Any `FAIL`/`none` is your culprit.
2.  **Enable DKIM** in Google Admin: *Apps → Google Workspace → Gmail → Authenticate email*. Generate the key and add the published `TXT` record to your DNS. **Workspace does not enable real DKIM by default** — this is the most common cause.
3.  **Check SPF.** Your domain's DNS needs a `TXT` record containing `include:_spf.google.com`.
4.  **Check DMARC.** A `p=reject` or `p=quarantine` policy will silently drop unauthenticated mail. Make sure SPF + DKIM pass *before* tightening DMARC.
5.  **Watch volume & reputation.** Since Feb 2024, Gmail/Yahoo require bulk senders to pass SPF + DKIM + DMARC. Sudden large sends from a "cold" domain also hurt. Sending limits are fixed by Google (Workspace **1,500/day**, consumer Gmail **500/day**) and **cannot be raised by the script**. For higher volume or reliable inbox placement, use a dedicated service (Brevo, SendGrid, Amazon SES).

**Built-in tools to diagnose & confirm sending:**

- **Send Log sheet** — every send is recorded in a `Send Log` tab with `SENT` / `FAILED` / `SKIPPED` status and the exact error. No more silent failures.
- **Scan for Bounced Emails** — scans your inbox for delivery-failure notices and flags exactly which participants hard-bounced (logged as `BOUNCED`). This is your real "did it arrive?" check. *Note: silent spam-foldering and DKIM/DMARC drops do not bounce, so they won't appear here — use step 1 above for those.*
- **Check Email Quota** — shows how many sends you have left today.
- **Set Sender Name & Reply-To** — sets a friendly `From` name and a real reply-to address, which improves recognition and inbox placement.

---

## 📜 License

This project is licensed under the Apache License, Version 2.0. The original concept is credited to [Martin Hawksey's mail-merge automation sample](https://developers.google.com/apps-script/samples/automations/mail-merge).
