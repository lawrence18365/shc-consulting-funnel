/**
 * Deployment steps:
 * script.new → paste → Deploy → New deployment → type "Web app" →
 * Execute as "Me" → Who has access "Anyone" → Deploy → authorize →
 * copy the /exec URL.
 *
 * After editing the script, redeploy with "New deployment" or bump the version.
 */

const NOTIFY_EMAIL = "saint@shcmarketing.info";

// Google Sheet ID from its URL; clear this value to disable Sheet logging.
const SHEET_ID = "1ioLCHQLXtl2li5yt70B0gcSsDqpHLCXOuRJ_hYG0psM";

const MAX_FIELD_LENGTH = 2000;

function doPost(e) {
  try {
    var submission;

    try {
      submission = JSON.parse(e.postData.contents);
    } catch (parseError) {
      submission = (e && e.parameter) || {};
    }

    submission = submission && typeof submission === "object" ? submission : {};

    var fields = {
      fullName: cleanValue_(submission.fullName),
      email: cleanValue_(submission.email),
      company: cleanValue_(submission.company),
      preferredDay: cleanValue_(submission.preferredDay),
      notes: cleanValue_(submission.notes),
      companyFax: cleanValue_(submission.companyFax)
    };

    // Silently accept honeypot submissions without emailing or saving them.
    if (fields.companyFax) {
      return jsonResponse_({ ok: true });
    }

    var summary = [
      "Full name: " + fields.fullName,
      "Email: " + fields.email,
      "Company: " + fields.company,
      "Preferred day: " + fields.preferredDay,
      "Notes: " + (fields.notes || "Not provided")
    ].join("\n");

    var message = {
      to: NOTIFY_EMAIL,
      subject: fields.company
        ? "Working Session Request — " + fields.company
        : "Working Session Request",
      body: summary
    };

    if (looksLikeEmail_(fields.email)) {
      message.replyTo = fields.email;
    }

    MailApp.sendEmail(message);

    if (SHEET_ID) {
      try {
        var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

        if (sheet.getLastRow() === 0) {
          sheet.appendRow([
            "Timestamp",
            "Full name",
            "Email",
            "Company",
            "Preferred day",
            "Notes"
          ]);
        }

        sheet.appendRow([
          new Date(),
          fields.fullName,
          fields.email,
          fields.company,
          fields.preferredDay,
          fields.notes
        ]);
      } catch (sheetError) {
        // Email delivery succeeded; a Sheet error must not change the response.
      }
    }

    return jsonResponse_({ ok: true });
  } catch (unexpectedError) {
    return jsonResponse_({ ok: false });
  }
}

function doGet() {
  return jsonResponse_({ ok: true, service: "shc-intake" });
}

function cleanValue_(value) {
  if (value === undefined || value === null) {
    return "";
  }

  return String(value).trim().slice(0, MAX_FIELD_LENGTH);
}

function looksLikeEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
    ContentService.MimeType.JSON
  );
}
