/**
 * Deployment steps:
 * script.new → paste → Deploy → New deployment → type "Web app" →
 * Execute as "Me" → Who has access "Anyone" → Deploy → authorize →
 * copy the /exec URL.
 *
 * After editing the script, redeploy with "New deployment" or bump the version.
 */

/**
 * Screening layers: honeypot, human timing, JavaScript token, link flooding,
 * solicitation terms, and per-email/global rate limits. Blocked requests are
 * written to the "Blocked" tab; check that tab if an expected lead never arrives.
 */

const NOTIFY_EMAIL = "saint@shcmarketing.info";

// Google Sheet ID from its URL; clear this value to disable Sheet logging.
const SHEET_ID = "1ioLCHQLXtl2li5yt70B0gcSsDqpHLCXOuRJ_hYG0psM";

const MAX_FIELD_LENGTH = 2000;

/**
 * A phrase blocks on its own only when it targets the recipient's website or
 * offers help to the recipient, or when it is an unmistakable spam/scam term.
 * SHC's prospects are service-business owners who naturally describe their own
 * companies with offering language, so those self-descriptions are service
 * topics and cannot satisfy both halves of the co-occurrence rule by themselves.
 * Service topics otherwise require a URL or first-person offering construction;
 * "seo" is checked as a whole word.
 */
const PITCH_PHRASES = [
  "i came across your website",
  "came across your website",
  "i visited your website",
  "visited your website",
  "i noticed your website",
  "noticed your website",
  "came across your site",
  "stumbled upon your website",
  "stumbled across your website",
  "hope this email finds you",
  "hope this finds you well",
  "i can help you",
  "we can help you",
  "we would like to offer",
  "we'd like to offer",
  "i would like to offer",
  "i'd like to offer",
  "let me know if you are interested",
  "let me know if you're interested",
  "are you interested in",
  "would you be interested",
  "backlink",
  "link building",
  "guest post",
  "guest blogging",
  "guest posting",
  "our team of developers",
  "dedicated developers",
  "hire developers",
  "hire our",
  "dear sir",
  "dear madam",
  "dear sir/madam",
  "dear sir or madam",
  "business proposal",
  "kindly revert",
  "kindly reply",
  "crypto",
  "bitcoin",
  "forex",
  "casino",
  "loan offer",
  "nft",
  "telegram",
  "whatsapp me",
  "whatsapp us",
  "contact me on whatsapp",
  "skype id",
  "skype username"
];

const SERVICE_TOPICS = [
  "seo",
  "search engine optimi",
  "web design",
  "website design",
  "website redesign",
  "web development",
  "app development",
  "mobile app",
  "software development",
  "digital marketing",
  "online marketing",
  "marketing services",
  "lead generation",
  "rank higher",
  "ranking",
  "first page of google",
  "top of google",
  "increase traffic",
  "website traffic",
  "social media management",
  "ppc",
  "google ads",
  // Sender self-descriptions: neutral unless another pitch signal is present.
  "we offer",
  "we provide",
  "we specialize",
  "we specialise",
  "our services include",
  "our company provides",
  "our agency provides",
  "we are a leading",
  "we are an agency",
  "we are a team of",
  "outsourcing",
  "white label",
  "free audit",
  "complimentary audit",
  "free quote",
  "free consultation offer",
  "no obligation quote",
  "affordable price",
  "affordable rate",
  "affordable rates",
  "cheap rate",
  "cheap rates",
  "best price guaranteed",
  "best prices guaranteed",
  "guarantee results",
  "guaranteed results",
  "money back guarantee"
];

const SELF_DESCRIPTION_TOPIC_START_INDEX = SERVICE_TOPICS.indexOf("we offer");

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
      companyFax: cleanValue_(submission.companyFax),
      elapsedMs: Object.prototype.hasOwnProperty.call(submission, "elapsedMs")
        ? cleanValue_(submission.elapsedMs)
        : null,
      jsToken: cleanValue_(submission.jsToken)
    };

    var screenReason = screenSubmission_(fields);

    if (screenReason) {
      logBlockedSubmission_(fields, screenReason);
      return jsonResponse_({ ok: true });
    }

    var summary = [
      "Full name: " + fields.fullName,
      "Email: " + fields.email,
      "Company: " + fields.company,
      "Preferred day: " + fields.preferredDay,
      "Notes: " + (fields.notes || "Not provided")
    ].join("\n");

    var htmlBody;

    try {
      htmlBody = buildIntakeEmailHtml_(fields, new Date());
    } catch (htmlError) {
      // Plain-text delivery must still succeed if HTML construction fails.
    }

    var message = {
      to: NOTIFY_EMAIL,
      subject: fields.company
        ? "Working Session Request — " + fields.company
        : "Working Session Request",
      body: summary
    };

    if (htmlBody) {
      message.htmlBody = htmlBody;
    }

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

function buildIntakeEmailHtml_(fields, submittedAt) {
  var fullName = escapeHtml_(fields.fullName);
  var email = escapeHtml_(fields.email);
  var company = escapeHtml_(fields.company);
  var preferredDay = escapeHtml_(fields.preferredDay);
  var notes = fields.notes
    ? escapeHtml_(fields.notes).replace(/\r\n|\r|\n/g, "<br>")
    : '<span style="color:#606265;font-style:italic;">Not provided</span>';
  var submittedAtText = escapeHtml_(formatSubmissionTimestamp_(submittedAt));
  var hasValidEmail = looksLikeEmail_(fields.email);
  var emailHtml = email;
  var replyButtonHtml = "";
  var preheaderParts = [];

  if (fields.company) {
    preheaderParts.push(fields.company);
  }

  if (fields.fullName) {
    preheaderParts.push(fields.fullName);
  }

  var preheader = preheaderParts.length
    ? preheaderParts.join(" — ") + " requested a working session"
    : "New working session request";

  if (hasValidEmail) {
    var mailtoHref = escapeHtml_("mailto:" + fields.email);
    var replyHref = escapeHtml_(
      "mailto:" +
        fields.email +
        "?subject=" +
        encodeURIComponent("Re: Your Working Session Request")
    );
    var firstName = fields.fullName
      ? fields.fullName.split(/\s+/)[0]
      : "";
    var replyLabel = firstName
      ? "Reply to " + escapeHtml_(firstName)
      : "Reply";

    emailHtml =
      '<a href="' +
      mailtoHref +
      '" style="color:#d9232e;text-decoration:underline;">' +
      email +
      "</a>";
    replyButtonHtml =
      '<table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;">' +
      '<tr><td style="padding-top:24px;">' +
      '<table role="presentation" border="0" cellspacing="0" cellpadding="0" style="border-collapse:separate;">' +
      '<tr><td bgcolor="#d9232e" style="background-color:#d9232e;border-radius:6px;">' +
      '<a href="' +
      replyHref +
      '" style="display:inline-block;padding:12px 24px;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:15px;line-height:20px;font-weight:bold;color:#ffffff;text-decoration:none;border-radius:6px;">' +
      replyLabel +
      "</a></td></tr></table></td></tr></table>";
  }

  return (
    '<!doctype html><html><body style="margin:0;padding:0;background-color:#f8f6f0;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">' +
    '<div style="display:none;font-size:1px;color:#f8f6f0;max-height:0;overflow:hidden;">' +
    escapeHtml_(preheader) +
    "</div>" +
    '<table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#f8f6f0" style="width:100%;border-collapse:collapse;background-color:#f8f6f0;">' +
    '<tr><td align="center" style="padding:24px 12px;">' +
    '<table role="presentation" width="600" border="0" cellspacing="0" cellpadding="0" align="center" style="width:100%;max-width:600px;border-collapse:collapse;">' +
    '<tr><td bgcolor="#090b0d" style="padding:20px 28px;background-color:#090b0d;">' +
    '<table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;">' +
    '<tr><td width="60" valign="middle" style="width:60px;vertical-align:middle;">' +
    '<img src="https://shcmarketing.info/assets/shc-logo-email.png" width="44" height="44" border="0" alt="SHC" style="display:block;width:44px;height:44px;border:0;">' +
    '</td><td valign="middle" style="vertical-align:middle;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:17px;line-height:22px;font-weight:bold;letter-spacing:2px;color:#ffffff;text-transform:uppercase;">SHC CONSULTING</td></tr>' +
    "</table></td></tr>" +
    '<tr><td height="4" bgcolor="#d9232e" style="height:4px;font-size:0;line-height:4px;background-color:#d9232e;">&nbsp;</td></tr>' +
    '<tr><td style="padding:24px 0 0 0;">' +
    '<table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#ffffff" style="width:100%;border-collapse:separate;background-color:#ffffff;border:1px solid #d9d6cf;">' +
    '<tr><td style="padding:30px 28px;">' +
    '<table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;">' +
    '<tr><td style="font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:12px;line-height:16px;font-weight:bold;letter-spacing:1.5px;color:#d9232e;text-transform:uppercase;">NEW WORKING SESSION REQUEST</td></tr>' +
    '<tr><td style="padding-top:8px;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:22px;line-height:28px;font-weight:bold;color:#25282a;">Lead details</td></tr>' +
    '<tr><td height="10" style="height:10px;font-size:0;line-height:10px;">&nbsp;</td></tr></table>' +
    '<table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;">' +
    '<tr><td width="35%" valign="top" style="width:35%;padding:12px 12px 12px 0;border-bottom:1px solid #efede7;vertical-align:top;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:11px;line-height:16px;font-weight:bold;letter-spacing:1px;color:#606265;text-transform:uppercase;">Full name</td>' +
    '<td width="65%" valign="top" style="width:65%;padding:12px 0;border-bottom:1px solid #efede7;vertical-align:top;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:16px;line-height:23px;color:#111315;">' +
    fullName +
    "</td></tr>" +
    '<tr><td width="35%" valign="top" style="width:35%;padding:12px 12px 12px 0;border-bottom:1px solid #efede7;vertical-align:top;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:11px;line-height:16px;font-weight:bold;letter-spacing:1px;color:#606265;text-transform:uppercase;">Email</td>' +
    '<td width="65%" valign="top" style="width:65%;padding:12px 0;border-bottom:1px solid #efede7;vertical-align:top;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:16px;line-height:23px;color:#111315;word-break:break-word;">' +
    emailHtml +
    "</td></tr>" +
    '<tr><td width="35%" valign="top" style="width:35%;padding:12px 12px 12px 0;border-bottom:1px solid #efede7;vertical-align:top;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:11px;line-height:16px;font-weight:bold;letter-spacing:1px;color:#606265;text-transform:uppercase;">Company</td>' +
    '<td width="65%" valign="top" style="width:65%;padding:12px 0;border-bottom:1px solid #efede7;vertical-align:top;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:16px;line-height:23px;color:#111315;">' +
    company +
    "</td></tr>" +
    '<tr><td width="35%" valign="top" style="width:35%;padding:12px 12px 12px 0;border-bottom:1px solid #efede7;vertical-align:top;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:11px;line-height:16px;font-weight:bold;letter-spacing:1px;color:#606265;text-transform:uppercase;">Preferred day</td>' +
    '<td width="65%" valign="top" style="width:65%;padding:12px 0;border-bottom:1px solid #efede7;vertical-align:top;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:16px;line-height:23px;color:#111315;">' +
    preferredDay +
    "</td></tr>" +
    '<tr><td width="35%" valign="top" style="width:35%;padding:12px 12px 0 0;vertical-align:top;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:11px;line-height:16px;font-weight:bold;letter-spacing:1px;color:#606265;text-transform:uppercase;">Notes</td>' +
    '<td width="65%" valign="top" style="width:65%;padding:12px 0 0 0;vertical-align:top;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:16px;line-height:23px;color:#111315;word-break:break-word;">' +
    notes +
    "</td></tr></table>" +
    replyButtonHtml +
    "</td></tr></table></td></tr>" +
    '<tr><td align="center" style="padding:20px 20px 4px 20px;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;font-size:12px;line-height:18px;color:#606265;">Submitted via the working-session intake form on <a href="https://shcmarketing.info/" style="color:#ad1821;text-decoration:underline;">shcmarketing.info</a><br>Submitted ' +
    submittedAtText +
    "</td></tr>" +
    "</table></td></tr></table></body></html>"
  );
}

function formatSubmissionTimestamp_(value) {
  return Utilities.formatDate(
    value,
    Session.getScriptTimeZone(),
    "MMMM d, yyyy 'at' h:mm a z"
  );
}

function screenSubmission_(fields) {
  if (fields.companyFax) {
    return "honeypot";
  }

  if (fields.elapsedMs !== null) {
    var elapsedMs = Number(fields.elapsedMs);

    if (isFinite(elapsedMs) && elapsedMs < 4000) {
      return "too-fast";
    }
  }

  if (fields.jsToken !== "shc-intake-v1") {
    return "no-js-token";
  }

  var linkText = [fields.notes, fields.company, fields.fullName].join(" ");
  var urlMatches = linkText.match(/(?:https?:\/\/|www\.)\S+/gi) || [];

  if (urlMatches.length >= 2) {
    return "link-spam";
  }

  var combinedText = [fields.fullName, fields.company, fields.notes].join(" ");
  var normalizedText = combinedText.toLowerCase();
  var hasPitchPhrase = false;

  for (var i = 0; i < PITCH_PHRASES.length; i += 1) {
    if (normalizedText.indexOf(PITCH_PHRASES[i]) !== -1) {
      hasPitchPhrase = true;
      break;
    }
  }

  if (hasPitchPhrase) {
    return "solicitation";
  }

  // "seo" remains a whole-word match; other service topics are substrings.
  var hasServiceTopic = /\bseo\b/i.test(combinedText);
  var hasOfferingContextTopic = hasServiceTopic;

  for (var j = 0; j < SERVICE_TOPICS.length; j += 1) {
    if (
      SERVICE_TOPICS[j] !== "seo" &&
      normalizedText.indexOf(SERVICE_TOPICS[j]) !== -1
    ) {
      hasServiceTopic = true;

      if (j < SELF_DESCRIPTION_TOPIC_START_INDEX) {
        hasOfferingContextTopic = true;
      }
    }
  }

  var hasOfferingConstruction = /\b(?:we|i)\s+(?:can|could|will|would|offer|offers|provide|provides|specialize|specialise|deliver|do)\b/i.test(
    combinedText
  );

  if (
    (hasServiceTopic && urlMatches.length >= 1) ||
    (hasOfferingContextTopic && hasOfferingConstruction) ||
    (hasServiceTopic && hasPitchPhrase)
  ) {
    return "solicitation";
  }

  try {
    var cache = CacheService.getScriptCache();
    var emailDigest = Utilities.base64EncodeWebSafe(
      Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        fields.email.toLowerCase()
      )
    ).replace(/=+$/, "");
    var now = Date.now();
    var emailWindow = Math.floor(now / 21600000);
    var globalWindow = Math.floor(now / 3600000);
    var emailKey = "intake-email:" + emailDigest + ":" + emailWindow;
    var globalKey = "intake-global:" + globalWindow;
    var emailCount = parseInt(cache.get(emailKey) || "0", 10);
    var globalCount = parseInt(cache.get(globalKey) || "0", 10);

    emailCount = isNaN(emailCount) ? 0 : emailCount;
    globalCount = isNaN(globalCount) ? 0 : globalCount;

    if (emailCount >= 3) {
      return "rate-limit-email";
    }

    if (globalCount >= 30) {
      return "rate-limit-global";
    }

    cache.put(emailKey, String(emailCount + 1), 21600);
    cache.put(globalKey, String(globalCount + 1), 3600);
  } catch (cacheError) {
    // Cache failures must never block a legitimate submission.
  }

  return null;
}

function logBlockedSubmission_(fields, reason) {
  try {
    if (!SHEET_ID) {
      return;
    }

    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    var blockedSheet = spreadsheet.getSheetByName("Blocked");

    if (!blockedSheet) {
      blockedSheet = spreadsheet.insertSheet("Blocked");
    }

    if (blockedSheet.getLastRow() === 0) {
      blockedSheet.appendRow([
        "Timestamp",
        "Reason",
        "Full name",
        "Email",
        "Company",
        "Preferred day",
        "Notes"
      ]);
    }

    blockedSheet.appendRow([
      new Date(),
      reason,
      fields.fullName,
      fields.email,
      fields.company,
      fields.preferredDay,
      fields.notes
    ]);
  } catch (blockedLogError) {
    // A blocked-log failure must not expose screening or change the response.
  }
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
    ContentService.MimeType.JSON
  );
}
