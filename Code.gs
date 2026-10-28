/**
 * ASTROMUL — Final Google Apps Script Backend
 * Apps Script HTML file must be named: index
 * Spreadsheet: ASTROMUL Registrations
 */
const SPREADSHEET_ID = "12AU_DqW4KwlL-JkOIIHroWqpIqxptBULUOUuKUilsvY";
const SHEET_NAME = "ASTROMUL Registrations";

const HEADERS = [
  "SUBMISSION DATE",
  "FULL NAME",
  "MOBILE / WHATSAPP",
  "FULL RESIDENTIAL ADDRESS",
  "DOES ASTROLOGY WORK",
  "SPECIALTIES",
  "CONSULTATION MODES",
  "CURRENTLY CONSULTS PEOPLE",
  "WANTS TO JOIN ASTROMUL",
  "PROFILE / SOCIAL LINK",
  "CONSENT",
  "ASTROLOGY CERTIFICATE / DIPLOMA",
  "LANGUAGES KNOWN"
];

const ALLOWED_YES_NO = ["हाँ", "नहीं"];
const ALLOWED_SPECIALTIES = [
  "वैदिक ज्योतिष एवं कुंडली", "विवाह मिलान", "अंक ज्योतिष", "हस्तरेखा",
  "प्रश्न ज्योतिष", "रत्न एवं रुद्राक्ष", "वास्तु", "लाल किताब", "टैरो एवं अन्य"
];
const ALLOWED_LANGUAGES = ["हिन्दी", "English", "अन्य (Other)"];

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("ASTROMUL · Astrologer Network");
}

function submitRegistration(data) {
  validateRegistration_(data);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getOrCreateRegistrationSheet_(ss);

    const upper = value => String(value == null ? "" : value).trim().toUpperCase();
    const safeText = value => {
      const text = upper(value);
      return /^[=+@-]/.test(text) ? "'" + text : text;
    };
    const phone = String(data.phone).trim();

    sheet.appendRow([
      new Date(),
      safeText(data.name),
      "'" + phone,
      safeText(data.address),
      data.experience,
      data.specialties.map(String).join(", "),
      "", // Consultation mode section was intentionally removed.
      data.currentConsultation,
      data.interest,
      "", // Optional profile/social link is not currently collected.
      "हाँ", // Consent checkbox was checked.
      data.certificate,
      data.languages.map(String).join(", ")
    ]);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: "आपका आवेदन सफलतापूर्वक Google Sheet में सेव हो गया है।"
    };
  } catch (err) {
    return {
      success: false,
      message: "Google Sheet में आवेदन सेव नहीं हो सका। " +
        (err && err.message ? err.message : String(err))
    };
  } finally {
    lock.releaseLock();
  }
}

function validateRegistration_(data) {
  if (!data || typeof data !== "object") {
    throw new Error("कृपया फॉर्म भरें।");
  }

  ["name", "phone", "address", "experience", "currentConsultation", "interest", "certificate"].forEach(key => {
    if (!String(data[key] || "").trim()) {
      throw new Error("कृपया सभी आवश्यक जानकारी भरें।");
    }
  });

  if (!/^[0-9+() -]{10,15}$/.test(String(data.phone).trim())) {
    throw new Error("कृपया सही मोबाइल नंबर दर्ज करें।");
  }
  if (String(data.name).trim().length > 100 || String(data.address).trim().length > 400) {
    throw new Error("नाम या पता निर्धारित सीमा से बड़ा है।");
  }

  ["experience", "currentConsultation", "interest", "certificate"].forEach(key => {
    if (!ALLOWED_YES_NO.includes(String(data[key]))) {
      throw new Error("कृपया हाँ या नहीं का विकल्प चुनें।");
    }
  });

  if (!Array.isArray(data.specialties) || data.specialties.length < 1 ||
      data.specialties.some(value => !ALLOWED_SPECIALTIES.includes(String(value)))) {
    throw new Error("कृपया सूची से कम-से-कम एक मान्य विशेषज्ञता चुनें।");
  }

  if (!Array.isArray(data.languages) || data.languages.length < 1 ||
      data.languages.some(value => !ALLOWED_LANGUAGES.includes(String(value)))) {
    throw new Error("कृपया सूची से कम-से-कम एक भाषा चुनें।");
  }

  if (!(data.consent === true || String(data.consent).toLowerCase() === "true")) {
    throw new Error("कृपया सहमति वाले चेकबॉक्स पर टिक करें।");
  }
}

function getOrCreateRegistrationSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
  const isEmpty = firstRow.every(value => !String(value).trim());

  if (isEmpty) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#07111f")
      .setFontColor("#f7d789");
    sheet.autoResizeColumns(1, HEADERS.length);
  } else {
    // Preserve existing headers and add any missing columns at the end.
    for (let col = 1; col <= HEADERS.length; col++) {
      const cell = sheet.getRange(1, col);
      if (!String(cell.getDisplayValue()).trim()) {
        cell.setValue(HEADERS[col - 1]);
        cell.setFontWeight("bold")
          .setBackground("#07111f")
          .setFontColor("#f7d789");
      }
    }
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, HEADERS.length);
  }
  return sheet;
}

function testSpreadsheetConnection() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateRegistrationSheet_(ss);
  Logger.log("Connected to spreadsheet: " + ss.getName());
  Logger.log("Registration tab: " + sheet.getName());
  return "कनेक्शन सफल है: " + ss.getName() + " / " + sheet.getName();
}
