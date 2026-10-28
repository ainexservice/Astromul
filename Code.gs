/**
 * ASTROMUL — GitHub Pages to Google Sheets bridge.
 * SETUP: Run setup() once in the Apps Script editor and approve the requested Google permissions.
 * Then deploy as Web App (Execute as: Me; Who has access: Anyone where available).
 */
const ASTROMUL_SHEET_NAME = "ASTROMUL Registrations";
const ASTROMUL_PROPERTY_KEY = "ASTROMUL_SPREADSHEET_ID";
const ASTROMUL_HEADERS = [
  "SUBMISSION DATE", "REQUEST ID", "FULL NAME", "MOBILE / WHATSAPP", "FULL RESIDENTIAL ADDRESS",
  "DOES ASTROLOGY WORK", "CURRENTLY CONSULTS PEOPLE", "WANTS TO JOIN ASTROMUL",
  "ASTROLOGY CERTIFICATE / DIPLOMA", "SPECIALTIES", "LANGUAGES KNOWN", "CONSENT"
];
const ASTROMUL_YES_NO = ["हाँ", "नहीं"];
const ASTROMUL_SPECIALTIES = [
  "वैदिक ज्योतिष एवं कुंडली", "विवाह मिलान", "अंक ज्योतिष", "हस्तरेखा",
  "प्रश्न ज्योतिष", "रत्न एवं रुद्राक्ष", "वास्तु", "लाल किताब", "टैरो एवं अन्य"
];
const ASTROMUL_LANGUAGES = ["हिन्दी", "English", "अन्य (Other)"];

/** Run manually once. Creates a private Google spreadsheet and the registration tab. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(ASTROMUL_PROPERTY_KEY);
  let ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); }
    catch (err) { id = ""; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create("ASTROMUL Registrations");
    id = ss.getId();
    props.setProperty(ASTROMUL_PROPERTY_KEY, id);
  }
  const sheet = getOrCreateSheet_(ss);
  Logger.log("ASTROMUL SETUP COMPLETE");
  Logger.log("Spreadsheet URL: " + ss.getUrl());
  Logger.log("Registration tab: " + sheet.getName());
  return "Setup successful. Open the execution log to view your Google Sheet URL: " + ss.getUrl();
}

/** Public health-check page. Does not expose the spreadsheet URL. */
function doGet() {
  const ready = Boolean(PropertiesService.getScriptProperties().getProperty(ASTROMUL_PROPERTY_KEY));
  const status = ready ? "Backend Ready" : "Setup Required";
  const instruction = ready
    ? "ASTROMUL Google Sheets backend is online. Submit a registration from the GitHub Pages website to test it."
    : "Open this Apps Script project, run setup() once, approve permissions, then redeploy this web app.";
  const html = '<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>ASTROMUL Backend</title>' +
    '<style>body{font:16px Arial,sans-serif;max-width:680px;margin:60px auto;padding:0 20px;color:#102039}h1{color:#9a6a1b}code{background:#f4eee2;padding:3px 6px;border-radius:4px}</style></head><body><h1>ASTROMUL Google Sheets Backend</h1><p><b>' + status + '</b></p><p>' + instruction + '</p></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Receives an HTML form POST from GitHub Pages via a hidden iframe; no CORS fetch is needed. */
function doPost(e) {
  let requestId = "";
  try {
    const raw = e && e.parameter && e.parameter.payload;
    if (!raw) throw new Error("फॉर्म का डेटा प्राप्त नहीं हुआ। नवीनतम index.html इस्तेमाल करें।");
    const data = JSON.parse(raw);
    requestId = String(data.requestId || "").slice(0, 100);
    validateRegistration_(data);

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const ss = openRegistrationSpreadsheet_();
      const sheet = getOrCreateSheet_(ss);
      sheet.appendRow([
        new Date(),
        safeCell_(requestId),
        safeCell_(data.name),
        safeCell_(data.phone),
        safeCell_(data.address),
        data.experience,
        data.currentConsultation,
        data.interest,
        data.certificate,
        data.specialties.join(", "),
        data.languages.join(", "),
        "हाँ"
      ]);
      const row = sheet.getLastRow();
      sheet.getRange(row, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    return responsePage_(requestId, true, "आपका आवेदन सफलतापूर्वक Google Sheet में सेव हो गया है।");
  } catch (err) {
    return responsePage_(requestId, false, "आवेदन सेव नहीं हो सका: " + (err && err.message ? err.message : String(err)));
  }
}

function openRegistrationSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(ASTROMUL_PROPERTY_KEY);
  if (!id) throw new Error("Apps Script setup() एक बार चलाएँ और फिर Web App को redeploy करें।");
  return SpreadsheetApp.openById(id);
}

function getOrCreateSheet_(ss) {
  let sheet = ss.getSheetByName(ASTROMUL_SHEET_NAME);
  if (!sheet) {
    const sheets = ss.getSheets();
    if (sheets.length === 1 && sheets[0].getName() === "Sheet1" && sheets[0].getLastRow() === 0) {
      sheet = sheets[0];
      sheet.setName(ASTROMUL_SHEET_NAME);
    } else {
      sheet = ss.insertSheet(ASTROMUL_SHEET_NAME);
    }
  }
  const current = sheet.getRange(1, 1, 1, ASTROMUL_HEADERS.length).getDisplayValues()[0];
  const empty = current.every(v => !String(v).trim());
  if (empty) {
    sheet.getRange(1, 1, 1, ASTROMUL_HEADERS.length).setValues([ASTROMUL_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, ASTROMUL_HEADERS.length)
      .setFontWeight("bold").setBackground("#07111f").setFontColor("#f7d789").setWrap(true);
  } else {
    for (let col = 1; col <= ASTROMUL_HEADERS.length; col++) {
      const cell = sheet.getRange(1, col);
      if (!String(cell.getDisplayValue()).trim()) {
        cell.setValue(ASTROMUL_HEADERS[col - 1]);
        cell.setFontWeight("bold").setBackground("#07111f").setFontColor("#f7d789");
      }
    }
    sheet.setFrozenRows(1);
  }
  sheet.autoResizeColumns(1, ASTROMUL_HEADERS.length);
  return sheet;
}

function validateRegistration_(data) {
  if (!data || typeof data !== "object") throw new Error("फॉर्म का डेटा सही नहीं है।");
  ["name", "phone", "address", "experience", "currentConsultation", "interest", "certificate"].forEach(key => {
    if (!String(data[key] || "").trim()) throw new Error("कृपया सभी आवश्यक जानकारी भरें।");
  });
  if (String(data.name).trim().length > 100) throw new Error("नाम 100 अक्षरों से छोटा रखें।");
  if (String(data.address).trim().length > 400) throw new Error("पता 400 अक्षरों से छोटा रखें।");
  if (!/^[0-9+() -]{10,15}$/.test(String(data.phone).trim())) throw new Error("कृपया सही मोबाइल नंबर दर्ज करें।");
  ["experience", "currentConsultation", "interest", "certificate"].forEach(key => {
    if (ASTROMUL_YES_NO.indexOf(String(data[key])) === -1) throw new Error("हाँ/नहीं वाले प्रश्नों में एक विकल्प चुनें।");
  });
  if (!Array.isArray(data.specialties) || data.specialties.length < 1 || data.specialties.length > ASTROMUL_SPECIALTIES.length ||
      data.specialties.some(value => ASTROMUL_SPECIALTIES.indexOf(String(value)) === -1)) {
    throw new Error("कम-से-कम एक मान्य विशेषज्ञता चुनें।");
  }
  if (!Array.isArray(data.languages) || data.languages.length < 1 || data.languages.length > ASTROMUL_LANGUAGES.length ||
      data.languages.some(value => ASTROMUL_LANGUAGES.indexOf(String(value)) === -1)) {
    throw new Error("कम-से-कम एक मान्य भाषा चुनें।");
  }
  if (!(data.consent === true || String(data.consent).toLowerCase() === "true")) throw new Error("सहमति आवश्यक है।");
}

// Prevent spreadsheet formula injection in visitor-submitted text.
function safeCell_(value) {
  const text = String(value == null ? "" : value).trim();
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

// The iframe response posts a result message back to the GitHub Pages window.
function responsePage_(requestId, success, message) {
  const payload = JSON.stringify({
    app: "ASTROMUL_SHEET_RESPONSE",
    requestId: String(requestId || ""),
    success: Boolean(success),
    message: String(message || "")
  }).replace(/</g, "\\u003c");
  const html = '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head><body>' +
    '<script>(function(){var result=' + payload + ';try{window.top.postMessage(result,"*");}catch(e){}document.body.textContent=result.message;})();</script>' +
    '<noscript>ASTROMUL request received.</noscript></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Optional diagnostic: run after setup() to verify that the Sheet is readable and writable. */
function testSpreadsheetConnection() {
  const ss = openRegistrationSpreadsheet_();
  const sheet = getOrCreateSheet_(ss);
  Logger.log("Spreadsheet URL: " + ss.getUrl());
  Logger.log("Registration tab: " + sheet.getName());
  return "कनेक्शन सफल: " + ss.getName() + " / " + sheet.getName();
}
