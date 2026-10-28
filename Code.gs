const SPREADSHEET_ID = "12AU_DqW4KwlL-JkOIIHroWqpIqxptBULUOUuKUilsvY";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("ASTROMUL | Astrologer Registration")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function submitRegistration(data) {
  if (!data || !data.name || !data.phone || !data.city ||
      !data.experience || !data.currentConsultation || !data.interest) {
    throw new Error("कृपया सभी आवश्यक जानकारी भरें।");
  }

  if (!Array.isArray(data.specialties) || data.specialties.length === 0) {
    throw new Error("कम-से-कम एक विशेषज्ञता चुनें।");
  }
  if (!Array.isArray(data.modes) || data.modes.length === 0) {
    throw new Error("कम-से-कम एक परामर्श माध्यम चुनें।");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Astrologer Registrations");

    if (!sheet) {
      sheet = ss.insertSheet("Astrologer Registrations");
      sheet.appendRow([
        "Registration Date", "Full Name", "WhatsApp Number", "City / State",
        "Astrology Experience", "Specialties", "Consultation Modes",
        "Current Consultation", "ASTROMUL Interest", "Profile Link"
      ]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 10)
        .setFontWeight("bold")
        .setBackground("#101827")
        .setFontColor("#f2dba9");
    }

    // Prefix phone with apostrophe to preserve leading zeros and plus signs.
    sheet.appendRow([
      new Date(),
      String(data.name).trim(),
      "'" + String(data.phone).trim(),
      String(data.city).trim(),
      String(data.experience),
      data.specialties.join(", "),
      data.modes.join(", "),
      String(data.currentConsultation),
      String(data.interest),
      String(data.profileLink || "").trim()
    ]);

    return "आपका रजिस्ट्रेशन सफलतापूर्वक जमा हो गया है। ASTROMUL से जुड़ने के लिए धन्यवाद!";
  } finally {
    lock.releaseLock();
  }
}
