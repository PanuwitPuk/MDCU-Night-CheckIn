const CONFIG = {
  MAIN_SHEET_NAME: "DATA_82",
  PLACE_SHEET_NAME: "dropdown_place",
  TIME_SHEET_NAME: "dropdown_time",
  EVENT_DATE_SHEET_NAME: "dropdown_place",
  EVENT_DATE_CELL: "C2",
  TIMEZONE: "Asia/Bangkok"
};

function getHeaderMap(headers) {
  const map = {};
  headers.forEach((h, i) => {
    map[String(h).trim().toLowerCase()] = i;
  });
  return map;
}

function extractColumnA_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];

  const rawValues = sheet.getRange(1, 1, lastRow, 1).getValues().flat().filter(val => val !== "" && val != null);
  const placeholder = "กรุณาเลือก";

  // 🟢 เพิ่มการ sort แบบเดียวกับบ้านและห้องนอน
  return rawValues.sort((a, b) => {
    const textA = String(a).trim();
    const textB = String(b).trim();
    if (textA === placeholder && textB !== placeholder) return -1;
    if (textB === placeholder && textA !== placeholder) return 1;
    return textA.localeCompare(textB, undefined, { numeric: true, sensitivity: "base" });
  });
}

function parseThaiDateValue_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return {
      year: value.getFullYear(),
      month: value.getMonth() + 1,
      day: value.getDate(),
      displayValue: Utilities.formatDate(value, CONFIG.TIMEZONE, "d/M/yyyy")
    };
  }

  const text = String(value || "").trim();
  const match = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (!match) {
    const parsedDate = new Date(text);
    if (!isNaN(parsedDate.getTime())) {
      return {
        year: parsedDate.getFullYear(),
        month: parsedDate.getMonth() + 1,
        day: parsedDate.getDate(),
        displayValue: text
      };
    }

    throw new Error(`อ่านวันที่ใน ${CONFIG.EVENT_DATE_SHEET_NAME}!${CONFIG.EVENT_DATE_CELL} ไม่ได้ กรุณาใส่เป็นวันที่ เช่น 12/5/2026`);
  }

  return {
    day: Number(match[1]),
    month: Number(match[2]),
    year: Number(match[3]),
    displayValue: text
  };
}

function getEventDateConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.EVENT_DATE_SHEET_NAME);
  if (!sheet) throw new Error("ไม่พบชีต: " + CONFIG.EVENT_DATE_SHEET_NAME);

  // 🟢 ดึงวันที่เปลี่ยนกิจกรรมจาก dropdown_place!C2
  // ส่วนเวลา 10:00 น. และการตัดสินว่าเป็นกิจกรรมไหน จะไปคำนวณต่อในหน้าเว็บ
  const rawDate = sheet.getRange(CONFIG.EVENT_DATE_CELL).getValue();
  const displayDate = sheet.getRange(CONFIG.EVENT_DATE_CELL).getDisplayValue();
  const parsedDate = parseThaiDateValue_(rawDate || displayDate);

  return {
    dateCell: `${CONFIG.EVENT_DATE_SHEET_NAME}!${CONFIG.EVENT_DATE_CELL}`,
    dateValue: parsedDate.displayValue,
    day: parsedDate.day,
    month: parsedDate.month,
    year: parsedDate.year,
    timezone: CONFIG.TIMEZONE
  };
}

function doGet(e) {
  return HtmlService.createHtmlOutput(
    "<script>window.top.location.href='https://freshyxnight2026.vercel.app/';</script>"
  );
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const data = payload.data || {};
    let result = {};

    if (action === "getConfigs") {
      result = getConfigs();
    } else if (action === "verifyUser") {
      result = verifyUser(data.id);
    } else if (action === "processCheckIn") {
      result = processCheckIn(data.id, data.round, data.place);
    } else if (action === "getDashboardData") {
      result = getDashboardData(data.keyword, data.baan, data.dorm, data.round);
    } else {
      result = { status: "error", message: "ไม่รู้จักคำสั่ง" };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 🟢 1. โหลดการตั้งค่า: ส่งข้อมูลบ้าน ห้องนอน และวันที่เปลี่ยนกิจกรรมกลับไป
function getConfigs() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
    if (!mainSheet) throw new Error("ไม่พบชีต: " + CONFIG.MAIN_SHEET_NAME);

    const data = mainSheet.getDataRange().getValues();
    const headers = data[0];
    const hMap = getHeaderMap(headers);

    let usersDict = {};
    let allBaans = new Set();
    let allDorms = new Set(); // 🟢 เพิ่ม Set สำหรับเก็บห้องนอน

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = String(row[hMap["id"]] || "").trim();
      if (!id) continue;

      const baan = String(row[hMap["baan"]] || "").trim();
      if (baan && baan !== "-") allBaans.add(baan);

      const dorm = String(row[hMap["dorm"]] || "").trim();
      if (dorm && dorm !== "-") allDorms.add(dorm); // 🟢 เก็บข้อมูลห้องนอนที่เอาเฉพาะที่มีค่า

      usersDict[id] = {
        id: id,
        baan: baan,
        dorm: dorm,
        codeline: String(row[hMap["codeline"]] || "").trim(),
        title: String(row[hMap["title"]] || "").trim(),
        firstName: String(row[hMap["first_name"]] || "").trim(),
        lastName: String(row[hMap["last_name"]] || "").trim(),
        nickname: String(row[hMap["nickname"]] || "").trim(),
        phone: String(row[hMap["phone"]] || "").trim()
      };
    }

    // เรียงชื่อบ้าน/เลขบ้าน ให้ถูกต้องตามหลักเลขคณิตและพจนานุกรม
    const sortedBaans = Array.from(allBaans).sort((a, b) => {
      return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
    });

    // เรียงเลขห้องนอน (เช่น 1, 2, 10) ให้ถูกต้องตามลำดับเลขคณิต (ถ้ามีเฉพาะตัวเลข) หรือเรียงอักษรปกติ
    const sortedDorms = Array.from(allDorms).sort((a, b) => {
      return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
    });

    return {
      status: "success",
      places: extractColumnA_(CONFIG.PLACE_SHEET_NAME),
      times: extractColumnA_(CONFIG.TIME_SHEET_NAME),
      eventDate: getEventDateConfigSafe_(),
      baans: sortedBaans, // ส่งบ้านที่ sort แล้ว
      dorms: sortedDorms, // 🟢 ส่งห้องนอนที่ sort แล้ว
      users: usersDict
    };
  } catch (err) {
    return { status: "error", message: err.message };
  }
}

function getEventDateConfigSafe_() {
  try {
    return getEventDateConfig_();
  } catch (err) {
    return {
      dateCell: `${CONFIG.EVENT_DATE_SHEET_NAME}!${CONFIG.EVENT_DATE_CELL}`,
      dateValue: "",
      day: null,
      month: null,
      year: null,
      timezone: CONFIG.TIMEZONE,
      error: err.message
    };
  }
}

function findUserInSystem_(targetId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);

  if (!sheet) return { found: false };

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { found: false };

  const headers = data[0];
  const hMap = getHeaderMap(headers);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = String(row[hMap["id"]] || "").trim();

    if (id === targetId) {
      return {
        found: true, sheet: sheet, headers: headers, hMap: hMap, rowIndex: i + 1,
        userData: {
          nickname: String(row[hMap["nickname"]] || "-").trim(),
          firstName: String(row[hMap["first_name"]] || "").trim(),
          lastName: String(row[hMap["last_name"]] || "").trim(),
          baan: String(row[hMap["baan"]] || "-").trim(),
          dorm: String(row[hMap["dorm"]] || "-").trim(),
          codeline: String(row[hMap["codeline"]] || "-").trim(),
          phone: String(row[hMap["phone"]] || "-").trim()
        }
      };
    }
  }
  return { found: false };
}

function verifyUser(id) {
  try {
    const result = findUserInSystem_(String(id).trim());
    if (result.found) return { status: "success", data: result.userData };
    return { status: "not_found", message: `ไม่พบข้อมูลรหัส: ${id}` };
  } catch (error) {
    return { status: "error", message: error.message };
  }
}

function processCheckIn(id, round, place) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(1800000);

    const targetId = String(id).trim();
    const targetRound = String(round).trim().toLowerCase();

    const result = findUserInSystem_(targetId);
    if (!result.found) return { status: "not_found", message: `ไม่พบข้อมูลรหัส: ${targetId}` };

    const { sheet, hMap, rowIndex, userData } = result;

    const timeColIndex = hMap[targetRound];

    if (timeColIndex === undefined) {
      return { status: "error", message: `ไม่พบคอลัมน์ชื่อ '${round}' ในตาราง` };
    }

    const timeCol = timeColIndex + 1;
    const placeCol = timeCol + 1;

    const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "HH:mm");

    sheet.getRange(rowIndex, timeCol).setValue(timestamp);
    sheet.getRange(rowIndex, placeCol).setValue(String(place).trim());

    return { status: "success", data: userData };
  } catch (error) {
    return { status: "error", message: error.message };
  } finally {
    lock.releaseLock();
  }
}

// 🟢 2. ดึงข้อมูลสรุปยอด: รองรับการเลือก "ทุกบ้าน" หรือ "ทุกห้องนอน"
function getDashboardData(keyword, targetBaan, targetDorm, targetRound) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);

    if (!sheet) throw new Error("ไม่พบชีตหลัก");

    let total = 0, checkedIn = 0, missing = 0;
    let checkedInList = [], missingList = [];

    const searchRound = String(targetRound || "").trim().toLowerCase();
    // หากเลือก "ทุกบ้าน" ค่าที่ส่งมาจะเป็น "" (String ว่าง) เราจะไม่นำไปกรอง
    const searchBaan = String(targetBaan || "").trim();
    const searchDorm = String(targetDorm || "").trim();
    const searchKeyword = String(keyword || "").trim().toLowerCase();

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) throw new Error("ไม่มีข้อมูลในตาราง");

    const headers = data[0];
    const hMap = getHeaderMap(headers);

    const timeColIndex = hMap[searchRound];
    if (timeColIndex === undefined) throw new Error("ไม่พบคอลัมน์เวลานี้ในตาราง");

    const timeCol = timeColIndex;
    const placeCol = timeCol + 1;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = String(row[hMap["id"]] || "").trim();
      if (!id) continue;

      const rowBaan = String(row[hMap["baan"]] || "").trim();
      const rowDorm = String(row[hMap["dorm"]] || "").trim();
      const firstName = String(row[hMap["first_name"]] || "").trim();
      const lastName = String(row[hMap["last_name"]] || "").trim();
      const nickname = String(row[hMap["nickname"]] || "").trim();

      const name = `${firstName} ${lastName} (${nickname || "-"})`.trim();
      const fullNameLower = `${firstName} ${lastName} ${nickname}`.toLowerCase();

      // 🟢 ตรวจสอบเงื่อนไขว่าตรงกับสิ่งที่ค้นหาหรือไม่
      let match = true;

      // ถ้า searchBaan ไม่ใช่ค่าว่าง ("") และ บ้านในแถวนี้ไม่ตรงกับที่หา -> ไม่เอา
      if (searchBaan !== "" && rowBaan !== searchBaan) match = false;

      // ถ้า searchDorm ไม่ใช่ค่าว่าง ("") และ ห้องนอนในแถวนี้ไม่ตรงกับที่หา -> ไม่เอา
      if (searchDorm !== "" && rowDorm !== searchDorm) match = false;

      // ถ้า searchKeyword ไม่ใช่ค่าว่าง ("") และ ไม่มีคำค้นหาในชื่อ หรือ รหัสไม่ตรง -> ไม่เอา
      if (searchKeyword !== "" && !fullNameLower.includes(searchKeyword) && id !== searchKeyword) match = false;

      // ถ้าตรงกับเงื่อนไขทั้งหมด (หรือไม่ได้กรองอะไรเลย)
      if (match) {
        total++;

        let time = row[timeCol];
        const place = row[placeCol];

        if (time) {
          if (time instanceof Date) {
            // กรณีเป็น Date Object โดยตรง
            time = Utilities.formatDate(time, CONFIG.TIMEZONE, "HH:mm");
          } else {
            const timeStr = String(time).trim();
            const parsedDate = new Date(timeStr);

            // เช็คว่าสามารถแปลงเป็น Date ได้ และเป็นข้อความยาว (เช่น Sat Dec 30...)
            if (!isNaN(parsedDate.getTime()) && timeStr.length > 10) {
              time = Utilities.formatDate(parsedDate, CONFIG.TIMEZONE, "HH:mm");
            } else {
              time = timeStr.split(" ")[0];
            }
          }
        }

        if (time && time !== "") {
          checkedIn++;
          checkedInList.push({ id, name, time, place });
        } else {
          missing++;
          missingList.push({ id, name });
        }
      }
    }

    return {
      status: "success",
      summary: { total, checkedIn, missing },
      checkedInList,
      missingList
    };

  } catch (error) {
    return { status: "error", message: error.message };
  }
}
