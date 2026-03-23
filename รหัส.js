/**
 * ROOM BOOKING SYSTEM - Backend
 * Optimized version with error handling and timeout management
 */

const ROOMS_CONFIG = {
  "2person": {
    label: "ห้อง 2 คน",
    sheet: "2person",
    rooms: {
      "gardenpino": { label: "Garden Pino", total: 3, emoji: "🌺" },
      "beachnest": { label: "Beach Nest", total: 2, emoji: "🏖️" },
      "gardennest": { label: "Garden Nest", total: 8, emoji: "🌿" },
      "cococube": { label: "Coco Cube", total: 10, emoji: "🥥" },
      "superiongarden": { label: "Superion Garden", total: 4, emoji: "🌳" },
      "standard": { label: "Standard", total: 8, emoji: "🛏️" }
    }
  },
  "4person": {
    label: "ห้อง 4 คน",
    sheet: "4person",
    rooms: {
      "seafront": { label: "Seafront", total: 6, emoji: "🌊" },
      "gardennestconnect": { label: "Gardennest Connect", total: 2, emoji: "🌴" }
    }
  }
};

const MAP_IMAGE_FILE_ID = "1uzRXK6CaT1msIed9uFWeM-WLNr2wZG7W";

/**
 * ================== PUBLIC FUNCTIONS ==================
 */

function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile("Index")
      .setTitle("ระบบจองห้องพัก")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    Logger.log("doGet error: " + e.message);
    return HtmlService.createHtmlOutput("<h1>Error</h1><p>" + e.message + "</p>");
  }
}

/**
 * User login function
 */
function login(studentId, name) {
  try {
    studentId = String(studentId || "").trim();
    name = String(name || "").trim();

    if (!name || !studentId) {
      return { 
        success: false, 
        message: "กรุณากรอกข้อมูลให้ครบ" 
      };
    }

    if (!/^\d{9}$/.test(studentId)) {
      return { 
        success: false, 
        message: "รหัสนักศึกษาต้องเป็นตัวเลข 9 หลัก" 
      };
    }

    const booking = getBookingByStudentId(studentId);

    return {
      success: true,
      studentId: studentId,
      name: name,
      booking: booking ? sanitizeBooking(booking) : null
    };
  } catch (e) {
    Logger.log("login error: " + e.message);
    return { 
      success: false, 
      message: "เกิดข้อผิดพลาด: " + e.message 
    };
  }
}

/**
 * Get all rooms status for a room type
 */
function getRoomsStatus(type) {
  try {
    type = String(type || "").trim();
    const config = ROOMS_CONFIG[type];
    
    if (!config) {
      return {};
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(config.sheet);

    if (!sheet) {
      sheet = createSheet(type);
      return buildEmptyStatus(config);
    }

    const lastRow = sheet.getLastRow();
    const bookedMap = {};

    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
      for (let i = 0; i < data.length; i++) {
        const roomKey = String(data[i][1] || "").trim();
        if (roomKey) {
          bookedMap[roomKey] = (bookedMap[roomKey] || 0) + 1;
        }
      }
    }

    const result = {};
    for (const key in config.rooms) {
      const room = config.rooms[key];
      const booked = Number(bookedMap[key] || 0);
      const remaining = Math.max(0, room.total - booked);

      result[key] = {
        label: String(room.label),
        emoji: String(room.emoji),
        total: Number(room.total),
        booked: booked,
        remaining: remaining,
        available: remaining > 0
      };
    }

    return result;
  } catch (e) {
    Logger.log("getRoomsStatus error: " + e.message);
    return {};
  }
}

/**
 * Submit a new booking
 */
function submitBooking(studentId, name, roomKey, type) {
  const lock = LockService.getScriptLock();
  
  try {
    if (!lock.tryLock(30000)) {
      return { 
        success: false, 
        message: "ระบบกำลังประมวลผลคำขอจากผู้ใช้อื่น กรุณาลองใหม่ในอีกสักครู่" 
      };
    }

    studentId = String(studentId || "").trim();
    name = String(name || "").trim();
    roomKey = String(roomKey || "").trim();
    type = String(type || "").trim();

    // Validation
    if (!studentId || !name || !roomKey || !type) {
      return { success: false, message: "ข้อมูลไม่ครบ" };
    }

    if (!/^\d{9}$/.test(studentId)) {
      return { success: false, message: "รหัสนักศึกษาต้องเป็นตัวเลข 9 หลัก" };
    }

    const existing = getBookingByStudentId(studentId);
    if (existing) {
      return {
        success: false,
        message: "คุณจองห้อง " + existing.roomLabel + " ไว้แล้ว"
      };
    }

    const config = ROOMS_CONFIG[type];
    if (!config) {
      return { success: false, message: "ไม่พบประเภทห้อง" };
    }

    const roomConfig = config.rooms[roomKey];
    if (!roomConfig) {
      return { success: false, message: "ไม่พบห้องนี้" };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(config.sheet);
    if (!sheet) {
      sheet = createSheet(type);
    }

    const lastRow = sheet.getLastRow();
    let bookedCnt = 0;

    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

      for (let i = 0; i < data.length; i++) {
        const rowRoomKey = String(data[i][1] || "").trim();
        const rowStudentId = String(data[i][2] || "").trim();

        if (rowRoomKey === roomKey) {
          bookedCnt++;
        }
        if (rowStudentId === studentId) {
          return { success: false, message: "คุณมีรายการจองอยู่แล้ว" };
        }
      }
    }

    if (bookedCnt >= roomConfig.total) {
      return {
        success: false,
        message: "ห้อง " + roomConfig.label + " เต็มแล้ว"
      };
    }

    const now = new Date();

    sheet.appendRow([
      now,
      roomKey,
      studentId,
      roomConfig.label,
      name,
      type,
      bookedCnt + 1
    ]);

    return {
      success: true,
      message: "จองสำเร็จ",
      booking: sanitizeBooking({
        type: type,
        roomKey: roomKey,
        roomLabel: roomConfig.label,
        name: name,
        studentId: studentId,
        emoji: roomConfig.emoji,
        timestamp: now.toISOString()
      })
    };
  } catch (e) {
    Logger.log("submitBooking error: " + e.message);
    return { 
      success: false, 
      message: "เกิดข้อผิดพลาด: " + e.message 
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cancel an existing booking
 */
function cancelBooking(studentId) {
  const lock = LockService.getScriptLock();
  
  try {
    if (!lock.tryLock(30000)) {
      return { 
        success: false, 
        message: "ระบบกำลังประมวลผลคำขอจากผู้ใช้อื่น กรุณาลองใหม่ในอีกสักครู่" 
      };
    }

    studentId = String(studentId || "").trim();
    if (!studentId) {
      return { success: false, message: "ไม่พบรหัสนักศึกษา" };
    }

    const booking = getBookingByStudentId(studentId);
    if (!booking) {
      return { success: false, message: "ไม่พบการจอง" };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(booking.sheet);
    if (!sheet) {
      return { success: false, message: "ไม่พบชีตข้อมูล" };
    }

    sheet.deleteRow(Number(booking.row));

    return { success: true, message: "ยกเลิกการจองสำเร็จ" };
  } catch (e) {
    Logger.log("cancelBooking error: " + e.message);
    return { 
      success: false, 
      message: "เกิดข้อผิดพลาด: " + e.message 
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Get map image as base64
 */
function getMapImageBase64() {
  try {
    if (!MAP_IMAGE_FILE_ID || MAP_IMAGE_FILE_ID === "YOUR_FILE_ID_HERE") {
      return "";
    }

    const file = DriveApp.getFileById(MAP_IMAGE_FILE_ID);
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    
    return "data:" + blob.getContentType() + ";base64," + base64;
  } catch (e) {
    Logger.log("getMapImageBase64 error: " + e.message);
    return "";
  }
}

/**
 * ================== PRIVATE HELPER FUNCTIONS ==================
 */

function getBookingByStudentId(studentId) {
  try {
    studentId = String(studentId || "").trim();
    if (!studentId) return null;

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    for (const type in ROOMS_CONFIG) {
      const config = ROOMS_CONFIG[type];
      const sheet = ss.getSheetByName(config.sheet);
      if (!sheet) continue;

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) continue;

      const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const rowStudentId = String(row[2] || "").trim();

        if (rowStudentId === studentId) {
          const roomKey = String(row[1] || "").trim();
          const roomInfo = ROOMS_CONFIG[type].rooms[roomKey] || {};
          
          return {
            type: String(type),
            roomKey: roomKey,
            roomLabel: String(row[3] || roomInfo.label || ""),
            name: String(row[4] || ""),
            studentId: String(row[2] || ""),
            timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0] || ""),
            row: Number(i + 2),
            sheet: String(config.sheet),
            emoji: String(roomInfo.emoji || "🛏️")
          };
        }
      }
    }

    return null;
  } catch (e) {
    Logger.log("getBookingByStudentId error: " + e.message);
    return null;
  }
}

function createSheet(type) {
  type = String(type || "").trim();
  const config = ROOMS_CONFIG[type];
  
  if (!config) {
    throw new Error("Invalid room type");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(config.sheet);
  
  if (sheet) {
    return sheet;
  }

  sheet = ss.insertSheet(config.sheet);
  sheet.appendRow([
    "Timestamp",
    "Room Key",
    "Student ID",
    "Room Label",
    "Name",
    "Type",
    "No."
  ]);

  const header = sheet.getRange(1, 1, 1, 7);
  header.setBackground("#1a6b5a");
  header.setFontColor("white");
  header.setFontWeight("bold");

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 80);

  return sheet;
}

function buildEmptyStatus(config) {
  const result = {};
  
  for (const key in config.rooms) {
    const r = config.rooms[key];
    result[key] = {
      label: String(r.label),
      emoji: String(r.emoji),
      total: Number(r.total),
      booked: 0,
      remaining: Number(r.total),
      available: true
    };
  }
  
  return result;
}

function sanitizeBooking(booking) {
  if (!booking) return null;
  
  return {
    type: String(booking.type || ""),
    roomKey: String(booking.roomKey || ""),
    roomLabel: String(booking.roomLabel || ""),
    name: String(booking.name || ""),
    studentId: String(booking.studentId || ""),
    timestamp: String(booking.timestamp || ""),
    emoji: String(booking.emoji || "🛏️"),
    row: booking.row ? Number(booking.row) : null,
    sheet: booking.sheet ? String(booking.sheet) : ""
  };
}