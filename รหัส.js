/**
 * Mudchim Room Booking
 * Centralized room configuration with lightweight caching and locking
 * for safer concurrent usage.
 */

const ROOM_GROUPS = Object.freeze({
  "6person": {
    capacity: 6,
    label: "ห้อง 6 คน",
    sheet: "6person",
    emoji: "👨‍👩‍👧‍👦",
    accent: "#0f766e",
    rooms: Object.freeze({
      gardennestconnect: {
        label: "Gardennest Connect",
        total: 2,
        emoji: "🌿"
      }
    })
  },
  "5person": {
    capacity: 5,
    label: "ห้อง 5 คน",
    sheet: "5person",
    emoji: "🧑‍🤝‍🧑",
    accent: "#0f766e",
    rooms: Object.freeze({
      seafront: {
        label: "Seafront",
        total: 1,
        emoji: "🌊"
      }
    })
  },
  "4person": {
    capacity: 4,
    label: "ห้อง 4 คน",
    sheet: "4person",
    emoji: "👨‍👩‍👧",
    accent: "#1d4ed8",
    rooms: Object.freeze({
      seafront: {
        label: "Seafront",
        total: 5,
        emoji: "🌊"
      }
    })
  },
  "3person": {
    capacity: 3,
    label: "ห้อง 3 คน",
    sheet: "3person",
    emoji: "👨‍👩‍👦",
    accent: "#2563eb",
    rooms: Object.freeze({
      gardennest: {
        label: "Garden Nest",
        total: 8,
        emoji: "🌿"
      },
      cococube: {
        label: "Coco Cube",
        total: 4,
        emoji: "🥥"
      }
    })
  },
  "2person": {
    capacity: 2,
    label: "ห้อง 2 คน",
    sheet: "2person",
    emoji: "👫",
    accent: "#7c3aed",
    rooms: Object.freeze({
      gardenpino: {
        label: "Garden Pino",
        total: 3,
        emoji: "🌱"
      },
      beachnest: {
        label: "Beach Nest",
        total: 2,
        emoji: "🏖️"
      },
      superiongarden: {
        label: "Superion Garden",
        total: 4,
        emoji: "🌳"
      },
      superior: {
        label: "Superior",
        total: 8,
        emoji: "🛏️"
      },
      standard: {
        label: "Standard",
        total: 8,
        emoji: "🪟"
      }
    })
  }
});

const GROUP_ORDER = Object.freeze(
  Object.keys(ROOM_GROUPS).sort(function(a, b) {
    return ROOM_GROUPS[b].capacity - ROOM_GROUPS[a].capacity;
  })
);

const MAP_IMAGE_FILE_ID = "1uzRXK6CaT1msIed9uFWeM-WLNr2wZG7W";
const ROOM_STATUS_CACHE_TTL = 8;
const BOOKING_CACHE_TTL = 120;
const USER_REGISTRY_SHEET = "Users";
const LOG_BOOK_SHEET = "LogBook";

function doGet() {
  try {
    const template = HtmlService.createTemplateFromFile("Index");
    template.roomCatalogJson = JSON.stringify(getRoomCatalog());

    return template
      .evaluate()
      .setTitle("ระบบจองห้องพัก");
  } catch (error) {
    Logger.log("doGet error: " + error.message);
    return HtmlService.createHtmlOutput(
      "<h1>Error</h1><p>" + sanitizeHtml_(error.message) + "</p>"
    );
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function login(studentId, name, email) {
  try {
    email = normalizeEmail_(email);
    studentId = normalizeStudentId_(studentId);
    name = normalizeName_(name);

    if (!name || !studentId || !email) {
      return {
        success: false,
        message: "กรุณากรอกชื่อ รหัสนักศึกษา และอีเมลให้ครบ"
      };
    }

    if (!/^\d{9}$/.test(studentId)) {
      return {
        success: false,
        message: "รหัสนักศึกษาต้องเป็นตัวเลข 9 หลัก"
      };
    }

    if (!/@su\.ac\.th$/i.test(email)) {
      return {
        success: false,
        message: "กรุณาใช้อีเมล @su.ac.th เท่านั้น"
      };
    }

    const identity = resolveUserIdentity_(email, studentId, name);
    let booking = getBookingByStudentId_(studentId, { useCache: true });
    booking = backfillBookingEmailIfMissing_(booking, email);

    return {
      success: true,
      studentId: identity.studentId,
      name: identity.name,
      email: identity.email,
      booking: booking ? sanitizeBooking_(booking) : null
    };
  } catch (error) {
    Logger.log("login error: " + error.message);
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message
    };
  }
}

function getRoomCatalog() {
  return GROUP_ORDER.map(function(typeKey) {
    const config = ROOM_GROUPS[typeKey];
    const roomKeys = Object.keys(config.rooms);
    const totalInventory = roomKeys.reduce(function(sum, roomKey) {
      return sum + Number(config.rooms[roomKey].total || 0);
    }, 0);

    return {
      key: String(typeKey),
      capacity: Number(config.capacity),
      label: String(config.label),
      sheet: String(config.sheet),
      emoji: String(config.emoji),
      accent: String(config.accent),
      roomTypeCount: roomKeys.length,
      totalInventory: totalInventory,
      rooms: roomKeys.map(function(roomKey) {
        const room = config.rooms[roomKey];
        return {
          key: String(roomKey),
          label: String(room.label),
          total: Number(room.total),
          emoji: String(room.emoji)
        };
      })
    };
  });
}

function getRoomsStatus(typeKey) {
  try {
    typeKey = String(typeKey || "").trim();
    const config = ROOM_GROUPS[typeKey];

    if (!config) {
      return null;
    }

    const cache = CacheService.getScriptCache();
    const cacheKey = buildRoomStatusCacheKey_(typeKey);
    const cached = cache.get(cacheKey);

    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (cacheError) {
        Logger.log("getRoomsStatus cache parse error: " + cacheError.message);
      }
    }

    const sheet = ensureSheet_(typeKey);
    const bookedMap = {};
    const lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (let index = 0; index < values.length; index++) {
        const roomKey = String(values[index][0] || "").trim();
        if (config.rooms[roomKey]) {
          bookedMap[roomKey] = (bookedMap[roomKey] || 0) + 1;
        }
      }
    }

    const rooms = Object.keys(config.rooms).map(function(roomKey) {
      const room = config.rooms[roomKey];
      const booked = Number(bookedMap[roomKey] || 0);
      const total = Number(room.total || 0);
      const remaining = Math.max(0, total - booked);

      return {
        key: String(roomKey),
        label: String(room.label),
        emoji: String(room.emoji),
        total: total,
        booked: booked,
        remaining: remaining,
        available: remaining > 0
      };
    });

    const status = {
      key: String(typeKey),
      label: String(config.label),
      capacity: Number(config.capacity),
      emoji: String(config.emoji),
      accent: String(config.accent),
      roomTypeCount: rooms.length,
      totalInventory: rooms.reduce(function(sum, room) {
        return sum + room.total;
      }, 0),
      totalBooked: rooms.reduce(function(sum, room) {
        return sum + room.booked;
      }, 0),
      totalRemaining: rooms.reduce(function(sum, room) {
        return sum + room.remaining;
      }, 0),
      rooms: rooms
    };

    cache.put(cacheKey, JSON.stringify(status), ROOM_STATUS_CACHE_TTL);
    return status;
  } catch (error) {
    Logger.log("getRoomsStatus error: " + error.message);
    return null;
  }
}

function submitBooking(studentId, name, email, roomKey, typeKey) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;

  studentId = normalizeStudentId_(studentId);
  name = normalizeName_(name);
  email = normalizeEmail_(email);
  roomKey = String(roomKey || "").trim();
  typeKey = String(typeKey || "").trim();

  try {
    if (!lock.tryLock(25000)) {
      appendLogBook_({
        action: "BOOK",
        status: "REJECTED",
        studentId: studentId,
        name: name,
        email: email,
        type: typeKey,
        roomKey: roomKey,
        message: "ระบบกำลังประมวลผลคำขอจำนวนมาก กรุณาลองใหม่อีกครั้ง"
      });
      return {
        success: false,
        message: "ระบบกำลังประมวลผลคำขอจำนวนมาก กรุณาลองใหม่อีกครั้ง"
      };
    }
    lockAcquired = true;

    if (!studentId || !name || !email || !roomKey || !typeKey) {
      appendLogBook_({
        action: "BOOK",
        status: "REJECTED",
        studentId: studentId,
        name: name,
        email: email,
        type: typeKey,
        roomKey: roomKey,
        message: "ข้อมูลไม่ครบ"
      });
      return { success: false, message: "ข้อมูลไม่ครบ" };
    }

    if (!/^\d{9}$/.test(studentId)) {
      appendLogBook_({
        action: "BOOK",
        status: "REJECTED",
        studentId: studentId,
        name: name,
        email: email,
        type: typeKey,
        roomKey: roomKey,
        message: "รหัสนักศึกษาต้องเป็นตัวเลข 9 หลัก"
      });
      return {
        success: false,
        message: "รหัสนักศึกษาต้องเป็นตัวเลข 9 หลัก"
      };
    }

    if (!/@su\.ac\.th$/i.test(email)) {
      appendLogBook_({
        action: "BOOK",
        status: "REJECTED",
        studentId: studentId,
        name: name,
        email: email,
        type: typeKey,
        roomKey: roomKey,
        message: "กรุณาใช้อีเมล @su.ac.th เท่านั้น"
      });
      return { success: false, message: "กรุณาใช้อีเมล @su.ac.th เท่านั้น" };
    }

    const identity = resolveUserIdentity_(email, studentId, name);
    studentId = identity.studentId;
    name = identity.name;
    email = identity.email;

    const existing = getBookingByStudentId_(studentId, { useCache: false });
    if (existing) {
      appendLogBook_({
        action: "BOOK",
        status: "REJECTED",
        studentId: studentId,
        name: name,
        email: email,
        type: existing.type,
        roomKey: existing.roomKey,
        roomLabel: existing.roomLabel,
        message: "คุณจองห้อง " + existing.roomLabel + " ไว้แล้ว"
      });
      return {
        success: false,
        message: "คุณจองห้อง " + existing.roomLabel + " ไว้แล้ว"
      };
    }

    const config = ROOM_GROUPS[typeKey];
    if (!config) {
      return { success: false, message: "ไม่พบประเภทห้อง" };
    }

    const roomConfig = config.rooms[roomKey];
    if (!roomConfig) {
      appendLogBook_({
        action: "BOOK",
        status: "REJECTED",
        studentId: studentId,
        name: name,
        email: email,
        type: typeKey,
        roomKey: roomKey,
        message: "ไม่พบห้องนี้"
      });
      return { success: false, message: "ไม่พบห้องนี้" };
    }

    const sheet = ensureSheet_(typeKey);
    const bookedCount = countBookingsForRoom_(sheet, roomKey);

    if (bookedCount >= Number(roomConfig.total || 0)) {
      appendLogBook_({
        action: "BOOK",
        status: "REJECTED",
        studentId: studentId,
        name: name,
        email: email,
        type: typeKey,
        roomKey: roomKey,
        roomLabel: roomConfig.label,
        message: "ห้อง " + roomConfig.label + " เต็มแล้ว"
      });
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
      email,
      typeKey,
      bookedCount + 1
    ]);
    SpreadsheetApp.flush();

    const booking = sanitizeBooking_({
      type: typeKey,
      typeLabel: config.label,
      roomKey: roomKey,
      roomLabel: roomConfig.label,
      name: name,
      studentId: studentId,
      email: email,
      emoji: roomConfig.emoji,
      timestamp: now.toISOString(),
      sheet: config.sheet
    });

    invalidateRoomStatusCache_(typeKey);
    setBookingCache_(studentId, booking);
    appendLogBook_({
      action: "BOOK",
      status: "SUCCESS",
      studentId: studentId,
      name: name,
      email: email,
      type: typeKey,
      roomKey: roomKey,
      roomLabel: roomConfig.label,
      message: "จองสำเร็จ"
    });

    return {
      success: true,
      message: "จองสำเร็จ",
      booking: booking
    };
  } catch (error) {
    Logger.log("submitBooking error: " + error.message);
    appendLogBook_({
      action: "BOOK",
      status: "ERROR",
      studentId: studentId,
      name: name,
      email: email,
      type: typeKey,
      roomKey: roomKey,
      message: error.message
    });
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message
    };
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

function cancelBooking(studentId) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;

  studentId = normalizeStudentId_(studentId);

  try {
    if (!lock.tryLock(25000)) {
      appendLogBook_({
        action: "CANCEL",
        status: "REJECTED",
        studentId: studentId,
        message: "ระบบกำลังประมวลผลคำขอจำนวนมาก กรุณาลองใหม่อีกครั้ง"
      });
      return {
        success: false,
        message: "ระบบกำลังประมวลผลคำขอจำนวนมาก กรุณาลองใหม่อีกครั้ง"
      };
    }
    lockAcquired = true;
    if (!studentId) {
      appendLogBook_({
        action: "CANCEL",
        status: "REJECTED",
        message: "ไม่พบรหัสนักศึกษา"
      });
      return { success: false, message: "ไม่พบรหัสนักศึกษา" };
    }

    const booking = getBookingByStudentId_(studentId, { useCache: false });
    if (!booking) {
      appendLogBook_({
        action: "CANCEL",
        status: "REJECTED",
        studentId: studentId,
        message: "ไม่พบการจอง"
      });
      return { success: false, message: "ไม่พบการจอง" };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(booking.sheet);
    if (!sheet) {
      return { success: false, message: "ไม่พบชีตข้อมูล" };
    }

    sheet.deleteRow(Number(booking.row));
    SpreadsheetApp.flush();

    invalidateRoomStatusCache_(booking.type);
    removeBookingCache_(studentId);
    appendLogBook_({
      action: "CANCEL",
      status: "SUCCESS",
      studentId: booking.studentId,
      name: booking.name,
      email: booking.email,
      type: booking.type,
      roomKey: booking.roomKey,
      roomLabel: booking.roomLabel,
      message: "ยกเลิกการจองสำเร็จ"
    });

    return {
      success: true,
      message: "ยกเลิกการจองสำเร็จ"
    };
  } catch (error) {
    Logger.log("cancelBooking error: " + error.message);
    appendLogBook_({
      action: "CANCEL",
      status: "ERROR",
      studentId: studentId,
      message: error.message
    });
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message
    };
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

function getMapImageBase64() {
  try {
    if (!MAP_IMAGE_FILE_ID || MAP_IMAGE_FILE_ID === "YOUR_FILE_ID_HERE") {
      return "";
    }

    const file = DriveApp.getFileById(MAP_IMAGE_FILE_ID);
    const blob = file.getBlob();
    return "data:" + blob.getContentType() + ";base64," +
      Utilities.base64Encode(blob.getBytes());
  } catch (error) {
    Logger.log("getMapImageBase64 error: " + error.message);
    return "";
  }
}

function getBookingByStudentId_(studentId, options) {
  studentId = normalizeStudentId_(studentId);
  const useCache = !options || options.useCache !== false;

  if (!studentId) {
    return null;
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = buildBookingCacheKey_(studentId);

  if (useCache) {
    const cached = cache.get(cacheKey);
    if (cached) {
      if (cached === "null") {
        return null;
      }
      try {
        return JSON.parse(cached);
      } catch (cacheError) {
        Logger.log("getBookingByStudentId cache parse error: " + cacheError.message);
      }
    }
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  for (let groupIndex = 0; groupIndex < GROUP_ORDER.length; groupIndex++) {
    const typeKey = GROUP_ORDER[groupIndex];
    const config = ROOM_GROUPS[typeKey];
    const sheet = spreadsheet.getSheetByName(config.sheet);

    if (!sheet) {
      continue;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      continue;
    }

    ensureBookingSheetSchema_(sheet);
    const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

    for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
      const row = values[rowIndex];
      const rowStudentId = normalizeStudentId_(row[2]);

      if (rowStudentId !== studentId) {
        continue;
      }

      const roomKey = String(row[1] || "").trim();
      const roomInfo = config.rooms[roomKey] || null;
      const booking = sanitizeBooking_({
        type: typeKey,
        typeLabel: config.label,
        roomKey: roomKey,
        roomLabel: String(row[3] || (roomInfo ? roomInfo.label : "")),
        name: normalizeName_(row[4]),
        studentId: rowStudentId,
        email: normalizeEmail_(row[5]),
        timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0] || ""),
        row: rowIndex + 2,
        sheet: config.sheet,
        emoji: roomInfo ? roomInfo.emoji : "🛏️"
      });

      if (useCache) {
        setBookingCache_(studentId, booking);
      }

      return booking;
    }
  }

  if (useCache) {
    cache.put(cacheKey, "null", 30);
  }

  return null;
}

function ensureSheet_(typeKey) {
  const config = ROOM_GROUPS[typeKey];
  if (!config) {
    throw new Error("Invalid room type");
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(config.sheet);

  if (sheet) {
    ensureBookingSheetSchema_(sheet);
    return sheet;
  }

  sheet = spreadsheet.insertSheet(config.sheet);
  sheet.appendRow([
    "Timestamp",
    "Room Key",
    "Student ID",
    "Room Label",
    "Name",
    "Email",
    "Type",
    "No."
  ]);

  const header = sheet.getRange(1, 1, 1, 8);
  header
    .setBackground("#0f766e")
    .setFontColor("white")
    .setFontWeight("bold");

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 165);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 170);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 230);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 70);

  return sheet;
}

function countBookingsForRoom_(sheet, roomKey) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 0;
  }

  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  let count = 0;

  for (let index = 0; index < values.length; index++) {
    if (String(values[index][0] || "").trim() === roomKey) {
      count++;
    }
  }

  return count;
}

function sanitizeBooking_(booking) {
  if (!booking) {
    return null;
  }

  return {
    type: String(booking.type || ""),
    typeLabel: String(booking.typeLabel || ""),
    roomKey: String(booking.roomKey || ""),
    roomLabel: String(booking.roomLabel || ""),
    name: String(booking.name || ""),
    studentId: String(booking.studentId || ""),
    email: String(booking.email || ""),
    timestamp: String(booking.timestamp || ""),
    emoji: String(booking.emoji || "🛏️"),
    row: booking.row ? Number(booking.row) : null,
    sheet: String(booking.sheet || "")
  };
}

function normalizeStudentId_(value) {
  return String(value || "").replace(/\D/g, "").trim();
}

function normalizeName_(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function normalizeEmail_(value) {
  return String(value || "").trim().toLowerCase();
}

function enforceAllowedEmailDomain_(email) {
  email = normalizeEmail_(email);

  if (!email) {
    throw new Error("กรุณากรอกอีเมลให้ถูกต้อง");
  }
  if (!/@su\.ac\.th$/i.test(email)) {
    throw new Error("กรุณาใช้อีเมล @su.ac.th เท่านั้น");
  }
}

function ensureBookingSheetSchema_(sheet) {
  const expectedHeaders = [
    "Timestamp",
    "Room Key",
    "Student ID",
    "Room Label",
    "Name",
    "Email",
    "Type",
    "No."
  ];
  const currentMaxColumns = sheet.getMaxColumns();
  const legacyHeaders = sheet.getRange(1, 1, 1, 7).getValues()[0];
  const isLegacySchema =
    String(legacyHeaders[0] || "").trim() === "Timestamp" &&
    String(legacyHeaders[1] || "").trim() === "Room Key" &&
    String(legacyHeaders[2] || "").trim() === "Student ID" &&
    String(legacyHeaders[3] || "").trim() === "Room Label" &&
    String(legacyHeaders[4] || "").trim() === "Name" &&
    String(legacyHeaders[5] || "").trim() === "Type" &&
    String(legacyHeaders[6] || "").trim() === "No.";

  if (isLegacySchema) {
    sheet.insertColumnAfter(5);
  } else if (currentMaxColumns < 8) {
    sheet.insertColumnsAfter(currentMaxColumns, 8 - currentMaxColumns);
  }

  const headerValues = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];

  for (let index = 0; index < expectedHeaders.length; index++) {
    if (String(headerValues[index] || "").trim() !== expectedHeaders[index]) {
      sheet.getRange(1, index + 1).setValue(expectedHeaders[index]);
    }
  }

  sheet.setFrozenRows(1);
}

function ensureUserRegistrySheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(USER_REGISTRY_SHEET);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(USER_REGISTRY_SHEET);
  }

  const headers = ["Email", "Student ID", "Name", "Bound At", "Last Login At"];
  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];

  for (let index = 0; index < headers.length; index++) {
    if (String(currentHeaders[index] || "").trim() !== headers[index]) {
      sheet.getRange(1, index + 1).setValue(headers[index]);
    }
  }

  sheet.setFrozenRows(1);
  return sheet;
}

function ensureLogBookSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(LOG_BOOK_SHEET);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(LOG_BOOK_SHEET);
  }

  const headers = [
    "Timestamp",
    "Action",
    "Status",
    "Student ID",
    "Name",
    "Email",
    "Type",
    "Room Key",
    "Room Label",
    "Message"
  ];
  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];

  for (let index = 0; index < headers.length; index++) {
    if (String(currentHeaders[index] || "").trim() !== headers[index]) {
      sheet.getRange(1, index + 1).setValue(headers[index]);
    }
  }

  sheet.setFrozenRows(1);
  return sheet;
}

function appendLogBook_(entry) {
  try {
    const sheet = ensureLogBookSheet_();
    sheet.appendRow([
      new Date(),
      String(entry.action || ""),
      String(entry.status || ""),
      String(entry.studentId || ""),
      String(entry.name || ""),
      String(entry.email || ""),
      String(entry.type || ""),
      String(entry.roomKey || ""),
      String(entry.roomLabel || ""),
      String(entry.message || "")
    ]);
  } catch (error) {
    Logger.log("appendLogBook error: " + error.message);
  }
}

function upsertUserIdentity_(email, studentId, name) {
  email = normalizeEmail_(email);
  studentId = normalizeStudentId_(studentId);
  name = normalizeName_(name);

  const sheet = ensureUserRegistrySheet_();
  const lastRow = sheet.getLastRow();
  const values = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 5).getValues()
    : [];
  let emailRow = null;
  let studentRow = null;

  for (let index = 0; index < values.length; index++) {
    const row = values[index];
    const rowEmail = normalizeEmail_(row[0]);
    const rowStudentId = normalizeStudentId_(row[1]);

    if (rowEmail === email) {
      emailRow = { index: index + 2, studentId: rowStudentId, name: normalizeName_(row[2]) };
    }
    if (rowStudentId === studentId) {
      studentRow = { index: index + 2, email: rowEmail };
    }
  }

  if (studentRow && studentRow.email !== email) {
    throw new Error("รหัสนักศึกษานี้ถูกผูกกับอีเมลอื่นแล้ว");
  }

  const now = new Date();

  if (emailRow) {
    if (emailRow.studentId !== studentId) {
      throw new Error("อีเมลนี้ถูกผูกกับรหัสนักศึกษาอื่นแล้ว");
    }

    sheet.getRange(emailRow.index, 3, 1, 3).setValues([[
      name || emailRow.name,
      values[emailRow.index - 2][3] || now,
      now
    ]]);

    return {
      email: email,
      studentId: studentId,
      name: name || emailRow.name
    };
  }

  sheet.appendRow([email, studentId, name, now, now]);
  return {
    email: email,
    studentId: studentId,
    name: name
  };
}

function resolveUserIdentity_(email, studentId, name) {
  email = normalizeEmail_(email);
  studentId = normalizeStudentId_(studentId);
  name = normalizeName_(name);

  const sheet = ensureUserRegistrySheet_();
  const lastRow = sheet.getLastRow();
  const values = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 5).getValues()
    : [];
  let exactRow = null;
  let emailRow = null;
  let studentRow = null;

  for (let index = 0; index < values.length; index++) {
    const row = values[index];
    const rowEmail = normalizeEmail_(row[0]);
    const rowStudentId = normalizeStudentId_(row[1]);
    const rowName = normalizeName_(row[2]);
    const rowIndex = index + 2;

    if (rowEmail === email && rowStudentId === studentId) {
      exactRow = { index: rowIndex, email: rowEmail, studentId: rowStudentId, name: rowName };
      break;
    }
    if (!emailRow && rowEmail === email) {
      emailRow = { index: rowIndex, email: rowEmail, studentId: rowStudentId, name: rowName };
    }
    if (!studentRow && rowStudentId === studentId) {
      studentRow = { index: rowIndex, email: rowEmail, studentId: rowStudentId, name: rowName };
    }
  }

  if (exactRow) {
    const now = new Date();
    sheet.getRange(exactRow.index, 3, 1, 3).setValues([[
      name || exactRow.name,
      values[exactRow.index - 2][3] || now,
      now
    ]]);
    return {
      email: email,
      studentId: studentId,
      name: name || exactRow.name
    };
  }

  if (emailRow || studentRow) {
    if (emailRow && studentRow && emailRow.index === studentRow.index) {
      const now = new Date();
      sheet.getRange(emailRow.index, 3, 1, 3).setValues([[
        name || emailRow.name || studentRow.name,
        values[emailRow.index - 2][3] || now,
        now
      ]]);
      return {
        email: email,
        studentId: studentId,
        name: name || emailRow.name || studentRow.name
      };
    }
    if (studentRow && studentRow.email !== email && canRebindIdentityRow_(studentRow, name)) {
      const now = new Date();
      sheet.getRange(studentRow.index, 1, 1, 5).setValues([[
        email,
        studentId,
        name || studentRow.name,
        values[studentRow.index - 2][3] || now,
        now
      ]]);
      return {
        email: email,
        studentId: studentId,
        name: name || studentRow.name
      };
    }
    if (emailRow && emailRow.studentId !== studentId && canRebindIdentityRow_(emailRow, name)) {
      const now = new Date();
      sheet.getRange(emailRow.index, 1, 1, 5).setValues([[
        email,
        studentId,
        name || emailRow.name,
        values[emailRow.index - 2][3] || now,
        now
      ]]);
      return {
        email: email,
        studentId: studentId,
        name: name || emailRow.name
      };
    }
    if (studentRow && studentRow.email !== email) {
      throw new Error("รหัสนักศึกษานี้ถูกผูกกับอีเมลอื่นแล้ว");
    }
    if (emailRow && emailRow.studentId !== studentId) {
      throw new Error("อีเมลนี้ถูกผูกกับรหัสนักศึกษาอื่นแล้ว");
    }
  }

  return upsertUserIdentity_(email, studentId, name);
}

function canRebindIdentityRow_(row, name) {
  const existingName = normalizeName_(row && row.name);
  const inputName = normalizeName_(name);

  if (!existingName || !inputName) {
    return false;
  }

  return existingName === inputName;
}

function getUserIdentityByEmail_(email) {
  email = normalizeEmail_(email);
  if (!email) {
    return null;
  }

  const sheet = ensureUserRegistrySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  for (let index = 0; index < values.length; index++) {
    const row = values[index];
    if (normalizeEmail_(row[0]) === email) {
      sheet.getRange(index + 2, 5).setValue(new Date());
      return {
        email: email,
        studentId: normalizeStudentId_(row[1]),
        name: normalizeName_(row[2])
      };
    }
  }

  return null;
}

function backfillBookingEmailIfMissing_(booking, email) {
  if (!booking || normalizeEmail_(booking.email) || !booking.row || !booking.sheet) {
    return booking;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(booking.sheet);
  if (!sheet) {
    return booking;
  }

  ensureBookingSheetSchema_(sheet);
  sheet.getRange(Number(booking.row), 6).setValue(normalizeEmail_(email));
  booking.email = normalizeEmail_(email);
  setBookingCache_(booking.studentId, booking);
  return booking;
}

function buildRoomStatusCacheKey_(typeKey) {
  return "room_status:" + typeKey;
}

function buildBookingCacheKey_(studentId) {
  return "booking:" + studentId;
}

function invalidateRoomStatusCache_(typeKey) {
  CacheService.getScriptCache().remove(buildRoomStatusCacheKey_(typeKey));
}

function setBookingCache_(studentId, booking) {
  CacheService.getScriptCache().put(
    buildBookingCacheKey_(studentId),
    JSON.stringify(sanitizeBooking_(booking)),
    BOOKING_CACHE_TTL
  );
}

function removeBookingCache_(studentId) {
  CacheService.getScriptCache().remove(buildBookingCacheKey_(studentId));
}

function sanitizeHtml_(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

