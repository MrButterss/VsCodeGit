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
const ROOM_STATUS_CACHE_TTL = 20;
const BOOKING_CACHE_TTL = 120;

function doGet() {
  try {
    const template = HtmlService.createTemplateFromFile("Index");
    template.roomCatalogJson = JSON.stringify(getRoomCatalog());

    return template
      .evaluate()
      .setTitle("ระบบจองห้องพัก")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

function login(studentId, name) {
  try {
    studentId = normalizeStudentId_(studentId);
    name = normalizeName_(name);

    if (!name || !studentId) {
      return {
        success: false,
        message: "กรุณากรอกชื่อและรหัสนักศึกษาให้ครบ"
      };
    }

    if (!/^\d{9}$/.test(studentId)) {
      return {
        success: false,
        message: "รหัสนักศึกษาต้องเป็นตัวเลข 9 หลัก"
      };
    }

    const booking = getBookingByStudentId_(studentId, { useCache: true });

    return {
      success: true,
      studentId: studentId,
      name: name,
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

function submitBooking(studentId, name, roomKey, typeKey) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;

  try {
    if (!lock.tryLock(25000)) {
      return {
        success: false,
        message: "ระบบกำลังประมวลผลคำขอจำนวนมาก กรุณาลองใหม่อีกครั้ง"
      };
    }
    lockAcquired = true;

    studentId = normalizeStudentId_(studentId);
    name = normalizeName_(name);
    roomKey = String(roomKey || "").trim();
    typeKey = String(typeKey || "").trim();

    if (!studentId || !name || !roomKey || !typeKey) {
      return { success: false, message: "ข้อมูลไม่ครบ" };
    }

    if (!/^\d{9}$/.test(studentId)) {
      return {
        success: false,
        message: "รหัสนักศึกษาต้องเป็นตัวเลข 9 หลัก"
      };
    }

    const existing = getBookingByStudentId_(studentId, { useCache: false });
    if (existing) {
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
      return { success: false, message: "ไม่พบห้องนี้" };
    }

    const sheet = ensureSheet_(typeKey);
    const bookedCount = countBookingsForRoom_(sheet, roomKey);

    if (bookedCount >= Number(roomConfig.total || 0)) {
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
      emoji: roomConfig.emoji,
      timestamp: now.toISOString(),
      sheet: config.sheet
    });

    invalidateRoomStatusCache_(typeKey);
    setBookingCache_(studentId, booking);

    return {
      success: true,
      message: "จองสำเร็จ",
      booking: booking
    };
  } catch (error) {
    Logger.log("submitBooking error: " + error.message);
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

  try {
    if (!lock.tryLock(25000)) {
      return {
        success: false,
        message: "ระบบกำลังประมวลผลคำขอจำนวนมาก กรุณาลองใหม่อีกครั้ง"
      };
    }
    lockAcquired = true;

    studentId = normalizeStudentId_(studentId);
    if (!studentId) {
      return { success: false, message: "ไม่พบรหัสนักศึกษา" };
    }

    const booking = getBookingByStudentId_(studentId, { useCache: false });
    if (!booking) {
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

    return {
      success: true,
      message: "ยกเลิกการจองสำเร็จ"
    };
  } catch (error) {
    Logger.log("cancelBooking error: " + error.message);
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

    const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

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
    return sheet;
  }

  sheet = spreadsheet.insertSheet(config.sheet);
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
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 70);

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

