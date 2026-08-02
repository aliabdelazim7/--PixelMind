// ===================================================================
//  ليد فلو (LeadFlow) — Google Apps Script Backend
// ===================================================================
//  الأسرار (التوكنات) لم تعد مكتوبة داخل الكود. يتم قراءتها من
//  Script Properties. لضبطها: Project Settings > Script Properties
//  المطلوب ضبطه:
//    - API_SECRET          : كلمة سر مشتركة تحمي عمليات الكتابة (doPost)
//    - SHEET_NAME          : اسم تبويب العملاء (اختياري، الافتراضي أول تبويب)
//    - TELEGRAM_BOT_TOKEN  : توكن بوت تلجرام
//    - TELEGRAM_CHAT_IDS   : معرفات المحادثات مفصولة بفواصل، مثال: 7416290524,5507184715
//    - META_PIXEL_ID       : معرف بكسل فيسبوك
//    - META_ACCESS_TOKEN   : توكن Meta Conversions API
// ===================================================================

// رابط شيت جوجل الخاص بك (ليس سراً — مجرد عنوان المستند)
const SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1VdSlXUBJtpMqkuow4Fg98XGrVpwFhRMc_X9QvsI9H_s/edit";

// ---------- أدوات مساعدة ----------

// قراءة سر من Script Properties بأمان
function getProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

// إخراج JSON موحّد
function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
}

// تحييد حقن الصيغ في جداول جوجل (Formula/CSV Injection)
// أي نص يبدأ بـ = + - @ يُسبق بعلامة اقتباس ليُخزَّن كنص لا كصيغة
function sanitizeCell_(v) {
  var s = (v === null || v === undefined) ? "" : String(v);
  if (/^[=+\-@\t\r]/.test(s)) {
    return "'" + s;
  }
  return s;
}

// تشفير SHA-256 وإرجاعه Hex (لبيانات ميتا CAPI)
function sha256Hex_(str) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(str || ""), Utilities.Charset.UTF_8);
  var out = "";
  for (var i = 0; i < raw.length; i++) {
    var b = (raw[i] < 0 ? raw[i] + 256 : raw[i]).toString(16);
    out += (b.length === 1 ? "0" + b : b);
  }
  return out;
}

// ---------- نقاط الدخول ----------

// جلب البيانات (قراءة فقط)
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getLeads') {
    return jsonOut_(getLeads());
  }
  return ContentService.createTextOutput("ليد فلو API يعمل بنجاح!")
      .setMimeType(ContentService.MimeType.TEXT);
}

// استقبال الكتابة/التعديل/الحذف والبروكسي (ميتا/تلجرام)
function doPost(e) {
  try {
    var requestData = JSON.parse(e.postData.contents);

    // حماية عمليات الكتابة بكلمة سر مشتركة (تُفعَّل فور ضبط API_SECRET)
    var expectedSecret = getProp_('API_SECRET');
    if (expectedSecret && requestData.secret !== expectedSecret) {
      return jsonOut_({ error: 'Unauthorized' });
    }
    if (!expectedSecret) {
      Logger.log('تحذير أمني: لم يتم ضبط API_SECRET في Script Properties — عمليات الكتابة مفتوحة.');
    }

    var action = requestData.action;
    var result;

    if (action === 'addOrUpdate') {
      result = addOrUpdateLead(requestData.lead);
    } else if (action === 'delete') {
      result = deleteLead(requestData.leadId);
    } else if (action === 'addBulk') {
      result = addLeadsBulk(requestData.leads);
    } else if (action === 'deleteBulk') {
      result = deleteLeadsBulk(requestData.leadIds);
    } else if (action === 'updateStatusBulk') {
      result = updateStatusBulk(requestData.leadIds, requestData.status);
    } else if (action === 'metaEvent') {
      result = sendMetaEvent(requestData.eventName, requestData.lead);
    } else if (action === 'telegram') {
      result = sendTelegramText(requestData.text);
    } else {
      result = { error: 'Action parameter invalid' };
    }

    return jsonOut_(result);
  } catch (error) {
    // لا نسرّب تفاصيل الخطأ الداخلية للعميل
    Logger.log('doPost error: ' + error.toString());
    return jsonOut_({ error: 'Internal error' });
  }
}

// الحصول على تبويب العملاء بشكل حتمي (اسم ثابت من الخصائص وإلا أول تبويب)
function getTargetSheet() {
  var ss = (SPREADSHEET_URL && SPREADSHEET_URL !== "")
      ? SpreadsheetApp.openByUrl(SPREADSHEET_URL)
      : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("لم يتم العثور على جدول بيانات. ضع رابط الشيت في SPREADSHEET_URL.");
  }
  var name = getProp_('SHEET_NAME');
  var sheet = name ? ss.getSheetByName(name) : ss.getSheets()[0];
  if (!sheet) {
    throw new Error("لم يتم العثور على تبويب العملاء (" + (name || "الأول") + ").");
  }
  return sheet;
}

// ---------- القراءة ----------

function getLeads() {
  try {
    var sheet = getTargetSheet();
    checkAndInitHeaders(sheet);

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var leads = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[1] && !row[2]) continue;

      var appTimeVal = "";
      if (row[5]) {
        var strVal = String(row[5]);
        if (Object.prototype.toString.call(row[5]) === '[object Date]' || typeof row[5].getTime === 'function') {
          appTimeVal = Utilities.formatDate(new Date(row[5]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
        } else if (strVal.indexOf('GMT') !== -1 || strVal.indexOf('توقيت') !== -1) {
          try {
            appTimeVal = Utilities.formatDate(new Date(strVal), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
          } catch (e) {
            appTimeVal = strVal;
          }
        } else {
          appTimeVal = strVal.trim();
        }
      }

      leads.push({
        id: row[0] || "",
        fullname: row[1] || "",
        phone: row[2] || "",
        status: row[3] || "",
        created_at: row[4] ? Utilities.formatDate(new Date(row[4]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
        appointment_time: appTimeVal,
        notes: row[6] || ""
      });
    }
    return leads;
  } catch (err) {
    Logger.log('getLeads error: ' + err.toString());
    return { error: 'Failed to read leads' };
  }
}

function findRowById(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      return i + 1;
    }
  }
  return -1;
}

// ---------- الكتابة (محميّة بقفل لمنع تعارض العمليات المتزامنة) ----------

function addOrUpdateLead(lead) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getTargetSheet();
    checkAndInitHeaders(sheet);

    var cleanName = sanitizeCell_(lead.fullname);
    var cleanNotes = sanitizeCell_(lead.notes || "");
    var phone = lead.phone || "";
    var status = lead.status || "";
    var appt = lead.appointment_time || "";

    // 1) تحديث بالمعرّف
    if (lead.id) {
      var rowNum = findRowById(sheet, lead.id);
      if (rowNum !== -1) {
        sheet.getRange(rowNum, 2).setValue(cleanName);
        sheet.getRange(rowNum, 3).setValue(phone);
        sheet.getRange(rowNum, 4).setValue(status);
        sheet.getRange(rowNum, 6).setValue(appt);
        sheet.getRange(rowNum, 7).setValue(cleanNotes);
        return getLeads();
      }
    }

    // 2) منع التكرار بالاسم + الهاتف
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var sheetFullname = String(data[i][1]).trim();
      var sheetPhone = String(data[i][2]).trim();
      if (sheetFullname === String(lead.fullname).trim() && sheetPhone === String(phone).trim()) {
        var r = i + 1;
        sheet.getRange(r, 4).setValue(status);
        sheet.getRange(r, 6).setValue(appt);
        sheet.getRange(r, 7).setValue(cleanNotes);
        return getLeads();
      }
    }

    // 3) إضافة صف جديد بمعرّف فريد
    var uniqueId = lead.id || genId_();
    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.appendRow([uniqueId, cleanName, phone, status, dateStr, appt, cleanNotes]);
    return getLeads();
  } finally {
    lock.releaseLock();
  }
}

// معرّف فريد أقوى (وقت + عشوائي أطول) لتفادي التصادم في الإضافة الجماعية
function genId_() {
  return "L-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000000);
}

function deleteLead(leadId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getTargetSheet();
    var rowNum = findRowById(sheet, leadId);
    if (rowNum !== -1) {
      sheet.deleteRow(rowNum);
    }
    return getLeads();
  } finally {
    lock.releaseLock();
  }
}

function checkAndInitHeaders(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow === 0 || (lastRow === 1 && sheet.getRange(1, 1).getValue() === "")) {
    sheet.getRange(1, 1, 1, 7).setValues([["المعرف", "الاسم بالكامل", "رقم الهاتف", "الحالة", "تاريخ الإضافة", "تاريخ ووقت الموعد", "الملاحظات"]]);
    var headerRange = sheet.getRange(1, 1, 1, 7);
    headerRange.setFontWeight("bold").setBackground("#efefef").setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 7);
  }
}

function addLeadsBulk(newLeads) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getTargetSheet();
    checkAndInitHeaders(sheet);

    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var data = sheet.getDataRange().getValues();

    var existing = {};
    for (var i = 1; i < data.length; i++) {
      var key = String(data[i][1]).trim() + "_" + String(data[i][2]).trim();
      existing[key] = true;
    }

    var rowsToAppend = [];
    newLeads.forEach(function (lead) {
      var key = String(lead.fullname).trim() + "_" + String(lead.phone).trim();
      if (!existing[key]) {
        rowsToAppend.push([
          lead.id || genId_(),
          sanitizeCell_(lead.fullname),
          lead.phone || "",
          lead.status || "New Lead",
          dateStr,
          lead.appointment_time || "",
          sanitizeCell_(lead.notes || "")
        ]);
        existing[key] = true;
      }
    });

    // كتابة دفعة واحدة بدل سطر بسطر (أسرع وأخف على الحصة)
    if (rowsToAppend.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rowsToAppend.length, 7).setValues(rowsToAppend);
    }
    return getLeads();
  } finally {
    lock.releaseLock();
  }
}

function deleteLeadsBulk(leadIds) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getTargetSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      var rowId = String(data[i][0]);
      if (leadIds.indexOf(rowId) !== -1) {
        sheet.deleteRow(i + 1);
      }
    }
    return getLeads();
  } finally {
    lock.releaseLock();
  }
}

// تغيير حالة مجموعة عملاء دفعة واحدة (مع الحفظ في الشيت)
function updateStatusBulk(leadIds, newStatus) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getTargetSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (leadIds.indexOf(String(data[i][0])) !== -1) {
        sheet.getRange(i + 1, 4).setValue(newStatus);
      }
    }
    return getLeads();
  } finally {
    lock.releaseLock();
  }
}

// ---------- بروكسي ميتا CAPI (التوكن يبقى على السيرفر فقط) ----------

function sendMetaEvent(eventName, lead) {
  try {
    var pixelId = getProp_('META_PIXEL_ID');
    var token = getProp_('META_ACCESS_TOKEN');
    if (!pixelId || !token) {
      return { ok: false, error: 'META_NOT_CONFIGURED' };
    }
    lead = lead || {};
    var phone = String(lead.phone || "").replace(/[\+\s\-\(\)]/g, '');
    var name = String(lead.fullname || "").trim().toLowerCase();

    var evt = {
      event_name: eventName || "Lead",
      event_time: Math.floor(new Date().getTime() / 1000),
      action_source: "website",
      user_data: {
        ph: [sha256Hex_(phone)],
        fn: [sha256Hex_(name)]
      }
    };
    // قيمة الصفقة تُرسل فقط مع حدث الشراء الفعلي
    if (evt.event_name === 'Purchase') {
      evt.custom_data = { value: 1.0, currency: "EGP" };
    }

    var url = "https://graph.facebook.com/v19.0/" + pixelId + "/events?access_token=" + encodeURIComponent(token);
    var res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ data: [evt] }),
      muteHttpExceptions: true
    });
    return { ok: true, code: res.getResponseCode() };
  } catch (e) {
    Logger.log('sendMetaEvent error: ' + e.toString());
    return { ok: false, error: 'META_SEND_FAILED' };
  }
}

// ---------- بروكسي تلجرام (التوكن يبقى على السيرفر فقط) ----------

function sendTelegramText(text) {
  if (!text) return { ok: false, error: 'EMPTY_TEXT' };
  sendTelegramMessageToAll(text);
  return { ok: true };
}

// ==================== التقرير اليومي التلقائي ====================

function onOpen() {
  try {
    setupDailyReportTrigger();
  } catch (e) {
    Logger.log("فشل إعداد مشغل التقرير اليومي: " + e.toString());
  }
}

function setupDailyReportTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendDailyReportTelegram') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('sendDailyReportTelegram').timeBased().everyDays(1).atHour(0).create();
}

function sendDailyReportTelegram() {
  try {
    var sheet = getTargetSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    var yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    var targetDateStr = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "yyyy-MM-dd");

    var addedToday = 0;
    var statusCounts = {};

    for (var i = 1; i < data.length; i++) {
      var rowDate = "";
      if (data[i][4]) {
        try {
          rowDate = Utilities.formatDate(new Date(data[i][4]), Session.getScriptTimeZone(), "yyyy-MM-dd");
        } catch (err) {
          rowDate = String(data[i][4]).split(" ")[0];
        }
      }
      if (rowDate === targetDateStr) {
        addedToday++;
        var status = data[i][3] || "New Lead";
        statusCounts[status] = (statusCounts[status] || 0) + 1;
      }
    }

    var text = "📅 <b>التقرير اليومي ليد فلو - " + targetDateStr + "</b>\n\n" +
               "➕ <b>العملاء الجدد المضافين اليوم:</b> " + addedToday + "\n";

    if (addedToday > 0) {
      text += "\n📊 <b>تصنيف العملاء الجدد اليوم:</b>\n";
      for (var key in statusCounts) {
        text += "  • " + translateStatusToArabic(key) + ": " + statusCounts[key] + "\n";
      }
    }

    var totalLeads = data.length - 1;
    var totalWon = 0;
    for (var j = 1; j < data.length; j++) {
      if (data[j][3] === 'Won') totalWon++;
    }
    var generalConversion = totalLeads > 0 ? Math.round((totalWon / totalLeads) * 100) : 0;

    text += "\n💼 <b>إحصائيات النظام العامة:</b>\n" +
            "🗂️ <b>إجمالي عملاء النظام:</b> " + totalLeads + " عميل\n" +
            "🟢 <b>إجمالي الصفقات الناجحة (Won):</b> " + totalWon + "\n" +
            "🎯 <b>معدل التحويل العام:</b> " + generalConversion + "%\n\n" +
            "✨ طابت ليلتكم! تم إرسال التقرير تلقائياً.";

    sendTelegramMessageToAll(text);
  } catch (e) {
    Logger.log("فشل إرسال التقرير اليومي: " + e.toString());
  }
}

// إرسال رسالة لكل المعرفات (التوكن والمعرفات من Script Properties)
function sendTelegramMessageToAll(text) {
  var botToken = getProp_('TELEGRAM_BOT_TOKEN');
  var chatIdsRaw = getProp_('TELEGRAM_CHAT_IDS') || "";
  var chatIds = chatIdsRaw.split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; });

  if (!botToken || chatIds.length === 0) {
    Logger.log('تلجرام غير مضبوط: تأكد من TELEGRAM_BOT_TOKEN و TELEGRAM_CHAT_IDS.');
    return;
  }

  chatIds.forEach(function (chatId) {
    try {
      var url = "https://api.telegram.org/bot" + botToken + "/sendMessage";
      var payload = { "chat_id": chatId, "text": text, "parse_mode": "HTML" };
      UrlFetchApp.fetch(url, {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
      });
    } catch (e) {
      Logger.log("فشل إرسال للمعرف " + chatId + ": " + e.toString());
    }
  });
}

function translateStatusToArabic(status) {
  var mapping = {
    "New Lead": "⚪ عميل جديد",
    "Contacted": "🔵 تم الاتصال",
    "Follow Up Required": "🟠 يحتاج متابعة",
    "Appointment Booked": "📅 تم حجز موعد",
    "Proposal Sent": "📄 تم إرسال العرض",
    "Negotiating": "🟡 تفاوض",
    "In Progress": "⚡ جاري العمل",
    "Won": "🟢 تم البيع بنجاح (Won)",
    "Lost": "🔴 صفقة خاسرة (Lost)",
    "Ghosted": "⚫ لم يرد / اختفى (Ghosted)",
    "Not Interested": "❌ غير مهتم"
  };
  return mapping[status] || status;
}
