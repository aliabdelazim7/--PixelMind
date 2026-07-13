// ضع رابط شيت جوجل (Google Sheet URL) الخاص بك هنا للربط المضمون والسريع
// مثال: "https://docs.google.com/spreadsheets/d/.../edit"
const SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1VdSlXUBJtpMqkuow4Fg98XGrVpwFhRMc_X9QvsI9H_s/edit"; 

// دالة doGet لاستقبال طلبات جلب البيانات من Vercel أو الفتح المباشر
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getLeads') {
    var data = getLeads();
    return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput("ليد فلو API يعمل بنجاح! قاعدة البيانات متصلة ومستعدة لاستقبال البيانات من Vercel.")
      .setMimeType(ContentService.MimeType.TEXT);
}

// دالة doPost لاستقبال الإضافات والتعديلات والحذف من Vercel عبر الـ API
function doPost(e) {
  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    var result;
    
    if (action === 'addOrUpdate') {
      result = addOrUpdateLead(requestData.lead);
    } else if (action === 'delete') {
      result = deleteLead(requestData.leadId);
    } else {
      result = { error: 'Action parameter invalid' };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
        
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
  }
}

// دالة مساعدة للحصول على شيت جوجل النشط بأمان
function getTargetSheet() {
  if (SPREADSHEET_URL && SPREADSHEET_URL !== "") {
    return SpreadsheetApp.openByUrl(SPREADSHEET_URL).getActiveSheet();
  }
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSpreadsheet) {
    return activeSpreadsheet.getActiveSheet();
  }
  throw new Error("لم يتم العثور على جدول بيانات نشط. يرجى وضع رابط شيت جوجل الخاص بك في متغير SPREADSHEET_URL بأعلى الكود.");
}

// دالة لجلب العملاء من جدول البيانات الحالي (Google Sheet)
function getLeads() {
  try {
    var sheet = getTargetSheet();
    checkAndInitHeaders(sheet);
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; 
    
    var leads = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // تخطي الصفوف الفارغة تماماً في شيت جوجل
      if (!row[0] && !row[1] && !row[2]) continue;
      
      leads.push({
        id: row[0] || "",
        fullname: row[1] || "",
        phone: row[2] || "",
        status: row[3] || "",
        created_at: row[4] ? Utilities.formatDate(new Date(row[4]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
        appointment_time: (row[5] instanceof Date) ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : String(row[5] || "").trim() // تنسيق الوقت لـ YYYY-MM-DD HH:MM لمنع ترحيل الساعات
      });
    }
    return leads;
  } catch (err) {
    return { error: err.toString() };
  }
}

// دالة للبحث عن رقم الصف باستخدام المعرف الفريد (ID) لمنع أخطاء تداخل وحذف الأسطر
function findRowById(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      return i + 1; // إرجاع رقم الصف 1-indexed
    }
  }
  return -1;
}

// دالة لإضافة عميل جديد أو تعديل عميل موجود في الشيت مع حماية كاملة من التكرار
function addOrUpdateLead(lead) {
  var sheet = getTargetSheet();
  checkAndInitHeaders(sheet);
  
  // 1. إذا كان العميل يمتلك معرفاً فريداً، نقوم بالبحث عنه وتحديث صفه
  if (lead.id) {
    var rowNum = findRowById(sheet, lead.id);
    if (rowNum !== -1) {
      sheet.getRange(rowNum, 2).setValue(lead.fullname);
      sheet.getRange(rowNum, 3).setValue(lead.phone);
      sheet.getRange(rowNum, 4).setValue(lead.status);
      sheet.getRange(rowNum, 6).setValue(lead.appointment_time || ""); // حفظ الموعد في العمود السادس (F)
      return getLeads();
    }
  }
  
  // 2. فحص مكررات الأسماء والأرقام لتفادي التكرار عند حدوث مشاكل الشبكة أو المزامنة الأوفلاين
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var sheetFullname = String(data[i][1]).trim();
    var sheetPhone = String(data[i][2]).trim();
    if (sheetFullname === String(lead.fullname).trim() && sheetPhone === String(lead.phone).trim()) {
      // العميل موجود بالفعل بنفس الاسم ورقم الهاتف! نقوم بتحديث حالته وموعده فقط لمنع التكرار
      var rowNum = i + 1;
      sheet.getRange(rowNum, 4).setValue(lead.status);
      sheet.getRange(rowNum, 6).setValue(lead.appointment_time || "");
      return getLeads();
    }
  }
  
  // 3. إنشاء معرف فريد للعميل الجديد وإضافته كصف جديد
  var uniqueId = lead.id || ("L-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000));
  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  sheet.appendRow([uniqueId, lead.fullname, lead.phone, lead.status, dateStr, lead.appointment_time || ""]);
  
  return getLeads(); 
}

// دالة لحذف عميل من الشيت باستخدام المعرف الفريد (ID)
function deleteLead(leadId) {
  var sheet = getTargetSheet();
  var rowNum = findRowById(sheet, leadId);
  if (rowNum !== -1) {
    sheet.deleteRow(rowNum);
  }
  return getLeads(); 
}

// دالة المساعدة للتحقق وإنشاء العناوين تلقائياً وتنسيق جدول البيانات
function checkAndInitHeaders(sheet) {
  var lastRow = sheet.getLastRow();
  
  // إذا كان الشيت فارغاً تماماً
  if (lastRow === 0 || (lastRow === 1 && sheet.getRange(1, 1).getValue() === "")) {
    // كتابة العناوين باللغة العربية مع إدراج المعرف الفريد وتاريخ الموعد
    sheet.getRange(1, 1, 1, 6).setValues([["المعرف", "الاسم بالكامل", "رقم الهاتف", "الحالة", "تاريخ الإضافة", "تاريخ ووقت الموعد"]]);
    
    // تنسيق الصف الأول (عريض، خلفية رمادية فاتحة، محاذاة في المنتصف)
    var headerRange = sheet.getRange(1, 1, 1, 6);
    headerRange.setFontWeight("bold")
               .setBackground("#efefef")
               .setHorizontalAlignment("center");
               
    // تجميد الصف الأول ليبقى ظاهراً أثناء التمرير
    sheet.setFrozenRows(1);
    
    // ضبط تلقائي لعرض الأعمدة لتناسب حجم النصوص
    sheet.autoResizeColumns(1, 6);
  }
}
