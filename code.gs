// ضع رابط شيت جوجل (Google Sheet URL) الخاص بك هنا للربط المضمون والسريع
// مثال: "https://docs.google.com/spreadsheets/d/.../edit"
const SPREADSHEET_URL = ""; 

// دالة doGet لاستقبال طلبات جلب البيانات من Vercel أو الفتح المباشر
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getLeads') {
    var data = getLeads();
    return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
  }
  
  // رسالة تأكيدية بدلاً من تحميل ملف HTML خارجي
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
      result = deleteLead(requestData.rowIndex);
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
    
    // التحقق المباشر وإنشاء العناوين وتنسيق الشيت تلقائياً إذا كان فارغاً
    checkAndInitHeaders(sheet);
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; // إذا كان الجدول فارغاً أو يحتوي فقط على العناوين
    
    var leads = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      leads.push({
        rowIndex: i + 1, // حفظ رقم الصف للتعديل أو الحذف لاحقاً
        fullname: row[0] || "",
        phone: row[1] || "",
        status: row[2] || "",
        created_at: row[3] ? Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "yyyy-MM-dd") : ""
      });
    }
    return leads;
  } catch (err) {
    return { error: err.toString() };
  }
}

// دالة لإضافة عميل جديد أو تعديل عميل موجود في الشيت
function addOrUpdateLead(lead) {
  var sheet = getTargetSheet();
  
  // التحقق المباشر وإنشاء العناوين وتنسيق الشيت تلقائياً إذا كان فارغاً
  checkAndInitHeaders(sheet);
  
  if (lead.rowIndex) {
    var rowNum = parseInt(lead.rowIndex);
    sheet.getRange(rowNum, 1).setValue(lead.fullname);
    sheet.getRange(rowNum, 2).setValue(lead.phone);
    sheet.getRange(rowNum, 3).setValue(lead.status);
  } else {
    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.appendRow([lead.fullname, lead.phone, lead.status, dateStr]);
  }
  
  return getLeads(); 
}

// دالة لحذف عميل من الشيت بناءً على رقم الصف
function deleteLead(rowIndex) {
  var sheet = getTargetSheet();
  var rowNum = parseInt(rowIndex);
  sheet.deleteRow(rowNum);
  
  return getLeads(); 
}

// دالة المساعدة للتحقق وإنشاء العناوين تلقائياً وتنسيق جدول البيانات
function checkAndInitHeaders(sheet) {
  var lastRow = sheet.getLastRow();
  
  // إذا كان الشيت فارغاً تماماً
  if (lastRow === 0 || (lastRow === 1 && sheet.getRange(1, 1).getValue() === "")) {
    // كتابة العناوين باللغة العربية في الصف الأول
    sheet.getRange(1, 1, 1, 4).setValues([["الاسم بالكامل", "رقم الهاتف", "الحالة", "تاريخ الإضافة"]]);
    
    // تنسيق احترافي للصف الأول (عريض، خلفية رمادية فاتحة، محاذاة في المنتصف)
    var headerRange = sheet.getRange(1, 1, 1, 4);
    headerRange.setFontWeight("bold")
               .setBackground("#efefef")
               .setHorizontalAlignment("center");
               
    // تجميد الصف الأول ليبقى ظاهراً أثناء التمرير
    sheet.setFrozenRows(1);
    
    // ضبط تلقائي لعرض الأعمدة لتناسب حجم النصوص
    sheet.autoResizeColumns(1, 4);
  }
}
