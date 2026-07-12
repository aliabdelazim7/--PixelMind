// دالة doGet لتشغيل وعرض واجهة المستخدم الرسومية الخاصة بالـ CRM أو إرجاع البيانات بصيغة JSON
function doGet(e) {
  // إذا كان الطلب استدعاء للبيانات عبر الـ API
  if (e && e.parameter && e.parameter.action === 'getLeads') {
    var data = getLeads();
    return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
  }
  
  // الوضع الافتراضي: عرض صفحة الويب للـ CRM
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ليد فلو - نظام إدارة العملاء البسيط')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// دالة doPost لاستقبال الإضافات والتعديلات والحذف من خارج الشيت عبر الـ API
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

// دالة لجلب العملاء من جدول البيانات الحالي (Google Sheet)
function getLeads() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
}

// دالة لإضافة عميل جديد أو تعديل عميل موجود في الشيت
function addOrUpdateLead(lead) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowNum = parseInt(rowIndex);
  sheet.deleteRow(rowNum);
  
  return getLeads(); 
}
