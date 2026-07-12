// دالة doGet لتشغيل وعرض واجهة المستخدم الرسومية الخاصة بالـ CRM
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ليد فلو - نظام إدارة العملاء البسيط')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// دالة لجلب العملاء من جدول البيانات الحالي (Google Sheet)
function getLeads() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // إذا كان الجدول فارغاً أو يحتوي فقط على العناوين
  
  var leads = [];
  
  // حلقة تكرارية تبدأ من الصف الثاني (لتخطي العناوين)
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
  
  // إذا كان العميل يمتلك رقم صف (rowIndex) فهذا يعني عملية تعديل
  if (lead.rowIndex) {
    var rowNum = parseInt(lead.rowIndex);
    sheet.getRange(rowNum, 1).setValue(lead.fullname);
    sheet.getRange(rowNum, 2).setValue(lead.phone);
    sheet.getRange(rowNum, 3).setValue(lead.status);
  } else {
    // إذا لم يمتلك رقم صف، نقوم بإنشائه كعميل جديد في آخر الجدول
    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.appendRow([lead.fullname, lead.phone, lead.status, dateStr]);
  }
  
  return getLeads(); // إرجاع القائمة المحدثة
}

// دالة لحذف عميل من الشيت بناءً على رقم الصف
function deleteLead(rowIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowNum = parseInt(rowIndex);
  sheet.deleteRow(rowNum);
  
  return getLeads(); // إرجاع القائمة المحدثة
}
