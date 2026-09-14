/**
 * سكربت الربط بين "صفحة تقييم الإدارات" و "داشبورد الإدارات"
 * جامعة تبوك — عمادة شؤون الطلاب
 *
 * طريقة الاستخدام:
 * 1) أنشئ Google Sheet جديد (فاضي).
 * 2) من القائمة: الإضافات (Extensions) > Apps Script.
 * 3) احذف أي كود موجود، ولصق هذا الملف كامل.
 * 4) من الأعلى: Deploy > New deployment > اختر النوع "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5) اضغط Deploy، وسينشئ لك رابط مثل:
 *    https://script.google.com/macros/s/XXXXXXXX/exec
 * 6) هذا الرابط هو API_URL — ضعه بنفس القيمة في كل من:
 *    - ملف "تقييم-الادارات.html"  (متغيّر API_URL)
 *    - ملف "داشبورد-الادارات.html" (متغيّر API_URL)
 *    لازم يكون نفس الرابط بالضبط في الملفين.
 *
 * ملاحظة: أول مرة تنشر السكربت غالبًا سيطلب منك صلاحيات — وافق عليها.
 * أي تعديل على الكود لاحقًا يتطلب: Deploy > Manage deployments > تعديل (Edit) > رفع نسخة جديدة (New version) > Deploy.
 */

const SHEET_RATINGS   = "Ratings";
const SHEET_COMPLAINTS = "Complaints";
const SHEET_NAMES     = "Names";

/* ================= توكن الحماية =================
 * غيّر هذي القيمة لأي نص عشوائي طويل تختاره أنت (حروف وأرقام، بدون مسافات).
 * نفس القيمة بالضبط لازم تُنسخ لمتغيّر API_TOKEN بملفي:
 * departments-evaluation.html و departments-dashboard.html
 */
const SECRET_TOKEN = "TabukSA2026-XyQ9pLmZ7";

function isAuthorized_(token){
  return token === SECRET_TOKEN;
}

function getSheet_(name, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if(!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

function ratingsSheet_(){
  return getSheet_(SHEET_RATINGS, ["date","deptId","deptName","service","raterType","quality","speed","staff","clarity","access","comment","complaintType","complaintText"]);
}
function complaintsSheet_(){
  return getSheet_(SHEET_COMPLAINTS, ["date","deptId","complaintType","complaintText"]);
}
function namesSheet_(){
  return getSheet_(SHEET_NAMES, ["deptId","name"]);
}

function jsonOut_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- قراءة البيانات (تُستخدم من صفحة التقييم وصفحة الداشبورد) ---------------- */
function doGet(e){
  const token = e && e.parameter ? e.parameter.token : "";
  if(!isAuthorized_(token)){
    return jsonOut_({ ok:false, error:"unauthorized" });
  }

  const namesRows = namesSheet_().getDataRange().getValues();
  const names = {};
  for(let i=1;i<namesRows.length;i++){
    const [deptId, name] = namesRows[i];
    if(deptId !== "" && deptId !== undefined) names[deptId] = name;
  }

  const ratingsRows = ratingsSheet_().getDataRange().getValues();
  const ratings = [];
  for(let i=1;i<ratingsRows.length;i++){
    const r = ratingsRows[i];
    if(r[1] === "" || r[1] === undefined) continue;
    ratings.push({
      date: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      deptId: r[1], deptName: r[2] || "", service: r[3] || "", raterType: r[4] || "",
      quality: r[5], speed: r[6], staff: r[7], clarity: r[8], access: r[9],
      comment: r[10] || "", complaintType: r[11] || "", complaintText: r[12] || ""
    });
  }

  const complaintsRows = complaintsSheet_().getDataRange().getValues();
  const complaints = [];
  for(let i=1;i<complaintsRows.length;i++){
    const c = complaintsRows[i];
    if(c[1] === "" || c[1] === undefined) continue;
    complaints.push({
      date: c[0] instanceof Date ? c[0].toISOString() : String(c[0]),
      deptId: c[1], complaintType: c[2] || "", complaintText: c[3] || ""
    });
  }

  return jsonOut_({ ok:true, names, ratings, complaints });
}

/* ---------------- استقبال البيانات (تقييم جديد / شكوى يدوية / تعديل اسم إدارة) ---------------- */
function doPost(e){
  let data;
  try{
    data = JSON.parse(e.postData.contents);
  }catch(err){
    return jsonOut_({ ok:false, error:"invalid_json" });
  }

  if(!isAuthorized_(data.token)){
    return jsonOut_({ ok:false, error:"unauthorized" });
  }

  const now = new Date();

  if(data.action === "rename"){
    const sh = namesSheet_();
    const rows = sh.getDataRange().getValues();
    let found = false;
    for(let i=1;i<rows.length;i++){
      if(String(rows[i][0]) === String(data.deptId)){
        sh.getRange(i+1, 2).setValue(data.name);
        found = true;
        break;
      }
    }
    if(!found) sh.appendRow([data.deptId, data.name]);
    return jsonOut_({ ok:true });
  }

  if(data.action === "addComplaint"){
    complaintsSheet_().appendRow([now, data.deptId, data.complaintType || "", data.complaintText || ""]);
    return jsonOut_({ ok:true });
  }

  // الحالة الافتراضية: تقييم جديد من صفحة التقييم
  ratingsSheet_().appendRow([
    now, data.deptId, data.deptName || "", data.service || "", data.raterType || "",
    Number(data.quality)||0, Number(data.speed)||0, Number(data.staff)||0,
    Number(data.clarity)||0, Number(data.access)||0,
    data.comment || "", data.complaintType || "", data.complaintText || ""
  ]);
  return jsonOut_({ ok:true });
}
