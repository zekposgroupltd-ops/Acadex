/* ============================================================
   ACADEX MULTI-SCHOOL — Google Apps Script Backend v4
   © 2026 ACADEX — Designed by Zekpo Destiny
   NEW: Classes management, staff subject assignments,
        duplicate result deduplication, RID delete,
        staff sees only assigned subjects
   ============================================================ */
const SPREADSHEET_ID = SpreadsheetApp.getActive().getId();
const REQUIRED_SHEETS = {
  RESULTS:       ["school_code","student_id","subject_code","first_test","second_test","exam","total","term","session","submitted_by","date_recorded"],
  STUDENTS:      ["school_code","student_id","full_name","class","arm","exam_pin","gender","date_of_birth","parent_phone","photo_url","status","date_added"],
  SCHOOLS:       ["school_code","school_name","address","town","state","admin_username","admin_password","staff_password","super_admin","contact_email","contact_phone","logo_url","status","date_added"],
  RESULT_IDS:    ["school_code","result_id","student_id","attempts_used","max_attempts","term","session","status","date_created"],
  SUBJECTS:      ["school_code","subject_code","subject_name","class","arm","duration_minutes","term","session","status","cbt_enabled"],
  SCHOOL_SETTINGS:["school_code","school_name","address","town","state","principal_name","term","session","first_test_max","second_test_max","exam_max","result_checking","exam_open","motto","logo_url","date_updated","result_display","position_scope"],
  ACTIVITY_LOG:  ["school_code","user_id","user_role","action","details","notes","timestamp"],
  ATTENDANCE:    ["school_code","student_id","class","date","status","recorded_by","term","session"],
  ANNOUNCEMENTS: ["school_code","title","content","priority","target_role","target_class","start_date","end_date","created_by","created_at"],
  FEES:          ["school_code","class","fee_type","amount","description","term","session","due_date","status","priority"],
  FEE_PAYMENTS:  ["school_code","student_id","fee_id","amount_paid","payment_date","payment_method","recorded_by","term","session","balance_after"],
  LIBRARY:       ["school_code","book_id","title","author","isbn","category","total_copies","available","status","date_added"],
  LIBRARY_LOANS: ["school_code","loan_id","book_id","student_id","borrow_date","due_date","return_date","status","recorded_by"],
  TIMETABLE:     ["school_code","class","arm","day","period","subject_code","teacher_id","room","term","session"],
  EVENTS:        ["school_code","event_id","title","description","event_type","start_date","end_date","all_day","location","target_class","target_role","created_by","created_at"],
  CLASSES:       ["school_code","class_name","class_level","arm","capacity","status","date_added"],
  STAFF:         ["school_code","staff_id","full_name","email","phone","subject_code","class","role","password","status","date_added"],
  QUESTIONS:     ["school_code","question_id","subject_code","question","option_a","option_b","option_c","option_d","correct","explanation","date_added"],
  STAFF_SUBJECTS:["school_code","staff_id","subject_code","class","date_assigned"],
  WEBSITE_CONFIG:["school_code","config_json","published","published_at","updated_at"]
};
function doGet(e) {
  try {
    ensureRequiredSheets();
    const action = e.parameter.action;
    if (!action) return output({ success: false, message: "No action specified" });
    if (action === "studentLogin")          return studentLogin(e.parameter);
    if (action === "parentLogin")           return parentLogin(e.parameter);
    if (action === "adminLogin")            return adminLogin(e.parameter);
    if (action === "staffLogin")            return staffLogin(e.parameter);
    if (action === "superAdminLogin")       return superAdminLogin(e.parameter);
    if (action === "getSettings")           return getSettings(e.parameter);
    if (action === "updateSettings")        return updateSettings(e.parameter);
    if (action === "getStudents")           return getStudents(e.parameter);
    if (action === "getStudentsWithPins")   return getStudentsWithPins(e.parameter);
    if (action === "addStudent")            return addStudent(e.parameter);
    if (action === "updateStudent")         return updateStudent(e.parameter);
    if (action === "deleteStudent")         return deleteStudent(e.parameter);
    if (action === "generateStudentId")     return generateStudentId(e.parameter);
    if (action === "getStaff")              return getStaff(e.parameter);
    if (action === "addStaff")              return addStaff(e.parameter);
    if (action === "updateStaff")           return updateStaff(e.parameter);
    if (action === "deleteStaff")           return deleteStaff(e.parameter);
    if (action === "getClasses")            return getClasses(e.parameter);
    if (action === "addClass")              return addClass(e.parameter);
    if (action === "updateClass")           return updateClass(e.parameter);
    if (action === "deleteClass")           return deleteClass(e.parameter);
    if (action === "getSubjects")           return getSubjects(e.parameter);
    if (action === "addSubject")            return addSubject(e.parameter);
    if (action === "updateSubject")         return updateSubject(e.parameter);
    if (action === "deleteSubject")         return deleteSubject(e.parameter);
    if (action === "getStaffSubjects")      return getStaffSubjects(e.parameter);
    if (action === "assignStaffSubject")    return assignStaffSubject(e.parameter);
    if (action === "removeStaffSubject")    return removeStaffSubject(e.parameter);
    if (action === "getQuestions")          return getQuestions(e.parameter);
    if (action === "addQuestion")           return addQuestion(e.parameter);
    if (action === "deleteQuestion")        return deleteQuestion(e.parameter);
    if (action === "saveResult")            return saveResult(e.parameter);
    if (action === "getResults")            return getResults(e.parameter);
    if (action === "saveManualResult")      return saveManualResult(e.parameter);
    if (action === "getClassResults")       return getClassResults(e.parameter);
    if (action === "getClassPositions")     return getClassPositions(e.parameter);
    if (action === "generateResultId")      return generateResultId(e.parameter);
    if (action === "validateResultId")      return validateResultId(e.parameter);
    if (action === "getResultIds")          return getResultIds(e.parameter);
    if (action === "deleteResultId")        return deleteResultId(e.parameter);
    if (action === "deleteBulkResultIds")   return deleteBulkResultIds(e.parameter);
    if (action === "getSchools")            return getSchools(e.parameter);
    if (action === "addSchool")             return addSchool(e.parameter);
    if (action === "updateSchool")          return updateSchool(e.parameter);
    if (action === "getDashboard")          return getDashboard(e.parameter);
    if (action === "deduplicateResults")    return deduplicateResults(e.parameter);
    if (action === "getAttendance")         return getAttendance(e.parameter);
    if (action === "markAttendance")        return markAttendance(e.parameter);
    if (action === "getAnnouncements")      return getAnnouncements(e.parameter);
    if (action === "addAnnouncement")       return addAnnouncement(e.parameter);
    if (action === "deleteAnnouncement")    return deleteAnnouncement(e.parameter);
    if (action === "getFees")               return getFees(e.parameter);
    if (action === "addFee")                return addFee(e.parameter);
    if (action === "updateFee")             return updateFee(e.parameter);
    if (action === "deleteFee")             return deleteFee(e.parameter);
    if (action === "recordPayment")         return recordPayment(e.parameter);
    if (action === "getFeePayments")        return getFeePayments(e.parameter);
    if (action === "getParentResults")      return getParentResults(e.parameter);
    if (action === "getParentAttendance")   return getParentAttendance(e.parameter);
    if (action === "getParentFees")         return getParentFees(e.parameter);
    if (action === "getLibraryBooks")       return getLibraryBooks(e.parameter);
    if (action === "addLibraryBook")        return addLibraryBook(e.parameter);
    if (action === "getLibraryLoans")       return getLibraryLoans(e.parameter);
    if (action === "borrowBook")            return borrowBook(e.parameter);
    if (action === "returnBook")            return returnBook(e.parameter);
    if (action === "getTimetable")          return getTimetable(e.parameter);
    if (action === "addTimetable")          return addTimetable(e.parameter);
    if (action === "getEvents")             return getEvents(e.parameter);
    if (action === "addEvent")              return addEvent(e.parameter);
    if (action === "deleteEvent")           return deleteEvent(e.parameter);
    if (action === "getStudentProfile")     return getStudentProfile(e.parameter);
    if (action === "updateStudentProfile")  return updateStudentProfile(e.parameter);
    if (action === "getStaffProfile")       return getStaffProfile(e.parameter);
    if (action === "updateStaffProfile")    return updateStaffProfile(e.parameter);
    if (action === "bulkImportStudents")    return bulkImportStudents(e.parameter);
    if (action === "bulkImportStaff")       return bulkImportStaff(e.parameter);
    if (action === "getFeeBalance")         return getFeeBalance(e.parameter);
    if (action === "recordPaymentAllocated")return recordPaymentAllocated(e.parameter);
    if (action === "getStudentLedger")      return getStudentLedger(e.parameter);
    if (action === "applyFeeWaiver")        return applyFeeWaiver(e.parameter);
    if (action === "rolloverTerm")          return rolloverTerm(e.parameter);
    if (action === "getFinanceDashboard")   return getFinanceDashboard(e.parameter);
    if (action === "getClassArms")          return getClassArms(e.parameter);
    if (action === "addClassArm")           return addClassArm(e.parameter);
    if (action === "deleteClassArm")        return deleteClassArm(e.parameter);
    if (action === "updateAnnouncement")    return updateAnnouncement(e.parameter);
    if (action === "saveWebsiteConfig")     return saveWebsiteConfig(e.parameter);
    if (action === "getWebsiteConfig")      return getWebsiteConfig(e.parameter);
    if (action === "getStudentFeeReport")   return getStudentFeeReport(e.parameter);
    if (action === "deleteFeePayment")      return deleteFeePayment(e.parameter);
    return output({ success: false, message: "Invalid action: " + action });
  } catch (err) {
    logError("doGet", err.message);
    return output({ success: false, message: "Server error: " + err.message });
  }
}
function doPost(e) { return doGet(e); }
// ============================================================
// CORE HELPERS
// ============================================================
function output(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function getSheet(name) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  if (!sheet) throw new Error("Sheet not found: " + name);
  return sheet;
}
function ensureRequiredSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.keys(REQUIRED_SHEETS).forEach(name => {
    if (!ss.getSheetByName(name)) {
      const sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, REQUIRED_SHEETS[name].length).setValues([REQUIRED_SHEETS[name]]);
    }
  });
}
function sheetToObjects(sheetName) {
  const sheet = getSheet(sheetName);
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim().toLowerCase().replace(/ /g, "_"));
  return data.slice(1)
    .filter(r => r.some(cell => cell !== "" && cell !== null && cell !== undefined))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = (r[i] !== undefined && r[i] !== null) ? r[i] : ""; });
      return obj;
    });
}
function getHeaders(sheetName) {
  const sheet = getSheet(sheetName);
  const h     = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return h.map(h => String(h).trim().toLowerCase().replace(/ /g, "_"));
}
function logActivity(schoolCode, userId, userRole, action, details) {
  try {
    getSheet("ACTIVITY_LOG").appendRow([schoolCode||"",userId||"",userRole||"",action||"",details||"","",new Date().toISOString()]);
  } catch (e) {}
}
function logError(location, message) { try { logActivity("SYSTEM","system","error",location,message); } catch(e){} }
function generatePin()     { return String(Math.floor(1000 + Math.random() * 9000)); }
function randomSuffix(len) { return Math.random().toString(36).substring(2,2+len).toUpperCase(); }
function currentYear()     { return new Date().getFullYear(); }
function today()           { return new Date().toISOString().split("T")[0]; }
function matchCode(a,b)    { return String(a).trim().toUpperCase() === String(b).trim().toUpperCase(); }
function getGrade(score) {
  const s = parseFloat(score) || 0;
  if (s >= 75) return { grade:"A1", remark:"Excellent"  };
  if (s >= 70) return { grade:"B2", remark:"Very Good"  };
  if (s >= 65) return { grade:"B3", remark:"Good"       };
  if (s >= 60) return { grade:"C4", remark:"Credit"     };
  if (s >= 55) return { grade:"C5", remark:"Credit"     };
  if (s >= 50) return { grade:"C6", remark:"Credit"     };
  if (s >= 45) return { grade:"D7", remark:"Pass"       };
  if (s >= 40) return { grade:"E8", remark:"Pass"       };
  return              { grade:"F9", remark:"Fail"       };
}
function scaleExamScore(correct, total, examMax) {
  if (!total || parseInt(total) === 0) return 0;
  return Math.round((parseInt(correct) / parseInt(total)) * (parseFloat(examMax) || 60) * 10) / 10;
}
function calcTotal(t1, t2, ex) {
  return (parseFloat(t1)||0) + (parseFloat(t2)||0) + (parseFloat(ex)||0);
}
function getSchoolSettings(schoolCode) {
  const s = sheetToObjects("SCHOOL_SETTINGS").find(s => matchCode(s.school_code, schoolCode));
  if (!s) throw new Error("School settings not found for: " + schoolCode);
  return s;
}
function getSchoolInfo(schoolCode) {
  return sheetToObjects("SCHOOLS").find(s => matchCode(s.school_code, schoolCode)) || {};
}
// Ensure CLASSES sheet exists
function ensureClassesSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("CLASSES");
  if (!sheet) {
    sheet = ss.insertSheet("CLASSES");
    sheet.getRange(1,1,1,7).setValues([["school_code","class_name","class_level","arm","capacity","status","date_added"]]);
  }
  return sheet;
}
// Ensure STAFF_SUBJECTS sheet exists
function ensureStaffSubjectsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("STAFF_SUBJECTS");
  if (!sheet) {
    sheet = ss.insertSheet("STAFF_SUBJECTS");
    sheet.getRange(1,1,1,5).setValues([["school_code","staff_id","subject_code","class","date_assigned"]]);
  }
  return sheet;
}
// ============================================================
// AUTH — STUDENT LOGIN
// ============================================================
function studentLogin(params) {
  try {
    const studentId = String(params.student_id||"").trim();
    const pin       = String(params.exam_pin  ||"").trim();
    if (!studentId||!pin) return output({success:false,message:"Student ID and PIN are required"});
    const student = sheetToObjects("STUDENTS").find(s=>
      String(s.student_id).trim()===studentId &&
      String(s.exam_pin).trim()  ===pin       &&
      String(s.status).trim().toLowerCase()==="active"
    );
    if (!student) return output({success:false,message:"Invalid Student ID or PIN, or account inactive"});
    let settings = {};
    try { settings = getSchoolSettings(student.school_code); } catch(e) {}
    if (String(settings.exam_open||"yes").trim().toLowerCase()==="no")
      return output({success:false,message:"Exams are closed. Contact your admin."});
    const safe = Object.assign({},student); delete safe.exam_pin;
    logActivity(student.school_code,studentId,"student","login","Student logged in");
    return output({success:true,student:safe,settings:{
      school_name:    settings.school_name    ||"",
      term:           settings.term           ||"",
      session:        settings.session        ||"",
      exam_max:       settings.exam_max       ||60,
      first_test_max: settings.first_test_max ||20,
      second_test_max:settings.second_test_max||20,
      result_display: settings.result_display ||"grade_only",
      result_checking:settings.result_checking||"yes"
    }});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// AUTH — PARENT LOGIN
// ============================================================
function parentLogin(params) {
  try {
    const phoneRaw = String(params.parent_phone||"").trim();
    const studentId = String(params.student_id||"").trim();
    if (!phoneRaw) return output({success:false,message:"Parent phone is required"});
    // Normalize phone: keep digits only, also try last-10-digit match (handles +234 / 0 prefixes)
    const norm = function(p){ return String(p||"").replace(/\D/g,""); };
    const phoneN = norm(phoneRaw);
    const phoneTail = phoneN.slice(-10);
    const all = sheetToObjects("STUDENTS").filter(function(s){
      if (String(s.status||"").trim().toLowerCase() !== "active") return false;
      const sp = norm(s.parent_phone);
      if (!sp) return false;
      if (sp === phoneN) return true;
      if (phoneTail && sp.slice(-10) === phoneTail) return true;
      return false;
    });
    // Optional: if a student_id was supplied, restrict to that one
    const matches = studentId ? all.filter(function(s){ return String(s.student_id).trim() === studentId; }) : all;
    if (!matches.length) return output({success:false,message:"No active student found for this phone number"});
    const primary = matches[0];
    let settings = {};
    try { settings = getSchoolSettings(primary.school_code); } catch(e) {}
    logActivity(primary.school_code, primary.student_id, "parent", "login", "Parent logged in (" + matches.length + " child" + (matches.length>1?"ren":"") + ")");
    const children = matches.map(function(s){
      return {
        student_id: s.student_id,
        student_name: s.full_name,
        student_class: s.class,
        arm: s.arm || "",
        photo_url: s.photo_url || "",
        school_code: s.school_code
      };
    });
    return output({
      success: true,
      parent: {
        student_id: primary.student_id,
        student_name: primary.full_name,
        student_class: primary.class,
        parent_phone: primary.parent_phone,
        school_code: primary.school_code,
        children: children
      },
      children: children,
      settings: {
        school_name: settings.school_name || "",
        term: settings.term || "",
        session: settings.session || ""
      }
    });
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// PARENT DATA FUNCTIONS
// ============================================================
function getParentResults(params) {
  try {
    const studentId = String(params.student_id||"").trim();
    const schoolCode = String(params.school_code||"").trim();
    const term = String(params.term||"").trim();
    const session = String(params.session||"").trim();
    if (!studentId || !schoolCode) return output({success:false,message:"Student ID and school code required"});
    const results = sheetToObjects("RESULTS").filter(r =>
      matchCode(r.school_code, schoolCode) &&
      String(r.student_id).trim() === studentId &&
      (!term || String(r.term).trim() === term) &&
      (!session || String(r.session).trim() === session)
    );
    const subjects = sheetToObjects("SUBJECTS");
    const resultsWithNames = results.map(r => {
      const subj = subjects.find(s => matchCode(s.school_code, schoolCode) && String(s.subject_code).trim() === String(r.subject_code).trim());
      return { ...r, subject_name: subj ? subj.subject_name : r.subject_code };
    });
    return output({success:true, results: resultsWithNames});
  } catch(err) { return output({success:false,message:err.message}); }
}
function getParentAttendance(params) {
  try {
    const studentId = String(params.student_id||"").trim();
    const schoolCode = String(params.school_code||"").trim();
    const term = String(params.term||"").trim();
    const session = String(params.session||"").trim();
    if (!studentId || !schoolCode) return output({success:false,message:"Student ID and school code required"});
    const attendance = sheetToObjects("ATTENDANCE").filter(a =>
      matchCode(a.school_code, schoolCode) &&
      String(a.student_id).trim() === studentId &&
      (!term || String(a.term).trim() === term) &&
      (!session || String(a.session).trim() === session)
    );
    return output({success:true, attendance: attendance});
  } catch(err) { return output({success:false,message:err.message}); }
}
function getParentFees(params) {
  try {
    const studentId = String(params.student_id||"").trim();
    const schoolCode = String(params.school_code||"").trim();
    const studentClass = String(params.student_class||"").trim();
    if (!studentId || !schoolCode) return output({success:false,message:"Student ID and school code required"});
    const fees = sheetToObjects("FEES").filter(f =>
      matchCode(f.school_code, schoolCode) &&
      (!studentClass || String(f.class).trim() === studentClass)
    );
    const payments = sheetToObjects("FEE_PAYMENTS").filter(p =>
      matchCode(p.school_code, schoolCode) &&
      String(p.student_id).trim() === studentId
    );
    const totalFees = fees.reduce((sum, f) => sum + (parseFloat(f.amount)||0), 0);
    const totalPaid = payments.reduce((sum, p) => sum + (parseFloat(p.amount_paid)||0), 0);
    return output({success:true, fees: fees, payments: payments, totalFees: totalFees, totalPaid: totalPaid, balance: totalFees - totalPaid});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// AUTH — ADMIN LOGIN
// ============================================================
function adminLogin(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const username   = String(params.username   ||"").trim();
    const password   = String(params.password   ||"").trim();
    if (!schoolCode||!username||!password)
      return output({success:false,message:"School code, username and password required"});
    const school = sheetToObjects("SCHOOLS").find(s=>
      matchCode(s.school_code,schoolCode) &&
      String(s.admin_username).trim()===username &&
      String(s.admin_password).trim()===password &&
      String(s.status).trim().toLowerCase()==="active"
    );
    if (!school) return output({success:false,message:"Invalid credentials or inactive school"});
    let settings = {};
    try { settings = getSchoolSettings(schoolCode); } catch(e) {}
    logActivity(schoolCode,username,"admin","login","Admin logged in");
    return output({success:true,role:"admin",school_code:schoolCode,
      school_name: school.school_name||"",
      logo_url:    school.logo_url   ||"",
      address:     school.address    ||"",
      town:        school.town       ||"",
      state:       school.state      ||"",
      settings:    settings
    });
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// AUTH — STAFF LOGIN (returns assigned subjects)
// ============================================================
function staffLogin(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const staffId    = String(params.staff_id   ||"").trim();
    const password   = String(params.password   ||"").trim();
    if (!schoolCode||!staffId||!password)
      return output({success:false,message:"School code, Staff ID and password required"});
    const staff = sheetToObjects("STAFF").find(s=>
      matchCode(s.school_code,schoolCode) &&
      String(s.staff_id).trim()===staffId &&
      String(s.password).trim()===password&&
      String(s.status).trim().toLowerCase()==="active"
    );
    if (!staff) return output({success:false,message:"Invalid Staff ID or password"});
    let settings = {};
    try { settings = getSchoolSettings(schoolCode); } catch(e) {}
    // Get subjects assigned to this staff member
    let assignedSubjects = [];
    try {
      ensureStaffSubjectsSheet();
      const assignments = sheetToObjects("STAFF_SUBJECTS").filter(a=>
        matchCode(a.school_code,schoolCode) && String(a.staff_id).trim()===staffId
      );
      const allSubjects = sheetToObjects("SUBJECTS").filter(s=>
        matchCode(s.school_code,schoolCode) && String(s.status).trim().toLowerCase()==="active"
      );
      assignedSubjects = assignments.map(a=>{
        const subj = allSubjects.find(s=>
          String(s.subject_code).trim().toUpperCase()===String(a.subject_code).trim().toUpperCase() &&
          String(s.class).trim().toLowerCase()===String(a.class).trim().toLowerCase()
        );
        return subj ? subj : null;
      }).filter(Boolean);
      // Fallback: if no assignments, use subject_code from staff record (legacy)
      if (assignedSubjects.length===0 && String(staff.subject_code||"").trim()) {
        assignedSubjects = allSubjects.filter(s=>
          String(s.subject_code).trim().toUpperCase()===String(staff.subject_code).trim().toUpperCase()
        );
      }
    } catch(e) {}
    const safe = Object.assign({},staff); delete safe.password;
    logActivity(schoolCode,staffId,"teacher","login","Staff logged in");
    return output({success:true,role:"teacher",staff:safe,
      school_name:     settings.school_name||"",
      settings:        settings,
      assignedSubjects:assignedSubjects
    });
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// AUTH — SUPER ADMIN LOGIN
// ============================================================
function superAdminLogin(params) {
  try {
    const username = String(params.username||"").trim();
    const password = String(params.password||"").trim();
    if (!username||!password) return output({success:false,message:"Username and password required"});
    const sa = sheetToObjects("SCHOOLS").find(s=>
      String(s.super_admin).trim().toLowerCase()==="yes" &&
      String(s.admin_username).trim()===username &&
      String(s.admin_password).trim()===password
    );
    if (!sa) return output({success:false,message:"Invalid super admin credentials"});
    logActivity("SYSTEM",username,"super_admin","login","Super admin logged in");
    return output({success:true,role:"super_admin",name:sa.school_name||"Super Admin"});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// SETTINGS
// ============================================================
function getSettings(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    const settings = getSchoolSettings(schoolCode);
    const info = getSchoolInfo(schoolCode);
    return output({success:true,settings:Object.assign({},settings,{logo_url:info.logo_url||settings.logo_url||""})});
  } catch(err) { return output({success:false,message:err.message}); }
}
function updateSettings(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    const sheet = getSheet("SCHOOL_SETTINGS"), headers = getHeaders("SCHOOL_SETTINGS");
    const data  = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i=1;i<data.length;i++) { if (matchCode(data[i][0],schoolCode)){rowIndex=i+1;break;} }
    const updatable = ["school_name","address","town","state","principal_name","term","session",
      "first_test_max","second_test_max","exam_max","result_checking","exam_open","motto",
      "result_display","position_scope"];
    if (rowIndex===-1) {
      const newRow = headers.map(h=>{
        if (h==="school_code") return schoolCode;
        if (h==="date_updated") return today();
        return params[h]!==undefined?params[h]:"";
      });
      sheet.appendRow(newRow);
    } else {
      updatable.forEach(f=>{ if(params[f]!==undefined){const c=headers.indexOf(f);if(c!==-1)sheet.getRange(rowIndex,c+1).setValue(params[f]);} });
      const dc=headers.indexOf("date_updated"); if(dc!==-1)sheet.getRange(rowIndex,dc+1).setValue(today());
    }
    if (params.logo_url!==undefined) {
      try {
        const ss=getSheet("SCHOOLS"),sh=getHeaders("SCHOOLS"),sd=ss.getDataRange().getValues();
        for (let i=1;i<sd.length;i++) {
          if (matchCode(sd[i][0],schoolCode)){const c=sh.indexOf("logo_url");if(c!==-1)ss.getRange(i+1,c+1).setValue(params.logo_url);break;}
        }
      } catch(e){}
    }
    logActivity(schoolCode,"admin","admin","updateSettings","Settings updated");
    return output({success:true,message:"Settings updated"});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// CLASSES MANAGEMENT
// ============================================================
function getClasses(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    ensureClassesSheet();
    let classes = sheetToObjects("CLASSES").filter(c=>matchCode(c.school_code,schoolCode));
    // Also derive classes from students if CLASSES sheet is empty
    if (classes.length===0) {
      const students = sheetToObjects("STUDENTS").filter(s=>matchCode(s.school_code,schoolCode));
      const unique = [...new Set(students.map(s=>String(s.class).trim()).filter(Boolean))];
      classes = unique.map(name=>({school_code:schoolCode,class_name:name,class_level:"",arm:"",capacity:"",status:"active",date_added:""}));
    }
    return output({success:true,classes:classes,count:classes.length});
  } catch(err) { return output({success:false,message:err.message}); }
}
function addClass(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const className  = String(params.class_name ||"").trim();
    if (!schoolCode||!className) return output({success:false,message:"school_code and class_name required"});
    ensureClassesSheet();
    const existing = sheetToObjects("CLASSES").filter(c=>matchCode(c.school_code,schoolCode));
    if (existing.find(c=>String(c.class_name).trim().toLowerCase()===className.toLowerCase()))
      return output({success:false,message:"Class already exists: "+className});
    getSheet("CLASSES").appendRow([schoolCode,className,params.class_level||"",params.arm||"",params.capacity||"","active",today()]);
    return output({success:true,message:"Class added: "+className,class_name:className});
  } catch(err) { return output({success:false,message:err.message}); }
}
function updateClass(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const className  = String(params.class_name ||"").trim();
    if (!schoolCode||!className) return output({success:false,message:"school_code and class_name required"});
    ensureClassesSheet();
    const sheet=getSheet("CLASSES"),headers=getHeaders("CLASSES"),data=sheet.getDataRange().getValues();
    let rowIndex=-1;
    for (let i=1;i<data.length;i++){if(matchCode(data[i][0],schoolCode)&&String(data[i][1]).trim().toLowerCase()===className.toLowerCase()){rowIndex=i+1;break;}}
    if (rowIndex===-1) return output({success:false,message:"Class not found"});
    ["class_level","arm","capacity","status"].forEach(f=>{if(params[f]!==undefined){const c=headers.indexOf(f);if(c!==-1)sheet.getRange(rowIndex,c+1).setValue(params[f]);}});
    return output({success:true,message:"Class updated"});
  } catch(err) { return output({success:false,message:err.message}); }
}
function deleteClass(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const className  = String(params.class_name ||"").trim();
    if (!schoolCode||!className) return output({success:false,message:"school_code and class_name required"});
    ensureClassesSheet();
    const sheet=getSheet("CLASSES"),data=sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++){
      if(matchCode(data[i][0],schoolCode)&&String(data[i][1]).trim().toLowerCase()===className.toLowerCase()){
        sheet.deleteRow(i+1); return output({success:true,message:"Class deleted"});
      }
    }
    return output({success:false,message:"Class not found"});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// STUDENT ID GENERATION — COLLISION-SAFE WITH LOCK
// ============================================================
function generateStudentId(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const props = PropertiesService.getScriptProperties();
      const key   = "STUDENT_COUNTER_"+schoolCode+"_"+currentYear();
      let counter = parseInt(props.getProperty(key)||"0");
      if (counter===0) {
        counter = sheetToObjects("STUDENTS").filter(s=>matchCode(s.school_code,schoolCode)).length;
      }
      counter++;
      props.setProperty(key,String(counter));
      const studentId = `${schoolCode}-${currentYear()}-${String(counter).padStart(4,"0")}`;
      return output({success:true,student_id:studentId,exam_pin:generatePin()});
    } finally { lock.releaseLock(); }
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// STUDENTS
// ============================================================
function getStudents(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const className  = String(params.class      ||"").trim();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    let students = sheetToObjects("STUDENTS").filter(s=>matchCode(s.school_code,schoolCode));
    if (className) students = students.filter(s=>String(s.class).trim().toLowerCase()===className.toLowerCase());
    students = students.map(s=>{const c=Object.assign({},s);delete c.exam_pin;return c;});
    return output({success:true,students:students,count:students.length});
  } catch(err) { return output({success:false,message:err.message}); }
}
function getStudentsWithPins(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const className  = String(params.class      ||"").trim();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    let students = sheetToObjects("STUDENTS").filter(s=>matchCode(s.school_code,schoolCode));
    if (className) students = students.filter(s=>String(s.class).trim().toLowerCase()===className.toLowerCase());
    return output({success:true,students:students,count:students.length});
  } catch(err) { return output({success:false,message:err.message}); }
}
function addStudent(params) {
  try {
    const schoolCode = String(params.school_code  ||"").trim().toUpperCase();
    const studentId  = String(params.student_id   ||"").trim();
    const fullName   = String(params.full_name    ||"").trim();
    const className  = String(params.class        ||"").trim();
    const pin        = String(params.exam_pin     ||generatePin()).trim();
    const gender     = String(params.gender       ||"").trim();
    const dob        = String(params.date_of_birth||"").trim();
    const phone      = String(params.parent_phone ||"").trim();
    const status     = String(params.status       ||"active").trim();
    if (!schoolCode||!studentId||!fullName||!className)
      return output({success:false,message:"school_code, student_id, full_name and class are required"});
    const existing = sheetToObjects("STUDENTS");
    if (existing.find(s=>String(s.student_id).trim()===studentId))
      return output({success:false,message:"Student ID already exists: "+studentId});
    getSheet("STUDENTS").appendRow([schoolCode,studentId,fullName,className,pin,gender,dob,phone,status,today()]);
    logActivity(schoolCode,"admin","admin","addStudent","Added: "+studentId);
    return output({success:true,message:"Student added",student_id:studentId,exam_pin:pin});
  } catch(err) { return output({success:false,message:err.message}); }
}
function updateStudent(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const studentId  = String(params.student_id ||"").trim();
    if (!schoolCode||!studentId) return output({success:false,message:"school_code and student_id required"});
    const sheet=getSheet("STUDENTS"),headers=getHeaders("STUDENTS"),data=sheet.getDataRange().getValues();
    let rowIndex=-1;
    for (let i=1;i<data.length;i++){if(matchCode(data[i][0],schoolCode)&&String(data[i][1]).trim()===studentId){rowIndex=i+1;break;}}
    if (rowIndex===-1) return output({success:false,message:"Student not found"});
    ["full_name","class","exam_pin","gender","date_of_birth","parent_phone","status"].forEach(f=>{
      if(params[f]!==undefined){const c=headers.indexOf(f);if(c!==-1)sheet.getRange(rowIndex,c+1).setValue(params[f]);}
    });
    logActivity(schoolCode,"admin","admin","updateStudent","Updated: "+studentId);
    return output({success:true,message:"Student updated"});
  } catch(err) { return output({success:false,message:err.message}); }
}
function deleteStudent(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const studentId  = String(params.student_id ||"").trim();
    if (!schoolCode||!studentId) return output({success:false,message:"school_code and student_id required"});
    const sheet=getSheet("STUDENTS"),data=sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++){
      if(matchCode(data[i][0],schoolCode)&&String(data[i][1]).trim()===studentId){
        sheet.deleteRow(i+1);
        logActivity(schoolCode,"admin","admin","deleteStudent","Deleted: "+studentId);
        return output({success:true,message:"Student deleted"});
      }
    }
    return output({success:false,message:"Student not found"});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// STAFF
// ============================================================
function getStaff(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    let staff = sheetToObjects("STAFF").filter(s=>matchCode(s.school_code,schoolCode));
    // Enrich each staff with their assigned subjects
    try {
      ensureStaffSubjectsSheet();
      const assignments = sheetToObjects("STAFF_SUBJECTS").filter(a=>matchCode(a.school_code,schoolCode));
      staff = staff.map(s=>{
        const c=Object.assign({},s); delete c.password;
        c.assigned_subjects = assignments.filter(a=>String(a.staff_id).trim()===String(s.staff_id).trim())
          .map(a=>a.subject_code+"@"+a.class);
        return c;
      });
    } catch(e) {
      staff = staff.map(s=>{const c=Object.assign({},s);delete c.password;return c;});
    }
    return output({success:true,staff:staff,count:staff.length});
  } catch(err) { return output({success:false,message:err.message}); }
}
function addStaff(params) {
  try {
    const schoolCode  = String(params.school_code ||"").trim().toUpperCase();
    const fullName    = String(params.full_name   ||"").trim();
    const email       = String(params.email       ||"").trim();
    const phone       = String(params.phone       ||"").trim();
    const subjectCode = String(params.subject_code||"").trim();
    const className   = String(params.class       ||"").trim();
    const role        = String(params.role        ||"teacher").trim();
    const password    = String(params.password    ||generatePin()).trim();
    const status      = String(params.status      ||"active").trim();
    if (!schoolCode||!fullName) return output({success:false,message:"school_code and full_name required"});
    const existing = sheetToObjects("STAFF").filter(s=>matchCode(s.school_code,schoolCode));
    const staffId  = `${schoolCode}-TCH-${String(existing.length+1).padStart(3,"0")}`;
    getSheet("STAFF").appendRow([schoolCode,staffId,fullName,email,phone,subjectCode,className,role,password,status,today()]);
    // Auto-assign subject if provided
    if (subjectCode && className) {
      try {
        ensureStaffSubjectsSheet();
        const fullCode = subjectCode.startsWith(schoolCode+"-")?subjectCode:`${schoolCode}-${subjectCode}`;
        getSheet("STAFF_SUBJECTS").appendRow([schoolCode,staffId,fullCode,className,today()]);
      } catch(e){}
    }
    logActivity(schoolCode,"admin","admin","addStaff","Added: "+staffId);
    return output({success:true,message:"Staff added",staff_id:staffId,password:password});
  } catch(err) { return output({success:false,message:err.message}); }
}
function updateStaff(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const staffId    = String(params.staff_id   ||"").trim();
    if (!schoolCode||!staffId) return output({success:false,message:"school_code and staff_id required"});
    const sheet=getSheet("STAFF"),headers=getHeaders("STAFF"),data=sheet.getDataRange().getValues();
    let rowIndex=-1;
    for (let i=1;i<data.length;i++){if(matchCode(data[i][0],schoolCode)&&String(data[i][1]).trim()===staffId){rowIndex=i+1;break;}}
    if (rowIndex===-1) return output({success:false,message:"Staff not found"});
    ["full_name","email","phone","subject_code","class","role","password","status"].forEach(f=>{
      if(params[f]!==undefined){const c=headers.indexOf(f);if(c!==-1)sheet.getRange(rowIndex,c+1).setValue(params[f]);}
    });
    return output({success:true,message:"Staff updated"});
  } catch(err) { return output({success:false,message:err.message}); }
}
function deleteStaff(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const staffId    = String(params.staff_id   ||"").trim();
    if (!schoolCode||!staffId) return output({success:false,message:"school_code and staff_id required"});
    const sheet=getSheet("STAFF"),data=sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++){
      if(matchCode(data[i][0],schoolCode)&&String(data[i][1]).trim()===staffId){
        sheet.deleteRow(i+1); return output({success:true,message:"Staff deleted"});
      }
    }
    return output({success:false,message:"Staff not found"});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// STAFF SUBJECT ASSIGNMENTS
// ============================================================
function getStaffSubjects(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const staffId    = String(params.staff_id   ||"").trim();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    ensureStaffSubjectsSheet();
    let assignments = sheetToObjects("STAFF_SUBJECTS").filter(a=>matchCode(a.school_code,schoolCode));
    if (staffId) assignments = assignments.filter(a=>String(a.staff_id).trim()===staffId);
    return output({success:true,assignments:assignments,count:assignments.length});
  } catch(err) { return output({success:false,message:err.message}); }
}
function assignStaffSubject(params) {
  try {
    const schoolCode  = String(params.school_code ||"").trim().toUpperCase();
    const staffId     = String(params.staff_id    ||"").trim();
    const subjectCode = String(params.subject_code||"").trim().toUpperCase();
    const className   = String(params.class       ||"").trim();
    if (!schoolCode||!staffId||!subjectCode||!className)
      return output({success:false,message:"school_code, staff_id, subject_code and class required"});
    ensureStaffSubjectsSheet();
    const existing = sheetToObjects("STAFF_SUBJECTS");
    if (existing.find(a=>matchCode(a.school_code,schoolCode)&&String(a.staff_id).trim()===staffId&&String(a.subject_code).trim().toUpperCase()===subjectCode&&String(a.class).trim().toLowerCase()===className.toLowerCase()))
      return output({success:false,message:"Assignment already exists"});
    getSheet("STAFF_SUBJECTS").appendRow([schoolCode,staffId,subjectCode,className,today()]);
    return output({success:true,message:"Subject assigned to staff"});
  } catch(err) { return output({success:false,message:err.message}); }
}
function removeStaffSubject(params) {
  try {
    const schoolCode  = String(params.school_code ||"").trim().toUpperCase();
    const staffId     = String(params.staff_id    ||"").trim();
    const subjectCode = String(params.subject_code||"").trim().toUpperCase();
    const className   = String(params.class       ||"").trim();
    if (!schoolCode||!staffId||!subjectCode) return output({success:false,message:"school_code, staff_id and subject_code required"});
    ensureStaffSubjectsSheet();
    const sheet=getSheet("STAFF_SUBJECTS"),data=sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++){
      if(matchCode(data[i][0],schoolCode)&&String(data[i][1]).trim()===staffId&&String(data[i][2]).trim().toUpperCase()===subjectCode&&(!className||String(data[i][3]).trim().toLowerCase()===className.toLowerCase())){
        sheet.deleteRow(i+1); return output({success:true,message:"Assignment removed"});
      }
    }
    return output({success:false,message:"Assignment not found"});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// SUBJECTS
// ============================================================
function getSubjects(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const className  = String(params.class      ||"").trim();
    const term       = String(params.term       ||"").trim();
    const staffId    = String(params.staff_id   ||"").trim();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    let subjects = sheetToObjects("SUBJECTS").filter(s=>
      matchCode(s.school_code,schoolCode)&&String(s.status).trim().toLowerCase()==="active"
    );
    if (className) subjects=subjects.filter(s=>String(s.class).trim().toLowerCase()===className.toLowerCase());
    if (term) subjects=subjects.filter(s=>String(s.term).trim().toLowerCase()===term.toLowerCase());
    // If staffId provided, filter to only assigned subjects
    if (staffId) {
      try {
        ensureStaffSubjectsSheet();
        const assignments = sheetToObjects("STAFF_SUBJECTS").filter(a=>
          matchCode(a.school_code,schoolCode)&&String(a.staff_id).trim()===staffId
        );
        if (assignments.length>0) {
          subjects = subjects.filter(s=>assignments.some(a=>
            String(a.subject_code).trim().toUpperCase()===String(s.subject_code).trim().toUpperCase()&&
            String(a.class).trim().toLowerCase()===String(s.class).trim().toLowerCase()
          ));
        }
      } catch(e){}
    }
    return output({success:true,subjects:subjects,count:subjects.length});
  } catch(err) { return output({success:false,message:err.message}); }
}
function addSubject(params) {
  try {
    const schoolCode  = String(params.school_code ||"").trim().toUpperCase();
    const subjectCode = String(params.subject_code||"").trim().toUpperCase();
    const subjectName = String(params.subject_name||"").trim();
    const className   = String(params.class       ||"").trim();
    const arm         = String(params.arm         ||"").trim(); // "", "Science", "Arts", "Commercial", "General", "All"
    const duration    = parseInt(params.duration_minutes)||45;
    const term        = String(params.term        ||"First Term").trim();
    const session     = String(params.session     ||"").trim();
    const cbtEnabled  = String(params.cbt_enabled ||"yes").trim();
    if (!schoolCode||!subjectCode||!subjectName||!className)
      return output({success:false,message:"school_code, subject_code, subject_name and class required"});
    const fullCode = subjectCode.startsWith(schoolCode+"-")?subjectCode:`${schoolCode}-${subjectCode}`;
    const existing = sheetToObjects("SUBJECTS");
    if (existing.find(s=>String(s.subject_code).trim().toUpperCase()===fullCode
        &&String(s.class).trim().toLowerCase()===className.toLowerCase()
        &&String(s.arm||"").trim().toLowerCase()===arm.toLowerCase()))
      return output({success:false,message:"Subject already exists for this class/arm"});
    // Build row by header to tolerate older sheets missing the `arm` column.
    const sh = getSheet("SUBJECTS");
    let headers = getHeaders("SUBJECTS");
    if (headers.indexOf("arm")===-1) {
      // Insert "arm" column right after "class"
      const classIdx = headers.indexOf("class");
      const insertAt = (classIdx===-1 ? headers.length : classIdx+1) + 1; // 1-based
      sh.insertColumnBefore(insertAt);
      sh.getRange(1, insertAt).setValue("arm");
      headers = getHeaders("SUBJECTS");
    }
    const valMap = {school_code:schoolCode, subject_code:fullCode, subject_name:subjectName, class:className, arm:arm, duration_minutes:duration, term:term, session:session, status:"active", cbt_enabled:cbtEnabled};
    const row = headers.map(h=> valMap[h]!==undefined ? valMap[h] : "");
    sh.appendRow(row);
    return output({success:true,message:"Subject added",subject_code:fullCode});
  } catch(err) { return output({success:false,message:err.message}); }
}
function updateSubject(params) {
  try {
    const schoolCode  = String(params.school_code ||"").trim().toUpperCase();
    const subjectCode = String(params.subject_code||"").trim().toUpperCase();
    if (!schoolCode||!subjectCode) return output({success:false,message:"school_code and subject_code required"});
    const sheet=getSheet("SUBJECTS"),headers=getHeaders("SUBJECTS"),data=sheet.getDataRange().getValues();
    // Optional row disambiguator when same code exists across classes/arms
    const matchClass = String(params.match_class||"").trim().toLowerCase();
    const matchArm   = String(params.match_arm  ||"").trim().toLowerCase();
    let rowIndex=-1;
    for (let i=1;i<data.length;i++){
      if(!matchCode(data[i][0],schoolCode)) continue;
      if(String(data[i][1]).trim().toUpperCase()!==subjectCode) continue;
      if(matchClass && String(data[i][3]||"").trim().toLowerCase()!==matchClass) continue;
      if(matchArm   && String(data[i][4]||"").trim().toLowerCase()!==matchArm)   continue;
      rowIndex=i+1; break;
    }
    if (rowIndex===-1) return output({success:false,message:"Subject not found"});
    ["subject_name","class","arm","duration_minutes","term","session","status","cbt_enabled"].forEach(f=>{
      if(params[f]!==undefined){const c=headers.indexOf(f);if(c!==-1)sheet.getRange(rowIndex,c+1).setValue(params[f]);}
    });
    return output({success:true,message:"Subject updated"});
  } catch(err) { return output({success:false,message:err.message}); }
}
function deleteSubject(params) {
  try {
    const schoolCode  = String(params.school_code ||"").trim().toUpperCase();
    const subjectCode = String(params.subject_code||"").trim().toUpperCase();
    if (!schoolCode||!subjectCode) return output({success:false,message:"school_code and subject_code required"});
    const sheet=getSheet("SUBJECTS"),data=sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++){
      if(matchCode(data[i][0],schoolCode)&&String(data[i][1]).trim().toUpperCase()===subjectCode){
        sheet.deleteRow(i+1); return output({success:true,message:"Subject deleted"});
      }
    }
    return output({success:false,message:"Subject not found"});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// QUESTIONS
// ============================================================
function getQuestions(params) {
  try {
    const schoolCode  = String(params.school_code ||"").trim().toUpperCase();
    const subjectCode = String(params.subject_code||"").trim().toUpperCase();
    if (!schoolCode||!subjectCode) return output({success:false,message:"school_code and subject_code required"});
    const questions = sheetToObjects("QUESTIONS").filter(q=>
      matchCode(q.school_code,schoolCode)&&String(q.subject_code).trim().toUpperCase()===subjectCode
    );
    return output({success:true,questions:questions,count:questions.length});
  } catch(err) { return output({success:false,message:err.message}); }
}
function addQuestion(params) {
  try {
    const schoolCode  = String(params.school_code ||"").trim().toUpperCase();
    const subjectCode = String(params.subject_code||"").trim().toUpperCase();
    const question    = String(params.question    ||"").trim();
    const optionA     = String(params.option_a    ||"").trim();
    const optionB     = String(params.option_b    ||"").trim();
    const optionC     = String(params.option_c    ||"").trim();
    const optionD     = String(params.option_d    ||"").trim();
    const correct     = String(params.correct     ||"").trim().toUpperCase();
    const difficulty  = String(params.difficulty  ||"easy").trim();
    if (!schoolCode||!subjectCode||!question||!optionA||!optionB||!optionC||!optionD||!correct)
      return output({success:false,message:"All question fields are required"});
    if (!["A","B","C","D"].includes(correct)) return output({success:false,message:"Correct must be A/B/C/D"});
    getSheet("QUESTIONS").appendRow([schoolCode,subjectCode,question,optionA,optionB,optionC,optionD,correct,difficulty,today()]);
    return output({success:true,message:"Question added"});
  } catch(err) { return output({success:false,message:err.message}); }
}
function deleteQuestion(params) {
  try {
    const schoolCode  = String(params.school_code ||"").trim().toUpperCase();
    const subjectCode = String(params.subject_code||"").trim().toUpperCase();
    const qText       = String(params.question    ||"").trim();
    if (!schoolCode||!subjectCode||!qText) return output({success:false,message:"school_code, subject_code and question required"});
    const sheet=getSheet("QUESTIONS"),data=sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++){
      if(matchCode(data[i][0],schoolCode)&&String(data[i][1]).trim().toUpperCase()===subjectCode&&String(data[i][2]).trim()===qText){
        sheet.deleteRow(i+1); return output({success:true,message:"Question deleted"});
      }
    }
    return output({success:false,message:"Question not found"});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// RESULTS — CBT SAVE (deduplication built-in)
// ============================================================
function saveResult(params) {
  try {
    const schoolCode     = String(params.school_code    ||"").trim().toUpperCase();
    const studentId      = String(params.student_id     ||"").trim();
    const subjectCode    = String(params.subject_code   ||"").trim().toUpperCase();
    const correctAnswers = parseInt(params.correct_answers)||0;
    const totalQuestions = parseInt(params.total_questions)||0;
    if (!schoolCode||!studentId||!subjectCode)
      return output({success:false,message:"school_code, student_id and subject_code required"});
    let examMax=60,term="",session="";
    try { const s=getSchoolSettings(schoolCode); examMax=parseFloat(s.exam_max)||60; term=s.term||""; session=s.session||""; } catch(e){}
    const examScore = scaleExamScore(correctAnswers,totalQuestions,examMax);
    const sheet=getSheet("RESULTS"),data=sheet.getDataRange().getValues();
    // Find ALL matching rows and keep only the latest — deduplicate
    const matchingRows = [];
    for (let i=1;i<data.length;i++){
      if(matchCode(data[i][0],schoolCode)&&String(data[i][1]).trim()===studentId&&String(data[i][2]).trim().toUpperCase()===subjectCode&&String(data[i][7]).trim()===term){
        matchingRows.push(i+1);
      }
    }
    // Delete duplicates (keep first, delete rest) then update first
    if (matchingRows.length>1) {
      for (let k=matchingRows.length-1;k>=1;k--) { sheet.deleteRow(matchingRows[k]); }
    }
    // Reload after potential deletions
    const freshData = sheet.getDataRange().getValues();
    for (let i=1;i<freshData.length;i++){
      if(matchCode(freshData[i][0],schoolCode)&&String(freshData[i][1]).trim()===studentId&&String(freshData[i][2]).trim().toUpperCase()===subjectCode&&String(freshData[i][7]).trim()===term){
        const t1=parseFloat(freshData[i][3])||0, t2=parseFloat(freshData[i][4])||0;
        const total=calcTotal(t1,t2,examScore);
        sheet.getRange(i+1,6).setValue(examScore);
        sheet.getRange(i+1,7).setValue(total);
        sheet.getRange(i+1,10).setValue("CBT");
        sheet.getRange(i+1,11).setValue(today());
        logActivity(schoolCode,studentId,"student","submitExam","Subject:"+subjectCode+" Score:"+examScore);
        return output({success:true,message:"Exam saved",exam_score:examScore,total:total});
      }
    }
    // No row found — create new
    const total=calcTotal(0,0,examScore);
    sheet.appendRow([schoolCode,studentId,subjectCode,0,0,examScore,total,term,session,"CBT",today()]);
    return output({success:true,message:"Exam saved",exam_score:examScore,total:total});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// RESULTS — MANUAL (deduplication built-in)
// ============================================================
function saveManualResult(params) {
  try {
    const schoolCode  = String(params.school_code ||"").trim().toUpperCase();
    const studentId   = String(params.student_id  ||"").trim();
    const subjectCode = String(params.subject_code||"").trim().toUpperCase();
    const firstTest   = parseFloat(params.first_test) ||0;
    const secondTest  = parseFloat(params.second_test)||0;
    const submittedBy = String(params.submitted_by||"admin").trim();
    const manualExam  = (params.exam!==undefined&&params.exam!=="") ? parseFloat(params.exam) : null;
    if (!schoolCode||!studentId||!subjectCode)
      return output({success:false,message:"school_code, student_id and subject_code required"});
    let firstMax=20,secondMax=20,examMax=60,term="",session="";
    try { const s=getSchoolSettings(schoolCode); firstMax=parseFloat(s.first_test_max)||20; secondMax=parseFloat(s.second_test_max)||20; examMax=parseFloat(s.exam_max)||60; term=s.term||""; session=s.session||""; } catch(e){}
    if (firstTest>firstMax) return output({success:false,message:`1st test exceeds max ${firstMax}`});
    if (secondTest>secondMax) return output({success:false,message:`2nd test exceeds max ${secondMax}`});
    if (manualExam!==null&&manualExam>examMax) return output({success:false,message:`Exam exceeds max ${examMax}`});
    const sheet=getSheet("RESULTS"),data=sheet.getDataRange().getValues();
    // Deduplicate first
    const matchingRows=[];
    for (let i=1;i<data.length;i++){
      if(matchCode(data[i][0],schoolCode)&&String(data[i][1]).trim()===studentId&&String(data[i][2]).trim().toUpperCase()===subjectCode&&String(data[i][7]).trim()===term){
        matchingRows.push(i+1);
      }
    }
    if (matchingRows.length>1) { for (let k=matchingRows.length-1;k>=1;k--){sheet.deleteRow(matchingRows[k]);} }
    const freshData=sheet.getDataRange().getValues();
    for (let i=1;i<freshData.length;i++){
      if(matchCode(freshData[i][0],schoolCode)&&String(freshData[i][1]).trim()===studentId&&String(freshData[i][2]).trim().toUpperCase()===subjectCode&&String(freshData[i][7]).trim()===term){
        const existExam=(manualExam!==null)?manualExam:(parseFloat(freshData[i][5])||0);
        const total=calcTotal(firstTest,secondTest,existExam);
        sheet.getRange(i+1,4).setValue(firstTest);
        sheet.getRange(i+1,5).setValue(secondTest);
        if (manualExam!==null) sheet.getRange(i+1,6).setValue(manualExam);
        sheet.getRange(i+1,7).setValue(total);
        sheet.getRange(i+1,10).setValue(submittedBy);
        sheet.getRange(i+1,11).setValue(today());
        return output({success:true,message:"Scores updated",total:total});
      }
    }
    const examScore=(manualExam!==null)?manualExam:0;
    const total=calcTotal(firstTest,secondTest,examScore);
    sheet.appendRow([schoolCode,studentId,subjectCode,firstTest,secondTest,examScore,total,term,session,submittedBy,today()]);
    return output({success:true,message:"Scores saved",total:total});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// RESULTS — GET (deduplication + name enrichment)
// ============================================================
function getResults(params) {
  try {
    const schoolCode  = String(params.school_code ||"").trim().toUpperCase();
    const studentId   = String(params.student_id  ||"").trim();
    const subjectCode = String(params.subject_code||"").trim().toUpperCase();
    const term        = String(params.term        ||"").trim();
    const classFilter = String(params.class       ||"").trim();
    let resolvedCode = schoolCode;
    if (!resolvedCode&&studentId) {
      const match=sheetToObjects("STUDENTS").find(s=>String(s.student_id).trim()===studentId);
      if (match) resolvedCode=String(match.school_code).trim().toUpperCase();
    }
    if (!resolvedCode) return output({success:false,message:"school_code required"});
    let results=sheetToObjects("RESULTS").filter(r=>matchCode(r.school_code,resolvedCode));
    if (studentId)   results=results.filter(r=>String(r.student_id).trim()===studentId);
    if (subjectCode) results=results.filter(r=>String(r.subject_code).trim().toUpperCase()===subjectCode);
    if (term)        results=results.filter(r=>String(r.term).trim().toLowerCase()===term.toLowerCase());
    if (classFilter) results=results.filter(r=>String(r.class||"").trim()===classFilter);
    // DEDUPLICATION on read — keep the row with the highest total per student+subject combo
    const seen = {};
    const deduped = [];
    results.forEach(r=>{
      const key = String(r.student_id).trim()+"||"+String(r.subject_code).trim().toUpperCase();
      const total = calcTotal(r.first_test,r.second_test,r.exam);
      if (!seen[key] || total > seen[key].total) { seen[key]=Object.assign({},r,{_total:total}); }
    });
    results = Object.values(seen);
    const subjects=sheetToObjects("SUBJECTS").filter(s=>matchCode(s.school_code,resolvedCode));
    const students=sheetToObjects("STUDENTS").filter(s=>matchCode(s.school_code,resolvedCode));
    const studentMap={};
    students.forEach(s=>{const sid=String(s.student_id).trim();const safe=Object.assign({},s);delete safe.exam_pin;studentMap[sid]=safe;});
    const studentInfo=studentId?(studentMap[studentId]||{}):{};
    const schoolInfo=getSchoolInfo(resolvedCode);
    let schoolSettings={};
    try { schoolSettings=getSchoolSettings(resolvedCode); } catch(e){}
    // Build per-student eligibility (school must offer subject for student's class + arm/stream)
    function _norm(v){ return String(v==null?"":v).trim().toLowerCase(); }
    function _isGeneralArm(a){ a=_norm(a); return !a || a==="general" || a==="all" || a==="-"; }
    function _eligibleFor(stu){
      const stuClass=_norm(stu.class), stuArm=_norm(stu.arm);
      const codes=new Set();
      subjects.forEach(s=>{
        if (_norm(s.class)!==stuClass) return;
        const subjArm=_norm(s.arm);
        // subject without arm = general (everyone in that class). With arm = only that stream.
        if (_isGeneralArm(subjArm) || subjArm===stuArm) {
          codes.add(String(s.subject_code).trim().toUpperCase());
        }
      });
      return codes;
    }
    const _eligCache={};
    results=results.map(r=>{
      const subj=subjects.find(s=>String(s.subject_code).trim().toUpperCase()===String(r.subject_code).trim().toUpperCase());
      const rowStu=studentMap[String(r.student_id).trim()]||{};
      // Drop entirely unrecorded rows so averages reflect only recorded subjects.
      const t1=String(r.first_test==null?"":r.first_test).trim();
      const t2=String(r.second_test==null?"":r.second_test).trim();
      const ex=String(r.exam==null?"":r.exam).trim();
      const hasAny = (t1!==""&&t1!=="0"&&t1!=="0.0") || (t2!==""&&t2!=="0"&&t2!=="0.0") || (ex!==""&&ex!=="0"&&ex!=="0.0");
      if (!hasAny) return null;
      const total=calcTotal(r.first_test,r.second_test,r.exam);
      const g=getGrade(total);
      // Stream/offered filter — drop results whose subject is not offered for this student's class+arm.
      const sid=String(r.student_id).trim();
      if (!_eligCache[sid]) _eligCache[sid]=_eligibleFor(rowStu);
      const elig=_eligCache[sid];
      // If we have any subject defs for the student's class, enforce eligibility; otherwise pass through (legacy data).
      const anyForClass=subjects.some(s=>_norm(s.class)===_norm(rowStu.class));
      if (anyForClass && !elig.has(String(r.subject_code).trim().toUpperCase())) return null;
      return Object.assign({},r,{
        subject_name:subj?subj.subject_name:r.subject_code,
        full_name:   rowStu.full_name||"",
        class:       rowStu.class   ||r.class||"",
        arm:         rowStu.arm     ||"",
        total:total, grade:g.grade, remark:g.remark
      });
    }).filter(Boolean);
    const grandTotal=results.reduce((acc,r)=>acc+(parseFloat(r.total)||0),0);
    const average=results.length>0?Math.round((grandTotal/results.length)*10)/10:0;
    const overallG=getGrade(average);
    return output({success:true,results:results,student:studentInfo,
      grand_total:grandTotal,average:average,
      overall_grade:overallG.grade,overall_remark:overallG.remark,subject_count:results.length,
      schoolSettings:{
        school_name:    schoolInfo.school_name    ||schoolSettings.school_name||resolvedCode,
        address:        schoolInfo.address        ||schoolSettings.address    ||"",
        town:           schoolInfo.town           ||schoolSettings.town       ||"",
        state:          schoolInfo.state          ||schoolSettings.state      ||"",
        logo_url:       schoolInfo.logo_url       ||"",
        principal_name: schoolSettings.principal_name||"",
        motto:          schoolSettings.motto      ||"",
        term:           schoolSettings.term       ||"",
        session:        schoolSettings.session    ||"",
        first_test_max: schoolSettings.first_test_max ||20,
        second_test_max:schoolSettings.second_test_max||20,
        exam_max:       schoolSettings.exam_max       ||60
      }
    });
  } catch(err) { return output({success:false,message:err.message}); }
}
function getClassResults(params) {
  try {
    const schoolCode=String(params.school_code||"").trim().toUpperCase();
    const className =String(params.class      ||"").trim();
    const term      =String(params.term       ||"").trim();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    let students=sheetToObjects("STUDENTS").filter(s=>matchCode(s.school_code,schoolCode));
    if (className) students=students.filter(s=>String(s.class).trim().toLowerCase()===className.toLowerCase());
    let allResults=sheetToObjects("RESULTS").filter(r=>matchCode(r.school_code,schoolCode));
    if (term) allResults=allResults.filter(r=>String(r.term).trim().toLowerCase()===term.toLowerCase());
    const subjects=sheetToObjects("SUBJECTS").filter(s=>matchCode(s.school_code,schoolCode));
    function _norm(v){ return String(v==null?"":v).trim().toLowerCase(); }
    function _isGeneralArm(a){ a=_norm(a); return !a||a==="general"||a==="all"||a==="-"; }
    function _eligibleFor(stu){
      const stuClass=_norm(stu.class), stuArm=_norm(stu.arm);
      const codes=new Set();
      subjects.forEach(s=>{
        if (_norm(s.class)!==stuClass) return;
        const sa=_norm(s.arm);
        if (_isGeneralArm(sa)||sa===stuArm) codes.add(String(s.subject_code).trim().toUpperCase());
      });
      return codes;
    }
    function _hasAny(r){
      const a=String(r.first_test==null?"":r.first_test).trim();
      const b=String(r.second_test==null?"":r.second_test).trim();
      const c=String(r.exam==null?"":r.exam).trim();
      return (a!==""&&a!=="0"&&a!=="0.0")||(b!==""&&b!=="0"&&b!=="0.0")||(c!==""&&c!=="0"&&c!=="0.0");
    }
    const summary=students.map(stu=>{
      const elig=_eligibleFor(stu);
      const anyForClass=subjects.some(s=>_norm(s.class)===_norm(stu.class));
      const stuResults=allResults.filter(r=>String(r.student_id).trim()===String(stu.student_id).trim());
      const seen={};
      stuResults.forEach(r=>{
        if (!_hasAny(r)) return;
        const key=String(r.subject_code).trim().toUpperCase();
        if (anyForClass && !elig.has(key)) return;
        const total=calcTotal(r.first_test,r.second_test,r.exam);
        if(!seen[key]||total>seen[key]._total) seen[key]=Object.assign({},r,{_total:total});
      });
      const enriched=Object.values(seen).map(r=>{
        const subj=subjects.find(s=>String(s.subject_code).trim().toUpperCase()===String(r.subject_code).trim().toUpperCase());
        const total=calcTotal(r.first_test,r.second_test,r.exam);
        const g=getGrade(total);
        return Object.assign({},r,{subject_name:subj?subj.subject_name:r.subject_code,total,grade:g.grade,remark:g.remark});
      });
      const grandTotal=enriched.reduce((acc,r)=>acc+(parseFloat(r.total)||0),0);
      const average=enriched.length>0?Math.round((grandTotal/enriched.length)*10)/10:0;
      return {student_id:stu.student_id,full_name:stu.full_name,class:stu.class,arm:stu.arm||"",results:enriched,grand_total:grandTotal,average:average,subject_count:enriched.length,grade:getGrade(average).grade,remark:getGrade(average).remark};
    });
    return output({success:true,class_results:summary,count:summary.length});
  } catch(err) { return output({success:false,message:err.message}); }
}
function getClassPositions(params) {
  try {
    const schoolCode=String(params.school_code||"").trim().toUpperCase();
    const className =String(params.class      ||"").trim();
    const term      =String(params.term       ||"").trim();
    if (!schoolCode||!className) return output({success:false,message:"school_code and class required"});
    const settings=getSchoolSettings(schoolCode);
    if (String(settings.result_display||"grade_only").trim().toLowerCase()==="grade_only")
      return output({success:true,positions:[],position_enabled:false});
    const students=sheetToObjects("STUDENTS").filter(s=>matchCode(s.school_code,schoolCode)&&String(s.class).trim().toLowerCase()===className.toLowerCase());
    let allResults=sheetToObjects("RESULTS").filter(r=>matchCode(r.school_code,schoolCode));
    if (term) allResults=allResults.filter(r=>String(r.term).trim().toLowerCase()===term.toLowerCase());
    const subjects=sheetToObjects("SUBJECTS").filter(s=>matchCode(s.school_code,schoolCode));
    function _norm(v){ return String(v==null?"":v).trim().toLowerCase(); }
    function _isGeneralArm(a){ a=_norm(a); return !a||a==="general"||a==="all"||a==="-"; }
    function _eligibleFor(stu){
      const stuClass=_norm(stu.class), stuArm=_norm(stu.arm);
      const codes=new Set();
      subjects.forEach(s=>{
        if (_norm(s.class)!==stuClass) return;
        const sa=_norm(s.arm);
        if (_isGeneralArm(sa)||sa===stuArm) codes.add(String(s.subject_code).trim().toUpperCase());
      });
      return codes;
    }
    function _hasAny(r){
      const a=String(r.first_test==null?"":r.first_test).trim();
      const b=String(r.second_test==null?"":r.second_test).trim();
      const c=String(r.exam==null?"":r.exam).trim();
      return (a!==""&&a!=="0"&&a!=="0.0")||(b!==""&&b!=="0"&&b!=="0.0")||(c!==""&&c!=="0"&&c!=="0.0");
    }
    const anyForClass=subjects.some(s=>_norm(s.class)===_norm(className));
    const averages=students.map(stu=>{
      const elig=_eligibleFor(stu);
      const stuResults=allResults.filter(r=>String(r.student_id).trim()===String(stu.student_id).trim());
      const seen={};
      stuResults.forEach(r=>{
        if (!_hasAny(r)) return;
        const k=String(r.subject_code).trim().toUpperCase();
        if (anyForClass && !elig.has(k)) return;
        const t=calcTotal(r.first_test,r.second_test,r.exam);
        if(!seen[k]||t>seen[k])seen[k]=t;
      });
      const totals=Object.values(seen);
      const grandTotal=totals.reduce((a,b)=>a+b,0);
      const average=totals.length>0?Math.round((grandTotal/totals.length)*10)/10:0;
      return {student_id:stu.student_id,full_name:stu.full_name,arm:stu.arm||"",average,position:0};
    });
    averages.sort((a,b)=>b.average-a.average);
    let pos=1;
    averages.forEach((s,i)=>{if(i>0&&s.average<averages[i-1].average)pos=i+1;s.position=pos;});
    return output({success:true,positions:averages,position_enabled:true,class:className,total_students:averages.length});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// DEDUPLICATE RESULTS — Admin utility to clean existing data
// ============================================================
function deduplicateResults(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    let settings = {};
    try { settings = getSchoolSettings(schoolCode); } catch(e){}
    const term = settings.term || "";
    const sheet = getSheet("RESULTS");
    const data  = sheet.getDataRange().getValues();
    const headers = data[0].map(h=>String(h).trim().toLowerCase().replace(/ /g,"_"));
    const scIdx=headers.indexOf("school_code"), sidIdx=headers.indexOf("student_id"),
          subIdx=headers.indexOf("subject_code"), termIdx=headers.indexOf("term"),
          totalIdx=headers.indexOf("total");
    // Group rows by school+student+subject+term
    const groups = {};
    for (let i=1;i<data.length;i++) {
      if (!matchCode(data[i][scIdx],schoolCode)) continue;
      const key = String(data[i][sidIdx]).trim()+"||"+String(data[i][subIdx]).trim().toUpperCase()+"||"+String(data[i][termIdx]).trim();
      if (!groups[key]) groups[key]=[];
      groups[key].push({rowNum:i+1,total:parseFloat(data[i][totalIdx])||0});
    }
    // For groups with >1 row, delete all but the one with highest total
    let deleted = 0;
    const rowsToDelete = [];
    Object.values(groups).forEach(rows=>{
      if (rows.length<=1) return;
      rows.sort((a,b)=>b.total-a.total);
      for (let k=1;k<rows.length;k++) rowsToDelete.push(rows[k].rowNum);
    });
    // Delete from bottom up
    rowsToDelete.sort((a,b)=>b-a);
    rowsToDelete.forEach(r=>{sheet.deleteRow(r);deleted++;});
    return output({success:true,message:`Deduplication complete. Removed ${deleted} duplicate row(s).`,removed:deleted});
  } catch(err) { return output({success:false,message:err.message}); }
}

// ============================================================
// ATTENDANCE
// ============================================================
function getAttendance(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const classFilter = params.class;
    const dateFilter = params.date;
    const term = params.term;
    const session = params.session;
    
    let sheet;
    try { sheet = getSheet("ATTENDANCE"); } catch(e) { return output({success:true,attendance:[]}); }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return output({success:true,attendance:[]});
    
    const headers = data[0].map(h=>String(h).trim().toLowerCase().replace(/ /g,"_"));
    const scIdx=headers.indexOf("school_code"), clsIdx=headers.indexOf("class"),
          dateIdx=headers.indexOf("date"), statusIdx=headers.indexOf("status"),
          termIdx=headers.indexOf("term"), sessIdx=headers.indexOf("session"),
          stuIdx=headers.indexOf("student_id");
    
    let results = [];
    for (let i=1;i<data.length;i++) {
      if (!matchCode(data[i][scIdx],schoolCode)) continue;
      if (classFilter && String(data[i][clsIdx]).trim() !== classFilter) continue;
      if (dateFilter && String(data[i][dateIdx]).trim() !== dateFilter) continue;
      if (term && String(data[i][termIdx]).trim() !== term) continue;
      if (session && String(data[i][sessIdx]).trim() !== session) continue;
      
      results.push({
        student_id: data[i][stuIdx],
        class: data[i][clsIdx],
        date: data[i][dateIdx],
        status: data[i][statusIdx],
        term: data[i][termIdx],
        session: data[i][sessIdx]
      });
    }
    return output({success:true,attendance:results});
  } catch(err) { return output({success:false,message:err.message}); }
}

function markAttendance(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    
    const studentIds = params.student_ids ? JSON.parse(params.student_ids) : [];
    const className = params.class;
    const date = params.date;
    const status = params.status;
    const recordedBy = params.recorded_by || "admin";
    const term = params.term;
    const session = params.session;
    
    if (!studentIds.length || !className || !date || !status) {
      return output({success:false,message:"student_ids, class, date, status required"});
    }
    
    let sheet;
    try { sheet = getSheet("ATTENDANCE"); } catch(e) {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      sheet = ss.insertSheet("ATTENDANCE");
      sheet.getRange(1,1,1,8).setValues([["school_code","student_id","class","date","status","recorded_by","term","session"]]);
    }
    
    const existingData = sheet.getDataRange().getValues();
    const existingMap = {};
    for (let i=1;i<existingData.length;i++) {
      const key = String(existingData[i][1]).trim() + "||" + String(existingData[i][3]).trim();
      existingMap[key] = i+1;
    }
    
    const newRows = [];
    studentIds.forEach(sid => {
      const key = sid + "||" + date;
      if (existingMap[key]) {
        sheet.getRange(existingMap[key], 5).setValue(status);
      } else {
        newRows.push([schoolCode, sid, className, date, status, recordedBy, term || "", session || ""]);
      }
    });
    
    if (newRows.length) {
      sheet.getRange(existingData.length+1, 1, newRows.length, 8).setValues(newRows);
    }
    
    return output({success:true,message:`Attendance marked for ${studentIds.length} students`});
  } catch(err) { return output({success:false,message:err.message}); }
}

// ============================================================
// ANNOUNCEMENTS
// ============================================================
function getAnnouncements(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const role = params.target_role;
    
    let sheet;
    try { sheet = getSheet("ANNOUNCEMENTS"); } catch(e) { return output({success:true,announcements:[]}); }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return output({success:true,announcements:[]});
    
    const headers = data[0].map(h=>String(h).trim().toLowerCase().replace(/ /g,"_"));
    const scIdx=headers.indexOf("school_code"), titleIdx=headers.indexOf("title"),
          contIdx=headers.indexOf("content"), prioIdx=headers.indexOf("priority"),
          roleIdx=headers.indexOf("target_role"), startIdx=headers.indexOf("start_date"),
          endIdx=headers.indexOf("end_date"), crByIdx=headers.indexOf("created_by"),
          crAtIdx=headers.indexOf("created_at");
    
    const now = new Date();
    let results = [];
    for (let i=1;i<data.length;i++) {
      if (!matchCode(data[i][scIdx],schoolCode)) continue;
      if (role && String(data[i][roleIdx]).trim() !== role && data[i][roleIdx] !== "all") continue;
      
      const startDate = data[i][startIdx] ? new Date(data[i][startIdx]) : null;
      const endDate = data[i][endIdx] ? new Date(data[i][endIdx]) : null;
      if (startDate && now < startDate) continue;
      if (endDate && now > endDate) continue;
      
      results.push({
        id: i,
        title: data[i][titleIdx],
        content: data[i][contIdx],
        priority: data[i][prioIdx],
        target_role: data[i][roleIdx],
        start_date: data[i][startIdx],
        end_date: data[i][endIdx],
        created_by: data[i][crByIdx],
        created_at: data[i][crAtIdx]
      });
    }
    results.sort((a,b) => (b.priority||"").localeCompare(a.priority||""));
    return output({success:true,announcements:results});
  } catch(err) { return output({success:false,message:err.message}); }
}

function addAnnouncement(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    
    const title = params.title;
    const content = params.content;
    const priority = params.priority || "normal";
    const targetRole = params.target_role || "all";
    const targetClass = params.target_class || "";
    const startDate = params.start_date || new Date().toISOString().split("T")[0];
    const endDate = params.end_date || "";
    const createdBy = params.created_by || "admin";
    
    if (!title || !content) return output({success:false,message:"title and content required"});
    
    let sheet;
    try { sheet = getSheet("ANNOUNCEMENTS"); } catch(e) {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      sheet = ss.insertSheet("ANNOUNCEMENTS");
      sheet.getRange(1,1,1,10).setValues([["school_code","title","content","priority","target_role","target_class","start_date","end_date","created_by","created_at"]]);
    }
    
    const now = new Date().toISOString();
    sheet.appendRow([schoolCode, title, content, priority, targetRole, targetClass, startDate, endDate, createdBy, now]);
    return output({success:true,message:"Announcement created"});
  } catch(err) { return output({success:false,message:err.message}); }
}

function deleteAnnouncement(params) {
  try {
    const rowNum = parseInt(params.row_num);
    if (!rowNum) return output({success:false,message:"row_num required"});
    
    const sheet = getSheet("ANNOUNCEMENTS");
    sheet.deleteRow(rowNum);
    return output({success:true,message:"Announcement deleted"});
  } catch(err) { return output({success:false,message:err.message}); }
}

// ============================================================
// EVENT CALENDAR
// ============================================================
function getEvents(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const startDate = params.start_date;
    const endDate = params.end_date;
    const eventType = params.event_type;
    const targetRole = params.target_role;
    const targetClass = params.target_class;
    
    if (!schoolCode) return output({success:false,message:"school_code required"});
    
    let sheet;
    try { sheet = getSheet("EVENTS"); } catch(e) { return output({success:true,events:[]}); }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return output({success:true,events:[]});
    
    const headers = data[0].map(h=>String(h).trim().toLowerCase().replace(/ /g,"_"));
    const scIdx=headers.indexOf("school_code"), idIdx=headers.indexOf("event_id"),
          titleIdx=headers.indexOf("title"), descIdx=headers.indexOf("description"),
          typeIdx=headers.indexOf("event_type"), sDateIdx=headers.indexOf("start_date"),
          eDateIdx=headers.indexOf("end_date"), allDayIdx=headers.indexOf("all_day"),
          locIdx=headers.indexOf("location"), clsIdx=headers.indexOf("target_class"),
          roleIdx=headers.indexOf("target_role"), crByIdx=headers.indexOf("created_by"),
          crAtIdx=headers.indexOf("created_at");
    
    let results = [];
    for (let i=1;i<data.length;i++) {
      if (!matchCode(data[i][scIdx],schoolCode)) continue;
      
      const eventStart = data[i][sDateIdx] ? new Date(data[i][sDateIdx]) : null;
      const eventEnd = data[i][eDateIdx] ? new Date(data[i][eDateIdx]) : null;
      
      // Filter by date range
      if (startDate && eventEnd && eventEnd < new Date(startDate)) continue;
      if (endDate && eventStart && eventStart > new Date(endDate)) continue;
      
      // Filter by event type
      if (eventType && String(data[i][typeIdx]).trim().toLowerCase() !== eventType.toLowerCase()) continue;
      
      // Filter by target role
      if (targetRole && String(data[i][roleIdx]).trim() !== targetRole && data[i][roleIdx] !== "all") continue;
      
      // Filter by target class
      if (targetClass && String(data[i][clsIdx]).trim() !== targetClass && data[i][clsIdx] !== "all") continue;
      
      results.push({
        id: i,
        event_id: data[i][idIdx],
        title: data[i][titleIdx],
        description: data[i][descIdx],
        event_type: data[i][typeIdx],
        start_date: data[i][sDateIdx],
        end_date: data[i][eDateIdx],
        all_day: data[i][allDayIdx] === true || data[i][allDayIdx] === "yes" || data[i][allDayIdx] === "true",
        location: data[i][locIdx],
        target_class: data[i][clsIdx],
        target_role: data[i][roleIdx],
        created_by: data[i][crByIdx],
        created_at: data[i][crAtIdx]
      });
    }
    
    // Sort by start date
    results.sort((a,b) => {
      if (!a.start_date) return 1;
      if (!b.start_date) return -1;
      return new Date(a.start_date) - new Date(b.start_date);
    });
    
    return output({success:true,events:results,count:results.length});
  } catch(err) { return output({success:false,message:err.message}); }
}

function addEvent(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    
    const title = String(params.title||"").trim();
    const description = String(params.description||"").trim();
    const eventType = String(params.event_type||"general").trim();
    const startDate = params.start_date;
    const endDate = params.end_date || params.start_date;
    const allDay = params.all_day === true || params.all_day === "yes" || params.all_day === "true" || params.all_day === "on";
    const location = String(params.location||"").trim();
    const targetRole = String(params.target_role||"all").trim();
    const targetClass = String(params.target_class||"all").trim();
    const createdBy = String(params.created_by||"admin").trim();
    
    if (!title || !startDate) return output({success:false,message:"title and start_date required"});
    
    let sheet;
    try { sheet = getSheet("EVENTS"); } catch(e) {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      sheet = ss.insertSheet("EVENTS");
      sheet.getRange(1,1,1,13).setValues([["school_code","event_id","title","description","event_type","start_date","end_date","all_day","location","target_class","target_role","created_by","created_at"]]);
    }
    
    const eventId = "EVT-" + schoolCode.substring(0,3) + "-" + randomSuffix(4).toUpperCase();
    const now = new Date().toISOString();
    
    sheet.appendRow([schoolCode, eventId, title, description, eventType, startDate, endDate, allDay ? "yes" : "no", location, targetClass, targetRole, createdBy, now]);
    
    logActivity(schoolCode, createdBy, "admin", "addEvent", "Event created: " + title);
    return output({success:true,message:"Event created",event_id:eventId});
  } catch(err) { return output({success:false,message:err.message}); }
}

function deleteEvent(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const eventId = String(params.event_id||"").trim();
    const rowNum = parseInt(params.row_num);
    
    if (!schoolCode) return output({success:false,message:"school_code required"});
    if (!eventId && !rowNum) return output({success:false,message:"event_id or row_num required"});
    
    const sheet = getSheet("EVENTS");
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h=>String(h).trim().toLowerCase().replace(/ /g,"_"));
    const idIdx = headers.indexOf("event_id");
    
    for (let i=1;i<data.length;i++) {
      if (!matchCode(data[i][0],schoolCode)) continue;
      if (eventId && String(data[i][idIdx]).trim() !== eventId) continue;
      if (rowNum && i+1 !== rowNum) continue;
      
      sheet.deleteRow(i+1);
      logActivity(schoolCode, "admin", "admin", "deleteEvent", "Event deleted: " + eventId);
      return output({success:true,message:"Event deleted"});
    }
    
    return output({success:false,message:"Event not found"});
  } catch(err) { return output({success:false,message:err.message}); }
}

// ============================================================
// STUDENT PROFILE
// ============================================================
function getStudentProfile(params) {
  try {
    const studentId = String(params.student_id||"").trim();
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    
    if (!studentId) return output({success:false,message:"student_id required"});
    
    let resolvedCode = schoolCode;
    if (!resolvedCode) {
      const student = sheetToObjects("STUDENTS").find(s=>String(s.student_id).trim()===studentId);
      if (student) resolvedCode = String(student.school_code).trim().toUpperCase();
    }
    
    if (!resolvedCode) return output({success:false,message:"school_code required or student not found"});
    
    const students = sheetToObjects("STUDENTS").filter(s=>matchCode(s.school_code,resolvedCode) && String(s.student_id).trim()===studentId);
    if (students.length === 0) return output({success:false,message:"Student not found"});
    
    const student = students[0];
    const safe = Object.assign({},student);
    delete safe.exam_pin;
    
    // Get additional info
    let settings = {};
    try { settings = getSchoolSettings(resolvedCode); } catch(e) {}
    
    return output({success:true,student:safe,settings:settings});
  } catch(err) { return output({success:false,message:err.message}); }
}

function updateStudentProfile(params) {
  try {
    const studentId = String(params.student_id||"").trim();
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    
    if (!studentId) return output({success:false,message:"student_id required"});
    
    let resolvedCode = schoolCode;
    if (!resolvedCode) {
      const student = sheetToObjects("STUDENTS").find(s=>String(s.student_id).trim()===studentId);
      if (student) resolvedCode = String(student.school_code).trim().toUpperCase();
    }
    
    if (!resolvedCode) return output({success:false,message:"school_code required or student not found"});
    
    const sheet = getSheet("STUDENTS");
    const headers = getHeaders("STUDENTS");
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i=1;i<data.length;i++) {
      if (matchCode(data[i][0],resolvedCode) && String(data[i][1]).trim() === studentId) {
        rowIndex = i+1;
        break;
      }
    }
    
    if (rowIndex === -1) return output({success:false,message:"Student not found"});
    
    // Update allowed fields
    const updatable = ["full_name","gender","date_of_birth","parent_phone","photo_url","address","emergency_contact","emergency_phone"];
    updatable.forEach(field => {
      if (params[field] !== undefined) {
        const colIdx = headers.indexOf(field);
        if (colIdx !== -1) {
          sheet.getRange(rowIndex, colIdx+1).setValue(params[field]);
        }
      }
    });
    
    logActivity(resolvedCode, studentId, "student", "updateProfile", "Profile updated");
    return output({success:true,message:"Profile updated"});
  } catch(err) { return output({success:false,message:err.message}); }
}

// ============================================================
// STAFF PROFILE
// ============================================================
function getStaffProfile(params) {
  try {
    const staffId = String(params.staff_id||"").trim();
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    
    if (!staffId) return output({success:false,message:"staff_id required"});
    
    let resolvedCode = schoolCode;
    if (!resolvedCode) {
      const staff = sheetToObjects("STAFF").find(s=>String(s.staff_id).trim()===staffId);
      if (staff) resolvedCode = String(staff.school_code).trim().toUpperCase();
    }
    
    if (!resolvedCode) return output({success:false,message:"school_code required or staff not found"});
    
    const staffList = sheetToObjects("STAFF").filter(s=>matchCode(s.school_code,resolvedCode) && String(s.staff_id).trim()===staffId);
    if (staffList.length === 0) return output({success:false,message:"Staff not found"});
    
    const staff = staffList[0];
    const safe = Object.assign({},staff);
    delete safe.password;
    
    // Get assigned subjects
    let assignedSubjects = [];
    try {
      ensureStaffSubjectsSheet();
      const assignments = sheetToObjects("STAFF_SUBJECTS").filter(a=>
        matchCode(a.school_code,resolvedCode) && String(a.staff_id).trim()===staffId
      );
      const allSubjects = sheetToObjects("SUBJECTS").filter(s=>matchCode(s.school_code,resolvedCode));
      assignedSubjects = assignments.map(a => {
        const subj = allSubjects.find(s =>
          String(s.subject_code).trim().toUpperCase() === String(a.subject_code).trim().toUpperCase()
        );
        return subj ? subj : {subject_code: a.subject_code, class: a.class};
      });
    } catch(e) {}
    
    return output({success:true,staff:safe,assigned_subjects:assignedSubjects});
  } catch(err) { return output({success:false,message:err.message}); }
}

function updateStaffProfile(params) {
  try {
    const staffId = String(params.staff_id||"").trim();
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    
    if (!staffId) return output({success:false,message:"staff_id required"});
    
    let resolvedCode = schoolCode;
    if (!resolvedCode) {
      const staff = sheetToObjects("STAFF").find(s=>String(s.staff_id).trim()===staffId);
      if (staff) resolvedCode = String(staff.school_code).trim().toUpperCase();
    }
    
    if (!resolvedCode) return output({success:false,message:"school_code required or staff not found"});
    
    const sheet = getSheet("STAFF");
    const headers = getHeaders("STAFF");
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i=1;i<data.length;i++) {
      if (matchCode(data[i][0],resolvedCode) && String(data[i][1]).trim() === staffId) {
        rowIndex = i+1;
        break;
      }
    }
    
    if (rowIndex === -1) return output({success:false,message:"Staff not found"});
    
    // Update allowed fields
    const updatable = ["full_name","email","phone","photo_url","qualifications","bio","address"];
    updatable.forEach(field => {
      if (params[field] !== undefined) {
        const colIdx = headers.indexOf(field);
        if (colIdx !== -1) {
          sheet.getRange(rowIndex, colIdx+1).setValue(params[field]);
        }
      }
    });
    
    logActivity(resolvedCode, staffId, "staff", "updateProfile", "Profile updated");
    return output({success:true,message:"Profile updated"});
  } catch(err) { return output({success:false,message:err.message}); }
}

// ============================================================
// BULK IMPORT
// ============================================================
function bulkImportStudents(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const csvData = String(params.csv_data||"").trim();
    
    if (!schoolCode) return output({success:false,message:"school_code required"});
    if (!csvData) return output({success:false,message:"csv_data required"});
    
    const lines = csvData.split('\n').filter(l=>l.trim());
    if (lines.length < 2) return output({success:false,message:"CSV must have header and at least one data row"});
    
    const headers = lines[0].split(',').map(h=>String(h).trim().toLowerCase());
    const requiredHeaders = ['student_id','full_name','class'];
    const missing = requiredHeaders.filter(h=>!headers.includes(h));
    if (missing.length > 0) return output({success:false,message:"Missing required headers: " + missing.join(', ')});
    
    const existingStudents = sheetToObjects("STUDENTS");
    const results = {success:0,failed:0,errors:[]};
    
    for (let i=1;i<lines.length;i++) {
      try {
        const values = lines[i].split(',').map(v=>String(v).trim());
        const row = {};
        headers.forEach((h,idx) => row[h] = values[idx] || "");
        
        const studentId = row.student_id;
        const fullName = row.full_name;
        const className = row.class;
        
        if (!studentId || !fullName || !className) {
          results.failed++;
          results.errors.push("Row " + (i+1) + ": Missing required fields");
          continue;
        }
        
        // Check for duplicates
        if (existingStudents.find(s=>String(s.student_id).trim() === studentId)) {
          results.failed++;
          results.errors.push("Row " + (i+1) + ": Student ID already exists - " + studentId);
          continue;
        }
        
        const pin = row.exam_pin || generatePin();
        const gender = row.gender || "";
        const dob = row.date_of_birth || "";
        const phone = row.parent_phone || "";
        
        getSheet("STUDENTS").appendRow([schoolCode, studentId, fullName, className, pin, gender, dob, phone, "active", today()]);
        results.success++;
      } catch(e) {
        results.failed++;
        results.errors.push("Row " + (i+1) + ": " + e.message);
      }
    }
    
    logActivity(schoolCode, "admin", "admin", "bulkImportStudents", `Imported ${results.success} students, ${results.failed} failed`);
    return output({success:true,message:`Import complete: ${results.success} added, ${results.failed} failed`,results:results});
  } catch(err) { return output({success:false,message:err.message}); }
}

function bulkImportStaff(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const csvData = String(params.csv_data||"").trim();
    
    if (!schoolCode) return output({success:false,message:"school_code required"});
    if (!csvData) return output({success:false,message:"csv_data required"});
    
    const lines = csvData.split('\n').filter(l=>l.trim());
    if (lines.length < 2) return output({success:false,message:"CSV must have header and at least one data row"});
    
    const headers = lines[0].split(',').map(h=>String(h).trim().toLowerCase());
    const requiredHeaders = ['full_name','role'];
    const missing = requiredHeaders.filter(h=>!headers.includes(h));
    if (missing.length > 0) return output({success:false,message:"Missing required headers: " + missing.join(', ')});
    
    const existingStaff = sheetToObjects("STAFF").filter(s=>matchCode(s.school_code,schoolCode));
    const results = {success:0,failed:0,errors:[]};
    
    for (let i=1;i<lines.length;i++) {
      try {
        const values = lines[i].split(',').map(v=>String(v).trim());
        const row = {};
        headers.forEach((h,idx) => row[h] = values[idx] || "");
        
        const fullName = row.full_name;
        const role = row.role || "teacher";
        
        if (!fullName) {
          results.failed++;
          results.errors.push("Row " + (i+1) + ": Missing full_name");
          continue;
        }
        
        const staffId = `${schoolCode}-TCH-${String(existingStaff.length + results.success + 1).padStart(3,"0")}`;
        const email = row.email || "";
        const phone = row.phone || "";
        const subjectCode = row.subject_code || "";
        const className = row.class || "";
        const password = row.password || generatePin();
        
        getSheet("STAFF").appendRow([schoolCode, staffId, fullName, email, phone, subjectCode, className, role, password, "active", today()]);
        
        // Auto-assign subject if provided
        if (subjectCode && className) {
          try {
            ensureStaffSubjectsSheet();
            const fullCode = subjectCode.startsWith(schoolCode+"-") ? subjectCode : `${schoolCode}-${subjectCode}`;
            getSheet("STAFF_SUBJECTS").appendRow([schoolCode, staffId, fullCode, className, today()]);
          } catch(e) {}
        }
        
        results.success++;
      } catch(e) {
        results.failed++;
        results.errors.push("Row " + (i+1) + ": " + e.message);
      }
    }
    
    logActivity(schoolCode, "admin", "admin", "bulkImportStaff", `Imported ${results.success} staff, ${results.failed} failed`);
    return output({success:true,message:`Import complete: ${results.success} added, ${results.failed} failed`,results:results});
  } catch(err) { return output({success:false,message:err.message}); }
}

// ============================================================
// FEE BALANCE
// ============================================================
function getFeeBalance(params) {
  try {
    const studentId = String(params.student_id||"").trim();
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    
    if (!studentId || !schoolCode) return output({success:false,message:"student_id and school_code required"});
    
    const student = sheetToObjects("STUDENTS").find(s=>matchCode(s.school_code,schoolCode) && String(s.student_id).trim()===studentId);
    if (!student) return output({success:false,message:"Student not found"});
    
    const studentClass = String(student.class).trim();
    const settings = getSchoolSettings(schoolCode);
    const term = settings.term || "";
    const session = settings.session || "";
    
    // Get all fees for this class
    const fees = sheetToObjects("FEES").filter(f=>
      matchCode(f.school_code,schoolCode) &&
      String(f.class).trim().toLowerCase() === studentClass.toLowerCase() &&
      (!term || String(f.term).trim() === term)
    );
    
    // Get all payments for this student
    const payments = sheetToObjects("FEE_PAYMENTS").filter(p=>
      matchCode(p.school_code,schoolCode) &&
      String(p.student_id).trim() === studentId &&
      (!term || String(p.term).trim() === term)
    );
    
    const totalFees = fees.reduce((sum,f) => sum + (parseFloat(f.amount)||0), 0);
    const totalPaid = payments.reduce((sum,p) => sum + (parseFloat(p.amount_paid)||0), 0);
    const balance = totalFees - totalPaid;
    
    // Get fee breakdown by priority
    const feeBreakdown = fees.map(f => ({
      fee_type: f.fee_type,
      amount: parseFloat(f.amount)||0,
      priority: parseInt(f.priority)||99,
      due_date: f.due_date,
      status: balance >= parseFloat(f.amount) ? "paid" : "pending"
    })).sort((a,b) => a.priority - b.priority);
    
    return output({
      success:true,
      student_id:studentId,
      student_class:studentClass,
      total_fees:totalFees,
      total_paid:totalPaid,
      balance:balance,
      fees:feeBreakdown,
      is_owing:balance > 0
    });
  } catch(err) { return output({success:false,message:err.message}); }
}

// ============================================================
// CLASS ARMS
// ============================================================
function getClassArms(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const className = params.class;
    
    if (!schoolCode) return output({success:false,message:"school_code required"});
    
    // Get all classes with arms
    const classes = sheetToObjects("CLASSES").filter(c=>matchCode(c.school_code,schoolCode));
    let arms = [];
    
    if (className) {
      // Get arms for specific class
      arms = classes.filter(c=>String(c.class_name).trim().toLowerCase() === className.toLowerCase())
        .map(c => ({
          class_name: c.class_name,
          arm: c.arm,
          class_level: c.class_level,
          capacity: c.capacity,
          status: c.status
        }));
    } else {
      // Get all unique class-arm combinations
      const seen = {};
      classes.forEach(c => {
        const key = String(c.class_name).trim() + "||" + String(c.arm||"").trim();
        if (!seen[key]) {
          seen[key] = true;
          arms.push({
            class_name: c.class_name,
            arm: c.arm,
            class_level: c.class_level,
            capacity: c.capacity,
            status: c.status
          });
        }
      });
    }
    
    return output({success:true,arms:arms,count:arms.length});
  } catch(err) { return output({success:false,message:err.message}); }
}

function addClassArm(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const className = String(params.class_name||"").trim();
    const arm = String(params.arm||"").trim();
    
    if (!schoolCode || !className || !arm) return output({success:false,message:"school_code, class_name and arm required"});
    
    ensureClassesSheet();
    
    // Check if class with this arm already exists
    const existing = sheetToObjects("CLASSES").filter(c=>
      matchCode(c.school_code,schoolCode) &&
      String(c.class_name).trim().toLowerCase() === className.toLowerCase() &&
      String(c.arm||"").trim().toLowerCase() === arm.toLowerCase()
    );
    
    if (existing.length > 0) return output({success:false,message:"Class arm already exists: " + className + " " + arm});
    
    getSheet("CLASSES").appendRow([schoolCode, className, params.class_level||"", arm, params.capacity||"", "active", today()]);
    
    logActivity(schoolCode, "admin", "admin", "addClassArm", "Added arm: " + className + " " + arm);
    return output({success:true,message:"Class arm added: " + className + " " + arm});
  } catch(err) { return output({success:false,message:err.message}); }
}

function deleteClassArm(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const className = String(params.class_name||"").trim();
    const arm = String(params.arm||"").trim();
    
    if (!schoolCode || !className || !arm) return output({success:false,message:"school_code, class_name and arm required"});
    
    const sheet = getSheet("CLASSES");
    const data = sheet.getDataRange().getValues();
    
    for (let i=1;i<data.length;i++) {
      if (matchCode(data[i][0],schoolCode) &&
          String(data[i][1]).trim().toLowerCase() === className.toLowerCase() &&
          String(data[i][3]||"").trim().toLowerCase() === arm.toLowerCase()) {
        sheet.deleteRow(i+1);
        logActivity(schoolCode, "admin", "admin", "deleteClassArm", "Deleted arm: " + className + " " + arm);
        return output({success:true,message:"Class arm deleted"});
      }
    }
    
    return output({success:false,message:"Class arm not found"});
  } catch(err) { return output({success:false,message:err.message}); }
}

// ============================================================
// FEES
// ============================================================
function getFees(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const classFilter = params.class;
    const term = params.term;
    const session = params.session;
    
    let sheet;
    try { sheet = getSheet("FEES"); } catch(e) { return output({success:true,fees:[]}); }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return output({success:true,fees:[]});
    
    const headers = data[0].map(h=>String(h).trim().toLowerCase().replace(/ /g,"_"));
    const scIdx=headers.indexOf("school_code"), clsIdx=headers.indexOf("class"),
          typeIdx=headers.indexOf("fee_type"), amtIdx=headers.indexOf("amount"),
          descIdx=headers.indexOf("description"), termIdx=headers.indexOf("term"),
          sessIdx=headers.indexOf("session"), dueIdx=headers.indexOf("due_date"),
          statIdx=headers.indexOf("status"), prioIdx=headers.indexOf("priority");
    
    let results = [];
    for (let i=1;i<data.length;i++) {
      if (!matchCode(data[i][scIdx],schoolCode)) continue;
      if (classFilter && String(data[i][clsIdx]).trim() !== classFilter) continue;
      if (term && String(data[i][termIdx]).trim() !== term) continue;
      if (session && String(data[i][sessIdx]).trim() !== session) continue;
      
      results.push({
        id: i,
        class: data[i][clsIdx],
        fee_type: data[i][typeIdx],
        amount: data[i][amtIdx],
        description: data[i][descIdx],
        term: data[i][termIdx],
        session: data[i][sessIdx],
        due_date: data[i][dueIdx],
        status: data[i][statIdx],
        priority: data[i][prioIdx] || 99
      });
    }
    // Sort by priority (lower number = higher priority)
    results.sort((a,b) => (parseInt(a.priority)||99) - (parseInt(b.priority)||99));
    return output({success:true,fees:results});
  } catch(err) { return output({success:false,message:err.message}); }
}

function addFee(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    
    const className = params.class;
    const feeType = params.fee_type;
    const amount = parseFloat(params.amount);
    const description = params.description || "";
    const term = params.term;
    const session = params.session;
    const dueDate = params.due_date || "";
    const priority = params.priority || "5";
    
    if (!className || !feeType || isNaN(amount)) {
      return output({success:false,message:"class, fee_type, amount required"});
    }
    
    let sheet;
    try { sheet = getSheet("FEES"); } catch(e) {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      sheet = ss.insertSheet("FEES");
      sheet.getRange(1,1,1,10).setValues([["school_code","class","fee_type","amount","description","term","session","due_date","status","priority"]]);
    }
    
    sheet.appendRow([schoolCode, className, feeType, amount, description, term, session, dueDate, "active", priority]);
    return output({success:true,message:"Fee added"});
  } catch(err) { return output({success:false,message:err.message}); }
}

function updateFee(params) {
  try {
    const rowNum = parseInt(params.row_num);
    const amount = parseFloat(params.amount);
    const status = params.status;
    
    if (!rowNum) return output({success:false,message:"row_num required"});
    
    const sheet = getSheet("FEES");
    if (!isNaN(amount)) sheet.getRange(rowNum, 4).setValue(amount);
    if (status) sheet.getRange(rowNum, 9).setValue(status);
    
    return output({success:true,message:"Fee updated"});
  } catch(err) { return output({success:false,message:err.message}); }
}

function deleteFee(params) {
  try {
    const rowNum = parseInt(params.row_num);
    if (!rowNum) return output({success:false,message:"row_num required"});
    
    const sheet = getSheet("FEES");
    sheet.deleteRow(rowNum);
    return output({success:true,message:"Fee deleted"});
  } catch(err) { return output({success:false,message:err.message}); }
}

function recordPayment(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    
    const studentId = params.student_id;
    const feeId = parseInt(params.fee_id);
    const amountPaid = parseFloat(params.amount_paid);
    const paymentDate = params.payment_date || new Date().toISOString().split("T")[0];
    const paymentMethod = params.payment_method || "cash";
    const recordedBy = params.recorded_by || "admin";
    const term = params.term;
    const session = params.session;
    
    if (!studentId || isNaN(amountPaid)) {
      return output({success:false,message:"student_id, fee_id, amount_paid required"});
    }
    
    let sheet;
    try { sheet = getSheet("FEE_PAYMENTS"); } catch(e) {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      sheet = ss.insertSheet("FEE_PAYMENTS");
      sheet.getRange(1,1,1,9).setValues([["school_code","student_id","fee_id","amount_paid","payment_date","payment_method","recorded_by","term","session"]]);
    }
    
    sheet.appendRow([schoolCode, studentId, feeId, amountPaid, paymentDate, paymentMethod, recordedBy, term, session]);
    return output({success:true,message:"Payment recorded"});
  } catch(err) { return output({success:false,message:err.message}); }
}

function getFeePayments(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const studentId = params.student_id;
    
    let sheet;
    try { sheet = getSheet("FEE_PAYMENTS"); } catch(e) { return output({success:true,payments:[]}); }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return output({success:true,payments:[]});
    
    const headers = data[0].map(h=>String(h).trim().toLowerCase().replace(/ /g,"_"));
    const scIdx=headers.indexOf("school_code"), stuIdx=headers.indexOf("student_id"),
          feeIdx=headers.indexOf("fee_id"), amtIdx=headers.indexOf("amount_paid"),
          dateIdx=headers.indexOf("payment_date"), methIdx=headers.indexOf("payment_method"),
          termIdx=headers.indexOf("term"), sessIdx=headers.indexOf("session");
    
    let results = [];
    for (let i=1;i<data.length;i++) {
      if (!matchCode(data[i][scIdx],schoolCode)) continue;
      if (studentId && String(data[i][stuIdx]).trim() !== studentId) continue;
      
      results.push({
        id: i,
        student_id: data[i][stuIdx],
        fee_id: data[i][feeIdx],
        amount_paid: data[i][amtIdx],
        payment_date: data[i][dateIdx],
        payment_method: data[i][methIdx],
        term: data[i][termIdx],
        session: data[i][sessIdx]
      });
    }
    return output({success:true,payments:results});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// LIBRARY SYSTEM
// ============================================================
function getLibraryBooks(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const category = params.category;
    let sheet;
    try { sheet = getSheet("LIBRARY"); } catch(e) { return output({success:true,books:[]}); }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return output({success:true,books:[]});
    const headers = data[0].map(h=>String(h).trim().toLowerCase().replace(/ /g,"_"));
    const scIdx=headers.indexOf("school_code"), idIdx=headers.indexOf("book_id"),
          titleIdx=headers.indexOf("title"), authIdx=headers.indexOf("author"),
          catIdx=headers.indexOf("category"), totalIdx=headers.indexOf("total_copies"),
          availIdx=headers.indexOf("available"), statIdx=headers.indexOf("status");
    let results = [];
    for (let i=1;i<data.length;i++) {
      if (!matchCode(data[i][scIdx],schoolCode)) continue;
      if (category && String(data[i][catIdx]).trim() !== category) continue;
      results.push({ id:i, book_id:data[i][idIdx], title:data[i][titleIdx], author:data[i][authIdx], category:data[i][catIdx], total_copies:data[i][totalIdx], available:data[i][availIdx], status:data[i][statIdx] });
    }
    return output({success:true,books:results});
  } catch(err) { return output({success:false,message:err.message}); }
}
function addLibraryBook(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const title = String(params.title||"").trim();
    const author = String(params.author||"").trim();
    const isbn = String(params.isbn||"").trim();
    const category = String(params.category||"").trim();
    const totalCopies = parseInt(params.total_copies) || 1;
    if (!schoolCode||!title) return output({success:false,message:"School code and title required"});
    const bookId = "BK" + schoolCode.substring(0,3) + String(Math.floor(1000 + Math.random() * 9000));
    getSheet("LIBRARY").appendRow([schoolCode,bookId,title,author,isbn,category,totalCopies,totalCopies,"active",today()]);
    return output({success:true,book_id:bookId,message:"Book added successfully"});
  } catch(err) { return output({success:false,message:err.message}); }
}
function getLibraryLoans(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const studentId = params.student_id;
    let sheet;
    try { sheet = getSheet("LIBRARY_LOANS"); } catch(e) { return output({success:true,loans:[]}); }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return output({success:true,loans:[]});
    const headers = data[0].map(h=>String(h).trim().toLowerCase().replace(/ /g,"_"));
    const scIdx=headers.indexOf("school_code"), lidIdx=headers.indexOf("loan_id"),
          bidIdx=headers.indexOf("book_id"), stuIdx=headers.indexOf("student_id"),
          bDateIdx=headers.indexOf("borrow_date"), dDateIdx=headers.indexOf("due_date"),
          rDateIdx=headers.indexOf("return_date"), statIdx=headers.indexOf("status");
    let results = [];
    for (let i=1;i<data.length;i++) {
      if (!matchCode(data[i][scIdx],schoolCode)) continue;
      if (studentId && String(data[i][stuIdx]).trim() !== studentId) continue;
      results.push({ id:i, loan_id:data[i][lidIdx], book_id:data[i][bidIdx], student_id:data[i][stuIdx], borrow_date:data[i][bDateIdx], due_date:data[i][dDateIdx], return_date:data[i][rDateIdx], status:data[i][statIdx] });
    }
    return output({success:true,loans:results});
  } catch(err) { return output({success:false,message:err.message}); }
}
function borrowBook(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const bookId = String(params.book_id||"").trim();
    const studentId = String(params.student_id||"").trim();
    if (!schoolCode||!bookId||!studentId) return output({success:false,message:"All fields required"});
    const books = sheetToObjects("LIBRARY").filter(b=>matchCode(b.school_code,schoolCode) && String(b.book_id).trim()===bookId);
    if (!books.length) return output({success:false,message:"Book not found"});
    const book = books[0];
    if (parseInt(book.available||0) <= 0) return output({success:false,message:"No copies available"});
    const loanId = "LN" + schoolCode.substring(0,3) + String(Math.floor(10000 + Math.random() * 90000));
    const dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + 14);
    getSheet("LIBRARY_LOANS").appendRow([schoolCode,loanId,bookId,studentId,today(),dueDate.toISOString().split("T")[0],"","borrowed",S.admin?S.admin.username:"admin"]);
    return output({success:true,loan_id:loanId,message:"Book borrowed successfully"});
  } catch(err) { return output({success:false,message:err.message}); }
}
function returnBook(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const loanId = String(params.loan_id||"").trim();
    if (!schoolCode||!loanId) return output({success:false,message:"Loan ID required"});
    const loans = sheetToObjects("LIBRARY_LOANS").filter(l=>matchCode(l.school_code,schoolCode) && String(l.loan_id).trim()===loanId);
    if (!loans.length) return output({success:false,message:"Loan not found"});
    const loan = loans[0];
    const sheet = getSheet("LIBRARY_LOANS");
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h=>String(h).trim().toLowerCase().replace(/ /g,"_"));
    const lidIdx = headers.indexOf("loan_id");
    for (let i=1;i<data.length;i++) {
      if (String(data[i][lidIdx]).trim() === loanId) {
        sheet.getRange(i+1, headers.indexOf("return_date")+1).setValue(today());
        sheet.getRange(i+1, headers.indexOf("status")+1).setValue("returned");
        break;
      }
    }
    return output({success:true,message:"Book returned successfully"});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// TIMETABLE SYSTEM
// ============================================================
function getTimetable(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const className = params.class;
    const day = params.day;
    const term = params.term;
    const session = params.session;
    let sheet;
    try { sheet = getSheet("TIMETABLE"); } catch(e) { return output({success:true,timetable:[]}); }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return output({success:true,timetable:[]});
    const headers = data[0].map(h=>String(h).trim().toLowerCase().replace(/ /g,"_"));
    const scIdx=headers.indexOf("school_code"), clsIdx=headers.indexOf("class"),
          dayIdx=headers.indexOf("day"), perIdx=headers.indexOf("period"),
          subIdx=headers.indexOf("subject_code"), teaIdx=headers.indexOf("teacher_id"),
          roomIdx=headers.indexOf("room"), termIdx=headers.indexOf("term"), sessIdx=headers.indexOf("session");
    let results = [];
    for (let i=1;i<data.length;i++) {
      if (!matchCode(data[i][scIdx],schoolCode)) continue;
      if (className && String(data[i][clsIdx]).trim() !== className) continue;
      if (day && String(data[i][dayIdx]).trim().toLowerCase() !== day.toLowerCase()) continue;
      if (term && String(data[i][termIdx]).trim() !== term) continue;
      if (session && String(data[i][sessIdx]).trim() !== session) continue;
      results.push({ id:i, class:data[i][clsIdx], day:data[i][dayIdx], period:data[i][perIdx], subject_code:data[i][subIdx], teacher_id:data[i][teaIdx], room:data[i][roomIdx], term:data[i][termIdx], session:data[i][sessIdx] });
    }
    return output({success:true,timetable:results});
  } catch(err) { return output({success:false,message:err.message}); }
}
function addTimetable(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const className = String(params.class||"").trim();
    const day = String(params.day||"").trim();
    const period = String(params.period||"").trim();
    const subjectCode = String(params.subject_code||"").trim();
    const teacherId = String(params.teacher_id||"").trim();
    const room = String(params.room||"").trim();
    const term = String(params.term||"").trim();
    const session = String(params.session||"").trim();
    if (!schoolCode||!className||!day||!period) return output({success:false,message:"Class, day and period required"});
    getSheet("TIMETABLE").appendRow([schoolCode,className,day,period,subjectCode,teacherId,room,term,session]);
    return output({success:true,message:"Timetable entry added"});
  } catch(err) { return output({success:false,message:err.message}); }
}

// ============================================================
// RESULT IDs
// ============================================================
function generateResultId(params) {
  try {
    const schoolCode=String(params.school_code||"").trim().toUpperCase();
    const studentId =String(params.student_id ||"").trim();
    if (!schoolCode||!studentId) return output({success:false,message:"school_code and student_id required"});
    const student=sheetToObjects("STUDENTS").find(s=>String(s.student_id).trim()===studentId&&matchCode(s.school_code,schoolCode));
    if (!student) return output({success:false,message:"Student not found"});
    let term="",session="";
    try { const s=getSchoolSettings(schoolCode); term=s.term||""; session=s.session||""; } catch(e){}
    const resultId=`RID-${schoolCode}-${randomSuffix(4)}-${randomSuffix(4)}`;
    getSheet("RESULT_IDS").appendRow([schoolCode,resultId,studentId,0,5,term,session,"active",today()]);
    return output({success:true,result_id:resultId,student_id:studentId,term:term,session:session});
  } catch(err) { return output({success:false,message:err.message}); }
}
function validateResultId(params) {
  try {
    const resultId=String(params.result_id||"").trim().toUpperCase();
    if (!resultId) return output({success:false,message:"result_id required"});
    const sheet=getSheet("RESULT_IDS"),data=sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++){
      if (String(data[i][1]).trim().toUpperCase()!==resultId) continue;
      const used=parseInt(data[i][3])||0, max=parseInt(data[i][4])||5;
      const status=String(data[i][7]).trim().toLowerCase();
      if (status==="expired"||used>=max){sheet.getRange(i+1,8).setValue("expired");return output({success:false,message:"Result ID expired"});}
      const newCount=used+1;
      sheet.getRange(i+1,4).setValue(newCount);
      if (newCount>=max) sheet.getRange(i+1,8).setValue("expired");
      return output({success:true,student_id:String(data[i][2]).trim(),school_code:String(data[i][0]).trim().toUpperCase(),
        attempts_used:newCount,attempts_remaining:max-newCount,
        term:String(data[i][5]).trim(),session:String(data[i][6]).trim()});
    }
    return output({success:false,message:"Result ID not found"});
  } catch(err) { return output({success:false,message:err.message}); }
}
function getResultIds(params) {
  try {
    const schoolCode=String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    const ids=sheetToObjects("RESULT_IDS").filter(r=>matchCode(r.school_code,schoolCode));
    return output({success:true,result_ids:ids,count:ids.length});
  } catch(err) { return output({success:false,message:err.message}); }
}
function deleteResultId(params) {
  try {
    const schoolCode=String(params.school_code||"").trim().toUpperCase();
    const resultId  =String(params.result_id  ||"").trim().toUpperCase();
    if (!schoolCode||!resultId) return output({success:false,message:"school_code and result_id required"});
    const sheet=getSheet("RESULT_IDS"),data=sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++){
      if (matchCode(data[i][0],schoolCode)&&String(data[i][1]).trim().toUpperCase()===resultId){
        sheet.deleteRow(i+1);
        logActivity(schoolCode,"admin","admin","deleteResultId","Deleted RID: "+resultId);
        return output({success:true,message:"Result ID deleted"});
      }
    }
    return output({success:false,message:"Result ID not found"});
  } catch(err) { return output({success:false,message:err.message}); }
}

// ============================================================
// BULK DELETE RESULT IDs
// ============================================================
function deleteBulkResultIds(params) {
  try {
    const schoolCode=String(params.school_code||"").trim().toUpperCase();
    const resultIds =String(params.result_ids||"").trim();
    if (!schoolCode||!resultIds) return output({success:false,message:"school_code and result_ids required"});
    const idsToDelete = resultIds.split(',').map(id=>String(id).trim().toUpperCase()).filter(Boolean);
    if (idsToDelete.length===0) return output({success:false,message:"No valid result IDs provided"});
    const sheet=getSheet("RESULT_IDS"),data=sheet.getDataRange().getValues();
    const rowsToDelete = [];
    for (let i=1;i<data.length;i++){
      if (matchCode(data[i][0],schoolCode)&&idsToDelete.includes(String(data[i][1]).trim().toUpperCase())){
        rowsToDelete.push(i+1);
      }
    }
    if (rowsToDelete.length===0) return output({success:false,message:"No matching Result IDs found"});
    // Delete from bottom up to preserve row indices
    rowsToDelete.sort((a,b)=>b-a).forEach(r=>sheet.deleteRow(r));
    logActivity(schoolCode,"admin","admin","deleteBulkResultIds","Deleted "+rowsToDelete.length+" RIDs");
    return output({success:true,message:"Deleted "+rowsToDelete.length+" Result ID(s)",deleted:rowsToDelete.length});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// SCHOOLS
// ============================================================
function getSchools(params) {
  try {
    const schools=sheetToObjects("SCHOOLS").map(s=>{const c=Object.assign({},s);delete c.admin_password;delete c.staff_password;return c;});
    return output({success:true,schools:schools,count:schools.length});
  } catch(err) { return output({success:false,message:err.message}); }
}
function addSchool(params) {
  try {
    const schoolCode=String(params.school_code   ||"").trim().toUpperCase();
    const schoolName=String(params.school_name   ||"").trim();
    const address   =String(params.address       ||"").trim();
    const town      =String(params.town          ||"").trim();
    const state     =String(params.state         ||"").trim();
    const adminUser =String(params.admin_username||"admin").trim();
    const adminPass =String(params.admin_password||generatePin()).trim();
    const staffPass =String(params.staff_password||generatePin()).trim();
    const email     =String(params.contact_email ||"").trim();
    const phone     =String(params.contact_phone ||"").trim();
    if (!schoolCode||!schoolName) return output({success:false,message:"school_code and school_name required"});
    if (schoolCode.length>8) return output({success:false,message:"School code max 8 chars"});
    if (sheetToObjects("SCHOOLS").find(s=>matchCode(s.school_code,schoolCode)))
      return output({success:false,message:"School code exists: "+schoolCode});
    getSheet("SCHOOLS").appendRow([schoolCode,schoolName,address,town,state,adminUser,adminPass,staffPass,"no",email,phone,"","active",today()]);
    getSheet("SCHOOL_SETTINGS").appendRow([schoolCode,schoolName,address,town,state,"","First Term","",20,20,60,"yes","yes","","",today(),"grade_only","class"]);
    return output({success:true,message:"School registered",school_code:schoolCode,admin_username:adminUser,admin_password:adminPass,staff_password:staffPass});
  } catch(err) { return output({success:false,message:err.message}); }
}
function updateSchool(params) {
  try {
    const schoolCode=String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    const sheet=getSheet("SCHOOLS"),headers=getHeaders("SCHOOLS"),data=sheet.getDataRange().getValues();
    let rowIndex=-1;
    for (let i=1;i<data.length;i++){if(matchCode(data[i][0],schoolCode)){rowIndex=i+1;break;}}
    if (rowIndex===-1) return output({success:false,message:"School not found"});
    ["school_name","address","town","state","admin_username","admin_password","staff_password","contact_email","contact_phone","logo_url","status"].forEach(f=>{
      if(params[f]!==undefined){const c=headers.indexOf(f);if(c!==-1)sheet.getRange(rowIndex,c+1).setValue(params[f]);}
    });
    return output({success:true,message:"School updated"});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// DASHBOARD
// ============================================================
function getDashboard(params) {
  try {
    const schoolCode=String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    const students =sheetToObjects("STUDENTS") .filter(s=>matchCode(s.school_code,schoolCode));
    const staff    =sheetToObjects("STAFF")    .filter(s=>matchCode(s.school_code,schoolCode));
    const subjects =sheetToObjects("SUBJECTS") .filter(s=>matchCode(s.school_code,schoolCode));
    const questions=sheetToObjects("QUESTIONS").filter(q=>matchCode(q.school_code,schoolCode));
    const results  =sheetToObjects("RESULTS")  .filter(r=>matchCode(r.school_code,schoolCode));
    const resultIds=sheetToObjects("RESULT_IDS").filter(r=>matchCode(r.school_code,schoolCode));
    const activeStudents=[...new Set(results.map(r=>r.student_id))].length;
    const allTotals=results.map(r=>calcTotal(r.first_test,r.second_test,r.exam));
    const overallAvg=allTotals.length>0?Math.round(allTotals.reduce((a,b)=>a+b,0)/allTotals.length*10)/10:0;
    const classes=[...new Set(students.map(s=>s.class))].filter(Boolean);
    // Enhanced analytics
    const genderStats={male:0,female:0,other:0};
    students.forEach(s=>{const g=String(s.gender||"").toLowerCase();if(g==="male")genderStats.male++;else if(g==="female")genderStats.female++;else genderStats.other++;});
    const subjectStats={};
    results.forEach(r=>{const sc=String(r.subject_code).trim();if(!subjectStats[sc])subjectStats[sc]={total:0,count:0};subjectStats[sc].total+=calcTotal(r.first_test,r.second_test,r.exam);subjectStats[sc].count++;});
    const subjectAverages=Object.keys(subjectStats).map(sc=>({subject_code:sc,avg:Math.round(subjectStats[sc].total/subjectStats[sc].count*10)/10,count:subjectStats[sc].count})).sort((a,b)=>b.avg-a.avg);
    const gradeDistribution={A1:0,B2:0,B3:0,C4:0,C5:0,C6:0,D7:0,E8:0,F9:0};
    results.forEach(r=>{const g=getGrade(calcTotal(r.first_test,r.second_test,r.exam));gradeDistribution[g.grade]=(gradeDistribution[g.grade]||0)+1;});
    const classPerformance=classes.map(cls=>{const clsStudents=students.filter(s=>String(s.class).trim().toLowerCase()===cls.toLowerCase());const clsResults=results.filter(r=>clsStudents.some(s=>String(s.student_id).trim()===String(r.student_id).trim()));const totals=clsResults.map(r=>calcTotal(r.first_test,r.second_test,r.exam));const avg=totals.length>0?Math.round(totals.reduce((a,b)=>a+b,0)/totals.length*10)/10:0;return{class:cls,students:clsStudents.length,results:clsResults.length,average:avg};});
    return output({success:true,dashboard:{
      total_students:   students.length,
      active_students:  students.filter(s=>String(s.status).trim().toLowerCase()==="active").length,
      total_staff:      staff.length,
      total_subjects:   subjects.length,
      total_questions:  questions.length,
      total_results:    results.length,
      exams_submitted:  activeStudents,
      result_ids_issued:resultIds.length,
      overall_average:  overallAvg,
      classes:          classes,
      total_classes:    classes.length,
      gender_stats:     genderStats,
      subject_averages: subjectAverages.slice(0,10),
      grade_distribution:gradeDistribution,
      class_performance: classPerformance
    }});
  } catch(err) { return output({success:false,message:err.message}); }
}
// ============================================================
// ANNOUNCEMENTS — UPDATE
// ============================================================
function updateAnnouncement(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const rowNum = parseInt(params.row_num);
    if (!schoolCode || !rowNum) return output({success:false,message:"school_code and row_num required"});
    const sheet = getSheet("ANNOUNCEMENTS");
    const headers = getHeaders("ANNOUNCEMENTS");
    const updatable = ["title","content","priority","target_role","target_class","start_date","end_date"];
    updatable.forEach(field => {
      if (params[field] !== undefined) {
        const c = headers.indexOf(field);
        if (c !== -1) sheet.getRange(rowNum, c+1).setValue(params[field]);
      }
    });
    logActivity(schoolCode, "admin", "admin", "updateAnnouncement", "Updated announcement row " + rowNum);
    return output({success:true, message:"Announcement updated"});
  } catch(err) { return output({success:false, message:err.message}); }
}

// ============================================================
// WEBSITE CONFIG (for School Website Builder)
// ============================================================
function saveWebsiteConfig(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const configJson = String(params.config_json||"").trim();
    const published = String(params.published||"no").trim();
    if (!schoolCode) return output({success:false, message:"school_code required"});
    if (!configJson) return output({success:false, message:"config_json required"});
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("WEBSITE_CONFIG");
    if (!sheet) {
      sheet = ss.insertSheet("WEBSITE_CONFIG");
      sheet.getRange(1,1,1,5).setValues([["school_code","config_json","published","published_at","updated_at"]]);
    }
    
    const data = sheet.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (matchCode(data[i][0], schoolCode)) { rowIdx = i + 1; break; }
    }
    
    const now = new Date().toISOString();
    const publishedAt = published === "yes" ? now : (rowIdx > 0 ? data[rowIdx-1][3] : "");
    const row = [schoolCode, configJson, published, publishedAt, now];
    
    if (rowIdx > 0) {
      sheet.getRange(rowIdx, 1, 1, 5).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    
    logActivity(schoolCode, "admin", "admin", "saveWebsiteConfig", "Website config saved");
    return output({success:true, message:"Website config saved", published: published});
  } catch(err) { return output({success:false, message:err.message}); }
}

function getWebsiteConfig(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false, message:"school_code required"});
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("WEBSITE_CONFIG");
    if (!sheet) return output({success:false, message:"No website config found"});
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (matchCode(data[i][0], schoolCode)) {
        return output({
          success: true,
          config: data[i][1],
          published: data[i][2],
          published_at: data[i][3],
          updated_at: data[i][4]
        });
      }
    }
    return output({success:false, message:"No config found for school: " + schoolCode});
  } catch(err) { return output({success:false, message:err.message}); }
}

// ============================================================
// FEE REPORT — Student-level fee status
// ============================================================
function getStudentFeeReport(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const cls = String(params.class||"").trim();
    const term = String(params.term||"").trim();
    const session = String(params.session||"").trim();
    
    if (!schoolCode) return output({success:false, message:"school_code required"});
    
    // Get all fees for filter
    let fees = sheetToObjects("FEES").filter(f => matchCode(f.school_code, schoolCode));
    if (cls) fees = fees.filter(f => String(f.class).trim().toLowerCase() === cls.toLowerCase());
    if (term) fees = fees.filter(f => String(f.term).trim() === term);
    if (session) fees = fees.filter(f => String(f.session).trim() === session);
    
    // Get students for the class
    let students = sheetToObjects("STUDENTS").filter(s =>
      matchCode(s.school_code, schoolCode) &&
      String(s.status).trim().toLowerCase() === "active"
    );
    if (cls) students = students.filter(s => String(s.class).trim().toLowerCase() === cls.toLowerCase());
    
    // Get all payments
    const payments = sheetToObjects("FEE_PAYMENTS").filter(p =>
      matchCode(p.school_code, schoolCode) &&
      (!term || String(p.term).trim() === term) &&
      (!session || String(p.session).trim() === session)
    );
    
    const totalFeeAmount = fees.reduce((sum, f) => sum + (parseFloat(f.amount)||0), 0);
    
    const report = students.map(s => {
      const stuPayments = payments.filter(p => String(p.student_id).trim() === String(s.student_id).trim());
      const totalPaid = stuPayments.reduce((sum, p) => sum + (parseFloat(p.amount_paid)||0), 0);
      const balance = totalFeeAmount - totalPaid;
      return {
        student_id: s.student_id,
        full_name: s.full_name,
        class: s.class,
        total_fees: totalFeeAmount,
        total_paid: totalPaid,
        balance: balance,
        status: balance <= 0 ? "cleared" : totalPaid > 0 ? "partial" : "owing",
        payments: stuPayments
      };
    });
    
    return output({success:true, report:report, fees:fees, total_fee_amount:totalFeeAmount});
  } catch(err) { return output({success:false, message:err.message}); }
}

function deleteFeePayment(params) {
  try {
    const rowNum = parseInt(params.row_num);
    if (!rowNum) return output({success:false, message:"row_num required"});
    const sheet = getSheet("FEE_PAYMENTS");
    sheet.deleteRow(rowNum);
    return output({success:true, message:"Payment deleted"});
  } catch(err) { return output({success:false, message:err.message}); }
}

// ============================================================
// APEX-STYLE FINANCE — priority allocation, ledger, waivers, rollover
// Sheets used (auto-created): FEES, FEE_PAYMENTS, FEE_LEDGER, FEE_WAIVERS, FEE_OLD_DEBT
// Allocation priority (highest -> lowest):
//   1. Old Debt (previous term)
//   2. Graduation / Party fees
//   3. Lesson fees
//   4. Current School Fees (everything else)
// ============================================================
function _ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return sh;
}
function _feePriorityBucket_(feeType) {
  const t = String(feeType||"").trim().toLowerCase();
  if (/grad|party/.test(t)) return 2;
  if (/lesson|extra|coaching/.test(t)) return 3;
  return 4; // school fee / tuition / default
}
function _bucketLabel_(b){ return ({1:"Old Debt",2:"Graduation/Party",3:"Lesson",4:"School Fee"})[b]||"Other"; }

function getStudentLedger(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const studentId  = String(params.student_id ||"").trim();
    if (!schoolCode || !studentId) return output({success:false,message:"school_code and student_id required"});
    const stu = sheetToObjects("STUDENTS").find(s=>matchCode(s.school_code,schoolCode) && String(s.student_id).trim()===studentId);
    if (!stu) return output({success:false,message:"Student not found"});

    const cls = String(stu.class||"").trim();
    const fees = sheetToObjects("FEES").filter(f=>matchCode(f.school_code,schoolCode) && String(f.class).trim().toLowerCase()===cls.toLowerCase());
    const payments = sheetToObjects("FEE_PAYMENTS").filter(p=>matchCode(p.school_code,schoolCode) && String(p.student_id).trim()===studentId);
    let waivers = [];
    try { waivers = sheetToObjects("FEE_WAIVERS").filter(w=>matchCode(w.school_code,schoolCode) && String(w.student_id).trim()===studentId); } catch(e){}
    let oldDebt = 0;
    try {
      const od = sheetToObjects("FEE_OLD_DEBT").find(o=>matchCode(o.school_code,schoolCode) && String(o.student_id).trim()===studentId);
      if (od) oldDebt = parseFloat(od.amount)||0;
    } catch(e){}

    // Build per-bucket totals
    const buckets = {1:{label:"Old Debt",charged:oldDebt,paid:0},2:{label:"Graduation/Party",charged:0,paid:0},3:{label:"Lesson",charged:0,paid:0},4:{label:"School Fee",charged:0,paid:0}};
    fees.forEach(f=>{ const b=_feePriorityBucket_(f.fee_type); buckets[b].charged += parseFloat(f.amount)||0; });
    waivers.forEach(w=>{ const b=parseInt(w.bucket)||4; buckets[b].charged -= parseFloat(w.amount)||0; });

    let ledger = [];
    try { ledger = sheetToObjects("FEE_LEDGER").filter(l=>matchCode(l.school_code,schoolCode) && String(l.student_id).trim()===studentId); } catch(e){}
    // Fallback: rebuild allocations from payments if ledger empty
    if (!ledger.length && payments.length) {
      ledger = payments.map(p=>({date:p.payment_date, type:"PAYMENT", method:p.payment_method, amount:p.amount_paid, allocation:"(unallocated)", note:""}));
    }

    payments.forEach(p=>{
      // Crude paid distribution if no ledger detail: by priority order
      // (skipped if ledger has explicit splits)
    });

    // Compute paid per bucket from ledger if it has bucket col
    ledger.forEach(l=>{
      const alloc = String(l.allocation||"");
      // allocation format: "1:5000|3:2000|4:3000"
      alloc.split("|").forEach(seg=>{
        const m = seg.match(/^(\d+):(-?\d+(\.\d+)?)$/);
        if (m) { const b=parseInt(m[1]); if (buckets[b]) buckets[b].paid += parseFloat(m[2]); }
      });
    });

    const totalCharged = Object.values(buckets).reduce((s,b)=>s+Math.max(0,b.charged),0);
    const totalPaid = payments.reduce((s,p)=>s+(parseFloat(p.amount_paid)||0),0);
    const balance = totalCharged - totalPaid;

    return output({
      success:true,
      student:{student_id:stu.student_id,full_name:stu.full_name,class:stu.class,arm:stu.arm||""},
      buckets:buckets, fees:fees, payments:payments, waivers:waivers, ledger:ledger,
      old_debt:oldDebt, total_charged:totalCharged, total_paid:totalPaid, balance:balance,
      status: balance<=0 ? "cleared" : (totalPaid>0 ? "partial" : "owing")
    });
  } catch(err){ return output({success:false,message:err.message}); }
}

function recordPaymentAllocated(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const studentId  = String(params.student_id ||"").trim();
    const amount     = parseFloat(params.amount_paid);
    const method     = params.payment_method || "Cash";
    const date       = params.payment_date || new Date().toISOString().split("T")[0];
    const recordedBy = params.recorded_by || "admin";
    const term       = params.term || "";
    const session    = params.session || "";
    if (!schoolCode || !studentId || isNaN(amount) || amount<=0)
      return output({success:false,message:"school_code, student_id and positive amount_paid required"});

    // Pull current ledger to compute outstanding per bucket
    const lr = JSON.parse(getStudentLedger({school_code:schoolCode,student_id:studentId}).getContent());
    if (!lr.success) return output(lr);
    const out = {1:Math.max(0,(lr.buckets[1].charged||0)-(lr.buckets[1].paid||0)),
                 2:Math.max(0,(lr.buckets[2].charged||0)-(lr.buckets[2].paid||0)),
                 3:Math.max(0,(lr.buckets[3].charged||0)-(lr.buckets[3].paid||0)),
                 4:Math.max(0,(lr.buckets[4].charged||0)-(lr.buckets[4].paid||0))};

    let remaining = amount;
    const splits = {};
    [1,2,3,4].forEach(b=>{
      if (remaining<=0) return;
      const take = Math.min(out[b], remaining);
      if (take>0) { splits[b]=take; remaining -= take; }
    });
    // Any overflow → credit toward future School Fee bucket (4) as advance
    if (remaining>0) { splits[4] = (splits[4]||0) + remaining; remaining = 0; }

    const allocStr = Object.keys(splits).map(b=>b+":"+splits[b]).join("|");
    const allocReadable = Object.keys(splits).map(b=>_bucketLabel_(parseInt(b))+" ₦"+Number(splits[b]).toLocaleString()).join(", ");

    // Append to FEE_PAYMENTS
    const paySheet = _ensureSheet_("FEE_PAYMENTS",["school_code","student_id","fee_id","amount_paid","payment_date","payment_method","recorded_by","term","session","balance_after"]);
    const newBalance = (lr.balance||0) - amount;
    paySheet.appendRow([schoolCode,studentId,"AUTO",amount,date,method,recordedBy,term,session,newBalance]);

    // Append to FEE_LEDGER
    const ledSheet = _ensureSheet_("FEE_LEDGER",["school_code","student_id","date","type","method","amount","allocation","balance_after","recorded_by","note"]);
    ledSheet.appendRow([schoolCode,studentId,date,"PAYMENT",method,amount,allocStr,newBalance,recordedBy,allocReadable]);

    // Pay down old debt sheet if bucket 1 received any
    if (splits[1]) {
      try {
        const sh = _ensureSheet_("FEE_OLD_DEBT",["school_code","student_id","amount","term","session","date_set"]);
        const data = sh.getDataRange().getValues();
        for (let i=1;i<data.length;i++) {
          if (matchCode(data[i][0],schoolCode) && String(data[i][1]).trim()===studentId) {
            const cur = parseFloat(data[i][2])||0;
            sh.getRange(i+1,3).setValue(Math.max(0,cur - splits[1]));
            break;
          }
        }
      } catch(e){}
    }

    return output({
      success:true, message:"Payment recorded",
      receipt:{
        student_id:studentId, full_name:lr.student.full_name, class:lr.student.class,
        amount:amount, method:method, date:date,
        allocation:splits, allocation_readable:allocReadable,
        balance_before:lr.balance, balance_after:newBalance,
        ref:"RCT-"+Utilities.formatDate(new Date(),"GMT","yyyyMMddHHmmss")
      }
    });
  } catch(err){ return output({success:false,message:err.message}); }
}

function applyFeeWaiver(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const studentId  = String(params.student_id ||"").trim();
    const bucket     = parseInt(params.bucket); // 1..4
    const amount     = parseFloat(params.amount);
    const reason     = String(params.reason||"").trim();
    const by         = params.recorded_by || "admin";
    if (!schoolCode||!studentId||!bucket||isNaN(amount)||amount<=0||!reason)
      return output({success:false,message:"school_code, student_id, bucket, amount, reason required"});
    const sh = _ensureSheet_("FEE_WAIVERS",["school_code","student_id","bucket","amount","reason","recorded_by","date"]);
    sh.appendRow([schoolCode,studentId,bucket,amount,reason,by,new Date().toISOString().split("T")[0]]);
    // also log into ledger
    const ledSheet = _ensureSheet_("FEE_LEDGER",["school_code","student_id","date","type","method","amount","allocation","balance_after","recorded_by","note"]);
    ledSheet.appendRow([schoolCode,studentId,new Date().toISOString().split("T")[0],"WAIVER","",amount,bucket+":"+amount,"",by,reason]);
    return output({success:true,message:"Waiver applied"});
  } catch(err){ return output({success:false,message:err.message}); }
}

function rolloverTerm(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    const newTerm    = String(params.new_term||"").trim();
    const newSession = String(params.new_session||"").trim();
    if (!schoolCode) return output({success:false,message:"school_code required"});

    const students = sheetToObjects("STUDENTS").filter(s=>matchCode(s.school_code,schoolCode) && String(s.status||"active").toLowerCase()==="active");
    const oldDebtSheet = _ensureSheet_("FEE_OLD_DEBT",["school_code","student_id","amount","term","session","date_set"]);
    const ledSheet = _ensureSheet_("FEE_LEDGER",["school_code","student_id","date","type","method","amount","allocation","balance_after","recorded_by","note"]);

    let processed=0, totalRolled=0;
    students.forEach(stu=>{
      const lr = JSON.parse(getStudentLedger({school_code:schoolCode,student_id:stu.student_id}).getContent());
      if (!lr.success) return;
      const bal = lr.balance||0;
      if (bal>0) {
        // upsert old_debt
        const data = oldDebtSheet.getDataRange().getValues();
        let found=false;
        for (let i=1;i<data.length;i++) {
          if (matchCode(data[i][0],schoolCode) && String(data[i][1]).trim()===stu.student_id) {
            oldDebtSheet.getRange(i+1,3).setValue((parseFloat(data[i][2])||0)+bal);
            oldDebtSheet.getRange(i+1,4).setValue(newTerm);
            oldDebtSheet.getRange(i+1,5).setValue(newSession);
            oldDebtSheet.getRange(i+1,6).setValue(new Date().toISOString().split("T")[0]);
            found=true; break;
          }
        }
        if (!found) oldDebtSheet.appendRow([schoolCode,stu.student_id,bal,newTerm,newSession,new Date().toISOString().split("T")[0]]);
        ledSheet.appendRow([schoolCode,stu.student_id,new Date().toISOString().split("T")[0],"ROLLOVER","",bal,"1:"+bal,"","admin","Carried forward to "+newTerm]);
        totalRolled += bal;
      }
      processed++;
    });
    return output({success:true, message:"Rollover complete", processed:processed, total_carried:totalRolled});
  } catch(err){ return output({success:false,message:err.message}); }
}

function getFinanceDashboard(params) {
  try {
    const schoolCode = String(params.school_code||"").trim().toUpperCase();
    if (!schoolCode) return output({success:false,message:"school_code required"});
    const fees = sheetToObjects("FEES").filter(f=>matchCode(f.school_code,schoolCode));
    const payments = sheetToObjects("FEE_PAYMENTS").filter(p=>matchCode(p.school_code,schoolCode));
    const totalExpected = fees.reduce((s,f)=>s+(parseFloat(f.amount)||0),0);
    const totalCollected = payments.reduce((s,p)=>s+(parseFloat(p.amount_paid)||0),0);
    // recent ledger
    let ledger = [];
    try { ledger = sheetToObjects("FEE_LEDGER").filter(l=>matchCode(l.school_code,schoolCode)); } catch(e){}
    ledger.sort((a,b)=>String(b.date).localeCompare(String(a.date)));
    return output({
      success:true,
      total_expected:totalExpected,
      total_collected:totalCollected,
      outstanding: totalExpected - totalCollected,
      collection_rate: totalExpected>0 ? Math.round((totalCollected/totalExpected)*1000)/10 : 0,
      recent: ledger.slice(0,10),
      fee_count: fees.length, payment_count: payments.length
    });
  } catch(err){ return output({success:false,message:err.message}); }
}
