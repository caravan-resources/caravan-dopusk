// ============================================================
// Google Apps Script v2 — Caravan Resources
// getTests() с протяжкой testId по пустым строкам блока (fix)
// ============================================================

const SHEET_ID       = "1nyyPGV0f7-7rio7LgD5E-MhTGEjn_QL8Fks4cfdy3P4";
const SHEET_EMP      = "Сотрудники";
const SHEET_RESULTS  = "Результаты";
const SHEET_TESTS    = "Тесты";
const SHEET_TRAINING = "Обучение";
const SHEET_TRAINING_PLAN = "ПланОбучения";
const SHEET_COURSE_CATALOG = "КаталогКурсов";
const SHEET_REQUIREMENTS = "Требования";
const SHEET_DOCUMENTS = "Документы";
const SHEET_PERSONNEL_EVENTS = "Личное дело";
const PHOTO_FOLDER_ID = "1rX0jetKQqm_Lsym8JX6YE8d4Js-Ye4y_";

// Табельные номера: префикс закреплён за основным ТОО сотрудника.
// SRC — ТОО "Saryarka Resources Capital", КУ — ТОО "Караван Улытау",
// CL — ТОО "Караван Логистик". Работники подрядных организаций получают
// префикс ПОД — их анкета приходит с полем too вида "Подрядчик: <название>",
// чтобы их можно было отличать и фильтровать по табельному номеру.
// Если передано что-то другое — общий "СОТР".
const TOO_PREFIXES = { "SRC": "SRC", "КУ": "КУ", "CL": "CL" };
const CONTRACTOR_PREFIX = "ПОД";

const SHEET_EQUIPMENT = "Техника";
const SHEET_CHECKLIST_TEMPLATES = "Чек-листы-Шаблоны";
const SHEET_CHECKLIST_RECORDS = "Чек-листы-Записи";

const SHEET_EVAL_CRITERIA   = "КаталогКритериевОценки";
const SHEET_EVALUATIONS     = "ОценкиОператоров";
const SHEET_EVAL_ANSWERS    = "ОтветыОценки";
// Тип техники для пунктов блока 1 («Обязанности оператора») — общий для
// всех профессий, не зависит от типа техники.
const EVAL_UNIVERSAL_TYPE = "ВСЕ";

function doGet(e) {
  const action   = e.parameter.action;
  const callback = e.parameter.callback;

  let result;
  if (action === "list")            result = getEmployees(e.parameter.site);
  else if (action === "getPersonnelStats") result = getPersonnelStats();
  else if (action === "getTests")   result = getTests();
  else if (action === "getResults") result = getResults();
  else if (action === "getTraining") result = getTraining();
  else if (action === "getRequirements") result = getRequirements();
  else if (action === "getDocuments") result = getDocuments(e.parameter.linkedId);
  else if (action === "getTrainingPlan") result = getTrainingPlan(e.parameter.site);
  else if (action === "getCourseCatalog") result = getCourseCatalog();
  else if (action === "getPersonnelEvents") result = getPersonnelEvents();
  else if (action === "getSurveys") result = getSurveys();
  else if (action === "getAllSurveys") result = getAllSurveys();
  else if (action === "getEquipment") result = getEquipment(e.parameter.site);
  else if (action === "getChecklistTemplates") result = getChecklistTemplates();
  else if (action === "getChecklistRecords") result = getChecklistRecords(e.parameter.days);
  else if (action === "getChecklistStats") result = getChecklistStats(e.parameter);
  else if (action === "getFuelRecords") result = getFuelRecords(e.parameter.site, e.parameter.days);
  else if (action === "getEquipmentDayDetail") result = getEquipmentDayDetail(e.parameter);
  else if (action === "getShiftAssignments") result = getShiftAssignments(e.parameter.site);
  else if (action === "getActiveVahtas") result = getActiveVahtas();
  else if (action === "getEvaluationCriteria") result = getEvaluationCriteria(e.parameter.equipmentType);
  else if (action === "getEvaluationEquipmentTypes") result = getEvaluationEquipmentTypes();
  else if (action === "getEvaluationWorkTypes") result = getEvaluationWorkTypes();
  else if (action === "getEvaluationEquipmentModels") result = getEvaluationEquipmentModels(e.parameter);
  else if (action === "getEvaluationInstructors") result = getEvaluationInstructors();
  else if (action === "getEvaluationSupervisors") result = getEvaluationSupervisors();
  else if (action === "getOperatorTimingRecords") result = getOperatorTimingRecords(e.parameter.empId);
  else if (action === "getTimingTruckModels") result = getTimingTruckModels();
  else if (action === "getOperatorEvaluations") result = getOperatorEvaluations(e.parameter.empId);
  else if (action === "getEvaluationDetail") result = getEvaluationDetail(e.parameter.evalId);
  else if (action === "getEvaluationAlerts") result = getEvaluationAlerts();
  else if (action === "getEvaluationStats") result = getEvaluationStats(e.parameter);
  else result = json({ ok: false, error: "unknown action" });

  if (callback) {
    const text = result.getContent();
    return ContentService
      .createTextOutput(callback + "(" + text + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return result;
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === "saveResult")     return saveResult(d.result);
    if (d.action === "updateAccess")   return updateAccess(d.empName, d.passDate, d.validUntil, d.position, d.site, d.empId);
    if (d.action === "saveTraining")   return saveTraining(d.empId, d.empName, d.course, d.date, d.validUntil);
    if (d.action === "getTraining")    return getTraining();
    if (d.action === "uploadPhoto")    return uploadPhoto(d.empId, d.empName, d.base64, d.mime);
    if (d.action === "importTest")     return importTest(d.test);
    if (d.action === "hireEmployee")   return hireEmployee(d);
    if (d.action === "saveRequirement") return saveRequirement(d.position, d.requiredTests, d.requiredCourses);
    if (d.action === "addTrainingPlan") return addTrainingPlan(d.empId, d.empName, d.course, d.dueDate, d.site);
    if (d.action === "bulkAddTrainingFacts") return bulkAddTrainingFacts(d.records);
    if (d.action === "recognizeRoster") return recognizeRoster(d.image, d.mimeType);
    if (d.action === "uploadDocument") return uploadDocument(d);
    if (d.action === "completeTrainingPlan") return completeTrainingPlan(d);
    if (d.action === "deleteTrainingPlan") return deleteTrainingPlan(d);
    if (d.action === "addCourseCatalog") return addCourseCatalog(d.course);
    if (d.action === "deleteCourseCatalog") return deleteCourseCatalog(d.course);
    if (d.action === "savePersonnelEvent") return savePersonnelEvent(d.empId, d.empName, d.type, d.date, d.description, d.issuedBy);
    if (d.action === "updateEmployee")        return updateEmployee(d);
    if (d.action === "updatePersonnelEvent")  return updatePersonnelEvent(d);
    if (d.action === "deletePersonnelEvent")  return deletePersonnelEvent(d);
    if (d.action === "submitSurvey")   return submitSurvey(d);
    if (d.action === "approveSurvey")  return approveSurvey(d.row, d.force);
    if (d.action === "rejectSurvey")   return rejectSurvey(d.row);
    if (d.action === "registerEquipment") return registerEquipment(d);
    if (d.action === "bulkRegisterEquipment") return bulkRegisterEquipment(d);
    if (d.action === "bulkAssignTemplate") return bulkAssignTemplate(d);
    if (d.action === "bulkUpdateEquipmentSite") return bulkUpdateEquipmentSite(d);
    if (d.action === "deleteEquipment") return deleteEquipment(d);
    if (d.action === "renameEquipmentId") return renameEquipmentId(d);
    if (d.action === "deleteEmployee") return deleteEmployee(d);
    if (d.action === "deleteSurveyRow") return deleteSurveyRow(d);
    if (d.action === "importChecklistTemplate") return importChecklistTemplate(d);
    if (d.action === "submitChecklistRecord") return submitChecklistRecord(d);
    if (d.action === "submitFuelRecord") return submitFuelRecord(d);
    if (d.action === "saveShiftAssignment") return saveShiftAssignment(d);
    if (d.action === "updateShiftAssignment") return updateShiftAssignment(d);
    if (d.action === "deleteShiftAssignmentRow") return deleteShiftAssignmentRow(d);
    if (d.action === "swapDayNight") return swapDayNight(d);
    if (d.action === "setActiveVahta") return setActiveVahta(d);
    if (d.action === "setEquipmentStatus") return setEquipmentStatus(d);
    if (d.action === "clearShiftAssignments") return clearShiftAssignments();
    if (d.action === "saveEvaluation") return saveEvaluation(d);
    if (d.action === "updateEvaluation") return updateEvaluation(d);
    if (d.action === "saveTimingRecord") return saveTimingRecord(d);
    if (d.action === "updateTimingRecord") return updateTimingRecord(d);
    if (d.action === "deleteTimingRecord") return deleteTimingRecord(d);
    if (d.action === "deleteEvaluation") return deleteEvaluation(d);
    if (d.action === "generateEvaluationSummary") return generateEvaluationSummary(d);
    if (d.action === "importEvaluationCriteria") return importEvaluationCriteria(d);
    return json({ ok: false, error: "unknown action" });
  } catch(err) {
    return json({ ok: false, error: err.toString() });
  }
}

// ── Загрузка фото в Google Drive ────────────────────────────
function uploadPhoto(empId, empName, base64Data, mime) {
  try {
    const bytes  = Utilities.base64Decode(base64Data);
    const ext    = mime === "image/png" ? "png" : "jpg";
    const fname  = (empName || empId || "employee").replace(/\s+/g,"_") + "." + ext;
    const blob   = Utilities.newBlob(bytes, mime, fname);
    const folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);

    const it = folder.getFilesByName(fname);
    while (it.hasNext()) { it.next().setTrashed(true); }

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const photoUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w300-h300";

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_EMP);
    const rows  = sheet.getDataRange().getValues();
    const h     = rows[0];
    const nameCol = h.indexOf("name");
    const urlCol  = h.indexOf("photoUrl");
    const idCol   = h.indexOf("id");

    const hasRealId = empId && String(empId).trim() !== "";
    const srcName = String(empName || "").trim().toLowerCase();

    for (let i = 1; i < rows.length; i++) {
      const rowId   = String(rows[i][idCol]).trim();
      const rowName = String(rows[i][nameCol]).trim().toLowerCase();

      const matched = hasRealId
        ? (rowId === String(empId).trim())
        : (srcName !== "" && (rowName === srcName || rowName.includes(srcName) || srcName.includes(rowName)));

      if (matched) {
        sheet.getRange(i + 1, urlCol + 1).setValue(photoUrl);
        return json({ ok: true, url: photoUrl });
      }
    }
    return json({ ok: true, url: photoUrl, warning: "employee not found in sheet, url not saved" });
  } catch(err) {
    return json({ ok: false, error: err.toString() });
  }
}

// ── Обновить даты допуска при успешной сдаче теста ──────────
function updateAccess(empName, passDate, validUntil, position, site, empId) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EMP);
  if (!sheet) return json({ ok: false, error: "no sheet" });

  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  const nameCol = headers.indexOf("name");
  const passCol = headers.indexOf("passDate");
  const valCol  = headers.indexOf("validUntil");
  const posCol  = headers.indexOf("position");
  const siteCol = headers.indexOf("site");
  const idCol   = headers.indexOf("id");

  if (nameCol < 0 || passCol < 0 || valCol < 0)
    return json({ ok: false, error: "columns not found" });

  const hasRealId = empId && empId !== "manual" && idCol >= 0;
  const srcName = String(empName || "").trim().toLowerCase();

  for (let i = 1; i < rows.length; i++) {
    const rowId   = idCol >= 0 ? String(rows[i][idCol]).trim() : "";
    const rowName = String(rows[i][nameCol]).trim().toLowerCase();

    const matched = hasRealId
      ? (rowId === String(empId).trim())
      : (srcName !== "" && (rowName === srcName || rowName.includes(srcName) || srcName.includes(rowName)));

    if (matched) {
      if (passDate)   sheet.getRange(i + 1, passCol + 1).setValue(new Date(passDate));
      if (validUntil) sheet.getRange(i + 1, valCol  + 1).setValue(new Date(validUntil));
      if (position && posCol >= 0 && !rows[i][posCol]) {
        sheet.getRange(i + 1, posCol + 1).setValue(position);
      }
      if (site && siteCol >= 0 && !rows[i][siteCol]) {
        sheet.getRange(i + 1, siteCol + 1).setValue(site);
      }
      return json({ ok: true, updated: rows[i][nameCol] });
    }
  }

  return json({
    ok: false,
    notFound: true,
    error: "Сотрудник не найден в базе. Обратитесь к инструктору или заполните анкету, чтобы вас приняли в систему, затем вернитесь для сдачи теста.",
  });
}

// ── Список сотрудников ───────────────────────────────────────

// ══════════════════════════════════════════════════════
// СВОДКА ПО КАДРАМ — для дашборда, вместо полного списка сотрудников.
// «Требует внимания» показывает только просроченных/истекающих — их
// обычно единицы из сотен, поэтому достаточно отдать именно их данные,
// а не все 446 карточек (~300 КБ). Полный список по-прежнему доступен
// через action=list, когда он реально нужен (Сотрудники, Вахта и т.д.).
// ══════════════════════════════════════════════════════
function getPersonnelStats() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EMP);
  if (!sheet) return json({ ok: true, total: 0, ok_: 0, warning: 0, expired: 0, attention: [] });

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json({ ok: true, total: 0, ok_: 0, warning: 0, expired: 0, attention: [] });

  const headers = rows[0];
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);
  const WARN_DAYS = 30;
  const now = new Date(); now.setHours(0,0,0,0);

  let total = 0, okCount = 0, warning = 0, expired = 0;
  const attention = [];

  rows.slice(1).forEach(r => {
    if (!r[0]) return;
    total++;
    const rawDate = r[idx["validUntil"]];
    let status = "none";
    let daysLeft = null;
    if (rawDate) {
      const d = rawDate instanceof Date ? rawDate : new Date(rawDate);
      if (!isNaN(d.getTime())) {
        daysLeft = Math.round((d - now) / 86400000);
        status = daysLeft < 0 ? "expired" : (daysLeft <= WARN_DAYS ? "warning" : "ok");
      }
    }
    if (status === "ok") okCount++;
    else if (status === "warning") { warning++; }
    else if (status === "expired") { expired++; }

    if (status === "expired" || status === "warning") {
      attention.push({
        id: r[idx["id"]] || "", name: r[idx["name"]] || "",
        position: r[idx["position"]] || "", site: r[idx["site"]] || "",
        validUntil: rawDate instanceof Date ? Utilities.formatDate(rawDate, "Asia/Almaty", "yyyy-MM-dd") : String(rawDate||""),
        daysLeft: daysLeft, status: status,
      });
    }
  });

  attention.sort((a,b) => (a.daysLeft||0) - (b.daysLeft||0));
  return json({ ok: true, total, ok_: okCount, warning, expired, attention });
}

function getEmployees(site) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EMP);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const DATE_COLS = ["passDate","validUntil"];
  // site необязателен — без него отдаём всех, как раньше (для instructor.html и т.п.)
  const siteCol = headers.indexOf("site");
  const siteFilter = site ? String(site).trim() : "";
  const data = rows.slice(1).filter(r => r[0])
    .filter(r => !siteFilter || (siteCol >= 0 && String(r[siteCol] || "").trim() === siteFilter))
    .map(r => {
    const o = {};
    headers.forEach((h,i) => {
      const val = r[i];
      if (DATE_COLS.includes(h) && val instanceof Date) {
        o[h] = Utilities.formatDate(val, "Asia/Almaty", "yyyy-MM-dd");
      } else {
        o[h] = val !== undefined ? String(val) : "";
      }
    });
    return o;
  });
  return json(data);
}

// ── Загрузка тестов из листа «Тесты» ────────────────────────
function getTests() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TESTS);
  if (!sheet) return json([]);

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);

  const testsMap = {};
  const testsOrder = [];

  let curTestId    = "";
  let curTitle     = "";
  let curSite      = "all";
  let curPassing   = 80;
  let curTimeLimit = 20;
  let curTestType  = "dopusk";

  rows.slice(1).forEach(r => {
    const rawTestId = String(r[0]).trim();
    if (rawTestId) {
      curTestId    = rawTestId;
      curTitle     = String(r[1]).trim();
      curSite      = String(r[2]).trim() || "all";
      curPassing   = Number(r[3]) || 80;
      curTimeLimit = Number(r[4]) || 20;
      curTestType  = String(r[11] || "").trim() || "dopusk";
    }

    const qId      = String(r[5]).trim();
    const qText    = String(r[6]).trim();
    const qType    = String(r[7]).trim() || "single";
    const qOptions = String(r[8]).split("|").map(s => s.trim()).filter(Boolean);
    const qCorrect = String(r[9]).split("|").map(s => Number(s.trim())).filter(s => !isNaN(s));
    const qExplain = String(r[10]).trim();

    if (!curTestId || !qId || !qText) return;

    if (!testsMap[curTestId]) {
      testsMap[curTestId] = {
        id: curTestId, title: curTitle, site: curSite,
        passingScore: curPassing, timeLimit: curTimeLimit, active: true,
        testType: curTestType, questions: []
      };
      testsOrder.push(curTestId);
    }

    testsMap[curTestId].questions.push({
      id: qId, text: qText, type: qType,
      options: qOptions, correct: qCorrect, explanation: qExplain
    });
  });

  return json(testsOrder.map(id => testsMap[id]));
}

// ── Сохранить результат ──────────────────────────────────────
function saveResult(r) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_RESULTS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_RESULTS);
    sheet.appendRow([
      "ID","Дата","Время","Сотрудник","Тест",
      "Результат","Балл (%)","Правильных","Всего вопросов",
      "Проходной (%)","Время (мин)","Истекло время","Ответы"
    ]);
    sheet.getRange(1,1,1,13)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  const now  = new Date();
  const date = Utilities.formatDate(now,"Asia/Almaty","dd.MM.yyyy");
  const time = Utilities.formatDate(now,"Asia/Almaty","HH:mm");

  // Если у листа ещё 12 колонок (старый формат) — дописываем заголовок «Ответы»,
  // чтобы разбор ошибок заработал без пересоздания листа
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("Ответы") < 0) {
    sheet.getRange(1, headers.length+1).setValue("Ответы")
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
  }

  ensureCapacity(sheet, 1);
  sheet.appendRow([
    r.id||"", date, time,
    r.empName||"", r.testTitle||"",
    r.passed?"Сдал ✓":"Не сдал ✗",
    r.score||0, r.correctQ||0, r.totalQ||0,
    r.passingScore||0, r.duration||0,
    r.expired?"Да":"Нет",
    // Ответы работника: {"q1":[0],"q2":[2],...} — нужны, чтобы видеть,
    // на каких именно вопросах человек ошибается
    r.answers ? JSON.stringify(r.answers) : "",
  ]);

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow,1,1,13)
    .setBackground(r.passed?"#E8F5E9":"#FFEBEE");

  return json({ ok: true });
}


// ══════════════════════════════════════════════════════
// Гарантирует, что в листе есть свободные строки под запись.
// Google Sheets НЕ расширяет лист автоматически: когда строки кончаются,
// appendRow молча ничего не делает, а вызывающий код получает "ok".
// Именно так были потеряны результаты тестов с 15.07 по 02.08.2026.
// ══════════════════════════════════════════════════════
function ensureCapacity(sheet, needRows) {
  if (!sheet) return;
  const need = needRows || 1;
  const free = sheet.getMaxRows() - sheet.getLastRow();
  if (free < need) {
    // добавляем с запасом, чтобы не дёргать это на каждой записи
    sheet.insertRowsAfter(sheet.getMaxRows(), Math.max(need - free, 500));
  }
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// CORS
function doOptions(e) {
  return ContentService
    .createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── Получить все результаты тестов ──────────────────
function getResults() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_RESULTS);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);

  const HEADERS = ["id","date","time","empName","testTitle","result","score","correctQ","totalQ","passingScore","duration","expired","answers"];

  const data = rows.slice(1).filter(r=>r[0]).map(r=>{
    const o = {};
    HEADERS.forEach((h,i) => {
      const val = r[i];
      if (val instanceof Date) {
        if (h === "time") {
          o[h] = Utilities.formatDate(val, "Asia/Almaty", "HH:mm");
        } else {
          o[h] = Utilities.formatDate(val, "Asia/Almaty", "dd.MM.yyyy");
        }
      } else {
        o[h] = val !== undefined ? String(val) : "";
      }
    });
    return o;
  });
  return json(data);
}

// ── Сохранить запись об обучении по профессии / инструктажу ─
function saveTraining(empId, empName, course, date, validUntil) {
  appendTrainingFact(empId, empName, course, date, validUntil);
  return json({ ok: true });
}

// Общая логика дозаписи факта прохождения курса — используется и прямой
// фиксацией (saveTraining), и завершением плана (completeTrainingPlan),
// чтобы не дублировать создание/миграцию листа «Обучение» в двух местах.
function appendTrainingFact(empId, empName, course, date, validUntil) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_TRAINING);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TRAINING);
    sheet.appendRow(["empId","empName","course","date","validUntil"]);
    sheet.getRange(1,1,1,5)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 300);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 120);
  } else {
    const headerRow = sheet.getRange(1,1,1,Math.max(5, sheet.getLastColumn())).getValues()[0];
    if (headerRow[4] !== "validUntil") {
      sheet.getRange(1,5).setValue("validUntil");
      sheet.getRange(1,5).setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
      sheet.setColumnWidth(5, 120);
    }
  }

  const now = new Date();
  const dateStr = date || Utilities.formatDate(now, "Asia/Almaty", "dd.MM.yyyy");

  ensureCapacity(sheet, 1);
  sheet.appendRow([empId||"", empName||"", course||"", dateStr, validUntil||""]);
}

// Групповое зачисление после очного обучения (ростер): вся группа записывается
// ОДНИМ запросом (batch setValues), а не по одному человеку — на слабой связи
// 20-30 отдельных POST-запросов подряд просто не доедут все.
// Использует ту же логику создания/миграции листа, что и appendTrainingFact,
// чтобы не дублировать её — не переиспользую appendTrainingFact саму
// (она пишет по одной строке через appendRow), а готовим лист один раз и
// потом пишем весь блок разом через getRange().setValues().
function bulkAddTrainingFacts(records) {
  if (!Array.isArray(records) || !records.length) return json({ ok: false, error: "Нет записей" });

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_TRAINING);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TRAINING);
    sheet.appendRow(["empId","empName","course","date","validUntil","batchId"]);
    sheet.getRange(1,1,1,6)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 300);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 160);
  } else {
    const headerRow = sheet.getRange(1,1,1,Math.max(6, sheet.getLastColumn())).getValues()[0];
    if (headerRow[4] !== "validUntil") {
      sheet.getRange(1,5).setValue("validUntil");
      sheet.getRange(1,5).setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
      sheet.setColumnWidth(5, 120);
    }
    // batchId — общий идентификатор группового зачисления, по нему потом
    // находим прикреплённый скан ростера (см. uploadDocument/getDocuments).
    if (headerRow[5] !== "batchId") {
      sheet.getRange(1,6).setValue("batchId");
      sheet.getRange(1,6).setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
      sheet.setColumnWidth(6, 160);
    }
  }

  const defaultDateStr = Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy");
  const batchId = "batch" + Date.now();
  const rows = records.map(r => [
    r.empId||"", r.empName||"", r.course||"", r.date||defaultDateStr, r.validUntil||"", batchId,
  ]);

  ensureCapacity(sheet, rows.length);
  sheet.getRange(sheet.getLastRow()+1, 1, rows.length, 6).setValues(rows);
  return json({ ok: true, added: rows.length, batchId });
}

// ══════════════════════════════════════════════════════════
// РАСПОЗНАВАНИЕ РОСТЕРА ПО ФОТО — через Claude API (модель с vision).
// Ключ хранится в Script Properties (Project Settings → Script Properties,
// имя ANTHROPIC_API_KEY), НЕ в самом коде — иначе он утечёт при любом
// экспорте/просмотре файла. Пользователь вписывает его сам через интерфейс
// Google, сюда он не попадает.
// ══════════════════════════════════════════════════════════
function recognizeRoster(imageBase64, mimeType) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
  if (!apiKey) {
    return json({ ok: false, error: "API-ключ не настроен. Project Settings → Script Properties → ANTHROPIC_API_KEY." });
  }
  if (!imageBase64) return json({ ok: false, error: "Нет фото" });

  const payload = {
    model: "claude-haiku-4-5-20251001",
    max_tokens: 1024,
    messages: [{
      role: "user",
      content: [
        {
          type: "image",
          source: { type: "base64", media_type: mimeType || "image/jpeg", data: imageBase64 },
        },
        {
          type: "text",
          text: "На фото — рукописный ростер (список сотрудников, прошедших обучение). " +
                "Извлеки ТОЛЬКО список ФИО из колонки «ФИО», каждое имя на отдельной строке, " +
                "в том порядке, как они идут в таблице сверху вниз. Ничего больше не пиши — " +
                "ни номеров строк, ни заголовков, ни пояснений, ни названия курса, ни дат. " +
                "Если какое-то имя прочитать не полностью уверенно — всё равно дай лучший вариант, " +
                "не пропускай строку.",
        },
      ],
    }],
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", options);
    const code = response.getResponseCode();
    const body = JSON.parse(response.getContentText());
    if (code !== 200) {
      const msg = (body.error && body.error.message) || `HTTP ${code}`;
      return json({ ok: false, error: msg });
    }
    const text = (body.content && body.content[0] && body.content[0].text) || "";
    return json({ ok: true, text: text.trim() });
  } catch (e) {
    return json({ ok: false, error: e.message });
  }
}

// ══════════════════════════════════════════════════════════
// АРХИВ ДОКУМЕНТОВ — сканы/фото ростеров, протоколов экзаменов и т.д.
// Файлы лежат на Google Диске (в папке рядом с самой таблицей), метаданные —
// в листе «Документы». type и linkedIds намеренно свободные строки, а не
// жёсткий enum/схема — чтобы новые виды документов добавлялись без правок
// бэкенда. linkedIds — то, что связывает скан с конкретной записью
// (batchId группового зачисления, id теста и т.п.), сравнение по вхождению
// в запятую-разделённый список, т.к. один скан может относиться к
// нескольким записям сразу.
// ══════════════════════════════════════════════════════════
function getOrCreateDocsFolder() {
  const file = DriveApp.getFileById(SHEET_ID);
  const parents = file.getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const existing = parentFolder.getFoldersByName("Документы — Caravan Resources");
  if (existing.hasNext()) return existing.next();
  return parentFolder.createFolder("Документы — Caravan Resources");
}

function uploadDocument(p) {
  const { image, mimeType, fileName, type, title, date, site, linkedIds, note } = p || {};
  if (!image) return json({ ok: false, error: "Нет файла" });
  if (!type) return json({ ok: false, error: "Нужен тип документа" });

  try {
    const folder = getOrCreateDocsFolder();
    const bytes = Utilities.base64Decode(image);
    const blob = Utilities.newBlob(bytes, mimeType || "image/jpeg", fileName || ("doc_" + Date.now()));
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = file.getUrl();

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_DOCUMENTS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_DOCUMENTS);
      sheet.appendRow(["docId","type","title","date","site","linkedIds","fileUrl","fileName","uploadedDate","note"]);
      sheet.getRange(1,1,1,10)
        .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(2, 130);
      sheet.setColumnWidth(3, 220);
      sheet.setColumnWidth(6, 160);
      sheet.setColumnWidth(7, 260);
    }

    const docId = "doc" + Date.now();
    const uploadedDate = Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy HH:mm");
    ensureCapacity(sheet, 1);
    sheet.appendRow([docId, type, title||"", date||"", site||"", linkedIds||"", fileUrl, fileName||"", uploadedDate, note||""]);

    return json({ ok: true, docId, fileUrl });
  } catch (e) {
    return json({ ok: false, error: e.message });
  }
}

// linkedId необязателен — без него отдаёт весь архив (для будущего общего просмотра).
function getDocuments(linkedId) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_DOCUMENTS);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const data = rows.slice(1).filter(r => r[0]).map(r => {
    const o = {}; headers.forEach((h,i) => { o[h] = r[i]; }); return o;
  });
  if (!linkedId) return json(data);
  const filtered = data.filter(d =>
    String(d.linkedIds||"").split(",").map(s=>s.trim()).includes(String(linkedId).trim())
  );
  return json(filtered);
}

// ── Получить все записи об обучении ─────────────────────────
function getTraining() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TRAINING);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const data = rows.slice(1).filter(r=>r[0]||r[1]).map(r=>{
    const o={}; headers.forEach((h,i)=>{o[h]=r[i];}); return o;
  });
  return json(data);
}

// ══════════════════════════════════════════════════════════
// ПЛАН ОБУЧЕНИЯ — курсы, назначенные, но ещё не пройденные.
// Отдельный лист от «Обучение» специально: «Обучение» активно используется
// в instructor.html (validUntil там уже различает периодический инструктаж
// от разового обучения по профессии) — плановые записи туда подмешивать
// нельзя, испортит существующую статистику. Как план выполняется —
// строка отсюда удаляется, а факт пишется в «Обучение» через
// appendTrainingFact (тот же путь, что и ручная фиксация).
// ══════════════════════════════════════════════════════════
function getTrainingPlan(site) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TRAINING_PLAN);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const siteCol = headers.indexOf("site");
  const siteFilter = site ? String(site).trim() : "";
  const data = [];
  rows.slice(1).forEach((r, i) => {
    if (!r[0] && !r[1]) return;
    if (siteFilter && (siteCol < 0 || String(r[siteCol] || "").trim() !== siteFilter)) return;
    const o = {};
    headers.forEach((h,j) => { o[h] = r[j]; });
    o.__row = i + 2;
    data.push(o);
  });
  return json(data);
}

function addTrainingPlan(empId, empName, course, dueDate, site) {
  if (!empId || !course) return json({ ok: false, error: "Нужны сотрудник и курс" });
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_TRAINING_PLAN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TRAINING_PLAN);
    sheet.appendRow(["empId","empName","course","dueDate","createdDate","site"]);
    sheet.getRange(1,1,1,6)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  const now = Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy");
  ensureCapacity(sheet, 1);
  sheet.appendRow([empId||"", empName||"", course||"", dueDate||"", now, site||""]);
  return json({ ok: true });
}

// План выполнен: пишем факт в «Обучение» и убираем строку плана.
function completeTrainingPlan(p) {
  const { row, empId, empName, course, date, validUntil } = p || {};
  if (!row) return json({ ok: false, error: "Нужен row" });
  appendTrainingFact(empId, empName, course, date, validUntil);
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TRAINING_PLAN);
  if (sheet) sheet.deleteRow(row);
  return json({ ok: true });
}

// План отменён без прохождения — просто убрать строку.
function deleteTrainingPlan(p) {
  const { row } = p || {};
  if (!row) return json({ ok: false, error: "Нужен row" });
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TRAINING_PLAN);
  if (!sheet) return json({ ok: false, error: "Лист «ПланОбучения» не найден" });
  sheet.deleteRow(row);
  return json({ ok: true });
}

// ══════════════════════════════════════════════════════════
// КАТАЛОГ КУРСОВ — фиксированный список названий, чтобы не печатать
// вручную и избежать разнобоя («Пожарная безопасность» / «пожарная без-ть»).
// ══════════════════════════════════════════════════════════
function getCourseCatalog() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_COURSE_CATALOG);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const data = rows.slice(1).map(r => String(r[0]||"").trim()).filter(Boolean);
  return json(data);
}

function addCourseCatalog(course) {
  const name = String(course||"").trim();
  if (!name) return json({ ok: false, error: "Нужно название курса" });
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_COURSE_CATALOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_COURSE_CATALOG);
    sheet.appendRow(["course"]);
    sheet.getRange(1,1,1,1)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 300);
  }
  const existing = sheet.getDataRange().getValues().slice(1).map(r => String(r[0]||"").trim());
  if (existing.includes(name)) return json({ ok: true, duplicate: true });
  ensureCapacity(sheet, 1);
  sheet.appendRow([name]);
  return json({ ok: true });
}

function deleteCourseCatalog(course) {
  const name = String(course||"").trim();
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_COURSE_CATALOG);
  if (!sheet) return json({ ok: false, error: "Каталог не найден" });
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]||"").trim() === name) {
      sheet.deleteRow(i + 1);
      return json({ ok: true });
    }
  }
  return json({ ok: false, error: "Не найдено" });
}

// ══════════════════════════════════════════════════════════
// УНИВЕРСАЛЬНЫЙ ИМПОРТ ТЕСТА (для панели инструктора)
// ══════════════════════════════════════════════════════════
function importTest(test) {
  if (!test || !test.testTitle || !Array.isArray(test.questions) || test.questions.length === 0) {
    return json({ ok: false, error: "Некорректные данные теста: нужны testTitle и хотя бы один вопрос" });
  }

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_TESTS);
  if (!sheet) sheet = ss.insertSheet(SHEET_TESTS);

  const HEADER = ["testId","testTitle","site","passingScore","timeLimit",
                  "qId","qText","qType","qOptions (через |)","qCorrect (индекс с 0)",
                  "qExplanation","testType"];

  let rows = sheet.getDataRange().getValues();
  if (rows.length === 0 || !rows[0][0]) {
    sheet.getRange(1,1,1,HEADER.length).setValues([HEADER]);
    sheet.getRange(1,1,1,HEADER.length)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
    rows = [HEADER];
  }

  const body = rows.slice(1);

  let testId = String(test.testId || "").trim();
  if (!testId) {
    let maxN = 0;
    body.forEach(r => {
      const m = String(r[0] || "").match(/^t(\d+)$/);
      if (m) maxN = Math.max(maxN, parseInt(m[1], 10));
    });
    testId = "t" + (maxN + 1);
  }

  const kept = [];
  let skipping = false;
  body.forEach(r => {
    const rid = String(r[0] || "").trim();
    if (rid !== "") skipping = (rid === testId);
    if (!skipping) kept.push(r);
  });

  const newRows = test.questions.map((q, i) => {
    const qId       = "q" + (i + 1);
    const qOptions  = (q.options || []).join("|");
    const qCorrect  = (q.correct || []).join("|");
    const first     = i === 0;
    return [
      first ? testId : "",
      first ? test.testTitle : "",
      first ? (test.site || "all") : "",
      first ? (test.passingScore || 80) : "",
      first ? (test.timeLimit || 20) : "",
      qId,
      q.text || "",
      q.type || "single",
      qOptions,
      qCorrect,
      q.explanation || "",
      first ? (test.testType || "dopusk") : "",
    ];
  });

  const finalRows = kept.concat(newRows);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, HEADER.length).clearContent();
  }
  if (finalRows.length > 0) {
    sheet.getRange(2, 1, finalRows.length, HEADER.length).setValues(finalRows);
  }

  return json({ ok: true, testId: testId, questions: newRows.length });
}

// ══════════════════════════════════════════════════════════
// ПРИЁМ СОТРУДНИКА — присвоение табельного номера по ТОО
// ══════════════════════════════════════════════════════════
function nextEmployeeId(too) {
  const tooStr = String(too || "");
  const prefix = /^Подрядчик:/.test(tooStr)
    ? CONTRACTOR_PREFIX
    : (TOO_PREFIXES[tooStr] || "СОТР");
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EMP);
  if (!sheet) return prefix + "-0001";

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return prefix + "-0001";

  const headers = rows[0];
  const idCol = headers.indexOf("id");
  if (idCol < 0) return prefix + "-0001";

  let maxN = 0;
  const re = new RegExp("^" + prefix + "-(\\d+)$");
  rows.slice(1).forEach(r => {
    const m = String(r[idCol] || "").match(re);
    if (m) maxN = Math.max(maxN, parseInt(m[1], 10));
  });

  const next = maxN + 1;
  return prefix + "-" + String(next).padStart(4, "0");
}

function ensureEmpColumns(sheet, headers) {
  const needed = ["tooPrimary", "tooSecondary", "hireDate", "address", "education", "experienceYears", "specialties", "additionalSkills", "email", "certificates", "courses"];
  let lastCol = headers.length;
  needed.forEach(col => {
    if (headers.indexOf(col) < 0) {
      lastCol += 1;
      sheet.getRange(1, lastCol).setValue(col);
      sheet.getRange(1, lastCol).setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
      headers.push(col);
    }
  });
  return headers;
}

function hireEmployee(p) {
  const { name, position, site, too, tooSecondary, hireDate, phone, address,
          education, experienceYears, specialties, additionalSkills, email, certificates, courses,
          photoUrl, force } = p || {};

  if (!name || !too) {
    return json({ ok: false, error: "Нужны как минимум имя и основное ТОО" });
  }

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_EMP);
  if (!sheet) return json({ ok: false, error: "Лист «Сотрудники» не найден" });

  const rows0 = sheet.getDataRange().getValues();
  let headers = rows0[0];

  if (!force) {
    const nameCol = headers.indexOf("name");
    const idCol   = headers.indexOf("id");
    const srcName = String(name).trim().toLowerCase();
    const candidates = [];
    if (nameCol >= 0) {
      rows0.slice(1).forEach(r => {
        const rowName = String(r[nameCol] || "").trim().toLowerCase();
        if (!rowName) return;
        if (rowName === srcName || rowName.includes(srcName) || srcName.includes(rowName)) {
          candidates.push({ id: idCol >= 0 ? String(r[idCol]) : "", name: r[nameCol] });
        }
      });
    }
    if (candidates.length > 0) {
      return json({ ok: false, warning: true, candidates: candidates,
        message: "Похожее имя уже есть в базе. Проверьте, не тот же ли это человек. Для создания нового профиля повторите запрос с force=true." });
    }
  }

  headers = ensureEmpColumns(sheet, headers);

  const newId = nextEmployeeId(too);
  const dateStr = hireDate || Utilities.formatDate(new Date(), "Asia/Almaty", "yyyy-MM-dd");

  const newRow = headers.map(h => {
    if (h === "id")             return newId;
    if (h === "name")           return name;
    if (h === "position")       return position || "";
    if (h === "site")           return site || "";
    if (h === "phone")          return phone || "";
    if (h === "tooPrimary")     return too;
    if (h === "tooSecondary")   return tooSecondary || "";
    if (h === "hireDate")       return dateStr;
    if (h === "address")        return address || "";
    if (h === "education")      return education || "";
    if (h === "experienceYears") return experienceYears || "";
    if (h === "specialties")    return specialties || "";
    if (h === "additionalSkills") return additionalSkills || "";
    if (h === "email")          return email || "";
    if (h === "certificates")   return certificates || "";
    if (h === "courses")        return courses || "";
    if (h === "photoUrl")       return photoUrl || "";
    return "";
  });
  ensureCapacity(sheet, 1);
  // Пишем через явный getRange, а не appendRow — иначе новая ячейка "наследует"
  // формат соседних ячеек в столбце (где-то Дата, где-то 0.00 с запятой), и
  // Sheets переинтерпретирует, например, "5" как дату 05.01.1900. Формат
  // ставим ДО записи значения, иначе Sheets успевает один раз неверно
  // истолковать строку при вставке, до того как формат применится.
  const nextRow = sheet.getLastRow() + 1;
  const expCol = headers.indexOf("experienceYears");
  if (expCol >= 0) sheet.getRange(nextRow, expCol + 1).setNumberFormat("@");
  sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

  return json({ ok: true, id: newId, name: name });
}

// ══════════════════════════════════════════════════════════
// ТРЕБОВАНИЯ ПО ДОЛЖНОСТЯМ — какие тесты обязательны для профессии
// ══════════════════════════════════════════════════════════
function getRequirements() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_REQUIREMENTS);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const data = rows.slice(1).filter(r => r[0]).map(r => ({
    position: String(r[0]).trim(),
    requiredTests: String(r[1] || "").split(",").map(s => s.trim()).filter(Boolean),
    requiredCourses: String(r[2] || "").split(",").map(s => s.trim()).filter(Boolean),
  }));
  return json(data);
}

// requiredTests/requiredCourses необязательны и обновляются НЕЗАВИСИМО:
// если параметр не передан (undefined) — соответствующая колонка не трогается.
// Так старые вызовы (instructor.html шлёт только requiredTests) не затирают
// requiredCourses, заданные отдельно через другой интерфейс, и наоборот.
function saveRequirement(position, requiredTests, requiredCourses) {
  if (!position) return json({ ok: false, error: "Нужна должность" });

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_REQUIREMENTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_REQUIREMENTS);
    sheet.appendRow(["position", "requiredTests", "requiredCourses"]);
    sheet.getRange(1,1,1,3)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 300);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 200);
  } else {
    const headerRow = sheet.getRange(1,1,1,Math.max(3, sheet.getLastColumn())).getValues()[0];
    if (headerRow[2] !== "requiredCourses") {
      sheet.getRange(1,3).setValue("requiredCourses");
      sheet.getRange(1,3).setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
      sheet.setColumnWidth(3, 200);
    }
  }

  const rows = sheet.getDataRange().getValues();
  const testsStr = requiredTests !== undefined
    ? (Array.isArray(requiredTests) ? requiredTests.join(",") : String(requiredTests || ""))
    : undefined;
  const coursesStr = requiredCourses !== undefined
    ? (Array.isArray(requiredCourses) ? requiredCourses.join(",") : String(requiredCourses || ""))
    : undefined;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(position).trim()) {
      if (testsStr   !== undefined) sheet.getRange(i + 1, 2).setValue(testsStr);
      if (coursesStr !== undefined) sheet.getRange(i + 1, 3).setValue(coursesStr);
      return json({ ok: true, updated: position });
    }
  }
  ensureCapacity(sheet, 1);
  sheet.appendRow([position, testsStr || "", coursesStr || ""]);
  return json({ ok: true, created: position });
}

// ══════════════════════════════════════════════════════════
// ЛИЧНОЕ ДЕЛО — взыскания и поощрения
// ══════════════════════════════════════════════════════════
// Лист: empId | empName | date | type (взыскание/поощрение) | description | issuedBy
function getPersonnelEvents() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PERSONNEL_EVENTS);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];

  // __row — номер строки в листе, нужен фронтенду для точечного
  // редактирования/удаления конкретной записи (как в getSurveys/approveSurvey).
  // Считается ДО фильтрации пустых строк, иначе после filter() индексы
  // "уедут" и будут указывать не на ту строку листа.
  const data = [];
  rows.slice(1).forEach((r, i) => {
    if (!r[0] && !r[1]) return;
    const o = {};
    headers.forEach((h,j) => { o[h] = r[j]; });
    o.__row = i + 2;
    data.push(o);
  });
  return json(data);
}

// evalId (необязательный) — если событие создано автоматически из
// saveEvaluation/updateEvaluation, связывает запись личного дела с
// конкретной оценкой, чтобы при редактировании оценки обновлять эту же
// строку, а не плодить дубликаты в личном деле на каждое сохранение.
function savePersonnelEvent(empId, empName, type, date, description, issuedBy, evalId) {
  if (!empName || !type) {
    return json({ ok: false, error: "Нужны как минимум имя сотрудника и тип события" });
  }

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_PERSONNEL_EVENTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PERSONNEL_EVENTS);
    sheet.appendRow(["empId", "empName", "date", "type", "description", "issuedBy", "evalId"]);
    sheet.getRange(1,1,1,7)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 350);
    sheet.setColumnWidth(6, 180);
  } else {
    const headerRow = sheet.getRange(1,1,1,Math.max(7, sheet.getLastColumn())).getValues()[0];
    if (headerRow[6] !== "evalId") {
      sheet.getRange(1,7).setValue("evalId");
      sheet.getRange(1,7).setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    }
  }

  const dateStr = date || Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy");
  ensureCapacity(sheet, 1);
  sheet.appendRow([empId || "", empName, dateStr, type, description || "", issuedBy || "", evalId || ""]);
  return json({ ok: true });
}

// ── Редактирование карточки сотрудника ───────────────────────
// Только уже существующие поля из EDITABLE — id/passDate/validUntil
// сюда намеренно не входят: даты допуска трогает updateAccess при сдаче
// теста, id вообще не редактируется. Пустая строка в поле — осознанная
// очистка значения, а не "поле не прислали".
function updateEmployee(p) {
  const { id, name, position, site, tooPrimary, tooSecondary, phone, email,
          address, hireDate, education, experienceYears, specialties,
          additionalSkills, certificates, courses } = p || {};

  if (!id)   return json({ ok: false, error: "Нужен id сотрудника" });
  if (!name) return json({ ok: false, error: "Нужно ФИО" });

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EMP);
  if (!sheet) return json({ ok: false, error: "Лист «Сотрудники» не найден" });

  const rows0   = sheet.getDataRange().getValues();
  const headers = ensureEmpColumns(sheet, rows0[0]);
  const idCol   = headers.indexOf("id");
  if (idCol < 0) return json({ ok: false, error: "Не найдена колонка id" });

  const values = { name, position, site, tooPrimary, tooSecondary, phone, email,
    address, hireDate, education, experienceYears, specialties, additionalSkills,
    certificates, courses };

  for (let i = 1; i < rows0.length; i++) {
    if (String(rows0[i][idCol]).trim() === String(id).trim()) {
      Object.keys(values).forEach(field => {
        const col = headers.indexOf(field);
        if (col >= 0 && values[field] !== undefined) {
          const cell = sheet.getRange(i + 1, col + 1);
          // Тот же баг с "наследованием" формата ячейки — фиксируем текстом
          // именно перед experienceYears, чтобы Sheets не превратил число в дату.
          if (field === "experienceYears") cell.setNumberFormat("@");
          cell.setValue(values[field]);
        }
      });
      return json({ ok: true, id, updated: name });
    }
  }
  return json({ ok: false, error: "Сотрудник с таким id не найден" });
}

// ── Редактирование / удаление записи личного дела ────────────
// Адресация по номеру строки (row), как у approveSurvey/rejectSurvey.
// empId — подстраховка от гонки: если между чтением списка и сохранением
// кто-то удалил строку выше этой, номер строки "уедет" на другого
// человека. Сверяем empId перед записью, тем же принципом, что и
// остальной матчинг в этом файле.
function updatePersonnelEvent(p) {
  const { row, empId, type, description, issuedBy, date } = p || {};
  if (!row) return json({ ok: false, error: "Нужен номер строки (row)" });

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PERSONNEL_EVENTS);
  if (!sheet) return json({ ok: false, error: "Лист «Личное дело» не найден" });

  const rowNum = Number(row);
  if (!rowNum || rowNum < 2 || rowNum > sheet.getLastRow()) {
    return json({ ok: false, error: "Строка не найдена — список успел измениться, обновите страницу." });
  }

  const headers  = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const empIdCol = headers.indexOf("empId");
  const typeCol  = headers.indexOf("type");
  const dateCol  = headers.indexOf("date");
  const descCol  = headers.indexOf("description");
  const issCol   = headers.indexOf("issuedBy");

  if (empId && empIdCol >= 0) {
    const actualEmpId = String(sheet.getRange(rowNum, empIdCol + 1).getValue()).trim();
    if (actualEmpId !== String(empId).trim()) {
      return json({ ok: false, error: "Строка сдвинулась (список изменился параллельно). Обновите страницу и повторите." });
    }
  }

  if (type !== undefined && typeCol >= 0)        sheet.getRange(rowNum, typeCol + 1).setValue(type);
  if (date !== undefined && dateCol >= 0)        sheet.getRange(rowNum, dateCol + 1).setValue(date);
  if (description !== undefined && descCol >= 0) sheet.getRange(rowNum, descCol + 1).setValue(description);
  if (issuedBy !== undefined && issCol >= 0)     sheet.getRange(rowNum, issCol + 1).setValue(issuedBy);

  return json({ ok: true, row: rowNum });
}

function deletePersonnelEvent(p) {
  const { row, empId } = p || {};
  if (!row) return json({ ok: false, error: "Нужен номер строки (row)" });

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PERSONNEL_EVENTS);
  if (!sheet) return json({ ok: false, error: "Лист «Личное дело» не найден" });

  const rowNum = Number(row);
  if (!rowNum || rowNum < 2 || rowNum > sheet.getLastRow()) {
    return json({ ok: false, error: "Строка не найдена — список успел измениться, обновите страницу." });
  }

  const headers  = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const empIdCol = headers.indexOf("empId");

  if (empId && empIdCol >= 0) {
    const actualEmpId = String(sheet.getRange(rowNum, empIdCol + 1).getValue()).trim();
    if (actualEmpId !== String(empId).trim()) {
      return json({ ok: false, error: "Строка сдвинулась (список изменился параллельно). Обновите страницу и повторите." });
    }
  }

  sheet.deleteRow(rowNum);
  return json({ ok: true, deleted: rowNum });
}

// ══════════════════════════════════════════════════════════
// АНКЕТЫ — самозаполнение через публичную страницу, с очередью
// подтверждения. Ничего не попадает в «Сотрудники» напрямую —
// только после ручного одобрения инструктором (approveSurvey).
// ══════════════════════════════════════════════════════════
const SHEET_SURVEYS = "Анкеты";

function submitSurvey(p) {
  const { empId, empName, position, site, too, tooSecondary, hireDate, address, phone,
          education, experienceYears, specialties, additionalSkills, email, certificates, courses,
          photoUrl } = p || {};

  if (!empName) return json({ ok: false, error: "Нужно ФИО" });

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_SURVEYS);
  let headers;

  const FULL_HEADER = ["empId","empName","position","site","too","tooSecondary","hireDate","address","phone",
                        "education","experienceYears","specialties","additionalSkills",
                        "email","certificates","courses",
                        "photoUrl","submittedAt","status"];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SURVEYS);
    sheet.appendRow(FULL_HEADER);
    sheet.getRange(1,1,1,FULL_HEADER.length)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(6, 300);
    sheet.setColumnWidth(10, 220);
    headers = FULL_HEADER;
  } else {
    headers = sheet.getDataRange().getValues()[0];
    let lastCol = headers.length;
    FULL_HEADER.forEach(col => {
      if (headers.indexOf(col) < 0) {
        lastCol += 1;
        sheet.getRange(1, lastCol).setValue(col);
        sheet.getRange(1, lastCol).setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
        headers.push(col);
      }
    });
  }

  const submittedAt = Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy HH:mm");
  const values = {
    empId: empId||"", empName, position: position||"", site: site||"", too: too||"", tooSecondary: tooSecondary||"",
    hireDate: hireDate||"", address: address||"", phone: phone||"",
    education: education||"", experienceYears: experienceYears||"", specialties: specialties||"",
    additionalSkills: additionalSkills||"",
    email: email||"", certificates: certificates||"", courses: courses||"",
    photoUrl: photoUrl||"",
    submittedAt, status: "pending",
  };
  const row = headers.map(h => values[h] !== undefined ? values[h] : "");
  ensureCapacity(sheet, 1);
  sheet.appendRow(row);
  return json({ ok: true });
}

function getSurveys() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SURVEYS);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const statusCol = headers.indexOf("status");

  const data = [];
  rows.slice(1).forEach((r, i) => {
    if (statusCol >= 0 && r[statusCol] !== "pending") return;
    const o = {};
    headers.forEach((h,j) => { o[h] = r[j]; });
    o.__row = i + 2;
    data.push(o);
  });
  return json(data);
}

function getAllSurveys() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SURVEYS);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];

  const data = [];
  rows.slice(1).forEach((r, i) => {
    const o = {};
    headers.forEach((h,j) => { o[h] = r[j]; });
    o.__row = i + 2;
    data.push(o);
  });
  return json(data);
}

function deleteSurveyRow(p) {
  const { row } = p || {};
  if (!row) return json({ ok: false, error: "Нужен номер строки" });

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SURVEYS);
  if (!sheet) return json({ ok: false, error: "Лист «Анкеты» не найден" });

  sheet.deleteRow(Number(row));
  return json({ ok: true, deleted: row });
}

function approveSurvey(rowNum, force) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return json({ ok: false, error: "Система занята обработкой другой заявки, попробуйте через пару секунд." });
  }

  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const surveySheet = ss.getSheetByName(SHEET_SURVEYS);
    if (!surveySheet) return json({ ok: false, error: "Лист «Анкеты» не найден" });

  const headers = surveySheet.getDataRange().getValues()[0];
  const row = surveySheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  const rec = {};
  headers.forEach((h,i) => { rec[h] = row[i]; });

  if (rec.status && rec.status !== "pending") {
    return json({ ok: false, error: "Эта анкета уже обработана (статус: " + rec.status + "). Обновите список." });
  }

  let result;
  if (rec.empId) {
    const empSheet = ss.getSheetByName(SHEET_EMP);
    const empRows = empSheet.getDataRange().getValues();
    const empHeaders = ensureEmpColumns(empSheet, empRows[0]);
    const idCol = empHeaders.indexOf("id");
    let found = false;
    for (let i = 1; i < empRows.length; i++) {
      if (String(empRows[i][idCol]) === String(rec.empId)) {
        ["tooPrimary","tooSecondary","hireDate","address","phone","education","experienceYears","specialties","additionalSkills","email","certificates","courses","photoUrl","position","site"].forEach(field => {
          const col = empHeaders.indexOf(field);
          const val = field === "tooPrimary" ? rec.too
                    : field === "tooSecondary" ? rec.tooSecondary
                    : rec[field];
          if (col >= 0 && val) {
            const cell = empSheet.getRange(i + 1, col + 1);
            if (field === "experienceYears") cell.setNumberFormat("@");
            cell.setValue(val);
          }
        });
        found = true;
        break;
      }
    }
    result = found
      ? { ok: true, updated: rec.empName }
      : { ok: false, error: "Сотрудник с ID " + rec.empId + " не найден в «Сотрудники»" };
  } else {
    const hireResult = hireEmployee({
      name: rec.empName, position: rec.position || "", site: rec.site || "", too: rec.too, tooSecondary: rec.tooSecondary,
      hireDate: rec.hireDate, phone: rec.phone, address: rec.address,
      education: rec.education, experienceYears: rec.experienceYears, specialties: rec.specialties,
      additionalSkills: rec.additionalSkills,
      email: rec.email, certificates: rec.certificates, courses: rec.courses,
      photoUrl: rec.photoUrl,
      force: !!force,
    });
    result = JSON.parse(hireResult.getContent());
  }

  if (result.ok) {
    surveySheet.getRange(rowNum, headers.indexOf("status") + 1).setValue("approved");
  }
  return json(result);
  } finally {
    lock.releaseLock();
  }
}

function rejectSurvey(rowNum) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SURVEYS);
  if (!sheet) return json({ ok: false, error: "Лист «Анкеты» не найден" });
  const headers = sheet.getDataRange().getValues()[0];
  const statusCol = headers.indexOf("status");
  if (statusCol < 0) return json({ ok: false, error: "Нет колонки status" });
  sheet.getRange(rowNum, statusCol + 1).setValue("rejected");
  return json({ ok: true });
}

// ══════════════════════════════════════════════════════════
// ТЕХНИКА — реестр физических единиц оборудования
// ══════════════════════════════════════════════════════════
function getEquipment(site) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EQUIPMENT);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const siteCol = headers.indexOf("site");
  const siteFilter = site ? String(site).trim() : "";
  const data = rows.slice(1).filter(r => r[0])
    .filter(r => !siteFilter || (siteCol >= 0 && String(r[siteCol] || "").trim() === siteFilter))
    .map(r => {
      const o = {}; headers.forEach((h,i) => { o[h] = r[i]; }); return o;
    });
  return json(data);
}

function registerEquipment(p) {
  const { equipmentId, model, type, site, inventoryNumber, templateId } = p || {};
  if (!equipmentId || !model) {
    return json({ ok: false, error: "Нужны как минимум номер техники и модель" });
  }

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_EQUIPMENT);
  const HEADER = ["equipmentId","model","type","site","inventoryNumber","status","addedDate","templateId"];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_EQUIPMENT);
    sheet.appendRow(HEADER);
    sheet.getRange(1,1,1,HEADER.length)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
  } else {
    const existingHeader = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    if (existingHeader.indexOf("templateId") < 0) {
      const col = existingHeader.length + 1;
      sheet.getRange(1, col).setValue("templateId");
      sheet.getRange(1, col).setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    }
  }

  const rows = sheet.getDataRange().getValues();
  const idCol = HEADER.indexOf("equipmentId");
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]).trim() === String(equipmentId).trim()) {
      return json({ ok: false, error: "Единица техники с таким номером уже зарегистрирована" });
    }
  }

  const dateStr = Utilities.formatDate(new Date(), "Asia/Almaty", "yyyy-MM-dd");
  ensureCapacity(sheet, 1);
  sheet.appendRow([equipmentId, model, type || "", site || "", inventoryNumber || "", "в работе", dateStr, templateId || ""]);
  return json({ ok: true, equipmentId });
}

function bulkRegisterEquipment(p) {
  const { items } = p || {};
  if (!Array.isArray(items) || items.length === 0) {
    return json({ ok: false, error: "Нужен непустой список items" });
  }

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_EQUIPMENT);
  const HEADER = ["equipmentId","model","type","site","inventoryNumber","status","addedDate","templateId"];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_EQUIPMENT);
    sheet.appendRow(HEADER);
    sheet.getRange(1,1,1,HEADER.length)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
  } else {
    const existingHeader = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    if (existingHeader.indexOf("templateId") < 0) {
      const col = existingHeader.length + 1;
      sheet.getRange(1, col).setValue("templateId");
      sheet.getRange(1, col).setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    }
  }

  const rows = sheet.getDataRange().getValues();
  const existingIds = new Set(rows.slice(1).map(r => String(r[0]).trim()));

  const dateStr = Utilities.formatDate(new Date(), "Asia/Almaty", "yyyy-MM-dd");
  let added = 0, skipped = 0;

  items.forEach(it => {
    const id = String(it.equipmentId || "").trim();
    if (!id || !it.model || existingIds.has(id)) { skipped++; return; }
    ensureCapacity(sheet, 1);
    sheet.appendRow([id, it.model, it.type || "", it.site || "", it.inventoryNumber || "", "в работе", dateStr, it.templateId || ""]);
    existingIds.add(id);
    added++;
  });

  return json({ ok: true, added, skipped });
}

function bulkAssignTemplate(p) {
  const { items } = p || {};
  if (!Array.isArray(items) || items.length === 0) {
    return json({ ok: false, error: "Нужен непустой список items" });
  }

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EQUIPMENT);
  if (!sheet) return json({ ok: false, error: "Лист «Техника» не найден" });

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf("equipmentId");
  let tplCol = headers.indexOf("templateId");
  if (idCol < 0) return json({ ok: false, error: "Не найдена колонка equipmentId" });
  if (tplCol < 0) {
    tplCol = headers.length;
    sheet.getRange(1, tplCol + 1).setValue("templateId");
    sheet.getRange(1, tplCol + 1).setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
  }

  const idToTemplate = {};
  items.forEach(it => { if (it.equipmentId) idToTemplate[String(it.equipmentId).trim()] = it.templateId || ""; });

  let updated = 0, notFound = 0;
  for (let i = 1; i < rows.length; i++) {
    const rowId = String(rows[i][idCol]).trim();
    if (idToTemplate.hasOwnProperty(rowId)) {
      sheet.getRange(i + 1, tplCol + 1).setValue(idToTemplate[rowId]);
      updated++;
      delete idToTemplate[rowId];
    }
  }
  notFound = Object.keys(idToTemplate).length;

  return json({ ok: true, updated, notFound });
}

// ── Массовая простановка участка для уже зарегистрированной техники ──
// equipmentIds не обязателен: если не передан — проставляет всем.
function bulkUpdateEquipmentSite(p) {
  const { site, equipmentIds } = p || {};
  if (!site) return json({ ok: false, error: "Нужен site" });

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EQUIPMENT);
  if (!sheet) return json({ ok: false, error: "Лист «Техника» не найден" });

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf("equipmentId");
  const siteCol = headers.indexOf("site");
  if (idCol < 0 || siteCol < 0) return json({ ok: false, error: "Не найдены нужные колонки" });

  const idSet = Array.isArray(equipmentIds) && equipmentIds.length
    ? new Set(equipmentIds.map(String))
    : null;

  let updated = 0;
  for (let i = 1; i < rows.length; i++) {
    const rowId = String(rows[i][idCol]).trim();
    if (idSet && !idSet.has(rowId)) continue;
    sheet.getRange(i + 1, siteCol + 1).setValue(site);
    updated++;
  }
  return json({ ok: true, updated });
}

function deleteEquipment(p) {
  const { equipmentId } = p || {};
  if (!equipmentId) return json({ ok: false, error: "Нужен equipmentId" });

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EQUIPMENT);
  if (!sheet) return json({ ok: false, error: "Лист «Техника» не найден" });

  const rows = sheet.getDataRange().getValues();
  const idCol = rows[0].indexOf("equipmentId");
  if (idCol < 0) return json({ ok: false, error: "Не найдена колонка equipmentId" });

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]).trim() === String(equipmentId).trim()) {
      sheet.deleteRow(i + 1);
      return json({ ok: true, deleted: equipmentId });
    }
  }
  return json({ ok: false, error: "Единица техники с таким номером не найдена" });
}

function renameEquipmentId(p) {
  const { oldId, newId } = p || {};
  if (!oldId || !newId) return json({ ok: false, error: "Нужны oldId и newId" });
  if (String(oldId).trim() === String(newId).trim()) return json({ ok: false, error: "Новый номер совпадает со старым" });

  const ss = SpreadsheetApp.openById(SHEET_ID);

  const eqSheet = ss.getSheetByName(SHEET_EQUIPMENT);
  if (!eqSheet) return json({ ok: false, error: "Лист «Техника» не найден" });
  const eqRows = eqSheet.getDataRange().getValues();
  const eqIdCol = eqRows[0].indexOf("equipmentId");
  if (eqIdCol < 0) return json({ ok: false, error: "Не найдена колонка equipmentId" });

  for (let i = 1; i < eqRows.length; i++) {
    if (String(eqRows[i][eqIdCol]).trim() === String(newId).trim()) {
      return json({ ok: false, error: "Единица техники с номером «" + newId + "» уже существует" });
    }
  }

  let found = false;
  for (let i = 1; i < eqRows.length; i++) {
    if (String(eqRows[i][eqIdCol]).trim() === String(oldId).trim()) {
      eqSheet.getRange(i + 1, eqIdCol + 1).setValue(newId);
      found = true;
      break;
    }
  }
  if (!found) return json({ ok: false, error: "Единица техники «" + oldId + "» не найдена" });

  let updatedRecords = 0;
  const recSheet = ss.getSheetByName(SHEET_CHECKLIST_RECORDS);
  if (recSheet) {
    const recRows = recSheet.getDataRange().getValues();
    const recIdCol = recRows[0].indexOf("equipmentId");
    if (recIdCol >= 0) {
      for (let i = 1; i < recRows.length; i++) {
        if (String(recRows[i][recIdCol]).trim() === String(oldId).trim()) {
          recSheet.getRange(i + 1, recIdCol + 1).setValue(newId);
          updatedRecords++;
        }
      }
    }
  }

  return json({ ok: true, oldId, newId, updatedRecords });
}

function deleteEmployee(p) {
  const { id } = p || {};
  if (!id) return json({ ok: false, error: "Нужен id" });

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EMP);
  if (!sheet) return json({ ok: false, error: "Лист «Сотрудники» не найден" });

  const rows = sheet.getDataRange().getValues();
  const idCol = rows[0].indexOf("id");
  if (idCol < 0) return json({ ok: false, error: "Не найдена колонка id" });

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 1);
      return json({ ok: true, deleted: id });
    }
  }
  return json({ ok: false, error: "Сотрудник с таким id не найден" });
}

// ══════════════════════════════════════════════════════════
// ЧЕК-ЛИСТЫ — ШАБЛОНЫ (по модели/типу техники, не по единице)
// ══════════════════════════════════════════════════════════
function getChecklistTemplates() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CHECKLIST_TEMPLATES);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);

  const templatesMap = {};
  const order = [];

  rows.slice(1).forEach(r => {
    const templateId = String(r[0] || "").trim();
    const title       = String(r[1] || "").trim();
    const section     = String(r[2] || "").trim();
    const itemId      = String(r[3] || "").trim();
    const itemText    = String(r[4] || "").trim();
    const metaStr     = String(r[5] || "").trim();
    if (!templateId || !itemId || !itemText) return;

    if (!templatesMap[templateId]) {
      templatesMap[templateId] = { id: templateId, title, sections: {} };
      order.push(templateId);
      if (metaStr) {
        try {
          const meta = JSON.parse(metaStr);
          if (meta.fuelLevelField) templatesMap[templateId].fuelLevelField = meta.fuelLevelField;
        } catch(e) { /* повреждённый meta — просто пропускаем, не роняем весь ответ */ }
      }
    }
    if (!templatesMap[templateId].sections[section]) {
      templatesMap[templateId].sections[section] = [];
    }
    templatesMap[templateId].sections[section].push({ id: itemId, text: itemText });
  });

  return json(order.map(id => templatesMap[id]));
}

function importChecklistTemplate(p) {
  const { templateId, title, sections, fuelLevelField } = p || {};
  if (!templateId || !title || !sections || typeof sections !== "object") {
    return json({ ok: false, error: "Нужны templateId, title и sections ({раздел: [пункты]})" });
  }

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_CHECKLIST_TEMPLATES);
  const HEADER = ["templateId","title","section","itemId","itemText","meta"];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CHECKLIST_TEMPLATES);
    sheet.appendRow(HEADER);
    sheet.getRange(1,1,1,HEADER.length)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
  } else {
    const existingHeader = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    if (existingHeader.indexOf("meta") < 0) {
      const col = existingHeader.length + 1;
      sheet.getRange(1, col).setValue("meta");
      sheet.getRange(1, col).setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    }
  }

  const rows = sheet.getDataRange().getValues();
  const body = rows.slice(1);

  const kept = body.filter(r => String(r[0] || "").trim() !== String(templateId).trim());

  const metaStr = fuelLevelField ? JSON.stringify({ fuelLevelField }) : "";

  const newRows = [];
  let itemCounter = 0;
  Object.keys(sections).forEach(sectionName => {
    (sections[sectionName] || []).forEach(itemText => {
      itemCounter += 1;
      const isFirst = itemCounter === 1;
      newRows.push([templateId, title, sectionName, "i" + itemCounter, itemText, isFirst ? metaStr : ""]);
    });
  });

  const finalRows = kept.concat(newRows);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, HEADER.length).clearContent();
  }
  if (finalRows.length > 0) {
    sheet.getRange(2, 1, finalRows.length, HEADER.length).setValues(finalRows);
  }

  return json({ ok: true, templateId, items: newRows.length });
}

// ══════════════════════════════════════════════════════════
// ЧЕК-ЛИСТЫ — ЗАПИСИ (заполненные осмотры)
// ══════════════════════════════════════════════════════════
function submitChecklistRecord(p) {
  const { equipmentId, templateId, empId, empName, items } = p || {};
  if (!equipmentId || !empName || !Array.isArray(items) || items.length === 0) {
    return json({ ok: false, error: "Нужны equipmentId, empName и хотя бы один пункт осмотра" });
  }

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_CHECKLIST_RECORDS);
  const HEADER = ["recordId","equipmentId","templateId","empId","empName","date","section","itemText","status","comment"];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CHECKLIST_RECORDS);
    sheet.appendRow(HEADER);
    sheet.getRange(1,1,1,HEADER.length)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  const recordId = "chk" + Date.now();
  const dateStr = Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy HH:mm");

  const rows = items.map(it => [
    recordId, equipmentId, templateId || "", empId || "", empName, dateStr,
    it.section || "", it.itemText || "", it.status || "", it.comment || "",
  ]);
  ensureCapacity(sheet, rows.length);
  rows.forEach(r => sheet.appendRow(r));

  return json({ ok: true, recordId, items: rows.length });
}

// days — за сколько последних дней отдавать осмотры. Без параметра или
// days=0 отдаётся вся история (обратная совместимость со старыми вызовами).
// Понадобилось, потому что полная выгрузка выросла до ~7 МБ и перестала
// доходить на слабой связи в карьере.

// ══════════════════════════════════════════════════════
// СЕРВЕРНАЯ АГРЕГАЦИЯ ОСМОТРОВ
// Считаем аналитику здесь, а не в браузере. Панель получает готовые цифры
// (десятки КБ) вместо всех записей (мегабайты, растут каждый день).
// Размер ответа почти не зависит от числа осмотров — это снимает потолок,
// в который мы упирались: 1300 осмотров = 7 МБ, 20 000 было бы ~100 МБ.
// ══════════════════════════════════════════════════════
function getChecklistStats(p) {
  const days = Number((p && p.days) || 0);
  // Точное окно [from, to) — для наград за конкретную неделю/месяц/квартал.
  // Без этого days=N трактовалось как "последние N дней от сегодня", и
  // прошлые периоды показывали накопленный итог до сегодняшнего дня вместо
  // данных именно за тот период.
  const fromParam = p && p.from ? new Date(p.from) : null;
  const toParam   = p && p.to   ? new Date(p.to)   : null;
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CHECKLIST_RECORDS);
  if (!sheet) return json({ ok: true, empty: true });

  // Участок берём из реестра техники
  const eqSheet = ss.getSheetByName(SHEET_EQUIPMENT);
  const eqSite = {}, eqModel = {}, eqType = {};
  if (eqSheet) {
    const er = eqSheet.getDataRange().getValues();
    const eh = er[0];
    const ci = n => eh.indexOf(n);
    const iId = ci("equipmentId"), iSite = ci("site"), iModel = ci("model"), iType = ci("type");
    er.slice(1).forEach(r => {
      const id = String(r[iId] || "").trim();
      if (!id) return;
      eqSite[id]  = String(r[iSite]  || "").trim();
      eqModel[id] = String(r[iModel] || "");
      eqType[id]  = String(r[iType]  || "");
    });
  }

  let since = null, until = null;
  if (fromParam && !isNaN(fromParam.getTime())) {
    since = fromParam;
    until = (toParam && !isNaN(toParam.getTime())) ? toParam : new Date();
  } else if (days > 0) {
    since = new Date();
    since.setDate(since.getDate() - days);
    since.setHours(0,0,0,0);
  }
  const now = new Date();
  const d7 = new Date(); d7.setDate(d7.getDate() - 7);

  // Читаем только нужные колонки (A..J), а не весь лист — заметно быстрее
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1
    ? sheet.getRange(1, 1, lastRow, 10).getValues()
    : [];

  // Накопители по участкам: "Алмалы", "Ашыктас" и "__all"
  function blank() {
    return {
      records: 0, items: 0, defects: 0, comments: 0,
      defects7: 0, records7: 0,
      equipment: {},        // id -> {osmotrs, defects, lastDate, lastRecordId}
      section: {},          // раздел -> {total, defects}
      defectText: {},       // текст -> count
      operator: {},         // имя -> {empId, osmotrs, defects, comments, eq:{}}
      day: {},              // дата -> {osmotrs, defects, eq:{}}
      lastRec: {},          // id -> последняя запись целиком (для «открытых замечаний»)
      recentDefects: [],    // свежие неисправности
      commentsList: [],     // свежие комментарии операторов
    };
  }
  const acc = { __all: blank() };

  // Группируем строки по recordId. Строки одного осмотра в листе НЕ обязательно
  // идут подряд, поэтому собираем через карту, а не «по соседству» — иначе один
  // осмотр разбивается на несколько и счётчик завышается.
  const recMap = {};
  const pending = [];
  for (let ri = 1; ri < rows.length; ri++) {
    const r = rows[ri];
    const recordId = String(r[0] || "").trim();
    if (!recordId) continue;
    const date = r[5] instanceof Date ? r[5] : new Date(r[5]);
    if (since && (!(date instanceof Date) || isNaN(date.getTime()) || date < since)) continue;
    if (until && date instanceof Date && !isNaN(date.getTime()) && date >= until) continue;

    let rec = recMap[recordId];
    if (!rec) {
      rec = recMap[recordId] = {
        recordId, equipmentId: String(r[1] || "").trim(),
        empId: r[3], empName: String(r[4] || ""), date: date, items: [],
      };
      pending.push(rec);
    }
    rec.items.push({
      section: String(r[6] || ""), itemText: String(r[7] || ""),
      status: String(r[8] || ""), comment: String(r[9] || ""),
    });
  }

  pending.forEach(rec => {
    const site = eqSite[rec.equipmentId] || "";
    const targets = [acc.__all];
    if (site) {
      if (!acc[site]) acc[site] = blank();
      targets.push(acc[site]);
    }
    const dateMs = rec.date.getTime();
    const dayKey = Utilities.formatDate(rec.date, "Asia/Almaty", "yyyy-MM-dd");
    const isWeek = rec.date >= d7;

    let recDefects = 0;
    const recDefectItems = [];
    rec.items.forEach(it => {
      const isDefect = it.status === "defect";
      const hasComment = String(it.comment || "").trim() !== "";
      if (isDefect) { recDefects++; recDefectItems.push(it); }

      targets.forEach(a => {
        a.items++;
        const s = it.section || "Без раздела";
        if (!a.section[s]) a.section[s] = { total: 0, defects: 0 };
        a.section[s].total++;
        if (isDefect) {
          a.defects++;
          a.section[s].defects++;
          a.defectText[it.itemText] = (a.defectText[it.itemText] || 0) + 1;
          if (isWeek) a.defects7++;
        }
        if (hasComment) a.comments++;
      });
    });

    targets.forEach(a => {
      a.records++;
      if (isWeek) a.records7++;

      const eq = a.equipment[rec.equipmentId] || (a.equipment[rec.equipmentId] = { osmotrs: 0, defects: 0, lastMs: 0 });
      eq.osmotrs++;
      eq.defects += recDefects;
      if (dateMs >= eq.lastMs) { eq.lastMs = dateMs; a.lastRec[rec.equipmentId] = rec; }

      const op = a.operator[rec.empName] || (a.operator[rec.empName] = { empId: rec.empId || "", osmotrs: 0, defects: 0, comments: 0, eq: {}, days: {} });
      if (rec.empId && !op.empId) op.empId = rec.empId;
      op.osmotrs++;
      op.defects += recDefects;
      op.comments += rec.items.filter(i => String(i.comment || "").trim()).length;
      op.eq[rec.equipmentId] = 1;
      op.days[dayKey] = 1;   // реальные уникальные дни, а не число осмотров

      const dd = a.day[dayKey] || (a.day[dayKey] = { osmotrs: 0, defects: 0, eq: {} });
      dd.osmotrs++;
      dd.defects += recDefects;
      dd.eq[rec.equipmentId] = 1;

      if (recDefects) {
        a.recentDefects.push({
          equipmentId: rec.equipmentId, empName: rec.empName,
          date: rec.date.toISOString(),
          items: recDefectItems.map(i => ({ itemText: i.itemText, comment: i.comment })),
        });
      }
      rec.items.forEach(it => {
        if (!String(it.comment || "").trim()) return;
        if (!a.commentsList) a.commentsList = [];
        a.commentsList.push({
          equipmentId: rec.equipmentId, empName: rec.empName,
          date: rec.date.toISOString(), itemText: it.itemText,
          status: it.status, comment: it.comment,
        });
      });
    });
  });

  // Приводим к компактному виду
  function pack(a) {
    if (!a) return null;
    const openDefects = [];
    Object.keys(a.lastRec).forEach(eid => {
      const rec = a.lastRec[eid];
      const defs = rec.items.filter(i => i.status === "defect");
      if (defs.length) {
        openDefects.push({
          equipmentId: eid, model: eqModel[eid] || "", empName: rec.empName,
          date: rec.date.toISOString(),
          items: defs.map(i => ({ itemText: i.itemText, comment: i.comment })),
        });
      }
    });
    openDefects.sort((x, y) => y.items.length - x.items.length);

    return {
      records: a.records, items: a.items, defects: a.defects, commentsCount: a.comments,
      records7: a.records7, defects7: a.defects7,
      inspectedCount: Object.keys(a.equipment).length,
      openDefects: openDefects,
      byEquipment: Object.keys(a.equipment).map(id => {
        const lr = a.lastRec[id];
        const lastDefects = lr ? lr.items.filter(i => i.status === "defect").length : 0;
        return {
          equipmentId: id, model: eqModel[id] || "", type: eqType[id] || "",
          osmotrs: a.equipment[id].osmotrs, defects: a.equipment[id].defects,
          // Данные последнего осмотра — для карточек техники на вкладках «Вахта»
          lastDate: lr ? lr.date.toISOString() : "",
          lastEmpName: lr ? lr.empName : "",
          lastDefects: lastDefects,
        };
      }).sort((x, y) => y.defects - x.defects || y.osmotrs - x.osmotrs),
      bySection: Object.keys(a.section).map(s => ({
        section: s, total: a.section[s].total, defects: a.section[s].defects,
      })).sort((x, y) => y.defects - x.defects),
      topDefects: Object.keys(a.defectText).map(t => ({
        itemText: t, count: a.defectText[t],
      })).sort((x, y) => y.count - x.count).slice(0, 30),
      byOperator: Object.keys(a.operator).map(n => ({
        empName: n, empId: a.operator[n].empId,
        osmotrs: a.operator[n].osmotrs, defects: a.operator[n].defects,
        comments: a.operator[n].comments,
        equipmentCount: Object.keys(a.operator[n].eq).length,
        daysCount: Object.keys(a.operator[n].days).length,
      })).sort((x, y) => y.defects - x.defects || y.osmotrs - x.osmotrs),
      byDay: Object.keys(a.day).map(d => ({
        date: d, osmotrs: a.day[d].osmotrs, defects: a.day[d].defects,
        equipmentCount: Object.keys(a.day[d].eq).length,
      })).sort((x, y) => y.date.localeCompare(x.date)).slice(0, 60),
      recentDefects: a.recentDefects.sort((x, y) => y.date.localeCompare(x.date)).slice(0, 60),
      comments: (a.commentsList || []).sort((x, y) => y.date.localeCompare(x.date)).slice(0, 40),
    };
  }

  const out = { ok: true, days: days, generatedAt: new Date().toISOString(), sites: {} };
  Object.keys(acc).forEach(k => {
    if (k === "__all") out.all = pack(acc[k]);
    else out.sites[k] = pack(acc[k]);
  });
  return json(out);
}

// Полная (несжатая — все пункты чек-листа, а не только неисправности/
// комментарии) выгрузка по одной единице техники за один календарный день.
// В отличие от getChecklistRecords/getChecklistStats тут нет ни отсечек
// по количеству, ни свёртки «ok»-пунктов — объём всегда маленький
// (техника + день), поэтому отдаём всё как есть. Используется из карточки
// техники на вкладках «Вахта» — «вытащить чек-лист и топливо на любую дату».
function getEquipmentDayDetail(p) {
  const equipmentId = String((p && p.equipmentId) || "").trim();
  const dateStr = String((p && p.date) || "").trim(); // yyyy-MM-dd, по Asia/Almaty
  if (!equipmentId || !dateStr) {
    return json({ ok: false, error: "Нужны equipmentId и date (yyyy-MM-dd)" });
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ── Чек-листы ──
  const chkSheet = ss.getSheetByName(SHEET_CHECKLIST_RECORDS);
  const checklists = [];
  if (chkSheet) {
    const lastRow = chkSheet.getLastRow();
    const rows = lastRow > 1 ? chkSheet.getRange(1, 1, lastRow, 10).getValues() : [];
    const recMap = {};
    const order = [];
    for (let ri = 1; ri < rows.length; ri++) {
      const r = rows[ri];
      const recordId = String(r[0] || "").trim();
      if (!recordId) continue;
      if (String(r[1] || "").trim() !== equipmentId) continue;
      const d = r[5] instanceof Date ? r[5] : new Date(r[5]);
      if (isNaN(d.getTime())) continue;
      if (Utilities.formatDate(d, "Asia/Almaty", "yyyy-MM-dd") !== dateStr) continue;

      if (!recMap[recordId]) {
        recMap[recordId] = {
          recordId: recordId, equipmentId: equipmentId,
          empId: r[3], empName: String(r[4] || ""), date: d.toISOString(),
          items: [],
        };
        order.push(recordId);
      }
      recMap[recordId].items.push({
        section: String(r[6] || ""), itemText: String(r[7] || ""),
        status: r[8], comment: String(r[9] || ""),
      });
    }
    order.forEach(id => checklists.push(recMap[id]));
    checklists.sort((a, b) => a.date.localeCompare(b.date));
  }

  // ── Заправки ──
  const fuelSheet = ss.getSheetByName(SHEET_FUEL_RECORDS);
  const fuel = [];
  if (fuelSheet) {
    const rows = fuelSheet.getDataRange().getValues();
    if (rows.length > 1) {
      const headers = rows[0];
      const iEq = headers.indexOf("equipmentId"), iDate = headers.indexOf("date");
      rows.slice(1).forEach(r => {
        if (String(r[iEq] || "").trim() !== equipmentId) return;
        const raw = r[iDate];
        const d = raw instanceof Date ? raw : new Date(raw);
        if (isNaN(d.getTime())) return;
        if (Utilities.formatDate(d, "Asia/Almaty", "yyyy-MM-dd") !== dateStr) return;
        const o = {}; headers.forEach((h, i) => { o[h] = r[i]; });
        o.date = d.toISOString();
        fuel.push(o);
      });
      fuel.sort((a, b) => a.date.localeCompare(b.date));
    }
  }

  return json({ ok: true, equipmentId: equipmentId, date: dateStr, checklists: checklists, fuel: fuel });
}

function getChecklistRecords(days) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CHECKLIST_RECORDS);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);

  const nDays = Number(days) || 0;
  let since = null;
  if (nDays > 0) {
    since = new Date();
    since.setDate(since.getDate() - nDays);
    since.setHours(0, 0, 0, 0);
  }

  const recordsMap = {};
  const order = [];

  rows.slice(1).forEach(r => {
    const recordId = String(r[0] || "").trim();
    if (!recordId) return;

    if (since) {
      const d = r[5] instanceof Date ? r[5] : new Date(r[5]);
      if (!(d instanceof Date) || isNaN(d.getTime()) || d < since) return;
    }

    if (!recordsMap[recordId]) {
      recordsMap[recordId] = {
        recordId,
        equipmentId: r[1], templateId: r[2], empId: r[3], empName: r[4], date: r[5],
        items: [],
      };
      order.push(recordId);
    }
    const status  = r[8];
    const comment = r[9];
    const rec = recordsMap[recordId];
    rec.totalItems = (rec.totalItems || 0) + 1;

    // Передаём только значимые пункты: неисправности и всё, где оператор
    // что-то написал. Исправные без комментария сворачиваем в счётчик —
    // их 90% от объёма, а для аналитики важно лишь их количество.
    // Это уменьшает выгрузку примерно в 9 раз (7 МБ -> 0.8 МБ).
    if (String(status) !== "ok" || String(comment || "").trim()) {
      rec.items.push({ section: r[6], itemText: r[7], status: status, comment: comment });
    } else {
      rec.okCount = (rec.okCount || 0) + 1;
    }
  });

  return json(order.reverse().map(id => {
    const rec = recordsMap[id];
    rec.okCount    = rec.okCount || 0;
    rec.totalItems = rec.totalItems || 0;
    return rec;
  }));
}

// ══════════════════════════════════════════════════════════
// ЗАПРАВКИ
// ══════════════════════════════════════════════════════════
const SHEET_FUEL_RECORDS = "Заправки";

function submitFuelRecord(p) {
  const { equipmentId, empId, empName, amountLiters } = p || {};
  if (!equipmentId || !empName || !amountLiters) {
    return json({ ok: false, error: "Нужны equipmentId, empName и amountLiters" });
  }

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_FUEL_RECORDS);
  const HEADER = ["recordId","equipmentId","empId","empName","date","amountLiters"];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FUEL_RECORDS);
    sheet.appendRow(HEADER);
    sheet.getRange(1,1,1,HEADER.length)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  const recordId = "fuel" + Date.now();
  const dateStr = Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy HH:mm");
  ensureCapacity(sheet, 1);
  sheet.appendRow([recordId, equipmentId, empId || "", empName, dateStr, amountLiters]);

  return json({ ok: true, recordId });
}

function getFuelRecords(site, days) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_FUEL_RECORDS);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const siteFilter = site ? String(site).trim() : "";
  // days необязателен — без него отдаём весь период, как раньше
  // (instructor.html дёргает этот же эндпоинт без параметра).
  const daysNum = days ? parseInt(days, 10) : 0;
  let sinceDate = null;
  if (daysNum > 0) {
    sinceDate = new Date();
    sinceDate.setDate(sinceDate.getDate() - daysNum);
    sinceDate.setHours(0, 0, 0, 0);
  }

  // У «Заправки» своей колонки site нет — участок узнаём через
  // equipmentId по листу «Техника» (тот же приём, что в getChecklistStats).
  let eqSite = null;
  if (siteFilter) {
    eqSite = {};
    const eqSheet = ss.getSheetByName(SHEET_EQUIPMENT);
    if (eqSheet) {
      const er = eqSheet.getDataRange().getValues();
      const eh = er[0];
      const iId = eh.indexOf("equipmentId"), iSite = eh.indexOf("site");
      er.slice(1).forEach(r => {
        const id = String(r[iId] || "").trim();
        if (id) eqSite[id] = String(r[iSite] || "").trim();
      });
    }
  }

  const eqCol = headers.indexOf("equipmentId");
  const dateCol = headers.indexOf("date");
  const data = rows.slice(1).filter(r => r[0])
    .filter(r => !siteFilter || eqSite[String(r[eqCol] || "").trim()] === siteFilter)
    .filter(r => {
      if (!sinceDate) return true;
      const raw = r[dateCol];
      const d = raw instanceof Date ? raw : new Date(raw);
      return !isNaN(d) && d >= sinceDate;
    })
    .map(r => {
      const o = {}; headers.forEach((h,i) => { o[h] = r[i]; }); return o;
    });
  return json(data.reverse());
}

// ══════════════════════════════════════════════════════
// СМЕНЫ — привязка операторов к технике по вахте/смене
// ══════════════════════════════════════════════════════
const SHEET_SHIFTS = "Смены";

// ══════════════════════════════════════════════════════
// АКТИВНАЯ ВАХТА ПО УЧАСТКУ
// Переключатель "Вахта-1 / Вахта-2" в master.html раньше был чисто
// локальным состоянием экрана — сбрасывался при перезагрузке страницы и
// нигде не хранился. Теперь горный мастер переключает его на сервере,
// чтобы другие панели (instructor.html — дисциплина по чек-листам и
// топливу) знали, кто реально сейчас физически на площадке, а не путали
// с ростером вахты, которая в этот момент отдыхает. Хранится в отдельном
// листе "Настройки": участок → активная вахта. Добавлено 24.08.2026.
// ══════════════════════════════════════════════════════
const SHEET_SETTINGS = "Настройки";

function getActiveVahtas() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  const out = {};
  if (sheet) {
    const rows = sheet.getDataRange().getValues();
    if (rows.length > 1) {
      const headers = rows[0];
      const iSite = headers.indexOf("site"), iVahta = headers.indexOf("activeVahta");
      rows.slice(1).forEach(r => {
        const site = String(r[iSite]||"").trim();
        if (site) out[site] = String(r[iVahta]||"").trim();
      });
    }
  }
  return json(out);
}

function setActiveVahta(p) {
  const { site, vahta } = p || {};
  if (!site || !vahta) return json({ ok:false, error:"Нужны site и vahta" });

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_SETTINGS);
  const HEADER = ["site","activeVahta","updatedAt"];
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SETTINGS);
    sheet.appendRow(HEADER);
    sheet.getRange(1,1,1,HEADER.length)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const iSite = headers.indexOf("site"), iVahta = headers.indexOf("activeVahta"), iUpd = headers.indexOf("updatedAt");
  const nowStr = Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy HH:mm");

  let rowIdx = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][iSite]||"").trim() === String(site).trim()) { rowIdx = i + 1; break; }
  }

  if (rowIdx > 0) {
    sheet.getRange(rowIdx, iVahta+1).setNumberFormat("@").setValue(vahta);
    sheet.getRange(rowIdx, iUpd+1).setNumberFormat("@").setValue(nowStr);
  } else {
    ensureCapacity(sheet, 1);
    sheet.appendRow([site, vahta, nowStr]);
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 2, 1, 2).setNumberFormat("@");
  }

  return json({ ok:true, site, vahta });
}

// ══════════════════════════════════════════════════════
// СТАТУС ТЕХНИКИ ЧЕРЕЗ QR — работник в checklist.html сам ставит "В работе"/
// "Ремонт" (автоматически при сдаче чек-листа, или вручную третьей плашкой
// для ремонта). В отличие от workStatus (статус РАБОТНИКА, ставит мастер в
// master.html), это статус ТЕХНИКИ — они разделены (26.08.2026), потому что
// характеризуют разные вещи и должны меняться независимо.
//
// Один серверный вызов делает всё сам: находит участок техники (по листу
// "Техника"), активную вахту участка (по листу "Настройки"), смену — по
// текущему времени Алматы (07:00–19:00 = день). Если для участка активная
// вахта не задана (пока это Ашыктас) — возвращает unavailable:true, а не
// ошибку; checklist.html должен на это не ругаться, а тихо промолчать.
// Если подходящей строки в "Смены" ещё нет — создаёт минимальную (только
// слот + статус, без сотрудника — назначением сотрудников занимается мастер
// отдельно через master.html, тут его сфера не расширяем).
// ══════════════════════════════════════════════════════
function setEquipmentStatus(p) {
  const { equipmentId, status } = p || {};
  if (!equipmentId || !status) return json({ ok: false, error: "Нужны equipmentId и status" });

  const ss = SpreadsheetApp.openById(SHEET_ID);

  const eqSheet = ss.getSheetByName(SHEET_EQUIPMENT);
  if (!eqSheet) return json({ ok: false, error: "Лист «Техника» не найден" });
  const eqRows = eqSheet.getDataRange().getValues();
  const eqHeaders = eqRows[0];
  const eqIdCol = eqHeaders.indexOf("equipmentId"), eqSiteCol = eqHeaders.indexOf("site");
  let site = "";
  for (let i = 1; i < eqRows.length; i++) {
    if (String(eqRows[i][eqIdCol]||"").trim() === String(equipmentId).trim()) {
      site = String(eqRows[i][eqSiteCol]||"").trim();
      break;
    }
  }
  if (!site) return json({ ok: false, error: "Техника не найдена" });

  const settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  let activeVahta = "";
  if (settingsSheet) {
    const rows = settingsSheet.getDataRange().getValues();
    if (rows.length > 1) {
      const headers = rows[0];
      const iSite = headers.indexOf("site"), iVahta = headers.indexOf("activeVahta");
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][iSite]||"").trim() === site) { activeVahta = String(rows[i][iVahta]||"").trim(); break; }
      }
    }
  }
  if (!activeVahta) return json({ ok: false, unavailable: true, error: "Активная вахта для участка «"+site+"» не задана" });

  const hour = Number(Utilities.formatDate(new Date(), "Asia/Almaty", "H"));
  const smena = (hour >= 7 && hour < 19) ? 1 : 2;

  const shiftSheet = ss.getSheetByName(SHEET_SHIFTS);
  if (!shiftSheet) return json({ ok: false, error: "Лист «Смены» не найден" });
  const shiftRows = shiftSheet.getDataRange().getValues();
  const shiftHeaders = shiftRows[0];
  const col = h => shiftHeaders.indexOf(h);
  const cEq = col("equipmentId"), cSm = col("smena"), cVah = col("vahta"), cSite = col("site"), cEqStatus = col("equipmentStatus");
  if (cEqStatus < 0) return json({ ok: false, error: "Столбец equipmentStatus не найден на листе «Смены» — добавь его вручную как в saveShiftAssignment" });

  let rowIdx = -1;
  for (let i = 1; i < shiftRows.length; i++) {
    if (String(shiftRows[i][cEq]||"").trim() === String(equipmentId).trim()
        && Number(shiftRows[i][cSm]) === smena
        && String(shiftRows[i][cVah]||"").trim() === activeVahta
        && String(shiftRows[i][cSite]||"").trim() === site) { rowIdx = i + 1; break; }
  }

  if (rowIdx > 0) {
    shiftSheet.getRange(rowIdx, cEqStatus + 1).setValue(status);
    const afterRow = shiftSheet.getRange(rowIdx, 1, 1, shiftHeaders.length).getValues()[0];
    const snapshot = {};
    shiftHeaders.forEach((h, i) => { snapshot[h] = afterRow[i]; });
    logShiftEvents([{ ...snapshot, action: "updated" }]);
  } else {
    ensureCapacity(shiftSheet, 1);
    const newRow = new Array(shiftHeaders.length).fill("");
    newRow[cEq] = equipmentId; newRow[cSm] = smena; newRow[cVah] = activeVahta; newRow[cSite] = site; newRow[cEqStatus] = status;
    shiftSheet.appendRow(newRow);
    const snapshot = {};
    shiftHeaders.forEach((h, i) => { snapshot[h] = newRow[i]; });
    logShiftEvents([{ ...snapshot, action: "created" }]);
  }

  return json({ ok: true, site, vahta: activeVahta, smena, status });
}

function getShiftAssignments(site) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SHIFTS);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const siteCol = headers.indexOf("site");
  const siteFilter = site ? String(site).trim() : "";
  const data = [];
  rows.slice(1).forEach((r, i) => {
    if (!r[0] && !r[1] && !r[4]) return;
    if (siteFilter && (siteCol < 0 || String(r[siteCol] || "").trim() !== siteFilter)) return;
    const o = {};
    headers.forEach((h,j) => { o[h] = r[j]; });
    o.__row = i + 2;
    data.push(o);
  });
  return json(data);
}

// ══════════════════════════════════════════════════════
// ЖУРНАЛ ИЗМЕНЕНИЙ СМЕН — история назначений по слотам (техника+смена+вахта+
// участок) для восстановления "наряда на дату X" за прошедшие дни. Раньше
// назначения были только живым состоянием (см. открытый вопрос №3 в
// оперативной базе) — при каждом изменении старое значение просто исчезало.
// Теперь на каждую мутацию листа "Смены" (создание/правка/своп день-ночь/
// удаление) пишется ОТДЕЛЬНАЯ строка-снимок нового состояния слота сюда.
// Реконструкция дня X = для каждого слота взять последнюю запись журнала
// с timestamp <= конец дня X. Важно: работает только ВПЕРЁД с момента
// добавления (24.08.2026) — восстановить историю за уже прошедшие дни
// нельзя, для них данных просто не существует.
const SHEET_SHIFT_LOG = "ИсторияСмен";
const SHIFT_LOG_HEADERS = ["timestamp","equipmentId","position","smena","vahta","site","empId","empName","status","dismissDate","skill","workStatus","reassignNote","action"];

// Баг 26.08.2026: при быстрых последовательных вызовах первый запрос
// создавал лист (insertSheet), но заголовок ("appendRow" сразу следом)
// иногда не успевал зафиксироваться до возврата функции — а все
// последующие вызовы видели через getSheetByName, что лист уже существует,
// и пропускали инициализацию заголовка целиком (она была привязана только
// к ветке "лист не найден"). Итог: данные писались, но без шапки, и место
// для заголовка навсегда пропадало. Исправлено: 1) заголовок пишется по
// признаку "лист пуст" (getLastRow()===0), а не "лист только что создан";
// 2) LockService сериализует конкурентные вызовы — без него getLastRow()+1
// для двух одновременных запросов может вернуть одно и то же число, и один
// перезапишет другого.
function logShiftEvents(entries) {
  if (!entries || !entries.length) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_SHIFT_LOG);
    if (!sheet) sheet = ss.insertSheet(SHEET_SHIFT_LOG);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(SHIFT_LOG_HEADERS);
      sheet.getRange(1, 1, 1, SHIFT_LOG_HEADERS.length)
        .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
    const now = new Date();
    const rows = entries.map(e => SHIFT_LOG_HEADERS.map(h => {
      if (h === "timestamp") return e.timestamp || now;
      return (e[h] !== undefined && e[h] !== null) ? e[h] : "";
    }));
    ensureCapacity(sheet, rows.length);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, SHIFT_LOG_HEADERS.length).setValues(rows);
  } finally {
    lock.releaseLock();
  }
}

function saveShiftAssignment(p) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_SHIFTS);
  // workStatus/reassignNote — оперативный статус РАБОТНИКА на смену (вакансия/
  // командировка/хозработы и т.п.) + свободная заметка о временной переброске.
  // Добавлены 22.08.2026 взамен бумажного наряда, который мастера заполняли вручную
  // каждую смену. equipmentStatus — статус ТЕХНИКИ (в работе/ремонт), отдельное
  // поле от workStatus: ставит сам работник через QR в checklist.html, а не мастер.
  // Разделены 26.08.2026 — раньше были одним общим списком, но "в работе"/"ремонт"
  // характеризуют машину, а не человека, который на ней сейчас числится.
  // ВАЖНО: если лист «Смены» уже существует (обычный случай), новые столбцы нужно
  // добавить туда вручную как последние — иначе появятся не в тех колонках. Если
  // листа ещё нет — appendRow создаст его сразу со всеми этими столбцами.
  const headers = ["equipmentId","position","smena","vahta","empId","empName","status","dismissDate","skill","site","workStatus","reassignNote","equipmentStatus"];
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SHIFTS);
    sheet.appendRow(headers);
  }
  ensureCapacity(sheet, 1);
  sheet.appendRow(headers.map(h => (p[h] !== undefined && p[h] !== null) ? p[h] : ""));
  logShiftEvents([{ ...Object.fromEntries(headers.map(h => [h, p[h] !== undefined && p[h] !== null ? p[h] : ""])), action: "created" }]);
  return json({ ok: true });
}

// Изменить существующее назначение по номеру строки. equipmentId тоже можно
// менять — нужно при переименовании техники, чтобы записи смен не осиротели.
function updateShiftAssignment(p) {
  const { row, equipmentId, empId, empName, status, dismissDate, skill, workStatus, reassignNote, equipmentStatus } = p || {};
  if (!row) return json({ ok: false, error: "Нужен row" });
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SHIFTS);
  if (!sheet) return json({ ok: false, error: "Лист «Смены» не найден" });
  const headers = sheet.getDataRange().getValues()[0];
  const col = h => headers.indexOf(h) + 1;
  if (equipmentId !== undefined && col("equipmentId") > 0) sheet.getRange(row, col("equipmentId")).setValue(equipmentId);
  if (empId       !== undefined && col("empId")       > 0) sheet.getRange(row, col("empId")).setValue(empId);
  if (empName     !== undefined && col("empName")     > 0) sheet.getRange(row, col("empName")).setValue(empName);
  if (status      !== undefined && col("status")      > 0) sheet.getRange(row, col("status")).setValue(status);
  if (dismissDate !== undefined && col("dismissDate") > 0) sheet.getRange(row, col("dismissDate")).setValue(dismissDate);
  if (skill       !== undefined && col("skill")       > 0) sheet.getRange(row, col("skill")).setValue(skill);
  // Пишутся, только если такие столбцы реально есть на листе (см. комментарий
  // в saveShiftAssignment) — на старом листе без них просто ничего не произойдёт,
  // без ошибки.
  if (workStatus  !== undefined && col("workStatus")  > 0) sheet.getRange(row, col("workStatus")).setValue(workStatus);
  if (reassignNote!== undefined && col("reassignNote")> 0) sheet.getRange(row, col("reassignNote")).setValue(reassignNote);
  if (equipmentStatus !== undefined && col("equipmentStatus") > 0) sheet.getRange(row, col("equipmentStatus")).setValue(equipmentStatus);

  // Полный снимок строки ПОСЛЕ изменения — не только то, что поменялось в
  // этом вызове, иначе журнал будет неполным (например, при правке только
  // workStatus в записи журнала выпадет empName, хотя он не менялся).
  const afterRow = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const snapshot = {};
  headers.forEach((h, i) => { snapshot[h] = afterRow[i]; });
  logShiftEvents([{ ...snapshot, action: "updated" }]);

  return json({ ok: true });
}

// Удалить назначение целиком (строку)
function deleteShiftAssignmentRow(p) {
  const { row } = p || {};
  if (!row) return json({ ok: false, error: "Нужен row" });
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SHIFTS);
  if (!sheet) return json({ ok: false, error: "Лист «Смены» не найден" });

  // Снимок ДО удаления — иначе журнал не будет знать, какой слот пропал.
  const headers = sheet.getDataRange().getValues()[0];
  const beforeRow = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const snapshot = {};
  headers.forEach((h, i) => { snapshot[h] = beforeRow[i]; });

  sheet.deleteRow(row);
  logShiftEvents([{ ...snapshot, action: "deleted" }]);
  return json({ ok: true });
}

// Новый заезд: вся вахта на участке меняется местами день↔ночь одним действием.
// Для каждой единицы техники меняем местами человека (и его status/dismissDate/skill)
// между строкой смены=1 (день) и смены=2 (ночь) — сам "слот" (equipmentId/position)
// остаётся на месте, местами меняются только люди в нём.
// Если для техники существует только одна из двух строк (день ИЛИ ночь) — вторую
// создаём заново с этими данными, а исходную опустошаем, чтобы человек физически
// переехал на другую смену, а не продублировался.
// Пишем изменения ОДНИМ batch-запросом (не по одной ячейке) — это может быть
// 100+ единиц техники, построчная запись на слабой связи не доедет.
function swapDayNight(p) {
  const { site, vahta } = p || {};
  const siteFilter = String(site||"").trim();
  const vahtaFilter = String(vahta||"").trim();
  if (!siteFilter || !vahtaFilter) return json({ ok: false, error: "Нужны site и vahta" });

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SHIFTS);
  if (!sheet) return json({ ok: false, error: "Лист «Смены» не найден" });

  const range = sheet.getDataRange();
  const rows  = range.getValues();
  if (rows.length < 2) return json({ ok: true, swapped: 0, created: 0 });
  const headers = rows[0];
  const col = h => headers.indexOf(h);
  const cEq=col("equipmentId"), cPos=col("position"), cSm=col("smena"), cVah=col("vahta"),
        cEmpId=col("empId"), cEmpName=col("empName"), cStatus=col("status"),
        cDismiss=col("dismissDate"), cSkill=col("skill"), cSite=col("site");

  // Снимок строки для журнала — по индексам колонок листа, а не по порядку в
  // SHIFT_LOG_HEADERS, чтобы не разъехаться, если когда-нибудь поменяется
  // порядок столбцов на листе.
  function rowToLogEntry(rowArr, action) {
    const e = { action };
    headers.forEach((h, i) => { e[h] = rowArr[i]; });
    return e;
  }

  // Группируем строки нужного участка+вахты по технике: день/ночь -> индекс в rows
  const groups = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[cSite]||"").trim() !== siteFilter) continue;
    if (String(r[cVah]||"").trim() !== vahtaFilter) continue;
    const eq = String(r[cEq]||"").trim();
    if (!eq) continue;
    if (!groups[eq]) groups[eq] = {};
    const sm = Number(r[cSm]);
    if (sm === 1) groups[eq].day = i;
    else if (sm === 2) groups[eq].night = i;
  }

  const personOf = (idx) => idx === undefined ? { empId:"", empName:"", status:"", dismissDate:"", skill:"" } : {
    empId: rows[idx][cEmpId], empName: rows[idx][cEmpName],
    status: rows[idx][cStatus], dismissDate: rows[idx][cDismiss], skill: rows[idx][cSkill],
  };
  const applyPerson = (idx, person) => {
    rows[idx][cEmpId] = person.empId; rows[idx][cEmpName] = person.empName;
    rows[idx][cStatus] = person.status; rows[idx][cDismiss] = person.dismissDate; rows[idx][cSkill] = person.skill;
  };
  const buildRow = (eq, position, smena, person) => {
    const r = new Array(headers.length).fill("");
    r[cEq]=eq; r[cPos]=position||""; r[cSm]=smena; r[cVah]=vahtaFilter; r[cSite]=siteFilter;
    r[cEmpId]=person.empId; r[cEmpName]=person.empName; r[cStatus]=person.status;
    r[cDismiss]=person.dismissDate; r[cSkill]=person.skill;
    return r;
  };

  const newRows = [];
  let swapped = 0;
  const logEntries = [];

  Object.entries(groups).forEach(([eq, g]) => {
    if (g.day === undefined && g.night === undefined) return;
    const dayPerson   = personOf(g.day);
    const nightPerson = personOf(g.night);
    const hadAny = dayPerson.empId || dayPerson.empName || nightPerson.empId || nightPerson.empName;
    if (!hadAny) return; // оба слота пусты — менять нечего

    if (g.day !== undefined) {
      applyPerson(g.day, nightPerson);
    } else if (nightPerson.empId || nightPerson.empName) {
      const nr = buildRow(eq, rows[g.night][cPos], 1, nightPerson);
      newRows.push(nr);
      logEntries.push(rowToLogEntry(nr, "created"));
    }
    if (g.night !== undefined) {
      applyPerson(g.night, dayPerson);
    } else if (dayPerson.empId || dayPerson.empName) {
      const nr = buildRow(eq, rows[g.day][cPos], 2, dayPerson);
      newRows.push(nr);
      logEntries.push(rowToLogEntry(nr, "created"));
    }
    // Существующие строки логируем ПОСЛЕ applyPerson выше — rows[idx] уже
    // содержит новое (переставленное) состояние на этот момент.
    if (g.day !== undefined) logEntries.push(rowToLogEntry(rows[g.day], "swapped"));
    if (g.night !== undefined) logEntries.push(rowToLogEntry(rows[g.night], "swapped"));
    swapped++;
  });

  // Одна запись всего изменённого диапазона разом — не по ячейке за раз
  range.setValues(rows);
  if (newRows.length) {
    ensureCapacity(sheet, newRows.length);
    sheet.getRange(sheet.getLastRow()+1, 1, newRows.length, headers.length).setValues(newRows);
  }
  logShiftEvents(logEntries);

  return json({ ok: true, swapped, created: newRows.length });
}

// Полностью очищает лист «Смены» (кроме заголовка) — перед чистой переимпортацией
function clearShiftAssignments() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SHIFTS);
  if (!sheet) return json({ ok: true, cleared: 0 });
  const lastRow = sheet.getLastRow();
  const cleared = Math.max(0, lastRow - 1);
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  return json({ ok: true, cleared });
}

// ══════════════════════════════════════════════════════════
// ОЦЕНКА РАБОТЫ ОПЕРАТОРА
// Три блока критериев (по образцу утверждённого чек-листа
// «Машинист буровой установки v9»):
//   Блок 1 «Обязанности оператора» — общий для ВСЕХ профессий (equipmentType
//   в каталоге = "ВСЕ"), шкала 0/1/"-", итог блока — % выполненных пунктов,
//   переведённый в 1-5 по порогам (совпадают с эталонным файлом).
//   Блок 2 «Состояние техники» и Блок 3 «Эксплуатация» — свои пункты под
//   каждый тип техники (equipmentType в каталоге = конкретный тип), шкала
//   1-5/"-", итог блока — среднее.
// Итоговой свёртки трёх блоков в одно число НЕТ (осознанно, по решению
// пользователя) — три балла хранятся раздельно, плюс текстовое резюме.
// Оценка может быть любой, включая целиком отрицательную — сервер ничего
// не сглаживает и не подтягивает к среднему.
// ══════════════════════════════════════════════════════════

// Пороги перевода % выполненных пунктов блока 1 в шкалу 1-5.
// Взяты из эталонного файла (лист «Списки» J2:J6) как фиксированные —
// если понадобится их менять по профессиям, выносить в отдельный лист
// настроек, но пока для всех профессий они общие.
const EVAL_BLOCK1_THRESHOLDS = [
  { min: 0.85, score: 5 },
  { min: 0.75, score: 4 },
  { min: 0.65, score: 3 },
  { min: 0.55, score: 2 },
  { min: 0,    score: 1 },
];

function computeEvalBlock1Score(values) {
  const evaluated = values.filter(v => v === "1" || v === "0" || v === 1 || v === 0);
  if (!evaluated.length) return "-";
  const ones = evaluated.filter(v => String(v) === "1").length;
  const pct = ones / evaluated.length;
  for (const t of EVAL_BLOCK1_THRESHOLDS) {
    if (pct >= t.min) return t.score;
  }
  return 1;
}

function computeEvalAvgScore(values) {
  const evaluated = values
    .filter(v => v !== "-" && v !== "" && v !== null && v !== undefined)
    .map(Number)
    .filter(n => !isNaN(n));
  if (!evaluated.length) return "-";
  const avg = evaluated.reduce((a, b) => a + b, 0) / evaluated.length;
  return Math.round(avg * 100) / 100;
}

// ── Каталог критериев ────────────────────────────────────────
// Лист: equipmentType | block | itemNum | itemText | scale (binary/scale5) |
//       relatedCourse | active
// equipmentType = "ВСЕ" — пункты блока 1, общие для всех профессий.
// equipmentType = конкретный тип — пункты блоков 2 и 3 под эту профессию.
// relatedCourse — необязательное имя курса из «КаталогКурсов»: если по
// этому пункту оценка низкая, соответствующий курс автоматически попадает
// в «ПланОбучения» (см. saveEvaluation).
function getEvaluationCriteria(equipmentType) {
  const type = String(equipmentType || "").trim();
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EVAL_CRITERIA);
  if (!sheet) return json({ ok: true, block1: [], block2: [], block3: [] });

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json({ ok: true, block1: [], block2: [], block3: [] });
  const headers = rows[0];
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);

  const out = { ok: true, block1: [], block2: [], block3: [] };
  rows.slice(1).forEach(r => {
    if (!r[idx["equipmentType"]]) return;
    const active = String(r[idx["active"]] || "да").trim().toLowerCase();
    if (active === "нет" || active === "false") return;
    const rowType = String(r[idx["equipmentType"]]).trim();
    if (rowType !== EVAL_UNIVERSAL_TYPE && rowType !== type) return;

    const block = String(r[idx["block"]] || "").trim();
    const item = {
      itemNum: String(r[idx["itemNum"]] || "").trim(),
      itemText: String(r[idx["itemText"]] || "").trim(),
      scale: String(r[idx["scale"]] || "").trim() || (block === "1" ? "binary" : "scale5"),
      relatedCourse: idx["relatedCourse"] >= 0 ? String(r[idx["relatedCourse"]] || "").trim() : "",
    };
    if (block === "1") out.block1.push(item);
    else if (block === "2") out.block2.push(item);
    else if (block === "3") out.block3.push(item);
  });

  const byItemNum = (a,b) => a.itemNum.localeCompare(b.itemNum, undefined, { numeric: true });
  out.block1.sort(byItemNum);
  out.block2.sort(byItemNum);
  out.block3.sort(byItemNum);

  return json(out);
}

// Список типов техники/профессий, под которые уже есть свои пункты
// блоков 2-3 в каталоге (для выпадающего списка в форме оценки).
function getEvaluationEquipmentTypes() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EVAL_CRITERIA);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const typeCol = headers.indexOf("equipmentType");
  if (typeCol < 0) return json([]);

  const types = new Set();
  rows.slice(1).forEach(r => {
    const t = String(r[typeCol] || "").trim();
    if (t && t !== EVAL_UNIVERSAL_TYPE) types.add(t);
  });
  return json(Array.from(types).sort());
}

// Подсказки для поля «Вид выполняемых работ» — свободный текст, не жёсткий
// каталог (в отличие от equipmentType), поэтому просто отдаём уникальные
// значения, которые уже вводили раньше, для датлиста на фронтенде.
function getEvaluationWorkTypes() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const col = headers.indexOf("workType");
  if (col < 0) return json([]);

  const values = new Set();
  rows.slice(1).forEach(r => {
    const v = String(r[col] || "").trim();
    if (v) values.add(v);
  });
  return json(Array.from(values).sort());
}

// Подсказки для поля «Модель техники» — тоже свободный текст, из уже
// сохранённых оценок. equipmentType (необязательно) сужает список: для
// буровой установки предлагаем модели буровых, а не самосвалов.
function getEvaluationEquipmentModels(p) {
  const typeFilter = (p && p.equipmentType) ? String(p.equipmentType).trim() : "";
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const modelCol = headers.indexOf("equipmentModel");
  const typeCol = headers.indexOf("equipmentType");
  if (modelCol < 0) return json([]);

  const values = new Set();
  rows.slice(1).forEach(r => {
    if (typeFilter && typeCol >= 0 && String(r[typeCol] || "").trim() !== typeFilter) return;
    const v = String(r[modelCol] || "").trim();
    if (v) values.add(v);
  });
  return json(Array.from(values).sort());
}

// Общий помощник: уникальные непустые значения одной колонки листа
// «ОценкиОператоров» — используется и для инструктора, и для супервайзера,
// чтобы не дублировать одну и ту же логику дважды.
function getEvaluationDistinctColumn(columnName) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const col = headers.indexOf(columnName);
  if (col < 0) return json([]);

  const values = new Set();
  rows.slice(1).forEach(r => {
    const v = String(r[col] || "").trim();
    if (v) values.add(v);
  });
  return json(Array.from(values).sort());
}

function getEvaluationInstructors() {
  return getEvaluationDistinctColumn("instructorName");
}

function getEvaluationSupervisors() {
  return getEvaluationDistinctColumn("supervisorName");
}

// Заменяет весь набор пунктов конкретного типа техники разом (как
// importChecklistTemplate) — используется при заведении нового типа
// или полной переработке существующего.
// items: [{ block: "2"|"3", itemNum: "2.1", itemText, scale, relatedCourse }]
// Для типа "ВСЕ" (блок 1) — тем же способом, отдельным вызовом.
function importEvaluationCriteria(p) {
  const { equipmentType, items } = p || {};
  const type = String(equipmentType || "").trim();
  if (!type || !Array.isArray(items) || !items.length) {
    return json({ ok: false, error: "Нужны equipmentType и непустой список items" });
  }

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_EVAL_CRITERIA);
  const HEADER = ["equipmentType","block","itemNum","itemText","scale","relatedCourse","active"];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_EVAL_CRITERIA);
    sheet.appendRow(HEADER);
    sheet.getRange(1,1,1,HEADER.length)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(4, 400);
  }

  const rows = sheet.getDataRange().getValues();
  const body = rows.length > 1 ? rows.slice(1) : [];
  const kept = body.filter(r => String(r[0] || "").trim() !== type);

  const newRows = items.map(it => [
    type,
    String(it.block || ""),
    String(it.itemNum || ""),
    String(it.itemText || ""),
    String(it.scale || (String(it.block) === "1" ? "binary" : "scale5")),
    String(it.relatedCourse || ""),
    "да",
  ]);

  const finalRows = kept.concat(newRows);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, HEADER.length).clearContent();
  if (finalRows.length > 0) {
    // itemNum вида "1.1"/"2.3" Sheets самостоятельно интерпретирует как дату
    // (день.месяц) при записи через setValues — тот же баг, что был с
    // experienceYears. Фиксируем колонку C текстовым форматом ДО записи,
    // иначе она один раз успевает неверно истолковать значение.
    const itemNumCol = HEADER.indexOf("itemNum") + 1;
    sheet.getRange(2, itemNumCol, finalRows.length, 1).setNumberFormat("@");
    sheet.getRange(2, 1, finalRows.length, HEADER.length).setValues(finalRows);
  }

  return json({ ok: true, equipmentType: type, items: newRows.length });
}

// ── Сохранение заполненной оценки ────────────────────────────
// p: { empId, empName, position, date, equipmentType, equipmentModel, site,
//      instructorName, supervisorName, workType, summary,
//      answers: [{ block, itemNum, itemText, value, comment }] }
// Баллы блоков считаются на сервере из ответов — не доверяем клиенту,
// чтобы рейтинг и личное дело всегда были согласованы с фактическими
// отметками, даже если фронтенд посчитал что-то иначе.
function saveEvaluation(p) {
  const { empId, empName, position, date, equipmentType, equipmentModel, site,
          instructorName, supervisorName, workType, summary, answers } = p || {};

  if (!empName || !equipmentType || !workType || !Array.isArray(answers) || !answers.length) {
    return json({ ok: false, error: "Нужны как минимум empName, equipmentType, workType (вид выполняемых работ) и ответы по пунктам" });
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Баллы по блокам
  const b1values = answers.filter(a => String(a.block) === "1").map(a => a.value);
  const b2values = answers.filter(a => String(a.block) === "2").map(a => a.value);
  const b3values = answers.filter(a => String(a.block) === "3").map(a => a.value);
  const block1Score = computeEvalBlock1Score(b1values);
  const block2Score = computeEvalAvgScore(b2values);
  const block3Score = computeEvalAvgScore(b3values);

  // Заголовок оценки
  let evalSheet = ss.getSheetByName(SHEET_EVALUATIONS);
  const EVAL_HEADER = ["evalId","empId","empName","position","date","equipmentType","equipmentModel",
                       "site","instructorName","supervisorName","workType",
                       "block1Score","block2Score","block3Score","summary","createdAt"];
  if (!evalSheet) {
    evalSheet = ss.insertSheet(SHEET_EVALUATIONS);
    evalSheet.appendRow(EVAL_HEADER);
    evalSheet.getRange(1,1,1,EVAL_HEADER.length)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    evalSheet.setFrozenRows(1);
    evalSheet.setColumnWidth(3, 200);
    evalSheet.setColumnWidth(15, 400);
  }

  const evalId = "eval" + Date.now();
  const dateStr = date || Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy");
  const createdAt = Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy HH:mm");

  ensureCapacity(evalSheet, 1);
  evalSheet.appendRow([
    evalId, empId || "", empName, position || "", dateStr, equipmentType, equipmentModel || "",
    site || "", instructorName || "", supervisorName || "", workType || "",
    block1Score, block2Score, block3Score, summary || "", createdAt,
  ]);

  // Детализация по пунктам
  let ansSheet = ss.getSheetByName(SHEET_EVAL_ANSWERS);
  const ANS_HEADER = ["evalId","block","itemNum","itemText","value","comment"];
  if (!ansSheet) {
    ansSheet = ss.insertSheet(SHEET_EVAL_ANSWERS);
    ansSheet.appendRow(ANS_HEADER);
    ansSheet.getRange(1,1,1,ANS_HEADER.length)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    ansSheet.setFrozenRows(1);
    ansSheet.setColumnWidth(4, 400);
  }
  const ansRows = answers.map(a => [
    evalId, String(a.block || ""), String(a.itemNum || ""), String(a.itemText || ""),
    a.value === undefined || a.value === null ? "" : String(a.value), a.comment || "",
  ]);
  ensureCapacity(ansSheet, ansRows.length);
  // itemNum вида "1.1"/"2.3" — та же ловушка с автоопределением даты,
  // что и в КаталогКритериевОценки (см. importEvaluationCriteria). Текстовый
  // формат ставим ДО записи значений.
  const ansItemNumCol = ANS_HEADER.indexOf("itemNum") + 1;
  const ansStartRow = ansSheet.getLastRow() + 1;
  ansSheet.getRange(ansStartRow, ansItemNumCol, ansRows.length, 1).setNumberFormat("@");
  ansSheet.getRange(ansStartRow, 1, ansRows.length, ANS_HEADER.length).setValues(ansRows);

  // Низкий балл по пункту с привязанным курсом -> автоматически в план обучения.
  // Низким считаем: блок 1 значение "0", блоки 2-3 значение 1 или 2.
  // Критерии с relatedCourse ищем по каталогу для этого типа техники.
  let coursesAdded = 0;
  try {
    const critResp = JSON.parse(getEvaluationCriteria(equipmentType).getContent());
    const courseByKey = {};
    ["block1","block2","block3"].forEach(bk => {
      (critResp[bk] || []).forEach(it => {
        if (it.relatedCourse) courseByKey[it.itemNum] = it.relatedCourse;
      });
    });
    answers.forEach(a => {
      const isLow = String(a.block) === "1" ? String(a.value) === "0"
                  : (Number(a.value) === 1 || Number(a.value) === 2);
      if (!isLow) return;
      const course = courseByKey[String(a.itemNum)];
      if (!course) return;
      addTrainingPlan(empId || "", empName, course, "", site || "");
      coursesAdded++;
    });
  } catch (e) {
    // Не блокируем сохранение оценки, если авто-рекомендация не удалась
  }

  // Личное дело — та же запись видна в истории сотрудника рядом со взысканиями/поощрениями.
  const scoresText = "Обязанности: " + block1Score + ", Техническое состояние: " + block2Score + ", Эксплуатация: " + block3Score;
  savePersonnelEvent(
    empId || "", empName, "Оценка",
    dateStr,
    scoresText + (summary ? (". " + summary) : ""),
    instructorName || "",
    evalId
  );

  return json({ ok: true, evalId, block1Score, block2Score, block3Score, coursesAdded });
}

// ── Редактирование уже сохранённой оценки ────────────────────
// В отличие от saveEvaluation: находит существующую строку по evalId и
// перезаписывает её (не создаёт новую), удаляет и пишет заново детализацию
// по пунктам, обновляет СВЯЗАННУЮ запись личного дела (не плодит вторую).
// Авто-рекомендацию курса при низком балле НЕ повторяет — иначе повторное
// редактирование одной и той же слабой оценки плодило бы дубликаты в
// «ПланОбучения»; рекомендация ставится только при первом создании.
function updateEvaluation(p) {
  const { evalId, empId, empName, position, date, equipmentType, equipmentModel, site,
          instructorName, supervisorName, workType, summary, answers } = p || {};

  if (!evalId) return json({ ok: false, error: "Нужен evalId" });
  if (!empName || !equipmentType || !workType || !Array.isArray(answers) || !answers.length) {
    return json({ ok: false, error: "Нужны как минимум empName, equipmentType, workType (вид выполняемых работ) и ответы по пунктам" });
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const evalSheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!evalSheet) return json({ ok: false, error: "Лист «ОценкиОператоров» не найден" });

  const rows = evalSheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf("evalId");
  if (idCol < 0) return json({ ok: false, error: "Не найдена колонка evalId" });

  let rowNum = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]).trim() === String(evalId).trim()) { rowNum = i + 1; break; }
  }
  if (rowNum < 0) return json({ ok: false, error: "Оценка с таким evalId не найдена — возможно, была удалена. Обновите список." });

  const b1values = answers.filter(a => String(a.block) === "1").map(a => a.value);
  const b2values = answers.filter(a => String(a.block) === "2").map(a => a.value);
  const b3values = answers.filter(a => String(a.block) === "3").map(a => a.value);
  const block1Score = computeEvalBlock1Score(b1values);
  const block2Score = computeEvalAvgScore(b2values);
  const block3Score = computeEvalAvgScore(b3values);

  const dateStr = date || Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy");
  const col = h => headers.indexOf(h) + 1;
  // evalId и createdAt намеренно не трогаем — сохраняем исходную идентичность записи
  evalSheet.getRange(rowNum, col("empId")).setValue(empId || "");
  evalSheet.getRange(rowNum, col("empName")).setValue(empName);
  evalSheet.getRange(rowNum, col("position")).setValue(position || "");
  evalSheet.getRange(rowNum, col("date")).setValue(dateStr);
  evalSheet.getRange(rowNum, col("equipmentType")).setValue(equipmentType);
  evalSheet.getRange(rowNum, col("equipmentModel")).setValue(equipmentModel || "");
  evalSheet.getRange(rowNum, col("site")).setValue(site || "");
  evalSheet.getRange(rowNum, col("instructorName")).setValue(instructorName || "");
  evalSheet.getRange(rowNum, col("supervisorName")).setValue(supervisorName || "");
  evalSheet.getRange(rowNum, col("workType")).setValue(workType || "");
  evalSheet.getRange(rowNum, col("block1Score")).setValue(block1Score);
  evalSheet.getRange(rowNum, col("block2Score")).setValue(block2Score);
  evalSheet.getRange(rowNum, col("block3Score")).setValue(block3Score);
  evalSheet.getRange(rowNum, col("summary")).setValue(summary || "");

  // Детализация: старые строки по этому evalId удаляем, пишем актуальные заново
  let ansSheet = ss.getSheetByName(SHEET_EVAL_ANSWERS);
  const ANS_HEADER = ["evalId","block","itemNum","itemText","value","comment"];
  if (!ansSheet) {
    ansSheet = ss.insertSheet(SHEET_EVAL_ANSWERS);
    ansSheet.appendRow(ANS_HEADER);
    ansSheet.getRange(1,1,1,ANS_HEADER.length)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    ansSheet.setFrozenRows(1);
    ansSheet.setColumnWidth(4, 400);
  } else {
    const ansRows = ansSheet.getDataRange().getValues();
    for (let i = ansRows.length - 1; i >= 1; i--) {
      if (String(ansRows[i][0]).trim() === String(evalId).trim()) ansSheet.deleteRow(i + 1);
    }
  }
  const newAnsRows = answers.map(a => [
    evalId, String(a.block || ""), String(a.itemNum || ""), String(a.itemText || ""),
    a.value === undefined || a.value === null ? "" : String(a.value), a.comment || "",
  ]);
  ensureCapacity(ansSheet, newAnsRows.length);
  const ansItemNumCol = ANS_HEADER.indexOf("itemNum") + 1;
  const ansStartRow = ansSheet.getLastRow() + 1;
  ansSheet.getRange(ansStartRow, ansItemNumCol, newAnsRows.length, 1).setNumberFormat("@");
  ansSheet.getRange(ansStartRow, 1, newAnsRows.length, ANS_HEADER.length).setValues(newAnsRows);

  // Личное дело: обновляем связанную запись (по evalId), а не создаём новую.
  // Для оценок, сохранённых до появления поля evalId, связи нет — тогда
  // создаём запись, как при первом сохранении.
  const scoresText = "Обязанности: " + block1Score + ", Техническое состояние: " + block2Score + ", Эксплуатация: " + block3Score;
  const peSheet = ss.getSheetByName(SHEET_PERSONNEL_EVENTS);
  let peUpdated = false;
  if (peSheet) {
    const peRows = peSheet.getDataRange().getValues();
    const peHeaders = peRows[0];
    const peEvalIdCol = peHeaders.indexOf("evalId");
    if (peEvalIdCol >= 0) {
      for (let i = 1; i < peRows.length; i++) {
        if (String(peRows[i][peEvalIdCol]).trim() === String(evalId).trim()) {
          const dateColNum = peHeaders.indexOf("date") + 1;
          const descColNum = peHeaders.indexOf("description") + 1;
          const issColNum  = peHeaders.indexOf("issuedBy") + 1;
          peSheet.getRange(i + 1, dateColNum).setValue(dateStr);
          peSheet.getRange(i + 1, descColNum).setValue(scoresText + (summary ? (". " + summary) : ""));
          if (issColNum > 0) peSheet.getRange(i + 1, issColNum).setValue(instructorName || "");
          peUpdated = true;
          break;
        }
      }
    }
  }
  if (!peUpdated) {
    savePersonnelEvent(
      empId || "", empName, "Оценка", dateStr,
      scoresText + (summary ? (". " + summary) : ""),
      instructorName || "", evalId
    );
  }

  return json({ ok: true, evalId, block1Score, block2Score, block3Score });
}

function deleteEvaluation(p) {
  const { row, evalId } = p || {};
  if (!row) return json({ ok: false, error: "Нужен номер строки (row)" });

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return json({ ok: false, error: "Лист «ОценкиОператоров» не найден" });

  const rowNum = Number(row);
  if (!rowNum || rowNum < 2 || rowNum > sheet.getLastRow()) {
    return json({ ok: false, error: "Строка не найдена — список успел измениться, обновите страницу." });
  }

  if (evalId) {
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const idCol = headers.indexOf("evalId");
    const actualId = idCol >= 0 ? String(sheet.getRange(rowNum, idCol + 1).getValue()).trim() : "";
    if (actualId !== String(evalId).trim()) {
      return json({ ok: false, error: "Строка сдвинулась (список изменился параллельно). Обновите страницу и повторите." });
    }
  }

  sheet.deleteRow(rowNum);

  // Чистим детализацию по этой оценке, чтобы не копить осиротевшие строки
  if (evalId) {
    const ansSheet = ss.getSheetByName(SHEET_EVAL_ANSWERS);
    if (ansSheet) {
      const ansRows = ansSheet.getDataRange().getValues();
      for (let i = ansRows.length - 1; i >= 1; i--) {
        if (String(ansRows[i][0]).trim() === String(evalId).trim()) ansSheet.deleteRow(i + 1);
      }
    }
  }

  return json({ ok: true, deleted: rowNum });
}

// ── История оценок сотрудника (для личного дела) ─────────────
function getOperatorEvaluations(empId) {
  if (!empId) return json([]);
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const idCol = headers.indexOf("empId");

  const data = [];
  rows.slice(1).forEach((r, i) => {
    if (!r[0]) return;
    if (idCol >= 0 && String(r[idCol]).trim() !== String(empId).trim()) return;
    const o = { __sortDate: r[headers.indexOf("date")] };
    headers.forEach((h,j) => {
      const val = r[j];
      if ((h === "date" || h === "createdAt") && val instanceof Date) {
        o[h] = Utilities.formatDate(val, "Asia/Almaty", h === "createdAt" ? "dd.MM.yyyy HH:mm" : "dd.MM.yyyy");
      } else {
        o[h] = val;
      }
    });
    o.__row = i + 2;
    data.push(o);
  });
  // Сортируем по исходному значению даты (Date или сравнимая строка), а не
  // по уже отформатированной "dd.MM.yyyy" — иначе лексикографический порядок
  // ломается (день впереди года).
  data.sort((a,b) => {
    const av = a.__sortDate, bv = b.__sortDate;
    if (av instanceof Date && bv instanceof Date) return bv - av;
    return String(bv).localeCompare(String(av));
  });
  data.forEach(o => delete o.__sortDate);
  return json(data);
}

// Детализация по пунктам одной конкретной оценки — подгружается отдельно,
// чтобы список истории (getOperatorEvaluations) оставался лёгким.
function getEvaluationDetail(evalId) {
  if (!evalId) return json([]);
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EVAL_ANSWERS);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const idCol = headers.indexOf("evalId");

  const data = rows.slice(1)
    .filter(r => idCol >= 0 && String(r[idCol]).trim() === String(evalId).trim())
    .map(r => {
      const o = {}; headers.forEach((h,i) => { o[h] = r[i]; }); return o;
    });
  return json(data);
}

// ── Сводка для дашборда: недавние оценки с низкими баллами ───
// Низким считаем блок 1 или 2 или 3, если ≤2 — независимо, по любому блоку.
// Не выбираем среднее, не сглаживаем: любой слабый блок = запись в списке.
function getEvaluationAlerts() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return json({ ok: true, total: 0, lowCount: 0, alerts: [] });
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json({ ok: true, total: 0, lowCount: 0, alerts: [] });
  const headers = rows[0];
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);

  const isLow = v => v !== "-" && v !== "" && !isNaN(Number(v)) && Number(v) <= 2;

  const alerts = [];
  let total = 0;
  rows.slice(1).forEach(r => {
    if (!r[idx["evalId"]]) return;
    total++;
    const b1 = r[idx["block1Score"]], b2 = r[idx["block2Score"]], b3 = r[idx["block3Score"]];
    const lowBlocks = [];
    if (isLow(b1)) lowBlocks.push("1");
    if (isLow(b2)) lowBlocks.push("2");
    if (isLow(b3)) lowBlocks.push("3");
    if (!lowBlocks.length) return;
    const rawDate = r[idx["date"]];
    alerts.push({
      evalId: r[idx["evalId"]], empId: r[idx["empId"]], empName: r[idx["empName"]],
      date: rawDate instanceof Date ? Utilities.formatDate(rawDate, "Asia/Almaty", "dd.MM.yyyy") : rawDate,
      __sortDate: rawDate,
      equipmentType: r[idx["equipmentType"]],
      block1Score: b1, block2Score: b2, block3Score: b3,
      lowBlocks: lowBlocks, summary: r[idx["summary"]] || "",
    });
  });

  alerts.sort((a,b) => {
    const av = a.__sortDate, bv = b.__sortDate;
    if (av instanceof Date && bv instanceof Date) return bv - av;
    return String(bv).localeCompare(String(av));
  });
  alerts.forEach(a => delete a.__sortDate);
  return json({ ok: true, total, lowCount: alerts.length, alerts: alerts.slice(0, 100) });
}

// ── Рейтинг по оценкам работы — независимый от рейтинга по осмотрам ──
// Усредняет три блока по каждому оператору за период/участок (те же
// фильтры, что уже есть на вкладке «Рейтинг операторов»). Сортировка по
// среднему трёх блоков — только для порядка вывода (кто выше в списке),
// пользователю всегда показываются три отдельных балла, не единая цифра —
// сводить их в один рейтинг с осмотрами техники осознанно не стали.
function getEvaluationStats(p) {
  const siteFilter = (p && p.site && p.site !== "all") ? String(p.site).trim() : "";
  const days = Number((p && p.days) || 0);
  // Точное окно [from, to) — для отчётов за конкретный период (тот же
  // принцип, что уже используется в getChecklistStats). from/to в приоритете
  // над days, если переданы оба.
  const fromParam = p && p.from ? new Date(p.from) : null;
  const toParam   = p && p.to   ? new Date(p.to)   : null;

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EVALUATIONS);
  if (!sheet) return json({ ok: true, operators: [] });
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json({ ok: true, operators: [] });
  const headers = rows[0];
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);

  let since = null, until = null;
  if (fromParam && !isNaN(fromParam.getTime())) {
    since = fromParam;
    until = (toParam && !isNaN(toParam.getTime())) ? toParam : new Date();
  } else if (days > 0) {
    since = new Date();
    since.setDate(since.getDate() - days);
    since.setHours(0,0,0,0);
  }

  const byOp = {}; // empName -> { empId, count, lastDate, evaluations[] }
  rows.slice(1).forEach(r => {
    if (!r[idx["evalId"]]) return;
    const site = String(r[idx["site"]] || "").trim();
    if (siteFilter && site !== siteFilter) return;

    const rawDate = r[idx["date"]];
    const d = rawDate instanceof Date ? rawDate : new Date(rawDate);
    if (since && (isNaN(d) || d < since)) return;
    if (until && !isNaN(d) && d >= until) return;

    const empName = String(r[idx["empName"]] || "").trim();
    if (!empName) return;
    const empId = r[idx["empId"]] || "";

    if (!byOp[empName]) byOp[empName] = { empId: empId, count: 0, lastDate: null, evaluations: [] };
    const op = byOp[empName];
    op.count++;
    if (empId && !op.empId) op.empId = empId;
    const b1 = r[idx["block1Score"]], b2 = r[idx["block2Score"]], b3 = r[idx["block3Score"]];
    if (!isNaN(d) && (!op.lastDate || d > op.lastDate)) op.lastDate = d;

    // Резюме и детали по каждой конкретной оценке (не только агрегат) — для
    // отображения под общей карточкой оператора, если у него их несколько.
    op.evaluations.push({
      date: !isNaN(d) ? Utilities.formatDate(d, "Asia/Almaty", "dd.MM.yyyy") : String(rawDate || ""),
      __sortDate: !isNaN(d) ? d.getTime() : 0,
      block1Score: b1, block2Score: b2, block3Score: b3,
      equipmentType: r[idx["equipmentType"]] || "",
      workType: r[idx["workType"]] || "",
      summary: r[idx["summary"]] || "",
    });
  });

  // Средний балл — с ранговым взвешиванием по давности: оценки сортируются
  // от старой к новой, вес растёт линейно (1, 2, 3...N), т.е. самая свежая
  // оценка весит в N раз больше самой первой. Один и тот же вес (по дате
  // самой оценки) применяется сразу ко всем трём блокам этой оценки — так
  // три параметра одной проверки не расходятся по весам между собой.
  // Простое среднее было бы несправедливо к тем, кто улучшился со временем:
  // одна давняя слабая оценка топила бы рейтинг наравне со свежей хорошей.
  function weightedAvg(evalsChronological, field) {
    let sumW = 0, sumWV = 0;
    evalsChronological.forEach((e, i) => {
      const v = e[field];
      if (v === "-" || v === "" || v === undefined || v === null || isNaN(Number(v))) return;
      const weight = i + 1; // 1..N, старая -> новая
      sumW += weight;
      sumWV += weight * Number(v);
    });
    return sumW ? Math.round((sumWV / sumW) * 100) / 100 : "-";
  }

  const operators = Object.keys(byOp).map(name => {
    const o = byOp[name];
    const chronological = o.evaluations.slice().sort((a,b) => a.__sortDate - b.__sortDate);
    const b1 = weightedAvg(chronological, "block1Score");
    const b2 = weightedAvg(chronological, "block2Score");
    const b3 = weightedAvg(chronological, "block3Score");
    const nums = [b1,b2,b3].filter(v => v !== "-");
    // Средний балл по трём блокам — теперь отдаём его и в ответе API, не
    // только используем для сортировки (раньше намеренно скрывали).
    const overallAvg = nums.length ? Math.round((nums.reduce((a,b)=>a+b,0)/nums.length)*100)/100 : "-";
    const evaluations = o.evaluations
      .sort((a,b) => b.__sortDate - a.__sortDate)
      .map(e => { const { __sortDate, ...rest } = e; return rest; });
    return {
      empName: name, empId: o.empId, count: o.count,
      block1Avg: b1, block2Avg: b2, block3Avg: b3, overallAvg: overallAvg,
      lastDate: o.lastDate ? Utilities.formatDate(o.lastDate, "Asia/Almaty", "dd.MM.yyyy") : "",
      evaluations: evaluations,
      __sort: nums.length ? nums.reduce((a,b)=>a+b,0)/nums.length : 0,
    };
  }).sort((a,b) => b.__sort - a.__sort);
  operators.forEach(o => delete o.__sort);

  return json({ ok: true, operators });
}

// ── Резюме оценки через Claude API ────────────────────────────
// Тот же ключ и модель, что и recognizeRoster (Script Properties,
// ANTHROPIC_API_KEY). Резюме объективное — без обязательного баланса
// плюсов и минусов: если оценка низкая, текст должен прямо это отражать,
// а не сглаживать дежурными формулировками.
function generateEvaluationSummary(p) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
  if (!apiKey) {
    return json({ ok: false, error: "API-ключ не настроен. Project Settings → Script Properties → ANTHROPIC_API_KEY." });
  }

  const { empName, position, equipmentType, block1Score, block2Score, block3Score, weakItems, strongItems } = p || {};

  const weakText = Array.isArray(weakItems) && weakItems.length
    ? weakItems.map(i => "- " + i).join("\n") : "нет отмеченных слабых пунктов";
  const strongText = Array.isArray(strongItems) && strongItems.length
    ? strongItems.map(i => "- " + i).join("\n") : "нет отдельно отмеченных сильных пунктов";

  const prompt =
    "Составь краткое объективное резюме (2-4 предложения, на русском) по результатам оценки " +
    "профессиональной деятельности оператора на месте работы.\n\n" +
    "Оператор: " + (empName || "не указан") + ", должность: " + (position || "не указана") + ".\n" +
    "Тип техники: " + (equipmentType || "не указан") + ".\n" +
    "Балл по блоку «Обязанности» (1-5, или «-» если не оценивалось): " + block1Score + "\n" +
    "Балл по блоку «Техническое состояние» (1-5): " + block2Score + "\n" +
    "Балл по блоку «Эксплуатация» (1-5): " + block3Score + "\n\n" +
    "Пункты с низкой оценкой:\n" + weakText + "\n\n" +
    "Пункты с высокой оценкой:\n" + strongText + "\n\n" +
    "ВАЖНО: резюме должно быть объективным и точным отражением баллов. Если баллы низкие — " +
    "резюме должно прямо и honestly это отражать, без искусственного смягчения, без обязательного " +
    "упоминания \"положительных сторон\" ради баланса, если их по факту нет. Не используй общие " +
    "фразы-клише. Пиши по существу, для внесения в личное дело сотрудника.";

  const payload = {
    model: "claude-haiku-4-5-20251001",
    max_tokens: 400,
    messages: [{ role: "user", content: prompt }],
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", options);
    const code = response.getResponseCode();
    const body = JSON.parse(response.getContentText());
    if (code !== 200) {
      const msg = (body.error && body.error.message) || `HTTP ${code}`;
      return json({ ok: false, error: msg });
    }
    const text = (body.content && body.content[0] && body.content[0].text) || "";
    return json({ ok: true, summary: text.trim() });
  } catch (e) {
    return json({ ok: false, error: e.message });
  }
}

// ══════════════════════════════════════════════════════════
// ХРОНОМЕТРАЖ ПОГРУЗКИ АВТОТРАНСПОРТА — для машинистов экскаватора.
// Отдельная сущность от «Оценка работы»: не баллы 1-5, а прямые замеры
// времени (в секундах) — общее время погрузки одного самосвала, число
// ковшей, и цикл одного ковша (время между выгрузкой 1-го и 2-го ковша).
// Замеряется внешне (по видео/скриншотам с таймкодами), в систему заносятся
// уже готовые числа. Архитектура и стиль функций — зеркало «Оценка работы»
// (saveEvaluation/updateEvaluation/getOperatorEvaluations), чтобы не
// изобретать второй паттерн для того же по сути класса задач.
// ══════════════════════════════════════════════════════════
const SHEET_TIMING = "Хронометраж";

function saveTimingRecord(p) {
  const { empId, empName, position, date, equipmentType, equipmentModel, truckModel,
          totalLoadTimeSec, bucketCount, bucketFillPercent, cycleTimeSec, properLoading,
          site, instructorName, summary } = p || {};

  if (!empName || !totalLoadTimeSec || !bucketCount || !cycleTimeSec) {
    return json({ ok: false, error: "Нужны как минимум empName, totalLoadTimeSec, bucketCount и cycleTimeSec" });
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_TIMING);
  const HEADER = ["recordId","empId","empName","position","date","equipmentType","equipmentModel",
                  "truckModel","totalLoadTimeSec","bucketCount","bucketFillPercent","cycleTimeSec",
                  "properLoading","site","instructorName","summary","createdAt"];
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TIMING);
    sheet.appendRow(HEADER);
    sheet.getRange(1,1,1,HEADER.length)
      .setBackground("#0D1B3E").setFontColor("#F4A52A").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(16, 400);
  }

  const recordId = "timing" + Date.now();
  const dateStr = date || Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy");
  const createdAt = Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy HH:mm");

  ensureCapacity(sheet, 1);
  sheet.appendRow([
    recordId, empId || "", empName, position || "", dateStr, equipmentType || "", equipmentModel || "",
    truckModel || "", totalLoadTimeSec, bucketCount, bucketFillPercent || "", cycleTimeSec,
    properLoading || "", site || "", instructorName || "", summary || "", createdAt,
  ]);

  return json({ ok: true, recordId });
}

function updateTimingRecord(p) {
  const { recordId, empId, empName, position, date, equipmentType, equipmentModel, truckModel,
          totalLoadTimeSec, bucketCount, bucketFillPercent, cycleTimeSec, properLoading,
          site, instructorName, summary } = p || {};

  if (!recordId) return json({ ok: false, error: "Нужен recordId" });
  if (!empName || !totalLoadTimeSec || !bucketCount || !cycleTimeSec) {
    return json({ ok: false, error: "Нужны как минимум empName, totalLoadTimeSec, bucketCount и cycleTimeSec" });
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TIMING);
  if (!sheet) return json({ ok: false, error: "Лист «Хронометраж» не найден" });

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf("recordId");
  if (idCol < 0) return json({ ok: false, error: "Не найдена колонка recordId" });

  let rowNum = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]).trim() === String(recordId).trim()) { rowNum = i + 1; break; }
  }
  if (rowNum < 0) return json({ ok: false, error: "Запись с таким recordId не найдена — возможно, была удалена. Обновите список." });

  const dateStr = date || Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy");
  const col = h => headers.indexOf(h) + 1;
  sheet.getRange(rowNum, col("empId")).setValue(empId || "");
  sheet.getRange(rowNum, col("empName")).setValue(empName);
  sheet.getRange(rowNum, col("position")).setValue(position || "");
  sheet.getRange(rowNum, col("date")).setValue(dateStr);
  sheet.getRange(rowNum, col("equipmentType")).setValue(equipmentType || "");
  sheet.getRange(rowNum, col("equipmentModel")).setValue(equipmentModel || "");
  sheet.getRange(rowNum, col("truckModel")).setValue(truckModel || "");
  sheet.getRange(rowNum, col("totalLoadTimeSec")).setValue(totalLoadTimeSec);
  sheet.getRange(rowNum, col("bucketCount")).setValue(bucketCount);
  sheet.getRange(rowNum, col("bucketFillPercent")).setValue(bucketFillPercent || "");
  sheet.getRange(rowNum, col("cycleTimeSec")).setValue(cycleTimeSec);
  sheet.getRange(rowNum, col("properLoading")).setValue(properLoading || "");
  sheet.getRange(rowNum, col("site")).setValue(site || "");
  sheet.getRange(rowNum, col("instructorName")).setValue(instructorName || "");
  sheet.getRange(rowNum, col("summary")).setValue(summary || "");

  return json({ ok: true, recordId });
}

function deleteTimingRecord(p) {
  const { row, recordId } = p || {};
  if (!row) return json({ ok: false, error: "Нужен номер строки (row)" });

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TIMING);
  if (!sheet) return json({ ok: false, error: "Лист «Хронометраж» не найден" });

  const rowNum = Number(row);
  if (!rowNum || rowNum < 2 || rowNum > sheet.getLastRow()) {
    return json({ ok: false, error: "Строка не найдена — список успел измениться, обновите страницу." });
  }

  if (recordId) {
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const idCol = headers.indexOf("recordId");
    const actualId = idCol >= 0 ? String(sheet.getRange(rowNum, idCol + 1).getValue()).trim() : "";
    if (actualId !== String(recordId).trim()) {
      return json({ ok: false, error: "Строка сдвинулась (список изменился параллельно). Обновите страницу и повторите." });
    }
  }

  sheet.deleteRow(rowNum);
  return json({ ok: true, deleted: rowNum });
}

function getOperatorTimingRecords(empId) {
  if (!empId) return json([]);
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TIMING);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const idCol = headers.indexOf("empId");
  const dateCol = headers.indexOf("date");

  const data = [];
  rows.slice(1).forEach((r, i) => {
    if (!r[0]) return;
    if (idCol >= 0 && String(r[idCol]).trim() !== String(empId).trim()) return;
    const o = { __sortDate: r[dateCol] };
    headers.forEach((h,j) => {
      const val = r[j];
      o[h] = (h === "date" && val instanceof Date) ? Utilities.formatDate(val, "Asia/Almaty", "dd.MM.yyyy") : val;
    });
    o.__row = i + 2;
    data.push(o);
  });
  data.sort((a,b) => {
    const av = a.__sortDate, bv = b.__sortDate;
    if (av instanceof Date && bv instanceof Date) return bv - av;
    return String(bv).localeCompare(String(av));
  });
  data.forEach(o => delete o.__sortDate);
  return json(data);
}

// Подсказки для полей «Модель самосвала» — тоже свободный текст, из уже
// внесённых записей хронометража.
function getTimingTruckModels() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TIMING);
  if (!sheet) return json([]);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  const headers = rows[0];
  const col = headers.indexOf("truckModel");
  if (col < 0) return json([]);
  const values = new Set();
  rows.slice(1).forEach(r => {
    const v = String(r[col] || "").trim();
    if (v) values.add(v);
  });
  return json(Array.from(values).sort());
}


// Запустить ОДИН раз вручную из редактора Apps Script (выбрать функцию
// seedEvaluationCriteria → Run), чтобы наполнить/обновить каталог. Повторный
// запуск безопасен — importEvaluationCriteria полностью заменяет набор
// пунктов по каждому типу техники, а не дублирует. Пункты блоков 2-3 взяты
// из структуры пяти присланных файлов (коды LC.x.x.x сохранены в тексте
// пункта для прослеживаемости к исходному регламенту), баллы из заполненных
// образцов проигнорированы — использована только структура критериев.
function seedEvaluationCriteria() {
  importEvaluationCriteria({
    equipmentType: EVAL_UNIVERSAL_TYPE,
    items: [
      { block: "1", itemNum: "1.1", itemText: "Обязанности оператора (наряд-допуск, инструктаж, предсменные проверки)", scale: "binary" },
      { block: "1", itemNum: "1.2", itemText: "Поддержание чистоты и порядка на рабочем месте", scale: "binary" },
      { block: "1", itemNum: "1.3", itemText: "Средства индивидуальной защиты (СИЗ)", scale: "binary" },
      { block: "1", itemNum: "1.4", itemText: "Использование радиосвязи, включая аварийную", scale: "binary" },
      { block: "1", itemNum: "1.5", itemText: "Аварийная остановка станка", scale: "binary" },
    ],
  });
  importEvaluationCriteria({
    equipmentType: "Буровая установка",
    items: [
      { block: "2", itemNum: "2.1", itemText: "Диагностика неисправностей бурового станка", scale: "scale5" },
      { block: "2", itemNum: "2.2", itemText: "Установка станка к бурению скважины", scale: "scale5" },
      { block: "2", itemNum: "2.3", itemText: "Работа с бортовой системой диспетчеризации/мониторинга", scale: "scale5" },
      { block: "2", itemNum: "2.4", itemText: "Техническое обслуживание и устранение неисправностей оборудования", scale: "scale5" },
      { block: "2", itemNum: "2.5", itemText: "Поддержание работоспособности оборудования при бурении", scale: "scale5" },
      { block: "3", itemNum: "3.1", itemText: "Планирование и подготовка работ по бурению", scale: "scale5" },
      { block: "3", itemNum: "3.2", itemText: "Поверхностное бурение взрывных скважин", scale: "scale5" },
      { block: "3", itemNum: "3.3", itemText: "Вращательное и ударно-вращательное бурение", scale: "scale5" },
      { block: "3", itemNum: "3.4", itemText: "Бурение наклонных скважин", scale: "scale5" },
      { block: "3", itemNum: "3.5", itemText: "Реагирование на возникающие проблемы при бурении", scale: "scale5" },
      { block: "3", itemNum: "3.6", itemText: "Завершение работ и уборка рабочего места", scale: "scale5" },
    ],
  });
  importEvaluationCriteria({
    equipmentType: "Колесный погрузчик",
    items: [
      { block: "2", itemNum: "2.1", itemText: "LC-2-1-1 Общее устройтва погрузчика", scale: "scale5" },
      { block: "2", itemNum: "2.2", itemText: "LC-2-1-2 Технические характеристики погрузчика САТ993К /САТ994К", scale: "scale5" },
      { block: "2", itemNum: "2.3", itemText: "LC-2-1-3 Устройство двигателя 1", scale: "scale5" },
      { block: "2", itemNum: "2.4", itemText: "LC-2-1-5 Гидравлическая система", scale: "scale5" },
      { block: "2", itemNum: "2.5", itemText: "LC-2-1-6 Осмотр обходом", scale: "scale5" },
      { block: "2", itemNum: "2.6", itemText: "LC-2-1-7 Электрическая система", scale: "scale5" },
      { block: "2", itemNum: "2.7", itemText: "LC-2-1-8 Противопожарная система", scale: "scale5" },
      { block: "2", itemNum: "2.8", itemText: "LC-2-1-10  Доступ к верхней площадке", scale: "scale5" },
      { block: "2", itemNum: "2.9", itemText: "LC-2-1-11 Рабочее оборудование / LC-2-1-13 Освещение", scale: "scale5" },
      { block: "2", itemNum: "2.10", itemText: "LC-2-1-12 Колеса", scale: "scale5" },
      { block: "2", itemNum: "2.11", itemText: "LC-2-1-14 Задняя часть/ LC-2-1-15 Правая сторона / LC-2-1-16 Левая сторона", scale: "scale5" },
      { block: "2", itemNum: "2.12", itemText: "LC-2-1-17 Сервисная площадка (палуба)", scale: "scale5" },
      { block: "2", itemNum: "2.13", itemText: "LC-2-2-2 Электронный прибор САТ 994К993К/ LC-2-2-5 Клавиатура панели управления", scale: "scale5" },
      { block: "2", itemNum: "2.14", itemText: "LC-2-2-3 Джойстик рулевого управление и трансмиссией/ LC-2-2-4 Управление рабочим оборудованием", scale: "scale5" },
      { block: "2", itemNum: "2.15", itemText: "LC-2-2-6 Внешние органы управления", scale: "scale5" },
      { block: "3", itemNum: "3.1", itemText: "LC-3-1-1 Планирование и подготовка к работам", scale: "scale5" },
      { block: "3", itemNum: "3.2", itemText: "LC-3-1-2 Подъем и спуск", scale: "scale5" },
      { block: "3", itemNum: "3.3", itemText: "LC-3-1-3 Эффективность и безопасность работы", scale: "scale5" },
      { block: "3", itemNum: "3.4", itemText: "LC-3-1-5 Важность поддержания эффективной работы", scale: "scale5" },
      { block: "3", itemNum: "3.5", itemText: "LC-3-2-1 Действие перед началом работы", scale: "scale5" },
      { block: "3", itemNum: "3.6", itemText: "LC-3-2-2 Действие при запуске двигателя/ LC-3-2-3 Проверка после запуска двигателя", scale: "scale5" },
      { block: "3", itemNum: "3.7", itemText: "LC-3-2-4 Проверка тормозной системы", scale: "scale5" },
      { block: "3", itemNum: "3.8", itemText: "LC-3-2-5 Начало движение/ LC-3-2-6 Перегон погрузчика", scale: "scale5" },
      { block: "3", itemNum: "3.9", itemText: "LC-3-2-8 Погрузка самосвала/ LC-3-2-9 Работа на уклонах", scale: "scale5" },
      { block: "3", itemNum: "3.10", itemText: "LC-3-2-11 Общая практика эксплуатации погрузчика", scale: "scale5" },
      { block: "3", itemNum: "3.11", itemText: "LC-3-3-1 Предупреждение опасности/ LC-3-3-2 Электрические проводники", scale: "scale5" },
      { block: "3", itemNum: "3.12", itemText: "LC-3-3-3 Управление в неблагоприятных условиях", scale: "scale5" },
      { block: "3", itemNum: "3.13", itemText: "LC-3-3-4 Минимизация воздействия на окружающую среду", scale: "scale5" },
      { block: "3", itemNum: "3.14", itemText: "LC-3-4-1 Дорожно-транспортные происшествия/ LC-3-4-2 Процедура аварийного реагирования по рации", scale: "scale5" },
      { block: "3", itemNum: "3.15", itemText: "LC-3-4-3 Пожар на борту погрузчика/ LC-3-4-5 Процедура эвакуации при опрокидывании", scale: "scale5" },
    ],
  });
  importEvaluationCriteria({
    equipmentType: "Карьерный самосвал",
    items: [
      { block: "2", itemNum: "2.1", itemText: "LC 2.1.1 Технические характеристики самосвала", scale: "scale5" },
      { block: "2", itemNum: "2.2", itemText: "LC 2.1.2 Органы управления и контрольные приборы", scale: "scale5" },
      { block: "2", itemNum: "2.3", itemText: "LC 2.1.3 Осмотр обходом", scale: "scale5" },
      { block: "2", itemNum: "2.4", itemText: "LC 2.1.4 Передняя сторона/ LC 2.1.7 Правая сторона/ LC 2.1.11 Задняя часть/ LC 2.1.12 Левая сторона", scale: "scale5" },
      { block: "2", itemNum: "2.5", itemText: "LC 2.1.5 Освещение/ LC 2.1.6 Кузов", scale: "scale5" },
      { block: "2", itemNum: "2.6", itemText: "LC 2.1.10 Подвески", scale: "scale5" },
      { block: "2", itemNum: "2.7", itemText: "LC 2.1.8 Колеса", scale: "scale5" },
      { block: "2", itemNum: "2.8", itemText: "LC 2.1.9 Система рулевого управления", scale: "scale5" },
      { block: "2", itemNum: "2.9", itemText: "LC 2.1.13 Сервисная площадка (палуба)", scale: "scale5" },
      { block: "2", itemNum: "2.10", itemText: "LC 2.1.14 Система охлаждения двигателя", scale: "scale5" },
      { block: "2", itemNum: "2.11", itemText: "LC 2.1.16 ДВС", scale: "scale5" },
      { block: "2", itemNum: "2.12", itemText: "LC 2.2.1 Кабина оператора", scale: "scale5" },
      { block: "2", itemNum: "2.13", itemText: "LC 2.2.2 Доступ в кабину", scale: "scale5" },
      { block: "2", itemNum: "2.14", itemText: "LC 2.2.3 Приборы и элементы управления в кабине", scale: "scale5" },
      { block: "2", itemNum: "2.15", itemText: "LC 2.2.4 Панель системы мониторинга", scale: "scale5" },
      { block: "3", itemNum: "3.1", itemText: "LC 3.1.5 Тестирование рабочих тормозов", scale: "scale5" },
      { block: "3", itemNum: "3.2", itemText: "LC 3.1.6 Тестирование стояночного тормоза", scale: "scale5" },
      { block: "3", itemNum: "3.3", itemText: "LC 3.1.7 Тестирование вспомогательной системы тормозов", scale: "scale5" },
      { block: "3", itemNum: "3.4", itemText: "LC 3.1.8 Тестирование тормоза замедлителя", scale: "scale5" },
      { block: "3", itemNum: "3.5", itemText: "LC 3.2.1 Начало движения", scale: "scale5" },
      { block: "3", itemNum: "3.6", itemText: "LC 3.2.2 Навыки вождения", scale: "scale5" },
      { block: "3", itemNum: "3.7", itemText: "LC 3.2.3 Процедуры по безопасности во время поворотов", scale: "scale5" },
      { block: "3", itemNum: "3.8", itemText: "LC 3.2.4 Использования замедлителя", scale: "scale5" },
      { block: "3", itemNum: "3.9", itemText: "LC 3.2.5 Остановка движения и парковка", scale: "scale5" },
      { block: "3", itemNum: "3.10", itemText: "LC 3.2.6 Эффективная производительность и ожидание загрузку", scale: "scale5" },
      { block: "3", itemNum: "3.11", itemText: "LC 3.2.7 Общие положения по загрузке", scale: "scale5" },
      { block: "3", itemNum: "3.12", itemText: "LC 3.2.8 Загрузка экскаваторам", scale: "scale5" },
      { block: "3", itemNum: "3.13", itemText: "LC 3.2.9 Разгрузка", scale: "scale5" },
      { block: "3", itemNum: "3.14", itemText: "LC 3.3.0 Категории предупреждения VIMS", scale: "scale5" },
      { block: "3", itemNum: "3.15", itemText: "LC 3.3.1 Индикаторы и система сигнализации", scale: "scale5" },
    ],
  });
  importEvaluationCriteria({
    equipmentType: "Экскаватор с обратной лопатой",
    items: [
      { block: "2", itemNum: "2.1", itemText: "LC.2.1.1 Общее устройство экскаватора", scale: "scale5" },
      { block: "2", itemNum: "2.2", itemText: "LC.2.1.3 Устройство двигателя 1", scale: "scale5" },
      { block: "2", itemNum: "2.3", itemText: "LC.2.1.5 Гидравлическая система", scale: "scale5" },
      { block: "2", itemNum: "2.4", itemText: "LC.2.1.6 Электрическая система", scale: "scale5" },
      { block: "2", itemNum: "2.5", itemText: "LC.2.1.7 Противопожарная система", scale: "scale5" },
      { block: "2", itemNum: "2.6", itemText: "LC.2.1.8 Ходовая часть/ LC.2.1.9 Поворотная платформа/ LC.2.1.10 Рабочее оборудование", scale: "scale5" },
      { block: "2", itemNum: "2.7", itemText: "LC.2.2.1 Контрольные приборы и органы управления", scale: "scale5" },
      { block: "2", itemNum: "2.8", itemText: "LC.2.2.6 Панель управления/ LC.2.2.7 Другие приборы", scale: "scale5" },
      { block: "2", itemNum: "2.9", itemText: "LC.2.2.8 Внешние органы управления", scale: "scale5" },
      { block: "2", itemNum: "2.10", itemText: "LC.2.3.1 Профилактическое обслуживание", scale: "scale5" },
      { block: "2", itemNum: "2.11", itemText: "LC.2.3.2 Наружный осмотр/ LC.2.3.3 Правая сторона/ LC.2.3.4 Левая сторона/ LC.2.3.5 Передняя и задняя сторон", scale: "scale5" },
      { block: "2", itemNum: "2.12", itemText: "LC.2.3.6 Сервисная платформа/ LC.2.3.7 Кабина оператора", scale: "scale5" },
      { block: "2", itemNum: "2.13", itemText: "LC.2.3.8 Плановое техническое обслуживания", scale: "scale5" },
      { block: "2", itemNum: "2.14", itemText: "LC.2.3.10 Процедура запуска/ LC.2.3.11 Эксплуатационные проверки/ LC.2.3.12 Начало движения", scale: "scale5" },
      { block: "2", itemNum: "2.15", itemText: "LC.2.3.16 Процедура парковки и остановки/ LC.2.3.17 Аварийная остановка", scale: "scale5" },
      { block: "3", itemNum: "3.1", itemText: "LC.3.1.1 Планирование и подготовка к работам/ LC.3.1.2 Подъем и спуск", scale: "scale5" },
      { block: "3", itemNum: "3.2", itemText: "LC.3.1.3 Эффективность и безопасность работы/ LC.3.1.4 Ключевые производственные показатели", scale: "scale5" },
      { block: "3", itemNum: "3.3", itemText: "LC.3.1.7 Перемещение/ LC.3.1.8 Повороты и передвижение", scale: "scale5" },
      { block: "3", itemNum: "3.4", itemText: "LC.3.1.12 Подъем по съезду/ LC.3.1.12 Подъем по съезду", scale: "scale5" },
      { block: "3", itemNum: "3.5", itemText: "LC.3.2.1 Рабочий цикл/ LC.3.2.2 Земляные работы/ LC.3.2.3 Навыки черпания", scale: "scale5" },
      { block: "3", itemNum: "3.6", itemText: "LC.3.2.4 Цикл поворота и выгрузка материала/ LC.3.2.5 Режимы черпания", scale: "scale5" },
      { block: "3", itemNum: "3.7", itemText: "LC.3.2.6 Разработка остаточной части уступа", scale: "scale5" },
      { block: "3", itemNum: "3.8", itemText: "LC.3.2.7 Подъем на рабочий уступ/ LC.3.2.8 Спуск с рабочего уступа", scale: "scale5" },
      { block: "3", itemNum: "3.9", itemText: "LC.3.2.9 Техника погрузки/ LC.3.2.10 Равномерное погрузка самосвалов", scale: "scale5" },
      { block: "3", itemNum: "3.10", itemText: "LC.3.2.11 Сквозная схема подачи самосвалов под погрузку/ LC.3.2.12 Верхняя погрузка/ LC.3.2.13 Нижняя погрузка", scale: "scale5" },
      { block: "3", itemNum: "3.11", itemText: "LC.3.2.14 Складирование горный массы/ LC.3.2.16 Методы зачистки", scale: "scale5" },
      { block: "3", itemNum: "3.12", itemText: "LC.3.2.17 Срезка откосов уступа/ LC.3.2.18 Заоткоска уступов", scale: "scale5" },
      { block: "3", itemNum: "3.13", itemText: "LC.3.2.20 Погрузка экскаватора под транспортировку/ LC.3.3.2 Работа в близи ЛЭП", scale: "scale5" },
      { block: "3", itemNum: "3.14", itemText: "LC.3.3.3 Управление в неблагоприятных условиях/ LC.3.4.1 Дорожно.транспортные происшествия", scale: "scale5" },
      { block: "3", itemNum: "3.15", itemText: "LC.3.4.2 Процедура аварийного реагирования по рации/ LC.3.4.3 Пожар на борту экскаватора", scale: "scale5" },
    ],
  });
  importEvaluationCriteria({
    equipmentType: "Экскаватор с прямой лопатой",
    items: [
      { block: "2", itemNum: "2.1", itemText: "LC.2.1.1 ОБЩЕЕ УСТРОЙСТВА/ LC.2.1.2 Технические характеристики ЕХ5600", scale: "scale5" },
      { block: "2", itemNum: "2.2", itemText: "LC.2.1.3 Гидравлическая система", scale: "scale5" },
      { block: "2", itemNum: "2.3", itemText: "LC.2.1.6 Противопожарная система", scale: "scale5" },
      { block: "2", itemNum: "2.4", itemText: "LC.2.1.7 Ходовая часть/ LC.2.1.8 Поворотная часть экскаватора/", scale: "scale5" },
      { block: "2", itemNum: "2.5", itemText: "LC.2.1.9 Навесное оборудование экскаватора", scale: "scale5" },
      { block: "2", itemNum: "2.6", itemText: "LC.2.2.1 ПРИБОРЫ И ОРГАНЫ УПРАВЛЕНИЯ", scale: "scale5" },
      { block: "2", itemNum: "2.7", itemText: "LC.2.2.3 Экран панель бортовой системы управления/ LC.2.2.5 Внешние органы управления", scale: "scale5" },
      { block: "2", itemNum: "2.8", itemText: "LC.2.3.2 Осмотр обходом", scale: "scale5" },
      { block: "2", itemNum: "2.9", itemText: "LC.2.3.3 Левая сторона/ LC.2.3.6 Задняя сторона/ LC.2.3.4 Передняя сторона/ LC.2.3.5 Правая сторона экскаватора", scale: "scale5" },
      { block: "2", itemNum: "2.10", itemText: "LC.2.3.7 Сервисная платформа экскаватора/ LC.2.3.8 Кабина оператора", scale: "scale5" },
      { block: "2", itemNum: "2.11", itemText: "LC.2.3.9 Плановое техническое обслуживание", scale: "scale5" },
      { block: "2", itemNum: "2.12", itemText: "LC.2.3.11 Процедура запуска экскаватора", scale: "scale5" },
      { block: "2", itemNum: "2.13", itemText: "LC.2.3.12 Эксплуатационные проверки", scale: "scale5" },
      { block: "2", itemNum: "2.14", itemText: "LC.2.3.13 Начало движения/ LC.2.3.14 Органы управления", scale: "scale5" },
      { block: "2", itemNum: "2.15", itemText: "LC.2.3.15 Парковка/ LC.2.3.16 Аварийная остановка", scale: "scale5" },
      { block: "3", itemNum: "3.1", itemText: "LC.3.1.1 ПЛАНИРОВАНИЕ И ПОДГОТОВКА К РАБОТЕ/ LC.3.1.2 Подъем и спуск", scale: "scale5" },
      { block: "3", itemNum: "3.2", itemText: "LC.3.1.3 Эффективность и безопасность работы", scale: "scale5" },
      { block: "3", itemNum: "3.3", itemText: "LC.3.1.4 Ключевые производственные показатели/LC.3.1.7 Производственные обязанности", scale: "scale5" },
      { block: "3", itemNum: "3.4", itemText: "LC.3.1.8 Передвижение экскаватора", scale: "scale5" },
      { block: "3", itemNum: "3.5", itemText: "LC.3.1.13 Подъём по съезду/ LC.3.1.14 Спуск по съезду", scale: "scale5" },
      { block: "3", itemNum: "3.6", itemText: "LC.3.2.1 РАБОЧИЕ ЦИКЛЫ/ LC.3.2.2 Работа с забоем/ LC.3.2.3 Рабочее направление", scale: "scale5" },
      { block: "3", itemNum: "3.7", itemText: "LC.3.2.4 Фронтальный доступ/ LC.3.2.5 Параллельный доступ", scale: "scale5" },
      { block: "3", itemNum: "3.8", itemText: "LC.3.2.10 Разбор твердой породы/ LC.3.2.11 Работа с высоким забоем", scale: "scale5" },
      { block: "3", itemNum: "3.9", itemText: "LC.3.2.12 Установка самосвала под погрузку", scale: "scale5" },
      { block: "3", itemNum: "3.10", itemText: "LC.3.2.16 Загрузка негабаритов", scale: "scale5" },
      { block: "3", itemNum: "3.11", itemText: "LC.3.2.18 Содержание подошвы карьера в чистоте", scale: "scale5" },
      { block: "3", itemNum: "3.12", itemText: "LC.3.3.1 ПРЕДУПРЕЖДЕНИЕ ОПАСНОСТИ", scale: "scale5" },
      { block: "3", itemNum: "3.13", itemText: "LC.3.3.3 Управление экскаватором в неблагоприятных условиях 1", scale: "scale5" },
      { block: "3", itemNum: "3.14", itemText: "LC.3.4.1 ДОРОЖНО-ТРАНСПОРТНЫЕ ПРОИШЕСТВИЯ", scale: "scale5" },
      { block: "3", itemNum: "3.15", itemText: "LC.3.4.2 Процедура аварийного реагирования по рации/ LC.3.4.4 Пожар на борту экскаватора/ LC.3.3.2 Электрический контакт", scale: "scale5" },
    ],
  });
  importEvaluationCriteria({
    equipmentType: "Гусеничный бульдозер",
    items: [
      { block: "2", itemNum: "2.1", itemText: "LC.2.1.2 Технические характеристики", scale: "scale5" },
      { block: "2", itemNum: "2.2", itemText: "LC.2.1.6 Шасси и ходовая часть/ LC.2.1.7 Гидравлическая система", scale: "scale5" },
      { block: "2", itemNum: "2.3", itemText: "LC.2.1.8 Рабочая оборудование", scale: "scale5" },
      { block: "2", itemNum: "2.4", itemText: "LC.2.1.9 Отсек быстрый заправки/ LC.2.1.10 Система пожаротушение", scale: "scale5" },
      { block: "2", itemNum: "2.5", itemText: "LC.2.2.1 Инструменты и органы управления/ LC.2.2.2 Система обработки основной информации (VIMS)", scale: "scale5" },
      { block: "2", itemNum: "2.6", itemText: "LC.2.2.5 Функции панели управления", scale: "scale5" },
      { block: "2", itemNum: "2.7", itemText: "LC.2.2.10 Внешние переключатели", scale: "scale5" },
      { block: "2", itemNum: "2.8", itemText: "LC.2.3.1 Планово-предупредительное обслуживание", scale: "scale5" },
      { block: "2", itemNum: "2.9", itemText: "LC.2.3.2 Пред сменный осмотр", scale: "scale5" },
      { block: "2", itemNum: "2.10", itemText: "LC.2.3.3 Передняя и задняя стороны/ LC.2.3.4 Левая сторона/ LC.2.3.4 Левая сторона", scale: "scale5" },
      { block: "2", itemNum: "2.11", itemText: "LC.2.3.6 Кабина оператора/ LC.2.3.10 Эксплуатационные проверки", scale: "scale5" },
      { block: "2", itemNum: "2.12", itemText: "LC.2.3.11 Проверка тормозной системы", scale: "scale5" },
      { block: "2", itemNum: "2.13", itemText: "LC.2.3.12 Начало движения", scale: "scale5" },
      { block: "2", itemNum: "2.14", itemText: "LC.2.3.18 Остановка и парковка", scale: "scale5" },
      { block: "2", itemNum: "2.15", itemText: "LC.2.3.20 Аварийная остановка бульдозера", scale: "scale5" },
      { block: "3", itemNum: "3.1", itemText: "LC.3.1.1 Планирование и подготовка к работам", scale: "scale5" },
      { block: "3", itemNum: "3.2", itemText: "LC.3.1.2 Подъем и спуск/ LC.3.1.3 Эффективная эксплуатация бульдозера", scale: "scale5" },
      { block: "3", itemNum: "3.3", itemText: "LC.3.1.4 Опасности, связанные с ограниченней видимостью", scale: "scale5" },
      { block: "3", itemNum: "3.4", itemText: "LC.3.1.6 Рабочий цикл бульдозера/ LC.3.2.3 Перемещение материала", scale: "scale5" },
      { block: "3", itemNum: "3.5", itemText: "LC.3.2.11 Прокладка дороги на склоне/ LC.3.2.12 Крутые склоны", scale: "scale5" },
      { block: "3", itemNum: "3.6", itemText: "LC.3.2.13 Продольный наклон отвала", scale: "scale5" },
      { block: "3", itemNum: "3.7", itemText: "LC.3.2.14 Обустройство дорог и подготовка бурового блока", scale: "scale5" },
      { block: "3", itemNum: "3.8", itemText: "LC.3.2.15 Складирование", scale: "scale5" },
      { block: "3", itemNum: "3.9", itemText: "LC.3.2.19 Работы на отвале/ LC.3.2.20 Работа в зоне разгрузки/ LC.3.2.21 Предохранительная бровка на отвале", scale: "scale5" },
      { block: "3", itemNum: "3.10", itemText: "LC.3.2.22 Рыхление/ LC.3.2.24 Зачистка рабочей зоны экскаватора", scale: "scale5" },
      { block: "3", itemNum: "3.11", itemText: "LC.3.2.26 Взрывчатые вещества", scale: "scale5" },
      { block: "3", itemNum: "3.12", itemText: "LC.3.2.33 Погрузка бульдозера для транспортировки", scale: "scale5" },
      { block: "3", itemNum: "3.13", itemText: "LC.3.2.34 Защита окружающей среды/ LC.3.2.35 Безопасность и выявление опасностей", scale: "scale5" },
      { block: "3", itemNum: "3.14", itemText: "LC.3.2.36 Электрический контакт/ LC.3.2.37 Управление бульдозером в неблагоприятных условиях 1", scale: "scale5" },
      { block: "3", itemNum: "3.15", itemText: "LC.3.3.1 Дорожно-транспортные происшествия/ LC.3.3.2 Процедура реагирования по рации в случае аварийной ситуации", scale: "scale5" },
    ],
  });
}
