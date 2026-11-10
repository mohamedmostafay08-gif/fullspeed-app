// =========================================================================
// Code.gs — السيرفر والراوتر الأساسي لتطبيق المناديب (Full Speed Payroll Hub)
// -------------------------------------------------------------------------
// ملاحظة أمان مهمة: أي دالة اسمها بيخلص بـ "_" مش ممكن تتنادى من المتصفح.
// الدوال اللي المتصفح يقدر يناديها هي بس:
//   checkSalary / loginWithToken / logout / clearDataCache
// والتطبيق المستقل (GitHub Pages) بيكلم السيرفر عن طريق doPost.
// =========================================================================

var USERS_SHEET = "System_Users";
var SESSION_DAYS = 30;        // مدة "افتكرني على الموبايل ده"
var MAX_SESSIONS = 3;         // أقصى عدد موبايلات مفتوح عليها الحساب في نفس الوقت
var MAX_FAILS = 5;            // عدد المحاولات الغلط قبل القفل المؤقت
var FAIL_LOCK_SECONDS = 600;  // مدة القفل (10 دقايق)
var CACHE_SECONDS = 300;      // بيانات الشيتات بتتخزن مؤقتًا 5 دقايق
var CACHE_CHUNK = 40000;      // حجم الجزء الواحد في الكاش (الحد 100KB)
var LOG_SHEET = "Logs";
var LOG_MAX_ROWS = 20000;     // لما اللوج يعدّي الرقم ده بنمسح أقدم 5000 سطر
var CLIENT_ = "web";          // "web" = رابط Apps Script ، "app" = التطبيق المستقل

var COMPANY_SHEETS = [
  "Talabat", "Amazon now", "Noon Minuts", "Noon Minuts_Pickers", "Noon Minuts_Logistics", "Noon Core",
  "Bosta", "Mylerz", "Breadfast", "InstaShop", "Jumia", "Flextock", "Halan", "Khazenly",
  "Lychee", "Talabya", "PDC Express", "Hassan Allam", "Carrefour", "Zadna", "MaxAB",
  "Tagaddod", "Rabbit", "J&T Express", "JD Express", "Nodust", "Yalla Fel Sekka", "Rumi's"
];

// =========================================================================
// فتح التطبيق
// =========================================================================
function doGet(e) {
  if (e && e.parameter && e.parameter.api === 'ping') {
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  }
  ensureDatabase_();
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Full Speed - Payroll Hub')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover');
}

// API للتطبيق المستقل: body = {"action": "...", "args": [...]}
function doPost(e) {
  var out;
  try {
    CLIENT_ = "app";
    var body = JSON.parse(e.postData.contents);
    var a = Array.isArray(body.args) ? body.args : [];
    switch (body.action) {
      case 'checkSalary': out = checkSalary(a[0], a[1], a[2], a[3] === true); break;
      case 'loginWithToken': out = loginWithToken(a[0], a[1], a[2]); break;
      case 'logout': out = logout(a[0], a[1]); break;
      default: out = fail_("عملية مش معروفة.");
    }
  } catch (err) {
    out = fail_("الطلب مش مفهوم.");
    log_("API", "", "", false, String(err));
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// بيشغّل setupDatabase مرة واحدة بس بدل كل فتحة (أسرع بكتير)
function ensureDatabase_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('DB_READY_V3') === '1') return;
  setupDatabase();
  props.setProperty('DB_READY_V3', '1');
}

// لو ضفت شركة جديدة في COMPANY_SHEETS شغّل الدالة دي يدوي من المحرر
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var users = ss.getSheetByName(USERS_SHEET);
  if (!users) {
    users = ss.insertSheet(USERS_SHEET);
    users.appendRow(["رقم المندوب (ID)", "الشركة", "كلمة المرور (مشفرة)", "تاريخ التسجيل", "جلسات الدخول (مشفرة)"]);
    users.getRange("A1:E1").setFontWeight("bold").setBackground("#00A651").setFontColor("#ffffff");
    users.setFrozenRows(1);
    users.hideSheet();
  } else if (!users.getRange("E1").getValue()) {
    users.getRange("E1").setValue("جلسات الدخول (مشفرة)").setFontWeight("bold").setBackground("#00A651").setFontColor("#ffffff");
  }

  if (!ss.getSheetByName(LOG_SHEET)) {
    var logs = ss.insertSheet(LOG_SHEET);
    logs.appendRow(["الوقت", "العملية", "الشركة", "كود المندوب", "النتيجة", "التفاصيل", "المصدر"]);
    logs.getRange("A1:G1").setFontWeight("bold").setBackground("#1f2937").setFontColor("#ffffff");
    logs.setFrozenRows(1);
  }

  for (var i = 0; i < COMPANY_SHEETS.length; i++) {
    if (!ss.getSheetByName(COMPANY_SHEETS[i])) {
      ss.insertSheet(COMPANY_SHEETS[i]).appendRow(["كود المندوب", "الاسم", "رقم الموبايل", "الصافي"]);
    }
  }
}

// =========================================================================
// أدوات مساعدة (خاصة بالسيرفر)
// =========================================================================
// الكود أو رقم الموبايل بشكل موحّد: الشيت بيخزن الموبايل رقم من غير الصفر (1149008444)
// والمندوب ممكن يكتبه 01149008444 أو +201149008444 أو بمسافات، فكلهم بيبقوا 1149008444
function canonId_(v) {
  if (v === null || v === undefined) return "";
  var s = String(v).trim().toUpperCase();
  var digits = s.replace(/[\s\-+(),]/g, "");
  if (/^\d+$/.test(digits)) {
    var m = digits.match(/^(?:0020|20)?0?(1\d{9})$/);   // موبايل مصري
    if (m) return m[1];
    return digits;
  }
  return s;
}

// أعمدة الدخول (كود المندوب / الموبايل). لازم الاسم يكون مطابق عشان "كود الفرع" مثلًا متتحسبش
var ID_HEADER_RE_ = /^(الكود|كود|كود الطيار|كود المندوب|كود الموظف|id|rider id|rider_id|employee id|phone|mobile|phone number|رقم الهاتف|رقم الهاتف الشخصي|رقم الموبايل|الموبايل|الهاتف|رقم التليفون|التليفون)$/;
var ID_EXCLUDE_RE_ = /فرع|branch|order|اوردر|مشرف|supervisor|طلب|شركه|company/;
function idCols_(headers) {
  var out = [];
  headers.forEach(function (h, i) { if (ID_HEADER_RE_.test(norm_(h))) out.push(i); });
  if (!out.length) headers.forEach(function (h, i) {
    var x = norm_(h);
    if (!ID_EXCLUDE_RE_.test(x) && /كود|\bid\b|phone|الهاتف|الموبايل/.test(x)) out.push(i);
  });
  if (!out.length) out.push(0);
  return out;
}
// المفتاح الأساسي للمندوب في الصف ده: أول عمود دخول فيه قيمة (الكود قبل الموبايل لو الاتنين موجودين)
function rowKey_(row, cols) {
  for (var i = 0; i < cols.length; i++) { var k = canonId_(row[cols[i]]); if (k) return k; }
  return "";
}
function rowMatches_(row, cols, id) {
  for (var i = 0; i < cols.length; i++) if (canonId_(row[cols[i]]) === id) return true;
  return false;
}

function norm_(t) {
  return t ? t.toString().replace(/ة/g, 'ه').replace(/[أإآ]/g, 'ا').trim().toLowerCase() : "";
}

// google.script.run مش بيرجّع Date صح، فبنحوّلها نص
function clean_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (v === null || v === undefined) return "";
  return v;
}

// =========================================================================
// الكاش: بدل ما نقرا الشيت كله مع كل دخول، بنقراه مرة كل 5 دقايق
// =========================================================================
function cacheVer_() {
  return PropertiesService.getScriptProperties().getProperty('CACHE_VER') || '1';
}

// بيرجّع كل قيم الشيت (التواريخ متحولة لنص) أو null لو الشيت مش موجود
var MEMO_ = {};   // نفس الطلب ميقراش نفس الشيت مرتين
function readSheet_(name) {
  if (MEMO_.hasOwnProperty(name)) return MEMO_[name];
  return (MEMO_[name] = readSheetCached_(name));
}

function readSheetCached_(name) {
  var cache = CacheService.getScriptCache();
  var key = 'sh2_' + cacheVer_() + '_' + Utilities.base64EncodeWebSafe(name);
  try {
    var n = Number(cache.get(key) || 0);
    if (n > 0) {
      var keys = [];
      for (var i = 0; i < n; i++) keys.push(key + '_' + i);
      var parts = cache.getAll(keys), str = '';
      for (var j = 0; j < n; j++) {
        if (parts[keys[j]] == null) { str = null; break; }
        str += parts[keys[j]];
      }
      if (str !== null) return JSON.parse(str);
    }
  } catch (e) {}

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) return null;
  // القيم زي ما هي ظاهرة في الشيت (58% بدل 0.58، و 1,000.00 بنفس التنسيق)
  var data = sheet.getDataRange().getDisplayValues();
  try {
    var s = JSON.stringify(data);
    var count = Math.max(1, Math.ceil(s.length / CACHE_CHUNK));
    if (count <= 60) {
      var obj = {};
      for (var k = 0; k < count; k++) obj[key + '_' + k] = s.substr(k * CACHE_CHUNK, CACHE_CHUNK);
      cache.putAll(obj, CACHE_SECONDS);
      cache.put(key, String(count), CACHE_SECONDS);
    }
  } catch (e) {}
  return data;
}

// شغّلها بعد ما ترفع بيانات جديدة لو مش عايز تستنى 5 دقايق
function clearDataCache() {
  PropertiesService.getScriptProperties().setProperty('CACHE_VER', Utilities.getUuid().slice(0, 8));
  return { success: true };
}

// أي تعديل يدوي في الشيت بيمسح الكاش أوتوماتيك
function onEdit(e) {
  clearDataCache();
}

// =========================================================================
// تسجيل العمليات في شيت Logs
// =========================================================================
function log_(action, company, riderId, ok, details) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(LOG_SHEET);
    if (!sh) { setupDatabase(); sh = ss.getSheetByName(LOG_SHEET); }
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    sh.appendRow([now, action, company || "", riderId ? "'" + String(riderId).slice(0, 40) : "", ok ? "نجح" : "فشل",
      String(details || "").slice(0, 300), CLIENT_]);
    if (Math.random() < 0.02 && sh.getLastRow() > LOG_MAX_ROWS) sh.deleteRows(2, 5000);
  } catch (e) {}
}

function hashPassword_(text) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text));
  var hex = '';
  for (var i = 0; i < digest.length; i++) {
    var b = digest[i]; if (b < 0) b += 256;
    var h = b.toString(16); if (h.length == 1) h = '0' + h;
    hex += h;
  }
  return hex;
}

function fail_(msg, extra) {
  var r = { success: false, message: msg };
  if (extra) for (var k in extra) r[k] = extra[k];
  return r;
}

// منع تخمين الرقم السري
function tooManyFails_(id) { return Number(CacheService.getScriptCache().get('fail_' + id) || 0) >= MAX_FAILS; }
function recordFail_(id) {
  var c = CacheService.getScriptCache();
  c.put('fail_' + id, String(Number(c.get('fail_' + id) || 0) + 1), FAIL_LOCK_SECONDS);
}
function clearFails_(id) { CacheService.getScriptCache().remove('fail_' + id); }

function sheetsFor_(companyName) {
  var n = companyName.toLowerCase();
  if (n === "talabat") return ["Talabat", "Talabat_Canceled", "Talabat_Compensations", "Talabat_Deductions"];
  if (n === "noon minuts") return ["Noon Minuts", "Noon Minuts_Pickers", "Noon Minuts_Logistics"];
  return [companyName];
}

// بيرجّع اسم الشيت اللي المندوب موجود فيه، أو null
function findRider_(ss, companyName, riderId) {
  var list = sheetsFor_(companyName);
  for (var i = 0; i < list.length; i++) {
    var data = readSheet_(list[i]);
    if (!data) continue;
    if (data.length <= 1) continue;
    var cols = idCols_(data[0]);
    for (var r = 1; r < data.length; r++) {
      if (rowMatches_(data[r], cols, riderId)) return { sheet: list[i], key: rowKey_(data[r], cols) || riderId };
    }
  }
  return null;
}

// نفس المندوب لو فتح تاني خلال 5 دقايق بياخد الرد جاهز
function buildPayload_(companyName, sheetName, riderId) {
  var cache = CacheService.getScriptCache();
  var key = 'pl2_' + cacheVer_() + '_' + Utilities.base64EncodeWebSafe(companyName + '|' + sheetName + '|' + riderId);
  try { var hit = cache.get(key); if (hit) return JSON.parse(hit); } catch (e) {}
  var data = buildPayloadFresh_(companyName, sheetName, riderId);
  if (data) { try { var j = JSON.stringify(data); if (j.length < 90000) cache.put(key, j, CACHE_SECONDS); } catch (e) {} }
  return data;
}

function buildPayloadFresh_(companyName, sheetName, riderId) {
  if (companyName.toLowerCase() === "talabat") {
    var t = getTalabatSuperAppData_(riderId);
    t.type = "talabat";
    return t;
  }
  var g = getGenericCompanyData_(sheetName, riderId);
  if (g) g.type = "generic";
  return g;
}

// بيضيف جلسة جديدة للمندوب (عمود E) ويرجّع التوكن
function addSession_(sheet, rowNum, existing) {
  var token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  var now = Date.now();
  var list = String(existing || "").split(",").filter(function (s) {
    var p = s.split("|"); return p.length === 2 && Number(p[1]) > now;
  });
  list.push(hashPassword_(token) + "|" + (now + SESSION_DAYS * 86400000));
  while (list.length > MAX_SESSIONS) list.shift();
  sheet.getRange(rowNum, 5).setValue(list.join(","));
  return token;
}

// =========================================================================
// الدخول بالكود والرقم السري
// confirmNew = true لما المندوب يأكد إنه بيختار رقم سري لأول مرة
// =========================================================================
function checkSalary(companyName, riderId, password, confirmNew) {
  var t0 = Date.now();
  var r = checkSalaryCore_(companyName, riderId, password, confirmNew);
  log_(confirmNew ? "أول دخول (اختيار رقم سري)" : "دخول بالرقم السري", companyName, riderId,
    r.success, (r.success ? "" : (r.needsSetup ? "أول مرة - مستني تأكيد الرقم السري" : r.message)) + " (" + (Date.now() - t0) + "ms)");
  return r;
}

function checkSalaryCore_(companyName, riderId, password, confirmNew) {
  try {
    var id = canonId_(riderId);
    var pass = password ? password.toString().trim() : "";
    companyName = companyName ? companyName.toString().trim() : "";
    if (!companyName || !id || !pass) return fail_("اكتب الكود والرقم السري.");
    if (tooManyFails_(id)) return fail_("محاولات غلط كتير. استنى 10 دقايق وجرب تاني.");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var found = findRider_(ss, companyName, id);
    if (!found) return fail_("الكود ده مش متسجل في بيانات " + companyName + " لحد دلوقتي. اتأكد منه أو كلم المشرف.");
    var sheetName = found.sheet;
    id = found.key;   // لو دخل بالموبايل والشيت فيه كود، الحساب بيتسجل على الكود (حساب واحد للاتنين)
    if (tooManyFails_(id)) return fail_("محاولات غلط كتير. استنى 10 دقايق وجرب تاني.");

    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return fail_("السيرفر عليه ضغط دلوقتي، جرب كمان شوية."); }

    var token;
    try {
      var users = ss.getSheetByName(USERS_SHEET);
      var rows = users.getDataRange().getValues();
      var rowIdx = -1, dup = [];
      for (var u = 1; u < rows.length; u++) {
        if (rows[u][0] != null && canonId_(rows[u][0]) === id) {
          if (rowIdx === -1) rowIdx = u; else dup.push(u + 1);
        }
      }
      for (var d = dup.length - 1; d >= 0; d--) users.deleteRow(dup[d]);

      var hashed = hashPassword_(pass);
      if (rowIdx === -1) {
        if (!confirmNew) {
          return fail_("ده أول دخول ليك. الرقم اللي كتبته هيبقى الرقم السري بتاعك، اكتبه تاني للتأكيد.", { needsSetup: true });
        }
        if (pass.length < 4) return fail_("الرقم السري لازم يكون 4 حروف أو أرقام على الأقل.");
        var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        users.appendRow([id, companyName, hashed, now, ""]);
        token = addSession_(users, users.getLastRow(), "");
      } else {
        if (String(rows[rowIdx][2]) !== hashed) {
          recordFail_(id);
          return fail_("الرقم السري غلط.");
        }
        var reg = rows[rowIdx][1] ? rows[rowIdx][1].toString().trim() : "";
        if (reg && reg.toLowerCase() !== companyName.toLowerCase()) {
          return fail_("إنت متسجل قبل كده على شركة " + reg + ". ادخل من لوجو " + reg + ".");
        }
        token = addSession_(users, rowIdx + 1, rows[rowIdx][4]);
      }
    } finally {
      lock.releaseLock();
    }

    clearFails_(id);
    var data = buildPayload_(companyName, sheetName, id);
    if (!data) return fail_("مفيش تفاصيل مرتب متسجلة لحسابك دلوقتي.");
    return { success: true, data: data, token: token, riderId: id };
  } catch (err) {
    return fail_("حصل خطأ: " + err.message);
  }
}

// الدخول التلقائي من غير رقم سري (بالتوكن المحفوظ على الموبايل)
function loginWithToken(companyName, riderId, token) {
  var t0 = Date.now();
  var r = loginWithTokenCore_(companyName, riderId, token);
  log_("دخول تلقائي", companyName, riderId, r.success, (r.success ? "" : r.message) + " (" + (Date.now() - t0) + "ms)");
  return r;
}

function loginWithTokenCore_(companyName, riderId, token) {
  try {
    var id = canonId_(riderId);
    var tok = token ? token.toString().trim() : "";
    if (!id || !tok || !companyName) return fail_("ادخل بالرقم السري.", { expired: true });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rows = ss.getSheetByName(USERS_SHEET).getDataRange().getValues();
    var h = hashPassword_(tok), now = Date.now(), ok = false;
    for (var u = 1; u < rows.length; u++) {
      if (rows[u][0] != null && canonId_(rows[u][0]) === id) {
        var reg = rows[u][1] ? rows[u][1].toString().trim().toLowerCase() : "";
        if (reg === companyName.toLowerCase()) {
          ok = String(rows[u][4] || "").split(",").some(function (s) {
            var p = s.split("|"); return p[0] === h && Number(p[1]) > now;
          });
        }
        break;
      }
    }
    if (!ok) return fail_("الجلسة خلصت. ادخل بالرقم السري تاني.", { expired: true });

    // نون مينتس بس اللي ليها أكتر من شيت، فمحتاجة ندوّر المندوب في أنهي واحد
    var hit = sheetsFor_(companyName).length > 1 && companyName.toLowerCase() !== "talabat"
      ? findRider_(ss, companyName, id) : { sheet: companyName };
    var sheetName = hit && hit.sheet;
    if (!sheetName) return fail_("الكود ده مش متسجل في بيانات " + companyName + " دلوقتي.");
    var data = buildPayload_(companyName, sheetName, id);
    if (!data) return fail_("مفيش تفاصيل مرتب متسجلة لحسابك دلوقتي.");
    return { success: true, data: data, riderId: id };
  } catch (err) {
    return fail_("حصل خطأ: " + err.message);
  }
}

function logout(riderId, token) {
  try {
    var id = canonId_(riderId);
    if (!id || !token) return { success: true };
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var users = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
      var rows = users.getDataRange().getValues();
      var h = hashPassword_(token.toString());
      for (var u = 1; u < rows.length; u++) {
        if (rows[u][0] != null && canonId_(rows[u][0]) === id) {
          var left = String(rows[u][4] || "").split(",").filter(function (s) { return s && s.split("|")[0] !== h; });
          users.getRange(u + 1, 5).setValue(left.join(","));
          break;
        }
      }
    } finally { lock.releaseLock(); }
    log_("خروج", "", id, true, "");
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

// =========================================================================
// طلبات (Super App)
// =========================================================================
function getTalabatSuperAppData_(riderId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var id = canonId_(riderId);

  var salary = getGenericCompanyData_("Talabat", id) || { rider_name: "-", vehicle_type: "-", weeks: [] };
  var base = { name: salary.rider_name, city: "-", sp: "-", vehicle: salary.vehicle_type };

  var list = ["Talabat", "Talabat_Canceled", "Talabat_Compensations", "Talabat_Deductions"];
  for (var s = 0; s < list.length; s++) {
    var data = readSheet_(list[s]);
    if (!data) continue;
    if (data.length <= 1) continue;
    var h = data[0];
    var find = function (word) { return h.findIndex(function (x) { return norm_(x).indexOf(word) !== -1; }); };
    var idC = h.findIndex(function (x) { var n = norm_(x); return n.indexOf("كود") !== -1 || n.indexOf("id") !== -1; });
    var nameC = find("اسم"), cityC = find("مدين"), spC = find("فتح"), vC = find("مركب");
    if (idC === -1) idC = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][idC] != null && canonId_(data[i][idC]) === id) {
        if ((base.name === "-" || base.name === "غير متوفر") && nameC !== -1) base.name = String(clean_(data[i][nameC]) || "-");
        if (base.city === "-" && cityC !== -1) base.city = String(clean_(data[i][cityC]) || "-");
        if (base.sp === "-" && spC !== -1) base.sp = String(clean_(data[i][spC]) || "-");
        if ((!base.vehicle || base.vehicle === "-" || base.vehicle === "Talabat") && vC !== -1) base.vehicle = String(clean_(data[i][vC]) || "-");
        break;
      }
    }
  }

  return {
    rider_base_info: base,
    salaries: salary.weeks || [],
    canceled: getTalabatSubSheetData_(ss, "Talabat_Canceled", id, ["Order code", "Order Date", "Starting Point", "City", "comment", "Cancelled orders Deduction Amount", "Week of deduction"]),
    compensations: getTalabatSubSheetData_(ss, "Talabat_Compensations", id, ["Order ID", "Order Date", "Starting Point", "City", "Comment", "Compensation Amount", "Week of Compensation"]),
    deductions: getTalabatSubSheetData_(ss, "Talabat_Deductions", id, ["المشرف", "المنطقة", "نقطة الفتح", "خصم المشرف", "الاسبوع", "الشهر"])
  };
}

function getTalabatSubSheetData_(ss, sheetName, riderId, reqCols) {
  var data = readSheet_(sheetName);
  if (!data) return {};
  if (data.length <= 1) return {};

  var headers = data[0], colMap = {};
  for (var i = 0; i < reqCols.length; i++) {
    var want = norm_(reqCols[i]);
    var idx = headers.findIndex(function (h) { return norm_(h).indexOf(want) !== -1; });
    colMap[reqCols[i]] = idx;
  }
  var riderCol = headers.findIndex(function (h) { var n = norm_(h); return n.indexOf("rider id") !== -1 || n.indexOf("كود") !== -1; });
  if (riderCol === -1) riderCol = 0;

  var groupCol = reqCols[reqCols.length - 1], grouped = {};
  for (var r = 1; r < data.length; r++) {
    if (data[r][riderCol] == null || canonId_(data[r][riderCol]) !== riderId) continue;
    var g = colMap[groupCol] !== -1 ? String(clean_(data[r][colMap[groupCol]])).trim() : "";
    if (!g) g = "أخرى";
    if (!grouped[g]) grouped[g] = [];
    var item = {};
    for (var c = 0; c < reqCols.length - 1; c++) {
      var ci = colMap[reqCols[c]];
      var val = ci !== -1 ? clean_(data[r][ci]) : "-";
      if (val === "") val = "-";
      if (reqCols[c].toLowerCase().indexOf("comment") !== -1) val = translateTalabatReason_(String(val));
      item[reqCols[c]] = val;
    }
    grouped[g].push(item);
  }
  return grouped;
}

function translateTalabatReason_(text) {
  if (!text || text === "-") return "غير محدد";
  var l = text.toLowerCase();
  if (l.indexOf("unprofessional behavior") !== -1) return "سلوك غير مهني من المندوب";
  if (l.indexOf("spilt") !== -1 || l.indexOf("damaged") !== -1) return "تلف أو انسكاب الأوردر";
  if (l.indexOf("late delivery") !== -1) return "تأخير في التوصيل";
  if (l.indexOf("mixed up") !== -1) return "تبديل الأوردرات بالخطأ";
  if (l.indexOf("por is not acceptable") !== -1 || l.indexOf("no por") !== -1) return "إثبات التوصيل (POR) غير مقبول أو مفقود";
  if (l.indexOf("picked up by another") !== -1) return "تم استلام الأوردر بواسطة مندوب آخر";
  if (l.indexOf("lack of delivery") !== -1) return "نقص في عدد المناديب";
  if (l.indexOf("missing/wrong items") !== -1) return "أصناف ناقصة أو خاطئة (خطأ من المطعم)";
  if (l.indexOf("failing to punch") !== -1 || l.indexOf("missed") !== -1) return "عدم تسجيل/تأكيد الأوردر (Missed)";
  if (l.indexOf("successfully delivery") !== -1) return "تم التوصيل بنجاح";
  if (l.indexOf("duplicate order") !== -1) return "أوردر مكرر بالخطأ";
  return text;
}

// =========================================================================
// الدالة العامة لكل الشركات
// =========================================================================
function getGenericCompanyData_(sheetName, riderId) {
  var data = readSheet_(sheetName);
  if (!data) return null;
  if (data.length <= 1) return null;
  var headers = data[0];

  function findCol(names) {
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i].toString().trim();
      for (var j = 0; j < names.length; j++) if (h === names[j] || h.indexOf(names[j]) !== -1) return i;
    }
    return -1;
  }

  // عمود بالاسم بالظبط أو بيبدأ بيه (عشان "عمولة الشهر" متتحسبش عمود الشهر)
  function findColExact(names) {
    for (var i = 0; i < headers.length; i++) {
      var h = norm_(headers[i]);
      for (var j = 0; j < names.length; j++) { var n = norm_(names[j]); if (h === n || h.indexOf(n + " ") === 0) return i; }
    }
    return -1;
  }
  var idCols = idCols_(headers);
  var colName = findCol(["اسم الطيار", "اسم المندوب", "الاسم", "Name", "الاسـم"]);
  if (colName === -1 && headers.length > 1) colName = 1;
  var colVehicle = findCol(["نوع المركبة", "المركبة", "Vehicle", "المعده"]);
  var colComp = findColExact(["اسم الشركة", "الشركة", "الفرع", "كود الفرع"]);
  var colWeek = findColExact(["الاسبوع", "الأسبوع", "الفترة", "الشهر", "Week"]);
  var colWallet = findCol(["صافي القبض الظاهر على محفظة الطيار", "صافي المحفظة", "صافي القبض", "صافي الراتب", "صافي الراتب المستحق"]);
  var colNet = findCol(["الصافي", "المبلغ المستحق", "صافي الراتب"]);

  var weeks = [], riderName = "-", vehicle = "", branch = "";
  // عناوين الأعمدة المتكررة (زي "الراتب" مرتين) بتاخد رقم عشان تتفرق
  var titles = [], seenT = {};
  headers.forEach(function (h) {
    var t = h ? h.toString().trim() : "";
    if (t) { seenT[t] = (seenT[t] || 0) + 1; if (seenT[t] > 1) t += " (" + seenT[t] + ")"; }
    titles.push(t);
  });
  var id = canonId_(riderId);

  for (var r = 1; r < data.length; r++) {
    if (!rowMatches_(data[r], idCols, id)) continue;
    if (colName !== -1 && data[r][colName]) riderName = String(clean_(data[r][colName])).replace(/\s*[-–_]+\s*full\s*speed.*$/i, "").trim() || riderName;
    if (colVehicle !== -1 && data[r][colVehicle]) vehicle = String(clean_(data[r][colVehicle])).trim();
    if (colComp !== -1 && data[r][colComp]) branch = String(clean_(data[r][colComp])).trim();

    var week = colWeek !== -1 ? String(clean_(data[r][colWeek])).trim() : "";
    var total = "0", top = week ? "الأسبوع: " + week : "صافي المستحقات", sub = "صافي المستحقات", usedCol = -1;

    if (colWallet !== -1 && String(clean_(data[r][colWallet])).trim() !== "") {
      total = data[r][colWallet]; usedCol = colWallet;
      // جملة المحفظة لطلبات بس، الباقي بياخد اسم العمود نفسه
      sub = sheetName === "Talabat" ? "صافي القبض الظاهر على المحفظة" : String(headers[colWallet]).trim();
    } else if (colNet !== -1 && String(clean_(data[r][colNet])).trim() !== "") {
      total = data[r][colNet]; usedCol = colNet; sub = String(headers[colNet]).trim();
    } else {
      total = data[r][headers.length - 1];
    }

    var details = [], skip = idCols.concat([colName, colVehicle, colWeek, colComp]);
    headers.forEach(function (h, i) { if (/^(vendor|المورد|الشركه الموردة)$/i.test(norm_(h))) skip.push(i); });
    if (usedCol !== -1) skip.push(usedCol);
    for (var c = 0; c < headers.length; c++) {
      var title = titles[c];
      if (skip.indexOf(c) !== -1 || !title) continue;
      var val = clean_(data[r][c]);
      if (val.toString().trim() === "") val = "-";
      if (typeof val === 'number' && val % 1 !== 0) val = Number(val.toFixed(2));
      details.push({ title: title, value: val });
    }

    weeks.push({
      week_id: "week_" + r,
      week_label: week ? "الأسبوع " + week : (weeks.length > 0 ? "دفعة " + (weeks.length + 1) : "الراتب الحالي"),
      top_title: top,
      sub_title: sub,
      final_total: clean_(total) === "" ? 0 : clean_(total),
      details: details
    });
  }
  if (!weeks.length) return null;
  return { rider_name: riderName, vehicle_type: vehicle, branch: branch, weeks: weeks };
}
