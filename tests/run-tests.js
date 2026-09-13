/**
 * ============================================================
 *  مجموعه تست شبیه‌سازی کامل سامانه پرمیت (بک‌اند v16.5)
 *  اجرا:  TZ=Asia/Tehran node tests/run-tests.js
 * ============================================================
 *  همه اکشن‌ها، مجوزها، جریان‌های وضعیت، بستن خودکار پرمیت در
 *  مرز شیفت‌ها (۱۹:۰۰ روز / ۰۷:۰۰ شب)، عبور از مرز ماه/سال شمسی،
 *  قابلیت‌های جدید (چک‌این، گزارش عملکرد، کاتالوگ، لوگو، خروجی)،
 *  و یکپارچگی سامانه چک‌لیست درکاو شبیه‌سازی و تست می‌شوند.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const mock = require('./gas-mock');

/* ---------------- بارگذاری اسکریپت بک‌اند در محیط شبیه‌سازی‌شده ---------------- */
function loadBackend() {
  const src = fs.readFileSync(path.join(__dirname, '..', 'app.html'), 'utf8');
  const ctx = {
    SpreadsheetApp: mock.SpreadsheetApp,
    DriveApp: mock.DriveApp,
    PropertiesService: mock.PropertiesService,
    ScriptApp: mock.ScriptApp,
    LockService: mock.LockService,
    ContentService: mock.ContentService,
    HtmlService: mock.HtmlService,
    Session: mock.Session,
    Utilities: mock.Utilities,
    Logger: mock.Logger,
    UrlFetchApp: mock.UrlFetchApp,
    DocumentApp: mock.DocumentApp,
    CacheService: mock.CacheService,
    MimeType: mock.MimeType,
    Date: mock.FakeDate, JSON: JSON, Math: Math, console: console,
    parseInt: parseInt, parseFloat: parseFloat, isNaN: isNaN, String: String,
    Object: Object, Array: Array, RegExp: RegExp, Error: Error
  };
  vm.createContext(ctx);
  vm.runInContext(src, ctx, { filename: 'app.html' });
  return ctx;
}

/* ---------------- ابزارهای تست ---------------- */
let PASS = 0, FAIL = 0;
const failures = [];
function check(cond, title) {
  if (cond) { PASS++; console.log('  ✅ ' + title); }
  else { FAIL++; failures.push(title); console.log('  ❌ ' + title); }
}
function section(t) { console.log('\n━━━ ' + t + ' ━━━'); }

function respParse(out) {
  // خروجی ContentService → JSON
  const txt = out.getContent();
  try { return JSON.parse(txt); } catch (e) { return { _raw: txt }; }
}
function T(y, mo, d, h, mi) { return new Date(y, mo - 1, d, h || 12, mi || 0, 0); }

/* ==================================================================== */
console.log('🧪 شبیه‌سازی کامل بک‌اند سامانه پرمیت درکاو (v16.5 بازسازی‌شده)');
console.log('   امروز در تست: 2026-09-13 (= 1405-06-22 شمسی)\n');

const G = loadBackend();
const doGet = G.doGet, doPost = G.doPost;
function post(body) { return respParse(doPost({ postData: { contents: JSON.stringify(body) } })); }
function get(params) { return respParse(doGet({ parameter: params })); }

const ADMIN = '96792582';
const USER_A = '13619873'; // بدون مجوز
const USER_B = '15459515'; // بعداً مجوزها به او داده می‌شود

/* ---------- گروه ۰: آماده‌سازی ---------- */
section('گروه ۰ — آماده‌سازی شیت‌ها');
mock.resetWorld({});
G._doSetup(false);
check(G._ss().getSheetByName('پرمیت') !== null, 'شیت «پرمیت» ساخته شد');
check(G._ss().getSheetByName('کاربران') !== null, 'شیت «کاربران» ساخته شد');
check(G._ss().getSheetByName('بازدیدها') !== null, 'شیت «بازدیدها» ساخته شد');
check(G._ss().getSheetByName('دانش_HSE') !== null, 'شیت «دانش_HSE» ساخته شد');

/* ---------- گروه ۱: پینگ و لایه‌های ارتباطی ---------- */
section('گروه ۱ — پینگ، لایه‌ها و سازگاری‌ها');
mock.setFakeNow(T(2026, 9, 13, 9, 0));
let r = get({ action: 'ping' });
check(r.ok === true, 'پینگ GET موفق است');
check(r.version === 'ptw-v16.5', 'نسخه سرور ptw-v16.5 اعلام می‌شود');
check(r.apiVersion === 'v16.5', 'فیلد apiVersion = v16.5 (انتظار اپ)');
check(r.status === 'success', 'فیلد status=success (سازگاری چک‌لیست)');
r = post({ action: 'ping' });
check(r.ok === true && r.version === 'ptw-v16.5', 'پینگ POST (لایه ۴) موفق است');
r = post({ action: 'test' });
check(r.ok === true, 'اکشن test هم پاسخ می‌دهد');
// لایه ۲ (JSONP)
let out = doGet({ parameter: { action: 'ping', callback: 'cb123' } });
check(out.getContent().indexOf('cb123(') === 0, 'لایه ۲: پاسخ JSONP با کال‌بک پیچیده می‌شود');
// لایه ۱ (iFrame)
out = doGet({ parameter: { action: 'ping', msgId: 'M-1' } });
check(out.getContent().indexOf('postMessage') !== -1 && out.getContent().indexOf('M-1') !== -1, 'لایه ۱: پاسخ postMessage برای iFrame');
// اکشن ناشناخته
r = post({ action: 'doesnotexist' });
check(r.ok === false && String(r.error).length > 0, 'اکشن ناشناخته در POST خطا می‌دهد');
r = get({ action: 'doesnotexist' });
check(r.ok === true && Array.isArray(r.rows), 'اکشن ناشناخته در GET = لیست پرمیت‌ها (سازگاری قدیمی)');

/* ---------- گروه ۲: ورود و احراز هویت ---------- */
section('گروه ۲ — ورود، احراز هویت و مدیریت کاربران');
r = post({ action: 'login', code: ADMIN, firstName: 'فرشید', lastName: 'مدیر' });
check(r.ok === true && r.user && r.user.isAdmin === true, 'ورود مدیر ارشد موفق + فلگ isAdmin');
check(r.user.canApprove === true && r.user.canCreatePermit === true, 'مدیر ارشد همه مجوزها را دارد');
r = post({ action: 'login', code: USER_A, firstName: 'علی', lastName: 'بازرس' });
check(r.ok === true && r.user.isAdmin === false, 'ورود کاربر عادی موفق');
check(r.user.canApprove === false, 'کاربر عادی به‌طور پیش‌فرض مجوز تأیید ندارد');
check(r.user.loginCount === 1, 'تعداد ورود کاربر جدید = ۱');
r = post({ action: 'login', code: '99999999', firstName: 'x', lastName: 'y' });
check(r.ok === false, 'کد خارج از لیست مجاز رد می‌شود');
r = post({ action: 'login', code: '1234567', firstName: 'x', lastName: 'y' });
check(r.ok === false, 'کد غیر ۸ رقمی رد می‌شود');
r = post({ action: 'login', code: USER_A, firstName: '', lastName: '' });
check(r.ok === false, 'ورود بدون نام رد می‌شود');
// لیست کاربران فقط مدیر
r = post({ action: 'listUsers', code: USER_A });
check(r.ok === false, 'listUsers برای کاربر عادی ممنوع است');
r = post({ action: 'listUsers', code: ADMIN });
check(r.ok === true && r.users.length === 12 && r.permDefs.length === 10, 'listUsers مدیر: ۱۲ کاربر + ۱۰ تعریف مجوز');
// اعطای مجوز
r = post({ action: 'setgrant', code: ADMIN, targetCode: USER_B, perm: 'createPermit', grant: true });
check(r.ok === true, 'اعطای مجوز ثبت پرمیت به کاربر B');
post({ action: 'setgrant', code: ADMIN, targetCode: USER_B, perm: 'approve', grant: true });
post({ action: 'setgrant', code: ADMIN, targetCode: USER_B, perm: 'reject', grant: true });
post({ action: 'setgrant', code: ADMIN, targetCode: USER_B, perm: 'close', grant: true });
post({ action: 'setgrant', code: ADMIN, targetCode: USER_B, perm: 'reactivate', grant: true });
post({ action: 'setgrant', code: ADMIN, targetCode: USER_B, perm: 'patrol', grant: true });
post({ action: 'setgrant', code: ADMIN, targetCode: USER_B, perm: 'reinspect', grant: true });
post({ action: 'setgrant', code: ADMIN, targetCode: USER_B, perm: 'delete', grant: true });
post({ action: 'setgrant', code: ADMIN, targetCode: USER_B, perm: 'units', grant: true });
post({ action: 'setgrant', code: ADMIN, targetCode: USER_B, perm: 'analyze', grant: true });
r = post({ action: 'setgrant', code: USER_A, targetCode: USER_B, perm: 'approve', grant: true });
check(r.ok === false, 'اعطای مجوز توسط غیرمدیر رد می‌شود');
r = post({ action: 'setgrant', code: ADMIN, targetCode: ADMIN, perm: 'approve', grant: true });
check(r.ok === false, 'تنظیم مجوز برای خود مدیر ارشد بی‌معنی است و رد می‌شود');
r = post({ action: 'setuseraccess', code: ADMIN, targetCode: ADMIN, active: false });
check(r.ok === false, 'نمی‌توان دسترسی مدیر ارشد را قطع کرد');
// سازگاری اکشن‌های قدیمی
r = post({ action: 'setreactivation', code: ADMIN, targetCode: '25502866', grant: 'true' });
check(r.ok === true && r.perm === 'reactivate', 'اکشن قدیمی setreactivation هنوز کار می‌کند');
r = post({ action: 'setcreatepermit', code: ADMIN, targetCode: '25502866', grant: 'true' });
check(r.ok === true && r.perm === 'createPermit', 'اکشن قدیمی setcreatepermit هنوز کار می‌کند');

/* ---------- گروه ۳: ثبت و ویرایش پرمیت ---------- */
section('گروه ۳ — ثبت، ویرایش و حذف پرمیت');
mock.setFakeNow(T(2026, 9, 13, 8, 30));
r = post({ action: 'createPermit', code: USER_A, permitNo: '1405-0601', startDate: '1405/06/22', startTime: '08:00', place: 'سالن فلوتاسیون', type: 'روتین', desc: 'تعمیر پمپ', unit: 'مکانیک', contractor: 'زرین صنعت' });
check(r.ok === false && r.needPerm === 'createPermit', 'ثبت پرمیت بدون مجوز → رد با پیام مجوز');
r = post({ action: 'createPermit', code: ADMIN, firstName: 'فرشید', lastName: 'مدیر', permitNo: '1405-0601', startDate: '1405/06/22', startTime: '08:00', place: 'سالن فلوتاسیون', type: 'روتین', desc: 'تعمیر پمپ اسلاری', unit: 'مکانیک', endDate: '', endTime: '19:00', contractor: 'زرین صنعت' });
check(r.ok === true && r.status === 'در انتظار تایید', 'ثبت پرمیت توسط مدیر → در انتظار تایید');
r = post({ action: 'createPermit', code: ADMIN, permitNo: '1405-0601', startDate: '1405-06-22', startTime: '08:00', place: 'x', type: 'y', desc: 'z', contractor: 'c' });
check(r.ok === false && /استفاده/.test(r.error), 'شماره پرمیت تکراری رد می‌شود');
r = post({ action: 'createPermit', code: ADMIN, permitNo: '1405-0602', startDate: '', startTime: '08:00', place: 'x', type: 'y', desc: 'z', contractor: 'c' });
check(r.ok === false, 'ثبت بدون تاریخ شروع رد می‌شود');
r = post({ action: 'createPermit', code: ADMIN, permitNo: '1405-0603', startDate: '1405/06/22', startTime: '25:99', place: 'x', type: 'y', desc: 'z', contractor: 'c' });
check(r.ok === false, 'ساعت شروع نامعتبر رد می‌شود');
// پرمیت شیفت شب (شروع ۲۰:۰۰) برای تست‌های بستن خودکار
post({ action: 'createPermit', code: ADMIN, permitNo: '1405-0604', startDate: '1405/06/22', startTime: '20:00', place: 'تیکندر', type: 'روتین', desc: 'تمیزکاری', unit: 'تولید', endDate: '', endTime: '19:00', contractor: 'رسم آرا' });
// پرمیت بامدادی (شروع ۰۳:۰۰)
post({ action: 'createPermit', code: ADMIN, permitNo: '1405-0605', startDate: '1405/06/22', startTime: '03:00', place: 'سنگ‌شکن', type: 'روتین', desc: 'نگهداری', unit: 'مکانیک', endDate: '', endTime: '19:00', contractor: 'نبکو/NEPCO' });
// پرمیت با اتمام صریح فردا
post({ action: 'createPermit', code: ADMIN, permitNo: '1405-0606', startDate: '1405/06/22', startTime: '09:00', place: 'انبار', type: 'روتین', desc: 'بارگیری', unit: 'انبار', endDate: '1405-06-23', endTime: '12:00', contractor: 'سایر' });
r = post({ action: 'createPermit', code: USER_B, firstName: 'رضا', lastName: 'کارگر', permitNo: '1405-0607', startDate: '1405/06/22', startTime: '10:00', place: 'کارگاه', type: 'روتین', desc: 'جوشکاری نرده', unit: 'ساخت/کارگاه ساخت', contractor: 'پیمانکاری جلال مرادی' });
check(r.ok === true, 'کاربر دارای مجوز می‌تواند پرمیت ثبت کند');
// ویرایش
r = post({ action: 'editPermit', code: ADMIN, permitNo: '1405-0601', place: 'سالن آسیاب', desc: 'تعویض لاینر بالمیل', type: 'غیر روتین' });
check(r.ok === true && r.editedFields >= 2, 'ویرایش پرمیت (اکشن جدید editPermit) موفق');
let rowsNow = post({ action: 'permits', code: ADMIN }).rows;
let p601 = rowsNow.find(function(x) { return x.permitNo === '1405-0601'; });
check(p601.place === 'سالن آسیاب' && p601.desc === 'تعویض لاینر بالمیل' && p601.type === 'غیر روتین', 'فیلدهای ویرایش‌شده در شیت ذخیره شدند');
check(p601.permitStatus === 'در انتظار تایید', 'وضعیت پرمیت بعد از ویرایش دست‌نخورده ماند');
r = post({ action: 'editPermit', code: USER_A, permitNo: '1405-0601', place: 'هک' });
check(r.ok === false, 'ویرایش بدون مجوز رد می‌شود');
r = post({ action: 'editPermit', code: ADMIN, permitNo: '9999', place: 'x' });
check(r.ok === false, 'ویرایش پرمیت ناموجود خطا می‌دهد');
// حذف
r = post({ action: 'deletePermit', code: USER_A, permitNo: '1405-0607' });
check(r.ok === false, 'حذف بدون مجوز رد می‌شود');
r = post({ action: 'deletePermit', code: USER_B, permitNo: '1405-0607', reason: 'تست حذف' });
check(r.ok === true, 'حذف پرمیت با مجوز موفق');
check(post({ action: 'permits', code: ADMIN }).rows.every(function(x) { return x.permitNo !== '1405-0607'; }), 'پرمیت حذف‌شده دیگر در لیست نیست');

/* ---------- گروه ۴: جریان وضعیت‌ها ---------- */
section('گروه ۴ — تأیید / رد / بازرسی مجدد / بستن دستی');
r = post({ action: 'approvePermit', code: USER_A, permitNo: '1405-0601' });
check(r.ok === false && r.needPerm === 'approve', 'تأیید بدون مجوز رد می‌شود');
r = post({ action: 'approvePermit', code: USER_B, permitNo: '1405-0601' });
check(r.ok === true, 'تأیید با مجوز موفق');
rowsNow = post({ action: 'permits', code: ADMIN }).rows;
p601 = rowsNow.find(function(x) { return x.permitNo === '1405-0601'; });
check(p601.permitStatus === 'فعال', 'وضعیت به «فعال» تغییر کرد');
r = post({ action: 'rejectPermit', code: USER_B, permitNo: '1405-0604', reason: 'عدم رعایت LOTO' });
check(r.ok === true, 'رد پرمیت با علت موفق');
rowsNow = post({ action: 'permits', code: ADMIN }).rows;
let p604 = rowsNow.find(function(x) { return x.permitNo === '1405-0604'; });
check(p604.permitStatus === 'رد شده' && /عدم رعایت LOTO/.test(p604.techReject), 'وضعیت «رد شده» + علت در ستون رد تایید فنی');
r = post({ action: 'reinspectPermit', code: USER_B, permitNo: '1405-0604', notes: 'ایرادات رفع شد', newStatus: 'فعال' });
check(r.ok === true && r.newStatus === 'فعال', 'بازرسی مجدد → فعال');
r = post({ action: 'closePermit', code: USER_B, permitNo: '1405-0604', reason: 'اتمام کار' });
check(r.ok === true, 'بستن دستی پرمیت موفق');
rowsNow = post({ action: 'permits', code: ADMIN }).rows;
p604 = rowsNow.find(function(x) { return x.permitNo === '1405-0604'; });
check(p604.permitStatus === 'بسته', 'وضعیت «بسته» ثبت شد');
r = post({ action: 'approvePermit', code: ADMIN, permitNo: '1404-9999' });
check(r.ok === false && /یافت نشد/.test(r.error), 'تأیید پرمیت ناموجود خطای مناسب می‌دهد');

/* ---------- گروه ۵: تمدید / بازفعال‌سازی ---------- */
section('گروه ۵ — تمدید پرمیت تا پایان شیفت (سقف ۴ بار)');
mock.setFakeNow(T(2026, 9, 13, 10, 0));
r = post({ action: 'reactivatePermit', code: USER_A, permitNo: '1405-0604' });
check(r.ok === false && r.needPerm === 'reactivate', 'تمدید بدون مجوز رد می‌شود');
r = post({ action: 'reactivatePermit', code: USER_B, permitNo: '1405-0604' });
check(r.ok === true && r.endDate === '1405/06/22' && r.endTime === '19:00', 'تمدید در شیفت روز → تا ۱۹:۰۰ امروز');
check(r.extendCount === 1 && r.remainingExtends === 3, 'شمارش تمدید: دفعه ۱ از ۴');
r = post({ action: 'reactivatePermit', code: USER_B, permitNo: '1405-0601' });
check(r.ok === false && /فقط پرمیت‌های/.test(r.error), 'تمدید پرمیت فعال رد می‌شود');
// رسیدن به سقف
post({ action: 'closePermit', code: USER_B, permitNo: '1405-0604' });
post({ action: 'reactivatePermit', code: USER_B, permitNo: '1405-0604' }); // ۲
post({ action: 'closePermit', code: USER_B, permitNo: '1405-0604' });
post({ action: 'reactivatePermit', code: USER_B, permitNo: '1405-0604' }); // ۳
post({ action: 'closePermit', code: USER_B, permitNo: '1405-0604' });
post({ action: 'reactivatePermit', code: USER_B, permitNo: '1405-0604' }); // ۴
post({ action: 'closePermit', code: USER_B, permitNo: '1405-0604' });
r = post({ action: 'reactivatePermit', code: USER_B, permitNo: '1405-0604' }); // ۵ → باید رد شود
check(r.ok === false && /حداکثر/.test(r.error), 'کاربر عادی در تمدید پنجم متوقف می‌شود (سقف ۴)');
r = post({ action: 'reactivatePermit', code: ADMIN, permitNo: '1405-0604' });
check(r.ok === true && r.isAdmin === true, 'مدیر ارشد بدون محدودیت تمدید می‌کند');
// تمدید در شیفت شب → بامداد روز بعد
mock.setFakeNow(T(2026, 9, 13, 21, 30));
post({ action: 'closePermit', code: ADMIN, permitNo: '1405-0604' });
r = post({ action: 'reactivatePermit', code: ADMIN, permitNo: '1405-0604' });
check(r.endDate === '1405/06/23' && r.endTime === '07:00', 'تمدید در شیفت شب → تا ۰۷:۰۰ روز بعد');

/* ---------- گروه ۶: ⏱️ بستن خودکار پرمیت (قلب درخواست کاربر) ---------- */
section('گروه ۶ — بستن خودکار پرمیت: همه حالات شیفت و مرزهای تقویمی');
// صحنه را بچینیم: پرمیت‌های تازه با وضعیت فعال
mock.resetWorld({});
G._doSetup(false);
function seed(permitNo, startDate, startTime, endDate, endTime, status) {
  const sh = G._permitSheet();
  sh.appendRow([1, startDate, startTime, 'محل-' + permitNo, 'روتین', 'شرح', 'تولید', endDate || '', endTime || '', '', '', '', permitNo, 'مجری-' + permitNo, status || 'فعال']);
}
seed('D-1', '1405-06-22', '08:00', '', '', 'فعال');          // روز: ۱۹:۰۰ همان روز بسته می‌شود
seed('N-1', '1405-06-22', '20:00', '', '19:00', 'فعال');      // شب: باگ قدیمی! باید تا ۰۷:۰۰ فردا باز بماند
seed('M-1', '1405-06-22', '03:00', '', '', 'فعال');           // بامدادی: ۰۷:۰۰ همان روز
seed('E-1', '1405-06-22', '09:00', '1405-06-24', '12:00', 'فعال'); // اتمام صریح
seed('R-1', '1405-06-22', '09:00', '', '', 'رد شده');          // هرگز دست نمی‌خورد
seed('C-1', '1405-06-22', '09:00', '', '', 'بسته');            // هرگز دست نمی‌خورد

function statusOf(no) {
  const vals = G._permitSheet().getDataRange().getDisplayValues();
  for (let i = 1; i < vals.length; i++) if (String(vals[i][12]).trim() === no) return String(vals[i][14]).trim();
  return null;
}

// ۱) صبح زود (۰۶:۵۹) — هنوز هیچ‌کدام بسته نشوند
mock.setFakeNow(T(2026, 9, 13, 6, 59));
check(G.autoCloseExpiredPermits() === 0, '۰۶:۵۹ — هنوز هیچ پرمیتی منقضی نشده');
check(statusOf('M-1') === 'فعال', '۰۶:۵۹ — پرمیت بامدادی (شروع ۰۳:۰۰) باز است');

// ۲) بعد از ۰۷:۰۰ — فقط بامدادی بسته شود
mock.setFakeNow(T(2026, 9, 13, 7, 1));
check(G.autoCloseExpiredPermits() === 1, '۰۷:۰۱ — فقط پرمیت بامدادی بسته شد');
check(statusOf('M-1') === 'بسته', '۰۷:۰۱ — پرمیت بامدادی بسته شد');
check(statusOf('D-1') === 'فعال' && statusOf('N-1') === 'فعال', 'پرمیت‌های روز و شب هنوز بازند');

// ۳) قبل از ۱۹:۰۰ — روزی بسته نشود
mock.setFakeNow(T(2026, 9, 13, 18, 59));
check(G.autoCloseExpiredPermits() === 0, '۱۸:۵۹ — پرمیت شیفت روز هنوز باز است');

// ۴) بعد از ۱۹:۰۰ — روز بسته شود، شب (با وجود ساعت اتمام ۱۹:۰۰ بدون تاریخ) باز بماند
mock.setFakeNow(T(2026, 9, 13, 19, 1));
check(G.autoCloseExpiredPermits() === 1, '۱۹:۰۱ — فقط پرمیت شیفت روز بسته شد');
check(statusOf('D-1') === 'بسته', '۱۹:۰۱ — پرمیت شیفت روز بسته شد (۱۲ ساعت کامل شد)');
check(statusOf('N-1') === 'فعال', '۱۹:۰۱ — پرمیت شیفت شب با وجود ساعت اتمام ۱۹:۰۰ (بدون تاریخ) بسته نشد ← رفع باگ قدیمی');
check(statusOf('E-1') === 'فعال', 'پرمیت با اتمام صریح هنوز باز است');
check(statusOf('R-1') === 'رد شده' && statusOf('C-1') === 'بسته', 'وضعیت‌های رد/بسته هرگز تغییر نمی‌کنند');

// ۵) بامداد روز بعد: شبی تا ۰۷:۰۰ باز بماند و سپس بسته شود
mock.setFakeNow(T(2026, 9, 14, 6, 59));
check(G.autoCloseExpiredPermits() === 0, '۰۶:۵۹ فردا — پرمیت شیفت شب هنوز معتبر است');
mock.setFakeNow(T(2026, 9, 14, 7, 1));
check(G.autoCloseExpiredPermits() === 1, '۰۷:۰۱ فردا — پرمیت شیفت شب سر ساعت مقرر بسته شد');
check(statusOf('N-1') === 'بسته', 'پرمیت شیفت شب بسته شد');

// ۶) اتمام صریح: بعدازظهر روز سوم
mock.setFakeNow(T(2026, 9, 15, 11, 59));
check(G.autoCloseExpiredPermits() === 0, 'روز سوم ۱۱:۵۹ — پرمیت اتمام‌صریح باز');
mock.setFakeNow(T(2026, 9, 15, 12, 1));
check(G.autoCloseExpiredPermits() === 1, 'روز سوم ۱۲:۰۱ — در لحظه اتمام صریح بسته شد');
check(statusOf('E-1') === 'بسته', 'پرمیت اتمام‌صریح بسته شد');

// ۷) عبور از مرز ماه شمسی (شب ۳۱ شهریور → ۷ صبح ۱ مهر)
mock.resetWorld({});
G._doSetup(false);
seed('NY-1', '1405-06-31', '20:00', '', '', 'فعال');
mock.setFakeNow(T(2026, 9, 22, 23, 0)); // شب ۳۱ شهریور
check(G.autoCloseExpiredPermits() === 0, 'شب ۳۱ شهریور — پرمیت شب باز است');
mock.setFakeNow(T(2026, 9, 23, 7, 15)); // بامداد ۱ مهر
check(G.autoCloseExpiredPermits() === 1, 'بامداد ۱ مهر — عبور از مرز ماه شمسی درست محاسبه شد و بسته شد');
// ۸) عبور از مرز سال کبیسه شمسی (شب ۳۰ اسفند ۱۴۰۳ → ۱ فروردین ۱۴۰۴)
mock.resetWorld({});
G._doSetup(false);
seed('YR-1', '1403-12-30', '20:00', '', '', 'فعال');
const expYR = G._computeExpiry('1403-12-30', '20:00', '', '');
check(expYR.date === '1404/01/01' && expYR.time === '07:00', 'انقضای شیفت شبِ ۳۰ اسفند کبیسه = ۰۷:۰۰ اول فروردین سال بعد');
check(G._computeExpiry('1403-12-29', '20:00', '', '').date === '1403/12/30', 'سال کبیسه: ۲۹ اسفند → ۳۰ اسفند (اسفند ۳۰ روزه)');
mock.setFakeNow(new Date(2025, 2, 20, 23, 0)); // شب ۳۰ اسفند ۱۴۰۳
check(G.autoCloseExpiredPermits() === 0, 'شب پایان سال کبیسه — پرمیت باز است');
mock.setFakeNow(new Date(2025, 2, 21, 7, 10)); // بامداد ۱ فروردین ۱۴۰۴
check(G.autoCloseExpiredPermits() === 1, 'بامداد ۱ فروردین — پرمیت شبِ پایان سال بسته شد');

// ۷) ماتریس واحد _computeExpiry
section('گروه ۶ب — ماتریس تابع انقضا (_computeExpiry)');
function exp(sD, sT, eD, eT) { const x = G._computeExpiry(sD, sT, eD, eT); return x ? (x.date + ' ' + x.time) : 'null'; }
check(exp('1405-06-22', '07:00', '', '') === '1405/06/22 19:00', 'شروع ۰۷:۰۰ → ۱۹:۰۰ همان روز');
check(exp('1405-06-22', '18:59', '', '') === '1405/06/22 19:00', 'شروع ۱۸:۵۹ → ۱۹:۰۰ همان روز');
check(exp('1405-06-22', '19:00', '', '') === '1405/06/23 07:00', 'شروع ۱۹:۰۰ → ۰۷:۰۰ روز بعد');
check(exp('1405-06-22', '23:30', '', '') === '1405/06/23 07:00', 'شروع ۲۳:۳۰ → ۰۷:۰۰ روز بعد');
check(exp('1405-06-22', '02:00', '', '') === '1405/06/22 07:00', 'شروع ۰۲:۰۰ → ۰۷:۰۰ همان روز');
check(exp('1405-06-22', '', '', '') === '1405/06/22 19:00', 'بدون ساعت شروع → پیش‌فرض شیفت روز');
check(exp('1405-06-22', '08:00', '1405-06-25', '') === '1405/06/25 19:00', 'تاریخ اتمام بدون ساعت → ۱۹:۰۰ همان تاریخ');
check(exp('1405-06-22', '08:00', '1405-06-25', '14:30') === '1405/06/25 14:30', 'اتمام صریح (تاریخ+ساعت) دقیقاً رعایت می‌شود');
check(exp('1405-06-31', '20:00', '', '') === '1405/07/01 07:00', 'عبور از مرز ماه (۳۱→۱)');
check(exp('1405-12-29', '20:00', '', '') === G._addDaysToJalali('1405/12/29', 1) + ' 07:00', 'اسفند: عبور روزانه با تقویم شمسی سازگار است');
check(exp('', '08:00', '', '') === 'null', 'بدون تاریخ شروع → انقضا محاسبه نمی‌شود');

/* ---------- گروه ۷: بستن خودکار «تنبل» از طریق خود اپ (بدون تریگر) ---------- */
section('گروه ۷ — بستن خودکار تنبل هنگام دریافت لیست (حتی بدون تریگر)');
mock.resetWorld({});
G._doSetup(false);
seed('LZ-1', '1405-06-22', '08:00', '', '', 'فعال');
mock.setFakeNow(T(2026, 9, 13, 20, 0)); // پرمیت منقضی شده
let res = post({ action: 'permits', code: '' });
check(res.autoClosed === 1, 'اپ لیست گرفت → پرمیت منقضی خودبه‌خود بسته شد (بدون تریگر)');
check(res.rows.find(function(x) { return x.permitNo === 'LZ-1'; }).permitStatus === 'بسته', 'در همان پاسخ، وضعیت «بسته» نمایش داده شد');
res = post({ action: 'permits', code: '' });
check(res.autoClosed === 0, 'تراتل ۵ دقیقه‌ای: اجرای مجدد بلافاصله تکرار نمی‌شود');
const audit = G._auditSheet().getDataRange().getDisplayValues();
check(audit.some(function(rw) { return rw[1] === 'CLOSE_AUTO' && /LZ-1/.test(rw[4]); }), 'بستن خودکار در لاگ ضد دستکاری ثبت شد');
res = get({ action: 'permits' });
check(res.ok === true, 'GET لیست پرمیت‌ها (لایه ۳) نیز مسیر بستن خودکار را طی می‌کند');

/* ---------- گروه ۸: تریگر ۱۵ دقیقه‌ای ---------- */
section('گروه ۸ — تریگر زمان‌بندی‌شده');
mock._internal.triggers.length = 0;
G.installAutoCloseTrigger();
let trg = mock._internal.triggers.find(function(t) { return t.getHandlerFunction() === 'autoCloseExpiredPermits'; });
check(!!trg && trg._minutes === 15, 'تریگر ۱۵ دقیقه‌ای بستن خودکار نصب شد');
G.installAutoCloseTrigger();
check(mock._internal.triggers.filter(function(t) { return t.getHandlerFunction() === 'autoCloseExpiredPermits'; }).length === 1, 'نصب مجدد، تریگر تکراری نمی‌سازد');

/* ---------- گروه ۹: گشت و بازدیدها ---------- */
section('گروه ۹ — گشت HSE، عکس و تاریخچه بازدیدها');
mock.resetWorld({});
G._doSetup(false);
post({ action: 'login', code: ADMIN, firstName: 'فرشید', lastName: 'مدیر' });
post({ action: 'login', code: USER_B, firstName: 'رضا', lastName: 'کارگر' });
post({ action: 'setgrant', code: ADMIN, targetCode: USER_B, perm: 'patrol', grant: true });
post({ action: 'setgrant', code: ADMIN, targetCode: USER_B, perm: 'close', grant: true });
seed('P-1', '1405-06-22', '08:00', '', '', 'فعال');
mock.setFakeNow(T(2026, 9, 13, 10, 30));
const fakePhoto = 'data:image/jpeg;base64,' + Buffer.from('PHBob3RvPnRlc3Q8L3Bob3RvPg==', 'base64').toString('base64');
r = post({ action: 'savePatrolVisit', code: USER_B, inspectorCode: USER_B, permitNo: 'P-1', result: 'مطابق', scoreText: '6/7', notes: 'همه موارد رعایت شده', itemsText: 'چک‌لیست', inspectorName: 'رضا کارگر', photoBase64: fakePhoto, photosBase64: [fakePhoto], closePermitAlso: false });
check(r.ok === true && r.visitId, 'ثبت گشت با عکس موفق');
check(String(r.photoUrl || '').indexOf('drive.google.com') !== -1, 'عکس در گوگل‌درایو ذخیره و لینک برگشت');
r = post({ action: 'visits', code: USER_B, permitNo: 'P-1' });
check(r.ok === true && r.visits.length === 1, 'تاریخچه بازدید برای کاربر لاگین‌شده');
check(r.visits[0].inspectorCode === '', 'کد کاربر (رمز) هرگز در پاسخ بازدیدها نیست');
const visitRow = G._ss().getSheetByName('بازدیدها').getDataRange().getDisplayValues()[1];
check(String(visitRow[8] || '').trim() === '', 'ستون «کد بازدیدکننده» در شیت خالی است (امنیت)');
r = post({ action: 'visits' });
check(r.ok === false, 'تاریخچه بازدید بدون ورود → خطا');
r = post({ action: 'savePatrolVisit', code: USER_B, permitNo: 'P-1', result: 'مطابق', scoreText: '7/7', closePermitAlso: true });
check(r.ok === true && r.closedPermit === true, 'گشت با بستن هم‌زمان پرمیت');
check(statusOf('P-1') === 'بسته', 'پرمیت هم‌زمان با گشت بسته شد');

/* ---------- گروه ۱۰: قابلیت‌های جدید (چک‌این، گزارش، کاتالوگ، لوگو، خروجی) ---------- */
section('گروه ۱۰ — چک‌این، گزارش عملکرد، کاتالوگ، لوگو، خروجی');
mock.setFakeNow(T(2026, 9, 13, 11, 0));
r = post({ action: 'checkin', code: ADMIN, kind: 'open', ip: '5.1.2.3', ips: '192.168.1.5', ua: 'Android WebView' });
check(r.ok === true, 'چک‌این دستگاه ثبت شد');
r = post({ action: 'checkin', code: '00000000', kind: 'open' });
check(r.ok === false, 'چک‌این با کد نامعتبر رد می‌شود');

// سناریوی کامل گزارش عملکرد مدیر
post({ action: 'createPermit', code: ADMIN, firstName: 'فرشید', lastName: 'مدیر', permitNo: 'S-100', startDate: '1405/06/22', startTime: '09:00', place: 'تست', type: 'روتین', desc: 'تست گزارش', unit: 'تولید', contractor: 'سایر' });
post({ action: 'approvePermit', code: ADMIN, permitNo: 'S-100' });
post({ action: 'createPermit', code: ADMIN, firstName: 'فرشید', lastName: 'مدیر', permitNo: 'S-101', startDate: '1405/06/22', startTime: '09:00', place: 'تست', type: 'روتین', desc: 'تست', unit: 'تولید', contractor: 'سایر' });
post({ action: 'rejectPermit', code: ADMIN, permitNo: 'S-101', reason: 'تست' });
post({ action: 'closePermit', code: ADMIN, permitNo: 'S-100' });
post({ action: 'reactivatePermit', code: ADMIN, permitNo: 'S-100' });
post({ action: 'savePatrolVisit', code: ADMIN, permitNo: 'S-100', result: 'مطابق', scoreText: '7/7' });
post({ action: 'reinspectPermit', code: ADMIN, permitNo: 'S-101', newStatus: 'فعال' });
r = post({ action: 'mystats', code: ADMIN });
check(r.ok === true, 'mystats پاسخ می‌دهد');
check(r.todayDate === '1405/06/22', 'تاریخ امروز شمسی درست است (1405-06-22)');
check(String(r.month).indexOf('شهریور') === 0, 'نام ماه شمسی «شهریور» در گزارش ماه');
check(r.today.created.indexOf('S-100') !== -1 && r.today.created.indexOf('S-101') !== -1, 'ثبت‌های امروز در سطل «ثبت پرمیت»');
check(r.today.approved.indexOf('S-100') !== -1, 'تأیید امروز شمرده شد');
check(r.today.rejected.indexOf('S-101') !== -1, 'رد امروز شمرده شد');
check(r.today.closed.indexOf('S-100') !== -1, 'بستن امروز شمرده شد');
check(r.today.reactivated.indexOf('S-100') !== -1, 'تمدید امروز شمرده شد');
check(r.today.patrol.indexOf('S-100') !== -1, 'گشت امروز شمرده شد');
check(r.today.reinspect.indexOf('S-101') !== -1, 'بازرسی مجدد امروز شمرده شد');
check(r.today.opens >= 1, 'ورود/چک برنامه شمرده شد');
check(r.monthStats.created.indexOf('S-100') !== -1, 'آمار ماه نیز پر شده است');
r = post({ action: 'mystats' });
check(r.ok === false, 'mystats بدون کد → خطا');

// کاتالوگ
r = post({ action: 'gethsecatalog' });
check(r.ok === true && r.types.length === 8 && r.checklist.length === 7, 'کاتالوگ پیش‌فرض: ۸ نوع کار + ۷ آیتم چک‌لیست');
check(r.types.some(function(t) { return t.name === 'کار گرم (جوشکاری/برشکاری)' && t.warning; }), 'انواع کار با هشدار ایمنی همراه‌اند');
r = post({ action: 'savehsecatalog', code: USER_B, types: '[]' });
check(r.ok === false, 'ذخیره کاتالوگ توسط غیرمدیر رد می‌شود');
r = post({ action: 'savehsecatalog', code: ADMIN, types: JSON.stringify([{ name: 'تست نوع', warning: 'هشدار تست', keywords: 'k' }]), checklist: JSON.stringify(['آیتم ۱', 'آیتم ۲']), knowledge: JSON.stringify([{ topic: 'موضوع', text: 'متن' }]) });
check(r.ok === true && r.types.length === 1 && r.types[0].name === 'تست نوع', 'ذخیره کاتالوگ توسط مدیر');
r = post({ action: 'gethsecatalog' });
check(r.checklist.length === 2 && r.knowledge.length === 1, 'کاتالوگ ذخیره‌شده در دفعات بعد خوانده می‌شود (ماندگاری)');

// لوگو
r = post({ action: 'savelogo', code: USER_B, logoBase64: 'data:image/png;base64,iVBORw0KGgo=' });
check(r.ok === false, 'ذخیره لوگو توسط غیرمدیر رد می‌شود');
const png1x1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
r = post({ action: 'savelogo', code: ADMIN, logoBase64: 'data:image/png;base64,' + png1x1 });
check(r.ok === true && String(r.logoUrl || '').indexOf('drive.google.com') !== -1, 'لوگو در درایو ذخیره شد و لینک برگشت');
let pingR = get({ action: 'ping' });
check(pingR.logoUrl === r.logoUrl, 'لوگو در پاسخ پینگ ارسال می‌شود (نمایش در اپ)');
let permR = post({ action: 'permits', code: '' });
check(permR.logoUrl === r.logoUrl, 'لوگو در پاسخ لیست پرمیت‌ها نیز هست');

// خروجی
r = post({ action: 'exportpermits', code: ADMIN });
check(r.ok === true && String(r.fileUrl || '').indexOf('drive.google.com') !== -1, 'خروجی CSV پرمیت‌ها در درایو ساخته شد');
check(r.rows >= 2, 'تعداد ردیف‌های خروجی درست است');
r = post({ action: 'exportpermits' });
check(r.ok === false, 'خروجی بدون ورود → خطا');

/* ---------- گروه ۱۱: ضد ثبت تکراری (txnId) ---------- */
section('گروه ۱۱ — جلوگیری از ثبت تکراری با شناسه تراکنش');
const txn = 'TXN_TEST_123';
r = post({ action: 'closePermit', code: ADMIN, permitNo: 'S-100', txnId: txn });
check(r.ok === true, 'اجرای اول با txnId موفق');
r = post({ action: 'closePermit', code: ADMIN, permitNo: 'S-100', txnId: txn });
check(r.ok === true && r.duplicate === true, 'اجرای تکراری همان txnId بدون اثر جانبی پاسخ موفق می‌دهد');

/* ---------- گروه ۱۲: پنل مدیریت و لاگ ---------- */
section('گروه ۱۲ — پنل مدیریت (لاگ/تاریخچه/آمار/خروجی/پشتیبان)');
r = post({ action: 'adminlogs', code: USER_B });
check(r.ok === false, 'adminlogs برای غیرمدیر ممنوع');
r = post({ action: 'adminlogs', code: ADMIN, limit: '100' });
check(r.ok === true && r.logs.length > 0, 'لاگ‌ها برای مدیر قابل خواندن‌اند');
check(r.logs.every(function(l) { return l.hash && l.hash.length === 64; }), 'هش زنجیره‌ای لاگ‌ها سالم است');
// یک بستن خودکار در همین دنیا رقم بزن تا در آمار دیده شود
G._permitSheet().appendRow([99, '1405-06-22', '03:00', 'محل-آ', 'روتین', 'تست بستن خودکار', 'تولید', '', '', '', '', '', 'AC-1', 'مجری', 'فعال']);
mock.setFakeNow(T(2026, 9, 13, 21, 30));
G.autoCloseExpiredPermits();
r = post({ action: 'adminstats', code: ADMIN });
check(r.ok === true && r.stats.total > 0 && r.stats.loginCount >= 2, 'آمار لاگ (ورودها/عملیات) درست است');
check(r.stats.byAction['CLOSE_AUTO'] >= 1, 'بستن‌های خودکار در آمار شمرده شده‌اند');
r = post({ action: 'adminhistory', code: ADMIN });
check(r.ok === true && r.items.length > 0, 'تاریخچه تغییرات پر شده است');
r = post({ action: 'adminexport', code: ADMIN });
check(r.ok === true && r.rows > 0 && String(r.fileUrl || '').indexOf('drive.google.com') !== -1, 'خروجی اکسل لاگ ساخته شد');
r = post({ action: 'adminhistoryexport', code: ADMIN });
check(r.ok === true, 'خروجی اکسل تاریخچه تغییرات');
G.backupPermitSheet();
r = post({ action: 'adminbackups', code: ADMIN });
check(r.ok === true && r.backups.length === 1 && /^بایگانی_/.test(r.backups[0].name), 'پشتیبان‌گیری + فهرست بایگانی‌ها');
// زنجیره لاگ سالم است؟
const logVals = G._auditSheet().getDataRange().getDisplayValues();
let chainOk = true;
for (let i = 2; i < logVals.length; i++) {
  if (String(logVals[i][7]) !== String(logVals[i - 1][8])) { chainOk = false; break; }
}
check(chainOk, 'زنجیره هش لاگ (ضد دستکاری) سالم است');

/* ---------- گروه ۱۳: واحدها/پیمانکاران ---------- */
section('گروه ۱۳ — واحدها و پیمانکاران');
r = post({ action: 'getunitssettings' });
check(r.ok === true && r.units.indexOf('تولید') !== -1 && r.contractors.indexOf('زرین صنعت') !== -1, 'تنظیمات پیش‌فرض واحدها/پیمانکاران');
r = post({ action: 'saveunitssettings', code: USER_A, units: '["الف"]' });
check(r.ok === false, 'ویرایش واحدها بدون مجوز رد می‌شود');
post({ action: 'setgrant', code: ADMIN, targetCode: USER_B, perm: 'units', grant: true });
r = post({ action: 'saveunitssettings', code: USER_B, units: JSON.stringify(['واحد جدید']), contractors: JSON.stringify(['پیمانکار جدید']) });
check(r.ok === true, 'ویرایش واحدها با مجوز موفق');
r = post({ action: 'getunitssettings' });
check(r.units.indexOf('واحد جدید') !== -1 && r.contractors.indexOf('پیمانکار جدید') !== -1, 'تنظیمات جدید خوانده شد');

/* ---------- گروه ۱۴: تحلیل AI (مسیر آفلاین/بازگشتی) ---------- */
section('گروه ۱۴ — تحلیل هوش مصنوعی (بدون کلید = موتور بازگشتی)');
r = post({ action: 'analyzepermit', code: USER_A, workType: 'کار گرم', place: 'نوار نقاله', desc: 'جوشکاری' });
check(r.ok === false && r.needPerm === 'analyze', 'تحلیل بدون مجوز رد می‌شود');
r = post({ action: 'analyzepermit', code: ADMIN, workType: 'باربرداری', place: 'سالن آسیاب', desc: 'تعویض فراز پمپ' });
check(r.ok === true && r.analysis && r.analysis.risk === 'پرریسک', 'تحلیل با موتور بازگشتی: پرریسک + جزئیات');
check(String(r.analysis.source).length > 0, 'منبع تحلیل مشخص است');

/* ---------- گروه ۱۵: یکپارچگی سامانه چک‌لیست درکاو ---------- */
section('گروه ۱۵ — سامانه چک‌لیست درکاو (یک استقرار مشترک)');
mock.setFakeNow(T(2026, 9, 13, 14, 0));
const checklistRec = { id: 'abc123def456', templateTitle: 'چک‌لیست پمپ', category: 'مکانیک', assetCode: 'P-101', inspector: 'رضا', score: 80, maxScore: 100, percent: 80, completedAt: new Date(2026, 8, 13, 14, 0).toISOString() };
r = post({ action: 'save_record', record: checklistRec, html: '<html><body><h1>Test</h1></body></html>', folderName: 'سامانه چک‌لیست درکاو' });
check(r.ok === true && r.status === 'success', 'ذخیره چک‌لیست (از روی HTML) موفق');
check(r.sheetUrl.length > 0, 'ثبت در شیت تجمیعی انجام شد');
// ساختار پوشه شمسی
const rootCl = mock.DriveApp.getFoldersByName('سامانه چک‌لیست درکاو');
check(rootCl.hasNext(), 'پوشه اصلی چک‌لیست در درایو ساخته شد');
let yearF = rootCl.next().getFoldersByName('1405');
check(yearF.hasNext(), 'پوشه سال شمسی ۱۴۰۵');
let monthF = yearF.next().getFoldersByName('06-شهریور');
check(monthF.hasNext(), 'پوشه ماه شمسی ۰۶-شهریور');
let catF = monthF.next().getFoldersByName('مکانیک');
check(catF.hasNext(), 'پوشه دسته‌بندی «مکانیک»');
const clFile = catF.next().getFiles();
check(clFile.hasNext() && /\.pdf$/.test(clFile.next().getName()), 'فایل PDF با نام شمسی ذخیره شد');
// رکورد با PDF آماده باینری
const pdfB64 = Buffer.from('%PDF-1.4 fake pdf bytes').toString('base64');
r = post({ action: 'save_record', record: { id: 'zzz999', templateTitle: 'PDF آماده', category: 'برق' }, pdfBase64: pdfB64 });
check(r.ok === true, 'ذخیره PDF آماده (ساخته‌شده در مرورگر) موفق');
// پشتیبان
r = post({ action: 'backup', payload: { templates: [{ id: 1 }] } });
check(r.ok === true, 'اکشن پشتیبان پاسخ موفق داد');
let backupFound = false;
const backupFolderIt = mock.DriveApp.getFoldersByName('پشتیبان کامل');
if (backupFolderIt.hasNext()) {
  const bf = backupFolderIt.next().getFiles();
  while (bf.hasNext()) { if (/پشتیبان_1405-06-22/.test(bf.next().getName())) backupFound = true; }
}
check(backupFound, 'فایل پشتیبان با نام شمسی «پشتیبان_1405-06-22» در پوشه «پشتیبان کامل» ذخیره شد');
// پایگاه ابری
r = post({ action: 'cloud_pull', apiKey: '' });
check(r.ok === true && r.meta === null, 'دریافت از ابرِ خالی');
r = post({ action: 'cloud_push', apiKey: 'SECRET-1', meta: { v: 7 }, data: { templates: [1, 2], records: [1], users: [], assets: [], assignments: [] } });
check(r.ok === true, 'ارسال به ابر با کلید اولیه (کلید ثبت شد)');
r = post({ action: 'cloud_push', apiKey: 'WRONG', meta: {}, data: {} });
check(r.ok === false && /کلید امنیتی/.test(r.error), 'کلید نادرست ابر رد می‌شود');
r = post({ action: 'cloud_meta', apiKey: 'SECRET-1' });
check(r.ok === true && r.exists === true && r.counts.templates === 2 && r.counts.records === 1, 'فراداده ابر با شمارش‌ها');
r = post({ action: 'cloud_pull', apiKey: 'SECRET-1' });
check(r.ok === true && r.data.templates.length === 2, 'دریافت کامل از ابر');
// پینگ با نام پوشه (سازگاری اپ چک‌لیست)
r = post({ action: 'test', folderName: 'سامانه چک‌لیست درکاو' });
check(r.ok === true && r.folderName === 'سامانه چک‌لیست درکاو', 'پینگ چک‌لیست با اطلاعات پوشه');

/* ---------- گروه ۱۶: پایداری در شرایط خاص ---------- */
section('گروه ۱۶ — شرایط خاص و پایداری');
// شیت خالی
mock.resetWorld({});
G._doSetup(false);
r = post({ action: 'permits', code: '' });
check(r.ok === true && r.rows.length === 0 && r.count === 0, 'لیست خالی بدون خطا');
check(G.autoCloseExpiredPermits() === 0, 'بستن خودکار روی شیت خالی بدون خطا');
// بدنه خراب در POST
out = doPost({ postData: { contents: '{json خراب' } });
check(out.getContent().indexOf('"ok"') !== -1 || out.getContent().indexOf('error') !== -1, 'POST با بدنه خراب شکست سخت ندارد');
// بدون هیچ پارامتری
out = doPost({});
check(out.getContent().indexOf('rows') !== -1 || out.getContent().indexOf('ok') !== -1, 'POST خالی پاسخ مدیریت‌شده می‌دهد');
// دسترسی معلق
mock.resetWorld({});
G._doSetup(false);
post({ action: 'login', code: ADMIN, firstName: 'فرشید', lastName: 'مدیر' });
post({ action: 'login', code: USER_A, firstName: 'علی', lastName: 'بازرس' });
post({ action: 'setuseraccess', code: ADMIN, targetCode: USER_A, active: false });
r = post({ action: 'permits', code: USER_A });
check(r.ok === true && r.active === false, 'کاربر معلق‌شده در لیست، غیرفعال گزارش می‌شود');
r = post({ action: 'login', code: USER_A, firstName: 'علی', lastName: 'بازرس' });
check(r.ok === false && r.active === false, 'ورود کاربر معلق‌شده رد می‌شود');
r = post({ action: 'closePermit', code: USER_A, permitNo: 'X' });
check(r.ok === false, 'هیچ نوشتنی از کاربر معلق پذیرفته نمی‌شود');
post({ action: 'setuseraccess', code: ADMIN, targetCode: USER_A, active: true });
r = post({ action: 'login', code: USER_A, firstName: 'علی', lastName: 'بازرس' });
check(r.ok === true, 'وصل مجدد دسترسی کار می‌کند');

/* ---------------- جمع‌بندی ---------------- */
console.log('\n════════════════════════════════════════════');
console.log('📊 نتیجه تست‌ها:  ✅ موفق: ' + PASS + '   ❌ ناموفق: ' + FAIL);
console.log('════════════════════════════════════════════');
if (failures.length) {
  console.log('\nموارد ناموفق:');
  failures.forEach(function(f) { console.log('  - ' + f); });
  process.exit(1);
}
