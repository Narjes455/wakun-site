/**
 * ماتشا وكّن الجمل 🐪 — نظام الطلبات
 * الخادم: Google Apps Script مربوط بملف قوقل شيت.
 * التركيب: الصق هذا الملف في (Extensions > Apps Script)، شغّل الدالة setup مرة وحدة،
 * بعدها Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * تطوير: نرجس العمري
 */

const SH = {
  MENU: 'المنيو',
  ADD: 'الإضافات',
  MILK: 'الحليب',
  SET: 'الإعدادات',
  ORD: 'الطلبات'
};
const MILK_MODES = ['اختيار بسعر', 'اختيار مجاني', 'بدون'];
const CATS = ['ماتشا', 'قهوة باردة', 'قهوة حارة', 'احتفالي', 'حلويات', 'ماتشا باودر'];
const PAYS = ['كاش عند الاستلام', 'شبكة أو Apple Pay عند الاستلام', 'تحويل بنكي'];
const STATUSES = ['جديد', 'يتجهز', 'جاهز', 'استلم', 'ملغي'];
const MENU_HEAD = ['المعرف', 'الصنف', 'القسم', 'السعر', 'السعرات', 'متوفر', 'الوصف', 'شارة', 'الصورة', 'الحليب', 'الإضافات', 'الأحجام', 'صور إضافية'];
const POWDER_ROW = ['powder', 'بوكس ماتشا باودر', 'ماتشا باودر', 86, '', 'نعم', 'ماتشا احتفالية يابانية', '', 'powder.webp', 'بدون', '',
  '30 جرام: 86، 60 جرام: 105، 100 جرام: 170', 'g_bag.webp، g_whisk.webp، g_whisk2.webp، g_balls.webp'];

/* ترقية الشيت القديم: يضيف عمودي الأحجام والصور، ويدمج أصناف الباودر في صنف واحد بأحجام */
const GREEN_ROWS = [
  ['واتساب: رقم استقبال الطلبات', '966509116220', 'الرقم اللي توصله رسالة كل طلب جديد (بالصيغة الدولية بدون +). اتركيه فاضي لإيقاف الرسائل'],
  ['Green API: idInstance', '', 'من صفحة الإنستانس في console.green-api.com'],
  ['Green API: apiTokenInstance', '', 'من نفس الصفحة، لا تشاركيه مع أحد'],
  ['Green API: apiUrl', 'https://api.green-api.com', 'من نفس الصفحة (مثال: https://7107.api.greenapi.com)']
];
/* يضيف صفوف إعدادات Green API للشيت لو ما كانت موجودة */
function upgradeSettings_() {
  const sh = sheet_(SH.SET); if (!sh) return;
  const keys = rows_(SH.SET).map(r => String(r[0]).trim());
  GREEN_ROWS.forEach(r => { if (keys.indexOf(r[0]) < 0) sh.appendRow(r); });
}

/* رسالة واتساب لرقم البزنس مع كل طلب جديد (عن طريق Green API) */
function notify_(o) {
  try {
    const s = settings_();
    const to = String(s['واتساب: رقم استقبال الطلبات'] || '').replace(/[^\d]/g, '');
    const id = String(s['Green API: idInstance'] || '').trim(), tok = String(s['Green API: apiTokenInstance'] || '').trim();
    if (!to || !id || !tok) { Logger.log('ناقص في الإعدادات: ' + [!to && 'رقم استقبال الطلبات', !id && 'idInstance', !tok && 'apiTokenInstance'].filter(Boolean).join('، ')); return false; }
    const base = String(s['Green API: apiUrl'] || 'https://api.green-api.com').trim().replace(/\/+$/, '');
    const msg = ['🐪 *طلب جديد ' + o.no + '*', '',
      '👤 ' + o.name + ' · ' + o.phone,
      '📍 ' + (o.branch || '-') + ' · ⏰ ' + (o.time || '-'),
      '💳 ' + o.pay, '',
      o.lines.map(l => '• ' + l).join('\n'), '',
      '💰 *الإجمالي: ' + o.total + ' ريال*',
      o.note ? '📝 ' + o.note : '',
      o.photo ? '🖼️ صورة الطباعة: ' + o.photo : ''].join('\n').replace(/\n{3,}/g, '\n\n').trim();
    const r = UrlFetchApp.fetch(base + '/waInstance' + id + '/sendMessage/' + tok, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ chatId: to + '@c.us', message: msg })
    });
    if (r.getResponseCode() !== 200) Logger.log('رد Green API: ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 300));
    return r.getResponseCode() === 200;
  } catch (e) { return false; }
}
/* شغّليها مرة وحدة من المحرر عشان توافقين على الصلاحية وتجربين الرسالة */
function testNotify() {
  upgradeSettings_();
  const ok = notify_({ no: 'تجربة', name: 'تجربة', phone: '05xxxxxxxx', branch: 'خميس مشيط', time: 'اليوم 5 م', pay: 'كاش عند الاستلام',
    lines: ['1× ايس ماتشا لاتيه (حليب عادي) = 20'], total: 20, note: 'هذي رسالة تجربة من الموقع', photo: '' });
  Logger.log(ok ? 'وصلت الرسالة ✓' : 'ما وصلت: تأكدي من إعدادات Green API في تبويب الإعدادات');
}

function upgrade_() {
  upgradeSettings_();
  const sh = sheet_(SH.MENU); if (!sh) return;
  const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  if (head.indexOf('الأحجام') >= 0) return;
  sh.getRange(1, 12, 1, 2).setValues([['الأحجام', 'صور إضافية']]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#5E6447');
  sh.setColumnWidth(12, 260); sh.setColumnWidth(13, 260);
  sh.getRange(1, 12).setNote('للمنتج اللي له أحجام: الاسم: السعر، بينها فاصلة. مثال: 30 جرام: 86، 60 جرام: 105');
  sh.getRange(1, 13).setNote('صور تتقلب في نافذة المنتج: أسماء ملفات من images أو روابط قوقل درايف، بينها فاصلة');
  const ids = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getValues().map(r => String(r[0]).trim());
  const old = ['powder30', 'powder60', 'powder100'].map(id => ids.indexOf(id)).filter(i => i >= 0).sort((a, b) => b - a);
  if (old.length) {
    const at = old[old.length - 1] + 2;
    sh.getRange(at, 1, 1, POWDER_ROW.length).setValues([POWDER_ROW]);
    old.slice(0, -1).forEach(i => sh.deleteRow(i + 2));
  }
}
function sizes_(v) {
  return String(v || '').split(/[,،\n]+/).map(x => x.trim()).filter(String).map(x => {
    const m = x.match(/^(.*?)[:：=]\s*([\d.]+)\s*$/); return m ? { name: m[1].trim(), price: Number(m[2]) } : null;
  }).filter(Boolean);
}

/* =============== التركيب الأول =============== */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const make = (name, head, rows) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(head);
    if (sh.getLastRow() <= 1 && rows && rows.length) {
      const fixed = rows.map(r => { const x = r.slice(0, head.length); while (x.length < head.length) x.push(''); return x; });
      sh.getRange(2, 1, fixed.length, head.length).setValues(fixed);
    }
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    return sh;
  };
  make(SH.MENU, MENU_HEAD, [
    ['latte', 'ايس ماتشا لاتيه', 'ماتشا', 20, 187, 'نعم', '', 'الأكثر طلبًا', 'latte.webp', 'اختيار بسعر', ''],
    ['ccaramel', 'ايس ماتشا كريم كراميل', 'ماتشا', 20, 220, 'نعم', '', '', 'ccaramel.webp', 'اختيار بسعر', ''],
    ['wmocha', 'ايس ماتشا وايت موكا', 'ماتشا', 20, 335, 'نعم', '', '', 'wmocha.webp', 'اختيار بسعر', ''],
    ['vhoney', 'ايس ماتشا فانيلا عسل', 'ماتشا', 20, 180, 'نعم', '', '🍯', 'vhoney.webp', 'اختيار بسعر', ''],
    ['coconut', 'ايس ماتشا ماء جوز الهند', 'ماتشا', 22, 70, 'نعم', 'بدون حليب', 'خفيف', 'coconut.webp', 'بدون', ''],
    ['spanish', 'ايس سبانش السحابة', 'قهوة باردة', 19, 190, 'نعم', '', '☁️', 'spanish.webp', 'اختيار بسعر', ''],
    ['toffee', 'ايس توفي لاتيه', 'قهوة باردة', 19, 330, 'نعم', '', '', 'toffee.webp', 'اختيار بسعر', ''],
    ['k_wmocha', 'ايس وايت موكا', 'قهوة باردة', 19, 160, 'نعم', '', '', 'k_wmocha.webp', 'اختيار بسعر', ''],
    ['k_ccaramel', 'ايس كريم كراميل لاتيه', 'قهوة باردة', 20, 250, 'نعم', '', '', 'k_ccaramel.webp', 'اختيار بسعر', ''],
    ['coldfoam', 'امريكانو كولد فوم', 'قهوة باردة', 20, 110, 'نعم', '', '', 'cold_foam.webp', 'بدون', ''],
    ['freddo', 'الفريدو', 'قهوة باردة', 14, '', 'نعم', '', '', 'freddo.webp', 'بدون', ''],
    ['drip', 'ايس دريب', 'قهوة باردة', 19, '', 'نعم', '', '', 'cold_day.webp', 'بدون', ''],
    ['cold_day', 'قهوة اليوم بارد بالشمام', 'قهوة باردة', 10, '', 'نعم', '', '', 'cold_day.webp', 'بدون', ''],
    ['v60', 'V60', 'قهوة حارة', 17, '', 'نعم', '', '', 'hot_day.webp', 'بدون', ''],
    ['hot_day', 'قهوة اليوم بالشمام', 'قهوة حارة', 10, '', 'نعم', '', '', 'hot_day.webp', 'بدون', ''],
    ['cele_matcha', 'ماتشا احتفالية', 'احتفالي', 25, 178, 'نعم', 'بشنطة هدية وزينة المناسبة', '🎉', 'cele_matcha.webp', 'اختيار بسعر', 'print'],
    ['cele_spanish', 'ايس سبانش احتفالي', 'احتفالي', 20, 110, 'نعم', 'هدية جاهزة لأي مناسبة', '🎁', 'cele_spanish.webp', 'اختيار بسعر', 'print'],
    ['crunchy', 'كرانشي كيك', 'حلويات', 25, 490, 'نعم', '', '', 'crunchy.webp', 'بدون', ''],
    ['lava', 'لافا كيك', 'حلويات', 25, 360, 'نعم', '', '', 'lava.webp', 'بدون', ''],
    ['matilda', 'ماتيلدا بيكان كوب', 'حلويات', 20, 380, 'نعم', '', '', 'matilda.webp', 'بدون', ''],
    POWDER_ROW
  ]);
  make(SH.ADD, ['المعرف', 'الإضافة', 'السعر', 'متوفر'], [
    ['print', 'طباعة صورة خاصة', 15, 'نعم']
  ]);
  make(SH.MILK, ['المعرف', 'الحليب', 'السعر', 'متوفر'], [
    ['regular', 'حليب عادي', 0, 'نعم'],
    ['lactose', 'حليب خالي اللاكتوز', 3, 'نعم'],
    ['oat', 'حليب الشوفان', 4, 'نعم'],
    ['coconut', 'حليب جوز الهند', 4, 'نعم']
  ]);
  make(SH.SET, ['الإعداد', 'القيمة', 'ملاحظة'], [
    ['اسم المحل', 'ماتشا وكّن الجمل', ''],
    ['يفتح الساعة', 16, 'بنظام 24 ساعة (16 = 4 العصر)'],
    ['يسكر الساعة', 24, 'بنظام 24 ساعة (24 = 12 منتصف الليل)'],
    ['يوم الإجازة', '', 'اكتبي اسم اليوم، أو اتركيه فاضي'],
    ['استقبال الطلبات', 'نعم', 'اكتبي «لا» لإيقاف الطلبات مؤقتًا'],
    ['رسالة البانر', '', 'رسالة تظهر فوق المنيو، اتركيها فاضية لإخفائها'],
    ['لما يخلص صنف', 'يظهر باهت', 'يختفي / يظهر باهت'],
    ['الفروع', 'خميس مشيط\nأبها - أحد رفيدة', 'كل فرع بسطر. ولو تبون رابط اللوكيشن: اسم الفرع | الرابط'],
    ['المناسبات', 'تخرج، عيد ميلاد، نجاح، بدون مناسبة', 'خيارات الكوب الاحتفالي، بينها فاصلة'],
    ['طرق الدفع', PAYS.join('، '), 'اللي تبونه يظهر للعميل، بينها فاصلة'],
    ['بيانات التحويل', 'الآيبان: SA1780000107608016907588', 'اسم البنك واسم صاحب الحساب والآيبان، تظهر لما العميل يختار تحويل'],
    ['رقم التواصل', '966509116220', 'بالصيغة الدولية بدون +'],
    ['كلمة مرور اللوحة', 'wk' + Math.floor(1000 + Math.random() * 9000), 'غيّريها لكلمة تعرفونها أنتم بس']
  ]);
  make(SH.ORD, ['رقم الطلب', 'التاريخ', 'الفرع', 'وقت الاستلام', 'الاسم', 'الجوال', 'الطلب', 'المبلغ', 'الدفع', 'الحالة', 'العدد', 'ملاحظات', 'الصورة المطبوعة']);
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('الورقة1') || ss.getSheetByName('ورقة1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  photoFolder_();
  upgrade_();
  tidy();
}

/* =============== ترتيب الشيت وتنسيقه =============== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🐪 وكّن الجمل').addItem('ترتيب الشيت', 'tidy').addToUi();
}

function tidy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const INK = '#5E6447', BAND = '#F2F0E6';
  const list = (vals) => SpreadsheetApp.newDataValidation().requireValueInList(vals, true).setAllowInvalid(false).build();
  const style = (name, widths, notes) => {
    const sh = ss.getSheetByName(name); if (!sh) return null;
    const cols = widths.length, rows = Math.max(sh.getMaxRows(), 200);
    if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
    widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setRightToLeft(true); sh.setFrozenRows(1); sh.setHiddenGridlines(true);
    const all = sh.getRange(1, 1, rows, cols);
    all.setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('right').setFontSize(11);
    sh.setRowHeight(1, 38);
    sh.getBandings().forEach(b => b.remove());
    all.applyRowBanding(SpreadsheetApp.BandingTheme.GREEN, true, false).setHeaderRowColor(INK).setFirstRowColor('#ffffff').setSecondRowColor(BAND);
    sh.getRange(1, 1, 1, cols).setFontWeight('bold').setFontColor('#ffffff').setFontSize(12);
    if (notes) notes.forEach((n, i) => { if (n) sh.getRange(1, i + 1).setNote(n); });
    return { sh, rows };
  };
  const cf = (range, pairs) => pairs.map(([txt, bg, fg]) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(txt).setBackground(bg).setFontColor(fg).setBold(true).setRanges([range]).build());
  const yesNo = (sh, col, n) => {
    const r = sh.getRange(2, col, n);
    r.setDataValidation(list(['نعم', 'لا'])).setHorizontalAlignment('center');
    return cf(r, [['نعم', '#dcefe2', '#1f6b3a'], ['لا', '#F6E7DA', '#8A4B2A']]);
  };

  let t = style(SH.MENU, [100, 190, 110, 80, 80, 80, 220, 100, 150, 120, 120, 260, 260], [
    'حروف إنجليزية بدون مسافات وما تتكرر', 'اسم المنتج كما يظهر', 'القسم اللي يظهر فيه المنتج', 'بالريال', 'سعرة حرارية (اختياري)',
    'نعم = ظاهر، لا = خلص', 'وصف قصير (اختياري)', 'كلمة صغيرة فوق المنتج (اختياري)', 'اسم صورة من مجلد images، أو رابط قوقل درايف',
    'اختيار بسعر = العميل يختار الحليب بسعره من تبويب الحليب\nاختيار مجاني = يختار بدون زيادة\nبدون = ما فيه اختيار حليب',
    'معرفات الإضافات المسموحة من تبويب الإضافات، بينها فاصلة. فاضي = بدون إضافات',
    'للمنتج اللي له أحجام: الاسم: السعر، بينها فاصلة. مثال: 30 جرام: 86، 60 جرام: 105',
    'صور تتقلب في نافذة المنتج: أسماء ملفات من images أو روابط قوقل درايف، بينها فاصلة']);
  if (t) {
    const n = t.rows - 1;
    t.sh.getRange(2, 4, n).setNumberFormat('0 "ر.س"').setHorizontalAlignment('center');
    t.sh.getRange(2, 5, n).setHorizontalAlignment('center');
    t.sh.getRange(2, 3, n).setDataValidation(list(CATS));
    t.sh.getRange(2, 10, n).setDataValidation(list(MILK_MODES));
    t.sh.getRange(2, 1, n).setFontColor('#888888');
    t.sh.setConditionalFormatRules(yesNo(t.sh, 6, n));
  }
  [SH.ADD, SH.MILK].forEach(name => {
    const x = style(name, [100, 190, 80, 80]);
    if (x) { const n = x.rows - 1; x.sh.getRange(2, 3, n).setNumberFormat('0 "ر.س"').setHorizontalAlignment('center'); x.sh.getRange(2, 1, n).setFontColor('#888888'); x.sh.setConditionalFormatRules(yesNo(x.sh, 4, n)); }
  });
  t = style(SH.SET, [200, 300, 380]);
  if (t) {
    t.sh.getRange(2, 1, t.rows - 1).setFontWeight('bold');
    t.sh.getRange(2, 3, t.rows - 1).setFontColor('#888888').setFontStyle('italic').setFontSize(10);
    const keys = t.sh.getRange(2, 1, Math.max(1, t.sh.getLastRow() - 1)).getValues().map(r => String(r[0]).trim());
    const at = k => keys.indexOf(k) + 2;
    if (at('استقبال الطلبات') > 1) t.sh.getRange(at('استقبال الطلبات'), 2).setDataValidation(list(['نعم', 'لا']));
    if (at('لما يخلص صنف') > 1) t.sh.getRange(at('لما يخلص صنف'), 2).setDataValidation(list(['يختفي', 'يظهر باهت']));
    if (at('يوم الإجازة') > 1) t.sh.getRange(at('يوم الإجازة'), 2).setDataValidation(list(['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت']));
  }
  t = style(SH.ORD, [95, 140, 120, 110, 120, 120, 330, 90, 150, 100, 70, 200, 200]);
  if (t) {
    const n = t.rows - 1;
    t.sh.getRange(2, 1, n).setFontWeight('bold');
    t.sh.getRange(2, 8, n).setNumberFormat('0 "ر.س"');
    const st = t.sh.getRange(2, 10, n);
    st.setDataValidation(list(STATUSES)).setHorizontalAlignment('center');
    t.sh.setConditionalFormatRules(cf(st, [['جديد', '#E3E6CF', '#4A4E38'], ['يتجهز', '#fff3c4', '#7a5a00'], ['جاهز', '#dcefe2', '#1f6b3a'], ['استلم', '#eeeeee', '#555555'], ['ملغي', '#eeeeee', '#999999']]));
  }
  ss.getSheets().forEach(sh => { if (sh.getLastRow() > 1) sh.autoResizeRows(2, sh.getLastRow() - 1); });
}

/* =============== أدوات =============== */
function sheet_(n) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); }
function rows_(n) {
  const sh = sheet_(n); if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}
function settings_() { const o = {}; rows_(SH.SET).forEach(r => { if (r[0] !== '') o[String(r[0]).trim()] = r[1]; }); return o; }
function yes_(v) { return String(v).trim() === 'نعم'; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function phone_(p) {
  p = String(p || '').replace(/[^\d]/g, '');
  if (p.indexOf('966') === 0) p = '0' + p.slice(3);
  if (p.length === 9 && p[0] === '5') p = '0' + p;
  return p;
}
function now_() { return Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'); }
function today_() { return Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd'); }
function clean_(s, n) { return String(s == null ? '' : s).replace(/[<>]/g, '').slice(0, n || 200); }
function ids_(v) { return String(v || '').split(/[,،\s]+/).map(x => x.trim()).filter(String); }
function list_(v) { return String(v || '').split(/[,،\n]+/).map(x => x.trim()).filter(String); }
function milkMode_(v) {
  const t = String(v || '').trim();
  if (t.indexOf('مجاني') >= 0) return 'free';
  if (t.indexOf('اختيار') >= 0) return 'paid';
  return 'none';
}
function branches_(v) {
  return String(v || '').split(/\n+/).map(x => x.trim()).filter(String).map(line => {
    const p = line.split('|'); return { name: p[0].trim(), map: (p[1] || '').trim() };
  });
}

function publicData_() {
  const s = settings_();
  const menu = rows_(SH.MENU).filter(r => r[0] !== '' && r[1] !== '').map(r => ({
    id: String(r[0]).trim(), name: String(r[1]), cat: String(r[2] || '').trim() || 'ماتشا', price: Number(r[3]) || 0,
    cal: r[4] === '' ? '' : Number(r[4]) || '', on: yes_(r[5]), desc: String(r[6] || ''), badge: String(r[7] || ''),
    img: String(r[8] || '').trim(), milk: milkMode_(r[9]), adds: ids_(r[10]),
    sizes: sizes_(r[11]), gallery: String(r[12] || '').split(/[,،\n]+/).map(x => x.trim()).filter(String)
  }));
  menu.forEach(m => { if (m.sizes.length) m.price = Math.min.apply(null, m.sizes.map(z => z.price)); });
  const opt = n => rows_(n).filter(r => r[0] !== '' && r[1] !== '').map(r => ({ id: String(r[0]).trim(), name: String(r[1]), price: Number(r[2]) || 0, on: yes_(r[3]) }));
  const close = s['يسكر الساعة'] === '' || s['يسكر الساعة'] == null ? 24 : Number(s['يسكر الساعة']);
  const pays = list_(s['طرق الدفع']).filter(p => PAYS.indexOf(p) >= 0);
  return {
    menu: menu, adds: opt(SH.ADD), milks: opt(SH.MILK),
    settings: {
      name: s['اسم المحل'] || 'ماتشا وكّن الجمل',
      open: Number(s['يفتح الساعة']) || 16, close: close === 0 ? 24 : close,
      off: String(s['يوم الإجازة'] || ''), accepting: s['استقبال الطلبات'] === '' ? true : yes_(s['استقبال الطلبات']),
      banner: String(s['رسالة البانر'] || ''),
      hideOut: String(s['لما يخلص صنف'] || 'يظهر باهت').trim() === 'يختفي',
      branches: branches_(s['الفروع']).map(b => b.name),
      occasions: list_(s['المناسبات']),
      pays: pays.length ? pays : PAYS.slice(0, 1),
      contact: String(s['رقم التواصل'] || '').replace(/[^\d]/g, ''),
      transfer: String(s['بيانات التحويل'] || '')
    }
  };
}

/* سعر المنتج يتحسب هنا في الخادم من الشيت، مو من المتصفح */
function priceItem_(m, it, data) {
  let unit = m.price; const parts = [];
  if (m.sizes && m.sizes.length) {
    const z = m.sizes.find(x => x.name === String(it.size)) || m.sizes[0];
    unit = z.price; parts.push(z.name);
  }
  if (m.milk === 'paid' || m.milk === 'free') {
    const mk = data.milks.find(x => x.id === String(it.milk) && x.on) || data.milks.find(x => x.on);
    if (mk) { parts.push(mk.name); if (m.milk === 'paid') unit += mk.price; }
  }
  (Array.isArray(it.adds) ? it.adds : []).slice(0, 10).forEach(id => {
    if (m.adds.indexOf(String(id)) < 0) return;
    const a = data.adds.find(x => x.id === String(id) && x.on); if (!a) return;
    unit += a.price; parts.push(a.name);
  });
  if (m.cat === 'احتفالي') {
    if (it.occ) parts.push('المناسبة: ' + clean_(it.occ, 30));
    if (it.card) parts.push('الكرت: ' + clean_(it.card, 30));
  }
  return { unit: unit, text: parts.join(' · ') };
}

/* =============== GET =============== */
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.a === 'data') { upgrade_(); return json_(publicData_()); }
    if (p.a === 'track') {
      const r = rows_(SH.ORD).find(x => String(x[0]) === String(p.no));
      return json_(r ? { ok: true, status: String(r[9]) } : { ok: false });
    }
    return json_({ ok: true, name: 'Wakun API' });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}

/* =============== POST =============== */
function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents || '{}'); } catch (x) { return json_({ ok: false, error: 'bad json' }); }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (b.a === 'order') return json_(placeOrder_(b));
    if (String(b.pass || '') !== String(settings_()['كلمة مرور اللوحة'])) return json_({ ok: false, error: 'auth' });
    if (b.a === 'admin') return json_(adminData_());
    if (b.a === 'status') return json_(setStatus_(b.no, b.status));
    if (b.a === 'stock') return json_(setOn_(b.kind === 'add' ? SH.ADD : b.kind === 'milk' ? SH.MILK : SH.MENU, b.id, b.on));
    if (b.a === 'accepting') return json_(setSetting_('استقبال الطلبات', b.on ? 'نعم' : 'لا'));
    return json_({ ok: false, error: 'unknown' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally { try { lock.releaseLock(); } catch (x) {} }
}

function placeOrder_(b) {
  const data = publicData_(), s = data.settings;
  if (!s.accepting) return { ok: false, msg: 'متوقفين عن استقبال الطلبات حاليًا' };
  const ph = phone_(b.phone);
  if (!/^05\d{8}$/.test(ph)) return { ok: false, msg: 'اكتب رقم جوال صحيح' };
  const branch = s.branches.indexOf(String(b.branch)) >= 0 ? String(b.branch) : (s.branches[0] || '');
  const pay = s.pays.indexOf(String(b.pay)) >= 0 ? String(b.pay) : s.pays[0];
  let total = 0, count = 0, hasPrint = false; const lines = [];
  (b.items || []).slice(0, 30).forEach(it => {
    const m = data.menu.find(x => x.id === String(it.id));
    const q = Math.max(1, Math.min(20, Number(it.q) || 1));
    if (!m || !m.on) return;
    const pr = priceItem_(m, it, data);
    total += pr.unit * q; count += q;
    if ((it.adds || []).indexOf('print') >= 0) hasPrint = true;
    lines.push(q + '× ' + m.name + (pr.text ? ' (' + pr.text + ')' : '') + ' = ' + (pr.unit * q));
  });
  if (!lines.length) return { ok: false, msg: 'المنتجات المطلوبة خلصت' };

  let photo = '';
  if (hasPrint && b.photo) photo = savePhoto_(b.photo, ph);

  const sh = sheet_(SH.ORD);
  const no = 'WK-' + (100 + sh.getLastRow());
  sh.appendRow([no, now_(), branch, clean_(b.time, 30), clean_(b.name, 40) || 'ضيف', "'" + ph, lines.join('\n'), total,
    pay, 'جديد', count, clean_(b.note, 300), photo]);
  notify_({ no: no, name: clean_(b.name, 40) || 'ضيف', phone: ph, branch: branch, time: clean_(b.time, 30), pay: pay,
    lines: lines, total: total, note: clean_(b.note, 300), photo: photo });
  const map = (branches_(settings_()['الفروع']).find(x => x.name === branch) || {}).map || '';
  return { ok: true, no: no, total: total, branch: branch, map: map, pay: pay, transfer: pay === 'تحويل بنكي' ? s.transfer : '' };
}

function photoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PHOTO_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('وكّن الجمل — صور الأكواب الاحتفالية');
  props.setProperty('PHOTO_FOLDER', f.getId());
  return f;
}
function savePhoto_(dataUrl, ph) {
  try {
    const m = String(dataUrl).match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
    if (!m || m[2].length > 2800000) return '';
    const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'print-' + ph + '-' + Date.now() + '.jpg');
    return photoFolder_().createFile(blob).getUrl();
  } catch (e) { return ''; }
}

function adminData_() {
  const t = today_();
  const ord = rows_(SH.ORD).map(r => ({ no: String(r[0]), at: String(r[1]), branch: String(r[2]), time: String(r[3]), name: String(r[4]),
    phone: String(r[5]), items: String(r[6]), total: Number(r[7]) || 0, pay: String(r[8]), status: String(r[9]), count: Number(r[10]) || 0,
    note: String(r[11] || ''), photo: String(r[12] || '') }));
  const todayOrd = ord.filter(o => o.at.indexOf(t) === 0 && o.status !== 'ملغي');
  const active = ord.filter(o => o.status !== 'استلم' && o.status !== 'ملغي').slice(-60);
  const pub = publicData_();
  return { ok: true, today: { orders: todayOrd.length, count: todayOrd.reduce((a, o) => a + o.count, 0), total: todayOrd.reduce((a, o) => a + o.total, 0) },
    orders: active.reverse(), menu: pub.menu, adds: pub.adds, milks: pub.milks, settings: pub.settings };
}
function setStatus_(no, status) {
  if (STATUSES.indexOf(status) < 0) return { ok: false };
  const all = rows_(SH.ORD); const i = all.findIndex(r => String(r[0]) === String(no));
  if (i < 0) return { ok: false };
  sheet_(SH.ORD).getRange(i + 2, 10).setValue(status);
  return { ok: true };
}
function setOn_(name, id, on) {
  const col = name === SH.MENU ? 6 : 4;
  const all = rows_(name); const i = all.findIndex(r => String(r[0]).trim() === String(id));
  if (i < 0) return { ok: false };
  sheet_(name).getRange(i + 2, col).setValue(on ? 'نعم' : 'لا');
  return { ok: true };
}
function setSetting_(key, val) {
  const all = rows_(SH.SET); const i = all.findIndex(r => String(r[0]).trim() === key);
  if (i < 0) return { ok: false };
  sheet_(SH.SET).getRange(i + 2, 2).setValue(val);
  return { ok: true };
}
