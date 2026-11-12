/*!
 * OSOS XLSB Server Converter
 * --------------------------
 * سيرفر بسيط بيستقبل ملف إكسل/CSV/TXT ويرجّعه محوّل لصيغة XLSB.
 *
 * مبادئ الخصوصية المتبعة هنا (مهم تفهمها قبل ما تنشره):
 * - الملف بيتعالج بالكامل في الذاكرة (RAM) فقط — مفيش أي كتابة على القرص.
 * - مفيش أي console.log لمحتوى الملف أو حتى اسمه.
 * - بعد إرسال الرد، النسخة اللي في الذاكرة بتتمسح تلقائيًا (garbage collection)
 *   لأننا مش محتفظين بأي reference ليها في أي متغيّر عام.
 * - السيرفر بياخد الملف "مؤقتًا" بس وقت الطلب نفسه، ومفيش تخزين دائم أو قاعدة بيانات.
 *
 * لازم تعرف: ده معناه إن بيانات الملف بتعدي فعليًا على السيرفر ده (اللي هو
 * حسابك انت على المنصة اللي هتختارها)، حتى لو مش بيتحفظ. لو عايز خصوصية
 * كاملة بدون أي استثناء، سيب التحويل يحصل محليًا فقط (زي ما كان قبل كده)
 * ومتحطش رابط سيرفر في الواجهة.
 */

const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const cors = require("cors");
const path = require("path");
const { extractText, getDocumentProxy, renderPageAsImage, createIsomorphicCanvasFactory } = require("unpdf");
const { createCanvas } = require("@napi-rs/canvas");
const { createWorker } = require("tesseract.js");

const app = express();

// تحديد مين مسموح له يستخدم السيرفر (حط دومين موقعك هنا بدل * لو حابب تقفلها أكتر)
app.use(cors({ origin: "*" }));
app.use("/api/vin/read", express.json({ limit: "8mb" })); // صورة الشاصي (base64) أكبر من 1 ميجا أحيانًا
app.use(express.json({ limit: "1mb" })); // لازمة عشان endpoint استخراج أرقام اللوحات (بياخد JSON مش ملف)

// نخزن الملف في الذاكرة مباشرة، مش على القرص أبدًا.
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 120 * 1024 * 1024 } // سقف أمان 120 ميجا للملف الواحد
});

app.get("/", (req, res) => {
  res.send("OSOS XLSB converter server is running.");
});

/* =================================================================
 * استخراج أرقام اللوحات من شهادات الـ PDF (بدل ما يحصل ده في المتصفح)
 * -----------------------------------------------------------------
 * الفكرة: الموبايل بيبعتلنا هنا بس IDs بتاعة الملفات، إحنا بنجيب كل PDF
 * من نفس رابط التنزيل الموجود أصلًا في Worker الشهادات (osos-certificates)
 * وبعدين نستخرج النص ونلاقي فيه رقم اللوحة، ونرجّع النتيجة الجاهزة.
 *
 * مهم: الـ Worker مقفول ومحتاج إثبات دخول Firebase، فالموقع بيبعت التوكن في
 * هيدر Authorization وإحنا بنمرره زي ما هو للـ Worker عند تنزيل كل PDF.
 * ================================================================= */

const CERT_DOWNLOAD_BASE = "https://osos-certificates.ososapp.workers.dev";
const CERT_RESOLVE_MAX_IDS = 20; // سقف أمان لعدد الملفات في الطلب الواحد (كان 30 — قللناه)
const CERT_RESOLVE_CONCURRENCY = 3; // كام ملف بيتفتح بالتوازي جوه نفس الطلب (كان 6 — قللناه عشان Render Free)

// فاصل مسموح بين حروف اللوحة أو بينها وبين الأرقام: مسافة أو شرطة (كان بس مسافة قبل كده)
const SEP = "[\\s\\-]*";

// نفس ترتيب الأنماط اللي كانت شغالة في المتصفح، منقولة هنا بالظبط عشان النتيجة متطابقة
// (دي طريقة احتياطية، بتتفتش لو الطريقة الأدق اللي تحت (بالعنوان) مالقتش حاجة)
const CERT_PLATE_PATTERNS = [
  // العربي: {2,4} مش {3} بالظبط — مرونة لضوضاء الـOCR (حرف زيادة أو ناقص)
  { re: new RegExp("(?<![\\u0621-\\u064A])((?:[\\u0621-\\u064A]" + SEP + "){2,4})(?![\\u0621-\\u064A])" + SEP + "(?:^|[^\\d])(\\d{4})(?!\\d)"), lettersFirst: true },
  { re: new RegExp("(?:^|[^\\d])(\\d{4})(?!\\d)" + SEP + "((?:[\\u0621-\\u064A]" + SEP + "){2,4})(?![\\u0621-\\u064A])"), lettersFirst: false },
  // الإنجليزي: فاضل {3} بالظبط لأنه من نص حقيقي مش OCR، دقيق أصلًا
  { re: new RegExp("(?:^|[^\\d])(\\d{4})(?!\\d)" + SEP + "((?:[A-Za-z]" + SEP + "){3})(?![A-Za-z])"), lettersFirst: false },
  { re: new RegExp("(?:^|[^A-Za-z])((?:[A-Za-z]" + SEP + "){3})(?![A-Za-z])" + SEP + "(?:^|[^\\d])(\\d{4})(?!\\d)"), lettersFirst: true }
];

// الطريقة الأدق: نلاقي عنوان الحقل نفسه ("رقم اللوحة باللغة العربية:" أو الإنجليزية)
// ونفتش بس في الجزء اللي بعده مباشرة — عشان منتأثرش بأي نص تاني في الشهادة
const AR_PLATE_LABEL = /رقم\s*اللوحة\s*باللغة\s*العربية\s*:?/;
const EN_PLATE_LABEL = /رقم\s*اللوحة\s*باللغة\s*الإنجليزية\s*:?/;
const LABEL_WINDOW = 60; // عدد الحروف اللي بنفتش فيها بعد العنوان مباشرة

function grabPlateAfterLabel(text, labelRe, letterClass, stopRe) {
  const m = text.match(labelRe);
  if (!m) return null;
  const start = m.index + m[0].length;
  let end = start + LABEL_WINDOW;
  // مهم: لو الحقل ده فاضي (مفيش رقم لوحة عربي مثلاً)، النص هيروح على طول لعنوان الحقل
  // اللي بعده (زي "رقم اللوحة باللغة الإنجليزية:") — لازم نوقف قبله عشان منلخبطش
  // ونفتكر إن كلمة "رقم" بتاعت العنوان التاني هي حروف اللوحة بالغلط
  if (stopRe) {
    const stopM = text.slice(start, end).match(stopRe);
    if (stopM) end = start + stopM.index;
  }
  const windowText = text.slice(start, end);
  // {2,4} مش {3} بالظبط — عشان الـOCR أحيانًا بيقرا حرف زيادة أو ناقص غلط، فمرونة بسيطة
  // في عدد الحروف بتخلينا نمسك اللوحة برضه بدل ما نرفضها تمامًا لمجرد فرق حرف واحد
  const lettersRe = new RegExp("(?:[" + letterClass + "]" + SEP + "){2,4}(?![" + letterClass + "])");
  const lettersMatch = windowText.match(lettersRe);
  const digitsMatch = windowText.match(/(\d{4})(?!\d)/);
  if (!lettersMatch || !digitsMatch) return null;
  const letters = lettersMatch[0].replace(/[\s\-]+/g, "");
  return letters + " " + digitsMatch[1];
}

// الحقل الإنجليزي: ساعات بيبقى مكتوب فيه "None" (مفيش لوحة إنجليزي) — ده مش لوحة، ولازم نتجاهله
// ونرجع للعربي. كمان مبنمسكش حروف وأرقام من الحقل اللي بعده (زي رقم الهيكل VIN) لأن الحروف والأرقام
// لازم يكونوا ملزوقين في بعض (3 حروف إنجليزي + 4 أرقام، بأي ترتيب).
const EN_FIELD_STOP = /رقم\s*الهيكل|لون\s*المركبة|نوع\s*المركبة|المصنع/;
function grabEnglishPlateAfterLabel(text) {
  const m = text.match(EN_PLATE_LABEL);
  if (!m) return null;
  const start = m.index + m[0].length;
  let win = text.slice(start, start + LABEL_WINDOW);
  const stopM = win.match(EN_FIELD_STOP);
  if (stopM) win = win.slice(0, stopM.index);
  win = win.replace(/\b(?:none|null|nil|n\/a)\b/gi, " ");
  const L3 = "((?:[A-Za-z][\\s\\-]*){3})(?![A-Za-z])";
  let r = win.match(new RegExp("(?<!\\d)(\\d{4})(?!\\d)[\\s\\-]*(?<![A-Za-z])" + L3));
  if (r) return r[2].replace(/[\s\-]+/g, "") + " " + r[1];
  r = win.match(new RegExp("(?<![A-Za-z])" + L3 + "[\\s\\-]*(?<!\\d)(\\d{4})(?!\\d)"));
  if (r) return r[1].replace(/[\s\-]+/g, "") + " " + r[2];
  return null;
}

// شهادات "توثيق" (ARN / مستخرج السند التنفيذي): العنوان "رقم اللوحة:" والقيمة حروف منفصلة
// بمسافات + من 1 لـ 4 أرقام، زي "أ ط ط 1637". مكتبة قراءة الـPDF أحيانًا بتطلّع كلمات العنوان
// بالمقلوب ("أ ط ط 1637 :اللوحة رقم")، فبنجرب الشكلين. الحروف لازم تكون منفصلة (حرف حرف)
// عشان منمسكش كلمة عادية من الشهادة بالغلط.
const TW_LETTERS = "(?<![\\u0621-\\u064A])([\\u0621-\\u064A](?:[ \\t]+[\\u0621-\\u064A]){1,3})(?![\\u0621-\\u064A])";
const TW_DIGITS = "(\\d{1,4})";
const TAWTHEEQ_PATTERNS = [
  { re: new RegExp(TW_LETTERS + "[ \\t]+" + TW_DIGITS + "\\s*:\\s*اللوحة\\s+رقم"), l: 1, d: 2 },
  { re: new RegExp("(?<!\\d)" + TW_DIGITS + "[ \\t]+" + TW_LETTERS + "\\s*:\\s*اللوحة\\s+رقم"), l: 2, d: 1 },
  { re: new RegExp("رقم\\s+اللوحة\\s*:[ \\t]*" + TW_LETTERS + "[ \\t]+" + TW_DIGITS + "(?!\\d)"), l: 1, d: 2 },
  { re: new RegExp("رقم\\s+اللوحة\\s*:[ \\t]*" + TW_DIGITS + "[ \\t]+" + TW_LETTERS), l: 2, d: 1 },
];
function extractTawtheeqPlate(text) {
  for (const p of TAWTHEEQ_PATTERNS) {
    const m = text.match(p.re);
    if (m) return m[p.l].replace(/\s+/g, "") + " " + m[p.d];
  }
  return null;
}

function extractPlateFromText(text) {
  // 1) جرب العنوان العربي الصريح (ولازم نوقف قبل عنوان الإنجليزي لو الحقل العربي فاضي)
  const ar = grabPlateAfterLabel(text, AR_PLATE_LABEL, "\\u0621-\\u064A", EN_PLATE_LABEL);
  if (ar) return ar;
  // 2) لو مفيش، جرب العنوان الإنجليزي الصريح (بيتجاهل "None" وأي قيمة فاضية، ولازم الحروف والأرقام يكونوا ملزوقين في بعض)
  const en = grabEnglishPlateAfterLabel(text);
  if (en) return en;
  // 2.5) شهادات توثيق (ARN): "رقم اللوحة:" من غير "باللغة العربية"
  const tw = extractTawtheeqPlate(text);
  if (tw) return tw;
  // 3) أخيرًا، الطريقة القديمة (فحص النص كله بدون الاعتماد على العنوان)
  for (let i = 0; i < CERT_PLATE_PATTERNS.length; i++) {
    const { re, lettersFirst } = CERT_PLATE_PATTERNS[i];
    const m = text.match(re);
    if (!m) continue;
    const letters = (lettersFirst ? m[1] : m[2]).replace(/[\s\-]+/g, "");
    const digits = lettersFirst ? m[2] : m[1];
    if (digits && letters) return letters + " " + digits;
  }
  return null;
}

/* =================================================================
 * شهادات سجل (REPO/CRN): قراءة اللوحة من مكان الحقل في الصفحة مباشرة
 * -----------------------------------------------------------------
 * الحروف العربي في شهادات سجل متخزنة كأرقام "شكل الحرف" (glyph) في خط
 * Microsoft Sans Serif من غير جدول تحويل لحروف، فالنص بيطلع رموز غريبة.
 * أرقام الأشكال دي ثابتة في الخط (اتأكدنا منها من cmap الخط المدمج في الشهادات
 * نفسها)، فبنحولها لحروف عربي بجدول ثابت — من غير OCR ومن غير تخمين.
 * وبنمسك قيمة الحقل من السطر بتاع عنوانه ("رقم اللوحة باللغة العربية/الإنجليزية")
 * فالمسافات بين الحروف أو بين الأرقام (زي "3 4 2 0 U N R") مش بتفرق.
 * ================================================================= */
const SIJIL_GID_AR = {"787":"ا","788":"أ","789":"آ","790":"إ","791":"ب","792":"ت","794":"ج","796":"ح","797":"ح","798":"خ","799":"خ","800":"س","801":"ش","802":"ص","813":"ض","814":"ع","815":"ع","816":"ع","818":"غ","819":"غ","820":"ف","821":"ف","823":"ق","824":"ق","825":"ك","826":"ل","827":"م","828":"ن","829":"ه","830":"ه","831":"ه","832":"ي","833":"ي","834":"ة","835":"ى","837":"ئ","838":"لا","839":"لا","843":"لأ","844":"لأ","845":"لآ","848":"لآ","851":"ء","852":"آ","853":"أ","854":"ؤ","855":"إ","857":"ا","858":"ب","859":"ة","860":"ت","863":"ح","865":"د","866":"ذ","867":"ر","868":"ز","869":"س","871":"ص","872":"ض","873":"ط","874":"ظ","875":"ع","877":"ـ","878":"ف","879":"ق","880":"ك","881":"ل","882":"م","883":"ن","884":"ه","885":"و","886":"ى","887":"ي","896":"لإ","897":"لإ","903":"ى","1022":"ل","2694":"الله","2695":"رال","3035":"ر","3036":"ز","3050":"ت","3051":"ث","3052":"ي","3162":"محمد","3163":"صلىاللهعليهوسلم","3165":"جلجلاله"};
// الحروف اللي بتيجي في اللوح السعودية بس (شكل الحرف المنفصل في الخط)
const SIJIL_PLATE_GID = { 857: "ا", 858: "ب", 863: "ح", 865: "د", 867: "ر", 869: "س", 871: "ص", 873: "ط", 875: "ع", 879: "ق", 880: "ك", 881: "ل", 882: "م", 883: "ن", 884: "ه", 885: "و", 886: "ى" };
const SIJIL_PLATE_LETTERS = new Set("ابحدرسصطعقكلمنهوىي".split(""));
// تحويل حروف اللوحة الإنجليزي للعربي (جدول اللوح السعودية الرسمي)
const EN_TO_AR_PLATE = { A: "ا", B: "ب", J: "ح", D: "د", R: "ر", S: "س", X: "ص", T: "ط", E: "ع", G: "ق", K: "ك", L: "ل", Z: "م", N: "ن", H: "ه", U: "و", V: "ي" };
function englishPlateToArabic(plate) {
  const m = String(plate || "").trim().match(/^([A-Za-z]{3})\s*(\d{1,4})$/);
  if (!m) return plate;
  const letters = m[1].toUpperCase().split("");
  if (!letters.every((c) => EN_TO_AR_PLATE[c])) return plate; // حرف مش في جدول اللوح السعودية: نسيبها زي ما هي
  // الحروف الإنجليزي بتتقري من الشمال لليمين والعربي من اليمين للشمال، وكل حرف عربي فوق اللي يقابله
  // في اللوحة → الترتيب العربي = عكس الإنجليزي (اتأكدنا: "V N R" في نفس الشهادة عربيها "ر ن ي")
  return letters.reverse().map((c) => EN_TO_AR_PLATE[c]).join("") + " " + m[2];
}
// نص العنصر (بالترتيب المرئي من الشمال لليمين) → نص عربي مقروء (من اليمين للشمال)
function sijilDecodeLabel(str) {
  const out = [];
  for (const ch of String(str || "")) {
    const cp = ch.codePointAt(0);
    if (cp < 128) { if (ch !== " " && ch !== ":") out.push(ch); continue; }
    out.push(SIJIL_GID_AR[cp] || "?");
  }
  return out.reverse().join("");
}
async function plateFromSijilItems(pdf, want) {
  const page = await pdf.getPage(1);
  const content = await page.getTextContent();
  const items = content.items.filter((it) => it.str && it.str.trim());
  let arLabel = null, enLabel = null;
  for (const it of items) {
    const t = sijilDecodeLabel(it.str);
    if (!/اللوحةباللغة/.test(t)) continue;
    if (/العربية/.test(t)) arLabel = it;
    else if (/نجليزية/.test(t)) enLabel = it;
  }
  // قيمة الحقل = العناصر اللي على نفس السطر وعلى شمال العنوان (لحد العنوان اللي قبله في نفس السطر)
  function valueOf(label) {
    if (!label) return null;
    const ly = label.transform[5], lx = label.transform[4];
    const row = items.filter((it) => it !== label && Math.abs(it.transform[5] - ly) < 3 && it.transform[4] < lx)
      .sort((a, b) => a.transform[4] - b.transform[4]);
    // نوقف عند أي عنوان تاني على نفس السطر (بيخلص بـ ":")
    let start = 0;
    row.forEach((it, i) => { if (/:\s*$/.test(it.str) || /^\s*:/.test(it.str)) start = i + 1; });
    return row.slice(start).map((it) => it.str).join(" ");
  }
  // الإنجليزي الأول (بأي مسافات) → نحوله عربي
  const enVal = want === "ar" ? null : valueOf(enLabel);
  if (enVal) {
    const v = enVal.replace(/[\s\-]+/g, "").toUpperCase();
    let m = v.match(/^([A-Z]{3})(\d{1,4})$/);
    if (m) return englishPlateToArabic(m[1] + " " + m[2]);
    m = v.match(/^(\d{1,4})([A-Z]{3})$/);
    if (m) return englishPlateToArabic(m[2] + " " + m[1]);
  }
  if (want === "en") return null;
  // بعده العربي: أرقام + أشكال حروف (بالترتيب المرئي، فبنقلب الحروف عشان تبقى من اليمين)
  const arVal = valueOf(arLabel);
  if (arVal) {
    const digits = (arVal.match(/\d/g) || []).join("");
    const letters = [];
    let bad = false;
    for (const ch of arVal) {
      const cp = ch.codePointAt(0);
      if (cp < 128) { if (!/[\d\s\-]/.test(ch)) bad = true; continue; }
      // أي شكل حرف في الخط (من جدول الخط) طالما هو حرف لوح سعودي -- زي الياء ي (887) اللي ماكانتش في الجدول الصغير
      let lt = SIJIL_PLATE_GID[cp] || SIJIL_GID_AR[cp];
      if (lt === "أ" || lt === "إ" || lt === "آ") lt = "ا";
      if (lt && SIJIL_PLATE_LETTERS.has(lt)) letters.push(lt); else bad = true;
    }
    if (!bad && letters.length === 3 && digits.length >= 1 && digits.length <= 4) return letters.reverse().join("") + " " + digits;
  }
  return null;
}

/* =================================================================
 * OCR للعربي (حل أخير، بس للشهادات اللي فشل معاها استخراج النص العادي)
 * -----------------------------------------------------------------
 * المشكلة: الفونت العربي في الشهادات دي بيتحول لرموز عشوائية لما نحاول
 * نقرأه كنص من جوه الـ PDF (مش مشكلة ترتيب أو مسافات، المشكلة في الفونت
 * نفسه). الحل الوحيد الشغال: نحوّل الصفحة لصورة، ونشغّل عليها "قراءة
 * ضوئية" (OCR) بالعربي — تمامًا زي ما العين البشرية بتقرا صورة.
 *
 * ده أبطأ بكتير من قراءة النص، فبنستخدمه كملاذ أخير بس، لما كل الطرق
 * التانية (العنوان الصريح + الأنماط القديمة) تفشل تمامًا.
 * ================================================================= */

const OCR_LANG_PATH = path.join(__dirname, "node_modules", "@tesseract.js-data", "ara", "4.0.0_best_int");
const OCR_RENDER_SCALE = 8; // دقة عالية — مضمونة سريعة دلوقتي لأننا بنقص المنطقة المطلوبة بس غالبًا
const OCR_TIMEOUT_MS = 30000; // 30 ثانية — كافية جدًا (بالتجربة الفعلية بياخد ثانية-اتنين بس)

let ocrWorkerPromise = null;
function getOcrWorker() {
  if (!ocrWorkerPromise) {
    ocrWorkerPromise = createWorker("ara", 1, {
      langPath: OCR_LANG_PATH,
      gzip: true,
      cacheMethod: "none"
    });
  }
  return ocrWorkerPromise;
}

// طابور بسيط بيضمن إن ملف واحد بس بيتعمل له OCR في نفس اللحظة — عشان
// الاستخدام الكتير للمعالج (CPU) في نفس الوقت ممكن يوقع سيرفر Render الضعيف
let ocrQueueTail = Promise.resolve();
function runInOcrQueue(fn) {
  const run = ocrQueueTail.then(fn, fn);
  ocrQueueTail = run.catch(() => {}); // فشل ملف واحد متوقفش اللي بعده
  return run;
}

// نلاقي مكان سطر رقم اللوحة العربي من غير ما نحتاج نقرا الحروف المشوهة خالص:
// بندور على "عنصر نص" شكله رقم من 4 خانات ومعاه كام رمز مش إنجليزي بعده مباشرة
// (ده بالظبط شكل سطر اللوحة العربي زي ما بيطلع من استخراج النص العادي، حتى لو
// الحروف نفسها مش مفهومة) — واستخدام موقعه (y) عشان نعرف نقص الصورة عليه بس.
function findArabicPlateAnchor(items) {
  for (const it of items) {
    const s = it.str || "";
    const m = s.match(/^(\d{4})\s+(.+)$/);
    if (!m) continue;
    if (/[A-Za-z]/.test(m[2])) continue; // ده سطر إنجليزي (زي "9496 A J S")، مش اللي محتاجينه
    if (m[2].replace(/\s+/g, "").length > 10) continue; // طويل قوي، مش شكل 3-4 حروف لوحة
    return it;
  }
  return null;
}

async function ocrArabicPlate(pdf, debugId) {
  return runInOcrQueue(async () => {
    const t0 = Date.now();
    console.log("[ocr] id=" + debugId + " step=start");
    const worker = await getOcrWorker();

    const page = await pdf.getPage(1);
    const viewport1x = page.getViewport({ scale: 1.0 });
    const content = await page.getTextContent();
    const anchor = findArabicPlateAnchor(content.items);

    const scale = OCR_RENDER_SCALE;
    let targetBuffer;

    if (anchor) {
      // مهم لاستهلاك الذاكرة على السيرفر: بدل ما نرسم الصفحة A4 كاملة بدقة عالية
      // (ممكن توصل لأكتر من 100 ميجا في الذاكرة للصفحة الواحدة) وبعدين نقص منها،
      // بنرسم على كانفاس صغير بحجم السطر المطلوب بس من الأول (كذا ميجا بالكتير)
      // — بنزيح نقطة الرسم لفوق (translate) عشان السطر اللي عايزينه يظهر جوه
      // حدود الكانفاس الصغير، وأي حاجة برة الحدود دي متترسمش أصلًا (متتاخدش
      // مساحة في الذاكرة).
      const fontSize = anchor.transform[0];
      const yBaseline = anchor.transform[5];
      const yTopPdf = yBaseline + fontSize * 1.6; // هامش فوق السطر (يغطي الهمزة وعلامات التشكيل)
      const yBottomPdf = yBaseline - fontSize * 0.8; // هامش تحت السطر
      const viewport = page.getViewport({ scale });
      // لازم نعمل الخطوة دي مرة واحدة قبل أي رسم يدوي بتاعنا إحنا (مش عن طريق
      // دوال unpdf الجاهزة) — بتظبط شوية أوامر لازمة لمكتبة الرسم تشتغل في Node
      await createIsomorphicCanvasFactory(() => import("@napi-rs/canvas"));
      const cropTopPx = Math.max(0, Math.round((viewport1x.height - yTopPdf) * scale));
      const cropBottomPx = Math.min(viewport.height, Math.round((viewport1x.height - yBottomPdf) * scale));
      const cropHeightPx = Math.max(8, cropBottomPx - cropTopPx);

      const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(cropHeightPx));
      const ctx = canvas.getContext("2d");
      ctx.translate(0, -cropTopPx);
      await page.render({ canvas, canvasContext: ctx, viewport }).promise;
      targetBuffer = canvas.toBuffer("image/png");
    } else {
      // لو ملقيناش مكان السطر بدقة (تصميم شهادة مختلف مثلًا)، منحاولش نصور الصفحة
      // كاملة كحل بديل — ده كان بيسبب مشكلة خطيرة: قراءة صورة الصفحة كاملة بالـOCR
      // أحيانًا بتاخد وقت طويل جدًا (شفنا حالات فضلت شغالة أكتر من 40 ثانية)، وطول
      // الوقت ده كان بيقفل طابور المعالجة بالكامل ويأخر كل الملفات اللي وراه. أفضل
      // بكتير نسيب الملف ده باسمه الأصلي (REPO/CRN) بدل ما نخاطر بتعليق الباقي كله.
      console.log("[ocr] id=" + debugId + " step=no-anchor-skip");
      return null;
    }
    console.log("[ocr] id=" + debugId + " step=rendered ms=" + (Date.now() - t0) + " cropped=" + !!anchor + " bytes=" + targetBuffer.length);

    // حماية إضافية: لو القراءة نفسها علّقت لأي سبب (حتى بعد القص)، نوقف الـworker
    // بالقوة (terminate) بدل ما نسيبه معلق يقفل الطابور للأبد — وهنعمل واحد جديد
    // بدل منه تلقائيًا في المرة الجاية.
    const RECOGNIZE_TIMEOUT_MS = 15000;
    let recognizeTimer;
    const recognizePromise = worker.recognize(targetBuffer);
    const timeoutPromise = new Promise((_, reject) => {
      recognizeTimer = setTimeout(() => reject(new Error("recognize-timeout")), RECOGNIZE_TIMEOUT_MS);
    });
    let text;
    try {
      const result = await Promise.race([recognizePromise, timeoutPromise]);
      text = result.data.text;
    } catch (e) {
      console.log("[ocr] id=" + debugId + " step=recognize-failed ms=" + (Date.now() - t0) + " error=" + (e && e.message));
      // الـworker ممكن يكون لسه شغال جوه، نتخلص منه ونعمل واحد جديد بدل منه
      worker.terminate().catch(() => {});
      ocrWorkerPromise = null;
      return null;
    } finally {
      clearTimeout(recognizeTimer);
    }
    console.log("[ocr] id=" + debugId + " step=recognized ms=" + (Date.now() - t0) + " text=" + JSON.stringify(String(text || "").slice(0, 300)));
    return text ? extractPlateFromText(text) : null;
  });
}


/* =================================================================
 * اسم المؤجر (طالب التنفيذ) — عشان نعرف كل شهادة تبع أنهي شركة
 * -----------------------------------------------------------------
 * - توثيق (ARN): الاسم نص عادي مقروء → بنقراه مباشرة من النص.
 * - سجل (REPO/CRN): الاسم العربي متشفّر، لكن رقم السجل التجاري للمؤجر أرقام
 *   عادية → بنقراه سريع، ولو أول مرة نشوف الرقم ده بنعمل OCR لسطر الاسم بس
 *   مرة واحدة ونحفظ (رقم السجل → الاسم) في الذاكرة، وكل شهادة بعدها بنفس الرقم
 *   بتاخد الاسم من الذاكرة من غير OCR.
 * ================================================================= */
const lessorNameByCr = new Map();      // رقم السجل التجاري → اسم الشركة
const lessorOcrPending = new Map();    // رقم السجل → Promise شغالة (عشان منكررش OCR لنفس الشركة)

function cleanLessorName(s) {
  if (!s) return null;
  s = String(s)
    .replace(/[\u064B-\u065F\u0670\u200E\u200F\u202A-\u202E]/g, "")
    .replace(/اسم\s*المؤجر\s*:?/g, " ")
    .replace(/(^|\s)الاسم(?=\s|$|:)/g, " ")
    .replace(/الرقم\s*الوطني\s*الموحد|رقم\s*التواصل|رقم\s*الترخيص|رقم\s*السجل\s*التجاري|السجل\s*التجاري|المدينة|العنوان\s*الوطني/g, " ")
    .replace(/[A-Za-z0-9\u0660-\u0669]/g, " ")
    .replace(/[|:؛;_\[\]{}()<>«»"'`~^*=+\\\/\-.,،؟?!]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const arLetters = (s.match(/[\u0621-\u064A]/g) || []).length;
  if (arLetters < 3 || s.length > 80) return null;
  return s;
}

// توثيق: تحت "المؤجر (طالب التنفيذ)" في سطر "الاسم: ..." (بنجرب الشكلين: عادي ومقلوب)
function extractTawtheeqLessor(text) {
  if (!text) return null;
  const m = text.match(/المؤجر\s*\(?\s*طالب\s*التنفيذ\s*\)?/);
  const from = m ? m.index + m[0].length : text.search(/طالب\s*التنفيذ/);
  if (from < 0) return null;
  const win = text.slice(from, from + 400);
  const STOP = /\s*(?::|\d|الرقم|رقم|المدينة|العنوان|التواصل|الترخيص|السجل|الصفة)/;
  let mm = win.match(/(?:^|\s)الاسم\s*:\s*([^\n\r]+)/);
  if (mm) {
    const v = cleanLessorName(mm[1].split(STOP)[0]);
    if (v) return v;
  }
  mm = win.match(/([^\n\r:]+?)\s*:\s*الاسم/);
  if (mm) {
    const v = cleanLessorName(mm[1]);
    if (v) return v;
  }
  return null;
}

// بيحسب مكان اسم المؤجر في الصفحة (من مواضع عناصر النص، مش من ترتيب النص):
// - توثيق: تحت عنوان "المؤجر (طالب التنفيذ)" أول سطر "الاسم" — بنقص السطر ده من النص الأيمن للصفحة
// - سجل: فوق سطر السجل التجاري للمؤجر مباشرة (أعلى رقم من 10 خانات مش بيبدأ بـ 700)
// بيرجّع { page, cr, band: { yTop, yBottom, x0 } | null, kind }
async function locateLessor(pdf) {
  const page = await pdf.getPage(1);
  const content = await page.getTextContent();
  const W = page.getViewport({ scale: 1.0 }).width;
  const items = content.items.filter((it) => it.str && it.str.trim());

  let crBest = null;
  for (const it of items) {
    const m = it.str.match(/(?<![\d\-])(\d{10})(?!\d)/);
    if (!m || /^700/.test(m[1])) continue;
    const y = it.transform[5];
    if (!crBest || y > crBest.y) crBest = { it, cr: m[1], y };
  }

  // توثيق؟
  let header = null;
  for (const it of items) {
    if (!/(?:طالب\s*التنفيذ|التنفيذ\s*طالب)/.test(it.str)) continue;
    const y = it.transform[5];
    if (!header || y > header.transform[5]) header = it;
  }
  if (header) {
    const hy = header.transform[5];
    let label = null;
    for (const it of items) {
      if (!/الاسم/.test(it.str)) continue;
      const y = it.transform[5];
      if (y < hy - 2 && (!label || y > label.transform[5])) label = it;
    }
    if (label) {
      const f = Math.abs(label.height || label.transform[0]) || 10;
      const ly = label.transform[5];
      return { page, items, cr: crBest ? crBest.cr : null, kind: "tawtheeq",
               band: { yTop: ly + f * 1.1, yBottom: ly - f * 0.5, x0: W * 0.4 } };
    }
  }
  // سجل
  if (crBest) {
    const f = Math.abs(crBest.it.height || crBest.it.transform[0]) || 10;
    const yb = crBest.it.transform[5];
    return { page, items, cr: crBest.cr, kind: "sijil",
             band: { yTop: yb + f * 3.3, yBottom: yb + f * 1.1, x0: 0 } };
  }
  return { page, items, cr: null, band: null, kind: null };
}

// OCR لشريط صغير من الصفحة (سطر اسم المؤجر بس)
async function ocrBand(page, band, debugId) {
  return runInOcrQueue(async () => {
    const worker = await getOcrWorker();
    const viewport1x = page.getViewport({ scale: 1.0 });
    const scale = OCR_RENDER_SCALE;
    const viewport = page.getViewport({ scale });
    await createIsomorphicCanvasFactory(() => import("@napi-rs/canvas"));
    const x0px = Math.max(0, Math.round((band.x0 || 0) * scale));
    const cropTopPx = Math.max(0, Math.round((viewport1x.height - band.yTop) * scale));
    const cropBottomPx = Math.min(viewport.height, Math.round((viewport1x.height - band.yBottom) * scale));
    const cropHeightPx = Math.max(8, cropBottomPx - cropTopPx);
    const cropWidthPx = Math.max(8, Math.ceil(viewport.width) - x0px);
    const canvas = createCanvas(cropWidthPx, Math.ceil(cropHeightPx));
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, cropWidthPx, Math.ceil(cropHeightPx));
    ctx.translate(-x0px, -cropTopPx);
    await page.render({ canvas, canvasContext: ctx, viewport }).promise;
    const buf = canvas.toBuffer("image/png");

    let timer;
    try {
      const result = await Promise.race([
        worker.recognize(buf),
        new Promise((_, reject) => { timer = setTimeout(() => reject(new Error("recognize-timeout")), 15000); })
      ]);
      const raw = String(result.data.text || "");
      const name = cleanLessorName(raw);
      console.log("[lessor] id=" + debugId + " ocr=" + JSON.stringify(raw.slice(0, 120)) + " -> " + JSON.stringify(name));
      return name;
    } catch (e) {
      console.log("[lessor] id=" + debugId + " ocr-failed " + (e && e.message));
      worker.terminate().catch(() => {});
      ocrWorkerPromise = null;
      return null;
    } finally {
      clearTimeout(timer);
    }
  });
}

// بيرجّع { name, cr } أو null
// توثيق: اسم المؤجر من مكانه في الصفحة. الكلمات في النص بتطلع مقلوبة الترتيب في السطر
// ("الراجحي مصرف :االسم") وكلمة الاسم ساعات بتطلع "االسم"، فبنقرا السطر من مواضع الكلمات:
// عنوان "المؤجر (طالب التنفيذ)" → أول سطر تحته فيه "الاسم" → الكلمات اللي على شماله لحد الحقل اللي بعده.
async function tawtheeqLessorFromItems(pdf) {
  const page = await pdf.getPage(1);
  const content = await page.getTextContent();
  const items = content.items.filter((it) => it.str && it.str.trim()).map((it) => ({ s: it.str.trim(), x: it.transform[4], y: it.transform[5], w: it.width || 0 }));
  const lines = new Map();
  for (const it of items) { const k = Math.round(it.y); let key = null; for (const kk of lines.keys()) if (Math.abs(kk - k) <= 2) { key = kk; break; } if (key == null) { key = k; lines.set(key, []); } lines.get(key).push(it); }
  const lineText = (arr) => arr.slice().sort((a, b) => b.x - a.x).map((i) => i.s).join(" ");
  let headY = null;
  for (const [y, arr] of lines) {
    const t = lineText(arr);
    if (/المؤجر/.test(t) && /طالب/.test(t) && /التنفيذ/.test(t) && !/المستأجر/.test(t)) { if (headY == null || y > headY) headY = y; }
  }
  if (headY == null) return null;
  const isNameLabel = (s) => /^(?:ا?الاسم|االسم|الإسم)\s*:?$/.test(s);
  const below = [...lines.keys()].filter((y) => y < headY && headY - y < 120).sort((a, b) => b - a);
  for (const y of below) {
    const arr = lines.get(y).slice().sort((a, b) => b.x - a.x);   // من اليمين للشمال
    const li = arr.findIndex((i) => isNameLabel(i.s));
    if (li < 0) continue;
    const words = [];
    let prevX = arr[li].x;
    for (let j = li + 1; j < arr.length; j++) {
      const it = arr[j];
      if (it.s === ":" ) { if (!words.length) { prevX = it.x; continue; } break; }
      if (prevX - (it.x + it.w) > 40) break;            // فراغ كبير = حقل تاني
      if (/[:0-9]/.test(it.s)) break;
      words.push(it.s); prevX = it.x;
    }
    const name = cleanLessorName(words.join(" "));
    if (name) return name;
  }
  return null;
}

async function resolveLessor(pdf, text, id) {
  // توثيق (النص مقروء): الاسم من مكانه في الصفحة الأول
  if (/توثيق|Tawtheeq/i.test(text || "")) {
    try {
      const nm = await tawtheeqLessorFromItems(pdf);
      if (nm) { const loc0 = await locateLessor(pdf); return { name: nm, cr: loc0.cr || null }; }
    } catch (e) {}
  }
  const loc = await locateLessor(pdf);
  // شهادات توثيق: الاسم نص عادي مقروء → نقراه مباشرة من النص (أسرع وأدق من OCR، ومن غير ما نخمّن رقم السجل)
  let name = null;
  if (loc.kind === "tawtheeq") name = extractTawtheeqLessor(text);
  if (!name) name = loc.cr ? (lessorNameByCr.get(loc.cr) || null) : null;
  if (!name && loc.band) {
    const run = () => ocrBand(loc.page, loc.band, id);
    if (loc.cr) {
      let pending = lessorOcrPending.get(loc.cr);
      if (!pending) {
        pending = run()
          .then((n) => { if (n) lessorNameByCr.set(loc.cr, n); return n; })
          .finally(() => lessorOcrPending.delete(loc.cr));
        lessorOcrPending.set(loc.cr, pending);
      }
      name = await withTimeout(pending, OCR_TIMEOUT_MS, null);
    } else {
      name = await withTimeout(run(), OCR_TIMEOUT_MS, null);
    }
  }
  if (!name && loc.kind !== "sijil") name = extractTawtheeqLessor(text); // احتياطي أخير (من النص)
  if (!name && !loc.cr) return null;
  return { name: name || null, cr: loc.cr || null };
}

async function mapWithConcurrency(items, limit, fn) {
  const results = new Array(items.length);
  let idx = 0;
  async function worker() {
    while (idx < items.length) {
      const current = idx++;
      try { results[current] = await fn(items[current], current); }
      catch (e) { results[current] = null; }
    }
  }
  const workers = [];
  for (let i = 0; i < Math.min(limit, items.length); i++) workers.push(worker());
  await Promise.all(workers);
  return results;
}

const CERT_RESOLVE_PER_ITEM_TIMEOUT_MS = 15000; // 15 ثانية أقصى حد للملف الواحد — لو زاد، نعتبره فشل ونكمل اللي بعده

function withTimeout(promise, ms, fallbackValue) {
  return new Promise((resolve) => {
    let settled = false;
    const timer = setTimeout(() => {
      if (!settled) { settled = true; resolve(fallbackValue); }
    }, ms);
    promise.then((v) => {
      if (!settled) { settled = true; clearTimeout(timer); resolve(v); }
    }).catch(() => {
      if (!settled) { settled = true; clearTimeout(timer); resolve(fallbackValue); }
    });
  });
}

// authHeader: هيدر Authorization اللي جه من الموقع (Bearer توكن Firebase) — بنمرره زي ما هو للـ Worker
async function resolveOnePlateRaw(id, mode, authHeader) {
  const controller = new AbortController();
  const abortTimer = setTimeout(() => controller.abort(), CERT_RESOLVE_PER_ITEM_TIMEOUT_MS);
  let pdf = null;
  try {
    const res = await fetch(CERT_DOWNLOAD_BASE + "/api/certificates/" + encodeURIComponent(id) + "/download", {
      signal: controller.signal,
      headers: (authHeader && typeof authHeader === "object") ? authHeader : (authHeader ? { Authorization: authHeader } : {})
    });
    if (!res.ok) return { plate: null, lessor: null, error: "download HTTP " + res.status };
    const buf = await res.arrayBuffer();
    pdf = await getDocumentProxy(new Uint8Array(buf));
    const { text } = await extractText(pdf, { mergePages: true });
    const wantPlate = mode !== "lessor";   // mode: "lessor" = أسماء الشركات بس | "plate" = أرقام اللوحات بس | غير كده = الاتنين
    const wantLessor = mode !== "plate";
    let result = null;
    // شهادات سجل: من مكان الحقل في الصفحة (الأدق، وبيقرا العربي من غير OCR)
    // الترتيب: الإنجليزي الأول (سجل بالمكان، وبعده عنوان الحقل في النص)، وبعدين العربي
    if (wantPlate) { try { result = await plateFromSijilItems(pdf, "en"); } catch (e) { result = null; } }
    if (wantPlate && !result && text) result = grabEnglishPlateAfterLabel(text);
    if (result) result = englishPlateToArabic(result);
    if (wantPlate) {
      let ar = null;
      try { ar = await plateFromSijilItems(pdf, "ar"); } catch (e) { ar = null; }
      // لو الشهادة فيها الاتنين واختلفوا، العربي المكتوب في الشهادة هو المرجع
      if (ar && (!result || ar.replace(/\s+/g, "") !== result.replace(/\s+/g, ""))) result = ar;
    }
    if (wantPlate && !result && text) result = extractPlateFromText(text);
    if (wantPlate && !result) {
      // ملاذ أخير: الفونت العربي في الشهادة ده مش قابل للقراءة كنص (مشكلة معروفة
      // في الفونت نفسه)، فنحول الصفحة لصورة ونقراها بالعربي (OCR) بدل النص
      result = await withTimeout(ocrArabicPlate(pdf, id), OCR_TIMEOUT_MS, null);
    }
    // اسم المؤجر (طالب التنفيذ) — فشله مايأثرش على رقم اللوحة
    let lessor = null;
    if (wantLessor) { try { lessor = await resolveLessor(pdf, text || "", id); } catch (e) { lessor = null; } }
    if (result) result = englishPlateToArabic(result); // اللوحة الإنجليزي تتكتب بالحروف العربي
    return { plate: result, lessor };
  } finally {
    clearTimeout(abortTimer);
  }
}

// طبقة حماية إضافية: حتى لو الـ fetch نجح بس استخراج النص نفسه علّق (ملف تالف مثلًا)،
// بعد المهلة القصوى (نص عادي + OCR لو احتاج) بنعتبره فشل ونرجع null بدل ما نستنى للأبد
// المهلة هنا لازم تستحمل وقت الـ OCR كمان (لو احتجناه) مش بس وقت التنزيل والنص العادي
async function resolveOnePlate(id, mode, authHeader) {
  try {
    return await withTimeout(resolveOnePlateRaw(id, mode, authHeader), CERT_RESOLVE_PER_ITEM_TIMEOUT_MS + OCR_TIMEOUT_MS * 2, null);
  } catch (e) {
    return null;
  }
}

// جسم الطلب: { ids: ["driveFileId1", "driveFileId2", ...] }
// الرد: { ok:true, results: { "driveFileId1": "AJS 9496" | null, ... }, lessors: { "driveFileId1": {name, cr} | null }, errors: { id: "سبب الفشل" } }
app.post("/api/certificates/resolve-plates", async (req, res) => {
  try {
    const ids = Array.isArray(req.body && req.body.ids) ? req.body.ids : [];
    const cleanIds = Array.from(new Set(ids.filter((x) => typeof x === "string" && x.trim()))).slice(0, CERT_RESOLVE_MAX_IDS);
    if (!cleanIds.length) return res.status(400).json({ ok: false, error: "لا يوجد ids" });

    const mode = (req.body && typeof req.body.mode === "string") ? req.body.mode : "both";
    const authHeader = req.headers.authorization || "";
    const results = {};
    const lessors = {};
    const errors = {};
    await mapWithConcurrency(cleanIds, CERT_RESOLVE_CONCURRENCY, async (id) => {
      try {
        const r = await resolveOnePlate(id, mode, authHeader);
        results[id] = r ? r.plate : null;
        lessors[id] = r ? r.lessor : null;
        if (r && r.error) errors[id] = r.error;
      } catch (e) { results[id] = null; lessors[id] = null; errors[id] = "exception: " + String((e && e.message) || e).slice(0, 80); }
    });

    // تشخيص (بدون أي محتوى ملفات): عدد الملفات، هل جه توكن، كام لوحة اتقرت، وأسباب الفشل
    try {
      const okCount = cleanIds.filter((i) => results[i]).length;
      console.log("[resolve] n=" + cleanIds.length + " mode=" + mode + " hasAuth=" + !!authHeader +
        " plates=" + okCount + " errors=" + JSON.stringify(errors).slice(0, 300));
    } catch (_) {}
    // lessors: { id: { name: "شركة ..." | null, cr: "4030206631" | null } | null }
    res.json({ ok: true, v: 2, results, lessors, errors });
  } catch (e) {
    res.status(500).json({ ok: false, error: String((e && e.message) || e) });
  }
});


/* ============================================================
   استخراج تلقائي بالليل (من غير ما حد يفتح التطبيق)
   - Worker الشهادات بينده على POST /api/certificates/auto-run كل 5 دقايق بمفتاح سري (SERVER_KEY)
   - السيرفر بيجيب الشهادات الجديدة، يستخرج لوحاتها وأسماء مؤجريها، ويحفظ النتايج في Firestore
     (collection: certMetaDay ، مستند لكل يوم) والتطبيق بيقراها أول ما يفتح
   - لو SERVER_KEY مش متضاف في Variables: الـ endpoint مقفول تمامًا
   ============================================================ */
const crypto = require("crypto");
const SERVER_KEY = process.env.SERVER_KEY || "";
const FS_PROJECT = process.env.FIRESTORE_PROJECT || "jta-cloud-daeb1";
const FS_BASE = "https://firestore.googleapis.com/v1/projects/" + FS_PROJECT + "/databases/(default)/documents/certMetaDay/";
const AUTO_HOURS = 72;          // بنراجع آخر 72 ساعة
// مفيش حد لعدد الشهادات: الجولة بتستخرج كل اللي لقته لحد ما تخلّص.
// الحد الوحيد مهلة الطلب في Cloud Run (أقصاها ساعة -- لازم --timeout=3600): بنقف قبلها بـ 5 دقايق،
// واللي فضل (لو فضل) بيكمّل في الجولة اللي بعدها على طول.
const AUTO_MAX_PER_RUN = Number(process.env.AUTO_MAX_PER_RUN || 0) || Infinity;
const AUTO_BUDGET_MS = Math.max(60000, Number(process.env.AUTO_BUDGET_MS || 3300000));   // 55 دقيقة
const AUTO_FLUSH_MS = 30000;    // كل 30 ثانية: اللي اتقرا يتضاف للتقارير على طول (مايستناش آخر الجولة)
// كام شهادة بتتفتح في نفس الوقت في الجولة التلقائية (كانت 3 -> 120 شهادة بتاخد دقيقة ونص).
// الافتراضي 12 (محتاج --cpu 2 --memory 2Gi). ممكن تتغير من المتغير AUTO_CONCURRENCY من غير رفع كود.
const AUTO_CONCURRENCY = Math.max(1, Math.min(32, Number(process.env.AUTO_CONCURRENCY || 12)));
const AUTO_MAX_ATTEMPTS = 3;
let autoRunning = false;
let gcpTok = { v: "", exp: 0 };

function keyEq(a, b) {
  const x = crypto.createHash("sha256").update(String(a)).digest();
  const y = crypto.createHash("sha256").update(String(b)).digest();
  return crypto.timingSafeEqual(x, y);
}

async function gcpToken() {
  if (gcpTok.v && Date.now() < gcpTok.exp) return gcpTok.v;
  const r = await fetch("http://metadata.google.internal/computeMetadata/v1/instance/service-accounts/default/token", { headers: { "Metadata-Flavor": "Google" } });
  if (!r.ok) throw new Error("metadata token HTTP " + r.status);
  const d = await r.json();
  gcpTok = { v: d.access_token, exp: Date.now() + Math.max(60, (d.expires_in || 600) - 120) * 1000 };
  return gcpTok.v;
}

function dayOf(f) {
  const t = Date.parse(f.createdTime || f.modifiedTime || "");
  return new Date(isNaN(t) ? Date.now() : t).toISOString().slice(0, 10); // يوم UTC
}

// بيرجع { id: { p, l, c, a } } لليوم ده
async function fsGetDay(day) {
  const r = await fetch(FS_BASE + day, { headers: { Authorization: "Bearer " + (await gcpToken()) } });
  if (r.status === 404) return {};
  if (!r.ok) throw new Error("firestore get HTTP " + r.status + " " + (await r.text()).slice(0, 150));
  const d = await r.json();
  const m = (((d.fields || {}).m || {}).mapValue || {}).fields || {};
  const out = {};
  for (const id of Object.keys(m)) {
    const f = ((m[id] || {}).mapValue || {}).fields || {};
    out[id] = {
      p: (f.p && f.p.stringValue) || "",
      l: (f.l && f.l.stringValue) || "",
      c: (f.c && f.c.stringValue) || "",
      a: Number((f.a && f.a.integerValue) || 0),
      // r: 1 = اتضافت للتقرير، 0 = مستنية تتضاف، -1 = قديمة (قبل الحقل ده -- بنعتبرها اتضافت)
      r: (f.r && f.r.integerValue != null) ? Number(f.r.integerValue) : -1
    };
  }
  return out;
}

async function fsSaveDay(day, entries) {
  const ids = Object.keys(entries);
  if (!ids.length) return;
  const fields = {};
  for (const id of ids) {
    const e = entries[id];
    fields[id] = { mapValue: { fields: {
      p: { stringValue: e.p || "" },
      l: { stringValue: e.l || "" },
      c: { stringValue: e.c || "" },
      a: { integerValue: String(e.a || 0) },
      r: { integerValue: String(e.r == null ? -1 : e.r) }
    } } };
  }
  const mask = ids.map((id) => "updateMask.fieldPaths=" + encodeURIComponent("m.`" + id.replace(/`/g, "") + "`")).join("&");
  const r = await fetch(FS_BASE + day + "?" + mask, {
    method: "PATCH",
    headers: { Authorization: "Bearer " + (await gcpToken()), "Content-Type": "application/json" },
    body: JSON.stringify({ fields: { m: { mapValue: { fields } } } })
  });
  if (!r.ok) throw new Error("firestore save HTTP " + r.status + " " + (await r.text()).slice(0, 150));
}

/* ============================================================
   إشعارات الآيفون (Push) لحسابات الأدمن: "نزلت شهادات جديدة"
   - تطبيق الآيفون بيحفظ رقم الجهاز في users/{uid}.pushTokens (للأدمن بس)
   - بعد كل جولة auto-run: لو فيه شهادات أول مرة تتشاف -> إشعار لكل أجهزة الأدمن
   - محتاج 3 Variables في Cloud Run (من حساب Apple Developer):
       APNS_KEY     = محتوى ملف AuthKey_XXXX.p8 كامل
       APNS_KEY_ID  = Key ID (10 حروف)
       APNS_TEAM_ID = Team ID (10 حروف)
     ولو مش موجودين: الإشعارات بتتقفل بهدوء وكل حاجة تانية شغالة عادي.
   ============================================================ */
const http2 = require("http2");
const APNS_KEY = String(process.env.APNS_KEY || "").replace(/\\n/g, "\n").trim();
const APNS_KEY_ID = String(process.env.APNS_KEY_ID || "").trim();
const APNS_TEAM_ID = String(process.env.APNS_TEAM_ID || "").trim();
const APNS_TOPIC = String(process.env.APNS_TOPIC || "com.osos.app").trim();
const FS_DOCS = "https://firestore.googleapis.com/v1/projects/" + FS_PROJECT + "/databases/(default)/documents";
let apnsJwt = { v: "", at: 0 };

function b64url(buf) {
  return Buffer.from(buf).toString("base64").replace(/=+$/, "").replace(/\+/g, "-").replace(/\//g, "_");
}

function apnsAuth() {
  // Apple بتقبل نفس التوكن لحد ساعة -- بنجدده كل 40 دقيقة
  if (apnsJwt.v && Date.now() - apnsJwt.at < 40 * 60 * 1000) return apnsJwt.v;
  const head = b64url(JSON.stringify({ alg: "ES256", kid: APNS_KEY_ID }));
  const claims = b64url(JSON.stringify({ iss: APNS_TEAM_ID, iat: Math.floor(Date.now() / 1000) }));
  const sig = crypto.sign("sha256", Buffer.from(head + "." + claims), { key: APNS_KEY, dsaEncoding: "ieee-p1363" });
  apnsJwt = { v: head + "." + claims + "." + b64url(sig), at: Date.now() };
  return apnsJwt.v;
}

// اتصال واحد لكل سيرفر Apple (HTTP/2 بيبعت كتير على نفس الاتصال)
function apnsSession(host, cache) {
  let c = cache[host];
  if (c && !c.closed && !c.destroyed) return c;
  c = http2.connect("https://" + host);
  c.on("error", () => {});
  cache[host] = c;
  return c;
}

function apnsSendOne(host, token, payload, cache) {
  return new Promise((resolve) => {
    let done = false;
    const finish = (r) => { if (!done) { done = true; resolve(r); } };
    let req;
    try {
      req = apnsSession(host, cache).request({
        ":method": "POST",
        ":path": "/3/device/" + token,
        "authorization": "bearer " + apnsAuth(),
        "apns-topic": APNS_TOPIC,
        "apns-push-type": "alert",
        "apns-priority": "10",
        "content-type": "application/json"
      });
    } catch (e) { return finish({ status: 0, reason: String((e && e.message) || e) }); }
    let status = 0, body = "";
    req.on("response", (h) => { status = Number(h[":status"]) || 0; });
    req.setEncoding("utf8");
    req.on("data", (c) => { body += c; });
    req.on("end", () => {
      let reason = "";
      try { reason = (JSON.parse(body || "{}").reason) || ""; } catch (_) {}
      finish({ status, reason });
    });
    req.on("error", (e) => finish({ status: 0, reason: String((e && e.message) || e) }));
    req.setTimeout(15000, () => { try { req.close(); } catch (_) {} finish({ status: 0, reason: "timeout" }); });
    req.end(JSON.stringify(payload));
  });
}

// النسخة من المتجر/TestFlight بتستخدم سيرفر Apple الأساسي، ونسخة التجربة من Xcode بتستخدم sandbox
async function apnsSend(token, payload, cache) {
  let r = await apnsSendOne("api.push.apple.com", token, payload, cache);
  if (r.status === 400 && r.reason === "BadDeviceToken") r = await apnsSendOne("api.sandbox.push.apple.com", token, payload, cache);
  return r;
}

const apnsReady = () => !!(APNS_KEY && APNS_KEY_ID && APNS_TEAM_ID);

/* ---------- إشعارات الأندرويد (Firebase Cloud Messaging) ----------
   مش محتاجة أي مفتاح: السيرفر بيبعت بصلاحية حساب Cloud Run نفسه (نفس اللي بيكتب في Firestore) */
const FCM_URL = "https://fcm.googleapis.com/v1/projects/" + (process.env.FIRESTORE_PROJECT || "jta-cloud-daeb1") + "/messages:send";
async function fcmSend(token, payload) {
  const alert = ((payload && payload.aps) || {}).alert || {};
  const data = {};
  if (payload && payload.tab) data.tab = String(payload.tab);
  const message = {
    token,
    notification: { title: String(alert.title || "OS OS"), body: String(alert.body || "") },
    data,
    android: { priority: "HIGH", notification: { channel_id: "osos_alerts", sound: "default" } }
  };
  try {
    const r = await fetch(FCM_URL, {
      method: "POST",
      headers: { Authorization: "Bearer " + (await gcpToken()), "Content-Type": "application/json" },
      body: JSON.stringify({ message })
    });
    if (r.ok) return { status: 200, reason: "" };
    const t = await r.text().catch(() => "");
    const gone = r.status === 404 || /UNREGISTERED|registration-token-not-registered|not a valid FCM registration token/i.test(t);
    return { status: r.status, reason: gone ? "Unregistered" : t.slice(0, 160) };
  } catch (e) {
    return { status: 0, reason: String((e && e.message) || e).slice(0, 160) };
  }
}
// في حاجة نقدر نبعت بيها؟ (الأندرويد دايمًا، والآيفون لما مفتاح Apple يتحط)
const pushReady = () => true;

// يبعت لقايمة أجهزة [{uid, token}] -- 20 مع بعض -- ويشيل الأرقام اللي Apple قالت إنها مبقتش شغالة
async function pushToDevices(list, payload) {
  const seen = new Set();
  list = list.filter((x) => x && x.token && !seen.has(x.token) && seen.add(x.token));
  if (!list.length) return { sent: 0, devices: 0 };
  const cache = {};
  let sent = 0, failed = 0, removed = 0;
  try {
    for (let i = 0; i < list.length; i += 20) {
      const part = list.slice(i, i + 20);
      const rs = await Promise.all(part.map((it) => it.p === "android"
        ? fcmSend(it.token, payload)
        : (apnsReady() ? apnsSend(it.token, payload, cache) : Promise.resolve({ status: -1, reason: "no apns key" }))));
      for (let k = 0; k < part.length; k++) {
        const r = rs[k], it = part[k];
        if (r.status === 200) { sent++; continue; }
        if (r.status === -1) continue;   // آيفون ومفتاح Apple لسه مش متحط
        failed++;
        if (r.status === 410 || r.reason === "Unregistered" || r.reason === "BadDeviceToken" || r.reason === "DeviceTokenNotForTopic") {
          await removePushToken(it.uid, it.token);
          removed++;
        } else {
          console.log("[push] fail " + r.status + " " + r.reason);
        }
      }
    }
  } finally {
    for (const h of Object.keys(cache)) { try { cache[h].close(); } catch (_) {} }
  }
  return { sent, failed, removed, devices: list.length };
}

function tokensOfDoc(doc) {
  const uid = String((doc && doc.name) || "").split("/").pop();
  const toks = ((((doc && doc.fields) || {}).pushTokens || {}).mapValue || {}).fields || {};
  const out = [];
  for (const tok of Object.keys(toks)) {
    const f = ((toks[tok] || {}).mapValue || {}).fields || {};
    const p = (f.p && f.p.stringValue) || "ios";
    if (p === "ios" && /^[0-9a-f]{32,200}$/i.test(tok)) out.push({ uid, token: tok, p });
    else if (p === "android" && /^[A-Za-z0-9_:\-]{32,400}$/.test(tok)) out.push({ uid, token: tok, p });
  }
  return out;
}

async function fsRunQuery(structuredQuery) {
  const r = await fetch(FS_DOCS + ":runQuery", {
    method: "POST",
    headers: { Authorization: "Bearer " + (await gcpToken()), "Content-Type": "application/json" },
    body: JSON.stringify({ structuredQuery })
  });
  if (!r.ok) throw new Error("firestore query HTTP " + r.status + " " + (await r.text()).slice(0, 150));
  const rows = await r.json();
  return (Array.isArray(rows) ? rows : []).map((x) => x && x.document).filter(Boolean);
}

async function fsGetDoc(path) {
  const r = await fetch(FS_DOCS + "/" + path, { headers: { Authorization: "Bearer " + (await gcpToken()) } });
  if (r.status === 404) return null;
  if (!r.ok) throw new Error("firestore get HTTP " + r.status);
  return r.json();
}

const fv = (doc, k) => {
  const f = ((doc && doc.fields) || {})[k];
  if (!f) return undefined;
  if ("stringValue" in f) return f.stringValue;
  if ("booleanValue" in f) return f.booleanValue;
  if ("integerValue" in f) return Number(f.integerValue);
  if ("arrayValue" in f) return ((f.arrayValue || {}).values || []).map((v) => v.stringValue || "");
  if ("timestampValue" in f) return f.timestampValue;
  return undefined;
};

// أجهزة الأدمن: [{ uid, token }]
async function adminPushTokens() {
  const docs = await fsRunQuery({
    from: [{ collectionId: "users" }],
    where: { fieldFilter: { field: { fieldPath: "isAdmin" }, op: "EQUAL", value: { booleanValue: true } } }
  });
  return docs.flatMap(tokensOfDoc);
}

// رقم جهاز Apple قالت إنه مبقاش شغال (التطبيق اتمسح مثلاً) -> نشيله من الحساب
async function removePushToken(uid, tok) {
  const mask = "updateMask.fieldPaths=" + encodeURIComponent("pushTokens.`" + tok + "`");
  await fetch(FS_DOCS + "/users/" + encodeURIComponent(uid) + "?" + mask + "&currentDocument.exists=true", {
    method: "PATCH",
    headers: { Authorization: "Bearer " + (await gcpToken()), "Content-Type": "application/json" },
    body: JSON.stringify({ fields: {} })
  }).catch(() => {});
}

async function pushNewCertsToAdmins(n) {
  if (!n) return { sent: 0 };
  if (!pushReady()) return { sent: 0, skipped: "no push" };
  return pushToDevices(await adminPushTokens(), {
    aps: { alert: { title: "📄 لوحات تثبيت جديدة (" + n + ")" }, sound: "default" },
    tab: "cert"
  });
}

/* ---------- حساب اللي بيكلّم السيرفر (من التطبيق) ----------
   التطبيق بيبعت Firebase ID token في Authorization، وإحنا بنتأكد منه عند Google */
const FB_API_KEY = String(process.env.FIREBASE_API_KEY || "AIzaSyBdV7HnjXHJfnvRCFVOcE1MMetCBov7-L8").trim();
async function uidFromReq(req) {
  const m = /^Bearer\s+(.+)$/i.exec(String(req.headers.authorization || ""));
  if (!m) return "";
  const r = await fetch("https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=" + encodeURIComponent(FB_API_KEY), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ idToken: m[1] })
  });
  if (!r.ok) return "";
  const d = await r.json().catch(() => null);
  return (d && Array.isArray(d.users) && d.users[0] && d.users[0].localId) || "";
}

// تسجيل/شيل رقم جهاز الآيفون للحساب (أي مشترك)
// POST /api/push/register  { token, on: true|false, label }
app.post("/api/push/register", async (req, res) => {
  try {
    const uid = await uidFromReq(req);
    if (!uid) return res.status(401).json({ ok: false, error: "auth" });
    const plat = String((req.body && req.body.platform) || "ios") === "android" ? "android" : "ios";
    const tok = plat === "android" ? String((req.body && req.body.token) || "") : String((req.body && req.body.token) || "").toLowerCase();
    if (plat === "ios" ? !/^[0-9a-f]{32,200}$/.test(tok) : !/^[A-Za-z0-9_:\-]{32,400}$/.test(tok)) return res.status(400).json({ ok: false, error: "token" });
    const on = !(req.body && req.body.on === false);
    const mask = "updateMask.fieldPaths=" + encodeURIComponent("pushTokens.`" + tok + "`") + "&currentDocument.exists=true";
    const body = on ? { fields: { pushTokens: { mapValue: { fields: { [tok]: { mapValue: { fields: {
      p: { stringValue: plat },
      at: { stringValue: new Date().toISOString() },
      label: { stringValue: String((req.body && req.body.label) || "").slice(0, 80) }
    } } } } } } } } : { fields: {} };
    const r = await fetch(FS_DOCS + "/users/" + encodeURIComponent(uid) + "?" + mask, {
      method: "PATCH",
      headers: { Authorization: "Bearer " + (await gcpToken()), "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
    if (!r.ok) return res.status(500).json({ ok: false, error: "save HTTP " + r.status });
    res.json({ ok: true, push: plat === "android" ? true : apnsReady() });
  } catch (e) {
    res.status(500).json({ ok: false, error: String((e && e.message) || e).slice(0, 200) });
  }
});

// إشعار الإدارة الجديد -> ستارة الآيفون للمشتركين اللي اتحددوا (أو للكل)
// POST /api/push/post  { id }   (الأدمن بس)
app.post("/api/push/post", async (req, res) => {
  try {
    const uid = await uidFromReq(req);
    if (!uid) return res.status(401).json({ ok: false, error: "auth" });
    const me = await fsGetDoc("users/" + encodeURIComponent(uid));
    if (!me || fv(me, "isAdmin") !== true) return res.status(403).json({ ok: false, error: "admin only" });
    const id = String((req.body && req.body.id) || "");
    if (!/^[\w-]{1,80}$/.test(id)) return res.status(400).json({ ok: false, error: "id" });
    const post = await fsGetDoc("announcements/" + id);
    if (!post) return res.status(404).json({ ok: false, error: "post not found" });
    if (fv(post, "pushedAt")) return res.json({ ok: true, skipped: "already sent" });
    if (!pushReady()) return res.json({ ok: true, sent: 0, skipped: "no push" });

    // نعلّم الإشعار إنه اتبعت الأول (عشان مايتبعتش مرتين لو الطلب اتكرر)
    await fetch(FS_DOCS + "/announcements/" + id + "?updateMask.fieldPaths=pushedAt&currentDocument.exists=true", {
      method: "PATCH",
      headers: { Authorization: "Bearer " + (await gcpToken()), "Content-Type": "application/json" },
      body: JSON.stringify({ fields: { pushedAt: { stringValue: new Date().toISOString() } } })
    });

    const report = fv(post, "report") === true;
    const masked = fv(post, "masked") === true;
    const text = String(fv(post, "text") || "");
    const companies = fv(post, "companies") || [];
    const payload = report
      ? { aps: { alert: { title: "📊 تقرير يومي جديد: " + companies.join("، "), body: "افتح شاشة الشهادات ← التقارير اليومية" }, sound: "default" }, tab: "reports" }
      : { aps: { alert: { title: "📢 إشعار من الإدارة", body: masked ? "وصل إشعار جديد من الإدارة — افتح التطبيق لعرضه" : text.slice(0, 400) }, sound: "default" } };

    // المستلمين: الكل، أو الحسابات اللي الأدمن اختارها (+ الأدمن بيشوف كل الإشعارات)
    let docs;
    if (fv(post, "audience") === "some") {
      const targets = (fv(post, "targets") || []).filter(Boolean).slice(0, 500);
      docs = [];
      for (let i = 0; i < targets.length; i += 10) {
        const part = await Promise.all(targets.slice(i, i + 10).map((t) => fsGetDoc("users/" + encodeURIComponent(t)).catch(() => null)));
        docs.push(...part.filter(Boolean));
      }
    } else {
      docs = await fsRunQuery({ from: [{ collectionId: "users" }], select: { fields: [{ fieldPath: "pushTokens" }, { fieldPath: "status" }] } });
    }
    docs = docs.filter((d) => fv(d, "status") !== "paused");
    const out = await pushToDevices(docs.flatMap(tokensOfDoc), payload);
    res.json(Object.assign({ ok: true }, out));
  } catch (e) {
    res.status(500).json({ ok: false, error: String((e && e.message) || e).slice(0, 200) });
  }
});

/* ============================================================
   التقرير اليومي التلقائي
   كل شهادة جديدة اتقرا رقم لوحتها واسم الشركة المؤجرة بتاعتها، بتنضاف لوحتها لتقرير الشركة دي
   في «التقارير اليومية» للمشتركين على طول (من غير ما الأدمن يعمل حاجة).
   - الشركة اللي اسمها في قايمة شركات التقارير بتنضاف لنفس كارتها (مع تقارير الأدمن).
   - أي شركة تانية ليها كارت باسمها.
   - الشهادة اللي اسم شركتها مااتقراش: مابتدخلش تقرير (بتفضل في شاشة الشهادات بس).
   AUTO_REPORTS=0 في Cloud Run بيقفل الخاصية دي.
   ============================================================ */
const AUTO_REPORTS = String(process.env.AUTO_REPORTS || "1") !== "0";
// نفس قايمة التطبيق (index.html: REPORT_COMPANIES)
const REPORT_COMPANIES = ["البنك الأهلي السعودي", "شركة الجبر للتمويل", "شركة راية للتمويل", "مصرف الراجحي", "شركة عبداللطيف جميل المتحدة", "شركة التيسير العربية للتمويل", "الفرنسي للتأجير", "شركة إيجارة للتمويل", "بنك البلاد", "الإمارات", "تأجير تمويلي", "جبر تمويلي", "بنك الرياض", "مصرف الإنماء", "بنك الجزيرة", "البنك العربي الوطني", "شركة امكان للتمويل"];

// الاسم الرسمي في الشهادة -> كارت الشركة في التقارير
const REPORT_ALIASES = [
  ["الراجحي", "مصرف الراجحي"],   // "شركة الراجحي المصرفية للاستثمار" = مصرف الراجحي
  // «تمويل الأولى»: الـ PDF بيطلّع «لا» مقلوبة والـ «ى» بتضيع -> «تمويل االول»
  ["تمويل االول", "تمويل الأولى"],
  ["تمويل االوىل", "تمويل الأولى"],   // ده اللي السيرفر بيقراه فعلًا (certMetaDay 9/11)
  ["تمويل الاول", "تمويل الأولى"]
];

function coNorm(s) {
  return String(s || "")
    .replace(/[ً-ٰٟـ]/g, "")
    .replace(/[أإآ]/g, "ا").replace(/ة/g, "ه").replace(/ى/g, "ي")
    .replace(/(^|\s)شركه(?=\s|$)/g, " ")
    .replace(/\s+/g, "");
}

// اسم الشركة في التقرير: لو الاسم المقروء هو واحدة من شركات القايمة -> اسمها في القايمة (عشان كارت واحد)
function reportCompanyOf(lessor) {
  const raw = String(lessor || "").replace(/\s+/g, " ").trim();
  if (!raw) return "";
  const n = coNorm(raw);
  if (n.length < 3) return "";
  // أسماء رسمية مختلفة عن اسم الكارت (متأكدين منها)
  for (const [k, co] of REPORT_ALIASES) if (n.includes(coNorm(k))) return co;
  let best = "", bestLen = 0;
  for (const c of REPORT_COMPANIES) {
    const cn = coNorm(c);
    if (cn.length < 4) continue;
    if (n === cn || n.includes(cn) || (cn.includes(n) && n.length >= 6)) {
      if (cn.length > bestLen) { best = c; bestLen = cn.length; }
    }
  }
  return best || raw;
}

// اللوحات اللي موجودة فعلاً في التقارير اليومية (آخر 24 ساعة) -> مانعملش بيها تقرير ولا إشعار تاني
async function recentReportPlateKeys() {
  const have = new Set();
  const since = Date.now() - AUTO_REPORT_MAX_AGE_MS;
  const docs = await fsRunQuery({
    from: [{ collectionId: "announcements" }],
    where: { fieldFilter: { field: { fieldPath: "report" }, op: "EQUAL", value: { booleanValue: true } } }
  });
  for (const d of docs) {
    const idN = Number(fv(d, "id"));
    const t = isFinite(idN) && idN > 1e12 ? idN : (Date.parse(fv(d, "createdAt") || "") || 0);
    if (t && t < since) continue;
    String(fv(d, "text") || "").split(/\r?\n/).forEach((l) => { const k = plateKeyLoose(l.trim()); if (k) have.add(k); });
  }
  return have;
}

async function createAutoReports(byCompany) {
  const cosIn = Object.keys(byCompany);
  if (!cosIn.length) return { reports: 0, handled: true };
  // نشيل اللوحات اللي اتكتبت قبل كده (نسخة تانية من نفس الشهادة «(1).pdf» أو اترفعت تاني)
  const have = await recentReportPlateKeys();
  const fresh = {};   // الشركة -> اللوحات الجديدة بس
  let dup = 0;
  for (const co of cosIn) {
    const seenHere = new Set();
    for (const p of byCompany[co] || []) {
      if (!p) continue;
      const k = plateKeyLoose(p) || String(p).replace(/\s+/g, "");
      if (seenHere.has(k)) continue;
      seenHere.add(k);
      if (have.has(k)) { dup++; continue; }
      have.add(k);
      (fresh[co] = fresh[co] || []).push(p);
    }
  }
  const cos = Object.keys(fresh);
  if (!cos.length) return { reports: 0, dup, handled: true };   // مفيش جديد -> لا تقرير ولا إشعار
  const base = Date.now();
  let made = 0;
  const madeCos = [];
  for (let i = 0; i < cos.length; i++) {
    const co = cos[i];
    const plates = fresh[co];
    const id = String(base + i);
    const fields = {
      text: { stringValue: plates.join("\n") },
      id: { stringValue: id },
      showClose: { booleanValue: true },
      showSort: { booleanValue: true },
      masked: { booleanValue: false },
      showNote: { booleanValue: true },
      report: { booleanValue: true },
      auto: { booleanValue: true },
      companies: { arrayValue: { values: [{ stringValue: co }] } },
      audience: { stringValue: "all" },
      targets: { arrayValue: { values: [] } },
      targetNames: { arrayValue: { values: [] } },
      pushPending: { booleanValue: true },   // الإشعار بيطلع بعد ما السيرفر يخلّص كل الشهادات (pushPendingReports)
      createdAt: { timestampValue: new Date().toISOString() }
    };
    const r = await fetch(FS_DOCS + "/announcements/" + id + "?currentDocument.exists=false", {
      method: "PATCH",
      headers: { Authorization: "Bearer " + (await gcpToken()), "Content-Type": "application/json" },
      body: JSON.stringify({ fields })
    });
    if (r.ok) { made++; madeCos.push(co); }
    else console.log("[auto-report] save HTTP " + r.status + " " + (await r.text()).slice(0, 150));
  }
  return { reports: made, companies: cos.length, dup, handled: made === cos.length };
}

// بعد ما السيرفر يخلّص كل الشهادات اللي عنده: إشعار واحد بكل اللي اتضاف من آخر إشعار
// (التقارير اللي لسه pushPending=true)، وبعدها بتتعلّم إنها اتبلّغت. لو شغالين على 500 شهادة مايطلعش إشعار كل شوية.
async function pushPendingReports() {
  const docs = await fsRunQuery({
    from: [{ collectionId: "announcements" }],
    where: { fieldFilter: { field: { fieldPath: "pushPending" }, op: "EQUAL", value: { booleanValue: true } } }
  });
  if (!docs.length) return { pending: 0 };
  const byCo = {}, seenK = new Set();
  docs.sort((x, y) => String(fv(x, "id") || "").localeCompare(String(fv(y, "id") || ""), undefined, { numeric: true }));
  for (const d of docs) {
    const co = (fv(d, "companies") || [])[0] || "";
    String(fv(d, "text") || "").split(/\r?\n/).map((l) => l.trim()).filter(Boolean).forEach((pl) => {
      const k = co + "|" + (plateKeyLoose(pl) || pl.replace(/\s+/g, ""));
      if (seenK.has(k)) return;
      seenK.add(k);
      (byCo[co] = byCo[co] || []).push(pl);
    });
  }
  let push = null;
  const groups = Object.keys(byCo).map((co) => ({ co, plates: byCo[co] })).filter((g) => g.plates.length);
  if (groups.length && pushReady()) {
    try {
      const msg = reportPushText(groups);
      let users = await fsRunQuery({ from: [{ collectionId: "users" }], select: { fields: [{ fieldPath: "pushTokens" }, { fieldPath: "status" }] } });
      users = users.filter((d) => fv(d, "status") !== "paused");
      push = await pushToDevices(users.flatMap(tokensOfDoc), {
        aps: { alert: { title: msg.title, body: msg.body }, sound: "default" },
        tab: "reports"   // الضغط على الإشعار بيفتح «التقارير اليومية» على طول
      });
    } catch (e) { push = { error: String((e && e.message) || e).slice(0, 150) }; }
  }
  // نعلّمهم «اتبلّغوا» (حتى لو الإشعار مااتبعتش، علشان التقارير تظهر في التطبيق ومانفضلش نحاول)
  const now = new Date().toISOString();
  for (const d of docs) {
    try {
      await fetch("https://firestore.googleapis.com/v1/" + d.name + "?updateMask.fieldPaths=pushPending&updateMask.fieldPaths=pushedAt&currentDocument.exists=true", {
        method: "PATCH",
        headers: { Authorization: "Bearer " + (await gcpToken()), "Content-Type": "application/json" },
        body: JSON.stringify({ fields: { pushPending: { booleanValue: false }, pushedAt: { stringValue: now } } })
      });
    } catch (e) {}
  }
  return { pending: docs.length, companies: groups.length, plates: groups.reduce((n, g) => n + g.plates.length, 0), push };
}

// نص الإشعار: «📊 جديد في التقارير اليومية: 5 لوحات» + «تمويل +3، إجارة +2» + أول اللوحات
function reportPushText(groups) {
  const total = groups.reduce((n, g) => n + g.plates.length, 0);
  const per = groups.map((g) => g.co + ": " + g.plates.length).join("، ");
  const all = groups.flatMap((g) => g.plates);
  const show = all.slice(0, 6).join(" ، ");
  const more = all.length > 6 ? " … و" + (all.length - 6) + " كمان" : "";
  return {
    title: "📊 جديد في التقارير اليومية: " + total + (total === 1 ? " لوحة" : total <= 10 ? " لوحات" : " لوحة"),
    body: per + "\n" + show + more
  };
}

function needsOpenName(name) { return /^\s*(repo|crn|arn)/i.test(String(name || "")); }

// ملف اسمه رقم لوحة ("ابح 1234.pdf" أو "ABC 1234.pdf") -> اللوحة بالعربي، وإلا ""
function plateFromFileName(name) {
  // «د ص ط 5508 (1).pdf»: الـ (1) اللي بيضيفها الموبايل لما الملف يتنزل مرتين بنشيلها
  let n = String(name || "").replace(/\.pdf$/i, "").replace(/\s*\(\d{1,3}\)\s*$/, "").replace(/[_\-]+/g, " ").replace(/\s+/g, " ").trim();
  n = englishPlateToArabic(n);
  let m = n.match(/^([\u0621-\u064A])\s?([\u0621-\u064A])\s?([\u0621-\u064A])\s?(\d{1,4})$/);
  if (m) return m[1] + m[2] + m[3] + " " + m[4];
  // «9272 د ه أ.pdf»: الرقم الأول (زي ما التطبيق بيقراه بالظبط)
  m = n.match(/^(\d{1,4})\s?([\u0621-\u064A])\s?([\u0621-\u064A])\s?([\u0621-\u064A])$/);
  return m ? (m[2] + m[3] + m[4] + " " + m[1]) : "";
}

const AUTO_REPORT_MAX_AGE_MS = 24 * 3600 * 1000;   // التقرير اليومي: شهادات آخر 24 ساعة بس

/* ============================================================
   «الشهادات اللي منزلتش في الإحالات»: الأدمن الرئيسي محمّل ملف إحالة ولقى فيه لوحات من التقرير،
   التطبيق بيبعتها هنا والسيرفر بيشيلها من التقارير نهائي (عند كل الناس).
   - الأدمن الرئيسي بس (isAdmin)
   - بيشيل بس اللوحات اللي اتبعتت، من التقرير اليومي و«الشهادات اللي منزلتش» الاتنين
   - الإحالة = ملفات الإحالة + اللوحات الملصوقة في كارت البحث (زرار الفرز)
   - بيتسجّل مين شال إيه في reportRefPurgeLog
   ============================================================ */
function plateKeyLoose(s) {
  let n = String(s || "")
    .replace(/[٠-٩]/g, (d) => String(d.charCodeAt(0) - 0x0660))
    .replace(/[۰-۹]/g, (d) => String(d.charCodeAt(0) - 0x06F0))
    .replace(/[ً-ٰٟـ]/g, "")
    .replace(/[أإآ]/g, "ا").replace(/ة/g, "ه").replace(/ى/g, "ي")
    .replace(/[\s,،\-_.]+/g, "");
  const m = n.match(/^([^\d]+)(\d{1,4})$/);
  if (!m) return "";
  return m[1] + m[2].padStart(4, "0");
}

app.post("/api/reports/referral-hit", async (req, res) => {
  try {
    const uid = await uidFromReq(req);
    if (!uid) return res.status(401).json({ ok: false, error: "auth" });
    const me = await fsGetDoc("users/" + encodeURIComponent(uid));
    // الأدمن الرئيسي بس هو اللي إحالته بتمسح من التقارير عند الكل
    if (!me || fv(me, "isAdmin") !== true) return res.status(403).json({ ok: false, error: "admin only" });
    const sent = Array.isArray(req.body && req.body.plates) ? req.body.plates.slice(0, 3000) : [];
    const keys = new Set(sent.map(plateKeyLoose).filter(Boolean));
    if (!keys.size) return res.json({ ok: true, removed: 0 });

    const docs = await fsRunQuery({
      from: [{ collectionId: "announcements" }],
      where: { fieldFilter: { field: { fieldPath: "report" }, op: "EQUAL", value: { booleanValue: true } } }
    });
    let removed = 0, touched = 0;
    const removedPlates = [];
    for (const d of docs) {
      const id = String(d.name || "").split("/").pop();
      const t = Number(fv(d, "id") || id);
      if (!(isFinite(t) && t > 1e12)) continue;   // اليومي و«الشهادات اللي منزلتش» الاتنين
      const text = String(fv(d, "text") || "");
      const lines = text.split(/\r?\n/);
      const keep = lines.filter((l) => {
        const k = plateKeyLoose(l.trim());
        if (k && keys.has(k)) { removedPlates.push(l.trim()); return false; }
        return true;
      });
      if (keep.length === lines.length) continue;
      removed += lines.length - keep.length;
      touched++;
      const hdr = { Authorization: "Bearer " + (await gcpToken()), "Content-Type": "application/json" };
      const docUrl = FS_DOCS + "/announcements/" + encodeURIComponent(id);
      if (keep.some((l) => /\d/.test(l))) {
        await fetch(docUrl + "?updateMask.fieldPaths=text&currentDocument.exists=true", {
          method: "PATCH", headers: hdr, body: JSON.stringify({ fields: { text: { stringValue: keep.join("\n").trim() } } })
        });
      } else {
        await fetch(docUrl, { method: "DELETE", headers: hdr });
      }
    }
    if (removed) {
      const logId = Date.now() + "-" + uid.slice(0, 8);
      await fetch(FS_DOCS + "/reportRefPurgeLog/" + logId, {
        method: "PATCH",
        headers: { Authorization: "Bearer " + (await gcpToken()), "Content-Type": "application/json" },
        body: JSON.stringify({ fields: {
          uid: { stringValue: uid },
          email: { stringValue: String(fv(me, "email") || "") },
          plates: { arrayValue: { values: removedPlates.slice(0, 500).map((x) => ({ stringValue: x })) } },
          at: { timestampValue: new Date().toISOString() }
        } })
      }).catch(() => {});
    }
    res.json({ ok: true, removed, reports: touched });
  } catch (e) {
    res.status(500).json({ ok: false, error: String((e && e.message) || e).slice(0, 200) });
  }
});

// استرجاع التقارير التلقائية اللي اتمسحت بالغلط (مرة واحدة): POST /api/reports/rebuild-auto
// بالمفتاح السري. ?dry=1 بيعرض الأعداد بس من غير ما يكتب حاجة. ?since=ISO (الافتراضي بداية التقارير التلقائية)
// بيرجّع لوحات الشهادات (لوحة + شركة) من certMetaDay، ومابيكررش أي لوحة موجودة في تقرير دلوقتي. من غير إشعار.
app.post("/api/reports/rebuild-auto", async (req, res) => {
  if (!SERVER_KEY || !keyEq(req.headers["x-server-key"] || "", SERVER_KEY)) return res.status(403).json({ ok: false, error: "forbidden" });
  try {
    const dry = String(req.query.dry || "") === "1";
    const since = Date.parse(String(req.query.since || "2026-11-07T07:30:00Z"));
    if (!isFinite(since)) return res.status(400).json({ ok: false, error: "bad since" });
    const hours = Math.min(96, Math.ceil((Date.now() - since) / 3600000) + 1);
    const serverHdr = { "X-Server-Key": SERVER_KEY };
    const lr = await fetch(CERT_DOWNLOAD_BASE + "/api/certificates/recent?hours=" + hours + "&limit=3000", { headers: serverHdr });
    const ld = await lr.json().catch(() => null);
    if (!lr.ok || !ld || ld.ok === false) throw new Error("list failed: HTTP " + lr.status);
    const files = (Array.isArray(ld.files) ? ld.files : []).filter((f) => {
      const t = Date.parse(f && (f.createdTime || f.modifiedTime) || "");
      return f && f.id && isFinite(t) && t >= since;
    });
    const days = {};
    for (const f of files) { const d = dayOf(f); (days[d] = days[d] || []).push(f); }
    // اللوحات الموجودة في أي تقرير دلوقتي (يومي أو اللي منزلتش) -> مانكررهاش
    const have = new Set();
    const docs = await fsRunQuery({
      from: [{ collectionId: "announcements" }],
      where: { fieldFilter: { field: { fieldPath: "report" }, op: "EQUAL", value: { booleanValue: true } } }
    });
    for (const d of docs) String(fv(d, "text") || "").split(/\r?\n/).forEach((l) => { const k = plateKeyLoose(l.trim()); if (k) have.add(k); });
    const by = {};   // الشركة -> { plates: Map(key->plate), newest }
    let withMeta = 0, skippedHave = 0;
    for (const d of Object.keys(days)) {
      const st = await fsGetDay(d);
      for (const f of days[d]) {
        const e = st[f.id];
        if (!e || !e.p || !e.l) continue;
        withMeta++;
        const k = plateKeyLoose(e.p);
        if (!k) continue;
        if (have.has(k)) { skippedHave++; continue; }
        const co = reportCompanyOf(e.l);
        if (!co) continue;
        const g = by[co] = by[co] || { plates: new Map(), newest: 0 };
        g.plates.set(k, e.p);
        g.newest = Math.max(g.newest, Date.parse(f.createdTime || f.modifiedTime || "") || 0);
      }
    }
    const summary = Object.keys(by).map((co) => ({ company: co, plates: by[co].plates.size }));
    if (dry) return res.json({ ok: true, dry: true, files: files.length, withMeta, skippedHave, companies: summary });
    let made = 0;
    const base = Date.now();
    const cos = Object.keys(by);
    for (let i = 0; i < cos.length; i++) {
      const co = cos[i];
      const plates = Array.from(by[co].plates.values());
      if (!plates.length) continue;
      // وقت التقرير = وقت أحدث شهادة فيه (عشان مدة الـ24 ساعة تتحسب صح)، ومايزيدش عن دلوقتي
      const id = String(Math.min(base, Math.max(since, by[co].newest || base)) + i);
      const fields = {
        text: { stringValue: plates.join("\n") },
        id: { stringValue: id },
        showClose: { booleanValue: true }, showSort: { booleanValue: true }, masked: { booleanValue: false }, showNote: { booleanValue: true },
        report: { booleanValue: true }, auto: { booleanValue: true }, restored: { booleanValue: true },
        companies: { arrayValue: { values: [{ stringValue: co }] } },
        audience: { stringValue: "all" },
        targets: { arrayValue: { values: [] } }, targetNames: { arrayValue: { values: [] } },
        createdAt: { timestampValue: new Date().toISOString() }
      };
      const r = await fetch(FS_DOCS + "/announcements/" + id + "?currentDocument.exists=false", {
        method: "PATCH",
        headers: { Authorization: "Bearer " + (await gcpToken()), "Content-Type": "application/json" },
        body: JSON.stringify({ fields })
      });
      if (r.ok) made++;
      else console.log("[rebuild-auto] save HTTP " + r.status + " " + (await r.text()).slice(0, 150));
    }
    console.log("[rebuild-auto] " + JSON.stringify({ made, summary }));
    res.json({ ok: true, made, companies: summary, skippedHave });
  } catch (e) {
    res.status(500).json({ ok: false, error: String((e && e.message) || e).slice(0, 200) });
  }
});

// قفل بسيط في Firestore: جولة واحدة بس في نفس الوقت (حتى لو Cloud Run فتح نسختين)
async function autoLock(ms) {
  const url = FS_DOCS + "/serverState/autoRun";
  const hdr = { Authorization: "Bearer " + (await gcpToken()), "Content-Type": "application/json" };
  try {
    const g = await fetch(url, { headers: hdr });
    if (g.ok) { const d = await g.json(); const until = Number((((d.fields || {}).until || {}).integerValue) || 0); if (until > Date.now()) return false; }
  } catch (e) {}
  const r = await fetch(url + "?updateMask.fieldPaths=until", { method: "PATCH", headers: hdr, body: JSON.stringify({ fields: { until: { integerValue: String(Date.now() + ms) } } }) });
  return r.ok;
}
async function autoUnlock() {
  try {
    await fetch(FS_DOCS + "/serverState/autoRun?updateMask.fieldPaths=until", { method: "PATCH", headers: { Authorization: "Bearer " + (await gcpToken()), "Content-Type": "application/json" }, body: JSON.stringify({ fields: { until: { integerValue: "0" } } }) });
  } catch (e) {}
}

async function autoRunOnce() {
  if (!(await autoLock(AUTO_BUDGET_MS + 120000))) return { skipped: "locked" };
  try { return await autoRunBody(); } finally { await autoUnlock(); }
}

async function autoRunBody() {
  const t0 = Date.now();
  const serverHdr = { "X-Server-Key": SERVER_KEY };
  const lr = await fetch(CERT_DOWNLOAD_BASE + "/api/certificates/recent?hours=" + AUTO_HOURS + "&limit=3000", { headers: serverHdr });
  const ld = await lr.json().catch(() => null);
  if (!lr.ok || !ld || ld.ok === false) throw new Error("list failed: HTTP " + lr.status + " " + String((ld && ld.error) || ""));
  // كل ملفات PDF (زي زرار «استخراج» في التطبيق بالظبط):
  //  - REPO/CRN/ARN: اللوحة + اسم الشركة من جوه الشهادة
  //  - أي اسم تاني: الاسم نفسه هو رقم اللوحة، والشركة بتتقرا من جوه الشهادة
  const files = (Array.isArray(ld.files) ? ld.files : []).filter((f) => f && f.id && (needsOpenName(f.name) || /\.pdf$/i.test(String(f.name || ""))));

  const days = {};
  for (const f of files) { const d = dayOf(f); (days[d] = days[d] || []).push(f); }
  const state = {};
  for (const d of Object.keys(days)) state[d] = await fsGetDay(d);

  const todo = [];
  for (const d of Object.keys(days)) {
    for (const f of days[d]) {
      const e = state[d][f.id];
      // لوحة من اسم الملف (من غير ما نفتحه): لو اتقرت دلوقتي ومكانتش متسجلة -> ترجع تتعالج حتى لو المحاولات خلصت
      // (شهادات «9272 د ه أ.pdf» و«... (1).pdf» كانت بتتساب من غير لوحة)
      const fp = needsOpenName(f.name) ? "" : plateFromFileName(f.name);
      // اتقرت (لوحة + شركة) بس لسه ماتضافتش للتقرير (الجولة اللي فاتت اتقطعت قبل ما تضيفها)
      const fAge = Date.now() - (Date.parse(f.createdTime || f.modifiedTime || "") || Date.now());
      const pendingRep = !!(e && e.p && e.l && e.r === 0 && fAge <= AUTO_REPORT_MAX_AGE_MS);
      if (!e || ((!e.p || !e.l) && e.a < AUTO_MAX_ATTEMPTS) || (!e.p && fp) || pendingRep) todo.push({ f, d, prev: e || null, fp, pendingRep });
    }
  }
  todo.sort((x, y) => (Date.parse(y.f.createdTime || y.f.modifiedTime || "") || 0) - (Date.parse(x.f.createdTime || x.f.modifiedTime || "") || 0));
  const queue = isFinite(AUTO_MAX_PER_RUN) ? todo.slice(0, AUTO_MAX_PER_RUN) : todo;

  let processed = 0, plates = 0, lessors = 0, fresh = 0, freshHidden = 0;
  let repAdd = {};     // الشركة -> لوحات جديدة للتقرير اليومي (من آخر إضافة)
  let repIds = [];     // الشهادات اللي في repAdd: [{ d, id, e }] -> بعد ما تتضاف بنعلّمها r=1
  const reportsOut = [];
  let lastFlush = Date.now();
  async function flushReports() {
    if (!Object.keys(repAdd).length) return;
    const add = repAdd, ids = repIds;
    repAdd = {}; repIds = [];
    let rr = null;
    try { rr = await createAutoReports(add); } catch (e) { rr = { error: String((e && e.message) || e).slice(0, 200) }; }
    reportsOut.push(rr);
    if (rr && (rr.reports || rr.handled)) {
      const byDay = {};
      ids.forEach((x) => { (byDay[x.d] = byDay[x.d] || {})[x.id] = Object.assign({}, x.e, { r: 1 }); });
      for (const d of Object.keys(byDay)) { try { await fsSaveDay(d, byDay[d]); } catch (e) {} }
    }
    lastFlush = Date.now();
  }
  const CHUNK = AUTO_CONCURRENCY * 3;
  for (let i = 0; i < queue.length; i += CHUNK) {
    if (Date.now() - t0 > AUTO_BUDGET_MS) break;
    const chunk = queue.slice(i, i + CHUNK);
    const outs = await mapWithConcurrency(chunk, AUTO_CONCURRENCY, async (it) => {
      if (it.pendingRep) return { plate: it.prev.p, lessor: { name: it.prev.l, cr: it.prev.c || null } };
      if (needsOpenName(it.f.name)) return resolveOnePlate(it.f.id, "both", serverHdr);
      // اسم الشركة متقري قبل كده واللوحة من الاسم -> مش محتاجين نفتح الملف
      if (it.prev && it.prev.l && it.fp) return { plate: it.fp, lessor: { name: it.prev.l, cr: it.prev.c || null } };
      // اسم الملف هو اللوحة: نقرا اسم الشركة بس (أسرع)
      const r = await resolveOnePlate(it.f.id, "lessor", serverHdr);
      return Object.assign({}, r || {}, { plate: plateFromFileName(it.f.name) || (r && r.plate) || null });
    });
    const save = {};
    chunk.forEach((it, k) => {
      const r = outs[k];
      const prev = it.prev || { p: "", l: "", c: "", a: 0 };
      const p = (r && r.plate) || prev.p || "";
      const l = (r && r.lessor && r.lessor.name) || prev.l || "";
      const c = (r && r.lessor && r.lessor.cr) || prev.c || "";
      if (p) plates++;
      if (l) lessors++;
      if (!it.prev) fresh++;   // أول مرة تتشاف = شهادة جديدة
      if (!it.prev && !(p && l)) freshHidden++;   // جديدة ومادخلتش التقارير -> إشعار للأدمن
      // اللوحة والشركة اتقروا الاتنين ومادخلوش تقرير قبل كده -> التقرير اليومي (مرة واحدة بس)
      const fAge = Date.now() - (Date.parse(it.f.createdTime || it.f.modifiedTime || "") || Date.now());
      const reportedBefore = !!(it.prev && prev.p && prev.l && prev.r !== 0);
      let added = false;
      if (AUTO_REPORTS && p && l && !reportedBefore && fAge <= AUTO_REPORT_MAX_AGE_MS) {
        const co = reportCompanyOf(l);
        if (co) { (repAdd[co] = repAdd[co] || []).push(p); added = true; }
      }
      // r: 0 = مستنية تتضاف (بتتعلّم 1 بعد ما التقرير يتحفظ)، غير كده 1 (خلصت أو مش هتدخل تقرير)
      const rf = added ? 0 : (reportedBefore ? prev.r : (p && l ? 1 : 0));
      const entry = { p, l, c, a: it.pendingRep ? prev.a : prev.a + 1, r: rf };
      (save[it.d] = save[it.d] || {})[it.f.id] = entry;
      if (added) repIds.push({ d: it.d, id: it.f.id, e: entry });
      processed++;
    });
    for (const d of Object.keys(save)) await fsSaveDay(d, save[d]);
    if (Date.now() - lastFlush >= AUTO_FLUSH_MS) await flushReports();
  }
  await flushReports();
  // خلّصنا كل الشهادات اللي كانت مستنية؟ -> دلوقتي بس إشعار التقارير اليومية (مرة واحدة بكل الجديد)
  let repPush = null;
  if (AUTO_REPORTS && processed >= todo.length) {
    try { repPush = await pushPendingReports(); } catch (e) { repPush = { error: String((e && e.message) || e).slice(0, 200) }; }
  }
  let push = null;
  try { push = await pushNewCertsToAdmins(freshHidden); } catch (e) { push = { error: String((e && e.message) || e).slice(0, 200) }; }
  const reports = reportsOut.length ? { flushes: reportsOut.length, reports: reportsOut.reduce((n, x) => n + ((x && x.reports) || 0), 0), last: reportsOut[reportsOut.length - 1] } : { reports: 0 };
  const out = { listed: files.length, todo: todo.length, processed, left: Math.max(0, todo.length - processed), fresh, plates, lessors, push, reports, repPush, ms: Date.now() - t0 };
  console.log("[auto-run] " + JSON.stringify(out));
  return out;
}

/* ---------- قراية رقم الشاصي من صورة (تشيك الشاصي بالكاميرا) ----------
   التطبيق بيبعت صورة الملصق (JPEG base64) + Firebase token -> السيرفر بيبعتها لقارئ الصور
   ويرجّع رقم الشاصي (17 خانة) ومعاه هل خانة التأكيد (الخانة 9) مظبوطة.
   المفتاح في المتغير VISION_API_KEY على السيرفر بس -- عمره ما بيروح للتطبيق.
   الصورة مش بتتحفظ ولا بتتكتب في اللوج. */
const VISION_API_KEY = String(process.env.VISION_API_KEY || "").trim();
// الموديلات بالترتيب من الأرخص للأقوى: بنبدأ بالأول، ولو مالقاش رقم سليم بنطلع للي بعده
const VIN_MODELS = String(process.env.VIN_MODELS || "claude-haiku-5-5,claude-sonnet-5-5,claude-opus-5-5")
  .split(",").map((x) => x.trim()).filter(Boolean);
const VIN_PER_HOUR = Math.max(10, Number(process.env.VIN_PER_HOUR || 300));         // حد لكل حساب في الساعة
const vinUse = new Map();                                                           // uid -> [أوقات]
const VIN_MAP = { A:1,B:2,C:3,D:4,E:5,F:6,G:7,H:8,J:1,K:2,L:3,M:4,N:5,P:7,R:9,S:2,T:3,U:4,V:5,W:6,X:7,Y:8,Z:9 };
const VIN_W = [8,7,6,5,4,3,2,10,0,9,8,7,6,5,4,3,2];
function vinCheckOk(v) {
  if (!/^[A-HJ-NPR-Z0-9]{17}$/.test(v)) return false;
  let s = 0;
  for (let i = 0; i < 17; i++) { const c = v[i]; s += (/\d/.test(c) ? +c : VIN_MAP[c]) * VIN_W[i]; }
  const r = s % 11;
  return (r === 10 ? "X" : String(r)) === v[8];
}
function vinClean(t) {
  return String(t || "").toUpperCase().replace(/O/g, "0").replace(/Q/g, "0").replace(/I/g, "1").replace(/[^A-HJ-NPR-Z0-9]/g, "");
}
const vinNoThink = {};                                                              // موديل رفض إعداد التفكير -> مانبعتهوش تاني
async function visionReadVin(model, b64, mime) {
  // الموديلات الجديدة بتفكر قبل ما ترد، والتفكير بياكل من max_tokens (كانت 200 فالرد كان بيطلع فاضي
  // stop_reason=max_tokens). بنقفل التفكير ونسيب مساحة كفاية للرد.
  const think = !vinNoThink[model];
  const r = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "x-api-key": VISION_API_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json" },
    body: JSON.stringify(Object.assign({
      model,
      max_tokens: 1500,
      messages: [{ role: "user", content: [
        { type: "image", source: { type: "base64", media_type: mime, data: b64 } },
        { type: "text", text:
          "The photo shows a vehicle VIN (chassis number) label, usually seen through a dusty or reflective windshield; it may be rotated or upside down. " +
          "Read the VIN exactly, character by character (17 characters; VINs never contain I, O or Q). Ignore other numbers (glass codes, dates, barcodes, short codes under the VIN). " +
          "If a character is unreadable put ? in its place. If there are two different VINs, give both. " +
          "Reply with JSON only, no other text: {\"vins\":[\"...\"]}  -- or {\"vins\":[]} if there is no VIN in the photo." }
      ] }]
    }, think ? { thinking: { type: "disabled" } } : {}))
  });
  const txt = await r.text();
  if (!r.ok && r.status === 400 && think && /thinking/i.test(txt)) { vinNoThink[model] = true; return visionReadVin(model, b64, mime); }
  if (!r.ok) throw new Error("vision HTTP " + r.status + " " + txt.slice(0, 160));
  let d = null; try { d = JSON.parse(txt); } catch (e) {}
  const out = String(((d && d.content) || []).map((c) => c.text || "").join(" "));
  const m = /\{[\s\S]*\}/.exec(out);
  let j = null; try { j = m ? JSON.parse(m[0]) : null; } catch (e) {}
  const raw = (j && Array.isArray(j.vins)) ? j.vins : (out.match(/[A-Z0-9?]{11,17}/gi) || []);
  // لو مالقاش رقم: نكتب رد القارئ في اللوج (من غير الصورة) عشان نعرف ليه
  if (!raw.length) console.log("🔎 vin رد فاضي (" + model + ", " + ((d && d.stop_reason) || "") + "): " + out.replace(/\s+/g, " ").slice(0, 300));
  return raw.map((x) => String(x || ""));
}

// POST /api/vin/read  { image: "<base64 jpeg>", mime?: "image/jpeg" }  (أي مشترك اشتراكه شغال)
app.post("/api/vin/read", async (req, res) => {
  try {
    if (!VISION_API_KEY) return res.status(503).json({ ok: false, error: "off" });
    const uid = await uidFromReq(req);
    if (!uid) return res.status(401).json({ ok: false, error: "auth" });
    const me = await fsGetDoc("users/" + encodeURIComponent(uid));
    if (!me) return res.status(403).json({ ok: false, error: "user" });
    const isAdmin = fv(me, "isAdmin") === true;
    const exp = Date.parse(String(fv(me, "expiresAt") || ""));
    if (!isAdmin && !isNaN(exp) && exp < Date.now()) return res.status(403).json({ ok: false, error: "expired" });
    const now = Date.now();
    const used = (vinUse.get(uid) || []).filter((t) => now - t < 3600000);
    if (!isAdmin && used.length >= VIN_PER_HOUR) return res.status(429).json({ ok: false, error: "limit" });
    used.push(now); vinUse.set(uid, used);
    const b64 = String((req.body && req.body.image) || "").replace(/^data:[^,]*,/, "");
    const mime = /^image\/(jpeg|png|webp)$/.test(String(req.body && req.body.mime)) ? req.body.mime : "image/jpeg";
    if (b64.length < 1000 || b64.length > 7_500_000 || !/^[A-Za-z0-9+/=\s]+$/.test(b64.slice(0, 2000))) return res.status(400).json({ ok: false, error: "image" });
    const t0 = Date.now();
    // من الأرخص للأقوى: أول ما موديل يلاقي رقم سليم (17 خانة وخانة التأكيد صح) بنقف
    const goodIn = (arr) => arr.some((x) => { const v = vinClean(String(x || "").replace(/\?/g, "")); return v.length === 17 && vinCheckOk(v); });
    let raw = [], model = "", lastErr = null;
    for (const m of VIN_MODELS) {
      try { raw = (await visionReadVin(m, b64, mime)).concat(raw); model = m; }
      catch (e) { lastErr = e; console.log("⚠️ vin " + m + ": " + String((e && e.message) || e).slice(0, 160)); }
      if (goodIn(raw)) break;
    }
    if (!model && lastErr) throw lastErr;
    const seen = new Set();
    const vins = [];
    for (const x of raw) {
      const unk = /\?/.test(x);
      const v = vinClean(x.replace(/\?/g, ""));
      if (!v || v.length < 6 || seen.has(v)) continue;
      seen.add(v);
      vins.push({ v, ok: v.length === 17 && vinCheckOk(v), unk });
    }
    vins.sort((a, b) => (b.ok - a.ok) || (b.v.length - a.v.length));
    // من غير الرقم نفسه في اللوج -- بس النتيجة والوقت
    console.log("🔎 vin [" + uid.slice(0, 6) + "] " + model + ": " + vins.length + " رقم" + (vins[0] ? (vins[0].ok ? " ✓" : " ✗") : "") + " في " + (Date.now() - t0) + "ms");
    res.json({ ok: true, vins });
  } catch (e) {
    console.log("⚠️ vin: " + String((e && e.message) || e).slice(0, 200));
    res.status(502).json({ ok: false, error: "read" });
  }
});

// تجربة إشعار الآيفون: POST /api/push/test بنفس المفتاح السري (SERVER_KEY)
app.post("/api/push/test", async (req, res) => {
  if (!SERVER_KEY || !keyEq(req.headers["x-server-key"] || "", SERVER_KEY)) return res.status(403).json({ ok: false, error: "forbidden" });
  try { res.json(Object.assign({ ok: true }, await pushNewCertsToAdmins(1))); }
  catch (e) { res.status(500).json({ ok: false, error: String((e && e.message) || e).slice(0, 300) }); }
});

app.post("/api/certificates/auto-run", async (req, res) => {
  if (!SERVER_KEY || !keyEq(req.headers["x-server-key"] || "", SERVER_KEY)) return res.status(403).json({ ok: false, error: "forbidden" });
  if (autoRunning) return res.json({ ok: true, skipped: "already running" });
  autoRunning = true;
  try {
    res.json(Object.assign({ ok: true }, await autoRunOnce()));
  } catch (e) {
    console.log("[auto-run] error: " + String((e && e.message) || e).slice(0, 300));
    res.status(500).json({ ok: false, error: String((e && e.message) || e).slice(0, 300) });
  } finally {
    autoRunning = false;
  }
});

// تشخيص: افتح الرابط ده في المتصفح (حط id شهادة) وابعتلي الناتج لو أسماء الشركات مش طالعة
// GET /api/certificates/debug-lessor?id=DRIVE_FILE_ID
app.get("/api/certificates/debug-lessor", async (req, res) => {
  try {
    const id = String(req.query.id || "").trim();
    if (!id) return res.status(400).json({ ok: false, error: "id مطلوب" });
    const authHeader = req.headers.authorization || "";
    const r = await fetch(CERT_DOWNLOAD_BASE + "/api/certificates/" + encodeURIComponent(id) + "/download", {
      headers: authHeader ? { Authorization: authHeader } : {}
    });
    if (!r.ok) return res.status(502).json({ ok: false, error: "HTTP " + r.status });
    const pdf = await getDocumentProxy(new Uint8Array(await r.arrayBuffer()));
    const { text } = await extractText(pdf, { mergePages: true });
    const loc = await locateLessor(pdf);
    const bandItems = loc.band
      ? loc.items.filter((it) => it.transform[5] <= loc.band.yTop + 6 && it.transform[5] >= loc.band.yBottom - 6)
          .map((it) => ({ str: it.str, x: Math.round(it.transform[4]), y: Math.round(it.transform[5]) }))
      : [];
    const lessor = await resolveLessor(pdf, text || "", id);
    res.json({ ok: true, kind: loc.kind, cr: loc.cr, band: loc.band, bandItems, lessor });
  } catch (e) {
    res.status(500).json({ ok: false, error: String((e && e.message) || e) });
  }
});

app.post("/convert", upload.single("file"), async (req, res) => {
  let buffer = req.file ? req.file.buffer : null;
  try {
    if (!buffer) return res.status(400).json({ error: "لا يوجد ملف مرفق" });

    const kind = String((req.body && req.body.kind) || "xlsx").toLowerCase();

    let wb;
    if (kind === "csv" || kind === "txt") {
      const text = buffer.toString("utf-8");
      wb = XLSX.read(text, { type: "string", raw: true });
    } else {
      wb = XLSX.read(buffer, {
        type: "buffer",
        raw: true,
        dense: true,
        cellDates: false,
        cellStyles: false,
        cellNF: false,
        cellHTML: false,
        cellFormula: true,
        bookVBA: kind === "xlsm",
        bookDeps: false,
        sheetStubs: false
      });
    }

    if (!wb || !wb.SheetNames || !wb.SheetNames.length) {
      return res.status(400).json({ error: "الملف فاضي أو غير مدعوم" });
    }

    wb.Workbook = wb.Workbook || { Views: [{ RTL: true }] };
    const out = XLSX.write(wb, {
      bookType: "xlsb",
      type: "buffer",
      compression: true,
      cellStyles: false,
      bookSST: false,
      bookVBA: kind === "xlsm"
    });

    res.setHeader("Content-Type", "application/vnd.ms-excel.sheet.binary.macroEnabled.12");
    res.setHeader("Content-Disposition", 'attachment; filename="converted.xlsb"');
    res.send(out);
  } catch (e) {
    res.status(500).json({ error: String((e && e.message) || e) });
  } finally {
    // نفضّي المرجع الوحيد لمحتوى الملف فور انتهاء الطلب.
    buffer = null;
    if (req.file) req.file.buffer = null;
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("OSOS XLSB converter listening on port " + PORT);
});
