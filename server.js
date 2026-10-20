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
// تحويل حروف اللوحة الإنجليزي للعربي (جدول اللوح السعودية الرسمي)
const EN_TO_AR_PLATE = { A: "ا", B: "ب", J: "ح", D: "د", R: "ر", S: "س", X: "ص", T: "ط", E: "ع", G: "ق", K: "ك", L: "ل", Z: "م", N: "ن", H: "ه", U: "و", V: "ى" };
function englishPlateToArabic(plate) {
  const m = String(plate || "").trim().match(/^([A-Za-z]{3})\s*(\d{1,4})$/);
  if (!m) return plate;
  const letters = m[1].toUpperCase().split("");
  if (!letters.every((c) => EN_TO_AR_PLATE[c])) return plate; // حرف مش في جدول اللوح السعودية: نسيبها زي ما هي
  return letters.map((c) => EN_TO_AR_PLATE[c]).join("") + " " + m[2];
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
async function plateFromSijilItems(pdf) {
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
  // 1) العربي: أرقام + أشكال حروف (بالترتيب المرئي، فبنقلب الحروف عشان تبقى من اليمين)
  const arVal = valueOf(arLabel);
  if (arVal) {
    const digits = (arVal.match(/\d/g) || []).join("");
    const letters = [];
    let bad = false;
    for (const ch of arVal) {
      const cp = ch.codePointAt(0);
      if (cp < 128) { if (!/[\d\s\-]/.test(ch)) bad = true; continue; }
      if (SIJIL_PLATE_GID[cp]) letters.push(SIJIL_PLATE_GID[cp]); else bad = true;
    }
    if (!bad && letters.length === 3 && digits.length >= 1 && digits.length <= 4) return letters.reverse().join("") + " " + digits;
  }
  // 2) الإنجليزي (بأي مسافات) → نحوله عربي
  const enVal = valueOf(enLabel);
  if (enVal) {
    const v = enVal.replace(/[\s\-]+/g, "").toUpperCase();
    let m = v.match(/^([A-Z]{3})(\d{1,4})$/);
    if (m) return englishPlateToArabic(m[1] + " " + m[2]);
    m = v.match(/^(\d{1,4})([A-Z]{3})$/);
    if (m) return englishPlateToArabic(m[2] + " " + m[1]);
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
async function resolveLessor(pdf, text, id) {
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
    if (wantPlate) { try { result = await plateFromSijilItems(pdf); } catch (e) { result = null; } }
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
const AUTO_MAX_PER_RUN = 120;   // أقصى عدد شهادات في الجولة الواحدة (الباقي في الجولة اللي بعدها)
const AUTO_BUDGET_MS = 200000;  // نوقف قبل مهلة الطلب
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
      a: Number((f.a && f.a.integerValue) || 0)
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
      a: { integerValue: String(e.a || 0) }
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

async function autoRunOnce() {
  const t0 = Date.now();
  const serverHdr = { "X-Server-Key": SERVER_KEY };
  const lr = await fetch(CERT_DOWNLOAD_BASE + "/api/certificates/recent?hours=" + AUTO_HOURS + "&limit=3000", { headers: serverHdr });
  const ld = await lr.json().catch(() => null);
  if (!lr.ok || !ld || ld.ok === false) throw new Error("list failed: HTTP " + lr.status + " " + String((ld && ld.error) || ""));
  const files = (Array.isArray(ld.files) ? ld.files : []).filter((f) => f && f.id && /^\s*(repo|crn|arn)/i.test(String(f.name || "")));

  const days = {};
  for (const f of files) { const d = dayOf(f); (days[d] = days[d] || []).push(f); }
  const state = {};
  for (const d of Object.keys(days)) state[d] = await fsGetDay(d);

  const todo = [];
  for (const d of Object.keys(days)) {
    for (const f of days[d]) {
      const e = state[d][f.id];
      if (!e || ((!e.p || !e.l) && e.a < AUTO_MAX_ATTEMPTS)) todo.push({ f, d, prev: e || null });
    }
  }
  todo.sort((x, y) => (Date.parse(y.f.createdTime || y.f.modifiedTime || "") || 0) - (Date.parse(x.f.createdTime || x.f.modifiedTime || "") || 0));
  const queue = todo.slice(0, AUTO_MAX_PER_RUN);

  let processed = 0, plates = 0, lessors = 0;
  const CHUNK = 9;
  for (let i = 0; i < queue.length; i += CHUNK) {
    if (Date.now() - t0 > AUTO_BUDGET_MS) break;
    const chunk = queue.slice(i, i + CHUNK);
    const outs = await mapWithConcurrency(chunk, CERT_RESOLVE_CONCURRENCY, (it) => resolveOnePlate(it.f.id, "both", serverHdr));
    const save = {};
    chunk.forEach((it, k) => {
      const r = outs[k];
      const prev = it.prev || { p: "", l: "", c: "", a: 0 };
      const p = (r && r.plate) || prev.p || "";
      const l = (r && r.lessor && r.lessor.name) || prev.l || "";
      const c = (r && r.lessor && r.lessor.cr) || prev.c || "";
      if (p) plates++;
      if (l) lessors++;
      (save[it.d] = save[it.d] || {})[it.f.id] = { p, l, c, a: prev.a + 1 };
      processed++;
    });
    for (const d of Object.keys(save)) await fsSaveDay(d, save[d]);
  }
  const out = { listed: files.length, todo: todo.length, processed, plates, lessors, ms: Date.now() - t0 };
  console.log("[auto-run] " + JSON.stringify(out));
  return out;
}

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
