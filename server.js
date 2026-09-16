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
const { extractText, getDocumentProxy, renderPageAsImage } = require("unpdf");
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
 * — من غير ما نحتاج أي توكنز Google جديدة هنا خالص — وبعدين نستخرج النص
 * ونلاقي فيه رقم اللوحة، ونرجّع النتيجة الجاهزة.
 * ================================================================= */

const CERT_DOWNLOAD_BASE = "https://osos-certificates.ososapp.workers.dev";
const CERT_RESOLVE_MAX_IDS = 20; // سقف أمان لعدد الملفات في الطلب الواحد (كان 30 — قللناه)
const CERT_RESOLVE_CONCURRENCY = 3; // كام ملف بيتفتح بالتوازي جوه نفس الطلب (كان 6 — قللناه عشان Render Free)

// فاصل مسموح بين حروف اللوحة أو بينها وبين الأرقام: مسافة أو شرطة (كان بس مسافة قبل كده)
const SEP = "[\\s\\-]*";

// نفس ترتيب الأنماط اللي كانت شغالة في المتصفح، منقولة هنا بالظبط عشان النتيجة متطابقة
// (دي طريقة احتياطية، بتتفتش لو الطريقة الأدق اللي تحت (بالعنوان) مالقتش حاجة)
const CERT_PLATE_PATTERNS = [
  { re: new RegExp("((?:[\\u0621-\\u064A]" + SEP + "){3})(?![\\u0621-\\u064A])" + SEP + "(?:^|[^\\d])(\\d{4})(?!\\d)"), lettersFirst: true },
  { re: new RegExp("(?:^|[^\\d])(\\d{4})(?!\\d)" + SEP + "((?:[\\u0621-\\u064A]" + SEP + "){3})(?![\\u0621-\\u064A])"), lettersFirst: false },
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
  const lettersRe = new RegExp("(?:[" + letterClass + "]" + SEP + "){3}(?![" + letterClass + "])");
  const lettersMatch = windowText.match(lettersRe);
  const digitsMatch = windowText.match(/(\d{4})(?!\d)/);
  if (!lettersMatch || !digitsMatch) return null;
  const letters = lettersMatch[0].replace(/[\s\-]+/g, "");
  return letters + " " + digitsMatch[1];
}

function extractPlateFromText(text) {
  // 1) جرب العنوان العربي الصريح (ولازم نوقف قبل عنوان الإنجليزي لو الحقل العربي فاضي)
  const ar = grabPlateAfterLabel(text, AR_PLATE_LABEL, "\\u0621-\\u064A", EN_PLATE_LABEL);
  if (ar) return ar;
  // 2) لو مفيش، جرب العنوان الإنجليزي الصريح
  const en = grabPlateAfterLabel(text, EN_PLATE_LABEL, "A-Za-z");
  if (en) return en;
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
const OCR_RENDER_SCALE = 2.5; // كل ما زاد، وضحت الصورة أكتر لكن استخرج أبطأ وأتقل على الرام
const OCR_TIMEOUT_MS = 45000; // 45 ثانية أقصى حد لعملية التصوير+القراءة الواحدة

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

async function ocrArabicPlate(pdf) {
  return runInOcrQueue(async () => {
    const worker = await getOcrWorker();
    const imageBuffer = await renderPageAsImage(pdf, 1, {
      canvasImport: () => import("@napi-rs/canvas"),
      scale: OCR_RENDER_SCALE
    });
    const { data: { text } } = await worker.recognize(Buffer.from(imageBuffer));
    return text ? extractPlateFromText(text) : null;
  });
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

async function resolveOnePlateRaw(id) {
  const controller = new AbortController();
  const abortTimer = setTimeout(() => controller.abort(), CERT_RESOLVE_PER_ITEM_TIMEOUT_MS);
  let pdf = null;
  try {
    const res = await fetch(CERT_DOWNLOAD_BASE + "/api/certificates/" + encodeURIComponent(id) + "/download", { signal: controller.signal });
    if (!res.ok) return null;
    const buf = await res.arrayBuffer();
    pdf = await getDocumentProxy(new Uint8Array(buf));
    const { text } = await extractText(pdf, { mergePages: true });
    let result = text ? extractPlateFromText(text) : null;
    if (!result) {
      // ملاذ أخير: الفونت العربي في الشهادة ده مش قابل للقراءة كنص (مشكلة معروفة
      // في الفونت نفسه)، فنحول الصفحة لصورة ونقراها بالعربي (OCR) بدل النص
      result = await withTimeout(ocrArabicPlate(pdf), OCR_TIMEOUT_MS, null);
    }
    return result;
  } finally {
    clearTimeout(abortTimer);
  }
}

// طبقة حماية إضافية: حتى لو الـ fetch نجح بس استخراج النص نفسه علّق (ملف تالف مثلًا)،
// بعد المهلة القصوى (نص عادي + OCR لو احتاج) بنعتبره فشل ونرجع null بدل ما نستنى للأبد
// المهلة هنا لازم تستحمل وقت الـ OCR كمان (لو احتجناه) مش بس وقت التنزيل والنص العادي
async function resolveOnePlate(id) {
  try {
    return await withTimeout(resolveOnePlateRaw(id), CERT_RESOLVE_PER_ITEM_TIMEOUT_MS + OCR_TIMEOUT_MS, null);
  } catch (e) {
    return null;
  }
}

// جسم الطلب: { ids: ["driveFileId1", "driveFileId2", ...] }
// الرد: { ok:true, results: { "driveFileId1": "AJS 9496" | null, ... } }
app.post("/api/certificates/resolve-plates", async (req, res) => {
  try {
    const ids = Array.isArray(req.body && req.body.ids) ? req.body.ids : [];
    const cleanIds = Array.from(new Set(ids.filter((x) => typeof x === "string" && x.trim()))).slice(0, CERT_RESOLVE_MAX_IDS);
    if (!cleanIds.length) return res.status(400).json({ ok: false, error: "لا يوجد ids" });

    const results = {};
    await mapWithConcurrency(cleanIds, CERT_RESOLVE_CONCURRENCY, async (id) => {
      try { results[id] = await resolveOnePlate(id); }
      catch (e) { results[id] = null; }
    });

    res.json({ ok: true, results });
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
