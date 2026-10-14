# نقل سيرفر OSOS إلى Google Cloud Run

ضيف الملفين `Dockerfile` و `.dockerignore` جنب `server.js` و `package.json` في نفس الـ repo (osos-xlsb).
ملف `server.js` ما بيتغيّرش.

## الطريقة السهلة (من المتصفح، مربوطة بـ GitHub)

1. ادخل https://console.cloud.google.com وأنشئ Project جديد (مثلًا `osos`).
2. من Billing: اربط حساب فواتير (محتاج كارت بنكي حتى لو مش هتدفع).
   وبعدها من Budgets & alerts اعمل تنبيه بقيمة صغيرة (مثلًا 5 دولار).
3. افتح **Cloud Run** ← **Create service** ← **Continuously deploy from a repository** ← **Set up with Cloud Build**.
   - اربط حساب GitHub واختار repo `osos-xlsb` والفرع `main`.
   - Build type: **Dockerfile** (المسار `/Dockerfile`).
4. إعدادات الخدمة:
   - Service name: `osos-xlsb-server`
   - Region: `us-central1`
   - Authentication: **Allow unauthenticated invocations** (لازم، لأن التطبيق بيكلم السيرفر مباشرة)
   - Billing: **Request-based** (الافتراضي)
   - CPU: 1 | Memory: **1 GiB**
   - Request timeout: 300 ثانية | Maximum concurrent requests: 4
   - Minimum instances: **0** | Maximum instances: **2** (ده اللي بيحدد سقف الفاتورة)
5. اضغط **Create** وانتظر البناء (حوالي 3–6 دقايق أول مرة).
6. هيطلع لك رابط شكله `https://osos-xlsb-server-xxxxx.us-central1.run.app`.
   افتحه في المتصفح، لازم تشوف: `OSOS XLSB converter server is running.`
7. ابعت الرابط لـ Claude عشان يحطه في `index.html`.
8. بعد ما تتأكد إن الشهادات والتحويل شغالين، شيل خدمة Render.

بعد كده أي Commit على فرع `main` بيتنشر تلقائيًا.

## بديل: من الطرفية (لو عندك gcloud)

```
gcloud run deploy osos-xlsb-server --source . --region us-central1 \
  --allow-unauthenticated --memory 1Gi --cpu 1 \
  --concurrency 4 --timeout 300 --min-instances 0 --max-instances 2
```
