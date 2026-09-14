# سيرفر تحويل XLSB لموقع OS OS

سيرفر بسيط بيستقبل ملف إكسل/CSV كبير ويرجّعه محوّل لصيغة XLSB بسرعة أعلى من
التحويل على المتصفح. الملف بيتعالج في الذاكرة فقط ومفيش أي تخزين دائم له.

## 1) تجربته على جهازك أولًا (اختياري)

```
cd server
npm install
npm start
```

هيشتغل على `http://localhost:3000`. لو فتحت الرابط في المتصفح المفروض تشوف:
`OSOS XLSB converter server is running.`

## 2) نشره مجانًا على Render (بدون بطاقة ائتمان)

1. اعمل حساب على https://render.com (تسجيل بالإيميل أو GitHub).
2. ارفع مجلد `server` ده على مستودع GitHub خاص بيك (أو استخدم خيار
   "Deploy from a public Git repository" لو حاطط الكود على GitHub).
3. من لوحة Render: New > Web Service > اختار المستودع.
4. الإعدادات:
   - Environment: **Node**
   - Build Command: `npm install`
   - Start Command: `npm start`
   - Instance Type: **Free**
5. اضغط Create Web Service وانتظر لحد ما يخلص Deploy.
6. Render هيديك رابط شكله زي كده:
   `https://osos-xlsb-server.onrender.com`

## 3) ربطه بموقع OS OS

افتح `index.html` ودوّر على السطر ده (قريب من دالة compressFromFile):

```js
window.OSOS_SERVER_URL = window.OSOS_SERVER_URL || "";
```

واستبدلها بـ:

```js
window.OSOS_SERVER_URL = window.OSOS_SERVER_URL || "https://osos-xlsb-server.onrender.com";
```

كده أي ملف أكبر من 5 ميجا هيتحول عن طريق السيرفر تلقائيًا، والملفات الأصغر
هتفضل تتحول محليًا زي ما هي (أسرع أصلًا للملفات الصغيرة).

## ملاحظات مهمة

- **الخصوصية**: الملف بيتبعت للسيرفر مؤقتًا وقت التحويل بس، ومفيش أي حفظ
  دائم أو Logs لمحتواه. لكن برضو هو بيعدي على سيرفر خارج جهاز المستخدم،
  فلو محتاج خصوصية كاملة بدون استثناء، سيب `OSOS_SERVER_URL` فاضي.
- **Free tier في Render بينام بعد فترة عدم استخدام** (Cold Start)، فأول طلب
  بعد فترة راحة ممكن ياخد 30-60 ثانية لحد ما السيرفر يصحى. الطلبات اللي
  بعد كده هتكون سريعة عادي.
- لو حابب تقفل السيرفر على دومين موقعك بس (بدل ما يكون مفتوح للكل)، غيّر
  السطر ده في `server.js`:
  ```js
  app.use(cors({ origin: "*" }));
  ```
  إلى:
  ```js
  app.use(cors({ origin: "https://your-site-domain.com" }));
  ```
