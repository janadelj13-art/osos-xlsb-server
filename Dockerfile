# صورة Node خفيفة (Debian) — متوافقة مع @napi-rs/canvas و tesseract.js
FROM node:22-slim

ENV NODE_ENV=production

# خطوط أساسية لو الـ PDF محتاج خط بديل وقت الرسم قبل الـ OCR
RUN apt-get update \
 && apt-get install -y --no-install-recommends fontconfig fonts-dejavu-core \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# نثبّت المكتبات الأول (عشان البناء يتخزّن لو الكود بس هو اللي اتغيّر)
COPY package*.json ./
RUN npm install --omit=dev && npm cache clean --force

COPY . .

# Cloud Run بيحط PORT=8080 لوحده، والسيرفر أصلًا بيقراه من process.env.PORT
EXPOSE 8080

USER node
CMD ["node", "server.js"]
