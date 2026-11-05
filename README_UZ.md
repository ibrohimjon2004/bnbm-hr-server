# BNBM Central Asia HR — Telegram Mini App + API

## Muhim
Bu paket avvalgi `bot.py` va SQLite `vacancies`, `extra_questions`, `applications` jadvallari bilan mos kelishga mo‘ljallangan. Ishga tushirishdan oldin ma'lumotlar bazangizning zaxira nusxasini oling. Fayllarni almashtirishda eski `.env` va `bnbm_hr.sqlite3` ni o‘chirmang.

**Telegram Mini App:** `https://t.me/bnbmhr_bot/bnbm_hr` — BotFather orqali berilgan URL hozir Render STATIC SITE ga ulangan. Bu paketdagi `static/index.html` API talab qiladi, shuning uchun uni yolg‘iz eski statik hostingga yuklasangiz ishlamaydi. Mini App URL'ni **yangi Python web-service HTTPS manziliga** almashtirish kerak. Bu loyihani bitta Python web-service sifatida ishga tushiring; statik frontend ham aynan shu serverdan chiqadi.

## Windowsda lokal ishga tushirish
1. ZIP arxivni alohida papkaga chiqaring.
2. Eski bazangizni (ixtiyoriy, agar mavjud bo‘lsa) shu papkaga NUSXALANG, aslini o‘chirmang.
3. `.env.example` faylini `.env` deb nusxalang, `BOT_TOKEN` va `ADMIN_ID` ni kiriting. **Bu faylni hech qachon GitHub'ga yuklamang.**
4. CMD: `py -m pip install -r requirements.txt`
5. CMD: `py -m uvicorn server:app --host 127.0.0.1 --port 8000`
6. Brauzer: `http://127.0.0.1:8000` — ko‘rinadi, lekin Telegram autentifikatsiyasi bo‘lmagani uchun anketa ishlamaydi. Haqiqiy test uchun HTTPS va Telegram ichida ochish kerak.
7. Botni alohida CMD oynasida ishga tushirish mumkin: `py bot.py`. Biroq eski bot `.env` bilan ishlaydi va Telegram polling jarayoni kompyuter yoqilganida ishlaydi.

## Render joylashtirish
1. Kodni **alohida GitHub repository** ga yuklang. `.env`, `.sqlite3` va nomzodlar fayllarini yuklamang.
2. Render → New → Web Service → repository tanlang. Runtime: Python; Build command: `pip install -r requirements.txt`; Start command: `uvicorn server:app --host 0.0.0.0 --port $PORT`.
3. Environment Variables: `BOT_TOKEN`, `ADMIN_ID`, `HR_IDS` (ixtiyoriy). Bot tokenini faqat Render Environment bo‘limiga kiriting.
4. Render bergan yangi HTTPS manzilni BotFather orqali Mini App URL sifatida almashtiring.
5. **DIQQAT:** Render free web service diskida SQLite ma'lumotlari restart/deployda yo‘qolishi mumkin. Shuning uchun bu konfiguratsiya faqat test uchun. Real nomzod ma'lumotlari uchun persistent PostgreSQL va xavfsiz fotosurat saqlash tizimi kerak; bu paket hali PostgreSQL bilan ishlamaydi. Free web service uxlab qolishi mumkin.
6. Render Web Service `server.py` ni ishga tushiradi, **`bot.py` pollingni avtomatik ishga tushirmaydi**. Mini App arizalarni server orqali qabul qilib Telegram API bilan PDF yuboradi; botning eski chat menyusi ishlashi uchun pollingni alohida jarayonda ishga tushirish kerak.

## Tizim
- Til: UZ / 中文; bo‘limlar: Savdo, Ishlab chiqarish.
- HR: faqat `ADMIN_ID` va `HR_IDS` da ko‘rsatilgan Telegram ID'lar; vakansiya yaratish/tahrirlash/yopish, qo‘shimcha savol, ariza holati, Excel, xitoycha PDF.
- Nomzod: 12 asosiy savol, kran uchun 2 qo‘shimcha, HR qo‘shgan savollar, fotosurat (5 MB gacha), rozilik.
- Telegram Mini App `initData` serverda HMAC bilan tekshiriladi. Oddiy brauzerda anketa topshirish mumkin emas.
- PDF'ga asosiy savollar xitoycha yoziladi; erkin matn avtomatik tarjima qilinmaydi, asl tilida saqlanadi.
- Zavod rasmi, aniq lokatsiya va HR aloqa ma'lumotlari berilmaganligi sababli **kiritilmagan**. Ularni tasdiqlangan ma'lumot bilan qo‘shish kerak.
- Xavfsizlik va ma'lumotlar bazasini production darajasida audit qilish kerak. 100% xatosizlik kafolati yo‘q.
