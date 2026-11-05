import asyncio
import logging
import os
from dotenv import load_dotenv
load_dotenv()
import sqlite3
from datetime import datetime
from io import BytesIO

from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart, Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    Message, CallbackQuery, ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton, BufferedInputFile
)
from aiogram.utils.keyboard import InlineKeyboardBuilder
from openpyxl import Workbook

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))
DB_PATH = os.getenv("DB_PATH", "bnbm_hr.sqlite3")

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN topilmadi. .env faylini sozlang.")

logging.basicConfig(level=logging.INFO)
bot = Bot(BOT_TOKEN)
dp = Dispatcher()


def db():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    with db() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS vacancies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            department TEXT NOT NULL,
            title_uz TEXT NOT NULL,
            title_zh TEXT NOT NULL,
            description_uz TEXT DEFAULT '',
            description_zh TEXT DEFAULT '',
            active INTEGER DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS extra_questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vacancy_id INTEGER NOT NULL,
            question_uz TEXT NOT NULL,
            question_zh TEXT NOT NULL,
            FOREIGN KEY(vacancy_id) REFERENCES vacancies(id)
        );
        CREATE TABLE IF NOT EXISTS applications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            telegram_id INTEGER NOT NULL,
            language TEXT NOT NULL,
            vacancy_id INTEGER NOT NULL,
            answers TEXT DEFAULT '',
            photo_file_id TEXT DEFAULT '',
            status TEXT DEFAULT 'Yangi',
            created_at TEXT NOT NULL
        );
        """)
        count = con.execute("SELECT COUNT(*) FROM vacancies").fetchone()[0]
        if count == 0:
            con.executemany("""
                INSERT INTO vacancies
                (department,title_uz,title_zh,description_uz,description_zh)
                VALUES (?,?,?,?,?)
            """, [
                ("Ishlab chiqarish", "Kran haydovchisi", "起重机司机",
                 "Vazifalar: yuklarni ko‘tarish, tushirish va ko‘chirish; kranni xavfsiz boshqarish; texnik holatini tekshirish. Talablar: kran boshqarish tajribasi, mas’uliyat va texnika xavfsizligiga rioya qilish. Tegishli guvohnoma bo‘lsa, afzallik beradi. Ish tartibi: 12 soatlik smena. 3 mahal bepul ovqat. Ish haqi suhbatda kelishiladi.",
                 "工作职责：负责货物吊装、卸载和转运；安全操作起重机；检查设备技术状况。任职要求：具有起重机操作经验，工作认真负责并遵守安全规定。持有相关操作证者优先。工作安排：12小时轮班，提供每日三餐。薪资面议。"),
                ("Savdo", "Savdo mutaxassisi", "销售专员",
                 "Savdo bo‘yicha mijozlar bilan ishlash. Batafsil ma’lumot HR menejeridan olinadi.",
                 "负责销售及客户沟通。详细信息请咨询人力资源经理。")
            ])


class Apply(StatesGroup):
    answering = State()
    photo = State()


class AddVacancy(StatesGroup):
    department = State()
    title_uz = State()
    title_zh = State()
    desc_uz = State()
    desc_zh = State()


class AddQuestion(StatesGroup):
    vacancy_id = State()
    question_uz = State()
    question_zh = State()


QUESTIONS = [
    {
        "key": "gender",
        "uz": "1. Jinsingizni tanlang:",
        "zh": "1. 请选择您的性别：",
        "options": [("Erkak", "男"), ("Ayol", "女")]
    },
    {"key": "first_name", "uz": "2. Pasportingiz bo‘yicha ismingizni kiriting.",
     "zh": "2. 请输入您护照上的名字。"},
    {"key": "last_name", "uz": "3. Pasportingiz bo‘yicha familiyangizni kiriting.",
     "zh": "3. 请输入您护照上的姓氏。"},
    {"key": "birth_date", "uz": "4. Tug‘ilgan sanangizni kiriting (kun.oy.yil). Masalan: 18.03.1995",
     "zh": "4. 请输入出生日期（日.月.年），例如：18.03.1995"},
    {
        "key": "marital",
        "uz": "5. Oilaviy holatingizni tanlang:",
        "zh": "5. 请选择您的婚姻状况：",
        "options": [("Uylangan / Turmushga chiqqan", "已婚"),
                    ("Bo‘ydoq / Turmushga chiqmagan", "未婚"),
                    ("Ajrashgan", "离异"), ("Beva", "丧偶")]
    },
    {"key": "address", "uz": "6. Turar joy manzilingizni kiriting (viloyat, tuman, ko‘cha).",
     "zh": "6. 请输入您的居住地址（州/市、区、街道）。"},
    {"key": "phone", "uz": "7. Telefon raqamingizni kiriting. Masalan: +998901234567",
     "zh": "7. 请输入您的联系电话，例如：+998901234567"},
    {
        "key": "russian",
        "uz": "8. Rus tilini qay darajada bilasiz?",
        "zh": "8. 您的俄语水平如何？",
        "options": [("Umuman bilmayman", "不会"),
                    ("Boshlang‘ich", "初级"),
                    ("O‘rta", "中级"),
                    ("Yaxshi", "良好"),
                    ("Erkin muloqot qila olaman", "流利")]
    },
    {
        "key": "student",
        "uz": "9. Hozirda talabamisiz?",
        "zh": "9. 您目前是在校学生吗？",
        "options": [("Ha", "是"), ("Yo‘q", "否")]
    },
    {
        "key": "shifts",
        "uz": "10. O‘zgaruvchan smenada ishlashga tayyormisiz?",
        "zh": "10. 您是否能够接受轮班工作？",
        "options": [("Ha", "是"), ("Yo‘q", "否")]
    },
    {"key": "experience", "uz": "11. Ish tajribangiz haqida yozing. Ish joyi, lavozim va ishlagan muddatingizni ko‘rsating. Tajribangiz bo‘lmasa, «Ish tajribam yo‘q» deb yozing.",
     "zh": "11. 请介绍您的工作经历，包括工作单位、职位和工作时间。如无工作经验，请填写“无工作经验”。"},
    {"key": "salary", "uz": "12. Qancha oylik ish haqi kutmoqdasiz? Masalan: 6 000 000 so‘m.",
     "zh": "12. 您期望的月薪是多少？例如：6,000,000苏姆。"},
]


def lang_text(lang, uz, zh):
    return uz if lang == "uz" else zh


def main_menu(lang):
    if lang == "uz":
        return ReplyKeyboardMarkup(
            keyboard=[[KeyboardButton(text="Bo‘sh ish o‘rinlari")],
                      [KeyboardButton(text="Biz haqimizda"), KeyboardButton(text="Bog‘lanish")]],
            resize_keyboard=True
        )
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="职位空缺")],
                  [KeyboardButton(text="关于我们"), KeyboardButton(text="联系方式")]],
        resize_keyboard=True
    )


def departments_kb(lang):
    b = InlineKeyboardBuilder()
    if lang == "uz":
        b.button(text="Savdo bo‘limi", callback_data="dep|Savdo")
        b.button(text="Ishlab chiqarish bo‘limi", callback_data="dep|Ishlab chiqarish")
    else:
        b.button(text="销售部", callback_data="dep|Savdo")
        b.button(text="生产部", callback_data="dep|Ishlab chiqarish")
    b.adjust(1)
    return b.as_markup()


def question_keyboard(q, lang):
    b = InlineKeyboardBuilder()
    for uz, zh in q.get("options", []):
        b.button(text=uz if lang == "uz" else zh, callback_data=f"answer|{uz}")
    b.adjust(1)
    return b.as_markup()


async def send_question(message_or_call, state: FSMContext):
    data = await state.get_data()
    lang = data["lang"]
    idx = data["q_index"]
    all_q = data["all_questions"]
    if idx >= len(all_q):
        await state.set_state(Apply.photo)
        prompt = ("13. Iltimos, rasmingizni yuboring. Oddiy telefon orqali olingan surat bo‘lishi mumkin."
                  if lang == "uz" else
                  "13. 请发送您的个人照片。可以使用手机拍摄的普通照片。")
        if isinstance(message_or_call, CallbackQuery):
            await message_or_call.message.answer(prompt)
        else:
            await message_or_call.answer(prompt)
        return

    q = all_q[idx]
    prompt = lang_text(lang, q["uz"], q["zh"])
    markup = question_keyboard(q, lang) if q.get("options") else None
    if isinstance(message_or_call, CallbackQuery):
        await message_or_call.message.answer(prompt, reply_markup=markup)
    else:
        await message_or_call.answer(prompt, reply_markup=markup)


@dp.message(CommandStart())
async def start(message: Message, state: FSMContext):
    await state.clear()
    kb = InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="O‘zbekcha", callback_data="lang|uz"),
        InlineKeyboardButton(text="中文", callback_data="lang|zh")
    ]])
    await message.answer("BNBM Central Asia HR botiga xush kelibsiz.\n欢迎使用 BNBM Central Asia 招聘机器人。\n\nTilni tanlang / 请选择语言：", reply_markup=kb)


@dp.callback_query(F.data.startswith("lang|"))
async def set_language(call: CallbackQuery, state: FSMContext):
    lang = call.data.split("|")[1]
    await state.update_data(lang=lang)
    await call.message.answer(
        "Kerakli bo‘limni tanlang." if lang == "uz" else "请选择所需服务。",
        reply_markup=main_menu(lang)
    )
    await call.answer()


@dp.message(F.text.in_({"Bo‘sh ish o‘rinlari", "职位空缺"}))
async def vacancies_menu(message: Message, state: FSMContext):
    data = await state.get_data()
    lang = data.get("lang", "uz")
    await message.answer(
        "BNBM Central Asia zavodi." if lang == "uz" else "BNBM Central Asia 工厂。",
        reply_markup=departments_kb(lang)
    )
    await message.answer(
        "Bo‘limni tanlang:" if lang == "uz" else "请选择部门："
    )


@dp.callback_query(F.data.startswith("dep|"))
async def select_department(call: CallbackQuery, state: FSMContext):
    dep = call.data.split("|", 1)[1]
    lang = (await state.get_data()).get("lang", "uz")
    with db() as con:
        rows = con.execute("SELECT * FROM vacancies WHERE department=? AND active=1 ORDER BY id DESC", (dep,)).fetchall()
    if not rows:
        await call.message.answer("Hozircha vakansiya mavjud emas." if lang == "uz" else "目前暂无职位空缺。")
    else:
        b = InlineKeyboardBuilder()
        for row in rows:
            title = row["title_uz"] if lang == "uz" else row["title_zh"]
            b.button(text=title, callback_data=f"vac|{row['id']}")
        b.adjust(1)
        await call.message.answer("Lavozimni tanlang:" if lang == "uz" else "请选择职位：", reply_markup=b.as_markup())
    await call.answer()


@dp.callback_query(F.data.startswith("vac|"))
async def choose_vacancy(call: CallbackQuery, state: FSMContext):
    vacancy_id = int(call.data.split("|")[1])
    lang = (await state.get_data()).get("lang", "uz")
    with db() as con:
        v = con.execute("SELECT * FROM vacancies WHERE id=? AND active=1", (vacancy_id,)).fetchone()
        extras = con.execute("SELECT * FROM extra_questions WHERE vacancy_id=? ORDER BY id", (vacancy_id,)).fetchall()
    if not v:
        await call.message.answer("Vakansiya topilmadi.")
        await call.answer()
        return
    title = v["title_uz"] if lang == "uz" else v["title_zh"]
    desc = v["description_uz"] if lang == "uz" else v["description_zh"]
    await call.message.answer(f"{title}\n\n{desc}")
    all_q = list(QUESTIONS)
    if "kran" in v["title_uz"].lower():
        all_q.extend([
            {"key": "crane_type", "uz": "Kranlarning qaysi turini boshqargansiz va necha yil tajribangiz bor?",
             "zh": "您操作过哪种类型的起重机？有几年操作经验？"},
            {"key": "certificate", "uz": "Kran haydovchisi guvohnomangiz yoki malaka sertifikatingiz bormi? (Bo‘lsa, afzallik beradi.)",
             "zh": "您是否持有起重机操作证或资格证书？（持有者优先。）",
             "options": [("Ha", "有"), ("Yo‘q", "没有"), ("Olishni rejalashtirganman", "计划考取")]}
        ])
    for e in extras:
        all_q.append({"key": f"extra_{e['id']}", "uz": e["question_uz"], "zh": e["question_zh"]})
    await state.set_state(Apply.answering)
    await state.update_data(vacancy_id=vacancy_id, answers={}, q_index=0, all_questions=all_q)
    await send_question(call, state)
    await call.answer()


async def save_answer(message: Message, state: FSMContext, answer: str):
    data = await state.get_data()
    idx = data["q_index"]
    q = data["all_questions"][idx]
    answers = data.get("answers", {})
    answers[q["key"]] = answer
    await state.update_data(answers=answers, q_index=idx + 1)
    await send_question(message, state)


@dp.callback_query(Apply.answering, F.data.startswith("answer|"))
async def answer_button(call: CallbackQuery, state: FSMContext):
    answer = call.data.split("|", 1)[1]
    await save_answer(call.message, state, answer)
    await call.answer()


@dp.message(Apply.answering)
async def answer_text(message: Message, state: FSMContext):
    await save_answer(message, state, message.text or "")


@dp.message(Apply.photo, F.photo)
async def receive_photo(message: Message, state: FSMContext):
    data = await state.get_data()
    answers = data["answers"]
    vacancy_id = data["vacancy_id"]
    photo_id = message.photo[-1].file_id
    lang = data["lang"]
    with db() as con:
        cur = con.execute("""
            INSERT INTO applications (telegram_id,language,vacancy_id,answers,photo_file_id,created_at)
            VALUES (?,?,?,?,?,?)
        """, (message.from_user.id, lang, vacancy_id, repr(answers), photo_id, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        app_id = cur.lastrowid
        vacancy = con.execute("SELECT * FROM vacancies WHERE id=?", (vacancy_id,)).fetchone()
    title = vacancy["title_uz"]
    summary = "\n".join([f"{k}: {v}" for k, v in answers.items()])
    admin_caption = (f"Yangi anketa №{app_id}\nLavozim: {title}\nTelegram ID: {message.from_user.id}\n"
                     f"Til: {'O‘zbekcha' if lang == 'uz' else 'Xitoycha'}\n\n{summary}")
    if ADMIN_ID:
        try:
            await bot.send_photo(ADMIN_ID, photo_id, caption=admin_caption[:1024])
            if len(admin_caption) > 1024:
                await bot.send_message(ADMIN_ID, admin_caption[1024:])
        except Exception:
            logging.exception("Admin xabari yuborilmadi")
    await message.answer(
        f"Anketangiz qabul qilindi. Ariza raqami: {app_id}. Rahmat!"
        if lang == "uz" else
        f"您的申请已提交。申请编号：{app_id}。感谢您！",
        reply_markup=main_menu(lang)
    )
    await state.clear()


@dp.message(Apply.photo)
async def photo_required(message: Message, state: FSMContext):
    lang = (await state.get_data()).get("lang", "uz")
    await message.answer("Iltimos, rasm yuboring." if lang == "uz" else "请发送照片。")


@dp.message(F.text.in_({"Biz haqimizda", "关于我们"}))
async def about(message: Message, state: FSMContext):
    lang = (await state.get_data()).get("lang", "uz")
    await message.answer(
        "BNBM Central Asia — qurilish materiallari ishlab chiqarish korxonasi."
        if lang == "uz" else
        "BNBM Central Asia 是一家建筑材料生产企业。"
    )


@dp.message(F.text.in_({"Bog‘lanish", "联系方式"}))
async def contact(message: Message, state: FSMContext):
    lang = (await state.get_data()).get("lang", "uz")
    await message.answer(
        "Bog‘lanish uchun HR bo‘limiga murojaat qiling."
        if lang == "uz" else
        "如需联系，请咨询人力资源部。"
    )


def admin_only(user_id):
    return ADMIN_ID != 0 and user_id == ADMIN_ID


@dp.message(Command("admin"))
async def admin_panel(message: Message):
    if not admin_only(message.from_user.id):
        return await message.answer("Sizda admin huquqi mavjud emas.")
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Vakansiya qo‘shish", callback_data="adm|addvac")],
        [InlineKeyboardButton(text="Vakansiyalar", callback_data="adm|listvac")],
        [InlineKeyboardButton(text="Qo‘shimcha savol qo‘shish", callback_data="adm|addq")],
        [InlineKeyboardButton(text="Excel hisobot", callback_data="adm|excel")],
        [InlineKeyboardButton(text="Arizalar soni", callback_data="adm|count")]
    ])
    await message.answer("HR boshqaruv paneli:", reply_markup=kb)


@dp.callback_query(F.data.startswith("adm|"))
async def admin_actions(call: CallbackQuery, state: FSMContext):
    if not admin_only(call.from_user.id):
        await call.answer("Ruxsat yo‘q", show_alert=True)
        return
    action = call.data.split("|")[1]
    if action == "addvac":
        await state.set_state(AddVacancy.department)
        await call.message.answer("Bo‘limni kiriting: Savdo yoki Ishlab chiqarish")
    elif action == "listvac":
        with db() as con:
            rows = con.execute("SELECT * FROM vacancies ORDER BY id DESC").fetchall()
        if not rows:
            await call.message.answer("Vakansiyalar yo‘q.")
        for r in rows:
            await call.message.answer(f"#{r['id']} | {r['department']} | {r['title_uz']} | {'Faol' if r['active'] else 'Yopiq'}")
    elif action == "addq":
        with db() as con:
            rows = con.execute("SELECT id,title_uz FROM vacancies WHERE active=1").fetchall()
        if not rows:
            await call.message.answer("Avval vakansiya yarating.")
        else:
            b = InlineKeyboardBuilder()
            for r in rows:
                b.button(text=f"#{r['id']} {r['title_uz']}", callback_data=f"qvac|{r['id']}")
            b.adjust(1)
            await call.message.answer("Savol qo‘shiladigan vakansiyani tanlang:", reply_markup=b.as_markup())
    elif action == "excel":
        await export_excel(call.message)
    elif action == "count":
        with db() as con:
            n = con.execute("SELECT COUNT(*) FROM applications").fetchone()[0]
        await call.message.answer(f"Jami arizalar: {n}")
    await call.answer()


@dp.callback_query(F.data.startswith("qvac|"))
async def choose_question_vacancy(call: CallbackQuery, state: FSMContext):
    if not admin_only(call.from_user.id):
        return await call.answer("Ruxsat yo‘q", show_alert=True)
    vid = int(call.data.split("|")[1])
    await state.set_state(AddQuestion.question_uz)
    await state.update_data(q_vacancy_id=vid)
    await call.message.answer("Qo‘shimcha savolni o‘zbek tilida yuboring:")
    await call.answer()


@dp.message(AddQuestion.question_uz)
async def extra_question_uz(message: Message, state: FSMContext):
    await state.update_data(extra_uz=message.text)
    await state.set_state(AddQuestion.question_zh)
    await message.answer("Savolning xitoycha tarjimasini yuboring:")


@dp.message(AddQuestion.question_zh)
async def extra_question_zh(message: Message, state: FSMContext):
    data = await state.get_data()
    with db() as con:
        con.execute("INSERT INTO extra_questions(vacancy_id,question_uz,question_zh) VALUES(?,?,?)",
                    (data["q_vacancy_id"], data["extra_uz"], message.text))
    await state.clear()
    await message.answer("Qo‘shimcha savol saqlandi.")


@dp.message(AddVacancy.department)
async def addvac_department(message: Message, state: FSMContext):
    dep = message.text.strip()
    if dep not in ("Savdo", "Ishlab chiqarish"):
        return await message.answer("Faqat Savdo yoki Ishlab chiqarish deb kiriting.")
    await state.update_data(dep=dep)
    await state.set_state(AddVacancy.title_uz)
    await message.answer("Lavozim nomini o‘zbek tilida kiriting:")


@dp.message(AddVacancy.title_uz)
async def addvac_title_uz(message: Message, state: FSMContext):
    await state.update_data(title_uz=message.text)
    await state.set_state(AddVacancy.title_zh)
    await message.answer("Lavozim nomini xitoy tilida kiriting:")


@dp.message(AddVacancy.title_zh)
async def addvac_title_zh(message: Message, state: FSMContext):
    await state.update_data(title_zh=message.text)
    await state.set_state(AddVacancy.desc_uz)
    await message.answer("Vakansiya tavsifini o‘zbek tilida kiriting:")


@dp.message(AddVacancy.desc_uz)
async def addvac_desc_uz(message: Message, state: FSMContext):
    await state.update_data(desc_uz=message.text)
    await state.set_state(AddVacancy.desc_zh)
    await message.answer("Vakansiya tavsifini xitoy tilida kiriting:")


@dp.message(AddVacancy.desc_zh)
async def addvac_desc_zh(message: Message, state: FSMContext):
    data = await state.get_data()
    with db() as con:
        con.execute("""
            INSERT INTO vacancies(department,title_uz,title_zh,description_uz,description_zh)
            VALUES(?,?,?,?,?)
        """, (data["dep"], data["title_uz"], data["title_zh"], data["desc_uz"], message.text))
    await state.clear()
    await message.answer("Yangi vakansiya muvaffaqiyatli qo‘shildi.")


async def export_excel(message: Message):
    with db() as con:
        rows = con.execute("""
            SELECT a.id,a.telegram_id,a.language,v.department,v.title_uz,a.answers,
                   a.status,a.created_at
            FROM applications a JOIN vacancies v ON v.id=a.vacancy_id
            ORDER BY a.id DESC
        """).fetchall()
    wb = Workbook()
    ws = wb.active
    ws.title = "Arizalar"
    ws.append(["Ariza №", "Telegram ID", "Til", "Bo‘lim", "Lavozim", "Javoblar", "Holat", "Sana"])
    for r in rows:
        ws.append([r["id"], r["telegram_id"], r["language"], r["department"],
                   r["title_uz"], r["answers"], r["status"], r["created_at"]])
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    await message.answer_document(BufferedInputFile(stream.read(), filename="BNBM_arizalar.xlsx"))


async def main():
    init_db()
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
