import asyncio
import logging
from datetime import datetime
from typing import Dict, List, Optional
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.types import FSInputFile  # Не забудьте импортировать FSInputFile
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
    CallbackQuery,
)
from aiogram.utils.keyboard import InlineKeyboardBuilder
import openpyxl
from openpyxl import Workbook
import io
import pandas as pd
from aiogram.types import BufferedInputFile
from aiogram.filters import Command
import os
from dotenv import load_dotenv

# ------------------- КОНФИГ -------------------
load_dotenv()

# Получаем токен из окружения
BOT_TOKEN = os.getenv('BOT_TOKEN')
EXCEL_FILENAME = "survey_results.xlsx"

# ------------------- ЛОГИРОВАНИЕ -------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ------------------- СОЗДАЕМ ДИСПЕТЧЕР -------------------
dp = Dispatcher()

# ------------------- СОСТОЯНИЯ FSM -------------------
class SurveyStates(StatesGroup):
    start_question = State()
    full_name = State()
    age = State()
    phone = State()
    goal = State()
    goal_other = State()
    current_format = State()
    desired_format = State()
    frequency = State()
    weekdays = State()
    time_slots = State()
    morning_intervals = State()
    afternoon_intervals = State()
    evening_intervals = State()
    unavailable_time = State()
    weekend = State()
    weekend_time = State()
    teacher_importance = State()
    teacher_change_reason = State()
    teacher_like = State()
    teacher_feedback = State()
    schedule_flexibility = State()
    wishes = State()
    wishes_other = State()
    need_manager = State()
    review_choice = State()
    rules_confirmation = State()

# ------------------- КЛАВИАТУРЫ -------------------
def make_inline_keyboard(options: Dict[str, str], callback_prefix: str) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    for cb_suffix, label in options.items():
        builder.button(text=label, callback_data=f"{callback_prefix}{cb_suffix}")
    builder.adjust(1)
    return builder.as_markup()

# ------------------- ВАРИАНТЫ ОТВЕТОВ -------------------
START_OPTIONS = {
    'positive': 'Да, хочу пройти опрос',
    'negative': 'Нет, отложу на потом'
}
GOAL_OPTIONS = {
    "general": "Общий английский / язык для себя",
    "spoken": "Разговорный язык",
    "raise_level": "Повысить текущий уровень",
    "school": "Английский для школы",
    "progress": "Повысить успеваемость",
    "oge": "Подготовка к ОГЭ",
    "ege": "Подготовка к ЕГЭ",
    "international": "Подготовка к международному экзамену",
    "work": "Английский для работы",
    "travel": "Английский для путешествий",
    "relocate": "Подготовка к переезду",
    "other": "Другая цель",
}
CURRENT_FORMAT_OPTIONS = {
    "individual": "Индивидуально",
    "split": "Сплит - парные занятия",
    "mini_group": "Мини-группа",
}
DESIRED_FORMAT_OPTIONS = {
    "individual": "Индивидуально",
    "split": "Сплит - парные занятия",
    "mini_group": "Мини-группа",
    "leave_current": "Хочу оставить текущий формат",
    "consider_other": "Готов рассмотреть другой формат",
}
FREQUENCY_OPTIONS = {
    "1": "1 раз в неделю",
    "2": "2 раза в неделю",
    "3": "3 раза в неделю",
    "more": "Более 3 раз в неделю",
    "leave_current": "Хочу оставить текущую частоту занятий",
}
WEEKDAYS_OPTIONS = {
    "mon": "Понедельник",
    "tue": "Вторник",
    "wed": "Среда",
    "thu": "Четверг",
    "fri": "Пятница",
}
TIME_SLOTS_OPTIONS = {
    "morning": "Первая половина дня",
    "afternoon": "Вторая половина дня",
    "evening": "Вечер",
    "any": "Время не принципиально",
}
MORNING_INTERVALS = {
    "08-09": "08:00-09:00",
    "09-10": "09:00-10:00",
    "10-11": "10:00-11:00",
    "11-12": "11:00-12:00",
    "12-13": "12:00-13:00",
    "13-14": "13:00-14:00",
}
AFTERNOON_INTERVALS = {
    "14-15": "14:00-15:00",
    "15-16": "15:00-16:00",
    "16-17": "16:00-17:00",
    "17-18": "17:00-18:00",
}
EVENING_INTERVALS = {
    "18-19": "18:00-19:00",
    "19-20": "19:00-20:00",
    "20-21": "20:00-21:00",
    "21plus": "21:00 и позже",
}
WEEKEND_OPTIONS = {
    "sat": "Да, в субботу",
    "sun": "Да, в воскресенье",
    "both": "Подойдут оба дня",
    "only_weekday": "Если не получится в будни",
    "no": "Нет",
}
WEEKEND_TIME_OPTIONS = {
    "morning": "Утро",
    "day": "День",
    "evening": "Вечер",
    "any": "Время не принципиально",
}
TEACHER_IMPORTANCE_OPTIONS = {
    "very_important": "Очень важно",
    "desirable": "Желательно сохранить преподавателя",
    "other_teacher": "Готов рассмотреть другого",
    "change": "Хотел бы сменить преподавателя",
}
TEACHER_LIKE_OPTIONS = {
    "great": "Да, все отлично ❤️",
    "good_but_wishes": "В целом да, но есть пожелания",
    "discuss": "Есть моменты для обсуждения",
    "change": "Хотел бы сменить преподавателя",
}
SCHEDULE_FLEXIBILITY_OPTIONS = {
    "30min": "Да, на 30 минут раньше или позже",
    "other_days": "Да, могу рассмотреть другие дни",
    "other_format": "Могу рассмотреть другой формат",
    "no": "Нет, не готов(а)",
}
WISHES_OPTIONS = {
        "more_talk": "Больше разговорной практики",
        "more_grammar": "Больше грамматики",
        "pronunciation": "Работа над произношением",
        "vocabulary": "Увеличить словарный запас",
        "school_program": "Помощь со школьной программой",
        "exam_prep": "Подготовка к экзамену",
        "homework": "Больше домашних заданий",
        "less_homework": "Уменьшить объем домашних заданий",
        "wishes_other": "Другие пожелания"
}
NEED_MANAGER_OPTIONS = {
    "yes": "Да",
    "no": "Нет, все указал в анкете",
}
REVIEW_SERVICE_OPTIONS = {
    "yandex": "Яндекс Карты - 200 бонусов",
    "2gis": "2ГИС - 200 бонусов",
    "otzovik": "Отзовик - 300 бонусов",
    "zoon": "Zoon - 300 бонусов",
    "later": "Оставлю позже",
}
REVIEW_LINKS = {
    "yandex": "https://yandex.ru/maps/org/lingva_family/209910153141?si=jjyzmmr1qvu81w1u3qehf50u40",
    "2gis": "https://2gis.ru/moscow/firm/70000001087290873/tab/reviews?m=37.579667%2C55.759481%2F16",
    "otzovik": "https://otzovik.com/postreview.php?pid=2480908",
    "zoon": "https://zoon.ru/msk/trainings/kursy_inostrannyh_yazykov_lingva_family",
}

# ------------------- РАБОТА С EXCEL -------------------
def get_answers_dict() -> Dict:
    return {
        "user_id": None,
        "username": None,
        "first_name": None,
        "full_name": "",
        "age": "",
        "phone": "",
        "goal": "",
        "goal_other": "",
        "current_format": "",
        "desired_format": "",
        "frequency": "",
        "weekdays": "",
        "time_slots": "",
        "morning_intervals": "",
        "afternoon_intervals": "",
        "evening_intervals": "",
        "unavailable_time": "",
        "weekend": "",
        "weekend_time": "",
        "teacher_importance": "",
        "teacher_change_reason": "",
        "teacher_like": "",
        "teacher_feedback": "",
        "schedule_flexibility": "",
        "wishes": "",
        "need_manager": "",
        "review_choice": "",
        "rules_confirmed": False,
        "timestamp": "",
    }

def save_to_excel(answers: Dict):
    row_data = [
        answers.get("user_id"),
        answers.get("username"),
        answers.get("first_name"),
        answers.get("full_name"),
        answers.get("age"),
        answers.get("phone"),
        answers.get("goal"),
        answers.get("goal_other"),
        answers.get("current_format"),
        answers.get("desired_format"),
        answers.get("frequency"),
        answers.get("weekdays"),
        answers.get("time_slots"),
        answers.get("morning_intervals"),
        answers.get("afternoon_intervals"),
        answers.get("evening_intervals"),
        answers.get("unavailable_time"),
        answers.get("weekend"),
        answers.get("weekend_time"),
        answers.get("teacher_importance"),
        answers.get("teacher_change_reason"),
        answers.get("teacher_like"),
        answers.get("teacher_feedback"),
        answers.get("schedule_flexibility"),
        answers.get("wishes"),
        answers.get("need_manager"),
        answers.get("review_choice"),
        answers.get("rules_confirmed"),
        answers.get("timestamp"),
    ]
    headers = [
        "user_id", "username", "first_name", "full_name", "age", "phone",
        "goal", "goal_other", "current_format", "desired_format",
        "frequency", "weekdays", "time_slots", "morning_intervals",
        "afternoon_intervals", "evening_intervals", "unavailable_time",
        "weekend", "weekend_time", "teacher_importance",
        "teacher_change_reason", "teacher_like", "teacher_feedback",
        "schedule_flexibility", "wishes", "need_manager",
        "review_choice", "rules_confirmed", "timestamp"
    ]
    try:
        try:
            wb = openpyxl.load_workbook(EXCEL_FILENAME)
            ws = wb.active
        except FileNotFoundError:
            wb = Workbook()
            ws = wb.active
            ws.append(headers)
        ws.append(row_data)
        wb.save(EXCEL_FILENAME)
        logger.info(f"Ответы сохранены для user_id={answers.get('user_id')}")
    except PermissionError:
        logger.error(f"Файл {EXCEL_FILENAME} заблокирован! Записываем во временный файл.")
        temp_filename = f"survey_results_temp_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        wb.save(temp_filename)

# ------------------- ХЭНДЛЕРЫ -------------------

@dp.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    answers = get_answers_dict()
    answers["user_id"] = message.from_user.id
    answers["username"] = message.from_user.username or ""
    answers["first_name"] = message.from_user.first_name or ""
    answers["timestamp"] = datetime.now().isoformat()
    await state.update_data(answers=answers)
    text = """
🎓 Начинаем новый учебный год в Lingva Family!

Приветствуем! ❤️ Мы очень рады, что вы с нами!

Впереди новый учебный год - будем учиться, двигаться к своим целям, видеть результат и получать удовольствие от занятий ✨

Мы уже составляем расписание и хотим учесть ваши планы и пожелания.
Если вам нужно поменять дни или время занятий, попробовать другой формат или у вас появилась новая цель - расскажите нам об этом.
Ответьте, пожалуйста, на несколько коротких вопросов. Так мы сможем сделать новый учебный год удобным и комфортным для вас ❤️
Начнем? 👇"""
    kb = make_inline_keyboard(START_OPTIONS, "start_")

    await message.answer(text, reply_markup=kb)
    await state.set_state(SurveyStates.start_question)


@dp.callback_query(SurveyStates.start_question, F.data.startswith("start_"))
async def process_start_choice(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data.get("answers", {})
    answers["start_choice"] = START_OPTIONS.get(choice, choice)
    await state.update_data(answers=answers)

    if answers["start_choice"] == START_OPTIONS["positive"]:
        await callback.message.edit_text("Укажите ФИО студента.")
        await state.set_state(SurveyStates.full_name)
    else:
        await callback.message.edit_text(
            "Хорошо, вы можете пройти опрос позже."
        )
        await state.clear()

    await callback.answer()

@dp.message(SurveyStates.full_name)
async def process_full_name(message: Message, state: FSMContext):
    data = await state.get_data()
    answers = data["answers"]
    answers["full_name"] = message.text.strip()
    await state.update_data(answers=answers)
    await message.answer("Укажите возраст студента.")

@dp.message(SurveyStates.age)
async def process_age(message: Message, state: FSMContext):
    data = await state.get_data()
    answers = data["answers"]
    answers["age"] = message.text.strip()
    await state.update_data(answers=answers)
    await message.answer("Напишите ваш номер телефона.")
    await state.set_state(SurveyStates.phone)

@dp.message(SurveyStates.phone)
async def process_phone(message: Message, state: FSMContext):
    data = await state.get_data()
    answers = data["answers"]
    answers["phone"] = message.text.strip()
    await state.update_data(answers=answers)
    text = "Какая у вас сейчас основная цель в обучении?"
    kb = make_inline_keyboard(GOAL_OPTIONS, "goal_")
    await message.answer(text, reply_markup=kb)
    await state.set_state(SurveyStates.goal)

@dp.callback_query(SurveyStates.goal, F.data.startswith("goal_"))
async def process_goal(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    answers["goal"] = GOAL_OPTIONS.get(choice, choice)
    await state.update_data(answers=answers)
    if choice == "other":
        await callback.message.edit_text(
            "Расскажите коротко, какой результат вы хотите получить от занятий?"
        )
        await state.set_state(SurveyStates.goal_other)
    else:
        await callback.message.edit_text(f"Выбрано: {answers['goal']}")
        await ask_current_format(callback.message, state)
    await callback.answer()

@dp.message(SurveyStates.goal_other)
async def process_goal_other(message: Message, state: FSMContext):
    data = await state.get_data()
    answers = data["answers"]
    answers["goal_other"] = message.text.strip()
    await state.update_data(answers=answers)
    await ask_current_format(message, state)

async def ask_current_format(msg: Message, state: FSMContext):
    text = "В каком формате вы занимаетесь сейчас?"
    kb = make_inline_keyboard(CURRENT_FORMAT_OPTIONS, "curfmt_")
    await msg.answer(text, reply_markup=kb)
    await state.set_state(SurveyStates.current_format)

@dp.callback_query(SurveyStates.current_format, F.data.startswith("curfmt_"))
async def process_current_format(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    answers["current_format"] = CURRENT_FORMAT_OPTIONS.get(choice, choice)
    await state.update_data(answers=answers)
    await callback.message.edit_text(f"Выбрано: {answers['current_format']}")
    await ask_desired_format(callback.message, state)
    await callback.answer()

async def ask_desired_format(msg: Message, state: FSMContext):
    data = await state.get_data()
    answers = data["answers"]
    selected = answers.get("desired_format", "")
    selected_list = [s.strip() for s in selected.split(",") if s.strip()] if selected else []
    text = "В каком формате вы хотели бы продолжить обучение в новом учебном году? (можно выбрать несколько)"
    kb = InlineKeyboardBuilder()
    for key, label in DESIRED_FORMAT_OPTIONS.items():
        checked = " ✅" if label in selected_list else ""
        kb.button(text=f"{label}{checked}", callback_data=f"desfmt_{key}")
    kb.button(text="✅ Готово", callback_data="desfmt_done")
    kb.adjust(1)
    await msg.answer(text, reply_markup=kb.as_markup())
    await state.set_state(SurveyStates.desired_format)

@dp.callback_query(SurveyStates.desired_format, F.data.startswith("desfmt_"))
async def process_desired_format(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    selected = answers.get("desired_format", "")
    selected_list = [s.strip() for s in selected.split(",") if s.strip()] if selected else []
    if choice == "done":
        await callback.message.edit_text(f"Выбрано: {answers['desired_format']}")
        await ask_frequency(callback.message, state)
        await callback.answer()
        return
    label = DESIRED_FORMAT_OPTIONS.get(choice, choice)
    if label in selected_list:
        selected_list.remove(label)
    else:
        selected_list.append(label)
    answers["desired_format"] = ", ".join(selected_list)
    await state.update_data(answers=answers)
    text = "В каком формате вы хотели бы продолжить обучение в новом учебном году? (можно выбрать несколько)"
    kb = InlineKeyboardBuilder()
    for key, lbl in DESIRED_FORMAT_OPTIONS.items():
        checked = " ✅" if lbl in selected_list else ""
        kb.button(text=f"{lbl}{checked}", callback_data=f"desfmt_{key}")
    kb.button(text="✅ Готово", callback_data="desfmt_done")
    kb.adjust(1)
    await callback.message.edit_text(text, reply_markup=kb.as_markup())
    await callback.answer()

async def ask_frequency(msg: Message, state: FSMContext):
    text = "Сколько раз в неделю вы хотели бы заниматься?"
    kb = make_inline_keyboard(FREQUENCY_OPTIONS, "freq_")
    await msg.answer(text, reply_markup=kb)
    await state.set_state(SurveyStates.frequency)

@dp.callback_query(SurveyStates.frequency, F.data.startswith("freq_"))
async def process_frequency(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    answers["frequency"] = FREQUENCY_OPTIONS.get(choice, choice)
    await state.update_data(answers=answers)
    await callback.message.edit_text(f"Выбрано: {answers['frequency']}")
    await ask_weekdays(callback.message, state)
    await callback.answer()

async def ask_weekdays(msg: Message, state: FSMContext):
    data = await state.get_data()
    answers = data["answers"]
    selected = answers.get("weekdays", "")
    selected_list = [s.strip() for s in selected.split(",") if s.strip()] if selected else []
    text = "Какие дни в будни вам предпочтительнее? (можно выбрать несколько)"
    kb = InlineKeyboardBuilder()
    for key, label in WEEKDAYS_OPTIONS.items():
        checked = " ✅" if label in selected_list else ""
        kb.button(text=f"{label}{checked}", callback_data=f"wk_{key}")
    kb.button(text="✅ Готово", callback_data="wk_done")
    kb.adjust(1)
    await msg.answer(text, reply_markup=kb.as_markup())
    await state.set_state(SurveyStates.weekdays)

@dp.callback_query(SurveyStates.weekdays, F.data.startswith("wk_"))
async def process_weekdays(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    selected = answers.get("weekdays", "")
    selected_list = [s.strip() for s in selected.split(",") if s.strip()] if selected else []
    if choice == "done":
        await callback.message.edit_text(f"Выбрано: {answers['weekdays']}")
        await ask_time_slots(callback.message, state)
        await callback.answer()
        return
    label = WEEKDAYS_OPTIONS.get(choice, choice)
    if label in selected_list:
        selected_list.remove(label)
    else:
        selected_list.append(label)
    answers["weekdays"] = ", ".join(selected_list)
    await state.update_data(answers=answers)
    text = "Какие дни в будни вам предпочтительнее? (можно выбрать несколько)"
    kb = InlineKeyboardBuilder()
    for key, lbl in WEEKDAYS_OPTIONS.items():
        checked = " ✅" if lbl in selected_list else ""
        kb.button(text=f"{lbl}{checked}", callback_data=f"wk_{key}")
    kb.button(text="✅ Готово", callback_data="wk_done")
    kb.adjust(1)
    await callback.message.edit_text(text, reply_markup=kb.as_markup())
    await callback.answer()

async def ask_time_slots(msg: Message, state: FSMContext):
    await state.update_data(temp_time_slots=[])
    text = "Какое время занятий вам удобнее? (можно выбрать несколько)"
    kb = InlineKeyboardBuilder()
    for key, label in TIME_SLOTS_OPTIONS.items():
        kb.button(text=label, callback_data=f"ts_{key}")
    kb.button(text="✅ Готово", callback_data="ts_done")
    kb.adjust(1)
    await msg.answer(text, reply_markup=kb.as_markup())
    await state.set_state(SurveyStates.time_slots)

@dp.callback_query(SurveyStates.time_slots, F.data == "ts_done")
async def process_time_slots_done(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    temp_slots = data.get("temp_time_slots", [])
    answers = data["answers"]
    answers["time_slots"] = ", ".join(temp_slots)
    await state.update_data(answers=answers)
    await callback.message.edit_text(f"Выбрано: {answers['time_slots'] if answers['time_slots'] else 'Не выбрано'}")
    
    if "Первая половина дня" in temp_slots:
        await ask_morning_intervals(callback.message, state)
    elif "Вторая половина дня" in temp_slots:
        await ask_afternoon_intervals(callback.message, state)
    elif "Вечер" in temp_slots:
        await ask_evening_intervals(callback.message, state)
    else:
        await ask_unavailable_time(callback.message, state)
        
    await callback.answer()

@dp.callback_query(SurveyStates.time_slots, F.data.startswith("ts_"))
async def process_time_slots(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    selected_list = data.get("temp_time_slots", [])
    label = TIME_SLOTS_OPTIONS.get(choice, choice)
    if label in selected_list:
        selected_list.remove(label)
    else:
        selected_list.append(label)
    await state.update_data(temp_time_slots=selected_list)
    text = "Какое время занятий вам удобнее? (можно выбрать несколько)"
    kb = InlineKeyboardBuilder()
    for key, lbl in TIME_SLOTS_OPTIONS.items():
        checked = " ✅" if lbl in selected_list else ""
        kb.button(text=f"{lbl}{checked}", callback_data=f"ts_{key}")
    kb.button(text="✅ Готово", callback_data="ts_done")
    kb.adjust(1)
    await callback.message.edit_text(text, reply_markup=kb.as_markup())
    await callback.answer()

async def ask_morning_intervals(msg: Message, state: FSMContext):
    data = await state.get_data()
    answers = data["answers"]
    selected = answers.get("morning_intervals", "")
    selected_list = [s.strip() for s in selected.split(",") if s.strip()] if selected else []
    text = "Отметьте все подходящие интервалы в первой половине дня:"
    kb = InlineKeyboardBuilder()
    for key, label in MORNING_INTERVALS.items():
        checked = " ✅" if label in selected_list else ""
        kb.button(text=f"{label}{checked}", callback_data=f"mint_{key}")
    kb.button(text="✅ Готово", callback_data="mint_done")
    kb.adjust(1)
    await msg.answer(text, reply_markup=kb.as_markup())
    await state.set_state(SurveyStates.morning_intervals)

@dp.callback_query(SurveyStates.morning_intervals, F.data.startswith("mint_"))
async def process_morning_intervals(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    selected = answers.get("morning_intervals", "")
    selected_list = [s.strip() for s in selected.split(",") if s.strip()] if selected else []
    if choice == "done":
        answers["morning_intervals"] = ", ".join(selected_list)
        await state.update_data(answers=answers)
        await callback.message.edit_text(f"Выбрано: {answers['morning_intervals']}")
        slot_names = [s.strip() for s in answers.get("time_slots", "").split(",") if s.strip()]
        if "Вторая половина дня" in slot_names:
            await ask_afternoon_intervals(callback.message, state)
        elif "Вечер" in slot_names:
            await ask_evening_intervals(callback.message, state)
        else:
            await ask_unavailable_time(callback.message, state)
        await callback.answer()
        return
    label = MORNING_INTERVALS.get(choice, choice)
    if label in selected_list:
        selected_list.remove(label)
    else:
        selected_list.append(label)
    answers["morning_intervals"] = ", ".join(selected_list)
    await state.update_data(answers=answers)
    text = "Отметьте все подходящие интервалы в первой половине дня:"
    kb = InlineKeyboardBuilder()
    for key, lbl in MORNING_INTERVALS.items():
        checked = " ✅" if lbl in selected_list else ""
        kb.button(text=f"{lbl}{checked}", callback_data=f"mint_{key}")
    kb.button(text="✅ Готово", callback_data="mint_done")
    kb.adjust(1)
    await callback.message.edit_text(text, reply_markup=kb.as_markup())
    await callback.answer()

async def ask_afternoon_intervals(msg: Message, state: FSMContext):
    data = await state.get_data()
    answers = data["answers"]
    selected = answers.get("afternoon_intervals", "")
    selected_list = [s.strip() for s in selected.split(",") if s.strip()] if selected else []
    text = "Отметьте все подходящие интервалы во второй половине дня:"
    kb = InlineKeyboardBuilder()
    for key, label in AFTERNOON_INTERVALS.items():
        checked = " ✅" if label in selected_list else ""
        kb.button(text=f"{label}{checked}", callback_data=f"aint_{key}")
    kb.button(text="✅ Готово", callback_data="aint_done")
    kb.adjust(1)
    await msg.answer(text, reply_markup=kb.as_markup())
    await state.set_state(SurveyStates.afternoon_intervals)

@dp.callback_query(SurveyStates.afternoon_intervals, F.data.startswith("aint_"))
async def process_afternoon_intervals(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    selected = answers.get("afternoon_intervals", "")
    selected_list = [s.strip() for s in selected.split(",") if s.strip()] if selected else []
    if choice == "done":
        answers["afternoon_intervals"] = ", ".join(selected_list)
        await state.update_data(answers=answers)
        await callback.message.edit_text(f"Выбрано: {answers['afternoon_intervals']}")
        slot_names = [s.strip() for s in answers.get("time_slots", "").split(",") if s.strip()]
        if "Вечер" in slot_names:
            await ask_evening_intervals(callback.message, state)
        else:
            await ask_unavailable_time(callback.message, state)
        await callback.answer()
        return
    label = AFTERNOON_INTERVALS.get(choice, choice)
    if label in selected_list:
        selected_list.remove(label)
    else:
        selected_list.append(label)
    answers["afternoon_intervals"] = ", ".join(selected_list)
    await state.update_data(answers=answers)
    text = "Отметьте все подходящие интервалы во второй половине дня:"
    kb = InlineKeyboardBuilder()
    for key, lbl in AFTERNOON_INTERVALS.items():
        checked = " ✅" if lbl in selected_list else ""
        kb.button(text=f"{lbl}{checked}", callback_data=f"aint_{key}")
    kb.button(text="✅ Готово", callback_data="aint_done")
    kb.adjust(1)
    await callback.message.edit_text(text, reply_markup=kb.as_markup())
    await callback.answer()

async def ask_evening_intervals(msg: Message, state: FSMContext):
    data = await state.get_data()
    answers = data["answers"]
    selected = answers.get("evening_intervals", "")
    selected_list = [s.strip() for s in selected.split(",") if s.strip()] if selected else []
    text = "Отметьте все подходящие интервалы вечером:"
    kb = InlineKeyboardBuilder()
    for key, label in EVENING_INTERVALS.items():
        checked = " ✅" if label in selected_list else ""
        kb.button(text=f"{label}{checked}", callback_data=f"eint_{key}")
    kb.button(text="✅ Готово", callback_data="eint_done")
    kb.adjust(1)
    await msg.answer(text, reply_markup=kb.as_markup())
    await state.set_state(SurveyStates.evening_intervals)

@dp.callback_query(SurveyStates.evening_intervals, F.data.startswith("eint_"))
async def process_evening_intervals(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    selected = answers.get("evening_intervals", "")
    selected_list = [s.strip() for s in selected.split(",") if s.strip()] if selected else []
    if choice == "done":
        answers["evening_intervals"] = ", ".join(selected_list)
        await state.update_data(answers=answers)
        await callback.message.edit_text(f"Выбрано: {answers['evening_intervals']}")
        await ask_unavailable_time(callback.message, state)
        await callback.answer()
        return
    label = EVENING_INTERVALS.get(choice, choice)
    if label in selected_list:
        selected_list.remove(label)
    else:
        selected_list.append(label)
    answers["evening_intervals"] = ", ".join(selected_list)
    await state.update_data(answers=answers)
    text = "Отметьте все подходящие интервалы вечером:"
    kb = InlineKeyboardBuilder()
    for key, lbl in EVENING_INTERVALS.items():
        checked = " ✅" if lbl in selected_list else ""
        kb.button(text=f"{lbl}{checked}", callback_data=f"eint_{key}")
    kb.button(text="✅ Готово", callback_data="eint_done")
    kb.adjust(1)
    await callback.message.edit_text(text, reply_markup=kb.as_markup())
    await callback.answer()

async def ask_unavailable_time(msg: Message, state: FSMContext):
    text = "Есть ли время, в которое вы точно НЕ можете заниматься?\nНапример: понедельник до 17:00, среда весь день, пятница после 19:00."
    await msg.answer(text)
    await state.set_state(SurveyStates.unavailable_time)

@dp.message(SurveyStates.unavailable_time)
async def process_unavailable_time(message: Message, state: FSMContext):
    data = await state.get_data()
    answers = data["answers"]
    answers["unavailable_time"] = message.text.strip()
    await state.update_data(answers=answers)
    await ask_weekend(message, state)

async def ask_weekend(msg: Message, state: FSMContext):
    text = "Хотели бы вы рассмотреть занятия в выходные дни?"
    kb = make_inline_keyboard(WEEKEND_OPTIONS, "wknd_")
    await msg.answer(text, reply_markup=kb)
    await state.set_state(SurveyStates.weekend)

@dp.callback_query(SurveyStates.weekend, F.data.startswith("wknd_"))
async def process_weekend(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    answers["weekend"] = WEEKEND_OPTIONS.get(choice, choice)
    await state.update_data(answers=answers)
    if choice in ("sat", "sun", "both"):
        await callback.message.edit_text(f"Выбрано: {answers['weekend']}")
        await ask_weekend_time(callback.message, state)
    else:
        await callback.message.edit_text(f"Выбрано: {answers['weekend']}")
        await ask_teacher_importance(callback.message, state)
    await callback.answer()

async def ask_weekend_time(msg: Message, state: FSMContext):
    text = "Какое время вам удобнее в выходные?"
    kb = make_inline_keyboard(WEEKEND_TIME_OPTIONS, "wkt_")
    await msg.answer(text, reply_markup=kb)
    await state.set_state(SurveyStates.weekend_time)

@dp.callback_query(SurveyStates.weekend_time, F.data.startswith("wkt_"))
async def process_weekend_time(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    answers["weekend_time"] = WEEKEND_TIME_OPTIONS.get(choice, choice)
    await state.update_data(answers=answers)
    await callback.message.edit_text(f"Выбрано: {answers['weekend_time']}")
    await ask_teacher_importance(callback.message, state)
    await callback.answer()

async def ask_teacher_importance(msg: Message, state: FSMContext):
    text = "Насколько важно сохранить вашего текущего преподавателя?"
    kb = make_inline_keyboard(TEACHER_IMPORTANCE_OPTIONS, "timp_")
    await msg.answer(text, reply_markup=kb)
    await state.set_state(SurveyStates.teacher_importance)

@dp.callback_query(SurveyStates.teacher_importance, F.data.startswith("timp_"))
async def process_teacher_importance(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    answers["teacher_importance"] = TEACHER_IMPORTANCE_OPTIONS.get(choice, choice)
    await state.update_data(answers=answers)
    if choice == "change":
        await callback.message.edit_text("Расскажите, пожалуйста, почему вы рассматриваете смену преподавателя?")
        await state.set_state(SurveyStates.teacher_change_reason)
    else:
        await callback.message.edit_text(f"Выбрано: {answers['teacher_importance']}")
        await ask_teacher_like(callback.message, state)
    await callback.answer()

@dp.message(SurveyStates.teacher_change_reason)
async def process_teacher_change_reason(message: Message, state: FSMContext):
    data = await state.get_data()
    answers = data["answers"]
    answers["teacher_change_reason"] = message.text.strip()
    await state.update_data(answers=answers)
    await ask_teacher_like(message, state)

# --- 12. Показ вопроса про отношение к преподавателю с кнопками ---
async def ask_teacher_like(msg: Message, state: FSMContext):
    text = "Нравится ли вам заниматься с текущим преподавателем?"
    kb = make_inline_keyboard(TEACHER_LIKE_OPTIONS, "tlike_")
    await msg.answer(text, reply_markup=kb)
    await state.set_state(SurveyStates.teacher_like)

@dp.callback_query(SurveyStates.teacher_like, F.data.startswith("tlike_"))
async def process_teacher_like(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    answers["teacher_like"] = TEACHER_LIKE_OPTIONS.get(choice, choice)
    await state.update_data(answers=answers)

    # Если выбраны "change" или "discuss", запрашиваем развернутый ввод
    if choice in ("change", "discuss"):
        await callback.message.edit_text(
            f"Выбрано: {answers['teacher_like']}\n\n"
            "Расскажите, пожалуйста, подробнее, с чем это связано?"
        )
        await state.set_state(SurveyStates.teacher_feedback)
    else:
        await callback.message.edit_text(f"Выбрано: {answers['teacher_like']}")
        await ask_schedule_flexibility(callback.message, state)
        
    await callback.answer()

@dp.message(SurveyStates.teacher_feedback)
async def process_teacher_feedback(message: Message, state: FSMContext):
    data = await state.get_data()
    answers = data["answers"]
    answers["teacher_feedback"] = message.text.strip()
    await state.update_data(answers=answers)
    await ask_schedule_flexibility(message, state)

async def ask_schedule_flexibility(msg: Message, state: FSMContext):
    text = "Если подходящего времени не окажется, готовы ли вы немного скорректировать расписание?"
    kb = make_inline_keyboard(SCHEDULE_FLEXIBILITY_OPTIONS, "flex_")
    await msg.answer(text, reply_markup=kb)
    await state.set_state(SurveyStates.schedule_flexibility)

@dp.callback_query(SurveyStates.schedule_flexibility, F.data.startswith("flex_"))
async def process_schedule_flexibility(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    answers["schedule_flexibility"] = SCHEDULE_FLEXIBILITY_OPTIONS.get(choice, choice)
    await state.update_data(answers=answers)
    await callback.message.edit_text(f"Выбрано: {answers['schedule_flexibility']}")
    await ask_wishes(callback.message, state)
    await callback.answer()

async def ask_wishes(msg: Message, state: FSMContext):
    text = "Есть ли у вас пожелания по обучению на новый учебный год?"
    kb = make_inline_keyboard(WISHES_OPTIONS, "wishes_")
    await msg.answer(text, reply_markup=kb)
    await state.set_state(SurveyStates.wishes)

# 2. Обработка клика по кнопкам 14-го вопроса
@dp.callback_query(SurveyStates.wishes, F.data.startswith("wishes_"))
async def process_wishes(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data.get("answers", {})
    
    answers["wishes"] = WISHES_OPTIONS.get(choice, choice)
    await state.update_data(answers=answers)
    
    if choice == "other":
        await callback.message.edit_text("Напишите в свободной форме ваши пожелания по обучению:")
        await state.set_state(SurveyStates.wishes_other)
    else:
        await callback.message.edit_text(f"Выбрано: {answers['wishes']}")
        await ask_need_manager(callback.message, state)
        
    await callback.answer()

# 3. Обработка текстового ввода, если выбран вариант "Другие пожелания"
@dp.message(SurveyStates.wishes_other)
async def process_wishes_other(message: Message, state: FSMContext):
    data = await state.get_data()
    answers = data.get("answers", {})
    
    answers["wishes_other"] = message.text.strip()
    await state.update_data(answers=answers)
    
    await ask_need_manager(message, state)

async def ask_need_manager(msg: Message, state: FSMContext):
    text = "Хотели бы вы, чтобы менеджер связался с вами и помог подобрать оптимальный формат или расписание?"
    kb = make_inline_keyboard(NEED_MANAGER_OPTIONS, "mgr_")
    await msg.answer(text, reply_markup=kb)
    await state.set_state(SurveyStates.need_manager)

@dp.callback_query(SurveyStates.need_manager, F.data.startswith("mgr_"))
async def process_need_manager(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    answers["need_manager"] = NEED_MANAGER_OPTIONS.get(choice, choice)
    await state.update_data(answers=answers)
    await callback.message.edit_text(f"Выбрано: {answers['need_manager']}")
    await ask_review(callback.message, state)
    await callback.answer()

async def ask_review(msg: Message, state: FSMContext):
    text = (
        "Если вам нравится заниматься в Lingva Family, поделитесь своим впечатлением о школе.\n"
        "Ваш отзыв помогает другим ученикам и родителям сделать выбор, а нам - становиться еще лучше.\n"
        "За опубликованный отзыв мы начислим вам бонусы от Lingva Family 🎁\n"
        "Где вам удобнее оставить отзыв?"
    )
    kb = make_inline_keyboard(REVIEW_SERVICE_OPTIONS, "rev_")
    await msg.answer(text, reply_markup=kb)
    await state.set_state(SurveyStates.review_choice)

@dp.callback_query(SurveyStates.review_choice, F.data.startswith("rev_"))
async def process_review_choice(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split("_", 1)[1]
    data = await state.get_data()
    answers = data["answers"]
    answers["review_choice"] = REVIEW_SERVICE_OPTIONS.get(choice, choice)
    await state.update_data(answers=answers)
    link = REVIEW_LINKS.get(choice)
    if link:
        await callback.message.edit_text(
            f"Вы выбрали: {answers['review_choice']}\n"
            f"Ссылка для отзыва: {link}\n\n"
            "После того как оставите отзыв, нажмите кнопку ниже, чтобы продолжить."
        )
        await callback.message.answer("Теперь перейдём к ознакомлению с правилами.")
    else:
        await callback.message.edit_text("Хорошо, вы можете оставить отзыв позже.")
    await ask_rules(callback.message, state)
    await callback.answer()

async def ask_rules(msg: Message, state: FSMContext):
    text = (
        "Правила обучения в Lingva Family\n\n"
        "Перед началом нового учебного сезона, пожалуйста, ознакомьтесь с Правилами ученика Lingva Family.\n"
        "В них собрана важная информация о занятиях, переносах и отменах, отработках, заморозке абонемента, оплате и других организационных моментах.\n\n"
        "📌 Правила ученика Lingva Family\n"
        "[Ознакомиться с правилами](https://lingvafamily.ru/rules)\n\n"
        "После ознакомления нажмите кнопку:"
    )
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="☑ С правилами ознакомлен(а)", callback_data="rules_ok")]
        ]
    )
    await msg.answer(text, reply_markup=kb, disable_web_page_preview=True)
    await state.set_state(SurveyStates.rules_confirmation)

@dp.callback_query(SurveyStates.rules_confirmation, F.data == "rules_ok")
async def process_rules_ok(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    answers = data["answers"]
    answers["rules_confirmed"] = True
    await state.update_data(answers=answers)
    save_to_excel(answers)
    final_text = (
        '''Спасибо! ❤️ Все готово.

Мы получили ваши ответы и обязательно учтем пожелания при формировании расписания на новый учебный год.
Если потребуется что-то уточнить, наш менеджер свяжется с вами.

💬 И самое главное - мы всегда открыты для обратной связи.
Если в процессе обучения у вас возникнет вопрос, пожелание или ситуация, которую хотелось бы обсудить - обязательно сообщите нам. Мы готовы внимательно разобраться в любой ситуации и вместе найти оптимальное решение.
Нам важно, чтобы обучение в Lingva Family было не только эффективным, но и комфортным для вас.

Спасибо, что вы с нами ❤️ 
До встречи в новом учебном сезоне!
Команда Lingva Family'''
    )
    await callback.message.edit_text(final_text)
    await state.clear()
    await callback.answer()

# ------------------- ЭКСПОРТ В EXCEL (АДМИН) -------------------
ADMIN_IDS = [8071127858, 711314367]

@dp.message(Command("export"))
async def export_to_excel(message: Message):
    # Проверка на права администратора
    if message.from_user.id not in ADMIN_IDS:
        await message.answer("❌ У вас нет прав для выполнения этой команды.")
        return

    try:
        # Чтение Excel-файла в память
        with open(EXCEL_FILENAME, "rb") as file:
            file_data = file.read()
        
        # Подготовка файла для отправки через aiogram
        document = BufferedInputFile(file_data, filename=EXCEL_FILENAME)
        
        # Отправка файла в чат
        await message.answer_document(
            document=document,
            caption=f"📊 Экспорт ответов на {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        )
    except FileNotFoundError:
        await message.answer("⚠️ Файл с результатами опроса еще не создан. Никто пока не прошел опрос.")
    except Exception as e:
        logger.error(f"Ошибка при экспорте файла: {e}")
        await message.answer("❌ Произошла ошибка при выгрузке файла.")



# ------------------- ЗАПУСК -------------------
async def main():
    bot = Bot(token=BOT_TOKEN)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
    
