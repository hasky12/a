import logging
import pandas as pd
from telegram import Update, KeyboardButton, ReplyKeyboardMarkup, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, CallbackContext, CallbackQueryHandler
from datetime import datetime, time
import os
import random

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Путь к файлу Excel
EXCEL_FILE = os.path.expanduser("~/Desktop/work_log.xlsx")

# Имена сотрудников (начальные значения)
EMPLOYEES = ["Влад Горяной", "Влад Задарий", "Юрий", "Даня"]

# ID администратора и токен бота
ADMIN_CHAT_ID = 1358116229
BOT_TOKEN = "7460801290:AAFJcdJa-tlusH4u3JXfqw55Q7fDagvLzEE"

# Создание Excel файла, если он не существует
if not os.path.isfile(EXCEL_FILE):
    df = pd.DataFrame(columns=["Дата"] + EMPLOYEES)
    df.to_excel(EXCEL_FILE, index=False)

# Список поощрительных слов
ENCOURAGEMENTS = [
    "Красавчик!",
    "ото вьебав!",
    "нихуя себе!",
    "Замечательно выполнено!",
    "да ты монстр",
    "оставь в покое работу, всё переделаешь так",
    "иди получи премию",
    "да ты ваще",
    "запись в колонку премия будет выполнена ты получил 100грн",
    "Дай рядом стоящему по шеи",
    "Фантастическая работа!",
    "я в шоке ",
    "ебаааааааать",
    "Хорошая работа АлеГ",
    "нужно больше!!!!",
    "хуль так мало?",
    "можешь перекурить",
    "выпей кофе, наверное устал",
    "ого!",
    "Ты настоящий мастер!",
    "Круто!"
]

# Функция для сохранения chat_id пользователя
def save_user_chat_id(update: Update, context: CallbackContext) -> None:
    user_id = update.effective_user.id
    chat_id = update.effective_chat.id
    context.bot_data[user_id] = chat_id

# Функция для обработки команды /start
async def start(update: Update, context: CallbackContext) -> None:
    if 'has_started' not in context.user_data:
        save_user_chat_id(update, context)  # Сохранение chat_id пользователя
        buttons = [
            [KeyboardButton("/work")],
            [KeyboardButton("/feedback")],
            [KeyboardButton("/admin_panel")] if update.effective_chat.id == ADMIN_CHAT_ID else []
        ]
        reply_markup = ReplyKeyboardMarkup(buttons, resize_keyboard=True)
        context.user_data['stage'] = None  # Сбрасываем этап при запуске
        
        await update.message.reply_text(
            "Этот бот помогает вам вести учет выполненной работы. Вот что он умеет:\n\n"
            "1. /work - Введите описание вашей выполненной работы. Выберите свое имя и введите описание.\n\n\n\n\n"
            "2. /feedback - МОЖНО НАПИСАТЬ О ПРОБЛЕМАХ ИЛИ ЖАЛОБАХ. СООБЩЕНИЕ НЕ ВИДЕТ НИКТО КРОМЕ АНДРЕЯ. ТАК ЖЕ МОЖНО НАПИСАТЬ ЧТО ЗАКАНЧИВАЕТСЯ ИЛИ ЧТО НУЖНО КУПИТЬ.\n\n\n\n\n\n"
            "3. /admin_panel - Доступно только для администратора. Позволяет отправлять уведомления всем пользователям и настраивать автоматические напоминания.\n\n"
            "Используйте команды ниже для взаимодействия с ботом.",
            reply_markup=reply_markup
        )
        context.user_data['has_started'] = True  # Устанавливаем флаг, что описание уже было показано
    else:
        buttons = [
            [KeyboardButton("/work")],
            [KeyboardButton("/feedback")],
            [KeyboardButton("/admin_panel")] if update.effective_chat.id == ADMIN_CHAT_ID else []
        ]
        reply_markup = ReplyKeyboardMarkup(buttons, resize_keyboard=True)
        await update.message.reply_text("Выберите команду:", reply_markup=reply_markup)


# Функция для обработки команды /work
async def work(update: Update, context: CallbackContext) -> None:
    save_user_chat_id(update, context)  # Сохранение chat_id пользователя
    keyboard = [[KeyboardButton(employee)] for employee in EMPLOYEES]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    context.user_data['stage'] = 'select_employee'
    await update.message.reply_text('Выберите свое имя:', reply_markup=reply_markup)

# Функция для обработки команды /feedback
async def feedback(update: Update, context: CallbackContext) -> None:
    save_user_chat_id(update, context)  # Сохранение chat_id пользователя
    context.user_data['stage'] = 'feedback'
    await update.message.reply_text("Пожалуйста, введите ваше сообщение для администратора:")

# Функция для обработки админ панели
async def admin_panel(update: Update, context: CallbackContext) -> None:
    keyboard = [
        [InlineKeyboardButton("Отправить уведомление", callback_data='send_notification')],
        [InlineKeyboardButton("Отправить уведомление о записи", callback_data='send_record_reminder')],
        [InlineKeyboardButton("Настройки автоматического уведомления", callback_data='settings_notification')]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    context.user_data['stage'] = 'admin_panel'
    await update.message.reply_text("Выберите действие:", reply_markup=reply_markup)

# Функция для обработки нажатий на кнопки админ панели
async def button_click(update: Update, context: CallbackContext) -> None:
    query = update.callback_query
    await query.answer()

    if query.data == 'send_notification':
        context.user_data['stage'] = 'send_notification'
        await query.edit_message_text("Введите сообщение для отправки пользователям:")

    elif query.data == 'send_record_reminder':
        for user_id, chat_id in context.bot_data.items():
            if chat_id != ADMIN_CHAT_ID:  # Исключаем администратора
                await context.bot.send_message(chat_id=chat_id, text="Не забудьте сделать запись о выполненной работе!")
        await query.edit_message_text("Уведомление о необходимости сделать запись отправлено.")
        context.user_data['stage'] = None
        await start(update, context)

    elif query.data == 'settings_notification':
        context.user_data['stage'] = 'settings_notification'
        await query.edit_message_text("Введите время в формате HH:MM для настройки автоматической отправки уведомлений:")

# Функция для обработки ввода текста для уведомлений и настроек
async def handle_response(update: Update, context: CallbackContext) -> None:
    stage = context.user_data.get('stage')
    save_user_chat_id(update, context)  # Сохранение chat_id пользователя
    
    if stage == 'select_employee':
        employee = update.message.text
        if employee in EMPLOYEES:
            context.user_data['employee'] = employee
            context.user_data['stage'] = 'enter_description'
            await update.message.reply_text("Введите описание выполненной работы:")
        else:
            await update.message.reply_text("Произошла ошибка. Пожалуйста, выберите имя из списка.")
    
    elif stage == 'enter_description':
        employee = context.user_data.get('employee')
        description = update.message.text
        
        if not description.strip():
            await update.message.reply_text("Описание работы не может быть пустым. Пожалуйста, введите описание.")
            return
        
        date_str = datetime.now().strftime("%Y-%m-%d")
        try:
            df = pd.read_excel(EXCEL_FILE)
        except ImportError as e:
            await update.message.reply_text("Ошибка импорта библиотеки. Пожалуйста, обновите библиотеку openpyxl: pip install --upgrade openpyxl")
            logger.error("Ошибка импорта библиотеки openpyxl: %s", e)
            return

        if date_str in df['Дата'].values:
            row_idx = df.index[df['Дата'] == date_str].tolist()[0]
        else:
            row_idx = len(df)
            df.loc[row_idx, 'Дата'] = date_str

        if pd.isna(df.at[row_idx, employee]):
            df.at[row_idx, employee] = description
        else:
            df.at[row_idx, employee] += f"; {description}"

        try:
            df.to_excel(EXCEL_FILE, index=False)
            await update.message.reply_text("Данные успешно сохранены.")
            
            # Уведомление администратора
            await context.bot.send_message(chat_id=ADMIN_CHAT_ID, text=f"Новая запись от {employee}: {description}")
            
            # Отправка поощрительного сообщения
            encouragement = random.choice(ENCOURAGEMENTS)
            await update.message.reply_text(encouragement)
            context.user_data['stage'] = None  # Сбрасываем этап после завершения

            # Возвращаем пользователя в главное меню
            await start(update, context)
        except PermissionError:
            await update.message.reply_text("Ошибка доступа к файлу. Пожалуйста, закройте файл Excel и попробуйте снова.")
    
    elif stage == 'feedback':
        feedback_message = update.message.text
        await context.bot.send_message(chat_id=ADMIN_CHAT_ID, text=f"Обратная связь от {update.effective_user.first_name}:\n{feedback_message}")
        await update.message.reply_text("Ваше сообщение отправлено администратору. Спасибо за ваш отзыв!")
        context.user_data['stage'] = None  # Сбрасываем этап после завершения обработки обратной связи

        # Возвращаем пользователя в главное меню
        await start(update, context)

    elif stage == 'send_notification':
        notification_message = update.message.text
        for user_id, chat_id in context.bot_data.items():
            if chat_id != ADMIN_CHAT_ID:  # Исключаем администратора
                await context.bot.send_message(chat_id=chat_id, text=notification_message)
        await update.message.reply_text("Уведомление отправлено всем пользователям.")
        context.user_data['stage'] = None  # Сбрасываем этап после отправки уведомления

        # Возвращаем пользователя в главное меню
        await start(update, context)

    elif stage == 'settings_notification':
        try:
            notification_time = datetime.strptime(update.message.text, "%H:%M").time()
            job_queue = context.job_queue
            job_queue.run_daily(reminder, notification_time)
            await update.message.reply_text(f"Время автоматической отправки уведомлений настроено на {notification_time}.")
        except ValueError:
            await update.message.reply_text("Неверный формат времени. Пожалуйста, введите время в формате HH:MM.")
        context.user_data['stage'] = None  # Сбрасываем этап после настройки времени

        # Возвращаем пользователя в главное меню
        await start(update, context)

    else:
        await update.message.reply_text("Произошла ошибка. Пожалуйста, начните с команды /work или выберите нужную команду.")

# Функция для напоминания о неотправленных отчетах
async def reminder(context: CallbackContext) -> None:
    df = pd.read_excel(EXCEL_FILE)
    date_str = datetime.now().strftime("%Y-%m-%d")
    
    if date_str in df['Дата'].values:
        row = df[df['Дата'] == date_str].iloc[0]
        for employee in EMPLOYEES:
            if pd.isna(row[employee]):
                chat_id = context.bot_data.get(f"{employee}_chat_id")
                if chat_id:
                    await context.bot.send_message(chat_id=chat_id, text=f"{employee}, вы забыли отправить отчет о выполненной работе сегодня.")
    else:
        for employee in EMPLOYEES:
            chat_id = context.bot_data.get(f"{employee}_chat_id")
            if chat_id:
                await context.bot.send_message(chat_id=chat_id, text=f"{employee}, вы забыли отправить отчет о выполненной работе сегодня.")

# Функция для обработки ошибок
async def error_handler(update: Update, context: CallbackContext) -> None:
    logger.error(msg="Exception while handling an update:", exc_info=context.error)

# Основная функция
def main() -> None:
    application = Application.builder().token(BOT_TOKEN).build()
    
    # Обработчики команд
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("work", work))
    application.add_handler(CommandHandler("feedback", feedback))
    application.add_handler(CommandHandler("admin_panel", admin_panel))
    application.add_handler(CallbackQueryHandler(button_click))
    
    # Обработчик этапов работы
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_response))

    # Настройка напоминания
    job_queue = application.job_queue
    job_queue.run_daily(reminder, time(hour=17, minute=0))
    
    # Обработчик ошибок
    application.add_error_handler(error_handler)

    application.run_polling()

if __name__ == '__main__':
    main()
