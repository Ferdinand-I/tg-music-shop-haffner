"""Telegram магазин.

Для правильного функционирования, все товары должны быть упакованы
в файл "shop.xlsx" согласно заданному формату.

Telegram bot token и payment provider token следует указать в файле .env

Для отправки email используется smtp сервер Yandex. Для использования функций
отправки необходимо в переменные окружения добавить значения логина и пароля
для Yandex.
"""
import datetime
import os
import re

from dotenv import load_dotenv
from telegram import (
    InlineKeyboardButton, InlineKeyboardMarkup, LabeledPrice, ShippingOption
)
from telegram.ext import (
    ApplicationBuilder, CommandHandler, CallbackQueryHandler,
    PreCheckoutQueryHandler, MessageHandler, filters, ConversationHandler,
    ShippingQueryHandler
)

from email_utils import (
    build_email, send_built_msg, build_message_from_kwargs
)
from xlsx_parser import (
    collect_items, collect_admins_id, add_admin_to_excel, add_transaction,
    calculate_total_marge
)

load_dotenv()

TOKEN = os.getenv('BOT_TOKEN')
PROVIDER_TOKEN = os.getenv('PROVIDER_TOKEN')
CONST_EXCEL_NAME = 'shop'

# Словарбь товаров и их характеристик, полученный после парсинга
# соответствущего xlsx-файла "shop.xlsx"
ITEMS = collect_items('shop.xlsx', 'shop')

# Список id пользователей, у которых есть права администратора
ADMINS = collect_admins_id('shop.xlsx', 'admin')

# ID возвращает функция start_add_admin(), как маркер состояния диалога
ID = 0

# Опции отправки
SHIPPING_OPTIONS = [
    ShippingOption(
        '1', 'CDEK Москва', [
            LabeledPrice('Доставка в Москву', 354 * 100),
        ]
    ),
    ShippingOption(
        '2', 'CDEK Новосибирск', [
            LabeledPrice('Доставка в Новосибирск', 404 * 100),
        ]
    )
]


async def start(update, context):
    """Велкам мессадж."""
    keyboard = [
        [InlineKeyboardButton('Показать товары', callback_data='show_items')]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    text = (
        'Для того, чтобы посмотреть список всех товаров нажмите на кнопку '
        '<b>"Показать товары"</b> или отправьте в чат сообщение /show'
    )
    await update.message.reply_text(
        text=text,
        reply_markup=reply_markup,
        parse_mode='html'
    )


async def show(update, context):
    """Показать все доступные товары магазина в виде инлайн кнопок."""
    if ITEMS:
        # Параметры клавиатуры берутся из словаря характеристик товаров,
        # который формируется при запуске скрипта или
        # при вызове функции update()
        keyboard = [
            [
                InlineKeyboardButton(
                    ITEMS.get(item).get('title'),
                    callback_data=item
                )
            ] for item in ITEMS
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text='Сейчас у нас в наличии: ',
            reply_markup=reply_markup,
        )
    else:
        await update.callback_query.answer(
            text=(
                'Список товаров пока недоступен для просмотра. '
                'Попробуйте позже.'
            ),
            show_alert=True
        )


async def send_invoice(update, context, shipping: bool, key: str):
    """Отправляет счёт фактуру."""
    keyboard = [
        [
            InlineKeyboardButton(
                'Оплатить',
                pay=True,
            )
        ],
        [
            InlineKeyboardButton(
                'Показать все товары',
                callback_data='show_items'
            )
        ]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    # Параметры инвойса берутся из словаря характеристик товаров,
    # который формируется при запуске скрипта или
    # при вызове функции update()
    title = ITEMS.get(key).get('title')
    description = ITEMS.get(key).get('description')
    currency = ITEMS.get(key).get('currency')
    prices = ITEMS.get(key).get('prices')
    options = {
        'chat_id': update.effective_chat.id,
        'title': title,
        'description': description,
        'payload': key,
        'provider_token': PROVIDER_TOKEN,
        'currency': currency,
        'prices': prices,
        'photo_url': 'https://i.ibb.co/JzS1x9r/photo-2023-01-14-18-49-35.jpg',
        'photo_height': 400,
        'photo_width': 400,
        'reply_markup': reply_markup,
        'need_name': True,
        'need_email': True,
        'need_phone_number': True,
    }
    if shipping:
        options['description'] += ' (с доставкой)'
        await context.bot.send_invoice(
            **options, need_shipping_address=True, is_flexible=True
        )
    else:
        options['description'] += ' (без доставки)'
        await context.bot.send_invoice(**options)


async def callback_button(update, context):
    """Реакция на инлайн-кнопки."""
    if update.callback_query.data == 'show_items':
        await show(update, context)
    elif 'need_shipping' in update.callback_query.data:
        key = update.callback_query.data.split()[0]
        await send_invoice(update, context, shipping=True, key=key)
    elif 'no_shipping' in update.callback_query.data:
        key = update.callback_query.data.split()[0]
        await send_invoice(update, context, shipping=False, key=key)
    else:
        item_title = ITEMS.get(update.callback_query.data).get('title')
        keyboard = [
            [
                InlineKeyboardButton(
                    'Да',
                    callback_data=update.callback_query.data + ' need_shipping'
                ),
                InlineKeyboardButton(
                    'Нет',
                    callback_data=update.callback_query.data + ' no_shipping'
                )
            ]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=f'Вы хотите оформить доставку "{item_title}"?',
            reply_markup=reply_markup,
        )


async def precheckout_callback(update, context):
    """Проверка счёт-фактуры."""
    query = update.pre_checkout_query
    if query.invoice_payload not in ITEMS:
        await query.answer(ok=False, error_message='Что-то пошло не так...')
    else:
        await query.answer(ok=True)


async def successful_payment_callback(update, context):
    """Сообщение об успешной оплате."""
    product_key = update.message.successful_payment.invoice_payload
    order_info = update.message.successful_payment.order_info
    kwargs = {
        'title': ITEMS.get(product_key).get('title'),
        'price': update.message.successful_payment.total_amount / 100,
        'currency': update.message.successful_payment.currency,
        'name': order_info.name,
        'email': order_info.email,
        'phone_number': order_info.phone_number,
        'status': 'Оплачено',
        'datetime': datetime.datetime.now().strftime('%d-%m-%y %H:%M'),
        'shipping_address': order_info.shipping_address
    }
    add_transaction('shop.xlsx', sheet_name='transactions', **kwargs)
    keyboard = [
        [InlineKeyboardButton('Показать товары', callback_data='show_items')]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text(
        'Спасибо, что приобрели товар в нашем магазине!',
        reply_markup=reply_markup,
    )
    # Отправляем письмо с сообщением об успешной оплате
    text_message = build_message_from_kwargs(kwargs)
    email_message = build_email(text_message, kwargs.get('email'))
    send_built_msg(email_message)


async def venue(update, context):
    """Сообщение с меткой на карте."""
    await update.message.reply_venue(
        latitude='59.927513',
        longitude='30.315841',
        title='Самовывоз доступен в Санкт-Петербурге по адресу',
        address='Канал Грибоедова, 52',
    )


async def update_shop(update, context):
    """Обновление списка товаров в магазине через парсинг эксель-файла."""
    if update.effective_user.id in ADMINS:
        # Перезаписываем глобальную переменную, содержащую словарь с товарами
        global ITEMS
        ITEMS = collect_items('shop.xlsx', 'shop')
        await update.message.reply_text(
            text=(
                'Ревизия успешно проведена!'
            ),
        )
        return ITEMS
    await update.message.reply_text(
        text=(
            'У вас не достаточно прав для инициализации ревизии.'
            'За подробностями обратитесь к @ferdinand_the_second'
        ),
    )


async def download_excel_admin(update, context):
    """Скачать эксель файл со списком товаров и админов."""
    if update.effective_user.id in ADMINS:
        with open('shop.xlsx', 'rb') as file:
            await update.message.reply_document(
                document=file
            )
    else:
        await update.message.reply_text(
            text=(
                'У вас не достаточно прав для получения рабочей '
                'документации магазина. '
                'За подробностями обратитесь к @ferdinand_the_second'
            ),
        )


async def upload_excel_admin(update, context):
    """Бот скачивает файл с товарами магазина."""
    if update.effective_user.id in ADMINS:
        if update.message.document.file_name == 'shop.xlsx':
            file = await context.bot.get_file(update.message.document.file_id)
            await file.download_to_drive('shop.xlsx')
            await update_shop(update, context)
        else:
            await update.message.reply_text(
                'Проверьте загружаемый файл. Он должен называться "shop.xlsx"'
            )
    else:
        await update.message.reply_text(
            text=(
                'У вас не достаточно прав для загрузки файлов на сервер. '
                'За подробностями обратитесь к @ferdinand_the_second'
            ),
        )


async def shipping(update, context):
    """Отвечает на запрос отправки товара."""
    query = update.shipping_query
    if query.invoice_payload not in ITEMS:
        await query.answer(ok=False, error_message='Что-то пошло не так...')
        return
    await query.answer(ok=True, shipping_options=SHIPPING_OPTIONS)


async def start_add_admin(update, context):
    """Энтри поинт начала диалога для добавления нового админа."""
    if update.effective_user.id in ADMINS:
        await update.message.reply_text(
            'Пришлите валидный id пользователя, который состоит из 6-9 цифр. '
            'Для завершения диалога введите /cancel'
        )
        return ID
    await update.message.reply_text(
        text=(
            'У вас не достаточно прав для добавления нового администратора. '
            'За подробностями обратитесь к @ferdinand_the_second'
        ),
    )
    return ConversationHandler.END


async def validate_add_admin(update, context):
    """Валидация и запись нового админ айди в память и в эксель (БД)."""
    new_admin_id = update.message.text
    if re.fullmatch(r'[+]?\d{6,10}', new_admin_id):
        ADMINS.append(int(new_admin_id))
        add_admin_to_excel('shop.xlsx', 'admin', admin_id=int(new_admin_id))
        await update.message.reply_text(
            'Новый админ успешно добавлен!'
        )
        return ConversationHandler.END
    elif update.message.text == '/cancel':
        await update.message.reply_text(
            'Диалог завершён'
        )
        return ConversationHandler.END
    else:
        await update.message.reply_text(
            'Id пользователя не валидный. '
            'Отправьте /add чтобы попробовать снова.'
        )
        return ConversationHandler.END


async def calculate_total(update, context):
    """Тотал маржа."""
    result = calculate_total_marge('shop.xlsx', total=True)
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=f'Общая выручка за всё время составляет: {result} рублей'
    )


async def cancel(update, context):
    """Конец конверсейшна с добавлением нового админа."""
    await update.message.reply_text(
        'Диалог завершён'
    )
    return ConversationHandler.END


def main():
    """Запуск бота."""
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler('start', start))
    app.add_handler(CommandHandler('show', show))
    app.add_handler(CommandHandler('venue', venue))
    app.add_handler(CommandHandler('update', update_shop))
    app.add_handler(CommandHandler('download', download_excel_admin))
    app.add_handler(
        MessageHandler(
            filters.Document.FileExtension('xlsx'),
            upload_excel_admin
        )
    )
    app.add_handler(
        ConversationHandler(
            entry_points=[CommandHandler('add', start_add_admin)],
            states={
                ID: [
                    MessageHandler(
                        filters.TEXT,
                        validate_add_admin,
                    )
                ]
            },
            fallbacks=[CommandHandler('cancel', cancel)],
        )
    )
    app.add_handler(CallbackQueryHandler(callback_button))
    app.add_handler(PreCheckoutQueryHandler(precheckout_callback))
    app.add_handler(
        MessageHandler(filters.SUCCESSFUL_PAYMENT, successful_payment_callback)
    )
    app.add_handler(ShippingQueryHandler(shipping))
    app.run_polling()


if __name__ == '__main__':
    main()
