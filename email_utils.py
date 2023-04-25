import os
import smtplib
from email.mime.text import MIMEText

from dotenv import load_dotenv

from xlsx_parser import get_string_shipping_address

load_dotenv()

YANDEX_MAIL_LOGIN = os.getenv('YANDEX_MAIL_LOGIN')
YANDEX_MAIL_PASSWORD = os.getenv('YANDEX_MAIL_PASSWORD')


def build_message_from_kwargs(kwargs):
    """Форма для типового сообщения после успешной оплаты."""
    if kwargs.get('shipping_address'):
        shipping_address = get_string_shipping_address(
            kwargs.get('shipping_address')
        )
        print(shipping_address)
        text_message = (
            f'{kwargs.get("name")}! \n'
            f'Ваш заказ "{kwargs.get("title")}" оплачен'
            f' и ожидает отправки по адресу: {shipping_address}'
        )
        return text_message
    text_message = (
        f'{kwargs.get("name")}! \nВаш заказ "{kwargs.get("title")}" оплачен'
    )
    return text_message


def build_email(message, to):
    """Билдит сообщение имэйл."""
    if to:
        msg = MIMEText(message)
        msg['Subject'] = 'Информация по заказу!'
        msg['From'] = YANDEX_MAIL_LOGIN + '@yandex.ru'
        msg['To'] = to
        return msg


def send_built_msg(message: MIMEText):
    """Коннектится к серверу и отправляет имэйл."""
    server = smtplib.SMTP_SSL('smtp.yandex.ru', port=465)
    server.login(
        YANDEX_MAIL_LOGIN,
        YANDEX_MAIL_PASSWORD
    )
    server.sendmail(
        from_addr=message.get('From'),
        to_addrs=message.get('To'),
        msg=message.as_string()
    )
    server.quit()
