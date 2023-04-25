# Telegram Music Shop Bot <img src="https://shoppingator.ru/pict/tovar/httpswwwcdn-front.kwork.ru/pics/t3/73/20180713-1649405373.jpg" width=64>

Телеграм магазин, написанный на **Python**.

Вместо БД бот использует excel таблицы для хранения и получения информации. В файле shop.xlsx можно указать id tg пользователей для получения административных прав.

Чтобы запустить проект локально:

1. Клонируйте репозиторий себе на компьютер, находясь в директории, откуда вы хотите в будущем запускать проект (в примере испоьзуется ссылка для подключения с помощью протокола **SSH** в консоли **BASH** для **WINDOWS**)

```BASH
git clone git@github.com:Ferdinand-I/tg-music-shop-haffner.git
```

2. Создайте и активируйте виртуальное окружение (в примере используется утилита **venv**), перейдите в директорию проекта

```BASH
python -m venv venv
source venv/Scripts/activate
cd tg-music-shop-haffner
```

3. Обновите **PIP** и установите зависимости *requirements.txt*

```BASH
python -m pip install --upgrade pip
pip install -r requirements.txt
```

4. Создайте файл *.env*, откройте его для редактирования с помощью редактора **nano**

```BASH
touch .env
nano .env
```

5. В редакторе опишите переменные по шаблону, сохраните изменения, выйдите из редактора

```nano
BOT_TOKEN=<ваш токен при регистрации бота telegram>
PROVIDER_TOKEN=<токен аккредитованного telegram банка эмитента для осуществления эквайринга>
YANDEX_MAIL_LOGIN=<для имэйл рассылки используется Яндекс SMTP сервер>
YANDEX_MAIL_PASSWORD=<пароль для подключения к SMTP серверу>
```

6. Запустите проект

```BASH
python shop_bot.py
```
