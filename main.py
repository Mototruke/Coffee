from flask import Flask, render_template_string, request, jsonify
import os
import re
from datetime import datetime

try:
    import pyodbc
except ImportError:
    pyodbc = None

app = Flask(__name__)

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'coffee_db.accdb')


def get_db_connection():
    if pyodbc is None or not os.path.exists(DB_PATH):
        return None
    conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + DB_PATH + ';'
    return pyodbc.connect(conn_str)


def validate_name_and_phone(name, phone):
    name = (name or '').strip()
    phone = (phone or '').strip()
    if not name or any(ch.isdigit() for ch in name):
        return False, "Имя не должно содержать цифры и не может быть пустым."
    if not re.fullmatch(r'[A-Za-zА-Яа-яЁё\s\-]+', name):
        return False, "В имени допускаются только буквы, пробелы и дефис."
    if not re.fullmatch(r'\d{11}', phone):
        return False, "Номер телефона должен состоять ровно из 11 цифр (например, 89991234567)."
    return True, ""


# Запасной список меню (если база Access недоступна)
DEFAULT_MENU = [
    {"id": 1, "title": "Капучино (300 мл)", "category": "Кофе", "description": "Классический двойной эспрессо с молочной пенкой", "price": 400, "img_url": "/static/kapus.webp"},
    {"id": 2, "title": "Латте (300 мл)", "category": "Кофе", "description": "Нежный кофейный напиток с большим количеством молока", "price": 400, "img_url": "/static/latte.webp"},
    {"id": 3, "title": "Раф Кофе (300 мл)", "category": "Кофе", "description": "С оттенками ванили, карамели или шоколада", "price": 450, "img_url": "/static/raf.webp"},
    {"id": 4, "title": "Бамбл Кофе (300 мл)", "category": "Кофе", "description": "Освежающий микс эспрессо, апельсинового сока и карамели", "price": 400, "img_url": "/static/bumble.webp"},
    {"id": 5, "title": "Айс Матча «Малина-кокос»", "category": "Кофе", "description": "Сезонная новинка (500 мл)", "price": 500, "img_url": "/static/matcha_malina.webp"},
    {"id": 6, "title": "Айс латте «Сникерс»", "category": "Кофе", "description": "С шоколадной пенкой и джус боллами (500 мл)", "price": 500, "img_url": "/static/snickers.webp"},
    {"id": 7, "title": "Раф Лимонный пирог", "category": "Кофе", "description": "Авторский кофе со вкусом домашней выпечки (300 мл)", "price": 450, "img_url": "/static/raf_lemon.webp"},
    {"id": 8, "title": "Латте Пряная груша", "category": "Кофе", "description": "Авторский согревающий латте (300 мл)", "price": 450, "img_url": "/static/latte_pear.webp"},
    {"id": 9, "title": "Облепиховый чай (700 мл)", "category": "Чай и напитки", "description": "С мёдом, имбирём и корицей", "price": 600, "img_url": "/static/tea_oblepiha.webp"},
    {"id": 10, "title": "Яблочный сбитень (700 мл)", "category": "Чай и напитки", "description": "Традиционный пряный горячий напиток", "price": 600, "img_url": "/static/sbiten.webp"},
    {"id": 11, "title": "Лимонад Манго-апельсин", "category": "Чай и напитки", "description": "Освежающий домашний лимонад (500 мл)", "price": 500, "img_url": "/static/limonad.webp"},
    {"id": 12, "title": "Смузи Зеленый микс", "category": "Чай и напитки", "description": "Яблоко, киви, сельдерей (300 мл)", "price": 550, "img_url": "/static/smoothie_green.webp"},
    {"id": 13, "title": "Смузи Клубника-банан", "category": "Чай и напитки", "description": "Клубника, банан, кокосовое молоко (300 мл)", "price": 550, "img_url": "/static/smoothie_berry.webp"},
    {"id": 14, "title": "Борщ с черносливом", "category": "Кухня и десерты", "description": "Подаётся со сметаной и гренками", "price": 420, "img_url": "/static/borsh.webp"},
    {"id": 15, "title": "Медовик с мороженым", "category": "Кухня и десерты", "description": "Фирменный десерт с шариком ванильного мороженого", "price": 350, "img_url": "/static/medowik.webp"}
]


def load_menu_items():
    conn = get_db_connection()
    if not conn:
        return DEFAULT_MENU
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, title, category, description, price, img_url FROM Menu ORDER BY id")
        rows = cursor.fetchall()
        items = []
        for r in rows:
            items.append({
                "id": r[0],
                "title": r[1],
                "category": r[2],
                "description": r[3],
                "price": r[4],
                "img_url": r[5]
            })
        cursor.close()
        conn.close()
        return items if items else DEFAULT_MENU
    except Exception as e:
        print("Ошибка чтения Menu из Access:", e)
        return DEFAULT_MENU


CART_MODAL_AND_SCRIPT = """
<!-- Модальное окно Корзины и доставки -->
<div class="modal" id="cartModal">
    <div class="modal-content" style="max-width: 480px;">
        <span class="close-btn" onclick="closeCart()">&times;</span>
        <h3 style="margin-top: 0;">Ваш заказ и доставка</h3>
        <div id="cartItemsList" style="max-height: 220px; overflow-y: auto; margin-bottom: 15px; border-bottom: 1px solid #e6d5c3; padding-bottom: 10px;">
            <p style="color: #7a6a6a;">Корзина пока пуста</p>
        </div>
        <div style="display: flex; justify-content: space-between; font-weight: bold; font-size: 1.15em; margin-bottom: 15px;">
            <span>Итого к оплате:</span>
            <span id="cartTotal">0 ₽</span>
        </div>
        <form onsubmit="submitOrder(event)">
            <label>Ваше имя (только буквы):</label>
            <input type="text" id="orderName" placeholder="Например: Анна" oninput="filterNameInput(this)" pattern="[A-Za-zА-Яа-яЁё\\s\\-]+" title="Имя не должно содержать цифры" required>
            <label>Телефон (ровно 11 цифр):</label>
            <input type="tel" id="orderPhone" placeholder="89991234567" maxlength="11" oninput="filterPhoneInput(this)" pattern="\\d{11}" title="Введите 11 цифр номера телефона" required>
            <label>Адрес доставки в Ижевске:</label>
            <input type="text" id="orderAddress" placeholder="Улица, дом, подъезд, квартира" required>
            <button type="submit" class="btn-primary" style="width: 100%; margin-top: 5px;">Оформить доставку</button>
        </form>
    </div>
</div>

<script>
    function filterNameInput(el) {
        el.value = el.value.replace(/[^A-Za-zА-Яа-яЁё\\s\\-]/g, '');
    }

    function filterPhoneInput(el) {
        el.value = el.value.replace(/\\D/g, '').slice(0, 11);
    }

    let cart = JSON.parse(localStorage.getItem('kff_cart') || '[]');

    function updateCartBadge() {
        const badge = document.getElementById('cartCount');
        if (badge) {
            badge.textContent = cart.length;
        }
    }

    function addToCart(title, price) {
        cart.push({ title: title, price: Number(price) });
        localStorage.setItem('kff_cart', JSON.stringify(cart));
        updateCartBadge();
        alert('«' + title + '» (' + price + ' ₽) добавлено в корзину!');
    }

    function removeFromCart(index) {
        cart.splice(index, 1);
        localStorage.setItem('kff_cart', JSON.stringify(cart));
        updateCartBadge();
        renderCartModal();
    }

    function renderCartModal() {
        const listEl = document.getElementById('cartItemsList');
        const totalEl = document.getElementById('cartTotal');
        if (cart.length === 0) {
            listEl.innerHTML = '<p style="color: #7a6a6a;">Корзина пока пуста. Добавьте напитки или блюда из меню.</p>';
            totalEl.textContent = '0 ₽';
            return;
        }
        let html = '';
        let sum = 0;
        cart.forEach((item, idx) => {
            sum += item.price;
            html += `
                <div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 0; border-bottom: 1px dashed #eee;">
                    <span>${item.title}</span>
                    <span>
                        <strong>${item.price} ₽</strong>
                        <button type="button" onclick="removeFromCart(${idx})" style="background: #d9534f; color: #fff; border: none; border-radius: 4px; padding: 2px 7px; margin-left: 8px; cursor: pointer; width: auto;">&times;</button>
                    </span>
                </div>
            `;
        });
        listEl.innerHTML = html;
        totalEl.textContent = sum + ' ₽';
    }

    function openCart() {
        renderCartModal();
        document.getElementById('cartModal').style.display = 'flex';
    }

    function closeCart() {
        document.getElementById('cartModal').style.display = 'none';
    }

    function submitOrder(e) {
        e.preventDefault();
        if (cart.length === 0) {
            alert('Добавьте хотя бы одну позицию в корзину!');
            return;
        }
        const name = document.getElementById('orderName').value.trim();
        const phone = document.getElementById('orderPhone').value.trim();
        const address = document.getElementById('orderAddress').value.trim();

        if (!name || /\\d/.test(name)) {
            alert('Имя должно содержать только буквы (без цифр).');
            return;
        }
        if (!/^\\d{11}$/.test(phone)) {
            alert('Номер телефона должен состоять ровно из 11 цифр (например, 89991234567).');
            return;
        }

        const total = cart.reduce((acc, item) => acc + item.price, 0);

        fetch('/order', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                name: name,
                phone: phone,
                address: address,
                items: cart,
                total: total
            })
        })
        .then(res => res.json())
        .then(data => {
            alert(data.message);
            if (data.status === 'ok') {
                cart = [];
                localStorage.setItem('kff_cart', '[]');
                updateCartBadge();
                closeCart();
            }
        });
    }

    document.addEventListener('DOMContentLoaded', updateCartBadge);
</script>
"""


HTML_PAGE = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Капучинофф — Уютная кофейня</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #faf6f2;
            color: #3b2f2f;
            line-height: 1.6;
        }
        header {
            background-color: #2c221e;
            color: #fff;
            padding: 18px 40px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            position: sticky;
            top: 0;
            z-index: 100;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
        }
        .logo a {
            color: #fff;
            text-decoration: none;
            font-size: 1.4em;
            letter-spacing: 2px;
        }
        nav { display: flex; align-items: center; flex-wrap: wrap; gap: 20px; }
        nav a {
            color: #e6d5c3;
            text-decoration: none;
            font-weight: 600;
            transition: color 0.2s;
        }
        nav a:hover { color: #fff; }
        .cart-btn {
            background-color: #c88a4b;
            color: #fff;
            border: none;
            padding: 8px 16px;
            border-radius: 20px;
            font-weight: bold;
            cursor: pointer;
            font-size: 0.95em;
        }
        .cart-btn:hover { background-color: #b07438; }

        .hero {
            background: linear-gradient(rgba(35, 25, 20, 0.65), rgba(35, 25, 20, 0.75)),
                        url('/static/hero_bg.webp') center/cover no-repeat, #3b2f2f;
            color: #fff;
            text-align: center;
            padding: 120px 20px;
        }
        .hero h1 { font-size: 2.8em; margin: 0 0 15px; letter-spacing: 1px; }
        .hero p { font-size: 1.2em; max-width: 650px; margin: 0 auto 30px; color: #f0e6dc; }
        .hero-btns { display: flex; justify-content: center; gap: 15px; flex-wrap: wrap; }
        .btn-primary {
            background-color: #c88a4b;
            color: #fff;
            padding: 14px 28px;
            border-radius: 8px;
            text-decoration: none;
            font-weight: bold;
            border: none;
            cursor: pointer;
            font-size: 1em;
            transition: background 0.2s;
        }
        .btn-primary:hover { background-color: #b07438; }
        .btn-outline {
            background-color: transparent;
            color: #fff;
            padding: 14px 28px;
            border-radius: 8px;
            text-decoration: none;
            font-weight: bold;
            border: 2px solid #e6d5c3;
            cursor: pointer;
            font-size: 1em;
            transition: all 0.2s;
        }
        .btn-outline:hover { background-color: #e6d5c3; color: #2c221e; }

        .section { padding: 60px 20px; max-width: 1100px; margin: auto; }
        .section-title { text-align: center; font-size: 2.1em; margin-bottom: 10px; color: #2c221e; }
        .section-subtitle { text-align: center; color: #7a6a6a; margin-bottom: 40px; }

        .about-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 40px;
            align-items: center;
        }
        .about-img {
            width: 100%;
            height: 340px;
            object-fit: cover;
            border-radius: 14px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
        }

        .menu-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
            gap: 25px;
        }
        .card {
            background: #fff;
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.05);
            text-align: center;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            transition: transform 0.2s;
        }
        .card:hover { transform: translateY(-5px); }
        .card img {
            width: 100%;
            height: 210px;
            object-fit: cover;
            border-radius: 8px;
            margin-bottom: 12px;
        }
        .card button {
            background-color: #6f4e37;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 1em;
            width: 100%;
            margin-top: 10px;
        }
        .card button:hover { background-color: #543b29; }

        .modal {
            display: none;
            position: fixed;
            z-index: 200;
            left: 0; top: 0; width: 100%; height: 100%;
            background: rgba(0,0,0,0.6);
            justify-content: center;
            align-items: center;
        }
        .modal-content {
            background: #fff;
            padding: 30px;
            border-radius: 12px;
            width: 90%;
            max-width: 420px;
            position: relative;
        }
        .modal-content input {
            width: 100%;
            padding: 10px;
            margin: 6px 0 14px;
            border: 1px solid #ccc;
            border-radius: 6px;
        }
        .close-btn {
            position: absolute;
            right: 15px; top: 10px;
            font-size: 1.5em;
            cursor: pointer;
            color: #888;
        }

        footer {
            background-color: #2c221e;
            color: #ccc;
            padding: 40px 20px;
            text-align: center;
        }
    </style>
</head>
<body>

<header>
    <div class="logo"><a href="/"><strong>КАПУЧИНОФФ</strong></a></div>
    <nav>
        <a href="/menu">Меню</a>
        <a href="#about">О нас</a>
        <a href="#contacts">Контакты</a>
        <button class="cart-btn" onclick="openCart()">🛒 Корзина (<span id="cartCount">0</span>)</button>
    </nav>
</header>

<div class="hero">
    <h1>Уютная кофейня в сердце города</h1>
    <p>Авторский кофе, домашняя кухня, свежая выпечка и тёплая атмосфера каждый день с 10:00 до 23:00</p>
    <div class="hero-btns">
        <a href="/menu" class="btn-primary">Открыть полное меню</a>
        <button class="btn-outline" onclick="openModal()">Забронировать столик</button>
    </div>
</div>

<div class="section" id="about">
    <div class="about-grid">
        <div>
            <h2 style="font-size: 2em; margin-top: 0;">Место, куда хочется возвращаться</h2>
            <p>«Капучинофф» — это семейная кофейня с уютной атмосферой и авторской кухней. Мы варим кофе на зерне свежей обжарки, готовим сытные завтраки весь день и печём фирменные десерты по домашним рецептам.</p>
            <p>Загляните к нам согреться облепиховым чаем или закажите быструю доставку любимых блюд прямо домой.</p>
        </div>
        <img src="/static/interior.webp" alt="Интерьер кофейни" class="about-img" onerror="this.src='https://placehold.co/600x400/e6d5c3/3b2f2f?text=Интерьер+Капучинофф'">
    </div>
</div>

<div class="section" id="menu">
    <h2 class="section-title">Хиты нашей кофейни</h2>
    <p class="section-subtitle">Самые популярные позиции среди наших гостей</p>
    <div class="menu-grid">
        {% for item in hits %}
        <div class="card">
            <div>
                <img src="{{ item.img_url }}" alt="{{ item.title }}" onerror="this.src='https://placehold.co/400x300/e6d5c3/3b2f2f?text={{ item.title }}'">
                <h3>{{ item.title }}</h3>
                <p style="font-size: 0.9em; color: #7a6a6a;">{{ item.description }}</p>
            </div>
            <div>
                <p style="font-weight: bold; font-size: 1.1em;">{{ item.price }} ₽</p>
                <button onclick="addToCart('{{ item.title }}', {{ item.price }})">Заказать</button>
            </div>
        </div>
        {% endfor %}
    </div>
    <div style="text-align: center; margin-top: 30px;">
        <a href="/menu" class="btn-primary" style="background-color: #3b2f2f;">Посмотреть все 15 позиций меню →</a>
    </div>
</div>

<footer id="contacts">
    <h3 style="color: #fff; margin-top: 0;">КАПУЧИНОФФ</h3>
    <p>Адрес: г. Ижевск, Улица Пушкина, 37</p>
    <p>Ежедневно с 10:00 до 23:00 | Телефон: +7 (968) 612-33-88</p>
</footer>

<!-- Модальное окно Бронирования столика -->
<div class="modal" id="bookModal">
    <div class="modal-content">
        <span class="close-btn" onclick="closeModal()">&times;</span>
        <h3 style="margin-top: 0;">Бронирование столика</h3>
        <form onsubmit="submitBooking(event)">
            <label>Ваше имя (только буквы):</label>
            <input type="text" id="guestName" placeholder="Например: Иван" oninput="filterNameInput(this)" pattern="[A-Za-zА-Яа-яЁё\s\-]+" title="Имя не должно содержать цифры" required>
            <label>Телефон (ровно 11 цифр):</label>
            <input type="tel" id="guestPhone" placeholder="89991234567" maxlength="11" oninput="filterPhoneInput(this)" pattern="\d{11}" title="Введите 11 цифр номера телефона" required>
            <div style="display: flex; gap: 10px;">
                <div style="flex: 1;">
                    <label>Дата визита:</label>
                    <input type="date" id="guestDate" required>
                </div>
                <div style="flex: 1;">
                    <label>Время визита:</label>
                    <input type="time" id="guestTime" required value="19:00" min="10:00" max="22:30">
                </div>
            </div>
            <button type="submit" class="btn-primary" style="width: 100%;">Подтвердить бронь</button>
        </form>
    </div>
</div>

<script>
    function openModal() {
        const dateInput = document.getElementById('guestDate');
        const today = new Date().toISOString().split('T')[0];
        dateInput.min = today;
        if (!dateInput.value) dateInput.value = today;
        document.getElementById('bookModal').style.display = 'flex';
    }
    function closeModal() { document.getElementById('bookModal').style.display = 'none'; }

    function submitBooking(e) {
        e.preventDefault();
        const name = document.getElementById('guestName').value.trim();
        const phone = document.getElementById('guestPhone').value.trim();
        const date = document.getElementById('guestDate').value;
        const time = document.getElementById('guestTime').value;

        if (!name || /\d/.test(name)) {
            alert('Имя должно содержать только буквы (без цифр).');
            return;
        }
        if (!/^\d{11}$/.test(phone)) {
            alert('Номер телефона должен состоять ровно из 11 цифр (например, 89991234567).');
            return;
        }

        fetch('/book', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name: name, phone: phone, date: date, time: time })
        })
        .then(res => res.json())
        .then(data => {
            alert(data.message);
            if (data.status === 'ok') {
                closeModal();
            }
        });
    }
</script>

""" + CART_MODAL_AND_SCRIPT + """
</body>
</html>
"""


MENU_PAGE = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Меню — Капучинофф</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #faf6f2;
            color: #3b2f2f;
        }
        header {
            background-color: #2c221e;
            color: #fff;
            padding: 18px 40px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            position: sticky;
            top: 0;
            z-index: 100;
        }
        .logo a { color: #fff; text-decoration: none; font-size: 1.4em; letter-spacing: 2px; }
        nav { display: flex; align-items: center; flex-wrap: wrap; gap: 20px; }
        nav a { color: #e6d5c3; text-decoration: none; font-weight: 600; }
        nav a:hover { color: #fff; }
        .cart-btn {
            background-color: #c88a4b;
            color: #fff;
            border: none;
            padding: 8px 16px;
            border-radius: 20px;
            font-weight: bold;
            cursor: pointer;
            font-size: 0.95em;
        }
        .section { padding: 40px 20px; max-width: 1100px; margin: auto; }
        h2.category-title {
            border-bottom: 2px solid #e6d5c3;
            padding-bottom: 8px;
            margin-top: 40px;
            margin-bottom: 20px;
            color: #543b29;
        }
        .menu-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
            gap: 20px;
        }
        .card {
            background: #fff;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
            text-align: center;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }
        .card img {
            width: 100%;
            height: 200px;
            object-fit: cover;
            border-radius: 8px;
            margin-bottom: 10px;
            background-color: #e6d5c3;
        }
        .card h3 { margin: 10px 0 5px; font-size: 1.15em; }
        .card .desc { font-size: 0.9em; color: #7a6a6a; margin: 5px 0; }
        .card .price { font-weight: bold; font-size: 1.1em; margin: 10px 0; }
        button.order-btn {
            background-color: #6f4e37;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 5px;
            cursor: pointer;
            font-size: 1em;
            width: 100%;
        }
        button.order-btn:hover { background-color: #543b29; }
        .btn-primary {
            background-color: #c88a4b;
            color: #fff;
            padding: 12px 24px;
            border-radius: 8px;
            font-weight: bold;
            border: none;
            cursor: pointer;
            font-size: 1em;
        }
        .modal {
            display: none;
            position: fixed;
            z-index: 200;
            left: 0; top: 0; width: 100%; height: 100%;
            background: rgba(0,0,0,0.6);
            justify-content: center;
            align-items: center;
        }
        .modal-content {
            background: #fff;
            padding: 30px;
            border-radius: 12px;
            width: 90%;
            max-width: 480px;
            position: relative;
        }
        .modal-content input {
            width: 100%;
            padding: 10px;
            margin: 6px 0 14px;
            border: 1px solid #ccc;
            border-radius: 6px;
        }
        .close-btn {
            position: absolute;
            right: 15px; top: 10px;
            font-size: 1.5em;
            cursor: pointer;
            color: #888;
        }
        footer {
            background-color: #2c221e;
            color: #ccc;
            text-align: center;
            padding: 30px 20px;
            margin-top: 40px;
        }
    </style>
</head>
<body>

<header>
    <div class="logo"><a href="/"><strong>КАПУЧИНОФФ</strong></a></div>
    <nav>
        <a href="/">Главная</a>
        <a href="/menu" style="color: #fff; text-decoration: underline;">Меню</a>
        <a href="/#contacts">Контакты</a>
        <button class="cart-btn" onclick="openCart()">🛒 Корзина (<span id="cartCount">0</span>)</button>
    </nav>
</header>

<div class="section">
    <h1>Полное меню кофейни</h1>

    {% for cat in ['Кофе', 'Чай и напитки', 'Кухня и десерты'] %}
    <h2 class="category-title">{{ cat }}</h2>
    <div class="menu-grid">
        {% for item in items if item.category == cat %}
        <div class="card">
            <div>
                <img src="{{ item.img_url }}" alt="{{ item.title }}" onerror="this.src='https://placehold.co/400x300/e6d5c3/3b2f2f?text={{ item.title }}'">
                <h3>{{ item.title }}</h3>
                <p class="desc">{{ item.description }}</p>
            </div>
            <div>
                <p class="price">{{ item.price }} ₽</p>
                <button class="order-btn" onclick="addToCart('{{ item.title }}', {{ item.price }})">Заказать</button>
            </div>
        </div>
        {% endfor %}
    </div>
    {% endfor %}
</div>

<footer>
    <p>Адрес: г. Ижевск, Улица Пушкина, 37 | Телефон: +7 (968) 612-33-88</p>
</footer>

""" + CART_MODAL_AND_SCRIPT + """
</body>
</html>
"""


@app.route('/')
def index():
    all_items = load_menu_items()
    hits = [i for i in all_items if i['id'] in (1, 14, 15)]
    if len(hits) < 3:
        hits = all_items[:3]
    return render_template_string(HTML_PAGE, hits=hits)


@app.route('/menu')
def menu():
    all_items = load_menu_items()
    return render_template_string(MENU_PAGE, items=all_items)


@app.route('/order', methods=['POST'])
def order():
    data = request.get_json()
    name = data.get('name', '').strip()
    phone = data.get('phone', '').strip()
    address = data.get('address', '').strip()
    items = data.get('items', [])
    total = data.get('total', 0)

    is_valid, err_msg = validate_name_and_phone(name, phone)
    if not is_valid:
        return jsonify({'status': 'error', 'message': err_msg}), 400

    items_str = ", ".join([f"{i['title']} ({i['price']} ₽)" for i in items])
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")

    print(f"Новый заказ на доставку [{now_str}]: {name}, тел: {phone}, адрес: {address} | Состав: {items_str} | Итого: {total} ₽")

    conn = get_db_connection()
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO Orders (customer_name, phone, delivery_address, items_list, total_price, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                (name, phone, address, items_str, total, now_str)
            )
            conn.commit()
            cursor.close()
            conn.close()
        except Exception as e:
            print("Примечание: не удалось записать в таблицу Orders Access (проверьте поля таблицы):", e)

    return jsonify({'status': 'ok', 'message': f'Спасибо, {name}! Ваш заказ на сумму {total} ₽ принят и будет доставлен по адресу: {address}.'})


@app.route('/book', methods=['POST'])
def book():
    data = request.get_json()
    name = data.get('name', '').strip()
    phone = data.get('phone', '').strip()
    date_raw = data.get('date', '')
    time_str = data.get('time', '')

    is_valid, err_msg = validate_name_and_phone(name, phone)
    if not is_valid:
        return jsonify({'status': 'error', 'message': err_msg}), 400

    date_formatted = '-'.join(reversed(date_raw.split('-'))) if date_raw else ''

    print(f"Новая бронь столика: {name}, тел: {phone}, дата: {date_formatted}, время: {time_str}")

    conn = get_db_connection()
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO Bookings (guest_name, phone, visit_date, visit_time) VALUES (?, ?, ?, ?)",
                (name, phone, date_formatted, time_str)
            )
            conn.commit()
            cursor.close()
            conn.close()
        except Exception as e:
            print("Примечание: не удалось записать в таблицу Bookings Access:", e)

    return jsonify({'status': 'ok', 'message': f"Спасибо, {name}! Столик забронирован на {date_formatted} в {time_str}."})


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
