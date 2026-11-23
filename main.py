import os
import re
from datetime import datetime, date
import pyodbc
from flask import Flask, render_template, request, redirect, url_for, session, flash

app = Flask(__name__)
app.secret_key = 'cappucinoff_secret_key'

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'coffee_db.accdb')


def get_db_connection():
    if not os.path.exists(DB_PATH):
        return None
    conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + DB_PATH + ';'
    return pyodbc.connect(conn_str)


def get_menu_items():
    conn = get_db_connection()
    items = []
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT id, title, category, description, price, img_url FROM Menu ORDER BY id")
            for row in cursor.fetchall():
                items.append({
                    'id': row[0],
                    'title': row[1],
                    'category': row[2],
                    'description': row[3],
                    'price': row[4],
                    'img_url': row[5]
                })
            cursor.close()
            conn.close()
        except Exception as e:
            print("Ошибка чтения меню из Access:", e)
    return items


def validate_name_and_phone(name, phone):
    if not name or not re.match(r'^[А-Яа-яЁёA-Za-z\s\-]+$', name):
        return False, "Ошибка: имя должно содержать только буквы (без цифр и символов)."
    if not phone or not re.match(r'^\d{11}$', phone):
        return False, "Ошибка: номер телефона должен состоять ровно из 11 цифр."
    return True, ""


@app.route('/')
def index():
    all_items = get_menu_items()
    hits = [i for i in all_items if any(k in i['title'] for k in ['Капучино', 'Борщ', 'Медовик'])][:3]
    if not hits:
        hits = all_items[:3]
    cart_count = len(session.get('cart', []))
    today_str = date.today().isoformat()
    return render_template('index.html', hits=hits, cart_count=cart_count, today=today_str)


@app.route('/menu')
def menu():
    all_items = get_menu_items()
    categories = {}
    for item in all_items:
        cat = item['category']
        categories.setdefault(cat, []).append(item)
    cart_count = len(session.get('cart', []))
    return render_template('menu.html', categories=categories, cart_count=cart_count)


@app.route('/add_to_cart', methods=['POST'])
def add_to_cart():
    title = request.form.get('title')
    price = int(request.form.get('price', 0))
    cart = session.get('cart', [])
    cart.append({'title': title, 'price': price})
    session['cart'] = cart
    flash(f'«{title}» добавлен в корзину!')
    return redirect(request.referrer or url_for('menu'))


@app.route('/remove_from_cart/<int:index>', methods=['POST'])
def remove_from_cart(index):
    cart = session.get('cart', [])
    if 0 <= index < len(cart):
        cart.pop(index)
        session['cart'] = cart
    return redirect(url_for('cart_page'))


@app.route('/cart')
def cart_page():
    cart = session.get('cart', [])
    total = sum(item['price'] for item in cart)
    return render_template('cart.html', cart=cart, total=total, cart_count=len(cart))


@app.route('/order', methods=['POST'])
def order():
    cart = session.get('cart', [])
    if not cart:
        flash('Корзина пуста! Добавьте хотя бы одну позицию.')
        return redirect(url_for('cart_page'))

    name = request.form.get('name', '').strip()
    phone = request.form.get('phone', '').strip()
    address = request.form.get('address', '').strip()

    is_valid, err_msg = validate_name_and_phone(name, phone)
    if not is_valid:
        flash(err_msg)
        return redirect(url_for('cart_page'))

    total = sum(item['price'] for item in cart)
    items_str = ", ".join([f"{i['title']} ({i['price']} ₽)" for i in cart])
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")

    conn = get_db_connection()
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO Orders (customer_name, phone, delivery_address, items_list, total_price, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                (name, phone, address, items_str, total, now_str)
            )
            cursor.execute("SELECT @@IDENTITY")
            new_order_id = int(cursor.fetchone()[0])

            for item in cart:
                cursor.execute("SELECT id FROM Menu WHERE title = ?", (item['title'],))
                row = cursor.fetchone()
                if row:
                    cursor.execute(
                        "INSERT INTO [Order] (order_id, menu_id, quantity) VALUES (?, ?, ?)",
                        (new_order_id, int(row[0]), 1)
                    )
            conn.commit()
            cursor.close()
            conn.close()
        except Exception as e:
            print("Ошибка записи заказа в Access:", e)

    session['cart'] = []
    flash(f'Спасибо, {name}! Ваш заказ на сумму {total} ₽ принят и будет доставлен по адресу: {address}.')
    return redirect(url_for('index'))


@app.route('/book', methods=['POST'])
def book():
    name = request.form.get('name', '').strip()
    phone = request.form.get('phone', '').strip()
    visit_date = request.form.get('date', '').strip()
    visit_time = request.form.get('time', '').strip()

    is_valid, err_msg = validate_name_and_phone(name, phone)
    if not is_valid:
        flash(err_msg)
        return redirect(url_for('index') + '#bookModal')

    conn = get_db_connection()
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO Bookings (guest_name, phone, visit_date, visit_time) VALUES (?, ?, ?, ?)",
                (name, phone, visit_date, visit_time)
            )
            conn.commit()
            cursor.close()
            conn.close()
        except Exception as e:
            print("Ошибка записи брони в Access:", e)

    formatted_date = "-".join(reversed(visit_date.split("-"))) if "-" in visit_date else visit_date
    flash(f'Спасибо, {name}! Столик забронирован на {formatted_date} в {visit_time}.')
    return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=True)
