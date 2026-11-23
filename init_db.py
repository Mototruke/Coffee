import pyodbc
import os

db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'coffee_db.accdb')
conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path + ';'

conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE Orders (
    id AUTOINCREMENT PRIMARY KEY,
    customer_name TEXT(100),
    phone TEXT(50),
    delivery_address TEXT(255),
    items_list LONGTEXT,
    total_price INTEGER,
    created_at TEXT(50)
)
""")

conn.commit()
cursor.close()
conn.close()
print("УСПЕХ! Новая таблица Orders создана.")


# 2. Обновлённая таблица Заказов с доставкой (используем LONGTEXT)
cursor.execute("""
CREATE TABLE Orders (
    id AUTOINCREMENT PRIMARY KEY,
    customer_name TEXT(100),
    phone TEXT(50),
    delivery_address TEXT(255),
    items_list LONGTEXT,
    total_price INTEGER,
    created_at TEXT(50)
)
""")

conn.commit()
cursor.close()
conn.close()
print("УСПЕХ! Новая таблица Orders создана и готова принимать заказы с доставкой.")

# 3. Таблица Бронирований
cursor.execute("""
CREATE TABLE Bookings (
    id AUTOINCREMENT PRIMARY KEY,
    guest_name TEXT(100),
    phone TEXT(50),
    visit_date TEXT(30),
    visit_time TEXT(20)
)
""")

# Заполняем 15 позиций меню
items = [
    ("Капучино (300 мл)", "Кофе", "Классический двойной эспрессо с молочной пенкой", 400, "/static/kapus.webp"),
    ("Латте (300 мл)", "Кофе", "Нежный кофейный напиток с большим количеством молока", 400, "/static/latte.webp"),
    ("Раф Кофе (300 мл)", "Кофе", "С оттенками ванили, карамели или шоколада", 450, "/static/raf.webp"),
    ("Бамбл Кофе (300 мл)", "Кофе", "Освежающий микс эспрессо, апельсинового сока и карамели", 400, "/static/bumble.webp"),
    ("Айс Матча «Малина-кокос»", "Кофе", "Сезонная новинка (500 мл)", 500, "/static/matcha_malina.webp"),
    ("Айс латте «Сникерс»", "Кофе", "С шоколадной пенкой и джус боллами (500 мл)", 500, "/static/snickers.webp"),
    ("Раф Лимонный пирог", "Кофе", "Авторский кофе со вкусом домашней выпечки (300 мл)", 450, "/static/raf_lemon.webp"),
    ("Латте Пряная груша", "Кофе", "Авторский согревающий латте (300 мл)", 450, "/static/latte_pear.webp"),
    ("Облепиховый чай (700 мл)", "Чай и напитки", "С мёдом, имбирём и корицей", 600, "/static/tea_oblepiha.webp"),
    ("Яблочный сбитень (700 мл)", "Чай и напитки", "Традиционный пряный горячий напиток", 600, "/static/sbiten.webp"),
    ("Лимонад Манго-апельсин", "Чай и напитки", "Освежающий домашний лимонад (500 мл)", 500, "/static/limonad.webp"),
    ("Смузи Зеленый микс", "Чай и напитки", "Яблоко, киви, сельдерей (300 мл)", 550, "/static/smoothie_green.webp"),
    ("Смузи Клубника-банан", "Чай и напитки", "Клубника, банан, кокосовое молоко (300 мл)", 550, "/static/smoothie_berry.webp"),
    ("Борщ с черносливом", "Кухня и десерты", "Подаётся со сметаной и гренками", 420, "/static/borsh.webp"),
    ("Медовик с мороженым", "Кухня и десерты", "Фирменный десерт с шариком ванильного мороженого", 350, "/static/medowik.webp")
]

for item in items:
    cursor.execute(
        "INSERT INTO Menu (title, category, description, price, img_url) VALUES (?, ?, ?, ?, ?)",
        item
    )

conn.commit()
cursor.close()
conn.close()
print("База Access успешно обновлена! Таблица Orders готова принимать заказы с доставкой.")
