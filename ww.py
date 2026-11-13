# -*- coding: utf-8 -*-
"""
KAPLAN SHOP — Telegram-бот магазина (aiogram 3.x + SQLite).

Зависимости: pip install "aiogram>=3.7,<4" openpyxl
Запуск:      python bot.py
"""

import asyncio
import html
import io
import json
import logging
import math
import os
import re
import sqlite3
import tempfile
import time
from contextlib import contextmanager
from datetime import datetime
import threading
from typing import Any, Dict, List, Mapping, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from aiogram import BaseMiddleware, Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.exceptions import TelegramBadRequest
from aiogram.filters import Command, CommandStart, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.base import BaseStorage, StorageKey
from aiogram.types import (
    BufferedInputFile,
    CallbackQuery,
    ErrorEvent,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
)

# =============================================================================
# 1. КОНФИГУРАЦИЯ
# =============================================================================

TOKEN = os.getenv("BOT_TOKEN", "9093252633:AAF_GGJn0tjtHI2OIJJVLtZaa3OON9I3llE")

# Админ 1 (главный): получает копии базы и уведомления о запуске
ADMIN_ID = 7101058884

# Админ 2 (менеджер): получает заказы и управляет складом
ADMIN_2_ID = 8730562013

# Все, кто получает заказы и имеет доступ к управлению складом
TRUSTED_IDS: List[int] = [ADMIN_ID, ADMIN_2_ID]

# Канал с товарами: "@username" или число (-100...)
CHANNEL_ID = os.getenv("CHANNEL_ID", "-1004489870357")

# База лежит РЯДОМ С ФАЙЛОМ БОТА. На хостинге можно задать DB_PATH на постоянный диск.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.getenv("DB_PATH", os.path.join(BASE_DIR, "kaplan_shop.db"))

DEFAULT_STOCK = 10          # остаток нового товара из канала
PAGE_SIZE = 8               # товаров на странице
BACKUP_EVERY_HOURS = 12     # автокопия базы админу (0 — отключить)
SHOP_NAME = "KAPLAN SHOP"

SERVICE_MARKERS = [
    "информативн",
    "публичной оферт",
    "не является оферт",
    "носит информативн",
]

# Если эти слова есть в названии/описании — клиент вписывает вкус сам
ASK_FLAVOR_KEYWORDS = [
    "жидк", "жиж", "liquid", "одноразк", "одноразов", "disposable",
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
log = logging.getLogger("kaplan_shop")

router = Router()


def is_trusted(user_id: int) -> bool:
    return user_id == ADMIN_ID or user_id in TRUSTED_IDS


def staff_ids() -> List[int]:
    return sorted({ADMIN_ID, *TRUSTED_IDS})


# =============================================================================
# 2. БАЗА ДАННЫХ
# =============================================================================

_db_lock = threading.RLock()
_conn: Optional[sqlite3.Connection] = None


def get_conn() -> sqlite3.Connection:
    global _conn
    if _conn is None:
        _conn = sqlite3.connect(DB_PATH, check_same_thread=False, isolation_level=None)
        _conn.row_factory = sqlite3.Row
        _conn.execute("PRAGMA journal_mode=WAL")
        _conn.execute("PRAGMA foreign_keys=ON")
    return _conn


def db_execute(sql: str, params: tuple = ()) -> int:
    with _db_lock:
        cur = get_conn().execute(sql, params)
        return cur.lastrowid or 0


def db_all(sql: str, params: tuple = ()) -> List[sqlite3.Row]:
    with _db_lock:
        return get_conn().execute(sql, params).fetchall()


def db_one(sql: str, params: tuple = ()) -> Optional[sqlite3.Row]:
    with _db_lock:
        return get_conn().execute(sql, params).fetchone()


@contextmanager
def transaction():
    """Атомарная транзакция: всё выполняется целиком или откатывается."""
    with _db_lock:
        conn = get_conn()
        conn.execute("BEGIN IMMEDIATE")
        try:
            yield conn
            conn.execute("COMMIT")
        except BaseException:
            if conn.in_transaction:
                conn.execute("ROLLBACK")
            raise


def _ensure_column(conn: sqlite3.Connection, table: str, column: str, ddl: str) -> None:
    cols = [r["name"] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
    if column not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {ddl}")


def init_db() -> None:
    with _db_lock:
        conn = get_conn()
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS products (
                id       INTEGER PRIMARY KEY AUTOINCREMENT,
                name     TEXT    NOT NULL,
                price    REAL    NOT NULL,
                flavors  TEXT    NOT NULL DEFAULT '',
                photo_id TEXT,
                post_id  INTEGER,
                stock    INTEGER NOT NULL DEFAULT 10,
                ask_flavor    INTEGER NOT NULL DEFAULT 0,
                manual_fields TEXT    NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS orders (
                id        INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id   INTEGER NOT NULL,
                username  TEXT,
                fullname  TEXT,
                phone     TEXT,
                address   TEXT,
                items     TEXT,
                total     REAL,
                date      TEXT,
                status    TEXT NOT NULL DEFAULT 'pending'
            );

            CREATE TABLE IF NOT EXISTS order_items (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id   INTEGER NOT NULL,
                product_id INTEGER NOT NULL,
                name       TEXT    NOT NULL DEFAULT '',
                flavor     TEXT    NOT NULL DEFAULT '',
                qty        INTEGER NOT NULL,
                price      REAL    NOT NULL,
                FOREIGN KEY (order_id) REFERENCES orders(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS cart (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id    INTEGER NOT NULL,
                product_id INTEGER NOT NULL,
                flavor     TEXT    NOT NULL DEFAULT '',
                qty        INTEGER NOT NULL DEFAULT 1,
                UNIQUE (user_id, product_id, flavor),
                FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS deleted_posts (
                post_id INTEGER PRIMARY KEY
            );

            CREATE TABLE IF NOT EXISTS verified_users (
                user_id INTEGER PRIMARY KEY,
                date    TEXT
            );

            CREATE TABLE IF NOT EXISTS fsm (
                key   TEXT PRIMARY KEY,
                state TEXT,
                data  TEXT NOT NULL DEFAULT '{}'
            );

            CREATE INDEX IF NOT EXISTS idx_products_post ON products(post_id);
            CREATE INDEX IF NOT EXISTS idx_cart_user ON cart(user_id);
            """
        )
        # Миграции для баз, созданных старыми версиями бота
        _ensure_column(conn, "products", "post_id", "INTEGER")
        _ensure_column(conn, "products", "stock", f"INTEGER NOT NULL DEFAULT {DEFAULT_STOCK}")
        _ensure_column(conn, "products", "ask_flavor", "INTEGER NOT NULL DEFAULT 0")
        _ensure_column(conn, "products", "manual_fields", "TEXT NOT NULL DEFAULT ''")
        _ensure_column(conn, "orders", "status", "TEXT NOT NULL DEFAULT 'pending'")


# ---------- возраст ----------

def is_verified(user_id: int) -> bool:
    if is_trusted(user_id):
        return True
    return db_one("SELECT 1 FROM verified_users WHERE user_id=?", (user_id,)) is not None


def set_verified(user_id: int) -> None:
    db_execute(
        "INSERT OR IGNORE INTO verified_users (user_id, date) VALUES (?, ?)",
        (user_id, datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    )


# ---------- товары ----------

def save_product(post_id: int, name: str, price: float, flavors: List[str],
                 photo_id: str, ask_flavor: bool) -> Tuple[int, bool]:
    """Создаёт/обновляет товар из поста канала. Ручные правки админа не затираются."""
    flavors_str = "\n".join(flavors)
    with _db_lock:
        if db_one("SELECT 1 FROM deleted_posts WHERE post_id=?", (post_id,)):
            return 0, False  # товар удалён админом — не воскрешаем

        row = db_one("SELECT * FROM products WHERE post_id=?", (post_id,))
        if row is None:
            for r in db_all("SELECT * FROM products"):
                if r["name"].casefold() == name.casefold():
                    row = r
                    break

        if row is not None:
            manual = {f for f in (row["manual_fields"] or "").split(",") if f}
            new_name = row["name"] if "name" in manual else name
            new_price = row["price"] if "price" in manual else price
            new_flavors = row["flavors"] if "flavors" in manual else flavors_str
            new_ask = row["ask_flavor"] if "ask" in manual else int(ask_flavor)
            db_execute(
                "UPDATE products SET name=?, price=?, flavors=?, photo_id=?, post_id=?, ask_flavor=? WHERE id=?",
                (new_name, new_price, new_flavors, photo_id, post_id, new_ask, row["id"]),
            )
            return row["id"], False

        new_id = db_execute(
            "INSERT INTO products (name, price, flavors, photo_id, post_id, stock, ask_flavor) "
            "VALUES (?,?,?,?,?,?,?)",
            (name, price, flavors_str, photo_id, post_id, DEFAULT_STOCK, int(ask_flavor)),
        )
        return new_id, True


def get_product(product_id: int) -> Optional[sqlite3.Row]:
    return db_one("SELECT * FROM products WHERE id = ?", (product_id,))


def count_products() -> int:
    row = db_one("SELECT COUNT(*) AS c FROM products")
    return int(row["c"]) if row else 0


def get_products_page(page: int) -> List[sqlite3.Row]:
    return db_all(
        "SELECT * FROM products ORDER BY id DESC LIMIT ? OFFSET ?",
        (PAGE_SIZE, page * PAGE_SIZE),
    )


def product_flavors(product: sqlite3.Row) -> List[str]:
    return [f for f in (product["flavors"] or "").split("\n") if f.strip()]


def add_manual(product_id: int, field: str) -> None:
    row = get_product(product_id)
    if not row:
        return
    fields = {f for f in (row["manual_fields"] or "").split(",") if f}
    fields.add(field)
    db_execute("UPDATE products SET manual_fields=? WHERE id=?", (",".join(sorted(fields)), product_id))


def clear_manual(product_id: int) -> None:
    db_execute("UPDATE products SET manual_fields='' WHERE id=?", (product_id,))


def update_product_stock(product_id: int, new_stock: int) -> None:
    db_execute("UPDATE products SET stock=? WHERE id=?", (max(0, new_stock), product_id))


def update_product_price(product_id: int, new_price: float) -> None:
    db_execute("UPDATE products SET price=? WHERE id=?", (max(0.0, new_price), product_id))
    add_manual(product_id, "price")


def update_product_name(product_id: int, name: str) -> None:
    db_execute("UPDATE products SET name=? WHERE id=?", (name, product_id))
    add_manual(product_id, "name")


def update_product_flavors(product_id: int, flavors: List[str]) -> None:
    db_execute("UPDATE products SET flavors=? WHERE id=?", ("\n".join(flavors), product_id))
    add_manual(product_id, "flavors")


def toggle_ask_flavor(product_id: int) -> None:
    db_execute("UPDATE products SET ask_flavor = 1 - ask_flavor WHERE id=?", (product_id,))
    add_manual(product_id, "ask")


def delete_product(product_id: int) -> None:
    with _db_lock:
        row = get_product(product_id)
        if row and row["post_id"] is not None:
            db_execute("INSERT OR IGNORE INTO deleted_posts (post_id) VALUES (?)", (row["post_id"],))
        db_execute("DELETE FROM products WHERE id=?", (product_id,))


# ---------- корзина ----------

def cart_product_total(user_id: int, product_id: int) -> int:
    row = db_one("SELECT COALESCE(SUM(qty),0) AS s FROM cart WHERE user_id=? AND product_id=?",
                 (user_id, product_id))
    return int(row["s"]) if row else 0


def available_for_user(user_id: int, product: sqlite3.Row) -> int:
    return max(0, int(product["stock"]) - cart_product_total(user_id, product["id"]))


def add_to_cart(user_id: int, product_id: int, flavor: str, qty: int) -> int:
    """Добавляет в корзину не больше остатка (суммарно по всем вкусам). Возвращает добавленное."""
    with _db_lock:
        prod = get_product(product_id)
        if not prod:
            return 0
        add = min(qty, available_for_user(user_id, prod))
        if add <= 0:
            return 0
        row = db_one("SELECT id FROM cart WHERE user_id=? AND product_id=? AND flavor=?",
                     (user_id, product_id, flavor))
        if row:
            db_execute("UPDATE cart SET qty = qty + ? WHERE id=?", (add, row["id"]))
        else:
            db_execute("INSERT INTO cart (user_id, product_id, flavor, qty) VALUES (?,?,?,?)",
                       (user_id, product_id, flavor, add))
        return add


def get_cart(user_id: int) -> List[sqlite3.Row]:
    return db_all(
        """
        SELECT c.id AS cart_id, c.product_id, c.flavor, c.qty, p.name, p.price, p.stock
        FROM cart c JOIN products p ON p.id = c.product_id
        WHERE c.user_id = ? ORDER BY c.id
        """,
        (user_id,),
    )


def change_cart_qty(user_id: int, cart_id: int, delta: int) -> bool:
    """False — если нельзя увеличить (нет остатка)."""
    with _db_lock:
        row = db_one(
            "SELECT c.id, c.qty, c.product_id, p.stock FROM cart c "
            "JOIN products p ON p.id = c.product_id WHERE c.id=? AND c.user_id=?",
            (cart_id, user_id),
        )
        if not row:
            return True
        if delta > 0 and cart_product_total(user_id, row["product_id"]) + delta > row["stock"]:
            return False
        new_qty = row["qty"] + delta
        if new_qty <= 0:
            db_execute("DELETE FROM cart WHERE id=?", (cart_id,))
        else:
            db_execute("UPDATE cart SET qty=? WHERE id=?", (new_qty, cart_id))
        return True


def delete_cart_item(user_id: int, cart_id: int) -> None:
    db_execute("DELETE FROM cart WHERE id=? AND user_id=?", (cart_id, user_id))


def clear_cart(user_id: int) -> None:
    db_execute("DELETE FROM cart WHERE user_id=?", (user_id,))


# ---------- заказы ----------

class OutOfStock(Exception):
    def __init__(self, name: str, available: int):
        super().__init__(name)
        self.name = name
        self.available = available


def create_order(user_id: int, username: str, fullname: str, phone: str,
                 address: str, date: str) -> Tuple[int, float, List[dict]]:
    """Проверка остатков + заказ + списание + очистка корзины — одной транзакцией."""
    with transaction() as conn:
        items = [dict(r) for r in conn.execute(
            "SELECT c.product_id, c.flavor, c.qty, p.name, p.price "
            "FROM cart c JOIN products p ON p.id = c.product_id "
            "WHERE c.user_id=? ORDER BY c.id", (user_id,)).fetchall()]
        if not items:
            raise ValueError("empty cart")

        need: Dict[int, int] = {}
        for it in items:
            need[it["product_id"]] = need.get(it["product_id"], 0) + it["qty"]
        for pid, qty in need.items():
            row = conn.execute("SELECT name, stock FROM products WHERE id=?", (pid,)).fetchone()
            if row["stock"] < qty:
                raise OutOfStock(row["name"], row["stock"])

        total = sum(it["price"] * it["qty"] for it in items)
        items_str = "; ".join(
            f"{it['name']}{' (' + it['flavor'] + ')' if it['flavor'] else ''} x{it['qty']} "
            f"= {fmt_money(it['price'] * it['qty'])}" for it in items
        )
        cur = conn.execute(
            "INSERT INTO orders (user_id, username, fullname, phone, address, items, total, date, status) "
            "VALUES (?,?,?,?,?,?,?,?, 'pending')",
            (user_id, username, fullname, phone, address, items_str, total, date),
        )
        order_id = cur.lastrowid
        for it in items:
            conn.execute(
                "INSERT INTO order_items (order_id, product_id, name, flavor, qty, price) VALUES (?,?,?,?,?,?)",
                (order_id, it["product_id"], it["name"], it["flavor"], it["qty"], it["price"]),
            )
        for pid, qty in need.items():
            conn.execute("UPDATE products SET stock = stock - ? WHERE id=?", (qty, pid))
        conn.execute("DELETE FROM cart WHERE user_id=?", (user_id,))
        return order_id, total, items


def confirm_order_db(order_id: int) -> bool:
    with _db_lock:
        cur = get_conn().execute(
            "UPDATE orders SET status='confirmed' WHERE id=? AND status='pending'", (order_id,))
        return cur.rowcount > 0


def fail_order_restock(order_id: int) -> bool:
    """Отмена заказа + возврат товара на склад. False — если заказ уже обработан."""
    with transaction() as conn:
        cur = conn.execute("UPDATE orders SET status='failed' WHERE id=? AND status='pending'", (order_id,))
        if cur.rowcount == 0:
            return False
        for r in conn.execute("SELECT product_id, qty FROM order_items WHERE order_id=?", (order_id,)).fetchall():
            conn.execute("UPDATE products SET stock = stock + ? WHERE id=?", (r["qty"], r["product_id"]))
        return True


# =============================================================================
# 3. ХРАНИЛИЩЕ СОСТОЯНИЙ (FSM) В SQLite — переживает перезапуск
# =============================================================================

class SQLiteStorage(BaseStorage):
    @staticmethod
    def _k(key: StorageKey) -> str:
        return f"{key.bot_id}:{key.chat_id}:{key.user_id}:{key.thread_id}:{key.destiny}"

    @staticmethod
    def _cleanup(k: str) -> None:
        db_execute("DELETE FROM fsm WHERE key=? AND state IS NULL AND data='{}'", (k,))

    async def set_state(self, key: StorageKey, state=None) -> None:
        value = state.state if isinstance(state, State) else state
        k = self._k(key)
        db_execute(
            "INSERT INTO fsm (key, state, data) VALUES (?, ?, '{}') "
            "ON CONFLICT(key) DO UPDATE SET state=excluded.state",
            (k, value),
        )
        self._cleanup(k)

    async def get_state(self, key: StorageKey) -> Optional[str]:
        row = db_one("SELECT state FROM fsm WHERE key=?", (self._k(key),))
        return row["state"] if row else None

    async def set_data(self, key: StorageKey, data: Mapping[str, Any]) -> None:
        k = self._k(key)
        db_execute(
            "INSERT INTO fsm (key, state, data) VALUES (?, NULL, ?) "
            "ON CONFLICT(key) DO UPDATE SET data=excluded.data",
            (k, json.dumps(dict(data), ensure_ascii=False)),
        )
        self._cleanup(k)

    async def get_data(self, key: StorageKey) -> Dict[str, Any]:
        row = db_one("SELECT data FROM fsm WHERE key=?", (self._k(key),))
        if not row:
            return {}
        try:
            return json.loads(row["data"] or "{}")
        except ValueError:
            return {}

    async def close(self) -> None:
        return None


# =============================================================================
# 4. РЕЗЕРВНЫЕ КОПИИ
# =============================================================================

RESTORE_TABLES = ["products", "orders", "order_items", "deleted_posts", "verified_users"]


def make_backup_bytes() -> bytes:
    fd, path = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    try:
        with _db_lock:
            dst = sqlite3.connect(path)
            try:
                get_conn().backup(dst)
            finally:
                dst.close()
        with open(path, "rb") as f:
            return f.read()
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def restore_from_bytes(data: bytes) -> Dict[str, int]:
    """Заменяет товары/заказы/подтверждения возраста данными из файла копии."""
    fd, path = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    src: Optional[sqlite3.Connection] = None
    try:
        with open(path, "wb") as f:
            f.write(data)
        src = sqlite3.connect(path)
        src.row_factory = sqlite3.Row
        tables = {r[0] for r in src.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if "products" not in tables:
            raise ValueError("в файле нет таблицы products — это не копия KAPLAN SHOP")
        counts: Dict[str, int] = {}
        with transaction() as conn:
            conn.execute("DELETE FROM cart")
            for t in reversed(RESTORE_TABLES):
                conn.execute(f"DELETE FROM {t}")
            for t in RESTORE_TABLES:
                if t not in tables:
                    continue
                src_cols = [r[1] for r in src.execute(f"PRAGMA table_info({t})")]
                dst_cols = [r["name"] for r in conn.execute(f"PRAGMA table_info({t})")]
                cols = [c for c in src_cols if c in dst_cols]
                if not cols:
                    continue
                rows = src.execute(f"SELECT {','.join(cols)} FROM {t}").fetchall()
                conn.executemany(
                    f"INSERT INTO {t} ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                    [tuple(r) for r in rows],
                )
                counts[t] = len(rows)
        return counts
    finally:
        if src is not None:
            src.close()
        try:
            os.remove(path)
        except OSError:
            pass


async def send_backup(bot: Bot, chat_id: int, reason: str) -> None:
    data = make_backup_bytes()
    name = f"kaplan_shop_backup_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.db"
    await bot.send_document(
        chat_id,
        BufferedInputFile(data, filename=name),
        caption=(
            f"💾 <b>Резервная копия базы</b> ({esc(reason)})\n"
            f"Товаров: {count_products()}\n\n"
            "Чтобы восстановить: отправьте этот файл боту с подписью <code>restore</code>."
        ),
    )


async def backup_loop(bot: Bot) -> None:
    while True:
        await asyncio.sleep(BACKUP_EVERY_HOURS * 3600)
        try:
            await send_backup(bot, ADMIN_ID, "автоматически")
        except Exception as e:
            log.error("Автобэкап не удался: %s", e)


# =============================================================================
# 5. EXCEL-ОТЧЁТ
# =============================================================================

def build_orders_excel() -> bytes:
    orders = db_all("SELECT * FROM orders WHERE status='confirmed' ORDER BY id ASC")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Подтвержденные заказы"

    headers = ["№ Заказа", "Дата и время", "Клиент", "Username", "Telegram ID",
               "Телефон", "Место/Адрес", "Состав заказа", "Сумма (₽)"]
    ws.append(headers)

    header_fill = PatternFill(start_color="2B2D42", end_color="2B2D42", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    thin = Side(style="thin", color="D3D3D3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font, cell.border = header_fill, header_font, border
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    zebra = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    for r, o in enumerate(orders, start=2):
        ws.append([
            o["id"], o["date"], o["fullname"], o["username"] or "—", o["user_id"],
            o["phone"], o["address"], o["items"], float(o["total"] or 0),
        ])
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=(c in (7, 8)))
            if r % 2 == 1:
                cell.fill = zebra
            if c == 9:
                cell.number_format = '#,##0.00 "₽"'

    total_row = len(orders) + 2
    ws.cell(row=total_row, column=8, value="ИТОГО ВЫРУЧКА:").font = Font(bold=True, size=11)
    ws.cell(row=total_row, column=8).alignment = Alignment(horizontal="right")
    sum_cell = ws.cell(row=total_row, column=9)
    sum_cell.value = float(sum(float(o["total"] or 0) for o in orders))  # число, не формула
    sum_cell.font = Font(bold=True, size=11, color="1B4332")
    sum_cell.fill = PatternFill(start_color="D8F3DC", end_color="D8F3DC", fill_type="solid")
    sum_cell.number_format = '#,##0.00 "₽"'
    sum_cell.border = border

    for i, w in enumerate([11, 20, 22, 18, 15, 18, 22, 52, 16], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =============================================================================
# 6. ПАРСЕР ПОСТОВ КАНАЛА
# =============================================================================

_NUM = r"(\d{1,3}(?:[ \u00a0\u202f]\d{3})+|\d+)(?:[.,](\d{1,2}))?"
_CUR = r"(?:₽|руб(?:лей|\.|л\.)?|р(?![A-Za-zА-Яа-яЁё0-9])|p(?![A-Za-zА-Яа-яЁё0-9])|rub)"
PRICE_SUFFIX_RE = re.compile(r"(?<![\w.,])" + _NUM + r"\s*" + _CUR, re.IGNORECASE)
PRICE_PREFIX_RE = re.compile(r"(?:₽|цена\s*[:\-–—]?)\s*" + _NUM + r"(?![\w])", re.IGNORECASE)
DASHES = "-–—−"


def _to_float(int_part: str, dec_part: Optional[str]) -> float:
    digits = re.sub(r"[ \u00a0\u202f]", "", int_part)
    return float(f"{digits}.{dec_part}") if dec_part else float(digits)


def find_price(line: str) -> Optional[Tuple[float, int, int]]:
    for regex in (PRICE_SUFFIX_RE, PRICE_PREFIX_RE):
        m = regex.search(line)
        if m:
            return _to_float(m.group(1), m.group(2)), m.start(), m.end()
    return None


def is_service_line(line: str) -> bool:
    low = line.lower()
    return any(marker in low for marker in SERVICE_MARKERS)


def _clean_name(text: str) -> str:
    text = re.sub(r"[*_`~]", "", text)
    text = re.sub(r"(?i)\bцена\b\s*[:\-–—]?", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" \t-–—:|·•,.;")


def detect_ask_flavor(text: str) -> bool:
    low = text.lower()
    return any(k in low for k in ASK_FLAVOR_KEYWORDS)


def parse_flavors_text(text: str) -> List[str]:
    out: List[str] = []
    for line in text.replace("\r", "").split("\n"):
        f = re.sub(r"\s+", " ", line.strip().lstrip(DASHES + "•* \t")).strip()
        if f and f not in out:
            out.append(f)
    return out


def parse_post(caption: Optional[str]) -> Optional[Dict[str, Any]]:
    if not caption:
        return None
    lines = [l.strip() for l in caption.replace("\r", "").split("\n")]
    lines = [l for l in lines if l and not is_service_line(l)]

    flavors: List[str] = []
    other: List[str] = []
    for line in lines:
        if line[0] in DASHES:
            flavor = re.sub(r"\s+", " ", line.lstrip(DASHES + " \t")).strip()
            if flavor and flavor not in flavors:
                flavors.append(flavor)
        else:
            other.append(line)

    price: Optional[float] = None
    name: Optional[str] = None
    for line in other:
        text = line
        found = find_price(line)
        if found and price is None:
            price, start, end = found
            text = line[:start] + " " + line[end:]
        if name is None:
            candidate = _clean_name(text)
            if candidate:
                name = candidate

    if not name or price is None or price <= 0:
        return None
    return {
        "name": name[:120],
        "price": price,
        "flavors": flavors,
        "ask_flavor": detect_ask_flavor("\n".join(other)),
    }


def is_target_channel(chat) -> bool:
    cid = str(CHANNEL_ID).strip()
    if cid.lstrip("-").isdigit():
        raw = cid.lstrip("-")
        return chat.id in (int(cid), -int(raw), int("-100" + raw))
    return bool(chat.username) and chat.username.lower() == cid.lstrip("@").lower()


async def process_channel_post(message: Message) -> None:
    if not is_target_channel(message.chat):
        log.info("Пост из чужого канала %s (%s) проигнорирован — проверьте CHANNEL_ID",
                 message.chat.id, message.chat.username)
        return
    if not message.photo or not message.caption:
        return
    parsed = parse_post(message.caption)
    if not parsed:
        log.warning("Пост %s не распознан (нет названия или цены)", message.message_id)
        return
    product_id, created = save_product(
        message.message_id, parsed["name"], parsed["price"], parsed["flavors"],
        message.photo[-1].file_id, parsed["ask_flavor"],
    )
    if product_id:
        log.info("Товар %s: #%s «%s» %s ₽ (вкус вводит клиент: %s)",
                 "добавлен" if created else "обновлён", product_id,
                 parsed["name"], parsed["price"], parsed["ask_flavor"])


@router.channel_post(F.photo)
async def on_channel_post(message: Message) -> None:
    await process_channel_post(message)


@router.edited_channel_post(F.photo)
async def on_edited_channel_post(message: Message) -> None:
    await process_channel_post(message)


# =============================================================================
# 7. ВСПОМОГАТЕЛЬНОЕ: форматирование, клавиатуры, middleware
# =============================================================================

def esc(text: Any) -> str:
    return html.escape(str(text), quote=False)


def fmt_money(value: float) -> str:
    value = float(value)
    if value.is_integer():
        return f"{int(value):,}".replace(",", " ") + " ₽"
    return f"{value:,.2f}".replace(",", " ") + " ₽"


def short(text: str, limit: int = 40) -> str:
    return text if len(text) <= limit else text[: limit - 1] + "…"


def ikb(rows: List[List[Tuple[str, str]]]) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[[InlineKeyboardButton(text=t, callback_data=d) for t, d in row] for row in rows]
    )


def main_menu_kb(user_id: int) -> InlineKeyboardMarkup:
    rows: List[List[Tuple[str, str]]] = [
        [("🛍 Открыть каталог", "catalog:0")],
        [("🛒 Корзина", "cart")],
    ]
    if is_trusted(user_id):
        rows.append([("⚙️ Управление каталогом (Склад)", "adm_stock:0")])
        rows.append([("📊 Скачать отчет Excel", "adm_excel"), ("💾 Бэкап базы", "adm_backup")])
    return ikb(rows)


def age_kb() -> InlineKeyboardMarkup:
    return ikb([[("✅ Мне исполнилось 18 лет", "age_yes")], [("❌ Нет", "age_no")]])


WELCOME_TEXT = (
    f"👋 Добро пожаловать в <b>{SHOP_NAME}</b>!\n\n"
    "🔞 Наш магазин предназначен только для совершеннолетних.\n"
    "Подтвердите, пожалуйста, что вам исполнилось 18 лет."
)
MENU_TEXT = f"🏠 <b>{SHOP_NAME}</b> — главное меню.\nВыберите действие:"


async def show(cb: CallbackQuery, text: str, markup: Optional[InlineKeyboardMarkup] = None) -> None:
    msg = cb.message
    if not isinstance(msg, Message):
        await cb.bot.send_message(cb.from_user.id, text, reply_markup=markup)
        return
    if msg.photo:
        try:
            await msg.delete()
        except TelegramBadRequest:
            pass
        await cb.bot.send_message(msg.chat.id, text, reply_markup=markup)
        return
    try:
        await msg.edit_text(text, reply_markup=markup)
    except TelegramBadRequest as e:
        if "message is not modified" in str(e).lower():
            return
        await cb.bot.send_message(msg.chat.id, text, reply_markup=markup)


async def edit_card(cb: CallbackQuery, text: str, markup: Optional[InlineKeyboardMarkup] = None) -> None:
    msg = cb.message
    if not isinstance(msg, Message):
        await cb.bot.send_message(cb.from_user.id, text, reply_markup=markup)
        return
    try:
        if msg.photo:
            await msg.edit_caption(caption=text, reply_markup=markup)
        else:
            await msg.edit_text(text, reply_markup=markup)
    except TelegramBadRequest as e:
        if "message is not modified" in str(e).lower():
            return
        await cb.bot.send_message(msg.chat.id, text, reply_markup=markup)


class AccessMiddleware(BaseMiddleware):
    """Антиспам + обязательная проверка возраста (кроме /start, /myid и кнопок возраста)."""

    def __init__(self) -> None:
        self._last: Dict[int, float] = {}

    async def __call__(self, handler, event, data):
        user = data.get("event_from_user")
        if user is None:
            return await handler(event, data)

        if isinstance(event, Message):
            if event.chat.type != "private":
                return None
            if (event.text or "").startswith(("/start", "/myid")):
                return await handler(event, data)

        if isinstance(event, CallbackQuery):
            now = time.monotonic()
            if now - self._last.get(user.id, 0.0) < 0.35:
                await event.answer()
                return None
            self._last[user.id] = now
            if event.data in ("age_yes", "age_no"):
                return await handler(event, data)

        if is_verified(user.id):
            return await handler(event, data)

        if isinstance(event, CallbackQuery):
            await event.answer("Сначала подтвердите возраст", show_alert=True)
            await event.bot.send_message(user.id, WELCOME_TEXT, reply_markup=age_kb())
        elif isinstance(event, Message):
            await event.answer(WELCOME_TEXT, reply_markup=age_kb())
        return None


access_mw = AccessMiddleware()
router.message.outer_middleware(access_mw)
router.callback_query.outer_middleware(access_mw)


# =============================================================================
# 8. СТАРТ, ВОЗРАСТ, МЕНЮ
# =============================================================================

class OrderForm(StatesGroup):
    phone = State()
    confirm = State()


class AddForm(StatesGroup):
    flavor = State()
    qty = State()


class AdminEdit(StatesGroup):
    value = State()


@router.message(CommandStart(), StateFilter("*"))
async def cmd_start(message: Message, state: FSMContext) -> None:
    await state.clear()
    try:
        tmp = await message.answer("⌨️", reply_markup=ReplyKeyboardRemove())
        await tmp.delete()
    except TelegramBadRequest:
        pass
    if is_verified(message.from_user.id):
        await message.answer(f"👋 С возвращением в <b>{SHOP_NAME}</b>!\n\n" + MENU_TEXT,
                             reply_markup=main_menu_kb(message.from_user.id))
    else:
        await message.answer(WELCOME_TEXT, reply_markup=age_kb())


@router.message(Command("myid"), StateFilter("*"))
async def cmd_myid(message: Message) -> None:
    await message.answer(f"Ваш Telegram ID: <code>{message.from_user.id}</code>")


@router.callback_query(F.data == "age_no")
async def cb_age_no(cb: CallbackQuery) -> None:
    await cb.answer()
    await show(cb, "🚫 К сожалению, доступ несовершеннолетним запрещён.", None)


@router.callback_query(F.data == "age_yes")
async def cb_age_yes(cb: CallbackQuery) -> None:
    set_verified(cb.from_user.id)
    await cb.answer()
    await show(cb, MENU_TEXT, main_menu_kb(cb.from_user.id))


@router.callback_query(F.data == "menu")
async def cb_menu(cb: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await cb.answer()
    await show(cb, MENU_TEXT, main_menu_kb(cb.from_user.id))


@router.callback_query(F.data == "noop")
async def cb_noop(cb: CallbackQuery) -> None:
    await cb.answer()


@router.message(Command("menu"), StateFilter("*"))
async def cmd_menu(message: Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer(MENU_TEXT, reply_markup=main_menu_kb(message.from_user.id))


@router.message(Command("cancel"), StateFilter("*"))
async def cmd_cancel(message: Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer("Отменено.", reply_markup=ReplyKeyboardRemove())
    await message.answer(MENU_TEXT, reply_markup=main_menu_kb(message.from_user.id))


# =============================================================================
# 9. КАТАЛОГ И КАРТОЧКА ТОВАРА
# =============================================================================

def render_catalog(page: int) -> Tuple[str, InlineKeyboardMarkup]:
    total = count_products()
    if total == 0:
        return (
            "✨ <b>Каталог пополняется!</b>\n\nСовсем скоро здесь появятся новые позиции. 📦",
            ikb([[("🛒 Корзина", "cart")], [("🏠 Меню", "menu")]]),
        )
    pages = max(1, math.ceil(total / PAGE_SIZE))
    page = max(0, min(page, pages - 1))
    rows: List[List[Tuple[str, str]]] = []
    for p in get_products_page(page):
        if p["stock"] > 0:
            label = f"{p['name']} — {fmt_money(p['price'])}"
        else:
            label = f"❌ [Нет в наличии] {p['name']}"
        rows.append([(short(label, 60), f"prod:{p['id']}")])
    if pages > 1:
        nav: List[Tuple[str, str]] = []
        if page > 0:
            nav.append(("◀️", f"catalog:{page - 1}"))
        nav.append((f"{page + 1}/{pages}", "noop"))
        if page < pages - 1:
            nav.append(("▶️", f"catalog:{page + 1}"))
        rows.append(nav)
    rows.append([("🛒 Корзина", "cart"), ("🏠 Меню", "menu")])
    return f"🛍 <b>Каталог {SHOP_NAME}</b>\nВыберите товар:", ikb(rows)


def product_card(product: sqlite3.Row) -> Tuple[str, InlineKeyboardMarkup]:
    in_stock = product["stock"] > 0
    stock_str = f"✅ В наличии: <b>{product['stock']} шт.</b>" if in_stock else "❌ <b>Товар закончился</b>"
    text = f"<b>{esc(product['name'])}</b>\n\n💰 Цена: <b>{fmt_money(product['price'])}</b>\n{stock_str}"

    flavors = product_flavors(product)
    ask = bool(product["ask_flavor"])
    buttons: List[List[Tuple[str, str]]] = []

    if ask:
        text += "\n\n✍️ <i>Вкус вы указываете сами при добавлении в корзину.</i>"
        if flavors:
            hint = "\n🍬 <b>Ориентиры по вкусам:</b>"
            for f in flavors:
                line = f"\n• {esc(f)}"
                if len(text) + len(hint) + len(line) > 850:
                    hint += "\n…"
                    break
                hint += line
            text += hint
        if in_stock:
            buttons.append([("🛒 Добавить в корзину", f"add:{product['id']}")])
    elif in_stock:
        if flavors:
            text += "\n\n🍬 <b>Выберите вкус:</b>"
            for i, f in enumerate(flavors[:40]):
                buttons.append([(short(f"🍬 {f}", 50), f"add_fl:{product['id']}:{i}")])
        else:
            buttons.append([("🛒 Добавить в корзину", f"add:{product['id']}")])

    buttons.append([("🛒 Корзина", "cart")])
    buttons.append([("⬅️ К каталогу", "catalog:0")])
    return text, ikb(buttons)


async def send_product_card(cb: CallbackQuery, product: sqlite3.Row) -> None:
    text, kb = product_card(product)
    msg = cb.message
    chat_id = msg.chat.id if isinstance(msg, Message) else cb.from_user.id

    if isinstance(msg, Message) and msg.photo:
        await edit_card(cb, text, kb)
        return
    if isinstance(msg, Message):
        try:
            await msg.delete()
        except TelegramBadRequest:
            pass
    try:
        if product["photo_id"]:
            await cb.bot.send_photo(chat_id, product["photo_id"], caption=text, reply_markup=kb)
        else:
            await cb.bot.send_message(chat_id, text, reply_markup=kb)
    except TelegramBadRequest:
        await cb.bot.send_message(chat_id, text, reply_markup=kb)


@router.callback_query(F.data.startswith("catalog:"))
async def cb_catalog(cb: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await cb.answer()
    text, kb = render_catalog(int(cb.data.split(":")[1]))
    await show(cb, text, kb)


@router.message(Command("catalog"), StateFilter("*"))
async def cmd_catalog(message: Message, state: FSMContext) -> None:
    await state.clear()
    text, kb = render_catalog(0)
    await message.answer(text, reply_markup=kb)


@router.callback_query(F.data.startswith("prod:"))
async def cb_product(cb: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    product = get_product(int(cb.data.split(":")[1]))
    if not product:
        await cb.answer("Товар не найден", show_alert=True)
        return
    await cb.answer()
    await send_product_card(cb, product)


# =============================================================================
# 10. ДОБАВЛЕНИЕ В КОРЗИНУ
# =============================================================================

def qty_kb(pid: int, available: int, mode: str, idx: int = -1) -> InlineKeyboardMarkup:
    """mode 'c' — вкус вписан клиентом (хранится в FSM), 'l' — вкус из списка/без вкуса."""
    opts = [q for q in (1, 2, 3, 4, 5, 10) if q <= available]
    rows: List[List[Tuple[str, str]]] = []
    if opts:
        rows.append([
            (str(q), f"qtyc:{pid}:{q}" if mode == "c" else f"qty:{pid}:{idx}:{q}") for q in opts
        ])
    if mode == "c":
        rows.append([("✏️ Изменить вкус", f"add:{pid}")])
    rows.append([("⬅️ Назад", f"prod:{pid}")])
    return ikb(rows)


def clean_flavor(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


@router.callback_query(F.data.startswith("add:"))
async def cb_add(cb: CallbackQuery, state: FSMContext) -> None:
    pid = int(cb.data.split(":")[1])
    product = get_product(pid)
    if not product or product["stock"] <= 0:
        await cb.answer("Товар закончился", show_alert=True)
        return
    if available_for_user(cb.from_user.id, product) <= 0:
        await cb.answer("Весь остаток уже лежит в вашей корзине", show_alert=True)
        return
    await cb.answer()
    await state.clear()
    if product["ask_flavor"]:
        await state.set_state(AddForm.flavor)
        await state.update_data(pid=pid)
        await edit_card(
            cb,
            f"✍️ <b>{esc(product['name'])}</b>\n\n"
            "Напишите, какой <b>вкус</b> вы хотите заказать (одним сообщением):",
            ikb([[("❌ Отмена", f"prod:{pid}")]]),
        )
    else:
        await edit_card(
            cb,
            f"🔢 Сколько штук <b>{esc(product['name'])}</b> добавить в корзину?",
            qty_kb(pid, available_for_user(cb.from_user.id, product), "l"),
        )


@router.message(AddForm.flavor, F.text, ~F.text.startswith("/"))
async def msg_flavor(message: Message, state: FSMContext) -> None:
    data = await state.get_data()
    product = get_product(int(data.get("pid", 0)))
    if not product or product["stock"] <= 0:
        await state.clear()
        await message.answer("Товар недоступен.", reply_markup=main_menu_kb(message.from_user.id))
        return
    flavor = clean_flavor(message.text)
    if len(flavor) < 2:
        await message.answer("⚠️ Напишите вкус текстом (минимум 2 символа).")
        return
    if len(flavor) > 60:
        await message.answer("⚠️ Слишком длинно — максимум 60 символов.")
        return
    available = available_for_user(message.from_user.id, product)
    if available <= 0:
        await state.clear()
        await message.answer("Весь остаток уже лежит в вашей корзине.", reply_markup=ikb([[("🛒 Корзина", "cart")]]))
        return
    await state.update_data(flavor=flavor)
    await state.set_state(AddForm.qty)
    await message.answer(
        f"✅ <b>{esc(product['name'])}</b>\n🍬 Вкус: <b>{esc(flavor)}</b>\n\n🔢 Выберите количество:",
        reply_markup=qty_kb(product["id"], available, "c"),
    )


@router.message(AddForm.flavor)
async def msg_flavor_wrong(message: Message) -> None:
    await message.answer("Пожалуйста, напишите вкус <b>текстом</b>.")


@router.callback_query(AddForm.qty, F.data.startswith("qtyc:"))
async def cb_qty_custom(cb: CallbackQuery, state: FSMContext) -> None:
    _, pid_s, qty_s = cb.data.split(":")
    data = await state.get_data()
    flavor = data.get("flavor", "")
    product = get_product(int(pid_s))
    if not product or product["stock"] <= 0 or not flavor:
        await state.clear()
        await cb.answer("Товар недоступен, откройте его заново", show_alert=True)
        return
    qty = int(qty_s)
    added = add_to_cart(cb.from_user.id, product["id"], flavor, qty)
    await state.clear()
    if added <= 0:
        await cb.answer("Нет в наличии", show_alert=True)
        return
    await cb.answer("Добавлено в корзину ✅")
    note = f"\n⚠️ Добавлено только {added} шт. — больше нет в наличии." if added < qty else ""
    await show(
        cb,
        f"✅ Добавлено в корзину:\n<b>{esc(product['name'])}</b>\n🍬 Вкус: {esc(flavor)}\n🔢 Количество: {added}{note}",
        ikb([[("🛒 Перейти в корзину", "cart")], [("🛍 Продолжить покупки", "catalog:0")]]),
    )


@router.callback_query(F.data.startswith("qtyc:"))
async def cb_qty_custom_stale(cb: CallbackQuery) -> None:
    await cb.answer("Сессия устарела. Откройте товар и начните заново.", show_alert=True)


@router.callback_query(F.data.startswith("add_fl:"))
async def cb_add_flavor_direct(cb: CallbackQuery) -> None:
    _, pid_s, idx_s = cb.data.split(":")
    product = get_product(int(pid_s))
    if not product or product["stock"] <= 0:
        await cb.answer("Товар закончился", show_alert=True)
        return
    flavors = product_flavors(product)
    idx = int(idx_s)
    if product["ask_flavor"] or not 0 <= idx < len(flavors):
        await cb.answer("Карточка устарела, откройте товар заново", show_alert=True)
        return
    available = available_for_user(cb.from_user.id, product)
    if available <= 0:
        await cb.answer("Весь остаток уже лежит в вашей корзине", show_alert=True)
        return
    await cb.answer()
    await edit_card(
        cb,
        f"✅ <b>{esc(product['name'])}</b>\n🍬 Вкус: <b>{esc(flavors[idx])}</b>\n\n🔢 Выберите количество:",
        qty_kb(product["id"], available, "l", idx),
    )


@router.callback_query(F.data.startswith("qty:"))
async def cb_qty(cb: CallbackQuery) -> None:
    _, pid_s, idx_s, qty_s = cb.data.split(":")
    product = get_product(int(pid_s))
    if not product or product["stock"] <= 0:
        await cb.answer("Товар закончился", show_alert=True)
        return
    if product["ask_flavor"]:
        await cb.answer("Карточка устарела, откройте товар заново", show_alert=True)
        return
    flavors = product_flavors(product)
    idx, qty = int(idx_s), int(qty_s)
    flavor = ""
    if flavors:
        if not 0 <= idx < len(flavors):
            await cb.answer("Выберите вкус заново", show_alert=True)
            return
        flavor = flavors[idx]
    added = add_to_cart(cb.from_user.id, product["id"], flavor, qty)
    if added <= 0:
        await cb.answer("Нет в наличии", show_alert=True)
        return
    await cb.answer("Добавлено в корзину ✅")
    flavor_line = f"\n🍬 Вкус: {esc(flavor)}" if flavor else ""
    note = f"\n⚠️ Добавлено только {added} шт. — больше нет в наличии." if added < qty else ""
    await show(
        cb,
        f"✅ Добавлено в корзину:\n<b>{esc(product['name'])}</b>{flavor_line}\n🔢 Количество: {added}{note}",
        ikb([[("🛒 Перейти в корзину", "cart")], [("🛍 Продолжить покупки", "catalog:0")]]),
    )


# =============================================================================
# 11. КОРЗИНА
# =============================================================================

def render_cart(user_id: int) -> Tuple[str, InlineKeyboardMarkup]:
    items = get_cart(user_id)
    if not items:
        return "🛒 Ваша корзина пуста.", ikb([[("🛍 Открыть каталог", "catalog:0")], [("🏠 Меню", "menu")]])

    lines = ["🛒 <b>Ваша корзина</b>\n"]
    total = 0.0
    rows: List[List[Tuple[str, str]]] = []
    for n, it in enumerate(items, 1):
        subtotal = it["price"] * it["qty"]
        total += subtotal
        flavor = f" ({esc(it['flavor'])})" if it["flavor"] else ""
        warn = f" ⚠️ на складе {it['stock']} шт." if it["qty"] > it["stock"] else ""
        lines.append(f"{n}. <b>{esc(it['name'])}</b>{flavor} — {it['qty']} шт. = <b>{fmt_money(subtotal)}</b>{warn}")
        rows.append([
            ("➖", f"ci:dec:{it['cart_id']}"),
            (short(f"{n}. {it['name']} x{it['qty']}", 20), "noop"),
            ("➕", f"ci:inc:{it['cart_id']}"),
            ("❌", f"ci:del:{it['cart_id']}"),
        ])
    lines.append(f"\n💰 <b>Итого: {fmt_money(total)}</b>")
    rows.append([("📦 Оформить заказ", "checkout")])
    rows.append([("🛍 В каталог", "catalog:0"), ("🗑 Очистить", "cart_clear")])
    rows.append([("🏠 Меню", "menu")])
    return "\n".join(lines), ikb(rows)


@router.callback_query(F.data == "cart")
async def cb_cart(cb: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await cb.answer()
    text, kb = render_cart(cb.from_user.id)
    await show(cb, text, kb)


@router.message(Command("cart"), StateFilter("*"))
async def cmd_cart(message: Message, state: FSMContext) -> None:
    await state.clear()
    text, kb = render_cart(message.from_user.id)
    await message.answer(text, reply_markup=kb)


@router.callback_query(F.data.startswith("ci:"))
async def cb_cart_item(cb: CallbackQuery) -> None:
    _, action, cart_id = cb.data.split(":")
    cid = int(cart_id)
    ok = True
    if action == "inc":
        ok = change_cart_qty(cb.from_user.id, cid, +1)
    elif action == "dec":
        change_cart_qty(cb.from_user.id, cid, -1)
    elif action == "del":
        delete_cart_item(cb.from_user.id, cid)
    if ok:
        await cb.answer()
    else:
        await cb.answer("Больше нет в наличии", show_alert=True)
    text, kb = render_cart(cb.from_user.id)
    await show(cb, text, kb)


@router.callback_query(F.data == "cart_clear")
async def cb_cart_clear(cb: CallbackQuery) -> None:
    clear_cart(cb.from_user.id)
    await cb.answer("Корзина очищена")
    text, kb = render_cart(cb.from_user.id)
    await show(cb, text, kb)


# =============================================================================
# 12. ОФОРМЛЕНИЕ ЗАКАЗА
# =============================================================================

CANCEL_TEXT = "❌ Отмена"


def normalize_phone(raw: str) -> str:
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 11 and digits.startswith("8"):
        digits = "7" + digits[1:]
    return "+" + digits


def phone_ok(raw: str) -> bool:
    digits = re.sub(r"\D", "", raw)
    return bool(re.fullmatch(r"\+?[\d\s\-()]{10,25}", raw.strip())) and 10 <= len(digits) <= 15


@router.callback_query(F.data == "checkout")
async def cb_checkout(cb: CallbackQuery, state: FSMContext) -> None:
    if not get_cart(cb.from_user.id):
        await cb.answer("Корзина пуста", show_alert=True)
        return
    await cb.answer()
    await state.clear()
    await state.set_state(OrderForm.phone)
    kb = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📱 Поделиться контактом", request_contact=True)],
            [KeyboardButton(text=CANCEL_TEXT)],
        ],
        resize_keyboard=True,
        one_time_keyboard=True,
        input_field_placeholder="Или введите номер: +7 999 123-45-67",
    )
    await cb.bot.send_message(
        cb.from_user.id,
        "📞 Отправьте номер телефона для связи:\nкнопкой «Поделиться контактом» или текстом.",
        reply_markup=kb,
    )


async def _go_confirm(message: Message, state: FSMContext, phone: str) -> None:
    items = get_cart(message.from_user.id)
    if not items:
        await state.clear()
        await message.answer("🛒 Корзина пуста.", reply_markup=ReplyKeyboardRemove())
        await message.answer(MENU_TEXT, reply_markup=main_menu_kb(message.from_user.id))
        return
    await state.update_data(phone=phone)
    await state.set_state(OrderForm.confirm)

    total = sum(i["price"] * i["qty"] for i in items)
    lines = ["📋 <b>Подтверждение заказа:</b>\n"]
    for it in items:
        flavor = f" ({esc(it['flavor'])})" if it["flavor"] else ""
        lines.append(f"• {esc(it['name'])}{flavor} x{it['qty']} = {fmt_money(it['price'] * it['qty'])}")
    lines.append(f"\n<b>Итого: {fmt_money(total)}</b>\n📞 Телефон: {esc(phone)}")

    await message.answer(f"✅ Телефон принят: <b>{esc(phone)}</b>", reply_markup=ReplyKeyboardRemove())
    await message.answer(
        "\n".join(lines),
        reply_markup=ikb([[("✅ Подтвердить заказ", "confirm_order")], [("❌ Отменить", "cancel_order")]]),
    )


@router.message(OrderForm.phone, F.contact)
async def get_phone_contact(message: Message, state: FSMContext) -> None:
    await _go_confirm(message, state, normalize_phone(message.contact.phone_number))


@router.message(OrderForm.phone, F.text, ~F.text.startswith("/"))
async def get_phone_text(message: Message, state: FSMContext) -> None:
    raw = message.text.strip()
    if raw == CANCEL_TEXT:
        await state.clear()
        await message.answer("Оформление отменено. Корзина сохранена.", reply_markup=ReplyKeyboardRemove())
        await message.answer(MENU_TEXT, reply_markup=main_menu_kb(message.from_user.id))
        return
    if not phone_ok(raw):
        await message.answer("⚠️ Не похоже на номер. Введите в формате <b>+7 999 123-45-67</b> "
                             "или нажмите «Поделиться контактом».")
        return
    await _go_confirm(message, state, normalize_phone(raw))


@router.message(OrderForm.phone)
async def get_phone_wrong(message: Message) -> None:
    await message.answer("Отправьте номер текстом или кнопкой «Поделиться контактом».")


@router.callback_query(F.data == "cancel_order", StateFilter("*"))
async def cb_cancel(cb: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await cb.answer("Отменено")
    await show(cb, "❌ Оформление отменено. Корзина сохранена.", main_menu_kb(cb.from_user.id))


async def notify_staff(bot: Bot, text: str, markup: Optional[InlineKeyboardMarkup]) -> None:
    for uid in staff_ids():
        try:
            await bot.send_message(uid, text, reply_markup=markup)
        except Exception as e:
            log.error("Не удалось отправить заказ пользователю %s: %s (он должен нажать /start у бота)", uid, e)


@router.callback_query(OrderForm.confirm, F.data == "confirm_order")
async def cb_confirm(cb: CallbackQuery, state: FSMContext) -> None:
    data = await state.get_data()
    await state.clear()  # защита от двойного нажатия
    user = cb.from_user
    phone = data.get("phone")
    if not phone:
        await cb.answer("Ошибка: данные потеряны", show_alert=True)
        return

    await cb.answer()
    date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    username = f"@{user.username}" if user.username else "—"
    fullname = user.full_name or "—"
    try:
        order_id, total, items = create_order(user.id, username, fullname, phone, "Самовывоз/Встреча", date)
    except OutOfStock as e:
        await show(
            cb,
            f"⚠️ Товара «{esc(e.name)}» осталось только {e.available} шт. Измените количество в корзине.",
            ikb([[("🛒 Корзина", "cart")]]),
        )
        return
    except ValueError:
        await show(cb, "🛒 Корзина пуста.", main_menu_kb(user.id))
        return

    await show(
        cb,
        f"🎉 <b>Заказ №{order_id} успешно оформлен!</b>\n\n"
        "🤝 <b>Мы с вами свяжемся для уточнения места встречи и для подтверждения заказа.</b>",
        main_menu_kb(user.id),
    )

    item_lines = []
    for it in items:
        flavor = f" (<b>{esc(it['flavor'])}</b>)" if it["flavor"] else ""
        item_lines.append(f"  • {esc(it['name'])}{flavor} — {it['qty']} шт. × {fmt_money(it['price'])}")

    pretty = (
        f"🆕 <b>НОВЫЙ ЗАКАЗ №{order_id}</b> (ожидает подтверждения)\n"
        f"━━━━━━━━━━━━━━━━━━\n"
        f"📅 <b>Дата:</b> {esc(date)}\n"
        f"👤 <b>Клиент:</b> {esc(fullname)}\n"
        f"🔗 <b>Username:</b> {esc(username)}\n"
        f"🆔 <b>ID:</b> <a href=\"tg://user?id={user.id}\">{user.id}</a>\n"
        f"📞 <b>Телефон:</b> {esc(phone)}\n"
        f"━━━━━━━━━━━━━━━━━━\n"
        f"🛒 <b>Состав:</b>\n" + "\n".join(item_lines) + "\n"
        f"━━━━━━━━━━━━━━━━━━\n"
        f"💰 <b>ИТОГО: {fmt_money(total)}</b>"
    )
    await notify_staff(
        cb.bot, pretty,
        ikb([[("✅ Заказ подтвержден", f"adm_ok:{order_id}"), ("❌ Встреча сорвалась", f"adm_fail:{order_id}")]]),
    )


@router.callback_query(F.data == "confirm_order")
async def cb_confirm_stale(cb: CallbackQuery) -> None:
    await cb.answer("Сессия устарела. Откройте корзину и оформите заказ заново.", show_alert=True)


# =============================================================================
# 13. КНОПКИ АДМИНА ПО ЗАКАЗАМ
# =============================================================================

async def _strip_buttons(cb: CallbackQuery) -> None:
    try:
        await cb.message.edit_reply_markup(reply_markup=None)
    except Exception:
        pass


@router.callback_query(F.data.startswith("adm_ok:"))
async def cb_admin_order_ok(cb: CallbackQuery) -> None:
    if not is_trusted(cb.from_user.id):
        await cb.answer("Нет доступа", show_alert=True)
        return
    order_id = int(cb.data.split(":")[1])
    if not confirm_order_db(order_id):
        await cb.answer("Заказ уже обработан", show_alert=True)
        await _strip_buttons(cb)
        return
    await cb.answer("Заказ подтверждён ✅ Готовлю Excel...")
    if isinstance(cb.message, Message):
        try:
            await cb.message.edit_text(cb.message.html_text + "\n\n🟢 <b>СТАТУС: Подтвержден</b>", reply_markup=None)
        except Exception:
            pass
    doc = BufferedInputFile(
        build_orders_excel(),
        filename=f"KAPLAN_SHOP_Orders_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx",
    )
    await cb.bot.send_document(
        cb.from_user.id, document=doc,
        caption=f"✅ <b>Заказ №{order_id} подтвержден!</b>\n📊 Таблица со всеми подтвержденными заказами и выручкой:",
    )


@router.callback_query(F.data.startswith("adm_fail:"))
async def cb_admin_order_fail(cb: CallbackQuery) -> None:
    if not is_trusted(cb.from_user.id):
        await cb.answer("Нет доступа", show_alert=True)
        return
    order_id = int(cb.data.split(":")[1])
    if not fail_order_restock(order_id):
        await cb.answer("Заказ уже обработан", show_alert=True)
        await _strip_buttons(cb)
        return
    await cb.answer("Заказ отменён, товар возвращён на склад ❌")
    if isinstance(cb.message, Message):
        try:
            await cb.message.edit_text(
                cb.message.html_text + "\n\n🔴 <b>СТАТУС: Встреча сорвалась (товар возвращён на склад)</b>",
                reply_markup=None,
            )
        except Exception:
            pass


# =============================================================================
# 14. ПАНЕЛЬ АДМИНА: СКЛАД, РЕДАКТИРОВАНИЕ, БЭКАП
# =============================================================================

def render_admin_stock(page: int) -> Tuple[str, InlineKeyboardMarkup]:
    total = count_products()
    if total == 0:
        return "📭 Каталог пока пуст.", ikb([[("🏠 Меню", "menu")]])
    pages = max(1, math.ceil(total / PAGE_SIZE))
    page = max(0, min(page, pages - 1))
    rows: List[List[Tuple[str, str]]] = []
    for p in get_products_page(page):
        status = f"[{p['stock']} шт.]" if p["stock"] > 0 else "[❌ НЕТ]"
        mark = "✍️" if p["ask_flavor"] else "🍬"
        rows.append([(short(f"{status} {mark} {p['name']} — {fmt_money(p['price'])}", 58), f"adm_item:{p['id']}")])
    if pages > 1:
        nav: List[Tuple[str, str]] = []
        if page > 0:
            nav.append(("◀️", f"adm_stock:{page - 1}"))
        nav.append((f"{page + 1}/{pages}", "noop"))
        if page < pages - 1:
            nav.append(("▶️", f"adm_stock:{page + 1}"))
        rows.append(nav)
    rows.append([("🏠 Главное меню", "menu")])
    return (
        "⚙️ <b>Управление каталогом и складом</b>\n"
        "✍️ — клиент пишет вкус сам, 🍬 — вкусы из списка.\nВыберите товар:",
        ikb(rows),
    )


def render_admin_item(product: sqlite3.Row) -> Tuple[str, InlineKeyboardMarkup]:
    stock_line = f"<b>{product['stock']} шт.</b>" if product["stock"] > 0 else "<b>❌ ЗАКОНЧИЛСЯ (0 шт.)</b>"
    flavors_preview = ", ".join(product_flavors(product)) or "нет"
    mode = "✍️ клиент пишет вкус сам" if product["ask_flavor"] else "🍬 выбор из списка вкусов"
    manual = (product["manual_fields"] or "").strip()
    text = (
        f"⚙️ <b>Редактирование товара</b>\n\n"
        f"📦 Название: <b>{esc(product['name'])}</b>\n"
        f"💰 Цена: <b>{fmt_money(product['price'])}</b>\n"
        f"📊 Остаток: {stock_line}\n"
        f"🔀 Вкус: {mode}\n"
        f"🍬 Список вкусов: {esc(flavors_preview)}"
    )
    if manual:
        text += "\n\n🔒 <i>Ручные правки защищены от перезаписи постом из канала.</i>"
    pid = product["id"]
    kb = ikb([
        [("❌ Закончился (0)", f"adm_set_st:{pid}:0"), ("🔄 Сброс (10)", f"adm_set_st:{pid}:10")],
        [("➖ 1", f"adm_adj_st:{pid}:-1"), ("➕ 1", f"adm_adj_st:{pid}:1"), ("➕ 5", f"adm_adj_st:{pid}:5")],
        [("🔢 Ввести точное кол-во", f"adm_ask:{pid}:stock")],
        [("💲 Цена", f"adm_ask:{pid}:price"), ("✏️ Название", f"adm_ask:{pid}:name")],
        [("🍬 Список вкусов", f"adm_ask:{pid}:flavors")],
        [("🔀 Вкус: клиент пишет ⇄ из списка", f"adm_toggle:{pid}")],
        [("🔄 Снова брать данные из канала", f"adm_unlock:{pid}")],
        [("🗑 Удалить товар", f"adm_del:{pid}")],
        [("⬅️ К списку товаров", "adm_stock:0")],
    ])
    return text, kb


async def _deny(cb: CallbackQuery) -> bool:
    if not is_trusted(cb.from_user.id):
        await cb.answer("Нет доступа", show_alert=True)
        return True
    return False


@router.message(Command("admin"), StateFilter("*"))
async def cmd_admin(message: Message, state: FSMContext) -> None:
    if not is_trusted(message.from_user.id):
        return
    await state.clear()
    text, kb = render_admin_stock(0)
    await message.answer(text, reply_markup=kb)


@router.callback_query(F.data == "adm_excel")
async def cb_adm_excel(cb: CallbackQuery) -> None:
    if await _deny(cb):
        return
    await cb.answer("Формирую таблицу...")
    doc = BufferedInputFile(
        build_orders_excel(),
        filename=f"KAPLAN_SHOP_Orders_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx",
    )
    await cb.bot.send_document(cb.from_user.id, document=doc,
                               caption="📊 <b>Актуальный отчет по подтвержденным заказам</b>")


@router.callback_query(F.data == "adm_backup")
async def cb_adm_backup(cb: CallbackQuery) -> None:
    if await _deny(cb):
        return
    await cb.answer("Готовлю копию...")
    await send_backup(cb.bot, cb.from_user.id, "по запросу")


@router.message(Command("backup"), StateFilter("*"))
async def cmd_backup(message: Message) -> None:
    if is_trusted(message.from_user.id):
        await send_backup(message.bot, message.from_user.id, "по запросу")


@router.message(F.document, StateFilter("*"))
async def on_document(message: Message) -> None:
    if not is_trusted(message.from_user.id):
        return
    name = (message.document.file_name or "").lower()
    if not name.endswith(".db"):
        return
    caption = (message.caption or "").lower()
    if "restore" not in caption and "восстан" not in caption:
        await message.answer(
            "Чтобы восстановить базу из этого файла, отправьте его ещё раз с подписью <code>restore</code>.\n"
            "⚠️ Текущие товары и заказы будут заменены содержимым файла."
        )
        return
    buf = io.BytesIO()
    await message.bot.download(message.document, destination=buf)
    try:
        counts = restore_from_bytes(buf.getvalue())
    except Exception as e:
        log.error("Восстановление не удалось: %s", e)
        await message.answer(f"❌ Не удалось восстановить: {esc(e)}")
        return
    await message.answer(
        "✅ <b>База восстановлена.</b>\n"
        f"Товаров: {counts.get('products', 0)}, заказов: {counts.get('orders', 0)}.",
        reply_markup=main_menu_kb(message.from_user.id),
    )


@router.callback_query(F.data.startswith("adm_stock:"))
async def cb_adm_stock(cb: CallbackQuery, state: FSMContext) -> None:
    if await _deny(cb):
        return
    await state.clear()
    await cb.answer()
    text, kb = render_admin_stock(int(cb.data.split(":")[1]))
    await show(cb, text, kb)


async def _show_item(cb: CallbackQuery, pid: int) -> None:
    prod = get_product(pid)
    if not prod:
        text, kb = render_admin_stock(0)
        await show(cb, text, kb)
        return
    text, kb = render_admin_item(prod)
    await show(cb, text, kb)


@router.callback_query(F.data.startswith("adm_item:"))
async def cb_adm_item(cb: CallbackQuery, state: FSMContext) -> None:
    if await _deny(cb):
        return
    await state.clear()
    pid = int(cb.data.split(":")[1])
    if not get_product(pid):
        await cb.answer("Товар не найден", show_alert=True)
        return
    await cb.answer()
    await _show_item(cb, pid)


@router.callback_query(F.data.startswith("adm_set_st:"))
async def cb_adm_set_stock(cb: CallbackQuery) -> None:
    if await _deny(cb):
        return
    _, pid_s, val_s = cb.data.split(":")
    update_product_stock(int(pid_s), int(val_s))
    await cb.answer(f"Остаток: {val_s} шт.")
    await _show_item(cb, int(pid_s))


@router.callback_query(F.data.startswith("adm_adj_st:"))
async def cb_adm_adj_stock(cb: CallbackQuery) -> None:
    if await _deny(cb):
        return
    _, pid_s, delta_s = cb.data.split(":")
    prod = get_product(int(pid_s))
    if not prod:
        await cb.answer("Товар не найден", show_alert=True)
        return
    new_val = max(0, prod["stock"] + int(delta_s))
    update_product_stock(prod["id"], new_val)
    await cb.answer(f"Остаток: {new_val} шт.")
    await _show_item(cb, prod["id"])


@router.callback_query(F.data.startswith("adm_toggle:"))
async def cb_adm_toggle(cb: CallbackQuery) -> None:
    if await _deny(cb):
        return
    pid = int(cb.data.split(":")[1])
    toggle_ask_flavor(pid)
    await cb.answer("Режим вкуса изменён")
    await _show_item(cb, pid)


@router.callback_query(F.data.startswith("adm_unlock:"))
async def cb_adm_unlock(cb: CallbackQuery) -> None:
    if await _deny(cb):
        return
    pid = int(cb.data.split(":")[1])
    clear_manual(pid)
    await cb.answer("Готово: при правке поста в канале данные обновятся", show_alert=True)
    await _show_item(cb, pid)


ADMIN_PROMPTS = {
    "stock": "🔢 Введите новое количество на складе (целое число):",
    "price": "💲 Введите новую цену в рублях:",
    "name": "✏️ Введите новое название товара:",
    "flavors": "🍬 Отправьте список вкусов — каждый с новой строки.\nЧтобы очистить список, отправьте <code>-</code>",
}


@router.callback_query(F.data.startswith("adm_ask:"))
async def cb_adm_ask(cb: CallbackQuery, state: FSMContext) -> None:
    if await _deny(cb):
        return
    _, pid_s, field = cb.data.split(":")
    if field not in ADMIN_PROMPTS or not get_product(int(pid_s)):
        await cb.answer("Товар не найден", show_alert=True)
        return
    await state.clear()
    await state.set_state(AdminEdit.value)
    await state.update_data(pid=int(pid_s), field=field)
    await cb.answer()
    await show(cb, ADMIN_PROMPTS[field], ikb([[("❌ Отмена", f"adm_item:{pid_s}")]]))


@router.message(AdminEdit.value, F.text, ~F.text.startswith("/"))
async def msg_adm_value(message: Message, state: FSMContext) -> None:
    if not is_trusted(message.from_user.id):
        return
    data = await state.get_data()
    pid, field = data.get("pid"), data.get("field")
    prod = get_product(pid) if pid else None
    if not prod:
        await state.clear()
        await message.answer("Товар не найден.")
        return
    raw = message.text.strip()

    if field == "stock":
        if not raw.isdigit():
            await message.answer("⚠️ Введите целое неотрицательное число.")
            return
        update_product_stock(pid, int(raw))
    elif field == "price":
        try:
            price = float(raw.replace(",", ".").replace(" ", ""))
            if price <= 0:
                raise ValueError
        except ValueError:
            await message.answer("⚠️ Введите корректную цену числом (больше нуля).")
            return
        update_product_price(pid, price)
    elif field == "name":
        if not 1 <= len(raw) <= 120:
            await message.answer("⚠️ Название должно быть от 1 до 120 символов.")
            return
        update_product_name(pid, raw)
    elif field == "flavors":
        update_product_flavors(pid, [] if raw in ("-", "нет", "—") else parse_flavors_text(message.text))
    else:
        await state.clear()
        return

    await state.clear()
    text, kb = render_admin_item(get_product(pid))
    await message.answer(f"✅ Сохранено!\n\n{text}", reply_markup=kb)


@router.message(AdminEdit.value)
async def msg_adm_value_wrong(message: Message) -> None:
    await message.answer("Отправьте значение текстом или нажмите «❌ Отмена».")


@router.callback_query(F.data.startswith("adm_del:"))
async def cb_adm_delete_ask(cb: CallbackQuery) -> None:
    if await _deny(cb):
        return
    pid = int(cb.data.split(":")[1])
    prod = get_product(pid)
    if not prod:
        await cb.answer("Товар не найден", show_alert=True)
        return
    await cb.answer()
    await show(
        cb,
        f"🗑 Точно удалить <b>{esc(prod['name'])}</b>?\nТовар исчезнет из каталога и корзин клиентов.",
        ikb([[("✅ Да, удалить", f"adm_delok:{pid}")], [("⬅️ Нет, назад", f"adm_item:{pid}")]]),
    )


@router.callback_query(F.data.startswith("adm_delok:"))
async def cb_adm_delete(cb: CallbackQuery) -> None:
    if await _deny(cb):
        return
    delete_product(int(cb.data.split(":")[1]))
    await cb.answer("🗑 Товар удалён", show_alert=True)
    text, kb = render_admin_stock(0)
    await show(cb, text, kb)


# =============================================================================
# 15. ЗАПАСНОЙ ОБРАБОТЧИК, ОШИБКИ, ЗАПУСК
# =============================================================================

@router.message(StateFilter(None), F.chat.type == "private")
async def fallback(message: Message) -> None:
    await message.answer(
        "Не совсем понял 🤔 Воспользуйтесь меню или нажмите /start.",
        reply_markup=main_menu_kb(message.from_user.id),
    )


async def on_error(event: ErrorEvent) -> bool:
    log.error("Ошибка при обработке апдейта: %s", event.exception, exc_info=event.exception)
    upd = event.update
    try:
        if upd.callback_query:
            await upd.callback_query.answer("Что-то пошло не так. Попробуйте ещё раз.", show_alert=True)
        elif upd.message:
            await upd.message.answer("⚠️ Что-то пошло не так. Нажмите /start")
    except Exception:
        pass
    return True


async def main() -> None:
    if not TOKEN or ":" not in TOKEN:
        raise SystemExit("Не задан TOKEN (в начале файла или переменная окружения BOT_TOKEN).")

    init_db()
    bot = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher(storage=SQLiteStorage())
    dp.include_router(router)
    dp.errors.register(on_error)

    me = await bot.get_me()
    log.info("Бот запущен: @%s | БД: %s | товаров: %d | канал: %s",
             me.username, DB_PATH, count_products(), CHANNEL_ID)

    # Уведомление главному админу + свежая копия базы при каждом старте
    try:
        n = count_products()
        hint = ("\n\n⚠️ Каталог пуст. Если вы ждали товары — отправьте боту файл копии "
                "<code>.db</code> с подписью <code>restore</code>.") if n == 0 else ""
        await bot.send_message(ADMIN_ID, f"✅ Бот запущен. Товаров в каталоге: <b>{n}</b>.{hint}")
        if n > 0:
            await send_backup(bot, ADMIN_ID, "при запуске")
    except Exception as e:
        log.warning("Не удалось написать админу (нажмите /start у бота): %s", e)

    backup_task = None
    if BACKUP_EVERY_HOURS > 0:
        backup_task = asyncio.create_task(backup_loop(bot))

    await bot.delete_webhook(drop_pending_updates=False)
    try:
        await dp.start_polling(bot, allowed_updates=dp.resolve_used_update_types())
    finally:
        if backup_task:
            backup_task.cancel()
        await bot.session.close()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        log.info("Бот остановлен")