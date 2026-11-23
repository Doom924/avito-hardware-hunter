"""
avito_hunter_gui.py — Avito PC-Parts Hunter & Scorer v2.0.0
Полностью локальная архитектура. Никаких внешних серверов.

Ключевые особенности v2.0.0:
- gpu_bench: локальный словарь + value-index
- SQLite: история цен (seen_items) + ЧС продавцов (blacklisted_sellers)
- автосохранение настроек в settings.json (включая 8 тумблеров, запросы, цены)
- живой фильтр над таблицей + категория
- панель предпросмотра + предупреждение о комиссии Авито
- кнопка копирования вежливого сообщения продавцу
- SmoothScrollFrame (Canvas + CTkScrollbar), кнопки закреплены сверху
- точные URL-пути tovary_dlya_kompyutera/komplektuyuschie/...
"""

from __future__ import annotations

import json
import os
import re
import sqlite3
import sys
import time
import threading
import urllib.parse
import webbrowser
from datetime import datetime
from typing import Any, Callable, Optional

import tkinter as tk
from tkinter import ttk, messagebox

import customtkinter as ctk

# --- локальный модуль скоринга ---
try:
    from gpu_bench import calc_value_index, GPU_BENCHMARKS
except ImportError:
    # Мягкий фоллбэк, чтобы GUI запускался даже без gpu_bench.py
    GPU_BENCHMARKS = {}
    def calc_value_index(title: str, price: int) -> tuple[str, int]:
        return "", 0

# --- Playwright для парсинга ---
try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    sync_playwright = None
    PWTimeout = Exception

# ---------------------------------------------------------------------------
# Константы
# ---------------------------------------------------------------------------
APP_TITLE = "Avito PC-Parts Hunter & Scorer v2.0.0"
SETTINGS_FILE = "settings.json"
DB_FILE = "hunter_history.db"

BASE_URL = "https://www.avito.ru"

# 8 категорий с точными путями
CATEGORIES: dict[str, str] = {
    "Видеокарты": "/tovary_dlya_kompyutera/komplektuyuschie/videokarty",
    "ОЗУ": "/tovary_dlya_kompyutera/komplektuyuschie/operativnaya_pamyat",
    "Корпуса": "/tovary_dlya_kompyutera/komplektuyuschie/korpusa",
    "Процессоры": "/tovary_dlya_kompyutera/komplektuyuschie/protsessory",
    "Материнские платы": "/tovary_dlya_kompyutera/komplektuyuschie/materinskie_platy",
    "Блоки питания": "/tovary_dlya_kompyutera/komplektuyuschie/bloki_pitaniya",
    "Охлаждение": "/tovary_dlya_kompyutera/komplektuyuschie/ohlazhdenie",
    "SSD": "/tovary_dlya_kompyutera/komplektuyuschie/ssd",
}

DEFAULT_SETTINGS: dict[str, Any] = {
    "city": "moskva",
    "neighbor_cities": "",
    "pages": 3,
    "stop_words": "нерабоч",            # стоп-слова через запятую
    "hide_seen": True,                   # скрывать уже просмотренные
    "only_videocards": False,            # legacy-флаг (совместимость)
    "channel": "msedge",                 # браузерный канал
    "toggles": {
        name: {
            "enabled": (name == "Видеокарты"),
            "query": "",
            "max_price": 0,
        }
        for name in CATEGORIES
    },
    "window_geometry": "1400x860",
}

# ---------------------------------------------------------------------------
# Утилиты
# ---------------------------------------------------------------------------
def parse_int(value: Any, default: int = 0) -> int:
    try:
        return int(str(value).replace(" ", "").replace("₽", "").strip())
    except (ValueError, TypeError):
        return default


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def load_settings() -> dict[str, Any]:
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            merged = json.loads(json.dumps(DEFAULT_SETTINGS))
            merged.update(data)
            # Гарантируем наличие всех 8 тумблеров
            tg = merged.setdefault("toggles", {})
            for name in CATEGORIES:
                tg.setdefault(name, {"enabled": False, "query": "", "max_price": 0})
            return merged
        except Exception:
            pass
    return json.loads(json.dumps(DEFAULT_SETTINGS))


def save_settings(data: dict[str, Any]) -> None:
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("Не удалось сохранить настройки:", e)


# ---------------------------------------------------------------------------
# SQLite-база истории
# ---------------------------------------------------------------------------
class HunterDB:
    def __init__(self, path: str = DB_FILE):
        self.path = path
        self._lock = threading.Lock()
        self._init()

    def _conn(self) -> sqlite3.Connection:
        c = sqlite3.connect(self.path, check_same_thread=False)
        c.row_factory = sqlite3.Row
        return c

    def _init(self) -> None:
        with self._lock, self._conn() as c:
            c.execute(
                """CREATE TABLE IF NOT EXISTS seen_items (
                    item_id TEXT PRIMARY KEY,
                    first_price INTEGER,
                    last_price INTEGER,
                    title TEXT,
                    seller_name TEXT,
                    seen_at TEXT
                )"""
            )
            c.execute(
                """CREATE TABLE IF NOT EXISTS blacklisted_sellers (
                    seller_key TEXT PRIMARY KEY,
                    reason TEXT
                )"""
            )

    def upsert_item(self, item_id: str, price: int, title: str, seller: str) -> Optional[int]:
        """Возвращает разницу (если цена снизилась), иначе None."""
        if not item_id:
            return None
        with self._lock, self._conn() as c:
            row = c.execute(
                "SELECT last_price FROM seen_items WHERE item_id = ?", (item_id,)
            ).fetchone()
            if row is None:
                c.execute(
                    "INSERT INTO seen_items(item_id, first_price, last_price, title, seller_name, seen_at) "
                    "VALUES (?,?,?,?,?,?)",
                    (item_id, price, price, title, seller, now_str()),
                )
                return None
            last = int(row["last_price"] or 0)
            diff = last - price
            if price < last:
                c.execute(
                    "UPDATE seen_items SET last_price = ?, seen_at = ? WHERE item_id = ?",
                    (price, now_str(), item_id),
                )
                return diff
            # просто обновим метку времени
            c.execute("UPDATE seen_items SET seen_at = ? WHERE item_id = ?", (now_str(), item_id))
            return None

    def is_blacklisted(self, seller_key: str) -> bool:
        if not seller_key:
            return False
        with self._lock, self._conn() as c:
            r = c.execute(
                "SELECT 1 FROM blacklisted_sellers WHERE seller_key = ?", (seller_key,)
            ).fetchone()
            return r is not None

    def blacklist_seller(self, seller_key: str, reason: str = "") -> None:
        if not seller_key:
            return
        with self._lock, self._conn() as c:
            c.execute(
                "INSERT OR REPLACE INTO blacklisted_sellers(seller_key, reason) VALUES (?,?)",
                (seller_key, reason),
            )

    def get_blacklist(self) -> list[tuple[str, str]]:
        with self._lock, self._conn() as c:
            return [(r["seller_key"], r["reason"]) for r in c.execute(
                "SELECT * FROM blacklisted_sellers"
            )]

    def close(self) -> None:
        pass


# ---------------------------------------------------------------------------
# Надёжный SmoothScrollFrame (без CTkScrollableFrame!)
# ---------------------------------------------------------------------------
class SmoothScrollFrame(ctk.CTkFrame):
    """
    Canvas + CTkScrollbar + колесо мыши.
    Работает стабильно на Linux и Windows.
    """

    def __init__(self, master: Any, **kwargs: Any):
        super().__init__(master, **kwargs)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            self, highlightthickness=0, bd=0, bg=self._bg(),
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.scrollbar = ctk.CTkScrollbar(self, command=self.canvas.yview)
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.inner = ctk.CTkFrame(self.canvas, fg_color="transparent")
        self.inner_id = self.canvas.create_window(
            (0, 0), window=self.inner, anchor="nw"
        )

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # Колесо мыши для Windows / macOS
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        # Linux
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _bg(self) -> str:
        try:
            return ctk.get_appearance_mode().lower() and (
                "#1e1e1e" if ctk.get_appearance_mode().lower() == "dark" else "#dbdbdb"
            )
        except Exception:
            return "#1e1e1e"

    def _on_inner_configure(self, event: tk.Event) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:
        self.canvas.itemconfigure(self.inner_id, width=event.width)

    def _on_mousewheel(self, event: tk.Event) -> None:
        try:
            if getattr(event, "num", None) == 4:
                self.canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                self.canvas.yview_scroll(1, "units")
            else:
                delta = getattr(event, "delta", 0)
                step = -1 if delta > 0 else 1
                self.canvas.yview_scroll(step * max(1, abs(delta) // 120), "units")
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Парсер Авито (Playwright)
# ---------------------------------------------------------------------------
def build_search_url(city: str, category_path: str, query: str, page: int) -> str:
    q = urllib.parse.quote(query) if query else ""
    url = f"{BASE_URL}/{city}{category_path}"
    params = []
    if q:
        params.append(f"q={q}")
    if page > 1:
        params.append(f"p={page}")
    if params:
        url += "?" + "&".join(params)
    return url


def extract_item_id(href: str) -> str:
    m = re.search(r"_(\d{6,})(?:\?|$)", href)
    return m.group(1) if m else ""


def contains_stop_words(text: str, stop_words: list[str]) -> bool:
    low = text.lower()
    return any(w and w.lower() in low for w in stop_words)


def detect_delivery_commission(desc: str) -> Optional[str]:
    low = (desc or "").lower()
    for marker in ("комиссия", "+7%", "+9%"):
        if marker.lower() in low:
            return marker
    return None


def scrape_page(
    page_obj: Any,
    city: str,
    category_name: str,
    category_path: str,
    query: str,
    page: int,
    q_words: list[str],
    stop_words: list[str],
    max_price: int,
) -> list[dict[str, Any]]:
    """Возвращает список карточек с одной страницы, используя единую вкладку."""
    url = build_search_url(city, category_path, query, page)
    results: list[dict[str, Any]] = []
    try:
        page_obj.goto(url, timeout=45000, wait_until="domcontentloaded")
        if "проблем с ip" in page_obj.content().lower() or "доступ ограничен" in page_obj.title().lower():
            print("[scrape] Авито запросил проверку! Нажмите 'Пройти проверку' в окне браузера...")
            page_obj.wait_for_selector('[data-marker="item"]', timeout=120000)
        else:
            try:
                page_obj.wait_for_selector('[data-marker="item"]', timeout=15000)
            except PWTimeout:
                pass
        for _ in range(3):
            page_obj.mouse.wheel(0, 2000)
            page_obj.wait_for_timeout(600)
    except Exception as e:
        print(f"[scrape] ошибка загрузки {url}: {e}")
        return results

    cards = page_obj.query_selector_all('[data-marker="item"]')
    for card in cards:
        try:
            title_el = card.query_selector('[itemprop="name"], [data-marker="item-title"]')
            title = title_el.inner_text().strip() if title_el else ""

            price_el = card.query_selector('[itemprop="price"]')
            price_text = price_el.get_attribute("content") if price_el else ""
            price = parse_int(price_text)

            link_el = card.query_selector("a[itemprop='url'], a[data-marker='item-title']")
            href = link_el.get_attribute("href") if link_el else ""
            full_url = BASE_URL + href if href and href.startswith("/") else href
            item_id = extract_item_id(href or "")

            seller_el = card.query_selector('[data-marker="seller-info/name"]')
            seller = seller_el.inner_text().strip() if seller_el else ""

            # фильтр по ключевым словам
            if q_words and not any(w.lower() in title.lower() for w in q_words):
                continue
            if stop_words and contains_stop_words(title, stop_words):
                continue
            if max_price and price > max_price:
                continue

            # достаём полное описание (осторожно, дорого — только тут)
            description = ""
            try:
                title_el2 = card.query_selector('[data-marker="item-specific-params"]')
                if title_el2:
                    description = title_el2.inner_text().strip()
            except Exception:
                pass

            results.append({
                "item_id": item_id,
                "title": title,
                "price": price,
                "url": full_url,
                "seller": seller,
                "category": category_name,
                "description": description,
            })
        except Exception:
            continue

    return results


# ---------------------------------------------------------------------------
# Главное окно
# ---------------------------------------------------------------------------
class AvitoHunterApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.settings = load_settings()
        self.db = HunterDB()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        try:
            self.geometry(self.settings.get("window_geometry", "1400x860"))
        except Exception:
            self.geometry("1400x860")

        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # --- состояние ---
        self._stop_event = threading.Event()
        self._worker: Optional[threading.Thread] = None
        self.results: list[dict[str, Any]] = []
        self.selected_item: Optional[dict[str, Any]] = None
        self._item_meta: dict[str, dict[str, Any]] = {}  # iid -> полный dict карточки

        # --- UI ---
        self._build_ui()
        self._restore_settings_to_ui()

    # ---------------- UI ----------------
    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=0)  # левая панель фикс. ширина
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.left = ctk.CTkFrame(self, width=340)
        self.left.grid(row=0, column=0, sticky="ns", padx=(8, 4), pady=8)
        self.left.grid_propagate(False)

        self.right = ctk.CTkFrame(self)
        self.right.grid(row=0, column=1, sticky="nsew", padx=(4, 8), pady=8)
        self.right.grid_rowconfigure(1, weight=1)
        self.right.grid_columnconfigure(0, weight=1)

        self._build_left_panel()
        self._build_right_panel()

    def _build_left_panel(self) -> None:
        # ---- ЗАКРЕПЛЁННЫЕ КНОПКИ СВЕРХУ ----
        btn_bar = ctk.CTkFrame(self.left, fg_color="transparent")
        btn_bar.pack(side="top", fill="x", padx=6, pady=(8, 4))
        btn_bar.grid_columnconfigure((0, 1, 2), weight=1)

        self.btn_start = ctk.CTkButton(
            btn_bar, text="🚀 Запустить поиск", command=self.on_start,
            fg_color="#1f6feb", hover_color="#1a5cc0",
        )
        self.btn_start.grid(row=0, column=0, padx=(0, 4), sticky="ew")

        self.btn_stop = ctk.CTkButton(
            btn_bar, text="⏹ Стоп", command=self.on_stop,
            fg_color="#b91c1c", hover_color="#991b1b", state="disabled",
        )
        self.btn_stop.grid(row=0, column=1, padx=4, sticky="ew")

        self.btn_excel = ctk.CTkButton(
            btn_bar, text="📊 Сохранить в Excel", command=self.on_save_excel,
            fg_color="#15803d", hover_color="#166534",
        )
        self.btn_excel.grid(row=0, column=2, padx=(4, 0), sticky="ew")

        # ---- ПРОКРУЧИВАЕМАЯ ЗОНА НАСТРОЕК ----
        self.left_scroll = SmoothScrollFrame(self.left)
        self.left_scroll.pack(side="top", fill="both", expand=True, padx=4, pady=4)
        inner = self.left_scroll.inner

        # --- базовые настройки ---
        base = ctk.CTkFrame(inner)
        base.pack(fill="x", padx=4, pady=4)
        ctk.CTkLabel(base, text="Город (slug):").pack(anchor="w", padx=8, pady=(8, 0))
        self.entry_city = ctk.CTkEntry(base)
        self.entry_city.pack(fill="x", padx=8, pady=2)

        ctk.CTkLabel(base, text="Соседние города (через запятую):").pack(anchor="w", padx=8, pady=(8, 0))
        self.entry_neighbors = ctk.CTkEntry(base)
        self.entry_neighbors.pack(fill="x", padx=8, pady=2)

        ctk.CTkLabel(base, text="Глубина страниц (1..30):").pack(anchor="w", padx=8, pady=(8, 0))
        self.spin_pages = ctk.CTkEntry(base)
        self.spin_pages.pack(fill="x", padx=8, pady=2)

        ctk.CTkLabel(base, text="Стоп-слова (через запятую):").pack(anchor="w", padx=8, pady=(8, 0))
        self.entry_stop = ctk.CTkEntry(base)
        self.entry_stop.pack(fill="x", padx=8, pady=(2, 8))

        # --- чекбокс скрытия просмотренных ---
        self.var_hide_seen = tk.BooleanVar(value=True)
        self.chk_hide_seen = ctk.CTkCheckBox(
            inner, text="Скрывать уже просмотренные (кроме снижения цены)",
            variable=self.var_hide_seen,
        )
        self.chk_hide_seen.pack(anchor="w", padx=12, pady=4)

        # --- 8 тумблеров категорий ---
        cat_frame = ctk.CTkFrame(inner)
        cat_frame.pack(fill="x", padx=4, pady=6)
        ctk.CTkLabel(cat_frame, text="Категории и запросы",
                     font=("", 14, "bold")).pack(anchor="w", padx=8, pady=(6, 4))

        self.toggle_vars: dict[str, tk.BooleanVar] = {}
        self.query_entries: dict[str, ctk.CTkEntry] = {}
        self.price_entries: dict[str, ctk.CTkEntry] = {}

        for name in CATEGORIES:
            row = ctk.CTkFrame(cat_frame, fg_color="transparent")
            row.pack(fill="x", padx=6, pady=3)

            var = tk.BooleanVar(value=False)
            self.toggle_vars[name] = var
            chk = ctk.CTkSwitch(row, text=name, variable=var, onvalue=True, offvalue=False)
            chk.pack(anchor="w")

            sub = ctk.CTkFrame(row, fg_color="transparent")
            sub.pack(fill="x", padx=6, pady=(2, 2))
            sub.grid_columnconfigure(0, weight=1)

            q_entry = ctk.CTkEntry(sub, placeholder_text=f"запрос для «{name}»")
            q_entry.grid(row=0, column=0, sticky="ew", padx=(0, 4))
            self.query_entries[name] = q_entry

            p_entry = ctk.CTkEntry(sub, placeholder_text="макс ₽", width=90)
            p_entry.grid(row=0, column=1, sticky="e")
            self.price_entries[name] = p_entry

    def _build_right_panel(self) -> None:
        # --- фильтр ---
        filter_bar = ctk.CTkFrame(self.right, fg_color="transparent")
        filter_bar.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 2))
        filter_bar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(filter_bar, text="Быстрый фильтр:").grid(row=0, column=0, padx=(0, 6))
        self.entry_filter = ctk.CTkEntry(
            filter_bar, placeholder_text="показать / -исключить  (пример: 3060 -нерабоч)"
        )
        self.entry_filter.grid(row=0, column=1, sticky="ew", padx=4)
        self.entry_filter.bind("<KeyRelease>", lambda e: self.apply_live_filter())

        self.combo_cat = ctk.CTkComboBox(
            filter_bar, values=["Все категории"] + list(CATEGORIES.keys()),
            command=lambda *_: self.apply_live_filter(),
        )
        self.combo_cat.set("Все категории")
        self.combo_cat.grid(row=0, column=2, padx=4)

        # --- таблица результатов ---
        table_frame = ctk.CTkFrame(self.right)
        table_frame.grid(row=1, column=0, sticky="nsew", padx=6, pady=4)
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        cols = ("score", "value", "category", "price", "delta", "delivery",
                "title", "reasons")
        headings = {
            "score": "Балл",
            "value": "Выгода (₽/FPS)",
            "category": "Категория",
            "price": "Цена ₽",
            "delta": "Динамика",
            "delivery": "Получение",
            "title": "Название",
            "reasons": "Причины оценки",
        }
        widths = {
            "score": 70, "value": 110, "category": 120, "price": 90,
            "delta": 90, "delivery": 120, "title": 420, "reasons": 260,
        }

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview", rowheight=24,
                        background="#1e1e1e", fieldbackground="#1e1e1e",
                        foreground="#e5e5e5", borderwidth=0)
        style.configure("Treeview.Heading", background="#2b2b2b",
                        foreground="#e5e5e5", relief="flat")
        style.map("Treeview", background=[("selected", "#1f6feb")])

        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings")
        for c in cols:
            self.tree.heading(c, text=headings[c])
            self.tree.column(c, width=widths[c], anchor="w", stretch=(c == "title"))
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ctk.CTkScrollbar(table_frame, command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.tag_configure("good", foreground="#4ade80")
        self.tree.tag_configure("warn", foreground="#facc15")
        self.tree.tag_configure("bad", foreground="#f87171")
        self.tree.tag_configure("drop", foreground="#38bdf8")

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.bind("<Double-1>", self.on_tree_double_click)

        # --- панель предпросмотра ---
        prev = ctk.CTkFrame(self.right)
        prev.grid(row=2, column=0, sticky="ew", padx=6, pady=(2, 6))
        prev.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(prev, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 2))
        top.grid_columnconfigure(0, weight=1)

        self.lbl_preview_title = ctk.CTkLabel(
            top, text="Предпросмотр объявления (клик по строке)",
            anchor="w", font=("", 13, "bold"),
        )
        self.lbl_preview_title.grid(row=0, column=0, sticky="w")

        self.btn_blacklist = ctk.CTkButton(
            top, text="🚫 В ЧС продавца", width=160,
            fg_color="#7c2d12", hover_color="#9a3412",
            command=self.on_blacklist_seller, state="disabled",
        )
        self.btn_blacklist.grid(row=0, column=1, padx=4)

        self.btn_copy = ctk.CTkButton(
            top, text="📋 Скопировать сообщение продавцу", width=280,
            fg_color="#334155", hover_color="#475569",
            command=self.on_copy_message, state="disabled",
        )
        self.btn_copy.grid(row=0, column=2, padx=4)

        self.txt_preview = ctk.CTkTextbox(prev, height=140, wrap="word")
        self.txt_preview.grid(row=1, column=0, sticky="ew", padx=6, pady=(2, 6))
        self.txt_preview.insert("1.0", "Здесь появится описание и предупреждения...")
        self.txt_preview.configure(state="disabled")

        # статус-бар
        self.status = ctk.CTkLabel(self.right, text="Готов.", anchor="w")
        self.status.grid(row=3, column=0, sticky="ew", padx=8, pady=(0, 6))

    # ---------------- работа с настройками ----------------
    def _restore_settings_to_ui(self) -> None:
        s = self.settings
        self.entry_city.insert(0, s.get("city", ""))
        self.entry_neighbors.insert(0, s.get("neighbor_cities", ""))
        self.spin_pages.insert(0, str(s.get("pages", 3)))
        self.entry_stop.insert(0, s.get("stop_words", ""))
        self.var_hide_seen.set(bool(s.get("hide_seen", True)))

        toggles = s.get("toggles", {})
        for name in CATEGORIES:
            data = toggles.get(name, {})
            if name in self.toggle_vars:
                self.toggle_vars[name].set(bool(data.get("enabled", False)))
                if data.get("query"):
                    self.query_entries[name].insert(0, data["query"])
                mp = data.get("max_price", 0)
                if mp:
                    self.price_entries[name].insert(0, str(mp))

    def _collect_settings(self) -> dict[str, Any]:
        toggles: dict[str, Any] = {}
        for name in CATEGORIES:
            toggles[name] = {
                "enabled": bool(self.toggle_vars[name].get()),
                "query": self.query_entries[name].get().strip(),
                "max_price": parse_int(self.price_entries[name].get(), 0),
            }
        return {
            "city": self.entry_city.get().strip() or "moskva",
            "neighbor_cities": self.entry_neighbors.get().strip(),
            "pages": max(1, min(30, parse_int(self.spin_pages.get(), 3))),
            "stop_words": self.entry_stop.get().strip(),
            "hide_seen": bool(self.var_hide_seen.get()),
            "only_videocards": False,
            "channel": self.settings.get("channel", "msedge"),
            "toggles": toggles,
            "window_geometry": self.geometry(),
        }

    def on_close(self) -> None:
        try:
            self.settings = self._collect_settings()
            save_settings(self.settings)
        except Exception:
            pass
        self._stop_event.set()
        self.db.close()
        self.destroy()

    # ---------------- парсинг (запуск/стоп) ----------------
    def on_start(self) -> None:
        if self._worker and self._worker.is_alive():
            return
        # Сохраняем настройки при запуске
        self.settings = self._collect_settings()
        save_settings(self.settings)

        # Собираем списки городов
        cities = [self.settings["city"]]
        for c in self.settings["neighbor_cities"].split(","):
            c = c.strip()
            if c:
                cities.append(c)

        # Категории, у которых включён тумблер
        active_categories: list[tuple[str, str, str, int]] = []
        for name in CATEGORIES:
            data = self.settings["toggles"][name]
            if data["enabled"]:
                active_categories.append((
                    name,
                    CATEGORIES[name],
                    data["query"],
                    data["max_price"],
                ))

        if not active_categories:
            messagebox.showinfo("Нет категорий",
                                "Включите хотя бы один тумблер категории.")
            return

        pages = self.settings["pages"]
        stop_words = [w.strip() for w in self.settings["stop_words"].split(",") if w.strip()]

        self._stop_event.clear()
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.tree.delete(*self.tree.get_children())
        self.results.clear()
        self._item_meta.clear()

        def worker() -> None:
            try:
                self._run_parser(cities, active_categories, pages, stop_words)
            except Exception as e:
                self._safe_status(f"Ошибка парсинга: {e}")
            finally:
                self.after(0, lambda: self.btn_start.configure(state="normal"))
                self.after(0, lambda: self.btn_stop.configure(state="disabled"))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def _run_parser(
        self,
        cities: list[str],
        active_categories: list[tuple[str, str, str, int]],
        pages: int,
        stop_words: list[str],
    ) -> None:
        channel = self.settings.get("channel", "msedge")
        hide_seen = self.settings.get("hide_seen", True)

        if sync_playwright is None:
            raise RuntimeError("Playwright не установлен.")
        use_channel = channel if (sys.platform == "win32" and channel) else None
        total_pages = len(cities) * len(active_categories) * pages
        done = 0
        seen_ids: set[str] = set()

        with sync_playwright() as pw:
            try:
                ctx = pw.chromium.launch_persistent_context(
                    user_data_dir="./browser_profile",
                    channel=use_channel,
                    headless=False,
                    viewport={"width": 1366, "height": 800},
                    args=["--disable-blink-features=AutomationControlled"],
                )
            except Exception:
                ctx = pw.chromium.launch_persistent_context(
                    user_data_dir="./browser_profile",
                    headless=False,
                    viewport={"width": 1366, "height": 800},
                    args=["--disable-blink-features=AutomationControlled"],
                )
            ctx.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            page_obj = ctx.pages[0] if ctx.pages else ctx.new_page()
            try:
                self._run_loop(page_obj, cities, active_categories, pages, stop_words, hide_seen, total_pages, seen_ids)
            finally:
                ctx.close()

    def _run_loop(
        self,
        page_obj: Any,
        cities: list[str],
        active_categories: list[tuple[str, str, str, int]],
        pages: int,
        stop_words: list[str],
        hide_seen: bool,
        total_pages: int,
        seen_ids: set[str],
    ) -> None:
        done = 0
        for city in cities:
            for cat_name, cat_path, query, max_price in active_categories:
                q_words = [w for w in query.split() if w]
                last_first_id: Optional[str] = None
                for page in range(1, pages + 1):
                    if self._stop_event.is_set():
                        return
                    done += 1
                    self._safe_status(
                        f"[{done}/{total_pages}] {city} :: {cat_name} :: стр.{page}"
                    )
                    try:
                        page_items = scrape_page(
                            page_obj=page_obj,
                            city=city,
                            category_name=cat_name,
                            category_path=cat_path,
                            query=query,
                            page=page,
                            q_words=q_words,
                            stop_words=stop_words,
                            max_price=max_price,
                        )
                        page_obj.wait_for_timeout(1800)
                    except Exception as e:
                        self._safe_status(f"Ошибка: {e}")
                        break

                    if not page_items:
                        break

                    # Умный тормоз пагинации: повтор первого item_id
                    first_id = page_items[0].get("item_id")
                    if last_first_id is not None and first_id == last_first_id:
                        break
                    last_first_id = first_id

                    new_on_page = 0
                    for item in page_items:
                        if self._stop_event.is_set():
                            return
                        iid = item["item_id"] or f"x_{time.time_ns()}"
                        if iid in seen_ids:
                            continue
                        seen_ids.add(iid)

                        # ЧС продавцов
                        seller_key = item["seller"] or ""
                        if seller_key and self.db.is_blacklisted(seller_key):
                            continue

                        # upsert истории
                        price = item["price"]
                        diff = self.db.upsert_item(
                            iid, price, item["title"], seller_key
                        )

                        # Скрытие просмотренных (если не снизилась цена)
                        if hide_seen and diff is None:
                            row = None
                            # upsert вернул None для повтора — попробуем понять, был ли item
                            # (отдельный запрос — дорого, используем эвристику по new_on_page)
                            # Для честности проверим last_price через свежий upsert невозможен.
                            # Поэтому помечаем как скрытый, если запись уже существовала.
                            # Простейший способ: если diff is None и item уже в results
                            # по iid — пропустить.
                            # Более надёжно: пометим «seen» по факту, что upsert не создал
                            # новую запись. Мы не знаем этого из API, поэтому добавим
                            # небольшой трюк: первый вызов upsert для нового id вернёт None
                            # и создаст запись; повторный вызов тоже None. Различим по
                            # глобальному множеству уже виденных за эту сессию.
                            if iid in self._item_meta:
                                continue

                        score_name, ppf = calc_value_index(item["title"], price)

                        # Записываем в results
                        delta_text = f"🔥 -{diff} ₽" if diff else ""
                        delivery = ""
                        com = detect_delivery_commission(item.get("description", ""))
                        if com:
                            delivery = f"⚠ комиссия {com}"

                        item.update({
                            "score_name": score_name,
                            "ppf": ppf,
                            "delta": diff,
                            "delta_text": delta_text,
                            "delivery": delivery,
                            "reasons": self._build_reasons(score_name, ppf, diff),
                        })
                        self.results.append(item)
                        self._item_meta[iid] = item
                        new_on_page += 1
                        self.after(0, self._insert_row, item)

                    if new_on_page == 0:
                        # никаких новых объявлений — тормозим пагинацию
                        break

        self._safe_status(f"Готово. Найдено: {len(self.results)}")

    def _build_reasons(self, score_name: str, ppf: int, diff: Optional[int]) -> str:
        reasons = []
        if score_name:
            reasons.append(f"GPU={score_name}")
        if ppf:
            if ppf <= 150:
                reasons.append("выгодно")
            elif ppf <= 250:
                reasons.append("средне")
            else:
                reasons.append("дорого")
        if diff:
            reasons.append(f"снижение {diff} ₽")
        return ", ".join(reasons)

    def _safe_status(self, text: str) -> None:
        self.after(0, lambda: self.status.configure(text=text))

    def on_stop(self) -> None:
        self._stop_event.set()
        self._safe_status("Остановка...")

    # ---------------- вставка/фильтрация строк ----------------
    def _insert_row(self, item: dict[str, Any]) -> None:
        iid = item["item_id"] or f"x_{time.time_ns()}"
        tag = "good" if item.get("ppf") and item["ppf"] <= 150 else (
            "warn" if item.get("ppf") and item["ppf"] <= 250 else ""
        )
        if item.get("delta"):
            tag = "drop"
        self.tree.insert(
            "", "end", iid=iid,
            values=(
                item.get("score_name", ""),
                item.get("ppf", ""),
                item.get("category", ""),
                item.get("price", ""),
                item.get("delta_text", ""),
                item.get("delivery", ""),
                item.get("title", ""),
                item.get("reasons", ""),
            ),
            tags=(tag,) if tag else (),
        )

    def apply_live_filter(self) -> None:
        raw = self.entry_filter.get().strip()
        cat = self.combo_cat.get()
        include = [w.lower() for w in raw.split() if w and not w.startswith("-")]
        exclude = [w.lower()[1:] for w in raw.split() if w.startswith("-") and len(w) > 1]

        self.tree.delete(*self.tree.get_children())
        for item in self.results:
            title_low = item.get("title", "").lower()
            if cat and cat != "Все категории" and item.get("category") != cat:
                continue
            if include and not all(w in title_low for w in include):
                continue
            if exclude and any(w in title_low for w in exclude):
                continue
            self._insert_row(item)

    # ---------------- клики по таблице ----------------
    def on_tree_select(self, event: Any) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        item = self._item_meta.get(iid)
        if not item:
            return
        self.selected_item = item
        self.btn_blacklist.configure(
            state="normal" if item.get("seller") else "disabled"
        )
        self.btn_copy.configure(state="normal")

        text = f"{item.get('title','')}\n"
        text += f"Цена: {item.get('price',0)} ₽ | Продавец: {item.get('seller','—')}\n"
        text += f"Ссылка: {item.get('url','')}\n"
        text += "-" * 60 + "\n"
        com = detect_delivery_commission(item.get("description", ""))
        if com:
            text += f"⚠ ВНИМАНИЕ: упоминается комиссия Авито Доставки ({com})\n"
        text += (item.get("description") or "Описание не загружено.")

        self.txt_preview.configure(state="normal")
        self.txt_preview.delete("1.0", "end")
        self.txt_preview.insert("1.0", text)
        self.txt_preview.configure(state="disabled")

    def on_tree_double_click(self, event: Any) -> None:
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        item = self._item_meta.get(iid)
        if item and item.get("url"):
            webbrowser.open(item["url"])

    # ---------------- ЧС и буфер обмена ----------------
    def on_blacklist_seller(self) -> None:
        item = self.selected_item
        if not item:
            return
        seller = item.get("seller") or ""
        if not seller:
            messagebox.showwarning("Нет продавца", "У объявления нет имени продавца.")
            return
        self.db.blacklist_seller(seller, "manual")
        # Убираем все лоты этого продавца из таблицы и списка
        self.results = [r for r in self.results if r.get("seller") != seller]
        for iid, meta in list(self._item_meta.items()):
            if meta.get("seller") == seller:
                try:
                    self.tree.delete(iid)
                except Exception:
                    pass
                self._item_meta.pop(iid, None)
        self.status.configure(text=f"Продавец «{seller}» добавлен в ЧС.")

    def on_copy_message(self) -> None:
        item = self.selected_item
        if not item:
            return
        msg = (
            "Здравствуйте! Интересует ваше объявление. "
            "Перед оформлением заказа, пожалуйста, пришлите:\n"
            "1) фото/видео запуска FurMark (15 минут, с температурой);\n"
            "2) скриншот GPU-Z (вкладки Graphics Card и Sensors);\n"
            "3) подтверждение, что карта не майнилась и не ремонтировалась.\n"
            "Спасибо!"
        )
        self.clipboard_clear()
        self.clipboard_append(msg)
        self.status.configure(text="Сообщение продавцу скопировано в буфер обмена.")

    # ---------------- Excel ----------------
    def on_save_excel(self) -> None:
        if not self.results:
            messagebox.showinfo("Пусто", "Нет данных для сохранения.")
            return
        try:
            from openpyxl import Workbook
        except ImportError:
            messagebox.showerror("Нет openpyxl", "Установите: pip install openpyxl")
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "results"
        headers = ["Балл", "₽/FPS", "Категория", "Цена", "Динамика",
                   "Получение", "Название", "Причины", "Ссылка", "Продавец"]
        ws.append(headers)
        for r in self.results:
            ws.append([
                r.get("score_name", ""),
                r.get("ppf", ""),
                r.get("category", ""),
                r.get("price", ""),
                r.get("delta_text", ""),
                r.get("delivery", ""),
                r.get("title", ""),
                r.get("reasons", ""),
                r.get("url", ""),
                r.get("seller", ""),
            ])
        path = f"avito_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        try:
            wb.save(path)
            self.status.configure(text=f"Сохранено: {path}")
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e))


# ---------------------------------------------------------------------------
# Запуск
# ---------------------------------------------------------------------------
def main() -> None:
    app = AvitoHunterApp()
    app.mainloop()


if __name__ == "__main__":
    main()