"""
Avito Hunter — десктопное приложение для поиска и скоринга
б/у комплектующих ПК на Авито.

Стек: CustomTkinter (dark theme) + Playwright (async, в отдельном потоке) + openpyxl.
"""

from __future__ import annotations

import asyncio
import os
import queue
import re
import sys
import threading
import time
import traceback
import webbrowser
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    from playwright.async_api import async_playwright, BrowserContext, Page
except ImportError:  # позволит UI запуститься даже без Playwright
    async_playwright = None  # type: ignore
    BrowserContext = Any  # type: ignore
    Page = Any  # type: ignore

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None  # type: ignore


# =============================================================================
#  КОНСТАНТЫ
# =============================================================================

APP_TITLE = "Avito Hunter — поиск и скоринг комплектующих"
APP_GEOMETRY = "1400x900"
BROWSER_PROFILE_DIR = "./browser_profile"
AVITO_BASE = "https://www.avito.ru"

DEFAULT_STOP_WORDS = (
    "запчаст, нерабоч, сломан, артефакт, коробк, ошибка 43, "
    "треугольник, дымок, отвал, только обмен, без претензий"
)

CATEGORIES: list[dict[str, str]] = [
    {
        "key": "gpu",
        "label": "Видеокарты",
        "query_ph": "Напр: rx 580 8gb",
        "url_path": "tovary_dlya_kompyutera/komplektuyuschie/videokarty",
    },
    {
        "key": "ram",
        "label": "Оперативная память",
        "query_ph": "Напр: ddr4 16gb",
        "url_path": "tovary_dlya_kompyutera/pamyat",
    },
    {
        "key": "case",
        "label": "Корпуса",
        "query_ph": "Напр: корпус atx",
        "url_path": "tovary_dlya_kompyutera/komplektuyuschie/korpusa",
    },
    {
        "key": "cpu",
        "label": "Процессоры",
        "query_ph": "Напр: ryzen 5 5600",
        "url_path": "tovary_dlya_kompyutera/komplektuyuschie_dlya_pk",
    },
    {
        "key": "mb",
        "label": "Материнские платы",
        "query_ph": "Напр: b450 am4",
        "url_path": "tovary_dlya_kompyutera/komplektuyuschie/materinskie_platy",
    },
    {
        "key": "psu",
        "label": "Блоки питания",
        "query_ph": "Напр: бп 600w",
        "url_path": "tovary_dlya_kompyutera/komplektuyuschie_dlya_pk",
    },
]

# Маркеры в описании/названии
POSITIVE_MARKERS = [
    "тест", "проверен", "проверено", "гарантия", "пломб", "чек", "рабоч",
    "рабочий", "работает", "без нареканий", "как новый",
]
RISK_MARKERS = [
    "рефаб", "refurb", "на запчасти", "не работает", "нерабоч",
    "не запускается", "битый", "сломан", "артефакт", "ошибка 43",
    "треугольник", "только обмен", "без претензий",
]


# =============================================================================
#  МОДЕЛИ ДАННЫХ
# =============================================================================

@dataclass
class SearchConfig:
    city_slug: str
    nearby_cities: list[str]
    stop_words: list[str]
    filter_booking_and_rating: bool
    categories: list[dict[str, Any]] = field(default_factory=list)
    # каждая категория: {"key","label","query","max_price","url_path"}


@dataclass
class Listing:
    title: str
    url: str
    price: Optional[int]
    category: str
    delivery: str          # "Самовывоз" / "Доставка" / "—"
    seller_rating: Optional[float]
    seller_reviews: Optional[int]
    is_booked: bool
    score: int = 0
    reasons: list[str] = field(default_factory=list)


# =============================================================================
#  УТИЛИТЫ
# =============================================================================

def parse_price(raw: str) -> Optional[int]:
    """'12 500 ₽' -> 12500; пусто/None -> None."""
    if not raw:
        return None
    digits = re.sub(r"[^\d]", "", raw)
    return int(digits) if digits else None


def parse_reviews(raw: str) -> Optional[int]:
    if not raw:
        return None
    digits = re.sub(r"[^\d]", "", raw)
    return int(digits) if digits else None


def parse_rating(raw: str) -> Optional[float]:
    if not raw:
        return None
    m = re.search(r"\d[.,]?\d*", raw)
    if not m:
        return None
    try:
        return float(m.group(0).replace(",", "."))
    except ValueError:
        return None


def matches_stop_words(text: str, stop_words: list[str]) -> bool:
    if not stop_words:
        return False
    low = text.lower()
    return any(w and w in low for w in stop_words)


# =============================================================================
#  СКОРИНГ
# =============================================================================

def score_listing(listing: Listing) -> None:
    """
    Рассчитываем 0..100 баллов надёжности продавца/объявления.
    Учитываем: самовывоз/доставку, рейтинг продавца, кол-во отзывов,
    маркеры тестов/пломб и маркеры риска/рефаба.
    """
    score = 50.0
    reasons: list[str] = []

    # --- Получение ---
    d = (listing.delivery or "").lower()
    if "самовывоз" in d:
        score += 10
        reasons.append("самовывоз (+10)")
    elif "доставка" in d:
        score += 4
        reasons.append("доставка (+4)")

    # --- Рейтинг продавца ---
    r = listing.seller_rating
    if r is not None:
        if r >= 4.8:
            score += 18
            reasons.append(f"рейтинг {r} (+18)")
        elif r >= 4.5:
            score += 12
            reasons.append(f"рейтинг {r} (+12)")
        elif r >= 4.0:
            score += 4
            reasons.append(f"рейтинг {r} (+4)")
        else:
            score -= 15
            reasons.append(f"низкий рейтинг {r} (-15)")

    # --- Кол-во отзывов ---
    n = listing.seller_reviews
    if n is not None:
        if n >= 50:
            score += 12
            reasons.append(f"отзывов {n} (+12)")
        elif n >= 10:
            score += 7
            reasons.append(f"отзывов {n} (+7)")
        elif n >= 1:
            score += 3
            reasons.append(f"отзывов {n} (+3)")

    # --- Маркеры в тексте ---
    text = (listing.title or "").lower()
    for m in POSITIVE_MARKERS:
        if m in text:
            score += 3
            reasons.append(f"маркер «{m}» (+3)")
            break
    for m in RISK_MARKERS:
        if m in text:
            score -= 12
            reasons.append(f"риск «{m}» (-12)")
            break

    # --- Бронь ---
    if listing.is_booked:
        score -= 20
        reasons.append("бронь (-20)")

    # --- Клэмп ---
    score = max(0, min(100, round(score)))
    listing.score = score
    listing.reasons = reasons or ["—"]


# =============================================================================
#  ПАРСЕР (Playwright, async)
# =============================================================================

class AvitoParser:
    """
    Асинхронный парсер. Запускается в отдельном потоке,
    результаты складывает в очередь (queue.Queue) для UI.
    """

    def __init__(
        self,
        config: SearchConfig,
        out_queue: "queue.Queue[tuple[str, Any]]",
        stop_event: threading.Event,
    ) -> None:
        self.config = config
        self.q = out_queue
        self.stop_event = stop_event

    # ---------- публичный запуск ----------
    def run(self) -> None:
        try:
            asyncio.run(self._run_async())
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            self.q.put(("error", f"{exc}\n\n{tb}"))
        finally:
            self.q.put(("done", None))

    # ---------- внутренняя асинхронная часть ----------
    async def _run_async(self) -> None:
        if async_playwright is None:
            self.q.put(
                ("error", "Playwright не установлен. Выполните:\n  pip install playwright\n  playwright install")
            )
            return

        async with async_playwright() as pw:
            context: BrowserContext = await self._launch(pw)
            try:
                page = await context.new_page()
                total_cats = len(self.config.categories)
                for idx, cat in enumerate(self.config.categories, 1):
                    if self.stop_event.is_set():
                        break
                    self.q.put(("log", f"[{idx}/{total_cats}] Категория: {cat['label']}"))
                    try:
                        await self._parse_category(page, cat)
                    except Exception as exc:  # noqa: BLE001
                        self.q.put(("log", f"  ! Ошибка в категории «{cat['label']}»: {exc}"))
                    await asyncio.sleep(0.5)
            finally:
                try:
                    await context.close()
                except Exception:  # noqa: BLE001
                    pass

    async def _launch(self, pw) -> BrowserContext:
        """Пытаемся Edge на Windows, иначе Chromium. Профиль общий."""
        profile_path = str(Path(BROWSER_PROFILE_DIR).resolve())
        Path(profile_path).mkdir(parents=True, exist_ok=True)

        common = dict(
            user_data_dir=profile_path,
            headless=False,
            viewport={"width": 1366, "height": 900},
            locale="ru-RU",
            args=["--disable-blink-features=AutomationControlled"],
        )

        if sys.platform == "win32":
            try:
                ctx = await pw.chromium.launch_persistent_context(channel="msedge", **common)
                self.q.put(("log", "Браузер: системный Edge (msedge)"))
                return ctx
            except Exception:  # noqa: BLE001
                self.q.put(("log", "Edge недоступен, откат на Chromium..."))

        ctx = await pw.chromium.launch_persistent_context(**common)
        self.q.put(("log", "Браузер: Chromium (Playwright)"))
        return ctx

    # ---------- сборка URL ----------
    def _build_url(self, cat: dict[str, Any], page_num: int = 1) -> str:
        slug = (self.config.city_slug or "rossiya").strip()
        path = cat["url_path"]
        url = f"{AVITO_BASE}/{slug}/{path}"
        params: list[str] = []
        q = (cat.get("query") or "").strip()
        if q:
            params.append(f"q={q.replace(' ', '+')}")
        max_price = cat.get("max_price")
        if max_price:
            params.append(f"pmax={max_price}")
        if page_num > 1:
            params.append(f"p={page_num}")
        if params:
            url += "?" + "&".join(params)
        return url

    # ---------- парсинг одной категории ----------
    async def _parse_category(self, page: Page, cat: dict[str, Any]) -> None:
        url = self._build_url(cat)
        self.q.put(("log", f"  → {url}"))
        await page.goto(url, wait_until="domcontentloaded", timeout=60_000)

        # Авито иногда просит "я не робот" — даём время
        try:
            await page.wait_for_selector('[data-marker="item"]', timeout=20_000)
        except Exception:
            self.q.put(
                ("log", "  ! Не дождались объявлений (возможно, капча/антибот). "
                        "Пройдите капчу в открывшемся окне браузера.")
            )
            return

        # Прокрутка для подгрузки
        for _ in range(3):
            if self.stop_event.is_set():
                return
            await page.mouse.wheel(0, 2000)
            await asyncio.sleep(0.7)

        cards = await page.query_selector_all('[data-marker="item"]')
        self.q.put(("log", f"  Найдено карточек: {len(cards)}"))

        parsed = 0
        for card in cards:
            if self.stop_event.is_set():
                return
            try:
                listing = await self._parse_card(card, cat)
            except Exception:  # noqa: BLE001
                continue
            if listing is None:
                continue

            # --- стоп-слова ---
            if matches_stop_words(listing.title, self.config.stop_words):
                continue

            # --- фильтр брони/рейтинга ---
            if self.config.filter_booking_and_rating:
                if listing.is_booked:
                    continue
                if listing.seller_rating is not None and listing.seller_rating < 4.5:
                    continue

            score_listing(listing)
            self.q.put(("row", listing))
            parsed += 1
            await asyncio.sleep(0.02)

        self.q.put(("log", f"  ✓ Прошло фильтры: {parsed}"))

    # ---------- парсинг одной карточки ----------
    async def _parse_card(self, card, cat: dict[str, Any]) -> Optional[Listing]:
        title_el = await card.query_selector('[data-marker="item-title"]')
        if not title_el:
            return None
        title = (await title_el.inner_text()).strip()
        rel = await title_el.get_attribute("href") or ""
        if rel.startswith("/"):
            url = AVITO_BASE + rel
        elif rel.startswith("http"):
            url = rel
        else:
            return None

        price_el = await card.query_selector('[data-marker="item-price"]')
        price = parse_price(await price_el.inner_text()) if price_el else None

        # самовывоз / доставка
        delivery = "—"
        body = (await card.inner_text()).lower()
        if "самовывоз" in body:
            delivery = "Самовывоз"
        elif "доставка" in body:
            delivery = "Доставка"

        # бронь (по наличию соответствующего бейджа)
        is_booked = "бронь" in body or "забронировано" in body

        # рейтинг и отзывы продавца (могут отсутствовать)
        seller_rating: Optional[float] = None
        seller_reviews: Optional[int] = None

        rating_el = await card.query_selector(
            '[data-marker="seller-info/rating"], [class*="rating"]'
        )
        if rating_el:
            seller_rating = parse_rating(await rating_el.inner_text())

        reviews_el = await card.query_selector(
            '[data-marker="seller-info/reviews"], [class*="reviews"]'
        )
        if reviews_el:
            seller_reviews = parse_reviews(await reviews_el.inner_text())

        return Listing(
            title=title,
            url=url,
            price=price,
            category=cat["label"],
            delivery=delivery,
            seller_rating=seller_rating,
            seller_reviews=seller_reviews,
            is_booked=is_booked,
        )


# =============================================================================
#  ЭКСПОРТ В EXCEL
# =============================================================================

def export_to_excel(rows: list[Listing], path: str) -> None:
    if Workbook is None:
        raise RuntimeError("openpyxl не установлен. Выполните: pip install openpyxl")

    wb = Workbook()
    ws = wb.active
    ws.title = "Avito Hunter"

    headers = ["Балл", "Категория", "Цена ₽", "Получение", "Название", "Причины оценки", "Ссылка"]
    ws.append(headers)

    header_fill = PatternFill("solid", fgColor="2B2B2B")
    header_font = Font(bold=True, color="FFFFFF")
    for col_idx, _ in enumerate(headers, 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")

    for r in rows:
        ws.append([
            r.score,
            r.category,
            r.price if r.price is not None else "",
            r.delivery,
            r.title,
            "; ".join(r.reasons),
            r.url,
        ])

    widths = [8, 22, 10, 14, 55, 45, 60]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    wb.save(path)


# =============================================================================
#  GUI
# =============================================================================

class AvitoHunterApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.title(APP_TITLE)
        self.geometry(APP_GEOMETRY)
        self.minsize(1200, 780)

        # --- данные ---
        self.results: list[Listing] = []
        self.result_queue: "queue.Queue[tuple[str, Any]]" = queue.Queue()
        self.stop_event = threading.Event()
        self.worker_thread: Optional[threading.Thread] = None
        self._worker_running = False
        self._sort_reverse: dict[str, bool] = {}

        # --- категории виджетов ---
        self.cat_widgets: dict[str, dict[str, Any]] = {}

        self._build_ui()
        self._poll_queue()

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_search = self.tabview.add("🔍 Поиск и скоринг")
        self.tab_help = self.tabview.add("📖 Инструкция и справочник")

        self._build_search_tab()
        self._build_help_tab()

    # -------------------------------------------------------- вкладка поиска
    def _build_search_tab(self) -> None:
        # --- левая панель (настройки) ---
        left = ctk.CTkFrame(self.tab_search, width=420)
        left.pack(side="left", fill="y", padx=(0, 8), pady=4)
        left
        # Кнопки
        self.start_btn = ctk.CTkButton(
            left, text="🚀 Запустить поиск", height=42,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._on_start_search,
        )
        self.start_btn.pack(fill="x", pady=(4, 2))

        self.save_btn = ctk.CTkButton(
            left, text="📊 Сохранить в Excel (.xlsx)", height=38,
            fg_color="#2E7D32", hover_color="#1B5E20",
            command=self._on_save_excel,
        )
        self.save_btn.pack(fill="x", pady=(0, 8))

        self.left_frame = left

        # Город и соседние города
        ctk.CTkLabel(left, text="Город (slug)", anchor="w").pack(fill="x", pady=(6, 2))
        self.city_entry = ctk.CTkEntry(
            left, placeholder_text="Slug города, напр: pervouralsk (пусто = rossiya)"
        )
        self.city_entry.pack(fill="x", pady=(0, 6))

        ctk.CTkLabel(left, text="Соседние города для самовывоза", anchor="w").pack(fill="x", pady=(6, 2))
        self.nearby_entry = ctk.CTkEntry(
            left, placeholder_text="Города рядом через запятую..."
        )
        self.nearby_entry.pack(fill="x", pady=(0, 8))

        # Категории
        ctk.CTkLabel(left, text="Категории", anchor="w",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(fill="x", pady=(8, 4))

        for cat in CATEGORIES:
            self._build_category_block(left, cat)

        # Стоп-слова
        ctk.CTkLabel(left, text="Стоп-слова", anchor="w").pack(fill="x", pady=(10, 2))
        self.stop_words_box = ctk.CTkTextbox(left, height=55)
        self.stop_words_box.pack(fill="x", pady=(0, 4))

        ctk.CTkButton(
            left, text="Заполнить базовые стоп-слова",
            command=self._fill_default_stop_words,
        ).pack(fill="x", pady=(0, 8))

        # Чекбокс
        self.filter_booking_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            left,
            text="Отсекать бронь и рейтинг продавца < 4.5",
            variable=self.filter_booking_var,
        ).pack(fill="x", pady=(4, 10))

        # статус
        self.status_label = ctk.CTkLabel(left, text="Готов к работе", anchor="w",
                                         text_color="#9E9E9E")
        self.status_label.pack(fill="x", pady=(8, 0))

        # --- правая панель (результаты) ---
        right = ctk.CTkFrame(self.tab_search)
        right.pack(side="left", fill="both", expand=True, pady=4)

        header = ctk.CTkFrame(right, fg_color="transparent")
        header.pack(fill="x", padx=8, pady=(8, 4))
        ctk.CTkLabel(header, text="📋 Результаты",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")
        self.counter_label = ctk.CTkLabel(header, text="0 объявлений", text_color="#9E9E9E")
        self.counter_label.pack(side="right")

        # таблица
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Avito.Treeview",
            background="#1F1F1F", fieldbackground="#1F1F1F",
            foreground="#E5E5E5", rowheight=26, borderwidth=0,
        )
        style.configure(
            "Avito.Treeview.Heading",
            background="#2B2B2B", foreground="#FFFFFF",
            relief="flat", font=("Segoe UI", 10, "bold"),
        )
        style.map("Avito.Treeview", background=[("selected", "#1F538D")])
        style.map("Avito.Treeview.Heading", background=[("active", "#3A3A3A")])

        tree_frame = ctk.CTkFrame(right)
        tree_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        columns = ("score", "category", "price", "delivery", "title", "reasons")
        self.tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings",
            style="Avito.Treeview", selectmode="browse",
        )
        headings = {
            "score": "Балл",
            "category": "Категория",
            "price": "Цена ₽",
            "delivery": "Получение",
            "title": "Название",
            "reasons": "Причины оценки",
        }
        widths = {
            "score": 70, "category": 130, "price": 90,
            "delivery": 100, "title": 380, "reasons": 280,
        }
        for col in columns:
            self.tree.heading(col, text=headings[col],
                              command=lambda c=col: self._sort_by(c))
            self.tree.column(col, width=widths[col],
                             anchor="w" if col in ("title", "reasons") else "center")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # подсветка строк
        self.tree.tag_configure("high", background="#1E4620")
        self.tree.tag_configure("mid", background="#1F1F1F")
        self.tree.tag_configure("low", background="#3A2020")
        self.tree.tag_configure("odd", background="#262626")

        self.tree.bind("<Double-1>", self._on_row_double_click)

    def _build_category_block(self, parent, cat: dict[str, str]) -> None:
        block = ctk.CTkFrame(parent, fg_color="#232323", corner_radius=8)
        block.pack(fill="x", pady=4, padx=2)

        var = ctk.BooleanVar(value=False)
        sw = ctk.CTkSwitch(block, text=cat["label"], variable=var)
        sw.pack(fill="x", padx=8, pady=(6, 2))

        query_entry = ctk.CTkEntry(block, placeholder_text=cat["query_ph"])
        query_entry.pack(fill="x", padx=8, pady=(0, 4))

        price_entry = ctk.CTkEntry(block, placeholder_text="До ₽", width=110)
        price_entry.pack(fill="x", padx=8, pady=(0, 6))

        self.cat_widgets[cat["key"]] = {
            "var": var,
            "query": query_entry,
            "price": price_entry,
            "url_path": cat["url_path"],
            "label": cat["label"],
        }

    # ---------------------------------------------------- вкладка инструкции
    def _build_help_tab(self) -> None:
        scroll = ctk.CTkFrame(self.tab_help)
        scroll.pack(fill="both", expand=True, padx=8, pady=8)

        # Блок 1
        self._make_card(
            scroll,
            "📝 Блок 1 — Как правильно вписывать модели комплектующих",
            [
                "Не нужно писать длинные названия вендоров и маркетинговые приставки — "
                "вы потеряете объявления обычных пользователей, которые не указывают полное название.",
                "",
                "❌ Плохо: «Gigabyte GeForce GTX 1060 Windforce OC 6GB»",
                "✅ Хорошо: «1060 6gb»",
                "",
                "❌ Плохо: «Kingston Fury Beast DDR4 3200MHz 16GB (2x8)»",
                "✅ Хорошо: «ddr4 16gb» или «ddr4 3200»",
                "",
                "❌ Плохо: «ASUS ROG STRIX B450-F GAMING II»",
                "✅ Хорошо: «b450 am4» (чипсет + сокет)",
                "",
                "❌ Плохо: «Блок питания Chieftec A-135 650W 80+ Bronze»",
                "✅ Хорошо: «600w» или «chieftec 650w»",
            ],
        )

        # Блок 2 — с кнопками
        card2 = ctk.CTkFrame(scroll, fg_color="#232323", corner_radius=10)
        card2.pack(fill="x", pady=8, padx=4)
        ctk.CTkLabel(
            card2,
            text="🔗 Блок 2 — Полезные ресурсы для проверки моделей и совместимости",
            font=ctk.CTkFont(size=15, weight="bold"), anchor="w",
        ).pack(fill="x", padx=12, pady=(10, 6))

        links = [
            ("TechPowerUp — база видеокарт (VRAM / шина / TDP)",
             "https://www.techpowerup.com/gpu-specs/"),
            ("TechPowerUp — база процессоров (сокеты / ядра / ОЗУ)",
             "https://www.techpowerup.com/cpu-specs/"),
            ("DNS Конфигуратор — проверка совместимости ПК",
             "https://www.dns-shop.ru/configurator/"),
        ]
        for label, url in links:
            row = ctk.CTkFrame(card2, fg_color="transparent")
            row.pack(fill="x", padx=12, pady=4)
            ctk.CTkLabel(row, text=label, anchor="w").pack(side="left", fill="x", expand=True)
            ctk.CTkButton(
                row, text="Открыть", width=100,
                command=lambda u=url: webbrowser.open(u),
            ).pack(side="right")
        ctk.CTkLabel(card2, text="").pack(pady=4)

        # Блок 3
        self._make_card(
            scroll,
            "🚀 Блок 3 — Пошаговая инструкция по запуску и работе",
            [
                "1. Slug города. Откройте Авито в браузере, выберите город. "
                "В адресной строке будет что-то вроде https://www.avito.ru/pervouralsk — "
                "«pervouralsk» и есть slug. Вставьте его в поле «Город». "
                "Если оставить пустым — поиск пойдёт по всей России (rossiya).",
                "",
                "2. Первый запуск и капча. При первом поиске откроется окно браузера. "
                "Если Авито покажет капчу — пройдите её один раз. Сессия сохранится в папке "
                "./browser_profile, и при следующих запусках капча обычно не требуется.",
                "",
                "3. Как читать баллы надёжности (0–100):",
                "   • 75–100 — надёжный продавец, есть самовывоз/тесты/пломбы;",
                "   • 50–74 — нейтрально, стоит проверить детали;",
                "   • 0–49 — риски: бронь, низкий рейтинг, рефаб, маркеры «на запчасти».",
                "",
                "4. Выгрузка в Excel. Нажмите «📊 Сохранить в Excel (.xlsx)» — "
                "будет сохранён файл с полным списком, причинами оценки и кликабельными ссылками.",
            ],
        )

    def _make_card(self, parent, title: str, lines: list[str]) -> None:
        card = ctk.CTkFrame(parent, fg_color="#232323", corner_radius=10)
        card.pack(fill="x", pady=8, padx=4)
        ctk.CTkLabel(
            card, text=title, anchor="w",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(fill="x", padx=12, pady=(10, 6))
        text = "\n".join(lines)
        ctk.CTkLabel(
            card, text=text, anchor="w", justify="left",
            wraplength=1100, font=ctk.CTkFont(size=13),
        ).pack(fill="x", padx=12, pady=(0, 12))

    # ---------------------------------------------------------- обработчики
    def _fill_default_stop_words(self) -> None:
        self.stop_words_box.delete("1.0", "end")
        self.stop_words_box.insert("1.0", DEFAULT_STOP_WORDS)

    def _collect_config(self) -> Optional[SearchConfig]:
        city = self.city_entry.get().strip()
        nearby_raw = self.nearby_entry.get().strip()
        nearby = [c.strip() for c in nearby_raw.split(",") if c.strip()] if nearby_raw else []

        stop_words_raw = self.stop_words_box.get("1.0", "end").strip()
        stop_words = [w.strip().lower() for w in re.split(r"[,\n]", stop_words_raw) if w.strip()]

        categories: list[dict[str, Any]] = []
        for key, w in self.cat_widgets.items():
            if not w["var"].get():
                continue
            query = w["query"].get().strip()
            price_raw = w["price"].get().strip()
            max_price = parse_price(price_raw) if price_raw else None
            categories.append({
                "key": key,
                "label": w["label"],
                "query": query,
                "max_price": max_price,
                "url_path": w["url_path"],
            })

        if not categories:
            messagebox.showwarning(
                "Нет категорий",
                "Включите хотя бы одну категорию для поиска.",
            )
            return None

        return SearchConfig(
            city_slug=city,
            nearby_cities=nearby,
            stop_words=stop_words,
            filter_booking_and_rating=self.filter_booking_var.get(),
            categories=categories,
        )

    def _on_start_search(self) -> None:
        if self._worker_running:
            return
        cfg = self._collect_config()
        if cfg is None:
            return

        # очистка
        self.results.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.counter_label.configure(text="0 объявлений")
        self.stop_event.clear()
        self._set_worker_running(True)
        self.status_label.configure(text="Запуск браузера...")

        self.worker_thread = threading.Thread(
            target=self._worker_target, args=(cfg,), daemon=True
        )
        self.worker_thread.start()

    def _worker_target(self, cfg: SearchConfig) -> None:
        parser = AvitoParser(cfg, self.result_queue, self.stop_event)
        parser.run()

    # --------- очередь UI ----------
    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self.result_queue.get_nowait()
                if kind == "log":
                    self.status_label.configure(text=str(payload))
                elif kind == "row":
                    self._add_result(payload)
                elif kind == "error":
                    messagebox.showerror("Ошибка", str(payload))
                elif kind == "done":
                    self._set_worker_running(False)
                    self.status_label.configure(
                        text=f"Готово. Найдено: {len(self.results)}"
                    )
        except queue.Empty:
            pass
        self.after(200, self._poll_queue)

    def _add_result(self, listing: Listing) -> None:
        self.results.append(listing)
        price_str = f"{listing.price:,}".replace(",", " ") if listing.price is not None else "—"
        if listing.score >= 75:
            tag = "high"
        elif listing.score >= 50:
            tag = "mid"
        else:
            tag = "low"

        self.tree.insert(
            "", "end",
            values=(
                listing.score,
                listing.category,
                price_str,
                listing.delivery,
                listing.title,
                "; ".join(listing.reasons),
            ),
            tags=(tag,),
        )
        self.counter_label.configure(text=f"{len(self.results)} объявлений")

    def _set_worker_running(self, running: bool) -> None:
        self._worker_running = running
        self.start_btn.configure(state="disabled" if running else "normal")
        if running:
            self.start_btn.configure(text="⏳ Поиск...")
        else:
            self.start_btn.configure(text="🚀 Запустить поиск")

    # --------- двойной клик ----------
    def _on_row_double_click(self, _event) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if 0 <= idx < len(self.results):
            webbrowser.open(self.results[idx].url)

    # --------- сортировка ----------
    def _sort_by(self, col: str) -> None:
        reverse = not self._sort_reverse.get(col, False)
        self._sort_reverse[col] = reverse

        def key_func(i: int) -> Any:
            r = self.results[i]
            return {
                "score": r.score,
                "category": r.category.lower(),
                "price": r.price if r.price is not None else 10**12,
                "delivery": r.delivery.lower(),
                "title": r.title.lower(),
                "reasons": ";".join(r.reasons).lower(),
            }[col]

        indexed = list(enumerate(self.results))
        indexed.sort(key=lambda t: key_func(t[0]), reverse=reverse)

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.results = [r for _, r in indexed]
        for r in self.results:
            price_str = f"{r.price:,}".replace(",", " ") if r.price is not None else "—"
            if r.score >= 75:
                tag = "high"
            elif r.score >= 50:
                tag = "mid"
            else:
                tag = "low"
            self.tree.insert(
                "", "end",
                values=(r.score, r.category, price_str, r.delivery,
                        r.title, "; ".join(r.reasons)),
                tags=(tag,),
            )

    # --------- Excel ----------
    def _on_save_excel(self) -> None:
        if not self.results:
            messagebox.showinfo("Нет данных", "Сначала выполните поиск.")
            return
        default_name = f"avito_hunter_{datetime.now():%Y-%m-%d_%H-%M}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Сохранить результаты",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return
        try:
            export_to_excel(self.results, path)
            messagebox.showinfo("Готово", f"Файл сохранён:\n{path}")
            if sys.platform == "win32":
                try:
                    os.startfile(path)  # type: ignore[attr-defined]
                except Exception:  # noqa: BLE001
                    pass
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Ошибка сохранения", str(exc))


# =============================================================================
#  ENTRYPOINT
# =============================================================================

def main() -> None:
    app = AvitoHunterApp()
    app.mainloop()


if __name__ == "__main__":
    main()