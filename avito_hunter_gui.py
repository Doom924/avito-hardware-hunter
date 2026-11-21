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
#  КОНСТАНТЫ / ПАЛИТРА PRO-СТИЛЯ
# =============================================================================

APP_TITLE = "Avito Hunter — поиск и скоринг комплектующих"
APP_GEOMETRY = "1400x900"
BROWSER_PROFILE_DIR = "./browser_profile"
AVITO_BASE = "https://www.avito.ru"

# --- Палитра ---
COLOR_BG            = "#0F172A"   # общий фон
COLOR_SURFACE       = "#1E293B"   # карточки/секции
COLOR_SURFACE_2     = "#111827"   # чуть темнее (поля)
COLOR_BORDER        = "#334155"   # тонкая обводка
COLOR_TEXT          = "#E2E8F0"   # основной текст
COLOR_TEXT_DIM      = "#94A3B8"   # второстепенный текст
COLOR_PRIMARY       = "#3B82F6"   # синий акцент
COLOR_PRIMARY_HOVER = "#2563EB"
COLOR_SUCCESS       = "#16A34A"
COLOR_SUCCESS_HOVER = "#15803D"
COLOR_WARN          = "#F59E0B"
COLOR_DANGER        = "#DC2626"

# --- Подсветка строк Treeview по баллам ---
ROW_HIGH_BG  = "#14361F"   # >= 75 — зелёный
ROW_MID_BG   = "#3B2F0B"   # 50–74 — жёлтый
ROW_LOW_BG   = "#3B1212"   # < 50 — тускло-красный
ROW_TEXT     = "#F1F5F9"

DEFAULT_STOP_WORDS = (
    "запчаст, нерабоч, сломан, артефакт, коробк, ошибка 43, "
    "треугольник, дымок, отвал, только обмен, без претензий"
)

CATEGORIES: list[dict[str, str]] = [
    {"key": "gpu",    "label": "Видеокарты",        "query_ph": "Напр: rx 580 8gb",
     "url_path": "tovary_dlya_kompyutera/videokarty"},
    {"key": "ram",    "label": "Оперативная память", "query_ph": "Напр: ddr4 16gb",
     "url_path": "tovary_dlya_kompyutera/pamyat"},
    {"key": "case",   "label": "Корпуса",            "query_ph": "Напр: корпус atx",
     "url_path": "tovary_dlya_kompyutera/korpusa"},
    {"key": "cpu",    "label": "Процессоры",         "query_ph": "Напр: ryzen 5 5600",
     "url_path": "tovary_dlya_kompyutera/komplektuyuschie_dlya_pk"},
    {"key": "mb",     "label": "Материнские платы",  "query_ph": "Напр: b450 am4",
     "url_path": "tovary_dlya_kompyutera/materinskie_platy"},
    {"key": "psu",    "label": "Блоки питания",      "query_ph": "Напр: бп 600w",
     "url_path": "tovary_dlya_kompyutera/komplektuyuschie_dlya_pk"},
    {"key": "cooler", "label": "Охлаждение",         "query_ph": "Напр: кулер башня",
     "url_path": "tovary_dlya_kompyutera/komplektuyuschie/sistemy_ohlazhdeniya"},
    {"key": "ssd",    "label": "SSD / Накопители",   "query_ph": "Напр: ssd m2 512gb",
     "url_path": "tovary_dlya_kompyutera/komplektuyuschie/zhestkie_diski"},
]

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
    filter_min_reviews: bool
    filter_avito_delivery: bool
    max_pages: int = 3
    categories: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class Listing:
    title: str
    url: str
    price: Optional[int]
    category: str
    delivery: str
    has_avito_delivery: bool
    seller_rating: Optional[float]
    seller_reviews: Optional[int]
    is_booked: bool
    score: int = 0
    reasons: list[str] = field(default_factory=list)


# =============================================================================
#  УТИЛИТЫ
# =============================================================================

def parse_price(raw: str) -> Optional[int]:
    if not raw:
        return None
    digits = re.sub(r"[^\d]", "", raw)
    return int(digits) if digits else None


def parse_reviews(raw: str) -> Optional[int]:
    if not raw:
        return None
    low = raw.lower()
    if "нет отзыв" in low or "без отзыв" in low:
        return 0
    digits = re.sub(r"[^\d]", "", raw)
    return int(digits) if digits else None


def parse_rating(raw: str) -> Optional[float]:
    if not raw:
        return None
    m = re.search(r"\d+[.,]?\d*", raw)
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


def parse_stop_words(raw: str) -> list[str]:
    if not raw:
        return []
    parts = re.split(r"[,;|\n]+", raw)
    return [p.strip().lower() for p in parts if p.strip()]


# =============================================================================
#  СКОРИНГ
# =============================================================================

def score_listing(listing: Listing) -> None:
    score = 50.0
    reasons: list[str] = []

    d = (listing.delivery or "").lower()
    if "самовывоз" in d:
        score += 10
        reasons.append("самовывоз (+10)")
    elif "доставка" in d:
        score += 4
        reasons.append("доставка (+4)")

    if listing.has_avito_delivery:
        score += 8
        reasons.append("Авито Доставка (+8)")

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

    if listing.is_booked:
        score -= 20
        reasons.append("бронь (-20)")

    score = max(0, min(100, round(score)))
    listing.score = score
    listing.reasons = reasons or ["—"]


# =============================================================================
#  СКРОЛЛИРУЕМЫЙ КОНТЕЙНЕР (Canvas + CTkScrollbar, кросс-платформенный)
# =============================================================================

class ScrollableFrame(ctk.CTkFrame):
    """
    Надёжный скроллируемый контейнер на базе tkinter.Canvas + CTkScrollbar.
    Не схлопывается по высоте на Arch Linux / Wayland (в отличие от CTkScrollableFrame).
    Поддерживает плавный скролл колесом мыши: Windows (<MouseWheel>) и Linux (<Button-4>/<Button-5>).
    """

    _SCROLL_STEP = 3  # единиц прокрутки за «щелчок»

    def __init__(self, master, **kwargs) -> None:
        super().__init__(master, fg_color="transparent", **kwargs)

        self._canvas = tk.Canvas(
            self,
            bg=COLOR_BG,
            highlightthickness=0,
            bd=0,
        )
        self._canvas.grid(row=0, column=0, sticky="nsew")

        self._scrollbar = ctk.CTkScrollbar(
            self, orientation="vertical", command=self._canvas.yview
        )
        self._scrollbar.grid(row=0, column=1, sticky="ns", padx=(4, 0))
        self._canvas.configure(yscrollcommand=self._scrollbar.set)

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Внутренний frame с содержимым
        self.inner = ctk.CTkFrame(self._canvas, fg_color="transparent")
        self._inner_id = self._canvas.create_window(
            (0, 0), window=self.inner, anchor="nw"
        )

        # Обновляем scrollregion при изменении размеров содержимого
        self.inner.bind("<Configure>", self._on_inner_configure)
        self._canvas.bind("<Configure>", self._on_canvas_configure)

        # Кросс-платформенная привязка колеса
        self._bind_mousewheel()

    def _on_inner_configure(self, _event=None) -> None:
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_configure(self, event) -> None:
        # внутренний frame всегда по ширине canvas
        self._canvas.itemconfigure(self._inner_id, width=event.width)

    def _bind_mousewheel(self) -> None:
        # Windows / macOS
        self._canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        # Linux (X11)
        self._canvas.bind_all("<Button-4>", self._on_button4, add="+")
        self._canvas.bind_all("<Button-5>", self._on_button5, add="+")

    def _should_scroll(self, event) -> bool:
        """
        Скроллим только если курсор мыши находится над этим контейнером
        или одним из его потомков. Иначе — не мешаем другим скроллам.
        """
        widget = event.widget
        try:
            while widget is not None:
                if widget == self._canvas or widget == self.inner:
                    return True
                widget = widget.master  # type: ignore[attr-defined]
        except Exception:  # noqa: BLE001
            pass
        # fallback: проверяем геометрию
        try:
            x_root, y_root = event.x_root, event.y_root
            cx = self._canvas.winfo_rootx()
            cy = self._canvas.winfo_rooty()
            cw = self._canvas.winfo_width()
            ch = self._canvas.winfo_height()
            return cx <= x_root <= cx + cw and cy <= y_root <= cy + ch
        except Exception:  # noqa: BLE001
            return False

    def _on_mousewheel(self, event) -> None:
        if not self._should_scroll(event):
            return
        delta = -1 if event.delta > 0 else 1
        self._canvas.yview_scroll(delta * self._SCROLL_STEP, "units")

    def _on_button4(self, event) -> None:
        if not self._should_scroll(event):
            return
        self._canvas.yview_scroll(-self._SCROLL_STEP, "units")

    def _on_button5(self, event) -> None:
        if not self._should_scroll(event):
            return
        self._canvas.yview_scroll(self._SCROLL_STEP, "units")


# =============================================================================
#  ПАРСЕР (Playwright, async) — с пагинацией
# =============================================================================

class AvitoParser:
    def __init__(
        self,
        config: SearchConfig,
        out_queue: "queue.Queue[tuple[str, Any]]",
        stop_event: threading.Event,
    ) -> None:
        self.config = config
        self.q = out_queue
        self.stop_event = stop_event

    def run(self) -> None:
        try:
            asyncio.run(self._run_async())
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            self.q.put(("error", f"{exc}\n\n{tb}"))
        finally:
            self.q.put(("done", None))

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
                    self.q.put((
                        "log",
                        f"[{idx}/{total_cats}] Категория: {cat['label']} "
                        f"(до {self.config.max_pages} стр.)",
                    ))
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

    async def _parse_category(self, page: Page, cat: dict[str, Any]) -> None:
        total_new = 0
        for page_num in range(1, max(1, self.config.max_pages) + 1):
            if self.stop_event.is_set():
                return
            url = self._build_url(cat, page_num)
            self.q.put(("log", f"  → стр. {page_num}: {url}"))
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=60_000)
            except Exception as exc:  # noqa: BLE001
                self.q.put(("log", f"  ! Не удалось открыть стр. {page_num}: {exc}"))
                break

            try:
                await page.wait_for_selector('[data-marker="item"]', timeout=20_000)
            except Exception:
                self.q.put((
                    "log",
                    f"  ! Стр. {page_num}: нет объявлений (капча/антибот или конец выдачи).",
                ))
                break

            # 3 плавных скролла для подгрузки всех 50 объявлений страницы
            for _ in range(3):
                if self.stop_event.is_set():
                    return
                try:
                    await page.mouse.wheel(0, 2000)
                except Exception:  # noqa: BLE001
                    pass
                await asyncio.sleep(0.7)

            cards = await page.query_selector_all('[data-marker="item"]')
            if not cards:
                self.q.put(("log", f"  · Стр. {page_num}: карточек 0, завершаем категорию."))
                break

            self.q.put(("log", f"  · Стр. {page_num}: найдено карточек {len(cards)}"))

            parsed_on_page = 0
            for card in cards:
                if self.stop_event.is_set():
                    return
                try:
                    listing = await self._parse_card(card, cat)
                except Exception:  # noqa: BLE001
                    continue
                if listing is None:
                    continue

                if matches_stop_words(listing.title, self.config.stop_words):
                    continue

                if self.config.filter_avito_delivery and not listing.has_avito_delivery:
                    continue

                if self.config.filter_booking_and_rating:
                    if listing.seller_rating is not None and listing.seller_rating < 4.5:
                        self.q.put((
                            "log",
                            f"  ✗ отсеян (рейтинг {listing.seller_rating}): {listing.title[:60]}",
                        ))
                        continue
                    if listing.is_booked:
                        continue

                if self.config.filter_min_reviews:
                    if listing.seller_reviews is None or listing.seller_reviews <= 0:
                        continue

                score_listing(listing)
                self.q.put(("row", listing))
                parsed_on_page += 1
                total_new += 1
                await asyncio.sleep(0.02)

            self.q.put(("log", f"  ✓ Стр. {page_num}: прошло фильтры {parsed_on_page}"))

            # Если на странице было мало карточек (< 40) — это последняя страница
            if len(cards) < 40:
                self.q.put(("log", f"  · Похоже, стр. {page_num} — последняя."))
                break

        self.q.put(("log", f"  ∑ По категории «{cat['label']}» добавлено: {total_new}"))

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

        body = (await card.inner_text()) or ""
        body_low = body.lower()

        delivery = "—"
        if "самовывоз" in body_low:
            delivery = "Самовывоз"
        elif "доставка" in body_low:
            delivery = "Доставка"

        has_avito_delivery = False
        avito_badge = await card.query_selector(
            '[data-marker*="delivery"], [data-marker*="badge/avito"], '
            '[class*="avito-delivery"], [class*="item-badge"][class*="delivery"]'
        )
        if avito_badge:
            has_avito_delivery = True
        if not has_avito_delivery and (
            "авито доставка" in body_low
            or "avito доставка" in body_low
            or "доставка авито" in body_low
        ):
            has_avito_delivery = True

        is_booked = "бронь" in body_low or "забронировано" in body_low

        seller_rating: Optional[float] = None
        seller_reviews: Optional[int] = None

        score_el = await card.query_selector('[data-marker="seller-rating/score"]')
        if score_el:
            seller_rating = parse_rating(await score_el.inner_text() or "")

        summary_el = await card.query_selector('[data-marker="seller-info/summary"]')
        if summary_el:
            summary_text = (await summary_el.inner_text()) or ""
            if seller_rating is None:
                seller_rating = parse_rating(summary_text)
            seller_reviews = parse_reviews(summary_text)

        if seller_rating is None:
            m = re.search(r"(\d+[.,]\d+)\s*★", body)
            if m:
                seller_rating = parse_rating(m.group(1))

        if seller_reviews is None:
            m = re.search(r"(\d+)\s*отзыв", body_low)
            if m:
                seller_reviews = int(m.group(1))
            elif "нет отзыв" in body_low:
                seller_reviews = 0

        return Listing(
            title=title,
            url=url,
            price=price,
            category=cat["label"],
            delivery=delivery,
            has_avito_delivery=has_avito_delivery,
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

    headers = [
        "Балл", "Категория", "Цена ₽", "Получение", "Авито Доставка",
        "Название", "Причины оценки", "Ссылка",
    ]
    ws.append(headers)

    header_fill = PatternFill("solid", fgColor="1E293B")
    header_font = Font(bold=True, color="E2E8F0")
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
            "да" if r.has_avito_delivery else "нет",
            r.title,
            "; ".join(r.reasons),
            r.url,
        ])

    widths = [8, 22, 10, 14, 14, 55, 45, 60]
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
        self.configure(fg_color=COLOR_BG)

        self.title(APP_TITLE)
        self.geometry(APP_GEOMETRY)
        self.minsize(1100, 700)

        self.results: list[Listing] = []
        self.result_queue: "queue.Queue[tuple[str, Any]]" = queue.Queue()
        self.stop_event = threading.Event()
        self.worker_thread: Optional[threading.Thread] = None
        self._worker_running = False
        self._sort_reverse: dict[str, bool] = {}

        self.cat_widgets: dict[str, dict[str, Any]] = {}

        # --- шрифты ---
        self.font_h1 = ctk.CTkFont(family="Segoe UI", size=16, weight="bold")
        self.font_h2 = ctk.CTkFont(family="Segoe UI", size=14, weight="bold")
        self.font_body = ctk.CTkFont(family="Segoe UI", size=12)
        self.font_small = ctk.CTkFont(family="Segoe UI", size=11)
        self.font_button = ctk.CTkFont(family="Segoe UI", size=14, weight="bold")

        self._build_ui()
        self._poll_queue()

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        self.tabview = ctk.CTkTabview(
            self,
            fg_color=COLOR_BG,
            segmented_button_selected_color=COLOR_PRIMARY,
            segmented_button_selected_hover_color=COLOR_PRIMARY_HOVER,
            segmented_button_unselected_color=COLOR_SURFACE,
            segmented_button_unselected_hover_color=COLOR_SURFACE_2,
            text_color=COLOR_TEXT,
        )
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_search = self.tabview.add("🔍 Поиск и скоринг")
        self.tab_help = self.tabview.add("📖 Инструкция и справочник")

        self._build_search_tab()
        self._build_help_tab()

    # -------------------------------------------------------- вкладка поиска
    def _build_search_tab(self) -> None:
        # Двухколоночный layout
        self.tab_search.grid_rowconfigure(0, weight=1)
        self.tab_search.grid_columnconfigure(0, weight=0, minsize=440)
        self.tab_search.grid_columnconfigure(1, weight=1)

        # === ЛЕВАЯ КОЛОНКА: закреплённые кнопки + скролл-область ===
        left_wrap = ctk.CTkFrame(
            self.tab_search,
            fg_color=COLOR_SURFACE,
            corner_radius=10,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        left_wrap.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=4)
        left_wrap.grid_rowconfigure(1, weight=1)
        left_wrap.grid_columnconfigure(0, weight=1)

        # --- Закреплённый верх: заголовок + кнопки ---
        top_sticky = ctk.CTkFrame(left_wrap, fg_color="transparent")
        top_sticky.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))

        ctk.CTkLabel(
            top_sticky, text="⚙️ Настройки поиска",
            font=self.font_h1, text_color=COLOR_TEXT, anchor="w",
        ).pack(fill="x", pady=(0, 8))

        self.start_btn = ctk.CTkButton(
            top_sticky, text="🚀 Запустить поиск", height=42,
            font=self.font_button,
            fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
            corner_radius=8,
            command=self._on_start_search,
        )
        self.start_btn.pack(fill="x", pady=(0, 6))

        self.save_btn = ctk.CTkButton(
            top_sticky, text="📊 Сохранить в Excel (.xlsx)", height=38,
            font=self.font_button,
            fg_color=COLOR_SUCCESS, hover_color=COLOR_SUCCESS_HOVER,
            corner_radius=8,
            command=self._on_save_excel,
        )
        self.save_btn.pack(fill="x", pady=(0, 6))

        # --- Прокручиваемая область ---
        self.left_scroll = ScrollableFrame(left_wrap)
        self.left_scroll.grid(row=1, column=0, sticky="nsew", padx=(6, 6), pady=(0, 8))
        left = self.left_scroll.inner

        # --- Поля города ---
        sec_city = self._make_section(left, "📍 Город и окружение")
        self.city_entry = ctk.CTkEntry(
            sec_city,
            placeholder_text="Slug города, напр: pervouralsk (пусто = rossiya)",
            fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER,
            text_color=COLOR_TEXT, placeholder_text_color=COLOR_TEXT_DIM,
        )
        self.city_entry.pack(fill="x", padx=8, pady=(0, 6))

        self.nearby_entry = ctk.CTkEntry(
            sec_city,
            placeholder_text="Города рядом через запятую...",
            fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER,
            text_color=COLOR_TEXT, placeholder_text_color=COLOR_TEXT_DIM,
        )
        self.nearby_entry.pack(fill="x", padx=8, pady=(0, 6))

        # Глубина поиска
        ctk.CTkLabel(
            sec_city, text="Глубина поиска (страниц)",
            font=self.font_body, text_color=COLOR_TEXT_DIM, anchor="w",
        ).pack(fill="x", padx=8, pady=(2, 2))
        self.pages_entry = ctk.CTkEntry(
            sec_city,
            placeholder_text="3 (≈150 объявлений на категорию)",
            fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER,
            text_color=COLOR_TEXT, placeholder_text_color=COLOR_TEXT_DIM,
        )
        self.pages_entry.insert(0, "3")
        self.pages_entry.pack(fill="x", padx=8, pady=(0, 8))

        # --- Категории ---
        sec_cats = self._make_section(left, "🧩 Категории")
        for cat in CATEGORIES:
            self._build_category_block(sec_cats, cat)

        # --- Стоп-слова ---
        sec_stop = self._make_section(left, "🛑 Стоп-слова")
        ctk.CTkLabel(
            sec_stop,
            text="Разделители: ,  ;  |  или Enter",
            font=self.font_small, text_color=COLOR_TEXT_DIM, anchor="w",
        ).pack(fill="x", padx=8, pady=(0, 4))

        self.stop_words_box = ctk.CTkTextbox(
            sec_stop, height=100,
            fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
        )
        self.stop_words_box.pack(fill="x", padx=8, pady=(0, 6))

        ctk.CTkButton(
            sec_stop, text="Заполнить базовые стоп-слова",
            font=self.font_body,
            fg_color=COLOR_SURFACE_2, hover_color="#475569",
            border_width=1, border_color=COLOR_BORDER,
            command=self._fill_default_stop_words,
        ).pack(fill="x", padx=8, pady=(0, 8))

        # --- Фильтры ---
        sec_filters = self._make_section(left, "🎯 Фильтры надёжности")
        self.filter_booking_var = ctk.BooleanVar(value=False)
        self.filter_min_reviews_var = ctk.BooleanVar(value=True)
        self.filter_avito_delivery_var = ctk.BooleanVar(value=False)

        for text, var in [
            ("Отсекать бронь и рейтинг продавца < 4.5", self.filter_booking_var),
            ("Только продавцы с историей продаж (отзывы > 0)", self.filter_min_reviews_var),
            ("Только с Авито Доставкой", self.filter_avito_delivery_var),
        ]:
            ctk.CTkCheckBox(
                sec_filters, text=text, variable=var,
                font=self.font_body, text_color=COLOR_TEXT,
                fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
                checkmark_color="white",
            ).pack(anchor="w", padx=8, pady=4)

        # --- Статус (внизу прокрутки) ---
        self.status_label = ctk.CTkLabel(
            left, text="Готов к работе",
            font=self.font_small, text_color=COLOR_TEXT_DIM, anchor="w",
            wraplength=380, justify="left",
        )
        self.status_label.pack(fill="x", padx=8, pady=(6, 12))

        # === ПРАВАЯ КОЛОНКА: результаты ===
        right = ctk.CTkFrame(
            self.tab_search,
            fg_color=COLOR_SURFACE,
            corner_radius=10,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        right.grid(row=0, column=1, sticky="nsew", pady=4)
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(right, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        ctk.CTkLabel(
            header, text="📋 Результаты",
            font=self.font_h1, text_color=COLOR_TEXT,
        ).pack(side="left")
        self.counter_label = ctk.CTkLabel(
            header, text="0 объявлений",
            font=self.font_body, text_color=COLOR_TEXT_DIM,
        )
        self.counter_label.pack(side="right")

        # --- Treeview ---
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Avito.Treeview",
            background=COLOR_SURFACE_2, fieldbackground=COLOR_SURFACE_2,
            foreground=ROW_TEXT, rowheight=28, borderwidth=0,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Avito.Treeview.Heading",
            background=COLOR_SURFACE, foreground=COLOR_TEXT,
            relief="flat", font=("Segoe UI", 10, "bold"),
        )
        style.map("Avito.Treeview", background=[("selected", COLOR_PRIMARY)])
        style.map("Avito.Treeview.Heading", background=[("active", "#475569")])

        tree_frame = ctk.CTkFrame(right, fg_color="transparent")
        tree_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        columns = ("score", "category", "price", "delivery", "avito_delivery", "title", "reasons")
        self.tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings",
            style="Avito.Treeview", selectmode="browse",
        )
        headings = {
            "score": "Балл",
            "category": "Категория",
            "price": "Цена ₽",
            "delivery": "Получение",
            "avito_delivery": "Авито Доставка",
            "title": "Название",
            "reasons": "Причины оценки",
        }
        widths = {
            "score": 60, "category": 120, "price": 85,
            "delivery": 100, "avito_delivery": 110,
            "title": 340, "reasons": 260,
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

        # Цветовая подсветка по баллам
        self.tree.tag_configure("high", background=ROW_HIGH_BG, foreground=ROW_TEXT)
        self.tree.tag_configure("mid",  background=ROW_MID_BG,  foreground=ROW_TEXT)
        self.tree.tag_configure("low",  background=ROW_LOW_BG,  foreground=ROW_TEXT)

        self.tree.bind("<Double-1>", self._on_row_double_click)

    # ----------- хелперы UI -----------
    def _make_section(self, parent, title: str) -> ctk.CTkFrame:
        """Карточка-секция со скруглённой рамкой и заголовком."""
        section = ctk.CTkFrame(
            parent,
            fg_color=COLOR_SURFACE_2,
            corner_radius=10,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        section.pack(fill="x", padx=2, pady=(0, 8))
        ctk.CTkLabel(
            section, text=title, font=self.font_h2,
            text_color=COLOR_TEXT, anchor="w",
        ).pack(fill="x", padx=10, pady=(8, 6))
        return section

    def _build_category_block(self, parent, cat: dict[str, str]) -> None:
        block = ctk.CTkFrame(
            parent,
            fg_color=COLOR_SURFACE,
            corner_radius=8,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        block.pack(fill="x", padx=8, pady=4)

        var = ctk.BooleanVar(value=False)
        ctk.CTkSwitch(
            block, text=cat["label"], variable=var,
            font=self.font_body, text_color=COLOR_TEXT,
            progress_color=COLOR_PRIMARY,
        ).pack(fill="x", padx=8, pady=(6, 2))

        query_entry = ctk.CTkEntry(
            block, placeholder_text=cat["query_ph"],
            fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER,
            text_color=COLOR_TEXT, placeholder_text_color=COLOR_TEXT_DIM,
        )
        query_entry.pack(fill="x", padx=8, pady=(0, 4))

        price_entry = ctk.CTkEntry(
            block, placeholder_text="До ₽",
            fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER,
            text_color=COLOR_TEXT, placeholder_text_color=COLOR_TEXT_DIM,
        )
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
        self.tab_help.grid_rowconfigure(0, weight=1)
        self.tab_help.grid_columnconfigure(0, weight=1)

        scroll = ScrollableFrame(self.tab_help)
        scroll.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        parent = scroll.inner

        self._make_card(
            parent,
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
                "",
                "❌ Плохо: «Noctua NH-D15 chromax.black»",
                "✅ Хорошо: «кулер башня» или «noctua»",
                "",
                "❌ Плохо: «Samsung 970 EVO Plus 500GB M.2 NVMe»",
                "✅ Хорошо: «ssd m2 512gb» или «ssd 500gb»",
            ],
        )

        # Блок 2 — с кнопками
        card2 = ctk.CTkFrame(
            parent,
            fg_color=COLOR_SURFACE_2,
            corner_radius=10,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        card2.pack(fill="x", pady=8, padx=4)
        ctk.CTkLabel(
            card2,
            text="🔗 Блок 2 — Полезные ресурсы для проверки моделей и совместимости",
            font=self.font_h2, text_color=COLOR_TEXT, anchor="w",
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
            ctk.CTkLabel(
                row, text=label, anchor="w",
                font=self.font_body, text_color=COLOR_TEXT,
            ).pack(side="left", fill="x", expand=True)
            ctk.CTkButton(
                row, text="Открыть", width=100,
                font=self.font_body,
                fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
                corner_radius=6,
                command=lambda u=url: webbrowser.open(u),
            ).pack(side="right")
        ctk.CTkLabel(card2, text="").pack(pady=4)

        self._make_card(
            parent,
            "🚀 Блок 3 — Пошаговая инструкция по запуску и работе",
            [
                "1. Slug города. Откройте Авито в браузере, выберите город. "
                "В адресной строке будет что-то вроде https://www.avito.ru/pervouralsk — "
                "«pervouralsk» и есть slug. Вставьте его в поле «Город». "
                "Если оставить пустым — поиск пойдёт по всей России (rossiya).",
                "",
                "2. Глубина поиска. Поле «Глубина поиска (страниц)» задаёт, сколько страниц выдачи "
                "обходить на каждую категорию (по умолчанию 3 — это до ~150 объявлений на категорию). "
                "На каждой странице делается 3 плавных скролла, чтобы подгрузились все 50 объявлений. "
                "Если карточек на странице 0 — поиск по категории досрочно завершается.",
                "",
                "3. Первый запуск и капча. При первом поиске откроется окно браузера. "
                "Если Авито покажет капчу — пройдите её один раз. Сессия сохранится в папке "
                "./browser_profile, и при следующих запусках капча обычно не требуется.",
                "",
                "4. Стоп-слова и фразы. В поле «Стоп-слова» можно вводить как одиночные слова, "
                "так и целые фразы (например: «работает но», «на запчасти»). "
                "Разделяйте их запятой, точкой с запятой, символом | или с новой строки. "
                "Регистр не важен — программа приводит всё к нижнему регистру.",
                "",
                "5. Фильтры надёжности. Чекбокс «Отсекать бронь и рейтинг продавца < 4.5» "
                "строго выбрасывает объявления с бронированием или с найденным рейтингом ниже 4.5. "
                "Чекбокс «Только продавцы с историей продаж (отзывы > 0)» (включён по умолчанию) "
                "пропускает только продавцов, у которых есть хотя бы один отзыв.",
                "",
                "6. Авито Доставка. Чекбокс «Только с Авито Доставкой» (по умолчанию выключен) "
                "оставляет только объявления с отметкой Авито Доставки. В таблице результатов "
                "есть столбец «Авито Доставка» (да/нет), и такие объявления получают +8 к баллу.",
                "",
                "7. Как читать баллы надёжности (0–100) и подсветку строк:",
                "   • зелёная строка (75–100) — надёжный продавец, есть самовывоз/тесты/пломбы;",
                "   • жёлтая строка (50–74) — нейтрально, стоит проверить детали;",
                "   • красная строка (0–49) — риски: бронь, низкий рейтинг, рефаб, «на запчасти».",
                "",
                "8. Выгрузка в Excel. Нажмите «📊 Сохранить в Excel (.xlsx)» — "
                "будет сохранён файл с полным списком, причинами оценки, столбцом Авито Доставки "
                "и кликабельными ссылками.",
            ],
        )

    def _make_card(self, parent, title: str, lines: list[str]) -> None:
        card = ctk.CTkFrame(
            parent,
            fg_color=COLOR_SURFACE_2,
            corner_radius=10,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        card.pack(fill="x", pady=8, padx=4)
        ctk.CTkLabel(
            card, text=title, anchor="w",
            font=self.font_h2, text_color=COLOR_TEXT,
        ).pack(fill="x", padx=12, pady=(10, 6))
        ctk.CTkLabel(
            card, text="\n".join(lines), anchor="w", justify="left",
            wraplength=1000,
            font=self.font_body, text_color=COLOR_TEXT,
        ).pack(fill="x", padx=12, pady=(0, 12))

    # ---------------------------------------------------------- обработчики
    def _fill_default_stop_words(self) -> None:
        self.stop_words_box.delete("1.0", "end")
        self.stop_words_box.insert("1.0", DEFAULT_STOP_WORDS)

    def _collect_config(self) -> Optional[SearchConfig]:
        city = self.city_entry.get().strip()
        nearby_raw = self.nearby_entry.get().strip()
        nearby = [c.strip() for c in nearby_raw.split(",") if c.strip()] if nearby_raw else []

        stop_words = parse_stop_words(self.stop_words_box.get("1.0", "end"))

        pages_raw = self.pages_entry.get().strip()
        try:
            max_pages = max(1, int(pages_raw)) if pages_raw else 3
        except ValueError:
            max_pages = 3

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
            filter_min_reviews=self.filter_min_reviews_var.get(),
            filter_avito_delivery=self.filter_avito_delivery_var.get(),
            max_pages=max_pages,
            categories=categories,
        )

    def _on_start_search(self) -> None:
        if self._worker_running:
            return
        cfg = self._collect_config()
        if cfg is None:
            return

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

    def _tag_for_score(self, score: int) -> str:
        if score >= 75:
            return "high"
        if score >= 50:
            return "mid"
        return "low"

    def _add_result(self, listing: Listing) -> None:
        self.results.append(listing)
        price_str = f"{listing.price:,}".replace(",", " ") if listing.price is not None else "—"
        self.tree.insert(
            "", "end",
            values=(
                listing.score,
                listing.category,
                price_str,
                listing.delivery,
                "да" if listing.has_avito_delivery else "нет",
                listing.title,
                "; ".join(listing.reasons),
            ),
            tags=(self._tag_for_score(listing.score),),
        )
        self.counter_label.configure(text=f"{len(self.results)} объявлений")

    def _set_worker_running(self, running: bool) -> None:
        self._worker_running = running
        self.start_btn.configure(state="disabled" if running else "normal")
        if running:
            self.start_btn.configure(text="⏳ Поиск...")
        else:
            self.start_btn.configure(text="🚀 Запустить поиск")

    def _on_row_double_click(self, _event) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if 0 <= idx < len(self.results):
            webbrowser.open(self.results[idx].url)

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
                "avito_delivery": 1 if r.has_avito_delivery else 0,
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
            self.tree.insert(
                "", "end",
                values=(r.score, r.category, price_str, r.delivery,
                        "да" if r.has_avito_delivery else "нет",
                        r.title, "; ".join(r.reasons)),
                tags=(self._tag_for_score(r.score),),
            )

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