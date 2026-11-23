"""
gpu_bench.py — локальный словарь производительности видеокарт и калькулятор
value-index (₽ за 1 балл FPS). 100 баллов = GTX 1060 6GB.

Никаких внешних API. Чистый Python + регулярки.
"""

from __future__ import annotations

import re
from typing import Optional

# ---------------------------------------------------------------------------
# База производительности (усреднённые FPS-индексы в игровых сценариях 1080p)
# 100 баллов = GTX 1060 6GB
# ---------------------------------------------------------------------------
GPU_BENCHMARKS: dict[str, int] = {
    # --- low / legacy ---
    "gtx 750 ti": 45,
    "gtx 1050": 55,
    "gtx 1050 ti": 65,
    "gtx 960": 60,
    "gtx 970": 85,
    "gtx 980": 100,
    "gtx 980 ti": 130,
    "rx 550": 35,
    "rx 560": 45,
    "rx 570": 78,
    "rx 580": 88,
    "rx 590": 92,

    # --- mainstream pascal ---
    "gtx 1060 3gb": 88,
    "gtx 1060": 100,
    "gtx 1060 6gb": 100,
    "gtx 1070": 130,
    "gtx 1070 ti": 145,
    "gtx 1080": 165,
    "gtx 1080 ti": 200,

    # --- turing ---
    "gtx 1650": 88,
    "gtx 1650 super": 105,
    "gtx 1660": 118,
    "gtx 1660 super": 130,
    "gtx 1660 ti": 135,
    "rtx 2060": 155,
    "rtx 2060 super": 175,
    "rtx 2070": 185,
    "rtx 2070 super": 200,
    "rtx 2080": 215,
    "rtx 2080 super": 230,
    "rtx 2080 ti": 260,

    # --- amd navi ---
    "rx 5500 xt": 100,
    "rx 5600 xt": 140,
    "rx 5700": 165,
    "rx 5700 xt": 180,
    "rx 6600": 175,
    "rx 6600 xt": 195,
    "rx 6700 xt": 230,
    "rx 6800": 265,
    "rx 6800 xt": 295,
    "rx 6900 xt": 320,
    "rx 7800 xt": 340,

    # --- ampere ---
    "rtx 3050": 120,
    "rtx 3060": 180,
    "rtx 3060 ti": 210,
    "rtx 3070": 245,
    "rtx 3070 ti": 260,
    "rtx 3080": 300,
    "rtx 3080 ti": 320,
    "rtx 3090": 340,

    # --- ada ---
    "rtx 4060": 195,
    "rtx 4060 ti": 235,
    "rtx 4070": 275,
}


# ---------------------------------------------------------------------------
# Регулярки для распознавания моделей в «грязных» заголовках Авито
# (латиница/кириллица, сокращения вроде 1660s, 3060ti, 580 8gb, gtx1060)
# ---------------------------------------------------------------------------

# Универсальный «токен»: буквы+цифры, допускаем пробел/дефис между ними.
_SPLIT_RE = re.compile(r"[^a-zа-я0-9]+", re.IGNORECASE)

# Числовые «хвосты» моделей для серии
_SERIES_RE = re.compile(
    r"\b(?:"
    r"750ti|750\s*ti|"
    r"1050ti|1050\s*ti|1050|"
    r"960|970|980ti|980\s*ti|980|"
    r"1060|1070ti|1070\s*ti|1070|1080ti|1080\s*ti|1080|"
    r"1650super|1650\s*super|1650|1660super|1660\s*super|1660ti|1660\s*ti|1660|"
    r"2060super|2060\s*super|2060|2070super|2070\s*super|2070|"
    r"2080super|2080\s*super|2080ti|2080\s*ti|2080|"
    r"3050|3060ti|3060\s*ti|3060|3070ti|3070\s*ti|3070|"
    r"3080ti|3080\s*ti|3080|3090|"
    r"4060ti|4060\s*ti|4060|4070|"
    r"5500xt|5500\s*xt|5500|5600xt|5600\s*xt|5600|"
    r"5700xt|5700\s*xt|5700|6600xt|6600\s*xt|6600|"
    r"6700xt|6700\s*xt|6700|6800xt|6800\s*xt|6800|6900xt|6900\s*xt|6900|"
    r"7800xt|7800\s*xt|7800|"
    r"580|590|570|560|550"
    r")\b",
    re.IGNORECASE,
)

# Маркеры памяти для уточнения (например 1060 3gb vs 6gb, 580 8gb)
_VRAM_RE = re.compile(r"\b([0-9]{1,2})\s*(?:gb|гб)\b", re.IGNORECASE)

# Префикс производителя/линейки
_BRAND_RE = re.compile(
    r"\b(?:gtx|rtx|rx|radeon|geforce)\b", re.IGNORECASE
)


def _normalize(text: str) -> str:
    """Приводим к нижнему регистру, убираем лишние разделители."""
    text = text.lower()
    text = text.replace("ё", "е")
    # Унификация сокращений, которые любят писать на Авито
    text = re.sub(r"\b(\d{4})\s*s\b", r"\1 super", text)   # 1660s -> 1660 super
    text = re.sub(r"\b(\d{4})\s*ti\b", r"\1 ti", text)
    text = re.sub(r"\b(\d{3,4})\s*xt\b", r"\1 xt", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _detect_series(text: str) -> Optional[str]:
    """Возвращает базовый ключ серии, например 'rtx 3060 ti' / 'rx 580'."""
    low = _normalize(text)

    # Ищем серию
    m = _SERIES_RE.search(low)
    if not m:
        return None
    series = re.sub(r"\s+", " ", m.group(0)).strip()

    # Определяем линейку (nvidia / amd)
    if "rtx" in low or "gtx" in low or "geforce" in low:
        brand = "nvidia"
    elif "rx" in low or "radeon" in low:
        brand = "amd"
    else:
        # Эвристики по номеру
        num = re.search(r"\d{3,4}", series)
        if not num:
            return None
        n = int(num.group(0))
        # 5xxx/6xxx/7xxx — AMD, остальное — NVIDIA
        brand = "amd" if n >= 5000 and not (n >= 10000) else "nvidia"

    # Спец-случаи RX 550/560/570/580/590 (3-значные)
    if re.fullmatch(r"\d{3}", series):
        brand = "amd"

    # Собираем финальный ключ
    if brand == "nvidia":
        if series in ("750", "750 ti", "960", "970", "980", "980 ti",
                      "1050", "1050 ti", "1060", "1070", "1070 ti",
                      "1080", "1080 ti", "1650", "1650 super",
                      "1660", "1660 super", "1660 ti",
                      "2060", "2060 super", "2070", "2070 super",
                      "2080", "2080 super", "2080 ti"):
            if series in ("750", "960", "970", "980", "980 ti",
                          "1050", "1050 ti", "1060", "1070", "1070 ti",
                          "1080", "1080 ti",
                          "1650", "1650 super", "1660", "1660 super", "1660 ti"):
                return f"gtx {series}"
            return f"rtx {series}"
        return f"rtx {series}"

    # AMD
    return f"rx {series}"


def calc_value_index(title: str, price: int) -> tuple[str, int]:
    """
    По заголовку распознаёт GPU и считает цену за 1 балл FPS.

    Возвращает (название_эталона, цена_за_1_балл_FPS) либо ("", 0),
    если модель не распознана или цена некорректна.
    """
    if not title or price is None or price <= 0:
        return "", 0

    low = _normalize(title)
    series = _detect_series(low)
    if not series:
        return "", 0

    # Уточнение по VRAM для 1060 и 580 (частые ловушки на Авито)
    vram_m = _VRAM_RE.search(low)
    vram = int(vram_m.group(1)) if vram_m else None

    candidates: list[str] = [series]

    # Спец-кейс GTX 1060 3GB/6GB — разный индекс
    if series == "gtx 1060":
        if vram == 3:
            candidates = ["gtx 1060 3gb"]
        elif vram == 6:
            candidates = ["gtx 1060 6gb"]
        else:
            candidates = ["gtx 1060 6gb", "gtx 1060 3gb"]

    # Спец-кейс GTX 1650 super / 1650
    if series in ("rtx 1650 super", "gtx 1650 super"):
        candidates = ["gtx 1650 super", "gtx 1650"]
    if series in ("rtx 1650", "gtx 1650"):
        candidates = ["gtx 1650"]

    # Ищем первый известный ключ в словаре
    for key in candidates:
        if key in GPU_BENCHMARKS:
            score = GPU_BENCHMARKS[key]
            if score <= 0:
                return "", 0
            price_per_fps = round(price / score)
            return key.upper(), price_per_fps

    # Фоллбэк: пробуем найти хотя бы подстроку в словаре
    for key, score in GPU_BENCHMARKS.items():
        if key in series or series in key:
            if score <= 0:
                continue
            return key.upper(), round(price / score)

    return "", 0


# ---------------------------------------------------------------------------
# Самопроверка при запуске модуля напрямую
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    tests = [
        ("Видеокарта RTX 3060 Ti 8GB MSI", 28000),
        ("GTX 1660 Super Palit", 12000),
        ("1660s Asus dual", 11500),
        ("RX 580 8GB Sapphire", 7000),
        ("GTX 1060 3GB", 5000),
        ("GTX 1060 6gb", 8000),
        ("RTX 4070", 55000),
        ("RX 7800 XT", 60000),
        ("GTX 750 Ti", 2500),
        ("RTX 2060 Super", 15000),
        ("Не видюха, а процессор i5", 5000),
    ]
    for t, p in tests:
        name, ppf = calc_value_index(t, p)
        print(f"{t!r:50} -> {name!r:20} {ppf} ₽/FPS")