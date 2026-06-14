#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
wb_reviews.py — выгрузка отзывов покупателей Wildberries (Россия) по артикулам/SKU.

Кратко о подходе (см. также README.md):
  Используется ПУБЛИЧНЫЙ JSON-API Wildberries, а НЕ браузерная автоматизация.
  Браузер (selenium/playwright) для этой задачи избыточен и медленный:
  страница /feedbacks внутри всё равно дёргает те же самые JSON-эндпоинты.

  Цепочка запросов:
    1) card.wb.ru/cards/v2/detail?nm=<артикул>   -> получаем root (imtId), бренд, название.
       imtId — это и есть "склейка": общий идентификатор для группы вариантов товара.
    2) feedbacks{1,2}.wb.ru/feedbacks/v{1,2}/<imtId> -> получаем ВСЕ отзывы склейки одним ответом.
       Отзывы приходят сразу целиком (без постраничности), поэтому мы получаем
       все текстовые отзывы и затем сами фильтруем по дате (последние N дней).

  Решение проблемы "склейки" (см. п.3 ТЗ):
    Ответ feedbacks содержит для КАЖДОГО отзыва поле nmId (и productDetails.nmId).
    Мы оставляем только те отзывы, у которых nmId == запрошенному артикулу.
    Так отзывы соседних вариантов из склейки не попадают в выгрузку.

Зависимости: requests (см. requirements.txt).
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
import time
import logging
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional

try:
    import requests
    from requests.adapters import HTTPAdapter
except ModuleNotFoundError:
    # Самая частая причина «скрипт не открывается / окно мигает и закрывается»:
    # не установлена библиотека requests. Даём понятное сообщение и не даём окну закрыться.
    print("\n[ОШИБКА] Не установлена библиотека 'requests'.")
    print("Установите её командой:\n")
    print("    pip install requests\n")
    print("(или: python -m pip install -r requirements.txt)\n")
    try:
        input("Нажмите Enter, чтобы закрыть окно…")
    except EOFError:
        pass
    raise SystemExit(1)

try:
    # urllib3 поставляется вместе с requests
    from urllib3.util.retry import Retry
except Exception:  # pragma: no cover
    Retry = None


# --------------------------------------------------------------------------- #
# Конфигурация по умолчанию (можно менять здесь или через аргументы запуска)
# --------------------------------------------------------------------------- #

DEFAULT_DAYS = 365          # выгружаем отзывы за последние N дней
MIN_TEXT_LEN = 2            # минимальная длина текста отзыва (символов)
REQUEST_TIMEOUT = 30        # таймаут одного HTTP-запроса, сек
REQUEST_DELAY = 0.7         # задержка между артикулами/запросами, сек (бережём сайт)
MAX_RETRIES = 4             # число повторных попыток при сетевых/временных ошибках
BACKOFF_BASE = 2            # экспоненциальная пауза: 2, 4, 8, 16 сек

# Эндпоинты карточки товара (для определения imtId/бренда/названия).
# Несколько доменов — на случай, если один недоступен.
CARD_HOSTS = [
    "https://card.wb.ru/cards/v2/detail",
    "https://card.wb.ru/cards/v1/detail",
]
CARD_PARAMS = {"appType": "1", "curr": "rub", "dest": "-1257786", "spp": "30"}

# Хосты/версии API отзывов. Данные шардированы между feedbacks1 и feedbacks2,
# поэтому перебираем все варианты и берём первый, где есть отзывы.
FEEDBACK_ENDPOINTS = [
    "https://feedbacks1.wb.ru/feedbacks/v1/{imt}",
    "https://feedbacks2.wb.ru/feedbacks/v2/{imt}",
    "https://feedbacks1.wb.ru/feedbacks/v2/{imt}",
    "https://feedbacks2.wb.ru/feedbacks/v1/{imt}",
]

# Заголовки CSV с отзывами (на русском). Это и есть "список полей" —
# чтобы добавить/убрать колонку, измените REVIEW_COLUMNS и функцию review_to_row().
REVIEW_COLUMNS = [
    "Артикул WB",
    "Бренд",
    "Название товара",
    "Цвет",
    "Размер",
    "Дата отзыва",
    "Оценка",
    "Есть фото",        # 1/0 — сами фото НЕ скачиваем, только признак наличия
    "Текст отзыва",
    "Достоинства",
    "Недостатки",
    "ID отзыва",
]

# Заголовки сопровождающего файла (сводки).
SUMMARY_COLUMNS = [
    "Артикул",
    "Бренд",
    "Название товара",
    "Ссылка на товар",
    "Найдено отзывов (всего)",
    "Найдено отзывов с текстом (API)",
    "Сохранено отзывов (после фильтра)",
    "Путь к CSV",
    "Статус",
    "Ошибка",
    "Дата и время выгрузки",
]

CSV_DELIMITER = ";"
CSV_ENCODING = "utf-8-sig"  # UTF-8 with BOM — корректно открывается в Excel


# --------------------------------------------------------------------------- #
# Логирование в консоль
# --------------------------------------------------------------------------- #

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("wb_reviews")


# --------------------------------------------------------------------------- #
# HTTP-сессия с повторными попытками
# --------------------------------------------------------------------------- #

def build_session() -> requests.Session:
    """Создаёт requests.Session с разумными заголовками и авто-ретраями на 5xx/429."""
    s = requests.Session()
    s.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
        ),
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "ru-RU,ru;q=0.9",
        "Origin": "https://www.wildberries.ru",
        "Referer": "https://www.wildberries.ru/",
    })
    if Retry is not None:
        retry = Retry(
            total=MAX_RETRIES,
            backoff_factor=1,                       # 0,2,4,8...
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=frozenset(["GET"]),
            raise_on_status=False,
        )
        adapter = HTTPAdapter(max_retries=retry)
        s.mount("https://", adapter)
        s.mount("http://", adapter)
    return s


def get_json(session: requests.Session, url: str, params: Optional[dict] = None,
             max_attempts: int = MAX_RETRIES, missing_codes=(404,),
             quiet: bool = False) -> Optional[dict]:
    """
    GET-запрос с возвратом JSON. Сам обрабатывает сетевые ошибки и делает
    повторные попытки с экспоненциальной паузой (2,4,8,16 сек).
    Возвращает dict или None (если так и не удалось получить данные).

    max_attempts  — число попыток (для лёгкого перебора CDN ставим 1);
    missing_codes — коды, которые трактуем как «здесь данных нет» (не ошибка);
    quiet         — не шуметь в лог при неудаче (для перебора хостов).
    """
    last_err = None
    for attempt in range(1, max_attempts + 1):
        try:
            resp = session.get(url, params=params, timeout=REQUEST_TIMEOUT)
            if resp.status_code in missing_codes:
                return None
            resp.raise_for_status()
            if not resp.content:
                return None
            return resp.json()
        except (requests.RequestException, ValueError) as e:
            last_err = e
            if attempt < max_attempts:
                wait = BACKOFF_BASE ** attempt
                if not quiet:
                    log.warning("Запрос не удался (попытка %d/%d): %s. Ждём %d сек…",
                                attempt, max_attempts, e, wait)
                time.sleep(wait)
    if not quiet:
        log.error("Не удалось получить данные с %s: %s", url, last_err)
    return None


# --------------------------------------------------------------------------- #
# Шаг 1. Определение imtId (склейки), бренда и названия по артикулу
# --------------------------------------------------------------------------- #

def resolve_card_via_api(session: requests.Session, nm_id: int) -> dict:
    """
    Способ 1 (быстрый): поисковый API карточки card.wb.ru.
    Работает для товаров, которые есть в продаже/в выдаче.
    Возвращает dict imt_id/brand/name (значения могут быть None).
    """
    for host in CARD_HOSTS:
        params = dict(CARD_PARAMS, nm=str(nm_id))
        data = get_json(session, host, params=params)
        if not data:
            continue
        products = (data.get("data") or {}).get("products") or []
        if not products:
            continue
        p = products[0]
        return {
            "imt_id": p.get("root"),          # root == imtId (идентификатор склейки)
            "brand": (p.get("brand") or "").strip() or None,
            "name": (p.get("name") or "").strip() or None,
        }
    return {"imt_id": None, "brand": None, "name": None}


def _basket_host_guess(vol: int) -> str:
    """Предполагаемый номер CDN-сервера (basket-XX) по диапазону vol. Это лишь стартовая
    догадка; если она не сработает, мы переберём остальные хосты."""
    ranges = [
        (0, 143, "01"), (144, 287, "02"), (288, 431, "03"), (432, 719, "04"),
        (720, 1007, "05"), (1008, 1061, "06"), (1062, 1115, "07"), (1116, 1169, "08"),
        (1170, 1313, "09"), (1314, 1601, "10"), (1602, 1655, "11"), (1656, 1919, "12"),
        (1920, 2045, "13"), (2046, 2189, "14"), (2190, 2405, "15"), (2406, 2621, "16"),
        (2622, 2837, "17"), (2838, 3053, "18"), (3054, 3269, "19"), (3270, 3485, "20"),
        (3486, 3701, "21"), (3702, 3917, "22"), (3918, 4133, "23"), (4134, 4349, "24"),
        (4350, 4565, "25"),
    ]
    for lo, hi, host in ranges:
        if lo <= vol <= hi:
            return host
    return "26"


def resolve_card_via_basket(session: requests.Session, nm_id: int) -> dict:
    """
    Способ 2 (надёжный): статический card.json на CDN-серверах WB (basket-XX).
    Здесь у КАЖДОГО товара лежит поле imt_id — работает, даже если товара нет
    в поисковой выдаче (распродан/скрыт). Это и позволяет указывать только артикул.

    Перебираем хосты, начиная с наиболее вероятного, и берём первый ответ с imt_id.
    """
    vol = nm_id // 100000
    part = nm_id // 1000
    guess = _basket_host_guess(vol)
    # порядок: сначала догадка, затем все остальные номера хостов
    host_numbers = [guess] + [f"{i:02d}" for i in range(1, 31) if f"{i:02d}" != guess]
    domains = ("wbbasket.ru", "wb.ru")  # WB сменил домен; пробуем оба
    for host in host_numbers:
        for dom in domains:
            url = (f"https://basket-{host}.{dom}/vol{vol}/part{part}/"
                   f"{nm_id}/info/ru/card.json")
            data = get_json(session, url, max_attempts=1,
                            missing_codes=(403, 404), quiet=True)
            if data and data.get("imt_id"):
                selling = data.get("selling") or {}
                brand = (selling.get("brand_name") or data.get("brand")
                         or selling.get("brand") or "").strip() or None
                name = (data.get("imt_name") or data.get("subj_name") or "").strip() or None
                return {"imt_id": data.get("imt_id"), "brand": brand, "name": name}
    return {"imt_id": None, "brand": None, "name": None}


def resolve_card(session: requests.Session, nm_id: int) -> dict:
    """
    По артикулу (nmId) определяет imt_id (склейку), бренд и название.
    Использует два независимых источника, чтобы не требовать ручной ввод imtId:
      1) card.wb.ru — быстро, но только для товаров «в выдаче»;
      2) basket CDN card.json — надёжно, работает и для скрытых/распроданных.
    Если первый не дал imt_id — автоматически пробуем второй и дополняем бренд/название.
    """
    res = resolve_card_via_api(session, nm_id)
    if res.get("imt_id"):
        return res

    log.info("Артикул %s: card.wb.ru не дал imtId, пробую CDN (card.json)…", nm_id)
    alt = resolve_card_via_basket(session, nm_id)
    # объединяем: imt_id берём из CDN, бренд/название — откуда есть
    return {
        "imt_id": alt.get("imt_id") or res.get("imt_id"),
        "brand": alt.get("brand") or res.get("brand"),
        "name": alt.get("name") or res.get("name"),
    }


# --------------------------------------------------------------------------- #
# Шаг 2. Загрузка всех отзывов склейки по imtId
# --------------------------------------------------------------------------- #

def fetch_feedbacks(session: requests.Session, imt_id: int) -> dict:
    """
    Перебирает хосты/версии feedbacks и возвращает первый непустой ответ.
    Возвращает dict с ключами:
      feedbacks (list), count (int), count_with_text (int).
    """
    for tmpl in FEEDBACK_ENDPOINTS:
        url = tmpl.format(imt=imt_id)
        data = get_json(session, url)
        if not data:
            continue
        feedbacks = data.get("feedbacks") or []
        if feedbacks:
            return {
                "feedbacks": feedbacks,
                "count": data.get("feedbackCount") or len(feedbacks),
                "count_with_text": data.get("feedbackCountWithText") or 0,
            }
        time.sleep(REQUEST_DELAY)
    return {"feedbacks": [], "count": 0, "count_with_text": 0}


# --------------------------------------------------------------------------- #
# Разбор и фильтрация отзывов
# --------------------------------------------------------------------------- #

def parse_date(raw: Optional[str]) -> Optional[datetime]:
    """Парсит дату отзыва (ISO 8601, например '2025-05-14T10:20:30Z') в datetime (UTC-aware)."""
    if not raw:
        return None
    s = str(raw).strip().replace("Z", "+00:00")
    try:
        dt = datetime.fromisoformat(s)
    except ValueError:
        # запасной разбор только даты
        try:
            dt = datetime.strptime(s[:10], "%Y-%m-%d")
        except ValueError:
            return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


def feedback_nm_id(fb: dict) -> Optional[int]:
    """Достаёт артикул (nmId) отзыва. Лежит либо на верхнем уровне, либо в productDetails."""
    nm = fb.get("nmId")
    if nm is None:
        nm = (fb.get("productDetails") or {}).get("nmId")
    try:
        return int(nm) if nm is not None else None
    except (TypeError, ValueError):
        return None


def review_to_row(fb: dict, dt: Optional[datetime], brand_fallback: str, name_fallback: str) -> list:
    """Преобразует один отзыв (dict из API) в строку CSV согласно REVIEW_COLUMNS."""
    pd = fb.get("productDetails") or {}
    brand = (pd.get("brandName") or brand_fallback or "").strip()
    name = (pd.get("productName") or name_fallback or "").strip()
    size = (pd.get("size") or fb.get("size") or fb.get("matchingSize") or "").strip()
    color = (pd.get("color") or fb.get("color") or "").strip()
    has_photo = 1 if (fb.get("photo") or fb.get("photos")) else 0
    date_str = dt.strftime("%d.%m.%Y") if dt else ""
    return [
        feedback_nm_id(fb) or "",
        brand,
        name,
        color,
        size,
        date_str,
        fb.get("productValuation") or fb.get("valuation") or "",
        has_photo,
        (fb.get("text") or "").strip(),
        (fb.get("pros") or "").strip(),
        (fb.get("cons") or "").strip(),
        fb.get("id") or "",
    ]


def filter_reviews(feedbacks: list, nm_id: int, days: int,
                   brand_fallback: str, name_fallback: str) -> list:
    """
    Применяет все фильтры ТЗ:
      • только нужный артикул (защита от склейки),
      • только отзывы с текстом длиной >= MIN_TEXT_LEN,
      • только за последние `days` дней.
    Возвращает список готовых строк CSV (отсортированных по дате, новые сверху).
    """
    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    rows = []
    for fb in feedbacks:
        # 1) фильтр по артикулу — главное правило против склейки
        if feedback_nm_id(fb) != nm_id:
            continue
        # 2) фильтр по тексту
        text = (fb.get("text") or "").strip()
        if len(text) < MIN_TEXT_LEN:
            continue
        # 3) фильтр по дате
        dt = parse_date(fb.get("createdDate") or fb.get("updatedDate"))
        if dt is not None and dt < cutoff:
            continue
        rows.append((dt, review_to_row(fb, dt, brand_fallback, name_fallback)))
    # новые отзывы сверху; отзывы без даты — в конец
    rows.sort(key=lambda r: r[0] or datetime.min.replace(tzinfo=timezone.utc), reverse=True)
    return [r[1] for r in rows]


# --------------------------------------------------------------------------- #
# Запись CSV
# --------------------------------------------------------------------------- #

def safe_folder_name(value: str) -> str:
    """Делает строку пригодной для имени папки в Windows/Unix."""
    value = (value or "").strip()
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
    return value or "Без бренда"


def write_reviews_csv(path: Path, rows: list) -> None:
    """Записывает отзывы в CSV (UTF-8 с BOM, разделитель ';')."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding=CSV_ENCODING, newline="") as f:
        writer = csv.writer(f, delimiter=CSV_DELIMITER, quoting=csv.QUOTE_MINIMAL)
        writer.writerow(REVIEW_COLUMNS)
        writer.writerows(rows)


def write_summary_csv(path: Path, summary_rows: list) -> None:
    """Записывает сопровождающий файл-сводку."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding=CSV_ENCODING, newline="") as f:
        writer = csv.writer(f, delimiter=CSV_DELIMITER, quoting=csv.QUOTE_MINIMAL)
        writer.writerow(SUMMARY_COLUMNS)
        writer.writerows(summary_rows)


# --------------------------------------------------------------------------- #
# Обработка одного артикула
# --------------------------------------------------------------------------- #

def process_article(session: requests.Session, nm_id: int, out_root: Path, days: int) -> dict:
    """
    Полный цикл по одному артикулу. Никогда не бросает исключение наружу —
    при любой ошибке возвращает статус с описанием, чтобы выгрузка продолжалась.
    """
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")
    url = f"https://www.wildberries.ru/catalog/{nm_id}/detail.aspx"
    result = {
        "nm_id": nm_id, "brand": "", "name": "", "url": url,
        "found": 0, "found_with_text": 0, "saved": 0,
        "csv_path": "", "status": "OK", "error": "", "time": now_str,
    }

    try:
        log.info("Артикул %s: получаю карточку…", nm_id)
        card = resolve_card(session, nm_id)
        result["brand"] = card["brand"] or ""
        result["name"] = card["name"] or ""

        if not card["imt_id"]:
            result["status"] = "Ошибка"
            result["error"] = "Не удалось определить imtId (карточка недоступна)"
            log.error("Артикул %s: %s", nm_id, result["error"])
            return result

        time.sleep(REQUEST_DELAY)
        log.info("Артикул %s: imtId=%s, загружаю отзывы…", nm_id, card["imt_id"])
        fb_data = fetch_feedbacks(session, card["imt_id"])
        result["found"] = fb_data["count"]
        result["found_with_text"] = fb_data["count_with_text"]

        rows = filter_reviews(
            fb_data["feedbacks"], nm_id, days,
            brand_fallback=result["brand"], name_fallback=result["name"],
        )
        result["saved"] = len(rows)

        # Папка: <выгрузка>/<Бренд>/<Артикул>/reviews.csv
        brand_folder = safe_folder_name(result["brand"]) if result["brand"] else "Без бренда"
        csv_path = out_root / brand_folder / str(nm_id) / "reviews.csv"
        write_reviews_csv(csv_path, rows)
        result["csv_path"] = str(csv_path)

        if result["saved"] == 0:
            result["status"] = "Нет отзывов"
            log.warning("Артикул %s: подходящих отзывов не найдено (создан пустой CSV).", nm_id)
        else:
            log.info("Артикул %s: сохранено %d отзывов -> %s", nm_id, result["saved"], csv_path)

    except Exception as e:  # ловим всё, чтобы не прерывать общую выгрузку
        result["status"] = "Ошибка"
        result["error"] = f"{type(e).__name__}: {e}"
        log.exception("Артикул %s: непредвиденная ошибка", nm_id)

    return result


# --------------------------------------------------------------------------- #
# Ввод артикулов: файл / аргументы / интерактивный (Ctrl+V)
# --------------------------------------------------------------------------- #

def extract_articles(text: str) -> list:
    """
    Достаёт артикулы из произвольного текста: строки, запятые, ссылки WB.
    Корректно работает со ссылками вида .../catalog/<артикул>/feedbacks?imtId=<склейка>:
    берётся именно артикул из catalog/<...>, а imtId игнорируется (это не артикул).
    """
    result = []
    # 1) Из ссылок берём артикул из сегмента catalog/<digits>
    for m in re.findall(r"catalog/(\d{4,})", text):
        result.append(int(m))
    # 2) Убираем из текста уже разобранные URL-части и параметр imtId,
    #    чтобы не принять imtId или хвост ссылки за отдельный артикул.
    cleaned = re.sub(r"https?://\S+", " ", text)        # целые ссылки
    cleaned = re.sub(r"imtId=\d+", " ", cleaned, flags=re.I)
    cleaned = re.sub(r"^\s*#.*$", " ", cleaned, flags=re.M)  # строки-комментарии
    # 3) Оставшиеся «голые» числа считаем артикулами
    for m in re.findall(r"\d{4,}", cleaned):
        result.append(int(m))
    return result


def read_articles_from_file(path: str) -> list:
    """Читает артикулы из TXT/CSV (любой разделитель, можно вставлять ссылки)."""
    with open(path, "r", encoding="utf-8-sig") as f:
        return extract_articles(f.read())


def interactive_input() -> list:
    """
    Минимальный UI: пользователь вставляет артикулы (Ctrl+V) — каждый с новой строки,
    через запятую/пробел или прямо ссылками WB. Завершение — пустая строка или Ctrl+D.
    """
    print("\n=== Выгрузка отзывов Wildberries ===")
    print("Вставьте артикулы WB (можно ссылками, через запятую или с новой строки).")
    print("Когда закончите — нажмите Enter на пустой строке (или Ctrl+D / Ctrl+Z).\n")
    lines = []
    try:
        while True:
            line = input("> ")
            if line.strip() == "":
                break
            lines.append(line)
    except EOFError:
        pass
    return extract_articles("\n".join(lines))


def dedup_keep_order(items: list) -> list:
    seen, out = set(), []
    for x in items:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #

def parse_args(argv=None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Выгрузка отзывов Wildberries по артикулам (через публичный API WB).",
    )
    p.add_argument("articles", nargs="*", help="Артикулы WB через пробел (можно ссылки).")
    p.add_argument("-f", "--file", help="Путь к TXT/CSV со списком артикулов.")
    p.add_argument("-o", "--output", default=".", help="Каталог, где создать папку выгрузки.")
    p.add_argument("-d", "--days", type=int, default=DEFAULT_DAYS,
                   help=f"За сколько последних дней брать отзывы (по умолчанию {DEFAULT_DAYS}).")
    p.add_argument("--delay", type=float, default=REQUEST_DELAY,
                   help="Задержка между запросами, сек.")
    return p.parse_args(argv)


def main(argv=None) -> int:
    global REQUEST_DELAY
    args = parse_args(argv)
    REQUEST_DELAY = args.delay

    # Собираем артикулы из всех источников
    articles: list = []
    if args.file:
        try:
            articles += read_articles_from_file(args.file)
        except OSError as e:
            log.error("Не удалось прочитать файл %s: %s", args.file, e)
            return 2
    if args.articles:
        articles += extract_articles(" ".join(args.articles))
    if not articles:
        articles = interactive_input()

    articles = dedup_keep_order(articles)
    if not articles:
        log.error("Не передано ни одного артикула. Завершаю работу.")
        return 1

    # Папка выгрузки: "выгрузка отзывов ДД-ММ-ГГГГ ЧЧ-ММ"
    stamp = datetime.now().strftime("%d-%m-%Y %H-%M")
    out_root = Path(args.output) / f"выгрузка отзывов {stamp}"
    out_root.mkdir(parents=True, exist_ok=True)
    log.info("Папка выгрузки: %s", out_root)
    log.info("Артикулов к обработке: %d | период: последние %d дн.", len(articles), args.days)

    session = build_session()
    summary_rows = []
    for i, nm_id in enumerate(articles, 1):
        log.info("[%d/%d] === Артикул %s ===", i, len(articles), nm_id)
        res = process_article(session, nm_id, out_root, args.days)
        summary_rows.append([
            res["nm_id"], res["brand"], res["name"], res["url"],
            res["found"], res["found_with_text"], res["saved"],
            res["csv_path"], res["status"], res["error"], res["time"],
        ])
        if i < len(articles):
            time.sleep(REQUEST_DELAY)

    # Сопровождающий файл-сводка в корне папки выгрузки
    summary_path = out_root / "сводка_выгрузки.csv"
    write_summary_csv(summary_path, summary_rows)
    log.info("Сводка сохранена: %s", summary_path)

    # Итог по тому, что не реализовалось (п.4 ТЗ)
    print_final_report(summary_rows)
    return 0


def print_final_report(summary_rows: list) -> None:
    """Печатает итог: сколько успешно, где были ошибки/пустые выгрузки."""
    ok = sum(1 for r in summary_rows if r[8] == "OK")
    empty = sum(1 for r in summary_rows if r[8] == "Нет отзывов")
    errors = [r for r in summary_rows if r[8] == "Ошибка"]
    total_saved = sum(int(r[6] or 0) for r in summary_rows)

    print("\n" + "=" * 60)
    print("ИТОГ ВЫГРУЗКИ")
    print("=" * 60)
    print(f"  Артикулов обработано : {len(summary_rows)}")
    print(f"  Успешно (с отзывами) : {ok}")
    print(f"  Без подходящих отзыв.: {empty}")
    print(f"  С ошибками           : {len(errors)}")
    print(f"  Всего отзывов сохран.: {total_saved}")
    if errors:
        print("\n  Не удалось выгрузить:")
        for r in errors:
            print(f"    • Артикул {r[0]}: {r[9]}")
    print("=" * 60 + "\n")


def _pause_if_double_click() -> None:
    """На Windows при запуске двойным кликом не даём окну закрыться сразу."""
    if os.name == "nt" and len(sys.argv) == 1:
        try:
            input("\nГотово. Нажмите Enter, чтобы закрыть окно…")
        except EOFError:
            pass


if __name__ == "__main__":
    code = 0
    try:
        code = main()
    except KeyboardInterrupt:
        log.warning("Прервано пользователем.")
        code = 130
    except Exception:
        log.exception("Критическая ошибка")
        code = 1
    finally:
        _pause_if_double_click()
    sys.exit(code)
