#!/usr/bin/env python3
"""
Локальний веб-інтерфейс до download_pleanary_pdf.py.

Запуск:
    python3 gui_app.py

Потім відкрити в браузері http://127.0.0.1:5005, завантажити PDF
порядку денного — скрипт сам скачає всі відповідні файли "На пленарне
засідання" у папку downloaded_plenary_pdfs (поруч зі скриптом).
"""

import html
import collections
import copy
import csv
import io
import json
import mimetypes
import os
import platform
import re
import shutil
import smtplib
import ssl
import subprocess
import sys
import tempfile
import threading
import uuid
import zipfile
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

import fitz  # PyMuPDF
import requests
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from flask import Flask, jsonify, render_template_string, request

import agenda_doc
import download_pleanary_pdf as core
from icons import sprite as icon_sprite

FROZEN = getattr(sys, "frozen", False)  # запущено як зібраний PyInstaller-застосунок
# Усередині зібраного застосунку писати поруч із кодом не можна (Program Files
# лише для читання), тому дані живуть у "Документи/Порядок денний".
BASE_DIR = (Path.home() / "Documents" / "Порядок денний") if FROZEN else Path(__file__).resolve().parent
BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
BASE_DIR.mkdir(parents=True, exist_ok=True)
DEFAULT_OUT_DIR = BASE_DIR / "downloaded_plenary_pdfs"
SETTINGS_PATH = BASE_DIR / "settings.json"
GENERATE_UPLOADS_DIR = BASE_DIR / "generation_uploads"
GENERATED_DOC_PATH = BASE_DIR / "Згенерованний ПР.docx"

NO_SPEAKER_LABEL = "БЕЗ_ДОПОВИДАЧА"

DEFAULT_SECTION_TITLES = {
    1: "Загальні питання",
    2: "Про питання найменування, перейменування, уточнення назв об’єктів міського підпорядкування",
    3: "Про питання власності",
    4: "Про питання містобудування та землекористування",
    5: "Про розгляд проєктів рішень Київської міської ради на виконання рішень/постанов судів",
    6: "Різне",
}
SECTION_TITLES_PATH = BASE_DIR / "section_titles.json"
# Розділи — дерево до 4 рівнів з ідентифікаторами "1", "1.1", "1.1.1", "1.1.1.1". Верхні (1–6) є завжди,
# підрозділи додає користувач; їхні номери йдуть підряд (1.1, 1.2, …) і перераховуються при видаленні.
MAX_SECTION_DEPTH = 4
NEW_SUBSECTION_TITLE = "Новий підрозділ"
SECTION_ID_RE = re.compile(r"^[1-9]\d*(\.[1-9]\d*){0,3}$")


def section_key(section_id: str) -> list:
    return [int(part) for part in section_id.split(".")]


def section_depth(section_id: str) -> int:
    return len(section_id.split("."))


def load_section_titles() -> dict:
    """Назви розділів і підрозділів: {"1": ..., "1.1": ...}. Підрозділ без наявного батька відкидається."""
    titles = {str(n): title for n, title in DEFAULT_SECTION_TITLES.items()}
    try:
        saved = json.loads(SECTION_TITLES_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return titles
    if not isinstance(saved, dict):
        return titles
    for sid in sorted((k for k in saved if SECTION_ID_RE.match(str(k))), key=section_key):
        title = saved[sid]
        if not isinstance(title, str) or not title.strip():
            continue
        if section_depth(sid) == 1:
            if sid in titles:
                titles[sid] = title
        elif sid.rsplit(".", 1)[0] in titles:
            titles[sid] = title
    return titles


def save_section_titles(titles: dict):
    SECTION_TITLES_PATH.write_text(
        json.dumps({sid: titles[sid] for sid in sorted(titles, key=section_key)}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def section_ids(titles: dict = None) -> list:
    """Усі розділи й підрозділи в порядку документа: 1, 1.1, 1.1.1, 1.2, 2, …"""
    return sorted(SECTION_TITLES if titles is None else titles, key=section_key)


def plan_section_delete(titles: dict, section_id: str):
    """Що станеться при видаленні підрозділу: (нові назви, {старий id: новий id}, множина видалених).
    Наступні сусіди (і їхні вкладені) отримують номери на один менший; видалені йдуть до батька."""
    parts = section_id.split(".")
    depth, parent = len(parts), ".".join(parts[:-1])
    removed = {k for k in titles if k == section_id or k.startswith(section_id + ".")}
    mapping = {k: parent for k in removed}
    position = int(parts[-1])
    for k in titles:
        kp = k.split(".")
        if k not in removed and len(kp) >= depth and kp[:depth - 1] == parts[:-1] and int(kp[depth - 1]) > position:
            shifted = list(kp)
            shifted[depth - 1] = str(int(kp[depth - 1]) - 1)
            mapping[k] = ".".join(shifted)
    new_titles = {mapping.get(k, k): v for k, v in titles.items() if k not in removed}
    return new_titles, mapping, removed


SECTION_TITLES = load_section_titles()

# Назва файлу проєкту: "183-25.pdf" — номер рішення 183, рік 2025 ("25" або "2025").
MAX_SEARCH_PAGES = 5   # скільки сторінок видачі пошуку переглядати
MAX_CANDIDATES = 40    # скільки знайдених проєктів перевіряти на рік


def check_decision_filename(filename: str):
    """Перевіряє назву файлу проєкту на формат НОМЕР-РІК.pdf.
    Повертає (номер, рік, помилка): помилка — текст для користувача або None."""
    m = re.match(r"^(\d+)(.*)$", filename)
    if not m:
        return None, None, "назва не починається з номера рішення"
    number, rest = m.groups()
    ym = re.match(r"-(\d+)", rest)
    if not ym:
        return number, None, f"після номера {number} немає дефіса та року"
    digits = ym.group(1)
    if len(digits) not in (2, 4):
        return number, None, f"рік «{digits}» має складатися з 2 або 4 цифр"
    year = 2000 + int(digits) if len(digits) == 2 else int(digits)
    if not (2000 <= year <= datetime.now().year + 1):
        return number, None, f"рік {year} виглядає неправильно"
    return number, year, None


def suggest_decision_filename(filename: str):
    """Підказка для назви на кшталт "рішення 183 від 2025.pdf" -> "183-25.pdf" (або None)."""
    m = re.match(r"^\D*?(\d{1,6})\D+(\d{4}|\d{2})\D*\.pdf$", filename, re.IGNORECASE)
    if not m:
        return None
    guess = f"{m.group(1)}-{m.group(2)[-2:]}.pdf"
    return guess if check_decision_filename(guess)[2] is None else None


def generate_name_entries(paths) -> list:
    """Зауваження до назв файлів (у такому ж вигляді, як зауваження до порядку денного)."""
    entries = []
    for path in sorted(paths, key=lambda p: p.name):
        _, _, error = check_decision_filename(path.name)
        if error:
            entries.append({
                "id": f"name|{path.name}", "kind": "некоректна назва файлу", "filename": path.name,
                "text": f"{error}. Очікується формат НОМЕР-РІК.pdf, напр. 183-25.pdf. "
                        "Перейменуйте файл або позначте галочкою, щоб пропустити його.",
                "suggest": suggest_decision_filename(path.name),
            })
    return entries


def validate_new_filename(raw: str, folder: Path, current: str) -> str:
    """Нова назва від користувача -> коректна назва файлу (".pdf" дописується сам)."""
    name = (raw or "").strip().strip('"')
    if not name:
        raise ValueError("Вкажіть нову назву файлу.")
    if re.search(r'[\\/:*?"<>|]', name):
        raise ValueError('Назва не може містити символи \\ / : * ? " < > |')
    name = re.sub(r"\s+(\.pdf)$", r"\1", name, flags=re.IGNORECASE)
    if "." not in name:
        name += ".pdf"
    elif not name.lower().endswith(".pdf"):
        raise ValueError("Файл має бути у форматі .pdf.")
    _, _, error = check_decision_filename(name)
    if error:
        raise ValueError(f"{error}. Приклад правильної назви: 183-25.pdf")
    if name != current and (folder / name).exists():
        raise ValueError("Файл з такою назвою вже є в цій папці.")
    return name


def _year_of(date_text):
    match = re.search(r"(?<!\d)(\d{4})(?!\d)", date_text or "")
    return int(match.group(1)) if match else None


def _search_candidates(session, number):
    """Посилання на проєкти з пошуку — по всіх сторінках видачі, у порядку видачі.
    Зупиняється, коли сторінка порожня або нічого нового не дала."""
    seen = set()
    for page in range(MAX_SEARCH_PAGES):
        url = SEARCH_URL_TEMPLATE.format(number=number) + (f"&page={page}" if page else "")
        resp = session.get(url, timeout=20)
        resp.raise_for_status()
        fresh = [h for h in dict.fromkeys(m.group(1) for m in SEARCH_RESULT_LINK_RE.finditer(resp.text))
                 if h not in seen]
        if not fresh:
            return
        seen.update(fresh)
        yield from fresh


def _number_on_page(html: str, number: str) -> bool:
    """Чи згадується на сторінці проєкту номер рішення ("08/231-183/ПР" або "183/ПР")."""
    n = re.escape(number)
    return bool(re.search(rf"(?<!\d){n}\s*/\s*ПР|{re.escape(DECISION_PREFIX)}\s*-\s*{n}(?!\d)", html, re.IGNORECASE))


def find_project_page(session, number: str, year):
    """Шукає на сайті проєкт рішення за номером і роком. Номер у різні роки повторюється
    (183/ПР є і за 2025, і за 2026), тому серед результатів пошуку береться той, чия дата
    реєстрації припадає на потрібний рік. Повертає (url, html сторінки, примітка або None);
    якщо нічого підходящого немає — LookupError з поясненням."""
    fallback, years_seen, checked = None, [], 0
    for href in _search_candidates(session, number):
        project_url = SEARCH_RESULT_BASE_URL + href
        if year is None:  # рік у назві не вказано — як і раніше, перший результат
            resp = session.get(project_url, timeout=20)
            resp.raise_for_status()
            return project_url, resp.text, "рік у назві файлу не вказано — взято перший результат пошуку"
        if checked >= MAX_CANDIDATES:
            break
        checked += 1
        try:
            resp = session.get(project_url, timeout=20)
            resp.raise_for_status()
        except Exception:
            continue
        page_year = _year_of(extract_project_date(resp.text))
        if page_year != year:
            years_seen.append(page_year)
            continue
        if _number_on_page(resp.text, number):
            return project_url, resp.text, None
        if fallback is None:
            fallback = (project_url, resp.text)
    if fallback:
        return fallback[0], fallback[1], (f"рік {year} збігається, але номер {number} на сторінці проєкту "
                                          f"не підтверджено — перевірте посилання")
    if not checked:
        raise LookupError("у результатах пошуку не знайдено жодного посилання на проєкт рішення")
    shown = ", ".join(sorted({str(y) if y else "?" for y in years_seen}))
    raise LookupError(f"серед знайдених проєктів ({checked}) немає за {year} рік (роки знайдених: {shown})")


SEARCH_URL_TEMPLATE = "https://old.kmr.gov.ua/uk/search-result?search_api_views_fulltext={number}&type=decision"
# Текст посилання на сторінці пошуку не завжди містить номер у кінці і
# буває або "Проєкт", або "Проект" (без "є") залежно від запису — тож
# орієнтуємось лише на href, а не на точний текст.
SEARCH_RESULT_LINK_RE = re.compile(
    r'<a href="(/uk/content/proy?ekt-rishennya-kyyivskoyi-miskoyi-rady-(\d+))"',
    re.IGNORECASE,
)
SEARCH_RESULT_BASE_URL = "https://old.kmr.gov.ua"

# Заголовок рішення — перший <p> після <h1 id="page-title"> на сторінці проєкту.
# Може бути обгорнутий у різну розмітку (div, table/tr/td тощо) залежно від
# документа, тому пропускаємо будь-що між ними, а не конкретний набір тегів.
PROJECT_TITLE_RE = re.compile(
    r'<h1[^>]*id="page-title"[^>]*>.*?</h1>.*?<p[^>]*>(.*?)</p>',
    re.DOTALL | re.IGNORECASE,
)
PROJECT_DATE_RE = re.compile(
    r'Дата\s+реєстрації\s+проекту\s+рішення.*?class="date-display-single"[^>]*>([^<]+)</span>',
    re.DOTALL | re.IGNORECASE,
)

# Індекс "08/231" — незмінний префікс номера рішення, не зчитується з файлу.
DECISION_PREFIX = "08/231"

# Різні документи описують доповідача по-різному, тому пробуємо кілька шаблонів
# (від найточнішого до найзагальнішого):
# 1) "Доповідачем проєкту рішення ... є <посада> Прізвище Ім'я По-батькові (тел. ...)."
# 2) "Доповідач на пленарному засіданні Київської міської ради - <посада> Прізвище Ім'я По-батькові, контактний телефон ..."
# 3) "...та доповідачем на пленарному засіданні є <посада> - Прізвище Ім'я По-батькові (тел. ...)."
# 4) "..., доповідачем - <посада> Прізвище І. І. (тел. ...)." — ім'я вже дане
#    ініціалами (не повними словами). Навмисно якір "доповідачем" в однині,
#    щоб не зачепити "співдоповідачі:" (кілька осіб одразу — там свідомо
#    лишаємо БЕЗ_ДОПОВИДАЧА, бо незрозуміло, кого з кількох обирати).
# 5) запасний варіант: будь-яке слово "доповідач..." і найближче після нього
#    ПІБ (3 слова з великої літери) перед комою чи дужкою — на випадок ще
#    не бачених формулювань.
SPEAKER_BLOCK_PATTERNS = [
    re.compile(
        r"Доповідачем\s+проєкту\s+рішення\s+Київської\s+міської\s+ради.*?"
        r"([А-ЯІЇЄҐ][’'а-яіїєґ]+)\s+([А-ЯІЇЄҐ][’'а-яіїєґ]+)\s+([А-ЯІЇЄҐ][’'а-яіїєґ]+)\s*\(тел\.",
        re.DOTALL,
    ),
    re.compile(
        r"Доповідач\s+на\s+пленарному\s+засіданні\s+Київської\s+міської\s+ради.*?"
        r"([А-ЯІЇЄҐ][’'а-яіїєґ]+)\s+([А-ЯІЇЄҐ][’'а-яіїєґ]+)\s+([А-ЯІЇЄҐ][’'а-яіїєґ]+)\s*,",
        re.DOTALL,
    ),
    re.compile(
        r"[Дд]оповідачем\s+на\s+пленарному\s+засіданні\s+є.*?"
        r"([А-ЯІЇЄҐ][’'а-яіїєґ]+)\s+([А-ЯІЇЄҐ][’'а-яіїєґ]+)\s+([А-ЯІЇЄҐ][’'а-яіїєґ]+)\s*\(тел\.",
        re.DOTALL,
    ),
    re.compile(
        r"(?<!спів)[Дд]оповідачем\b[\s\S]{0,300}?"
        r"([А-ЯІЇЄҐ][’'а-яіїєґ]+)\s+([А-ЯІЇЄҐ])\.\s*([А-ЯІЇЄҐ])\.",
    ),
    re.compile(
        r"[Дд]оповіда[чц][а-яіїєґ]*\b[\s\S]{0,400}?"
        r"([А-ЯІЇЄҐ][’'а-яіїєґ]+)\s+([А-ЯІЇЄҐ][’'а-яіїєґ]+)\s+([А-ЯІЇЄҐ][’'а-яіїєґ]+)\s*[,(]",
    ),
]
# Формулювання, де імʼя йде ПЕРЕД прізвищем, а прізвище — великими літерами, без по батькові:
# "Доповідач: директор Департаменту земельних ресурсів Валентина ПЕЛИХ." -> "Пелих В."
# Після прізвища обовʼязкове завершення (крапка, кома, дужка, кінець рядка), щоб не сплутати
# з назвою посади, де трапляється слово великими літерами ("Департамент ЗЕМЕЛЬНИХ ресурсів").
SPEAKER_NAME_FIRST_PATTERNS = [
    re.compile(
        r"[Дд]оповідач\s*:[\s\S]{0,250}?"
        r"([А-ЯІЇЄҐ][’'а-яіїєґ]+)\s+([А-ЯІЇЄҐ][А-ЯІЇЄҐ’'\-]{1,})(?=\s*(?:[.,(;]|$))",
        re.MULTILINE,
    ),
]
OCR_LANG = "ukr"
OCR_DPI = 300


def _surname_case(surname: str) -> str:
    """"ПЕЛИХ" -> "Пелих", "КОВАЛЬ-ПЕТРЕНКО" -> "Коваль-Петренко"."""
    return "-".join(part.capitalize() for part in surname.split("-"))

def _tesseract_candidates():
    """tesseract не завжди потрапляє в PATH після встановлення, а на Windows
    типове місце інсталяції ще й залежить від того, чи були права
    адміністратора (Program Files) чи ні (LOCALAPPDATA). Перевіряємо все."""
    candidates = [
        str(BUNDLE_DIR / "tesseract" / "tesseract.exe"),
        str(BUNDLE_DIR / "tesseract" / "tesseract"),
        "tesseract",
    ]
    program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
    program_files_x86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    local_appdata = os.environ.get("LOCALAPPDATA", "")
    candidates += [
        rf"{program_files}\Tesseract-OCR\tesseract.exe",
        rf"{program_files_x86}\Tesseract-OCR\tesseract.exe",
    ]
    if local_appdata:
        candidates += [
            rf"{local_appdata}\Programs\Tesseract-OCR\tesseract.exe",
            rf"{local_appdata}\Tesseract-OCR\tesseract.exe",
        ]
    candidates += [
        "/opt/homebrew/bin/tesseract",
        "/usr/local/bin/tesseract",
        "/usr/bin/tesseract",
    ]
    return candidates


def _tesseract_env(tesseract_cmd: str) -> dict:
    """Для вбудованого Tesseract вказуємо його власну папку tessdata."""
    env = os.environ.copy()
    bundled_tessdata = Path(tesseract_cmd).parent / "tessdata"
    if bundled_tessdata.is_dir():
        env["TESSDATA_PREFIX"] = str(bundled_tessdata)
    return env


_TESSERACT_CMD = None
_TESSERACT_RESOLVED = False


def find_tesseract_cmd():
    global _TESSERACT_CMD, _TESSERACT_RESOLVED
    if _TESSERACT_RESOLVED:
        return _TESSERACT_CMD
    for candidate in _tesseract_candidates():
        resolved = shutil.which(candidate) or (candidate if Path(candidate).is_file() else None)
        if resolved:
            _TESSERACT_CMD = resolved
            break
    _TESSERACT_RESOLVED = True
    return _TESSERACT_CMD

app = Flask(__name__)
current_agenda = {"name": None, "out_dir": None}
jobs = {}  # job_id -> {"log": [str], "done": bool, "out_dir": str, "audit": [dict]}

PAGE = """
<!doctype html>
<html lang="uk">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Порядок денний</title>
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'%3E%3Ctext y='.9em' font-size='90'%3E%F0%9F%93%84%3C/text%3E%3C/svg%3E">
<style>
  :root {
    --bg: #f2f4f8;
    --card: #ffffff;
    --text: #1a1d29;
    --text-muted: #6b7280;
    --border: #e5e7eb;
    --primary: #4f46e5;
    --primary-hover: #4338ca;
    --primary-soft: #eef2ff;
    --green: #059669;
    --green-soft: #ecfdf5;
    --amber: #d97706;
    --amber-soft: #fffbeb;
    --red: #dc2626;
    --red-soft: #fef2f2;
    --radius: 14px;
    --shadow: 0 1px 2px rgba(16,24,40,.04), 0 8px 24px rgba(16,24,40,.06);
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #0f1117;
      --card: #171a24;
      --text: #e8e9ee;
      --text-muted: #9298a8;
      --border: #2a2e3b;
      --primary: #6366f1;
      --primary-hover: #7c7ff2;
      --primary-soft: #1e2036;
      --green: #34d399;
      --green-soft: #0e2a22;
      --amber: #fbbf24;
      --amber-soft: #2e2410;
      --red: #f87171;
      --red-soft: #2e1616;
      --shadow: 0 1px 2px rgba(0,0,0,.3), 0 8px 24px rgba(0,0,0,.35);
    }
  }
  * { box-sizing: border-box; }
  .ic { width: 1em; height: 1em; fill: none; stroke: currentColor; stroke-width: 2; stroke-linecap: round;
    stroke-linejoin: round; vertical-align: -0.125em; flex: none; }
  .mp-link .ic, .mp-del .ic { width: 14px; height: 14px; }
  .mp-edit { border: 0; background: none; color: var(--text-muted); cursor: pointer; padding: 2px 6px;
    border-radius: 6px; display: flex; }
  .mp-edit .ic { width: 14px; height: 14px; }
  .mp-edit:hover { color: var(--primary); background: var(--primary-soft); }
  .mp-btns { display: flex; gap: 8px; margin-top: 8px; }
  .mp-btn.danger { background: var(--red); }
  .mp-btn.danger:hover:not(:disabled) { background: var(--red); filter: brightness(.9); }
  .modal-ov { position: fixed; inset: 0; z-index: 1100; background: rgba(0,0,0,.5); display: flex;
    align-items: center; justify-content: center; padding: 16px; }
  .modal { width: 100%; max-width: 380px; background: var(--card); border: 1px solid var(--border);
    border-radius: 14px; box-shadow: var(--shadow); padding: 20px; color: var(--text); }
  .modal h3 { margin: 0 0 8px; font-size: 16px; }
  .modal p { margin: 0 0 18px; font-size: 13.5px; line-height: 1.55; color: var(--text-muted); overflow-wrap: anywhere; }
  .modal .mp-actions { justify-content: flex-end; margin-top: 0; }
  .mp-link { display: inline-flex; align-items: center; gap: 5px; }
  .mp-del { display: flex; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: var(--bg);
    color: var(--text);
    margin: 0;
    display: flex;
    min-height: 100vh;
    -webkit-font-smoothing: antialiased;
  }
  .main { flex: 1; min-width: 0; padding: 48px 16px; display: flex; justify-content: center; }

  .sidebar {
    flex: none; width: 320px; height: 100vh; position: sticky; top: 0; overflow: auto;
    background: var(--card); border-right: 1px solid var(--border); padding: 18px 10px;
  }
  .sb-title { font-size: 11px; font-weight: 600; letter-spacing: .06em; text-transform: uppercase;
    color: var(--text-muted); padding: 0 8px 10px; }
  .tree ul { list-style: none; margin: 0; padding: 0; }
  /* Пунктирна напрямна під шевроном батьківської папки показує, що саме вона згортає. */
  .tree li > ul { margin-left: 13px; padding-left: 6px; border-left: 1px dotted var(--text-muted);
    border-left-color: color-mix(in srgb, var(--text-muted) 60%, transparent); }
  .tree li:not(.open) > ul { display: none; }
  .trow { display: flex; align-items: center; gap: 6px; padding: 4px 8px; border-radius: 8px;
    cursor: pointer; font-size: 13px; line-height: 1.35; }
  .trow:hover { background: var(--primary-soft); }
  .trow .chev { width: 12px; height: 12px; flex: none; display: flex; color: var(--text-muted); transition: transform .12s; }
  .tree li.open > .trow .chev { transform: rotate(90deg); }
  .trow .nm { flex: 1; min-width: 0; overflow-wrap: anywhere; }
  .trow.root .nm { font-weight: 600; }
  .trow .treveal { flex: none; display: flex; border: 0; background: none; color: var(--primary); cursor: pointer;
    padding: 4px 5px; border-radius: 6px; opacity: .75; }
  .trow .treveal:hover { opacity: 1; background: var(--primary-soft); }
  .trow .ttrash { color: var(--text-muted); }
  .trow .ttrash:hover { color: var(--red); background: var(--red-soft); }
  .mp-links { display: flex; flex-wrap: wrap; gap: 4px 16px; }
  .trow:hover .treveal { opacity: 1; }

  .mp { position: fixed; z-index: 900; width: 330px; max-height: calc(100vh - 16px); overflow: auto;
    background: var(--card); border: 1px solid var(--border); border-radius: 12px; box-shadow: var(--shadow);
    padding: 14px; font-size: 13px; color: var(--text); }
  .mp[hidden] { display: none; }
  .mp-title { font-weight: 600; font-size: 14px; }
  .mp-sub { color: var(--text-muted); font-size: 12px; margin: 2px 0 10px; overflow: hidden;
    text-overflow: ellipsis; white-space: nowrap; }
  .mp-list { max-height: 240px; overflow: auto; border: 1px solid var(--border); border-radius: 10px; }
  .mp-row { display: flex; align-items: center; gap: 10px; padding: 8px 10px; cursor: pointer; }
  .mp-row + .mp-row { border-top: 1px solid var(--border); }
  .mp-row:hover { background: var(--primary-soft); }
  .mp-row .who { flex: 1; min-width: 0; }
  .mp-row .who b { display: block; font-weight: 600; overflow: hidden; text-overflow: ellipsis; }
  .mp-row .who span { display: block; color: var(--text-muted); font-size: 12px; overflow: hidden; text-overflow: ellipsis; }
  .mp-del { border: 0; background: none; color: var(--text-muted); cursor: pointer; font-size: 14px;
    padding: 2px 6px; border-radius: 6px; }
  .mp-del:hover { color: var(--red); background: var(--red-soft); }
  .mp input[type=text], .mp input[type=email], .mp input[type=password] { width: 100%; padding: 8px 10px;
    border-radius: 8px; border: 1px solid var(--border); background: var(--bg); color: var(--text);
    font-size: 13px; margin-top: 6px; }
  .mp select { width: 100%; padding: 8px 10px; border-radius: 8px; border: 1px solid var(--border);
    background: var(--bg); color: var(--text); font-size: 13px; margin-top: 6px; }
  .mp input:focus, .mp select:focus { outline: none; border-color: var(--primary); }
  .mp-link { border: 0; background: none; color: var(--primary); cursor: pointer; font-size: 12.5px; padding: 6px 0 0; }
  .mp-actions { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-top: 10px; }
  .mp-btn { border: 0; border-radius: 9px; padding: 9px 14px; background: var(--primary); color: #fff;
    font-weight: 600; font-size: 13px; cursor: pointer; }
  .mp-btn:hover:not(:disabled) { background: var(--primary-hover); }
  .mp-btn:disabled { opacity: .5; cursor: not-allowed; }
  .mp-btn.ghost { background: transparent; color: var(--text); border: 1px solid var(--border); }
  .mp-status { font-size: 12.5px; margin-top: 8px; color: var(--text-muted); line-height: 1.45; }
  .mp-status.err { color: var(--red); }
  .mp-label { font-size: 11px; font-weight: 600; letter-spacing: .05em; text-transform: uppercase;
    color: var(--text-muted); margin: 2px 0 6px; }
  .mp-chips { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; margin-bottom: 10px; }
  .mp-chips .mp-link { padding: 0 2px; }
  .mp-chip { display: inline-flex; align-items: center; border: 1px solid var(--border); border-radius: 999px;
    font-size: 12px; overflow: hidden; }
  .mp-chip.on { background: var(--primary-soft); border-color: var(--primary); }
  .mp-chip button { border: 0; background: none; color: var(--text); cursor: pointer; font-size: 12px; }
  .mp-chip .chip-main { padding: 4px 4px 4px 10px; font-weight: 600; }
  .mp-chip .chip-edit { display: flex; padding: 4px 8px 4px 4px; color: var(--text-muted); }
  .mp-chip .chip-edit .ic { width: 12px; height: 12px; }
  .mp-chip .chip-edit:hover { color: var(--primary); }
  .mp-info { margin: 8px 0 6px; }
  .mp-info .mp-hint { padding: 0; }
  .mp-info .mp-link { padding: 0; }
  .mp-hint { font-size: 12px; color: var(--text-muted); line-height: 1.45; padding: 2px 0; }
  .trow .treveal .ic { width: 16px; height: 16px; }
  .trow .chev .ic { width: 12px; height: 12px; }
  .trow .ico { display: flex; flex: none; }
  .trow .ico .ic { width: 16px; height: 16px; }
  .trow .ico-root { color: var(--primary); }
  .trow .ico-dir { color: var(--amber); }
  .trow .ico-file { color: var(--text-muted); }
  .tree-empty { font-size: 13px; color: var(--text-muted); padding: 4px 8px; line-height: 1.5; }
  .shell { width: 100%; max-width: 640px; transition: max-width .2s ease; }
  .shell.wide { max-width: 1040px; }
  .head { text-align: center; margin-bottom: 28px; }
  .head .icon { font-size: 34px; line-height: 1; margin-bottom: 10px; color: var(--primary); }
  .head h1 { font-size: 21px; font-weight: 700; margin: 0 0 6px; letter-spacing: -0.01em; }
  .head p { font-size: 14px; color: var(--text-muted); margin: 0; }

  .card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    box-shadow: var(--shadow);
    padding: 24px;
    margin-bottom: 16px;
  }

  .dropzone {
    border: 2px dashed var(--border);
    border-radius: 12px;
    padding: 32px 20px;
    text-align: center;
    cursor: pointer;
    transition: border-color .15s, background .15s;
  }
  .dropzone:hover { border-color: var(--primary); }
  .dropzone.hover { border-color: var(--primary); background: var(--primary-soft); }
  .dropzone .drop-icon { font-size: 28px; line-height: 1; margin-bottom: 8px; color: var(--primary); }
  .dropzone .drop-main { font-size: 14.5px; font-weight: 600; }
  .dropzone .drop-sub { font-size: 12.5px; color: var(--text-muted); margin-top: 3px; }
  input[type=file] { display: none; }

  .file-chip {
    display: none;
    align-items: center;
    gap: 10px;
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 12px 14px;
    margin-top: 4px;
  }
  .file-chip .fi { font-size: 20px; line-height: 1; color: var(--primary); }
  .file-chip .fname { font-size: 13.5px; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .file-chip .fsize { font-size: 12px; color: var(--text-muted); }
  .file-chip .fclear {
    margin-left: auto; cursor: pointer; color: var(--text-muted);
    border: none; background: none; font-size: 16px; padding: 4px 8px; border-radius: 8px;
  }
  .file-chip .fclear:hover { background: var(--bg); color: var(--text); }

  label.field-label { display: block; font-size: 12.5px; font-weight: 600; color: var(--text-muted); margin: 18px 0 6px; text-transform: uppercase; letter-spacing: .03em; }
  .field-wrap { position: relative; }
  .field-wrap .fi-prefix { position: absolute; left: 12px; top: 50%; transform: translateY(-50%); font-size: 15px; line-height: 1; opacity: .7; }
  #outdir {
    width: 100%; padding: 10px 12px 10px 34px; border-radius: 10px;
    border: 1px solid var(--border); background: var(--bg); color: var(--text);
    font-family: Menlo, Consolas, monospace; font-size: 12.5px;
  }
  #outdir:focus { outline: none; border-color: var(--primary); }
  .field-row { display: flex; gap: 8px; align-items: stretch; }
  .field-row .field-wrap { flex: 1; min-width: 0; }
  .pick-btn { flex: none; padding: 0 14px; border-radius: 10px; border: 1px solid var(--border); background: transparent;
    color: var(--text); font-size: 13px; font-weight: 600; cursor: pointer; }
  .pick-btn:hover { background: var(--bg); }
  .pending-list { display: none; margin-top: 8px; max-height: 190px; overflow: auto; border: 1px solid var(--border);
    border-radius: 10px; }
  .p-row { display: flex; align-items: center; gap: 8px; padding: 6px 10px; font-size: 12.5px; }
  .p-row + .p-row { border-top: 1px solid var(--border); }
  .p-row .p-name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .p-row .p-size { color: var(--text-muted); font-size: 11.5px; flex: none; }
  .p-row .p-del { border: 0; background: none; color: var(--text-muted); cursor: pointer; padding: 2px 5px;
    border-radius: 6px; display: flex; }
  .p-row .p-del:hover { color: var(--red); background: var(--red-soft); }
  .p-row .p-del .ic { width: 13px; height: 13px; }
  .field-hint { font-size: 12px; color: var(--text-muted); margin-top: 6px; line-height: 1.45; }

  button.primary {
    width: 100%; margin-top: 20px; padding: 12px 16px; border: none; border-radius: 10px;
    background: var(--primary); color: #fff; font-size: 14.5px; font-weight: 600; cursor: pointer;
    transition: background .15s; display: flex; align-items: center; justify-content: center; gap: 8px;
  }
  button.primary:hover:not(:disabled) { background: var(--primary-hover); }
  button.primary:disabled { opacity: .45; cursor: not-allowed; }

  button.secondary {
    width: 100%; margin-top: 10px; padding: 10px 16px; border-radius: 10px;
    border: 1px solid var(--border); background: transparent; color: var(--text);
    font-size: 13.5px; font-weight: 600; cursor: pointer; transition: background .15s;
  }
  button.secondary:hover { background: var(--bg); }

  .spinner {
    width: 15px; height: 15px; border-radius: 50%;
    border: 2px solid rgba(255,255,255,.4); border-top-color: #fff;
    animation: spin .7s linear infinite; display: none;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  .progress-wrap { display: none; margin-top: 18px; }
  .status-line { font-size: 13px; color: var(--text-muted); margin-bottom: 8px; display: flex; justify-content: space-between; }
  .status-line .pct { font-variant-numeric: tabular-nums; font-weight: 600; color: var(--text); }
  .bar-track { height: 8px; border-radius: 99px; background: var(--bg); overflow: hidden; }
  .bar-fill { height: 100%; width: 0%; background: var(--primary); border-radius: 99px; transition: width .3s ease; }

  .stats { display: none; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-top: 18px; }
  .stat { border-radius: 10px; padding: 12px 10px; text-align: center; }
  .stat .n { font-size: 20px; font-weight: 700; line-height: 1.2; }
  .stat .l { font-size: 11px; color: var(--text-muted); margin-top: 2px; }
  .stat.found { background: var(--primary-soft); }
  .stat.found .n { color: var(--primary); }
  .stat.ok { background: var(--green-soft); }
  .stat.ok .n { color: var(--green); }
  .stat.skip { background: var(--amber-soft); }
  .stat.skip .n { color: var(--amber); }

  details.log-details { margin-top: 4px; }
  details.log-details summary {
    cursor: pointer; font-size: 13px; font-weight: 600; color: var(--text-muted);
    padding: 8px 0; list-style: none; display: flex; align-items: center; gap: 6px;
  }
  details.log-details summary::-webkit-details-marker { display: none; }
  details.log-details summary .chev { transition: transform .15s; font-size: 11px; }
  details.log-details[open] summary .chev { transform: rotate(90deg); }
  .log-box {
    background: #14161d; color: #d7d9e0; padding: 14px; border-radius: 10px;
    height: 280px; overflow-y: auto; white-space: pre-wrap; word-break: break-word;
    font-family: Menlo, Consolas, monospace; font-size: 11.5px; line-height: 1.55; margin-top: 4px;
  }
  .log-box .l-ok { color: #6ee7b7; }
  .log-box .l-skip { color: #fcd34d; }
  .log-box .l-err { color: #fca5a5; }
  .log-box .l-head { color: #93c5fd; font-weight: 600; }
  .log-box::-webkit-scrollbar { width: 8px; }
  .log-box::-webkit-scrollbar-thumb { background: #2f3340; border-radius: 8px; }

  details.audit-details { margin-top: 8px; display: none; }
  details.audit-details summary { color: var(--amber); }
  .audit-list {
    max-height: 320px; overflow-y: auto; margin-top: 4px; border: 1px solid var(--border, #e3e5ec);
    border-radius: 10px; background: var(--amber-soft);
  }
  .audit-item { padding: 9px 12px; font-size: 12.5px; line-height: 1.45; border-bottom: 1px solid rgba(0,0,0,.06); }
  .audit-item .a-head { display: flex; gap: 10px; align-items: flex-start; cursor: pointer; }
  .audit-item .a-fix { display: flex; flex-wrap: wrap; gap: 6px; margin: 8px 0 0 26px; }
  .audit-item .a-fix input[type=text] { flex: 1; min-width: 180px; padding: 6px 9px; border-radius: 8px;
    border: 1px solid var(--border); background: var(--card); color: var(--text); font-size: 12.5px; }
  .audit-item .a-fix select { padding: 6px 8px; border-radius: 8px; border: 1px solid var(--border);
    background: var(--card); color: var(--text); font-size: 12.5px; }
  .audit-item .a-fix input:focus, .audit-item .a-fix select:focus { outline: none; border-color: var(--primary); }
  .audit-item .a-fix button { padding: 6px 12px; border: 0; border-radius: 8px; background: var(--primary);
    color: #fff; font-weight: 600; font-size: 12.5px; cursor: pointer; }
  .audit-item .a-fix button:disabled { opacity: .5; cursor: default; }
  .audit-item .a-msg { margin: 6px 0 0 26px; font-size: 12px; color: var(--text-muted); }
  .audit-item .a-msg:empty { display: none; }
  .audit-item .a-msg.err { color: var(--red); }
  .audit-item .a-msg.ok { color: var(--green); display: flex; align-items: center; gap: 5px; }
  .audit-item:last-child { border-bottom: 0; }
  .audit-item .a-head input { margin-top: 3px; flex: none; cursor: pointer; }
  .audit-item .a-kind {
    display: inline-block; font-size: 10.5px; font-weight: 700; text-transform: uppercase; letter-spacing: .03em;
    color: var(--amber); margin-right: 6px;
  }
  .audit-item .a-where { font-weight: 600; margin-right: 6px; }
  .audit-item .a-pr { font-weight: 600; color: var(--primary); margin-right: 6px; }
  .audit-item .a-text { color: var(--text-muted); word-break: break-word; }
  .audit-item.fixed { opacity: .55; }
  .audit-item.fixed .a-where, .audit-item.fixed .a-text { text-decoration: line-through; }

  .file-summary { font-size: 13px; color: var(--text-muted); margin-top: 8px; }
  .file-summary b { color: var(--text); }

  .foot { text-align: center; font-size: 11.5px; color: var(--text-muted); margin-top: 20px; }

  .results-layout { display: none; gap: 20px; align-items: flex-start; margin-top: 20px; }
  .results-list {
    flex: 1; min-width: 0; display: flex; flex-direction: column; gap: 10px;
    max-height: 70vh; overflow-y: auto; padding: 2px 6px 2px 2px;
  }
  .results-list::-webkit-scrollbar { width: 8px; }
  .results-list::-webkit-scrollbar-thumb { background: var(--border); border-radius: 8px; }
  .result-item {
    background: var(--card); border: 1px solid var(--border); border-radius: 12px;
    padding: 14px 16px; position: relative;
  }
  .result-item .r-text {
    font-size: 13px; line-height: 1.6; white-space: pre-wrap; word-break: break-word;
    margin-bottom: 12px;
  }
  .result-item .r-file {
    font-size: 11px; color: var(--text-muted); margin-bottom: 6px;
    display: flex; align-items: center; gap: 8px;
  }
  .r-index {
    display: inline-flex; align-items: center; justify-content: center;
    min-width: 20px; height: 20px; padding: 0 6px; border-radius: 6px;
    background: var(--primary-soft); color: var(--primary);
    font-size: 11px; font-weight: 700; flex-shrink: 0;
  }
  .results-col { flex: 1; min-width: 0; display: flex; flex-direction: column; }
  .new-info { font-size: 13px; color: var(--text-muted); margin-bottom: 10px; line-height: 1.5; }
  .new-info:empty { display: none; }
  .result-item.is-new { border-color: var(--amber); background: var(--amber-soft); box-shadow: inset 4px 0 0 var(--amber); }
  .new-badge { display: inline-block; font-size: 10.5px; font-weight: 700; letter-spacing: .05em; text-transform: uppercase;
    color: var(--amber); border: 1px solid var(--amber); border-radius: 6px; padding: 1px 6px; margin-left: 8px; }
  .agenda-cards { margin-top: 14px; border: 1px solid var(--border); border-radius: 12px; background: var(--card); }
  .agenda-cards summary { cursor: pointer; padding: 12px 16px; font-weight: 600; font-size: 13px; }
  #agendaList { max-height: 60vh; overflow: auto; padding: 0 12px 12px; display: flex; flex-direction: column; gap: 6px; }
  .result-item.is-agenda { padding: 8px 12px; background: var(--bg); border-radius: 9px; font-size: 12.5px; line-height: 1.5; }
  .result-item.is-agenda .a-sec { font-weight: 700; color: var(--primary); margin-right: 8px; white-space: nowrap; }
  .arr-item.is-new { border-color: var(--amber); background: var(--amber-soft); }
  .arr-item.locked { cursor: default; background: transparent; }
  .arr-item.locked .arr-num { color: var(--text-muted); }
  .arr-section h4.arr-head { cursor: pointer; display: flex; align-items: center; gap: 6px; user-select: none; }
  .arr-section h4.arr-head .ic { width: 14px; height: 14px; transition: transform .12s; }
  .arr-section.open h4.arr-head .ic { transform: rotate(90deg); }
  .arr-count { color: var(--text-muted); font-weight: 500; font-size: 12px; }
  .arr-section:not(.open) .arr-list { min-height: 8px; }
  .arrange { display: none; margin-top: 20px; background: var(--card); border: 1px solid var(--border);
    border-radius: var(--radius); box-shadow: var(--shadow); padding: 20px 22px; }
  .arrange h3 { margin: 0 0 4px; font-size: 16px; }
  .arrange .arr-sub { color: var(--text-muted); font-size: 13px; line-height: 1.5; margin-bottom: 6px; }
  .arr-note { color: var(--text-muted); font-size: 13.5px; line-height: 1.55; padding: 6px 0; }
  .arr-section h4 { font-size: 14px; font-weight: 700; margin: 18px 0 8px; }
  .arr-list { min-height: 34px; border: 1px dashed transparent; border-radius: 10px; padding: 3px; transition: background .12s, border-color .12s; }
  .arr-list.over { border-color: var(--primary); background: var(--primary-soft); }
  .arr-empty { color: var(--text-muted); font-size: 12.5px; padding: 8px 10px; }
  .arr-item { display: flex; gap: 10px; align-items: flex-start; padding: 9px 10px; margin-bottom: 6px;
    border: 1px solid var(--border); border-radius: 10px; background: var(--bg); cursor: grab; }
  .arr-item:last-child { margin-bottom: 0; }
  .arr-item:active { cursor: grabbing; }
  .arr-item.dragging { opacity: .4; }
  .arr-item.drop-before { box-shadow: 0 -3px 0 0 var(--primary); }
  .arr-item.drop-after { box-shadow: 0 3px 0 0 var(--primary); }
  .arr-item.moved { animation: movedflash 1.4s ease-out; }
  @keyframes movedflash { from { box-shadow: 0 0 0 3px var(--primary); } to { box-shadow: 0 0 0 0 transparent; } }
  .arr-grip { color: var(--text-muted); flex: none; margin-top: 2px; display: flex; }
  .arr-grip .ic { width: 16px; height: 16px; }
  .arr-num { font-weight: 700; color: var(--primary); flex: none; min-width: 24px; font-size: 13px; padding-top: 1px; }
  .arr-body { flex: 1; min-width: 0; }
  .arr-title { font-size: 13px; line-height: 1.5; overflow-wrap: anywhere; }
  .arr-file { color: var(--text-muted); font-size: 11px; margin-top: 2px; }
  .arr-btns { display: flex; flex-direction: column; gap: 3px; flex: none; }
  .arr-btns button { border: 1px solid var(--border); background: transparent; color: var(--text-muted);
    border-radius: 6px; padding: 2px 5px; cursor: pointer; display: flex; }
  .arr-btns button:hover:not(:disabled) { border-color: var(--primary); color: var(--primary); }
  .arr-btns button:disabled { opacity: .3; cursor: default; }
  .arr-btns .ic { width: 13px; height: 13px; }
  .r-sections { display: flex; flex-wrap: wrap; gap: 6px; }
  .section-btn {
    border: 1px solid var(--border); background: transparent; color: var(--text-muted);
    font-size: 12px; font-weight: 600; padding: 6px 10px; border-radius: 8px;
    cursor: pointer; transition: background .15s, color .15s, border-color .15s;
  }
  .section-btn:hover { border-color: var(--primary); color: var(--primary); }
  .section-btn .ic { width: 12px; height: 12px; margin-left: 4px; vertical-align: -2px; }
  .sec-menu { position: fixed; z-index: 950; min-width: 240px; max-width: min(420px, calc(100vw - 16px));
    max-height: 60vh; overflow: auto; background: var(--card); border: 1px solid var(--border); border-radius: 10px;
    box-shadow: var(--shadow); padding: 4px; }
  .sec-menu[hidden] { display: none; }
  .sec-menu-item { display: block; width: 100%; text-align: left; border: 0; background: none; color: var(--text);
    font-size: 12.5px; line-height: 1.4; padding: 7px 12px; border-radius: 7px; cursor: pointer; white-space: normal; }
  .sec-menu-item:hover { background: var(--primary-soft); }
  .sec-menu-item.current { color: var(--primary); font-weight: 700; }
  .section-btn.active { background: var(--primary); border-color: var(--primary); color: #fff; }

  .reminder-panel {
    width: 260px; flex-shrink: 0; position: sticky; top: 20px;
    background: var(--primary-soft); border: 1px solid var(--border); border-radius: 12px;
    padding: 16px 18px;
  }
  .reminder-panel h3 {
    font-size: 11px; text-transform: uppercase; letter-spacing: .04em;
    color: var(--text-muted); margin: 0 0 12px; font-weight: 700;
  }
  .reminder-panel ol { margin: 0; padding-left: 0; list-style: none; }
  .reminder-panel li {
    font-size: 12.5px; line-height: 1.5; color: var(--text); margin-bottom: 10px;
    display: flex; align-items: flex-start; gap: 6px;
  }
  .reminder-panel li:last-child { margin-bottom: 0; }
  .reminder-panel li .s-num { flex-shrink: 0; font-weight: 700; color: var(--text-muted); }
  .reminder-panel li .s-text { flex: 1; min-width: 0; overflow-wrap: anywhere; }
  .reminder-panel li .s-edit-btn {
    flex-shrink: 0; border: none; background: none; cursor: pointer;
    font-size: 14px; color: var(--text-muted); opacity: .6; padding: 0 2px; line-height: 1.6; transition: opacity .15s;
  }
  .reminder-panel li .s-edit-btn:hover { opacity: 1; }
  .reminder-panel li .s-edit-btn .ic { width: 14px; height: 14px; }
  .reminder-panel li .s-del:hover { color: var(--red); }
  .reminder-panel li .s-add:hover { color: var(--primary); }
  .reminder-panel li .s-input {
    flex: 1; font: inherit; padding: 2px 6px; border: 1px solid var(--primary);
    border-radius: 5px; background: var(--card); color: var(--text);
  }

  @media (max-width: 900px) {
    body { flex-direction: column; }
    .sidebar { width: auto; height: auto; max-height: 40vh; position: static; border-right: 0;
      border-bottom: 1px solid var(--border); }
  }
  @media (max-width: 760px) {
    .results-layout { flex-direction: column; }
    .reminder-panel { width: 100%; position: static; }
  }

  #generateBtn { display: none; }

  .toast-stack {
    position: fixed; left: 50%; bottom: 24px; transform: translateX(-50%);
    display: flex; flex-direction: column; gap: 8px; z-index: 1000;
    width: 100%; max-width: 460px; padding: 0 16px; align-items: center;
  }
  .toast {
    width: 100%; background: var(--card); border: 1px solid var(--border);
    border-left: 4px solid var(--amber); border-radius: 10px; box-shadow: var(--shadow);
    padding: 12px 14px; font-size: 13px; line-height: 1.5; color: var(--text);
    opacity: 0; transform: translateY(8px); transition: opacity .2s, transform .2s;
  }
  .toast.show { opacity: 1; transform: translateY(0); }
  .toast.toast-ok { border-left-color: var(--green); }

  .tabs {
    display: flex; gap: 4px; background: var(--bg); border: 1px solid var(--border);
    border-radius: 12px; padding: 4px; margin-bottom: 20px;
  }
  .tab {
    flex: 1; border: none; background: transparent; color: var(--text-muted);
    font-size: 13.5px; font-weight: 600; padding: 9px 10px; border-radius: 9px;
    cursor: pointer; transition: background .15s, color .15s;
  }
  .tab:hover { color: var(--text); }
  .tab.active { background: var(--card); color: var(--primary); box-shadow: var(--shadow); }

  .panel-desc { text-align: center; font-size: 14px; color: var(--text-muted); margin: -8px 0 20px; }
</style>
</head>
<body>
{{ icon_sprite|safe }}
<aside class="sidebar">
  <div class="sb-title">Структура</div>
  <div class="tree" id="tree"></div>
</aside>
<main class="main">
<div class="shell">

  <div class="head">
    <div class="icon"><svg class="ic"><use href="#i-landmark"/></svg></div>
    <h1>Порядок денний</h1>
  </div>

  <div class="tabs" role="tablist">
    <button class="tab active" id="tabBtn-plenary" role="tab" aria-selected="true">Пленарні PDF</button>
    <button class="tab" id="tabBtn-generate" role="tab" aria-selected="false">Генерація ПД</button>
  </div>

  <div class="panel" id="panel-plenary">
  <p class="panel-desc">Завантажте PDF з порядком денним — скрипт сам знайде й скачає всі відповідні рішення "На пленарне засідання"</p>
  <div class="card">
    <div class="dropzone" id="drop">
      <div class="drop-icon"><svg class="ic"><use href="#i-upload"/></svg></div>
      <div class="drop-main">Перетягніть сюди PDF порядку денного</div>
      <div class="drop-sub">або натисніть, щоб обрати файл</div>
    </div>
    <input type="file" id="file" accept="application/pdf">

    <div class="file-chip" id="fileChip">
      <div class="fi"><svg class="ic"><use href="#i-paperclip"/></svg></div>
      <div style="min-width:0;">
        <div class="fname" id="fname"></div>
        <div class="fsize" id="fsize"></div>
      </div>
      <button class="fclear" id="fclear" title="Прибрати файл"><svg class="ic"><use href="#i-x"/></svg></button>
    </div>

    <label class="field-label" for="outdir">Робоча папка</label>
    <div class="field-row">
      <div class="field-wrap">
        <span class="fi-prefix"><svg class="ic"><use href="#i-folder"/></svg></span>
        <input type="text" id="outdir" value="{{ default_out_dir }}">
      </div>
      <button type="button" class="pick-btn" id="pickDirBtn" hidden>Обрати…</button>
    </div>
    <div class="field-hint">Сюди скачуються файли, і звідси дерево зліва їх читає. Можна вказати власну папку або обрати ту, де вже лежать скачані файли.</div>

    <button class="primary" id="startBtn" disabled>
      <span class="spinner" id="spinner"></span>
      <span id="startLabel">Завантажити</span>
    </button>

    <div class="progress-wrap" id="progressWrap">
      <div class="status-line" id="statusLine"><span id="statusText">Починаємо…</span><span class="pct" id="statusPct"></span></div>
      <div class="bar-track"><div class="bar-fill" id="barFill"></div></div>

      <div class="stats" id="stats">
        <div class="stat found"><div class="n" id="statFound">–</div><div class="l">знайдено</div></div>
        <div class="stat ok"><div class="n" id="statOk">–</div><div class="l">скачано</div></div>
        <div class="stat skip"><div class="n" id="statSkip">–</div><div class="l">пропущено</div></div>
      </div>

      <button class="secondary" id="openBtn" style="display:none">Показати папку</button>

      <details class="log-details audit-details" id="auditDetails" open>
        <summary><span class="chev"><svg class="ic"><use href="#i-chevron-right"/></svg></span> <span id="auditTitle">Зауваження до PDF</span></summary>
        <div class="audit-list" id="auditList"></div>
      </details>

      <details class="log-details" id="logDetails">
        <summary><span class="chev"><svg class="ic"><use href="#i-chevron-right"/></svg></span> Детальний журнал</summary>
        <pre class="log-box" id="log"></pre>
      </details>
    </div>
  </div>
  </div>

  <div class="panel" id="panel-generate" hidden>
  <p class="panel-desc">Перетягніть сюди проєкти рішень (PDF, можна одразу кілька сотень) — для кожного файлу скрипт визначить за назвою номер і рік рішення (напр. 183-25.pdf — №183 за 2025 рік) і знайде його на сайті КМР</p>
  <div class="card">
    <div class="dropzone" id="dropGen">
      <div class="drop-icon"><svg class="ic"><use href="#i-upload"/></svg></div>
      <div class="drop-main">Перетягніть сюди файли проєктів рішень</div>
      <div class="drop-sub">або натисніть, щоб обрати файли (напр. "524-26.pdf"). Можна одразу кілька, а можна по одному — список накопичується. Разом із ними можна додати й порядок денний (Word .docx або PDF): нові пункти вставляться в нього</div>
    </div>
    <input type="file" id="fileGen" accept="application/pdf,.docx" multiple>

    <div class="file-chip" id="fileChipGen">
      <div class="fi"><svg class="ic"><use href="#i-paperclip"/></svg></div>
      <div style="min-width:0;">
        <div class="fname" id="fnameGen"></div>
        <div class="fsize" id="fsizeGen"></div>
      </div>
      <button class="fclear" id="fclearGen" title="Очистити список"><svg class="ic"><use href="#i-x"/></svg></button>
    </div>
    <div class="pending-list" id="pendingListGen"></div>

    <button class="primary" id="startBtnGen" disabled>
      <span class="spinner" id="spinnerGen"></span>
      <span id="startLabelGen">Почати</span>
    </button>

    <div class="progress-wrap" id="progressWrapGen">
      <div class="status-line" id="statusLineGen"><span id="statusTextGen">Починаємо…</span><span class="pct" id="statusPctGen"></span></div>
      <div class="bar-track"><div class="bar-fill" id="barFillGen"></div></div>

      <div class="stats" id="statsGen">
        <div class="stat found"><div class="n" id="statFoundGen">–</div><div class="l">файлів</div></div>
        <div class="stat ok"><div class="n" id="statOkGen">–</div><div class="l">оброблено</div></div>
        <div class="stat skip"><div class="n" id="statSkipGen">–</div><div class="l">помилок</div></div>
      </div>

      <details class="log-details audit-details" id="auditDetailsGen" open>
        <summary><span class="chev"><svg class="ic"><use href="#i-chevron-right"/></svg></span> <span id="auditTitleGen">Назви файлів</span></summary>
        <div class="audit-list" id="auditListGen"></div>
      </details>

      <button class="secondary" id="openBtnGen" style="display:none">Показати папку</button>

      <details class="log-details" id="logDetailsGen">
        <summary><span class="chev"><svg class="ic"><use href="#i-chevron-right"/></svg></span> Детальний журнал</summary>
        <pre class="log-box" id="logGen"></pre>
      </details>
    </div>
  </div>

  <div class="results-layout" id="resultsLayout">
    <div class="results-col">
      <div class="new-info" id="newInfo"></div>
      <div class="results-list" id="resultsList"></div>
      <details class="agenda-cards" id="agendaCards" style="display:none">
        <summary>Пункти з порядку денного (<span id="agendaCount">0</span>) — розділи вже обрано, вони стоять на своїх місцях</summary>
        <div id="agendaList"></div>
      </details>
    </div>
    <div class="reminder-panel">
      <h3>Розділи порядку денного</h3>
      <ol id="sectionsList"></ol>
    </div>
  </div>

  <div class="arrange" id="arrangePanel">
    <h3>Порядок у документі</h3>
    <div class="arr-sub">Так пункти стоятимуть у документі. Перетягуйте їх мишкою (навіть в інший розділ) або користуйтеся стрілками справа.</div>
    <div id="arrangeBody"></div>
  </div>

  <button class="primary" id="generateBtn">Згенерувати ПД</button>
  <button class="secondary" id="newBatchBtn" style="display:none">Очистити список і почати заново</button>
  </div>

  <div class="foot">Працює локально · нічого не надсилається за межі цього комп'ютера</div>
</div>
</main>

<div class="toast-stack" id="toastStack"></div>

<script>
const shellEl = document.querySelector('.shell');
const tabs = {
  plenary: { btn: document.getElementById('tabBtn-plenary'), panel: document.getElementById('panel-plenary') },
  generate: { btn: document.getElementById('tabBtn-generate'), panel: document.getElementById('panel-generate') },
};
function selectTab(name) {
  for (const key in tabs) {
    const active = key === name;
    tabs[key].btn.classList.toggle('active', active);
    tabs[key].btn.setAttribute('aria-selected', active ? 'true' : 'false');
    tabs[key].panel.hidden = !active;
  }
  shellEl.classList.toggle('wide', name === 'generate');
}
tabs.plenary.btn.addEventListener('click', () => selectTab('plenary'));
tabs.generate.btn.addEventListener('click', () => selectTab('generate'));

const treeEl = document.getElementById('tree');
const treeOpen = new Set();  // розкриті папки (відносні шляхи; '' — корінь); за замовчуванням усе згорнуто
let treeData = { name: null };
let treeSig = '';

async function treeApi(url, path) {
  try {
    const resp = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ path }),
    });
    if (!resp.ok) {
      const data = await resp.json().catch(() => ({}));
      showToast(data.error || 'Не вдалося виконати дію.');
      refreshTree();  // дерево могло відстати від диска — оновлюємо одразу
    }
  } catch (e) {
    showToast('Немає зв’язку із застосунком.');
  }
}

// Іконки Lucide з прихованого SVG-спрайта на початку сторінки (див. icons.py).
function icon(name) {
  const ns = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(ns, 'svg');
  svg.setAttribute('class', 'ic');
  const use = document.createElementNS(ns, 'use');
  use.setAttribute('href', '#i-' + name);
  svg.appendChild(use);
  return svg;
}

// ---- Надсилання на пошту: підменю з контактами ----
const mp = document.createElement('div');
mp.className = 'mp';
mp.hidden = true;
document.body.appendChild(mp);
const mpState = { node: null, rect: null, view: 'list', contacts: [], selected: new Set(),
  addOpen: false, editId: null, settings: {}, status: '', statusErr: false,
  groups: [], search: '', groupEdit: null };

async function mpApi(method, url, body) {
  const resp = await fetch(url, {
    method,
    headers: { 'Content-Type': 'application/json' },
    body: body ? JSON.stringify(body) : undefined,
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) throw new Error(data.error || ('Помилка ' + resp.status));
  return data;
}

function mpEl(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
}

function mpClose() { mp.hidden = true; }

// Власне вікно підтвердження (системний confirm() у вікні застосунку може не працювати).
function mpConfirm(title, text, okLabel) {
  return new Promise(resolve => {
    const overlay = mpEl('div', 'modal-ov');
    const box = mpEl('div', 'modal');
    box.setAttribute('role', 'dialog');
    box.append(mpEl('h3', '', title), mpEl('p', '', text));
    const actions = mpEl('div', 'mp-actions');
    const cancel = mpEl('button', 'mp-btn ghost', 'Скасувати');
    const ok = mpEl('button', 'mp-btn danger', okLabel);
    actions.append(cancel, ok);
    box.appendChild(actions);
    overlay.appendChild(box);
    document.body.appendChild(overlay);
    const done = value => {
      document.removeEventListener('keydown', onKey, true);
      overlay.remove();
      resolve(value);
    };
    const onKey = e => {
      if (e.key === 'Escape') { e.stopPropagation(); done(false); }
    };
    document.addEventListener('keydown', onKey, true);
    cancel.addEventListener('click', () => done(false));
    ok.addEventListener('click', () => done(true));
    overlay.addEventListener('mousedown', e => { if (e.target === overlay) done(false); });
    cancel.focus();
  });
}

function mpPosition() {
  const sb = document.querySelector('.sidebar').getBoundingClientRect();
  const wide = window.innerWidth > 900;
  const h = mp.offsetHeight, w = mp.offsetWidth;
  const left = wide ? sb.right + 8 : Math.max(8, Math.min(mpState.rect.left, window.innerWidth - w - 8));
  const top = wide ? mpState.rect.top : mpState.rect.bottom + 6;
  mp.style.left = left + 'px';
  mp.style.top = Math.max(8, Math.min(top, window.innerHeight - h - 8)) + 'px';
}

async function mpOpen(anchor, node) {
  if (!mp.hidden && mpState.node && mpState.node.path === node.path) { mpClose(); return; }
  Object.assign(mpState, { node, rect: anchor.getBoundingClientRect(), view: 'list',
    selected: new Set(), addOpen: false, editId: null, search: '', groupEdit: null, status: '', statusErr: false });
  mp.hidden = false;
  try {
    const [c, cfg, gr] = await Promise.all([
      mpApi('GET', '/api/contacts'), mpApi('GET', '/api/mail/settings'), mpApi('GET', '/api/groups')]);
    mpState.contacts = c.contacts;
    mpState.settings = cfg;
    mpState.groups = gr.groups;
    if (!cfg.configured) mpState.status = 'Пошту ще не налаштовано — відкрийте «Налаштування пошти».';
  } catch (e) {
    mpState.status = e.message;
    mpState.statusErr = true;
  }
  mpRender();
}

function mpRender() {
  mp.textContent = '';
  if (mpState.view === 'settings') mpRenderSettings();
  else if (mpState.view === 'group') mpRenderGroup();
  else mpRenderList();
  mpPosition();
}

function mpInput(type, placeholder, value) {
  const i = document.createElement('input');
  i.type = type;
  i.placeholder = placeholder;
  i.value = value || '';
  return i;
}

function mpMatches(c, query) {
  const q = query.trim().toLowerCase();
  return !q || (c.name + ' ' + c.email).toLowerCase().includes(q);
}

function mpOpenGroup(group) {
  // Нова група стартує з поточного вибору — так 40 адрес відмічають один раз.
  mpState.groupEdit = { id: group ? group.id : null, name: group ? group.name : '',
    members: new Set(group ? group.contact_ids : [...mpState.selected]) };
  mpState.view = 'group';
  mpState.status = '';
  mpState.statusErr = false;
  mpRender();
}

function mpRenderList() {
  const node = mpState.node;
  mp.appendChild(mpEl('div', 'mp-title', 'Надіслати на пошту'));
  mp.appendChild(mpEl('div', 'mp-sub', node.name));

  const status = mpEl('div', 'mp-status' + (mpState.statusErr ? ' err' : ''), mpState.status);
  const setStatus = (msg, err) => {
    mpState.status = msg;
    mpState.statusErr = !!err;
    status.textContent = msg;
    status.className = 'mp-status' + (err ? ' err' : '');
  };
  const sendBtn = mpEl('button', 'mp-btn');
  const list = mpEl('div', 'mp-list');
  const countEl = mpEl('span', 'mp-hint');
  const allLink = mpEl('button', 'mp-link');
  const chips = [];

  const visible = () => mpState.contacts.filter(c => mpMatches(c, mpState.search));
  const memberIds = g => g.contact_ids.filter(id => mpState.contacts.some(c => c.id === id));

  // Оновлює все, що залежить від вибору: кнопку, лічильник, підсвітку груп.
  const refresh = () => {
    sendBtn.hidden = mpState.selected.size === 0;
    sendBtn.textContent = 'Надіслати (' + mpState.selected.size + ')';
    countEl.textContent = 'Обрано: ' + mpState.selected.size + ' з ' + mpState.contacts.length;
    const rows = visible();
    allLink.textContent = rows.length && rows.every(c => mpState.selected.has(c.id)) ? 'Зняти всіх' : 'Вибрати всіх';
    chips.forEach(({ g, chip }) => {
      const ids = memberIds(g);
      chip.classList.toggle('on', ids.length > 0 && ids.every(id => mpState.selected.has(id)));
    });
  };

  function contactRow(c) {
    const row = mpEl('label', 'mp-row');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = mpState.selected.has(c.id);
    cb.addEventListener('change', () => {
      if (cb.checked) mpState.selected.add(c.id); else mpState.selected.delete(c.id);
      refresh();
    });
    const who = mpEl('div', 'who');
    who.append(mpEl('b', '', c.name), mpEl('span', '', c.email));
    const edit = mpEl('button', 'mp-edit');
    edit.appendChild(icon('pencil'));
    edit.title = 'Змінити ім’я або пошту';
    edit.addEventListener('click', e => {
      e.preventDefault();
      e.stopPropagation();
      mpState.editId = c.id;
      mpState.addOpen = true;
      mpState.status = '';
      mpRender();
    });
    const del = mpEl('button', 'mp-del');
    del.appendChild(icon('x'));
    del.title = 'Видалити контакт';
    del.addEventListener('click', async e => {
      e.preventDefault();
      e.stopPropagation();
      const sure = await mpConfirm('Видалити контакт?',
        'Контакт «' + c.name + '» (' + c.email + ') буде видалено зі списку та з усіх груп. Цю дію не можна скасувати.',
        'Видалити');
      if (!sure) return;
      try {
        await mpApi('DELETE', '/api/contacts/' + c.id);
        mpState.contacts = mpState.contacts.filter(x => x.id !== c.id);
        mpState.groups.forEach(g => { g.contact_ids = g.contact_ids.filter(id => id !== c.id); });
        mpState.selected.delete(c.id);
        mpRender();
      } catch (err) { setStatus(err.message, true); }
    });
    row.append(cb, who, edit, del);
    return row;
  }

  function renderRows() {
    const keep = list.scrollTop;
    list.textContent = '';
    const rows = visible();
    if (!rows.length) list.appendChild(mpEl('div', 'mp-hint', 'Нічого не знайдено.'));
    rows.forEach(c => list.appendChild(contactRow(c)));
    list.scrollTop = keep;
  }

  if (!mpState.contacts.length) {
    mp.appendChild(mpEl('div', 'mp-hint', 'Контактів ще немає — додайте першого нижче.'));
  } else {
    // Групи: клік — обрати всіх із групи (повторний клік — зняти).
    mp.appendChild(mpEl('div', 'mp-label', 'Групи'));
    const chipsBox = mpEl('div', 'mp-chips');
    mpState.groups.forEach(g => {
      const chip = mpEl('span', 'mp-chip');
      const main = mpEl('button', 'chip-main', g.name + ' · ' + memberIds(g).length);
      main.title = 'Обрати всіх із групи (ще раз — зняти)';
      main.addEventListener('click', () => {
        const ids = memberIds(g);
        const allOn = ids.length > 0 && ids.every(id => mpState.selected.has(id));
        ids.forEach(id => { if (allOn) mpState.selected.delete(id); else mpState.selected.add(id); });
        renderRows();
        refresh();
      });
      const edit = mpEl('button', 'chip-edit');
      edit.appendChild(icon('pencil'));
      edit.title = 'Змінити групу';
      edit.addEventListener('click', () => mpOpenGroup(g));
      chip.append(main, edit);
      chips.push({ g, chip });
      chipsBox.appendChild(chip);
    });
    const newGroup = mpEl('button', 'mp-link');
    newGroup.append(icon('plus'), 'Нова група');
    newGroup.title = 'Створити групу з обраних контактів';
    newGroup.addEventListener('click', () => mpOpenGroup(null));
    chipsBox.appendChild(newGroup);
    mp.appendChild(chipsBox);

    const search = mpInput('text', 'Пошук за ім’ям або поштою', mpState.search);
    search.style.marginTop = '0';
    search.addEventListener('input', () => { mpState.search = search.value; renderRows(); refresh(); });
    mp.appendChild(search);

    allLink.addEventListener('click', () => {
      const rows = visible();
      const allOn = rows.every(c => mpState.selected.has(c.id));
      rows.forEach(c => { if (allOn) mpState.selected.delete(c.id); else mpState.selected.add(c.id); });
      renderRows();
      refresh();
    });
    const info = mpEl('div', 'mp-actions mp-info');
    info.append(countEl, allLink);
    mp.appendChild(info);
    mp.appendChild(list);
    renderRows();
  }

  if (mpState.addOpen) {
    const editing = mpState.contacts.find(c => c.id === mpState.editId);
    const name = mpInput('text', 'Ім’я та прізвище', editing ? editing.name : '');
    const email = mpInput('email', 'email@example.com', editing ? editing.email : '');
    const add = mpEl('button', 'mp-btn', editing ? 'Зберегти' : 'Додати');
    const cancel = mpEl('button', 'mp-btn ghost', 'Скасувати');
    const submit = async () => {
      const body = { name: name.value, email: email.value };
      try {
        if (editing) {
          const res = await mpApi('PUT', '/api/contacts/' + editing.id, body);
          Object.assign(editing, res.contact);
        } else {
          const res = await mpApi('POST', '/api/contacts', body);
          mpState.contacts.push(res.contact);
          mpState.selected.add(res.contact.id);
        }
        mpState.addOpen = false;
        mpState.editId = null;
        mpState.status = '';
        mpRender();
      } catch (err) { setStatus(err.message, true); }
    };
    add.addEventListener('click', submit);
    cancel.addEventListener('click', () => { mpState.addOpen = false; mpState.editId = null; mpRender(); });
    [name, email].forEach(i => i.addEventListener('keydown', e => { if (e.key === 'Enter') submit(); }));
    const btns = mpEl('div', 'mp-btns');
    btns.append(add, cancel);
    mp.append(name, email, btns);
    setTimeout(() => name.focus(), 0);
  } else {
    const links = mpEl('div', 'mp-links');
    const addLink = mpEl('button', 'mp-link');
    addLink.append(icon('plus'), 'Додати контакт');
    addLink.addEventListener('click', () => { mpState.addOpen = true; mpState.editId = null; mpRender(); });

    // Імпорт контактів з Excel (.xlsx) або CSV: колонки з іменем та поштою.
    const importLink = mpEl('button', 'mp-link');
    importLink.append(icon('file-spreadsheet'), 'Імпорт з Excel');
    importLink.title = 'Файл .xlsx або .csv: колонка з поштою та (бажано) з іменем; дані з першого аркуша';
    const picker = document.createElement('input');
    picker.type = 'file';
    picker.accept = '.xlsx,.xlsm,.csv';
    picker.hidden = true;
    importLink.addEventListener('click', () => picker.click());
    picker.addEventListener('change', async () => {
      if (!picker.files.length) return;
      const form = new FormData();
      form.append('file', picker.files[0]);
      setStatus('Імпортуємо контакти…', false);
      try {
        const resp = await fetch('/api/contacts/import', { method: 'POST', body: form });
        const data = await resp.json().catch(() => ({}));
        if (!resp.ok) throw new Error(data.error || ('Помилка ' + resp.status));
        mpState.contacts = data.contacts;
        mpState.status = 'Імпортовано: ' + data.added +
          (data.duplicates ? ' · уже були у списку: ' + data.duplicates : '') +
          (data.invalid ? ' · рядків без коректної пошти: ' + data.invalid : '');
        mpState.statusErr = false;
        mpRender();
      } catch (err) { setStatus(err.message, true); }
    });
    links.append(addLink, importLink, picker);
    mp.appendChild(links);
  }

  mp.appendChild(status);

  const actions = mpEl('div', 'mp-actions');
  const cfgLink = mpEl('button', 'mp-link');
  cfgLink.append(icon('settings'), 'Налаштування пошти');
  cfgLink.style.padding = '0';
  cfgLink.addEventListener('click', () => { mpState.view = 'settings'; mpState.status = ''; mpRender(); });
  sendBtn.addEventListener('click', async () => {
    sendBtn.disabled = true;
    sendBtn.textContent = 'Надсилаємо…';
    setStatus('', false);
    try {
      const res = await mpApi('POST', '/api/tree/send', { path: node.path, contact_ids: [...mpState.selected] });
      showToast('Надіслано отримувачам: ' + res.sent, 'ok');
      mpClose();
    } catch (err) {
      setStatus(err.message, true);
      sendBtn.disabled = false;
      refresh();
    }
  });
  actions.append(cfgLink, sendBtn);
  mp.appendChild(actions);
  refresh();
}

// Створення / редагування групи: назва + контакти, що до неї входять.
function mpRenderGroup() {
  const ge = mpState.groupEdit;
  mp.appendChild(mpEl('div', 'mp-title', ge.id ? 'Змінити групу' : 'Нова група'));
  const name = mpInput('text', 'Назва групи, напр. «для землі»', ge.name);
  name.style.marginBottom = '4px';
  name.addEventListener('input', () => { ge.name = name.value; });
  mp.appendChild(name);
  if (!ge.id && ge.members.size) {
    mp.appendChild(mpEl('div', 'mp-hint', 'Уже відмічено ваш поточний вибір — додайте чи приберіть контакти.'));
  }

  let query = '';
  const search = mpInput('text', 'Пошук за ім’ям або поштою', '');
  mp.appendChild(search);
  const countEl = mpEl('span', 'mp-hint');
  const allLink = mpEl('button', 'mp-link');
  const info = mpEl('div', 'mp-actions mp-info');
  info.append(countEl, allLink);
  mp.appendChild(info);
  const list = mpEl('div', 'mp-list');
  mp.appendChild(list);

  const visible = () => mpState.contacts.filter(c => mpMatches(c, query));
  const refresh = () => {
    countEl.textContent = 'У групі: ' + ge.members.size + ' з ' + mpState.contacts.length;
    const rows = visible();
    allLink.textContent = rows.length && rows.every(c => ge.members.has(c.id)) ? 'Зняти всіх' : 'Вибрати всіх';
  };
  const renderRows = () => {
    const keep = list.scrollTop;
    list.textContent = '';
    const rows = visible();
    if (!rows.length) list.appendChild(mpEl('div', 'mp-hint', 'Нічого не знайдено.'));
    rows.forEach(c => {
      const row = mpEl('label', 'mp-row');
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = ge.members.has(c.id);
      cb.addEventListener('change', () => {
        if (cb.checked) ge.members.add(c.id); else ge.members.delete(c.id);
        refresh();
      });
      const who = mpEl('div', 'who');
      who.append(mpEl('b', '', c.name), mpEl('span', '', c.email));
      row.append(cb, who);
      list.appendChild(row);
    });
    list.scrollTop = keep;
  };
  search.addEventListener('input', () => { query = search.value; renderRows(); refresh(); });
  allLink.addEventListener('click', () => {
    const rows = visible();
    const allOn = rows.every(c => ge.members.has(c.id));
    rows.forEach(c => { if (allOn) ge.members.delete(c.id); else ge.members.add(c.id); });
    renderRows();
    refresh();
  });
  renderRows();
  refresh();

  const status = mpEl('div', 'mp-status' + (mpState.statusErr ? ' err' : ''), mpState.status);
  mp.appendChild(status);
  const showError = msg => { status.textContent = msg; status.className = 'mp-status err'; };

  const btns = mpEl('div', 'mp-btns');
  const save = mpEl('button', 'mp-btn', 'Зберегти');
  const cancel = mpEl('button', 'mp-btn ghost', 'Скасувати');
  btns.append(save, cancel);
  mp.appendChild(btns);
  cancel.addEventListener('click', () => { mpState.view = 'list'; mpState.status = ''; mpRender(); });
  save.addEventListener('click', async () => {
    const body = { name: ge.name, contact_ids: [...ge.members] };
    try {
      const res = ge.id ? await mpApi('PUT', '/api/groups/' + ge.id, body) : await mpApi('POST', '/api/groups', body);
      const at = mpState.groups.findIndex(g => g.id === res.group.id);
      if (at >= 0) mpState.groups[at] = res.group; else mpState.groups.push(res.group);
      mpState.view = 'list';
      mpState.status = 'Групу «' + res.group.name + '» збережено (' + res.group.contact_ids.length + ').';
      mpState.statusErr = false;
      mpRender();
    } catch (err) { showError(err.message); }
  });

  if (ge.id) {
    const del = mpEl('button', 'mp-link', 'Видалити групу');
    del.style.color = 'var(--red)';
    del.addEventListener('click', async () => {
      const sure = await mpConfirm('Видалити групу?',
        'Групу «' + ge.name + '» буде видалено. Самі контакти залишаться у списку.', 'Видалити');
      if (!sure) return;
      try {
        await mpApi('DELETE', '/api/groups/' + ge.id);
        mpState.groups = mpState.groups.filter(g => g.id !== ge.id);
        mpState.view = 'list';
        mpState.status = '';
        mpRender();
      } catch (err) { showError(err.message); }
    });
    mp.appendChild(del);
  }
}

function mpRenderSettings() {
  const cfg = mpState.settings;
  mp.appendChild(mpEl('div', 'mp-title', 'Налаштування пошти'));

  const method = document.createElement('select');
  const optOutlook = new Option('Outlook на цьому комп’ютері', 'outlook');
  optOutlook.disabled = !cfg.outlook_available;
  method.append(optOutlook, new Option('Пошта через SMTP (Gmail тощо)', 'smtp'));
  method.value = cfg.method === 'outlook' && !cfg.outlook_available ? 'smtp' : cfg.method;
  mp.appendChild(method);
  const hint = mpEl('div', 'mp-hint');
  hint.style.marginTop = '8px';
  mp.appendChild(hint);

  const smtpBox = document.createElement('div');
  const host = mpInput('text', 'SMTP-сервер (smtp.gmail.com)', cfg.host);
  const port = mpInput('text', 'Порт (465)', cfg.port);
  const user = mpInput('email', 'Ваша пошта (логін)', cfg.username);
  const pass = mpInput('password', cfg.has_password ? 'Пароль збережено (залиште порожнім)' : 'Пароль застосунку', '');
  const from = mpInput('text', 'Ім’я відправника (необов’язково)', cfg.from_name);
  smtpBox.append(host, port, user, pass, from);
  mp.appendChild(smtpBox);

  const syncMethod = () => {
    const outlook = method.value === 'outlook';
    smtpBox.hidden = outlook;
    hint.textContent = outlook
      ? 'Листи надсилаються одразу від імені вашої пошти в класичному Outlook — паролі не потрібні, ' +
        'надіслане видно в «Надіслані». Outlook має бути встановлений і налаштований.'
      : (cfg.outlook_available ? '' : 'Outlook доступний лише у Windows із класичним Outlook. ') +
        'Пошта, з якої надсилатимуться листи (SMTP). Для Gmail: smtp.gmail.com, порт 465, пароль застосунку.';
  };
  method.addEventListener('change', syncMethod);
  syncMethod();

  const status = mpEl('div', 'mp-status' + (mpState.statusErr ? ' err' : ''), mpState.status);
  mp.appendChild(status);
  const actions = mpEl('div', 'mp-actions');
  const back = mpEl('button', 'mp-btn ghost', 'Назад');
  back.addEventListener('click', () => { mpState.view = 'list'; mpState.status = ''; mpRender(); });
  const save = mpEl('button', 'mp-btn', 'Зберегти');
  save.addEventListener('click', async () => {
    try {
      mpState.settings = await mpApi('POST', '/api/mail/settings', {
        method: method.value, host: host.value, port: port.value, username: user.value,
        password: pass.value, from_name: from.value,
      });
      mpState.view = 'list';
      mpState.status = 'Налаштування збережено.';
      mpState.statusErr = false;
      mpRender();
    } catch (err) {
      status.textContent = err.message;
      status.className = 'mp-status err';
    }
  });
  actions.append(back, save);
  mp.appendChild(actions);
}

document.addEventListener('mousedown', e => {
  if (!mp.hidden && !mp.contains(e.target) && !e.target.closest('.tmail, .modal-ov')) mpClose();
});
document.addEventListener('keydown', e => { if (e.key === 'Escape' && !document.querySelector('.modal-ov')) mpClose(); });

function countFiles(node) {
  return node.children.reduce((sum, c) => sum + (c.dir ? countFiles(c) : 1), 0);
}

// Видалення з підтвердженням. Файли йдуть у системний кошик, тож їх можна відновити.
async function treeDelete(n) {
  const text = n.dir
    ? 'Папку «' + n.name + '» разом з усім вмістом (файлів: ' + countFiles(n) + ') буде переміщено до кошика. ' +
      'Її можна буде відновити з кошика.'
    : 'Файл «' + n.name + '» буде переміщено до кошика. Його можна буде відновити з кошика.';
  const sure = await mpConfirm(n.dir ? 'Видалити папку?' : 'Видалити файл?', text, 'Видалити');
  if (!sure) return;
  try {
    await mpApi('POST', '/api/tree/delete', { path: n.path });
    if (!mp.hidden && mpState.node && (mpState.node.path === n.path || mpState.node.path.startsWith(n.path + '/'))) mpClose();
    showToast('Переміщено до кошика: ' + n.name, 'ok');
  } catch (e) {
    showToast(e.message);
  }
  refreshTree();  // одразу показуємо, що змінилося на диску
}

function treeItem(n, isRoot) {
  const li = document.createElement('li');
  const row = document.createElement('div');
  row.className = 'trow' + (isRoot ? ' root' : '');
  row.title = n.name;
  const chev = document.createElement('span');
  chev.className = 'chev';
  if (n.dir) chev.appendChild(icon('chevron-right'));
  const ico = document.createElement('span');
  ico.className = 'ico ' + (isRoot ? 'ico-root' : n.dir ? 'ico-dir' : 'ico-file');
  ico.appendChild(icon(isRoot ? 'landmark' : n.dir ? (treeOpen.has(n.path) ? 'folder-open' : 'folder') : 'file-text'));
  const nm = document.createElement('span');
  nm.className = 'nm';
  nm.textContent = n.name;
  row.append(chev, ico, nm);
  li.appendChild(row);

  const mail = document.createElement('button');
  mail.className = 'treveal tmail';
  mail.title = 'Надіслати на пошту';
  mail.appendChild(icon('mail'));
  mail.addEventListener('click', e => { e.stopPropagation(); mpOpen(mail, n); });
  row.appendChild(mail);

  const reveal = document.createElement('button');
  reveal.className = 'treveal';
  reveal.title = 'Показати розташування на комп’ютері';
  reveal.appendChild(icon('external-link'));
  reveal.addEventListener('click', e => { e.stopPropagation(); treeApi('/api/tree/reveal', n.path); });
  row.appendChild(reveal);

  if (isRoot) {
    // Корінь не видаляємо (це уся папка результатів), але місце лишаємо, щоб кнопки стояли рівно.
    const gap = document.createElement('span');
    gap.className = 'treveal';
    gap.style.visibility = 'hidden';
    gap.appendChild(icon('trash-2'));
    row.appendChild(gap);
  } else {
    const trash = document.createElement('button');
    trash.className = 'treveal ttrash';
    trash.title = 'Видалити (у кошик)';
    trash.appendChild(icon('trash-2'));
    trash.addEventListener('click', e => { e.stopPropagation(); treeDelete(n); });
    row.appendChild(trash);
  }

  if (n.dir) {
    li.classList.toggle('open', treeOpen.has(n.path));
    row.addEventListener('click', () => {
      if (treeOpen.has(n.path)) treeOpen.delete(n.path); else treeOpen.add(n.path);
      renderTree();
    });
    if (n.children.length) {
      const ul = document.createElement('ul');
      n.children.forEach(c => ul.appendChild(treeItem(c, false)));
      li.appendChild(ul);
    }
  } else {
    row.addEventListener('click', () => treeApi('/api/tree/open', n.path));
  }
  return li;
}

function renderTree() {
  const keepScroll = treeEl.parentElement.scrollTop;
  treeEl.textContent = '';
  if (!treeData.name) {
    const empty = document.createElement('div');
    empty.className = 'tree-empty';
    empty.textContent = 'Завантажте порядок денний — тут з’явиться структура папок і файлів.';
    treeEl.appendChild(empty);
    return;
  }
  const ul = document.createElement('ul');
  ul.appendChild(treeItem({ name: treeData.name, path: '', dir: true, children: treeData.children }, true));
  treeEl.appendChild(ul);
  treeEl.parentElement.scrollTop = keepScroll;
}

async function refreshTree() {
  try {
    const data = await (await fetch('/api/tree')).json();
    const sig = JSON.stringify(data);
    if (sig === treeSig) return;
    treeSig = sig;
    treeData = data;
    renderTree();
  } catch (e) { /* сервер ще стартує — спробуємо за мить */ }
}
refreshTree();
setInterval(refreshTree, 2500);

const drop = document.getElementById('drop');
const fileInput = document.getElementById('file');
const fileChip = document.getElementById('fileChip');
const fname = document.getElementById('fname');
const fsize = document.getElementById('fsize');
const fclear = document.getElementById('fclear');
const startBtn = document.getElementById('startBtn');
const startLabel = document.getElementById('startLabel');
const spinner = document.getElementById('spinner');
const openBtn = document.getElementById('openBtn');
const outdirInput = document.getElementById('outdir');
const pickDirBtn = document.getElementById('pickDirBtn');
let lastWorkDir = outdirInput.value;

// Зберігає робочу папку на сервері; після успіху дерево показує її вміст.
async function saveWorkDir() {
  const value = outdirInput.value.trim();
  if (value === lastWorkDir) return;
  try {
    const resp = await fetch('/api/settings', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ work_dir: value }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(data.error || ('Помилка ' + resp.status));
    lastWorkDir = outdirInput.value = data.work_dir;
    showToast('Робочу папку збережено: ' + data.work_dir, 'ok');
    refreshTree();
  } catch (e) {
    outdirInput.value = lastWorkDir;  // лишаємо останню робочу
    showToast(e.message);
  }
}
outdirInput.addEventListener('change', saveWorkDir);
outdirInput.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); saveWorkDir(); } });

// Системне вікно вибору папки є лише у вікні програми (pywebview); у браузері шлях вводять вручну.
function updatePickButton() {
  pickDirBtn.hidden = !(window.pywebview && window.pywebview.api && window.pywebview.api.pick_folder);
}
window.addEventListener('pywebviewready', updatePickButton);
updatePickButton();
pickDirBtn.addEventListener('click', async () => {
  try {
    const chosen = await window.pywebview.api.pick_folder(outdirInput.value);
    if (chosen) { outdirInput.value = chosen; await saveWorkDir(); }
  } catch (e) { showToast('Не вдалося відкрити вікно вибору папки: ' + e); }
});
const logEl = document.getElementById('log');
const progressWrap = document.getElementById('progressWrap');
const statusText = document.getElementById('statusText');
const statusPct = document.getElementById('statusPct');
const barFill = document.getElementById('barFill');
const statsEl = document.getElementById('stats');
const statFound = document.getElementById('statFound');
const statOk = document.getElementById('statOk');
const statSkip = document.getElementById('statSkip');
const logDetails = document.getElementById('logDetails');
const auditDetails = document.getElementById('auditDetails');
const auditList = document.getElementById('auditList');
const auditTitle = document.getElementById('auditTitle');
let auditSignature = '';
let logAutoOpened = false;
let auditItems = [];  // поточний список зауважень (як його повернув сервер)
let awaitingFixes = false;  // файл має зауваження: скачування чекає, поки їх опрацюють
let proceeding = false;
const auditNotes = new Map();   // id -> повідомлення під формою (напр. "зауваження лишилося")
const auditDrafts = new Map();  // id -> вставлене, але ще не збережене посилання
const auditFixed = new Set();

let selectedFile = null;
let currentJob = null;
let pollTimer = null;

drop.addEventListener('click', () => fileInput.click());
drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('hover'); });
drop.addEventListener('dragleave', () => drop.classList.remove('hover'));
drop.addEventListener('drop', e => {
  e.preventDefault();
  drop.classList.remove('hover');
  if (e.dataTransfer.files.length) setFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', () => { if (fileInput.files.length) setFile(fileInput.files[0]); });

fclear.addEventListener('click', e => {
  e.stopPropagation();
  selectedFile = null;
  fileInput.value = '';
  fileChip.style.display = 'none';
  drop.style.display = 'block';
  startBtn.disabled = true;
});

function humanSize(bytes) {
  if (bytes < 1024) return bytes + ' Б';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(0) + ' КБ';
  return (bytes / (1024 * 1024)).toFixed(1) + ' МБ';
}

function setFile(f) {
  selectedFile = f;
  fname.textContent = f.name;
  fsize.textContent = humanSize(f.size);
  drop.style.display = 'none';
  fileChip.style.display = 'flex';
  startBtn.disabled = false;
}

function setRunning(running) {
  startBtn.disabled = running;
  spinner.style.display = running ? 'inline-block' : 'none';
  startLabel.textContent = running ? 'Виконується…' : 'Завантажити';
}

function resetRunUi(message) {
  setRunning(true);
  openBtn.style.display = 'none';
  statsEl.style.display = 'none';
  progressWrap.style.display = 'block';
  logDetails.open = false;
  logAutoOpened = false;
  awaitingFixes = false;
  proceeding = false;
  auditItems = [];
  logEl.textContent = '';
  auditDetails.style.display = 'none';
  auditList.textContent = '';
  auditSignature = '';
  auditFixed.clear();
  auditNotes.clear();
  auditDrafts.clear();
  barFill.style.width = '0%';
  statusPct.textContent = '';
  statusText.textContent = message;
}

startBtn.addEventListener('click', async () => {
  if (!selectedFile) return;
  resetRunUi('Завантажуємо PDF і читаємо порядок денний…');

  const form = new FormData();
  form.append('pdf', selectedFile);
  form.append('out_dir', outdirInput.value);

  let resp, data;
  try {
    resp = await fetch('/upload', { method: 'POST', body: form });
    data = await resp.json();
  } catch (e) {
    statusText.textContent = 'Помилка мережі: ' + e;
    setRunning(false);
    return;
  }
  if (data.error) {
    statusText.textContent = 'Помилка: ' + data.error;
    setRunning(false);
    return;
  }
  currentJob = data.job_id;
  refreshTree();
  poll();
});

function escapeHtml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function renderLog(lines) {
  const html = lines.map(l => {
    const esc = escapeHtml(l);
    if (l.startsWith('[+]')) return '<span class="l-ok">' + esc + '</span>';
    if (l.startsWith('[~]')) return '<span class="l-skip">' + esc + '</span>';
    if (l.startsWith('[!]') || l.startsWith('  -')) return '<span class="l-err">' + esc + '</span>';
    if (l.startsWith('Знайдено') || l.startsWith('Готово') || l.startsWith('Не вдалося') || l.startsWith('Перевірка') || l.startsWith('Увага')) return '<span class="l-head">' + esc + '</span>';
    return esc;
  }).join('\\n');
  // Автовиправлення (жовті рядки) показуємо одразу — один раз за запуск.
  if (!logAutoOpened && lines.some(l => l.startsWith('[~]'))) { logDetails.open = true; logAutoOpened = true; }
  // Не смикаємо прокрутку, якщо користувач сам піднявся вище по журналу.
  const atBottom = logEl.scrollHeight - logEl.scrollTop - logEl.clientHeight < 40;
  logEl.innerHTML = html;
  if (atBottom) logEl.scrollTop = logEl.scrollHeight;
}

function renderAudit(items) {
  auditItems = items || [];
  if (!auditItems.length) {
    auditSignature = '';
    auditList.textContent = '';
    auditDetails.style.display = 'none';
    return;
  }
  auditDetails.style.display = 'block';
  const signature = auditItems.map(i => i.id).join('|');
  if (signature !== auditSignature) {
    auditSignature = signature;
    auditList.textContent = '';
    auditItems.forEach(i => auditList.appendChild(auditRow(i)));
  }
  updateAuditTitle();
}

function auditRow(i) {
  const row = document.createElement('div');
  row.className = 'audit-item';
  row.dataset.id = i.id;

  const head = document.createElement('label');
  head.className = 'a-head';
  const cb = document.createElement('input');
  cb.type = 'checkbox';
  cb.checked = auditFixed.has(i.id);
  row.classList.toggle('fixed', cb.checked);
  cb.addEventListener('change', () => {
    if (cb.checked) auditFixed.add(i.id); else auditFixed.delete(i.id);
    row.classList.toggle('fixed', cb.checked);
    updateAuditTitle();
    checkAllResolved();
  });
  const where = 'розділ ' + i.section + (i.item !== null ? ', пункт ' + i.item : '');
  const text = document.createElement('span');
  text.innerHTML = '<span class="a-kind">' + escapeHtml(i.kind) + '</span>' +
    '<span class="a-where">' + escapeHtml(where) + '</span>' +
    (i.pr ? '<span class="a-pr">' + escapeHtml(i.pr) + '</span>' : '') +
    '<span class="a-text">' + escapeHtml(i.text) + '</span>';
  head.append(cb, text);
  row.appendChild(head);
  if (i.fix) auditFixForm(i, row);
  return row;
}

// Поле "вставити посилання" для зауважень, що лікуються підстановкою адреси.
function auditFixForm(i, row) {
  const form = document.createElement('div');
  form.className = 'a-fix';
  let pick = null;
  if (i.extra && i.extra.items) {
    pick = document.createElement('select');
    pick.title = 'У якому пункті замінити посилання';
    i.extra.items.forEach((it, idx) => {
      const o = document.createElement('option');
      o.value = String(idx);
      o.textContent = 'пункт ' + it.num + (it.pr ? ' · ' + it.pr : '');
      pick.appendChild(o);
    });
    pick.value = String(i.extra.items.length - 1);
    form.appendChild(pick);
  }
  const url = document.createElement('input');
  url.type = 'text';
  url.placeholder = 'Вставте посилання на проєкт рішення';
  url.value = auditDrafts.get(i.id) || '';
  url.addEventListener('input', () => auditDrafts.set(i.id, url.value));
  const save = document.createElement('button');
  save.type = 'button';
  save.textContent = 'Зберегти';
  form.append(url, save);
  const msg = document.createElement('div');
  msg.className = 'a-msg';
  if (auditNotes.has(i.id)) { msg.classList.add('err'); msg.textContent = auditNotes.get(i.id); }
  row.append(form, msg);

  const submit = async () => {
    if (!url.value.trim()) { msg.className = 'a-msg err'; msg.textContent = 'Вставте посилання.'; return; }
    save.disabled = true;
    msg.className = 'a-msg';
    msg.textContent = 'Зберігаємо у файлі й перевіряємо…';
    try {
      const resp = await fetch('/api/audit/fix', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ job_id: currentJob, audit_id: i.id, url: url.value,
          item_index: pick ? parseInt(pick.value, 10) : null }),
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) throw new Error(data.error || ('Помилка ' + resp.status));
      auditDrafts.delete(i.id);
      if (data.resolved) {
        auditNotes.delete(i.id);
        showToast('Виправлено: файл перевірено заново, зауваження прибрано зі списку.', 'ok');
      } else {
        auditNotes.set(i.id, data.message);
      }
      if (data.warning) showToast(data.warning);
      const log = await (await fetch('/api/log/' + currentJob)).json();
      renderLog(log.log);
      renderAudit(data.audit);
      checkAllResolved();
    } catch (e) {
      save.disabled = false;
      msg.className = 'a-msg err';
      msg.textContent = e.message;
    }
  };
  save.addEventListener('click', submit);
  url.addEventListener('keydown', e => { if (e.key === 'Enter') submit(); });
}

// Усі зауваження виправлені (зникли зі списку) або позначені галочкою —
// просимо сервер перевірити файл ще раз і, якщо все гаразд, почати скачування.
async function checkAllResolved() {
  if (!awaitingFixes || proceeding) return;
  if (auditItems.some(i => !auditFixed.has(i.id))) return;
  proceeding = true;
  setRunning(true);
  statusText.textContent = 'Усі зауваження опрацьовано — перевіряємо файл ще раз…';
  let data;
  try {
    const resp = await fetch('/api/proceed', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ job_id: currentJob, acknowledged: [...auditFixed] }),
    });
    data = await resp.json();
    if (!resp.ok) throw new Error(data.error || ('Помилка ' + resp.status));
  } catch (e) {
    statusText.textContent = 'Помилка: ' + e.message;
    setRunning(false);
    proceeding = false;
    return;
  }
  if (data.started) {
    const jobId = data.job_id;
    resetRunUi('Файл перевірено, починаємо завантаження…');
    currentJob = jobId;
    poll();
    return;
  }
  // При повторній перевірці знайшлися ще зауваження.
  setRunning(false);
  proceeding = false;
  statusText.textContent = 'Після повторної перевірки лишилися зауваження — виправте їх або позначте галочкою.';
  const log = await (await fetch('/api/log/' + currentJob)).json();
  renderLog(log.log);
  renderAudit(data.audit);
}

function updateAuditTitle() {
  const done = auditItems.filter(i => auditFixed.has(i.id)).length;
  auditTitle.textContent = 'Зауваження до PDF: опрацьовано ' + done + ' з ' + auditItems.length;
}

function updateProgress(lines) {
  const text = lines.join('\\n');
  const foundMatch = text.match(/Знайдено посилань на проєкти рішень: (\\d+)/);
  const readyIdx = lines.findIndex(l => l.startsWith('Готово:'));
  const relevant = readyIdx === -1 ? lines : lines.slice(0, readyIdx);
  const processed = relevant.filter(l => /^\\[[+=!]\\]/.test(l)).length;

  if (foundMatch) {
    const total = parseInt(foundMatch[1], 10);
    const pct = total > 0 ? Math.min(100, Math.round(processed / total * 100)) : 0;
    barFill.style.width = pct + '%';
    statusPct.textContent = pct + '%';
    statusText.textContent = readyIdx === -1
      ? `Обробка пунктів… ${Math.min(processed, total)} / ${total}`
      : 'Формуємо звіти й папку "Для депутатів"…';
  }

  const doneMatch = text.match(/Готово: завантажено\\/наявно (\\d+) з (\\d+), помилок: (\\d+)/);
  const skipMatch = text.match(/Знайдено (\\d+) пунктів, скачано (\\d+)\\. Пропущено (\\d+):/);
  if (doneMatch) {
    const ok = parseInt(doneMatch[1], 10);
    const total = foundMatch ? parseInt(foundMatch[1], 10) : parseInt(doneMatch[2], 10);
    const skipped = skipMatch ? parseInt(skipMatch[3], 10) : Math.max(0, total - ok);
    statFound.textContent = total;
    statOk.textContent = ok;
    statSkip.textContent = skipped;
    statsEl.style.display = 'grid';
  }
}

function poll() {
  pollTimer = setInterval(async () => {
    const resp = await fetch('/api/log/' + currentJob);
    const data = await resp.json();
    renderLog(data.log);
    renderAudit(data.audit);
    updateProgress(data.log);
    if (data.done && data.awaiting) {
      clearInterval(pollTimer);
      setRunning(false);
      awaitingFixes = true;
      statusPct.textContent = '';
      statusText.textContent = 'Завантаження не розпочато: у файлі є зауваження (' + data.audit.length +
        '). Виправте їх або позначте галочкою — файл буде перевірено ще раз, і завантаження почнеться саме.';
      checkAllResolved();
      return;
    }
    if (data.done) {
      clearInterval(pollTimer);
      setRunning(false);
      barFill.style.width = '100%';
      statusPct.textContent = '100%';
      const hadSkips = statSkip.textContent && statSkip.textContent !== '0';
      const hadAudit = !!(data.audit && data.audit.length);
      statusText.textContent = hadSkips ? 'Готово, є пропущені пункти — див. журнал'
        : hadAudit ? 'Готово, але в PDF порядку денного є зауваження — див. журнал'
        : 'Готово! Усі файли скачано';
      if (hadSkips) logDetails.open = true;
      openBtn.style.display = 'block';
      openBtn.onclick = () => fetch('/api/open/' + currentJob, { method: 'POST' });
    }
  }, 1000);
}

// ---- Вкладка "Генерація ПД" ----
const dropGen = document.getElementById('dropGen');
const fileInputGen = document.getElementById('fileGen');
const fileChipGen = document.getElementById('fileChipGen');
const fnameGen = document.getElementById('fnameGen');
const fsizeGen = document.getElementById('fsizeGen');
const fclearGen = document.getElementById('fclearGen');
const startBtnGen = document.getElementById('startBtnGen');
const startLabelGen = document.getElementById('startLabelGen');
const spinnerGen = document.getElementById('spinnerGen');
const openBtnGen = document.getElementById('openBtnGen');
const logGen = document.getElementById('logGen');
const progressWrapGen = document.getElementById('progressWrapGen');
const statusTextGen = document.getElementById('statusTextGen');
const statusPctGen = document.getElementById('statusPctGen');
const barFillGen = document.getElementById('barFillGen');
const statsGen = document.getElementById('statsGen');
const statFoundGen = document.getElementById('statFoundGen');
const statOkGen = document.getElementById('statOkGen');
const statSkipGen = document.getElementById('statSkipGen');
const logDetailsGen = document.getElementById('logDetailsGen');
const resultsLayout = document.getElementById('resultsLayout');
const resultsList = document.getElementById('resultsList');
const generateBtn = document.getElementById('generateBtn');
const toastStack = document.getElementById('toastStack');

let agendaActive = false;           // у завданні є порядок денний: дерево розділів — його власне
let lockedSections = new Set();    // заголовки з порядку денного (не перейменовуються й не видаляються)
const arrToggle = new Map();       // розділ -> розгорнутий/згорнутий вручну (блок "Порядок у документі")
const newInfo = document.getElementById('newInfo');
const agendaCards = document.getElementById('agendaCards');
const agendaList = document.getElementById('agendaList');
const agendaCount = document.getElementById('agendaCount');
let selectedFilesGen = [];      // черга: файли, додані, але ще не відправлені на обробку
let runningGen = false;
const processedNamesGen = new Set();  // імена вже відправлених файлів (щоб не додати двічі)
const pendingListGen = document.getElementById('pendingListGen');
const newBatchBtn = document.getElementById('newBatchBtn');
let currentJobGen = null;
let pollTimerGen = null;
let renderedResultsCount = 0;

const auditDetailsGen = document.getElementById('auditDetailsGen');
const auditListGen = document.getElementById('auditListGen');
const auditTitleGen = document.getElementById('auditTitleGen');
let auditItemsGen = [];        // поточні зауваження до назв файлів (як їх повернув сервер)
let auditSignatureGen = '';
let awaitingGen = false;       // є некоректні назви: обробка чекає
let proceedingGen = false;
const skippedGen = new Set();  // id файлів, позначених галочкою (пропустити)
const draftsGen = new Map();   // id -> введена, але ще не збережена нова назва

function showToast(message, kind) {
  const toast = document.createElement('div');
  toast.className = 'toast' + (kind === 'ok' ? ' toast-ok' : '');
  toast.textContent = message;
  toastStack.appendChild(toast);
  requestAnimationFrame(() => toast.classList.add('show'));
  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => toast.remove(), 250);
  }, 6000);
}

// ---- Розділи порядку денного: дерево до 4 рівнів (1, 1.1, 1.1.1, 1.1.1.1), назви редагуються й зберігаються ----
const sectionsList = document.getElementById('sectionsList');
const MAX_SECTION_DEPTH = 4;
let sectionTitles = {};   // id -> назва
let sectionOrder = [];    // id у порядку документа: 1, 1.1, 1.1.1, 1.2, 2, …
const sectionDepth = id => id.split('.').length;
const sectionChildren = id => sectionOrder.filter(x => x.startsWith(id + '.') && sectionDepth(x) === sectionDepth(id) + 1);
const sectionBranch = id => sectionOrder.filter(x => x === id || x.startsWith(id + '.'));  // розділ і все вкладене
const topSections = () => sectionOrder.filter(id => sectionDepth(id) === 1);

function applySections(data) {
  sectionTitles = data.sections || {};
  sectionOrder = data.order || [];
  lockedSections = new Set(data.locked || []);
  renderSectionsList();
  rebuildCardSections();  // кнопки розділів на картках залежать від наявних підрозділів
  renderArrange();
}

function renderSectionsList() {
  sectionsList.textContent = '';
  sectionOrder.forEach(id => sectionsList.appendChild(renderSectionItem(id)));
}

function sectionIconButton(cls, iconName, title, handler) {
  const b = document.createElement('button');
  b.className = 's-edit-btn ' + cls;
  b.type = 'button';
  b.title = title;
  b.appendChild(icon(iconName));
  b.addEventListener('click', handler);
  return b;
}

function renderSectionItem(id) {
  const depth = sectionDepth(id);
  const li = document.createElement('li');
  li.dataset.section = id;
  li.style.paddingLeft = ((depth - 1) * 10) + 'px';

  const numEl = document.createElement('span');
  numEl.className = 's-num';
  numEl.textContent = id + '.';
  li.appendChild(numEl);

  const textEl = document.createElement('span');
  textEl.className = 's-text';
  textEl.textContent = sectionTitles[id];
  li.appendChild(textEl);

  const locked = lockedSections.has(id);
  if (!locked) li.appendChild(sectionIconButton('s-rename', 'pencil', 'Змінити назву', () => startEditingSection(li, id)));
  if (depth < MAX_SECTION_DEPTH) {
    li.appendChild(sectionIconButton('s-add', 'plus', 'Додати підрозділ', () => addSubsection(id)));
  }
  if (depth > 1 && !locked) {
    li.appendChild(sectionIconButton('s-del', 'trash-2', 'Видалити підрозділ', () => deleteSubsection(id)));
  }
  return li;
}

async function sectionsApi(url, body) {
  const resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(agendaActive ? { ...body, job_id: currentJobGen } : body),
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok || data.error) throw new Error(data.error || ('Помилка ' + resp.status));
  return data;
}

// Додає підрозділ (1 -> 1.1, 1.1 -> 1.1.1 …) і одразу відкриває поле для його назви.
async function addSubsection(parent) {
  try {
    const data = await sectionsApi('/api/sections/add', { parent });
    applySections(data);
    const li = sectionsList.querySelector('li[data-section="' + data.id + '"]');
    if (li) { li.scrollIntoView({ block: 'nearest' }); startEditingSection(li, data.id); }
  } catch (e) { showToast(e.message); }
}

async function deleteSubsection(id) {
  const kids = sectionBranch(id).length - 1;
  const used = resultsData.filter(r => r.section && (r.section === id || r.section.startsWith(id + '.'))).length;
  const text = 'Підрозділ «' + id + ' ' + sectionTitles[id] + '»' + (kids ? ' разом із вкладеними (' + kids + ')' : '') +
    ' буде видалено.' + (used ? ' Картки з нього (' + used + ') повернуться в розділ вище.' : '') +
    ' Номери наступних підрозділів зсунуться.';
  if (!(await mpConfirm('Видалити підрозділ?', text, 'Видалити'))) return;
  try {
    applySections(await sectionsApi('/api/sections/delete', { section: id }));
    await syncResultsFromServer();  // картки з видалених підрозділів сервер уже переніс
  } catch (e) { showToast(e.message); }
}

function startEditingSection(li, id) {
  const textEl = li.querySelector('.s-text');
  const buttons = li.querySelectorAll('.s-edit-btn');
  const currentTitle = sectionTitles[id] || '';

  const input = document.createElement('input');
  input.className = 's-input';
  input.type = 'text';
  input.value = currentTitle;

  textEl.replaceWith(input);
  buttons.forEach(b => { b.style.display = 'none'; });
  input.focus();
  input.select();

  let settled = false;
  const finish = async (save) => {
    if (settled) return;
    settled = true;
    const newTitle = input.value.trim();
    if (save && newTitle && newTitle !== currentTitle) {
      try {
        const data = await sectionsApi('/api/sections', { section: id, title: newTitle });
        sectionTitles = data.sections;
      } catch (e) {
        showToast(e.message.startsWith('Помилка') ? 'Помилка мережі під час збереження назви розділу' : 'Помилка: ' + e.message);
      }
    }
    const newTextEl = document.createElement('span');
    newTextEl.className = 's-text';
    newTextEl.textContent = sectionTitles[id];
    input.replaceWith(newTextEl);
    buttons.forEach(b => { b.style.display = ''; });
    renderArrange();
  };

  input.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); finish(true); }
    if (e.key === 'Escape') { e.preventDefault(); finish(false); }
  });
  input.addEventListener('blur', () => finish(true));
}

async function loadSections() {
  try {
    const resp = await fetch('/api/sections');
    applySections(await resp.json());
  } catch (e) { /* панель лишиться порожньою, не критично */ }
}
loadSections();

dropGen.addEventListener('click', () => fileInputGen.click());
dropGen.addEventListener('dragover', e => { e.preventDefault(); dropGen.classList.add('hover'); });
dropGen.addEventListener('dragleave', () => dropGen.classList.remove('hover'));
dropGen.addEventListener('drop', e => {
  e.preventDefault();
  dropGen.classList.remove('hover');
  if (e.dataTransfer.files.length) addFilesGen(e.dataTransfer.files);
});
fileInputGen.addEventListener('change', () => {
  if (fileInputGen.files.length) addFilesGen(fileInputGen.files);
  fileInputGen.value = '';  // щоб той самий файл можна було обрати ще раз після видалення зі списку
});

fclearGen.addEventListener('click', e => {
  e.stopPropagation();
  selectedFilesGen = [];
  renderPendingGen();
});

// Файли додаються до черги, а не замінюють її: можна кидати по одному, скільки потрібно.
function addFilesGen(fileList) {
  const notPdf = [], dup = [];
  for (const f of Array.from(fileList)) {
    if (!/[.](pdf|docx)$/i.test(f.name)) notPdf.push(f.name);
    else if (selectedFilesGen.some(x => x.name === f.name) || processedNamesGen.has(f.name)) dup.push(f.name);
    else selectedFilesGen.push(f);
  }
  if (notPdf.length) showToast('Лише PDF і Word (.docx). Пропущено: ' + notPdf.slice(0, 3).join(', ') + (notPdf.length > 3 ? '…' : ''));
  if (dup.length) showToast('Уже додано або оброблено: ' + dup.slice(0, 3).join(', ') + (dup.length > 3 ? '…' : ''));
  renderPendingGen();
}

function renderPendingGen() {
  const n = selectedFilesGen.length;
  pendingListGen.textContent = '';
  fileChipGen.style.display = n ? 'flex' : 'none';
  pendingListGen.style.display = n ? 'block' : 'none';
  if (n) {
    fnameGen.textContent = 'Файлів у черзі: ' + n;
    fsizeGen.textContent = humanSize(selectedFilesGen.reduce((sum, f) => sum + f.size, 0));
  }
  selectedFilesGen.forEach(f => {
    const row = document.createElement('div');
    row.className = 'p-row';
    const name = document.createElement('span');
    name.className = 'p-name';
    name.textContent = f.name;
    name.title = f.name;
    const size = document.createElement('span');
    size.className = 'p-size';
    size.textContent = humanSize(f.size);
    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'p-del';
    del.title = 'Прибрати з черги';
    del.appendChild(icon('x'));
    del.addEventListener('click', () => {
      selectedFilesGen = selectedFilesGen.filter(x => x !== f);
      renderPendingGen();
    });
    row.append(name, size);
    if (/[.]docx$/i.test(f.name) || !/^[0-9]+-[0-9]{2,4}/.test(f.name)) {
      const tag = document.createElement('span');
      tag.className = 'new-badge';
      tag.textContent = /[.]docx$/i.test(f.name) ? 'порядок денний' : 'можливо, порядок денний';
      tag.style.marginLeft = '0';
      row.appendChild(tag);
    }
    row.append(del);
    pendingListGen.appendChild(row);
  });
  updateStartGen();
}

function updateStartGen() {
  startBtnGen.disabled = runningGen || selectedFilesGen.length === 0;
  spinnerGen.style.display = runningGen ? 'inline-block' : 'none';
  startLabelGen.textContent = runningGen ? 'Виконується…'
    : (selectedFilesGen.length ? 'Обробити файлів: ' + selectedFilesGen.length : 'Почати');
}

function setRunningGen(running) {
  runningGen = running;
  updateStartGen();
}

// Виходить із режиму порядку денного: повертає загальне дерево розділів і прибирає його картки.
function resetAgendaUi() {
  const wasActive = agendaActive;
  agendaActive = false;
  lockedSections = new Set();
  arrToggle.clear();
  agendaList.textContent = '';
  agendaCards.style.display = 'none';
  agendaCount.textContent = '0';
  newInfo.textContent = '';
  if (wasActive) loadSections();
}

function resetRunUiGen(message, keepResults) {
  setRunningGen(true);
  openBtnGen.style.display = 'none';
  statsGen.style.display = 'none';
  progressWrapGen.style.display = 'block';
  logDetailsGen.open = false;
  logGen.textContent = '';
  barFillGen.style.width = '0%';
  statusPctGen.textContent = '';
  statusTextGen.textContent = message;
  if (!keepResults) {
    resetAgendaUi();
    resultsList.innerHTML = '';
    resultsData = [];
    renderArrange();
    resultsLayout.style.display = 'none';
    generateBtn.style.display = 'none';
    newBatchBtn.style.display = 'none';
    renderedResultsCount = 0;
  }
  awaitingGen = false;
  proceedingGen = false;
  auditItemsGen = [];
  auditSignatureGen = '';
  auditListGen.textContent = '';
  auditDetailsGen.style.display = 'none';
  skippedGen.clear();
  draftsGen.clear();
}

startBtnGen.addEventListener('click', async () => {
  if (!selectedFilesGen.length) return;
  // Якщо вже є оброблені результати — нові файли дописуються до них, а не замінюють їх.
  const append = !!currentJobGen && renderedResultsCount > 0;
  const prevJob = append ? currentJobGen : null;
  resetRunUiGen(`Завантажуємо файлів: ${selectedFilesGen.length}…`, append);

  const form = new FormData();
  for (const f of selectedFilesGen) form.append('files', f);
  if (prevJob) form.append('prev_job', prevJob);

  let resp, data;
  try {
    resp = await fetch('/generate/upload', { method: 'POST', body: form });
    data = await resp.json();
  } catch (e) {
    statusTextGen.textContent = 'Помилка мережі: ' + e;
    setRunningGen(false);
    return;
  }
  if (data.error) {
    statusTextGen.textContent = 'Помилка: ' + data.error;
    setRunningGen(false);
    return;
  }
  selectedFilesGen.forEach(f => processedNamesGen.add(f.name));
  selectedFilesGen = [];
  renderPendingGen();
  currentJobGen = data.job_id;
  pollGen();
});

newBatchBtn.addEventListener('click', async () => {
  const sure = await mpConfirm('Почати заново?',
    'Список оброблених проєктів буде очищено, щоб скласти новий. Уже створений документ залишиться.', 'Очистити');
  if (!sure) return;
  clearInterval(pollTimerGen);
  currentJobGen = null;
  renderedResultsCount = 0;
  resetAgendaUi();
  resultsList.innerHTML = '';
  resultsData = [];
  renderArrange();
  resultsLayout.style.display = 'none';
  generateBtn.style.display = 'none';
  newBatchBtn.style.display = 'none';
  progressWrapGen.style.display = 'none';
  processedNamesGen.clear();
  selectedFilesGen = [];
  setRunningGen(false);
  renderPendingGen();
});

function renderLogGen(lines) {
  const html = lines.map(l => {
    const esc = escapeHtml(l);
    if (l.startsWith('[+]')) return '<span class="l-ok">' + esc + '</span>';
    if (l.startsWith('[~]')) return '<span class="l-skip">' + esc + '</span>';
    if (l.startsWith('[!]')) return '<span class="l-err">' + esc + '</span>';
    if (l.startsWith('Знайдено') || l.startsWith('Готово') || l.startsWith('Перевірка')) return '<span class="l-head">' + esc + '</span>';
    return esc;
  }).join('\\n');
  logGen.innerHTML = html;
  logGen.scrollTop = logGen.scrollHeight;
}

function updateProgressGen(lines) {
  const text = lines.join('\\n');
  const foundMatch = text.match(/Знайдено файлів для обробки: (\\d+)/);
  const readyIdx = lines.findIndex(l => l.startsWith('Готово:'));
  const relevant = readyIdx === -1 ? lines : lines.slice(0, readyIdx);
  const processed = relevant.filter(l => /^\\[[+!]\\]/.test(l)).length;

  if (foundMatch) {
    const total = parseInt(foundMatch[1], 10);
    const pct = total > 0 ? Math.min(100, Math.round(processed / total * 100)) : 0;
    barFillGen.style.width = pct + '%';
    statusPctGen.textContent = pct + '%';
    statusTextGen.textContent = readyIdx === -1
      ? `Обробка файлів… ${Math.min(processed, total)} / ${total}`
      : 'Завершуємо…';
  }

  const doneMatch = text.match(/Готово: оброблено (\\d+) з (\\d+), помилок: (\\d+)/);
  if (doneMatch) {
    const ok = parseInt(doneMatch[1], 10);
    const total = parseInt(doneMatch[2], 10);
    const errors = parseInt(doneMatch[3], 10);
    statFoundGen.textContent = total;
    statOkGen.textContent = ok;
    statSkipGen.textContent = errors;
    statsGen.style.display = 'grid';
  }
}

// ---- Порядок у документі: пункти згруповані за розділами, як у документі ----
const arrangePanel = document.getElementById('arrangePanel');
const arrangeBody = document.getElementById('arrangeBody');
let dragIndex = null;

// Пункти розділу в порядку документа: спершу розставлені користувачем (rank), далі як завантажено.
function sectionItems(section) {
  return resultsData.map((r, i) => ({ r, i })).filter(x => x.r.section === section)
    .sort((a, b) => ((a.r.rank == null ? 1e9 : a.r.rank) - (b.r.rank == null ? 1e9 : b.r.rank)) || (a.i - b.i));
}

function renderArrange(movedIndex) {
  arrangeBody.textContent = '';
  if (!resultsData.length) { arrangePanel.style.display = 'none'; return; }
  arrangePanel.style.display = 'block';
  const missing = resultsData.filter(r => !r.locked && r.section == null).length;
  if (missing) {
    const note = document.createElement('div');
    note.className = 'arr-note';
    note.textContent = 'Оберіть розділ для всіх ' + (agendaActive ? 'нових карток' : 'карток') + ' вище (лишилось: ' + missing +
      ') — тоді тут з’явиться список у тому порядку, у якому пункти стоятимуть у документі, і їх можна буде переставляти.';
    arrangeBody.appendChild(note);
    return;
  }
  for (const section of sectionOrder) {
    const items = sectionItems(section);
    const fresh = items.filter(x => !x.r.locked).length;
    // Розділи з новими пунктами (і короткі) розгорнуті, великі розділи порядку денного — згорнуті.
    const open = arrToggle.has(section) ? arrToggle.get(section) : (fresh > 0 || items.length <= 6);
    const block = document.createElement('div');
    block.className = 'arr-section' + (open ? ' open' : '');
    block.style.marginLeft = ((sectionDepth(section) - 1) * 20) + 'px';
    const head = document.createElement('h4');
    head.className = 'arr-head';
    const count = document.createElement('span');
    count.className = 'arr-count';
    count.textContent = '(' + items.length + (agendaActive && fresh ? ', нових: ' + fresh : '') + ')';
    head.append(icon('chevron-right'), document.createTextNode(section + ') ' + (sectionTitles[section] || '') + ':'), count);
    head.addEventListener('click', () => { arrToggle.set(section, !open); renderArrange(); });
    const list = document.createElement('div');
    list.className = 'arr-list';
    list.dataset.section = section;
    if (open) {
      if (!items.length) {
        const empty = document.createElement('div');
        empty.className = 'arr-empty';
        empty.textContent = 'Немає пунктів — можна перетягнути сюди.';
        list.appendChild(empty);
      }
      items.forEach((entry, pos) => list.appendChild(arrangeRow(entry, pos, items.length, section, movedIndex)));
    }
    block.append(head, list);
    arrangeBody.appendChild(block);
  }
}

function arrangeRow(entry, pos, total, section, movedIndex) {
  const locked = !!entry.r.locked;
  const row = document.createElement('div');
  row.className = 'arr-item' + (entry.i === movedIndex ? ' moved' : '') + (locked ? ' locked' : (agendaActive ? ' is-new' : ''));
  row.draggable = !locked;
  row.dataset.index = String(entry.i);

  const grip = document.createElement('span');
  grip.className = 'arr-grip';
  if (!locked) grip.appendChild(icon('grip-vertical'));
  else grip.style.width = '16px';
  const num = document.createElement('span');
  num.className = 'arr-num';
  num.textContent = (pos + 1) + ')';
  const body = document.createElement('div');
  body.className = 'arr-body';
  const title = document.createElement('div');
  title.className = 'arr-title';
  title.textContent = (entry.r.line || '').trim();
  body.appendChild(title);
  if (!locked) {
    const file = document.createElement('div');
    file.className = 'arr-file';
    file.textContent = entry.r.filename;
    if (agendaActive) {
      const badge = document.createElement('span');
      badge.className = 'new-badge';
      badge.textContent = 'Новий';
      file.appendChild(badge);
    }
    body.appendChild(file);
  }
  row.append(grip, num, body);

  if (!locked) {  // пункти з порядку денного нерухомі; лише нові можна переставляти
    const btns = document.createElement('div');
    btns.className = 'arr-btns';
    const up = document.createElement('button');
    up.type = 'button';
    up.title = 'Вище в розділі';
    up.disabled = pos === 0;
    up.appendChild(icon('arrow-up'));
    up.addEventListener('click', () => arrangeMove(entry.i, section, pos - 1));
    const down = document.createElement('button');
    down.type = 'button';
    down.title = 'Нижче в розділі';
    down.disabled = pos === total - 1;
    down.appendChild(icon('arrow-down'));
    down.addEventListener('click', () => arrangeMove(entry.i, section, pos + 1));
    btns.append(up, down);
    row.appendChild(btns);
  }
  return row;
}

// position — місце серед пунктів розділу без самого пункту (0 — першим).
async function arrangeMove(index, section, position) {
  try {
    const resp = await fetch('/api/generate/arrange', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ job_id: currentJobGen, index, section, position }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(data.error || ('Помилка ' + resp.status));
    data.results.forEach((r, i) => { resultsData[i] = resultFields(r); });
    // картки лишаються на місці, але підсвітка розділу в них має відповідати
    resultsList.querySelectorAll('.result-item').forEach((card, i) => {
      /* кнопки розділів перебудовуються нижче */
    });
    rebuildCardSections();
    renderArrange(index);
    setTimeout(() => arrangeBody.querySelectorAll('.arr-item.moved').forEach(e => e.classList.remove('moved')), 1500);
  } catch (e) { showToast(e.message); }
}

// Список розділу під курсором: сам список або (для заголовка, зокрема згорнутого розділу) список цього розділу.
function dropList(el) {
  const list = el.closest('.arr-list');
  if (list) return list;
  const sec = el.closest('.arr-section');
  return sec ? sec.querySelector('.arr-list') : null;
}

function clearDropMarks() {
  arrangeBody.querySelectorAll('.drop-before, .drop-after, .over').forEach(e => e.classList.remove('drop-before', 'drop-after', 'over'));
}

// Куди впаде пункт: перед/після рядка, над яким курсор, або в кінець розділу (порожнє місце списку).
function dropTarget(e, list) {
  const row = e.target.closest('.arr-item');
  if (row && list.contains(row)) {
    const rect = row.getBoundingClientRect();
    return { row, before: e.clientY < rect.top + rect.height / 2 };
  }
  return null;
}

arrangeBody.addEventListener('dragstart', e => {
  const row = e.target.closest ? e.target.closest('.arr-item') : null;
  if (!row) return;
  dragIndex = parseInt(row.dataset.index, 10);
  row.classList.add('dragging');
  e.dataTransfer.effectAllowed = 'move';
  try { e.dataTransfer.setData('text/plain', String(dragIndex)); } catch (err) { /* не критично */ }
});
arrangeBody.addEventListener('dragend', () => {
  dragIndex = null;
  clearDropMarks();
  arrangeBody.querySelectorAll('.dragging').forEach(e => e.classList.remove('dragging'));
});
arrangeBody.addEventListener('dragover', e => {
  if (dragIndex === null) return;
  const list = dropList(e.target);
  if (!list) return;
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
  clearDropMarks();
  const target = dropTarget(e, list);
  if (target) target.row.classList.add(target.before ? 'drop-before' : 'drop-after');
  else list.classList.add('over');
});
arrangeBody.addEventListener('drop', e => {
  if (dragIndex === null) return;
  const list = dropList(e.target);
  if (!list) return;
  e.preventDefault();
  const section = list.dataset.section;
  const target = dropTarget(e, list);
  const index = dragIndex;
  dragIndex = null;
  clearDropMarks();
  const others = sectionItems(section).filter(x => x.i !== index);
  let position = others.length;
  if (target) {
    const targetIndex = parseInt(target.row.dataset.index, 10);
    if (targetIndex === index) return;  // кинули на самого себе
    const k = others.findIndex(x => x.i === targetIndex);
    if (k >= 0) position = target.before ? k : k + 1;
  }
  arrangeMove(index, section, position);
});

// Кнопка верхнього розділу, у якого є підрозділи, відкриває список "куди додати картку":
// сам розділ або будь-який вкладений підрозділ. Без підрозділів — звичайна кнопка.
const sectionMenu = document.createElement('div');
sectionMenu.className = 'sec-menu';
sectionMenu.hidden = true;
document.body.appendChild(sectionMenu);
function closeSectionMenu() { sectionMenu.hidden = true; }
document.addEventListener('mousedown', e => {
  if (!sectionMenu.hidden && !sectionMenu.contains(e.target) && !e.target.closest('.section-btn')) closeSectionMenu();
});
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeSectionMenu(); });
window.addEventListener('resize', closeSectionMenu);

function openSectionMenu(btn, index, top, wrap) {
  sectionMenu.textContent = '';
  const current = resultsData[index] ? resultsData[index].section : null;
  sectionBranch(top).forEach(id => {
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'sec-menu-item' + (id === current ? ' current' : '');
    item.style.paddingLeft = (12 + (sectionDepth(id) - 1) * 16) + 'px';
    item.textContent = id + ' · ' + (sectionTitles[id] || '');
    item.addEventListener('click', () => { closeSectionMenu(); selectResultSection(index, id, wrap); });
    sectionMenu.appendChild(item);
  });
  sectionMenu.hidden = false;
  const r = btn.getBoundingClientRect();
  const w = sectionMenu.offsetWidth, h = sectionMenu.offsetHeight;
  sectionMenu.style.left = Math.max(8, Math.min(r.left, window.innerWidth - w - 8)) + 'px';
  sectionMenu.style.top = (r.bottom + 4 + h > window.innerHeight ? Math.max(8, r.top - h - 4) : r.bottom + 4) + 'px';
}

function fillSectionButtons(wrap, index, selected) {
  wrap.textContent = '';
  topSections().forEach(top => {
    const hasKids = sectionChildren(top).length > 0;
    const inBranch = selected === top || (!!selected && selected.startsWith(top + '.'));
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'section-btn' + (inBranch ? ' active' : '');
    btn.textContent = inBranch && selected !== top ? selected : top + ' розділ';
    if (hasKids) btn.appendChild(icon('chevron-down'));
    btn.title = inBranch && selected !== top ? selected + ' ' + sectionTitles[selected] : '';
    btn.addEventListener('click', () => {
      if (hasKids) openSectionMenu(btn, index, top, wrap);
      else selectResultSection(index, top, wrap);
    });
    wrap.appendChild(btn);
  });
}

function makeSectionButtons(index, selected) {
  const wrap = document.createElement('div');
  wrap.className = 'r-sections';
  fillSectionButtons(wrap, index, selected);
  return wrap;
}

// Після зміни дерева розділів перебудовує кнопки на всіх картках.
function rebuildCardSections() {
  resultsList.querySelectorAll('.result-item').forEach(card => {
    const i = parseInt(card.dataset.index, 10);
    const wrap = card.querySelector('.r-sections');
    if (wrap && resultsData[i]) fillSectionButtons(wrap, i, resultsData[i].section);
  });
}

// Підтягує з сервера актуальні розділи карток (напр. після видалення підрозділу).
async function syncResultsFromServer() {
  if (!currentJobGen) return;
  try {
    const data = await (await fetch('/api/log/' + currentJobGen)).json();
    (data.results || []).forEach((r, i) => { resultsData[i] = resultFields(r); });
    rebuildCardSections();
    updateNewInfo();
    renderArrange();
  } catch (e) { /* наступний poll синхронізує */ }
}

async function selectResultSection(index, section, wrapEl) {
  if (resultsData[index]) {
    if (resultsData[index].section !== section) delete resultsData[index].rank;
    resultsData[index].section = section;
  }
  fillSectionButtons(wrapEl, index, section);
  updateNewInfo();
  renderArrange();
  try {
    await fetch('/api/generate/select', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ job_id: currentJobGen, index, section }),
    });
  } catch (e) { /* наступний poll підтвердить/синхронізує стан */ }
}

let resultsData = [];  // копія результатів: назва, рядок, розділ, місце в розділі (rank)

function resultFields(r) {
  return { filename: r.filename, line: r.line, section: r.section, rank: r.rank, locked: !!r.locked, source: r.source };
}

function appendResultItem(item, index) {
  const div = document.createElement('div');
  div.className = 'result-item';
  div.dataset.index = String(index);

  // Пункт із порядку денного: компактна картка в згорнутій групі, розділ уже обрано й змінити його не можна.
  if (item.locked) {
    div.classList.add('is-agenda');
    const sec = document.createElement('span');
    sec.className = 'a-sec';
    sec.textContent = item.section;
    div.append(sec, document.createTextNode((item.line || '').trim()));
    agendaList.appendChild(div);
    agendaCards.style.display = 'block';
    agendaCount.textContent = String(agendaList.children.length);
    return;
  }

  if (agendaActive) div.classList.add('is-new');
  const fileEl = document.createElement('div');
  fileEl.className = 'r-file';
  const indexEl = document.createElement('span');
  indexEl.className = 'r-index';
  indexEl.textContent = String(index + 1);
  fileEl.appendChild(indexEl);
  fileEl.appendChild(document.createTextNode(item.filename));
  if (agendaActive) {
    const badge = document.createElement('span');
    badge.className = 'new-badge';
    badge.textContent = 'Новий';
    fileEl.appendChild(badge);
  }
  div.appendChild(fileEl);

  const textEl = document.createElement('div');
  textEl.className = 'r-text';
  textEl.textContent = item.line;
  div.appendChild(textEl);

  div.appendChild(makeSectionButtons(index, item.section));
  resultsList.appendChild(div);
}

// Скільки нових пунктів і для скількох ще не обрано розділ.
function updateNewInfo() {
  const fresh = resultsData.filter(r => !r.locked);
  const todo = fresh.filter(r => r.section == null).length;
  newInfo.textContent = agendaActive && fresh.length
    ? 'Нові пункти з файлів: ' + fresh.length + (todo ? ' · обрати розділ ще треба для: ' + todo : ' · розділи обрано для всіх')
    : '';
}

function renderResults(results) {
  if (!results || !results.length) return;
  resultsLayout.style.display = 'flex';
  generateBtn.style.display = 'flex';
  newBatchBtn.style.display = 'block';
  for (let i = renderedResultsCount; i < results.length; i++) {
    resultsData[i] = resultFields(results[i]);
    appendResultItem(results[i], i);
  }
  renderedResultsCount = results.length;
  updateNewInfo();
  renderArrange();
}

generateBtn.addEventListener('click', async () => {
  if (!currentJobGen) return;
  generateBtn.disabled = true;
  try {
    const resp = await fetch('/api/log/' + currentJobGen);
    const data = await resp.json();
    const results = data.results || [];

    const missing = results
      .map((r, i) => (r.section == null ? i + 1 : null))
      .filter(n => n !== null);

    if (missing.length) {
      const list = missing.length > 20
        ? missing.slice(0, 20).join(', ') + `… (ще ${missing.length - 20})`
        : missing.join(', ');
      showToast(`Оберіть розділ для пунктів: ${list}`);
      return;
    }

    const buildResp = await fetch('/api/generate/build', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ job_id: currentJobGen }),
    });
    const buildData = await buildResp.json();
    if (buildData.error) {
      showToast('Помилка генерації: ' + buildData.error);
      return;
    }
    showToast(`Документ створено: ${buildData.path}`, 'ok');
  } finally {
    generateBtn.disabled = false;
  }
});

// Зауваження до назв файлів — той самий вигляд, що й зауваження до порядку денного,
// але замість посилання — нова назва файлу; "Перейменувати" перейменовує файл у програмі.
function renderAuditGen(items) {
  auditItemsGen = items || [];
  if (!auditItemsGen.length) {
    auditSignatureGen = '';
    auditListGen.textContent = '';
    auditDetailsGen.style.display = 'none';
    return;
  }
  auditDetailsGen.style.display = 'block';
  const signature = auditItemsGen.map(i => i.id).join('|');
  if (signature !== auditSignatureGen) {
    auditSignatureGen = signature;
    auditListGen.textContent = '';
    auditItemsGen.forEach(i => auditListGen.appendChild(auditRowGen(i)));
  }
  updateAuditTitleGen();
}

function updateAuditTitleGen() {
  const done = auditItemsGen.filter(i => skippedGen.has(i.id)).length;
  auditTitleGen.textContent = 'Назви файлів: пропущено ' + done + ' з ' + auditItemsGen.length;
}

function auditRowGen(i) {
  const row = document.createElement('div');
  row.className = 'audit-item';
  row.dataset.id = i.id;

  const head = document.createElement('label');
  head.className = 'a-head';
  head.title = 'Позначте, щоб пропустити цей файл';
  const cb = document.createElement('input');
  cb.type = 'checkbox';
  cb.checked = skippedGen.has(i.id);
  row.classList.toggle('fixed', cb.checked);
  cb.addEventListener('change', () => {
    if (cb.checked) skippedGen.add(i.id); else skippedGen.delete(i.id);
    row.classList.toggle('fixed', cb.checked);
    updateAuditTitleGen();
    checkAllResolvedGen();
  });
  const text = document.createElement('span');
  text.innerHTML = '<span class="a-kind">' + escapeHtml(i.kind) + '</span>' +
    '<span class="a-where">' + escapeHtml('файл «' + i.filename + '»') + '</span>' +
    '<span class="a-text">' + escapeHtml(i.text) + '</span>';
  head.append(cb, text);
  row.appendChild(head);

  const form = document.createElement('div');
  form.className = 'a-fix';
  const input = document.createElement('input');
  input.type = 'text';
  input.placeholder = 'Нова назва, напр. 183-25.pdf';
  input.value = draftsGen.has(i.id) ? draftsGen.get(i.id) : (i.suggest || i.filename);
  input.addEventListener('input', () => draftsGen.set(i.id, input.value));
  const save = document.createElement('button');
  save.type = 'button';
  save.textContent = 'Перейменувати';
  form.append(input, save);
  const msg = document.createElement('div');
  msg.className = 'a-msg';
  if (i.suggest) msg.textContent = 'Підказка: ' + i.suggest;
  row.append(form, msg);

  const submit = async () => {
    save.disabled = true;
    msg.className = 'a-msg';
    msg.textContent = 'Перейменовуємо…';
    try {
      const resp = await fetch('/api/generate/rename', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ job_id: currentJobGen, audit_id: i.id, new_name: input.value }),
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) throw new Error(data.error || ('Помилка ' + resp.status));
      draftsGen.delete(i.id);
      skippedGen.delete(i.id);
      showToast('Файл перейменовано: ' + i.filename + ' → ' + data.new_name, 'ok');
      const log = await (await fetch('/api/log/' + currentJobGen)).json();
      renderLogGen(log.log);
      renderAuditGen(data.audit);
      checkAllResolvedGen();
    } catch (e) {
      save.disabled = false;
      msg.className = 'a-msg err';
      msg.textContent = e.message;
    }
  };
  save.addEventListener('click', submit);
  input.addEventListener('keydown', e => { if (e.key === 'Enter') submit(); });
  return row;
}

// Усі назви виправлені (зникли зі списку) або пропущені галочкою — сервер перевіряє їх
// ще раз і, якщо все гаразд, починає обробку.
async function checkAllResolvedGen() {
  if (!awaitingGen || proceedingGen) return;
  if (auditItemsGen.some(i => !skippedGen.has(i.id))) return;
  proceedingGen = true;
  setRunningGen(true);
  statusTextGen.textContent = 'Усі назви опрацьовано — перевіряємо ще раз…';
  let data;
  try {
    const resp = await fetch('/api/generate/proceed', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ job_id: currentJobGen, acknowledged: [...skippedGen] }),
    });
    data = await resp.json();
    if (!resp.ok) throw new Error(data.error || ('Помилка ' + resp.status));
  } catch (e) {
    statusTextGen.textContent = 'Помилка: ' + e.message;
    setRunningGen(false);
    proceedingGen = false;
    return;
  }
  if (data.started) {
    const jobId = data.job_id;
    resetRunUiGen('Назви перевірено, починаємо обробку…', true);
    currentJobGen = jobId;
    pollGen();
    return;
  }
  setRunningGen(false);
  proceedingGen = false;
  statusTextGen.textContent = 'Після повторної перевірки лишилися некоректні назви — виправте їх або позначте галочкою.';
  const log = await (await fetch('/api/log/' + currentJobGen)).json();
  renderLogGen(log.log);
  renderAuditGen(data.audit);
}

function pollGen() {
  pollTimerGen = setInterval(async () => {
    const resp = await fetch('/api/log/' + currentJobGen);
    const data = await resp.json();
    renderLogGen(data.log);
    updateProgressGen(data.log);
    if (data.failed) {
      clearInterval(pollTimerGen);
      setRunningGen(false);
      statusPctGen.textContent = '';
      statusTextGen.textContent = 'Не вдалося прочитати порядок денний — причина в журналі.';
      logDetailsGen.open = true;
      return;
    }
    if (data.agenda && data.sections && !agendaActive) {  // розділи завдання — заголовки з порядку денного
      agendaActive = true;
      applySections(data.sections);
    }
    renderResults(data.results);
    renderAuditGen(data.audit);
    if (data.done && data.awaiting) {
      clearInterval(pollTimerGen);
      setRunningGen(false);
      awaitingGen = true;
      statusPctGen.textContent = '';
      statusTextGen.textContent = 'Обробку не розпочато: у назвах файлів є помилки (' + data.audit.length +
        '). Перейменуйте файли або позначте їх галочкою, щоб пропустити — назви буде перевірено ще раз, і обробка почнеться сама.';
      checkAllResolvedGen();
      return;
    }
    if (data.done) {
      clearInterval(pollTimerGen);
      setRunningGen(false);
      barFillGen.style.width = '100%';
      statusPctGen.textContent = '100%';
      const hadErrors = statSkipGen.textContent && statSkipGen.textContent !== '0';
      statusTextGen.textContent = hadErrors ? 'Готово, є помилки — див. журнал' : 'Готово! Усі файли оброблено';
      if (hadErrors) logDetailsGen.open = true;
      openBtnGen.style.display = 'block';
      openBtnGen.onclick = () => fetch('/api/open/' + currentJobGen, { method: 'POST' });
    }
  }, 1000);
}
</script>
</body>
</html>
"""


AGENDA_STATE_PATH = BASE_DIR / "last_agenda.json"
TREE_MAX_NODES = 5000


def load_current_agenda() -> dict:
    """Остання завантажена повістка: назва (корінь дерева) і папка з результатами."""
    try:
        data = json.loads(AGENDA_STATE_PATH.read_text(encoding="utf-8"))
        return {"name": data.get("name"), "out_dir": data.get("out_dir")}
    except (OSError, json.JSONDecodeError):
        return {"name": None, "out_dir": None}


def set_current_agenda(name: str, out_dir: Path):
    current_agenda.update(name=name, out_dir=str(out_dir))
    AGENDA_STATE_PATH.write_text(json.dumps(current_agenda, ensure_ascii=False), encoding="utf-8")


current_agenda.update(load_current_agenda())


def agenda_title(filename: str) -> str:
    """Назва порядку денного з імені файлу + рік завантаження в дужках,
    якщо рік (напр. у даті 26.10.2026) в назві вже не вказано."""
    stem = Path(filename).stem
    if re.search(r"\b(?:19|20)\d{2}\b", stem):
        return stem
    return f"{stem} ({datetime.now().year})"


def get_work_dir() -> Path:
    """Робоча папка: куди скачуються файли і звідки дерево їх читає. Обирає користувач."""
    saved = _read_json(SETTINGS_PATH, {}).get("work_dir")
    return Path(saved).expanduser() if saved else DEFAULT_OUT_DIR


def save_work_dir(folder: Path):
    cfg = _read_json(SETTINGS_PATH, {})
    cfg["work_dir"] = str(folder)
    _write_json(SETTINGS_PATH, cfg)


def validate_work_dir(raw: str) -> Path:
    """Шлях від користувача -> готова до роботи папка (створюється, якщо її ще немає)."""
    raw = (raw or "").strip().strip('"')
    if not raw:
        raise ValueError("Вкажіть шлях до папки.")
    folder = Path(raw).expanduser()
    if not folder.is_absolute():
        raise ValueError("Вкажіть повний шлях до папки, напр. C:\\Users\\Ім’я\\Documents\\Порядок.")
    try:
        folder.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise ValueError(f"Не вдалося створити папку: {exc}") from exc
    if not folder.is_dir():
        raise ValueError("Цей шлях — не папка.")
    if not os.access(folder, os.W_OK):
        raise ValueError("У цю папку не можна записувати — оберіть іншу.")
    return folder


def detect_agenda(folder: Path) -> str:
    """Назва порядку денного для папки, яку користувач обрав як робочу: найновіший PDF
    у корені, що не схожий на скачаний пункт; інакше — назва самої папки."""
    pdfs = [p for p in folder.glob("*.pdf") if p.is_file() and not core.MAIN_FILENAME_RE.match(p.name)]
    if pdfs:
        return agenda_title(max(pdfs, key=lambda p: p.stat().st_mtime).name)
    return folder.name or str(folder)


def _natural_key(name: str):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def build_tree(directory: Path, root: Path, budget: list) -> list:
    """Вкладені {name, path, dir, children}: спершу папки, потім файли,
    у "натуральному" порядку (1.2 перед 1.10)."""
    try:
        entries = [e for e in directory.iterdir() if not e.name.startswith(".")]
    except OSError:
        return []
    entries.sort(key=lambda e: (not e.is_dir(), _natural_key(e.name)))
    nodes = []
    for entry in entries:
        if budget[0] <= 0:
            break
        budget[0] -= 1
        is_dir = entry.is_dir()
        nodes.append({
            "name": entry.name,
            "path": entry.relative_to(root).as_posix(),
            "dir": is_dir,
            "children": build_tree(entry, root, budget) if is_dir else [],
        })
    return nodes


def open_in_os(path: str):
    """Відкриває файл/папку в системному застосунку за замовчуванням —
    команда для цього різна на macOS/Windows/Linux."""
    system = platform.system()
    if system == "Darwin":
        subprocess.run(["open", path])
    elif system == "Windows":
        os.startfile(path)  # type: ignore[attr-defined]
    else:
        subprocess.run(["xdg-open", path])


def _reveal_windows_api(path) -> bool:
    """Відкриває Провідник із виділеним елементом системним викликом SHOpenFolderAndSelectItems —
    без командного рядка, тож пробіли, кирилиця й лапки в шляху нічого не ламають."""
    import ctypes
    from ctypes import wintypes
    shell32, ole32 = ctypes.windll.shell32, ctypes.windll.ole32
    shell32.SHParseDisplayName.argtypes = [wintypes.LPCWSTR, ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p),
                                           wintypes.ULONG, ctypes.POINTER(wintypes.ULONG)]
    shell32.SHParseDisplayName.restype = ctypes.c_long
    shell32.SHOpenFolderAndSelectItems.argtypes = [ctypes.c_void_p, wintypes.UINT, ctypes.c_void_p, wintypes.DWORD]
    shell32.SHOpenFolderAndSelectItems.restype = ctypes.c_long
    ole32.CoTaskMemFree.argtypes = [ctypes.c_void_p]
    initialized = ole32.CoInitialize(None) in (0, 1)  # запит Flask іде у власному потоці
    pidl = ctypes.c_void_p()
    try:
        if shell32.SHParseDisplayName(str(path), None, ctypes.byref(pidl), 0, None) != 0:
            return False
        try:
            return shell32.SHOpenFolderAndSelectItems(pidl, 0, None, 0) == 0  # 0 елементів = виділити сам шлях
        finally:
            ole32.CoTaskMemFree(pidl)
    finally:
        if initialized:
            ole32.CoUninitialize()


def reveal_in_os(path: Path) -> str:
    """Показує файл/папку в провіднику (з виділенням), а не відкриває її вміст. Повертає, який спосіб
    спрацював (його видно в консолі). Якщо не спрацював жоден — OSError."""
    system = platform.system()
    text = os.path.normpath(str(path))  # на Windows — зворотні скісні риски, як того вимагає Провідник
    if system == "Windows":
        attempts = []
        # 1) системний виклик; 2) команда з лапками лише навколо шляху (якщо передати списком, subprocess
        # візьме в лапки весь "/select,<шлях>" — Explorer його не розбере й відкриє "Документи");
        # 3) хоча б відкрити папку, де лежить елемент.
        attempts.append(("SHOpenFolderAndSelectItems", lambda: _reveal_windows_api(text)))
        attempts.append(("explorer /select", lambda: subprocess.Popen(f'explorer /select,"{text}"') is not None))
        attempts.append(("відкрито папку", lambda: os.startfile(str(path.parent)) is None))  # type: ignore[attr-defined]
        for name, attempt in attempts:
            try:
                if attempt():
                    print(f"[показати місце] {text} -> {name}", flush=True)
                    return name
            except Exception as exc:
                print(f"[показати місце] {name}: не вдалося ({exc})", flush=True)
        raise OSError(f"Провідник не відкрив {text}")
    if system == "Darwin":
        subprocess.run(["open", "-R", str(path)], check=True)
        return "open -R"
    subprocess.run(["xdg-open", str(path.parent)], check=True)
    return "xdg-open"


CONTACTS_PATH = BASE_DIR / "contacts.json"
MAIL_SETTINGS_PATH = BASE_DIR / "mail_settings.json"
GROUPS_PATH = BASE_DIR / "contact_groups.json"
# Більшість поштових серверів відсікають листи від ~25 МБ, а base64 додає ~33%.
MAX_ATTACHMENT_BYTES = 18 * 1024 * 1024
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


class MailError(Exception):
    """Помилка, текст якої можна показати користувачу як є."""


def _read_json(path: Path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return default


def _write_json(path: Path, data, private: bool = False):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    if private:  # у файлі пароль до пошти — читати його має лише власник
        try:
            os.chmod(path, 0o600)
        except OSError:
            pass


def build_attachment(target: Path, root: Path, agenda_name: str):
    """Файл — як є; папка — один zip (з тією ж назвою папки всередині).
    Повертає (ім'я файлу, mime-тип, байти)."""
    if target.is_file():
        files, base_name = [target], None
    else:
        files = [f for f in target.rglob("*") if f.is_file() and not f.name.startswith(".")]
        base_name = agenda_name if target == root else target.name
        if not files:
            raise MailError("Папка порожня — нічого надсилати.")
    total = sum(f.stat().st_size for f in files)
    limit_mb = MAX_ATTACHMENT_BYTES // (1024 * 1024)
    too_big = (f"Вкладення завелике ({total / 1048576:.0f} МБ): пошта приймає листи до ~{limit_mb} МБ. "
               "Надішліть окремі файли або меншу папку.")
    if total > MAX_ATTACHMENT_BYTES:
        raise MailError(too_big)

    if base_name is None:
        mime = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        return target.name, mime, target.read_bytes()

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, (Path(base_name) / f.relative_to(target)).as_posix())
    if buf.tell() > MAX_ATTACHMENT_BYTES:
        raise MailError(too_big)
    return core.sanitize_filename(base_name) + ".zip", "application/zip", buf.getvalue()


MAX_IMPORT_BYTES = 5 * 1024 * 1024
MAX_IMPORT_ROWS = 20000
EMAIL_FIND_RE = re.compile(r"[^@\s;,<>()\[\]]+@[^@\s;,<>()\[\]]+\.[^@\s;,<>()\[\]]+")
# Заголовки колонок з іменем (у нижньому регістрі, за підрядком).
NAME_HEADER_KEYS = ("прізвищ", "фамил", "ім'я", "імя", "імʼя", "ім’я", "имя", "по батьков", "по-батьков",
                    "отчество", "піб", "пiб", "фіо", "фио", "name", "співробітник", "сотрудник", "отримувач",
                    "контакт", "посадов")
EMAIL_HEADER_KEYS = ("mail", "пошт", "почт", "адрес")


def read_contacts_table(filename: str, data: bytes) -> list:
    """Рядки таблиці (списки текстових клітинок) з .xlsx або .csv."""
    ext = Path(filename).suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            rows = []
            for row in wb.active.iter_rows(values_only=True):
                rows.append(["" if v is None else str(v).strip() for v in row])
                if len(rows) > MAX_IMPORT_ROWS:
                    break
            wb.close()
            return rows
        except Exception as exc:  # битий/не той файл: zipfile.BadZipFile, KeyError, ...
            raise ValueError("Не вдалося прочитати файл Excel — це справжній .xlsx?") from exc
    if ext == ".xls":
        raise ValueError("Старий формат .xls не підтримується — збережіть файл як .xlsx або .csv.")
    if ext == ".csv":
        try:
            text = data.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = data.decode("cp1251", errors="replace")
        try:
            dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        return [[c.strip() for c in row] for row in csv.reader(io.StringIO(text), dialect)][:MAX_IMPORT_ROWS]
    raise ValueError("Підтримуються файли .xlsx та .csv.")


def parse_contacts_rows(rows: list):
    """З рядків таблиці дістає (ім'я, пошта). Колонку пошти шукає за вмістом, колонки
    імені — за заголовком (Прізвище, Ім'я, ПІБ, Name …), а без заголовка бере першу
    текстову колонку; без імені підставляє частину пошти до "@".
    Повертає ([(name, email)], кількість рядків без коректної пошти)."""
    rows = [r for r in rows if any(c for c in r)]
    if not rows:
        return [], 0
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]

    first_has_email = any(EMAIL_FIND_RE.search(c) for c in rows[0])
    header = None if first_has_email else [c.lower() for c in rows[0]]
    data_rows = rows if header is None else rows[1:]

    counts = [sum(1 for r in data_rows if EMAIL_FIND_RE.search(r[i])) for i in range(width)]
    if not any(counts):
        return [], len(data_rows)
    email_col = counts.index(max(counts))

    name_cols = []
    if header:
        name_cols = [i for i, h in enumerate(header)
                     if i != email_col and any(k in h for k in NAME_HEADER_KEYS)
                     and not any(k in h for k in EMAIL_HEADER_KEYS)]
    if not name_cols:
        text_cols = [i for i in range(width) if i != email_col and counts[i] == 0 and any(
            r[i] and not r[i].replace(".", "").replace(",", "").isdigit() for r in data_rows)]
        name_cols = text_cols[:1]

    found, invalid = [], 0
    for r in data_rows:
        emails = EMAIL_FIND_RE.findall(r[email_col])
        if not emails:
            if r[email_col]:
                invalid += 1
            continue
        name = re.sub(r"\s+", " ", " ".join(r[i] for i in name_cols if r[i])).strip()
        for email in emails:
            found.append((name or email.split("@")[0], email))
    return found, invalid


def move_to_trash(path: Path):
    """У системний кошик (Windows — "Кошик", macOS — Trash), а не назавжди."""
    try:
        from send2trash import send2trash
    except ImportError:
        if platform.system() != "Darwin":
            raise RuntimeError("не встановлено бібліотеку send2trash")
        subprocess.run(["osascript", "-e", "on run argv", "-e", 'tell application "Finder" to delete (POSIX file (item 1 of argv))',
                        "-e", "end run", str(path)], check=True, capture_output=True)
        return
    send2trash(str(path))


def outlook_available() -> bool:
    """Класичний Outlook керується через COM — це можливо лише у Windows (pywin32)."""
    if platform.system() != "Windows":
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


def _mail_method(cfg: dict) -> str:
    """"outlook" або "smtp". Якщо користувач ще не обирав — Outlook, коли він доступний."""
    return cfg.get("method") or ("outlook" if outlook_available() else "smtp")


def send_via_outlook(recipients: list, subject: str, item_name: str, attachment: tuple):
    """Надсилає одразу, від імені облікового запису за замовчуванням у класичному
    Outlook (паролі не потрібні, листи видно в "Надіслані"). Кожному — окремий лист."""
    if not outlook_available():
        raise MailError("Відправка через Outlook доступна лише у Windows із класичним Outlook.")
    import pythoncom
    import win32com.client

    filename, _, data = attachment
    pythoncom.CoInitialize()  # запит Flask іде у власному потоці — йому потрібна своя ініціалізація COM
    try:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            path = Path(tmp) / filename
            path.write_bytes(data)
            outlook = win32com.client.Dispatch("Outlook.Application")
            for person in recipients:
                mail = outlook.CreateItem(0)  # 0 = olMailItem
                mail.To = person["email"]
                mail.Subject = subject
                mail.Body = f"Доброго дня, {person['name']}!\n\nУ вкладенні: {item_name}.\n"
                mail.Attachments.Add(str(path))
                mail.Send()
    except Exception as exc:  # pywintypes.com_error та інші збої COM
        raise MailError(
            f"Не вдалося надіслати через Outlook: {exc}. Переконайтеся, що встановлено класичний Outlook "
            "(не «новий Outlook»), що він налаштований і не очікує відповіді на запит безпеки."
        ) from exc
    finally:
        pythoncom.CoUninitialize()


def send_mail(recipients: list, subject: str, item_name: str, attachment: tuple):
    """Окремий лист кожному отримувачу (адреси одне одного не бачать)."""
    cfg = _read_json(MAIL_SETTINGS_PATH, {})
    if _mail_method(cfg) == "outlook":
        return send_via_outlook(recipients, subject, item_name, attachment)
    host, user, password = cfg.get("host"), cfg.get("username"), cfg.get("password")
    port = int(cfg.get("port") or 465)
    if not (host and user and password):
        raise MailError("Спершу налаштуйте пошту: «Налаштування пошти» у цьому вікні.")
    filename, mime, data = attachment
    maintype, _, subtype = mime.partition("/")
    context = ssl.create_default_context()
    try:
        smtp = (smtplib.SMTP_SSL(host, port, timeout=60, context=context) if port == 465
                else smtplib.SMTP(host, port, timeout=60))
        with smtp:
            if port != 465:
                smtp.starttls(context=context)
            smtp.login(user, password)
            for person in recipients:
                msg = EmailMessage()
                msg["Subject"] = subject
                msg["From"] = formataddr((cfg.get("from_name", ""), user))
                msg["To"] = formataddr((person["name"], person["email"]))
                msg.set_content(f"Доброго дня, {person['name']}!\n\nУ вкладенні: {item_name}.\n")
                msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)
                smtp.send_message(msg)
    except smtplib.SMTPAuthenticationError as exc:
        raise MailError("Пошта не прийняла логін або пароль. Для Gmail потрібен «пароль застосунку», "
                        "а не звичайний пароль.") from exc
    except (smtplib.SMTPException, OSError) as exc:
        raise MailError(f"Не вдалося надіслати лист: {exc}") from exc


# Види зауважень, для яких в інтерфейсі є поле "вставити посилання".
FIXABLE_AUDIT_KINDS = ("немає посилання", "однакове посилання")
PROJECT_LINK_RE = re.compile(r'^https?://[^\s()<>"\\]+$')


def audit_entries(problems) -> list:
    """Зауваження для інтерфейсу. id стабільний (вид|розділ|пункт|№), щоб після
    виправлення одного зауваження позначки в інших не з'їжджали."""
    seen, entries = {}, []
    for kind, section, item, text, extra in problems:
        key = f"{kind}|{section}|{item}"
        seen[key] = seen.get(key, 0) + 1
        entries.append({
            "id": f"{key}|{seen[key]}", "kind": kind, "section": section, "item": item, "text": text,
            "pr": (extra or {}).get("pr"),
            "fix": kind in FIXABLE_AUDIT_KINDS and extra is not None, "extra": extra,
        })
    return entries


def run_job(job_id: str, pdf_path: Path, out_dir: Path, acknowledged=frozenset()):
    """acknowledged — id зауважень, які користувач виправив або позначив галочкою.
    Поки є зауваження без позначки, скачування не починається (awaiting_fixes)."""
    job = jobs[job_id]

    def log(msg):
        job["log"].append(str(msg))

    try:
        def on_audit(problems):
            entries = audit_entries(problems)
            pending = [e for e in entries if e["id"] not in acknowledged]
            job["audit"] = entries if pending else []
            if pending:
                job["awaiting_fixes"] = True
                return False

        core.download_agenda_pdfs(pdf_path, out_dir, delay=0.3, log=log, on_audit=on_audit, autofix=True)
    except Exception as exc:
        log(f"[!] Критична помилка: {exc}")
    finally:
        job["done"] = True


def _clean_html_text(fragment: str) -> str:
    text = re.sub(r"<[^>]+>", " ", fragment)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def extract_project_title(page_html: str) -> str | None:
    match = PROJECT_TITLE_RE.search(page_html)
    return _clean_html_text(match.group(1)) if match else None


def extract_project_date(page_html: str) -> str | None:
    match = PROJECT_DATE_RE.search(page_html)
    return match.group(1).strip() if match else None


def extract_speaker_short_name(pdf_path: Path) -> str | None:
    """OCR-ить сторінки PDF (Tesseract, укр. мова) у пошуках абзацу
    "Доповідачем проєкту рішення ..." і повертає ім'я доповідача у форматі
    "Прізвище І.П.". Зупиняється на першій сторінці, де абзац знайдено."""
    tesseract_cmd = find_tesseract_cmd()
    if not tesseract_cmd:
        raise RuntimeError(
            "Tesseract OCR не знайдено на цьому комп'ютері. Встановіть його "
            "(macOS: brew install tesseract tesseract-lang; Windows: "
            "https://github.com/UB-Mannheim/tesseract/wiki) і спробуйте ще раз."
        )

    doc = fitz.open(pdf_path)
    try:
        with tempfile.TemporaryDirectory() as tmp_dir:
            for page_index in range(len(doc)):
                png_path = Path(tmp_dir) / f"page_{page_index}.png"
                pix = doc[page_index].get_pixmap(dpi=OCR_DPI)
                pix.save(png_path)

                try:
                    result = subprocess.run(
                        [tesseract_cmd, str(png_path), "stdout", "-l", OCR_LANG],
                        capture_output=True, text=True, encoding="utf-8", timeout=60,
                        env=_tesseract_env(tesseract_cmd),
                        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                    )
                except OSError as exc:
                    raise RuntimeError(f"не вдалося запустити Tesseract ({exc})") from exc
                finally:
                    png_path.unlink(missing_ok=True)

                for pattern in SPEAKER_BLOCK_PATTERNS:
                    match = pattern.search(result.stdout)
                    if match:
                        surname, first_name, patronymic = match.groups()
                        return f"{surname} {first_name[0]}.{patronymic[0]}."
                for pattern in SPEAKER_NAME_FIRST_PATTERNS:
                    match = pattern.search(result.stdout)
                    if match:
                        first_name, surname = match.groups()
                        return f"{_surname_case(surname)} {first_name[0]}."
    finally:
        doc.close()
    return None


def load_agenda_into_job(job: dict):
    """Читає порядок денний: його заголовки стають деревом розділів завдання, а пункти — готовими
    (заблокованими) картками з уже обраними розділами та порядком."""
    info = job["agenda"]
    model = agenda_doc.load_agenda(info["path"])
    items, headings = model.items(), model.sections()
    if not items:
        raise ValueError("не знайдено жодного пункту виду «1) Про …».")
    position, loaded = collections.Counter(), []
    for item in items:
        sid = item["section"]
        if sid not in headings:  # пункт до першого заголовка розділу
            continue
        loaded.append({"filename": info["name"], "line": item["text"], "url": item["url"], "section": sid,
                       "rank": position[sid], "source": "agenda", "locked": True, "agenda_index": item["index"]})
        position[sid] += 1
    titles = {sid: headings[sid].rstrip(": ").strip() for sid in sorted(headings, key=section_key)}
    job["results"][:0] = loaded
    job["sections"] = {"titles": titles, "locked": list(titles)}
    info["parsed"] = True
    job["log"].append(f"Порядок денний «{info['name']}»: розділів — {len(titles)}, пунктів — {len(loaded)}. "
                      "Нові пункти буде додано в нього на обрані вами місця.")


def build_agenda_doc(job: dict) -> Path:
    """Порядок денний + нові пункти: Word-файл правиться на місці (форматування й посилання старих
    пунктів не чіпаються), з PDF відновлюється Word-документ. Нові пункти стають там, де їх розставив
    користувач; нумерація "N)" у зачеплених розділах перераховується."""
    info, results, titles = job["agenda"], job["results"], job["sections"]["titles"]
    model = agenda_doc.load_agenda(info["path"])
    placements = {}
    for sid in titles:
        entries = []
        for r in section_order(results, sid):
            entries.append(("agenda", r["agenda_index"]) if r.get("locked") else ("new", r["line"].strip(), r["url"]))
        placements[sid] = entries
    agenda_doc.apply_new_items(model, placements, titles)
    out = BASE_DIR / (core.sanitize_filename(Path(info["name"]).stem) + " (з новими пунктами).docx")
    model.save(out)
    return out


def run_generate_job(job_id: str, file_paths: list, skip=frozenset()):
    """Спершу перевіряє назви файлів: поки є некоректні (крім тих, що користувач пропустив),
    обробка не починається — файли можна перейменувати в програмі (awaiting_fixes)."""
    job = jobs[job_id]

    def log(msg):
        job["log"].append(str(msg))

    if job.get("agenda") and not job["agenda"].get("parsed"):
        try:
            load_agenda_into_job(job)
        except Exception as exc:
            log(f"[!] Не вдалося прочитати порядок денний: {exc}")
            job["failed"] = True
            job["done"] = True
            return

    entries = generate_name_entries(file_paths)
    pending = [e for e in entries if e["filename"] not in skip]
    if pending:
        job["audit"] = entries
        job["awaiting_fixes"] = True
        log(f"Перевірка назв файлів: некоректних — {len(pending)} з {len(file_paths)}. Обробку не розпочато: "
            "перейменуйте файли або позначте їх галочкою, щоб пропустити — після цього назви буде перевірено "
            "ще раз і обробка почнеться.")
        job["done"] = True
        return
    for path in [p for p in file_paths if p.name in skip]:
        log(f"    [i] {path.name}: пропущено за вашим рішенням (некоректна назва файлу)")
    file_paths = [p for p in file_paths if p.name not in skip]

    log(f"Знайдено файлів для обробки: {len(file_paths)}")

    session = requests.Session()
    session.headers.update(core.HEADERS)

    try:
        ok, failed = 0, []
        for file_path in file_paths:
            filename = file_path.name
            number, year, name_error = check_decision_filename(filename)
            if name_error:  # мало б бути відсіяно перевіркою вище
                log(f"[!] {filename}: {name_error} (очікується формат НОМЕР-РІК.pdf, напр. 183-25.pdf)")
                failed.append(filename)
                continue
            label = f"номер {number}" + (f", {year} рік" if year else "")
            try:
                project_url, project_html, note = find_project_page(session, number, year)
            except LookupError as exc:
                log(f"[!] {filename} ({label}): {exc}")
                failed.append(filename)
                continue
            except Exception as exc:
                log(f"[!] {filename} ({label}): помилка запиту до пошуку чи сторінки проєкту — {exc}")
                failed.append(filename)
                continue
            if note:
                log(f"    [i] {filename}: {note}")

            title = extract_project_title(project_html)
            date = extract_project_date(project_html)
            if not title or not date:
                missing = ", ".join(
                    n for n, v in (("заголовок", title), ("дата реєстрації", date)) if not v
                )
                log(f"[!] {filename} ({label}): не вдалося знайти на сторінці проєкту: {missing} — {project_url}")
                failed.append(filename)
                continue

            try:
                short_name = extract_speaker_short_name(file_path)
            except Exception as exc:
                log(f"[!] {filename} ({label}): помилка розпізнавання доповідача — {exc}")
                failed.append(filename)
                continue
            if not short_name:
                short_name = NO_SPEAKER_LABEL
                log(f"    [i] {filename}: абзац \"Доповідачем проєкту рішення...\" не знайдено "
                    f"у файлі — використовую \"{NO_SPEAKER_LABEL}\"")

            line = (
                f"\t{title} (від {date} №  {DECISION_PREFIX}-{number}/ПР). "
                f"(доп. {short_name})."
            )
            log(f"[+] {filename} -> {line}")
            job["results"].append({
                "filename": filename, "line": line, "section": None, "url": project_url, "source": "new",
            })
            ok += 1

        log(f"\nГотово: оброблено {ok} з {len(file_paths)}, помилок: {len(failed)}")
    except Exception as exc:
        log(f"[!] Критична помилка: {exc}")
    finally:
        job["done"] = True


def add_hyperlink(paragraph, url: str, text: str):
    """Додає в абзац клікабельний гіперлінк на зовнішній url (python-docx
    не має для цього готового API — стандартний рецепт через oxml).
    Навмисно без кольору/підкреслення — звичайний чорний текст, як у
    справжніх документах КМР."""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    run = OxmlElement("w:r")
    text_el = OxmlElement("w:t")
    text_el.text = text
    run.append(text_el)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    return hyperlink


def section_order(results: list, section: int) -> list:
    """Пункти розділу в порядку документа: спершу ті, що користувач розставив (rank), потім
    решта — у порядку завантаження."""
    items = [(i, r) for i, r in enumerate(results) if r.get("section") == section]
    items.sort(key=lambda t: (t[1].get("rank", 10**9), t[0]))
    return [r for _, r in items]


def build_generated_doc(results: list) -> Path:
    """Будує "Згенерованний ПР.docx" у папці даних застосунку (BASE_DIR): 6 розділів з
    підказки, у кожному — обрані користувачем пункти, де сам текст пункту —
    гіперлінк на сторінку проєкту рішення. Нумерація пунктів у кожному
    розділі завжди починається з 1) (без префіксу номера розділу)."""
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)

    # Розділ, його власні пункти, далі підрозділи (у порядку номерів) — і так до 4-го рівня.
    for section_id in section_ids():
        header = doc.add_paragraph()
        header_run = header.add_run(f"{section_id}) {SECTION_TITLES[section_id]}:")
        header_run.bold = True

        for sub_index, item in enumerate(section_order(results, section_id), start=1):
            title = item["line"].lstrip("\t")
            para = doc.add_paragraph()
            para.paragraph_format.left_indent = Inches(0.4)
            para.add_run(f"{sub_index}) ")
            add_hyperlink(para, item["url"], title)

    doc.save(GENERATED_DOC_PATH)
    return GENERATED_DOC_PATH


@app.route("/")
def index():
    return render_template_string(PAGE, default_out_dir=str(get_work_dir()), icon_sprite=icon_sprite())


def sections_payload(titles: dict, locked=()) -> dict:
    """locked — заголовки з порядку денного: їх не перейменовують і не видаляють."""
    return {"sections": dict(titles), "order": section_ids(titles), "locked": sorted(locked, key=section_key)}


def tree_for(job_id):
    """Дерево розділів: якщо в завданні є порядок денний — його власні заголовки (і підрозділи, додані
    поверх них), інакше загальне, збережене на диску. Повертає (назви, заблоковані id, завдання)."""
    job = jobs.get(job_id) if job_id else None
    if job and job.get("sections"):
        return job["sections"]["titles"], set(job["sections"]["locked"]), job
    return SECTION_TITLES, set(), None


@app.route("/api/sections")
def api_sections_get():
    titles, locked, _ = tree_for(request.args.get("job_id"))
    return jsonify(**sections_payload(titles, locked))


@app.route("/api/sections", methods=["POST"])
def api_sections_set():
    data = request.get_json(silent=True) or {}
    titles, locked, job = tree_for(data.get("job_id"))
    section = str(data.get("section", ""))
    title = (data.get("title") or "").strip()
    if section not in titles:
        return jsonify(error="invalid section"), 400
    if section in locked:
        return jsonify(error="Заголовок із порядку денного змінювати не можна."), 400
    if not title:
        return jsonify(error="Назва розділу не може бути порожньою"), 400

    titles[section] = title
    if job is None:
        save_section_titles(SECTION_TITLES)
    return jsonify(ok=True, **sections_payload(titles, locked))


@app.route("/api/sections/add", methods=["POST"])
def api_sections_add():
    """Додає підрозділ до розділу: 1 -> 1.1, 1.1 -> 1.1.1 … (до 4 рівнів)."""
    data = request.get_json(silent=True) or {}
    titles, locked, job = tree_for(data.get("job_id"))
    parent = str(data.get("parent", ""))
    if parent not in titles:
        return jsonify(error="invalid section"), 400
    if section_depth(parent) >= MAX_SECTION_DEPTH:
        return jsonify(error=f"Глибше за {MAX_SECTION_DEPTH} рівні підрозділи створювати не можна."), 400
    siblings = [int(k.rsplit(".", 1)[1]) for k in titles
                if k.startswith(parent + ".") and section_depth(k) == section_depth(parent) + 1]
    new_id = f"{parent}.{max(siblings, default=0) + 1}"
    titles[new_id] = NEW_SUBSECTION_TITLE
    if job is None:
        save_section_titles(SECTION_TITLES)
    return jsonify(ok=True, id=new_id, **sections_payload(titles, locked))


@app.route("/api/sections/delete", methods=["POST"])
def api_sections_delete():
    """Видаляє підрозділ разом із вкладеними. Картки з них повертаються в батьківський розділ,
    номери наступних підрозділів зсуваються (1.3 -> 1.2), а розділи в списках результатів
    оновлюються відповідно. Заголовки з порядку денного й основні розділи 1–6 видаляти не можна."""
    data = request.get_json(silent=True) or {}
    titles, locked, job = tree_for(data.get("job_id"))
    section = str(data.get("section", ""))
    if section not in titles:
        return jsonify(error="invalid section"), 400
    if section in locked:
        return jsonify(error="Заголовок із порядку денного видаляти не можна."), 400
    if section_depth(section) == 1:
        return jsonify(error="Основні розділи 1–6 видаляти не можна."), 400
    new_titles, mapping, removed = plan_section_delete(titles, section)
    titles.clear()
    titles.update(new_titles)
    if job is None:
        save_section_titles(SECTION_TITLES)
    targets = [job] if job is not None else [j for j in list(jobs.values()) if not j.get("sections")]
    for target in targets:
        for result in target.get("results", []):
            current = result.get("section")
            if current in mapping:
                result["section"] = mapping[current]
                if current in removed:
                    result.pop("rank", None)  # у батьківському розділі — в кінець
    return jsonify(ok=True, **sections_payload(titles, locked))


@app.route("/upload", methods=["POST"])
def upload():
    uploaded = request.files.get("pdf")
    if not uploaded or not uploaded.filename:
        return jsonify(error="Файл не обрано"), 400

    try:
        out_dir = validate_work_dir(request.form.get("out_dir", "").strip() or str(get_work_dir()))
    except ValueError as exc:
        return jsonify(error=str(exc)), 400
    save_work_dir(out_dir)  # що використали — те й запам'ятали

    # Сам файл порядку денного лежить у корені папки результатів — його видно в дереві.
    job_id = uuid.uuid4().hex
    pdf_path = out_dir / Path(uploaded.filename).name

    # Той самий порядок денний вдруге просто замінює файл у папці. Але не під час його обробки.
    if any(not j["done"] and j.get("pdf_path") == str(pdf_path) for j in list(jobs.values())):
        return jsonify(error="Цей порядок денний зараз обробляється — дочекайтеся завершення й завантажте ще раз."), 400
    replaced = pdf_path.exists()
    try:
        uploaded.save(pdf_path)
    except OSError as exc:
        return jsonify(error=f"Не вдалося записати файл порядку денного (можливо, він відкритий в іншій програмі — "
                             f"закрийте його й спробуйте ще раз): {exc}"), 400

    set_current_agenda(agenda_title(uploaded.filename), out_dir)
    log = (["Файл порядку денного вже був у папці — його замінено новою версією. Раніше скачані документи "
            "не перезавантажуються (наявні файли пропускаються), звіти збираються заново."] if replaced else [])
    jobs[job_id] = {"log": log, "done": False, "out_dir": str(out_dir), "audit": [], "pdf_path": str(pdf_path)}
    thread = threading.Thread(target=run_job, args=(job_id, pdf_path, out_dir), daemon=True)
    thread.start()

    return jsonify(job_id=job_id)


@app.route("/generate/upload", methods=["POST"])
def generate_upload():
    uploaded_files = [f for f in request.files.getlist("files") if f.filename]
    if not uploaded_files:
        return jsonify(error="Файли не обрано"), 400

    # Файли можна додавати порціями (навіть по одному): нова порція дописується до вже
    # оброблених результатів попередньої (prev_job), і з усіх них збирається один документ.
    prev = jobs.get(request.form.get("prev_job", ""))
    results = [dict(r) for r in prev.get("results", [])] if prev else []
    already = {r["filename"] for r in results}
    log, fresh = [], []
    for uploaded in uploaded_files:
        name = Path(uploaded.filename).name
        if name in already:
            log.append(f"    [i] {name}: уже оброблено раніше в цьому списку — пропущено")
        else:
            fresh.append((uploaded, name))
    if not fresh:
        return jsonify(error="Усі ці файли вже оброблено раніше."), 400

    job_id = uuid.uuid4().hex
    job_dir = GENERATE_UPLOADS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    saved = []
    for uploaded, name in fresh:
        dest = job_dir / name
        uploaded.save(dest)
        saved.append(dest)

    # Порядок денний серед файлів: Word (.docx) або PDF з посиланнями на проєкти, назва якого не схожа
    # на "НОМЕР-РІК". Він один; його пункти стають готовими картками, а нові вставляються в нього.
    agendas = [p for p in saved if p.suffix.lower() == ".docx" or (
        p.suffix.lower() == ".pdf" and check_decision_filename(p.name)[2] and agenda_doc.looks_like_agenda_pdf(p))]
    error = None
    if len(agendas) > 1:
        error = "Можна додати лише один порядок денний за раз."
    elif agendas and prev and prev.get("agenda"):
        error = "Порядок денний уже завантажено. Щоб взяти інший, натисніть «Очистити список і почати заново»."
    elif agendas and results:
        error = ("Порядок денний додається разом із першими файлами. Натисніть «Очистити список і почати заново» "
                 "й завантажте його разом з усіма файлами.")
    if error:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify(error=error), 400
    agenda_path = agendas[0] if agendas else None
    file_paths = sorted((p for p in saved if p != agenda_path), key=lambda p: p.name)

    job = {"log": log, "done": False, "out_dir": str(job_dir), "results": results}
    if prev and prev.get("agenda"):  # порядок денний і його розділи переходять у наступні порції
        job["agenda"], job["sections"] = prev["agenda"], prev["sections"]
    if agenda_path:
        job["agenda"] = {"path": str(agenda_path), "name": agenda_path.name, "kind": agenda_path.suffix.lower().lstrip("."),
                         "parsed": False}
    jobs[job_id] = job
    thread = threading.Thread(target=run_generate_job, args=(job_id, file_paths), daemon=True)
    thread.start()

    return jsonify(job_id=job_id)


@app.route("/api/log/<job_id>")
def api_log(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify(error="unknown job"), 404
    extra = {}
    if job.get("agenda"):
        extra["agenda"] = {"name": job["agenda"]["name"], "kind": job["agenda"]["kind"],
                           "items": sum(1 for r in job.get("results", []) if r.get("locked"))}
        if job.get("sections"):
            extra["sections"] = sections_payload(job["sections"]["titles"], job["sections"]["locked"])
    return jsonify(log=job["log"], done=job["done"], results=job.get("results", []), audit=job.get("audit", []),
                   awaiting=job.get("awaiting_fixes", False), failed=job.get("failed", False), **extra)


def _generate_job_files(folder: Path) -> list:
    return sorted(p for p in folder.iterdir() if p.is_file())


@app.route("/api/generate/rename", methods=["POST"])
def api_generate_rename():
    """Перейменовує файл (його копію в робочій папці програми — оригінал не чіпає) і
    перевіряє назви ще раз."""
    data = request.get_json(silent=True) or {}
    job = jobs.get(data.get("job_id"))
    if not job or not job.get("awaiting_fixes") or not job["done"] or job.get("proceeded"):
        return jsonify(error="Немає чого перейменовувати — завантажте файли ще раз."), 400
    entry = next((e for e in job["audit"] if e["id"] == data.get("audit_id")), None)
    folder = Path(job["out_dir"])
    if entry is None or not (folder / entry["filename"]).is_file():
        return jsonify(error="Файл не знайдено — можливо, його вже видалили."), 400
    try:
        new_name = validate_new_filename(str(data.get("new_name", "")), folder, entry["filename"])
        (folder / entry["filename"]).rename(folder / new_name)
    except ValueError as exc:
        return jsonify(error=str(exc)), 400
    except OSError as exc:
        return jsonify(error=f"Не вдалося перейменувати файл: {exc}"), 400
    job["audit"] = generate_name_entries(_generate_job_files(folder))
    job["log"].append(f"[~] Файл перейменовано: {entry['filename']} → {new_name}")
    return jsonify(ok=True, new_name=new_name, audit=job["audit"])


@app.route("/api/generate/proceed", methods=["POST"])
def api_generate_proceed():
    """Усі зауваження до назв опрацьовано (перейменовано чи пропущено галочкою): назви
    перевіряються ще раз, і якщо все гаразд — починається обробка."""
    data = request.get_json(silent=True) or {}
    old = jobs.get(data.get("job_id"))
    if not old or not old.get("awaiting_fixes") or not old["done"]:
        return jsonify(error="Немає чого продовжувати."), 400
    if old.get("proceeded"):
        return jsonify(error="Обробку вже розпочато."), 400
    folder = Path(old["out_dir"])
    files = _generate_job_files(folder)
    entries = generate_name_entries(files)
    acknowledged = {str(i)[len("name|"):] for i in data.get("acknowledged", []) if str(i).startswith("name|")}
    pending = [e for e in entries if e["filename"] not in acknowledged]
    old["audit"] = entries
    if pending:
        old["log"].append(f"[~] Повторна перевірка: некоректних назв лишилося — {len(pending)}. Обробку не розпочато.")
        return jsonify(started=False, audit=entries)

    old["proceeded"] = True
    job_id = uuid.uuid4().hex
    jobs[job_id] = {
        "log": list(old["log"]) + ["Назви файлів перевірено — починаємо обробку."],
        "done": False, "out_dir": str(folder), "results": [dict(r) for r in old.get("results", [])],
    }
    if old.get("agenda"):
        jobs[job_id]["agenda"], jobs[job_id]["sections"] = old["agenda"], old.get("sections")
    skip = frozenset(e["filename"] for e in entries)
    threading.Thread(target=run_generate_job, args=(job_id, files, skip), daemon=True).start()
    return jsonify(started=True, job_id=job_id)


@app.route("/api/generate/select", methods=["POST"])
def api_generate_select():
    data = request.get_json(silent=True) or {}
    job = jobs.get(data.get("job_id"))
    if not job:
        return jsonify(error="unknown job"), 404

    results = job.get("results", [])
    index = data.get("index")
    section = data.get("section")
    if not isinstance(index, int) or not (0 <= index < len(results)):
        return jsonify(error="invalid index"), 400
    if section is not None and section not in tree_for(data.get("job_id"))[0]:
        return jsonify(error="invalid section"), 400
    if results[index].get("locked"):
        return jsonify(error="Пункт із порядку денного: його розділ змінювати не можна."), 400

    if results[index].get("section") != section:
        results[index].pop("rank", None)  # у новому розділі — в кінець, поки користувач не переставить
    results[index]["section"] = section
    return jsonify(ok=True)


@app.route("/api/generate/arrange", methods=["POST"])
def api_generate_arrange():
    """Ставить пункт на потрібне місце в розділі (можливо, переносячи з іншого розділу).
    position — номер місця серед пунктів розділу БЕЗ самого пункту (0 — першим). Картки в
    списку результатів не рухаються: порядок у розділі зберігається окремо (rank)."""
    data = request.get_json(silent=True) or {}
    job = jobs.get(data.get("job_id"))
    if not job:
        return jsonify(error="unknown job"), 404
    if not job["done"]:
        return jsonify(error="Зачекайте, поки завершиться обробка файлів."), 400
    results = job.get("results", [])
    index, section, position = data.get("index"), data.get("section"), data.get("position")
    if (not isinstance(index, int) or not (0 <= index < len(results))
            or section not in tree_for(data.get("job_id"))[0] or not isinstance(position, int) or position < 0):
        return jsonify(error="invalid arrange"), 400
    if results[index].get("locked"):
        return jsonify(error="Пункти з порядку денного стоять на своїх місцях — переставляти можна лише нові."), 400

    item = results[index]
    others = [r for r in section_order(results, section) if r is not item]
    others.insert(min(position, len(others)), item)
    item["section"] = section
    for rank, r in enumerate(others):
        r["rank"] = rank
    return jsonify(ok=True, results=results)


@app.route("/api/generate/build", methods=["POST"])
def api_generate_build():
    data = request.get_json(silent=True) or {}
    job = jobs.get(data.get("job_id"))
    if not job:
        return jsonify(error="unknown job"), 404

    results = job.get("results", [])
    if not results:
        return jsonify(error="Немає жодного опрацьованого пункту"), 400

    missing = [i + 1 for i, r in enumerate(results) if r.get("section") is None]
    if missing:
        return jsonify(error="not_all_selected", missing=missing), 400

    try:
        doc_path = build_agenda_doc(job) if job.get("agenda") else build_generated_doc(results)
    except Exception as exc:
        return jsonify(error=f"Не вдалося створити документ: {exc}"), 500

    open_in_os(str(doc_path))
    return jsonify(ok=True, path=str(doc_path))


@app.route("/api/settings", methods=["GET", "POST"])
def api_settings():
    """Робоча папка. Після зміни дерево показує вміст нової папки (зі своїм порядком денним)."""
    if request.method == "POST":
        try:
            folder = validate_work_dir(str((request.get_json(silent=True) or {}).get("work_dir", "")))
        except ValueError as exc:
            return jsonify(error=str(exc)), 400
        save_work_dir(folder)
        if current_agenda["out_dir"] != str(folder):
            set_current_agenda(detect_agenda(folder), folder)
    return jsonify(work_dir=str(get_work_dir()))


@app.route("/api/tree")
def api_tree():
    name, out_dir = current_agenda["name"], current_agenda["out_dir"]
    if not name:
        return jsonify(name=None, children=[])
    root = Path(out_dir)
    children = build_tree(root, root, [TREE_MAX_NODES]) if root.is_dir() else []
    return jsonify(name=name, children=children)


def _tree_target():
    """Шлях із запиту, гарантовано всередині папки результатів (або None)."""
    if not current_agenda["out_dir"]:
        return None
    root = Path(current_agenda["out_dir"]).resolve()
    rel = (request.get_json(silent=True) or {}).get("path", "")
    target = (root / rel).resolve()
    return target if target.is_relative_to(root) and target.exists() else None


@app.route("/api/tree/open", methods=["POST"])
def api_tree_open():
    target = _tree_target()
    if target is None:
        return jsonify(error="Файл або папку не знайдено — можливо, її видалили чи перемістили."), 400
    open_in_os(str(target))
    return jsonify(ok=True)


@app.route("/api/tree/delete", methods=["POST"])
def api_tree_delete():
    """Переміщує файл чи папку з дерева в кошик. Кореневу папку (усі результати,
    можливо й інших порядків денних) видаляти не дозволяємо."""
    target = _tree_target()
    if target is None:
        return jsonify(error="Файл або папку не знайдено — можливо, її вже видалили чи перемістили."), 400
    root = Path(current_agenda["out_dir"]).resolve()
    if target == root:
        return jsonify(error="Кореневу папку видаляти не можна — видаліть потрібні файли чи підпапки."), 400
    if any(not j["done"] and j.get("out_dir") and Path(j["out_dir"]).resolve() == root for j in list(jobs.values())):
        return jsonify(error="Зачекайте, поки завершиться обробка порядку денного, і спробуйте ще раз."), 400
    try:
        move_to_trash(target)
    except Exception as exc:
        return jsonify(error=f"Не вдалося видалити: {exc}"), 400
    return jsonify(ok=True)


@app.route("/api/tree/reveal", methods=["POST"])
def api_tree_reveal():
    target = _tree_target()
    if target is None:
        return jsonify(error="Файл або папку не знайдено — можливо, її видалили чи перемістили."), 400
    try:
        reveal_in_os(target)
    except (OSError, subprocess.CalledProcessError) as exc:
        return jsonify(error=f"Не вдалося відкрити Провідник: {exc}. Шлях: {target}"), 400
    return jsonify(ok=True, path=str(target))


def _clean_contact(data: dict):
    name = re.sub(r"\s+", " ", str(data.get("name", ""))).strip()
    email = str(data.get("email", "")).strip()
    if not name:
        raise ValueError("Вкажіть ім’я.")
    if not EMAIL_RE.match(email):
        raise ValueError("Невірна адреса пошти.")
    return name, email


@app.route("/api/contacts", methods=["GET", "POST"])
def api_contacts():
    contacts = _read_json(CONTACTS_PATH, [])
    if request.method == "GET":
        return jsonify(contacts=contacts)
    try:
        name, email = _clean_contact(request.get_json(silent=True) or {})
    except ValueError as exc:
        return jsonify(error=str(exc)), 400
    contact = {"id": uuid.uuid4().hex[:8], "name": name, "email": email}
    contacts.append(contact)
    _write_json(CONTACTS_PATH, contacts)
    return jsonify(contact=contact)


@app.route("/api/contacts/import", methods=["POST"])
def api_contacts_import():
    upload = request.files.get("file")
    if not upload or not upload.filename:
        return jsonify(error="Оберіть файл Excel (.xlsx) або CSV."), 400
    data = upload.read(MAX_IMPORT_BYTES + 1)
    if len(data) > MAX_IMPORT_BYTES:
        return jsonify(error="Файл завеликий (понад 5 МБ)."), 400
    try:
        found, invalid = parse_contacts_rows(read_contacts_table(upload.filename, data))
    except ValueError as exc:
        return jsonify(error=str(exc)), 400
    if not found:
        return jsonify(error="У файлі не знайдено адрес пошти. Потрібна колонка з email, бажано з колонкою імені."), 400

    contacts = _read_json(CONTACTS_PATH, [])
    known = {c["email"].casefold() for c in contacts}
    added = duplicates = 0
    for name, email in found:
        if email.casefold() in known:
            duplicates += 1
            continue
        known.add(email.casefold())
        contacts.append({"id": uuid.uuid4().hex[:8], "name": name, "email": email})
        added += 1
    if added:
        _write_json(CONTACTS_PATH, contacts)
    return jsonify(added=added, duplicates=duplicates, invalid=invalid, contacts=contacts)


@app.route("/api/contacts/<cid>", methods=["PUT"])
def api_contact_update(cid):
    try:
        name, email = _clean_contact(request.get_json(silent=True) or {})
    except ValueError as exc:
        return jsonify(error=str(exc)), 400
    contacts = _read_json(CONTACTS_PATH, [])
    contact = next((c for c in contacts if c["id"] == cid), None)
    if contact is None:
        return jsonify(error="Контакт не знайдено."), 404
    contact.update(name=name, email=email)
    _write_json(CONTACTS_PATH, contacts)
    return jsonify(contact=contact)


@app.route("/api/contacts/<cid>", methods=["DELETE"])
def api_contact_delete(cid):
    contacts = [c for c in _read_json(CONTACTS_PATH, []) if c["id"] != cid]
    _write_json(CONTACTS_PATH, contacts)
    groups = _read_json(GROUPS_PATH, [])
    for group in groups:
        group["contact_ids"] = [i for i in group["contact_ids"] if i != cid]
    _write_json(GROUPS_PATH, groups)
    return jsonify(ok=True)


def _clean_group(data: dict, groups: list, own_id=None):
    """Назва (унікальна, без урахування регістру) і id наявних контактів без повторів."""
    name = re.sub(r"\s+", " ", str(data.get("name", ""))).strip()
    if not name:
        raise ValueError("Вкажіть назву групи.")
    if len(name) > 60:
        raise ValueError("Назва групи задовга (до 60 символів).")
    if any(g["name"].casefold() == name.casefold() and g["id"] != own_id for g in groups):
        raise ValueError("Група з такою назвою вже є.")
    known = {c["id"] for c in _read_json(CONTACTS_PATH, [])}
    ids = list(dict.fromkeys(i for i in data.get("contact_ids", []) if i in known))
    return name, ids


@app.route("/api/groups", methods=["GET", "POST"])
def api_groups():
    groups = _read_json(GROUPS_PATH, [])
    if request.method == "GET":
        return jsonify(groups=groups)
    try:
        name, ids = _clean_group(request.get_json(silent=True) or {}, groups)
    except ValueError as exc:
        return jsonify(error=str(exc)), 400
    group = {"id": uuid.uuid4().hex[:8], "name": name, "contact_ids": ids}
    groups.append(group)
    _write_json(GROUPS_PATH, groups)
    return jsonify(group=group)


@app.route("/api/groups/<gid>", methods=["PUT", "DELETE"])
def api_group_item(gid):
    groups = _read_json(GROUPS_PATH, [])
    group = next((g for g in groups if g["id"] == gid), None)
    if group is None:
        return jsonify(error="Групу не знайдено."), 404
    if request.method == "DELETE":
        _write_json(GROUPS_PATH, [g for g in groups if g["id"] != gid])
        return jsonify(ok=True)
    try:
        name, ids = _clean_group(request.get_json(silent=True) or {}, groups, own_id=gid)
    except ValueError as exc:
        return jsonify(error=str(exc)), 400
    group.update(name=name, contact_ids=ids)
    _write_json(GROUPS_PATH, groups)
    return jsonify(group=group)


@app.route("/api/mail/settings", methods=["GET", "POST"])
def api_mail_settings():
    cfg = _read_json(MAIL_SETTINGS_PATH, {})
    if request.method == "POST":
        data = request.get_json(silent=True) or {}
        method = data.get("method") or _mail_method(cfg)
        if method not in ("outlook", "smtp"):
            return jsonify(error="Невідомий спосіб відправки."), 400
        if method == "outlook" and not outlook_available():
            return jsonify(error="Outlook доступний лише у Windows із встановленим класичним Outlook."), 400
        cfg["method"] = method
        if method == "smtp":
            try:
                port = int(str(data.get("port") or 465).strip())
            except ValueError:
                return jsonify(error="Порт має бути числом."), 400
            clean = lambda key: re.sub(r"[\r\n]+", " ", str(data.get(key, ""))).strip()
            cfg.update(
                host=clean("host"), port=port, username=clean("username"),
                # порожнє поле пароля = лишити збережений
                password=data.get("password") or cfg.get("password", ""),
                from_name=clean("from_name"),
            )
        _write_json(MAIL_SETTINGS_PATH, cfg, private=True)
    method = _mail_method(cfg)
    smtp_ready = bool(cfg.get("host") and cfg.get("username") and cfg.get("password"))
    return jsonify(
        method=method, outlook_available=outlook_available(),
        host=cfg.get("host", ""), port=cfg.get("port", 465), username=cfg.get("username", ""),
        from_name=cfg.get("from_name", ""), has_password=bool(cfg.get("password")),
        configured=outlook_available() if method == "outlook" else smtp_ready,
    )


@app.route("/api/tree/send", methods=["POST"])
def api_tree_send():
    target = _tree_target()
    if target is None:
        return jsonify(error="Файл або папку не знайдено."), 400
    ids = set((request.get_json(silent=True) or {}).get("contact_ids") or [])
    recipients = [c for c in _read_json(CONTACTS_PATH, []) if c["id"] in ids]
    if not recipients:
        return jsonify(error="Оберіть хоча б одного отримувача."), 400
    root = Path(current_agenda["out_dir"]).resolve()
    agenda = current_agenda["name"]
    item = agenda if target == root else target.name
    try:
        attachment = build_attachment(target, root, agenda)
        send_mail(recipients, subject=f"{agenda} — {item}" if target != root else agenda,
                  item_name=item, attachment=attachment)
    except MailError as exc:
        return jsonify(error=str(exc)), 400
    return jsonify(ok=True, sent=len(recipients))


@app.route("/api/audit/fix", methods=["POST"])
def api_audit_fix():
    """Прописує вставлене посилання в файл порядку денного (для зауважень
    "немає посилання" і "однакове посилання") і перераховує зауваження."""
    data = request.get_json(silent=True) or {}
    job = jobs.get(data.get("job_id"))
    if not job or "pdf_path" not in job:
        return jsonify(error="Невідоме завдання — завантажте порядок денний ще раз."), 404
    if not job["done"]:
        return jsonify(error="Дочекайтеся завершення обробки, потім виправляйте."), 400
    entry = next((e for e in job["audit"] if e["id"] == data.get("audit_id")), None)
    if entry is None or not entry["fix"]:
        return jsonify(error="Це зауваження не можна виправити з програми."), 400

    url = str(data.get("url", "")).strip()
    if not PROJECT_LINK_RE.match(url) or not core.PROJECT_URL_RE.search(url):
        return jsonify(error="Це не схоже на посилання на проєкт рішення — очікую адресу виду "
                             "https://kmr.gov.ua/uk/content/proekt-rishennya-kyyivskoyi-miskoyi-rady-…"), 400

    extra = entry["extra"]
    if "items" in extra:
        idx = data.get("item_index")
        if not isinstance(idx, int) or not (0 <= idx < len(extra["items"])):
            return jsonify(error="Оберіть пункт, у якому змінити посилання."), 400
        target = extra["items"][idx]
        where = f"розділ {target['section']}, пункт {target['num']}" + (f" · {target['pr']}" if target.get("pr") else "")
        pos = target["pos"]
    else:
        where = f"розділ {entry['section']}, пункт {entry['item']}" + (f" · {entry['pr']}" if entry.get("pr") else "")
        pos = extra["pos"]

    pdf_path = Path(job["pdf_path"])
    if not pdf_path.is_file():
        return jsonify(error="Файл порядку денного не знайдено — його видалили чи перемістили."), 400
    try:
        core.set_item_link(pdf_path, pos, url)
        # Файл змінено — перевіряємо його заново: зауваження має зникнути, а посилання
        # пункту — нормально розпізнаватися (дата й номер рішення поруч із ним).
        entries = audit_entries(core.audit_agenda(pdf_path))
        _, unresolved = core.extract_agenda_items(pdf_path)
    except (ValueError, RuntimeError, OSError) as exc:
        return jsonify(error=f"Не вдалося змінити файл: {exc}"), 400
    job["audit"] = entries
    job["log"].append(f"[~] Посилання збережено у файлі порядку денного: {where} → {url}")

    resolved = all(e["id"] != entry["id"] for e in entries)
    section_label, item_num = (target["section"], target["num"]) if "items" in extra else (entry["section"], entry["item"])
    unreadable = any(u[2] == item_num and u[1] and u[1][-1] == section_label for u in unresolved)
    warning = None
    if unreadable:
        warning = (f"{where}: посилання збережено, але програма не змогла розпізнати дату й номер рішення "
                   "поруч із ним — перевірте текст пункту в PDF.")
        job["log"].append("[~] " + warning)
    if resolved:
        job["log"].append(f"[~] Перевірено заново: зауваження прибрано зі списку ({where}).")
    return jsonify(ok=True, resolved=resolved, warning=warning, audit=entries,
                   message=None if resolved else "Посилання збережено, але зауваження лишилося — перевірте адресу.")


@app.route("/api/proceed", methods=["POST"])
def api_proceed():
    """Користувач опрацював усі зауваження (виправив або позначив галочкою): файл
    перевіряється ще раз, і якщо без зауважень (крім позначених) — починається
    завантаження. Інакше повертає оновлений список зауважень."""
    data = request.get_json(silent=True) or {}
    old = jobs.get(data.get("job_id"))
    if not old or not old.get("awaiting_fixes") or not old["done"]:
        return jsonify(error="Немає чого продовжувати."), 400
    if old.get("proceeded"):
        return jsonify(error="Завантаження вже розпочато."), 400
    pdf_path, out_dir = Path(old["pdf_path"]), Path(old["out_dir"])
    if not pdf_path.is_file():
        return jsonify(error="Файл порядку денного не знайдено — його видалили чи перемістили."), 400
    acknowledged = {str(x) for x in data.get("acknowledged", [])}
    try:
        entries = audit_entries(core.audit_agenda(pdf_path))
    except (ValueError, RuntimeError, OSError) as exc:
        return jsonify(error=f"Не вдалося перевірити файл: {exc}"), 400

    old["audit"] = entries
    pending = [e for e in entries if e["id"] not in acknowledged]
    if pending:
        old["log"].append(f"[~] Повторна перевірка: лишилося зауважень — {len(pending)}. Завантаження не розпочато.")
        return jsonify(started=False, audit=entries)

    old["proceeded"] = True
    job_id = uuid.uuid4().hex
    jobs[job_id] = {
        "log": list(old["log"]) + ["Усі зауваження опрацьовано, файл перевірено ще раз — починаємо завантаження."],
        "done": False, "out_dir": str(out_dir), "audit": [], "pdf_path": str(pdf_path),
    }
    threading.Thread(target=run_job, args=(job_id, pdf_path, out_dir, frozenset(acknowledged)), daemon=True).start()
    return jsonify(started=True, job_id=job_id)


@app.route("/api/open/<job_id>", methods=["POST"])
def api_open(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify(error="unknown job"), 404
    open_in_os(job["out_dir"])
    return jsonify(ok=True)


class DesktopApi:
    """Методи, які сторінка може викликати у вікні програми (window.pywebview.api)."""

    def pick_folder(self, start=""):
        import webview
        kind = webview.FileDialog.FOLDER if hasattr(webview, "FileDialog") else webview.FOLDER_DIALOG
        start_dir = start if start and Path(start).is_dir() else ""
        chosen = webview.windows[0].create_file_dialog(kind, directory=start_dir)
        return chosen[0] if chosen else None


def main():
    """Запускає Flask на вільному порту і показує інтерфейс у власному вікні
    (pywebview). Якщо pywebview недоступний — відкриває браузер."""
    from werkzeug.serving import make_server

    server = make_server("127.0.0.1", 0, app, threaded=True)
    url = f"http://127.0.0.1:{server.server_port}/"
    threading.Thread(target=server.serve_forever, daemon=True).start()
    try:
        import webview
    except ImportError:
        import webbrowser
        webbrowser.open(url)
        print(f"Відкрито в браузері: {url}  (Ctrl+C — вихід)")
        threading.Event().wait()
    else:
        webview.create_window("Порядок денний", url, width=1200, height=850, js_api=DesktopApi())
        webview.start()
    server.shutdown()


if __name__ == "__main__":
    main()
