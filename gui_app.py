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
import json
import os
import platform
import re
import shutil
import subprocess
import tempfile
import threading
import uuid
from pathlib import Path

import fitz  # PyMuPDF
import requests
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from flask import Flask, jsonify, render_template_string, request

import download_pleanary_pdf as core

BASE_DIR = Path(__file__).resolve().parent
UPLOADS_DIR = BASE_DIR / "uploaded_agendas"
DEFAULT_OUT_DIR = BASE_DIR / "downloaded_plenary_pdfs"
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


def load_section_titles() -> dict:
    if SECTION_TITLES_PATH.exists():
        try:
            saved = json.loads(SECTION_TITLES_PATH.read_text(encoding="utf-8"))
            return {n: saved.get(str(n), DEFAULT_SECTION_TITLES[n]) for n in range(1, 7)}
        except (json.JSONDecodeError, OSError):
            pass
    return dict(DEFAULT_SECTION_TITLES)


def save_section_titles(titles: dict):
    SECTION_TITLES_PATH.write_text(
        json.dumps({str(n): titles[n] for n in range(1, 7)}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


SECTION_TITLES = load_section_titles()

DECISION_NUMBER_RE = re.compile(r"^(\d+)-")
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
OCR_LANG = "ukr"
OCR_DPI = 300

def _tesseract_candidates():
    """tesseract не завжди потрапляє в PATH після встановлення, а на Windows
    типове місце інсталяції ще й залежить від того, чи були права
    адміністратора (Program Files) чи ні (LOCALAPPDATA). Перевіряємо все."""
    candidates = ["tesseract"]
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
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: var(--bg);
    color: var(--text);
    margin: 0;
    padding: 48px 16px;
    display: flex;
    justify-content: center;
    -webkit-font-smoothing: antialiased;
  }
  .shell { width: 100%; max-width: 640px; transition: max-width .2s ease; }
  .shell.wide { max-width: 1040px; }
  .head { text-align: center; margin-bottom: 28px; }
  .head .icon { font-size: 34px; line-height: 1; margin-bottom: 10px; }
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
  .dropzone .drop-icon { font-size: 28px; margin-bottom: 8px; }
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
  .file-chip .fi { font-size: 20px; }
  .file-chip .fname { font-size: 13.5px; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .file-chip .fsize { font-size: 12px; color: var(--text-muted); }
  .file-chip .fclear {
    margin-left: auto; cursor: pointer; color: var(--text-muted);
    border: none; background: none; font-size: 16px; padding: 4px 8px; border-radius: 8px;
  }
  .file-chip .fclear:hover { background: var(--bg); color: var(--text); }

  label.field-label { display: block; font-size: 12.5px; font-weight: 600; color: var(--text-muted); margin: 18px 0 6px; text-transform: uppercase; letter-spacing: .03em; }
  .field-wrap { position: relative; }
  .field-wrap .fi-prefix { position: absolute; left: 12px; top: 50%; transform: translateY(-50%); font-size: 14px; opacity: .7; }
  #outdir {
    width: 100%; padding: 10px 12px 10px 34px; border-radius: 10px;
    border: 1px solid var(--border); background: var(--bg); color: var(--text);
    font-family: Menlo, Consolas, monospace; font-size: 12.5px;
  }
  #outdir:focus { outline: none; border-color: var(--primary); }

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
  .audit-item {
    display: flex; gap: 10px; align-items: flex-start; padding: 9px 12px; font-size: 12.5px;
    line-height: 1.45; cursor: pointer; border-bottom: 1px solid rgba(0,0,0,.06);
  }
  .audit-item:last-child { border-bottom: 0; }
  .audit-item input { margin-top: 3px; flex: none; cursor: pointer; }
  .audit-item .a-kind {
    display: inline-block; font-size: 10.5px; font-weight: 700; text-transform: uppercase; letter-spacing: .03em;
    color: var(--amber); margin-right: 6px;
  }
  .audit-item .a-where { font-weight: 600; margin-right: 6px; }
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
  .r-sections { display: flex; flex-wrap: wrap; gap: 6px; }
  .section-btn {
    border: 1px solid var(--border); background: transparent; color: var(--text-muted);
    font-size: 12px; font-weight: 600; padding: 6px 10px; border-radius: 8px;
    cursor: pointer; transition: background .15s, color .15s, border-color .15s;
  }
  .section-btn:hover { border-color: var(--primary); color: var(--primary); }
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
  .reminder-panel li .s-text { flex: 1; }
  .reminder-panel li .s-edit-btn {
    flex-shrink: 0; border: none; background: none; cursor: pointer;
    font-size: 12px; opacity: .5; padding: 0 2px; line-height: 1.6; transition: opacity .15s;
  }
  .reminder-panel li .s-edit-btn:hover { opacity: 1; }
  .reminder-panel li .s-input {
    flex: 1; font: inherit; padding: 2px 6px; border: 1px solid var(--primary);
    border-radius: 5px; background: var(--card); color: var(--text);
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
<div class="shell">

  <div class="head">
    <div class="icon">🏛️</div>
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
      <div class="drop-icon">⬆️</div>
      <div class="drop-main">Перетягніть сюди PDF порядку денного</div>
      <div class="drop-sub">або натисніть, щоб обрати файл</div>
    </div>
    <input type="file" id="file" accept="application/pdf">

    <div class="file-chip" id="fileChip">
      <div class="fi">📎</div>
      <div style="min-width:0;">
        <div class="fname" id="fname"></div>
        <div class="fsize" id="fsize"></div>
      </div>
      <button class="fclear" id="fclear" title="Прибрати файл">✕</button>
    </div>

    <label class="field-label" for="outdir">Папка для збереження</label>
    <div class="field-wrap">
      <span class="fi-prefix">📁</span>
      <input type="text" id="outdir" value="{{ default_out_dir }}">
    </div>

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

      <button class="secondary" id="openBtn" style="display:none">Показати папку у Finder</button>

      <details class="log-details audit-details" id="auditDetails" open>
        <summary><span class="chev">▶</span> <span id="auditTitle">Зауваження до PDF</span></summary>
        <div class="audit-list" id="auditList"></div>
      </details>

      <details class="log-details" id="logDetails">
        <summary><span class="chev">▶</span> Детальний журнал</summary>
        <pre class="log-box" id="log"></pre>
      </details>
    </div>
  </div>
  </div>

  <div class="panel" id="panel-generate" hidden>
  <p class="panel-desc">Перетягніть сюди проєкти рішень (PDF, можна одразу кілька сотень) — для кожного файлу скрипт визначить номер рішення за назвою й перевірить його на сайті КМР</p>
  <div class="card">
    <div class="dropzone" id="dropGen">
      <div class="drop-icon">⬆️</div>
      <div class="drop-main">Перетягніть сюди файли проєктів рішень</div>
      <div class="drop-sub">або натисніть, щоб обрати файли (напр. "524-26.pdf") — підтримується масове завантаження</div>
    </div>
    <input type="file" id="fileGen" accept="application/pdf" multiple>

    <div class="file-chip" id="fileChipGen">
      <div class="fi">📎</div>
      <div style="min-width:0;">
        <div class="fname" id="fnameGen"></div>
        <div class="fsize" id="fsizeGen"></div>
      </div>
      <button class="fclear" id="fclearGen" title="Прибрати файли">✕</button>
    </div>

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

      <button class="secondary" id="openBtnGen" style="display:none">Показати папку у Finder</button>

      <details class="log-details" id="logDetailsGen">
        <summary><span class="chev">▶</span> Детальний журнал</summary>
        <pre class="log-box" id="logGen"></pre>
      </details>
    </div>
  </div>

  <div class="results-layout" id="resultsLayout">
    <div class="results-list" id="resultsList"></div>
    <div class="reminder-panel">
      <h3>Розділи порядку денного</h3>
      <ol id="sectionsList"></ol>
    </div>
  </div>

  <button class="primary" id="generateBtn">Згенерувати ПД</button>
  </div>

  <div class="foot">Працює локально · нічого не надсилається за межі цього комп'ютера</div>
</div>

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

startBtn.addEventListener('click', async () => {
  if (!selectedFile) return;
  setRunning(true);
  openBtn.style.display = 'none';
  statsEl.style.display = 'none';
  progressWrap.style.display = 'block';
  logDetails.open = false;
  logEl.textContent = '';
  auditDetails.style.display = 'none';
  auditList.textContent = '';
  auditSignature = '';
  auditFixed.clear();
  barFill.style.width = '0%';
  statusPct.textContent = '';
  statusText.textContent = 'Завантажуємо PDF і читаємо порядок денний…';

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
  poll();
});

function escapeHtml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function renderLog(lines) {
  const html = lines.map(l => {
    const esc = escapeHtml(l);
    if (l.startsWith('[+]')) return '<span class="l-ok">' + esc + '</span>';
    if (l.startsWith('[!]') || l.startsWith('  -')) return '<span class="l-err">' + esc + '</span>';
    if (l.startsWith('Знайдено') || l.startsWith('Готово') || l.startsWith('Не вдалося') || l.startsWith('Перевірка') || l.startsWith('Увага')) return '<span class="l-head">' + esc + '</span>';
    return esc;
  }).join('\\n');
  // Не смикаємо прокрутку, якщо користувач сам піднявся вище по журналу.
  const atBottom = logEl.scrollHeight - logEl.scrollTop - logEl.clientHeight < 40;
  logEl.innerHTML = html;
  if (atBottom) logEl.scrollTop = logEl.scrollHeight;
}

function renderAudit(items) {
  if (!items || !items.length) return;
  auditDetails.style.display = 'block';
  const signature = items.map(i => i.id + i.kind + i.section + i.item).join('|');
  if (signature !== auditSignature) {
    auditSignature = signature;
    auditList.textContent = '';
    items.forEach(i => {
      const row = document.createElement('label');
      row.className = 'audit-item';
      row.dataset.id = i.id;
      const where = 'розділ ' + i.section + (i.item !== null ? ', пункт ' + i.item : '');
      row.innerHTML = '<input type="checkbox">' +
        '<span><span class="a-kind">' + escapeHtml(i.kind) + '</span>' +
        '<span class="a-where">' + escapeHtml(where) + '</span>' +
        '<span class="a-text">' + escapeHtml(i.text) + '</span></span>';
      row.querySelector('input').addEventListener('change', e => {
        if (e.target.checked) auditFixed.add(i.id); else auditFixed.delete(i.id);
        row.classList.toggle('fixed', e.target.checked);
        updateAuditTitle(items.length);
      });
      auditList.appendChild(row);
    });
  }
  updateAuditTitle(items.length);
}

function updateAuditTitle(total) {
  auditTitle.textContent = 'Зауваження до PDF: виправлено ' + auditFixed.size + ' з ' + total;
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
      : 'Формуємо звіти й папку "Для депутатов"…';
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

let selectedFilesGen = [];
let currentJobGen = null;
let pollTimerGen = null;
let renderedResultsCount = 0;

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

// ---- Розділи порядку денного (редаговані назви, зберігаються на диску) ----
const sectionsList = document.getElementById('sectionsList');
let sectionTitles = {};

function renderSectionItem(num, title) {
  const li = document.createElement('li');
  li.dataset.section = num;

  const numEl = document.createElement('span');
  numEl.className = 's-num';
  numEl.textContent = num + '.';
  li.appendChild(numEl);

  const textEl = document.createElement('span');
  textEl.className = 's-text';
  textEl.textContent = title;
  li.appendChild(textEl);

  const editBtn = document.createElement('button');
  editBtn.className = 's-edit-btn';
  editBtn.type = 'button';
  editBtn.title = 'Змінити назву розділу';
  editBtn.textContent = '✏️';
  editBtn.addEventListener('click', () => startEditingSection(li, num));
  li.appendChild(editBtn);

  return li;
}

function startEditingSection(li, num) {
  const textEl = li.querySelector('.s-text');
  const editBtn = li.querySelector('.s-edit-btn');
  const currentTitle = sectionTitles[num] || '';

  const input = document.createElement('input');
  input.className = 's-input';
  input.type = 'text';
  input.value = currentTitle;

  textEl.replaceWith(input);
  editBtn.style.display = 'none';
  input.focus();
  input.select();

  let settled = false;
  const finish = async (save) => {
    if (settled) return;
    settled = true;
    const newTitle = input.value.trim();
    if (save && newTitle && newTitle !== currentTitle) {
      try {
        const resp = await fetch('/api/sections', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ section: num, title: newTitle }),
        });
        const data = await resp.json();
        if (data.error) {
          showToast('Помилка: ' + data.error);
          sectionTitles[num] = currentTitle;
        } else {
          sectionTitles[num] = newTitle;
        }
      } catch (e) {
        showToast('Помилка мережі під час збереження назви розділу');
        sectionTitles[num] = currentTitle;
      }
    }
    const newTextEl = document.createElement('span');
    newTextEl.className = 's-text';
    newTextEl.textContent = sectionTitles[num];
    input.replaceWith(newTextEl);
    editBtn.style.display = '';
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
    const data = await resp.json();
    sectionTitles = data.sections || {};
    sectionsList.innerHTML = '';
    for (let n = 1; n <= 6; n++) {
      sectionsList.appendChild(renderSectionItem(n, sectionTitles[n]));
    }
  } catch (e) { /* панель лишиться порожньою, не критично */ }
}
loadSections();

dropGen.addEventListener('click', () => fileInputGen.click());
dropGen.addEventListener('dragover', e => { e.preventDefault(); dropGen.classList.add('hover'); });
dropGen.addEventListener('dragleave', () => dropGen.classList.remove('hover'));
dropGen.addEventListener('drop', e => {
  e.preventDefault();
  dropGen.classList.remove('hover');
  if (e.dataTransfer.files.length) setFilesGen(e.dataTransfer.files);
});
fileInputGen.addEventListener('change', () => { if (fileInputGen.files.length) setFilesGen(fileInputGen.files); });

fclearGen.addEventListener('click', e => {
  e.stopPropagation();
  selectedFilesGen = [];
  fileInputGen.value = '';
  fileChipGen.style.display = 'none';
  dropGen.style.display = 'block';
  startBtnGen.disabled = true;
});

function setFilesGen(fileList) {
  selectedFilesGen = Array.from(fileList);
  const totalSize = selectedFilesGen.reduce((sum, f) => sum + f.size, 0);
  fnameGen.textContent = selectedFilesGen.length === 1
    ? selectedFilesGen[0].name
    : `${selectedFilesGen.length} файлів обрано`;
  fsizeGen.textContent = humanSize(totalSize);
  dropGen.style.display = 'none';
  fileChipGen.style.display = 'flex';
  startBtnGen.disabled = selectedFilesGen.length === 0;
}

function setRunningGen(running) {
  startBtnGen.disabled = running;
  spinnerGen.style.display = running ? 'inline-block' : 'none';
  startLabelGen.textContent = running ? 'Виконується…' : 'Почати';
}

startBtnGen.addEventListener('click', async () => {
  if (!selectedFilesGen.length) return;
  setRunningGen(true);
  openBtnGen.style.display = 'none';
  statsGen.style.display = 'none';
  progressWrapGen.style.display = 'block';
  logDetailsGen.open = false;
  logGen.textContent = '';
  barFillGen.style.width = '0%';
  statusPctGen.textContent = '';
  statusTextGen.textContent = `Завантажуємо ${selectedFilesGen.length} файлів…`;
  resultsList.innerHTML = '';
  resultsLayout.style.display = 'none';
  generateBtn.style.display = 'none';
  renderedResultsCount = 0;

  const form = new FormData();
  for (const f of selectedFilesGen) form.append('files', f);

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
  currentJobGen = data.job_id;
  pollGen();
});

function renderLogGen(lines) {
  const html = lines.map(l => {
    const esc = escapeHtml(l);
    if (l.startsWith('[+]')) return '<span class="l-ok">' + esc + '</span>';
    if (l.startsWith('[!]')) return '<span class="l-err">' + esc + '</span>';
    if (l.startsWith('Знайдено') || l.startsWith('Готово')) return '<span class="l-head">' + esc + '</span>';
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

function makeSectionButtons(index, selected) {
  const wrap = document.createElement('div');
  wrap.className = 'r-sections';
  for (let s = 1; s <= 6; s++) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'section-btn' + (selected === s ? ' active' : '');
    btn.textContent = s + ' розділ';
    btn.addEventListener('click', () => selectResultSection(index, s, wrap));
    wrap.appendChild(btn);
  }
  return wrap;
}

async function selectResultSection(index, section, wrapEl) {
  wrapEl.querySelectorAll('.section-btn').forEach((b, i) => b.classList.toggle('active', i + 1 === section));
  try {
    await fetch('/api/generate/select', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ job_id: currentJobGen, index, section }),
    });
  } catch (e) { /* наступний poll підтвердить/синхронізує стан */ }
}

function appendResultItem(item, index) {
  const div = document.createElement('div');
  div.className = 'result-item';

  const fileEl = document.createElement('div');
  fileEl.className = 'r-file';
  const indexEl = document.createElement('span');
  indexEl.className = 'r-index';
  indexEl.textContent = String(index + 1);
  fileEl.appendChild(indexEl);
  fileEl.appendChild(document.createTextNode(item.filename));
  div.appendChild(fileEl);

  const textEl = document.createElement('div');
  textEl.className = 'r-text';
  textEl.textContent = item.line;
  div.appendChild(textEl);

  div.appendChild(makeSectionButtons(index, item.section));
  resultsList.appendChild(div);
}

function renderResults(results) {
  if (!results || !results.length) return;
  resultsLayout.style.display = 'flex';
  generateBtn.style.display = 'flex';
  for (let i = renderedResultsCount; i < results.length; i++) {
    appendResultItem(results[i], i);
  }
  renderedResultsCount = results.length;
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

function pollGen() {
  pollTimerGen = setInterval(async () => {
    const resp = await fetch('/api/log/' + currentJobGen);
    const data = await resp.json();
    renderLogGen(data.log);
    updateProgressGen(data.log);
    renderResults(data.results);
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


def run_job(job_id: str, pdf_path: Path, out_dir: Path):
    job = jobs[job_id]

    def log(msg):
        job["log"].append(str(msg))

    try:
        def on_audit(problems):
            job["audit"] = [
                {"id": i, "kind": kind, "section": section, "item": item, "text": text}
                for i, (kind, section, item, text) in enumerate(problems)
            ]

        core.download_agenda_pdfs(pdf_path, out_dir, delay=0.3, log=log, on_audit=on_audit)
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
                        capture_output=True, text=True, timeout=60,
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
    finally:
        doc.close()
    return None


def run_generate_job(job_id: str, file_paths: list):
    job = jobs[job_id]

    def log(msg):
        job["log"].append(str(msg))

    log(f"Знайдено файлів для обробки: {len(file_paths)}")

    session = requests.Session()
    session.headers.update(core.HEADERS)

    try:
        ok, failed = 0, []
        for file_path in file_paths:
            filename = file_path.name
            match = DECISION_NUMBER_RE.match(filename)
            if not match:
                log(f"[!] {filename}: не вдалося визначити номер з назви файлу "
                    f"(очікується формат NNN-XX.pdf)")
                failed.append(filename)
                continue

            number = match.group(1)
            url = SEARCH_URL_TEMPLATE.format(number=number)
            try:
                resp = session.get(url, timeout=20)
                resp.raise_for_status()
            except Exception as exc:
                log(f"[!] {filename} (номер {number}): помилка запиту до пошуку — {exc}")
                failed.append(filename)
                continue

            link_match = SEARCH_RESULT_LINK_RE.search(resp.text)
            if not link_match:
                log(f"[!] {filename} (номер {number}): у результатах пошуку не знайдено "
                    f"жодного посилання на проєкт рішення")
                failed.append(filename)
                continue

            href, project_id = link_match.group(1), link_match.group(2)
            project_url = SEARCH_RESULT_BASE_URL + href

            try:
                project_resp = session.get(project_url, timeout=20)
                project_resp.raise_for_status()
            except Exception as exc:
                log(f"[!] {filename} (номер {number}): помилка запиту до сторінки проєкту — {exc}")
                failed.append(filename)
                continue

            title = extract_project_title(project_resp.text)
            date = extract_project_date(project_resp.text)
            if not title or not date:
                missing = ", ".join(
                    n for n, v in (("заголовок", title), ("дата реєстрації", date)) if not v
                )
                log(f"[!] {filename} (номер {number}): не вдалося знайти на сторінці проєкту: {missing} — {project_url}")
                failed.append(filename)
                continue

            try:
                short_name = extract_speaker_short_name(file_path)
            except Exception as exc:
                log(f"[!] {filename} (номер {number}): помилка розпізнавання доповідача — {exc}")
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
                "filename": filename, "line": line, "section": None, "url": project_url,
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


def build_generated_doc(results: list) -> Path:
    """Будує "Згенерованний ПР.docx" у корені застосунку: 6 розділів з
    підказки, у кожному — обрані користувачем пункти, де сам текст пункту —
    гіперлінк на сторінку проєкту рішення. Нумерація пунктів у кожному
    розділі завжди починається з 1) (без префіксу номера розділу)."""
    by_section = {n: [] for n in range(1, 7)}
    for item in results:
        section = item.get("section")
        if section in by_section:
            by_section[section].append(item)

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)

    for section_num in range(1, 7):
        header = doc.add_paragraph()
        header_run = header.add_run(f"{section_num}) {SECTION_TITLES[section_num]}:")
        header_run.bold = True

        for sub_index, item in enumerate(by_section[section_num], start=1):
            title = item["line"].lstrip("\t")
            para = doc.add_paragraph()
            para.paragraph_format.left_indent = Inches(0.4)
            para.add_run(f"{sub_index}) ")
            add_hyperlink(para, item["url"], title)

    doc.save(GENERATED_DOC_PATH)
    return GENERATED_DOC_PATH


@app.route("/")
def index():
    return render_template_string(PAGE, default_out_dir=str(DEFAULT_OUT_DIR))


@app.route("/api/sections")
def api_sections_get():
    return jsonify(sections={str(n): SECTION_TITLES[n] for n in range(1, 7)})


@app.route("/api/sections", methods=["POST"])
def api_sections_set():
    data = request.get_json(silent=True) or {}
    section = data.get("section")
    title = (data.get("title") or "").strip()
    if section not in range(1, 7):
        return jsonify(error="invalid section"), 400
    if not title:
        return jsonify(error="Назва розділу не може бути порожньою"), 400

    SECTION_TITLES[section] = title
    save_section_titles(SECTION_TITLES)
    return jsonify(ok=True, sections={str(n): SECTION_TITLES[n] for n in range(1, 7)})


@app.route("/upload", methods=["POST"])
def upload():
    uploaded = request.files.get("pdf")
    if not uploaded or not uploaded.filename:
        return jsonify(error="Файл не обрано"), 400

    out_dir_raw = request.form.get("out_dir", "").strip()
    out_dir = Path(out_dir_raw).expanduser() if out_dir_raw else DEFAULT_OUT_DIR

    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    job_id = uuid.uuid4().hex
    pdf_path = UPLOADS_DIR / f"{job_id}_{uploaded.filename}"
    uploaded.save(pdf_path)

    jobs[job_id] = {"log": [], "done": False, "out_dir": str(out_dir), "audit": []}
    thread = threading.Thread(target=run_job, args=(job_id, pdf_path, out_dir), daemon=True)
    thread.start()

    return jsonify(job_id=job_id)


@app.route("/generate/upload", methods=["POST"])
def generate_upload():
    uploaded_files = [f for f in request.files.getlist("files") if f.filename]
    if not uploaded_files:
        return jsonify(error="Файли не обрано"), 400

    job_id = uuid.uuid4().hex
    job_dir = GENERATE_UPLOADS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    file_paths = []
    for uploaded in uploaded_files:
        dest = job_dir / Path(uploaded.filename).name
        uploaded.save(dest)
        file_paths.append(dest)
    file_paths.sort(key=lambda p: p.name)

    jobs[job_id] = {"log": [], "done": False, "out_dir": str(job_dir), "results": []}
    thread = threading.Thread(target=run_generate_job, args=(job_id, file_paths), daemon=True)
    thread.start()

    return jsonify(job_id=job_id)


@app.route("/api/log/<job_id>")
def api_log(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify(error="unknown job"), 404
    return jsonify(log=job["log"], done=job["done"], results=job.get("results", []), audit=job.get("audit", []))


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
    if section is not None and section not in range(1, 7):
        return jsonify(error="invalid section"), 400

    results[index]["section"] = section
    return jsonify(ok=True)


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
        doc_path = build_generated_doc(results)
    except Exception as exc:
        return jsonify(error=f"Не вдалося створити документ: {exc}"), 500

    open_in_os(str(doc_path))
    return jsonify(ok=True, path=str(doc_path))


@app.route("/api/open/<job_id>", methods=["POST"])
def api_open(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify(error="unknown job"), 404
    open_in_os(job["out_dir"])
    return jsonify(ok=True)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5005, debug=False)
