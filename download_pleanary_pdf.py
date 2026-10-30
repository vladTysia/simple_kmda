#!/usr/bin/env python3
"""
Приймає на вхід PDF порядку денного (наприклад "Президія_...pdf") з
гіперпосиланнями на сторінки виду
https://kmr.gov.ua/uk/content/proekt-rishennya-kyyivskoyi-miskoyi-rady-<N>.

Для кожного посилання визначає дату та номер рішення з тексту поруч
(формат "(від DD.MM.YYYY № .../ПР)"), відкриває відповідну сторінку і
скачує з розділу "Файли" pdf, підпис якого містить TARGET_TEXT
("На пленарне засідання"), зберігаючи його під іменем "DD.MM.YYYY_NNN_ПР.pdf".
"""

import argparse
import os
import re
import shutil
import sys
import time
from pathlib import Path
from urllib.parse import urljoin

import fitz  # PyMuPDF
import requests

TARGET_TEXT = "На пленарне засідання"
SECOND_READING_TEXT = f"{TARGET_TEXT} (друге читання)"
ADDITIONAL_TARGET_TEXTS = [
    "Пропозиції",
    "Витяг з протоколу",
    "Висновок управління захисту прав та інтересів територіальної громади міста Києва",
    "Лист",  # будь-який "Лист департамента"/"Лист департаменту"/від кого завгодно
]
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}
PROJECT_URL_RE = re.compile(
    r"kmr\.gov\.ua/uk/(?:content/proy?ekt-rishennya-kyyivskoyi-miskoyi-rady-\d+|node/\d+)",
    re.IGNORECASE,
)
CITATION_RE = re.compile(
    # Номер перед "ПР" зазвичай відділений "/", але подекуди в документі
    # це одруківка з пробілом замість риски (напр. "524 ПР)") — приймаємо і те, і те.
    # "дата №" також подекуди задвоєна прямо в тексті документа
    # (напр. "31.12.2024 № 31.12.2024 № 08/231-2008/ПР)") — пропускаємо повтор.
    r"\(\s*від\s+(\d{2}\.\d{2}\.\d{4})\d*\s*№\s*(?:\d{2}\.\d{2}\.\d{4}\s*№\s*)?([0-9\s/\-]+?[/\s]\s*ПР)\s*\)",
    re.IGNORECASE,
)
MAJOR_HEADING_RE = re.compile(r"^\d+\.$")
# Підрозділи до 4 рівнів: "1.1.", "1.1.1.", "1.1.1.1." (+ текст заголовка в тому ж рядку).
MINOR_HEADING_RE = re.compile(r"^(\d+(?:\.\d+){1,3})\.\s+(?=[А-ЯЁІЇЄ])")
ITEM_NUMBER_RE = re.compile(r"^(\d+)\)\s+Про\b")
ITEM_NUMBER_ALONE_RE = re.compile(r"^(\d+)\)\s*$")
# На скільки pt (по вертикалі) початок посилання може бути вищим за рядок з
# номером пункту "N)" і все одно вважатися посиланням цього пункту.
ITEM_Y_TOLERANCE = 6.0
# Зсув посилання відносно номера пункту, більший за цей, вважаємо
# підозрілим вирівнюванням (у нормі ~0.5pt) — про нього попереджає audit_agenda.
MISALIGN_WARN_PT = 1.0
# Де має починатися посилання відносно рядка з номером пункту (pt по вертикалі).
LINK_ALIGNED_DY = -0.5


def normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def extract_section_headings(doc):
    """Повертає список (page_index, y0, мітка) для заголовків розділів на
    кшталт "1.", "1.1. Опис ...", "1.1.1. Опис ..." (до 4 рівнів), де мітка —
    повний номер ("1", "1.1", "1.1.1"), у порядку появи в документі."""
    headings = []
    for page_index in range(len(doc)):
        page_dict = doc[page_index].get_text("dict")
        for block in page_dict["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                text = "".join(span["text"] for span in line["spans"]).strip()
                y0 = line["bbox"][1]
                if MAJOR_HEADING_RE.match(text):
                    headings.append((page_index, y0, text[:-1]))
                    continue
                minor_match = MINOR_HEADING_RE.match(text)
                if minor_match:
                    headings.append((page_index, y0, minor_match.group(1)))
    return headings


def extract_item_number_markers(doc):
    """Повертає список (page_index, y0, номер) для маркерів пунктів порядку
    денного на кшталт "209) Про ..." (у порядку появи в документі). Номер
    пункту скидається на початку кожного розділу/підрозділу — це нормально,
    так позначено в самому документі."""
    lines_all = []
    for page_index in range(len(doc)):
        page_dict = doc[page_index].get_text("dict")
        for block in page_dict["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                text = "".join(span["text"] for span in line["spans"]).strip()
                if text:
                    lines_all.append((page_index, line["bbox"][1], text))

    markers = []
    n = len(lines_all)
    for i, (page_index, y0, text) in enumerate(lines_all):
        match = ITEM_NUMBER_RE.match(text)
        if match:
            markers.append((page_index, y0, int(match.group(1))))
            continue
        match = ITEM_NUMBER_ALONE_RE.match(text)
        if match and i + 1 < n and lines_all[i + 1][2].startswith("Про"):
            markers.append((page_index, y0, int(match.group(1))))
    return markers


def assign_item_numbers(item_positions, markers):
    """Для кожного пункту (за позицією його першого посилання) визначає
    номер пункту порядку денного, спираючись на останній маркер "N) Про ..."
    перед ним."""
    # Прямокутник посилання на той самий рядок, що й власний маркер пункту,
    # починається трохи вище за bbox рядка з маркером: зазвичай на ~0.5pt, але
    # в окремих пунктах до ~2.5pt (інший розмір/baseline маркера "10)").
    # Допуск ITEM_Y_TOLERANCE надійно ставить позицію пункту "після" маркера
    # того ж рядка, не зачіпаючи сусідні пункти (інтервал між ними ~17pt).
    events = [(pno, y0, 0, ("marker", num)) for pno, y0, num in markers]
    events += [(pno, y0 + ITEM_Y_TOLERANCE, 1, ("item", idx)) for idx, (pno, y0) in enumerate(item_positions)]
    events.sort(key=lambda e: (e[0], e[1], e[2]))

    result = [None] * len(item_positions)
    current = None
    for _, _, _, payload in events:
        if payload[0] == "marker":
            current = payload[1]
        else:
            _, idx = payload
            result[idx] = current
    return result


def section_path(label: str):
    """"1.1.1" -> ("1", "1.1", "1.1.1") — ланцюжок вкладених папок розділу."""
    parts = label.split(".")
    return tuple(".".join(parts[:i]) for i in range(1, len(parts) + 1))


def assign_sections(item_positions, headings):
    """Для кожного пункту (за позицією його першого посилання в документі)
    визначає шлях розділу — кортеж ("1", "1.1", "1.1.1") або () якщо
    заголовків перед пунктом не було, спираючись на останній заголовок перед ним."""
    events = [(pno, y0, 0, ("heading", label)) for pno, y0, label in headings]
    events += [(pno, y0, 1, ("item", idx)) for idx, (pno, y0) in enumerate(item_positions)]
    events.sort(key=lambda e: (e[0], e[1], e[2]))

    result = [()] * len(item_positions)
    current = ()
    for _, _, _, payload in events:
        if payload[0] == "heading":
            current = section_path(payload[1])
        else:
            result[payload[1]] = current
    return result


def _page_word_spans(doc, page_index, cache):
    """Кешований список (y0, x0, cx, cy, text) для текстових фрагментів
    сторінки — використовується, щоб вибирати лише слова, чий центр дійсно
    потрапляє в прямокутник посилання, а не все, що зачіпає його межу."""
    spans = cache.get(page_index)
    if spans is not None:
        return spans

    spans = []
    for block in doc[page_index].get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                x0, y0, x1, y1 = span["bbox"]
                spans.append((y0, x0, (x0 + x1) / 2, (y0 + y1) / 2, span["text"]))
    cache[page_index] = spans
    return spans


def _text_in_rect(doc, page_index, rect, cache):
    """Текст слів, чий центр лежить у межах rect. На відміну від
    get_textbox(rect), не захоплює сусідній рядок і не "розтягується" по
    всьому рядку, коли rect вузький (наприклад, лише права половина рядка)."""
    spans = _page_word_spans(doc, page_index, cache)
    selected = [
        (y0, x0, text)
        for y0, x0, cx, cy, text in spans
        if rect.x0 <= cx <= rect.x1 and rect.y0 <= cy <= rect.y1
    ]
    # Групуємо y0 по "рядку" з допуском: окремий гліф (напр. одинока дужка
    # "(") подеколи має інший baseline (~4pt) за сусідні слова того ж рядка,
    # хоча візуально стоїть на ньому — точне сортування за y0 переставляло
    # такі слова в текст сусіднього рядка й ламало порядок цитати.
    selected.sort(key=lambda t: (round(t[0] / 8), t[1]))
    return "".join(t[2] for t in selected)


def extract_agenda_items(pdf_path: Path):
    """Повертає список (project_url, date, number, section_path, item_number) з PDF порядку денного."""
    doc = fitz.open(pdf_path)
    span_cache = {}

    raw_links = []  # (page_index, uri, rect)
    for page_index in range(len(doc)):
        for link in doc[page_index].get_links():
            uri = link.get("uri")
            if not uri or not PROJECT_URL_RE.search(uri):
                continue
            rect = fitz.Rect(link["from"])
            if rect.is_empty:
                # Подекуди в документі трапляється вироджений Rect(0,0,0,0) —
                # без нього текст, зате з y0=0.0, яке "перетягує" позицію
                # пункту (first_y0 = min(...)) на верх сторінки й ламає
                # визначення розділу/номера пункту за позицією.
                continue
            raw_links.append((page_index, uri, rect))

    # Групуємо послідовні посилання з однаковим uri (можуть охоплювати кілька рядків/сторінок).
    groups = []
    current_uri = None
    current_entries = []
    for page_index, uri, rect in raw_links:
        if uri != current_uri:
            if current_uri is not None:
                groups.append((current_uri, current_entries))
            current_uri = uri
            current_entries = []
        current_entries.append((page_index, rect))
    if current_uri is not None:
        groups.append((current_uri, current_entries))

    parsed = []  # (uri, date, number, first_page, first_y0)
    unresolved = []  # (uri, first_page, first_y0)
    for uri, entries in groups:
        by_page = {}
        page_order = []
        for page_index, rect in entries:
            if page_index not in by_page:
                by_page[page_index] = []
                page_order.append(page_index)
            by_page[page_index].append(rect)

        texts = []
        for page_index in page_order:
            for rect in by_page[page_index]:
                # Беремо лише слова, чий центр потрапляє в прямокутник
                # посилання (не merge в один bounding box і не get_textbox):
                # коли перший рядок посилання ділить рядок з хвостом
                # попереднього пункту, обидва попередні підходи захоплювали
                # чужий текст зліва чи з сусіднього рядка.
                texts.append(_text_in_rect(doc, page_index, rect, span_cache).strip())
        text = " ".join(t for t in texts if t)

        # Якщо посилання все ж частково перетинає хвіст попереднього пункту,
        # в тексті може лишитися його "чужа" цитата — вона завжди йде першою,
        # а власна цитата пункту завжди останньою (прямо перед новим пунктом).
        matches = list(CITATION_RE.finditer(text))
        if not matches:
            first_page = page_order[0]
            first_y0 = min(r.y0 for r in by_page[first_page])
            unresolved.append((uri, first_page, first_y0))
            continue
        match = matches[-1]

        date = match.group(1)
        full_number = re.sub(r"\s+", "", match.group(2))
        full_number = re.sub(r"(?<!/)ПР$", "/ПР", full_number, flags=re.IGNORECASE)
        number = full_number.rsplit("-", 1)[-1]  # "08/231-855/ПР" -> "855/ПР"
        first_page = page_order[0]
        first_y0 = min(r.y0 for r in by_page[first_page])
        parsed.append((uri, date, number, first_page, first_y0))

    # Розділ і номер пункту рахуємо разом і для розпізнаних, і для
    # нерозпізнаних посилань (позиція в документі однаково відома для обох) —
    # щоб у підсумковому звіті можна було назвати пункт, а не лише URI.
    all_positions = [(p[3], p[4]) for p in parsed] + [(u[1], u[2]) for u in unresolved]

    headings = extract_section_headings(doc)
    all_sections = assign_sections(all_positions, headings)

    item_markers = extract_item_number_markers(doc)
    all_item_numbers = assign_item_numbers(all_positions, item_markers)

    n_parsed = len(parsed)
    items = [
        (uri, date, number, section, item_number)
        for (uri, date, number, _, _), section, item_number
        in zip(parsed, all_sections[:n_parsed], all_item_numbers[:n_parsed])
    ]
    unresolved = [
        (uri, section, item_number)
        for (uri, _, _), section, item_number
        in zip(unresolved, all_sections[n_parsed:], all_item_numbers[n_parsed:])
    ]

    doc.close()
    return items, unresolved


def audit_agenda(pdf_path: Path):
    """Шукає в PDF порядку денного помилки, які легко припустити вручну
    (пунктів сотні) і які ламають розбір. Повертає список зауважень
    (вид, розділ, номер пункту або None, повідомлення, extra), впорядкований за
    розташуванням у документі. extra — дані для виправлення (або None): у
    "немає посилання" {"pos": [сторінка, y]} пункту, в "однакове посилання"
    {"items": [{"num", "section", "pos"}, ...]} — див. set_item_link. Нічого не завантажує.

    Види зауважень:
      "немає посилання"      — у пункті "N) Про ..." немає гіперпосилання;
      "однакове посилання"   — те саме посилання у кількох пунктах;
      "кілька посилань"      — у пункті кілька різних посилань;
      "зсув посилання"       — посилання не вирівняне з номером пункту;
      "посилання без пункту" — посилання поза будь-яким пунктом "N) Про ...";
      "порушена нумерація"   — номери пунктів розділу не йдуть підряд 1,2,3...
    """
    doc = fitz.open(pdf_path)
    try:
        headings = extract_section_headings(doc)
        markers = extract_item_number_markers(doc)

        # Початок тексту пункту — щоб у зауваженні було видно, який це пункт.
        snippets = {}
        for page_index in range(len(doc)):
            lines = []
            for block in doc[page_index].get_text("dict")["blocks"]:
                if block.get("type") != 0:
                    continue
                for line in block["lines"]:
                    text = "".join(span["text"] for span in line["spans"]).strip()
                    if text:
                        lines.append((line["bbox"][1], text))
            lines.sort(key=lambda t: t[0])
            for i, (y0, text) in enumerate(lines):
                tail = " ".join(t for _, t in lines[i:i + 6])
                snippets[(page_index, y0)] = normalize_spaces(re.sub(r"^\d+\)\s*", "", tail))[:110]

        links = []  # (page_index, y0, uri)
        for page_index in range(len(doc)):
            for link in doc[page_index].get_links():
                uri = link.get("uri")
                rect = fitz.Rect(link["from"])
                if uri and PROJECT_URL_RE.search(uri) and not rect.is_empty:
                    links.append((page_index, rect.y0, uri))
    finally:
        doc.close()

    events = [(p, y, 0, ("heading", label)) for p, y, label in headings]
    events += [(p, y, 1, ("marker", (num, p, y))) for p, y, num in markers]
    events += [(p, y + ITEM_Y_TOLERANCE, 2, ("link", (p, y, uri))) for p, y, uri in links]
    events.sort(key=lambda e: (e[0], e[1], e[2]))

    section = "?"
    current = None  # поточний пункт: dict
    items = []  # пункти у порядку документа
    orphans = []  # (розділ, page_index, uri) — посилання поза пунктом
    for _, _, _, (kind, payload) in events:
        if kind == "heading":
            section = payload
            current = None
        elif kind == "marker":
            num, p, y = payload
            current = {"section": section, "num": num, "page": p, "y": y, "links": [],
                       "snippet": snippets.get((p, y), "")}
            items.append(current)
        elif current is None:
            orphans.append((section, payload[0], payload[2]))
        else:
            current["links"].append(payload)

    problems = []  # (порядок, вид, розділ, пункт, текст)

    def add(order_item, kind, section_label, num, text, extra=None):
        problems.append(((order_item["page"], order_item["y"]) if order_item else (0, 0), kind, section_label, num, text, extra))

    def item_label(it):
        return f"«{it['snippet']}»" if it["snippet"] else ""

    uri_items = {}
    last_num = {}
    for it in items:
        uris = list(dict.fromkeys(u for _, _, u in it["links"]))
        for u in uris:
            uri_items.setdefault(u, []).append(it)

        if not uris:
            add(it, "немає посилання", it["section"], it["num"],
                f"у пункті немає гіперпосилання на проєкт рішення — його не буде скачано. {item_label(it)}",
                {"pos": [it["page"], it["y"]]})
        elif len(uris) > 1:
            add(it, "кілька посилань", it["section"], it["num"],
                f"у пункті {len(uris)} різні посилання — усі проєкти скачаються з одним номером пункту. "
                f"{item_label(it)}")

        if it["links"]:
            first_page, first_y, _ = it["links"][0]
            dy = first_y - it["y"] if first_page == it["page"] else None
            if dy is None:
                add(it, "зсув посилання", it["section"], it["num"],
                    f"посилання починається на іншій сторінці ({first_page + 1}), ніж номер пункту "
                    f"({it['page'] + 1}). {item_label(it)}")
            elif abs(dy) > MISALIGN_WARN_PT:
                add(it, "зсув посилання", it["section"], it["num"],
                    f"посилання зсунуте відносно номера пункту на {dy:+.1f} pt (стор. {it['page'] + 1}; у нормі ≈ -0.5) — "
                    f"перевірте, що файл отримав правильний номер. {item_label(it)}")

        prev = last_num.get(it["section"])
        if prev is None and it["num"] != 1:
            add(it, "порушена нумерація", it["section"], it["num"],
                f"нумерація пунктів розділу починається не з 1, а з {it['num']}")
        elif prev is not None and it["num"] != prev + 1:
            add(it, "порушена нумерація", it["section"], it["num"],
                f"після пункту {prev} іде {it['num']} (пропущено або продубльовано номер)")
        last_num[it["section"]] = it["num"]

    for uri, owners in uri_items.items():
        if len(owners) < 2:
            continue
        nums = ", ".join(str(o["num"]) for o in owners)
        sections = sorted({o["section"] for o in owners})
        consecutive = len(sections) == 1 and all(
            b["num"] == a["num"] + 1 for a, b in zip(owners, owners[1:])
        )
        note = ("підряд — програма об'єднає їх в один запис і скачає проєкт один раз"
                if consecutive else "проєкт буде скачано окремо для кожного з них")
        add(owners[0], "однакове посилання", owners[0]["section"], owners[0]["num"],
            f"те саме посилання у пунктах {nums} (розділ {', '.join(sections)}): {note}. {uri}",
            {"items": [{"num": o["num"], "section": o["section"], "pos": [o["page"], o["y"]]} for o in owners]})

    for section_label, page_index, uri in orphans:
        problems.append(((page_index, 0), "посилання без пункту", section_label, None,
                         f"посилання на стор. {page_index + 1} не належить жодному пункту «N) Про ...»: {uri}", None))

    problems.sort(key=lambda p: p[0])
    return [(kind, sec, num, text, extra) for _, kind, sec, num, text, extra in problems]


def _document_layout(doc):
    """Пункти порядку денного у порядку документа: [{section, num, page, y,
    end}], де end = (сторінка, y) наступного маркера/заголовка — межа пункту.
    Розділ і межі рахуються так само, як у audit_agenda."""
    events = [(p, y, 0, ("heading", label)) for p, y, label in extract_section_headings(doc)]
    events += [(p, y, 1, ("marker", num)) for p, y, num in extract_item_number_markers(doc)]
    events.sort(key=lambda e: (e[0], e[1], e[2]))
    items, section = [], "?"
    for i, (p, y, _, (kind, payload)) in enumerate(events):
        if kind == "heading":
            section = payload
            continue
        nxt = events[i + 1] if i + 1 < len(events) else None
        items.append({"section": section, "num": payload, "page": p, "y": y,
                      "end": (nxt[0], nxt[1]) if nxt else (len(doc), 0.0)})
    return items


def _item_links(doc, item, cache=None):
    """Усі анотації-посилання пункту (будь-якого виду, зокрема "биті" без адреси):
    [(page_index, link_dict)]. Належність — за тим самим правилом, що й в audit."""
    start, end = (item["page"], item["y"]), item["end"]
    found = []
    for pno in range(item["page"], min(end[0], len(doc) - 1) + 1):
        if cache is not None and pno in cache:
            links = cache[pno]
        else:
            links = doc[pno].get_links()
            if cache is not None:
                cache[pno] = links
        for link in links:
            rect = fitz.Rect(link["from"])
            if not rect.is_empty and start <= (pno, rect.y0 + ITEM_Y_TOLERANCE) < end:
                found.append((pno, link))
    return found


def _save_in_place(doc, path: Path):
    """Зберігає зміни в той самий файл: інкрементально (швидко, не чіпає решту
    файлу), а якщо PDF так зберегти не можна — повним перезаписом."""
    try:
        doc.saveIncr()
    except (ValueError, RuntimeError):
        tmp = path.with_name(path.stem + ".tmp.pdf")
        doc.save(tmp)
        doc.close()
        os.replace(tmp, path)


def fix_link_misalignment(pdf_path: Path, log=print) -> int:
    """Вирівнює "зсунуті" посилання: якщо початок посилання пункту відхиляється від
    рядка з номером пункту більше ніж на MISALIGN_WARN_PT, усі прямокутники посилань
    пункту зсуваються так, щоб початок був на LINK_ALIGNED_DY. Правка — прямо у файлі;
    кожну виправлену помилку записує в журнал рядком "[~]". Повертає кількість пунктів."""
    doc = fitz.open(pdf_path)
    try:
        cache, fixed = {}, 0
        for item in _document_layout(doc):
            links = [(p, l) for p, l in _item_links(doc, item, cache)
                     if l.get("kind") == fitz.LINK_URI and PROJECT_URL_RE.search(l.get("uri", ""))]
            if not links:
                continue
            first_page, first = min(links, key=lambda t: (t[0], fitz.Rect(t[1]["from"]).y0))
            if first_page != item["page"]:
                continue  # інша сторінка — це не зсув по вертикалі, лишаємо для ручної перевірки
            dy = fitz.Rect(first["from"]).y0 - item["y"]
            if abs(dy) <= MISALIGN_WARN_PT:
                continue
            shift = LINK_ALIGNED_DY - dy
            for pno, link in links:
                r = fitz.Rect(link["from"])
                doc[pno].update_link({**link, "from": fitz.Rect(r.x0, r.y0 + shift, r.x1, r.y1 + shift)})
            fixed += 1
            log(f"[~] Виправлено автоматично: розділ {item['section']}, пункт {item['num']} — "
                f"посилання було зсунуте на {dy:+.1f} pt (стор. {item['page'] + 1}), вирівняно за номером пункту.")
        if fixed:
            _save_in_place(doc, pdf_path)
        return fixed
    finally:
        doc.close()


def _add_links_over_item(doc, item, uri: str) -> int:
    """Запасний шлях, коли в пункті взагалі немає прямокутників посилання: додає
    по одному на кожен візуальний рядок тексту пункту (як це робить Word)."""
    start, end = (item["page"], item["y"]), item["end"]
    added = 0
    for pno in range(item["page"], min(end[0], len(doc) - 1) + 1):
        rows = []  # [x0, y0, x1, y1]
        for block in doc[pno].get_text("dict")["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                text = "".join(sp["text"] for sp in line["spans"]).strip()
                x0, y0, x1, y1 = line["bbox"]
                if not text or text.isdigit() or not (start <= (pno, y0) < end):
                    continue
                if pno == item["page"] and abs(y0 - item["y"]) < 0.01:
                    # рядок із маркером: посилання починається від "Про", а не від "N)"
                    first = next((sp for sp in line["spans"] if sp["text"].strip().startswith("Про")), None)
                    if first:
                        x0 = first["bbox"][0] - 2.2
                rows.append([x0, y0, x1, y1])
        rows.sort(key=lambda r: (round(r[1]), r[0]))
        merged = []
        for r in rows:
            if merged and abs(r[1] - merged[-1][1]) < 2:
                merged[-1][0], merged[-1][2] = min(merged[-1][0], r[0]), max(merged[-1][2], r[2])
            else:
                merged.append(r)
        for x0, y0, x1, y1 in merged:
            doc[pno].insert_link({"kind": fitz.LINK_URI, "uri": uri,
                                  "from": fitz.Rect(x0, y0 + LINK_ALIGNED_DY, x1, y1 + 1.2)})
            added += 1
    return added


def set_item_link(pdf_path: Path, pos, uri: str) -> int:
    """Прописує посилання uri всім прямокутникам пункту, що починається на pos =
    [сторінка, y] (так його позначає audit_agenda): замінює "биті" посилання без
    адреси і дублікати. Правка — прямо у файлі. Повертає кількість прямокутників."""
    doc = fitz.open(pdf_path)
    try:
        item = next((it for it in _document_layout(doc)
                     if it["page"] == pos[0] and abs(it["y"] - pos[1]) < 0.01), None)
        if item is None:
            raise ValueError("Пункт не знайдено у файлі — можливо, файл змінився. Завантажте порядок денний ще раз.")
        links = _item_links(doc, item)
        for pno, link in links:
            doc[pno].update_link({"kind": fitz.LINK_URI, "uri": uri, "from": fitz.Rect(link["from"]),
                                  "xref": link["xref"]})
        count = len(links) or _add_links_over_item(doc, item, uri)
        if not count:
            raise ValueError("У пункті не знайдено тексту, на який можна поставити посилання.")
        _save_in_place(doc, pdf_path)
        return count
    finally:
        doc.close()


def log_audit(problems, log=print):
    """Виводить результат audit_agenda в журнал. Повертає кількість зауважень."""
    if not problems:
        log("Перевірка порядку денного: помилок не знайдено.")
        return 0
    log(f"Перевірка порядку денного: знайдено {len(problems)} зауважень (варто перевірити вручну в PDF):")
    for kind, sec, num, text, _ in problems:
        where = f"розділ {sec}" + (f", пункт {num}" if num is not None else "")
        log(f"  - [{kind}] {where}: {text}")
    return len(problems)


def pad_pdf_to_multiple_of_4(pdf_path: Path, log=print):
    """Додає в кінець pdf порожні сторінки, щоб загальна кількість сторінок
    стала кратною 4. Нічого не робить, якщо вона вже кратна 4."""
    doc = fitz.open(pdf_path)
    page_count = len(doc)
    remainder = page_count % 4
    if remainder == 0:
        doc.close()
        return

    pages_to_add = 4 - remainder
    width, height = doc[-1].rect.width, doc[-1].rect.height
    for _ in range(pages_to_add):
        doc.new_page(width=width, height=height)

    tmp_path = pdf_path.with_name(pdf_path.name + ".tmp")
    doc.save(tmp_path)
    doc.close()
    tmp_path.replace(pdf_path)

    log(f"    + додано {pages_to_add} порожніх стор. ({page_count} -> {page_count + pages_to_add})")


def find_all_files(html: str, page_url: str):
    """Повертає список (label, absolute_url) для всіх pdf-посилань у блоці 'Файли:'."""
    files_block_start = html.find('field-label field-content">Файли')
    if files_block_start == -1:
        return []

    block = html[files_block_start:]
    pattern = re.compile(r'<a[^>]+href="([^"]+\.pdf)"[^>]*>\s*([^<]*)</a>', re.IGNORECASE)

    return [
        (label.strip(), urljoin(page_url, href.strip()))
        for href, label in pattern.findall(block)
    ]


def fetch_page(session, page_url: str):
    """Завантажує сторінку; якщо хост kmr.gov.ua / old.kmr.gov.ua відмовляє
    (наприклад, 403 від Cloudflare), пробує той самий шлях на іншому хості."""
    alt_url = (
        page_url.replace("://old.kmr.gov.ua", "://kmr.gov.ua", 1)
        if "://old.kmr.gov.ua" in page_url
        else page_url.replace("://kmr.gov.ua", "://old.kmr.gov.ua", 1)
    )
    last_exc = None
    for url in dict.fromkeys([page_url, alt_url]):
        try:
            resp = session.get(url, timeout=30)
            resp.raise_for_status()
            return resp
        except requests.RequestException as exc:
            last_exc = exc
    raise last_exc


def sanitize_filename(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', "_", name).strip()
    return re.sub(r"\s+", " ", name)


REPORT_NAME = "Звіт_розділу.pdf"
MAIN_FILENAME_RE = re.compile(r"^\d+_\d{2}\.\d{2}\.\d{4}_")
MAIN_FILENAME_NUMBER_TAG_RE = re.compile(r"^(\d+)_\d{2}\.\d{2}\.\d{4}_(.+)\.pdf$")
FOR_DEPUTIES_DIR_NAME = "Для депутатів"
# Колишня (російська) назва: у вже скачаних наборах така папка лишилась від старих запусків.
LEGACY_FOR_DEPUTIES_DIR_NAMES = ("Для депутатов",)
# Згенеровані файли зі старими (російськими) назвами від попередніх запусків.
LEGACY_ARTIFACT_PATTERNS = ("Отчет_раздела.pdf", "Отчет_раздела_и_микс_*.pdf", "Микс_*.pdf", "Друк_раздела_*.pdf")


def remove_legacy_artifacts(out_dir: Path, log=print) -> int:
    """Прибирає звіти й мікси зі старими назвами: програма щоразу генерує їх
    заново вже з українськими назвами, а старі лишилися б поруч (і могли б потрапити
    в нові звіти)."""
    removed = 0
    for pattern in LEGACY_ARTIFACT_PATTERNS:
        for path in out_dir.rglob(pattern):
            if path.is_file():
                path.unlink()
                removed += 1
    if removed:
        log(f"[=] Прибрано {removed} згенерованих файлів зі старими назвами — замінено файлами з українськими назвами.")
    return removed


def build_section_reports(out_dir: Path, log=print):
    """У кожній папці розділу (out_dir/<major>) створює REPORT_NAME —
    об'єднання всіх pdf з цього розділу та його підрозділів, у порядку
    їхнього індексу в назві файлу."""
    if not out_dir.exists():
        return

    for section_dir in sorted(p for p in out_dir.iterdir() if p.is_dir()):
        # Копії пунктів у "Для депутатів" названі так само, як основні файли, — без
        # цього виключення при повторному запуску вони дублювалися б у звіті.
        pdf_files = [
            p for p in section_dir.rglob("*.pdf")
            if p.name != REPORT_NAME and MAIN_FILENAME_RE.match(p.name)
            and FOR_DEPUTIES_DIR_NAME not in p.relative_to(section_dir).parts
            and not any(name in p.relative_to(section_dir).parts for name in LEGACY_FOR_DEPUTIES_DIR_NAMES)
        ]
        if not pdf_files:
            continue

        def sort_key(p):
            match = re.match(r"^(\d+)_", p.name)
            return (int(match.group(1)) if match else 10**9, p.name)

        pdf_files.sort(key=sort_key)

        merged = fitz.open()
        for pdf_file in pdf_files:
            with fitz.open(pdf_file) as src:
                merged.insert_pdf(src)

        report_path = section_dir / REPORT_NAME
        merged.save(report_path)
        merged.close()
        log(f"[=] Зібрано звіт розділу {section_dir.name}: {report_path.relative_to(out_dir)} "
            f"({len(pdf_files)} файлів)")


def _mix_sort_index(mix_path: Path) -> int:
    """Визначає порядковий індекс мікс-файлу — він тепер зашитий прямо в
    назву ("Мікс_<папка>_<index_prefix>_<номер>.pdf")."""
    prefix = f"Мікс_{mix_path.parent.name}_"
    if not mix_path.stem.startswith(prefix):
        return 10**9
    match = re.match(r"^(\d+)_", mix_path.stem[len(prefix):])
    return int(match.group(1)) if match else 10**9


def build_section_mix_reports(out_dir: Path, log=print):
    """У кожній папці розділу (out_dir/<major>) створює
    "Звіт_розділу_і_мікс_<номер папки>.pdf" — об'єднання всіх файлів "Мікс_*"
    з цього розділу та його підрозділів, у порядку пунктів повістки."""
    if not out_dir.exists():
        return

    for section_dir in sorted(p for p in out_dir.iterdir() if p.is_dir()):
        mix_files = list(section_dir.rglob("Мікс_*.pdf"))
        if not mix_files:
            continue

        mix_files.sort(key=lambda p: (_mix_sort_index(p), p.name))

        merged = fitz.open()
        for mix_file in mix_files:
            with fitz.open(mix_file) as src:
                merged.insert_pdf(src)

        report_path = section_dir / f"Звіт_розділу_і_мікс_{section_dir.name}.pdf"
        merged.save(report_path)
        merged.close()
        log(f"[=] Зібрано звіт+мікс розділу {section_dir.name}: {report_path.relative_to(out_dir)} "
            f"({len(mix_files)} файлів)")


def build_for_deputies(out_dir: Path, log=print):
    """У кожній "листовій" папці розділу (де реально лежать скачані файли,
    напр. out_dir/1/1.1/1.1.1 чи out_dir/3) створює підпапку "Для депутатів" з
    рівно одним файлом на кожен пункт: Мікс-файлом, якщо для цього пункту
    є додаткові файли, інакше основним файлом пункту."""
    if not out_dir.exists():
        return

    # Стара папка з російською назвою — це згенерована копія (її й так щоразу
    # пересоздавали); якщо лишити, її сприйняло б за "розділ" і вклало б у неї нову.
    for legacy in sorted(p for p in out_dir.rglob("*") if p.is_dir() and p.name in LEGACY_FOR_DEPUTIES_DIR_NAMES):
        shutil.rmtree(legacy)
        log(f"[=] Прибрано стару папку «{legacy.name}»: {legacy.relative_to(out_dir)} (замінено на «{FOR_DEPUTIES_DIR_NAME}»)")

    for section_dir in sorted(p for p in out_dir.rglob("*") if p.is_dir() and p.name != FOR_DEPUTIES_DIR_NAME):
        # Ключ — index_prefix (номер пункту), а не number_tag: номер рішення
        # не унікальний по всьому розділу (нумерація рішень періодично
        # скидається, напр. щороку), тож два різні пункти можуть мати
        # однаковий "681_ПР" — ключ за number_tag тихо затер би один із них.
        main_by_index = {}
        for p in section_dir.iterdir():
            if not p.is_file():
                continue
            match = MAIN_FILENAME_NUMBER_TAG_RE.match(p.name)
            if match:
                index_prefix, number_tag = match.group(1), match.group(2)
                main_by_index[index_prefix] = (number_tag, p)
        if not main_by_index:
            continue

        deputies_dir = section_dir / FOR_DEPUTIES_DIR_NAME
        if deputies_dir.exists():
            shutil.rmtree(deputies_dir)
        deputies_dir.mkdir()

        for index_prefix, (number_tag, main_path) in main_by_index.items():
            mix_path = section_dir / f"Мікс_{section_dir.name}_{index_prefix}_{number_tag}.pdf"
            source = mix_path if mix_path.exists() else main_path

            # Номер пункту в назву — щоб Мікс-файл сортувався у "Для
            # депутатів" на своєму місці за порядком денним, а не за назвою
            # ("Мікс_..." впав би в кінець).
            base_name = re.sub(r"^\d+_", "", source.name)
            dest_name = f"{index_prefix}_{base_name}"
            shutil.copy2(source, deputies_dir / dest_name)

        log(f"[=] Для депутатів: {deputies_dir.relative_to(out_dir)} "
            f"({len(main_by_index)} пунктів)")

        deputies_files = sorted(deputies_dir.iterdir(), key=lambda p: p.name)
        merged = fitz.open()
        for p in deputies_files:
            with fitz.open(p) as src:
                merged.insert_pdf(src)
        merged.save(deputies_dir / "Друк.pdf")
        merged.close()
        log(f"    [=] Друк: {deputies_dir.relative_to(out_dir)}/Друк.pdf ({len(deputies_files)} файлів)")


def build_section_druk(out_dir: Path, log=print):
    """У кожній папці розділу верхнього рівня (out_dir/<major>) створює
    "Друк_розділу_<major>.pdf" — об'єднання файлів "Друк.pdf" з "Для
    депутатів" цього розділу й усіх його підрозділів, у порядку підрозділів
    (1.1, потім 1.2, ...)."""
    if not out_dir.exists():
        return

    def section_sort_key(druk_path: Path):
        # druk_path = .../<major>/[<підрозділи>/]Для депутатів/Друк.pdf
        section_name = druk_path.parent.parent.name
        try:
            return tuple(int(x) for x in section_name.split("."))
        except ValueError:
            return (10**9,)

    for major_dir in sorted(p for p in out_dir.iterdir() if p.is_dir()):
        druk_files = list(major_dir.rglob(f"{FOR_DEPUTIES_DIR_NAME}/Друк.pdf"))
        if not druk_files:
            continue

        druk_files.sort(key=section_sort_key)

        merged = fitz.open()
        for druk_file in druk_files:
            with fitz.open(druk_file) as src:
                merged.insert_pdf(src)

        report_path = major_dir / f"Друк_розділу_{major_dir.name}.pdf"
        merged.save(report_path)
        merged.close()
        log(f"[=] Друк розділу {major_dir.name}: {report_path.relative_to(out_dir)} "
            f"({len(druk_files)} підрозділів)")


def build_item_mix(out_dir: Path, section_dir: Path, main_path: Path, extra_paths: list, index_prefix: str, number_tag: str, log=print):
    """Об'єднує основний файл засідання (завжди першим) з усіма додатковими
    файлами того ж пункту в один pdf "Мікс_<папка>_<index_prefix>_<номер>.pdf".
    index_prefix (номер пункту) потрібен для унікальності: номер рішення
    (number_tag) сам по собі може повторюватись у різні роки.
    Якщо додаткових файлів немає, мікс не створюється (а старий, зайвий, видаляється)."""
    mix_path = section_dir / f"Мікс_{section_dir.name}_{index_prefix}_{number_tag}.pdf"

    if not main_path.exists() or not extra_paths:
        if mix_path.exists():
            mix_path.unlink()
            log(f"    [-] Видалено зайвий мікс без доп. файлів: {mix_path.relative_to(out_dir)}")
        return

    merged = fitz.open()
    for p in [main_path] + extra_paths:
        if p.exists():
            with fitz.open(p) as src:
                merged.insert_pdf(src)
    merged.save(mix_path)
    merged.close()
    log(f"    [=] Мікс: {mix_path.relative_to(out_dir)}")


def download_agenda_pdfs(pdf_path: Path, out_dir: Path, delay: float, log=print, on_audit=None, autofix=False):
    """on_audit — необов'язковий колбек: якщо заданий, зауваження перевірки PDF
    передаються йому списком (kind, section, item, text, extra) і не засмічують
    журнал (туди йде лише короткий рядок); інакше їх повністю виводить log.
    Якщо колбек повертає False — завантаження не починається (чекаємо, поки
    користувач виправить зауваження), функція одразу завершується.
    autofix — правити PDF прямо у файлі (зараз: вирівнювання зсунутих посилань)
    ще до перевірки; у GUI це копія порядку денного в папці результатів."""
    if autofix:
        fix_link_misalignment(pdf_path, log)
    audit_problems = audit_agenda(pdf_path)
    if on_audit is None:
        log_audit(audit_problems, log)
    else:
        hold = on_audit(audit_problems) is False
        if hold:
            log(f"Перевірка порядку денного: знайдено {len(audit_problems)} зауважень — див. блок «Зауваження до PDF».")
            log("Завантаження не розпочато: спершу виправте зауваження або позначте їх галочкою — "
                "після цього файл буде перевірено ще раз і завантаження почнеться.")
            return 0, 0, []
        log("Перевірка порядку денного: " + (
            "помилок не знайдено." if not audit_problems
            else f"зауваження ({len(audit_problems)}) опрацьовані — продовжуємо."))
    items, unresolved = extract_agenda_items(pdf_path)
    total_found = len(items) + len(unresolved)
    log(f"Знайдено посилань на проєкти рішень: {total_found}")

    def section_label(section):
        return "/".join(section) if section else "?"

    skipped = [
        (item_number, section_label(section), uri,
         "не вдалося розпізнати дату/номер рішення в тексті поруч із посиланням "
         "(ймовірно, помилка чи нестандартне форматування в самому документі)")
        for uri, section, item_number in unresolved
    ]
    if unresolved:
        log(f"Не вдалося визначити дату/номер для {len(unresolved)} посилань (пропущено):")
        for item_number, section, uri, reason in skipped:
            item_label = f"пункт {item_number}" if item_number is not None else "пункт ?"
            log(f"  - {item_label} (розділ {section}): {uri}")

    out_dir.mkdir(parents=True, exist_ok=True)
    remove_legacy_artifacts(out_dir, log)
    session = requests.Session()
    session.headers.update(HEADERS)

    index_width = max(3, len(str(len(items))))

    ok, failed = 0, []
    for position, (uri, date, number, section, item_number) in enumerate(items, start=1):
        index_prefix = f"{item_number:0{index_width}d}" if item_number is not None else f"{position:0{index_width}d}"
        filename = f"{index_prefix}_{date}_{number.replace('/', '_')}.pdf"
        section_dir = out_dir.joinpath(*section) if section else out_dir / "unsorted"
        section_dir.mkdir(parents=True, exist_ok=True)
        out_path = section_dir / filename
        rel_path = out_path.relative_to(out_dir)

        main_already_exists = out_path.exists()
        if main_already_exists:
            pad_pdf_to_multiple_of_4(out_path, log)

        page_url = "https://" + uri if not uri.startswith("http") else uri
        try:
            page_resp = fetch_page(session, page_url)

            all_files = find_all_files(page_resp.text, page_resp.url)

            if main_already_exists:
                log(f"[=] Вже існує, пропускаю: {rel_path}")
            else:
                main_url = next(
                    (url for label, url in all_files if label.strip().casefold() == SECOND_READING_TEXT.casefold()),
                    None,
                )
                if not main_url:
                    main_url = next(
                        (url for label, url in all_files if label.strip().casefold() == TARGET_TEXT.casefold()),
                        None,
                    )
                if not main_url:
                    raise RuntimeError(f"файл з підписом '{TARGET_TEXT}' не знайдено")

                file_resp = session.get(main_url, timeout=60)
                file_resp.raise_for_status()

                out_path.write_bytes(file_resp.content)
                log(f"[+] {rel_path}  <-  {page_url}")
                pad_pdf_to_multiple_of_4(out_path, log)

            number_tag = number.replace("/", "_")
            seen_extra_names = {}
            extra_paths = []
            for label, extra_url in all_files:
                if not any(target.casefold() in label.casefold() for target in ADDITIONAL_TARGET_TEXTS):
                    continue

                # Номер рішення (number_tag) не унікальний по всьому розділу —
                # нумерація рішень періодично скидається (напр. щороку), тож
                # той самий "681_ПР" може належати двом різним пунктам. Додаємо
                # index_prefix пункту, щоб файли різних пунктів не затирали
                # один одного.
                base_name = f"{sanitize_filename(label)}_{index_prefix}_{number_tag}"
                dup_count = seen_extra_names.get(base_name, 0)
                candidate_name = base_name if dup_count == 0 else f"{base_name}({dup_count})"
                seen_extra_names[base_name] = dup_count + 1

                extra_path = section_dir / f"{candidate_name}.pdf"
                extra_paths.append(extra_path)
                if extra_path.exists():
                    continue

                extra_resp = session.get(extra_url, timeout=60)
                extra_resp.raise_for_status()
                extra_path.write_bytes(extra_resp.content)
                log(f"    [+] {extra_path.relative_to(out_dir)}")
                pad_pdf_to_multiple_of_4(extra_path, log)

            build_item_mix(out_dir, section_dir, out_path, extra_paths, index_prefix, number_tag, log)

            ok += 1
        except Exception as exc:
            log(f"[!] Помилка для {page_url}: {exc}")
            failed.append((page_url, str(exc)))
            skipped.append((item_number, section_label(section), page_url, str(exc)))

        time.sleep(delay)

    log(f"\nГотово: завантажено/наявно {ok} з {len(items)}, помилок: {len(failed)}")
    if failed:
        log("Список помилок:")
        for url, err in failed:
            log(f"  - {url}: {err}")

    if total_found > ok:
        skipped.sort(key=lambda s: (s[0] is None, s[0] if s[0] is not None else 0))
        log(f"\nЗнайдено {total_found} пунктів, скачано {ok}. Пропущено {len(skipped)}:")
        for item_number, section, uri, reason in skipped:
            item_label = f"пункт {item_number}" if item_number is not None else "пункт ?"
            log(f"  - {item_label} (розділ {section}): {reason} — {uri}")

    if audit_problems and on_audit is None:
        log(f"\nУвага: перевірка порядку денного знайшла {len(audit_problems)} зауважень у самому PDF "
            f"(див. початок журналу) — перегляньте їх.")
        for kind, sec, num, text, _ in audit_problems:
            if kind in ("немає посилання", "порушена нумерація", "посилання без пункту"):
                log(f"  - [{kind}] розділ {sec}" + (f", пункт {num}" if num is not None else ""))

    build_section_reports(out_dir, log)
    build_section_mix_reports(out_dir, log)
    build_for_deputies(out_dir, log)
    build_section_druk(out_dir, log)

    return ok, len(items), failed


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdf", type=Path, help="Шлях до PDF порядку денного")
    parser.add_argument(
        "--out-dir", type=Path, default=Path("downloaded_plenary_pdfs"),
        help="Куди зберігати завантажені файли (за замовчуванням ./downloaded_plenary_pdfs)",
    )
    parser.add_argument(
        "--delay", type=float, default=0.3,
        help="Пауза між запитами в секундах (за замовчуванням 0.3)",
    )
    args = parser.parse_args()

    if not args.pdf.exists():
        print(f"Файл не знайдено: {args.pdf}")
        sys.exit(1)

    download_agenda_pdfs(args.pdf, args.out_dir, args.delay)


if __name__ == "__main__":
    main()
