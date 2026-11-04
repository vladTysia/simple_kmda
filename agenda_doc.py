"""Порядок денний як документ: читання з PDF або Word і вставка нових пунктів.

Для вкладки "Генерація ПД": користувач може додати до проєктів рішень ще й файл порядку денного.
З нього беруться розділи (заголовки "1.", "1.1.", … до 4 рівнів) і пункти ("1) Про …") — вони
стають готовими картками, а нові пункти вставляються у відповідні місця документа.

  * Word (.docx)  -> правиться сам документ: нові абзаци вставляються поруч із наявними й копіюють
                     їхнє форматування; решта документа (і гіперпосилання старих пунктів) не чіпається.
  * PDF           -> PDF не зберігає абзаци, тому з нього відновлюється Word-документ: абзаци, шрифти,
                     відступи, вирівнювання й гіперпосилання збираються з розмітки сторінок
                     (вигляд близький до оригіналу, але не ідентичний).
"""

import copy
import re
from collections import Counter
from pathlib import Path

import fitz  # PyMuPDF

PROJECT_URI_RE = re.compile(r"kmr\.gov\.ua/", re.IGNORECASE)
HEADING_RE = re.compile(r"^(\d+(?:\.\d+){0,3})\.\s+(\S.*)$", re.DOTALL)
ITEM_RE = re.compile(r"^(\d+)\)\s*(.*)$", re.DOTALL)
MAX_DEPTH = 4


def _section_key(section_id: str) -> list:
    return [int(p) for p in section_id.split(".")]


# ------------------------------------------------------------------------------------------
# PDF -> абзаци
# ------------------------------------------------------------------------------------------

def _page_rows(page):
    """Візуальні рядки сторінки: спани з однаковою базовою лінією, зліва направо. У PDF з Word
    вирівняний по ширині рядок часто розбитий на окремі слова — тут вони знову збираються."""
    links = [(fitz.Rect(l["from"]), l["uri"]) for l in page.get_links()
             if l.get("kind") == fitz.LINK_URI and l.get("uri")]
    spans = []
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            for sp in line["spans"]:
                if sp["text"] == "":
                    continue
                x0, y0, x1, y1 = sp["bbox"]
                cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
                url = next((u for r, u in links if r.x0 <= cx <= r.x1 and r.y0 <= cy <= r.y1), None)
                spans.append({"x0": x0, "x1": x1, "base": sp["origin"][1], "y0": y0, "text": sp["text"],
                              "size": sp["size"], "font": sp["font"], "bold": "bold" in sp["font"].lower() or bool(sp["flags"] & 16),
                              "url": url})
    spans.sort(key=lambda s: (s["base"], s["x0"]))
    rows, cur = [], []
    for sp in spans:
        # окремі гліфи (напр. "(") можуть стояти на 3-5 pt вище/нижче базової лінії рядка — це той самий рядок
        if cur and sp["base"] - cur[0]["base"] > 6.0:
            rows.append(cur)
            cur = []
        cur.append(sp)
    if cur:
        rows.append(cur)
    out = []
    for r in rows:
        r.sort(key=lambda s: s["x0"])
        out.append(r)
    return out


def _row_runs(spans):
    """Рядок -> [(текст, жирний, url)]: сусідні слова з однаковими властивостями зливаються."""
    runs = []
    prev = None
    for sp in spans:
        text = sp["text"]
        gap = sp["x0"] - prev["x1"] if prev else 0
        if prev and gap > 1.0 and not prev["text"].endswith(" ") and not text.startswith(" "):
            sep = " "
            # пробіл між словами належить до посилання, лише якщо обидва слова в ньому
            attrs = (prev["bold"], prev["url"]) if prev["url"] == sp["url"] else (prev["bold"], None)
            if runs and (runs[-1][1], runs[-1][2]) == attrs:
                runs[-1] = (runs[-1][0] + sep, runs[-1][1], runs[-1][2])
            else:
                runs.append((sep, attrs[0], attrs[1]))
        if runs and (runs[-1][1], runs[-1][2]) == (sp["bold"], sp["url"]):
            runs[-1] = (runs[-1][0] + text, sp["bold"], sp["url"])
        else:
            runs.append((text, sp["bold"], sp["url"]))
        prev = sp
    return runs


def pdf_rows(doc):
    """Усі рядки документа підряд (без номерів сторінок унизу), зі службовими полями."""
    rows = []
    for pno in range(len(doc)):
        page = doc[pno]
        height = page.rect.height
        for spans in _page_rows(page):
            text = "".join(s["text"] for s in spans)
            if text.strip().isdigit() and spans[0]["base"] > height - 90 and spans[0]["size"] <= 12.5:
                continue  # номер сторінки (у Word він — колонтитул, який дописується заново)
            rows.append({
                "page": pno, "base": spans[0]["base"], "x0": min(s["x0"] for s in spans), "x1": max(s["x1"] for s in spans),
                "size": max(s["size"] for s in spans), "runs": _row_runs(spans),
                "text": "".join(r[0] for r in _row_runs(spans)),
                "bold": all(s["bold"] for s in spans if s["text"].strip()),
            })
    return rows


def _mode(values, default):
    values = [round(v, 1) for v in values]
    return Counter(values).most_common(1)[0][0] if values else default


def _starts_item(row, left, strict=True):
    """Початок пункту "N) …": номер стоїть із відступом першого рядка. Без цієї умови рядок, що
    почався з перенесеного "2019) (464010216)" (розрив усередині дати), помилково став би пунктом."""
    if strict and row["x0"] <= left + 15:
        return False
    return bool(re.match(r"^\d+\)(\s|$)", row["text"].lstrip()))


def _starts_heading(row, left, strict=True):
    if strict and row["x0"] <= left + 15:
        return False
    return row["bold"] and bool(HEADING_RE.match(row["text"].lstrip()))


def pdf_to_blocks(path, strict=True) -> dict:
    """Відновлює структуру порядку денного з PDF.

    Повертає {"blocks": [...], "layout": {...}}, де кожен блок — абзац:
      kind: "heading" | "item" | "other"; section: id розділу ("3.1.1"); number: номер пункту;
      runs: [(текст, жирний, url)] (для пункту/заголовка — без номера й табуляції);
      fmt: вирівнювання, відступи, інтервал, розмір шрифту.
    """
    doc = fitz.open(path)
    try:
        rows = pdf_rows(doc)
        page_rect = doc[0].rect
    finally:
        doc.close()
    if not rows:
        raise ValueError("У PDF не знайдено тексту — це скан? Потрібен PDF з текстовим шаром.")

    left = _mode([r["x0"] for r in rows], 70.0)
    right = _mode([r["x1"] for r in rows], page_rect.width - 40)
    pitch = _mode([b["base"] - a["base"] for a, b in zip(rows, rows[1:]) if a["page"] == b["page"] and 10 < b["base"] - a["base"] < 25], 17.0)
    centre = (left + right) / 2

    # --- рядки -> абзаци ---
    paragraphs, cur = [], None
    for i, row in enumerate(rows):
        prev = rows[i - 1] if i else None
        indented = row["x0"] > left + 15
        centred = abs((row["x0"] + row["x1"]) / 2 - centre) < 8 and row["x0"] > left + 30 and row["x1"] < right - 30
        new = (cur is None or _starts_item(row, left, strict) or _starts_heading(row, left, strict)
               or centred or cur["centred"])
        if not new:
            gap = row["base"] - prev["base"]
            short_prev = prev["x1"] < right - 30
            if cur["block_indent"]:
                new = abs(row["x0"] - cur["rows"][0]["x0"]) > 5            # блок із відступом ліворуч (напр. "Попередній проєкт…")
            elif indented:
                new = True                                                  # нова червона стрічка
            elif prev["page"] == row["page"] and gap > pitch * 1.5:
                new = True                                                  # порожній проміжок між абзацами
            elif short_prev:
                new = True                                                  # попередній рядок короткий, а цей без відступу
        if new:
            cur = {"rows": [row], "centred": centred, "block_indent": row["x0"] > left + 60 and not centred}
            paragraphs.append(cur)
        else:
            cur["rows"].append(row)

    # --- абзаци -> блоки ---
    blocks, section = [], None
    for k, par in enumerate(paragraphs):
        rws = par["rows"]
        first = rws[0]
        flat = _join_rows(rws)
        text = "".join(r[0] for r in flat)
        lead = text.lstrip()
        offset = len(text) - len(lead)
        kind, number, sec_id = "other", None, None
        if _starts_item(first, left, strict):
            m = re.match(r"^(\d+)\)\s*", lead)
            kind, number, sec_id = "item", int(m.group(1)), section
            flat = _strip_prefix(flat, offset + m.end())
        elif _starts_heading(first, left, strict):
            m = re.match(r"^(\d+(?:\.\d+){0,3})\.\s*", lead)
            kind, sec_id = "heading", m.group(1)
            section = sec_id
            flat = _strip_prefix(flat, offset + m.end())
        multi = len(rws) > 1
        block_indent = par["block_indent"]
        first_indent = 0.0 if (block_indent or par["centred"]) else max(0.0, first["x0"] - left)
        left_indent = (first["x0"] - left) if block_indent else 0.0
        justified = kind in ("item", "heading") or (multi and rws[0]["x1"] >= right - 3)
        align = "center" if par["centred"] else ("justify" if justified and not block_indent else "left")
        space_before = 0.0
        if k:
            pb = paragraphs[k - 1]["rows"][-1]
            if pb["page"] == first["page"]:
                space_before = max(0.0, first["base"] - pb["base"] - pitch)
                space_before = space_before if space_before >= 2 else 0.0
        blocks.append({
            "kind": kind, "section": sec_id, "number": number, "runs": flat, "page": first["page"],
            "fmt": {"align": align, "first_indent": first_indent, "left_indent": left_indent,
                    "space_before": space_before, "size": first["size"]},
        })
    if strict and not any(b["kind"] == "item" for b in blocks):
        return pdf_to_blocks(path, strict=False)
    return {"blocks": blocks, "layout": {"left": left, "right": right, "pitch": pitch,
                                          "page": (page_rect.width, page_rect.height)}}


def _join_rows(rws):
    """Рядки абзацу -> runs через пробіл; пробіл між рядками входить у посилання, лише якщо воно триває."""
    flat = []

    def push(text, bold, url):
        if not text:
            return
        if flat and (flat[-1][1], flat[-1][2]) == (bold, url):
            flat[-1] = (flat[-1][0] + text, bold, url)
        else:
            flat.append((text, bold, url))

    for i, row in enumerate(rws):
        if i and flat:
            last, nxt = flat[-1], row["runs"][0]
            push(" ", last[1], last[2] if last[2] == nxt[2] else None)
        for r in row["runs"]:
            push(*r)
    return flat


def _strip_prefix(runs, count):
    """Прибирає перші count символів (номер "1)" / "1.1." і пробіли після нього) з серії runs."""
    out, left = [], count
    for text, bold, url in runs:
        if left >= len(text):
            left -= len(text)
            continue
        out.append((text[left:], bold, url))
        left = 0
    return out


# ------------------------------------------------------------------------------------------
# Спільна логіка: вставка нових пунктів у модель документа
# ------------------------------------------------------------------------------------------

def _squash(text: str) -> str:
    """Текст пункту для карток: пробіли й табуляції — в один пробіл."""
    return re.sub(r"\s+", " ", text).strip()


def _heading_depth(section_id: str) -> int:
    return len(section_id.split("."))


def _branch_nodes(nodes, section_id):
    prefix = section_id + "."
    return [n for n in nodes if n["kind"] in ("heading", "item") and n["section"] and
            (n["section"] == section_id or n["section"].startswith(prefix))]


def _heading_node(nodes, section_id):
    return next((n for n in nodes if n["kind"] == "heading" and n["section"] == section_id), None)


def _ensure_heading(model, section_id, titles):
    """Заголовок розділу, якого ще немає в документі (створений користувачем підрозділ), —
    ставиться після гілки попереднього підрозділу того ж рівня або після власних пунктів батька."""
    head = _heading_node(model.nodes, section_id)
    if head is not None:
        return head
    parent = section_id.rsplit(".", 1)[0] if "." in section_id else None
    if parent:
        _ensure_heading(model, parent, titles)
    siblings = [n["section"] for n in model.nodes if n["kind"] == "heading" and n["section"] != section_id and
                (n["section"].rsplit(".", 1)[0] if "." in n["section"] else None) == parent and
                _section_key(n["section"]) < _section_key(section_id)]
    if siblings:
        anchor = _branch_nodes(model.nodes, max(siblings, key=_section_key))[-1]
    elif parent:
        own = [n for n in model.nodes if n["kind"] in ("heading", "item") and n["section"] == parent]
        anchor = own[-1]
    else:  # новий розділ верхнього рівня — у кінець останньої гілки
        branches = [n for n in model.nodes if n["kind"] in ("heading", "item")]
        anchor = branches[-1] if branches else None
    same_depth = [n for n in model.nodes if n["kind"] == "heading" and _heading_depth(n["section"]) == _heading_depth(section_id)]
    template = same_depth[-1] if same_depth else next((n for n in model.nodes if n["kind"] == "heading"), None)
    title = titles.get(section_id, "").strip()
    if template is not None and model.heading_text(template).rstrip().endswith(":") and not title.endswith(":"):
        title += ":"
    node = model.new_heading(section_id, title, template)
    model.insert_after(anchor, node)
    return node


def _item_template(model, section_id, anchor):
    if anchor is not None and anchor["kind"] == "item":
        return anchor
    own = [n for n in model.nodes if n["kind"] == "item" and n["section"] == section_id]
    return own[0] if own else next((n for n in model.nodes if n["kind"] == "item"), None)


def _renumber(model, section_id):
    """Нумерація "1) 2) 3) …" у групі (заголовок і його власні пункти) знову йде підряд."""
    head = _heading_node(model.nodes, section_id)
    if head is None:
        return
    number = 0
    for node in model.nodes[model.nodes.index(head) + 1:]:
        if node["kind"] == "heading":
            break
        if node["kind"] == "item":
            number += 1
            model.set_number(node, number)


def apply_new_items(model, placements: dict, titles: dict):
    """placements[розділ] = [("agenda", індекс пункту порядку денного) | ("new", текст, url)] — те, як
    пункти розділу мають іти в підсумку. Нові пункти вставляються після попереднього пункту порядку
    денного (або після заголовка розділу), старі лишаються на своїх місцях; у розділах, куди щось
    додали, номери "N)" перераховуються."""
    items = {n["agenda_index"]: n for n in model.nodes if n["kind"] == "item"}
    touched = []
    for sid in sorted(placements, key=_section_key):
        entries = placements[sid]
        if not any(e[0] == "new" for e in entries):
            continue
        anchor = _ensure_heading(model, sid, titles)
        for entry in entries:
            if entry[0] == "agenda":
                anchor = items[entry[1]]
            else:
                node = model.new_item(entry[1], entry[2], _item_template(model, sid, anchor), sid)
                model.insert_after(anchor, node)
                anchor = node
        touched.append(sid)
    for sid in touched:
        _renumber(model, sid)


# ------------------------------------------------------------------------------------------
# Модель з PDF (відновлення Word-документа)
# ------------------------------------------------------------------------------------------

class PdfModel:
    kind = "pdf"

    def __init__(self, parsed: dict):
        self.layout = parsed["layout"]
        self.nodes = parsed["blocks"]
        index = 0
        for node in self.nodes:
            if node["kind"] == "item":
                node["agenda_index"] = index
                index += 1

    @classmethod
    def load(cls, path):
        return cls(pdf_to_blocks(path))

    # --- для карток ---
    def sections(self) -> dict:
        return {n["section"]: _squash("".join(r[0] for r in n["runs"])) for n in self.nodes if n["kind"] == "heading"}

    def items(self) -> list:
        out = []
        for n in self.nodes:
            if n["kind"] == "item":
                url = next((r[2] for r in n["runs"] if r[2]), None)
                out.append({"index": n["agenda_index"], "section": n["section"],
                            "text": _squash("".join(r[0] for r in n["runs"])), "url": url, "number": n["number"]})
        return out

    def heading_text(self, node) -> str:
        return "".join(r[0] for r in node["runs"])

    # --- правка ---
    def insert_after(self, anchor, node):
        self.nodes.insert(self.nodes.index(anchor) + 1 if anchor is not None else 0, node)

    def new_item(self, text, url, template, section_id):
        fmt = dict(template["fmt"]) if template else {"align": "justify", "first_indent": 35.4, "left_indent": 0.0,
                                                    "space_before": 0.0, "size": 14.0}
        fmt["space_before"] = 0.0
        return {"kind": "item", "section": section_id, "number": None, "runs": [(text, False, url)], "fmt": fmt, "page": None}

    def new_heading(self, section_id, title, template):
        fmt = dict(template["fmt"]) if template else {"align": "justify", "first_indent": 35.4, "left_indent": 0.0,
                                                    "space_before": 0.0, "size": 14.0}
        return {"kind": "heading", "section": section_id, "number": None, "runs": [(title, True, None)], "fmt": fmt, "page": None}

    def set_number(self, node, number):
        node["number"] = number

    # --- збереження ---
    def save(self, path):
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Pt

        width, height = self.layout["page"]
        doc = Document()
        sec = doc.sections[0]
        sec.page_width, sec.page_height = Pt(width), Pt(height)
        sec.left_margin, sec.right_margin = Pt(self.layout["left"]), Pt(width - self.layout["right"])
        sec.top_margin, sec.bottom_margin = Pt(42.5), Pt(56.7)
        sec.footer_distance = Pt(28)
        normal = doc.styles["Normal"]
        normal.font.name = "Times New Roman"
        normal.font.size = Pt(14)
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        normal.paragraph_format.space_after = Pt(0)
        normal.paragraph_format.space_before = Pt(0)

        def style_run(run, bold, size):
            run.bold = bool(bold)
            run.font.name = "Times New Roman"
            run.font.size = Pt(size)

        def add_link(par, url, text, bold, size):
            rid = par.part.relate_to(url, RT.HYPERLINK, is_external=True)
            link = OxmlElement("w:hyperlink")
            link.set(qn("r:id"), rid)
            run = OxmlElement("w:r")
            rpr = OxmlElement("w:rPr")
            fonts = OxmlElement("w:rFonts")
            for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                fonts.set(qn(attr), "Times New Roman")
            rpr.append(fonts)
            if bold:
                rpr.append(OxmlElement("w:b"))
            sz = OxmlElement("w:sz")
            sz.set(qn("w:val"), str(int(round(size * 2))))
            rpr.append(sz)
            run.append(rpr)
            t = OxmlElement("w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = text
            run.append(t)
            link.append(run)
            par._p.append(link)

        aligns = {"justify": WD_ALIGN_PARAGRAPH.JUSTIFY, "center": WD_ALIGN_PARAGRAPH.CENTER, "left": WD_ALIGN_PARAGRAPH.LEFT}
        for node in self.nodes:
            fmt = node["fmt"]
            par = doc.add_paragraph()
            pf = par.paragraph_format
            pf.alignment = aligns.get(fmt["align"], WD_ALIGN_PARAGRAPH.LEFT)
            pf.first_line_indent = Pt(fmt["first_indent"]) if fmt["first_indent"] else None
            pf.left_indent = Pt(fmt["left_indent"]) if fmt["left_indent"] else None
            pf.space_before = Pt(fmt["space_before"]) if fmt["space_before"] else None
            pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
            pf.line_spacing = Pt(self.layout["pitch"])
            size = fmt["size"]
            if node["kind"] == "item":
                style_run(par.add_run(f"{node['number']})\t"), False, size)
            elif node["kind"] == "heading":
                style_run(par.add_run(f"{node['section']}.\t"), True, size)
            for text, bold, url in node["runs"]:
                if url:
                    add_link(par, url, text, bold, size)
                elif text:
                    style_run(par.add_run(text), bold, size)

        # номер сторінки внизу по центру (як в оригіналі)
        fp = sec.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = fp.add_run()
        run.font.size, run.font.name = Pt(12), "Times New Roman"
        for kind, text in (("begin", None), (None, " PAGE "), ("end", None)):
            if kind:
                el = OxmlElement("w:fldChar")
                el.set(qn("w:fldCharType"), kind)
            else:
                el = OxmlElement("w:instrText")
                el.set(qn("xml:space"), "preserve")
                el.text = text
            run._r.append(el)
        doc.save(path)


def load_agenda(path):
    """Модель порядку денного з .docx (правиться сам файл) або .pdf (відновлюється Word-документ)."""
    suffix = Path(path).suffix.lower()
    if suffix == ".pdf":
        return PdfModel.load(path)
    if suffix == ".docx":
        return DocxModel.load(path)
    raise ValueError("Порядок денний має бути у форматі .docx або .pdf.")


# ------------------------------------------------------------------------------------------
# Модель із Word (.docx): правиться сам документ
# ------------------------------------------------------------------------------------------

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_FIELD_URL_RE = re.compile(r'HYPERLINK\s+"([^"]+)"', re.IGNORECASE)


def _p_text(p_el) -> str:
    out = []
    for el in p_el.iter():
        if el.tag == _W + "t":
            out.append(el.text or "")
        elif el.tag == _W + "tab" and el.getparent().tag != _W + "tabs":
            out.append("\t")
    return "".join(out)


def _p_runs(p_el) -> list:
    """Усі run абзацу в порядку документа (разом із тими, що в гіперпосиланнях)."""
    runs = []
    for r in p_el.iter(_W + "r"):
        text = "".join((t.text or "") if t.tag == _W + "t" else "\t" for t in r if t.tag in (_W + "t", _W + "tab"))
        runs.append({"el": r, "text": text, "in_link": r.getparent().tag == _W + "hyperlink"})
    return runs


def _is_bold(rpr) -> bool:
    if rpr is None:
        return False
    b = rpr.find(_W + "b")
    return b is not None and b.get(_W + "val") not in ("0", "false")


def _p_bold(p_el) -> bool:
    for run in _p_runs(p_el):
        if run["text"].strip():
            return _is_bold(run["el"].find(_W + "rPr"))
    return False


def _p_urls(part, p_el) -> list:
    urls = []
    for h in p_el.iter(_W + "hyperlink"):
        rid = h.get(_R + "id")
        if rid and rid in part.rels and part.rels[rid].is_external:
            urls.append(part.rels[rid].target_ref)
    for f in p_el.iter(_W + "fldSimple"):
        m = _FIELD_URL_RE.search(f.get(_W + "instr") or "")
        if m:
            urls.append(m.group(1))
    for it in p_el.iter(_W + "instrText"):
        m = _FIELD_URL_RE.search(it.text or "")
        if m:
            urls.append(m.group(1))
    return urls


class DocxModel:
    kind = "docx"

    def __init__(self, doc, nodes):
        self.doc = doc
        self.nodes = nodes

    @classmethod
    def load(cls, path):
        from docx import Document
        try:
            doc = Document(str(path))
        except Exception as exc:
            raise ValueError("Не вдалося відкрити Word-файл — це справжній .docx?") from exc
        nodes, section, index = [], None, 0
        for p in doc.paragraphs:
            el = p._p
            text = _p_text(el)
            lead = text.lstrip()
            node = {"kind": "other", "section": None, "number": None, "el": el}
            m_item = re.match(r"^(\d+)\)", lead)
            m_head = HEADING_RE.match(lead)
            if m_item:
                node.update(kind="item", section=section, number=int(m_item.group(1)), agenda_index=index)
                index += 1
            elif m_head and (_p_bold(el) or text.rstrip().endswith(":")):
                section = m_head.group(1)
                node.update(kind="heading", section=section)
            nodes.append(node)
        if not any(n["kind"] == "item" for n in nodes):
            auto = any(p._p.find(_W + "pPr") is not None and p._p.find(_W + "pPr").find(_W + "numPr") is not None
                       for p in doc.paragraphs)
            raise ValueError(
                "У Word-файлі не знайдено пунктів виду «1) Про …»."
                + (" Схоже, нумерація там автоматична (список Word) — програма читає лише номери, набрані текстом. "
                   "Збережіть документ у PDF і завантажте його, або перетворіть нумерацію на звичайний текст." if auto else ""))
        return cls(doc, nodes)

    # --- для карток ---
    def _title(self, node) -> str:
        text = _p_text(node["el"]).lstrip()
        text = re.sub(r"^\d+\)\s*", "", text) if node["kind"] == "item" else HEADING_RE.sub(lambda m: m.group(2), text)
        return re.sub(r"\s+", " ", text).strip()

    def sections(self) -> dict:
        return {n["section"]: self._title(n) for n in self.nodes if n["kind"] == "heading"}

    def items(self) -> list:
        out = []
        for n in self.nodes:
            if n["kind"] == "item":
                urls = _p_urls(self.doc.part, n["el"])
                out.append({"index": n["agenda_index"], "section": n["section"], "text": self._title(n),
                            "url": urls[0] if urls else None, "number": n["number"]})
        return out

    def heading_text(self, node) -> str:
        return self._title(node)

    # --- правка ---
    def insert_after(self, anchor, node):
        anchor["el"].addnext(node["el"])
        self.nodes.insert(self.nodes.index(anchor) + 1, node)

    @staticmethod
    def _rpr(run):
        rpr = run["el"].find(_W + "rPr") if run else None
        return copy.deepcopy(rpr) if rpr is not None else None

    def _make_run(self, rpr, text=None, tab=False):
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        r = OxmlElement("w:r")
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        if tab:
            r.append(OxmlElement("w:tab"))
        else:
            t = OxmlElement("w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = text
            r.append(t)
        return r

    def new_item(self, text, url, template, section_id):
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        p = OxmlElement("w:p")
        num_rpr = text_rpr = None
        has_tab = True
        if template is not None:
            tpl = template["el"]
            ppr = tpl.find(_W + "pPr")
            if ppr is not None:
                p.append(copy.deepcopy(ppr))
            runs = [r for r in _p_runs(tpl)]
            num_rpr = self._rpr(runs[0]) if runs else None
            link_run = next((r for r in runs if r["in_link"]), None)
            if link_run is None:
                link_run = next((r for r in runs[1:] if r["text"].strip()), runs[0] if runs else None)
            text_rpr = self._rpr(link_run)
            has_tab = bool(re.match(r"^\s*\d+\)\s*\t", _p_text(tpl))) or "\t" in _p_text(tpl)[:12]
        p.append(self._make_run(num_rpr, "0)"))
        p.append(self._make_run(num_rpr, tab=True) if has_tab else self._make_run(num_rpr, " "))
        link = OxmlElement("w:hyperlink")
        if url:
            link.set(qn("r:id"), self.doc.part.relate_to(url, RT.HYPERLINK, is_external=True))
            link.append(self._make_run(text_rpr, text))
            p.append(link)
        else:
            p.append(self._make_run(text_rpr, text))
        return {"kind": "item", "section": section_id, "number": None, "el": p}

    def new_heading(self, section_id, title, template):
        from docx.oxml import OxmlElement
        p = OxmlElement("w:p")
        label_rpr = title_rpr = None
        if template is not None:
            tpl = template["el"]
            ppr = tpl.find(_W + "pPr")
            if ppr is not None:
                p.append(copy.deepcopy(ppr))
            runs = _p_runs(tpl)
            label_rpr = self._rpr(runs[0]) if runs else None
            title_rpr = self._rpr(next((r for r in runs[1:] if r["text"].strip()), runs[0] if runs else None))
        p.append(self._make_run(label_rpr, f"{section_id}."))
        p.append(self._make_run(label_rpr, tab=True))
        p.append(self._make_run(title_rpr, title))
        return {"kind": "heading", "section": section_id, "number": None, "el": p}

    def set_number(self, node, number):
        node["number"] = number
        for t in node["el"].iter(_W + "t"):
            if t.text and re.match(r"^\s*\d+\)", t.text):
                t.text = re.sub(r"^(\s*)\d+\)", lambda m: f"{m.group(1)}{number})", t.text, count=1)
                return
            if t.text and re.match(r"^\s*\d+$", t.text):  # "1" і ")" у різних run
                t.text = re.sub(r"\d+", str(number), t.text, count=1)
                return

    def save(self, path):
        self.doc.save(str(path))


def looks_like_agenda_pdf(path, min_links: int = 3) -> bool:
    """PDF — порядок денний, якщо в ньому є кілька гіперпосилань на сторінки проєктів рішень kmr.gov.ua."""
    try:
        doc = fitz.open(path)
    except Exception:
        return False
    try:
        count = 0
        for page in doc:
            count += sum(1 for l in page.get_links() if PROJECT_URI_RE.search(l.get("uri") or ""))
            if count >= min_links:
                return True
        return False
    finally:
        doc.close()
