# ============================================================
#  kocderma.com — Belge Üretim Sunucusu (FastAPI)
#  Claude içeriği JSON şema olarak üretir; bu sunucu gerçek
#  docx / xlsx / pptx / pdf dosyasını kendi üretir ve indirtir.
#  Hiçbir beta/özel API özelliğine bağlı değildir.
#
#  Ortam değişkeni: ANTHROPIC_API_KEY (Render'da Secret olarak ekle)
# ============================================================

import os
import re
import json
import uuid
import time

import requests
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem, Table, TableStyle,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
MODEL = "claude-sonnet-4-6"
FILES_DIR = os.environ.get("FILES_DIR", "/tmp/genfiles")
os.makedirs(FILES_DIR, exist_ok=True)

ALLOWED_ORIGINS = [
    "https://kocderma.com",
    "https://www.kocderma.com",
]

# Türkçe karakterler için Unicode font (PDF). Linux'ta DejaVu genelde kuruludur.
PDF_FONT = "Helvetica"
for _fp in [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
]:
    if os.path.exists(_fp):
        try:
            pdfmetrics.registerFont(TTFont("DejaVu", _fp))
            PDF_FONT = "DejaVu"
            break
        except Exception:
            pass

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_methods=["*"],
    allow_headers=["*"],
)

SYSTEM = """Sen bir belge üreticisisin. Kullanıcının isteğine göre TEK BİR JSON nesnesi üret.
Yanıtın SADECE JSON olsun: açıklama yazma, kod bloğu işareti (```) kullanma.

Şema:
{
  "kind": "docx" | "xlsx" | "pptx" | "pdf",
  "filename": "uygun-ad.uzanti",
  "title": "Belge başlığı",
  "summary": "Kullanıcıya gösterilecek 1 cümlelik Türkçe özet",
  "blocks": [
     {"type":"heading","level":1,"text":"..."},
     {"type":"paragraph","text":"..."},
     {"type":"bullets","items":["...","..."]},
     {"type":"numbered","items":["...","..."]},
     {"type":"table","headers":["..."],"rows":[["...","..."]]}
  ],
  "sheets": [ {"name":"Sayfa1","headers":["..."],"rows":[["..."]]} ],
  "slides": [ {"title":"...","bullets":["...","..."]} ]
}

Kurallar:
- Kullanıcı format belirttiyse (Word=docx, Excel=xlsx, sunum/PowerPoint=pptx, PDF=pdf) onu kullan; belirtmediyse içeriğe en uygun olanı seç.
- docx ve pdf için "blocks" doldur. xlsx için "sheets". pptx için "slides".
- İçeriği Türkçe, dolu ve dermatoloji eğitimine uygun hazırla; önemli İngilizce terimleri parantez içinde ekle.
- Bu bir eğitim aracıdır; tanı koyma, kişiye özel tıbbi tavsiye verme.
- SADECE geçerli JSON döndür."""


class DocReq(BaseModel):
    messages: list


def call_claude(messages):
    body = {"model": MODEL, "max_tokens": 8192, "system": SYSTEM, "messages": messages}
    last = "Claude API hatasi"
    for attempt in range(4):
        try:
            r = requests.post(
                "https://api.anthropic.com/v1/messages",
                headers={
                    "x-api-key": ANTHROPIC_API_KEY,
                    "anthropic-version": "2023-06-01",
                    "content-type": "application/json",
                },
                json=body,
                timeout=90,
            )
        except Exception as e:
            last = "baglanti/zaman asimi: " + str(e)
            time.sleep(2 * (attempt + 1))
            continue
        if r.status_code == 200:
            data = r.json()
            parts = [b.get("text", "") for b in data.get("content", []) if b.get("type") == "text"]
            return "\n".join(parts)
        if r.status_code in (429, 529, 503):
            last = "Sunucu yogun (" + str(r.status_code) + ")"
            time.sleep(2 * (attempt + 1))
            continue
        try:
            last = r.json().get("error", {}).get("message", "") or ("hata " + str(r.status_code))
        except Exception:
            last = "hata " + str(r.status_code)
        break
    raise RuntimeError(last)


def parse_spec(text):
    t = (text or "").strip()
    t = re.sub(r"^```[a-zA-Z]*", "", t).strip()
    t = re.sub(r"```$", "", t).strip()
    start, end = t.find("{"), t.rfind("}")
    if start >= 0 and end > start:
        t = t[start:end + 1]
    return json.loads(t)


def build_docx(spec, path):
    doc = Document()
    if spec.get("title"):
        doc.add_heading(str(spec["title"]), 0)
    for b in spec.get("blocks", []) or []:
        tp = b.get("type")
        if tp == "heading":
            lvl = min(max(int(b.get("level", 1) or 1), 1), 4)
            doc.add_heading(str(b.get("text", "")), lvl)
        elif tp == "paragraph":
            doc.add_paragraph(str(b.get("text", "")))
        elif tp == "bullets":
            for it in b.get("items", []) or []:
                doc.add_paragraph(str(it), style="List Bullet")
        elif tp == "numbered":
            for it in b.get("items", []) or []:
                doc.add_paragraph(str(it), style="List Number")
        elif tp == "table":
            headers = b.get("headers", []) or []
            rows = b.get("rows", []) or []
            ncol = len(headers) or (len(rows[0]) if rows else 0)
            if ncol:
                table = doc.add_table(rows=0, cols=ncol)
                try:
                    table.style = "Table Grid"
                except Exception:
                    pass
                if headers:
                    cells = table.add_row().cells
                    for i, h in enumerate(headers):
                        if i < ncol:
                            cells[i].text = str(h)
                for row in rows:
                    cells = table.add_row().cells
                    for i, c in enumerate(row):
                        if i < ncol:
                            cells[i].text = str(c)
    doc.save(path)


def build_xlsx(spec, path):
    wb = Workbook()
    sheets = spec.get("sheets")
    if not sheets:
        sheets = [{"name": "Sayfa1", "headers": spec.get("headers", []), "rows": spec.get("rows", [])}]
    first = True
    for sh in sheets:
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = (str(sh.get("name") or "Sayfa"))[:31]
        headers = sh.get("headers", []) or []
        rows = sh.get("rows", []) or []
        if headers:
            ws.append([str(h) for h in headers])
        for row in rows:
            ws.append([str(c) for c in row])
    wb.save(path)


def build_pptx(spec, path):
    prs = Presentation()
    if spec.get("title"):
        s = prs.slides.add_slide(prs.slide_layouts[0])
        s.shapes.title.text = str(spec["title"])
        if len(s.placeholders) > 1:
            s.placeholders[1].text = str(spec.get("subtitle", "") or "")
    for sl in spec.get("slides", []) or []:
        s = prs.slides.add_slide(prs.slide_layouts[1])
        s.shapes.title.text = str(sl.get("title", ""))
        bullets = sl.get("bullets", []) or []
        if bullets:
            tf = s.placeholders[1].text_frame
            tf.text = str(bullets[0])
            for it in bullets[1:]:
                p = tf.add_paragraph()
                p.text = str(it)
    prs.save(path)


def build_pdf(spec, path):
    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm,
                            topMargin=18 * mm, bottomMargin=18 * mm)
    styles = getSampleStyleSheet()
    for sn in ["Title", "Heading1", "Heading2", "Heading3", "BodyText"]:
        styles[sn].fontName = PDF_FONT
    story = []
    if spec.get("title"):
        story.append(Paragraph(str(spec["title"]), styles["Title"]))
        story.append(Spacer(1, 8))
    for b in spec.get("blocks", []) or []:
        tp = b.get("type")
        if tp == "heading":
            lvl = min(max(int(b.get("level", 1) or 1), 1), 3)
            story.append(Paragraph(str(b.get("text", "")), styles["Heading" + str(lvl)]))
        elif tp == "paragraph":
            story.append(Paragraph(str(b.get("text", "")), styles["BodyText"]))
            story.append(Spacer(1, 4))
        elif tp in ("bullets", "numbered"):
            items = [ListItem(Paragraph(str(it), styles["BodyText"])) for it in (b.get("items", []) or [])]
            story.append(ListFlowable(items, bulletType="1" if tp == "numbered" else "bullet"))
            story.append(Spacer(1, 4))
        elif tp == "table":
            headers = b.get("headers", []) or []
            rows = b.get("rows", []) or []
            data = []
            if headers:
                data.append([str(h) for h in headers])
            for row in rows:
                data.append([str(c) for c in row])
            if data:
                tbl = Table(data, hAlign="LEFT")
                tbl.setStyle(TableStyle([
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F5FB3")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, -1), PDF_FONT),
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ]))
                story.append(tbl)
                story.append(Spacer(1, 6))
    if not story:
        story.append(Paragraph(str(spec.get("summary", "Belge")), styles["BodyText"]))
    doc.build(story)


BUILDERS = {"docx": build_docx, "xlsx": build_xlsx, "pptx": build_pptx, "pdf": build_pdf}


@app.get("/")
def health():
    return {"ok": True, "service": "kocderma-belge", "key": bool(ANTHROPIC_API_KEY)}


@app.post("/doc")
def make_doc(req: DocReq):
    if not ANTHROPIC_API_KEY:
        return {"reply": "Sunucuda API anahtari ayarli degil.", "files": []}
    if not req.messages:
        return {"reply": "Bos istek.", "files": []}

    try:
        text = call_claude(req.messages)
        spec = parse_spec(text)
    except Exception as e:
        return {"reply": "Belge olusturulamadi: " + str(e), "files": []}

    kind = str(spec.get("kind") or "docx").lower()
    if kind not in BUILDERS:
        kind = "docx"
    filename = str(spec.get("filename") or ("belge." + kind))
    filename = re.sub(r"[^A-Za-z0-9_.\-çğıöşüÇĞİÖŞÜ ]", "", filename) or ("belge." + kind)
    if not filename.lower().endswith("." + kind):
        filename = filename + "." + kind

    stored = uuid.uuid4().hex + "__" + filename
    path = os.path.join(FILES_DIR, stored)
    try:
        BUILDERS[kind](spec, path)
    except Exception as e:
        return {"reply": "Dosya uretiminde hata: " + str(e), "files": []}

    return {"reply": str(spec.get("summary") or "Belge hazir."),
            "files": [{"id": stored, "name": filename}]}


@app.get("/files/{file_id}")
def get_file(file_id: str):
    if "/" in file_id or "\\" in file_id or ".." in file_id:
        raise HTTPException(status_code=400, detail="gecersiz id")
    path = os.path.join(FILES_DIR, file_id)
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="dosya yok")
    name = file_id.split("__", 1)[1] if "__" in file_id else file_id
    return FileResponse(path, filename=name)
