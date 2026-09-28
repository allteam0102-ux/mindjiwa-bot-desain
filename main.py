# -*- coding: utf-8 -*-
"""
MINDJIWA BOT DESAIN  (gaya @mindjiwa.id — stabilo hijau)
========================================================
Alur: Notion (Status "Siap Desain") -> ambil gambar -> rakit desain gaya
mindjiwa -> kirim ke Telegram. Menghasilkan .pptx (Canva) + .svg (Figma).

Sumber gambar:
  - CAROUSEL     : otomatis dari Pexels (kata kunci per slide)
  - SINGLE POST  : dari link yang kamu tempel (Pinterest) -> diubah jadi
                   efek SKETSA PENSIL otomatis (filter gratis, bukan AI).
    (Kalau slide Single Post tidak diberi link, robot ambil dari Pexels lalu
     tetap dijadikan sketsa pensil.)

Tema per konten (kolom 'Tema'):
  - Krem  (default): background krem, foto di frame atas, teks di bawah
  - Gelap          : foto full, teks di bawah dengan overlay gelap
"""

import os
import re
import io
import html
import time
import base64
import random
import tempfile
import traceback

import requests
from PIL import Image, ImageDraw, ImageOps, ImageFilter, ImageChops
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass


# ======================================================================
# >>> PENGATURAN BRAND <<<
# ======================================================================
HASHTAG       = os.environ.get("HASHTAG")       or "#TempatmuJedaDanBercerita"
FOOTER_TEXT   = os.environ.get("FOOTER_TEXT")   or "Curhat dan Konseling\nOnline & Offline"
ACCENT_COLOR  = (os.environ.get("ACCENT_COLOR") or "A6CE39").lstrip("#")    # hijau stabilo
CREAM_COLOR   = (os.environ.get("CREAM_COLOR")  or "F3F1E7").lstrip("#")    # krem
LOGO_URL      = os.environ.get("LOGO_URL")      or ""                        # logo mindjiwa (PNG)
HEADLINE_FONT = os.environ.get("HEADLINE_FONT") or "Poppins"
BODY_FONT     = os.environ.get("BODY_FONT")     or "Poppins"
ARROW_TEXT    = os.environ.get("ARROW_TEXT")    or "→"
STYLE_HINT    = os.environ.get("STYLE_HINT")    or "aesthetic lifestyle candid natural light"
DEFAULT_THEME = (os.environ.get("DEFAULT_THEME") or "Krem").strip()

CANVAS_W = int(os.environ.get("CANVAS_W") or 1080)
CANVAS_H = int(os.environ.get("CANVAS_H") or 1350)

OUTPUT_PPTX = (os.environ.get("OUTPUT_PPTX") or "true").lower() == "true"
OUTPUT_SVG  = (os.environ.get("OUTPUT_SVG")  or "true").lower() == "true"


# ======================================================================
# 1. PENGATURAN TEKNIS (dari Secrets)
# ======================================================================
def env(name, default=None, required=False):
    val = os.environ.get(name, default)
    if required and (val is None or str(val).strip() == ""):
        raise SystemExit(f"[SETUP ERROR] Environment variable '{name}' belum diisi. Cek README.")
    return val

NOTION_TOKEN        = env("NOTION_TOKEN", required=True)
NOTION_DATABASE_ID  = env("NOTION_DATABASE_ID", required=True)
PEXELS_API_KEY      = env("PEXELS_API_KEY", required=True)
TELEGRAM_BOT_TOKEN  = env("TELEGRAM_BOT_TOKEN", required=True)
TELEGRAM_CHAT_ID    = env("TELEGRAM_CHAT_ID", required=True)

TITLE_PROPERTY   = env("TITLE_PROPERTY", "Judul")
STATUS_PROPERTY  = env("STATUS_PROPERTY", "Status")
FORMAT_PROPERTY  = env("FORMAT_PROPERTY", "Format")
CONTENT_PROPERTY = env("CONTENT_PROPERTY", "Konten")
KEYWORD_PROPERTY = env("KEYWORD_PROPERTY", "Kata Kunci Gambar")
THEME_PROPERTY   = env("THEME_PROPERTY", "Tema")

STATUS_READY = env("STATUS_READY", "Siap Desain")
STATUS_DONE  = env("STATUS_DONE",  "Terkirim")
STATUS_ERROR = env("STATUS_ERROR", "Gagal")
STATUS_TYPE  = env("STATUS_TYPE", "select").strip().lower()

PEXELS_ORIENTATION = "portrait" if CANVAS_H > CANVAS_W else "square"

def hex_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

ACCENT = hex_rgb(ACCENT_COLOR)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
INK    = "1A1A1A"    # teks gelap

EMU_W = CANVAS_W * 9525
EMU_H = CANVAS_H * 9525
def fx(f): return Emu(int(f * EMU_W))
def fy(f): return Emu(int(f * EMU_H))

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/122.0 Safari/537.36"}


# ======================================================================
# 2. HELPER NOTION
# ======================================================================
NOTION_BASE = "https://api.notion.com/v1"
NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Notion-Version": "2022-06-28",
    "Content-Type": "application/json",
}

def notion_find_ready():
    url = f"{NOTION_BASE}/databases/{NOTION_DATABASE_ID}/query"
    if STATUS_TYPE == "status":
        flt = {"property": STATUS_PROPERTY, "status": {"equals": STATUS_READY}}
    else:
        flt = {"property": STATUS_PROPERTY, "select": {"equals": STATUS_READY}}
    resp = requests.post(url, headers=NOTION_HEADERS, json={"filter": flt}, timeout=60)
    if resp.status_code != 200:
        raise SystemExit(f"[NOTION ERROR] {resp.status_code}: {resp.text}")
    return resp.json().get("results", [])

def notion_set_status(page_id, status_value, note=None):
    url = f"{NOTION_BASE}/pages/{page_id}"
    if STATUS_TYPE == "status":
        props = {STATUS_PROPERTY: {"status": {"name": status_value}}}
    else:
        props = {STATUS_PROPERTY: {"select": {"name": status_value}}}
    props_with_note = dict(props)
    if note:
        props_with_note["Catatan"] = {"rich_text": [{"text": {"content": note[:1900]}}]}
    r = requests.patch(url, headers=NOTION_HEADERS, json={"properties": props_with_note}, timeout=60)
    if r.status_code != 200 and note:
        requests.patch(url, headers=NOTION_HEADERS, json={"properties": props}, timeout=60)

def _plain_text(rich_list):
    return "".join(part.get("plain_text", "") for part in (rich_list or []))

def read_property(props, name, kind):
    p = props.get(name)
    if not p:
        return ""
    if kind == "title":
        return _plain_text(p.get("title", [])).strip()
    if kind == "rich_text":
        return _plain_text(p.get("rich_text", [])).strip()
    if kind == "select":
        return (p.get("select") or {}).get("name", "").strip()
    if kind == "status":
        return (p.get("status") or {}).get("name", "").strip()
    return ""


# ======================================================================
# 3. PARSING ISI KONTEN
# ======================================================================
def split_slides(content_text):
    raw = (content_text or "").replace("\r\n", "\n").replace("\r", "\n")
    blocks = re.split(r"(?m)^\s*---\s*$", raw)
    return [b.strip() for b in blocks if b.strip()]

URL_RE = re.compile(r"https?://\S+")

def extract_url_from_block(block):
    url = None
    kept = []
    for line in block.split("\n"):
        m = URL_RE.search(line)
        if m and url is None:
            url = m.group(0).rstrip(").,")
            sisa = URL_RE.sub("", line).strip()
            if sisa:
                kept.append(sisa)
        else:
            kept.append(line)
    return url, "\n".join(kept).strip()

def headline_and_body(block):
    lines = block.split("\n")
    idx = next((i for i, l in enumerate(lines) if l.strip()), None)
    if idx is None:
        return "", ""
    return lines[idx].strip(), "\n".join(lines[idx + 1:]).strip()

def keyword_for(index, keyword_blocks, headline, body):
    if index < len(keyword_blocks) and keyword_blocks[index].strip():
        q = keyword_blocks[index].strip()
    elif headline:
        q = headline
    elif body:
        q = body
    else:
        q = "aesthetic minimal"
    q = re.sub(r"==", "", q)
    words = re.sub(r"[^\w\s]", " ", q).split()
    return " ".join(words[:4]) if words else "aesthetic minimal"

def hl_words(line):
    """Pecah 1 baris jadi kata-kata + tanda apakah di-highlight (==...==)."""
    out = []
    for part in re.split(r"(==.+?==)", line):
        is_hl = len(part) >= 4 and part.startswith("==") and part.endswith("==")
        seg = part[2:-2] if is_hl else part
        for w in seg.split(" "):
            if w != "":
                out.append((w, is_hl))
    return out


# ======================================================================
# 4. GAMBAR
# ======================================================================
def fetch_image_bytes(url):
    r = requests.get(url, headers=UA, timeout=60)
    r.raise_for_status()
    ct = r.headers.get("Content-Type", "").lower()
    if "image" not in ct:
        raise ValueError(
            f"URL bukan gambar langsung (Content-Type: {ct or 'tidak diketahui'}). "
            f"Di Pinterest: klik kanan gambar -> 'Copy image address' (link diakhiri .jpg/.png)."
        )
    return r.content

def pexels_pick(query, used_ids):
    def _search(q):
        r = requests.get("https://api.pexels.com/v1/search",
                         headers={"Authorization": PEXELS_API_KEY},
                         params={"query": q, "per_page": 15, "orientation": PEXELS_ORIENTATION}, timeout=60)
        r.raise_for_status()
        return r.json().get("photos", [])
    photos = _search(f"{query} {STYLE_HINT}".strip()) or _search(query) or _search("aesthetic minimal calm")
    if not photos:
        return None, None
    fresh = [p for p in photos if p.get("id") not in used_ids]
    pool = fresh if fresh else photos
    photo = random.choice(pool[:8])
    if photo.get("id"):
        used_ids.add(photo["id"])
    img_url = photo["src"].get("large2x") or photo["src"].get("large") or photo["src"]["original"]
    return requests.get(img_url, timeout=60).content, f'Foto: {photo.get("photographer","-")} (Pexels)'

def crop_to(im, w, h):
    target = w / h
    iw, ih = im.size
    if iw / ih > target:
        nw = int(ih * target); x = (iw - nw) // 2; im = im.crop((x, 0, x + nw, ih))
    else:
        nh = int(iw / target); y = (ih - nh) // 2; im = im.crop((0, y, iw, y + nh))
    return im.resize((w, h), Image.LANCZOS)

def pencil_sketch(im):
    """Ubah foto jadi efek sketsa pensil (gratis, tanpa AI)."""
    import numpy as np
    gray = ImageOps.grayscale(im)
    inv = ImageOps.invert(gray)
    blur = inv.filter(ImageFilter.GaussianBlur(radius=max(6, im.size[0] // 45)))
    a = np.asarray(gray, dtype=np.float32)
    b = np.asarray(blur, dtype=np.float32)
    with np.errstate(divide="ignore", invalid="ignore"):
        res = a * 255.0 / (255.0 - b)          # color dodge
    res[b >= 255] = 255
    res = np.clip(res, 0, 255).astype("uint8")
    sketch = Image.fromarray(res, "L")
    sketch = ImageOps.autocontrast(sketch, cutoff=1)
    return sketch.convert("RGB")


# ======================================================================
# 5. WARNA TEMA
# ======================================================================
def is_dark(theme):
    return str(theme).strip().lower().startswith("gelap")

def theme_colors(theme):
    if is_dark(theme):
        return dict(bg="121212", head="FFFFFF", body="D8D8D8", foot="FFFFFF", hl_text=INK)
    return dict(bg=CREAM_COLOR, head=INK, body="555555", foot=INK, hl_text=INK)


def download_logo_path():
    if not LOGO_URL:
        return None
    try:
        data = requests.get(LOGO_URL, headers=UA, timeout=30).content
        path = os.path.join(tempfile.gettempdir(), f"logo_{int(time.time()*1000)}.png")
        with open(path, "wb") as fp:
            fp.write(data)
        Image.open(path).verify()
        return path
    except Exception as e:
        print("  ! logo gagal diambil:", e)
        return None


# ======================================================================
# 6. RAKIT .PPTX (untuk Canva)
# ======================================================================
def _set_highlight(run, hexcolor):
    """Tambah stabilo (highlight) ke sebuah run — didukung PowerPoint & Canva."""
    try:
        rPr = run._r.get_or_add_rPr()
        hl = rPr.makeelement(qn('a:highlight'), {})
        clr = hl.makeelement(qn('a:srgbClr'), {'val': hexcolor})
        hl.append(clr)
        rPr.append(hl)
    except Exception:
        pass

def _add_rich(slide, left, top, width, height, text, size_pt, bold, font, color_hex,
              align=PP_ALIGN.LEFT, highlight=True, italic=False):
    box = slide.shapes.add_textbox(fx(left), fy(top), fx(width), fy(height))
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for line in text.split("\n"):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        words = hl_words(line) if highlight else [(w, False) for w in line.split(" ") if w]
        if not words:
            words = [("", False)]
        for j, (w, is_hl) in enumerate(words):
            run = p.add_run()
            run.text = (w if j == 0 else " " + w)
            f = run.font
            f.size = Pt(size_pt)
            f.bold = bold
            f.italic = italic
            f.name = font
            if is_hl:
                f.color.rgb = hex_rgb(INK)
                _set_highlight(run, ACCENT_COLOR)
            else:
                f.color.rgb = hex_rgb(color_hex)
    return box

def build_pptx(slides_data, out_path, logo_path):
    prs = Presentation()
    prs.slide_width = Emu(EMU_W)
    prs.slide_height = Emu(EMU_H)
    blank = prs.slide_layouts[6]

    for s in slides_data:
        c = theme_colors(s["theme"])
        dark = is_dark(s["theme"])
        slide = prs.slides.add_slide(blank)

        # background
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb = hex_rgb(c["bg"])
        bg.line.fill.background(); bg.shadow.inherit = False

        if dark:
            # foto full-bleed + panah/teks di bawah
            slide.shapes.add_picture(s["image_path"], 0, 0, width=prs.slide_width, height=prs.slide_height)
            head_top, body_top = 0.60, 0.78
        else:
            # foto di frame atas (rounded), teks di bawah
            slide.shapes.add_picture(s["image_path"], fx(0.07), fy(0.11),
                                     width=fx(0.86), height=fy(0.34))
            head_top, body_top = 0.50, 0.67

        # logo + hashtag
        if logo_path:
            try:
                slide.shapes.add_picture(logo_path, fx(0.06), fy(0.05), height=fy(0.045))
            except Exception:
                pass
        _add_rich(slide, 0.55, 0.05, 0.39, 0.07, HASHTAG, 12, True, BODY_FONT, ACCENT_COLOR,
                  align=PP_ALIGN.RIGHT, highlight=False)

        # headline + body
        if s["headline"]:
            _add_rich(slide, 0.07, head_top, 0.86, 0.18, s["headline"], 28, True, HEADLINE_FONT, c["head"])
        if s["body"]:
            _add_rich(slide, 0.07, body_top, 0.86, 0.18, s["body"], 16, False, BODY_FONT, c["body"])

        # footer + panah
        _add_rich(slide, 0.08, 0.90, 0.55, 0.08, FOOTER_TEXT, 13, True, BODY_FONT, c["foot"], highlight=False)
        if s["total"] > 1:
            _add_rich(slide, 0.82, 0.90, 0.12, 0.07, ARROW_TEXT, 24, True, BODY_FONT, c["head"],
                      align=PP_ALIGN.RIGHT, highlight=False)

    prs.save(out_path)
    return out_path


# ======================================================================
# 7. RAKIT .SVG (untuk Figma) — stabilo hijau digambar sebagai blok
# ======================================================================
def _img_data_uri(path):
    im = Image.open(path).convert("RGB")
    buf = io.BytesIO(); im.save(buf, "JPEG", quality=85)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()

def _svg_escape(t):
    return html.escape(t or "", quote=True)

def _svg_rich(text, x0, y0, width, size, bold, font, color_hex, gap=1.34):
    """Render teks + stabilo hijau (blok menerus) di belakang frasa ber-highlight."""
    weight = "700" if bold else "400"
    char_w = size * (0.64 if bold else 0.58)
    space_w = size * 0.46
    out = []
    y = y0
    for src_line in text.split("\n"):
        words = hl_words(src_line)
        if not words:
            y += int(size * gap); continue
        # bungkus jadi baris visual sesuai lebar
        vlines, cur, cur_w = [], [], 0
        for w, is_hl in words:
            ww = len(w) * char_w
            add = ww + (space_w if cur else 0)
            if cur and cur_w + add > width:
                vlines.append(cur); cur, cur_w = [], 0
                add = ww
            cur.append((w, is_hl, ww)); cur_w += add
        if cur:
            vlines.append(cur)
        for vl in vlines:
            baseline = y + size
            # 1) hitung posisi x tiap kata
            positions = []
            cx = x0
            for k, (w, is_hl, ww) in enumerate(vl):
                if k > 0:
                    cx += space_w
                positions.append(cx)
                cx += ww
            # 2) gambar stabilo: gabung kata highlight yang berurutan jadi satu blok
            k = 0
            while k < len(vl):
                if vl[k][1]:
                    j = k
                    while j + 1 < len(vl) and vl[j + 1][1]:
                        j += 1
                    x_start = positions[k]
                    x_end = positions[j] + vl[j][2]
                    pad = size * 0.12
                    out.append(
                        f'<rect x="{int(x_start - pad)}" y="{int(baseline - size*0.86)}" '
                        f'width="{int(x_end - x_start + pad*2)}" height="{int(size*1.10)}" '
                        f'rx="{int(size*0.10)}" fill="#{ACCENT_COLOR}"/>'
                    )
                    k = j + 1
                else:
                    k += 1
            # 3) gambar teks
            for (w, is_hl, ww), px in zip(vl, positions):
                fill = INK if is_hl else color_hex
                out.append(
                    f'<text x="{int(px)}" y="{int(baseline)}" '
                    f'font-family="{_svg_escape(font)}, Arial, sans-serif" '
                    f'font-size="{size}" font-weight="{weight}" fill="#{fill}">{_svg_escape(w)}</text>'
                )
            y += int(size * gap)
    return "\n".join(out), y

def build_svg_single(slides_data, out_path, logo_path):
    total = len(slides_data)
    W = CANVAS_W * total
    H = CANVAS_H
    logo_uri = None
    if logo_path:
        try:
            im = Image.open(logo_path).convert("RGBA")
            b = io.BytesIO(); im.save(b, "PNG")
            logo_uri = "data:image/png;base64," + base64.b64encode(b.getvalue()).decode()
        except Exception:
            pass

    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">']
    for idx, s in enumerate(slides_data):
        xo = idx * CANVAS_W
        c = theme_colors(s["theme"])
        dark = is_dark(s["theme"])
        # background
        parts.append(f'<rect x="{xo}" y="0" width="{CANVAS_W}" height="{H}" fill="#{c["bg"]}"/>')
        if dark:
            parts.append(f'<image x="{xo}" y="0" width="{CANVAS_W}" height="{H}" '
                         f'preserveAspectRatio="xMidYMid slice" href="{_img_data_uri(s["image_path"])}"/>')
            parts.append(f'<rect x="{xo}" y="{int(0.45*H)}" width="{CANVAS_W}" height="{int(0.55*H)}" '
                         f'fill="#000000" opacity="0.45"/>')
            head_y, body_y = int(0.60*H), int(0.78*H)
        else:
            fx0 = xo + int(0.07*CANVAS_W); fy0 = int(0.11*H)
            fw = int(0.86*CANVAS_W); fh = int(0.34*H)
            parts.append(f'<clipPath id="cp{idx}"><rect x="{fx0}" y="{fy0}" width="{fw}" height="{fh}" rx="28"/></clipPath>')
            parts.append(f'<image x="{fx0}" y="{fy0}" width="{fw}" height="{fh}" '
                         f'preserveAspectRatio="xMidYMid slice" clip-path="url(#cp{idx})" '
                         f'href="{_img_data_uri(s["image_path"])}"/>')
            head_y, body_y = int(0.50*H), int(0.67*H)

        if logo_uri:
            parts.append(f'<image x="{xo + int(0.06*CANVAS_W)}" y="{int(0.05*H)}" '
                         f'height="{int(0.05*H)}" href="{logo_uri}"/>')
        parts.append(
            f'<text x="{xo + int(0.94*CANVAS_W)}" y="{int(0.075*H)}" text-anchor="end" '
            f'font-family="{_svg_escape(BODY_FONT)}, Arial, sans-serif" font-size="22" '
            f'font-weight="700" fill="#{ACCENT_COLOR}">{_svg_escape(HASHTAG)}</text>'
        )

        if s["headline"]:
            svg, _ = _svg_rich(s["headline"], xo + int(0.07*CANVAS_W), head_y,
                               int(0.86*CANVAS_W), 42, True, HEADLINE_FONT, c["head"])
            parts.append(svg)
        if s["body"]:
            svg, _ = _svg_rich(s["body"], xo + int(0.07*CANVAS_W), body_y,
                               int(0.86*CANVAS_W), 26, False, BODY_FONT, c["body"])
            parts.append(svg)

        # footer + garis hijau + panah
        parts.append(f'<rect x="{xo + int(0.06*CANVAS_W)}" y="{int(0.905*H)}" width="6" height="{int(0.05*H)}" fill="#{ACCENT_COLOR}"/>')
        fsvg, _ = _svg_rich(FOOTER_TEXT, xo + int(0.09*CANVAS_W), int(0.905*H),
                            int(0.5*CANVAS_W), 20, True, BODY_FONT, c["foot"])
        parts.append(fsvg)
        if total > 1:
            parts.append(
                f'<text x="{xo + int(0.93*CANVAS_W)}" y="{int(0.95*H)}" text-anchor="end" '
                f'font-family="{_svg_escape(BODY_FONT)}, Arial, sans-serif" font-size="40" '
                f'font-weight="700" fill="#{c["head"]}">{_svg_escape(ARROW_TEXT)}</text>'
            )
    parts.append('</svg>')
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))
    return out_path


# ======================================================================
# 8. TELEGRAM
# ======================================================================
TG_BASE = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}"

def _tg_call(method, data, files=None, label=""):
    try:
        r = requests.post(f"{TG_BASE}/{method}", data=data, files=files, timeout=300)
        try:
            j = r.json()
        except Exception:
            j = {}
        time.sleep(0.5)
        if not r.ok or not j.get("ok", False):
            print(f"    ! Telegram {method} {label} GAGAL (HTTP {r.status_code}): {r.text[:300]}")
            return False
        return True
    except Exception as e:
        print(f"    ! Telegram {method} {label} ERROR: {e}")
        return False

def tg_message(text):
    return _tg_call("sendMessage", {"chat_id": TELEGRAM_CHAT_ID, "text": text[:4000]}, label="msg")

def tg_photo(path, caption=""):
    with open(path, "rb") as f:
        return _tg_call("sendPhoto", {"chat_id": TELEGRAM_CHAT_ID, "caption": caption[:1000]},
                        {"photo": f}, label="photo")

def tg_document(path, caption=""):
    with open(path, "rb") as f:
        return _tg_call("sendDocument", {"chat_id": TELEGRAM_CHAT_ID, "caption": caption[:1000]},
                        {"document": f}, label="doc")


# ======================================================================
# 9. PROSES SATU KONTEN
# ======================================================================
def process_page(page, workdir):
    props = page["properties"]
    title = read_property(props, TITLE_PROPERTY, "title") or "Tanpa Judul"
    fmt = read_property(props, FORMAT_PROPERTY, "select") or "Single Post"
    content = read_property(props, CONTENT_PROPERTY, "rich_text")
    keywords = read_property(props, KEYWORD_PROPERTY, "rich_text")
    theme = read_property(props, THEME_PROPERTY, "select") or DEFAULT_THEME

    if not content.strip():
        raise ValueError("Kolom 'Konten' kosong.")

    is_single = "single" in fmt.lower()
    slide_blocks = split_slides(content)
    if is_single:
        slide_blocks = slide_blocks[:1]
    keyword_blocks = split_slides(keywords)

    print(f"  -> '{title}' | {fmt} | {len(slide_blocks)} slide | tema={theme} | pensil={is_single}")

    slides_data, preview_paths, credits = [], [], []
    used_ids = set()
    total = len(slide_blocks)
    for i, block in enumerate(slide_blocks, start=1):
        inline_url, clean = extract_url_from_block(block)
        headline, body = headline_and_body(clean)

        if inline_url:                                   # ada link -> pakai
            im = Image.open(io.BytesIO(fetch_image_bytes(inline_url))).convert("RGB")
        else:                                            # tidak ada -> ambil Pexels
            q = keyword_for(i - 1, keyword_blocks, headline, body)
            data, credit = pexels_pick(q, used_ids)
            if not data:
                raise ValueError(f"Slide {i}: gambar Pexels tidak ketemu untuk '{q}'.")
            im = Image.open(io.BytesIO(data)).convert("RGB")
            if credit:
                credits.append(f"Slide {i}: {credit}")

        # ukuran area gambar sesuai tema
        if is_dark(theme):
            im = crop_to(im, CANVAS_W, CANVAS_H)
        else:
            im = crop_to(im, int(0.86 * CANVAS_W), int(0.34 * CANVAS_H))

        if is_single:                                    # single post -> sketsa pensil
            im = pencil_sketch(im)

        img_path = os.path.join(workdir, f"slide_{i}.png")
        im.save(img_path, "PNG")
        preview_paths.append(img_path)
        slides_data.append({"headline": headline, "body": body, "image_path": img_path,
                            "index": i, "total": total, "theme": theme})

    safe_name = re.sub(r"[^\w\- ]", "", title).strip().replace(" ", "_")[:40] or "desain"
    logo_path = download_logo_path()

    lines = [f"🌿 {title}  ({fmt}, {total} slide, tema {theme})", ""]
    for i, s in enumerate(slides_data, start=1):
        h = s["headline"] or "(tanpa judul)"
        b = (s["body"][:150] + "…") if len(s["body"]) > 150 else s["body"]
        lines.append(f"— Slide {i}: {h}")
        if b:
            lines.append(f"  {b}")
    lines += ["", *credits]
    tg_message("\n".join(lines))

    for i, p in enumerate(preview_paths, start=1):
        tg_photo(p, caption=f"Preview slide {i}/{total}")

    if OUTPUT_PPTX:
        pptx_path = os.path.join(workdir, f"{safe_name}.pptx")
        build_pptx(slides_data, pptx_path, logo_path)
        if not tg_document(pptx_path, caption=f"{title} — .pptx: import ke Canva ✨"):
            tg_message("⚠️ File .pptx gagal dikirim (cek log).")

    if OUTPUT_SVG:
        svg_path = os.path.join(workdir, f"{safe_name}.svg")
        build_svg_single(slides_data, svg_path, logo_path)
        if not tg_document(svg_path, caption=f"{title} — .svg: tarik ke Figma (teks bisa diedit) ✨"):
            tg_message("⚠️ File .svg gagal dikirim (cek log).")

    return total


# ======================================================================
# 10. MAIN
# ======================================================================
def main():
    print(f"== MINDJIWA BOT DESAIN ({CANVAS_W}x{CANVAS_H}) | pptx={OUTPUT_PPTX} svg={OUTPUT_SVG} ==")
    if tg_message("✅ Mindjiwa bot desain — mulai jalan."):
        print("  Telegram OK.")
    else:
        print("  !! Telegram BERMASALAH — cek TELEGRAM_BOT_TOKEN & TELEGRAM_CHAT_ID.")
    pages = notion_find_ready()
    print(f"Ditemukan {len(pages)} konten berstatus '{STATUS_READY}'.")
    if not pages:
        print("Tidak ada yang perlu diproses. Selesai.")
        return

    for page in pages:
        page_id = page["id"]
        workdir = tempfile.mkdtemp()
        try:
            process_page(page, workdir)
            notion_set_status(page_id, STATUS_DONE)
            print("  v Sukses & status diubah jadi:", STATUS_DONE)
        except Exception as e:
            err = f"{type(e).__name__}: {e}"
            print("  x GAGAL:", err)
            traceback.print_exc()
            try:
                tg_message(f"⚠️ Gagal memproses '{read_property(page['properties'], TITLE_PROPERTY, 'title')}'.\n{err}")
            except Exception:
                pass
            notion_set_status(page_id, STATUS_ERROR, note=err)

    print("== Selesai ==")


if __name__ == "__main__":
    main()
