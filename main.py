# -*- coding: utf-8 -*-
"""
MINDJIWA BOT DESAIN  (gaya @mindjiwa.id — multi-style, stabilo hijau)
====================================================================
Notion (Status "Siap Desain") -> gambar -> rakit desain -> Telegram.
Output: .pptx (Canva) + .svg (Figma).

GAMBAR:
  - CAROUSEL     : otomatis dari Pexels (kata kunci per slide)
  - SINGLE POST  : dari link yang kamu tempel -> jadi SKETSA PENSIL (filter gratis)
  - (link boleh ditaruh langsung di dalam Konten, baris sendiri)

TEMA / STYLE (kolom 'Tema'):
  - Krem     : bg krem, foto di frame atas, teks di bawah
  - Gelap    : foto full, teks di bawah (overlay gelap)
  - Kertas   : bg krem, TANPA foto, teks di tengah, headline serif (gaya "Manifesting")
  - Kutipan  : foto gelap + tanda kutip hijau besar + teks putih

PENANDA TEKS (berlaku semua style):
  ==kata==   -> stabilo hijau
  _kata_     -> miring (italic)
  ((kata))   -> dilingkari hijau
  ~kata~     -> garis bawah hijau
"""

import os
import re
import io
import json
import html
import time
import base64
import random
import secrets
import tempfile
import traceback

import requests
from PIL import Image, ImageOps, ImageFilter
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
LOGO_URL      = os.environ.get("LOGO_URL")      or ""
HEADLINE_FONT = os.environ.get("HEADLINE_FONT") or "Plus Jakarta Sans"
SERIF_FONT    = os.environ.get("SERIF_FONT")    or "Playfair Display"
BODY_FONT     = os.environ.get("BODY_FONT")     or "Plus Jakarta Sans"
ARROW_TEXT    = os.environ.get("ARROW_TEXT")    or "→"
STYLE_HINT    = os.environ.get("STYLE_HINT")    or "editorial lifestyle film"
DEFAULT_THEME = (os.environ.get("DEFAULT_THEME") or "Krem").strip()

CANVAS_W = int(os.environ.get("CANVAS_W") or 1080)
CANVAS_H = int(os.environ.get("CANVAS_H") or 1350)

OUTPUT_PPTX = (os.environ.get("OUTPUT_PPTX") or "true").lower() == "true"
OUTPUT_SVG  = (os.environ.get("OUTPUT_SVG")  or "true").lower() == "true"
OUTPUT_SVG_PER_SLIDE = (os.environ.get("OUTPUT_SVG_PER_SLIDE") or "true").lower() == "true"

# ----------------------------------------------------------------------
# SISTEM BELAJAR (memory + feedback). File memori disimpan di repo.
# ----------------------------------------------------------------------
MEMORY_PATH = os.environ.get("MEMORY_FILE") or "memory.json"
LEARN_ENABLED = (os.environ.get("LEARN_ENABLED") or "true").lower() == "true"

# "knob" yang bisa digeser oleh feedback (nilai awal = netral)
G_HEAD_SCALE = 1.0   # pengali ukuran judul
G_BODY_SCALE = 1.0   # pengali ukuran body
G_SCRIM      = 1.0   # pengali kegelapan overlay (inert di Mindjiwa — overlay fix per-style)
G_SOFT       = 0.0   # tambahan "film/soft" (inert di Mindjiwa)

def _clamp(v, lo, hi):
    return max(lo, min(hi, v))

# Ukuran teks DASAR tetap ditentukan per-TEMA lewat style_of() (head_size/body_size).
# Knob sistem belajar cuma MENGALIKAN ukuran itu di titik pakai (pptx & svg), dengan
# lantai ~10pt / ~16px biar tetap kebaca walau knob mengecil. Base TIDAK diubah.
def head_pt(base): return max(10, int(round(base * G_HEAD_SCALE)))
def body_pt(base): return max(9,  int(round(base * G_BODY_SCALE)))
def head_px(base): return max(16, int(round(base * G_HEAD_SCALE)))
def body_px(base): return max(14, int(round(base * G_BODY_SCALE)))


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
UNSPLASH_ACCESS_KEY = os.environ.get("UNSPLASH_ACCESS_KEY", "")   # opsional

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
DONE_EMOJI   = env("DONE_EMOJI", "✅")      # ✅ ikon halaman saat selesai

PEXELS_ORIENTATION = "portrait" if CANVAS_H > CANVAS_W else "square"

def hex_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

ACCENT = hex_rgb(ACCENT_COLOR)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
INK    = "1A1A1A"

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

def notion_set_status(page_id, status_value, note=None, icon_emoji=None):
    url = f"{NOTION_BASE}/pages/{page_id}"
    if STATUS_TYPE == "status":
        props = {STATUS_PROPERTY: {"status": {"name": status_value}}}
    else:
        props = {STATUS_PROPERTY: {"select": {"name": status_value}}}
    props_with_note = dict(props)
    if note:
        props_with_note["Catatan"] = {"rich_text": [{"text": {"content": note[:1900]}}]}
    body = {"properties": props_with_note}
    if icon_emoji:
        body["icon"] = {"type": "emoji", "emoji": icon_emoji}   # ubah ikon halaman (penanda selesai)
    r = requests.patch(url, headers=NOTION_HEADERS, json=body, timeout=60)
    if r.status_code != 200:
        # ulangi tanpa catatan (kalau kolom Catatan tak ada), ikon tetap diubah
        body2 = {"properties": props}
        if icon_emoji:
            body2["icon"] = {"type": "emoji", "emoji": icon_emoji}
        requests.patch(url, headers=NOTION_HEADERS, json=body2, timeout=60)

def notion_get_page(page_id):
    """Ambil ulang 1 halaman Notion by ID (buat auto-revisi)."""
    try:
        r = requests.get(f"{NOTION_BASE}/pages/{page_id}", headers=NOTION_HEADERS, timeout=60)
        if r.status_code == 200:
            return r.json()
    except Exception as e:
        print("  ! notion_get_page gagal:", e)
    return None

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

def get_title(props):
    """Ambil judul: coba kolom bernama TITLE_PROPERTY, kalau kosong cari kolom bertipe 'title'."""
    t = read_property(props, TITLE_PROPERTY, "title")
    if t:
        return t
    for _name, p in props.items():
        if isinstance(p, dict) and p.get("type") == "title":
            return _plain_text(p.get("title", [])).strip()
    return ""


# ======================================================================
# 3. PARSING ISI KONTEN
# ======================================================================
def split_slides(content_text):
    raw = (content_text or "").replace("\r\n", "\n").replace("\r", "\n")
    # terima --- / — / –– dsb (Notion kadang mengubah --- jadi em-dash)
    blocks = re.split(r"(?m)^\s*(?:[-–—]{2,}|[–—])\s*$", raw)
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
    q = re.sub(r"[=~()_]", "", q)
    words = re.sub(r"[^\w\s]", " ", q).split()
    return " ".join(words[:4]) if words else "aesthetic minimal"

def rich_words(line):
    """Pecah baris jadi kata + flag: hl(stabilo) it(italic) ul(garis bawah) circ(lingkaran)."""
    results = []
    def emit(text, hl, it, ul, circ):
        for w in text.split(" "):
            if w != "":
                results.append((w, hl, it, ul, circ))
    for p1 in re.split(r"(==.+?==)", line):
        hl = len(p1) >= 4 and p1.startswith("==") and p1.endswith("==")
        s1 = p1[2:-2] if hl else p1
        if s1 == "":
            continue
        for p2 in re.split(r"(\(\(.+?\)\))", s1):
            circ = len(p2) >= 4 and p2.startswith("((") and p2.endswith("))")
            s2 = p2[2:-2] if circ else p2
            if s2 == "":
                continue
            for p3 in re.split(r"(~[^~]+~)", s2):
                ul = len(p3) >= 2 and p3.startswith("~") and p3.endswith("~")
                s3 = p3[1:-1] if ul else p3
                if s3 == "":
                    continue
                for p4 in re.split(r"(_[^_]+_)", s3):
                    it = len(p4) >= 2 and p4.startswith("_") and p4.endswith("_")
                    s4 = p4[1:-1] if it else p4
                    if s4 == "":
                        continue
                    emit(s4, hl, it, ul, circ)
    return results


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

def _short_q(query, add_hint=True):
    """Kata kunci dipendekin (maks 3 kata + 2 kata hint) -> hindari Pexels 500."""
    words = (query or "").split()[:3]
    q = " ".join(words) if words else "aesthetic calm"
    if add_hint and STYLE_HINT:
        q = (q + " " + " ".join(STYLE_HINT.split()[:2])).strip()
    return q

def pexels_pick(query, used_ids):
    def _search(q):
        r = requests.get("https://api.pexels.com/v1/search",
                         headers={"Authorization": PEXELS_API_KEY},
                         params={"query": q, "per_page": 15, "orientation": PEXELS_ORIENTATION}, timeout=60)
        r.raise_for_status()
        return r.json().get("photos", [])
    photos = _search(_short_q(query)) or _search(_short_q(query, add_hint=False)) or _search("aesthetic minimal calm")
    if not photos:
        return None, None
    fresh = [p for p in photos if p.get("id") not in used_ids]
    pool = fresh if fresh else photos
    photo = random.choice(pool[:8])
    if photo.get("id"):
        used_ids.add(photo["id"])
    img_url = photo["src"].get("large2x") or photo["src"].get("large") or photo["src"]["original"]
    return requests.get(img_url, timeout=60).content, f'Foto: {photo.get("photographer","-")} (Pexels)'

def unsplash_pick(query, used_ids):
    r = requests.get("https://api.unsplash.com/search/photos",
                     headers={"Authorization": f"Client-ID {UNSPLASH_ACCESS_KEY}"},
                     params={"query": _short_q(query), "per_page": 15, "orientation": PEXELS_ORIENTATION}, timeout=60)
    r.raise_for_status()
    results = r.json().get("results", [])
    if not results:
        return None, None
    fresh = [p for p in results if p.get("id") not in used_ids]
    pool = fresh if fresh else results
    photo = random.choice(pool[:8])
    if photo.get("id"):
        used_ids.add(photo["id"])
    urls = photo.get("urls", {})
    img_url = urls.get("regular") or urls.get("full") or urls.get("raw")
    name = (photo.get("user") or {}).get("name", "-")
    return requests.get(img_url, timeout=60).content, f"Foto: {name} (Unsplash)"

def image_pick(query, used_ids):
    """Ambil gambar: Unsplash dulu (utama), Pexels cadangan.
    Kalau kata kunci tak dapat gambar, coba kata kunci cadangan Inggris."""
    sources = (["unsplash", "pexels"] if UNSPLASH_ACCESS_KEY else ["pexels"])
    queries = [query, "calm aesthetic minimal lifestyle", "soft natural light nature"]
    for q in queries:
        for src in sources:
            try:
                data, credit = (unsplash_pick if src == "unsplash" else pexels_pick)(q, used_ids)
                if data:
                    return data, credit
            except Exception as e:
                print(f"    ! {src} gagal ('{q}'): {e}")
    return None, None

def make_placeholder():
    """Background netral (gradasi gelap) kalau benar-benar tak ada gambar."""
    im = Image.new("RGB", (CANVAS_W, CANVAS_H))
    import numpy as np
    top = np.array([40, 44, 38]); bot = np.array([12, 12, 12])
    arr = np.zeros((CANVAS_H, CANVAS_W, 3), dtype="uint8")
    for y in range(CANVAS_H):
        t = y / CANVAS_H
        arr[y, :] = (top * (1 - t) + bot * t).astype("uint8")
    return Image.fromarray(arr, "RGB")

def crop_to(im, w, h):
    target = w / h
    iw, ih = im.size
    if iw / ih > target:
        nw = int(ih * target); x = (iw - nw) // 2; im = im.crop((x, 0, x + nw, ih))
    else:
        nh = int(iw / target); y = (ih - nh) // 2; im = im.crop((0, y, iw, y + nh))
    return im.resize((w, h), Image.LANCZOS)

def pencil_sketch(im):
    import numpy as np
    gray = ImageOps.grayscale(im)
    inv = ImageOps.invert(gray)
    blur = inv.filter(ImageFilter.GaussianBlur(radius=max(6, im.size[0] // 45)))
    a = np.asarray(gray, dtype=np.float32)
    b = np.asarray(blur, dtype=np.float32)
    with np.errstate(divide="ignore", invalid="ignore"):
        res = a * 255.0 / (255.0 - b)
    res[b >= 255] = 255
    res = np.clip(res, 0, 255).astype("uint8")
    sketch = Image.fromarray(res, "L")
    sketch = ImageOps.autocontrast(sketch, cutoff=1)
    return sketch.convert("RGB")

def darken(im, amt=0.4):
    return Image.blend(im, Image.new("RGB", im.size, (0, 0, 0)), amt)


# ======================================================================
# 5. STYLE / TEMA
# ======================================================================
def style_of(theme):
    t = str(theme).strip().lower()
    if t.startswith("gelap"):
        return dict(kind="gelap", bg="121212", head="FFFFFF", body="D8D8D8", foot="FFFFFF",
                    photo="full", align="left", serif=False, quote=False,
                    head_top=0.60, body_top=0.78, head_size=40, body_size=26)
    if t.startswith("kertas") or t.startswith("manifest"):
        return dict(kind="kertas", bg=CREAM_COLOR, head=INK, body="444444", foot=INK,
                    photo="none", align="center", serif=True, quote=False,
                    head_top=0.34, body_top=0.46, head_size=52, body_size=28)
    if t.startswith("kutipan") or t.startswith("quote"):
        return dict(kind="kutipan", bg="121212", head="FFFFFF", body="D8D8D8", foot="FFFFFF",
                    photo="full", align="left", serif=False, quote=True,
                    head_top=0.40, body_top=0.40, head_size=38, body_size=34)
    if t.startswith("duo"):
        return dict(kind="duo", bg="121212", head="FFFFFF", body="FFFFFF", foot="FFFFFF",
                    photo="duo", align="left", serif=False, quote=False,
                    head_top=0.30, body_top=0.72, head_size=30, body_size=30)
    return dict(kind="krem", bg=CREAM_COLOR, head=INK, body="555555", foot=INK,
                photo="frame", align="left", serif=False, quote=False,
                head_top=0.50, body_top=0.67, head_size=40, body_size=26)


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
        words = rich_words(line) if highlight else [(w, False, False, False, False) for w in line.split(" ") if w]
        if not words:
            words = [("", False, False, False, False)]
        for j, (w, is_hl, is_it, is_ul, is_circ) in enumerate(words):
            run = p.add_run()
            run.text = (w if j == 0 else " " + w)
            f = run.font
            f.size = Pt(size_pt)
            f.bold = bold
            f.italic = italic or is_it
            f.underline = is_ul
            f.name = font
            if is_hl:
                f.color.rgb = hex_rgb(INK)
                _set_highlight(run, ACCENT_COLOR)
            elif is_ul:
                f.color.rgb = hex_rgb(ACCENT_COLOR)
            else:
                f.color.rgb = hex_rgb(color_hex)
    return box

def build_pptx(slides_data, out_path, logo_path):
    prs = Presentation()
    prs.slide_width = Emu(EMU_W)
    prs.slide_height = Emu(EMU_H)
    blank = prs.slide_layouts[6]

    for s in slides_data:
        st = style_of(s["theme"])
        slide = prs.slides.add_slide(blank)

        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb = hex_rgb(st["bg"])
        bg.line.fill.background(); bg.shadow.inherit = False

        # ---- DUO ----
        if s.get("duo"):
            panels = s["panels"]
            if len(panels) >= 1:
                slide.shapes.add_picture(panels[0]["image_path"], 0, 0, width=prs.slide_width, height=fy(0.5))
            if len(panels) >= 2:
                slide.shapes.add_picture(panels[1]["image_path"], 0, fy(0.5), width=prs.slide_width, height=fy(0.5))
            if logo_path:
                try:
                    slide.shapes.add_picture(logo_path, fx(0.06), fy(0.05), height=fy(0.045))
                except Exception:
                    pass
            _add_rich(slide, 0.55, 0.05, 0.39, 0.07, HASHTAG, 12, True, BODY_FONT, ACCENT_COLOR,
                      align=PP_ALIGN.RIGHT, highlight=False)
            if len(panels) >= 1 and panels[0]["caption"]:
                _add_rich(slide, 0.07, 0.30, 0.86, 0.16, panels[0]["caption"], head_pt(20), True, BODY_FONT, "FFFFFF")
            if len(panels) >= 2 and panels[1]["caption"]:
                _add_rich(slide, 0.07, 0.72, 0.86, 0.16, panels[1]["caption"], head_pt(20), True, BODY_FONT, "FFFFFF")
            _add_rich(slide, 0.08, 0.90, 0.55, 0.08, FOOTER_TEXT, 13, True, BODY_FONT, "FFFFFF", highlight=False)
            if s["total"] > 1:
                _add_rich(slide, 0.82, 0.90, 0.12, 0.07, ARROW_TEXT, 24, True, BODY_FONT, "FFFFFF",
                          align=PP_ALIGN.RIGHT, highlight=False)
            continue

        if st["photo"] == "full" and s.get("image_path"):
            slide.shapes.add_picture(s["image_path"], 0, 0, width=prs.slide_width, height=prs.slide_height)
        elif st["photo"] == "frame" and s.get("image_path"):
            slide.shapes.add_picture(s["image_path"], fx(0.07), fy(0.11), width=fx(0.86), height=fy(0.34))

        # logo + hashtag
        if logo_path:
            try:
                slide.shapes.add_picture(logo_path, fx(0.06), fy(0.05), height=fy(0.045))
            except Exception:
                pass
        _add_rich(slide, 0.55, 0.05, 0.39, 0.07, HASHTAG, 12, True, BODY_FONT, ACCENT_COLOR,
                  align=PP_ALIGN.RIGHT, highlight=False)

        # tanda kutip hijau besar (Kutipan)
        if st["quote"]:
            _add_rich(slide, 0.07, 0.30, 0.2, 0.12, "“", 90, True, SERIF_FONT, ACCENT_COLOR, highlight=False)

        align = PP_ALIGN.CENTER if st["align"] == "center" else PP_ALIGN.LEFT
        head_font = SERIF_FONT if st["serif"] else HEADLINE_FONT
        left = 0.07; width = 0.86
        if s["headline"]:
            _add_rich(slide, left, st["head_top"], width, 0.20, s["headline"],
                      head_pt(st["head_size"] * 0.66), True, head_font, st["head"], align=align)
        if s["body"]:
            _add_rich(slide, left, st["body_top"], width, 0.24, s["body"],
                      body_pt(st["body_size"] * 0.62), False, BODY_FONT, st["body"], align=align)

        _add_rich(slide, 0.08, 0.90, 0.55, 0.08, FOOTER_TEXT, 13, True, BODY_FONT, st["foot"], highlight=False)
        if s["total"] > 1:
            _add_rich(slide, 0.82, 0.90, 0.12, 0.07, ARROW_TEXT, 24, True, BODY_FONT, st["head"],
                      align=PP_ALIGN.RIGHT, highlight=False)

    prs.save(out_path)
    return out_path


# ======================================================================
# 7. RAKIT .SVG (untuk Figma)
# ======================================================================
def _img_data_uri(path):
    im = Image.open(path).convert("RGB")
    buf = io.BytesIO(); im.save(buf, "JPEG", quality=85)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()

def _svg_escape(t):
    return html.escape(t or "", quote=True)

def _svg_rich(text, x0, y0, width, size, bold, font, color_hex, align="left", gap=1.34):
    weight = "700" if bold else "400"
    char_w = size * (0.64 if bold else 0.58)
    space_w = size * 0.46
    out = []
    y = y0
    for src_line in text.split("\n"):
        words = rich_words(src_line)
        if not words:
            y += int(size * gap); continue
        vlines, cur, cur_w = [], [], 0
        for w, hl, it, ul, cc in words:
            ww = len(w) * char_w
            add = ww + (space_w if cur else 0)
            if cur and cur_w + add > width:
                vlines.append((cur, cur_w)); cur, cur_w = [], 0
                add = ww
            cur.append((w, hl, it, ul, cc, ww)); cur_w += add
        if cur:
            vlines.append((cur, cur_w))
        for vl, line_w in vlines:
            baseline = y + size
            start_x = x0
            if align == "center":
                start_x = x0 + max(0, (width - line_w) / 2)
            positions = []
            cx = start_x
            for k, tup in enumerate(vl):
                if k > 0:
                    cx += space_w
                positions.append(cx)
                cx += tup[5]
            # dekorasi: gabung run berurutan yang berflag sama
            def runs(flag_idx):
                k = 0
                while k < len(vl):
                    if vl[k][flag_idx]:
                        j = k
                        while j + 1 < len(vl) and vl[j + 1][flag_idx]:
                            j += 1
                        yield positions[k], positions[j] + vl[j][5]
                        k = j + 1
                    else:
                        k += 1
            # stabilo (hl) idx1
            for xs, xe in runs(1):
                pad = size * 0.12
                out.append(f'<rect x="{int(xs-pad)}" y="{int(baseline-size*0.86)}" '
                           f'width="{int(xe-xs+pad*2)}" height="{int(size*1.10)}" rx="{int(size*0.10)}" '
                           f'fill="#{ACCENT_COLOR}"/>')
            # lingkaran (circ) idx4
            for xs, xe in runs(4):
                pad = size * 0.22
                cxp = (xs + xe) / 2
                rxp = (xe - xs) / 2 + pad
                out.append(f'<ellipse cx="{int(cxp)}" cy="{int(baseline-size*0.34)}" '
                           f'rx="{int(rxp)}" ry="{int(size*0.72)}" fill="none" '
                           f'stroke="#{ACCENT_COLOR}" stroke-width="{max(3,int(size*0.06))}"/>')
            # garis bawah (ul) idx3
            for xs, xe in runs(3):
                out.append(f'<rect x="{int(xs)}" y="{int(baseline+size*0.14)}" '
                           f'width="{int(xe-xs)}" height="{max(3,int(size*0.09))}" fill="#{ACCENT_COLOR}"/>')
            # teks
            for (w, hl, it, ul, cc, ww), px in zip(vl, positions):
                fill = INK if hl else color_hex
                style = ' font-style="italic"' if it else ""
                out.append(f'<text x="{int(px)}" y="{int(baseline)}"{style} '
                           f'font-family="{_svg_escape(font)}, Arial, sans-serif" '
                           f'font-size="{size}" font-weight="{weight}" fill="#{fill}">{_svg_escape(w)}</text>')
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
        st = style_of(s["theme"])
        parts.append(f'<rect x="{xo}" y="0" width="{CANVAS_W}" height="{H}" fill="#{st["bg"]}"/>')

        # ---- DUO ----
        if s.get("duo"):
            panels = s["panels"]
            if len(panels) >= 1:
                parts.append(f'<image x="{xo}" y="0" width="{CANVAS_W}" height="{int(0.5*H)}" '
                             f'preserveAspectRatio="xMidYMid slice" href="{_img_data_uri(panels[0]["image_path"])}"/>')
            if len(panels) >= 2:
                parts.append(f'<image x="{xo}" y="{int(0.5*H)}" width="{CANVAS_W}" height="{int(0.5*H)}" '
                             f'preserveAspectRatio="xMidYMid slice" href="{_img_data_uri(panels[1]["image_path"])}"/>')
            if logo_uri:
                parts.append(f'<image x="{xo + int(0.06*CANVAS_W)}" y="{int(0.05*H)}" '
                             f'height="{int(0.05*H)}" href="{logo_uri}"/>')
            parts.append(f'<text x="{xo + int(0.94*CANVAS_W)}" y="{int(0.075*H)}" text-anchor="end" '
                         f'font-family="{_svg_escape(BODY_FONT)}, Arial, sans-serif" font-size="22" '
                         f'font-weight="700" fill="#{ACCENT_COLOR}">{_svg_escape(HASHTAG)}</text>')
            if len(panels) >= 1 and panels[0]["caption"]:
                svg, _ = _svg_rich(panels[0]["caption"], xo + int(0.07*CANVAS_W), int(0.30*H),
                                   int(0.86*CANVAS_W), head_px(30), True, BODY_FONT, "FFFFFF")
                parts.append(svg)
            if len(panels) >= 2 and panels[1]["caption"]:
                svg, _ = _svg_rich(panels[1]["caption"], xo + int(0.07*CANVAS_W), int(0.72*H),
                                   int(0.86*CANVAS_W), head_px(30), True, BODY_FONT, "FFFFFF")
                parts.append(svg)
            parts.append(f'<rect x="{xo + int(0.06*CANVAS_W)}" y="{int(0.905*H)}" width="6" height="{int(0.05*H)}" fill="#{ACCENT_COLOR}"/>')
            fsvg, _ = _svg_rich(FOOTER_TEXT, xo + int(0.09*CANVAS_W), int(0.905*H), int(0.5*CANVAS_W), 20, True, BODY_FONT, "FFFFFF")
            parts.append(fsvg)
            if total > 1:
                parts.append(f'<text x="{xo + int(0.93*CANVAS_W)}" y="{int(0.95*H)}" text-anchor="end" '
                             f'font-family="{_svg_escape(BODY_FONT)}, Arial, sans-serif" font-size="40" '
                             f'font-weight="700" fill="#FFFFFF">{_svg_escape(ARROW_TEXT)}</text>')
            continue

        if st["photo"] == "full" and s.get("image_path"):
            parts.append(f'<image x="{xo}" y="0" width="{CANVAS_W}" height="{H}" '
                         f'preserveAspectRatio="xMidYMid slice" href="{_img_data_uri(s["image_path"])}"/>')
            parts.append(f'<rect x="{xo}" y="{int(0.32*H)}" width="{CANVAS_W}" height="{int(0.68*H)}" '
                         f'fill="#000000" opacity="0.5"/>')
        elif st["photo"] == "frame" and s.get("image_path"):
            fx0 = xo + int(0.07*CANVAS_W); fy0 = int(0.11*H)
            fw = int(0.86*CANVAS_W); fh = int(0.34*H)
            parts.append(f'<clipPath id="cp{idx}"><rect x="{fx0}" y="{fy0}" width="{fw}" height="{fh}" rx="28"/></clipPath>')
            parts.append(f'<image x="{fx0}" y="{fy0}" width="{fw}" height="{fh}" '
                         f'preserveAspectRatio="xMidYMid slice" clip-path="url(#cp{idx})" '
                         f'href="{_img_data_uri(s["image_path"])}"/>')

        if logo_uri:
            parts.append(f'<image x="{xo + int(0.06*CANVAS_W)}" y="{int(0.05*H)}" '
                         f'height="{int(0.05*H)}" href="{logo_uri}"/>')
        parts.append(f'<text x="{xo + int(0.94*CANVAS_W)}" y="{int(0.075*H)}" text-anchor="end" '
                     f'font-family="{_svg_escape(BODY_FONT)}, Arial, sans-serif" font-size="22" '
                     f'font-weight="700" fill="#{ACCENT_COLOR}">{_svg_escape(HASHTAG)}</text>')

        if st["quote"]:
            parts.append(f'<text x="{xo + int(0.07*CANVAS_W)}" y="{int(0.36*H)}" '
                         f'font-family="{_svg_escape(SERIF_FONT)}, Georgia, serif" font-size="150" '
                         f'font-weight="700" fill="#{ACCENT_COLOR}">“</text>')

        hx = xo + int(0.07*CANVAS_W); hw = int(0.86*CANVAS_W)
        if s["headline"]:
            svg, _ = _svg_rich(s["headline"], hx, int(st["head_top"]*H), hw, head_px(st["head_size"]), True,
                               SERIF_FONT if st["serif"] else HEADLINE_FONT, st["head"], align=st["align"])
            parts.append(svg)
        if s["body"]:
            svg, _ = _svg_rich(s["body"], hx, int(st["body_top"]*H), hw, body_px(st["body_size"]), False,
                               BODY_FONT, st["body"], align=st["align"])
            parts.append(svg)

        parts.append(f'<rect x="{xo + int(0.06*CANVAS_W)}" y="{int(0.905*H)}" width="6" height="{int(0.05*H)}" fill="#{ACCENT_COLOR}"/>')
        fsvg, _ = _svg_rich(FOOTER_TEXT, xo + int(0.09*CANVAS_W), int(0.905*H), int(0.5*CANVAS_W), 20, True, BODY_FONT, st["foot"])
        parts.append(fsvg)
        if total > 1:
            parts.append(f'<text x="{xo + int(0.93*CANVAS_W)}" y="{int(0.95*H)}" text-anchor="end" '
                         f'font-family="{_svg_escape(BODY_FONT)}, Arial, sans-serif" font-size="40" '
                         f'font-weight="700" fill="#{st["head"]}">{_svg_escape(ARROW_TEXT)}</text>')
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

def tg_buttons(text, keyboard):
    """Kirim pesan + tombol inline. keyboard = list of rows, tiap tombol (label, callback_data)."""
    kb = {"inline_keyboard": [[{"text": t, "callback_data": d} for (t, d) in row] for row in keyboard]}
    return _tg_call("sendMessage", {"chat_id": TELEGRAM_CHAT_ID, "text": text[:3500],
                                    "reply_markup": json.dumps(kb)}, label="btn")

def tg_delete_webhook():
    """Pastikan nggak ada webhook nyangkut. Webhook bikin getUpdates MANDEK (balik kosong),
    jadi tombol/ketikan user nggak pernah kebaca walau bot tetap bisa KIRIM pesan.
    drop_pending_updates=false -> pesan yg ketahan TIDAK dibuang (langsung kebaca sesudahnya)."""
    try:
        r = requests.post(f"{TG_BASE}/deleteWebhook",
                          data={"drop_pending_updates": "false"}, timeout=20)
        j = r.json()
        print(f"    deleteWebhook: ok={j.get('ok')} {j.get('description','')}")
    except Exception as e:
        print("    ! deleteWebhook error:", e)

def tg_get_updates(offset):
    try:
        r = requests.get(f"{TG_BASE}/getUpdates",
                         params={"offset": offset, "timeout": 0, "allowed_updates": json.dumps(["callback_query", "message"])},
                         timeout=40)
        j = r.json()
        if not j.get("ok"):
            print(f"    ! getUpdates NOT ok: {str(j)[:200]}")
        return j.get("result", []) if j.get("ok") else []
    except Exception as e:
        print("    ! getUpdates error:", e)
        return []

def tg_answer_callback(cb_id, text=""):
    # notif kecil di tombol; kalau gagal (query kedaluwarsa krn batch) abaikan diam-diam
    try:
        requests.post(f"{TG_BASE}/answerCallbackQuery",
                      data={"callback_query_id": cb_id, "text": text[:180]}, timeout=20)
    except Exception:
        pass


# ======================================================================
# 8b. SISTEM BELAJAR — memori + feedback (approve/reject + alasan)
# ======================================================================
DEFAULT_MEMORY = {
    "preferences": {"head_scale": 1.0, "body_scale": 1.0, "scrim": 1.0, "soft": 0.0},
    "pending": {},        # token -> {title, page_id, ts, revision}
    "resolved": [],       # token yang sudah dinilai (anti dobel)
    "awaiting_text": None,  # token yang menunggu alasan ketik
    "revise_queue": {},   # page_id -> {title, revision, reasons[]} konten yg minta direvisi
    "log": [],            # riwayat {ts, title, status, reason}
    "stats": {"approved": 0, "rejected": 0, "weekly": []},
    "tg_offset": 0,
}

MAX_REVISI = int(os.environ.get("MAX_REVISI") or 10)   # batas aman auto-revisi per konten

# tombol alasan saat Reject: (label, kode)
REASON_BUTTONS = [
    ("Teks kegedean", "tbig"),
    ("Teks kekecilan", "tsmall"),
    ("Teks susah kebaca", "hard"),
    ("Overlay kegelapan", "dark"),
    ("Gambar kurang aesthetic", "img"),
]
REASON_LABEL = {k: v for v, k in REASON_BUTTONS}

def load_memory():
    try:
        with open(MEMORY_PATH, "r", encoding="utf-8") as f:
            mem = json.load(f)
        for k, v in DEFAULT_MEMORY.items():
            mem.setdefault(k, v if not isinstance(v, (dict, list)) else (dict(v) if isinstance(v, dict) else list(v)))
        for k, v in DEFAULT_MEMORY["preferences"].items():
            mem["preferences"].setdefault(k, v)
        return mem
    except Exception:
        import copy
        return copy.deepcopy(DEFAULT_MEMORY)

def save_memory(mem):
    try:
        with open(MEMORY_PATH, "w", encoding="utf-8") as f:
            json.dump(mem, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("    ! gagal simpan memori:", e)

def apply_reason(prefs, code):
    """Geser knob sesuai alasan. Mengembalikan deskripsi singkat perubahan."""
    if code == "tbig":
        prefs["head_scale"] = _clamp(prefs["head_scale"] - 0.06, 0.7, 1.3)
        prefs["body_scale"] = _clamp(prefs["body_scale"] - 0.05, 0.7, 1.3)
        return "teks dikecilin"
    if code == "tsmall":
        prefs["head_scale"] = _clamp(prefs["head_scale"] + 0.06, 0.7, 1.3)
        prefs["body_scale"] = _clamp(prefs["body_scale"] + 0.05, 0.7, 1.3)
        return "teks digedein"
    if code == "hard":
        prefs["scrim"] = _clamp(prefs["scrim"] + 0.12, 0.6, 1.6)
        return "overlay dipergelap (biar teks kebaca)"
    if code == "dark":
        prefs["scrim"] = _clamp(prefs["scrim"] - 0.12, 0.6, 1.6)
        return "overlay dipertipis"
    if code == "img":
        prefs["soft"] = _clamp(prefs["soft"] + 0.08, 0.0, 0.4)
        return "gambar dibikin lebih soft/film"
    return "dicatat (tanpa ubah setelan)"

def prefs_summary(prefs):
    return (f"Setelan belajar sekarang: judul {int(prefs['head_scale']*100)}%, "
            f"body {int(prefs['body_scale']*100)}%, overlay {int(prefs['scrim']*100)}%, "
            f"soft +{int(prefs['soft']*100)}%.")


# ---------- OTAK AI: Groq (utama) + Gemini (cadangan). Dua-duanya gratis ----------
GEMINI_API_KEY  = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL    = os.environ.get("GEMINI_MODEL") or "gemini-3.8-flash"
GROQ_API_KEY    = os.environ.get("GROQ_API_KEY", "")
GROQ_MODEL      = os.environ.get("GROQ_MODEL") or "openai/gpt-oss-20b"
AI_CHAT_ENABLED = ((os.environ.get("AI_CHAT_ENABLED") or "true").lower() == "true"
                   and (bool(GEMINI_API_KEY) or bool(GROQ_API_KEY)))
# cek mandiri pakai 'mata' AI (Gemini vision). gratis (gambar sbg INPUT), cuma butuh Gemini key.
VISION_CHECK_ENABLED = ((os.environ.get("VISION_CHECK_ENABLED") or "true").lower() == "true"
                        and bool(GEMINI_API_KEY))

def _gemini_json(prompt):
    if not GEMINI_API_KEY:
        return None
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    body = {"contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.3, "responseMimeType": "application/json"}}
    try:
        r = requests.post(url, json=body, timeout=45)
        if r.status_code != 200:
            print(f"    ! Gemini {r.status_code}: {r.text[:120]}")
            return None
        return json.loads(r.json()["candidates"][0]["content"]["parts"][0]["text"])
    except Exception as e:
        print("    ! Gemini gagal:", e); return None

def _groq_json(prompt):
    if not GROQ_API_KEY:
        return None
    try:
        r = requests.post("https://api.groq.com/openai/v1/chat/completions",
                          headers={"Authorization": f"Bearer {GROQ_API_KEY}"},
                          json={"model": GROQ_MODEL, "temperature": 0.3,
                                "response_format": {"type": "json_object"},
                                "messages": [{"role": "user", "content": prompt}]},
                          timeout=45)
        if r.status_code != 200:
            print(f"    ! Groq {r.status_code}: {r.text[:120]}")
            return None
        return json.loads(r.json()["choices"][0]["message"]["content"])
    except Exception as e:
        print("    ! Groq gagal:", e); return None

AI_PRIMARY = (os.environ.get("AI_PRIMARY") or "groq").lower()   # 'groq' (stabil) atau 'gemini'

def ai_json(prompt):
    """Minta jawaban JSON ke AI. Default: Groq dulu (lebih stabil), Gemini cadangan.
    Bisa dibalik lewat secret AI_PRIMARY=gemini."""
    if not AI_CHAT_ENABLED:
        return None
    if AI_PRIMARY == "gemini":
        first, second, second_name = _gemini_json, _groq_json, "Groq"
    else:
        first, second, second_name = _groq_json, _gemini_json, "Gemini"
    data = first(prompt)
    if data is None:
        data = second(prompt)   # cadangan kalau yang utama ngadat
        if data is not None:
            print(f"    (pakai {second_name} sebagai cadangan)")
    return data

AI_PROMPT = """Kamu asisten desain untuk bot konten Instagram (brand Mindjiwa — platform wellness mental).
User memberi perintah/feedback (bahasa Indonesia atau Inggris) untuk merevisi sebuah desain.
Terjemahkan jadi penyesuaian angka. Setelan sekarang (1.0 = normal):
- head_scale: ukuran JUDUL (batas 0.7-1.3)
- body_scale: ukuran teks BODY (batas 0.7-1.3)
- scrim: kegelapan overlay di belakang teks (batas 0.6-1.6; naik = lebih gelap = teks lebih terbaca)
- soft: kesan film/lembut pada gambar (batas 0.0-0.4; naik = lebih lembut/pudar; turun = lebih cerah/tajam)
PENTING: hampir SEMUA masukan user = permintaan revisi -> set action "revise".
Keluhan/komentar sekecil apapun soal desain (foto/teks/warna/overlay) = "revise".
Contoh yang HARUS "revise": "gambar kurang bagus", "kurang aesthetic", "ganti foto",
"fotonya jelek", "judul kegedean", "kurang jelas", "overlay kurang gelap".
Kalau user komentarin/minta ganti foto, set action "revise" (foto otomatis diganti baru saat revisi).
Pakai action "none" HANYA kalau pesan jelas cuma sapaan/basa-basi/terima kasih/pertanyaan umum
(mis. "halo", "makasih", "lagi apa") — BUKAN komentar soal desain.
Balas HANYA JSON valid, tanpa teks lain:
{"head_scale_delta": <angka -0.12..0.12>, "body_scale_delta": <angka>, "scrim_delta": <angka>, "soft_delta": <angka>, "note": "<ringkasan singkat dalam bahasa Indonesia, maks 12 kata>", "action": "revise" atau "none"}
Pesan user: "%s" """

def ai_parse_command(text):
    """Ubah kalimat bebas jadi penyesuaian angka (Groq/Gemini). None kalau gagal/mati."""
    return ai_json(AI_PROMPT % text[:400])

def apply_ai(prefs, data):
    """Terapkan delta dari AI (dibatasi aman). Kembalikan daftar yg berubah."""
    changed = []
    plan = [("head_scale_delta", "head_scale", 0.7, 1.3),
            ("body_scale_delta", "body_scale", 0.7, 1.3),
            ("scrim_delta", "scrim", 0.6, 1.6),
            ("soft_delta", "soft", 0.0, 0.4)]
    for key, knob, lo, hi in plan:
        try:
            d = float(data.get(key, 0) or 0)
        except Exception:
            d = 0.0
        d = _clamp(d, -0.15, 0.15)   # batasi per-perintah biar nggak loncat jauh
        if abs(d) >= 0.01:
            prefs[knob] = _clamp(prefs[knob] + d, lo, hi)
            changed.append(knob)
    return changed

def _latest_pending(mem):
    best = None
    for tok, rec in mem.get("pending", {}).items():
        if best is None or rec.get("ts", 0) > best[1].get("ts", 0):
            best = (tok, rec)
    return best

def ai_image_keywords(slide_texts):
    """1 panggilan AI (Groq/Gemini) -> kata kunci FOTO (Inggris, aesthetic) per slide. None kalau gagal."""
    if not AI_CHAT_ENABLED or not slide_texts:
        return None
    joined = "\n".join(f"{i+1}. {(t or '')[:160]}" for i, t in enumerate(slide_texts))
    prompt = (
        "Kamu art director untuk brand wellness mental (Mindjiwa) — gaya editorial/aesthetic lifestyle, "
        "nuansa tenang, soft, natural light, tone lembut (BUKAN religius).\n"
        "Untuk TIAP slide di bawah, buat 1 kata kunci pencarian FOTO STOK dalam BAHASA INGGRIS (3-5 kata), "
        "fokus ke suasana/visual (bukan terjemahan harfiah teks), hindari tulisan/logo, utamakan natural light & tone lembut.\n"
        'Balas HANYA JSON object: {"keywords": ["...", "..."]} dengan panjang array PERSIS sama dengan jumlah slide.\n\n' + joined)
    data = ai_json(prompt)
    try:
        arr = (data or {}).get("keywords")
        if isinstance(arr, list) and arr:
            return [str(x) for x in arr]
    except Exception:
        pass
    return None

def _gemini_vision_json(prompt, image_path):
    """Kirim gambar + pertanyaan ke Gemini (mode 'mata'/vision), minta jawaban JSON.
    Gambar dikecilin dulu (max 640px, JPEG) biar payload kecil & cepat."""
    if not GEMINI_API_KEY:
        return None
    try:
        im = Image.open(image_path).convert("RGB")
        im.thumbnail((640, 640))
        buf = io.BytesIO(); im.save(buf, "JPEG", quality=80)
        b64 = base64.b64encode(buf.getvalue()).decode()
    except Exception as e:
        print("    ! vision: gagal siapin gambar:", e); return None
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    body = {"contents": [{"parts": [
                {"text": prompt},
                {"inline_data": {"mime_type": "image/jpeg", "data": b64}}]}],
            "generationConfig": {"temperature": 0.2, "responseMimeType": "application/json"}}
    try:
        r = requests.post(url, json=body, timeout=60)
        if r.status_code != 200:
            print(f"    ! Gemini vision {r.status_code}: {r.text[:120]}")
            return None
        return json.loads(r.json()["candidates"][0]["content"]["parts"][0]["text"])
    except Exception as e:
        print("    ! Gemini vision gagal:", e); return None

def ai_vision_check(image_path, zone, theme):
    """Cek mandiri: kirim foto COVER slide ke Gemini, tanya apakah teks bakal mudah dibaca.
    Kembalikan {ok: bool, issue: str} atau None kalau mati/gagal."""
    if not VISION_CHECK_ENABLED:
        return None
    zmap = {"top": "atas", "bottom": "bawah", "left": "kiri", "right": "kanan", "center": "tengah"}
    warna = "putih" if theme == "dark" else "gelap (hampir hitam)"
    prompt = (
        "Kamu reviewer desain konten Instagram. Ini FOTO cover sebuah slide.\n"
        f"Rencana: logo/hashtag + teks diletakkan di area {zmap.get(zone, zone)}, warna teks {warna}.\n"
        "Nilai apakah di area itu teks akan MUDAH DIBACA: kontras cukup, area tidak terlalu ramai/berpola, "
        "dan tidak menutupi wajah/objek penting.\n"
        'Balas HANYA JSON: {"ok": true atau false, "issue": "<masalah singkat bahasa Indonesia maks 12 kata; '
        'kosongkan kalau sudah ok>"}'
    )
    data = _gemini_vision_json(prompt, image_path)   # 1x aja (retry malah boros kuota)
    if not isinstance(data, dict):
        return None
    return {"ok": bool(data.get("ok", True)), "issue": str(data.get("issue") or "").strip()}

def _queue_revision(mem, rec, reason_label):
    """Masukkan konten ke antrian revisi (biar run berikutnya di-generate ulang)."""
    page_id = ((rec or {}).get("page_id") or "").replace("-", "")   # normalisasi kunci
    if not page_id:
        return False
    rev = (rec or {}).get("revision", 0)
    q = mem.setdefault("revise_queue", {})
    entry = q.get(page_id) or {"title": (rec or {}).get("title", "desain"), "revision": rev, "reasons": []}
    entry["title"] = (rec or {}).get("title", entry.get("title", "desain"))
    entry["revision"] = rev + 1
    entry["reasons"].append(reason_label)
    entry["reasons"] = entry["reasons"][-8:]
    q[page_id] = entry
    return True

def _mark_resolved(mem, tok):
    res = mem.setdefault("resolved", [])
    if tok not in res:
        res.append(tok)
    if len(res) > 500:
        del res[:len(res) - 500]

def process_feedback(mem):
    """Baca pencetan tombol sejak run terakhir, update stats + knob.
    Knob diterapkan BERDASARKAN alasan yg dipencet (tidak tergantung 'pending'),
    biar tetap jalan walau catatan pending hilang. Anti-dobel pakai daftar 'resolved'."""
    prefs = mem["preferences"]
    tg_delete_webhook()   # jaga-jaga: webhook nyangkut bikin tombol/ketikan nggak kebaca
    updates = tg_get_updates(mem.get("tg_offset", 0) + 1)
    changed = []
    processed = 0
    for up in updates:
        mem["tg_offset"] = max(mem.get("tg_offset", 0), up.get("update_id", 0))

        # ---- pesan teks: perintah stats / reset / revisi bebas (AI) ----
        msg = up.get("message")
        if msg:
            text = (msg.get("text") or "").strip()
            low = text.lower()
            mem["awaiting_text"] = None   # fitur lama nggak dipakai lagi
            if low in ("stats", "/stats", "statistik"):
                a = mem["stats"]["approved"]; r = mem["stats"]["rejected"]; tot = a + r
                rate = int(100 * a / tot) if tot else 0
                tg_message(f"📊 Statistik belajar Mindjiwa:\n"
                           f"Approve {a} / Reject {r}  (approval rate {rate}%).\n"
                           f"{prefs_summary(mem['preferences'])}")
                continue
            if low in ("/reset", "reset belajar"):
                mem["preferences"] = dict(DEFAULT_MEMORY["preferences"])
                prefs = mem["preferences"]
                tg_message("🔄 Setelan belajar direset ke awal (netral). Statistik tetap tersimpan.")
                continue
            # ---- CHAT AI: perintah revisi bebas (ketik kalimat biasa) ----
            if text:
                data = ai_parse_command(text)
                # PENGAMAN: AI bilang "none" tapi teks jelas keluhan + ada desain nunggu -> tetap REVISI
                _rev_kata = ("ganti", "ubah", "revisi", "perbaiki", "kurang", "jelek", "buram",
                             "norak", "kegedean", "kekecilan", "lebih ", "jangan", "terlalu")
                if (data and data.get("action") == "none"
                        and any(k in low for k in _rev_kata) and _latest_pending(mem)):
                    data["action"] = "revise"
                if data and data.get("action") == "none":
                    tg_message("Oke, dicatat 🙂 Kalau mau gw revisi, kasih tau yang perlu diubah ya "
                               "(misal: “judul kekecilan”, “overlay kurang gelap”, “ganti foto lebih cerah”).")
                    continue
                note = (data or {}).get("note") or text[:80]
                ch = apply_ai(prefs, data) if data else []
                latest = _latest_pending(mem)
                if latest:
                    tok, rec = latest
                    _queue_revision(mem, rec, "chat: " + note)
                    mem["pending"].pop(tok, None); _mark_resolved(mem, tok)
                    mem["stats"]["rejected"] += 1
                    extra = ("\n" + prefs_summary(prefs)) if ch else ""
                    tail = "" if data else " (AI lagi nggak aktif, jadi gw revisi dgn foto baru aja)"
                    tg_message(f"🧠 Paham: {note}. Gw revisi '{rec.get('title','desain')}' ya{tail}.{extra}")
                else:
                    if ch:
                        tg_message(f"🧠 Paham: {note}. Setelan disesuaikan, kepakai di desain berikutnya.\n{prefs_summary(prefs)}")
                    else:
                        tg_message("Belum ada desain yang bisa direvisi. Kirim konten dulu ya 🙂")
                processed += 1
                continue
            continue

        # ---- pencetan tombol (callback) ----
        cb = up.get("callback_query")
        if not cb:
            continue
        data = cb.get("data", "")
        cb_id = cb.get("id", "")
        parts = data.split(":")
        kind = parts[0]
        tok = parts[1] if len(parts) > 1 else ""
        if tok and tok in mem.get("resolved", []):
            tg_answer_callback(cb_id, "Sudah dinilai sebelumnya 👍"); continue

        # page_id & code diambil dari DATA TOMBOL (nggak gantung ke pending yg bisa hilang)
        if kind == "rs":
            code = parts[2] if len(parts) > 2 else ""
            pid_cb = parts[3] if len(parts) > 3 else ""
        else:
            code = ""
            pid_cb = parts[2] if len(parts) > 2 else ""
        pend = mem["pending"].get(tok, {})
        rec = {"title": pend.get("title", "desain"),
               "page_id": pid_cb or pend.get("page_id"),
               "revision": pend.get("revision", 0)}

        if kind == "a":                      # approve (1x pencet) -> FINAL
            mem["pending"].pop(tok, None)
            mem["stats"]["approved"] += 1
            _mark_resolved(mem, tok); processed += 1
            pid_norm = (rec.get("page_id") or "").replace("-", "")
            if pid_norm:
                mem.get("revise_queue", {}).pop(pid_norm, None)   # batal revisi, udah oke
            mem["log"].append({"ts": int(time.time()), "title": rec["title"],
                               "status": "approved", "reason": ""})
            tg_answer_callback(cb_id, "✅ Disimpan sebagai contoh bagus!")
        elif kind == "rs":                   # reject + alasan (1x pencet) -> knob + antri revisi
            if code == "other":   # tombol lama: sekarang cukup KETIK langsung
                tg_answer_callback(cb_id, "Ketik aja perintah revisimu langsung ya 🙂")
                tg_message("💬 Ketik aja apa yang mau diubah (misal: “judul kekecilan, foto lebih cerah”) — "
                           "langsung gw proses, nggak usah pencet tombol ini.")
            else:
                mem["stats"]["rejected"] += 1
                desc = apply_reason(prefs, code)
                mem["pending"].pop(tok, None)
                _mark_resolved(mem, tok); processed += 1
                ok = _queue_revision(mem, rec, REASON_LABEL.get(code, code))
                print(f"  reject({code}) pid={rec.get('page_id')} queued={ok}")
                mem["log"].append({"ts": int(time.time()), "title": rec["title"],
                                   "status": "rejected", "reason": REASON_LABEL.get(code, code)})
                changed.append(desc)
                tg_answer_callback(cb_id, f"Paham. {desc}. Gw revisi ya.")
        elif kind == "r":                    # kompat tombol lama (Reject 2-langkah) -> abaikan halus
            tg_answer_callback(cb_id, "Pakai tombol alasan di desain baru ya 🙏")

    # log diagnostik di console (biar kelihatan apa yg kebaca)
    print(f"  feedback: {len(updates)} update, {processed} diproses, "
          f"revise_queue={len(mem.get('revise_queue', {}))}, offset={mem.get('tg_offset')}")

    # laporan ke Telegram biar KELIHATAN bot baca feedback
    if processed > 0:
        head = f"🔎 Feedback terbaca: {processed} pencetan diproses."
        if changed:
            head += "\n🧠 Bot menyesuaikan diri: " + "; ".join(changed) + "."
        head += "\n" + prefs_summary(prefs)
        tg_message(head)
    mem["_last_processed"] = processed
    return mem

def apply_prefs_to_globals(mem):
    global G_HEAD_SCALE, G_BODY_SCALE, G_SCRIM, G_SOFT
    p = mem.get("preferences", {})
    G_HEAD_SCALE = float(p.get("head_scale", 1.0))
    G_BODY_SCALE = float(p.get("body_scale", 1.0))
    G_SCRIM      = float(p.get("scrim", 1.0))
    G_SOFT       = float(p.get("soft", 0.0))


# ======================================================================
# 9. PROSES SATU KONTEN
# ======================================================================
def process_page(page, workdir, mem=None, revision=None):
    props = page["properties"]
    page_id = page.get("id", "")
    title = get_title(props) or "Tanpa Judul"
    fmt = read_property(props, FORMAT_PROPERTY, "select") or "Single Post"
    content = read_property(props, CONTENT_PROPERTY, "rich_text")
    keywords = read_property(props, KEYWORD_PROPERTY, "rich_text")
    theme = (read_property(props, THEME_PROPERTY, "select")
             or read_property(props, THEME_PROPERTY, "status")
             or DEFAULT_THEME)
    st = style_of(theme)
    print(f"  [tema dibaca: '{theme}' -> style: {st['kind']}]")

    if not content.strip():
        raise ValueError("Kolom 'Konten' kosong.")

    is_single = "single" in fmt.lower()
    slide_blocks = split_slides(content)
    if is_single:
        slide_blocks = slide_blocks[:1]
    keyword_blocks = split_slides(keywords)
    # kata kunci gambar pintar (AI) — 1 panggilan utk semua slide; None kalau AI mati/gagal.
    # cuma kepakai di jalur auto (tak tempel link), tak mengubah sumber/gaya gambar.
    ai_kws = ai_image_keywords(slide_blocks)
    if ai_kws:
        print(f"  [AI keywords: {ai_kws}]")

    print(f"  -> '{title}' | {fmt} | {len(slide_blocks)} slide | tema={theme} | pensil={is_single}")

    slides_data, preview_paths, credits = [], [], []
    vision_notes = []           # catatan cek mandiri (AI vision) — cuma cover
    used_ids = set()
    total = len(slide_blocks)
    for i, block in enumerate(slide_blocks, start=1):
        # ---- Style DUO: 2 foto ditumpuk + caption (panel dipisah '///') ----
        if st["kind"] == "duo":
            panels_raw = [p.strip() for p in re.split(r"\s*///\s*", block) if p.strip()][:2]
            if not panels_raw:
                raise ValueError(f"Slide {i} (Duo) kosong. Format: URL1 + caption /// URL2 + caption.")
            panels = []
            for pnl in panels_raw:
                purl, pclean = extract_url_from_block(pnl)
                cap = pclean.strip()
                if purl:
                    pim = Image.open(io.BytesIO(fetch_image_bytes(purl))).convert("RGB")
                else:
                    q = keyword_for(0, [], cap, "")
                    data, credit = image_pick(q, used_ids)
                    if data:
                        pim = Image.open(io.BytesIO(data)).convert("RGB")
                        if credit:
                            credits.append(f"Slide {i}: {credit}")
                    else:
                        pim = make_placeholder()
                    if credit:
                        credits.append(f"Slide {i}: {credit}")
                pim = darken(crop_to(pim, CANVAS_W, int(0.5 * CANVAS_H)), 0.42)
                ppath = os.path.join(workdir, f"slide_{i}_{len(panels)+1}.png")
                pim.save(ppath, "PNG")
                preview_paths.append(ppath)
                panels.append({"image_path": ppath, "caption": cap})
            slides_data.append({"duo": True, "panels": panels, "headline": "", "body": "",
                                "image_path": None, "index": i, "total": total, "theme": theme})
            continue

        inline_url, clean = extract_url_from_block(block)
        headline, body = headline_and_body(clean)

        img_path = None
        if st["photo"] != "none":
            if inline_url:
                im = Image.open(io.BytesIO(fetch_image_bytes(inline_url))).convert("RGB")
            else:
                explicit = keyword_blocks[i - 1].strip() if (i - 1) < len(keyword_blocks) else ""
                if explicit:
                    q = keyword_for(i - 1, keyword_blocks, headline, body)
                elif ai_kws and (i - 1) < len(ai_kws) and ai_kws[i - 1].strip():
                    q = ai_kws[i - 1].strip()          # kata kunci FOTO dari AI (Inggris)
                else:
                    q = keyword_for(i - 1, keyword_blocks, headline, body)
                data, credit = image_pick(q, used_ids)
                if data:
                    im = Image.open(io.BytesIO(data)).convert("RGB")
                    if credit:
                        credits.append(f"Slide {i}: {credit}")
                else:
                    im = make_placeholder()
                    credits.append(f"Slide {i}: (background netral — gambar '{q}' tak ditemukan)")
            if st["photo"] == "frame":
                im = crop_to(im, int(0.86 * CANVAS_W), int(0.34 * CANVAS_H))
            else:
                im = crop_to(im, CANVAS_W, CANVAS_H)
            if is_single:
                im = pencil_sketch(im)
            img_path = os.path.join(workdir, f"slide_{i}.png")
            im.save(img_path, "PNG")
            preview_paths.append(img_path)

        slides_data.append({"headline": headline, "body": body, "image_path": img_path,
                            "index": i, "total": total, "theme": theme})

    # --- CEK MANDIRI (AI vision) — CUKUP slide COVER biar hemat kuota gratis Gemini ---
    cover_img = next((s.get("image_path") for s in slides_data if s.get("image_path")), None)
    if cover_img:
        # warna teks di cover: tema gelap/kutipan/duo -> teks putih (dark); krem/kertas -> teks gelap (light)
        vis_theme = "dark" if st["head"] == "FFFFFF" else "light"
        chk = ai_vision_check(cover_img, "bottom", vis_theme)
        if chk is None:
            print("    [vision cover] dilewati/gagal (kuota/AI nggak jawab)")
        elif not chk["ok"]:
            vision_notes.append(f"Cover: ⚠️ {chk['issue'] or 'teks mungkin kurang terbaca'}")
            print(f"    [vision cover] ⚠️ {chk['issue']}")
        else:
            vision_notes.append("Cover: ✅ aman")
            print("    [vision cover] aman")

    safe_name = re.sub(r"[^\w\- ]", "", title).strip().replace(" ", "_")[:40] or "desain"
    logo_path = download_logo_path()

    lines = []
    if revision:
        rev_n = revision.get("revision", 0)
        prev = ", ".join(revision.get("reasons", [])[-4:]) or "-"
        lines += [f"🔄 Merevisi '{title}' (revisi ke-{rev_n})",
                  f"⚠️ Catatan dari sebelumnya: {prev}", ""]
    lines += [f"🌿 {title}  ({fmt}, {total} slide, tema {theme})", ""]
    for i, s in enumerate(slides_data, start=1):
        h = s["headline"] or "(tanpa judul)"
        b = (s["body"][:150] + "…") if len(s["body"]) > 150 else s["body"]
        lines.append(f"— Slide {i}: {h}")
        if b:
            lines.append(f"  {b}")
    lines += ["", *credits]
    if VISION_CHECK_ENABLED:
        warn = [n for n in vision_notes if "⚠️" in n]
        if warn:
            lines += ["", "🔎 Cek mandiri (AI vision):", *warn,
                      "   (kalau mau diperbaiki, pencet Reject atau ketik revisimu)"]
        elif vision_notes:
            lines += ["", "🔎 Cek mandiri (cover): aman ✅"]
        else:
            lines += ["", "🔎 Cek mandiri: dilewati — kuota AI lagi penuh, nanti dicoba lagi"]
    tg_message("\n".join(lines))

    for i, p in enumerate(preview_paths, start=1):
        tg_photo(p, caption=f"Preview gambar slide {i}")

    # pembatas + nama konten sebelum kirim file
    tg_message(f"📎 ====================\nFILE DESAIN: {title}\n====================")

    if OUTPUT_PPTX:
        pptx_path = os.path.join(workdir, f"{safe_name}.pptx")
        build_pptx(slides_data, pptx_path, logo_path)
        if not tg_document(pptx_path, caption=f"{title} — .pptx (semua slide): import ke Canva ✨"):
            tg_message("⚠️ File .pptx gagal dikirim (cek log).")

    if OUTPUT_SVG:
        svg_path = os.path.join(workdir, f"{safe_name}.svg")
        build_svg_single(slides_data, svg_path, logo_path)
        if not tg_document(svg_path, caption=f"{title} — .svg (semua slide): tarik ke Figma ✨"):
            tg_message("⚠️ File .svg gagal dikirim (cek log).")

    if OUTPUT_SVG_PER_SLIDE and total > 1:
        for s in slides_data:
            sp = os.path.join(workdir, f"{safe_name}_slide{s['index']}.svg")
            build_svg_single([s], sp, logo_path)
            if not tg_document(sp, caption=f"{title} — slide {s['index']}/{total} (.svg per slide)"):
                tg_message(f"⚠️ SVG slide {s['index']} gagal dikirim (cek log).")

    # ---- tombol penilaian (1x pencet langsung selesai) utk sistem belajar ----
    if LEARN_ENABLED and mem is not None:
        rev_n = (revision or {}).get("revision", 0)
        token = secrets.token_hex(4)
        pid_cb = (page_id or "").replace("-", "")     # page_id ditaruh di tombol (anti-hilang)
        mem["pending"][token] = {"title": title, "page_id": page_id,
                                 "ts": int(time.time()), "revision": rev_n}
        rows = [[("✅ Approve (oke!)", f"a:{token}:{pid_cb}")]]
        rr = []
        for label, code in REASON_BUTTONS:
            rr.append(("❌ " + label, f"rs:{token}:{code}:{pid_cb}"))
            if len(rr) == 2:
                rows.append(rr); rr = []
        if rr:
            rows.append(rr)
        head = f"Nilai desain '{title}' 👇"
        if rev_n:
            prev = ", ".join((revision or {}).get("reasons", [])[-4:]) or "-"
            head = (f"🔄 REVISI ke-{rev_n} dari '{title}' 👇\n"
                    f"⚠️ Catatan dari sebelumnya: {prev}")
        tip = ("\n💬 atau ketik aja perintahmu (misal: “judul kekecilan, foto lebih cerah”)"
               if AI_CHAT_ENABLED else "")
        tg_buttons(head + "\n✅ kalau udah oke — atau pencet alasannya kalau masih kurang pas:" + tip, rows)

    return total


# ======================================================================
# 10. MAIN
# ======================================================================
def main():
    print(f"== MINDJIWA BOT DESAIN ({CANVAS_W}x{CANVAS_H}) | unsplash={'on' if UNSPLASH_ACCESS_KEY else 'off'} | belajar={'on' if LEARN_ENABLED else 'off'} ==")

    # 1) muat memori + 2) proses feedback (pencetan tombol) sejak run lalu
    mem = load_memory()
    if LEARN_ENABLED:
        try:
            mem = process_feedback(mem)
        except Exception as e:
            print("  ! proses feedback gagal:", e)
        apply_prefs_to_globals(mem)
        print("  " + prefs_summary(mem["preferences"]))

    # 3) AUTO-REVISI: generate ulang konten yg tadi di-reject (sampai di-approve)
    if LEARN_ENABLED and mem.get("revise_queue"):
        for pid in list(mem["revise_queue"].keys()):
            info = mem["revise_queue"].pop(pid)   # pindah dari antrian -> jadi pending lagi
            if info.get("revision", 1) > MAX_REVISI:
                tg_message(f"🛑 '{info.get('title','desain')}' udah direvisi {MAX_REVISI}x tapi belum pas. "
                           f"Mungkin lebih enak kamu kasih contoh manual biar gw niru. Gw stop auto-revisi yg ini dulu.")
                continue
            page = notion_get_page(pid)
            if not page:
                tg_message(f"⚠️ Gagal ambil ulang '{info.get('title','desain')}' dari Notion buat revisi.")
                continue
            wd = tempfile.mkdtemp()
            try:
                tg_message(f"🔄 Merevisi '{info.get('title','desain')}' (revisi ke-{info.get('revision',1)})…")
                process_page(page, wd, mem, revision=info)
            except Exception as e:
                print("  x revisi gagal:", e); traceback.print_exc()
                tg_message(f"⚠️ Revisi '{info.get('title','desain')}' gagal: {type(e).__name__}. Gw coba lagi next.")
                mem["revise_queue"][pid] = info   # balikin ke antrian biar dicoba lagi

    # 4) konten BARU berstatus "Siap Desain"
    pages = notion_find_ready()
    print(f"Ditemukan {len(pages)} konten berstatus '{STATUS_READY}'.")
    if not pages:
        # tidak ada kerjaan -> jangan spam Telegram tiap run, cukup simpan memori
        print("Tidak ada konten baru. (Feedback tetap diproses.)")
        save_memory(mem)
        return

    tg_message(f"✅ Robot Mindjiwa jalan — ada {len(pages)} konten baru.")
    for page in pages:
        page_id = page["id"]
        workdir = tempfile.mkdtemp()
        try:
            process_page(page, workdir, mem)
            notion_set_status(page_id, STATUS_DONE, icon_emoji=DONE_EMOJI)
            print("  v Sukses & status diubah jadi:", STATUS_DONE, "| ikon ->", DONE_EMOJI)
        except Exception as e:
            err = f"{type(e).__name__}: {e}"
            print("  x GAGAL:", err)
            traceback.print_exc()
            try:
                tg_message(f"⚠️ Gagal memproses '{get_title(page['properties'])}'.\n{err}")
            except Exception:
                pass
            notion_set_status(page_id, STATUS_ERROR, note=err)

    save_memory(mem)
    print("== Selesai ==")


if __name__ == "__main__":
    main()
