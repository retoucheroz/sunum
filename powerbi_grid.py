"""Slayt 15 & 18 için orijinal Power BI görselini birebir taklit eden
kompozit görsel üretici: sol 'By Quantity Sold' lacivert şerit, üstte
ürün kodu + adı, ortada fotoğraf, altta Qty ve Retail Price satırları."""
from PIL import Image, ImageDraw, ImageFont
import os

NAVY = (31, 64, 99)
GRID = (200, 206, 214)
GRAY = (90, 95, 105)
BLACK = (25, 28, 33)
WHITE = (255, 255, 255)
NAVY_TXT = (31, 64, 99)
S = 2  # 2x ölçek

def _f(size, bold=False):
    for c in (("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
               else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),):
        if os.path.exists(c):
            return ImageFont.truetype(c, int(size*S))
    return ImageFont.load_default()

def _fmt(n):
    return f"{int(round(n)):,}".replace(",", ".")

def _trunc(txt, n):
    txt = str(txt)
    return txt if len(txt) <= n else txt[:n-1] + ""

def _draw_band(d, ox, oy, W, band_h, items, photos):
    """Tek bant (10 ürün) çizer. items:[(code,name,qty,price)], photos:[path|None]"""
    side_w = int(W*0.052)
    ncol = 10
    col_w = (W - side_w)/ncol
    hdr_h = int(band_h*0.13)
    qty_h = int(band_h*0.085)
    price_h = int(band_h*0.085)
    photo_h = band_h - hdr_h - qty_h - price_h

    f_code = _f(8); f_name = _f(8.5); f_val = _f(8.5, True); f_lbl = _f(8, True)
    f_side = _f(9, True)

    # başlık: kod + ad
    for i, (code, name, qty, price) in enumerate(items):
        cx = ox + side_w + i*col_w
        d.text((int((cx+6)*S), int((oy+4)*S)), str(code), font=f_code, fill=GRAY)
        d.text((int((cx+6)*S), int((oy+hdr_h*0.5)*S)), _trunc(name, 17), font=f_name, fill=BLACK)
    # başlık alt çizgi
    yln = oy + hdr_h
    d.line([(int((ox+side_w)*S), int(yln*S)), (int((ox+W)*S), int(yln*S))], fill=NAVY, width=2)

    # fotoğraf alanı
    py = yln
    for i, ph in enumerate(photos):
        cx = ox + side_w + i*col_w
        avail_w = col_w*0.82; avail_h = photo_h*0.94
        if ph and os.path.exists(ph):
            try:
                im = Image.open(ph).convert("RGB")
                iw, ih = im.size
                sc = min(avail_w/iw, avail_h/ih)
                nw, nh = max(1, int(iw*sc)), max(1, int(ih*sc))
                im = im.resize((int(nw*S), int(nh*S)))
                px = int((cx + (col_w-nw)/2)*S); ppy = int((py + (photo_h-nh)/2)*S)
                d._image.paste(im, (px, ppy))
                continue
            except Exception:
                pass
        # placeholder
        d.rectangle([int((cx+col_w*0.09)*S), int((py+photo_h*0.03)*S),
                     int((cx+col_w*0.91)*S), int((py+photo_h*0.97)*S)],
                    fill=(238, 240, 243), outline=GRID, width=1)

    # Qty satırı
    qy = yln + photo_h
    d.line([(int(ox*S), int(qy*S)), (int((ox+W)*S), int(qy*S))], fill=GRID, width=1)
    d.rectangle([int(ox*S), int(qy*S), int((ox+side_w)*S), int((qy+qty_h)*S)], fill=NAVY)
    d.text((int((ox+6)*S), int((qy+qty_h*0.22)*S)), "Qty", font=f_lbl, fill=WHITE)
    for i, (code, name, qty, price) in enumerate(items):
        cx = ox + side_w + i*col_w
        t = _fmt(qty) if qty is not None else ""
        bb = d.textbbox((0, 0), t, font=f_val); tw = bb[2]-bb[0]
        d.text((int((cx+col_w-8)*S-tw), int((qy+qty_h*0.22)*S)), t, font=f_val, fill=BLACK)

    # Retail Price satırı
    ry = qy + qty_h
    d.line([(int(ox*S), int(ry*S)), (int((ox+W)*S), int(ry*S))], fill=GRID, width=1)
    d.rectangle([int(ox*S), int(ry*S), int((ox+side_w)*S), int((ry+price_h)*S)], fill=NAVY)
    d.text((int((ox+6)*S), int((ry+price_h*0.22)*S)), "Retail Price", font=_f(6.5, True), fill=WHITE)
    for i, (code, name, qty, price) in enumerate(items):
        cx = ox + side_w + i*col_w
        t = (f"{price:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
             if price is not None else "")
        bb = d.textbbox((0, 0), t, font=f_val); tw = bb[2]-bb[0]
        d.text((int((cx+col_w-8)*S-tw), int((ry+price_h*0.22)*S)), t, font=f_val, fill=BLACK)

    # "By Quantity Sold" şeridi (fotoğraf yüksekliği boyunca)
    d.rectangle([int(ox*S), int(yln*S), int((ox+side_w)*S), int((yln+photo_h)*S)], fill=NAVY)
    for j, w in enumerate(["By", "Quantity", "Sold"]):
        bb = d.textbbox((0, 0), w, font=f_side); tw = bb[2]-bb[0]
        d.text((int((ox+side_w/2)*S-tw/2), int((yln+photo_h/2-18+j*16)*S)), w, font=f_side, fill=WHITE)

def build_powerbi_grid(erkek, kadin, out_path, photos_e=None, photos_k=None):
    """erkek/kadin: [(code,name,qty,price), ...] (10'ar).
    photos_*: [path|None,...] eşleşen sırada. Döner: (path,(W,H))."""
    W, H = 1300, 700
    img = Image.new("RGB", (W*S, H*S), WHITE)
    d = ImageDraw.Draw(img); d._image = img
    band_h = int((H-30)/2)
    _draw_band(d, 10, 10, W-20, band_h, erkek, photos_e or [None]*len(erkek))
    _draw_band(d, 10, 20+band_h, W-20, band_h, kadin, photos_k or [None]*len(kadin))
    img.save(out_path, "PNG")
    return out_path, (W, H)
