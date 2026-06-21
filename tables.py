"""Best-seller tablo görseli üretici — Mavi lacivert stil.
GY ve GH sıralaması sütunları opsiyonel olarak eklenir."""
from PIL import Image, ImageDraw, ImageFont
import os

NAVY  = (38, 79, 120)
LIGHT = (222, 233, 245)
WHITE = (255, 255, 255)
BLACK = (20, 20, 20)
GRID  = (150, 150, 150)
SCALE = 2

def _font(size, bold=False):
    for c in (["/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
               else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
               "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold
               else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"]):
        if os.path.exists(c): return ImageFont.truetype(c, int(size * SCALE))
    return ImageFont.load_default()

def _fmt_int(n):
    try: return f"{int(round(n)):,}".replace(",", ".")
    except: return str(n)

def _fmt_pct(x):
    return f"{x*100:.1f}".replace(".", ",") + "%"

def build_bestseller_table(rows, totals, weights, kind="jean", out_path="table.png",
                           gy_ranks=None, gh_ranks=None):
    """
    rows: [(kod, urun_ismi, satis, stok), ...]
    gy_ranks: {kod: rank_int} — None ise sütun gösterilmez
    gh_ranks: {kod: rank_int} — None ise sütun gösterilmez
    """
    show_gy = gy_ranks is not None and any(v is not None for v in gy_ranks.values())
    show_gh = gh_ranks is not None and any(v is not None for v in gh_ranks.values())

    # Sütun genişlikleri
    col_defs = [("1-31 Sıralaması / Ürün Kodları", 240)]
    if show_gy: col_defs.append(("GY Sıra", 70))
    if show_gh: col_defs.append(("GH Sıra", 70))
    col_defs += [("Ürün İsmi", 330), ("Toplam Satış", 115), ("Kalan Stok", 105)]

    headers  = [c[0] for c in col_defs]
    col_w    = [c[1] for c in col_defs]
    row_h    = 42
    pad      = 10
    n_rows   = 1 + len(rows) + 3
    W        = sum(col_w)
    H        = row_h * n_rows

    img = Image.new("RGB", (W * SCALE, H * SCALE), WHITE)
    d   = ImageDraw.Draw(img)

    f_hdr  = _font(15, bold=True)
    f_cell = _font(16, bold=True)
    f_sum  = _font(15, bold=True)
    f_sums = _font(12, bold=True)

    def cell(x, y, w, h, text, fill, fg, align="left", font=None):
        if font is None: font = f_cell
        d.rectangle([x*SCALE, y*SCALE, (x+w)*SCALE, (y+h)*SCALE],
                    fill=fill, outline=GRID, width=1)
        if not text: return
        bbox = d.textbbox((0,0), text, font=font)
        tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
        if align == "left":   tx = x*SCALE + pad*SCALE
        elif align == "right": tx = (x+w)*SCALE - tw - pad*SCALE
        else:                  tx = x*SCALE + (w*SCALE - tw)//2
        ty = y*SCALE + (h*SCALE - th)//2 - bbox[1]
        d.text((tx, ty), text, fill=fg, font=font)

    # başlık
    x = 0; y = 0
    aligns = ["left"] + (["center"] if show_gy else []) + (["center"] if show_gh else []) + ["left","center","center"]
    for i, hdr in enumerate(headers):
        cell(x, y, col_w[i], row_h, hdr, NAVY, WHITE, aligns[i], f_hdr)
        x += col_w[i]
    y += row_h

    # ürün satırları
    codes_list = [r[0] for r in rows]
    for ri, (kod, isim, satis, stok) in enumerate(rows):
        x = 0
        cell(x, y, col_w[0], row_h, str(kod), WHITE, BLACK, "left"); x += col_w[0]
        ci = 1
        if show_gy:
            gy = gy_ranks.get(str(kod).strip())
            cell(x, y, col_w[ci], row_h, str(gy) if gy else "-", WHITE, BLACK, "center"); x += col_w[ci]; ci += 1
        if show_gh:
            gh = gh_ranks.get(str(kod).strip())
            cell(x, y, col_w[ci], row_h, str(gh) if gh else "-", WHITE, BLACK, "center"); x += col_w[ci]; ci += 1
        cell(x, y, col_w[ci], row_h, str(isim), WHITE, BLACK, "left"); x += col_w[ci]; ci += 1
        cell(x, y, col_w[ci], row_h, _fmt_int(satis), WHITE, BLACK, "center"); x += col_w[ci]; ci += 1
        cell(x, y, col_w[ci], row_h, _fmt_int(stok),  WHITE, BLACK, "center")
        y += row_h

    # özet satırlar
    total_label  = "Total Jean Adedi" if kind == "jean" else "Total Adet"
    weight_label = "İlk 10 Total Jean Ağırlığı" if kind == "jean" else "İlk 10 Total Ağırlığı"
    summary = [
        ("İlk 10 Ürünün Toplam Adedi", _fmt_int(totals['top10_satis']), _fmt_int(totals['top10_stok'])),
        (total_label,                   _fmt_int(totals['total_satis']), _fmt_int(totals['total_stok'])),
        (weight_label,                  _fmt_pct(weights['satis']),      _fmt_pct(weights['stok'])),
    ]
    # Özet satır — GY/GH sütunları boş geçilir
    n_extra = (1 if show_gy else 0) + (1 if show_gh else 0)
    for label, v1, v2 in summary:
        x = 0
        cell(x, y, col_w[0], row_h, label, LIGHT, BLACK, "left", f_sums); x += col_w[0]
        for e in range(n_extra):
            cell(x, y, col_w[1+e], row_h, "", LIGHT, BLACK); x += col_w[1+e]
        ci = 1 + n_extra
        cell(x, y, col_w[ci], row_h, "", LIGHT, BLACK); x += col_w[ci]
        ci += 1
        cell(x, y, col_w[ci], row_h, v1, LIGHT, BLACK, "center", f_sum); x += col_w[ci]
        ci += 1
        cell(x, y, col_w[ci], row_h, v2, LIGHT, BLACK, "center", f_sum)
        y += row_h

    img.save(out_path, "PNG")
    return out_path, (W, H)
