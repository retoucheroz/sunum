"""Örnek Çalışma Dosyası Excel'inden GY/BY karşılaştırma verilerini okur.
Slayt 7, 9, 10, 11 için tablo ve bullet verileri üretir."""
import warnings
import pandas as pd
warnings.filterwarnings("ignore")

_SHEET_MAP = {
    'erkek_kategori':      'Sayfa3',   # Slayt 7
    'kadin_kategori':      'Sayfa2',   # Slayt 9
    'erkek_sub_aylik':     'Sayfa5',   # Slayt 10 (1 aylık)
    'kadin_sub_aylik':     'Sayfa4',   # Slayt 11 (1 aylık)
    'erkek_sub_std':       'Sayfa6',   # Slayt 10 (STD)
    'kadin_sub_std':       'Sayfa7',   # Slayt 11 (STD)
}

def _fmt_int(n):
    try: return f"{int(round(float(n))):,}".replace(",", ".")
    except: return str(n)

def _fmt_pct(v):
    try:
        f = float(v)
        s = f"{abs(f)*100:.1f}".replace('.', ',')
        return f"+{s}" if f >= 0 else f"-{s}"
    except: return str(v)

def _fmt_pct_disp(v):
    """Ekranda '40,8%' yerine '%40,8' formatı"""
    try:
        f = float(v)
        s = f"{abs(f)*100:.1f}".replace('.', ',')
        return (f"%{s}" if f >= 0 else f"-%{s}")
    except: return str(v)

def read_sheet(calisma_path, sheet_key):
    """Belirtilen sayfayı DataFrame olarak döndürür."""
    sheet_name = _SHEET_MAP.get(sheet_key)
    if not sheet_name:
        return None
    try:
        df = pd.read_excel(calisma_path, sheet_name=sheet_name, header=0)
        # Satış% ve Stok% sayısala çevir
        for col in df.columns:
            if '%' in str(col) or col in ('Satış%','Stok%'):
                df[col] = pd.to_numeric(df[col], errors='coerce')
        for col in ['GY Satış ', 'BY Satış ', 'GY Stok', 'BY Stok',
                    'Net Sales Quantity', 'Stock Quantity']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        return df
    except Exception as e:
        print(f"Hata ({sheet_name}): {e}")
        return None

def get_erkek_bullets(df):
    """Slayt 7 sağ kutu bullet listesi: [(kategori, satis_pct, stok_pct)]"""
    if df is None: return []
    bullets = []
    # Kategori önem sırası
    priority = ['Denim All','Ceket','Gömlek','Penye Üstler','Aksesuar',
                'Non Denim Altlar','Sweatshirt','Triko']
    for cat in priority:
        row = df[df['Category'] == cat]
        if row.empty: continue
        r = row.iloc[0]
        sat = _fmt_pct_disp(r.get('Satış%', r.get('Satış% ', 0)))
        stk = _fmt_pct_disp(r.get('Stok%', 0))
        bullets.append((cat, sat, stk))
    # Kalan kategoriler
    for _, r in df.iterrows():
        cat = str(r.get('Category',''))
        if cat and cat not in [b[0] for b in bullets] and cat != 'Genel Toplam':
            sat = _fmt_pct_disp(r.get('Satış%', r.get('Satış% ', 0)))
            stk = _fmt_pct_disp(r.get('Stok%', 0))
            bullets.append((cat, sat, stk))
    return bullets

def get_kadin_bullets(df):
    """Slayt 9 sağ kutu bullet listesi."""
    return get_erkek_bullets(df)  # Aynı format

def get_sub_bullets(df, section='Erkek'):
    """Slayt 10/11 sağ kutu — SubCategory bazlı."""
    if df is None: return []
    sub_df = df[df['Section'] == section] if 'Section' in df.columns else df
    bullets = []
    for cat in sub_df['Category'].unique() if 'Category' in sub_df.columns else []:
        cat_rows = sub_df[sub_df['Category'] == cat]
        total_gy_sat = cat_rows['GY Satış '].sum() if 'GY Satış ' in cat_rows else 0
        total_by_sat = cat_rows['BY Satış '].sum() if 'BY Satış ' in cat_rows else 0
        total_gy_stk = cat_rows['GY Stok'].sum()   if 'GY Stok' in cat_rows else 0
        total_by_stk = cat_rows['BY Stok'].sum()   if 'BY Stok' in cat_rows else 0
        sat_pct = (total_by_sat - total_gy_sat) / abs(total_gy_sat) if total_gy_sat else 0
        stk_pct = (total_by_stk - total_gy_stk) / abs(total_gy_stk) if total_gy_stk else 0
        bullets.append((cat, _fmt_pct_disp(sat_pct), _fmt_pct_disp(stk_pct)))
    return bullets

def build_category_table_image(df, section_filter, out_path, title_line=''):
    """
    DataFrame'den Mavi stilinde PNG tablo üretir.
    section_filter: 'Erkek' veya 'Kadın'
    """
    from PIL import Image, ImageDraw, ImageFont
    import os

    if df is None or df.empty:
        return None

    # Veriyi filtrele
    if 'Section' in df.columns:
        sub = df[df['Section'].astype(str).str.contains(section_filter, case=False, na=False)].copy()
    else:
        sub = df.copy()
    # Genel Toplam satırı
    total = df[df.get('Category','').astype(str).str.contains('Toplam|Total', case=False, na=False)] if 'Category' in df.columns else pd.DataFrame()

    # Sütunlar
    cols = ['Section','Category','GY Satış ','BY Satış ','GY Stok','BY Stok','Satış%','Stok%']
    cols = [c for c in cols if c in sub.columns]
    display_cols = {
        'Section':'Section','Category':'Category',
        'GY Satış ':'GY Satış','BY Satış ':'BY Satış',
        'GY Stok':'GY Stok','BY Stok':'BY Stok',
        'Satış%':'Satış%','Stok%':'Stok%'
    }

    SCALE = 2
    NAVY  = (31, 64, 99)
    WHITE = (255,255,255)
    BLACK = (20,20,20)
    LIGHT = (222,233,245)
    GREEN = (0,128,0)
    RED   = (192,0,0)

    col_w = [80,120,90,90,90,90,70,70][:len(cols)]
    row_h = 28
    W = sum(col_w)
    rows_data = [sub[c].tolist() for c in cols]
    n_rows = len(sub) + 2  # başlık + veri + toplam

    img = Image.new('RGB', (W*SCALE, row_h*SCALE*n_rows), WHITE)
    d   = ImageDraw.Draw(img)

    def _font(sz, bold=False):
        for p in ['/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf' if bold
                  else '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf']:
            import os
            if os.path.exists(p):
                from PIL import ImageFont
                return ImageFont.truetype(p, sz*SCALE)
        from PIL import ImageFont
        return ImageFont.load_default()

    f_h = _font(9, True); f_b = _font(9, True); f_r = _font(9)

    def cell(x, y, w, h, txt, fill, fg, align='center', font=None):
        if font is None: font = f_r
        d.rectangle([x*SCALE, y*SCALE, (x+w)*SCALE, (y+h)*SCALE], fill=fill, outline=(180,180,180), width=1)
        if not txt: return
        bb = d.textbbox((0,0), str(txt), font=font)
        tw, th = bb[2]-bb[0], bb[3]-bb[1]
        if align == 'right': tx = (x+w)*SCALE - tw - 4
        elif align == 'left': tx = x*SCALE + 4
        else: tx = x*SCALE + (w*SCALE-tw)//2
        ty = y*SCALE + (h*SCALE-th)//2 - bb[1]
        d.text((tx,ty), str(txt), fill=fg, font=font)

    # Başlık satırı
    x = 0
    headers = [display_cols.get(c,c) for c in cols]
    for i,(h,w) in enumerate(zip(headers,col_w)):
        cell(x, 0, w, row_h, h, NAVY, WHITE, font=f_h)
        x += w

    # Veri satırları
    for ri, (_, row) in enumerate(sub.iterrows()):
        y = (ri+1)*row_h
        x = 0
        fill = LIGHT if ri%2==0 else WHITE
        for ci, c in enumerate(cols):
            v = row[c]
            w = col_w[ci]
            if c in ('Satış%','Stok%'):
                try:
                    fv = float(v)
                    txt = _fmt_pct(v)
                    fg = GREEN if fv >= 0 else RED
                except: txt = str(v); fg = BLACK
            elif c in ('GY Satış ','BY Satış ','GY Stok','BY Stok'):
                txt = _fmt_int(v) if v else ''
                fg = BLACK
            else:
                txt = str(v) if v else ''
                fg = BLACK
            cell(x, y, w, row_h, txt, fill, fg,
                 align='right' if c in ('GY Satış ','BY Satış ','GY Stok','BY Stok','Satış%','Stok%') else 'left',
                 font=f_b if c=='Category' else f_r)
            x += w

    # Toplam satırı
    if not total.empty:
        y = (len(sub)+1)*row_h
        x = 0
        for ci, c in enumerate(cols):
            v = total.iloc[0][c] if c in total.columns else ''
            w = col_w[ci]
            if c in ('Satış%','Stok%'):
                try:
                    fv = float(v); txt = _fmt_pct(v)
                    fg = GREEN if fv >= 0 else RED
                except: txt=str(v); fg=WHITE
                cell(x,y,w,row_h,txt,NAVY,fg,align='right',font=f_b)
            else:
                cell(x,y,w,row_h,_fmt_int(v) if v else '',NAVY,WHITE,align='right' if ci>1 else 'left',font=f_b)
            x += w

    img.save(out_path)
    return out_path
