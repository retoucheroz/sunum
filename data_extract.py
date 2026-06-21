"""Best-seller Excel'lerden veri çekme modülü.
- Bu dönem ilk 10 (Sayfa1/Sayfa2'den)
- GY sıralaması ('Denim All Tamamı (2)' sayfasından — geçen yıl)
- GH sıralaması (opsiyonel ayrı GH dosyasından — verilmezse boş)
"""
import warnings
import pandas as pd
warnings.filterwarnings("ignore")

# ── Ham sayfa okuyucu ────────────────────────────────────────────────────────
def _read_product_sheet(path, sheet):
    """ProductCode, ProductName, Net Sales Quantity, Stock Quantity, Section içeren sayfayı okur."""
    try:
        df = pd.read_excel(path, sheet_name=sheet, header=None)
    except Exception:
        return None
    # Başlık satırını bul (Section + ProductCode içeren satır)
    header_row = None
    for i, row in df.iterrows():
        vals = [str(v).strip() for v in row if pd.notna(v)]
        if 'Section' in vals and 'ProductCode' in vals:
            header_row = i
            break
    if header_row is None:
        return None
    df.columns = [str(v).strip() for v in df.iloc[header_row]]
    df = df.iloc[header_row+1:].reset_index(drop=True)
    # Sütunları normalize et
    df = df.rename(columns={c: c for c in df.columns})
    for col in ['Net Sales Quantity', 'Stock Quantity']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df

# ── GY sıralaması hesapla ────────────────────────────────────────────────────
def _build_gy_rank(path, sheet, section_vals):
    """GY (Geçen Yıl) sayfasından section bazlı sıralama dict'i döner.
    {product_code: rank}  — rank 1 = en çok satan"""
    df = _read_product_sheet(path, sheet)
    if df is None or 'Section' not in df.columns:
        return {}
    mask = df['Section'].astype(str).str.strip().isin(section_vals)
    sub = df[mask & df['ProductCode'].notna()].copy()
    sub = sub[sub['ProductCode'].astype(str).str.strip() != '']
    sub = sub.sort_values('Net Sales Quantity', ascending=False).reset_index(drop=True)
    return {str(r['ProductCode']).strip(): i+1 for i, r in sub.iterrows()}

# ── GH (opsiyonel ayrı dosya) sıralaması ────────────────────────────────────
def _build_gh_rank(gh_path, section_vals):
    """GH (Geçen Hafta) ayrı dosyasından sıralama. Dosya yoksa {} döner."""
    if not gh_path:
        return {}
    import openpyxl
    try:
        wb = openpyxl.load_workbook(gh_path, read_only=True, data_only=True)
    except Exception:
        return {}
    # Ürün verisi içeren sayfayı bul
    for sn in wb.sheetnames:
        df = _read_product_sheet(gh_path, sn)
        if df is not None and 'Section' in df.columns and 'ProductCode' in df.columns:
            mask = df['Section'].astype(str).str.strip().isin(section_vals)
            sub = df[mask & df['ProductCode'].notna()].copy()
            sub = sub.sort_values('Net Sales Quantity', ascending=False).reset_index(drop=True)
            wb.close()
            return {str(r['ProductCode']).strip(): i+1 for i, r in sub.iterrows()}
    wb.close()
    return {}

# ── Ana fonksiyon ────────────────────────────────────────────────────────────
def extract_bestseller(path, gy_sheet='Denim All Tamamı (2)', gh_path=None):
    """
    İki sayfalı (Sayfa1/Sayfa2) best-seller dosyasını okur.
    Section sütununa göre Erkek/Kadın ayrımı yapar.
    GY sıralamasını gy_sheet'ten, GH'yi opsiyonel gh_path'ten hesaplar.
    
    Döner: {
      'erkek': {rows, totals, weights, fit_names, codes, gy_ranks, gh_ranks},
      'kadin': {...}
    }
    """
    import openpyxl
    xls = pd.ExcelFile(path)
    frames = []
    for sn in xls.sheet_names:
        try:
            df = _read_product_sheet(path, sn)
        except Exception:
            continue
        if df is None: continue
        cols = set(df.columns)
        if {'Section', 'ProductCode', 'Net Sales Quantity', 'Stock Quantity'}.issubset(cols):
            frames.append(df)
    if not frames:
        raise ValueError(f"{path}: beklenen sütunlar bulunamadı")
    allrows = pd.concat(frames, ignore_index=True)

def extract_bestseller(path, gy_sheet='Denim All Tamamı (2)', gh_path=None):
    """
    Sayfa1 (Kadın) ve Sayfa2 (Erkek) sayfalarından doğrudan okur.
    GY sıralamasını gy_sheet'ten hesaplar.
    """
    result = {}
    for sec_key, sheet_name, sec_vals in [
        ('erkek', 'Sayfa2', ['Erkek', 'Men']),
        ('kadin', 'Sayfa1', ['Kadın', 'Kadin', 'Women']),
    ]:
        df = _read_product_sheet(path, sheet_name)
        if df is None:
            raise ValueError(f"{path}: '{sheet_name}' sayfası okunamadı")
        
        # Genel Toplam satırını at
        df = df[~df['Section'].astype(str).str.contains('Toplam|Total|Genel', na=False)]
        df = df[df['ProductCode'].notna() & (df['ProductCode'].astype(str).str.strip() != '')]
        df['Net Sales Quantity'] = pd.to_numeric(df['Net Sales Quantity'], errors='coerce').fillna(0)
        df['Stock Quantity']     = pd.to_numeric(df['Stock Quantity'],     errors='coerce').fillna(0)
        df = df.sort_values('Net Sales Quantity', ascending=False)

        total_satis = df['Net Sales Quantity'].sum()
        total_stok  = df['Stock Quantity'].sum()
        top = df.head(10)
        top10_satis = top['Net Sales Quantity'].sum()
        top10_stok  = top['Stock Quantity'].sum()

        rows_data = [(str(r['ProductCode']).strip(), str(r['ProductName']).strip(),
                      r['Net Sales Quantity'], r['Stock Quantity'])
                     for _, r in top.iterrows()]
        codes = [r[0] for r in rows_data]

        # GY sıralaması
        gy_ranks = _build_gy_rank(path, gy_sheet, sec_vals)
        # GH sıralaması
        gh_ranks = _build_gh_rank(gh_path, sec_vals)

        result[sec_key] = {
            'rows':     rows_data,
            'totals':   {'top10_satis': top10_satis, 'top10_stok': top10_stok,
                         'total_satis': total_satis,  'total_stok': total_stok},
            'weights':  {'satis': top10_satis/total_satis if total_satis else 0,
                         'stok':  top10_stok/total_stok   if total_stok  else 0},
            'fit_names': [_fit_name(r[1]) for r in rows_data],
            'codes':     codes,
            'gy_ranks':  {code: gy_ranks.get(code) for code in codes},
            'gh_ranks':  {code: gh_ranks.get(code) for code in codes},
        }
    return result

def _fit_name(urun):
    import re
    if not isinstance(urun, str): return ''
    m = re.match(r'\s*([A-ZÇĞİÖŞÜ][A-ZÇĞİÖŞÜ\s]+?)\s', urun + ' ')
    if m: return m.group(1).strip().title()
    return urun.split()[0].title() if urun.split() else ''

def jean_narrative(data, date_range=''):
    e, k = data['erkek'], data['kadin']
    def uniq_fits(fits, n=5):
        seen = []
        for f in fits:
            if f and f not in seen: seen.append(f)
            if len(seen) >= n: break
        return seen
    ef = uniq_fits(e['fit_names']); kf = uniq_fits(k['fit_names'])
    def _k(n): return f"{int(round(n/1000))}K"
    dr = date_range if date_range else ""
    erkek = (f"{dr} tarihleri arasında erkek jean'lerinde toplamda {_k(e['totals']['total_satis'])} "
             f"adet satış gerçekleşmiştir. Satış sıralamasında ilk 10'a "
             f"{', '.join(ef[:-1])} ve {ef[-1]} jean fitleri girmiştir. "
             f"İlk 10 ürünün toplam satışı {_k(e['totals']['top10_satis'])} adettir "
             f"ve toplam satışın %{round(e['weights']['satis']*100)}'unu oluşturmaktadır.")
    kadin = (f"Kadın kategorisinde ise toplamda {_k(k['totals']['total_satis'])} adet satış "
             f"gerçekleşmiştir. Satış sıralamasında ilk 10'a "
             f"{', '.join(kf[:-1])} ve {kf[-1]} jean fitleri girmiştir. "
             f"İlk 10 ürünün toplam satışı {_k(k['totals']['top10_satis'])} adettir "
             f"ve toplam satışın %{round(k['weights']['satis']*100)}'ini oluşturmaktadır.")
    return erkek, kadin

def _k(n): return f"{int(round(n/1000))}K"
