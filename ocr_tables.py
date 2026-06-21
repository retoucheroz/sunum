"""LFL tablo EMF görsellerini OCR ile okuyup kategori bazlı değerleri çıkarır."""
import re, zipfile, tempfile, os
import pytesseract
import numpy as np
from PIL import Image, ImageEnhance

# Kategori eşleştirme: OCR çıktısı → standart key
_CAT_MAP = {
    'denim all':       'Denim All',
    'denim':           'Denim All',
    'jackets':         'Ceket',
    'jacket':          'Ceket',
    'ceket':           'Ceket',
    'shirts':          'Gömlek',
    'shirt':           'Gömlek',
    'gomlek':          'Gömlek',
    'non denim bottoms': 'Non Denim Alt',
    'non-denim':       'Non Denim Alt',
    'non denim':       'Non Denim Alt',
    'tees':            'Tees',
    'accessories':     'Aksesuar',
    'accessory':       'Aksesuar',
    'sweatshirts':     'Sweatshirt',
    'sweatshirt':      'Sweatshirt',
}

def _pct(s):
    """'59.9%' → 59.9"""
    try: return float(s.replace('%','').replace(',','.'))
    except: return None

def _ocr_emf(emf_path):
    """EMF → PNG → OCR metin"""
    import subprocess
    png = emf_path.replace('.emf', '.png')
    subprocess.run(['soffice','--headless','--convert-to','png',
                    '--outdir', os.path.dirname(emf_path), emf_path],
                   capture_output=True)
    if not os.path.exists(png):
        return ''
    im = Image.open(png).convert('L')
    im = ImageEnhance.Contrast(im).enhance(2.5)
    im = im.resize((im.width*3, im.height*3), Image.LANCZOS)
    return pytesseract.image_to_string(im, lang='tur+eng', config='--psm 6 --oem 3')

def _parse_lfl_table(txt):
    """OCR metninden {kategori: {grc_sdg, hdf_sdg}} dict'i çıkarır.
    Tablo yapısı: Kategori | SDG | Önceki | Eski ALL | ALL | SDG(hedef) | ...
    Sütun 1=SDG(grc), sütun 4=ALL(grc), sütun 5=SDG(hdf), sütun 8=ALL(hdf)
    Biz SDG sütununu alıyoruz (col 1 ve col 5)."""
    result = {}
    for line in txt.split('\n'):
        line = line.strip()
        if not line: continue
        # Yüzde değerlerini çıkar
        pcts = re.findall(r'-?\d+[\.,]\d+%', line)
        if len(pcts) < 4: continue
        # Kategori adını satırın başından al
        cat_raw = re.split(r'\s*\|?\s*-?\d', line)[0].strip().lower()
        cat_raw = re.sub(r'[^a-zA-ZğĞşŞıİçÇöÖüÜ\s]', '', cat_raw).strip()
        # Eşleştir
        cat = None
        for key, val in _CAT_MAP.items():
            if key in cat_raw:
                cat = val; break
        if not cat: continue
        # SDG = col 0 (Gerçekleşen SDG), Hedef SDG = col 4
        grc = _pct(pcts[0])
        hdf = _pct(pcts[4]) if len(pcts) > 4 else None
        if cat not in result:
            result[cat] = {'grc_sdg': grc, 'hdf_sdg': hdf}
    return result

def ocr_lfl_from_pptx(pptx_path, slide_idx, emf_idx=1):
    """PPTX'teki belirtilen slaytın LFL tablosunu OCR ile okur.
    emf_idx: 0=adet tablosu, 1=LFL tablosu (genelde ikinci EMF).
    Döner: {kategori: {grc_sdg, hdf_sdg}}"""
    with zipfile.ZipFile(pptx_path) as z:
        rels = z.read(f'ppt/slides/_rels/slide{slide_idx}.xml.rels').decode()
        emfs = re.findall(r'Target="\.\./media/([^"]+\.emf)"', rels)
        if not emfs or emf_idx >= len(emfs):
            return {}
        emf_name = emfs[emf_idx]
        tmpdir = tempfile.mkdtemp()
        emf_path = os.path.join(tmpdir, emf_name.replace('/', '_'))
        with open(emf_path, 'wb') as f:
            f.write(z.read(f'ppt/media/{emf_name}'))
    txt = _ocr_emf(emf_path)
    return _parse_lfl_table(txt)
