# Mavi Sunum Aracı — Deploy Rehberi

## Streamlit Community Cloud (önerilen, ücretsiz)

1. Bu klasördeki tüm dosyaları bir GitHub reposuna push et (repo public veya private olabilir).
   Gerekli dosyalar:
   - `app.py`, `builder.py`, `data_sources.py`, `data_extract.py`,
     `slide_updaters.py`, `slide_utils.py`, `tables.py`, `powerbi_grid.py`,
     `ocr_tables.py`, `calisma_dosyasi.py`
   - `requirements.txt` (Python paketleri)
   - `packages.txt` (sistem paketleri: tesseract, libreoffice)

2. https://share.streamlit.io adresine git, GitHub hesabınla giriş yap.

3. "New app" → reponu seç → main file olarak `app.py` seç → Deploy.

4. İlk build 3-5 dakika sürebilir çünkü `packages.txt`'teki LibreOffice
   kurulumu büyük bir paket. Sabırlı ol, sadece ilk seferde.

5. Build bitince uygulamanın linki hazır olur (örn. `xxx.streamlit.app`).
   Bu linki her salı kullanırsın, kurulum tekrar gerekmez.

### packages.txt içeriği (zaten eklendi)
```
tesseract-ocr
tesseract-ocr-tur
libreoffice
```

## Neden Vercel değil?

Vercel serverless mimaride çalışır — her istek kısa ömürlü, izole bir
fonksiyon olarak işlenir. Bu araç ise:
- Streamlit'in kalıcı sunucu/websocket modeline ihtiyaç duyuyor
- `soffice` (LibreOffice) ve `tesseract` gibi sistem binary'lerine
  ihtiyaç duyuyor — Vercel'de bunları kurmak mümkün değil

Streamlit Community Cloud, Railway, Render veya Hugging Face Spaces gibi
platformlar bu ihtiyaçları (apt paketleri + kalıcı süreç) destekler.

## Alternatif: Railway / Render (Docker ile)

Eğer Streamlit Cloud'da limitlere takılırsan (örn. büyük dosya işleme,
zaman aşımı), bir `Dockerfile` ile Railway veya Render'a da deploy
edebilirsin. İstersen bu Dockerfile'ı da hazırlarım.
