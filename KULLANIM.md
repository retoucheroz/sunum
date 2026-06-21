# Mavi Haftalık Özet Sunum Otomasyonu

Excel ve kaynak dosyalarını verince, mevcut özet sunumu **şablon olarak kullanıp** içeriği
güncelleyen ve hazır `.pptx` üreten araç. Tasarım/marka düzeni korunur.

## Kurulum (tek seferlik)

1. [Python 3.10+](https://www.python.org/downloads/) kurulu olmalı.
2. Bu klasörde terminal açıp:

```bash
pip install -r requirements.txt
```

## Çalıştırma

```bash
streamlit run app.py
```

Tarayıcıda arayüz açılır. Sırayla:

1. **5 dosyayı yükle:** özet şablon, satıştan gelen rapor, öne çıkan fitler, öne çıkan
   non-denim altlar, planlamadan gelen.
2. **Dönem bilgisini** gir (tarih aralığı, ay etiketi).
3. **"Sunumu Oluştur"** → işlem günlüğünü gör → **"Sunumu İndir"**.

## Araç şu an ne yapıyor

| Slayt | İşlem | Durum |
|------|-------|-------|
| Tümü | "Abi..." çalışma notlarını temizler | ✅ Otomatik |
| 3 | Satış raporundaki Hedef + KPI tablolarını resim olarak aktarır | ✅ Otomatik |
| 14 | Erkek/Kadın best-seller **jean** tablolarını Excel'den kurar + anlatı metni | ✅ Otomatik |
| 17 | Erkek/Kadın best-seller **non-denim** tablolarını kurar + anlatı metni | ✅ Otomatik |
| 15, 18 | İlk 10 ürünün fotoğrafını mavi.com'dan indirip dizer | ✅ (kurumsal ağda) |
| 2, 4 | Metrik/özet kutuları | ⏳ Form değerleri (sonraki sürümde planlamadan otomatik) |
| 5, 6, 8 | LFL Hedef/Gerçekleşen | ⏳ Planlama eşlemesi netleşince otomatik |

## Önemli notlar

- **Ürün görselleri (slayt 15/18):** `https://sky-static.mavi.com/<kod>_image_1.jpg`
  adresinden indirilir (üründeki baştaki "M" atılır). İnternet erişimi olan bir makinede
  çalışır. Erişim yoksa o slaytlar atlanır ve günlükte belirtilir.
- **Best-seller mantığı:** Erkek/Kadın ayrımı sayfa sırasından değil **`Section`**
  sütunundan yapılır. Ürünler `Net Sales Quantity`'e göre büyükten küçüğe sıralanır,
  ilk 10 alınır, alt toplamlar ve ağırlık % hesaplanır.
- **Başka ay için:** Arayüzdeki "Tekrar kullanım" bölümünden şablondaki tarih/ay
  yazılarını yeni döneme bul-değiştir yapabilirsin.

## Dosya yapısı

```
app.py            → web arayüzü (Streamlit)
builder.py        → ana motor (sunum oluşturma)
data_extract.py   → Excel okuma + best-seller hesaplama
tables.py         → best-seller tablo görseli üretici (Mavi lacivert stili)
requirements.txt  → bağımlılıklar
```

## Sonraki adımlar (planlanan)

- Slayt 5/6/8 Hedef/Gerçekleşen değerlerinin planlama dosyasından otomatik çekilmesi
  (kaynak sütun bir kez netleştirilince).
- GY/GH (geçen hafta & geçen ay sıralaması) sütunları — önceki dönem export'ları verilince.
