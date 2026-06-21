"""Mavi Haftalık Özet Sunum Otomasyonu — Web Arayüzü.
Çalıştırmak için:  python3 -m streamlit run app.py
"""
import streamlit as st
import tempfile, os, traceback
from builder import build

st.set_page_config(page_title="Mavi Özet Sunum", page_icon="📊", layout="centered")

st.title("📊 Mavi Haftalık Özet Sunum")
st.caption("5 dosyayı yükle, tarih aralığını gir, hazır sunumu indir.")

# ── 1) Dosyalar ──────────────────────────────────────────────────────────────
st.subheader("1) Dosyalar")
c1, c2 = st.columns(2)
with c1:
    f_template = st.file_uploader("Özet şablon (.pptx)", type="pptx", key="tpl",
                                  help="Önceki haftanın özet sunumu şablon olarak kullanılır")
    f_report   = st.file_uploader("Satıştan gelen rapor (.pptx)", type="pptx", key="rep")
    f_planlama = st.file_uploader("Planlamadan gelen (.xlsx)", type="xlsx", key="plan")
with c2:
    f_fitler   = st.file_uploader("Öne çıkan fitler (.xlsx)", type="xlsx", key="fit")
    f_nondenim = st.file_uploader("Öne çıkan non-denim altlar (.xlsx)", type="xlsx", key="non")
    f_calisma  = st.file_uploader("📊 Örnek çalışma dosyası (.xlsx)", type="xlsx", key="cal",
                                   help="Slayt 7, 9, 10, 11 için GY/BY karşılaştırma tabloları")

# ── 2) Dönem Bilgisi ─────────────────────────────────────────────────────────
st.subheader("2) Dönem Bilgisi")
c3, c4 = st.columns(2)
with c3:
    date_range = st.text_input("Bu haftanın tarih aralığı", value="1-17 Haziran")
    month_label = st.text_input("Bu haftanın ay etiketi", value="Haziran'26")
with c4:
    old_date_range = st.text_input("Şablondaki ESKİ tarih aralığı", value="1-31 Mayıs",
                                    help="Şablonda yazıyı bu ile değiştirecek")
    old_month_label = st.text_input("Şablondaki ESKİ ay etiketi", value="Mayıs'26")
    download_products = st.checkbox(
        "Ürün görsellerini mavi.com'dan indir (slayt 15 & 18)", value=True)

with st.expander("📅 GH Sıralaması — Geçen Hafta Excel'leri (opsiyonel)"):
    st.caption("Bu haftanın ilk 10 ürününün geçen haftaki sıralamasını tabloya ekler. "
               "Verilmezse GH sütunu görünmez.")
    gh1, gh2 = st.columns(2)
    with gh1:
        f_gh_fitler   = st.file_uploader("GH — Öne çıkan fitler (.xlsx)", type="xlsx", key="gh_fit")
    with gh2:
        f_gh_nondenim = st.file_uploader("GH — Non-denim altlar (.xlsx)", type="xlsx", key="gh_non")
    st.caption("Şablondaki tarih yazılarını yeni döneme çevirmek için. Boş bırakırsan dokunulmaz.")
    c5, c6 = st.columns(2)
    with c5:
        rep_old1 = st.text_input("Eski yazı 1", value="")
        rep_old2 = st.text_input("Eski yazı 2", value="")
    with c6:
        rep_new1 = st.text_input("Yeni yazı 1", value="")
        rep_new2 = st.text_input("Yeni yazı 2", value="")

# ── 3) Oluştur ───────────────────────────────────────────────────────────────
st.subheader("3) Oluştur")
if st.button("🚀 Sunumu Oluştur", type="primary"):
    missing = [n for n, f in [("şablon", f_template), ("rapor", f_report),
               ("fitler", f_fitler), ("non-denim", f_nondenim),
               ("planlama", f_planlama)] if f is None]
    if missing:
        st.error("Eksik dosya: " + ", ".join(missing))
    else:
        tmp = tempfile.mkdtemp()
        def _save(uploaded, name):
            p = os.path.join(tmp, name)
            with open(p, "wb") as fh: fh.write(uploaded.getbuffer())
            return p
        try:
            tpl  = _save(f_template, "tpl.pptx")
            rep  = _save(f_report,   "rep.pptx")
            fit  = _save(f_fitler,   "fit.xlsx")
            non  = _save(f_nondenim, "non.xlsx")
            plan = _save(f_planlama, "plan.xlsx")

            replacements = {}
            if rep_old1 and rep_new1: replacements[rep_old1] = rep_new1
            if rep_old2 and rep_new2: replacements[rep_old2] = rep_new2

            cfg = {"date_range": date_range, "month_label": month_label,
                   "old_date_range": old_date_range, "old_ay_etiket": old_month_label,
                   "replacements": replacements}
            out = os.path.join(tmp, f"Ozet_Sunum_{month_label.replace(chr(39),'').replace(' ','_')}.pptx")

            gh_fit_path = None
            gh_non_path = None
            if f_gh_fitler:
                gh_fit_path = _save(f_gh_fitler, "gh_fit.xlsx")
            if f_gh_nondenim:
                gh_non_path = _save(f_gh_nondenim, "gh_non.xlsx")

            calisma_path = None
            if f_calisma:
                calisma_path = _save(f_calisma, "calisma.xlsx")

            with st.spinner("Sunum oluşturuluyor..."):
                log = build(tpl, rep, fit, non, plan, cfg, out,
                            download_products=download_products,
                            gh_fitler_path=gh_fit_path,
                            gh_nondenim_path=gh_non_path,
                            calisma_path=calisma_path)

            st.success("✅ Hazır!")
            with st.expander("İşlem günlüğü", expanded=True):
                st.code("\n".join(log))
            with open(out, "rb") as fh:
                st.download_button(
                    "⬇️ Sunumu İndir (.pptx)", fh.read(),
                    file_name=os.path.basename(out),
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        except Exception as e:
            st.error(f"Hata: {e}")
            st.code(traceback.format_exc())

st.divider()
st.caption("Tüm değerler Excel ve rapor dosyalarından otomatik çekilir — manuel giriş gerekmez.")
