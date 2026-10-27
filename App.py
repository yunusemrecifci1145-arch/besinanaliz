import streamlit as st
import pandas as pd
import io

# 1. Sayfa Ayarları
st.set_page_config(page_title="Klinik Besin Tüketim Analizi", layout="wide")

# 2. Veri Okuma ve Önbellekleme (Caching)
@st.cache_data
def load_data():
    try:
        df = pd.read_excel('besinler.xlsx')
        # İlk sütunu isim, diğerlerini sayısal değer olarak ayarla
        isim_sutunu = df.columns[0]
        df[isim_sutunu] = df[isim_sutunu].astype(str)
        
        for col in df.columns[1:]:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        return df, isim_sutunu
    except FileNotFoundError:
        st.error("Hata: 'besinler.xlsx' dosyası bulunamadı. Lütfen app.py ile aynı dizine ekleyin.")
        return pd.DataFrame(), ""

df_besinler, isim_sutunu = load_data()

# 3. Session State (Oturum Değişkenleri) Başlatma
MEALS = ["Kahvaltı", "Kuşluk (Ara 1)", "Öğle Yemeği", "İkindi (Ara 2)", "Akşam Yemeği", "Gece (Ara 3)"]

if 'tuketim_listesi' not in st.session_state:
    st.session_state['tuketim_listesi'] = {meal: [] for meal in MEALS}

def sifirla():
    st.session_state['tuketim_listesi'] = {meal: [] for meal in MEALS}

# Uygulama Başlığı
st.title("Hastane Besin Tüketim Kaydı ve Analiz Sistemi")

# 4. Hasta Bilgi Alanı (Üst Kısım)
st.header("📋 Hasta Bilgileri ve Hedefler")
col1, col2, col3 = st.columns(3)
with col1:
    hasta_kodu = st.text_input("Hasta Kodu / No", placeholder="Örn: Göğüs-102")
with col2:
    hedef_kcal = st.number_input("Günlük Enerji Hedefi (kcal)", min_value=0.0, step=50.0)
with col3:
    hedef_protein = st.number_input("Günlük Protein Hedefi (g)", min_value=0.0, step=5.0)

# 5. Öğün Blokları ve Besin Girişi (Merkezi Kısım)
st.header("🍽️ Öğün Girişleri")

if not df_besinler.empty:
    besin_listesi = df_besinler[isim_sutunu].tolist()

    for meal in MEALS:
        with st.expander(f"🔽 {meal}", expanded=False):
            c1, c2, c3, c4 = st.columns([4, 2, 2, 2])
            
            with c1:
                secilen_besin = st.selectbox("Besin / Enteral Ürün Seçiniz", besin_listesi, key=f"secim_{meal}")
            with c2:
                miktar = st.number_input("Miktar (g/ml)", min_value=0, value=100, step=10, key=f"mik_{meal}")
            with c3:
                yuzde = st.select_slider("Tüketim Yüzdesi", options=[25, 50, 75, 100], value=100, key=f"yuzde_{meal}")
            with c4:
                st.markdown("<br>", unsafe_allow_html=True) # Butonu hizalamak için boşluk
                if st.button("Öğüne Ekle", key=f"ekle_{meal}", use_container_width=True):
                    # Besin verilerini çek
                    besin_data = df_besinler[df_besinler[isim_sutunu] == secilen_besin].iloc[0]
                    hesaplanan_degerler = {"Besin": secilen_besin, "Girilen Miktar": miktar, "Tüketim %": yuzde}
                    
                    # Makro ve mikro besinleri (Miktar/100 * Yüzde/100) formülüyle hesapla
                    for col in df_besinler.columns[1:]:
                        deger = besin_data[col]
                        hesaplanan_degerler[col] = (deger / 100) * miktar * (yuzde / 100)
                        
                    st.session_state['tuketim_listesi'][meal].append(hesaplanan_degerler)
                    st.success(f"{secilen_besin} eklendi!")

            # Eklenenleri Gösterme ve Silme Butonları
            if st.session_state['tuketim_listesi'][meal]:
                st.markdown("**Eklenenler:**")
                for i, item in enumerate(st.session_state['tuketim_listesi'][meal]):
                    cc1, cc2, cc3 = st.columns([6, 2, 2])
                    cc1.write(f"- {item['Besin']} ({item['Girilen Miktar']}g/ml, %{item['Tüketim %']})")
                    if cc3.button("Sil ❌", key=f"sil_{meal}_{i}", help="Besini sil"):
                        st.session_state['tuketim_listesi'][meal].pop(i)
                        st.rerun()

# 6. Hesaplama Motoru (Arka Plan)
# Tüm öğünlerdeki verileri tek bir tabloda (DataFrame) birleştiriyoruz
tum_tuketim = []
for meal in MEALS:
    for item in st.session_state['tuketim_listesi'][meal]:
        item_copy = item.copy()
        item_copy['Öğün'] = meal
        tum_tuketim.append(item_copy)

df_ozet = pd.DataFrame(tum_tuketim)

# Toplamları bulmak için sayısal sütunları tespit et
if not df_ozet.empty:
    sayisal_sutunlar = [col for col in df_ozet.columns if col not in ['Öğün', 'Besin', 'Girilen Miktar', 'Tüketim %']]
    toplamlar = df_ozet[sayisal_sutunlar].sum()
else:
    toplamlar = pd.Series(dtype=float)

# Enerji, CHO, Protein, Yağ, Lif isimlerini kendi dosyanıza göre esnek yakalamak için:
# Örn: sütun isimlerinde 'Enerji', 'Karbonhidrat' vs geçiyorsa onları alalım.
enerji_col = next((c for c in toplamlar.index if 'enerji' in c.lower() or 'kcal' in c.lower()), None)
cho_col = next((c for c in toplamlar.index if 'karbonhidrat' in c.lower() or 'cho' in c.lower()), None)
pro_col = next((c for c in toplamlar.index if 'protein' in c.lower()), None)
yag_col = next((c for c in toplamlar.index if 'yağ' in c.lower() or 'fat' in c.lower()), None)
lif_col = next((c for c in toplamlar.index if 'lif' in c.lower() or 'fiber' in c.lower()), None)

# 7. Canlı Sonuç Paneli ve Raporlama (Alt Kısım)
st.divider()
st.header("📊 Toplam Tüketim Analizi")

if not df_ozet.empty:
    # --- Makro Dağılım Kartları ---
    top_enerji = toplamlar[enerji_col] if enerji_col else 0
    top_cho = toplamlar[cho_col] if cho_col else 0
    top_pro = toplamlar[pro_col] if pro_col else 0
    top_yag = toplamlar[yag_col] if yag_col else 0
    top_lif = toplamlar[lif_col] if lif_col else 0

    # Enerji % Hesapları (CHO: 4, PRO: 4, YAĞ: 9)
    cho_yuzde = (top_cho * 4 / top_enerji * 100) if top_enerji > 0 else 0
    pro_yuzde = (top_pro * 4 / top_enerji * 100) if top_enerji > 0 else 0
    yag_yuzde = (top_yag * 9 / top_enerji * 100) if top_enerji > 0 else 0

    mc1, mc2, mc3, mc4, mc5 = st.columns(5)
    mc1.metric("Toplam Enerji", f"{top_enerji:.1f} kcal")
    mc2.metric("Karbonhidrat", f"{top_cho:.1f} g", f"% {cho_yuzde:.1f} enerjiden", delta_color="off")
    mc3.metric("Protein", f"{top_pro:.1f} g", f"% {pro_yuzde:.1f} enerjiden", delta_color="off")
    mc4.metric("Yağ", f"{top_yag:.1f} g", f"% {yag_yuzde:.1f} enerjiden", delta_color="off")
    mc5.metric("Lif", f"{top_lif:.1f} g")

    # --- Hedef Karşılama Göstergesi ---
    if hedef_kcal > 0 or hedef_protein > 0:
        st.subheader("🎯 Hedef Karşılama Oranları")
        hc1, hc2 = st.columns(2)
        if hedef_kcal > 0:
            oran_enerji = (top_enerji / hedef_kcal) * 100
            hc1.progress(min(oran_enerji / 100, 1.0))
            hc1.write(f"**Enerji:** Hedefin %{oran_enerji:.1f}'i karşılandı.")
        if hedef_protein > 0:
            oran_pro = (top_pro / hedef_protein) * 100
            hc2.progress(min(oran_pro / 100, 1.0))
            hc2.write(f"**Protein:** Hedefin %{oran_pro:.1f}'i karşılandı.")

    # --- Tüm Mikro Besin Ögeleri Tablosu ---
    st.subheader("🔬 Günlük Alınan Tüm Besin Ögeleri")
    mikro_df = pd.DataFrame({
        "Besin Ögesi": toplamlar.index,
        "Toplam Alınan Miktar": toplamlar.values.round(2)
    })
    st.dataframe(mikro_df.T, use_container_width=True) # Yatay gösterim için Transpoze edildi

    # --- Dışa Aktarma ve Sıfırlama ---
    st.divider()
    rapor_col1, rapor_col2 = st.columns([2, 1])
    
    with rapor_col1:
        # SPSS / Tez formatı için tek satırlık hasta verisi hazırlama
        spesifik_hasta_verisi = {
            "Hasta Kodu": hasta_kodu,
            "Hedef Kcal": hedef_kcal,
            "Hedef Protein": hedef_protein,
            **toplamlar.round(3).to_dict()
        }
        df_export = pd.DataFrame([spesifik_hasta_verisi])
        
        csv_data = df_export.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📄 Raporu CSV Olarak İndir (Tez / SPSS Veri Setine Uygun)",
            data=csv_data,
            file_name=f"{hasta_kodu if hasta_kodu else 'hasta'}_analiz.csv",
            mime='text/csv',
            type="primary"
        )
else:
    st.info("Henüz herhangi bir öğüne besin eklenmedi. Tez verisi veya SPSS çıktısı oluşturmak için yukarıdaki sekmelerden besin girmeye başlayın.")

with st.columns([1, 2, 1])[2]: # Sağ tarafa hizalamak için
    if st.button("🔄 Tümünü Sıfırla / Yeni Hasta", help="Mevcut hastanın kayıtlarını temizler."):
        sifirla()
        st.rerun()
