import streamlit as st
import pandas as pd
import io

# 1. Sayfa Ayarları
st.set_page_config(page_title="Klinik Besin Tüketim Analizi", layout="wide")

# 2. Veri Yükleme ve Birleştirme
@st.cache_data
def load_data():
    try:
        df = pd.read_excel('besinler.xlsx')
    except Exception as e:
        st.error(f"Excel yüklenirken hata oluştu: {e}")
        return None, None

    # Dağınık olan isim sütunlarını tek bir 'Besin Adı' altında topla
    isim_serisi = df['Besin Adı'].fillna(df.get('Yemek Adı', pd.Series(dtype=object))).fillna(df.get('Ürün Adı', pd.Series(dtype=object)))
    df['Besin Adı'] = isim_serisi

    # Dağınık olan enerji sütunlarını topla
    kalori_serisi = df.get('Kalori (kcal)', pd.Series(dtype=float))
    if 'Enerji (kcal)' in df.columns:
        kalori_serisi = kalori_serisi.fillna(df['Enerji (kcal)'])
    if 'Enerji (kcal/100ml)' in df.columns:
        kalori_serisi = kalori_serisi.fillna(df['Enerji (kcal/100ml)'])
    df['Kalori (kcal)'] = kalori_serisi

    # Protein sütunlarını birleştir
    if 'Protein (g/100ml)' in df.columns:
        df['Protein (g)'] = df['Protein (g)'].fillna(df['Protein (g/100ml)'])

    # İsmi olmayan boş satırları at, tekrarlayanları temizle
    df = df.dropna(subset=['Besin Adı']).copy()
    df['Besin Adı'] = df['Besin Adı'].astype(str).str.strip()
    df = df.drop_duplicates(subset=['Besin Adı'])

    # Sayısal olmayan değerleri sıfırla
    for col in df.columns:
        if col not in ['Besin Adı', 'Kategori', 'Kaynak Dosya']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    return df, 'Besin Adı'

df, isim_sutunu = load_data()

if df is None:
    st.stop()

# 3. Session State (Hafıza) Kurulumu
ogunler = ["Kahvaltı", "Kuşluk (Ara 1)", "Öğle Yemeği", "İkindi (Ara 2)", "Akşam Yemeği", "Gece (Ara 3)"]
if "tuketim" not in st.session_state:
    st.session_state.tuketim = {ogun: [] for ogun in ogunler}

# 4. Üst Başlık ve Hasta Bilgileri
st.title("🥗 Klinik Besin Tüketim Kaydı & Analiz")
col_h1, col_h2, col_h3 = st.columns(3)
with col_h1:
    hasta_no = st.text_input("Hasta No / Kodu", placeholder="Örn: G-101")
with col_h2:
    hedef_kcal = st.number_input("Günlük Enerji Hedefi (kcal)", min_value=0.0, value=0.0, step=50.0)
with col_h3:
    hedef_protein = st.number_input("Günlük Protein Hedefi (g)", min_value=0.0, value=0.0, step=5.0)

st.markdown("---")

# 5. Öğün Giriş Blokları
besin_listesi = ["Seçiniz..."] + sorted(df[isim_sutunu].tolist())

for ogun in ogunler:
    with st.expander(f"🍽️ {ogun} ({len(st.session_state.tuketim[ogun])} besin eklendi)", expanded=False):
        c1, c2, c3, c4 = st.columns([3, 1, 1, 1])
        with c1:
            secilen_besin = st.selectbox(f"Besin Ara / Seç ({ogun})", options=besin_listesi, key=f"secim_{ogun}")
        with c2:
            miktar = st.number_input(f"Miktar (g / ml)", min_value=1.0, value=100.0, step=10.0, key=f"miktar_{ogun}")
        with c3:
            yuzde = st.select_slider(f"Tüketim (%)", options=[25, 50, 75, 100], value=100, key=f"yuzde_{ogun}")
        with c4:
            st.write("")
            st.write("")
            if st.button("➕ Ekle", key=f"btn_{ogun}"):
                if secilen_besin != "Seçiniz...":
                    besin_verisi = df[df[isim_sutunu] == secilen_besin].iloc[0].to_dict()
                    carpan = (miktar / 100.0) * (yuzde / 100.0)
                    
                    hesaplanan = {"Besin": secilen_besin, "Miktar (g/ml)": miktar, "Tüketim (%)": yuzde}
                    for k, v in besin_verisi.items():
                        if isinstance(v, (int, float)):
                            hesaplanan[k] = v * carpan
                            
                    st.session_state.tuketim[ogun].append(hesaplanan)
                    st.rerun()

        # Eklenen besinleri listele
        if st.session_state.tuketim[ogun]:
            for idx, item in enumerate(st.session_state.tuketim[ogun]):
                k1, k2, k3, k4 = st.columns([4, 2, 2, 1])
                k1.write(f"• **{item['Besin']}**")
                k2.write(f"{item['Miktar (g/ml)']} g/ml (%{item['Tüketim (%)']})")
                k3.write(f"{item.get('Kalori (kcal)', 0):.1f} kcal | {item.get('Protein (g)', 0):.1f} g prot")
                if k4.button("❌", key=f"del_{ogun}_{idx}"):
                    st.session_state.tuketim[ogun].pop(idx)
                    st.rerun()

st.markdown("---")

# 6. Canlı Sonuç Paneli
st.subheader("📊 Günlük Toplam Besin Değerleri")

tum_besinler = []
for ogun in ogunler:
    tum_besinler.extend(st.session_state.tuketim[ogun])

if tum_besinler:
    df_toplam = pd.DataFrame(tum_besinler)
    
    toplam_kcal = df_toplam.get('Kalori (kcal)', pd.Series([0])).sum()
    toplam_cho = df_toplam.get('Karbonhidrat (g)', pd.Series([0])).sum()
    toplam_prot = df_toplam.get('Protein (g)', pd.Series([0])).sum()
    toplam_yag = df_toplam.get('Yağ (g)', pd.Series([0])).sum()
    toplam_lif = df_toplam.get('Lif (g)', pd.Series([0])).sum()

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Toplam Enerji", f"{toplam_kcal:.1f} kcal")
    m2.metric("Karbonhidrat", f"{toplam_cho:.1f} g")
    m3.metric("Protein", f"{toplam_prot:.1f} g")
    m4.metric("Yağ", f"{toplam_yag:.1f} g")
    m5.metric("Lif", f"{toplam_lif:.1f} g")

    # Hedef Karşılama
    if hedef_kcal > 0 or hedef_protein > 0:
        st.markdown("##### 🎯 Hedef Karşılama Durumu")
        c_h1, c_h2 = st.columns(2)
        if hedef_kcal > 0:
            oran_kcal = (toplam_kcal / hedef_kcal) * 100
            c_h1.info(f"Enerji Karşılama: **%{oran_kcal:.1f}** ({toplam_kcal:.0f} / {hedef_kcal:.0f} kcal)")
        if hedef_protein > 0:
            oran_prot = (toplam_protein / hedef_protein) * 100
            c_h2.info(f"Protein Karşılama: **%{oran_prot:.1f}** ({toplam_prot:.1f} / {hedef_protein:.1f} g)")

    # Mikro Besin Tablosu
    st.markdown("##### 🔬 Mikro ve Diğer Besin Ögeleri Dökümü")
    sayisal_sutunlar = [c for c in df_toplam.columns if c not in ['Besin', 'Miktar (g/ml)', 'Tüketim (%)', 'Besin Adı', 'Kategori', 'Kaynak Dosya']]
    mikro_toplamlar = df_toplam[sayisal_sutunlar].sum().round(2).reset_index()
    mikro_toplamlar.columns = ["Besin Ögesi", "Toplam Alınan Miktar"]
    st.dataframe(mikro_toplamlar, use_container_width=True)

    # Dışa Aktarma
    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        csv_veri = df_toplam.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 Raporu CSV İndir", data=csv_veri, file_name=f"{hasta_no if hasta_no else 'hasta'}_tuketim.csv", mime="text/csv")
    with col_btn2:
        if st.button("🔄 Tümünü Sıfırla / Yeni Hasta"):
            st.session_state.tuketim = {ogun: [] for ogun in ogunler}
            st.rerun()
else:
    st.info("Henüz hiçbir öğüne besin eklenmedi. Yukarıdaki öğün sekmelerini açarak besin ekleyebilirsiniz.")
