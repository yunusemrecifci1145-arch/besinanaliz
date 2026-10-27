import streamlit as st
import pandas as pd
import io
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ==========================================
# 1. SAYFA VE STİL AYARLARI
# ==========================================
st.set_page_config(page_title="Klinik Besin Tüketim Analizi", layout="wide", page_icon="🥗")

st.markdown("""
<style>
    .footer-box {
        margin-top: 60px;
        padding: 20px;
        border-top: 2px solid #e2e8f0;
        background-color: #f8fafc;
        border-radius: 10px;
        text-align: center;
    }
    .footer-name {
        font-size: 26px;
        font-weight: 800;
        color: #0f172a;
        margin: 5px 0;
        letter-spacing: 0.5px;
    }
    .footer-desc {
        font-size: 14px;
        color: #475569;
        margin-bottom: 0px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. VERİ YÜKLEME VE DÜZENLEME
# ==========================================
@st.cache_data
def load_data():
    try:
        df = pd.read_excel('besinler.xlsx')
    except Exception as e:
        st.error(f"Excel yüklenirken hata oluştu: {e}")
        return None, None

    isim_serisi = df['Besin Adı'].fillna(df.get('Yemek Adı', pd.Series(dtype=object))).fillna(df.get('Ürün Adı', pd.Series(dtype=object)))
    df['Besin Adı'] = isim_serisi

    kalori_serisi = df.get('Kalori (kcal)', pd.Series(dtype=float))
    if 'Enerji (kcal)' in df.columns:
        kalori_serisi = kalori_serisi.fillna(df['Enerji (kcal)'])
    if 'Enerji (kcal/100ml)' in df.columns:
        kalori_serisi = kalori_serisi.fillna(df['Enerji (kcal/100ml)'])
    df['Kalori (kcal)'] = kalori_serisi

    if 'Protein (g/100ml)' in df.columns:
        df['Protein (g)'] = df['Protein (g)'].fillna(df['Protein (g/100ml)'])

    if 'Porsiyon (g)' not in df.columns:
        df['Porsiyon (g)'] = 100.0
    else:
        df['Porsiyon (g)'] = pd.to_numeric(df['Porsiyon (g)'], errors='coerce').fillna(100.0)

    df = df.dropna(subset=['Besin Adı']).copy()
    df['Besin Adı'] = df['Besin Adı'].astype(str).str.strip()
    df = df.drop_duplicates(subset=['Besin Adı'])

    for col in df.columns:
        if col not in ['Besin Adı', 'Kategori', 'Kaynak Dosya']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    return df, 'Besin Adı'

df, isim_sutunu = load_data()
if df is None:
    st.stop()

# ==========================================
# 3. SESSION STATE KURULUMU
# ==========================================
ogunler = ["Kahvaltı", "Kuşluk (Ara 1)", "Öğle Yemeği", "İkindi (Ara 2)", "Akşam Yemeği", "Gece (Ara 3)"]

if "mod" not in st.session_state:
    st.session_state.mod = None

if "tuketim" not in st.session_state:
    st.session_state.tuketim = {
        "Gün 1": {ogun: [] for ogun in ogunler},
        "Gün 2": {ogun: [] for ogun in ogunler},
        "Gün 3": {ogun: [] for ogun in ogunler}
    }

# ==========================================
# 4. YARDIMCI FONKSİYONLAR (EXCEL & PDF)
# ==========================================
def generate_excel(df_data, ozet_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_ozet = pd.DataFrame(list(ozet_dict.items()), columns=["Parametre", "Değer"])
        df_ozet.to_excel(writer, sheet_name='Genel Özet', index=False)
        df_data.to_excel(writer, sheet_name='Tüketim Detayı', index=False)
    return output.getvalue()

def generate_pdf(hasta_kodu, ozet_dict, df_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor("#0f172a"), alignment=1)
    sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor("#64748b"), alignment=1)
    head_style = ParagraphStyle('HeadStyle', parent=styles['Heading3'], fontSize=12, leading=15, textColor=colors.HexColor("#1e293b"))
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=8, leading=10)

    elements.append(Paragraph("Klinik Besin Tüketim Raporu", title_style))
    elements.append(Paragraph("Geliştirici: Yunus Emre Çifçi - Beslenme ve Diyetetik", sub_style))
    elements.append(Spacer(1, 15))

    elements.append(Paragraph("Hasta & Değerlendirme Özeti", head_style))
    elements.append(Spacer(1, 6))

    ozet_table_data = [[k, str(v)] for k, v in ozet_dict.items()]
    t_ozet = Table(ozet_table_data, colWidths=[200, 300])
    t_ozet.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#f8fafc")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('PADDING', (0, 0), (-1, -1), 4),
    ]))
    elements.append(t_ozet)
    elements.append(Spacer(1, 15))

    elements.append(Paragraph("Tüketilen Besinler ve Dağılım", head_style))
    elements.append(Spacer(1, 6))

    temel_kolonlar = [c for c in ['Gün', 'Öğün', 'Besin', 'Miktar (g/ml)', 'Kalori (kcal)', 'Protein (g)'] if c in df_data.columns]
    table_content = [temel_kolonlar]

    for _, row in df_data[temel_kolonlar].iterrows():
        row_vals = []
        for col in temel_kolonlar:
            val = row[col]
            if isinstance(val, float):
                row_vals.append(f"{val:.1f}")
            else:
                row_vals.append(str(val)[:24])
        table_content.append(row_vals)

    col_widths = [50, 70, 180, 70, 70, 60] if 'Gün' in temel_kolonlar else [80, 200, 80, 80, 60]
    t_detay = Table(table_content, colWidths=col_widths)
    t_detay.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0f172a")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ('PADDING', (0, 0), (-1, -1), 3),
    ]))
    elements.append(t_detay)

    doc.build(elements)
    return buffer.getvalue()

# ==========================================
# 5. EKRAN 1: KARŞILAMA EKRANI
# ==========================================
if st.session_state.mod is None:
    st.markdown("<h1 style='text-align: center; color: #0f172a;'>🥗 Klinik Besin Tüketim Analizi</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #475569; font-size: 16px;'>Diyetisyenler, stajyerler ve beslenme uzmanları için geliştirilmiş hızlı tüketim analizi aracı.</p>", unsafe_allow_html=True)
    st.write("")

    col_box1, col_box2, col_box3 = st.columns([1, 4, 1])
    with col_box2:
        st.info("""
        Bu sistem üzerinden herhangi bir **üyelik oluşturmadan ve tamamen ücretsiz** olarak hastalarınızın veya danışanlarınızın 1 ya da 3 günlük besin tüketim kayıtlarını pratik bir şekilde girebilirsiniz.
        
        * Girilen besinlerin enerji, makro öge dağılımlarını ve mikro besin detaylarını anlık olarak hesaplar.
        * Sonuçları resmi klinik formata uygun **PDF Raporu** veya **Excel (.xlsx)** belgesi olarak tek tıkla indirmenizi sağlar.
        * 🔒 *Hasta ve tüketim verileri hiçbir harici sunucuya kaydedilmez, sadece tarayıcınız açıkken bellekte tutulur.*
        """)

        st.write("")
        b_c1, b_c2 = st.columns(2)
        with b_c1:
            if st.button("📋 1 Günlük Tüketim Analizi Yap", use_container_width=True):
                st.session_state.mod = "1 Günlük"
                st.rerun()
        with b_c2:
            if st.button("📊 3 Günlük Standart Kayıt Başlat", use_container_width=True):
                st.session_state.mod = "3 Günlük"
                st.rerun()

# ==========================================
# 6. EKRAN 2: VERİ GİRİŞİ VE ANALİZ
# ==========================================
else:
    top_c1, top_c2 = st.columns([4, 1])
    with top_c1:
        st.title(f"🥗 Klinik Besin Tüketim Kaydı ({st.session_state.mod})")
    with top_c2:
        st.write("")
        if st.button("🏠 Ana Sayfaya Dön"):
            st.session_state.mod = None
            st.rerun()

    # Hasta Bilgileri
    col_h1, col_h2, col_h3, col_h4 = st.columns(4)
    with col_h1:
        hasta_no = st.text_input("Hasta No / Ad-Soyad", placeholder="Örn: G-101 / Servis")
    with col_h2:
        agirlik = st.number_input("Ağırlık (kg)", min_value=0.0, value=70.0, step=1.0)
    with col_h3:
        hedef_kcal = st.number_input("Hedef Enerji (kcal)", min_value=0.0, value=2000.0, step=50.0)
    with col_h4:
        hedef_protein = st.number_input("Hedef Protein (g)", min_value=0.0, value=80.0, step=5.0)

    st.markdown("---")

    # Gün Seçimi
    aktif_gunler = ["Gün 1"] if st.session_state.mod == "1 Günlük" else ["Gün 1", "Gün 2", "Gün 3"]
    secilen_gun = st.radio("Kayıt Yapılacak Günü Seçin:", aktif_gunler, horizontal=True) if st.session_state.mod == "3 Günlük" else "Gün 1"

    besin_listesi = ["Seçiniz..."] + sorted(df[isim_sutunu].tolist())

    # Öğün Blokları
    for ogun in ogunler:
        ekli_sayi = len(st.session_state.tuketim[secilen_gun][ogun])
        with st.expander(f"🍽️ {ogun} ({ekli_sayi} besin eklendi)", expanded=False):
            c_besin, c_tip, c_miktar, c_yuzde, c_btn = st.columns([3, 1.5, 1.2, 1.2, 1])
            
            with c_besin:
                secilen_besin = st.selectbox("Besin Seç", options=besin_listesi, key=f"secim_{secilen_gun}_{ogun}")
            
            with c_tip:
                tip = st.radio("Ölçü Türü", ["Gram / ml", "Porsiyon"], horizontal=True, key=f"tip_{secilen_gun}_{ogun}")
            
            with c_miktar:
                if tip == "Porsiyon":
                    miktar_input = st.number_input("Porsiyon Adedi", min_value=0.25, value=1.0, step=0.25, key=f"pors_{secilen_gun}_{ogun}")
                else:
                    miktar_input = st.number_input("Miktar (g / ml)", min_value=1.0, value=100.0, step=10.0, key=f"mkt_{secilen_gun}_{ogun}")

            with c_yuzde:
                yuzde = st.select_slider("Tüketim (%)", options=[25, 50, 75, 100], value=100, key=f"yuz_{secilen_gun}_{ogun}")

            with c_btn:
                st.write("")
                st.write("")
                if st.button("➕ Ekle", key=f"btn_{secilen_gun}_{ogun}"):
                    if secilen_besin != "Seçiniz...":
                        besin_verisi = df[df[isim_sutunu] == secilen_besin].iloc[0].to_dict()
                        std_porsiyon = besin_verisi.get('Porsiyon (g)', 100.0)
                        
                        net_gram = (miktar_input * std_porsiyon) if tip == "Porsiyon" else miktar_input
                        carpan = (net_gram / 100.0) * (yuzde / 100.0)

                        hesaplanan = {
                            "Gün": secilen_gun,
                            "Öğün": ogun,
                            "Besin": secilen_besin,
                            "Giriş Tipi": f"{miktar_input} Porsiyon" if tip == "Porsiyon" else f"{net_gram:.0f} g",
                            "Miktar (g/ml)": net_gram,
                            "Tüketim (%)": yuzde
                        }

                        for k, v in besin_verisi.items():
                            if isinstance(v, (int, float)):
                                hesaplanan[k] = v * carpan

                        st.session_state.tuketim[secilen_gun][ogun].append(hesaplanan)
                        st.rerun()

            # Eklenenleri Listele
            if st.session_state.tuketim[secilen_gun][ogun]:
                for idx, item in enumerate(st.session_state.tuketim[secilen_gun][ogun]):
                    k1, k2, k3, k4 = st.columns([4, 2, 2, 1])
                    k1.write(f"• **{item['Besin']}** ({item['Giriş Tipi']})")
                    k2.write(f"{item['Miktar (g/ml)']:.0f} g/ml (%{item['Tüketim (%)']})")
                    k3.write(f"{item.get('Kalori (kcal)', 0):.1f} kcal | {item.get('Protein (g)', 0):.1f} g prot")
                    if k4.button("❌", key=f"del_{secilen_gun}_{ogun}_{idx}"):
                        st.session_state.tuketim[secilen_gun][ogun].pop(idx)
                        st.rerun()

    st.markdown("---")

    # ==========================================
    # 7. ANALİZ VE RAPORLAMA BÖLÜMÜ
    # ==========================================
    tum_besinler = []
    for g in aktif_gunler:
        for ogun in ogunler:
            tum_besinler.extend(st.session_state.tuketim[g][ogun])

    if tum_besinler:
        df_toplam = pd.DataFrame(tum_besinler)
        gun_sayisi = len(aktif_gunler)

        # Günlük Ortalamalar
        toplam_kcal = df_toplam.get('Kalori (kcal)', pd.Series([0])).sum() / gun_sayisi
        toplam_cho = df_toplam.get('Karbonhidrat (g)', pd.Series([0])).sum() / gun_sayisi
        toplam_prot = df_toplam.get('Protein (g)', pd.Series([0])).sum() / gun_sayisi
        toplam_yag = df_toplam.get('Yağ (g)', pd.Series([0])).sum() / gun_sayisi
        toplam_lif = df_toplam.get('Lif (g)', pd.Series([0])).sum() / gun_sayisi

        st.subheader("📊 Günlük Ortalama Tüketim & Karşılama")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Ortalama Enerji", f"{toplam_kcal:.1f} kcal")
        m2.metric("Karbonhidrat", f"{toplam_cho:.1f} g")
        m3.metric("Protein", f"{toplam_prot:.1f} g")
        m4.metric("Yağ", f"{toplam_yag:.1f} g")
        m5.metric("Lif", f"{toplam_lif:.1f} g")

        # Makro Yüzdeleri ve Klinik Metrikler
        kcal_cho = toplam_cho * 4
        kcal_prot = toplam_prot * 4
        kcal_yag = toplam_yag * 9
        toplam_makro_kcal = kcal_cho + kcal_prot + kcal_yag

        st.markdown("##### 🔬 Klinik Nutrisyonel Değerlendirme")
        c_k1, c_k2, c_k3, c_k4 = st.columns(4)
        
        prot_per_kg = (toplam_prot / agirlik) if agirlik > 0 else 0
        c_k1.info(f"**Protein / kg:** {prot_per_kg:.2f} g/kg")

        if toplam_makro_kcal > 0:
            c_k2.metric("% CHO", f"%{(kcal_cho / toplam_makro_kcal)*100:.1f}")
            c_k3.metric("% Protein", f"%{(kcal_prot / toplam_makro_kcal)*100:.1f}")
            c_k4.metric("% Yağ", f"%{(kcal_yag / toplam_makro_kcal)*100:.1f}")

        # Hedef Karşılama
        if hedef_kcal > 0 or hedef_protein > 0:
            st.markdown("##### 🎯 Hedef Karşılama")
            c_h1, c_h2 = st.columns(2)
            if hedef_kcal > 0:
                oran_kcal = (toplam_kcal / hedef_kcal) * 100
                c_h1.info(f"Enerji Karşılama: **%{oran_kcal:.1f}** ({toplam_kcal:.0f} / {hedef_kcal:.0f} kcal)")
            if hedef_protein > 0:
                oran_prot = (toplam_prot / hedef_protein) * 100
                c_h2.info(f"Protein Karşılama: **%{oran_prot:.1f}** ({toplam_prot:.1f} / {hedef_protein:.1f} g)")

        # Mikro Besin Tablosu
        with st.expander("🔬 Tüm Mikro ve Diğer Besin Ögeleri Dökümü"):
            haric_kolonlar = ['Gün', 'Öğün', 'Besin', 'Giriş Tipi', 'Miktar (g/ml)', 'Tüketim (%)', 'Besin Adı', 'Kategori', 'Kaynak Dosya', 'Porsiyon (g)']
            sayisal_sutunlar = [c for c in df_toplam.columns if c not in haric_kolonlar]
            mikro_toplamlar = (df_toplam[sayisal_sutunlar].sum() / gun_sayisi).round(2).reset_index()
            mikro_toplamlar.columns = ["Besin Ögesi", "Günlük Ortalama Miktar"]
            st.dataframe(mikro_toplamlar, use_container_width=True)

        st.markdown("---")

        # Rapor İndirme Butonları
        st.subheader("📥 Raporu Dışa Aktar")
        ozet_rapor_verisi = {
            "Hasta Kodu": hasta_no if hasta_no else "Belirtilmedi",
            "Kayıt Modeli": st.session_state.mod,
            "Ağırlık (kg)": agirlik,
            "Ortalama Enerji (kcal)": f"{toplam_kcal:.1f}",
            "Ortalama Protein (g)": f"{toplam_prot:.1f}",
            "Protein/Ağırlık (g/kg)": f"{prot_per_kg:.2f}",
            "Karbonhidrat (g)": f"{toplam_cho:.1f}",
            "Yağ (g)": f"{toplam_yag:.1f}",
            "Enerji Karşılama Oranı": f"%{(toplam_kcal / hedef_kcal)*100:.1f}" if hedef_kcal > 0 else "-",
            "Protein Karşılama Oranı": f"%{(toplam_prot / hedef_protein)*100:.1f}" if hedef_protein > 0 else "-"
        }

        col_d1, col_d2, col_d3 = st.columns([1.5, 1.5, 3])
        with col_d1:
            excel_data = generate_excel(df_toplam, ozet_rapor_verisi)
            st.download_button(
                label="📊 Excel Raporu İndir (.xlsx)",
                data=excel_data,
                file_name=f"{hasta_no if hasta_no else 'hasta'}_tuketim.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_d2:
            pdf_data = generate_pdf(hasta_no, ozet_rapor_verisi, df_toplam)
            st.download_button(
                label="📄 PDF Raporu İndir (.pdf)",
                data=pdf_data,
                file_name=f"{hasta_no if hasta_no else 'hasta'}_tuketim.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        with col_d3:
            st.write("")
            if st.button("🔄 Tümünü Sıfırla / Yeni Hasta", type="secondary"):
                st.session_state.tuketim = {
                    "Gün 1": {ogun: [] for ogun in ogunler},
                    "Gün 2": {ogun: [] for ogun in ogunler},
                    "Gün 3": {ogun: [] for ogun in ogunler}
                }
                st.rerun()
    else:
        st.info("Henüz hiçbir öğüne besin eklenmedi. Yukarıdaki öğün sekmelerini açarak besin ekleyebilirsiniz.")

# ==========================================
# 8. EN ALT İMZA ALANI (FOOTER)
# ==========================================
st.markdown("""
<div class="footer-box">
    <div class="footer-name">Yunus Emre Çifçi</div>
    <div class="footer-desc">Beslenme ve Diyetetik / Klinik Nutrisyonel Veri Analiz Platformu</div>
</div>
""", unsafe_allow_html=True)
