import streamlit as st
import pandas as pd
import io
import os

# ReportLab ve PDF Kütüphaneleri
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ==========================================
# 1. SAYFA VE STİL AYARLARI
# ==========================================
st.set_page_config(page_title="Klinik Besin Tüketim Analizi", layout="wide", page_icon="🥗")

st.markdown("""
<style>
    .footer-box {
        margin-top: 50px;
        padding: 24px;
        border-top: 1px solid rgba(128, 128, 128, 0.2);
        background-color: var(--secondary-background-color);
        border-radius: 12px;
        text-align: center;
    }
    .footer-name {
        font-size: 24px;
        font-weight: 800;
        color: var(--text-color);
        margin: 4px 0;
        letter-spacing: 0.5px;
    }
    .footer-desc {
        font-size: 13px;
        color: var(--text-color);
        opacity: 0.8;
        margin-bottom: 0px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. VERİ YÜKLEME (Sessiz ve Çift Excel Destekli)
# ==========================================
@st.cache_data
def load_data():
    dfs = []
    
    # Standart Besinler
    if os.path.exists('besinler.xlsx'):
        try:
            df_b = pd.read_excel('besinler.xlsx')
            dfs.append(df_b)
        except Exception:
            pass

    # Enteral Mamalar (Varsa otomatik içine katar)
    if os.path.exists('enteral.xlsx'):
        try:
            df_e = pd.read_excel('enteral.xlsx')
            dfs.append(df_e)
        except Exception:
            pass

    if not dfs:
        st.error("Veri tabanı bulunamadı! 'besinler.xlsx' dosyasının yüklü olduğundan emin olun.")
        return None, None

    df = pd.concat(dfs, ignore_index=True)

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
def tr_fix(text):
    if text is None:
        return ""
    s = str(text)
    ceviri = {
        'ı': 'i', 'İ': 'I',
        'ş': 's', 'Ş': 'S',
        'ğ': 'g', 'Ğ': 'G',
        'ü': 'u', 'Ü': 'U',
        'ö': 'o', 'Ö': 'O',
        'ç': 'c', 'Ç': 'C'
    }
    for tr, eng in ceviri.items():
        s = s.replace(tr, eng)
    return s

def generate_excel(df_data, ozet_dict, df_mikro, klinik_not):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_ozet = pd.DataFrame(list(ozet_dict.items()), columns=["Parametre", "Değer"])
        df_ozet.to_excel(writer, sheet_name='Genel Özet', index=False)
        df_data.to_excel(writer, sheet_name='Tüketim Detayı', index=False)
        df_mikro.to_excel(writer, sheet_name='Mikro Besinler', index=False)
        if klinik_not.strip():
            pd.DataFrame([{"Klinik Not": klinik_not}]).to_excel(writer, sheet_name='Klinik Yorum', index=False)
    return output.getvalue()

def generate_pdf(hasta_kodu, ozet_dict, df_data, df_mikro, klinik_not):
    font_secildi = False
    linux_font_yollari = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSans.ttf"
    ]
    
    secilen_font = "Helvetica"
    secilen_bold = "Helvetica-Bold"

    for font_yolu in linux_font_yollari:
        if os.path.exists(font_yolu):
            try:
                pdfmetrics.registerFont(TTFont('CustomTR', font_yolu))
                secilen_font = 'CustomTR'
                secilen_bold = 'CustomTR'
                font_secildi = True
                break
            except Exception:
                pass

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('TitleStyle', fontName=secilen_bold, fontSize=15, leading=18, textColor=colors.HexColor("#0f172a"), alignment=1)
    sub_style = ParagraphStyle('SubStyle', fontName=secilen_font, fontSize=8, leading=11, textColor=colors.HexColor("#64748b"), alignment=1)
    head_style = ParagraphStyle('HeadStyle', fontName=secilen_bold, fontSize=10, leading=13, textColor=colors.HexColor("#1e293b"))
    cell_style = ParagraphStyle('CellStyle', fontName=secilen_font, fontSize=7.5, leading=9.5, textColor=colors.HexColor("#0f172a"))
    cell_style_bold = ParagraphStyle('CellStyleBold', fontName=secilen_bold, fontSize=7.5, leading=9.5, textColor=colors.white)
    not_style = ParagraphStyle('NotStyle', fontName=secilen_font, fontSize=8, leading=11, textColor=colors.HexColor("#334155"))

    t_baslik = "Klinik Besin Tuketim Raporu" if not font_secildi else "Klinik Besin Tüketim Raporu"
    t_yazar = "Gelistirici: Yunus Emre Cifci - Beslenme ve Diyetetik" if not font_secildi else "Geliştirici: Yunus Emre Çifçi - Beslenme ve Diyetetik"
    t_ozet_baslik = "Hasta & Makro Besin Dagilim Ozeti" if not font_secildi else "Hasta & Makro Besin Dağılım Özeti"
    t_detay_baslik = "Tuketilen Besinler ve Dagilim" if not font_secildi else "Tüketilen Besinler ve Dağılım"
    t_mikro_baslik = "Gunluk Toplam Mikro ve Diger Besin Ogeleri Dokumu" if not font_secildi else "Günlük Toplam Mikro ve Diğer Besin Ögeleri Dökümü"
    t_not_baslik = "Diyetisyen Klinik Notu / Oneri" if not font_secildi else "Diyetisyen Klinik Notu / Öneri"

    elements.append(Paragraph(t_baslik, title_style))
    elements.append(Paragraph(t_yazar, sub_style))
    elements.append(Spacer(1, 10))

    # Özet Tablosu
    elements.append(Paragraph(t_ozet_baslik, head_style))
    elements.append(Spacer(1, 4))
    ozet_table_data = []
    for k, v in ozet_dict.items():
        k_metin = k if font_secildi else tr_fix(k)
        v_metin = str(v) if font_secildi else tr_fix(v)
        ozet_table_data.append([Paragraph(k_metin, cell_style), Paragraph(v_metin, cell_style)])

    t_ozet = Table(ozet_table_data, colWidths=[220, 280])
    t_ozet.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#f8fafc")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ('PADDING', (0, 0), (-1, -1), 2.5),
    ]))
    elements.append(t_ozet)
    elements.append(Spacer(1, 10))

    # Klinik Not (Varsa)
    if klinik_not.strip():
        elements.append(Paragraph(t_not_baslik, head_style))
        elements.append(Spacer(1, 4))
        p_not = Paragraph(klinik_not if font_secildi else tr_fix(klinik_not), not_style)
        t_not = Table([[p_not]], colWidths=[500])
        t_not.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#f1f5f9")),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#94a3b8")),
            ('PADDING', (0, 0), (-1, -1), 6),
        ]))
        elements.append(t_not)
        elements.append(Spacer(1, 10))

    # Tüketilen Besinler Tablosu
    elements.append(Paragraph(t_detay_baslik, head_style))
    elements.append(Spacer(1, 4))
    temel_kolonlar = [c for c in ['Gün', 'Öğün', 'Besin', 'Miktar (g/ml)', 'Kalori (kcal)', 'Protein (g)'] if c in df_data.columns]
    table_content = []
    baslik_satiri = []
    for col in temel_kolonlar:
        c_metin = col if font_secildi else tr_fix(col)
        baslik_satiri.append(Paragraph(c_metin, cell_style_bold))
    table_content.append(baslik_satiri)

    for _, row in df_data[temel_kolonlar].iterrows():
        row_vals = []
        for col in temel_kolonlar:
            val = row[col]
            text_val = f"{val:.1f}" if isinstance(val, float) else str(val)[:24]
            final_text = text_val if font_secildi else tr_fix(text_val)
            row_vals.append(Paragraph(final_text, cell_style))
        table_content.append(row_vals)

    col_widths = [50, 70, 180, 70, 70, 60] if 'Gün' in temel_kolonlar else [80, 200, 80, 80, 60]
    t_detay = Table(table_content, colWidths=col_widths)
    t_detay.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0f172a")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ('PADDING', (0, 0), (-1, -1), 2.5),
    ]))
    elements.append(t_detay)
    elements.append(Spacer(1, 10))

    # Mikro Besin Tablosu
    elements.append(Paragraph(t_mikro_baslik, head_style))
    elements.append(Spacer(1, 4))
    mikro_rows = []
    for _, r in df_mikro.iterrows():
        ad = str(r["Besin Ögesi"]) if font_secildi else tr_fix(r["Besin Ögesi"])
        deg = f"{r['Günlük Ortalama Miktar']}"
        mikro_rows.append((ad, deg))

    mikro_tablo_data = [[
        Paragraph("<b>Besin Ögesi</b>", cell_style_bold),
        Paragraph("<b>Miktar</b>", cell_style_bold),
        Paragraph("<b>Besin Ögesi</b>", cell_style_bold),
        Paragraph("<b>Miktar</b>", cell_style_bold)
    ]]

    yarisi = (len(mikro_rows) + 1) // 2
    for i in range(yarisi):
        sol_ad, sol_deg = mikro_rows[i]
        sag_ad, sag_deg = mikro_rows[i + yarisi] if (i + yarisi) < len(mikro_rows) else ("", "")
        mikro_tablo_data.append([
            Paragraph(sol_ad, cell_style),
            Paragraph(sol_deg, cell_style),
            Paragraph(sag_ad, cell_style),
            Paragraph(sag_deg, cell_style)
        ])

    t_mikro = Table(mikro_tablo_data, colWidths=[170, 80, 170, 80])
    t_mikro.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#334155")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ('PADDING', (0, 0), (-1, -1), 2),
    ]))
    elements.append(t_mikro)

    doc.build(elements)
    return buffer.getvalue()

# ==========================================
# 5. EKRAN 1: KARŞILAMA EKRANI
# ==========================================
if st.session_state.mod is None:
    st.markdown("<h1 style='text-align: center;'>🥗 Klinik Besin Tüketim Analizi</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.8; font-size: 16px;'>Diyetisyenler ve stajyerler için pratik tüketim analiz aracı.</p>", unsafe_allow_html=True)
    st.write("")

    col_box1, col_box2, col_box3 = st.columns([1, 4, 1])
    with col_box2:
        st.info("""
        Bu sistem üzerinden herhangi bir **üyelik oluşturmadan ve tamamen ücretsiz** olarak 1 ya da 3 günlük besin tüketim kayıtlarını pratik bir şekilde girebilirsiniz.
        
        * Girilen besinlerin enerji, makro öge dağılımlarını ve mikro besin detaylarını anlık olarak hesaplar.
        * Sonuçları resmi klinik formata uygun **PDF Raporu** veya **Excel (.xlsx)** belgesi olarak tek tıkla indirmenizi sağlar.
        * 🔒 *Hasta verileri sunucuya kaydedilmez, sadece tarayıcınız açıkken işlenir.*
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
    col_h1, col_h2 = st.columns(2)
    with col_h1:
        hasta_no = st.text_input("Hasta No / Ad-Soyad", placeholder="Örn: G-101 / Servis")
    with col_h2:
        agirlik = st.number_input("Vücut Ağırlığı (kg)", min_value=1.0, value=70.0, step=1.0)

    st.markdown("---")

    aktif_gunler = ["Gün 1"] if st.session_state.mod == "1 Günlük" else ["Gün 1", "Gün 2", "Gün 3"]
    secilen_gun = st.radio("Kayıt Yapılacak Günü Seçin:", aktif_gunler, horizontal=True) if st.session_state.mod == "3 Günlük" else "Gün 1"

    arama_placeholder = "Seçiniz veya aramak için yazınız..."
    besin_listesi = [arama_placeholder] + sorted(df[isim_sutunu].tolist())

    # Öğün Blokları
    for ogun in ogunler:
        ekli_sayi = len(st.session_state.tuketim[secilen_gun][ogun])
        with st.expander(f"🍽️ {ogun} ({ekli_sayi} besin eklendi)", expanded=False):
            c_besin, c_tip, c_miktar, c_yuzde, c_btn = st.columns([2.8, 1.2, 1.2, 1.8, 0.8])
            
            with c_besin:
                secilen_besin = st.selectbox("Besin Ara / Seç", options=besin_listesi, key=f"secim_{secilen_gun}_{ogun}")
            
            with c_tip:
                tip = st.radio("Ölçü Türü", ["Gram / ml", "Porsiyon"], horizontal=True, key=f"tip_{secilen_gun}_{ogun}")
            
            with c_miktar:
                if tip == "Porsiyon":
                    miktar_input = st.number_input("Porsiyon Adedi", min_value=0.1, value=1.0, step=0.25, key=f"pors_{secilen_gun}_{ogun}")
                else:
                    miktar_input = st.number_input("Miktar (g / ml)", min_value=1.0, value=100.0, step=10.0, key=f"mkt_{secilen_gun}_{ogun}")

            with c_yuzde:
                yuzde = st.slider(
                    "Tüketim Miktarı (%):",
                    min_value=5,
                    max_value=100,
                    value=100,
                    step=5,
                    help="Hastanın tabağında veya porsiyonunda tükettiği oran",
                    key=f"yuz_{secilen_gun}_{ogun}"
                )

            with c_btn:
                st.write("")
                st.write("")
                if st.button("➕ Ekle", key=f"btn_{secilen_gun}_{ogun}"):
                    if secilen_besin != arama_placeholder:
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

        st.subheader("📊 Günlük Ortalama Tüketim & Makro Dağılımı")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Ortalama Enerji", f"{toplam_kcal:.1f} kcal")
        m2.metric("Karbonhidrat", f"{toplam_cho:.1f} g")
        m3.metric("Protein", f"{toplam_prot:.1f} g")
        m4.metric("Yağ", f"{toplam_yag:.1f} g")
        m5.metric("Lif", f"{toplam_lif:.1f} g")

        # Makro Yüzde Hesapları
        kcal_cho = toplam_cho * 4
        kcal_prot = toplam_prot * 4
        kcal_yag = toplam_yag * 9
        toplam_makro_kcal = kcal_cho + kcal_prot + kcal_yag

        yuzde_cho = (kcal_cho / toplam_makro_kcal * 100) if toplam_makro_kcal > 0 else 0
        yuzde_prot = (kcal_prot / toplam_makro_kcal * 100) if toplam_makro_kcal > 0 else 0
        yuzde_yag = (kcal_yag / toplam_makro_kcal * 100) if toplam_makro_kcal > 0 else 0
        prot_per_kg = (toplam_prot / agirlik) if agirlik > 0 else 0

        st.markdown("##### 🔬 Enerjinin Makro Dağılımı (% Enerji)")
        c_k1, c_k2, c_k3, c_k4 = st.columns(4)
        c_k1.info(f"**Protein / kg:** {prot_per_kg:.2f} g/kg")
        c_k2.metric("Karbonhidrat Payı", f"%{yuzde_cho:.1f}")
        c_k3.metric("Protein Payı", f"%{yuzde_prot:.1f}")
        c_k4.metric("Yağ Payı", f"%{yuzde_yag:.1f}")

        # Mikro Besin Tablosu
        haric_kolonlar = ['Gün', 'Öğün', 'Besin', 'Giriş Tipi', 'Miktar (g/ml)', 'Tüketim (%)', 'Besin Adı', 'Kategori', 'Kaynak Dosya', 'Porsiyon (g)']
        sayisal_sutunlar = [c for c in df_toplam.columns if c not in haric_kolonlar]
        mikro_toplamlar = (df_toplam[sayisal_sutunlar].sum() / gun_sayisi).round(2).reset_index()
        mikro_toplamlar.columns = ["Besin Ögesi", "Günlük Ortalama Miktar"]

        with st.expander("🔬 Tüm Mikro ve Diğer Besin Ögeleri Dökümü", expanded=True):
            st.dataframe(mikro_toplamlar, use_container_width=True)

        st.markdown("---")

        # Diyetisyen Not Alanı
        st.subheader("📝 Diyetisyen Klinik Notu / Öneri")
        klinik_notu = st.text_area(
            "Hastaya veya epikrize eklemek istediğiniz değerlendirme notu (PDF ve Excel'e eklenecektir):",
            placeholder="Örn: Oral alım yetersiz olup enerji ihtiyacının %60'ı karşılanmıştır. Enteral ürün desteği önerilir."
        )

        st.markdown("---")

        # Rapor İndirme Butonları
        st.subheader("📥 Raporu Dışa Aktar")
        ozet_rapor_verisi = {
            "Hasta Kodu": hasta_no if hasta_no else "Belirtilmedi",
            "Kayit Modeli": st.session_state.mod,
            "Agirlik (kg)": agirlik,
            "Toplam Enerji (kcal)": f"{toplam_kcal:.1f}",
            "Protein (g)": f"{toplam_prot:.1f}",
            "Protein / Agirlik (g/kg)": f"{prot_per_kg:.2f}",
            "Karbonhidrat (g)": f"{toplam_cho:.1f}",
            "Yag (g)": f"{toplam_yag:.1f}",
            "Lif (g)": f"{toplam_lif:.1f}",
            "Enerjinin CHO Payi (%)": f"%{yuzde_cho:.1f}",
            "Enerjinin Protein Payi (%)": f"%{yuzde_prot:.1f}",
            "Enerjinin Yag Payi (%)": f"%{yuzde_yag:.1f}"
        }

        col_d1, col_d2, col_d3 = st.columns([1.5, 1.5, 3])
        with col_d1:
            excel_data = generate_excel(df_toplam, ozet_rapor_verisi, mikro_toplamlar, klinik_notu)
            st.download_button(
                label="📊 Excel Raporu İndir (.xlsx)",
                data=excel_data,
                file_name=f"{hasta_no if hasta_no else 'hasta'}_tuketim.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_d2:
            pdf_data = generate_pdf(hasta_no, ozet_rapor_verisi, df_toplam, mikro_toplamlar, klinik_notu)
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
