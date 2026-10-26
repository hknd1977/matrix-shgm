import io
import os
import re
import sqlite3
import networkx as nx
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# PDF Raporlama ve Font Yönetimi için ReportLab Kütüphaneleri
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

# 1. Sayfa Yapılandırması
st.set_page_config(
    page_title="Havacılık Lisans & Geçiş Rehberi",
    page_icon="✈️",
    layout="wide"
)

DB_PATH = "havacilik_egitim.db"
TABLE_NAME = "lisans_ders_matrisi"


# --- TÜRKÇE FONT VE ARAMA YARDIMCILARI ---

def register_turkish_pdf_font():
    """
    İşletim sistemindeki Unicode Türkçe destekli TrueType fontları bulup ReportLab'e kaydeder.
    """
    possible_fonts = [
        # Linux (Debian / Ubuntu / Fedora / Streamlit Cloud)
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"),
        ("/usr/share/fonts/TTF/DejaVuSans.ttf", "/usr/share/fonts/TTF/DejaVuSans-Bold.ttf"),
        # Windows
        ("C:\\Windows\\Fonts\\arial.ttf", "C:\\Windows\\Fonts\\arialbd.ttf"),
        ("C:\\Windows\\Fonts\\calibri.ttf", "C:\\Windows\\Fonts\\calibrib.ttf"),
        # macOS
        ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf"),
        ("/System/Library/Fonts/Supplemental/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
    ]

    for reg_path, bold_path in possible_fonts:
        if os.path.exists(reg_path):
            try:
                pdfmetrics.registerFont(TTFont('TRFont', reg_path))
                if os.path.exists(bold_path):
                    pdfmetrics.registerFont(TTFont('TRFont-Bold', bold_path))
                else:
                    pdfmetrics.registerFont(TTFont('TRFont-Bold', reg_path))
                return 'TRFont', 'TRFont-Bold'
            except Exception:
                continue
    return 'Helvetica', 'Helvetica-Bold'


def tr_normalize(text: str) -> str:
    """Türkçe karakter uyumsuzluğunu (İ/i, I/ı vb.) çözer ve küçük harfe dönüştürür."""
    if not text:
        return ""
    text = text.replace('İ', 'i').replace('I', 'ı')
    tr_map = str.maketrans("çğışöü", "cgisou")
    return text.lower().translate(tr_map)


def smart_search(query: str, choices: list[str]) -> list[str]:
    """Sorgudaki kelimeleri bağımsız analiz ederek Türkçe duyarlı akıllı arama yapar."""
    if not query or not query.strip():
        return choices

    q_norm = tr_normalize(query)
    q_tokens = q_norm.split()

    scored_results = []
    for choice in choices:
        c_norm = tr_normalize(choice)
        score = 0

        if q_norm in c_norm:
            score += 100

        matched_tokens = sum(1 for token in q_tokens if token in c_norm)

        if matched_tokens == len(q_tokens):
            score += 50 + (matched_tokens * 10)
        elif matched_tokens > 0:
            score += matched_tokens * 10

        if score > 0:
            scored_results.append((score, choice))

    scored_results.sort(key=lambda x: x[0], reverse=True)
    filtered = [item[1] for item in scored_results]
    return filtered if filtered else choices


@st.cache_data
def load_data():
    """Veritabanından tüm matrisi çeker ve veri tiplerini düzenler."""
    with sqlite3.connect(DB_PATH) as conn:
        df = pd.read_sql_query(f"SELECT * FROM {TABLE_NAME}", conn)

    df["gecme_notu"] = pd.to_numeric(df["gecme_notu"], errors="coerce").fillna(0).astype(int)
    return df


df_raw = load_data()

if df_raw.empty:
    st.error("Veritabanında henüz kayıtlı ders bulunamadı!")
    st.stop()

# State Yönetimi
tum_lisanslar = sorted(df_raw["lisans_adi"].unique().tolist())
tum_dersler = sorted(df_raw["ders_kodu_ve_adi"].unique().tolist())

if "selected_license" not in st.session_state:
    st.session_state.selected_license = tum_lisanslar[0]

if "selected_course" not in st.session_state:
    st.session_state.selected_course = tum_dersler[0]


# --- RAPOR ÜRETME FONKSİYONLARI ---

def create_excel_report(secilen_lisans, mevcut_dersler_df, df_tam_100, df_kismi):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        mevcut_dersler_df.to_excel(writer, sheet_name='Mevcut Müfredat', index=False)
        if not df_tam_100.empty:
            df_tam_100.to_excel(writer, sheet_name='100% Uyumlu Lisanslar', index=False)
        if not df_kismi.empty:
            df_kismi.to_excel(writer, sheet_name='Gereksinimli Geçişler', index=False)
    output.seek(0)
    return output.getvalue()


def create_pdf_report(secilen_lisans, mevcut_dersler_df, top_n_df):
    """Türkçe karakterleri düzgün basan PDF Raporu üretir."""
    font_reg, font_bold = register_turkish_pdf_font()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'], fontName=font_bold, fontSize=15, leading=19, textColor=colors.HexColor('#1A237E'), alignment=1, spaceAfter=15
    )
    subtitle_style = ParagraphStyle(
        'SubtitleStyle', parent=styles['Heading2'], fontName=font_bold, fontSize=11, leading=15, textColor=colors.HexColor('#0D47A1'), spaceBefore=10, spaceAfter=8
    )
    body_style = ParagraphStyle(
        'BodyStyle', parent=styles['Normal'], fontName=font_reg, fontSize=8, leading=11
    )
    header_cell_style = ParagraphStyle(
        'HeaderCellStyle', parent=styles['Normal'], fontName=font_bold, fontSize=8, leading=11, textColor=colors.HexColor('#1A237E')
    )

    story.append(Paragraph("<b>HAVACILIK LİSANS ANALİZ VE MÜFREDAT RAPORU</b>", title_style))
    story.append(Paragraph(f"<b>İncelenen Lisans:</b> {secilen_lisans}", subtitle_style))
    story.append(Spacer(1, 10))

    # 1. Mevcut Dersler Tablosu
    story.append(Paragraph("<b>1. Mevcut Müfredat Dersleri</b>", subtitle_style))

    table_data = [[
        Paragraph("<b>Ders Kodu ve Adı</b>", header_cell_style),
        Paragraph("<b>Eğitim Cinsi</b>", header_cell_style),
        Paragraph("<b>Veriliş Şekli</b>", header_cell_style),
        Paragraph("<b>Sınav Tipi</b>", header_cell_style),
        Paragraph("<b>Geçme Notu</b>", header_cell_style)
    ]]

    for _, row in mevcut_dersler_df.iterrows():
        table_data.append([
            Paragraph(str(row["ders_kodu_ve_adi"]), body_style),
            Paragraph(str(row["egitim_cinsi"]), body_style),
            Paragraph(str(row["verilis_sekli"]), body_style),
            Paragraph(str(row["sinav_tipi"]), body_style),
            Paragraph(str(row["gecme_notu"]), body_style)
        ])

    t1 = Table(table_data, colWidths=[200, 85, 85, 95, 55])
    t1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E8EAF6')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#C5CAE9')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t1)
    story.append(Spacer(1, 15))

    # 2. En Yakın Geçiş Yapılabilecek Lisanslar Tablosu
    if not top_n_df.empty:
        story.append(Paragraph("<b>2. En Az Ek Emekle Geçilebilecek Lisanslar</b>", subtitle_style))
        top_table_data = [[
            Paragraph("<b>Hedef Lisans Adı</b>", header_cell_style),
            Paragraph("<b>Tamamlama (%)</b>", header_cell_style),
            Paragraph("<b>Muaf Ders</b>", header_cell_style),
            Paragraph("<b>Eksik Ders</b>", header_cell_style),
            Paragraph("<b>Not Yükseltme</b>", header_cell_style)
        ]]

        for _, row in top_n_df.iterrows():
            top_table_data.append([
                Paragraph(str(row["Lisans Adı"]), body_style),
                Paragraph(f"%{row['Tamamlama Oranı (%)']}", body_style),
                Paragraph(str(row["Muaf / Ortak Ders"]), body_style),
                Paragraph(str(row["Sıfırdan Alınacak Ders"]), body_style),
                Paragraph(str(row["Not Yükseltme Gereken"]), body_style)
            ])

        t2 = Table(top_table_data, colWidths=[180, 80, 80, 80, 100])
        t2.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E8F5E9')),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#A5D6A7')),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        story.append(t2)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# 2. Yan Panel (Sidebar) & Akıllı Arama Süzgeçleri
st.sidebar.header("🎯 Analiz ve Süzgeç Filtreleri")

with st.sidebar.popover("ℹ️ Kullanım Kılavuzu & İpuçları"):
    st.markdown("### 🎓 Kim Nasıl Kullanmalı?")
    st.markdown("""
    **👨‍🏫 Akademisyenler / Danışmanlar İçin:**
    - Öğrencinizin mevcut lisansını seçerek en az ek yükle geçiş yapabileceği lisansları görün.
    - PDF ve Excel raporlarında Türkçe karakterler tam uyumlu üretilir.

    **🧑‍🎓 Öğrenciler İçin:**
    - Seçtiğiniz dersin en yüksek not barajı isteyen lisansını görün (**Tab 2**).
    """)

search_query_lisans = st.sidebar.text_input(
    "🔎 Akıllı Lisans Ara (ör. 'iha pilot'):",
    value="",
    help="Lisans adının tamamını bilmeseniz bile hatırladığınız kelimeleri aralarına boşluk koyarak yazabilirsiniz."
)

akilli_lisanslar = smart_search(search_query_lisans, tum_lisanslar)

secilen_lisans = st.sidebar.selectbox(
    "Mevcut / İncelemek İstediğiniz Ana Lisans:",
    options=akilli_lisanslar,
    index=0,
    key="license_selector",
    help="💡 Öğrencinin hali hazırda sahip olduğu veya ana hedefi olan lisansı seçin."
)

n_yakin = st.sidebar.slider(
    "Görüntülenecek Liste Büyüklüğü (N):",
    min_value=1,
    max_value=20,
    value=5,
    help="💡 Grafik ve tablolarda ilk kaç adet en yakın veya en yüksek barajlı lisansın listeleneceğini belirler."
)

# 3. Hesaplama Algoritması: Lisans Yakınlık Analizi
mevcut_dersler_df = df_raw[df_raw["lisans_adi"] == secilen_lisans]
mevcut_ders_dict = dict(zip(mevcut_dersler_df["ders_kodu_ve_adi"], mevcut_dersler_df["gecme_notu"]))

yakinlik_listesi = []

for lisans in tum_lisanslar:
    if lisans == secilen_lisans:
        continue

    hedef_df = df_raw[df_raw["lisans_adi"] == lisans]
    hedef_dersler = dict(zip(hedef_df["ders_kodu_ve_adi"], hedef_df["gecme_notu"]))

    toplam_hedef_ders = len(hedef_dersler)
    muaf_ders_sayisi = 0
    not_yukseltme_gereken_sayisi = 0
    eksik_ders_sayisi = 0

    for ders, hedef_not in hedef_dersler.items():
        if ders in mevcut_ders_dict:
            mevcut_not = mevcut_ders_dict[ders]
            if mevcut_not >= hedef_not:
                muaf_ders_sayisi += 1
            else:
                not_yukseltme_gereken_sayisi += 1
        else:
            eksik_ders_sayisi += 1

    tamamlama_yuzdesi = round((muaf_ders_sayisi / toplam_hedef_ders) * 100, 1) if toplam_hedef_ders > 0 else 0

    yakinlik_listesi.append({
        "Lisans Adı": lisans,
        "Tamamlama Oranı (%)": tamamlama_yuzdesi,
        "Muaf / Ortak Ders": muaf_ders_sayisi,
        "Not Yükseltme Gereken": not_yukseltme_gereken_sayisi,
        "Sıfırdan Alınacak Ders": eksik_ders_sayisi,
        "Toplam Ders": toplam_hedef_ders
    })

df_yakinlik = pd.DataFrame(yakinlik_listesi).sort_values(
    by=["Tamamlama Oranı (%)", "Sıfırdan Alınacak Ders"],
    ascending=[False, True]
)

df_tam_100 = df_yakinlik[df_yakinlik["Tamamlama Oranı (%)"] == 100.0]
df_kismi = df_yakinlik[df_yakinlik["Tamamlama Oranı (%)"] < 100.0]

# Üst Başlık ve İndirme Butonları Alanı
col_title, col_download = st.columns([3, 2])

with col_title:
    st.title("✈️ Havacılık Eğitimi Lisans Geçiş Rehberi")

with col_download:
    st.write("")
    st.write("##### 📤 Özet Rapor İndir")
    btn_col1, btn_col2 = st.columns(2)

    excel_bytes = create_excel_report(secilen_lisans, mevcut_dersler_df, df_tam_100, df_kismi)
    btn_col1.download_button(
        label="📊 Excel Raporu",
        data=excel_bytes,
        file_name=f"{secilen_lisans}_Lisans_Raporu.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    pdf_bytes = create_pdf_report(secilen_lisans, mevcut_dersler_df, df_kismi.head(n_yakin))
    btn_col2.download_button(
        label="📄 PDF Raporu",
        data=pdf_bytes,
        file_name=f"{secilen_lisans}_Lisans_Raporu.pdf",
        mime="application/pdf",
        use_container_width=True
    )

# 4. Sekmeli Arayüz (Tabs)
tab1, tab2, tab3, tab4 = st.tabs([
    "🎯 Lisans Geçiş Rehberi",
    "📚 Ders Sayfası & Not Barajı Filtresi",
    "📊 Sertifika Kaldıraç Analizi",
    "🕸️ Lisanslar Arası İlişki Ağı"
])

# --- TAB 1: LİSANS GEÇİŞ REHBERİ ---
with tab1:
    st.caption(f"Şu an incelenen lisans: **{secilen_lisans}**")
    st.subheader(f"'{secilen_lisans}' İle En Az Emekle Alınabilecek Lisanslar")

    top_n_df = df_kismi.head(n_yakin)

    if not top_n_df.empty:
        col1, col2 = st.columns([3, 2])

        with col1:
            dynamic_height = max(380, len(top_n_df) * 35)
            fig = px.bar(
                top_n_df,
                x="Tamamlama Oranı (%)",
                y="Lisans Adı",
                orientation="h",
                text="Tamamlama Oranı (%)",
                color="Tamamlama Oranı (%)",
                color_continuous_scale="Blues",
                title=f"En Yüksek Uyumluluğa Sahip İlk {len(top_n_df)} Lisans (< %100)"
            )
            fig.update_layout(
                height=dynamic_height,
                yaxis={"categoryorder": "total ascending"},
                margin=dict(l=10, r=10, t=40, b=10)
            )
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            st.write("##### Ek Ders Gerektiren Lisanslar Listesi")
            st.dataframe(
                top_n_df[["Lisans Adı", "Tamamlama Oranı (%)", "Sıfırdan Alınacak Ders", "Not Yükseltme Gereken"]],
                hide_index=True,
                use_container_width=True,
                height=dynamic_height
            )
    else:
        st.info("Kısmi ek ders gerektiren lisans bulunamadı.")

    st.divider()
    st.subheader("🎓 Doğrudan Alınabilir / %100 Uyumlu Lisanslar")
    if not df_tam_100.empty:
        st.success(f"**{secilen_lisans}** müfredatınız sayesinde hiçbir ek ders almadan doğrudan hak kazandığınız **{len(df_tam_100)}** adet lisans bulunmaktadır:")
        st.dataframe(
            df_tam_100[["Lisans Adı", "Tamamlama Oranı (%)", "Muaf / Ortak Ders", "Toplam Ders"]],
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("Mevcut müfredatınızla birebir aynı (%100) başka bir lisans bulunmuyor.")

    st.divider()
    st.subheader(f"📋 {secilen_lisans} — Mevcut Müfredat Dersleri")
    st.dataframe(mevcut_dersler_df, use_container_width=True, hide_index=True)


# --- TAB 2: MÜSTAKİL DERS SAYFASI & NOT BARAJI FİLTRESİ ---
with tab2:
    st.subheader("📚 Müstakil Ders Sayfası & Not Barajı Süzgeci")

    col_search_course, col_select_course, col_filter_type, col_filter_val = st.columns([1.5, 2, 1, 1])

    with col_search_course:
        search_query_ders = st.text_input(
            "🔎 Akıllı Ders Ara:",
            value="",
            help="Aradığınız dersin kodunu veya hatırladığınız kelimeleri yazın."
        )

    akilli_dersler = smart_search(search_query_ders, tum_dersler)

    with col_select_course:
        secilen_ders = st.selectbox(
            "Seçilen Ders:",
            options=akilli_dersler,
            index=0,
            key="course_selector"
        )

    with col_filter_type:
        not_filtre_yonu = st.radio(
            "Not Baraj Şartı:",
            options=["Tümü", "Barajı Üstünde (≥)", "Barajı Altında (≤)"],
            horizontal=False
        )

    with col_filter_val:
        esik_not = st.number_input("Not Değeri (X):", min_value=0, max_value=100, value=70, step=5)

    if secilen_ders:
        st.session_state.selected_course = secilen_ders
        ders_veri = df_raw[df_raw["ders_kodu_ve_adi"] == secilen_ders]

        if not_filtre_yonu == "Barajı Üstünde (≥)":
            filitreli_ders_veri = ders_veri[ders_veri["gecme_notu"] >= esik_not]
        elif not_filtre_yonu == "Barajı Altında (≤)":
            filitreli_ders_veri = ders_veri[ders_veri["gecme_notu"] <= esik_not]
        else:
            filitreli_ders_veri = ders_veri.copy()

        toplam_lisans_sayisi = ders_veri["lisans_adi"].nunique()
        filtreli_lisans_sayisi = filitreli_ders_veri["lisans_adi"].nunique()
        min_not = ders_veri["gecme_notu"].min()
        max_not = ders_veri["gecme_notu"].max()

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Kapsadığı Lisans Sayısı", f"{filtreli_lisans_sayisi} / {toplam_lisans_sayisi}")
        m2.metric("En Düşük Geçme Barajı", f"{min_not}")
        m3.metric("En Yüksek Geçme Barajı", f"{max_not}")
        m4.metric("Yaygın Veriliş Şekli", ders_veri["verilis_sekli"].mode()[0] if not ders_veri.empty else "-")

        st.divider()

        dersi_isteyenler_df = filitreli_ders_veri.sort_values(by="gecme_notu", ascending=False)
        top_n_ders_lisans = dersi_isteyenler_df.head(n_yakin)

        if not top_n_ders_lisans.empty:
            st.write(f"##### 🏆 '{secilen_ders}' Dersini ({not_filtre_yonu} {esik_not if not_filtre_yonu != 'Tümü' else ''}) Şartıyla İsteyen İlk {len(top_n_ders_lisans)} Lisans")

            col_chart, col_table = st.columns([3, 2])
            with col_chart:
                dynamic_course_height = max(380, len(top_n_ders_lisans) * 35)
                fig_ders = px.bar(
                    top_n_ders_lisans,
                    x="gecme_notu",
                    y="lisans_adi",
                    orientation="h",
                    text="gecme_notu",
                    color="gecme_notu",
                    color_continuous_scale="Reds",
                    title=f"Not Barajı Sıralaması (İlk {len(top_n_ders_lisans)})"
                )
                fig_ders.update_layout(
                    height=dynamic_course_height,
                    yaxis={"categoryorder": "total ascending"},
                    margin=dict(l=10, r=10, t=40, b=10)
                )
                st.plotly_chart(fig_ders, use_container_width=True)

            with col_table:
                st.dataframe(
                    top_n_ders_lisans[["lisans_adi", "gecme_notu", "egitim_cinsi", "verilis_sekli", "sinav_tipi"]],
                    use_container_width=True,
                    hide_index=True,
                    height=dynamic_course_height
                )
        else:
            st.warning(f"Belirlediğiniz not filtresine ({not_filtre_yonu} {esik_not}) uyan lisans bulunamadı.")


# --- TAB 3: SERTİFİKA KALDIRAÇ ANALİZİ (GENEL) ---
with tab3:
    st.subheader("💡 En Çok Lisansa Dönüşen (Ortak) Dersler ve Sertifikalar")

    ders_lisans_sayisi = df_raw.groupby("ders_kodu_ve_adi")["lisans_adi"].nunique().reset_index()
    ders_lisans_sayisi.columns = ["Ders Kodu ve Adı", "Kapsadığı Lisans Sayısı"]

    not_stats = df_raw.groupby("ders_kodu_ve_adi")["gecme_notu"].agg(["min", "max"]).reset_index()
    not_stats.columns = ["Ders Kodu ve Adı", "Min Geçme Notu", "Maks Geçme Notu"]

    kaldirac_df = pd.merge(ders_lisans_sayisi, not_stats, on="Ders Kodu ve Adı").sort_values(
        by="Kapsadığı Lisans Sayısı", ascending=False
    )

    top_n_kaldirac = kaldirac_df.head(max(10, n_yakin))
    dynamic_kaldirac_height = max(400, len(top_n_kaldirac) * 35)

    col_a, col_b = st.columns([3, 2])
    with col_a:
        fig_kaldirac = px.bar(
            top_n_kaldirac,
            x="Kapsadığı Lisans Sayısı",
            y="Ders Kodu ve Adı",
            orientation="h",
            color="Kapsadığı Lisans Sayısı",
            color_continuous_scale="Blues",
            title=f"En Çok Lisans Açan İlk {len(top_n_kaldirac)} Ders"
        )
        fig_kaldirac.update_layout(
            height=dynamic_kaldirac_height,
            yaxis={"categoryorder": "total ascending"},
            margin=dict(l=10, r=10, t=40, b=10)
        )
        st.plotly_chart(fig_kaldirac, use_container_width=True)

    with col_b:
        st.write("##### Sertifika / Ders Kapsama Tablosu")
        st.dataframe(
            kaldirac_df,
            hide_index=True,
            use_container_width=True,
            height=dynamic_kaldirac_height
        )


# --- TAB 4: LİSANSLAR ARASI İLİŞKİ AĞI (NETWORK GRAPH) ---
with tab4:
    st.subheader("🕸️ Lisanslar Arası Etkileşimli İlişki Ağı")
    st.caption("Ağ haritasında bir Ana Lisans etrafındaki muafiyet bağlarını iki farklı modda analiz edin.")

    ag_modu = st.radio(
        "Ağ Analiz Modunu Seçin:",
        options=[
            "🎯 Ana Lisans vs. Seçilen Lisanslar (Manuel Karşılaştırma)",
            "⚡ Ana Lisans vs. En Yakın N Lisans (Otomatik Uyum Köprüsü)"
        ],
        horizontal=True
    )

    col_net_ctrl1, col_net_ctrl2 = st.columns([1, 1])

    with col_net_ctrl1:
        min_ortak_ders = st.slider(
            "Eşik Ortak Ders Sayısı (Min Bağlantı Gücü):",
            min_value=1,
            max_value=15,
            value=2
        )

    if ag_modu.startswith("🎯"):
        with col_net_ctrl2:
            mukayese_lisanslari = st.multiselect(
                "Mukayese Edilecek Diğer Lisansları Seçin:",
                options=[l for l in tum_lisanslar if l != secilen_lisans],
                default=[l for l in tum_lisanslar if l != secilen_lisans][:min(4, len(tum_lisanslar)-1)]
            )
            hedef_lisans_kumesi = [secilen_lisans] + mukayese_lisanslari
    else:
        with col_net_ctrl2:
            n_ag_yakin = st.slider(
                "En Yakın Kaç Lisans Ağ Haritasına Eklensin (N):",
                min_value=1,
                max_value=15,
                value=5
            )
            en_yakin_lisanslar = df_yakinlik.head(n_ag_yakin)["Lisans Adı"].tolist()
            hedef_lisans_kumesi = [secilen_lisans] + en_yakin_lisanslar

    G = nx.Graph()

    for l in hedef_lisans_kumesi:
        c_count = len(df_raw[df_raw["lisans_adi"] == l])
        G.add_node(l, total_courses=c_count)

    for i in range(len(hedef_lisans_kumesi)):
        l1 = hedef_lisans_kumesi[i]
        s1 = set(df_raw[df_raw["lisans_adi"] == l1]["ders_kodu_ve_adi"])
        for j in range(i + 1, len(hedef_lisans_kumesi)):
            l2 = hedef_lisans_kumesi[j]
            s2 = set(df_raw[df_raw["lisans_adi"] == l2]["ders_kodu_ve_adi"])
            ortak_dersler = s1.intersection(s2)

            if len(ortak_dersler) >= min_ortak_ders:
                G.add_edge(l1, l2, weight=len(ortak_dersler), shared_list=list(ortak_dersler))

    if G.number_of_nodes() > 0:
        pos = nx.spring_layout(G, k=0.7, seed=42)

        edge_x, edge_y = [], []
        for edge in G.edges():
            x0, y0 = pos[edge[0]]
            x1, y1 = pos[edge[1]]
            edge_x.extend([x0, x1, None])
            edge_y.extend([y0, y1, None])

        edge_trace = go.Scatter(
            x=edge_x, y=edge_y,
            line=dict(width=2, color='#B0BEC5'),
            hoverinfo='none',
            mode='lines'
        )

        node_x, node_y, node_text, node_size, node_color = [], [], [], [], []

        for node in G.nodes():
            x, y = pos[node]
            node_x.append(x)
            node_y.append(y)
            deg = G.degree(node)
            tot_courses = G.nodes[node]['total_courses']

            node_text.append(f"<b>{node}</b><br>• Toplam Ders: {tot_courses}<br>• Bağlantı Sayısı: {deg}")

            if node == secilen_lisans:
                node_size.append(34)
                node_color.append("#D32F2F")
            else:
                node_size.append(18 + deg * 2)
                node_color.append(deg)

        node_trace = go.Scatter(
            x=node_x, y=node_y,
            mode='markers+text',
            hoverinfo='text',
            text=[f"⭐ {n}" if n == secilen_lisans else (n if len(n) <= 22 else n[:20]+"...") for n in G.nodes()],
            textposition="top center",
            hovertext=node_text,
            marker=dict(
                showscale=True,
                colorscale='YlGnBu',
                size=node_size,
                color=node_color,
                colorbar=dict(
                    thickness=15,
                    title=dict(text='Bağlantı Sayısı', side='right'),
                    xanchor='left'
                ),
                line=dict(width=2, color='#37474F')
            )
        )

        fig_network = go.Figure(
            data=[edge_trace, node_trace],
            layout=go.Layout(
                title=dict(
                    text=f"<b>Ana Lisans: {secilen_lisans} — (Eşik: Min {min_ortak_ders} Ortak Ders)</b>",
                    font=dict(size=15)
                ),
                showlegend=False,
                hovermode='closest',
                margin=dict(b=20, l=5, r=5, t=40),
                height=650,
                xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
                yaxis=dict(showgrid=False, zeroline=False, showticklabels=False)
            )
        )

        st.plotly_chart(fig_network, use_container_width=True)
    else:
        st.warning("Ağ oluşturmak için seçilen parametrelere uygun lisans bulunamadı.")
