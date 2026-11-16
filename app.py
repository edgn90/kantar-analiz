import streamlit as st
import pandas as pd
import io
import os
import base64
import zipfile
import urllib.request
import ssl
import re
import plotly.express as px
from PIL import Image
from xhtml2pdf import pisa
from xhtml2pdf.default import DEFAULT_FONT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams["font.sans-serif"] = ["DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False
    HAS_MATPLOTLIB = True
except Exception:
    HAS_MATPLOTLIB = False

# Sayfa Ayarları
st.set_page_config(page_title="KAN-TAR Analiz Sistemi", layout="wide")


def _is_valid_ttf(path):
    return os.path.exists(path) and os.path.getsize(path) > 100_000


def _download_font(url, dest_path):
    try:
        ssl._create_default_https_context = ssl._create_unverified_context
    except Exception:
        pass
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as resp, open(dest_path, "wb") as out:
            out.write(resp.read())
        return _is_valid_ttf(dest_path)
    except Exception:
        if os.path.exists(dest_path):
            try:
                os.remove(dest_path)
            except Exception:
                pass
        return False


def get_turkish_font_paths():
    """
    Türkçe karakter destekli DejaVuSans ve DejaVuSans-Bold fontlarının
    yerel tam yollarını bulur veya güvenilir CDN'den indirir.
    """
    candidate_pairs = [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        (os.path.abspath("DejaVuSans.ttf"), os.path.abspath("DejaVuSans-Bold.ttf")),
    ]

    try:
        mpl_ttf_dir = os.path.join(os.path.dirname(matplotlib.__file__), "mpl-data", "fonts", "ttf")
        candidate_pairs.insert(0, (
            os.path.join(mpl_ttf_dir, "DejaVuSans.ttf"),
            os.path.join(mpl_ttf_dir, "DejaVuSans-Bold.ttf"),
        ))
    except Exception:
        pass

    for reg_path, bold_path in candidate_pairs:
        if _is_valid_ttf(reg_path) and _is_valid_ttf(bold_path):
            return os.path.abspath(reg_path), os.path.abspath(bold_path)

    local_reg = os.path.abspath("DejaVuSans.ttf")
    local_bold = os.path.abspath("DejaVuSans-Bold.ttf")

    reg_ok = _is_valid_ttf(local_reg) or _download_font(
        "https://cdn.jsdelivr.net/npm/dejavu-fonts-ttf@2.37.3/ttf/DejaVuSans.ttf",
        local_reg,
    )
    bold_ok = _is_valid_ttf(local_bold) or _download_font(
        "https://cdn.jsdelivr.net/npm/dejavu-fonts-ttf@2.37.3/ttf/DejaVuSans-Bold.ttf",
        local_bold,
    )

    if reg_ok and bold_ok:
        return local_reg, local_bold
    return None, None


def setup_fonts():
    """
    Fontları hem ReportLab çekirdeğine hem de xhtml2pdf DEFAULT_FONT haritasına kaydeder
    ve HTML şablonu için @font-face CSS bloğunu döndürür.
    """
    reg_path, bold_path = get_turkish_font_paths()
    if not reg_path or not bold_path:
        return False, ""

    try:
        if "DejaVuSans" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("DejaVuSans", reg_path))
        if "DejaVuSans-Bold" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", bold_path))
        registerFontFamily("DejaVuSans", normal="DejaVuSans", bold="DejaVuSans-Bold")

        DEFAULT_FONT["helvetica"] = "DejaVuSans"
        DEFAULT_FONT["sans-serif"] = "DejaVuSans"

        font_face_css = f"""
        @font-face {{
            font-family: "DejaVuSans";
            src: url("{reg_path}");
            font-weight: normal;
            font-style: normal;
        }}
        @font-face {{
            font-family: "DejaVuSans";
            src: url("{bold_path}");
            font-weight: bold;
            font-style: normal;
        }}
        """
        return True, font_face_css
    except Exception:
        return False, ""


def _image_bytes_to_png_base64(raw_bytes):
    """
    Görseli xhtml2pdf ile %100 uyumlu PNG formatına dönüştürüp base64 string olarak döndürür.
    """
    try:
        img = Image.open(io.BytesIO(raw_bytes))
        if img.mode in ("RGBA", "LA", "P"):
            img = img.convert("RGBA")
            bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
            img = Image.alpha_composite(bg, img).convert("RGB")
        elif img.mode != "RGB":
            img = img.convert("RGB")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception:
        return None


def get_kantar_logo_base64(uploaded_logo=None):
    """
    Önce kenar çubuğundan yüklenen logoyu, yoksa proje klasöründeki logo dosyasını okur.
    """
    if uploaded_logo is not None:
        try:
            b64 = _image_bytes_to_png_base64(uploaded_logo.getvalue())
            if b64:
                return b64
        except Exception:
            pass

    candidate_files = [
        "kantar_logo.png",
        "kantar_logo.jpg",
        "kantar_logo.jpeg",
        "logo.png",
        "logo.jpg",
        "ChatGPT Görseli 16 Kas 2026 16_48_14.png",
    ]
    for fname in candidate_files:
        if os.path.exists(fname):
            try:
                with open(fname, "rb") as f:
                    b64 = _image_bytes_to_png_base64(f.read())
                if b64:
                    return b64
            except Exception:
                continue
    return None


@st.cache_data
def load_asm_mapping():
    try:
        df_map = pd.concat(pd.read_excel("ASM_AHB_Eslestirme.xlsx", sheet_name=None).values(), ignore_index=True)
        mapping = dict(zip(df_map['AHB BİRİMİ'].astype(str).str.strip().str.lower(), df_map['ASM'].astype(str).str.strip()))
        return mapping
    except Exception:
        return None


@st.cache_data
def load_data(files, asm_mapping):
    tum_veriler = []
    
    for file in files:
        dosya_adi = file.name
        xls = pd.ExcelFile(file)
        
        if 'Birim Bazında' in xls.sheet_names:
            df_test = pd.read_excel(xls, sheet_name='Birim Bazında', nrows=5)
            if any('kronik hastalik' in str(c).lower().replace('ı', 'i') for c in df_test.columns):
                df_raw = pd.read_excel(xls, sheet_name='Birim Bazında')
                df_raw.columns = [str(c).strip().lower().replace('ı', 'i') for c in df_raw.columns]
                
                df_pivot = df_raw.pivot_table(
                    index=['birim adi'], 
                    columns='kronik hastalik', 
                    values=['aylik gereken', 'yapilan'],
                    aggfunc='sum'
                ).reset_index()
                
                df_pivot.columns = ['_'.join(col).strip('_') for col in df_pivot.columns.values]
                
                temp_df = pd.DataFrame()
                temp_df['Birim Adı'] = df_pivot['birim adi']
                temp_df['Birim ID'] = temp_df['Birim Adı'].str.extract(r'(\d+)\s+Nolu').fillna("Bilinmiyor")
                temp_df['Dosya (Ay)'] = dosya_adi
                
                def get_col(df, keyword, val_type):
                    for c in df.columns:
                        if val_type in c.lower() and keyword.lower() in c.lower():
                            return pd.to_numeric(df[c], errors='coerce').fillna(0)
                    return pd.Series([0]*len(df))
                
                temp_df['Mamografi Gereken'] = get_col(df_pivot, 'mamograf', 'gereken')
                temp_df['Mamografi Yapılan'] = get_col(df_pivot, 'mamograf', 'yapilan')
                temp_df['Serviks Gereken'] = get_col(df_pivot, 'serv', 'gereken')
                temp_df['Serviks Yapılan'] = get_col(df_pivot, 'serv', 'yapilan')
                temp_df['Kolorektal Gereken'] = get_col(df_pivot, 'kolorek', 'gereken')
                temp_df['Kolorektal Yapılan'] = get_col(df_pivot, 'kolorek', 'yapilan')
                
                tum_veriler.append(temp_df)
                continue
                
        df_temp = pd.read_excel(file, header=None, nrows=15)
        header_idx = 1 
        for i, row in df_temp.iterrows():
            row_str = " ".join([str(h).lower() for h in row.tolist()])
            if "birim" in row_str and "oran" in row_str:
                header_idx = i
                break
                
        df_old = pd.read_excel(file, header=header_idx)
        df_old.columns = [str(c).strip() for c in df_old.columns]
        
        sutun_degisiklikleri = {}
        for col in df_old.columns:
            if col.lower() == "birim adi": sutun_degisiklikleri[col] = "Birim Adı"
            elif col.lower() == "birim id": sutun_degisiklikleri[col] = "Birim ID"
        if sutun_degisiklikleri:
            df_old = df_old.rename(columns=sutun_degisiklikleri)
            
        df_old = df_old.rename(columns={
            'Birim Id': 'Birim ID',
            'Gereken Tarama': 'Mamografi Gereken', 'Oran(%)': 'Mamografi Oranı',
            'Gereken Tarama.1': 'Serviks Gereken', 'Oran(%).1': 'Serviks Oranı',
            'Gereken Tarama.2': 'Kolorektal Gereken', 'Oran(%).2': 'Kolorektal Oranı'
        })
        
        for t in ['Mamografi', 'Serviks', 'Kolorektal']:
            gereken = pd.to_numeric(df_old[f'{t} Gereken'], errors='coerce').fillna(0)
            oran = pd.to_numeric(df_old[f'{t} Oranı'], errors='coerce').fillna(0)
            df_old[f'{t} Gereken'] = gereken
            df_old[f'{t} Yapılan'] = (oran * gereken / 100).round()
            
        df_old_final = df_old[['Birim Adı', 'Birim ID', 'Mamografi Gereken', 'Mamografi Yapılan', 'Serviks Gereken', 'Serviks Yapılan', 'Kolorektal Gereken', 'Kolorektal Yapılan']].copy()
        df_old_final['Dosya (Ay)'] = dosya_adi
        tum_veriler.append(df_old_final)

    combined_df = pd.concat(tum_veriler, ignore_index=True)
    
    if asm_mapping:
        combined_df['ASM Adı'] = combined_df['Birim Adı'].astype(str).str.strip().str.lower().map(asm_mapping).fillna("Bilinmeyen/Eksik ASM")
    else:
        combined_df['ASM Adı'] = "Sistemde ASM_AHB_Eslestirme.xlsx bulunamadı"

    districts = [
        "Adalar", "Arnavutköy", "Ataşehir", "Avcılar", "Bağcılar", "Bahçelievler", 
        "Bakırköy", "Başakşehir", "Bayrampaşa", "Beşiktaş", "Beykoz", "Beylikdüzü", 
        "Beyoğlu", "Büyükçekmece", "Çatalca", "Çekmeköy", "Esenler", "Esenyurt", 
        "Eyüpsultan", "Fatih", "Gaziosmanpaşa", "Güngören", "Kadıköy", "Kağıthane", 
        "Kartal", "Küçükçekmece", "Maltepe", "Pendik", "Sancaktepe", "Sarıyer", 
        "Silivri", "Sultanbeyli", "Sultangazi", "Şile", "Şişli", "Tuzla", "Ümraniye", 
        "Üsküdar", "Zeytinburnu"
    ]
    
    def get_district(birim_adi):
        birim_adi_upper = str(birim_adi).upper().replace('İ', 'I')
        for d in districts:
            if d.upper().replace('İ', 'I') in birim_adi_upper:
                return d
        if "EYÜP" in birim_adi_upper or "EYUP" in birim_adi_upper:
            return "Eyüpsultan"
        return "Diğer/Kurum Tabiplikleri"

    combined_df['İlçe'] = combined_df['Birim Adı'].apply(get_district)
    return combined_df


@st.cache_data
def get_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sıralama')
    return output.getvalue()


def get_pdf(df, title):
    fonts_loaded, font_face_css = setup_fonts()
    font_family = '"DejaVuSans", sans-serif' if fonts_loaded else "Helvetica, Arial, sans-serif"
    
    html_content = f"""
    <!DOCTYPE html>
    <html lang="tr">
    <head>
    <meta charset="UTF-8">
    <style>
        {font_face_css}
        @page {{ size: A4; margin: 1cm; }}
        body {{ font-family: {font_family}; font-size: 10px; color: #333; }}
        h2 {{ text-align: center; color: #333; font-family: {font_family}; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 5px; text-align: left; font-family: {font_family}; }}
        th {{ background-color: #f4f6f9; font-weight: bold; }}
    </style>
    </head>
    <body>
        <h2>{title}</h2>
        {df.to_html(index=False)}
    </body>
    </html>
    """
    pdf_buffer = io.BytesIO()
    pisa.CreatePDF(io.BytesIO(html_content.encode('utf-8')), dest=pdf_buffer, encoding='utf-8')
    return pdf_buffer.getvalue()


def build_ilce_asm_ranking_df(ilce_detay, tarama_turu):
    gereken_col = f"{tarama_turu} Gereken"
    yapilan_col = f"{tarama_turu} Yapılan"
    asm_df = ilce_detay.groupby('ASM Adı', as_index=False)[[gereken_col, yapilan_col]].sum()
    asm_df['Oran (%)'] = (asm_df[yapilan_col] / asm_df[gereken_col] * 100).fillna(0)
    asm_df = asm_df.sort_values(by=['Oran (%)', 'ASM Adı'], ascending=[False, True]).reset_index(drop=True)
    asm_df['İlçe İçi Sıra'] = asm_df.index + 1
    asm_df['Sıralı ASM Etiketi'] = asm_df.apply(lambda r: f"{int(r['İlçe İçi Sıra'])}. {r['ASM Adı']}", axis=1)
    return asm_df


def create_ilce_asm_chart_base64(ilce_detay, secilen_ilce, tarama_turu):
    if not HAS_MATPLOTLIB:
        return None
    asm_df = build_ilce_asm_ranking_df(ilce_detay, tarama_turu)
    if asm_df.empty:
        return None

    plot_df = asm_df.iloc[::-1]
    fig_height = min(10.5, max(2.4, len(plot_df) * 0.34))
    fig, ax = plt.subplots(figsize=(8.2, fig_height), dpi=150)

    bar_colors = ['#2980b9' if val >= 70 else ('#f39c12' if val >= 50 else '#c0392b') for val in plot_df['Oran (%)']]
    bars = ax.barh(plot_df['Sıralı ASM Etiketi'], plot_df['Oran (%)'], color=bar_colors, height=0.62)

    ax.axvline(70, color='#d63031', linestyle='--', linewidth=1.4, label='Bakanlık %70 Alt Sınır Hedefi')
    ax.set_xlim(0, max(105, plot_df['Oran (%)'].max() + 14))
    ax.set_xlabel('Başarı Oranı (%)', fontsize=8.5)
    ax.set_title(f"{secilen_ilce} İlçesi - {tarama_turu} Taraması ASM İlçe İçi Sıralaması ve Oranları", fontsize=10, fontweight='bold', pad=8)
    ax.tick_params(axis='both', labelsize=8)
    ax.grid(axis='x', linestyle=':', alpha=0.5)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    for bar, val in zip(bars, plot_df['Oran (%)']):
        ax.text(
            val + 1.2,
            bar.get_y() + bar.get_height() / 2,
            f"%{val:.2f}",
            va='center',
            ha='left',
            fontsize=8,
            fontweight='bold',
            color='#2c3e50'
        )

    ax.legend(loc='lower right', fontsize=8, frameon=True)
    plt.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight')
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode('utf-8')


def get_ilce_rapor_pdf(df, secilen_ilce, secilen_aylar, logo_b64=None):
    fonts_loaded, font_face_css = setup_fonts()
    font_family = '"DejaVuSans", sans-serif' if fonts_loaded else "Helvetica, Arial, sans-serif"
    
    def get_month_num(m):
        match = re.search(r'^(\d+)', str(m))
        return int(match.group(1)) if match else 99
    
    sirali_aylar = sorted(secilen_aylar, key=get_month_num)
    aylar_str = ", ".join(sirali_aylar)
    
    ilce_grup = df.groupby('İlçe')[['Mamografi Gereken', 'Mamografi Yapılan', 'Serviks Gereken', 'Serviks Yapılan', 'Kolorektal Gereken', 'Kolorektal Yapılan']].sum().reset_index()
    for t in ['Mamografi', 'Serviks', 'Kolorektal']:
        ilce_grup[f'{t} Oranı'] = (ilce_grup[f'{t} Yapılan'] / ilce_grup[f'{t} Gereken'] * 100).fillna(0)
        ilce_grup[f'{t} Sıra'] = ilce_grup[f'{t} Oranı'].rank(ascending=False, method='min').astype(int)
        
    ilce_ozet = ilce_grup[ilce_grup['İlçe'] == secilen_ilce].iloc[0]
    ilce_detay = df[df['İlçe'] == secilen_ilce]
    
    def format_oran(oran):
        if pd.isna(oran):
            oran = 0
        if oran < 50.0:
            return f"<span style='color: red; font-weight: bold;'>%{oran:.2f} (Uyarı)</span>"
        return f"%{oran:.2f}"

    if logo_b64:
        ust_baslik_html = f"""
        <table class="top-header-table">
            <tr>
                <td width="18%" style="border: none;"></td>
                <td width="64%" style="border: none; text-align: center; vertical-align: middle;">
                    <h1 style="margin: 0 0 4px 0;">{secilen_ilce} İlçesi KAN-TAR Değerlendirme Raporu</h1>
                    <div class="header-info" style="margin: 0;"><b>Kapsanan Aylar:</b> {aylar_str}</div>
                </td>
                <td width="18%" style="border: none; text-align: right; vertical-align: top; padding: 0;">
                    <img src="data:image/png;base64,{logo_b64}" width="105" />
                </td>
            </tr>
        </table>
        """
    else:
        ust_baslik_html = f"""
        <h1>{secilen_ilce} İlçesi KAN-TAR Değerlendirme Raporu</h1>
        <div class="header-info"><b>Kapsanan Aylar:</b> {aylar_str}</div>
        """

    html_content = f"""
    <!DOCTYPE html>
    <html lang="tr">
    <head>
    <meta charset="UTF-8">
    <style>
        {font_face_css}
        @page {{ size: A4; margin: 1cm; }}
        body {{ font-family: {font_family}; font-size: 10px; color: #333; }}
        .top-header-table {{ width: 100%; border-collapse: collapse; margin-bottom: 12px; }}
        h1 {{ text-align: center; color: #2c3e50; font-size: 16px; margin-bottom: 5px; font-family: {font_family}; }}
        .header-info {{ text-align: center; font-size: 11px; margin-bottom: 20px; color: #7f8c8d; font-family: {font_family}; }}
        h2 {{ color: #34495e; font-size: 13px; border-bottom: 1px solid #ccc; padding-bottom: 3px; margin-top: 15px; font-family: {font_family}; }}
        h3 {{ color: #16a085; font-size: 11px; margin-top: 10px; margin-bottom: 5px; font-family: {font_family}; }}
        table {{ border-collapse: collapse; width: 100%; margin-bottom: 5px; }}
        th, td {{ border: 1px solid #bdc3c7; padding: 3px; font-size: 8.5px; font-family: {font_family}; }}
        th {{ background-color: #ecf0f1; font-weight: bold; text-align: center; }}
        td {{ text-align: center; }}
        .left-align {{ text-align: left; }}
        .summary-table th {{ background-color: #34495e; color: white; }}
        .summary-table td {{ font-size: 11px; padding: 8px; }}
        .chart-box {{ text-align: center; margin-top: 8px; margin-bottom: 12px; }}
    </style>
    </head>
    <body>
        {ust_baslik_html}
        
        <h2>1. İlçe Genel Durumu ve İstanbul Sıralaması</h2>
        <table class="summary-table">
            <tr>
                <th>Mamografi Taraması</th>
                <th>Serviks Taraması</th>
                <th>Kolorektal Taraması</th>
            </tr>
            <tr>
                <td>{format_oran(ilce_ozet['Mamografi Oranı'])} <br><span style='font-size: 9px; color: #555;'>İstanbul Sırası: {ilce_ozet['Mamografi Sıra']}.</span></td>
                <td>{format_oran(ilce_ozet['Serviks Oranı'])} <br><span style='font-size: 9px; color: #555;'>İstanbul Sırası: {ilce_ozet['Serviks Sıra']}.</span></td>
                <td>{format_oran(ilce_ozet['Kolorektal Oranı'])} <br><span style='font-size: 9px; color: #555;'>İstanbul Sırası: {ilce_ozet['Kolorektal Sıra']}.</span></td>
            </tr>
        </table>
        
        <h2>2. ASM İlçe İçi Sıralaması ve Başarı Oranları Grafikleri</h2>
    """

    for tarama_turu in ['Mamografi', 'Serviks', 'Kolorektal']:
        chart_b64 = create_ilce_asm_chart_base64(ilce_detay, secilen_ilce, tarama_turu)
        if chart_b64:
            html_content += f"""
            <div class="chart-box">
                <img src="data:image/png;base64,{chart_b64}" width="530" />
            </div>
            """

    html_content += """
        <h2>3. ASM ve Birim (AHB) Bazlı Kırılımlar</h2>
    """
    
    asm_grup = ilce_detay.groupby('ASM Adı')[['Mamografi Gereken', 'Mamografi Yapılan', 'Serviks Gereken', 'Serviks Yapılan', 'Kolorektal Gereken', 'Kolorektal Yapılan']].sum().reset_index()
    asm_grup = asm_grup.sort_values(by='ASM Adı')
    
    for _, asm_row in asm_grup.iterrows():
        asm_adi = asm_row['ASM Adı']
        asm_m_ger = int(round(asm_row['Mamografi Gereken']))
        asm_m_yap = int(round(asm_row['Mamografi Yapılan']))
        asm_s_ger = int(round(asm_row['Serviks Gereken']))
        asm_s_yap = int(round(asm_row['Serviks Yapılan']))
        asm_k_ger = int(round(asm_row['Kolorektal Gereken']))
        asm_k_yap = int(round(asm_row['Kolorektal Yapılan']))

        m_oran = (asm_row['Mamografi Yapılan'] / asm_row['Mamografi Gereken'] * 100) if asm_row['Mamografi Gereken'] > 0 else 0
        s_oran = (asm_row['Serviks Yapılan'] / asm_row['Serviks Gereken'] * 100) if asm_row['Serviks Gereken'] > 0 else 0
        k_oran = (asm_row['Kolorektal Yapılan'] / asm_row['Kolorektal Gereken'] * 100) if asm_row['Kolorektal Gereken'] > 0 else 0
        
        html_content += f"<h3>[ASM] {asm_adi}</h3>"
        html_content += f"""
        <table style="background-color: #f8f9fa;">
            <tr>
                <td width="33%">
                    <b>ASM Ort. Mamografi:</b> {format_oran(m_oran)}<br>
                    <span style="font-size: 8px; color: #555;">Gereken: <b>{asm_m_ger}</b> | Yapılan: <b>{asm_m_yap}</b></span>
                </td>
                <td width="33%">
                    <b>ASM Ort. Serviks:</b> {format_oran(s_oran)}<br>
                    <span style="font-size: 8px; color: #555;">Gereken: <b>{asm_s_ger}</b> | Yapılan: <b>{asm_s_yap}</b></span>
                </td>
                <td width="33%">
                    <b>ASM Ort. Kolorektal:</b> {format_oran(k_oran)}<br>
                    <span style="font-size: 8px; color: #555;">Gereken: <b>{asm_k_ger}</b> | Yapılan: <b>{asm_k_yap}</b></span>
                </td>
            </tr>
        </table>
        <table>
            <tr>
                <th class="left-align" width="28%">Birim Adı</th>
                <th colspan="3" width="24%">Mamografi</th>
                <th colspan="3" width="24%">Serviks</th>
                <th colspan="3" width="24%">Kolorektal</th>
            </tr>
            <tr>
                <th class="left-align" width="28%"></th>
                <th width="7%">Gereken</th>
                <th width="7%">Yapılan</th>
                <th width="10%">Oran (%)</th>
                <th width="7%">Gereken</th>
                <th width="7%">Yapılan</th>
                <th width="10%">Oran (%)</th>
                <th width="7%">Gereken</th>
                <th width="7%">Yapılan</th>
                <th width="10%">Oran (%)</th>
            </tr>
        """
        
        asm_birimler = ilce_detay[ilce_detay['ASM Adı'] == asm_adi].sort_values(by='Birim Adı')
        for _, birim_row in asm_birimler.iterrows():
            m_ger = int(round(birim_row['Mamografi Gereken']))
            m_yap = int(round(birim_row['Mamografi Yapılan']))
            s_ger = int(round(birim_row['Serviks Gereken']))
            s_yap = int(round(birim_row['Serviks Yapılan']))
            k_ger = int(round(birim_row['Kolorektal Gereken']))
            k_yap = int(round(birim_row['Kolorektal Yapılan']))

            html_content += "<tr>"
            html_content += f"<td class='left-align'>{str(birim_row['Birim Adı'])}</td>"
            html_content += f"<td>{m_ger}</td><td>{m_yap}</td><td>{format_oran(birim_row['Mamografi Oranı'])}</td>"
            html_content += f"<td>{s_ger}</td><td>{s_yap}</td><td>{format_oran(birim_row['Serviks Oranı'])}</td>"
            html_content += f"<td>{k_ger}</td><td>{k_yap}</td><td>{format_oran(birim_row['Kolorektal Oranı'])}</td>"
            html_content += "</tr>"
        html_content += "</table>"
        
    html_content += "</body></html>"
    
    pdf_buffer = io.BytesIO()
    pisa.CreatePDF(io.BytesIO(html_content.encode('utf-8')), dest=pdf_buffer, encoding='utf-8')
    return pdf_buffer.getvalue()


st.title("📊 KAN-TAR Analiz Sistemi")

asm_mapping = load_asm_mapping()

st.sidebar.header("1. Veri Yükleme")
yuklenen_dosyalar = st.sidebar.file_uploader("KAN-TAR Excel Dosyalarını Yükleyin", type=["xlsx", "xls"], accept_multiple_files=True)
yuklenen_logo = st.sidebar.file_uploader("İlçe Raporu Logosu (Opsiyonel)", type=["png", "jpg", "jpeg"])
kantar_logo_b64 = get_kantar_logo_base64(yuklenen_logo)

if yuklenen_dosyalar:
    try:
        df_ham = load_data(yuklenen_dosyalar, asm_mapping)
        
        st.sidebar.header("2. Filtreleme Seçenekleri")
        
        with st.sidebar.form(key="filtre_form"):
            ay_secenekleri = df_ham['Dosya (Ay)'].unique().tolist()
            secilen_aylar = st.multiselect("Analiz Edilecek Ayları Seçin", options=ay_secenekleri, default=ay_secenekleri)
            
            st.markdown("---")
            
            ilce_secenekleri = ["Tüm İlçeler"] + sorted(df_ham['İlçe'].unique().tolist())
            secilen_ilce = st.selectbox("İlçe Seçin", ilce_secenekleri)
            
            st.markdown("---")
            
            st.subheader("3. Oran Filtresi (ASM ve Birimler İçin)")
            min_oran, max_oran = st.slider(
                "Gösterilecek Başarı Oranı Aralığı (%)",
                min_value=0, max_value=100, value=(0, 100), step=1
            )
            
            calistir_butonu = st.form_submit_button(label="Çalıştır 🚀")

        if secilen_aylar:
            df_filtrelenmis_ay = df_ham[df_ham['Dosya (Ay)'].isin(secilen_aylar)]
            
            agg_funcs = {
                'Birim ID': 'first',
                'İlçe': 'first',
                'ASM Adı': 'first',
                'Mamografi Gereken': 'sum', 'Mamografi Yapılan': 'sum',
                'Serviks Gereken': 'sum', 'Serviks Yapılan': 'sum',
                'Kolorektal Gereken': 'sum', 'Kolorektal Yapılan': 'sum',
            }
            df = df_filtrelenmis_ay.groupby('Birim Adı').agg(agg_funcs).reset_index()
            
            for t in ['Mamografi', 'Serviks', 'Kolorektal']:
                df[f'{t} Oranı'] = (df[f'{t} Yapılan'] / df[f'{t} Gereken'] * 100).fillna(0)

            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                "🎀 Mamografi Taraması", 
                "🎗️ Serviks Taraması", 
                "🧬 Kolorektal Taraması", 
                "📑 İlçe Bazlı Rapor", 
                "📈 Aylık Trend",
                "🚨 Acil Müdahale (Sıfır Tarama)"
            ])

            def tablo_goster(tarama_turu):
                if secilen_ilce != "Tüm İlçeler":
                    filtrelenmis_df = df[df['İlçe'] == secilen_ilce].copy()
                    ilce_baslik = f"{secilen_ilce} İlçesi"
                else:
                    filtrelenmis_df = df.copy()
                    ilce_baslik = "Tüm İstanbul Geneli"

                hedef_kolon = f"{tarama_turu} Oranı"
                gereken_kolon = f"{tarama_turu} Gereken"
                yapilan_kolon = f"{tarama_turu} Yapılan"
                
                # --- ÖZET METRİK KARTLARI ---
                toplam_gereken = filtrelenmis_df[gereken_kolon].sum()
                toplam_yapilan = filtrelenmis_df[yapilan_kolon].sum()
                genel_oran = (toplam_yapilan / toplam_gereken * 100) if toplam_gereken > 0 else 0

                st.markdown(f"#### 📌 {ilce_baslik} - {tarama_turu} Özet Metrikleri")
                m1, m2, m3 = st.columns(3)
                m1.metric("Toplam Gereken", f"{toplam_gereken:,.0f}")
                m2.metric("Toplam Yapılan", f"{toplam_yapilan:,.0f}")
                m3.metric("Genel Başarı Oranı", f"%{genel_oran:.2f}")

                st.markdown("---")
                
                st.subheader(f"📍 İlçe Başarı Sıralaması ({ilce_baslik})")
                ilce_grup = filtrelenmis_df.groupby('İlçe')[[gereken_kolon, yapilan_kolon]].sum().reset_index()
                ilce_grup['Genel Oran (%)'] = (ilce_grup[yapilan_kolon] / ilce_grup[gereken_kolon] * 100).fillna(0)
                ilce_grup_sirali = ilce_grup.sort_values(by='Genel Oran (%)', ascending=False).reset_index(drop=True)
                ilce_grup_sirali.index += 1
                
                if secilen_ilce == "Tüm İlçeler":
                    max_y_ilce = max(100, float(ilce_grup_sirali['Genel Oran (%)'].max()) * 1.12)
                    fig_ilce = px.bar(
                        ilce_grup_sirali, 
                        x='İlçe', 
                        y='Genel Oran (%)', 
                        text='Genel Oran (%)',
                        title=f"İlçelere Göre {tarama_turu} Başarı Oranları (Bakanlık %70 Alt Sınır Hedefi)",
                        color='Genel Oran (%)',
                        color_continuous_scale='Blues'
                    )
                    fig_ilce.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
                    fig_ilce.update_layout(xaxis_tickangle=-45, yaxis_range=[0, max_y_ilce])
                    fig_ilce.add_hline(
                        y=70,
                        line_dash="dash",
                        line_color="red",
                        line_width=2,
                        annotation_text="Bakanlık %70 Alt Sınır Hedefi",
                        annotation_position="top right"
                    )
                    st.plotly_chart(fig_ilce, use_container_width=True)

                ilce_grup_pdf = ilce_grup_sirali.copy()
                ilce_grup_pdf['Genel Oran (%)'] = ilce_grup_pdf['Genel Oran (%)'].apply(lambda x: f"%{x:.2f}")
                
                st.dataframe(
                    ilce_grup_sirali.style.format({'Genel Oran (%)': "{:.2f}%", gereken_kolon: "{:.0f}", yapilan_kolon: "{:.0f}"})
                    .background_gradient(subset=['Genel Oran (%)'], cmap='Blues'),
                    use_container_width=True
                )
                
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button("📥 Excel İndir (İlçe)", data=get_excel(ilce_grup_sirali), file_name=f"{tarama_turu}_Ilce_Siralama.xlsx", key=f"{tarama_turu}_ilce_ex")
                with col2:
                    st.download_button("📄 PDF İndir (İlçe)", data=get_pdf(ilce_grup_pdf, f"{ilce_baslik} - {tarama_turu} İlçe Sıralaması"), file_name=f"{tarama_turu}_Ilce_Siralama.pdf", mime="application/pdf", key=f"{tarama_turu}_ilce_pdf")
                
                st.markdown("---")

                st.subheader(f"🏥 ASM (Aile Sağlığı Merkezi) Başarı Sıralaması ({ilce_baslik} | %{min_oran} - %{max_oran} Arası)")
                asm_grup = filtrelenmis_df.groupby(['İlçe', 'ASM Adı'])[[gereken_kolon, yapilan_kolon]].sum().reset_index()
                asm_grup[hedef_kolon] = (asm_grup[yapilan_kolon] / asm_grup[gereken_kolon] * 100).fillna(0)
                
                asm_grup_aralikli = asm_grup[(asm_grup[hedef_kolon] >= min_oran) & (asm_grup[hedef_kolon] <= max_oran)]
                asm_grup_sirali = asm_grup_aralikli.sort_values(by=hedef_kolon, ascending=False).reset_index(drop=True)
                asm_grup_sirali.index += 1
                
                if not asm_grup_sirali.empty:
                    max_y_asm = max(100, float(asm_grup_sirali[hedef_kolon].max()) * 1.12)
                    fig_asm = px.bar(
                        asm_grup_sirali,
                        x='ASM Adı',
                        y=hedef_kolon,
                        text=hedef_kolon,
                        title=f"{ilce_baslik} - ASM Bazlı {tarama_turu} Başarı Oranları (Bakanlık %70 Alt Sınır Hedefi)",
                        color=hedef_kolon,
                        color_continuous_scale='Greens'
                    )
                    fig_asm.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
                    fig_asm.update_layout(xaxis_tickangle=-45, yaxis_range=[0, max_y_asm])
                    fig_asm.add_hline(
                        y=70,
                        line_dash="dash",
                        line_color="red",
                        line_width=2,
                        annotation_text="Bakanlık %70 Alt Sınır Hedefi",
                        annotation_position="top right"
                    )
                    st.plotly_chart(fig_asm, use_container_width=True)

                asm_grup_pdf = asm_grup_sirali.copy()
                asm_grup_pdf[hedef_kolon] = asm_grup_pdf[hedef_kolon].apply(lambda x: f"%{x:.2f}")
                
                st.dataframe(
                    asm_grup_sirali.style.format({hedef_kolon: "{:.2f}%", gereken_kolon: "{:.0f}", yapilan_kolon: "{:.0f}"})
                    .background_gradient(subset=[hedef_kolon], cmap='Blues'),
                    use_container_width=True
                )
                
                col3, col4 = st.columns(2)
                with col3:
                    st.download_button("📥 Excel İndir (ASM)", data=get_excel(asm_grup_sirali), file_name=f"{tarama_turu}_ASM_Siralama.xlsx", key=f"{tarama_turu}_asm_ex")
                with col4:
                    st.download_button("📄 PDF İndir (ASM)", data=get_pdf(asm_grup_pdf, f"{ilce_baslik} - {tarama_turu} ASM Sıralaması"), file_name=f"{tarama_turu}_ASM_Siralama.pdf", mime="application/pdf", key=f"{tarama_turu}_asm_pdf")

                st.markdown("---")

                # --- BİRİM BAZLI DETAYLI TABLO (EXPANDER İÇİNDE) ---
                with st.expander("Birim Bazlı Detaylı Tabloyu Gör"):
                    st.subheader(f"🩺 Birim (AHB) Bazlı Sıralama ({ilce_baslik} | %{min_oran} - %{max_oran} Arası)")
                    birim_df_aralikli = filtrelenmis_df[
                        (filtrelenmis_df[hedef_kolon] >= min_oran) & 
                        (filtrelenmis_df[hedef_kolon] <= max_oran)
                    ]
                    
                    birim_df_sirali = birim_df_aralikli.sort_values(by=hedef_kolon, ascending=False)
                    
                    gosterilecek_df = birim_df_sirali[[
                        'İlçe', 'ASM Adı', 'Birim Adı', 'Birim ID', gereken_kolon, yapilan_kolon, hedef_kolon
                    ]].reset_index(drop=True)
                    gosterilecek_df.index += 1
                    
                    gosterilecek_pdf = gosterilecek_df.copy()
                    gosterilecek_pdf[hedef_kolon] = gosterilecek_pdf[hedef_kolon].apply(lambda x: f"%{x:.2f}")
                    
                    st.dataframe(
                        gosterilecek_df.style.format({hedef_kolon: "{:.2f}%", gereken_kolon: "{:.0f}", yapilan_kolon: "{:.0f}"}), 
                        use_container_width=True
                    )
                    
                    col5, col6 = st.columns(2)
                    with col5:
                        st.download_button("📥 Excel İndir (Birim)", data=get_excel(gosterilecek_df), file_name=f"{tarama_turu}_Birim_Siralama.xlsx", key=f"{tarama_turu}_birim_ex")
                    with col6:
                        st.download_button("📄 PDF İndir (Birim)", data=get_pdf(gosterilecek_pdf, f"{ilce_baslik} - {tarama_turu} Birim Bazlı Sıralama"), file_name=f"{tarama_turu}_Birim_Siralama.pdf", mime="application/pdf", key=f"{tarama_turu}_birim_pdf")

            with tab1:
                tablo_goster("Mamografi")
            with tab2:
                tablo_goster("Serviks")
            with tab3:
                tablo_goster("Kolorektal")
                
            with tab4:
                def get_month_num(m):
                    match = re.search(r'^(\d+)', str(m))
                    return int(match.group(1)) if match else 99

                sirali_ay_isimleri = sorted(secilen_aylar, key=get_month_num)
                logo_hash = len(kantar_logo_b64) if kantar_logo_b64 else 0

                if secilen_ilce == "Tüm İlçeler":
                    st.subheader("📦 Tüm İlçeler İçin Toplu PDF Raporu (.ZIP)")
                    
                    mevcut_ilceler = sorted([
                        ilce for ilce in df['İlçe'].unique() 
                        if ilce != "Diğer/Kurum Tabiplikleri"
                    ])
                    
                    st.markdown(
                        f"Sol menüden belirli bir ilçe seçilmediği için seçili **{', '.join(sirali_ay_isimleri)}** dönemleri kapsamında "
                        f"veride yer alan **{len(mevcut_ilceler)} ilçenin** tamamına ait ayrı ayrı hazırlanmış Türkçe karakter uyumlu "
                        "PDF raporlarını tek bir `.zip` arşiv dosyası olarak oluşturup indirebilirsiniz."
                    )

                    zip_cache_key = f"toplu_zip_v5_{'_'.join(sirali_ay_isimleri)}_{len(mevcut_ilceler)}_{int(df['Mamografi Yapılan'].sum())}_{logo_hash}"

                    if st.button("📦 Tüm İlçelerin Raporlarını Toplu İndir (ZIP)", key="toplu_zip_olustur_btn"):
                        if not mevcut_ilceler:
                            st.warning("Rapor oluşturulacak geçerli ilçe verisi bulunamadı.")
                        else:
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            zip_buffer = io.BytesIO()
                            
                            with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                                toplam_ilce = len(mevcut_ilceler)
                                for idx, ilce_adi in enumerate(mevcut_ilceler, start=1):
                                    status_text.text(f"PDF hazırlanıyor ({idx}/{toplam_ilce}): {ilce_adi} İlçesi...")
                                    ilce_pdf_bytes = get_ilce_rapor_pdf(df, ilce_adi, secilen_aylar, logo_b64=kantar_logo_b64)
                                    zf.writestr(f"{ilce_adi}_Ilce_KANTAR_Raporu.pdf", ilce_pdf_bytes)
                                    progress_bar.progress(idx / toplam_ilce)
                            
                            status_text.empty()
                            progress_bar.empty()
                            st.session_state[zip_cache_key] = zip_buffer.getvalue()

                    if zip_cache_key in st.session_state:
                        st.success(f"✅ {len(mevcut_ilceler)} ilçe için PDF raporları arşivlendi ve indirilmeye hazır!")
                        st.download_button(
                            label="📥 Tüm İlçelerin Raporlarını Toplu İndir (ZIP)",
                            data=st.session_state[zip_cache_key],
                            file_name="Tum_Ilceler_KANTAR_Raporlari.zip",
                            mime="application/zip",
                            key="toplu_zip_indir_btn"
                        )
                else:
                    st.subheader(f"📑 {secilen_ilce} İlçesi Özel Değerlendirme Raporu")
                    st.markdown(f"Bu rapor **{', '.join(sirali_ay_isimleri)}** dönemlerini kapsamaktadır.")
                    st.markdown("Rapor İçeriği:")
                    st.markdown("- **1.** İlçenin İstanbul genelindeki başarı sıralaması ve 3 Kanser türündeki özet oranı")
                    st.markdown("- **2.** 3 tarama türü için ASM'lerin **ilçe içi sırasını** ve **oranlarını** gösteren 3 ayrı grafik")
                    st.markdown("- **3.** İlçeye bağlı tüm Aile Sağlığı Merkezlerinin (ASM) ve altındaki AHB birimlerinin **gereken**, **yapılan** tarama sayıları ile **başarı oranları**")
                    st.markdown("- **%50'nin altında kalan oranlar için kırmızı uyarı vurgusu** ⚠️")
                    
                    with st.spinner("PDF Raporu Hazırlanıyor, lütfen bekleyin..."):
                        pdf_data = get_ilce_rapor_pdf(df, secilen_ilce, secilen_aylar, logo_b64=kantar_logo_b64)
                    
                    st.download_button(
                        label=f"📥 {secilen_ilce} Raporunu İndir (PDF)",
                        data=pdf_data,
                        file_name=f"{secilen_ilce}_Ilce_KANTAR_Raporu.pdf",
                        mime="application/pdf"
                    )

                    st.markdown("---")
                    st.markdown(f"### 📊 {secilen_ilce} İlçesi - ASM İlçe İçi Sıralaması ve Oranları")
                    ilce_detay_ui = df[df['İlçe'] == secilen_ilce].copy()

                    for ikon, tarama_turu, renk_skalasi in [
                        ("🎀", "Mamografi", "Blues"),
                        ("🎗️", "Serviks", "Purples"),
                        ("🧬", "Kolorektal", "Greens"),
                    ]:
                        asm_rank_df = build_ilce_asm_ranking_df(ilce_detay_ui, tarama_turu)
                        if not asm_rank_df.empty:
                            plot_rank_df = asm_rank_df.iloc[::-1]
                            max_x_val = max(100, float(plot_rank_df['Oran (%)'].max()) * 1.15)
                            fig_ilce_asm = px.bar(
                                plot_rank_df,
                                x='Oran (%)',
                                y='Sıralı ASM Etiketi',
                                orientation='h',
                                text='Oran (%)',
                                title=f"{ikon} {secilen_ilce} İlçesi - {tarama_turu} Taraması ASM İlçe İçi Sıralaması ve Oranları",
                                color='Oran (%)',
                                color_continuous_scale=renk_skalasi,
                                labels={'Sıralı ASM Etiketi': 'İlçe İçi Sıra ve ASM Adı', 'Oran (%)': 'Başarı Oranı (%)'}
                            )
                            fig_ilce_asm.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
                            fig_ilce_asm.update_layout(
                                xaxis_range=[0, max_x_val],
                                height=max(320, len(plot_rank_df) * 34),
                                showlegend=False
                            )
                            fig_ilce_asm.add_vline(
                                x=70,
                                line_dash="dash",
                                line_color="red",
                                line_width=2,
                                annotation_text="Bakanlık %70 Alt Sınır Hedefi",
                                annotation_position="top right"
                            )
                            st.plotly_chart(fig_ilce_asm, use_container_width=True)
                    
            with tab5:
                st.subheader("📈 Aylık Trend ve Gelişim Analizi")
                
                if len(secilen_aylar) < 2:
                    st.info("Trend analizi yapabilmek için sol menüden en az 2 farklı ay (dosya) seçmelisiniz.")
                else:
                    def get_month_order(filename):
                        match = re.search(r'^(\d+)', str(filename))
                        return int(match.group(1)) if match else 99
                    
                    if secilen_ilce == "Tüm İlçeler":
                        st.markdown("**Tüm İstanbul Geneli Aylık Trend**")
                        trend_df_raw = df_filtrelenmis_ay.copy()
                        hedef_isim = "İstanbul Geneli"
                    else:
                        trend_df_raw = df_filtrelenmis_ay[df_filtrelenmis_ay['İlçe'] == secilen_ilce].copy()
                        trend_seviyesi = st.radio("Trend Seviyesi Seçin:", ["İlçe Geneli", "Belirli Bir ASM"])
                        
                        if trend_seviyesi == "Belirli Bir ASM":
                            asm_listesi = sorted(trend_df_raw['ASM Adı'].unique())
                            secilen_asm_trend = st.selectbox("Analiz edilecek ASM'yi seçin:", asm_listesi)
                            trend_df_raw = trend_df_raw[trend_df_raw['ASM Adı'] == secilen_asm_trend]
                            hedef_isim = secilen_asm_trend
                        else:
                            hedef_isim = f"{secilen_ilce} İlçe Geneli"
                    
                    trend_grup = trend_df_raw.groupby('Dosya (Ay)')[['Mamografi Gereken', 'Mamografi Yapılan', 'Serviks Gereken', 'Serviks Yapılan', 'Kolorektal Gereken', 'Kolorektal Yapılan']].sum().reset_index()
                    
                    trend_grup['Ay_Sira'] = trend_grup['Dosya (Ay)'].apply(get_month_order)
                    trend_grup = trend_grup.sort_values('Ay_Sira').drop('Ay_Sira', axis=1).reset_index(drop=True)
                    
                    for t in ['Mamografi', 'Serviks', 'Kolorektal']:
                        trend_grup[f'{t} Oranı (%)'] = (trend_grup[f'{t} Yapılan'] / trend_grup[f'{t} Gereken'] * 100).fillna(0)
                    
                    trend_uzun_df = trend_grup.melt(
                        id_vars=['Dosya (Ay)'],
                        value_vars=['Mamografi Oranı (%)', 'Serviks Oranı (%)', 'Kolorektal Oranı (%)'],
                        var_name='Tarama Türü',
                        value_name='Başarı Oranı (%)'
                    )
                    max_y_trend = max(100, float(trend_uzun_df['Başarı Oranı (%)'].max()) * 1.12)

                    st.markdown(f"#### 📊 {hedef_isim} - Tarama Oranları Değişimi")
                    fig_trend = px.line(
                        trend_uzun_df,
                        x='Dosya (Ay)',
                        y='Başarı Oranı (%)',
                        color='Tarama Türü',
                        markers=True,
                        title=f"{hedef_isim} - Aylık Tarama Oranları Değişimi (Bakanlık %70 Alt Sınır Hedefi)"
                    )
                    fig_trend.update_layout(yaxis_range=[0, max_y_trend])
                    fig_trend.add_hline(
                        y=70,
                        line_dash="dash",
                        line_color="red",
                        line_width=2,
                        annotation_text="Bakanlık %70 Alt Sınır Hedefi",
                        annotation_position="top right"
                    )
                    st.plotly_chart(fig_trend, use_container_width=True)
                    
                    st.markdown("#### 💡 Son Durum ve Otomatik İçgörüler")
                    
                    son_ay = trend_grup.iloc[-1]
                    onceki_ay = trend_grup.iloc[-2]
                    
                    son_ay_isim = str(son_ay['Dosya (Ay)']).replace('.xlsx', '').replace('Birim Bazında', '').strip()
                    onceki_ay_isim = str(onceki_ay['Dosya (Ay)']).replace('.xlsx', '').replace('Birim Bazında', '').strip()
                    
                    st.markdown(f"***{son_ay_isim}*** verileri ile ***{onceki_ay_isim}*** verilerinin karşılaştırması:")
                    
                    col_m, col_s, col_k = st.columns(3)
                    
                    def metric_goster(col, baslik, oran_kolon):
                        guncel_deger = son_ay[oran_kolon]
                        fark = son_ay[oran_kolon] - onceki_ay[oran_kolon]
                        col.metric(
                            label=baslik, 
                            value=f"%{guncel_deger:.2f}", 
                            delta=f"{fark:.2f}%"
                        )
                    
                    metric_goster(col_m, "🎀 Mamografi", "Mamografi Oranı (%)")
                    metric_goster(col_s, "🎗️ Serviks", "Serviks Oranı (%)")
                    metric_goster(col_k, "🧬 Kolorektal", "Kolorektal Oranı (%)")

            with tab6:
                st.subheader("🚨 Acil Müdahale Gerektiren Birimler (Tarama Yapılmayan / %0 Birimler)")
                st.markdown("Bu sekmede, seçilen ilçe ve aylar kapsamında **birim özelinde hangi aylarda sıfır tarama yapıldığı** ve **kaç kez sıfır alındığı** tarama türüne göre ayrı sekmelerde listelenmektedir.")
                
                if secilen_ilce != "Tüm İlçeler":
                    acil_df_base = df_filtrelenmis_ay[df_filtrelenmis_ay['İlçe'] == secilen_ilce].copy()
                else:
                    acil_df_base = df_filtrelenmis_ay.copy()

                def get_month_order_acil(filename):
                    match = re.search(r'^(\d+)', str(filename))
                    return int(match.group(1)) if match else 99

                def clean_month_name_acil(filename):
                    return str(filename).replace('.xlsx', '').replace('.xls', '').replace('Birim Bazında', '').strip()

                def acil_tarama_tablosu_goster(tarama_turu):
                    yapilan_col = f"{tarama_turu} Yapılan"
                    gereken_col = f"{tarama_turu} Gereken"

                    df_zero = acil_df_base[acil_df_base[yapilan_col] == 0].copy()

                    if df_zero.empty:
                        st.success(f"🎉 Tebrikler! Seçilen kriterlerde **{tarama_turu}** taraması sıfır olan birim bulunmuyor.")
                        return

                    df_zero['Ay_Sira'] = df_zero['Dosya (Ay)'].apply(get_month_order_acil)
                    df_zero['Ay_Temiz'] = df_zero['Dosya (Ay)'].apply(clean_month_name_acil)
                    df_zero = df_zero.sort_values(['Ay_Sira', 'Dosya (Ay)'])

                    df_birim_ozet = df_zero.groupby(['İlçe', 'ASM Adı', 'Birim Adı', 'Birim ID'], as_index=False).agg(
                        **{
                            'Sıfır Olduğu Ay Sayısı': ('Dosya (Ay)', 'nunique'),
                            'Sıfır Olduğu Aylar': ('Ay_Temiz', lambda x: ', '.join(dict.fromkeys(x))),
                            'Toplam Gereken Hedef': (gereken_col, 'sum')
                        }
                    )

                    df_birim_ozet = df_birim_ozet.sort_values(
                        by=['Sıfır Olduğu Ay Sayısı', 'Toplam Gereken Hedef', 'İlçe', 'ASM Adı', 'Birim Adı'],
                        ascending=[False, False, True, True, True]
                    ).reset_index(drop=True)
                    df_birim_ozet.index += 1

                    st.warning(
                        f"⚠️ **{tarama_turu}** taramasında toplam **{len(df_birim_ozet)}** birimin "
                        f"en az 1 ayda sıfır tarama yaptığı tespit edildi."
                    )

                    st.dataframe(
                        df_birim_ozet.style.format({
                            'Sıfır Olduğu Ay Sayısı': "{:.0f}",
                            'Toplam Gereken Hedef': "{:.0f}"
                        }).background_gradient(subset=['Sıfır Olduğu Ay Sayısı'], cmap='Reds'),
                        use_container_width=True
                    )

                    col_ex, col_pdf = st.columns(2)
                    with col_ex:
                        st.download_button(
                            f"📥 Excel İndir ({tarama_turu} Sıfır Listesi)",
                            data=get_excel(df_birim_ozet),
                            file_name=f"{tarama_turu}_Acil_Mudahale_Sifir_Taramalar.xlsx",
                            key=f"acil_excel_btn_{tarama_turu}"
                        )
                    with col_pdf:
                        st.download_button(
                            f"📄 PDF İndir ({tarama_turu} Sıfır Listesi)",
                            data=get_pdf(df_birim_ozet, f"{tarama_turu} - Acil Müdahale Sıfır Tarama Birimleri"),
                            file_name=f"{tarama_turu}_Acil_Mudahale_Sifir_Taramalar.pdf",
                            mime="application/pdf",
                            key=f"acil_pdf_btn_{tarama_turu}"
                        )

                subtab_m, subtab_s, subtab_k = st.tabs([
                    "🎀 Mamografi Sıfır Taramalar",
                    "🎗️ Serviks Sıfır Taramalar",
                    "🧬 Kolorektal Sıfır Taramalar"
                ])

                with subtab_m:
                    acil_tarama_tablosu_goster("Mamografi")
                with subtab_s:
                    acil_tarama_tablosu_goster("Serviks")
                with subtab_k:
                    acil_tarama_tablosu_goster("Kolorektal")

        else:
            st.warning("Lütfen sol menüden analiz edilecek en az bir ay (dosya) seçin.")

    except Exception as e:
        st.error(f"Dosyalar okunurken bir hata oluştu. Hata detayı: {e}")
else:
    st.info("👈 Lütfen sol menüden analiz edilecek Excel dosyasını (veya dosyalarını) sisteme yükleyin.")
