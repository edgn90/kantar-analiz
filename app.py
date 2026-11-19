import streamlit as st
import pandas as pd
import io
import os
import glob
import base64
import zipfile
import urllib.request
import ssl
import re
import plotly.express as px
from PIL import Image, ImageDraw, ImageFont
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


def safe_int(val, default=0):
    """
    NaN, inf veya None değerlerde 'cannot convert float NaN to integer'
    hatasını önleyerek güvenli tam sayı dönüşümü yapar.
    """
    try:
        if pd.isna(val) or val in (float("inf"), float("-inf")):
            return default
        return int(round(float(val)))
    except Exception:
        return default


def normalize_tr_text(text):
    s = str(text)
    tr_map = str.maketrans({
        'İ': 'i', 'I': 'i', 'ı': 'i',
        'Ş': 's', 'ş': 's',
        'Ğ': 'g', 'ğ': 'g',
        'Ü': 'u', 'ü': 'u',
        'Ö': 'o', 'ö': 'o',
        'Ç': 'c', 'ç': 'c',
    })
    return s.translate(tr_map).lower()


def _tr_title_district(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.lower() == 'nan':
        return None
    tr_lower_map = str.maketrans({'İ': 'i', 'I': 'ı', 'Ş': 'ş', 'Ğ': 'ğ', 'Ü': 'ü', 'Ö': 'ö', 'Ç': 'ç'})
    s_low = s.translate(tr_lower_map).lower()
    if not s_low:
        return None
    first = s_low[0]
    first_up_map = {'i': 'İ', 'ı': 'I', 'ş': 'Ş', 'ğ': 'Ğ', 'ü': 'Ü', 'ö': 'Ö', 'ç': 'Ç'}
    first_up = first_up_map.get(first, first.upper())
    return first_up + s_low[1:]


def _parse_month_components(filename):
    """
    Dosya adından yıl ve ay bilgisini (Türkçe ay isimleri, ASCII karşılıkları veya
    baştaki/yıl sonrasındaki ay numaraları üzerinden) ayrıştırır.
    """
    raw_name = str(filename).strip()
    clean_name = re.sub(r'\.(xlsx|xls)$', '', raw_name, flags=re.IGNORECASE)
    norm = normalize_tr_text(clean_name)

    year_match = re.search(r'(20\d{2})', norm)
    year = int(year_match.group(1)) if year_match else 0

    ay_sozlugu = [
        ('ocak', 1), ('subat', 2), ('mart', 3), ('nisan', 4),
        ('mayis', 5), ('haziran', 6), ('temmuz', 7), ('agustos', 8),
        ('eylul', 9), ('ekim', 10), ('kasim', 11), ('aralik', 12),
    ]

    month_num = None
    for ay_adi, ay_no in ay_sozlugu:
        if re.search(rf'(?:^|[^a-z]){ay_adi}(?:[^a-z]|$)', norm):
            month_num = ay_no
            break

    if month_num is None:
        lead_match = re.match(r'^0?(\d{1,2})(?:[^\d]|$)', norm)
        if lead_match:
            val = int(lead_match.group(1))
            if 1 <= val <= 12:
                month_num = val
        if month_num is None and year > 0:
            ym_match = re.search(r'20\d{2}[_\-\.\s]+0?(\d{1,2})(?:[^\d]|$)', norm)
            if ym_match:
                val = int(ym_match.group(1))
                if 1 <= val <= 12:
                    month_num = val

    if month_num is None:
        fallback_num = re.search(r'^(\d+)', norm)
        month_num = int(fallback_num.group(1)) if fallback_num else 99

    return year, month_num, raw_name


def build_month_order_map(labels):
    """
    Verilen dosya/ay etiketlerini Türkçe ay sırasına (ve birden fazla yıl varsa yıl sırasına)
    göre sıralayıp her etikete kronolojik sıra indeksi atar.
    """
    unique_labels = list(dict.fromkeys(labels))
    parsed = {lbl: _parse_month_components(lbl) for lbl in unique_labels}
    known_years = {y for (y, _, _) in parsed.values() if y > 0}

    if len(known_years) <= 1:
        ordered = sorted(unique_labels, key=lambda l: (parsed[l][1], parsed[l][2]))
    else:
        default_year = max(known_years)
        ordered = sorted(
            unique_labels,
            key=lambda l: (parsed[l][0] if parsed[l][0] > 0 else default_year, parsed[l][1], parsed[l][2])
        )
    return {lbl: idx for idx, lbl in enumerate(ordered)}


def sort_month_labels(labels):
    order_map = build_month_order_map(labels)
    return sorted(labels, key=lambda l: order_map.get(l, 999))


def style_gradient_safe(styler, subset_col, cmap_type="Blues"):
    """
    matplotlib kurulu olmasa dahi DataFrame hücrelerine mavi veya kırmızı
    renk gradyanı uygulayan saf Python yardımcı fonksiyonu.
    """
    if subset_col not in styler.data.columns or styler.data.empty:
        return styler

    s = pd.to_numeric(styler.data[subset_col], errors="coerce").replace([float("inf"), float("-inf")], float("nan"))
    if s.dropna().empty:
        return styler

    vmin, vmax = float(s.min()), float(s.max())

    def _color_val(val):
        if pd.isna(val) or val in (float("inf"), float("-inf")):
            return ""
        ratio = 0.5 if vmax == vmin else max(0.0, min(1.0, (float(val) - vmin) / (vmax - vmin)))
        if cmap_type == "Reds":
            r = int(255 - ratio * (255 - 203))
            g = int(245 - ratio * (245 - 24))
            b = int(240 - ratio * (240 - 29))
        else:
            r = int(247 - ratio * (247 - 33))
            g = int(251 - ratio * (251 - 113))
            b = int(255 - ratio * (255 - 181))
        text_color = "#ffffff" if ratio > 0.65 else "#111111"
        return f"background-color: rgb({r}, {g}, {b}); color: {text_color};"

    return styler.apply(lambda col: [_color_val(v) for v in col], subset=[subset_col])


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
        import matplotlib as _mpl
        mpl_ttf_dir = os.path.join(os.path.dirname(_mpl.__file__), "mpl-data", "fonts", "ttf")
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


def _mapping_file_priority(path):
    """
    GitHub deposunda birden fazla ASM_AHB_Eslestirme*.xlsx dosyası varsa
    dosya adındaki yıl ve ay bilgisini (ör. KASIM_2026) önceliklendirerek en güncelini seçer.
    """
    base = os.path.basename(path)
    norm = normalize_tr_text(re.sub(r'\.(xlsx|xls)$', '', base, flags=re.IGNORECASE))
    year_match = re.search(r'(20\d{2})', norm)
    year = int(year_match.group(1)) if year_match else 0

    ay_sozlugu = [
        ('ocak', 1), ('subat', 2), ('mart', 3), ('nisan', 4),
        ('mayis', 5), ('haziran', 6), ('temmuz', 7), ('agustos', 8),
        ('eylul', 9), ('ekim', 10), ('kasim', 11), ('aralik', 12),
    ]
    month_num = 0
    for ay_adi, ay_no in ay_sozlugu:
        if re.search(rf'(?:^|[^a-z]){ay_adi}(?:[^a-z]|$)', norm):
            month_num = ay_no
            break

    try:
        mtime = os.path.getmtime(path)
    except Exception:
        mtime = 0.0

    return (year, month_num, mtime, base)


def _parse_mapping_excel_source(source):
    """
    Eşleştirme Excel'indeki tüm sayfaları okur; ASM sütunu boş olan ('Eşleşmeyenler' vb.)
    satırları ayıklayarak yalnızca geçerli ASM ve İlçe eşleştirme sözlüklerini döndürür.
    """
    sheets = pd.read_excel(source, sheet_name=None)
    df_map = pd.concat(sheets.values(), ignore_index=True)
    df_map.columns = [str(c).strip() for c in df_map.columns]

    ahb_col = next((c for c in df_map.columns if 'ahb' in normalize_tr_text(c)), 'AHB BİRİMİ')
    asm_col = next((c for c in df_map.columns if normalize_tr_text(c) == 'asm'), 'ASM')
    ilce_col = next((c for c in df_map.columns if 'ilce' in normalize_tr_text(c)), None)

    valid_asm = df_map[df_map[ahb_col].notna() & df_map[asm_col].notna()].copy()
    valid_asm[ahb_col] = valid_asm[ahb_col].astype(str).str.strip()
    valid_asm[asm_col] = valid_asm[asm_col].astype(str).str.strip()
    valid_asm = valid_asm[
        (valid_asm[ahb_col] != '') &
        (valid_asm[asm_col] != '') &
        (valid_asm[asm_col].str.lower() != 'nan')
    ]

    asm_dict = dict(zip(valid_asm[ahb_col].apply(normalize_tr_text), valid_asm[asm_col]))

    ilce_dict = {}
    if ilce_col and ilce_col in df_map.columns:
        valid_ilce = df_map[df_map[ahb_col].notna() & df_map[ilce_col].notna()].copy()
        for _, r in valid_ilce.iterrows():
            d_clean = _tr_title_district(r[ilce_col])
            if d_clean:
                ilce_dict[normalize_tr_text(str(r[ahb_col]).strip())] = d_clean

    return asm_dict, ilce_dict


@st.cache_data
def load_asm_mapping_from_disk(candidate_signature):
    try:
        if not candidate_signature:
            return None, {}, None
        target_path = candidate_signature[0][0]
        asm_dict, ilce_dict = _parse_mapping_excel_source(target_path)
        return asm_dict, ilce_dict, os.path.basename(target_path)
    except Exception:
        return None, {}, None


def get_active_asm_mapping():
    """
    Proje (GitHub) klasöründeki ASM_AHB_Eslestirme*.xlsx dosyalarından
    en güncel olanını otomatik seçip yükler.
    """
    candidates = glob.glob("ASM_AHB_Eslestirme*.xlsx") + glob.glob("ASM_AHB_Eslestirme*.xls")
    candidates = [p for p in candidates if os.path.isfile(p) and not os.path.basename(p).startswith("~$")]
    candidates.sort(key=_mapping_file_priority, reverse=True)
    signature = tuple((p, os.path.getmtime(p)) for p in candidates)
    return load_asm_mapping_from_disk(signature)


@st.cache_data
def load_data(files, asm_mapping, ilce_mapping=None):
    tum_veriler = []
    ilce_mapping = ilce_mapping or {}
    
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
                temp_df['Birim ID'] = temp_df['Birim Adı'].astype(str).str.extract(r'(\d+)\s+Nolu').fillna("Bilinmiyor")
                temp_df['Dosya (Ay)'] = dosya_adi
                
                def get_col(df, keyword, val_type):
                    for c in df.columns:
                        if val_type in c.lower() and keyword.lower() in c.lower():
                            return (
                                pd.to_numeric(df[c], errors='coerce')
                                .replace([float('inf'), float('-inf')], 0)
                                .fillna(0)
                            )
                    return pd.Series([0.0]*len(df))
                
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

        gerekli_kolonlar = ['Birim Adı', 'Mamografi Gereken', 'Mamografi Oranı', 'Serviks Gereken', 'Serviks Oranı', 'Kolorektal Gereken', 'Kolorektal Oranı']
        eksik_kolonlar = [k for k in gerekli_kolonlar if k not in df_old.columns]
        if eksik_kolonlar:
            raise ValueError(f"'{dosya_adi}' dosyası geçerli bir KAN-TAR tablosu formatında değil (Eksik sütunlar: {', '.join(eksik_kolonlar)}).")

        if 'Birim ID' not in df_old.columns:
            df_old['Birim ID'] = df_old['Birim Adı'].astype(str).str.extract(r'(\d+)\s+Nolu').fillna("Bilinmiyor")
        
        for t in ['Mamografi', 'Serviks', 'Kolorektal']:
            gereken = (
                pd.to_numeric(df_old[f'{t} Gereken'], errors='coerce')
                .replace([float('inf'), float('-inf')], 0)
                .fillna(0)
            )
            oran = (
                pd.to_numeric(df_old[f'{t} Oranı'], errors='coerce')
                .replace([float('inf'), float('-inf')], 0)
                .fillna(0)
            )
            df_old[f'{t} Gereken'] = gereken
            df_old[f'{t} Yapılan'] = (oran * gereken / 100).round().fillna(0)
            
        df_old_final = df_old[['Birim Adı', 'Birim ID', 'Mamografi Gereken', 'Mamografi Yapılan', 'Serviks Gereken', 'Serviks Yapılan', 'Kolorektal Gereken', 'Kolorektal Yapılan']].copy()
        df_old_final['Dosya (Ay)'] = dosya_adi
        tum_veriler.append(df_old_final)

    combined_df = pd.concat(tum_veriler, ignore_index=True)

    combined_df = combined_df[combined_df['Birim Adı'].notna()].copy()
    combined_df['Birim Adı'] = combined_df['Birim Adı'].astype(str).str.strip()
    combined_df = combined_df[
        (combined_df['Birim Adı'] != '') &
        (combined_df['Birim Adı'].str.lower() != 'nan') &
        (~combined_df['Birim Adı'].str.lower().isin(['toplam', 'genel toplam']))
    ].reset_index(drop=True)

    sayisal_sutunlar = [
        'Mamografi Gereken', 'Mamografi Yapılan',
        'Serviks Gereken', 'Serviks Yapılan',
        'Kolorektal Gereken', 'Kolorektal Yapılan'
    ]
    for col in sayisal_sutunlar:
        combined_df[col] = (
            pd.to_numeric(combined_df[col], errors='coerce')
            .replace([float('inf'), float('-inf')], 0)
            .fillna(0)
        )

    norm_birim_keys = combined_df['Birim Adı'].apply(normalize_tr_text)
    
    if asm_mapping:
        combined_df['ASM Adı'] = norm_birim_keys.map(asm_mapping).fillna("Bilinmeyen/Eksik ASM")
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
        norm_key = normalize_tr_text(str(birim_adi).strip())
        if norm_key in ilce_mapping:
            return ilce_mapping[norm_key]
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
    asm_df['Oran (%)'] = (
        (asm_df[yapilan_col] / asm_df[gereken_col] * 100)
        .replace([float('inf'), float('-inf')], 0)
        .fillna(0)
    )
    asm_df = asm_df.sort_values(by=['Oran (%)', 'ASM Adı'], ascending=[False, True]).reset_index(drop=True)
    asm_df['İlçe İçi Sıra'] = asm_df.index + 1
    asm_df['Sıralı ASM Etiketi'] = asm_df.apply(lambda r: f"{safe_int(r['İlçe İçi Sıra'])}. {r['ASM Adı']}", axis=1)
    return asm_df


def _create_ilce_asm_chart_pillow_base64(asm_df, secilen_ilce, tarama_turu):
    """
    matplotlib bulunmadığında Pillow ile Türkçe karakter destekli yatay sıralama grafiği çizer.
    """
    reg_path, bold_path = get_turkish_font_paths()
    width = 1100
    row_h = 34
    top_pad = 55
    bottom_pad = 58
    left_pad = 350
    right_pad = 110
    height = max(240, top_pad + bottom_pad + len(asm_df) * row_h)

    img = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    try:
        title_font = ImageFont.truetype(bold_path, 17) if bold_path else ImageFont.load_default()
        label_font = ImageFont.truetype(reg_path, 13) if reg_path else ImageFont.load_default()
        bold_font = ImageFont.truetype(bold_path, 13) if bold_path else ImageFont.load_default()
        small_font = ImageFont.truetype(reg_path, 11) if reg_path else ImageFont.load_default()
    except Exception:
        title_font = label_font = bold_font = small_font = ImageFont.load_default()

    title = f"{secilen_ilce} İlçesi - {tarama_turu} Taraması ASM İlçe İçi Sıralaması ve Oranları"
    draw.text((width // 2, 22), title, fill=(44, 62, 80), font=title_font, anchor="mm")

    max_val = max(100.0, float(asm_df['Oran (%)'].max()) + 12.0)
    plot_x0 = left_pad
    plot_x1 = width - right_pad
    plot_w = plot_x1 - plot_x0
    plot_y0 = top_pad
    plot_y1 = height - bottom_pad

    for tick in [0, 25, 50, 70, 100]:
        if tick <= max_val:
            gx = plot_x0 + int((tick / max_val) * plot_w)
            draw.line([(gx, plot_y0), (gx, plot_y1)], fill=(230, 230, 230), width=1)
            draw.text((gx, plot_y1 + 14), f"%{tick}", fill=(100, 100, 100), font=small_font, anchor="mm")

    for idx, row in asm_df.iterrows():
        y_center = plot_y0 + idx * row_h + row_h // 2
        bar_h = 20
        val = 0.0 if pd.isna(row['Oran (%)']) else float(row['Oran (%)'])
        label = str(row['Sıralı ASM Etiketi'])
        if len(label) > 40:
            label = label[:37] + "..."
        draw.text((plot_x0 - 10, y_center), label, fill=(44, 62, 80), font=label_font, anchor="rm")

        bar_w = int((val / max_val) * plot_w)
        if val >= 70:
            color = (41, 128, 185)   # Mavi (#2980b9)
        elif val >= 50:
            color = (243, 156, 18)   # Turuncu (#f39c12)
        else:
            color = (192, 57, 43)    # Kırmızı (#c0392b)

        draw.rectangle([plot_x0, y_center - bar_h // 2, plot_x0 + bar_w, y_center + bar_h // 2], fill=color)
        draw.text((plot_x0 + bar_w + 8, y_center), f"%{val:.2f}", fill=(44, 62, 80), font=bold_font, anchor="lm")

    x70 = plot_x0 + int((70.0 / max_val) * plot_w)
    for y in range(plot_y0, plot_y1, 10):
        draw.line([(x70, y), (x70, min(y + 6, plot_y1))], fill=(214, 48, 49), width=2)
    draw.text((x70, plot_y1 + 34), "Bakanlık %70 Alt Sınır Hedefi", fill=(214, 48, 49), font=small_font, anchor="mm")

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode("utf-8")


def create_ilce_asm_chart_base64(ilce_detay, secilen_ilce, tarama_turu):
    asm_df = build_ilce_asm_ranking_df(ilce_detay, tarama_turu)
    if asm_df.empty:
        return None

    if not HAS_MATPLOTLIB:
        return _create_ilce_asm_chart_pillow_base64(asm_df, secilen_ilce, tarama_turu)

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


def _create_istanbul_baraj_chart_pillow_base64(tarama_turleri, gecenler, kalanlar, toplam_ilce):
    reg_path, bold_path = get_turkish_font_paths()
    width, height = 1000, 420
    img = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    try:
        title_font = ImageFont.truetype(bold_path, 18) if bold_path else ImageFont.load_default()
        label_font = ImageFont.truetype(bold_path, 14) if bold_path else ImageFont.load_default()
        small_font = ImageFont.truetype(reg_path, 12) if reg_path else ImageFont.load_default()
    except Exception:
        title_font = label_font = small_font = ImageFont.load_default()

    draw.text(
        (width // 2, 26),
        f"İstanbul Geneli ({toplam_ilce} İlçe) - %70 Hedef Barajını Geçen / Kalan İlçe Sayıları",
        fill=(44, 62, 80),
        font=title_font,
        anchor="mm"
    )

    plot_x0, plot_x1 = 110, width - 80
    plot_y0, plot_y1 = 75, height - 75
    plot_h = plot_y1 - plot_y0
    max_y = max(toplam_ilce + 5, max(gecenler + kalanlar + [10]))

    for tick in range(0, int(max_y) + 1, 10):
        gy = plot_y1 - int((tick / max_y) * plot_h)
        draw.line([(plot_x0, gy), (plot_x1, gy)], fill=(230, 230, 230), width=1)
        draw.text((plot_x0 - 12, gy), str(tick), fill=(100, 100, 100), font=small_font, anchor="rm")

    group_w = (plot_x1 - plot_x0) // len(tarama_turleri)
    bar_w = 68
    for idx, t in enumerate(tarama_turleri):
        cx = plot_x0 + idx * group_w + group_w // 2
        g_val = gecenler[idx]
        k_val = kalanlar[idx]

        gh = int((g_val / max_y) * plot_h)
        kh = int((k_val / max_y) * plot_h)

        draw.rectangle([cx - bar_w - 6, plot_y1 - gh, cx - 6, plot_y1], fill=(39, 174, 96))
        g_pct = (g_val / toplam_ilce * 100) if toplam_ilce > 0 else 0
        draw.text((cx - bar_w // 2 - 6, plot_y1 - gh - 12), f"{g_val} (%{g_pct:.1f})", fill=(30, 132, 73), font=small_font, anchor="mm")

        draw.rectangle([cx + 6, plot_y1 - kh, cx + bar_w + 6, plot_y1], fill=(192, 57, 43))
        k_pct = (k_val / toplam_ilce * 100) if toplam_ilce > 0 else 0
        draw.text((cx + bar_w // 2 + 6, plot_y1 - kh - 12), f"{k_val} (%{k_pct:.1f})", fill=(146, 43, 33), font=small_font, anchor="mm")

        draw.text((cx, plot_y1 + 22), t, fill=(44, 62, 80), font=label_font, anchor="mm")

    draw.rectangle([240, height - 32, 258, height - 16], fill=(39, 174, 96))
    draw.text((266, height - 24), "%70 Barajını Geçen İlçe", fill=(44, 62, 80), font=small_font, anchor="lm")
    draw.rectangle([520, height - 32, 538, height - 16], fill=(192, 57, 43))
    draw.text((546, height - 24), "%70 Altında Kalan İlçe", fill=(44, 62, 80), font=small_font, anchor="lm")

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode("utf-8")


def create_istanbul_baraj_chart_base64(ilce_ozet_df):
    """
    3 tarama türünde %70 hedef barajını geçen ve altında kalan ilçe sayılarını
    karşılaştırmalı grafik olarak üretir.
    """
    if ilce_ozet_df.empty:
        return None

    tarama_turleri = ['Mamografi', 'Serviks', 'Kolorektal']
    gecenler = [safe_int((ilce_ozet_df[f'{t} Oranı'] >= 70.0).sum()) for t in tarama_turleri]
    kalanlar = [safe_int((ilce_ozet_df[f'{t} Oranı'] < 70.0).sum()) for t in tarama_turleri]
    toplam_ilce = len(ilce_ozet_df)

    if not HAS_MATPLOTLIB:
        return _create_istanbul_baraj_chart_pillow_base64(tarama_turleri, gecenler, kalanlar, toplam_ilce)

    fig, ax = plt.subplots(figsize=(7.8, 3.1), dpi=150)
    x_pos = list(range(len(tarama_turleri)))
    bar_width = 0.34

    bars_gecen = ax.bar([x - bar_width / 2 for x in x_pos], gecenler, width=bar_width, label='%70 Barajını Geçen İlçe', color='#27ae60')
    bars_kalan = ax.bar([x + bar_width / 2 for x in x_pos], kalanlar, width=bar_width, label='%70 Altında Kalan İlçe', color='#c0392b')

    ax.set_xticks(x_pos)
    ax.set_xticklabels(tarama_turleri, fontsize=9.5, fontweight='bold')
    ax.set_ylabel('İlçe Sayısı', fontsize=8.5)
    ax.set_ylim(0, max(toplam_ilce + 7, max(gecenler + kalanlar) + 8))
    ax.set_title(f"İstanbul Geneli ({toplam_ilce} İlçe) - Tarama Türlerine Göre %70 Hedef Barajı Durumu", fontsize=10, fontweight='bold', pad=8)
    ax.grid(axis='y', linestyle=':', alpha=0.5)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    for bar in bars_gecen:
        h = safe_int(bar.get_height())
        pct = (h / toplam_ilce * 100) if toplam_ilce > 0 else 0
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            h + 0.6,
            f"{h} İlçe (%{pct:.1f})",
            ha='center',
            va='bottom',
            fontsize=8,
            fontweight='bold',
            color='#1e8449'
        )

    for bar in bars_kalan:
        h = safe_int(bar.get_height())
        pct = (h / toplam_ilce * 100) if toplam_ilce > 0 else 0
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            h + 0.6,
            f"{h} İlçe (%{pct:.1f})",
            ha='center',
            va='bottom',
            fontsize=8,
            fontweight='bold',
            color='#922b21'
        )

    ax.legend(loc='upper right', fontsize=8, frameon=True)
    plt.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight')
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode('utf-8')


def create_istanbul_ilce_ranking_chart_base64(ilce_ozet_df, tarama_turu):
    """
    Tüm İstanbul ilçelerinin ilgili tarama türündeki il sıralamasını ve oranlarını
    yatay çubuk grafik olarak üretir.
    """
    if ilce_ozet_df.empty:
        return None

    oran_col = f"{tarama_turu} Oranı"
    sira_col = f"{tarama_turu} Sıra"

    sub_df = ilce_ozet_df.sort_values(by=[oran_col, 'İlçe'], ascending=[False, True]).reset_index(drop=True)
    sub_df['Sıralı Etiket'] = sub_df.apply(lambda r: f"{safe_int(r[sira_col])}. {r['İlçe']}", axis=1)

    if not HAS_MATPLOTLIB:
        pillow_df = pd.DataFrame({
            'Sıralı ASM Etiketi': sub_df['Sıralı Etiket'],
            'Oran (%)': sub_df[oran_col]
        })
        return _create_ilce_asm_chart_pillow_base64(pillow_df, "İstanbul Geneli", tarama_turu)

    plot_df = sub_df.iloc[::-1]
    fig_height = min(10.2, max(3.2, len(plot_df) * 0.245))
    fig, ax = plt.subplots(figsize=(8.2, fig_height), dpi=150)

    bar_colors = ['#2980b9' if val >= 70 else ('#f39c12' if val >= 50 else '#c0392b') for val in plot_df[oran_col]]
    bars = ax.barh(plot_df['Sıralı Etiket'], plot_df[oran_col], color=bar_colors, height=0.66)

    ax.axvline(70, color='#d63031', linestyle='--', linewidth=1.4, label='Bakanlık %70 Alt Sınır Hedefi')
    ax.set_xlim(0, max(105, float(plot_df[oran_col].max()) + 14))
    ax.set_xlabel('Başarı Oranı (%)', fontsize=8.5)
    ax.set_title(f"İstanbul İl Geneli - {tarama_turu} Taraması İlçe Sıralaması ve Oranları", fontsize=10, fontweight='bold', pad=8)
    ax.tick_params(axis='both', labelsize=7.6)
    ax.grid(axis='x', linestyle=':', alpha=0.5)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    for bar, val in zip(bars, plot_df[oran_col]):
        ax.text(
            val + 1.1,
            bar.get_y() + bar.get_height() / 2,
            f"%{val:.2f}",
            va='center',
            ha='left',
            fontsize=7.5,
            fontweight='bold',
            color='#2c3e50'
        )

    ax.legend(loc='lower right', fontsize=8, frameon=True)
    plt.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight')
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode('utf-8')


def get_istanbul_yonetici_ozeti_pdf(df, secilen_aylar, logo_b64=None):
    """
    Tüm İstanbul ilçelerinin 3 tarama türündeki oranlarını, il sıralamalarını ve
    %70 barajını geçen/kalan ilçe sayılarını grafiklerle özetleyen Yönetici Özeti PDF'i üretir.
    """
    fonts_loaded, font_face_css = setup_fonts()
    font_family = '"DejaVuSans", sans-serif' if fonts_loaded else "Helvetica, Arial, sans-serif"

    sirali_aylar = sort_month_labels(secilen_aylar)
    aylar_str = ", ".join(sirali_aylar)

    df_ilceler = df[df['İlçe'] != "Diğer/Kurum Tabiplikleri"].copy()
    if df_ilceler.empty:
        df_ilceler = df.copy()

    ilce_grup = df_ilceler.groupby('İlçe', as_index=False)[
        ['Mamografi Gereken', 'Mamografi Yapılan', 'Serviks Gereken', 'Serviks Yapılan', 'Kolorektal Gereken', 'Kolorektal Yapılan']
    ].sum()

    for t in ['Mamografi', 'Serviks', 'Kolorektal']:
        ilce_grup[f'{t} Oranı'] = (
            (ilce_grup[f'{t} Yapılan'] / ilce_grup[f'{t} Gereken'] * 100)
            .replace([float('inf'), float('-inf')], 0)
            .fillna(0)
        )
        ilce_grup = ilce_grup.sort_values(by=[f'{t} Oranı', 'İlçe'], ascending=[False, True]).reset_index(drop=True)
        ilce_grup[f'{t} Sıra'] = ilce_grup.index + 1

    ilce_grup['Ortalama Oran'] = (
        ilce_grup['Mamografi Oranı'] + ilce_grup['Serviks Oranı'] + ilce_grup['Kolorektal Oranı']
    ) / 3.0
    ilce_grup['Gecilen Baraj Sayisi'] = (
        (ilce_grup['Mamografi Oranı'] >= 70.0).astype(int) +
        (ilce_grup['Serviks Oranı'] >= 70.0).astype(int) +
        (ilce_grup['Kolorektal Oranı'] >= 70.0).astype(int)
    )
    ilce_grup = ilce_grup.sort_values(by=['Ortalama Oran', 'İlçe'], ascending=[False, True]).reset_index(drop=True)
    ilce_grup['Genel Sıra'] = ilce_grup.index + 1

    toplam_ilce_sayisi = len(ilce_grup)

    def format_oran_exec(oran):
        if pd.isna(oran) or oran in (float("inf"), float("-inf")):
            oran = 0.0
        if oran >= 70.0:
            return f"<span style='color: #1e8449; font-weight: bold;'>%{oran:.2f}</span>"
        if oran < 50.0:
            return f"<span style='color: #c0392b; font-weight: bold;'>%{oran:.2f}</span>"
        return f"<b>%{oran:.2f}</b>"

    ozet_kartlar = {}
    for t in ['Mamografi', 'Serviks', 'Kolorektal']:
        top_ger = safe_int(df[f'{t} Gereken'].sum())
        top_yap = safe_int(df[f'{t} Yapılan'].sum())
        il_oran = (top_yap / top_ger * 100) if top_ger > 0 else 0.0
        gecen = safe_int((ilce_grup[f'{t} Oranı'] >= 70.0).sum())
        kalan = safe_int((ilce_grup[f'{t} Oranı'] < 70.0).sum())
        gecen_yuzde = (gecen / toplam_ilce_sayisi * 100) if toplam_ilce_sayisi > 0 else 0.0
        kalan_yuzde = (kalan / toplam_ilce_sayisi * 100) if toplam_ilce_sayisi > 0 else 0.0

        sirali_t = ilce_grup.sort_values(by=[f'{t} Oranı', 'İlçe'], ascending=[False, True]).reset_index(drop=True)
        en_iyi_ad = sirali_t.iloc[0]['İlçe'] if not sirali_t.empty else "-"
        en_iyi_oran = float(sirali_t.iloc[0][f'{t} Oranı']) if not sirali_t.empty else 0.0
        en_dusuk_ad = sirali_t.iloc[-1]['İlçe'] if not sirali_t.empty else "-"
        en_dusuk_oran = float(sirali_t.iloc[-1][f'{t} Oranı']) if not sirali_t.empty else 0.0

        ozet_kartlar[t] = {
            'top_ger': f"{top_ger:,}".replace(",", "."),
            'top_yap': f"{top_yap:,}".replace(",", "."),
            'il_oran_html': format_oran_exec(il_oran),
            'gecen': gecen,
            'kalan': kalan,
            'gecen_yuzde': gecen_yuzde,
            'kalan_yuzde': kalan_yuzde,
            'en_iyi': f"{en_iyi_ad} (%{en_iyi_oran:.2f})",
            'en_dusuk': f"{en_dusuk_ad} (%{en_dusuk_oran:.2f})",
        }

    if logo_b64:
        ust_baslik_html = f"""
        <table class="top-header-table">
            <tr>
                <td width="18%" style="border: none;"></td>
                <td width="64%" style="border: none; text-align: center; vertical-align: middle;">
                    <h1 style="margin: 0 0 4px 0;">İstanbul İl Geneli KAN-TAR Yönetici Özeti Raporu</h1>
                    <div class="header-info" style="margin: 0;"><b>Kapsanan Aylar:</b> {aylar_str} | <b>Değerlendirilen İlçe:</b> {toplam_ilce_sayisi}</div>
                </td>
                <td width="18%" style="border: none; text-align: right; vertical-align: top; padding: 0;">
                    <img src="data:image/png;base64,{logo_b64}" width="105" />
                </td>
            </tr>
        </table>
        """
    else:
        ust_baslik_html = f"""
        <h1>İstanbul İl Geneli KAN-TAR Yönetici Özeti Raporu</h1>
        <div class="header-info"><b>Kapsanan Aylar:</b> {aylar_str} | <b>Değerlendirilen İlçe:</b> {toplam_ilce_sayisi}</div>
        """

    baraj_chart_b64 = create_istanbul_baraj_chart_base64(ilce_grup)
    baraj_chart_html = (
        f'<div class="chart-box"><img src="data:image/png;base64,{baraj_chart_b64}" width="525" /></div>'
        if baraj_chart_b64 else ""
    )

    ranking_charts_html = ""
    for tarama_turu in ['Mamografi', 'Serviks', 'Kolorektal']:
        r_b64 = create_istanbul_ilce_ranking_chart_base64(ilce_grup, tarama_turu)
        if r_b64:
            ranking_charts_html += f"""
            <div class="chart-box">
                <img src="data:image/png;base64,{r_b64}" width="530" />
            </div>
            """

    tablo_satirlari_html = ""
    for _, r in ilce_grup.iterrows():
        m_ger = f"{safe_int(r['Mamografi Gereken']):,}".replace(",", ".")
        m_yap = f"{safe_int(r['Mamografi Yapılan']):,}".replace(",", ".")
        s_ger = f"{safe_int(r['Serviks Gereken']):,}".replace(",", ".")
        s_yap = f"{safe_int(r['Serviks Yapılan']):,}".replace(",", ".")
        k_ger = f"{safe_int(r['Kolorektal Gereken']):,}".replace(",", ".")
        k_yap = f"{safe_int(r['Kolorektal Yapılan']):,}".replace(",", ".")

        gecilen_sayi = safe_int(r['Gecilen Baraj Sayisi'])
        baraj_rozet_renk = "#1e8449" if gecilen_sayi == 3 else ("#d68910" if gecilen_sayi >= 1 else "#c0392b")

        tablo_satirlari_html += f"""
        <tr>
            <td><b>{safe_int(r['Genel Sıra'])}.</b></td>
            <td class="left-align"><b>{r['İlçe']}</b></td>
            <td>{format_oran_exec(r['Mamografi Oranı'])}</td>
            <td><b>{safe_int(r['Mamografi Sıra'])}.</b></td>
            <td style="font-size: 7.2px; color: #555;">{m_yap} / {m_ger}</td>
            <td>{format_oran_exec(r['Serviks Oranı'])}</td>
            <td><b>{safe_int(r['Serviks Sıra'])}.</b></td>
            <td style="font-size: 7.2px; color: #555;">{s_yap} / {s_ger}</td>
            <td>{format_oran_exec(r['Kolorektal Oranı'])}</td>
            <td><b>{safe_int(r['Kolorektal Sıra'])}.</b></td>
            <td style="font-size: 7.2px; color: #555;">{k_yap} / {k_ger}</td>
            <td>{format_oran_exec(r['Ortalama Oran'])}</td>
            <td style="color: {baraj_rozet_renk}; font-weight: bold;">{gecilen_sayi} / 3</td>
        </tr>
        """

    html_content = f"""
    <!DOCTYPE html>
    <html lang="tr">
    <head>
    <meta charset="UTF-8">
    <style>
        {font_face_css}
        @page {{ size: A4; margin: 0.9cm; }}
        body {{ font-family: {font_family}; font-size: 9.5px; color: #333; }}
        .top-header-table {{ width: 100%; border-collapse: collapse; margin-bottom: 10px; }}
        h1 {{ text-align: center; color: #2c3e50; font-size: 16px; margin-bottom: 4px; font-family: {font_family}; }}
        .header-info {{ text-align: center; font-size: 10.5px; margin-bottom: 14px; color: #7f8c8d; font-family: {font_family}; }}
        h2 {{ color: #2c3e50; font-size: 12.5px; border-bottom: 1.5px solid #34495e; padding-bottom: 3px; margin-top: 14px; margin-bottom: 8px; font-family: {font_family}; }}
        table {{ border-collapse: collapse; width: 100%; margin-bottom: 8px; }}
        th, td {{ border: 1px solid #bdc3c7; padding: 3.5px 3px; font-size: 8px; font-family: {font_family}; text-align: center; }}
        th {{ background-color: #ecf0f1; font-weight: bold; }}
        .left-align {{ text-align: left; padding-left: 5px; }}
        .summary-table th {{ background-color: #2c3e50; color: white; font-size: 10.5px; padding: 6px; }}
        .summary-table td {{ font-size: 9.5px; padding: 7px; line-height: 1.45; background-color: #f9fbfc; }}
        .chart-box {{ text-align: center; margin-top: 6px; margin-bottom: 10px; }}
        .legend-note {{ font-size: 8px; color: #555; margin-bottom: 6px; }}
    </style>
    </head>
    <body>
        {ust_baslik_html}

        <h2>1. İstanbul İl Geneli Performans ve %70 Hedef Barajı Özeti</h2>
        <table class="summary-table">
            <tr>
                <th width="33.3%">Mamografi Taraması</th>
                <th width="33.3%">Serviks Taraması</th>
                <th width="33.3%">Kolorektal Taraması</th>
            </tr>
            <tr>
                <td>
                    <div style="font-size: 12.5px; margin-bottom: 3px;">İl Geneli Oran: {ozet_kartlar['Mamografi']['il_oran_html']}</div>
                    <div>Gereken: <b>{ozet_kartlar['Mamografi']['top_ger']}</b> | Yapılan: <b>{ozet_kartlar['Mamografi']['top_yap']}</b></div>
                    <div style="margin-top: 4px; color: #1e8449;"><b>%70 Barajını Geçen:</b> {ozet_kartlar['Mamografi']['gecen']} / {toplam_ilce_sayisi} İlçe (%{ozet_kartlar['Mamografi']['gecen_yuzde']:.1f})</div>
                    <div style="color: #c0392b;"><b>%70 Altında Kalan:</b> {ozet_kartlar['Mamografi']['kalan']} / {toplam_ilce_sayisi} İlçe (%{ozet_kartlar['Mamografi']['kalan_yuzde']:.1f})</div>
                    <div style="margin-top: 4px; font-size: 8px; color: #555;">En Yüksek: <b>{ozet_kartlar['Mamografi']['en_iyi']}</b><br>En Düşük: <b>{ozet_kartlar['Mamografi']['en_dusuk']}</b></div>
                </td>
                <td>
                    <div style="font-size: 12.5px; margin-bottom: 3px;">İl Geneli Oran: {ozet_kartlar['Serviks']['il_oran_html']}</div>
                    <div>Gereken: <b>{ozet_kartlar['Serviks']['top_ger']}</b> | Yapılan: <b>{ozet_kartlar['Serviks']['top_yap']}</b></div>
                    <div style="margin-top: 4px; color: #1e8449;"><b>%70 Barajını Geçen:</b> {ozet_kartlar['Serviks']['gecen']} / {toplam_ilce_sayisi} İlçe (%{ozet_kartlar['Serviks']['gecen_yuzde']:.1f})</div>
                    <div style="color: #c0392b;"><b>%70 Altında Kalan:</b> {ozet_kartlar['Serviks']['kalan']} / {toplam_ilce_sayisi} İlçe (%{ozet_kartlar['Serviks']['kalan_yuzde']:.1f})</div>
                    <div style="margin-top: 4px; font-size: 8px; color: #555;">En Yüksek: <b>{ozet_kartlar['Serviks']['en_iyi']}</b><br>En Düşük: <b>{ozet_kartlar['Serviks']['en_dusuk']}</b></div>
                </td>
                <td>
                    <div style="font-size: 12.5px; margin-bottom: 3px;">İl Geneli Oran: {ozet_kartlar['Kolorektal']['il_oran_html']}</div>
                    <div>Gereken: <b>{ozet_kartlar['Kolorektal']['top_ger']}</b> | Yapılan: <b>{ozet_kartlar['Kolorektal']['top_yap']}</b></div>
                    <div style="margin-top: 4px; color: #1e8449;"><b>%70 Barajını Geçen:</b> {ozet_kartlar['Kolorektal']['gecen']} / {toplam_ilce_sayisi} İlçe (%{ozet_kartlar['Kolorektal']['gecen_yuzde']:.1f})</div>
                    <div style="color: #c0392b;"><b>%70 Altında Kalan:</b> {ozet_kartlar['Kolorektal']['kalan']} / {toplam_ilce_sayisi} İlçe (%{ozet_kartlar['Kolorektal']['kalan_yuzde']:.1f})</div>
                    <div style="margin-top: 4px; font-size: 8px; color: #555;">En Yüksek: <b>{ozet_kartlar['Kolorektal']['en_iyi']}</b><br>En Düşük: <b>{ozet_kartlar['Kolorektal']['en_dusuk']}</b></div>
                </td>
            </tr>
        </table>

        {baraj_chart_html}

        <h2>2. Tarama Türlerine Göre Tüm İlçelerin İl Sıralaması Grafikleri</h2>
        {ranking_charts_html}

        <h2>3. Tüm İlçelerin Karşılaştırmalı Başarı Oranları ve İl Sıralamaları</h2>
        <div class="legend-note">
            <b>Renk Göstergesi:</b> <span style="color: #1e8449; font-weight: bold;">Yeşil (%70 ve üzeri - Bakanlık Hedefi Üstü)</span> |
            <b>Siyah (%50 - %69,99)</b> |
            <span style="color: #c0392b; font-weight: bold;">Kırmızı (%50 altı - Kritik Uyarı)</span>
        </div>
        <table>
            <tr>
                <th rowspan="2" width="5%">Genel<br>Sıra</th>
                <th rowspan="2" class="left-align" width="13%">İlçe</th>
                <th colspan="3" width="22%">Mamografi</th>
                <th colspan="3" width="22%">Serviks</th>
                <th colspan="3" width="22%">Kolorektal</th>
                <th rowspan="2" width="8%">3 Tarama<br>Ort. (%)</th>
                <th rowspan="2" width="8%">%70 Geçilen<br>Tarama</th>
            </tr>
            <tr>
                <th width="7%">Oran</th>
                <th width="5%">Sıra</th>
                <th width="10%">Yapılan/Ger.</th>
                <th width="7%">Oran</th>
                <th width="5%">Sıra</th>
                <th width="10%">Yapılan/Ger.</th>
                <th width="7%">Oran</th>
                <th width="5%">Sıra</th>
                <th width="10%">Yapılan/Ger.</th>
            </tr>
            {tablo_satirlari_html}
        </table>
    </body>
    </html>
    """

    pdf_buffer = io.BytesIO()
    pisa.CreatePDF(io.BytesIO(html_content.encode('utf-8')), dest=pdf_buffer, encoding='utf-8')
    return pdf_buffer.getvalue()


def get_ilce_rapor_pdf(df, secilen_ilce, secilen_aylar, logo_b64=None):
    fonts_loaded, font_face_css = setup_fonts()
    font_family = '"DejaVuSans", sans-serif' if fonts_loaded else "Helvetica, Arial, sans-serif"
    
    sirali_aylar = sort_month_labels(secilen_aylar)
    aylar_str = ", ".join(sirali_aylar)
    
    ilce_grup = df.groupby('İlçe')[['Mamografi Gereken', 'Mamografi Yapılan', 'Serviks Gereken', 'Serviks Yapılan', 'Kolorektal Gereken', 'Kolorektal Yapılan']].sum().reset_index()
    for t in ['Mamografi', 'Serviks', 'Kolorektal']:
        ilce_grup[f'{t} Oranı'] = (
            (ilce_grup[f'{t} Yapılan'] / ilce_grup[f'{t} Gereken'] * 100)
            .replace([float('inf'), float('-inf')], 0)
            .fillna(0)
        )
        ilce_grup[f'{t} Sıra'] = ilce_grup[f'{t} Oranı'].rank(ascending=False, method='min').fillna(0).astype(int)
        
    ilce_ozet = ilce_grup[ilce_grup['İlçe'] == secilen_ilce].iloc[0]
    ilce_detay = df[df['İlçe'] == secilen_ilce]
    
    def format_oran(oran):
        if pd.isna(oran) or oran in (float("inf"), float("-inf")):
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
                <td>{format_oran(ilce_ozet['Mamografi Oranı'])} <br><span style='font-size: 9px; color: #555;'>İstanbul Sırası: {safe_int(ilce_ozet['Mamografi Sıra'])}.</span></td>
                <td>{format_oran(ilce_ozet['Serviks Oranı'])} <br><span style='font-size: 9px; color: #555;'>İstanbul Sırası: {safe_int(ilce_ozet['Serviks Sıra'])}.</span></td>
                <td>{format_oran(ilce_ozet['Kolorektal Oranı'])} <br><span style='font-size: 9px; color: #555;'>İstanbul Sırası: {safe_int(ilce_ozet['Kolorektal Sıra'])}.</span></td>
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
        asm_m_ger = safe_int(asm_row['Mamografi Gereken'])
        asm_m_yap = safe_int(asm_row['Mamografi Yapılan'])
        asm_s_ger = safe_int(asm_row['Serviks Gereken'])
        asm_s_yap = safe_int(asm_row['Serviks Yapılan'])
        asm_k_ger = safe_int(asm_row['Kolorektal Gereken'])
        asm_k_yap = safe_int(asm_row['Kolorektal Yapılan'])

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
            m_ger = safe_int(birim_row['Mamografi Gereken'])
            m_yap = safe_int(birim_row['Mamografi Yapılan'])
            s_ger = safe_int(birim_row['Serviks Gereken'])
            s_yap = safe_int(birim_row['Serviks Yapılan'])
            k_ger = safe_int(birim_row['Kolorektal Gereken'])
            k_yap = safe_int(birim_row['Kolorektal Yapılan'])

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

asm_mapping, ilce_mapping, aktif_eslestirme_adi = get_active_asm_mapping()

st.sidebar.header("1. Veri Yükleme")
yuklenen_dosyalar = st.sidebar.file_uploader("KAN-TAR Excel Dosyalarını Yükleyin", type=["xlsx", "xls"], accept_multiple_files=True)
yuklenen_logo = st.sidebar.file_uploader("İlçe Raporu Logosu (Opsiyonel)", type=["png", "jpg", "jpeg"])
kantar_logo_b64 = get_kantar_logo_base64(yuklenen_logo)

if yuklenen_dosyalar:
    try:
        df_ham = load_data(yuklenen_dosyalar, asm_mapping, ilce_mapping)
        
        st.sidebar.header("2. Filtreleme Seçenekleri")
        
        with st.sidebar.form(key="filtre_form"):
            ay_secenekleri = sort_month_labels(df_ham['Dosya (Ay)'].unique().tolist())
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
                df[f'{t} Oranı'] = (
                    (df[f'{t} Yapılan'] / df[f'{t} Gereken'] * 100)
                    .replace([float('inf'), float('-inf')], 0)
                    .fillna(0)
                )

            eslesmeyen_sayisi = (
                safe_int(df_ham[df_ham['ASM Adı'] == "Bilinmeyen/Eksik ASM"]['Birim Adı'].nunique())
                if asm_mapping is not None else 0
            )
            tab7_baslik = (
                f"⚠️ Eşleşmeyen AHB ({eslesmeyen_sayisi})"
                if eslesmeyen_sayisi > 0
                else "🔍 Eşleşmeyen / Yeni AHB"
            )

            tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
                "🎀 Mamografi Taraması", 
                "🎗️ Serviks Taraması", 
                "🧬 Kolorektal Taraması", 
                "📑 İlçe Bazlı Rapor", 
                "📈 Aylık Trend",
                "🚨 Acil Müdahale (Sıfır Tarama)",
                tab7_baslik
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
                ilce_grup['Genel Oran (%)'] = (
                    (ilce_grup[yapilan_kolon] / ilce_grup[gereken_kolon] * 100)
                    .replace([float('inf'), float('-inf')], 0)
                    .fillna(0)
                )
                ilce_grup_sirali = ilce_grup.sort_values(by='Genel Oran (%)', ascending=False).reset_index(drop=True)
                ilce_grup_sirali.index += 1
                
                if secilen_ilce == "Tüm İlçeler" and not ilce_grup_sirali.empty:
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
                
                styler_ilce = ilce_grup_sirali.style.format({'Genel Oran (%)': "{:.2f}%", gereken_kolon: "{:.0f}", yapilan_kolon: "{:.0f}"})
                styler_ilce = style_gradient_safe(styler_ilce, 'Genel Oran (%)', cmap_type='Blues')
                st.dataframe(styler_ilce, use_container_width=True)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button("📥 Excel İndir (İlçe)", data=get_excel(ilce_grup_sirali), file_name=f"{tarama_turu}_Ilce_Siralama.xlsx", key=f"{tarama_turu}_ilce_ex")
                with col2:
                    st.download_button("📄 PDF İndir (İlçe)", data=get_pdf(ilce_grup_pdf, f"{ilce_baslik} - {tarama_turu} İlçe Sıralaması"), file_name=f"{tarama_turu}_Ilce_Siralama.pdf", mime="application/pdf", key=f"{tarama_turu}_ilce_pdf")
                
                st.markdown("---")

                st.subheader(f"🏥 ASM (Aile Sağlığı Merkezi) Başarı Sıralaması ({ilce_baslik} | %{min_oran} - %{max_oran} Arası)")
                asm_grup = filtrelenmis_df.groupby(['İlçe', 'ASM Adı'])[[gereken_kolon, yapilan_kolon]].sum().reset_index()
                asm_grup[hedef_kolon] = (
                    (asm_grup[yapilan_kolon] / asm_grup[gereken_kolon] * 100)
                    .replace([float('inf'), float('-inf')], 0)
                    .fillna(0)
                )
                
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
                
                styler_asm = asm_grup_sirali.style.format({hedef_kolon: "{:.2f}%", gereken_kolon: "{:.0f}", yapilan_kolon: "{:.0f}"})
                styler_asm = style_gradient_safe(styler_asm, hedef_kolon, cmap_type='Blues')
                st.dataframe(styler_asm, use_container_width=True)
                
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
                sirali_ay_isimleri = sort_month_labels(secilen_aylar)
                logo_hash = len(kantar_logo_b64) if kantar_logo_b64 else 0

                def render_istanbul_yonetici_ozeti_bolumu(key_prefix="main"):
                    st.subheader("🏛️ İstanbul İl Geneli Yönetici Özeti Raporu (PDF)")
                    st.markdown(
                        f"Seçili **{', '.join(sirali_ay_isimleri)}** dönemleri kapsamında İstanbul genelindeki tüm ilçelerin "
                        "**3 tarama türündeki (Mamografi, Serviks, Kolorektal) başarı oranlarını**, **il sıralamalarını** "
                        "ve **%70 Bakanlık hedef barajını geçen / altında kalan ilçe sayılarını** grafikler ve karşılaştırmalı "
                        "özet tablolarla tek bir yönetici özeti PDF dosyası olarak indirebilirsiniz."
                    )
                    with st.spinner("İstanbul İl Geneli Yönetici Özeti PDF hazırlanıyor..."):
                        il_ozet_pdf = get_istanbul_yonetici_ozeti_pdf(df, secilen_aylar, logo_b64=kantar_logo_b64)

                    st.download_button(
                        label="📥 İstanbul İl Geneli Yönetici Özeti İndir (PDF)",
                        data=il_ozet_pdf,
                        file_name="Istanbul_Il_Geneli_KANTAR_Yonetici_Ozeti.pdf",
                        mime="application/pdf",
                        key=f"istanbul_yonetici_ozeti_pdf_btn_{key_prefix}"
                    )

                if secilen_ilce == "Tüm İlçeler":
                    render_istanbul_yonetici_ozeti_bolumu(key_prefix="all")

                    st.markdown("---")
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

                    zip_cache_key = f"toplu_zip_v7_{'_'.join(sirali_ay_isimleri)}_{len(mevcut_ilceler)}_{safe_int(df['Mamografi Yapılan'].sum())}_{logo_hash}"

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
                    render_istanbul_yonetici_ozeti_bolumu(key_prefix="single")

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
                    
                    month_order_map_trend = build_month_order_map(trend_grup['Dosya (Ay)'].tolist())
                    trend_grup['Ay_Sira'] = trend_grup['Dosya (Ay)'].map(month_order_map_trend)
                    trend_grup = trend_grup.sort_values('Ay_Sira').drop('Ay_Sira', axis=1).reset_index(drop=True)
                    
                    for t in ['Mamografi', 'Serviks', 'Kolorektal']:
                        trend_grup[f'{t} Oranı (%)'] = (
                            (trend_grup[f'{t} Yapılan'] / trend_grup[f'{t} Gereken'] * 100)
                            .replace([float('inf'), float('-inf')], 0)
                            .fillna(0)
                        )
                    
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

                def clean_month_name_acil(filename):
                    return str(filename).replace('.xlsx', '').replace('.xls', '').replace('Birim Bazında', '').strip()

                def acil_tarama_tablosu_goster(tarama_turu):
                    yapilan_col = f"{tarama_turu} Yapılan"
                    gereken_col = f"{tarama_turu} Gereken"

                    df_zero = acil_df_base[acil_df_base[yapilan_col] == 0].copy()

                    if df_zero.empty:
                        st.success(f"🎉 Tebrikler! Seçilen kriterlerde **{tarama_turu}** taraması sıfır olan birim bulunmuyor.")
                        return

                    month_order_map_acil = build_month_order_map(df_zero['Dosya (Ay)'].tolist())
                    df_zero['Ay_Sira'] = df_zero['Dosya (Ay)'].map(month_order_map_acil)
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

                    styler_acil = df_birim_ozet.style.format({
                        'Sıfır Olduğu Ay Sayısı': "{:.0f}",
                        'Toplam Gereken Hedef': "{:.0f}"
                    })
                    styler_acil = style_gradient_safe(styler_acil, 'Sıfır Olduğu Ay Sayısı', cmap_type='Reds')
                    st.dataframe(styler_acil, use_container_width=True)

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

            with tab7:
                st.subheader("🔍 Eşleşmeyen / Yeni AHB Birimleri Kontrolü")
                if asm_mapping is None:
                    st.warning("⚠️ Sistemde geçerli bir `ASM_AHB_Eslestirme*.xlsx` dosyası bulunamadığı için ASM eşleştirme kontrolü yapılamadı.")
                else:
                    st.caption(f"Aktif eşleştirme kaynağı: **{aktif_eslestirme_adi}** ({len(asm_mapping):,} kayıtlı AHB birimi)")
                    eslesmeyen_ham = df_ham[df_ham['ASM Adı'] == "Bilinmeyen/Eksik ASM"].copy()
                    if eslesmeyen_ham.empty:
                        st.success(f"🎉 Tebrikler! Yüklenen dosyalardaki tüm AHB birimleri `{aktif_eslestirme_adi}` listesiyle eksiksiz eşleşti.")
                    else:
                        eslesmeyen_ozet = eslesmeyen_ham.groupby(['İlçe', 'Birim Adı', 'Birim ID'], as_index=False).agg(
                            **{
                                'Görüldüğü Aylar': ('Dosya (Ay)', lambda x: ', '.join(sort_month_labels(list(dict.fromkeys(x)))))
                            }
                        ).sort_values(by=['İlçe', 'Birim Adı']).reset_index(drop=True)
                        eslesmeyen_ozet.index += 1

                        st.warning(
                            f"⚠️ `{aktif_eslestirme_adi}` listesinde ASM karşılığı bulunmayan **{len(eslesmeyen_ozet)}** adet AHB birimi tespit edildi. "
                            "Bu birimler raporlarda **'Bilinmeyen/Eksik ASM'** başlığı altında toplanmıştır."
                        )
                        st.markdown("Eşleştirme tablonuzu güncellemek için aşağıdaki listeyi inceleyebilir veya Excel olarak indirebilirsiniz:")
                        st.dataframe(eslesmeyen_ozet, use_container_width=True)
                        st.download_button(
                            label="📥 Eşleşmeyen AHB Birimlerini İndir (Excel)",
                            data=get_excel(eslesmeyen_ozet),
                            file_name="Eslesmeyen_Yeni_AHB_Birimleri.xlsx",
                            key="eslesmeyen_ahb_excel_btn"
                        )

        else:
            st.warning("Lütfen sol menüden analiz edilecek en az bir ay (dosya) seçin.")

    except Exception as e:
        st.error(f"Dosyalar okunurken bir hata oluştu. Hata detayı: {e}")
else:
    st.info("👈 Lütfen sol menüden analiz edilecek Excel dosyasını (veya dosyalarını) sisteme yükleyin.")
