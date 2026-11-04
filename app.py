import streamlit as st
import pandas as pd
import io
from xhtml2pdf import pisa

# Sayfa Ayarları
st.set_page_config(page_title="KAN-TAR Analiz Sistemi", layout="wide")

@st.cache_data
def load_data(file):
    # 1. Başlık satırını otomatik bulma (Tip hatalarına karşı %100 güvenli yöntem)
    df_temp = pd.read_excel(file, header=None, nrows=15)
    header_idx = 1 
    
    for i, row in df_temp.iterrows():
        row_degerleri = [str(hücre).lower() for hücre in row.tolist()]
        row_str = " ".join(row_degerleri)
        
        if "birim" in row_str and "oran" in row_str:
            header_idx = i
            break
            
    df = pd.read_excel(file, header=header_idx)
    
    # 2. Sütun isimlerindeki görünmez boşlukları temizleme
    df.columns = [str(c).strip() for c in df.columns]
    
    # 3. İsim farklılıklarını tolere etme
    sutun_degisiklikleri = {}
    for col in df.columns:
        if col.lower() == "birim adi": sutun_degisiklikleri[col] = "Birim Adı"
        elif col.lower() == "birim id": sutun_degisiklikleri[col] = "Birim ID"
    
    if sutun_degisiklikleri:
        df = df.rename(columns=sutun_degisiklikleri)
        
    df = df.rename(columns={
        'Birim Id': 'Birim ID',
        'Gereken Tarama': 'Mamografi Gereken', 'Oran(%)': 'Mamografi Oranı',
        'Gereken Tarama.1': 'Serviks Gereken', 'Oran(%).1': 'Serviks Oranı',
        'Gereken Tarama.2': 'Kolorektal Gereken', 'Oran(%).2': 'Kolorektal Oranı'
    })
    
    df = df.fillna(0)
    
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

    df['İlçe'] = df['Birim Adı'].apply(get_district)
    
    df['Mamografi Oranı'] = pd.to_numeric(df['Mamografi Oranı'], errors='coerce').fillna(0)
    df['Mamografi Gereken'] = pd.to_numeric(df['Mamografi Gereken'], errors='coerce').fillna(0)
    df['Serviks Oranı'] = pd.to_numeric(df['Serviks Oranı'], errors='coerce').fillna(0)
    df['Serviks Gereken'] = pd.to_numeric(df['Serviks Gereken'], errors='coerce').fillna(0)
    df['Kolorektal Oranı'] = pd.to_numeric(df['Kolorektal Oranı'], errors='coerce').fillna(0)
    df['Kolorektal Gereken'] = pd.to_numeric(df['Kolorektal Gereken'], errors='coerce').fillna(0)
    
    df['Mamografi Yapılan'] = (df['Mamografi Oranı'] * df['Mamografi Gereken'] / 100).round()
    df['Serviks Yapılan'] = (df['Serviks Oranı'] * df['Serviks Gereken'] / 100).round()
    df['Kolorektal Yapılan'] = (df['Kolorektal Oranı'] * df['Kolorektal Gereken'] / 100).round()
    
    return df

@st.cache_data
def get_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sıralama')
    return output.getvalue()

@st.cache_data
def get_pdf(df, title):
    html_content = f"""
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        @page {{ size: A4; margin: 1cm; }}
        body {{ font-family: Helvetica, Arial, sans-serif; font-size: 10px; }}
        h2 {{ text-align: center; color: #333; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 5px; text-align: left; }}
        th {{ background-color: #f4f6f9; }}
    </style>
    </head>
    <body>
        <h2>{title}</h2>
        {df.to_html(index=False)}
    </body>
    </html>
    """
    pdf_buffer = io.BytesIO()
    pisa.CreatePDF(io.StringIO(html_content), dest=pdf_buffer)
    return pdf_buffer.getvalue()

st.title("📊 KAN-TAR Analiz Sistemi")

st.sidebar.header("1. Veri Yükleme")
yuklenen_dosya = st.sidebar.file_uploader("KAN-TAR Excel Dosyasını Yükleyin", type=["xlsx", "xls"])

if yuklenen_dosya is not None:
    try:
        df = load_data(yuklenen_dosya)
        
        st.sidebar.header("2. Filtreleme Seçenekleri")
        
        # İlçe Filtresi
        ilce_secenekleri = ["Tüm İlçeler"] + sorted(df['İlçe'].unique().tolist())
        secilen_ilce = st.sidebar.selectbox("İlçe Seçin", ilce_secenekleri)
        
        st.sidebar.markdown("---")
        
        # YENİ: Oran Filtresi
        st.sidebar.subheader("3. Oran Filtresi (Birimler İçin)")
        min_oran, max_oran = st.sidebar.slider(
            "Gösterilecek Başarı Oranı Aralığı (%)",
            min_value=0, max_value=100, value=(0, 100), step=1
        )

        tab1, tab2, tab3 = st.tabs(["🎀 Mamografi Taraması", "🎗️ Serviks Taraması", "🧬 Kolorektal Taraması"])

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
            
            # --- ÜST TABLO: İLÇE SIRALAMASI (Filtresiz Toplam Başarı) ---
            st.subheader(f"📍 İlçe Başarı Sıralaması ({ilce_baslik})")
            ilce_grup = filtrelenmis_df.groupby('İlçe')[[gereken_kolon, yapilan_kolon]].sum().reset_index()
            ilce_grup['Genel Oran (%)'] = (ilce_grup[yapilan_kolon] / ilce_grup[gereken_kolon] * 100).fillna(0)
            ilce_grup_sirali = ilce_grup.sort_values(by='Genel Oran (%)', ascending=False).reset_index(drop=True)
            ilce_grup_sirali.index += 1
            
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
                st.download_button("📄 PDF İndir (İlçe)", data=get_pdf(ilce_grup_pdf, f"{ilce_baslik} - {tarama_turu} Ilce Siralamasi"), file_name=f"{tarama_turu}_Ilce_Siralama.pdf", mime="application/pdf", key=f"{tarama_turu}_ilce_pdf")
            
            st.markdown("---")

            # --- ALT TABLO: BİRİM SIRALAMASI (Slider Oranına Göre Filtrelenmiş) ---
            st.subheader(f"🏥 Birim Bazlı Sıralama ({ilce_baslik} | %{min_oran} - %{max_oran} Arası)")
            
            # Sadece slider'da seçilen aralıktaki birimleri filtreliyoruz
            birim_df_aralikli = filtrelenmis_df[
                (filtrelenmis_df[hedef_kolon] >= min_oran) & 
                (filtrelenmis_df[hedef_kolon] <= max_oran)
            ]
            
            birim_df_sirali = birim_df_aralikli.sort_values(by=hedef_kolon, ascending=False)
            
            gosterilecek_df = birim_df_sirali[[
                'İlçe', 'Birim Adı', 'Birim ID', gereken_kolon, yapilan_kolon, hedef_kolon
            ]].reset_index(drop=True)
            gosterilecek_df.index += 1
            
            gosterilecek_pdf = gosterilecek_df.copy()
            gosterilecek_pdf[hedef_kolon] = gosterilecek_pdf[hedef_kolon].apply(lambda x: f"%{x:.2f}")
            
            st.dataframe(
                gosterilecek_df.style.format({hedef_kolon: "{:.2f}%", gereken_kolon: "{:.0f}", yapilan_kolon: "{:.0f}"}), 
                use_container_width=True
            )
            
            col3, col4 = st.columns(2)
            with col3:
                st.download_button("📥 Excel İndir (Birim)", data=get_excel(gosterilecek_df), file_name=f"{tarama_turu}_Birim_Siralama.xlsx", key=f"{tarama_turu}_birim_ex")
            with col4:
                st.download_button("📄 PDF İndir (Birim)", data=get_pdf(gosterilecek_pdf, f"{ilce_baslik} - {tarama_turu} Birim Bazli Siralama"), file_name=f"{tarama_turu}_Birim_Siralama.pdf", mime="application/pdf", key=f"{tarama_turu}_birim_pdf")

        with tab1:
            tablo_goster("Mamografi")
        with tab2:
            tablo_goster("Serviks")
        with tab3:
            tablo_goster("Kolorektal")

    except Exception as e:
        st.error(f"Dosya okunurken bir hata oluştu. Lütfen dosyanızın HSYS standart KAN-TAR çıktısı olduğundan emin olun. Hata detayı: {e}")
else:
    st.info("👈 Lütfen sol menüden analiz edilecek Excel dosyasını sisteme yükleyin.")
