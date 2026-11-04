import streamlit as st
import pandas as pd
import io
from weasyprint import HTML

# Sayfa Ayarları
st.set_page_config(page_title="KAN-TAR Analiz Sistemi", layout="wide")

@st.cache_data
def load_data(file):
    df = pd.read_excel(file, header=1)
    df = df.rename(columns={
        'Birim Adı': 'Birim Adı', 'Birim Id': 'Birim ID',
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
    df['Mamografi Yapılan'] = (df['Mamografi Oranı'] * df['Mamografi Gereken'] / 100).round()
    df['Serviks Yapılan'] = (df['Serviks Oranı'] * df['Serviks Gereken'] / 100).round()
    df['Kolorektal Yapılan'] = (df['Kolorektal Oranı'] * df['Kolorektal Gereken'] / 100).round()
    
    return df

# Dışa Aktarma Fonksiyonları (Performans için Önbellekli)
@st.cache_data
def get_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sıralama')
    return output.getvalue()

@st.cache_data
def get_pdf(df, title):
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        body {{ font-family: Arial, sans-serif; font-size: 11px; }}
        h2 {{ text-align: center; color: #333; }}
        table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
        th, td {{ border: 1px solid #ddd; padding: 6px; text-align: left; }}
        th {{ background-color: #f4f6f9; }}
    </style>
    </head>
    <body>
        <h2>{title}</h2>
        {df.to_html(index=False)}
    </body>
    </html>
    """
    return HTML(string=html_content).write_pdf()

st.title("📊 KAN-TAR Analiz Sistemi")

st.sidebar.header("1. Veri Yükleme")
yuklenen_dosya = st.sidebar.file_uploader("KAN-TAR Excel Dosyasını Yükleyin", type=["xlsx", "xls"])

if yuklenen_dosya is not None:
    try:
        df = load_data(yuklenen_dosya)
        
        st.sidebar.header("2. Filtreleme Seçenekleri")
        ilce_secenekleri = ["Tüm İlçeler"] + sorted(df['İlçe'].unique().tolist())
        secilen_ilce = st.sidebar.selectbox("İlçe Seçin", ilce_secenekleri)

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
            
            # --- ÜST TABLO: İLÇE SIRALAMASI ---
            st.subheader(f"📍 İlçe Başarı Sıralaması ({ilce_baslik})")
            ilce_grup = filtrelenmis_df.groupby('İlçe')[[gereken_kolon, yapilan_kolon]].sum().reset_index()
            ilce_grup['Genel Oran (%)'] = (ilce_grup[yapilan_kolon] / ilce_grup[gereken_kolon] * 100).fillna(0)
            
            ilce_grup_sirali = ilce_grup.sort_values(by='Genel Oran (%)', ascending=False).reset_index(drop=True)
            ilce_grup_sirali.index += 1
            
            st.dataframe(
                ilce_grup_sirali.style.format({'Genel Oran (%)': "{:.2f}%", gereken_kolon: "{:.0f}", yapilan_kolon: "{:.0f}"})
                .background_gradient(subset=['Genel Oran (%)'], cmap='Blues'),
                use_container_width=True
            )
            
            # İlçe Tablosu İndirme Butonları
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "📥 Excel İndir (İlçe Sıralaması)", 
                    data=get_excel(ilce_grup_sirali), 
                    file_name=f"{tarama_turu}_Ilce_Siralama.xlsx",
                    key=f"{tarama_turu}_ilce_excel"
                )
            with col2:
                st.download_button(
                    "📄 PDF İndir (İlçe Sıralaması)", 
                    data=get_pdf(ilce_grup_sirali, f"{ilce_baslik} - {tarama_turu} İlçe Sıralaması"), 
                    file_name=f"{tarama_turu}_Ilce_Siralama.pdf",
                    mime="application/pdf",
                    key=f"{tarama_turu}_ilce_pdf"
                )
            
            st.markdown("---")

            # --- ALT TABLO: BİRİM SIRALAMASI ---
            st.subheader(f"🏥 Birim Bazlı Sıralama ({ilce_baslik})")
            birim_df_sirali = filtrelenmis_df.sort_values(by=hedef_kolon, ascending=False)
            
            gosterilecek_df = birim_df_sirali[[
                'İlçe', 'Birim Adı', 'Birim ID', gereken_kolon, yapilan_kolon, hedef_kolon
            ]].reset_index(drop=True)
            gosterilecek_df.index += 1
            
            st.dataframe(
                gosterilecek_df.style.format({hedef_kolon: "{:.2f}%", gereken_kolon: "{:.0f}", yapilan_kolon: "{:.0f}"}), 
                use_container_width=True
            )
            
            # Birim Tablosu İndirme Butonları
            col3, col4 = st.columns(2)
            with col3:
                st.download_button(
                    "📥 Excel İndir (Birim Sıralaması)", 
                    data=get_excel(gosterilecek_df), 
                    file_name=f"{tarama_turu}_Birim_Siralama.xlsx",
                    key=f"{tarama_turu}_birim_excel"
                )
            with col4:
                st.download_button(
                    "📄 PDF İndir (Birim Sıralaması)", 
                    data=get_pdf(gosterilecek_df, f"{ilce_baslik} - {tarama_turu} Birim Bazlı Sıralama"), 
                    file_name=f"{tarama_turu}_Birim_Siralama.pdf",
                    mime="application/pdf",
                    key=f"{tarama_turu}_birim_pdf"
                )

        with tab1:
            tablo_goster("Mamografi")
            
        with tab2:
            tablo_goster("Serviks")
            
        with tab3:
            tablo_goster("Kolorektal")

    except Exception as e:
        st.error(f"Dosya okunurken bir hata oluştu. Hata: {e}")
else:
    st.info("👈 Lütfen sol menüden analiz edilecek Excel dosyasını sisteme yükleyin.")
