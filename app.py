import streamlit as st
import pandas as pd

# Sayfa Ayarları
st.set_page_config(page_title="KAN-TAR Analiz Sistemi", layout="wide")

@st.cache_data
def load_data(file):
    # Sabit dosya yolu yerine yüklenen dosyayı okuyoruz
    df = pd.read_excel(file, header=1)
    
    # Kolonları isimlendirme
    df = df.rename(columns={
        'Birim Adı': 'Birim Adı',
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
    
    df['Mamografi Yapılan'] = (df['Mamografi Oranı'] * df['Mamografi Gereken'] / 100).round()
    df['Serviks Yapılan'] = (df['Serviks Oranı'] * df['Serviks Gereken'] / 100).round()
    df['Kolorektal Yapılan'] = (df['Kolorektal Oranı'] * df['Kolorektal Gereken'] / 100).round()
    
    return df

st.title("📊 KAN-TAR Analiz Sistemi")

# Dosya Yükleme Alanı
st.sidebar.header("1. Veri Yükleme")
yuklenen_dosya = st.sidebar.file_uploader("KAN-TAR Excel Dosyasını Yükleyin", type=["xlsx", "xls"])

if yuklenen_dosya is not None:
    # Dosya yüklendiyse veriyi işle ve arayüzü göster
    try:
        df = load_data(yuklenen_dosya)
        
        st.sidebar.header("2. Filtreleme Seçenekleri")
        tarama_turu = st.sidebar.selectbox("Tarama Türünü Seçin", ["Mamografi", "Serviks", "Kolorektal"])

        tab1, tab2 = st.tabs(["🏥 İlçe Bazlı Birim Sıralaması", "📍 Toplu İlçe Sıralaması"])

        with tab1:
            st.subheader(f"Seçilen İlçe İçindeki Birimlerin {tarama_turu} Sıralaması")
            secilen_ilce = st.sidebar.selectbox("İlçe Seçin", sorted(df['İlçe'].unique()))
            
            ilce_df = df[df['İlçe'] == secilen_ilce].copy()
            hedef_kolon = f"{tarama_turu} Oranı"
            
            ilce_df_sirali = ilce_df.sort_values(by=hedef_kolon, ascending=False)
            gosterilecek_df = ilce_df_sirali[[
                'Birim Adı', 'Birim ID', f"{tarama_turu} Gereken", f"{tarama_turu} Yapılan", hedef_kolon
            ]].reset_index(drop=True)
            gosterilecek_df.index += 1
            
            st.dataframe(
                gosterilecek_df.style.format({hedef_kolon: "{:.2f}%", f"{tarama_turu} Gereken": "{:.0f}", f"{tarama_turu} Yapılan": "{:.0f}"}), 
                use_container_width=True
            )

        with tab2:
            st.subheader(f"Tüm İlçelerin {tarama_turu} Başarı Sıralaması")
            gereken_kolon = f"{tarama_turu} Gereken"
            yapilan_kolon = f"{tarama_turu} Yapılan"
            
            ilce_grup = df.groupby('İlçe')[[gereken_kolon, yapilan_kolon]].sum().reset_index()
            ilce_grup['Genel Oran (%)'] = (ilce_grup[yapilan_kolon] / ilce_grup[gereken_kolon] * 100).fillna(0)
            
            ilce_grup_sirali = ilce_grup.sort_values(by='Genel Oran (%)', ascending=False).reset_index(drop=True)
            ilce_grup_sirali.index += 1
            
            st.dataframe(
                ilce_grup_sirali.style.format({'Genel Oran (%)': "{:.2f}%", gereken_kolon: "{:.0f}", yapilan_kolon: "{:.0f}"})
                .background_gradient(subset=['Genel Oran (%)'], cmap='Blues'),
                use_container_width=True
            )
            
    except Exception as e:
        st.error(f"Dosya okunurken bir hata oluştu. Lütfen doğru formatta bir dosya yüklediğinize emin olun. Hata: {e}")
else:
    st.info("👈 Lütfen sol menüden analiz edilecek Excel dosyasını sisteme yükleyin.")
