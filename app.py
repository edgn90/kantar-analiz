import streamlit as st
import pandas as pd

# Sayfa Ayarları
st.set_page_config(page_title="KAN-TAR Analiz Sistemi", layout="wide")

@st.cache_data
def load_data(file):
    df = pd.read_excel(file, header=1)
    
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

st.sidebar.header("1. Veri Yükleme")
yuklenen_dosya = st.sidebar.file_uploader("KAN-TAR Excel Dosyasını Yükleyin", type=["xlsx", "xls"])

if yuklenen_dosya is not None:
    try:
        df = load_data(yuklenen_dosya)
        
        st.sidebar.header("2. Filtreleme Seçenekleri")
        # "Tüm İlçeler" seçeneğini varsayılan olarak başa ekliyoruz
        ilce_secenekleri = ["Tüm İlçeler"] + sorted(df['İlçe'].unique().tolist())
        secilen_ilce = st.sidebar.selectbox("İlçe Seçin", ilce_secenekleri)

        # Tarama türlerine göre 3 ayrı sekme oluşturuyoruz
        tab1, tab2, tab3 = st.tabs(["🎀 Mamografi Taraması", "🎗️ Serviks Taraması", "🧬 Kolorektal Taraması"])

        def tablo_goster(tarama_turu):
            # İlçe "Tüm İlçeler" değilse tabloyu filtrele
            if secilen_ilce != "Tüm İlçeler":
                filtrelenmis_df = df[df['İlçe'] == secilen_ilce].copy()
                ilce_baslik = f"{secilen_ilce} İlçesi"
            else:
                filtrelenmis_df = df.copy()
                ilce_baslik = "Tüm İstanbul Geneli"

            hedef_kolon = f"{tarama_turu} Oranı"
            gereken_kolon = f"{tarama_turu} Gereken"
            yapilan_kolon = f"{tarama_turu} Yapılan"
            
            # ÜST TABLO: İlçe Toplu Başarı Sıralaması
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
            
            st.markdown("---")

            # ALT TABLO: AHB (Birim) Bazlı Sıralama
            st.subheader(f"🏥 Birim Bazlı Sıralama ({ilce_baslik})")
            birim_df_sirali = filtrelenmis_df.sort_values(by=hedef_kolon, ascending=False)
            
            # Filtreye göre birimlerin hangi ilçede olduğunu da gösteriyoruz
            gosterilecek_df = birim_df_sirali[[
                'İlçe', 'Birim Adı', 'Birim ID', gereken_kolon, yapilan_kolon, hedef_kolon
            ]].reset_index(drop=True)
            gosterilecek_df.index += 1
            
            st.dataframe(
                gosterilecek_df.style.format({hedef_kolon: "{:.2f}%", gereken_kolon: "{:.0f}", yapilan_kolon: "{:.0f}"}), 
                use_container_width=True
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
