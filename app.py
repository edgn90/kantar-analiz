import streamlit as st
import pandas as pd

# Sayfa Ayarları
st.set_page_config(page_title="KAN-TAR Analiz Sistemi", layout="wide")

@st.cache_data
def load_data():
    file_path = "Agustos Birim Bazında KAN-TAR.xlsx"
    
    # Excel dosyasının yapısı gereği asıl başlıklar 2. satırda (header=1)
    df = pd.read_excel(file_path, header=1)
    
    # Kolonları isimlendirme
    df = df.rename(columns={
        'Birim Adı': 'Birim Adı',
        'Birim Id': 'Birim ID',
        'Gereken Tarama': 'Mamografi Gereken', 'Oran(%)': 'Mamografi Oranı',
        'Gereken Tarama.1': 'Serviks Gereken', 'Oran(%).1': 'Serviks Oranı',
        'Gereken Tarama.2': 'Kolorektal Gereken', 'Oran(%).2': 'Kolorektal Oranı'
    })
    
    # Eksik verileri temizleme
    df = df.fillna(0)
    
    # İstanbul İlçe Tespiti
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
    
    # İlçe bazlı gerçek oran hesaplaması için 'Yapılan Tarama' sayılarını türetme
    df['Mamografi Yapılan'] = (df['Mamografi Oranı'] * df['Mamografi Gereken'] / 100).round()
    df['Serviks Yapılan'] = (df['Serviks Oranı'] * df['Serviks Gereken'] / 100).round()
    df['Kolorektal Yapılan'] = (df['Kolorektal Oranı'] * df['Kolorektal Gereken'] / 100).round()
    
    return df

# Veriyi Yükle
df = load_data()

st.title("📊 KAN-TAR Analiz Sistemi")

# Filtreler için Yan Menü
st.sidebar.header("Filtreleme Seçenekleri")
tarama_turu = st.sidebar.selectbox(
    "Tarama Türünü Seçin",
    ["Mamografi", "Serviks", "Kolorektal"]
)

# Arayüzü Sekmelere Ayırma
tab1, tab2 = st.tabs(["🏥 İlçe Bazlı Birim Sıralaması", "📍 Toplu İlçe Sıralaması"])

with tab1:
    st.subheader(f"Seçilen İlçe İçindeki Birimlerin {tarama_turu} Sıralaması")
    
    secilen_ilce = st.sidebar.selectbox("İlçe Seçin", sorted(df['İlçe'].unique()))
    
    # Veriyi filtrele ve seçilen orana göre sırala
    ilce_df = df[df['İlçe'] == secilen_ilce].copy()
    hedef_kolon = f"{tarama_turu} Oranı"
    
    ilce_df_sirali = ilce_df.sort_values(by=hedef_kolon, ascending=False)
    
    # Gösterilecek kolonları sadeleştirme
    gosterilecek_df = ilce_df_sirali[[
        'Birim Adı', 'Birim ID', f"{tarama_turu} Gereken", f"{tarama_turu} Yapılan", hedef_kolon
    ]].reset_index(drop=True)
    
    st.dataframe(
        gosterilecek_df.style.format({hedef_kolon: "{:.2f}%", f"{tarama_turu} Gereken": "{:.0f}", f"{tarama_turu} Yapılan": "{:.0f}"}), 
        use_container_width=True
    )

with tab2:
    st.subheader(f"Tüm İlçelerin {tarama_turu} Başarı Sıralaması")
    
    # Seçilen tarama türü için ilçe bazlı toplamları alma
    gereken_kolon = f"{tarama_turu} Gereken"
    yapilan_kolon = f"{tarama_turu} Yapılan"
    
    ilce_grup = df.groupby('İlçe')[[gereken_kolon, yapilan_kolon]].sum().reset_index()
    
    # İlçenin gerçek tarama oranını hesaplama
    ilce_grup['Genel Oran (%)'] = (ilce_grup[yapilan_kolon] / ilce_grup[gereken_kolon] * 100).fillna(0)
    
    # Orana göre azalan şekilde sıralama
    ilce_grup_sirali = ilce_grup.sort_values(by='Genel Oran (%)', ascending=False).reset_index(drop=True)
    ilce_grup_sirali.index += 1  # Sıra numarası 1'den başlasın
    
    st.dataframe(
        ilce_grup_sirali.style.format({'Genel Oran (%)': "{:.2f}%", gereken_kolon: "{:.0f}", yapilan_kolon: "{:.0f}"})
        .background_gradient(subset=['Genel Oran (%)'], cmap='Blues'),
        use_container_width=True
    )
