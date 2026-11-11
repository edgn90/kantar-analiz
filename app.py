import streamlit as st
import pandas as pd
import io
from xhtml2pdf import pisa

# Sayfa Ayarları
st.set_page_config(page_title="KAN-TAR Analiz Sistemi", layout="wide")

@st.cache_data
def load_data(files):
    tum_veriler = []
    
    for file in files:
        xls = pd.ExcelFile(file)
        
        # YENİ FORMAT KONTROLÜ (Ocak 2026 ve Sonrası - Uzun Format)
        if 'Birim Bazında' in xls.sheet_names:
            df_test = pd.read_excel(xls, sheet_name='Birim Bazında', nrows=5)
            # Eğer 'Kronik Hastalik' kolonu varsa bu yeni formattır
            if any('kronik hastalik' in str(c).lower().replace('ı', 'i') for c in df_test.columns):
                df_raw = pd.read_excel(xls, sheet_name='Birim Bazında')
                df_raw.columns = [str(c).strip().lower().replace('ı', 'i') for c in df_raw.columns]
                
                # Uzun formatı geniş formata (Pivot) çeviriyoruz
                df_pivot = df_raw.pivot_table(
                    index=['birim adi'], 
                    columns='kronik hastalik', 
                    values=['aylik gereken', 'yapilan'],
                    aggfunc='sum'
                ).reset_index()
                
                # Kolon isimlerini birleştirip düzleştiriyoruz
                df_pivot.columns = ['_'.join(col).strip('_') for col in df_pivot.columns.values]
                
                temp_df = pd.DataFrame()
                temp_df['Birim Adı'] = df_pivot['birim adi']
                temp_df['Birim ID'] = temp_df['Birim Adı'].str.extract(r'(\d+)\s+Nolu').fillna("Bilinmiyor")
                
                # İlgili kolonları kelime aramasıyla eşleştirme
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
                continue # Bu dosya bitti, diğerine geç
                
        # ESKİ FORMAT KONTROLÜ (Haziran 2026 Öncesi - Geniş Format)
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
        
        # Eski formattan "Yapılan" verisini geri hesaplayıp oluşturma
        for t in ['Mamografi', 'Serviks', 'Kolorektal']:
            gereken = pd.to_numeric(df_old[f'{t} Gereken'], errors='coerce').fillna(0)
            oran = pd.to_numeric(df_old[f'{t} Oranı'], errors='coerce').fillna(0)
            df_old[f'{t} Gereken'] = gereken
            df_old[f'{t} Yapılan'] = (oran * gereken / 100).round()
            
        tum_veriler.append(df_old[['Birim Adı', 'Birim ID', 'Mamografi Gereken', 'Mamografi Yapılan', 'Serviks Gereken', 'Serviks Yapılan', 'Kolorektal Gereken', 'Kolorektal Yapılan']])

    # BİRDEN FAZLA AY YÜKLENDİYSE BİRLEŞTİR VE TOPLA (KÜMÜLATİF)
    combined_df = pd.concat(tum_veriler, ignore_index=True)
    
    agg_funcs = {
        'Birim ID': 'first',
        'Mamografi Gereken': 'sum', 'Mamografi Yapılan': 'sum',
        'Serviks Gereken': 'sum', 'Serviks Yapılan': 'sum',
        'Kolorektal Gereken': 'sum', 'Kolorektal Yapılan': 'sum',
    }
    
    # Aynı isimli AHB'lerin aylık rakamlarını topluyoruz
    final_df = combined_df.groupby('Birim Adı').agg(agg_funcs).reset_index()
    
    # Toplamlar üzerinden gerçek, nihai oranları hesaplıyoruz
    for t in ['Mamografi', 'Serviks', 'Kolorektal']:
        final_df[f'{t} Oranı'] = (final_df[f'{t} Yapılan'] / final_df[f'{t} Gereken'] * 100).fillna(0)
    
    # İlçe Tespiti
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

    final_df['İlçe'] = final_df['Birim Adı'].apply(get_district)
    
    return final_df

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
# YENİ: accept_multiple_files=True eklendi. Artık aynı anda birden fazla ay Excel'i seçilebilir.
yuklenen_dosyalar = st.sidebar.file_uploader("KAN-TAR Excel Dosyalarını Yükleyin (Çoklu Seçim Yapabilirsiniz)", type=["xlsx", "xls"], accept_multiple_files=True)

if yuklenen_dosyalar: # Liste boş değilse
    try:
        df = load_data(yuklenen_dosyalar)
        
        st.sidebar.header("2. Filtreleme Seçenekleri")
        
        with st.sidebar.form(key="filtre_form"):
            ilce_secenekleri = ["Tüm İlçeler"] + sorted(df['İlçe'].unique().tolist())
            secilen_ilce = st.selectbox("İlçe Seçin", ilce_secenekleri)
            
            st.markdown("---")
            
            st.subheader("3. Oran Filtresi (Birimler İçin)")
            min_oran, max_oran = st.slider(
                "Gösterilecek Başarı Oranı Aralığı (%)",
                min_value=0, max_value=100, value=(0, 100), step=1
            )
            
            calistir_butonu = st.form_submit_button(label="Çalıştır 🚀")

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

            st.subheader(f"🏥 Birim Bazlı Sıralama ({ilce_baslik} | %{min_oran} - %{max_oran} Arası)")
            
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
        st.error(f"Dosyalar okunurken bir hata oluştu. Hata detayı: {e}")
else:
    st.info("👈 Lütfen sol menüden analiz edilecek Excel dosyasını (veya dosyalarını) sisteme yükleyin.")
