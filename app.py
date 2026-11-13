import streamlit as st
import pandas as pd
import io
import os
from xhtml2pdf import pisa

# Sayfa Ayarları
st.set_page_config(page_title="KAN-TAR Analiz Sistemi", layout="wide")

@st.cache_resource
def get_fonts():
    # Linux (Streamlit Cloud) sistemlerinde varsayılan olarak bulunan, Türkçe karakterleri tam destekleyen FreeSans fontları
    regular = "/usr/share/fonts/truetype/freefont/FreeSans.ttf"
    bold = "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf"
    
    # Eğer sistem Linux değilse (örneğin kendi Windows bilgisayarınızda çalıştırıyorsanız) Windows fontlarını kullanır
    if not os.path.exists(regular):
        regular = "C:/Windows/Fonts/arial.ttf"
        bold = "C:/Windows/Fonts/arialbd.ttf"
        
    return regular, bold

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

@st.cache_data
def get_pdf(df, title):
    reg_font, bold_font = get_fonts()
    html_content = f"""
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        @font-face {{ font-family: 'TrFont'; src: url('{reg_font}'); }}
        @font-face {{ font-family: 'TrFont'; src: url('{bold_font}'); font-weight: bold; }}
        @page {{ size: A4; margin: 1cm; }}
        body {{ font-family: 'TrFont', Helvetica, Arial, sans-serif; font-size: 10px; }}
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

def get_ilce_rapor_pdf(df, secilen_ilce, secilen_aylar):
    reg_font, bold_font = get_fonts()
    aylar_str = ", ".join(secilen_aylar)
    
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

    html_content = f"""
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        @font-face {{ font-family: 'TrFont'; src: url('{reg_font}'); }}
        @font-face {{ font-family: 'TrFont'; src: url('{bold_font}'); font-weight: bold; }}
        @page {{ size: A4; margin: 1cm; }}
        body {{ font-family: 'TrFont', Helvetica, Arial, sans-serif; font-size: 10px; color: #333; }}
        h1 {{ text-align: center; color: #2c3e50; font-size: 16px; margin-bottom: 5px; }}
        .header-info {{ text-align: center; font-size: 11px; margin-bottom: 20px; color: #7f8c8d; }}
        h2 {{ color: #34495e; font-size: 13px; border-bottom: 1px solid #ccc; padding-bottom: 3px; margin-top: 15px; }}
        h3 {{ color: #16a085; font-size: 11px; margin-top: 10px; margin-bottom: 5px; }}
        table {{ border-collapse: collapse; width: 100%; margin-bottom: 5px; }}
        th, td {{ border: 1px solid #bdc3c7; padding: 4px; font-size: 9px; }}
        th {{ background-color: #ecf0f1; font-weight: bold; text-align: center; }}
        td {{ text-align: center; }}
        .left-align {{ text-align: left; }}
        .summary-table th {{ background-color: #34495e; color: white; }}
        .summary-table td {{ font-size: 11px; padding: 8px; }}
    </style>
    </head>
    <body>
        <h1>{secilen_ilce} İlçesi KAN-TAR Değerlendirme Raporu</h1>
        <div class="header-info"><b>Kapsanan Aylar:</b> {aylar_str}</div>
        
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
        
        <h2>2. ASM ve Birim (AHB) Bazlı Kırılımlar</h2>
    """
    
    asm_grup = ilce_detay.groupby('ASM Adı')[['Mamografi Gereken', 'Mamografi Yapılan', 'Serviks Gereken', 'Serviks Yapılan', 'Kolorektal Gereken', 'Kolorektal Yapılan']].sum().reset_index()
    asm_grup = asm_grup.sort_values(by='ASM Adı')
    
    for _, asm_row in asm_grup.iterrows():
        asm_adi = asm_row['ASM Adı']
        m_oran = (asm_row['Mamografi Yapılan'] / asm_row['Mamografi Gereken'] * 100) if asm_row['Mamografi Gereken'] > 0 else 0
        s_oran = (asm_row['Serviks Yapılan'] / asm_row['Serviks Gereken'] * 100) if asm_row['Serviks Gereken'] > 0 else 0
        k_oran = (asm_row['Kolorektal Yapılan'] / asm_row['Kolorektal Gereken'] * 100) if asm_row['Kolorektal Gereken'] > 0 else 0
        
        html_content += f"<h3>🏥 {asm_adi}</h3>"
        html_content += f"""
        <table style="background-color: #f8f9fa;">
            <tr>
                <td width="33%"><b>ASM Ort. Mamografi:</b> {format_oran(m_oran)}</td>
                <td width="33%"><b>ASM Ort. Serviks:</b> {format_oran(s_oran)}</td>
                <td width="33%"><b>ASM Ort. Kolorektal:</b> {format_oran(k_oran)}</td>
            </tr>
        </table>
        <table>
            <tr>
                <th class="left-align" width="40%">Birim Adı</th>
                <th width="20%">Mamografi (%)</th>
                <th width="20%">Serviks (%)</th>
                <th width="20%">Kolorektal (%)</th>
            </tr>
        """
        
        asm_birimler = ilce_detay[ilce_detay['ASM Adı'] == asm_adi].sort_values(by='Birim Adı')
        for _, birim_row in asm_birimler.iterrows():
            html_content += "<tr>"
            html_content += f"<td class='left-align'>{birim_row['Birim Adı']}</td>"
            html_content += f"<td>{format_oran(birim_row['Mamografi Oranı'])}</td>"
            html_content += f"<td>{format_oran(birim_row['Serviks Oranı'])}</td>"
            html_content += f"<td>{format_oran(birim_row['Kolorektal Oranı'])}</td>"
            html_content += "</tr>"
        html_content += "</table>"
        
    html_content += "</body></html>"
    
    pdf_buffer = io.BytesIO()
    pisa.CreatePDF(io.StringIO(html_content), dest=pdf_buffer)
    return pdf_buffer.getvalue()

st.title("📊 KAN-TAR Analiz Sistemi")

asm_mapping = load_asm_mapping()

st.sidebar.header("1. Veri Yükleme")
yuklenen_dosyalar = st.sidebar.file_uploader("KAN-TAR Excel Dosyalarını Yükleyin", type=["xlsx", "xls"], accept_multiple_files=True)

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

            tab1, tab2, tab3, tab4 = st.tabs(["🎀 Mamografi Taraması", "🎗️ Serviks Taraması", "🧬 Kolorektal Taraması", "📑 İlçe Bazlı Rapor"])

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

                st.subheader(f"🏥 ASM (Aile Sağlığı Merkezi) Başarı Sıralaması ({ilce_baslik} | %{min_oran} - %{max_oran} Arası)")
                asm_grup = filtrelenmis_df.groupby(['İlçe', 'ASM Adı'])[[gereken_kolon, yapilan_kolon]].sum().reset_index()
                asm_grup[hedef_kolon] = (asm_grup[yapilan_kolon] / asm_grup[gereken_kolon] * 100).fillna(0)
                
                asm_grup_aralikli = asm_grup[(asm_grup[hedef_kolon] >= min_oran) & (asm_grup[hedef_kolon] <= max_oran)]
                asm_grup_sirali = asm_grup_aralikli.sort_values(by=hedef_kolon, ascending=False).reset_index(drop=True)
                asm_grup_sirali.index += 1
                
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
                    st.download_button("📄 PDF İndir (ASM)", data=get_pdf(asm_grup_pdf, f"{ilce_baslik} - {tarama_turu} ASM Siralamasi"), file_name=f"{tarama_turu}_ASM_Siralama.pdf", mime="application/pdf", key=f"{tarama_turu}_asm_pdf")
                
                st.markdown("---")

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
                    st.download_button("📄 PDF İndir (Birim)", data=get_pdf(gosterilecek_pdf, f"{ilce_baslik} - {tarama_turu} Birim Bazli Siralama"), file_name=f"{tarama_turu}_Birim_Siralama.pdf", mime="application/pdf", key=f"{tarama_turu}_birim_pdf")

            with tab1:
                tablo_goster("Mamografi")
            with tab2:
                tablo_goster("Serviks")
            with tab3:
                tablo_goster("Kolorektal")
                
            with tab4:
                if secilen_ilce == "Tüm İlçeler":
                    st.info("👈 İlçe bazlı PDF raporu oluşturabilmek için lütfen sol menüden belirli bir ilçeyi seçin ve Çalıştır'a basın.")
                else:
                    st.subheader(f"📑 {secilen_ilce} İlçesi Özel Değerlendirme Raporu")
                    st.markdown(f"Bu rapor **{', '.join(secilen_aylar)}** dönemlerini kapsamaktadır.")
                    st.markdown("Rapor İçeriği:")
                    st.markdown("- İlçenin İstanbul genelindeki başarı sıralaması ve 3 Kanser türündeki özet oranı")
                    st.markdown("- İlçeye bağlı tüm Aile Sağlığı Merkezlerinin (ASM) ortalama oranları")
                    st.markdown("- Her ASM'nin altına hiyerarşik olarak dizilmiş Aile Hekimliği Birim (AHB) detay oranları")
                    st.markdown("- **%50'nin altında kalan oranlar için kırmızı uyarı vurgusu** ⚠️")
                    
                    pdf_data = get_ilce_rapor_pdf(df, secilen_ilce, secilen_aylar)
                    
                    st.download_button(
                        label=f"📥 {secilen_ilce} Raporunu İndir (PDF)",
                        data=pdf_data,
                        file_name=f"{secilen_ilce}_Ilce_KANTAR_Raporu.pdf",
                        mime="application/pdf"
                    )
        else:
            st.warning("Lütfen sol menüden analiz edilecek en az bir ay (dosya) seçin.")

    except Exception as e:
        st.error(f"Dosyalar okunurken bir hata oluştu. Hata detayı: {e}")
else:
    st.info("👈 Lütfen sol menüden analiz edilecek Excel dosyasını (veya dosyalarını) sisteme yükleyin.")
