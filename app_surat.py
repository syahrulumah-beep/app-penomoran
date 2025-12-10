import streamlit as st
import pandas as pd
from datetime import datetime, date
import calendar
import os
from fpdf import FPDF
import io

# Konfigurasi Halaman
st.set_page_config(
    page_title="Sistem Penomoran Klinik Utama Rawat Inap Parung",
    layout="wide",
    initial_sidebar_state="collapsed"
)

DB_FILE = 'data_surat.csv'

# Inisialisasi Session State
if 'last_saved' not in st.session_state:
    st.session_state.last_saved = {}

def load_data():
    required_columns = ["No", "Jenis", "Tanggal", "Bulan", "Tahun", "Kode_Klasifikasi", "Kepada", "Perihal", "Keterangan", "Nomor_Surat"]
    
    if not os.path.exists(DB_FILE):
        return pd.DataFrame(columns=required_columns)
    
    df = pd.read_csv(DB_FILE)
    df['Tanggal'] = pd.to_datetime(df['Tanggal']).dt.date
    return df

def save_data(df):
    df.to_csv(DB_FILE, index=False)

def get_next_number(df, tanggal, jenis_surat):
    tahun = tanggal.year
    bulan = tanggal.month
    
    df_jenis = df[df['Jenis'] == jenis_surat]
    df_tahun = df_jenis[df_jenis['Tahun'] == tahun]
    
    if len(df_tahun) == 0:
        return 1
    
    current_max = df_tahun["No"].max()
    
    df_bulan = df_tahun[df_tahun['Bulan'] == bulan]
    
    if len(df_bulan) == 0 and bulan > 1:
        return current_max + 6
    
    return current_max + 1

def format_nomor(nomor):
    return str(int(nomor)).zfill(3)

def generate_single_pdf(nomor, perihal, tanggal, kepada, keterangan, jenis):
    pdf = FPDF()
    pdf.add_page()
    
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, "KLINIK UTAMA RAWAT INAP PARUNG", ln=True, align='C')
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 5, "Umum dan Kepegawaian", ln=True, align='C')
    pdf.cell(0, 5, "Dokumen ini digenerate secara otomatis", ln=True, align='C')
    pdf.line(10, 30, 200, 30)
    pdf.ln(20)

    if "Keputusan" in jenis or "Perjanjian" in jenis:
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 5, jenis.upper(), ln=True, align='C')
        pdf.cell(0, 5, f"NOMOR: {nomor}", ln=True, align='C')
        pdf.ln(10)
    
    pdf.set_font("Arial", size=12)
    
    tgl_str = tanggal.strftime('%d-%m-%Y')
    pdf.cell(0, 10, f"Tanggal: {tgl_str}", ln=True, align='R')

    if "Keputusan" not in jenis and "Perjanjian" not in jenis:
        pdf.cell(30, 8, "Nomor", 0, 0)
        pdf.cell(5, 8, ":", 0, 0)
        pdf.cell(0, 8, nomor, 0, 1)
        
        pdf.cell(30, 8, "Perihal", 0, 0)
        pdf.cell(5, 8, ":", 0, 0)
        pdf.cell(0, 8, perihal, 0, 1)
        pdf.ln(5)

    if kepada and kepada != "-":
        pdf.cell(0, 8, "Kepada Yth,", ln=True)
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, kepada, ln=True)
        pdf.set_font("Arial", size=12)
        pdf.ln(5)

    pdf.ln(5)
    pdf.multi_cell(0, 8, keterangan)
    
    pdf.ln(20)
    pdf.cell(120)
    pdf.cell(0, 8, "( __________________ )", ln=True)
    
    return pdf.output(dest='S').encode('latin-1', 'replace')

def generate_recap_pdf(df, start_date, end_date, jenis_surat):
    pdf = FPDF(orientation='L', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, f"LAPORAN REKAPITULASI {jenis_surat.upper()}", ln=True, align='C')
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, "KLINIK UTAMA RAWAT INAP PARUNG", ln=True, align='C')
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 6, "Umum dan Kepegawaian", ln=True, align='C')
    pdf.cell(0, 10, f"Periode: {start_date} s.d {end_date}", ln=True, align='C')
    pdf.ln(10)

    pdf.set_font("Arial", 'B', 10)
    pdf.set_fill_color(200, 220, 255)
    pdf.cell(15, 10, "No", 1, 0, 'C', 1)
    pdf.cell(30, 10, "Tanggal", 1, 0, 'C', 1)
    pdf.cell(60, 10, "Nomor Surat", 1, 0, 'C', 1)
    pdf.cell(40, 10, "Tujuan", 1, 0, 'C', 1)
    pdf.cell(130, 10, "Perihal", 1, 1, 'C', 1)

    pdf.set_font("Arial", size=9)
    for index, row in df.iterrows():
        tgl = row['Tanggal'].strftime('%d-%m-%Y') if isinstance(row['Tanggal'], date) else str(row['Tanggal'])
        
        perihal_txt = str(row['Perihal'])
        perihal_short = (perihal_txt[:75] + '...') if len(perihal_txt) > 75 else perihal_txt
        
        kepada_txt = str(row['Kepada'])
        kepada_short = (kepada_txt[:20] + '..') if len(kepada_txt) > 20 else kepada_txt

        pdf.cell(15, 8, str(int(row['No'])).zfill(3), 1, 0, 'C')
        pdf.cell(30, 8, tgl, 1, 0, 'C')
        pdf.cell(60, 8, str(row['Nomor_Surat']), 1, 0)
        pdf.cell(40, 8, kepada_short, 1, 0)
        pdf.cell(130, 8, perihal_short, 1, 1)

    return pdf.output(dest='S').encode('latin-1', 'replace')

def generate_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data Surat')
        worksheet = writer.sheets['Data Surat']
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)
    return output.getvalue()

def process_form(jenis_surat, kode_klasifikasi, tanggal, kepada, perihal, keterangan, df):
    if not kode_klasifikasi:
        return None, "Kode Klasifikasi wajib diisi!", None
    
    final_no_urut = get_next_number(df, tanggal, jenis_surat)
    nomor_formatted = format_nomor(final_no_urut)
    
    if jenis_surat == "Surat Masuk":
        final_nomor_surat = f"{kode_klasifikasi}/{nomor_formatted}-KURIP"
    elif jenis_surat == "Surat Keluar":
        final_nomor_surat = f"{kode_klasifikasi}/{nomor_formatted}-KURIP"
    elif jenis_surat == "Surat Keputusan (SK)":
        final_nomor_surat = f"{kode_klasifikasi}/SK-{nomor_formatted}/KURIP/{tanggal.year}"
    elif jenis_surat == "Perjanjian Kerjasama (MOU)":
        final_nomor_surat = f"{kode_klasifikasi}/{nomor_formatted}/KURIP/{tanggal.year}"
    else:
        return None, "Jenis surat tidak valid", None
        
    new_data = pd.DataFrame({
        "No": [final_no_urut],
        "Jenis": [jenis_surat],
        "Tanggal": [tanggal],
        "Bulan": [tanggal.month],
        "Tahun": [tanggal.year],
        "Kode_Klasifikasi": [kode_klasifikasi],
        "Kepada": [kepada],
        "Perihal": [perihal],
        "Keterangan": [keterangan],
        "Nomor_Surat": [final_nomor_surat]
    })
    
    new_data['Tanggal'] = pd.to_datetime(new_data['Tanggal']).dt.date
    
    updated_df = pd.concat([df, new_data], ignore_index=True)
    save_data(updated_df)
    
    pdf_bytes = generate_single_pdf(final_nomor_surat, perihal, tanggal, kepada, keterangan, jenis_surat)
    
    return final_nomor_surat, None, pdf_bytes


# --- UI LAYOUT ---
st.title("Sistem Penomoran Klinik Utama Rawat Inap Parung")
st.subheader("Umum dan Kepegawaian")

st.markdown("---")

df = load_data()

col1, col2 = st.columns(2)

# --- SURAT MASUK ---
with col1:
    with st.container(border=True):
        st.markdown("### Surat Masuk")
        st.caption("Format: Kode Klasifikasi/NomorSurat-KURIP")
        
        with st.form("form_surat_masuk", clear_on_submit=True):
            kode_sm = st.text_input("Kode Klasifikasi", placeholder="Cth: 005, ADM")
            tanggal_sm = st.date_input("Tanggal")
            kepada_sm = st.text_input("Kepada / Tujuan")
            perihal_sm = st.text_input("Perihal")
            keterangan_sm = st.text_area("Keterangan", height=80)
            
            calon_no = get_next_number(df, tanggal_sm, "Surat Masuk")
            st.info(f"Preview Next Number: **{format_nomor(calon_no)}**")
            
            submit_sm = st.form_submit_button("Simpan Surat Masuk")
        
        if submit_sm:
            df = load_data()
            nomor, error, pdf_bytes = process_form("Surat Masuk", kode_sm, tanggal_sm, kepada_sm, perihal_sm, keterangan_sm, df)
            if error:
                st.error(error)
            else:
                st.success(f"Tersimpan: {nomor}")
                st.session_state.last_saved['sm'] = {'nomor': nomor, 'pdf': pdf_bytes}
                st.rerun()
        
        if 'sm' in st.session_state.last_saved:
            data = st.session_state.last_saved['sm']
            st.download_button("Download Bukti PDF", data['pdf'], f"SM_{data['nomor'].replace('/', '_')}.pdf", "application/pdf", key="dl_sm")

# --- SURAT KELUAR ---
with col2:
    with st.container(border=True):
        st.markdown("### Surat Keluar")
        st.caption("Format: Kode Klasifikasi/NomorSurat-KURIP")
        
        with st.form("form_surat_keluar", clear_on_submit=True):
            kode_sk = st.text_input("Kode Klasifikasi", placeholder="Cth: 005, ADM")
            tanggal_sk = st.date_input("Tanggal")
            kepada_sk = st.text_input("Kepada / Tujuan")
            perihal_sk = st.text_input("Perihal")
            keterangan_sk = st.text_area("Keterangan", height=80)
            
            calon_no_sk = get_next_number(df, tanggal_sk, "Surat Keluar")
            st.info(f"Preview Next Number: **{format_nomor(calon_no_sk)}**")
            
            submit_sk = st.form_submit_button("Simpan Surat Keluar")
        
        if submit_sk:
            df = load_data()
            nomor, error, pdf_bytes = process_form("Surat Keluar", kode_sk, tanggal_sk, kepada_sk, perihal_sk, keterangan_sk, df)
            if error:
                st.error(error)
            else:
                st.success(f"Tersimpan: {nomor}")
                st.session_state.last_saved['sk'] = {'nomor': nomor, 'pdf': pdf_bytes}
                st.rerun()
        
        if 'sk' in st.session_state.last_saved:
            data = st.session_state.last_saved['sk']
            st.download_button("Download Bukti PDF", data['pdf'], f"SK_{data['nomor'].replace('/', '_')}.pdf", "application/pdf", key="dl_sk")

col3, col4 = st.columns(2)

# --- SURAT KEPUTUSAN (SK) ---
with col3:
    with st.container(border=True):
        st.markdown("### Surat Keputusan (SK)")
        st.caption("Format: Kode Klasifikasi/SK-NomorSurat/KURIP/Tahun")
        
        with st.form("form_sk", clear_on_submit=True):
            kode_skep = st.text_input("Kode Klasifikasi", placeholder="Cth: 005, ADM")
            tanggal_skep = st.date_input("Tanggal")
            kepada_skep = st.text_input("Kepada / Tujuan")
            perihal_skep = st.text_input("Perihal")
            keterangan_skep = st.text_area("Keterangan", height=80)
            
            calon_no_skep = get_next_number(df, tanggal_skep, "Surat Keputusan (SK)")
            st.info(f"Preview Next Number: **SK-{format_nomor(calon_no_skep)}**")
            
            submit_skep = st.form_submit_button("Simpan Surat Keputusan")
        
        if submit_skep:
            df = load_data()
            nomor, error, pdf_bytes = process_form("Surat Keputusan (SK)", kode_skep, tanggal_skep, kepada_skep, perihal_skep, keterangan_skep, df)
            if error:
                st.error(error)
            else:
                st.success(f"Tersimpan: {nomor}")
                st.session_state.last_saved['skep'] = {'nomor': nomor, 'pdf': pdf_bytes}
                st.rerun()
        
        if 'skep' in st.session_state.last_saved:
            data = st.session_state.last_saved['skep']
            st.download_button("Download Bukti PDF", data['pdf'], f"SKEP_{data['nomor'].replace('/', '_')}.pdf", "application/pdf", key="dl_skep")

# --- MOU ---
with col4:
    with st.container(border=True):
        st.markdown("### Perjanjian Kerjasama (MOU)")
        st.caption("Format: Kode Klasifikasi/NomorSurat/KURIP/Tahun")
        
        with st.form("form_mou", clear_on_submit=True):
            kode_mou = st.text_input("Kode Klasifikasi", placeholder="Cth: 005, ADM")
            tanggal_mou = st.date_input("Tanggal")
            kepada_mou = st.text_input("Kepada / Tujuan")
            perihal_mou = st.text_input("Perihal")
            keterangan_mou = st.text_area("Keterangan", height=80)
            
            calon_no_mou = get_next_number(df, tanggal_mou, "Perjanjian Kerjasama (MOU)")
            st.info(f"Preview Next Number: **{format_nomor(calon_no_mou)}**")
            
            submit_mou = st.form_submit_button("Simpan Perjanjian Kerjasama")
        
        if submit_mou:
            df = load_data()
            nomor, error, pdf_bytes = process_form("Perjanjian Kerjasama (MOU)", kode_mou, tanggal_mou, kepada_mou, perihal_mou, keterangan_mou, df)
            if error:
                st.error(error)
            else:
                st.success(f"Tersimpan: {nomor}")
                st.session_state.last_saved['mou'] = {'nomor': nomor, 'pdf': pdf_bytes}
                st.rerun()
        
        if 'mou' in st.session_state.last_saved:
            data = st.session_state.last_saved['mou']
            st.download_button("Download Bukti PDF", data['pdf'], f"MOU_{data['nomor'].replace('/', '_')}.pdf", "application/pdf", key="dl_mou")

st.markdown("---")
st.header("Laporan & Ekspor Data")

today = date.today()
df_report = load_data()

tab1, tab2, tab3, tab4 = st.tabs(["Surat Masuk", "Surat Keluar", "Surat Keputusan (SK)", "Perjanjian Kerjasama (MOU)"])

# FUNGSI UNTUK MENAMPILKAN DAN MENGHAPUS DATA
def render_report_tab(tab_name, jenis_filter, key_suffix):
    df_filtered = df_report[df_report['Jenis'] == jenis_filter]
    
    st.subheader(f"Laporan {tab_name}")
    c1, c2 = st.columns(2)
    with c1:
        start_d = st.date_input("Dari Tanggal", value=today.replace(day=1), key=f"start_{key_suffix}")
    with c2:
        end_d = st.date_input("Sampai Tanggal", value=today, key=f"end_{key_suffix}")
    
    if not df_filtered.empty:
        mask = (df_filtered['Tanggal'] >= start_d) & (df_filtered['Tanggal'] <= end_d)
        df_show = df_filtered.loc[mask]
    else:
        df_show = df_filtered
    
    st.info(f"Menampilkan **{len(df_show)}** dokumen")
    
    c_exp1, c_exp2 = st.columns(2)
    with c_exp1:
        if not df_show.empty:
            excel_data = generate_excel(df_show)
            st.download_button("Download Excel", excel_data, f'{key_suffix}_{start_d}_{end_d}.xlsx', 
                             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key=f"xl_{key_suffix}")
    with c_exp2:
        if not df_show.empty:
            pdf_data = generate_recap_pdf(df_show, start_d, end_d, jenis_filter)
            st.download_button("Download PDF Rekap", pdf_data, f'{key_suffix}_{start_d}_{end_d}.pdf', 'application/pdf', key=f"pdf_{key_suffix}")
    
    # --- TABEL DATA ---
    if not df_show.empty:
        display_df = df_show.drop(columns=['Bulan', 'Tahun', 'Keterangan'], errors='ignore').sort_values(by="No", ascending=False)
        st.dataframe(display_df, use_container_width=True, hide_index=True)
        
        # --- FITUR HAPUS DATA ---
        st.markdown("### 🗑️ Zona Hapus Data")
        with st.expander(f"Buka untuk menghapus data {tab_name}"):
            st.warning("⚠️ Perhatian: Data yang dihapus tidak dapat dikembalikan.")
            
            # Buat list opsi penghapusan yang informatif
            # Format: [Nomor Surat] - [Perihal]
            delete_options = df_show.apply(lambda x: f"{x['Nomor_Surat']} | {x['Perihal']}", axis=1).tolist()
            
            selected_option = st.selectbox("Pilih surat yang ingin dihapus:", ["-- Pilih Surat --"] + delete_options, key=f"del_sel_{key_suffix}")
            
            if selected_option != "-- Pilih Surat --":
                # Ambil Nomor Surat dari string yang dipilih (split berdasarkan " | ")
                nomor_to_delete = selected_option.split(" | ")[0]
                
                if st.button(f"Hapus Permanen {nomor_to_delete}", type="primary", key=f"btn_del_{key_suffix}"):
                    # Proses Hapus dari Database Utama
                    current_db = load_data()
                    new_db = current_db[current_db['Nomor_Surat'] != nomor_to_delete]
                    save_data(new_db)
                    st.success(f"Data {nomor_to_delete} berhasil dihapus!")
                    st.rerun()

with tab1:
    render_report_tab("Surat Masuk", "Surat Masuk", "sm")

with tab2:
    render_report_tab("Surat Keluar", "Surat Keluar", "sk")

with tab3:
    render_report_tab("Surat Keputusan", "Surat Keputusan (SK)", "skep")

with tab4:
    render_report_tab("MOU", "Perjanjian Kerjasama (MOU)", "mou")

st.caption("*Setiap jenis surat memiliki penomoran terpisah. Nomor reset ke 001 setiap awal tahun. 5 nomor dilewati setiap pergantian bulan.*")
