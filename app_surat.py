import streamlit as st
import pandas as pd
from datetime import datetime, date
import calendar
import os
from fpdf import FPDF
import io

st.set_page_config(page_title="Sistem Penomoran Klinik Utama Rawat Inap Parung", layout="wide")

DB_FILE = 'data_surat.csv'

def load_data():
    required_columns = ["No", "Jenis", "Tanggal", "Bulan", "Tahun", "Kode_Klasifikasi", "Kepada", "Perihal", "Keterangan", "Nomor_Surat"]
    
    if not os.path.exists(DB_FILE):
        return pd.DataFrame(columns=required_columns)
    
    df = pd.read_csv(DB_FILE)
    df['Tanggal'] = pd.to_datetime(df['Tanggal']).dt.date
    return df

def save_data(df):
    df.to_csv(DB_FILE, index=False)

def is_end_of_month(tanggal):
    last_day = calendar.monthrange(tanggal.year, tanggal.month)[1]
    return tanggal.day == last_day

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
    return str(nomor).zfill(3)

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
    
    return pdf.output(dest='S').encode('latin-1')

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
        
        perihal_short = (row['Perihal'][:75] + '...') if len(str(row['Perihal'])) > 75 else str(row['Perihal'])
        kepada_short = (row['Kepada'][:20] + '..') if len(str(row['Kepada'])) > 20 else str(row['Kepada'])

        pdf.cell(15, 8, str(row['No']).zfill(3), 1, 0, 'C')
        pdf.cell(30, 8, tgl, 1, 0, 'C')
        pdf.cell(60, 8, str(row['Nomor_Surat']), 1, 0)
        pdf.cell(40, 8, kepada_short, 1, 0)
        pdf.cell(130, 8, perihal_short, 1, 1)

    return pdf.output(dest='S').encode('latin-1')

def generate_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data Surat')
    return output.getvalue()

def process_form(jenis_surat, kode_klasifikasi, tanggal, kepada, perihal, keterangan, df):
    if not kode_klasifikasi:
        return None, "Kode Klasifikasi wajib diisi!"
    
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
        return None, "Jenis surat tidak valid"
        
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
    
    return final_nomor_surat, None


st.title("Sistem Penomoran Klinik Utama Rawat Inap Parung")
st.subheader("Umum dan Kepegawaian")

st.markdown("---")

df = load_data()

col1, col2 = st.columns(2)

with col1:
    with st.container(border=True):
        st.markdown("### Surat Masuk")
        st.caption("Format: Kode Klasifikasi/NomorSurat-KURIP")
        
        with st.form("form_surat_masuk"):
            kode_sm = st.text_input("Kode Klasifikasi", placeholder="Cth: 005, ADM", key="kode_sm")
            tanggal_sm = st.date_input("Tanggal", key="tgl_sm")
            kepada_sm = st.text_input("Kepada / Tujuan", key="kepada_sm")
            perihal_sm = st.text_input("Perihal", key="perihal_sm")
            keterangan_sm = st.text_area("Keterangan", height=80, key="ket_sm")
            
            calon_no = get_next_number(df, tanggal_sm, "Surat Masuk")
            st.info(f"Preview: **{kode_sm}/{format_nomor(calon_no)}-KURIP**")
            
            submit_sm = st.form_submit_button("Simpan Surat Masuk")
            
            if submit_sm:
                df = load_data()
                nomor, error = process_form("Surat Masuk", kode_sm, tanggal_sm, kepada_sm, perihal_sm, keterangan_sm, df)
                if error:
                    st.error(error)
                else:
                    st.success(f"Tersimpan: {nomor}")
                    pdf_bytes = generate_single_pdf(nomor, perihal_sm, tanggal_sm, kepada_sm, keterangan_sm, "Surat Masuk")
                    st.download_button("Download PDF", pdf_bytes, f"SM_{nomor.replace('/', '_')}.pdf", "application/pdf", key="dl_sm")

with col2:
    with st.container(border=True):
        st.markdown("### Surat Keluar")
        st.caption("Format: Kode Klasifikasi/NomorSurat-KURIP")
        
        with st.form("form_surat_keluar"):
            kode_sk = st.text_input("Kode Klasifikasi", placeholder="Cth: 005, ADM", key="kode_sk")
            tanggal_sk = st.date_input("Tanggal", key="tgl_sk")
            kepada_sk = st.text_input("Kepada / Tujuan", key="kepada_sk")
            perihal_sk = st.text_input("Perihal", key="perihal_sk")
            keterangan_sk = st.text_area("Keterangan", height=80, key="ket_sk")
            
            calon_no_sk = get_next_number(df, tanggal_sk, "Surat Keluar")
            st.info(f"Preview: **{kode_sk}/{format_nomor(calon_no_sk)}-KURIP**")
            
            submit_sk = st.form_submit_button("Simpan Surat Keluar")
            
            if submit_sk:
                df = load_data()
                nomor, error = process_form("Surat Keluar", kode_sk, tanggal_sk, kepada_sk, perihal_sk, keterangan_sk, df)
                if error:
                    st.error(error)
                else:
                    st.success(f"Tersimpan: {nomor}")
                    pdf_bytes = generate_single_pdf(nomor, perihal_sk, tanggal_sk, kepada_sk, keterangan_sk, "Surat Keluar")
                    st.download_button("Download PDF", pdf_bytes, f"SK_{nomor.replace('/', '_')}.pdf", "application/pdf", key="dl_sk")

col3, col4 = st.columns(2)

with col3:
    with st.container(border=True):
        st.markdown("### Surat Keputusan (SK)")
        st.caption("Format: Kode Klasifikasi/SK-NomorSurat/KURIP/Tahun")
        
        with st.form("form_sk"):
            kode_skep = st.text_input("Kode Klasifikasi", placeholder="Cth: 005, ADM", key="kode_skep")
            tanggal_skep = st.date_input("Tanggal", key="tgl_skep")
            kepada_skep = st.text_input("Kepada / Tujuan", key="kepada_skep")
            perihal_skep = st.text_input("Perihal", key="perihal_skep")
            keterangan_skep = st.text_area("Keterangan", height=80, key="ket_skep")
            
            calon_no_skep = get_next_number(df, tanggal_skep, "Surat Keputusan (SK)")
            st.info(f"Preview: **{kode_skep}/SK-{format_nomor(calon_no_skep)}/KURIP/{tanggal_skep.year}**")
            
            submit_skep = st.form_submit_button("Simpan Surat Keputusan")
            
            if submit_skep:
                df = load_data()
                nomor, error = process_form("Surat Keputusan (SK)", kode_skep, tanggal_skep, kepada_skep, perihal_skep, keterangan_skep, df)
                if error:
                    st.error(error)
                else:
                    st.success(f"Tersimpan: {nomor}")
                    pdf_bytes = generate_single_pdf(nomor, perihal_skep, tanggal_skep, kepada_skep, keterangan_skep, "Surat Keputusan (SK)")
                    st.download_button("Download PDF", pdf_bytes, f"SKEP_{nomor.replace('/', '_')}.pdf", "application/pdf", key="dl_skep")

with col4:
    with st.container(border=True):
        st.markdown("### Perjanjian Kerjasama (MOU)")
        st.caption("Format: Kode Klasifikasi/NomorSurat/KURIP/Tahun")
        
        with st.form("form_mou"):
            kode_mou = st.text_input("Kode Klasifikasi", placeholder="Cth: 005, ADM", key="kode_mou")
            tanggal_mou = st.date_input("Tanggal", key="tgl_mou")
            kepada_mou = st.text_input("Kepada / Tujuan", key="kepada_mou")
            perihal_mou = st.text_input("Perihal", key="perihal_mou")
            keterangan_mou = st.text_area("Keterangan", height=80, key="ket_mou")
            
            calon_no_mou = get_next_number(df, tanggal_mou, "Perjanjian Kerjasama (MOU)")
            st.info(f"Preview: **{kode_mou}/{format_nomor(calon_no_mou)}/KURIP/{tanggal_mou.year}**")
            
            submit_mou = st.form_submit_button("Simpan Perjanjian Kerjasama")
            
            if submit_mou:
                df = load_data()
                nomor, error = process_form("Perjanjian Kerjasama (MOU)", kode_mou, tanggal_mou, kepada_mou, perihal_mou, keterangan_mou, df)
                if error:
                    st.error(error)
                else:
                    st.success(f"Tersimpan: {nomor}")
                    pdf_bytes = generate_single_pdf(nomor, perihal_mou, tanggal_mou, kepada_mou, keterangan_mou, "Perjanjian Kerjasama (MOU)")
                    st.download_button("Download PDF", pdf_bytes, f"MOU_{nomor.replace('/', '_')}.pdf", "application/pdf", key="dl_mou")

st.markdown("---")
st.header("Laporan & Ekspor Data")

today = date.today()

tab1, tab2, tab3, tab4 = st.tabs(["Surat Masuk", "Surat Keluar", "Surat Keputusan (SK)", "Perjanjian Kerjasama (MOU)"])

with tab1:
    st.subheader("Laporan Surat Masuk")
    col_sm1, col_sm2 = st.columns(2)
    with col_sm1:
        start_sm = st.date_input("Dari Tanggal", value=today.replace(day=1), key="start_sm")
    with col_sm2:
        end_sm = st.date_input("Sampai Tanggal", value=today, key="end_sm")
    
    df_sm = load_data()
    df_sm = df_sm[df_sm['Jenis'] == "Surat Masuk"]
    if not df_sm.empty:
        mask = (df_sm['Tanggal'] >= start_sm) & (df_sm['Tanggal'] <= end_sm)
        df_sm_filtered = df_sm.loc[mask]
    else:
        df_sm_filtered = df_sm
    
    st.info(f"Menampilkan **{len(df_sm_filtered)}** Surat Masuk")
    
    col_exp_sm1, col_exp_sm2 = st.columns(2)
    with col_exp_sm1:
        if not df_sm_filtered.empty:
            excel_sm = generate_excel(df_sm_filtered)
            st.download_button("Download Excel", excel_sm, f'Surat_Masuk_{start_sm}_{end_sm}.xlsx', 
                             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key="exp_sm_xl")
    with col_exp_sm2:
        if not df_sm_filtered.empty:
            pdf_sm = generate_recap_pdf(df_sm_filtered, start_sm, end_sm, "Surat Masuk")
            st.download_button("Download PDF", pdf_sm, f'Surat_Masuk_{start_sm}_{end_sm}.pdf', 'application/pdf', key="exp_sm_pdf")
    
    if not df_sm_filtered.empty:
        st.dataframe(df_sm_filtered.sort_values(by="No", ascending=False), use_container_width=True, hide_index=True)

with tab2:
    st.subheader("Laporan Surat Keluar")
    col_sk1, col_sk2 = st.columns(2)
    with col_sk1:
        start_sk = st.date_input("Dari Tanggal", value=today.replace(day=1), key="start_sk")
    with col_sk2:
        end_sk = st.date_input("Sampai Tanggal", value=today, key="end_sk")
    
    df_sk = load_data()
    df_sk = df_sk[df_sk['Jenis'] == "Surat Keluar"]
    if not df_sk.empty:
        mask = (df_sk['Tanggal'] >= start_sk) & (df_sk['Tanggal'] <= end_sk)
        df_sk_filtered = df_sk.loc[mask]
    else:
        df_sk_filtered = df_sk
    
    st.info(f"Menampilkan **{len(df_sk_filtered)}** Surat Keluar")
    
    col_exp_sk1, col_exp_sk2 = st.columns(2)
    with col_exp_sk1:
        if not df_sk_filtered.empty:
            excel_sk = generate_excel(df_sk_filtered)
            st.download_button("Download Excel", excel_sk, f'Surat_Keluar_{start_sk}_{end_sk}.xlsx', 
                             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key="exp_sk_xl")
    with col_exp_sk2:
        if not df_sk_filtered.empty:
            pdf_sk = generate_recap_pdf(df_sk_filtered, start_sk, end_sk, "Surat Keluar")
            st.download_button("Download PDF", pdf_sk, f'Surat_Keluar_{start_sk}_{end_sk}.pdf', 'application/pdf', key="exp_sk_pdf")
    
    if not df_sk_filtered.empty:
        st.dataframe(df_sk_filtered.sort_values(by="No", ascending=False), use_container_width=True, hide_index=True)

with tab3:
    st.subheader("Laporan Surat Keputusan (SK)")
    col_skep1, col_skep2 = st.columns(2)
    with col_skep1:
        start_skep = st.date_input("Dari Tanggal", value=today.replace(day=1), key="start_skep")
    with col_skep2:
        end_skep = st.date_input("Sampai Tanggal", value=today, key="end_skep")
    
    df_skep = load_data()
    df_skep = df_skep[df_skep['Jenis'] == "Surat Keputusan (SK)"]
    if not df_skep.empty:
        mask = (df_skep['Tanggal'] >= start_skep) & (df_skep['Tanggal'] <= end_skep)
        df_skep_filtered = df_skep.loc[mask]
    else:
        df_skep_filtered = df_skep
    
    st.info(f"Menampilkan **{len(df_skep_filtered)}** Surat Keputusan")
    
    col_exp_skep1, col_exp_skep2 = st.columns(2)
    with col_exp_skep1:
        if not df_skep_filtered.empty:
            excel_skep = generate_excel(df_skep_filtered)
            st.download_button("Download Excel", excel_skep, f'SK_{start_skep}_{end_skep}.xlsx', 
                             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key="exp_skep_xl")
    with col_exp_skep2:
        if not df_skep_filtered.empty:
            pdf_skep = generate_recap_pdf(df_skep_filtered, start_skep, end_skep, "Surat Keputusan (SK)")
            st.download_button("Download PDF", pdf_skep, f'SK_{start_skep}_{end_skep}.pdf', 'application/pdf', key="exp_skep_pdf")
    
    if not df_skep_filtered.empty:
        st.dataframe(df_skep_filtered.sort_values(by="No", ascending=False), use_container_width=True, hide_index=True)

with tab4:
    st.subheader("Laporan Perjanjian Kerjasama (MOU)")
    col_mou1, col_mou2 = st.columns(2)
    with col_mou1:
        start_mou = st.date_input("Dari Tanggal", value=today.replace(day=1), key="start_mou")
    with col_mou2:
        end_mou = st.date_input("Sampai Tanggal", value=today, key="end_mou")
    
    df_mou = load_data()
    df_mou = df_mou[df_mou['Jenis'] == "Perjanjian Kerjasama (MOU)"]
    if not df_mou.empty:
        mask = (df_mou['Tanggal'] >= start_mou) & (df_mou['Tanggal'] <= end_mou)
        df_mou_filtered = df_mou.loc[mask]
    else:
        df_mou_filtered = df_mou
    
    st.info(f"Menampilkan **{len(df_mou_filtered)}** Perjanjian Kerjasama")
    
    col_exp_mou1, col_exp_mou2 = st.columns(2)
    with col_exp_mou1:
        if not df_mou_filtered.empty:
            excel_mou = generate_excel(df_mou_filtered)
            st.download_button("Download Excel", excel_mou, f'MOU_{start_mou}_{end_mou}.xlsx', 
                             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key="exp_mou_xl")
    with col_exp_mou2:
        if not df_mou_filtered.empty:
            pdf_mou = generate_recap_pdf(df_mou_filtered, start_mou, end_mou, "Perjanjian Kerjasama (MOU)")
            st.download_button("Download PDF", pdf_mou, f'MOU_{start_mou}_{end_mou}.pdf', 'application/pdf', key="exp_mou_pdf")
    
    if not df_mou_filtered.empty:
        st.dataframe(df_mou_filtered.sort_values(by="No", ascending=False), use_container_width=True, hide_index=True)

st.caption("*Setiap jenis surat memiliki penomoran terpisah. Nomor reset ke 001 setiap awal tahun. 5 nomor dilewati setiap pergantian bulan.*")