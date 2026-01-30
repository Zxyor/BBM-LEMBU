import streamlit as st
import mysql.connector
import pandas as pd
import io
import matplotlib
# PENTING: Set backend ke Agg agar tidak error thread
matplotlib.use('Agg') 
import matplotlib.pyplot as plt
from itertools import zip_longest 
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, portrait
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor 
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.drawing.image import Image as XLImage
import datetime
import re 

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Sistem BBM Proyek LEMBU", layout="wide")

# --- KONEKSI DATABASE ---
def init_connection():
    return mysql.connector.connect(
        host=st.secrets["db"]["host"],
        user=st.secrets["db"]["user"],
        password=st.secrets["db"]["password"],
        database=st.secrets["db"]["database"],
        port=st.secrets["db"]["port"]
    )

# --- HELPER FUNCTIONS ---
def get_hari_indonesia(tanggal):
    try:
        if pd.isnull(tanggal): return "-"
        kamus = {'Monday': 'Senin', 'Tuesday': 'Selasa', 'Wednesday': 'Rabu', 
                 'Thursday': 'Kamis', 'Friday': 'Jumat', 'Saturday': 'Sabtu', 'Sunday': 'Minggu'}
        return kamus[tanggal.strftime('%A')]
    except: return "-"

def cek_kategori(nama_alat):
    nama = str(nama_alat).upper()
    kata_kunci_mobil = ["TRUCK", "MOBIL", "TRITON", "DT", "FAW", "SANNY", "R6", "R10", "PICK UP", "HILUX", "STRADA", "GRAND MAX"]
    if any(k in nama for k in kata_kunci_mobil):
        return "MOBIL_TRUCK"
    return "ALAT_BERAT"

def set_cell_bg(cell, color_hex):
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color_hex))
    cell._tc.get_or_add_tcPr().append(shading_elm)

def process_rekap_data(df, excluded_list):
    if df.empty: return df
    df = df.copy()
    df['full_name'] = df['nama_alat'].astype(str) + " " + df['no_unit'].astype(str)
    mask = df['full_name'].isin(excluded_list)
    df.loc[mask, 'nama_alat'] = "Lainnya"
    df.loc[mask, 'no_unit'] = "-"
    return df

def filter_non_consumption(df):
    if df.empty: return df
    mask_donor = df['jumlah_liter'] < 0
    mask_recv = df['keterangan'].str.contains('Transfer|Pinjam', case=False, na=False)
    return df[~(mask_donor | mask_recv)]

# --- FUNGSI GRAFIK (DIPERBAIKI: Hapus List Comprehension) ---
def buat_grafik_final(df_alat, df_truck):
    try:
        buffer_img = io.BytesIO()
        if df_alat.empty and df_truck.empty: return None
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(7.5, 8)) 
        warna_alat, warna_truck = '#F4B084', '#9BC2E6' 
        
        if not df_alat.empty: df_alat = df_alat[df_alat['nama_alat'] != 'Lainnya']
        if not df_truck.empty: df_truck = df_truck[df_truck['nama_alat'] != 'Lainnya']
        
        # Grafik Alat Berat
        if not df_alat.empty:
            rekap = df_alat.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
            rekap['label'] = rekap.apply(lambda x: f"{x['nama_alat']} {x['no_unit']}", axis=1)
            data1 = rekap.groupby('label')['jumlah_liter'].sum().sort_values()
            data1.plot(kind='barh', ax=ax1, color=warna_alat, edgecolor='black', width=0.6)
            ax1.set_title(f"PEMAKAIAN ALAT BERAT", fontsize=10, fontweight='bold')
            ax1.set_ylabel("")
            # FIX: Ganti List Comprehension dengan For Loop biasa agar tidak nge-print text ke layar
            for i, v in enumerate(data1):
                ax1.text(v, i, f" {v:,.0f}", va='center', fontsize=8)
        else: 
            ax1.text(0.5, 0.5, "Data Kosong / Hidden", ha='center'); ax1.set_title("ALAT BERAT")
        
        # Grafik Mobil & Truck
        if not df_truck.empty:
            rekap = df_truck.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
            rekap['label'] = rekap.apply(lambda x: f"{x['nama_alat']} {x['no_unit']}", axis=1)
            data2 = rekap.groupby('label')['jumlah_liter'].sum().sort_values()
            data2.plot(kind='barh', ax=ax2, color=warna_truck, edgecolor='black', width=0.6)
            ax2.set_title(f"PEMAKAIAN MOBIL & TRUCK", fontsize=10, fontweight='bold')
            ax2.set_ylabel("")
            # FIX: Ganti List Comprehension dengan For Loop biasa
            for i, v in enumerate(data2):
                ax2.text(v, i, f" {v:,.0f}", va='center', fontsize=8)
        else: 
            ax2.text(0.5, 0.5, "Data Kosong / Hidden", ha='center'); ax2.set_title("MOBIL & TRUCK")
        
        plt.tight_layout(); plt.savefig(buffer_img, format='png', dpi=100); buffer_img.seek(0); plt.close(fig)
        return buffer_img
    except Exception as e: return None

# --- EXPORT FUNCTIONS (PDF, EXCEL, DOCX) ---

def generate_pdf_portrait(df_masuk, df_keluar, nama_lokasi, stok_awal, excluded_list):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=portrait(A4), rightMargin=15, leftMargin=15, topMargin=20, bottomMargin=20)
    elements = []
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='SmallHeader', parent=styles['Heading3'], fontSize=8, spaceAfter=2))
    styles.add(ParagraphStyle(name='TinyText', parent=styles['Normal'], fontSize=7))

    tm = df_masuk['jumlah_liter'].sum() if not df_masuk.empty else 0
    tk_real = df_keluar['jumlah_liter'].sum() if not df_keluar.empty else 0
    sisa = stok_awal + tm - tk_real
    df_keluar_report = filter_non_consumption(df_keluar)
    tk_report = df_keluar_report['jumlah_liter'].sum() if not df_keluar_report.empty else 0

    elements.append(Paragraph(f"LAPORAN BBM: {nama_lokasi}", styles['Title'])); elements.append(Spacer(1, 10))
    BATCH_SIZE = 30 
    
    col_w_A = [15, 25, 30, 65, 35, 30, 70]; left_chunks = []
    if not df_keluar_report.empty:
        data_rows = []; row_colors = []; last_date = None; is_grey = False 
        for i, row in df_keluar_report.iterrows():
            curr_date = row['tanggal']
            if last_date is not None and curr_date != last_date: is_grey = not is_grey
            last_date = curr_date; row_colors.append(is_grey)
            data_rows.append([i+1, row['HARI'][:3], row['tanggal'].strftime('%d/%m'), Paragraph(str(row['nama_alat']), styles['TinyText']), str(row['no_unit']), f"{row['jumlah_liter']:,.0f}", Paragraph(str(row['keterangan']), styles['TinyText'])])
        for i in range(0, len(data_rows), BATCH_SIZE):
            chunk_data = [['NO', 'HARI', 'TGL', 'ALAT', 'UNIT', 'LTR', 'KET']] + data_rows[i:i + BATCH_SIZE]
            chunk_colors = row_colors[i:i + BATCH_SIZE]
            if i + BATCH_SIZE >= len(data_rows): chunk_data.append(['', '', '', 'TOTAL', '', f"{tk_report:,.0f}", ''])
            t = Table(chunk_data, colWidths=col_w_A, repeatRows=1)
            t_style = [('FONTSIZE', (0,0), (-1,-1), 6), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'TOP')]
            for idx, use_grey in enumerate(chunk_colors):
                if use_grey: t_style.append(('BACKGROUND', (0, idx+1), (-1, idx+1), colors.whitesmoke))
            if i + BATCH_SIZE >= len(data_rows): t_style.append(('BACKGROUND', (0,-1), (-1,-1), colors.yellow)); t_style.append(('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'))
            t.setStyle(TableStyle(t_style))
            content = []; 
            if i == 0: content.append(Paragraph("A. PENGGUNAAN BBM (KELUAR)", styles['SmallHeader']))
            content.append(t); left_chunks.append(content)
    else:
        t = Table([['NO', '...'], ['-', 'DATA KOSONG']], colWidths=col_w_A)
        t.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('FONTSIZE', (0,0), (-1,-1), 6)]))
        left_chunks.append([Paragraph("A. PENGGUNAAN BBM (KELUAR)", styles['SmallHeader']), t])

    right_chunks = []; col_w_B = [15, 25, 30, 60, 40, 30, 70]; data_rows_B = []; row_colors_B = []
    if not df_masuk.empty:
        last_date = None; is_grey = False
        for i, row in df_masuk.iterrows():
            curr_date = row['tanggal']
            if last_date is not None and curr_date != last_date: is_grey = not is_grey
            last_date = curr_date; row_colors_B.append(is_grey)
            data_rows_B.append([i+1, row['HARI'][:3], row['tanggal'].strftime('%d/%m'), Paragraph(str(row['sumber']), styles['TinyText']), row['jenis_bbm'], f"{row['jumlah_liter']:,.0f}", Paragraph(str(row['keterangan']), styles['TinyText'])])
    chunked_B_tables = []
    if data_rows_B:
        for i in range(0, len(data_rows_B), BATCH_SIZE):
            chunk_B = [['NO', 'HARI', 'TGL', 'SUMBER', 'JENIS', 'LTR', 'KET']] + data_rows_B[i:i + BATCH_SIZE]; chunk_colors = row_colors_B[i:i + BATCH_SIZE]
            if i + BATCH_SIZE >= len(data_rows_B): chunk_B.append(['', '', '', 'TOTAL', '', f"{tm:,.0f}", ''])
            t = Table(chunk_B, colWidths=col_w_B, repeatRows=1)
            t_style = [('FONTSIZE', (0,0), (-1,-1), 6), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'TOP')]
            for idx, use_grey in enumerate(chunk_colors):
                if use_grey: t_style.append(('BACKGROUND', (0, idx+1), (-1, idx+1), colors.whitesmoke))
            if i + BATCH_SIZE >= len(data_rows_B): t_style.append(('BACKGROUND', (0,-1), (-1,-1), colors.yellow)); t_style.append(('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'))
            t.setStyle(TableStyle(t_style)); chunked_B_tables.append(t)
    else:
        t = Table([['NO', '...'], ['-', 'DATA KOSONG']], colWidths=col_w_B); t.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('FONTSIZE', (0,0), (-1,-1), 6)])); chunked_B_tables.append(t)

    b_start = [Paragraph("B. BBM MASUK", styles['SmallHeader']), chunked_B_tables[0]]; right_chunks.append(b_start)
    for tbl in chunked_B_tables[1:]: right_chunks.append([tbl])

    if not df_keluar.empty and 'kategori' not in df_keluar.columns: df_keluar['kategori'] = df_keluar['nama_alat'].apply(cek_kategori)
    df_rekap_src = process_rekap_data(df_keluar_report, excluded_list)
    df_alat = df_rekap_src[df_rekap_src['kategori'] == 'ALAT_BERAT'] if not df_rekap_src.empty else pd.DataFrame()
    df_truck = df_rekap_src[df_rekap_src['kategori'] == 'MOBIL_TRUCK'] if not df_rekap_src.empty else pd.DataFrame()

    def make_rekap_table_pdf(title, df_sub, color_fill):
        data = [[title, '', ''], ['NO', 'NAMA UNIT', 'LITER']]
        df_clean = df_sub[df_sub['nama_alat'] != 'Lainnya']; df_lainnya = df_sub[df_sub['nama_alat'] == 'Lainnya']; counter = 1
        if not df_clean.empty:
            rekap = df_clean.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
            for i, row in rekap.iterrows(): data.append([counter, Paragraph(f"{row['nama_alat']} {row['no_unit']}", styles['TinyText']), f"{row['jumlah_liter']:,.0f}"]); counter += 1
        grand_total = df_sub['jumlah_liter'].sum(); data.append(['', 'TOTAL', f"{grand_total:,.0f}"]); total_row_idx = len(data) - 1
        lainnya_row_idx = -1
        if not df_lainnya.empty:
            sum_lainnya = df_lainnya['jumlah_liter'].sum(); data.append(['', 'Lainnya', f"{sum_lainnya:,.0f}"]); lainnya_row_idx = len(data) - 1
        if df_sub.empty: data.append(['-', 'KOSONG', '-'])
        t = Table(data, colWidths=[25, 175, 70])
        s = [('SPAN', (0,0), (-1,0)), ('ALIGN', (0,0), (-1,0), 'CENTER'), ('BACKGROUND', (0,0), (-1,0), color_fill), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('FONTSIZE', (0,0), (-1,-1), 6), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('ALIGN', (0,1), (-1,1), 'CENTER'), ('ALIGN', (-1,2), (-1,-1), 'RIGHT'), ('BACKGROUND', (1, total_row_idx), (-1, total_row_idx), colors.yellow), ('FONTNAME', (1, total_row_idx), (-1, total_row_idx), 'Helvetica-Bold')]
        if lainnya_row_idx != -1: s.append(('BACKGROUND', (1, lainnya_row_idx), (-1, lainnya_row_idx), colors.mistyrose)); s.append(('FONTNAME', (1, lainnya_row_idx), (-1, lainnya_row_idx), 'Helvetica-Oblique'))
        t.setStyle(TableStyle(s)); return t

    rekap_chunk = [Spacer(1,5), Paragraph("C. REKAPITULASI UNIT", styles['SmallHeader'])]
    if not df_alat.empty: rekap_chunk.append(make_rekap_table_pdf("TOTAL ALAT BERAT", df_alat, colors.orange)); rekap_chunk.append(Spacer(1,5))
    if not df_truck.empty: rekap_chunk.append(make_rekap_table_pdf("TOTAL MOBIL & TRUCK", df_truck, colors.skyblue)); rekap_chunk.append(Spacer(1,5))
    right_chunks.append(rekap_chunk)
    
    t_stok = Table([[f"SISA STOK: {sisa:,.0f} L (Awal: {stok_awal:,.0f} + M: {tm:,.0f} - K: {tk_real:,.0f})"]], colWidths=[270])
    t_stok.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.lightyellow), ('GRID', (0,0), (-1,-1), 1, colors.black), ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'), ('TEXTCOLOR', (0,0), (-1,-1), colors.red), ('FONTSIZE', (0,0), (-1,-1), 8)]))
    last_chunk = [Paragraph("D. INFO STOK", styles['SmallHeader']), t_stok, Spacer(1,5)]
    img_buffer = buat_grafik_final(df_alat, df_truck)
    if img_buffer: last_chunk.append(Paragraph("E. DIAGRAM / GRAFIK", styles['SmallHeader'])); last_chunk.append(RLImage(img_buffer, width=270, height=300))
    right_chunks.append(last_chunk)

    rows_data = []
    for l_content, r_content in zip_longest(left_chunks, right_chunks, fillvalue=[]): rows_data.append([l_content, r_content])
    main_table = Table(rows_data, colWidths=[280, 280], splitByRow=1)
    main_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('LEFTPADDING', (0,0), (-1,-1), 2), ('RIGHTPADDING', (0,0), (-1,-1), 2), ('BOTTOMPADDING', (0,0), (-1,-1), 10)]))
    elements.append(main_table); doc.build(elements); buffer.seek(0)
    return buffer

def generate_excel_styled(df_masuk, df_keluar, nama_lokasi, stok_awal, excluded_list):
    output = io.BytesIO(); wb = Workbook()
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid"); pink_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid") 
    grey_fill = PatternFill(start_color="D3D3D3", fill_type="solid"); row_grey_fill = PatternFill(start_color="F2F2F2", fill_type="solid")
    header_font = Font(bold=True, size=11); title_font = Font(bold=True, size=14)
    tm = df_masuk['jumlah_liter'].sum() if not df_masuk.empty else 0; tk_real = df_keluar['jumlah_liter'].sum() if not df_keluar.empty else 0; sisa = stok_awal + tm - tk_real
    df_keluar_report = filter_non_consumption(df_keluar); tk_report = df_keluar_report['jumlah_liter'].sum() if not df_keluar_report.empty else 0
    ws = wb.active; ws.title = "Laporan BBM"; ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.fitToWidth = 1
    ws.merge_cells('A1:O1'); ws['A1'] = f"LAPORAN BBM: {nama_lokasi}"; ws['A1'].font = title_font; ws['A1'].alignment = Alignment(horizontal='center')
    if not df_keluar.empty and 'kategori' not in df_keluar.columns: df_keluar['kategori'] = df_keluar['nama_alat'].apply(cek_kategori)
    df_rekap_src = process_rekap_data(df_keluar_report, excluded_list)
    df_alat = df_rekap_src[df_rekap_src['kategori'] == 'ALAT_BERAT'] if not df_rekap_src.empty else pd.DataFrame()
    df_truck = df_rekap_src[df_rekap_src['kategori'] == 'MOBIL_TRUCK'] if not df_rekap_src.empty else pd.DataFrame()

    def buat_tabel_excel_colored(ws, start_row, start_col, title, headers, dataframe, data_cols, total_label, total_val):
        c = ws.cell(row=start_row, column=start_col, value=title); c.font = header_font; start_row += 1
        for col_i, h in enumerate(headers): c = ws.cell(row=start_row, column=start_col + col_i, value=h); c.fill = grey_fill; c.border = thin_border; c.font = header_font; c.alignment = Alignment(horizontal='center')
        start_row += 1
        if not dataframe.empty:
            last_date = None; is_grey = False 
            for i, row in dataframe.iterrows():
                curr_date = row['tanggal']; 
                if last_date is not None and curr_date != last_date: is_grey = not is_grey
                last_date = curr_date; current_fill = row_grey_fill if is_grey else None
                c_no = ws.cell(row=start_row, column=start_col, value=i+1); c_no.border = thin_border
                if current_fill: c_no.fill = current_fill
                for col_i, col_name in enumerate(data_cols):
                    val = row[col_name]; 
                    if col_name == 'tanggal': val = val.strftime('%d/%m/%Y')
                    c = ws.cell(row=start_row, column=start_col + 1 + col_i, value=val); c.border = thin_border
                    if current_fill: c.fill = current_fill
                    if col_name == 'jumlah_liter': c.number_format = '#,##0'
                start_row += 1
            ws.cell(row=start_row, column=start_col + 3, value=total_label).font = header_font 
            c_tot = ws.cell(row=start_row, column=start_col + 5, value=total_val); c_tot.font = header_font; c_tot.number_format = '#,##0'
            for i in range(len(headers)): c = ws.cell(row=start_row, column=start_col + i); c.fill = yellow_fill; c.border = thin_border
            start_row += 2
        else: ws.cell(row=start_row, column=start_col, value="Data Kosong"); start_row += 2
        return start_row 
    row_cursor_A = 3; row_cursor_A = buat_tabel_excel_colored(ws, row_cursor_A, 1, "A. PENGGUNAAN BBM (KELUAR)", ['NO', 'HARI', 'TANGGAL', 'NAMA ALAT', 'NO UNIT', 'LITER', 'KETERANGAN'], df_keluar_report, ['HARI', 'tanggal', 'nama_alat', 'no_unit', 'jumlah_liter', 'keterangan'], "TOTAL KELUAR", tk_report)
    row_cursor_B = 3; row_cursor_B = buat_tabel_excel_colored(ws, row_cursor_B, 9, "B. BBM MASUK", ['NO', 'HARI', 'TANGGAL', 'SUMBER', 'JENIS', 'LITER', 'KETERANGAN'], df_masuk, ['HARI', 'tanggal', 'sumber', 'jenis_bbm', 'jumlah_liter', 'keterangan'], "TOTAL MASUK", tm)

    ws.cell(row=row_cursor_B, column=9, value="C. REKAPITULASI UNIT").font = title_font; row_cursor_B += 1
    def buat_rekap_excel(ws, r_row, c_col, title, df_sub, color_fill):
        ws.merge_cells(start_row=r_row, start_column=c_col, end_row=r_row, end_column=c_col+2)
        c = ws.cell(row=r_row, column=c_col, value=title); c.font = header_font; c.fill = color_fill; c.border = thin_border; c.alignment = Alignment(horizontal='center'); r_row += 1
        ws.cell(row=r_row, column=c_col, value="NO").border = thin_border; ws.cell(row=r_row, column=c_col+1, value="NAMA UNIT").border = thin_border; ws.cell(row=r_row, column=c_col+2, value="LITER").border = thin_border; 
        for i in range(3): ws.cell(row=r_row, column=c_col+i).alignment = Alignment(horizontal='center')
        r_row += 1
        if not df_sub.empty:
            df_clean = df_sub[df_sub['nama_alat'] != 'Lainnya']; df_lainnya = df_sub[df_sub['nama_alat'] == 'Lainnya']; counter = 1
            if not df_clean.empty:
                rekap = df_clean.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
                for idx, row in rekap.iterrows():
                    nm = f"{row['nama_alat']} {row['no_unit']}"; c_no = ws.cell(row=r_row, column=c_col, value=counter); c_no.border = thin_border
                    c_nm = ws.cell(row=r_row, column=c_col+1, value=nm); c_nm.border = thin_border; c_val = ws.cell(row=r_row, column=c_col+2, value=row['jumlah_liter']); c_val.border = thin_border; c_val.number_format = '#,##0'
                    r_row += 1; counter += 1
            ws.merge_cells(start_row=r_row, start_column=c_col, end_row=r_row, end_column=c_col+1)
            c_tot_lbl = ws.cell(row=r_row, column=c_col, value="TOTAL"); c_tot_lbl.fill = yellow_fill; c_tot_lbl.border = thin_border; c_tot_lbl.alignment = Alignment(horizontal='center')
            c_tot_val = ws.cell(row=r_row, column=c_col+2, value=df_sub['jumlah_liter'].sum()); c_tot_val.fill = yellow_fill; c_tot_val.border = thin_border; c_tot_val.number_format = '#,##0'; r_row += 1
            if not df_lainnya.empty:
                sum_lainnya = df_lainnya['jumlah_liter'].sum(); ws.merge_cells(start_row=r_row, start_column=c_col, end_row=r_row, end_column=c_col+1)
                c_lain = ws.cell(row=r_row, column=c_col, value="Lainnya"); c_lain.fill = pink_fill; c_lain.border = thin_border; c_lain.alignment = Alignment(horizontal='center')
                c_lain_val = ws.cell(row=r_row, column=c_col+2, value=sum_lainnya); c_lain_val.fill = pink_fill; c_lain_val.border = thin_border; c_lain_val.number_format = '#,##0'; r_row += 1
            r_row += 1 
        else: ws.cell(row=r_row, column=c_col, value="- Kosong -"); r_row += 2
        return r_row

    row_cursor_B = buat_rekap_excel(ws, row_cursor_B, 9, "TOTAL ALAT BERAT", df_alat, PatternFill(start_color="F4B084", fill_type="solid"))
    row_cursor_B = buat_rekap_excel(ws, row_cursor_B, 9, "TOTAL MOBIL & TRUCK", df_truck, PatternFill(start_color="9BC2E6", fill_type="solid"))

    ws.cell(row=row_cursor_B, column=9, value="D. INFO STOK").font = title_font; row_cursor_B += 1
    stok_text = f"SISA STOK: {sisa:,.2f} L (Awal: {stok_awal:,.0f} + M: {tm:,.0f} - K: {tk_real:,.0f})"
    ws.merge_cells(start_row=row_cursor_B, start_column=9, end_row=row_cursor_B, end_column=14)
    c_stok = ws.cell(row=row_cursor_B, column=9, value=stok_text)
    c_stok.font = Font(bold=True, size=12, color="FF0000"); c_stok.alignment = Alignment(horizontal='left'); c_stok.fill = PatternFill(start_color="FFFFCC", fill_type="solid"); c_stok.border = thin_border; row_cursor_B += 2
    ws.cell(row=row_cursor_B, column=9, value="E. DIAGRAM / GRAFIK").font = title_font; row_cursor_B += 1
    img_buffer = buat_grafik_final(df_alat, df_truck)
    if img_buffer:
        try: img = XLImage(img_buffer); ws.add_image(img, f'I{row_cursor_B}') 
        except: pass
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']: ws.column_dimensions[col].width = 15
    ws.column_dimensions['A'].width = 5; ws.column_dimensions['D'].width = 30; ws.column_dimensions['G'].width = 25
    for col in ['I', 'J', 'K', 'L', 'M', 'N', 'O']: ws.column_dimensions[col].width = 15
    ws.column_dimensions['I'].width = 5; ws.column_dimensions['J'].width = 35; ws.column_dimensions['L'].width = 25; ws.column_dimensions['O'].width = 25
    wb.save(output); output.seek(0)
    return output

def generate_docx_fixed(df_masuk, df_keluar, nama_lokasi, stok_awal, excluded_list):
    doc = Document(); 
    for section in doc.sections: section.top_margin = Cm(1.27); section.bottom_margin = Cm(1.27); section.left_margin = Cm(1.0); section.right_margin = Cm(1.0)
    tm = df_masuk['jumlah_liter'].sum() if not df_masuk.empty else 0; tk_real = df_keluar['jumlah_liter'].sum() if not df_keluar.empty else 0; sisa = stok_awal + tm - tk_real
    df_keluar_report = filter_non_consumption(df_keluar); tk_report = df_keluar_report['jumlah_liter'].sum() if not df_keluar_report.empty else 0
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; r = p.add_run(f"LAPORAN BBM: {nama_lokasi}"); r.bold = True; r.font.size = Pt(14); doc.add_paragraph()
    master_table = doc.add_table(rows=1, cols=2); master_table.autofit = False; master_table.allow_autofit = False
    for cell in master_table.columns[0].cells: cell.width = Cm(9.5)
    for cell in master_table.columns[1].cells: cell.width = Cm(9.5)

    left_cell = master_table.cell(0, 0); left_cell.add_paragraph("A. PENGGUNAAN BBM (KELUAR)", style='Heading 3')
    headers_A = ['NO', 'TGL', 'ALAT', 'UNIT', 'LTR', 'KET']; table_A = left_cell.add_table(rows=1, cols=len(headers_A)); table_A.style = 'Table Grid'; hdr_cells = table_A.rows[0].cells
    for i, h in enumerate(headers_A): hdr_cells[i].text = h; hdr_cells[i].paragraphs[0].runs[0].bold = True; hdr_cells[i].paragraphs[0].runs[0].font.size = Pt(7)

    if not df_keluar_report.empty:
        last_date = None; is_grey = False
        for i, row in df_keluar_report.iterrows():
            curr_date = row['tanggal']; 
            if last_date is not None and curr_date != last_date: is_grey = not is_grey
            last_date = curr_date
            row_cells = table_A.add_row().cells
            if is_grey: 
                for c in row_cells: set_cell_bg(c, "F2F2F2")
            row_cells[0].text = str(i+1); row_cells[1].text = row['tanggal'].strftime('%d/%m'); row_cells[2].text = str(row['nama_alat']); row_cells[3].text = str(row['no_unit'])
            row_cells[4].text = f"{row['jumlah_liter']:.0f}"; row_cells[5].text = str(row['keterangan'])
            for c in row_cells: c.paragraphs[0].runs[0].font.size = Pt(7)
        row_tot = table_A.add_row().cells; row_tot[2].text = "TOTAL"; row_tot[4].text = f"{tk_report:.0f}"
        for c in row_tot: set_cell_bg(c, "FFFF00"); p = c.paragraphs[0]; 
        if not p.runs: p.add_run() 
        p.runs[0].font.size = Pt(7); p.runs[0].bold = True
    else: table_A.add_row().cells[0].text = "DATA KOSONG"

    right_cell = master_table.cell(0, 1); right_cell.add_paragraph("B. BBM MASUK", style='Heading 3')
    headers_B = ['NO', 'TGL', 'SUMBER', 'JNS', 'LTR', 'KET']; table_B = right_cell.add_table(rows=1, cols=len(headers_B)); table_B.style = 'Table Grid'; hdr_B = table_B.rows[0].cells
    for i, h in enumerate(headers_B): hdr_B[i].text = h; hdr_B[i].paragraphs[0].runs[0].bold = True; hdr_B[i].paragraphs[0].runs[0].font.size = Pt(7)

    if not df_masuk.empty:
        last_date = None; is_grey = False
        for i, row in df_masuk.iterrows():
            curr_date = row['tanggal']; 
            if last_date is not None and curr_date != last_date: is_grey = not is_grey
            last_date = curr_date
            row_cells = table_B.add_row().cells
            if is_grey: 
                for c in row_cells: set_cell_bg(c, "F2F2F2")
            row_cells[0].text = str(i+1); row_cells[1].text = row['tanggal'].strftime('%d/%m'); row_cells[2].text = str(row['sumber']); row_cells[3].text = str(row['jenis_bbm'])
            row_cells[4].text = f"{row['jumlah_liter']:.0f}"; row_cells[5].text = str(row['keterangan']); 
            for c in row_cells: c.paragraphs[0].runs[0].font.size = Pt(7)
        row_tot = table_B.add_row().cells; row_tot[2].text = "TOTAL"; row_tot[4].text = f"{tm:.0f}"
        for c in row_tot: set_cell_bg(c, "FFFF00"); p = c.paragraphs[0]; 
        if not p.runs: p.add_run()
        p.runs[0].font.size = Pt(7); p.runs[0].bold = True
    else: table_B.add_row().cells[0].text = "DATA KOSONG"

    right_cell.add_paragraph().add_run(""); right_cell.add_paragraph("C. REKAPITULASI UNIT", style='Heading 3')
    if not df_keluar.empty and 'kategori' not in df_keluar.columns: df_keluar['kategori'] = df_keluar['nama_alat'].apply(cek_kategori)
    df_rekap_src = process_rekap_data(df_keluar_report, excluded_list)
    df_alat = df_rekap_src[df_rekap_src['kategori'] == 'ALAT_BERAT'] if not df_rekap_src.empty else pd.DataFrame()
    df_truck = df_rekap_src[df_rekap_src['kategori'] == 'MOBIL_TRUCK'] if not df_rekap_src.empty else pd.DataFrame()

    def add_rekap_docx(container, title, df_sub, bg_color):
        container.add_paragraph(title, style='Heading 4'); t = container.add_table(rows=1, cols=3); t.style = 'Table Grid'
        h = t.rows[0].cells; h[0].text="NO"; h[1].text="NAMA UNIT"; h[2].text="LITER"
        if not df_sub.empty:
            df_clean = df_sub[df_sub['nama_alat'] != 'Lainnya']; df_lainnya = df_sub[df_sub['nama_alat'] == 'Lainnya']; counter = 1
            if not df_clean.empty:
                rekap = df_clean.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
                for idx, row in rekap.iterrows():
                    c = t.add_row().cells; nm = f"{row['nama_alat']} {row['no_unit']}"; c[0].text = str(counter); c[1].text = nm; c[2].text = f"{row['jumlah_liter']:.0f}"
                    for cell in c: cell.paragraphs[0].runs[0].font.size = Pt(8)
                    counter += 1
            rt = t.add_row().cells; rt[1].text = "TOTAL"; rt[2].text = f"{df_sub['jumlah_liter'].sum():.0f}"
            for c in rt: set_cell_bg(c, "FFFF00"); p = c.paragraphs[0]; 
            if not p.runs: p.add_run()
            p.runs[0].bold = True; p.runs[0].font.size = Pt(8)
            if not df_lainnya.empty:
                rl = t.add_row().cells; rl[1].text = "Lainnya"; rl[2].text = f"{df_lainnya['jumlah_liter'].sum():.0f}"
                for c in rl: set_cell_bg(c, "FFCCCC"); p = c.paragraphs[0]; 
                if not p.runs: p.add_run()
                p.runs[0].italic = True; p.runs[0].font.size = Pt(8)
        else: t.add_row().cells[0].text = "KOSONG"

    if not df_alat.empty: add_rekap_docx(right_cell, "TOTAL ALAT BERAT", df_alat, "F4B084"); right_cell.add_paragraph().add_run("")
    if not df_truck.empty: add_rekap_docx(right_cell, "TOTAL MOBIL & TRUCK", df_truck, "9BC2E6"); right_cell.add_paragraph().add_run("")
    right_cell.add_paragraph("D. INFO STOK", style='Heading 3')
    p_stok = right_cell.add_paragraph(f"SISA STOK: {sisa:,.2f} Liter\n(Awal: {stok_awal:,.0f} + Masuk: {tm:,.0f} - Keluar: {tk_real:,.0f})")
    p_stok.runs[0].bold = True; p_stok.runs[0].font.color.rgb = RGBColor(255, 0, 0)
    right_cell.add_paragraph().add_run(""); right_cell.add_paragraph("E. DIAGRAM", style='Heading 3')
    img_buffer = buat_grafik_final(df_alat, df_truck)
    if img_buffer: right_cell.add_paragraph().add_run().add_picture(img_buffer, width=Cm(8))
    buffer = io.BytesIO(); doc.save(buffer); buffer.seek(0)
    return buffer

# --- MAIN APP ---
def main():
    try: 
        conn = init_connection(); cursor = conn.cursor(buffered=True) 
        
        # --- DATABASE MIGRATION ---
        try: cursor.execute("SELECT stok_awal FROM lokasi_proyek LIMIT 1"); cursor.fetchall()
        except: cursor.execute("ALTER TABLE lokasi_proyek ADD COLUMN stok_awal FLOAT DEFAULT 0"); conn.commit()
        try: cursor.execute("SELECT kunci_lokasi FROM lokasi_proyek LIMIT 1"); cursor.fetchall()
        except: cursor.execute("ALTER TABLE lokasi_proyek ADD COLUMN kunci_lokasi VARCHAR(255) DEFAULT '123'"); conn.commit()
        cursor.execute("""CREATE TABLE IF NOT EXISTS log_aktivitas (id INT AUTO_INCREMENT PRIMARY KEY, lokasi_id INT, tanggal DATETIME DEFAULT CURRENT_TIMESTAMP, kategori VARCHAR(50), deskripsi TEXT, affected_ids TEXT)""")
        try: cursor.execute("SELECT affected_ids FROM log_aktivitas LIMIT 1"); cursor.fetchall()
        except: 
             try: cursor.execute("ALTER TABLE log_aktivitas ADD COLUMN affected_ids TEXT"); conn.commit()
             except: pass
        cursor.execute("""CREATE TABLE IF NOT EXISTS rekap_exclude (id INT AUTO_INCREMENT PRIMARY KEY, lokasi_id INT, nama_unit_full VARCHAR(255))"""); conn.commit()
    except Exception as e: st.error(f"Database Error: {e}"); st.stop()

    # --- 1. ADMIN LOGIN ---
    password_rahasia = "123" 
    if "logged_in" not in st.session_state: st.session_state.logged_in = False
    
    # State untuk menyimpan lokasi yang sedang aktif (masuk)
    if "active_project_id" not in st.session_state: st.session_state.active_project_id = None
    if "active_project_name" not in st.session_state: st.session_state.active_project_name = None

    if not st.session_state.logged_in:
        st.markdown("<h1 style='text-align: center;'>🔐 Login Admin</h1>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns([1,2,1])
        with c2:
            pwd = st.text_input("Masukkan Password Admin:", type="password")
            if st.button("Login", use_container_width=True):
                if pwd == password_rahasia: 
                    st.session_state.logged_in = True
                    st.rerun()
                else: 
                    st.error("Password Salah!")
        st.stop() 

    # --- 2. MENU UTAMA (LANDING PAGE) ---
    # Jika belum ada proyek aktif, tampilkan pilihan Masuk/Buat
    if st.session_state.active_project_id is None:
        st.title("🗂️ Menu Utama")
        st.write("Selamat Datang, Admin. Silakan pilih lokasi proyek atau buat baru.")
        st.divider()

        col_left, col_right = st.columns(2, gap="large")

        # FITUR MASUK LOKASI
        with col_left:
            with st.container(border=True):
                st.subheader("📂 Masuk ke Lokasi Proyek")
                try: df_lokasi = pd.read_sql("SELECT * FROM lokasi_proyek", conn)
                except: df_lokasi = pd.DataFrame()

                if not df_lokasi.empty:
                    pilih_nama = st.selectbox("Pilih Lokasi:", df_lokasi['nama_tempat'])
                    input_pass_lokasi = st.text_input("Password Lokasi:", type="password", key="pass_enter")
                    if st.button("Masuk Lokasi", type="primary", use_container_width=True):
                        # Cek Password
                        data_lok = df_lokasi[df_lokasi['nama_tempat'] == pilih_nama].iloc[0]
                        if input_pass_lokasi == data_lok['kunci_lokasi']:
                            st.session_state.active_project_id = int(data_lok['id'])
                            st.session_state.active_project_name = data_lok['nama_tempat']
                            st.success(f"Berhasil masuk ke {pilih_nama}")
                            st.rerun()
                        else:
                            st.error("Password Lokasi Salah!")
                else:
                    st.info("Belum ada data lokasi. Silakan buat baru di sebelah kanan.")

        # FITUR BUAT LOKASI BARU
        with col_right:
            with st.container(border=True):
                st.subheader("➕ Buat Lokasi Baru")
                new_name = st.text_input("Nama Lokasi Baru")
                new_pass = st.text_input("Buat Password Lokasi", type="password", key="pass_create")
                st.warning("⚠️ **Password nya diingat baik-baik !**")
                
                if st.button("Simpan Lokasi Baru", use_container_width=True):
                    if new_name and new_pass:
                        try:
                            cursor.execute("INSERT INTO lokasi_proyek (nama_tempat, kunci_lokasi) VALUES (%s, %s)", (new_name, new_pass))
                            conn.commit()
                            st.success("Lokasi berhasil dibuat! Silakan masuk melalui menu di sebelah kiri.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Gagal membuat lokasi: {e}")
                    else:
                        st.error("Nama dan Password wajib diisi!")
        
        st.stop() # Berhenti disini agar tidak load dashboard di bawah

    # --- 3. PROJECT DASHBOARD (Hanya muncul jika active_project_id terisi) ---
    
    # Set variable lokal dari session
    lokasi_id = st.session_state.active_project_id
    nama_proyek = st.session_state.active_project_name
    
    # Ambil stok awal terbaru dari DB (agar sinkron)
    cursor.execute("SELECT stok_awal FROM lokasi_proyek WHERE id=%s", (lokasi_id,))
    stok_awal_val = cursor.fetchone()[0]
    
    # --- SIDEBAR NAVIGASI ---
    with st.sidebar:
        st.header(f"📍 {nama_proyek}")
        
        # Tombol Kembali ke Menu Utama
        if st.button("⬅️ Kembali ke Menu Utama", use_container_width=True):
            st.session_state.active_project_id = None
            st.session_state.active_project_name = None
            st.rerun()
        
        st.divider()
        st.write(f"**Stok Awal:** {stok_awal_val:,.0f} Liter")
        new_sa = st.number_input("Atur Stok Awal (Saldo)", value=float(stok_awal_val))
        if st.button("Update Stok Awal"):
            cursor.execute("UPDATE lokasi_proyek SET stok_awal = %s WHERE id = %s", (new_sa, lokasi_id))
            conn.commit(); st.success("Tersimpan!"); st.rerun()

    # --- LOAD DATA PROYEK ---
    df_masuk = pd.read_sql(f"SELECT * FROM bbm_masuk WHERE lokasi_id={lokasi_id}", conn)
    df_keluar = pd.read_sql(f"SELECT * FROM bbm_keluar WHERE lokasi_id={lokasi_id}", conn)
    df_log = pd.read_sql(f"SELECT * FROM log_aktivitas WHERE lokasi_id={lokasi_id}", conn)
    df_ex = pd.read_sql(f"SELECT nama_unit_full FROM rekap_exclude WHERE lokasi_id={lokasi_id}", conn)
    excluded_list = df_ex['nama_unit_full'].tolist() if not df_ex.empty else []

    if not df_masuk.empty: df_masuk['HARI'] = pd.to_datetime(df_masuk['tanggal']).apply(get_hari_indonesia)
    if not df_keluar.empty: df_keluar['HARI'] = pd.to_datetime(df_keluar['tanggal']).apply(get_hari_indonesia)

    # --- TABS DASHBOARD ---
    st.title(f"🚜 Dashboard: {nama_proyek}")
    t1, t2, t3 = st.tabs(["📝 Input & History", "📊 Laporan & Grafik", "🖨️ Export Dokumen"])
    
    with t1:
        # --- FORM INPUT ---
        st.subheader("Input Transaksi BBM")
        mode_transaksi = st.radio("Pilih Jenis Transaksi:", ["📥 BBM MASUK", "📤 PENGGUNAAN BBM", "🔄 PINJAM / TRANSFER ANTAR UNIT"], horizontal=True)
        
        with st.container(border=True):
            if mode_transaksi == "📥 BBM MASUK":
                with st.form("form_masuk"):
                    c1, c2 = st.columns(2)
                    with c1: tg = st.date_input("Tanggal Masuk"); sm = st.text_input("Sumber / Supplier")
                    with c2: jn = st.selectbox("Jenis BBM", ["Dexlite","Solar","Bensin"]); jl = st.number_input("Jumlah Liter", 0.0); kt = st.text_area("Keterangan")
                    if st.form_submit_button("Simpan BBM Masuk"):
                        cursor.execute("SELECT id FROM bbm_masuk WHERE lokasi_id=%s AND tanggal=%s AND sumber=%s AND jumlah_liter=%s", (lokasi_id, tg, sm, jl))
                        if cursor.fetchall(): st.warning("⚠️ Data serupa sudah ada!") 
                        cursor.execute("INSERT INTO bbm_masuk (lokasi_id, tanggal, sumber, jenis_bbm, jumlah_liter, keterangan) VALUES (%s,%s,%s,%s,%s,%s)", (lokasi_id, tg, sm, jn, jl, kt))
                        conn.commit(); st.success("Data Masuk Tersimpan!"); st.rerun()
            
            elif mode_transaksi == "📤 PENGGUNAAN BBM":
                with st.form("form_keluar"):
                    c1, c2 = st.columns(2)
                    with c1: tg_p = st.date_input("Tanggal Pakai"); al = st.text_input("Nama Alat/Kendaraan"); un = st.text_input("Kode Unit (Ex: DT-01)")
                    with c2: jl_p = st.number_input("Liter Digunakan", 0.0); kt_p = st.text_area("Keterangan / Lokasi Kerja")
                    if st.form_submit_button("Simpan Penggunaan"):
                        cursor.execute("SELECT id FROM bbm_keluar WHERE lokasi_id=%s AND tanggal=%s AND nama_alat=%s AND no_unit=%s AND jumlah_liter=%s", (lokasi_id, tg_p, al, un, jl_p))
                        if cursor.fetchall(): st.warning("⚠️ Data serupa sudah ada!") 
                        cursor.execute("INSERT INTO bbm_keluar (lokasi_id, tanggal, nama_alat, no_unit, jumlah_liter, keterangan) VALUES (%s,%s,%s,%s,%s,%s)", (lokasi_id, tg_p, al, un, jl_p, kt_p))
                        conn.commit(); st.success("Data Penggunaan Tersimpan!"); st.rerun()

            elif mode_transaksi == "🔄 PINJAM / TRANSFER ANTAR UNIT":
                st.info("ℹ️ Mode ini memindahkan liter dari satu unit ke unit lain. Stok Total BBM tidak berubah.")
                with st.form("form_transfer"):
                    c1, c2 = st.columns(2)
                    with c1: tgl_tf = st.date_input("Tanggal Transfer"); donor_alat = st.text_input("DARI ALAT (Pemberi/Donor)"); donor_unit = st.text_input("No Unit Donor")
                    with c2: liter_tf = st.number_input("Jumlah Liter Dipinjam", min_value=0.0); recv_alat = st.text_input("KE ALAT (Penerima)"); recv_unit = st.text_input("No Unit Penerima"); ket_tf = st.text_area("Keterangan Tambahan")
                    if st.form_submit_button("Proses Transfer"):
                        if liter_tf > 0 and donor_alat and recv_alat:
                            ket_donor = f"Transfer ke {recv_alat} {recv_unit}. {ket_tf}"; cursor.execute("INSERT INTO bbm_keluar (lokasi_id, tanggal, nama_alat, no_unit, jumlah_liter, keterangan) VALUES (%s,%s,%s,%s,%s,%s)", (lokasi_id, tgl_tf, donor_alat, donor_unit, -liter_tf, ket_donor))
                            ket_recv = f"Pinjam dari {donor_alat} {donor_unit}. {ket_tf}"; cursor.execute("INSERT INTO bbm_keluar (lokasi_id, tanggal, nama_alat, no_unit, jumlah_liter, keterangan) VALUES (%s,%s,%s,%s,%s,%s)", (lokasi_id, tgl_tf, recv_alat, recv_unit, liter_tf, ket_recv))
                            conn.commit(); st.success(f"Berhasil transfer {liter_tf}L dari {donor_alat} ke {recv_alat}"); st.rerun()
                        else: st.error("Mohon lengkapi nama alat dan jumlah liter!")

        # --- RIWAYAT INPUT ---
        st.divider()
        with st.expander("⏳ RIWAYAT INPUT & UNDO (Filter & Sortir)", expanded=True):
            c_f1, c_f2, c_f3 = st.columns(3)
            filter_tipe = c_f1.multiselect("Filter Jenis:", ["MASUK", "PAKAI", "TRANSFER", "KOREKSI"], default=["MASUK", "PAKAI", "TRANSFER", "KOREKSI"])
            filter_sort = c_f2.selectbox("Urutkan:", ["Waktu Input Terbaru (ID)", "Waktu Input Terlama (ID)", "Tanggal Laporan Terbaru", "Tanggal Laporan Terlama"])
            use_date_filter = c_f3.checkbox("Filter Tanggal Tertentu")
            date_val = c_f3.date_input("Pilih Tanggal", value=datetime.date.today(), disabled=not use_date_filter)

            history_data = []
            if not df_masuk.empty:
                temp_m = df_masuk.copy(); temp_m['Tipe'] = 'MASUK'; temp_m['Kategori_Filter'] = 'MASUK'; temp_m['Detail'] = temp_m['sumber'] + " (" + temp_m['jenis_bbm'] + ")"; temp_m['Label_History'] = "📥 BBM MASUK (Beli)"
                history_data.append(temp_m[['id', 'tanggal', 'Tipe', 'Detail', 'jumlah_liter', 'Label_History', 'keterangan', 'Kategori_Filter']])
            if not df_keluar.empty:
                temp_k = df_keluar.copy(); temp_k['Tipe'] = 'KELUAR'; temp_k['Detail'] = temp_k['nama_alat'] + " " + temp_k['no_unit']
                def tentukan_label(row):
                    liter = row['jumlah_liter']; ket = str(row['keterangan']).lower()
                    if liter < 0: return ("🔄 TRANSFER KELUAR (Donor)", "TRANSFER")
                    elif "pinjam" in ket or "transfer" in ket: return ("🔄 TRANSFER MASUK (Terima)", "TRANSFER")
                    else: return ("📤 PENGGUNAAN (Pakai)", "PAKAI")
                applied = temp_k.apply(tentukan_label, axis=1)
                temp_k['Label_History'] = [x[0] for x in applied]; temp_k['Kategori_Filter'] = [x[1] for x in applied]
                history_data.append(temp_k[['id', 'tanggal', 'Tipe', 'Detail', 'jumlah_liter', 'Label_History', 'keterangan', 'Kategori_Filter']])
            if not df_log.empty:
                temp_l = df_log.copy(); temp_l['Tipe'] = 'LOG'; temp_l['Kategori_Filter'] = 'KOREKSI'; temp_l['Detail'] = temp_l['kategori']; temp_l['jumlah_liter'] = 0; temp_l['Label_History'] = "🛠️ ADMIN/KOREKSI"; temp_l['keterangan'] = temp_l['deskripsi']
                temp_l['affected_ids_val'] = temp_l['affected_ids'] if 'affected_ids' in temp_l.columns else None
                history_data.append(temp_l[['id', 'tanggal', 'Tipe', 'Detail', 'jumlah_liter', 'Label_History', 'keterangan', 'Kategori_Filter', 'affected_ids_val']])

            if history_data:
                df_history = pd.concat(history_data)
                if filter_tipe: df_history = df_history[df_history['Kategori_Filter'].isin(filter_tipe)]
                if use_date_filter: df_history = df_history[pd.to_datetime(df_history['tanggal']).dt.date == date_val]
                if "Waktu Input Terbaru" in filter_sort: df_history.sort_values(by='id', ascending=False, inplace=True)
                elif "Waktu Input Terlama" in filter_sort: df_history.sort_values(by='id', ascending=True, inplace=True)
                elif "Tanggal Laporan Terbaru" in filter_sort: df_history.sort_values(by='tanggal', ascending=False, inplace=True)
                elif "Tanggal Laporan Terlama" in filter_sort: df_history.sort_values(by='tanggal', ascending=True, inplace=True)

                limit_view = st.selectbox("Jumlah Data Ditampilkan:", ["10", "50", "100", "SEMUA"])
                df_view = df_history.head(int(limit_view)) if limit_view != "SEMUA" else df_history
                st.write(f"Menampilkan **{len(df_view)}** data."); st.info("💡 **Catatan Undo Transfer:** Jika membatalkan transfer, pastikan Anda menghapus **KEDUA** baris (Baris 'Donor' dan Baris 'Terima') agar stok kembali seimbang.")
                
                for index, row in df_view.iterrows():
                    col_a, col_b, col_c, col_d, col_e = st.columns([2, 2, 3, 1, 1])
                    with col_a: 
                        lbl = row['Label_History']
                        if "MASUK" in lbl and "TRANSFER" not in lbl: st.success(lbl)
                        elif "PENGGUNAAN" in lbl: st.warning(lbl)
                        elif "TRANSFER" in lbl: st.info(lbl)
                        elif "ADMIN" in lbl: st.error(lbl)
                        else: st.write(lbl)
                    with col_b: st.write(f"📅 {row['tanggal'].strftime('%d/%m/%Y')}")
                    with col_c: st.write(f"**{row['Detail']}**"); st.caption(f"ket: {row['keterangan']}")
                    with col_d: st.write(f"{row['jumlah_liter']:,.0f} L") if row['Tipe'] != 'LOG' else st.write("-")
                    with col_e:
                        if row['Tipe'] != 'LOG':
                            if st.button("❌ Hapus", key=f"hist_del_{row['Tipe']}_{row['id']}"):
                                table_del = "bbm_masuk" if row['Tipe'] == 'MASUK' else "bbm_keluar"
                                cursor.execute(f"DELETE FROM {table_del} WHERE id=%s", (row['id'],)); conn.commit(); st.success("Data berhasil dihapus!"); st.rerun()
                        else:
                            if st.button("↩️ Undo", key=f"hist_undo_{row['id']}"):
                                try:
                                    matches = re.findall(r"'(.*?)'", row['keterangan'])
                                    if len(matches) >= 2:
                                        old_val = matches[0]; affected_ids = row.get('affected_ids_val')
                                        if affected_ids:
                                            field = "nama_alat" if "NAMA ALAT" in row['Detail'] else "no_unit"
                                            cursor.execute(f"UPDATE bbm_keluar SET {field}='{old_val}' WHERE id IN ({affected_ids})")
                                        cursor.execute("DELETE FROM log_aktivitas WHERE id=%s", (row['id'],)); conn.commit(); st.success(f"Berhasil Undo."); st.rerun()
                                except Exception as e: st.error(f"Error Undo: {e}")
                    st.markdown("---")
            else: st.write("Belum ada riwayat input.")

    with t2:
        tm, tk = df_masuk['jumlah_liter'].sum(), df_keluar['jumlah_liter'].sum(); sisa_stok = stok_awal_val + tm - tk
        st.markdown(f"""<div style="background-color:#d4edda;padding:15px;border-radius:10px;border:1px solid #c3e6cb;text-align:center;margin-bottom:20px;"><h2 style="color:#155724;margin:0;">💰 SISA STOK: {sisa_stok:,.2f} Liter</h2><span style="color:#155724;font-weight:bold;">(Awal: {stok_awal_val:,.0f} + Masuk: {tm:,.0f} - Keluar: {tk:,.0f})</span></div>""", unsafe_allow_html=True)
        
        with st.expander("⚙️ ATUR REKAP (Sembunyikan Unit ke 'Lainnya')", expanded=False):
            st.write("Pilih Unit yang ingin digabung menjadi **'Lainnya'** di tabel Rekapitulasi.")
            if not df_keluar.empty:
                df_keluar['full_name'] = df_keluar['nama_alat'].astype(str) + " " + df_keluar['no_unit'].astype(str)
                unique_units = sorted(df_keluar['full_name'].unique().tolist())
                selected_excludes = st.multiselect("Pilih Unit:", unique_units, default=[u for u in unique_units if u in excluded_list])
                if st.button("Simpan Pengaturan Rekap"):
                    cursor.execute(f"DELETE FROM rekap_exclude WHERE lokasi_id={lokasi_id}")
                    for item in selected_excludes: cursor.execute("INSERT INTO rekap_exclude (lokasi_id, nama_unit_full) VALUES (%s, %s)", (lokasi_id, item))
                    conn.commit(); st.success("Pengaturan Disimpan!"); st.rerun()
            else: st.info("Belum ada data unit keluar.")

        # --- MENU KOREKSI ---
        with st.expander("🛠️ MENU ADMIN (Koreksi Nama/Unit)", expanded=False):
            t_kor, t_hap = st.tabs(["Koreksi Nama", "Hapus Data (Backup)"])
            with t_kor:
                c1, c2 = st.columns(2)
                with c1:
                    list_alat = sorted(df_keluar['nama_alat'].unique().tolist()) if not df_keluar.empty else []
                    if list_alat:
                        pilih_lama = st.selectbox("Alat Salah:", list_alat, key="ot"); input_baru = st.text_input("Nama Benar:", key="nt")
                        if st.button("Ganti Nama Alat"): 
                            cursor.execute(f"SELECT id FROM bbm_keluar WHERE nama_alat='{pilih_lama}' AND lokasi_id={lokasi_id}")
                            ids = [str(r[0]) for r in cursor.fetchall()]; ids_str = ",".join(ids)
                            if ids:
                                cursor.execute("UPDATE bbm_keluar SET nama_alat=%s WHERE nama_alat=%s AND lokasi_id=%s", (input_baru, pilih_lama, lokasi_id))
                                cursor.execute("INSERT INTO log_aktivitas (lokasi_id, kategori, deskripsi, affected_ids) VALUES (%s, %s, %s, %s)", (lokasi_id, "GANTI NAMA ALAT", f"Mengubah '{pilih_lama}' menjadi '{input_baru}'", ids_str))
                                conn.commit(); st.success("Nama Diganti & Dicatat!"); st.rerun()
                with c2:
                    list_unit = sorted(df_keluar['no_unit'].unique().tolist()) if not df_keluar.empty else []
                    if list_unit:
                        pl_u = st.selectbox("Unit Salah:", list_unit, key="ou"); ib_u = st.text_input("Unit Benar:", key="nu")
                        if st.button("Ganti No Unit"): 
                            cursor.execute(f"SELECT id FROM bbm_keluar WHERE no_unit='{pl_u}' AND lokasi_id={lokasi_id}")
                            ids = [str(r[0]) for r in cursor.fetchall()]; ids_str = ",".join(ids)
                            if ids:
                                cursor.execute("UPDATE bbm_keluar SET no_unit=%s WHERE no_unit=%s AND lokasi_id=%s", (ib_u, pl_u, lokasi_id))
                                cursor.execute("INSERT INTO log_aktivitas (lokasi_id, kategori, deskripsi, affected_ids) VALUES (%s, %s, %s, %s)", (lokasi_id, "GANTI NO UNIT", f"Mengubah '{pl_u}' menjadi '{ib_u}'", ids_str))
                                conn.commit(); st.success("Unit Diganti & Dicatat!"); st.rerun()
            with t_hap:
                c1, c2 = st.columns(2)
                with c1:
                    if not df_masuk.empty:
                        m_sel = st.selectbox("Hapus Masuk:", df_masuk.apply(lambda x: f"{x['id']}|{x['tanggal']}|{x['sumber']}", axis=1))
                        if st.button("Hapus Masuk"): cursor.execute(f"DELETE FROM bbm_masuk WHERE id={m_sel.split('|')[0]}"); conn.commit(); st.rerun()
                with c2:
                    if not df_keluar.empty:
                        k_sel = st.selectbox("Hapus Keluar:", df_keluar.apply(lambda x: f"{x['id']}|{x['tanggal']}|{x['nama_alat']}", axis=1))
                        if st.button("Hapus Keluar"): cursor.execute(f"DELETE FROM bbm_keluar WHERE id={k_sel.split('|')[0]}"); conn.commit(); st.rerun()

        # --- GRAFIK & DATA ---
        df_alat_g = pd.DataFrame(); df_truck_g = pd.DataFrame()
        if not df_keluar.empty:
            if 'kategori' not in df_keluar.columns: df_keluar['kategori'] = df_keluar['nama_alat'].apply(cek_kategori)
            df_screen_rekap = process_rekap_data(df_keluar, excluded_list)
            rekap_screen = df_screen_rekap[df_screen_rekap['nama_alat'] != 'Lainnya'].groupby(['kategori', 'nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
            c_rekap1, c_rekap2 = st.columns(2)
            with c_rekap1: st.write("**Rekap Alat Berat**"); st.dataframe(rekap_screen[rekap_screen['kategori']=='ALAT_BERAT'][['nama_alat', 'no_unit', 'jumlah_liter']], hide_index=True)
            with c_rekap2: st.write("**Rekap Mobil/Truck**"); st.dataframe(rekap_screen[rekap_screen['kategori']=='MOBIL_TRUCK'][['nama_alat', 'no_unit', 'jumlah_liter']], hide_index=True)
            df_rekap_for_graph = process_rekap_data(df_keluar, excluded_list)
            if not df_rekap_for_graph.empty:
                    df_alat_g = df_rekap_for_graph[df_rekap_for_graph['kategori'] == 'ALAT_BERAT']
                    df_truck_g = df_rekap_for_graph[df_rekap_for_graph['kategori'] == 'MOBIL_TRUCK']

        st.divider(); col_a, col_b = st.columns(2)
        with col_a: st.subheader("📋 Data Masuk BBM"); st.dataframe(df_masuk.sort_values('tanggal', ascending=False), use_container_width=True)
        with col_b: st.subheader("📋 Data Penggunaan BBM"); st.dataframe(df_keluar.sort_values('tanggal', ascending=False), use_container_width=True)
        st.divider(); st.subheader("📊 Monitoring"); 
        img_buffer = buat_grafik_final(df_alat_g, df_truck_g)
        if img_buffer: st.image(img_buffer, caption="Diagram Alat Berat & Mobil")
        else: st.info("Belum ada data untuk ditampilkan di grafik.")

    with t3:
        st.header("Download Laporan")
        c1, c2, c3 = st.columns(3)
        with c1: st.download_button("📕 PDF", generate_pdf_portrait(df_masuk, df_keluar, nama_proyek, stok_awal_val, excluded_list), f"Laporan_{nama_proyek}.pdf", "application/pdf")
        with c2: st.download_button("📗 Excel", generate_excel_styled(df_masuk, df_keluar, nama_proyek, stok_awal_val, excluded_list), f"Laporan_{nama_proyek}.xlsx")
        with c3: st.download_button("📘 Word", generate_docx_fixed(df_masuk, df_keluar, nama_proyek, stok_awal_val, excluded_list), f"Laporan_{nama_proyek}.docx")

if __name__ == "__main__":
    main()