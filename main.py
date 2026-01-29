import streamlit as st
import mysql.connector
import pandas as pd
import io
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, portrait
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from docx import Document
# Import lengkap untuk Word & XML (untuk warna background)
from docx.shared import Inches, Pt, Cm, RGBColor 
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Sistem BBM Proyek", layout="wide")

# --- KONEKSI DATABASE ---
def init_connection():
    return mysql.connector.connect(
        host=st.secrets["db"]["host"],
        user=st.secrets["db"]["user"],
        password=st.secrets["db"]["password"],
        database=st.secrets["db"]["database"],
        port=st.secrets["db"]["port"]
    )

# --- HELPER ---
def get_hari_indonesia(tanggal):
    try:
        if pd.isnull(tanggal): return "-"
        kamus = {'Monday': 'Senin', 'Tuesday': 'Selasa', 'Wednesday': 'Rabu', 
                 'Thursday': 'Kamis', 'Friday': 'Jumat', 'Saturday': 'Sabtu', 'Sunday': 'Minggu'}
        return kamus[tanggal.strftime('%A')]
    except: return "-"

def cek_kategori(nama_alat):
    nama = str(nama_alat).upper()
    kata_kunci_mobil = ["TRUCK", "MOBIL", "TRITON", "DT", "FAW", "SANNY", "R6", "R10", "PICK UP", "HILUX", "STRADA"]
    if any(k in nama for k in kata_kunci_mobil):
        return "MOBIL_TRUCK"
    return "ALAT_BERAT"

# Fungsi Bantu Warna Word
def set_cell_bg(cell, color_hex):
    """Mengubah warna background cell di Word (docx)"""
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color_hex))
    cell._tc.get_or_add_tcPr().append(shading_elm)

# --- FUNGSI GRAFIK ---
def buat_grafik_final(df_alat, df_truck):
    buffer_img = io.BytesIO()
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(7.5, 8)) 
    
    warna_alat = '#F4B084' 
    warna_truck = '#9BC2E6' 

    if not df_alat.empty:
        df_alat['label'] = df_alat['nama_alat'] + " " + df_alat['no_unit']
        data1 = df_alat.groupby('label')['jumlah_liter'].sum().sort_values()
        data1.plot(kind='barh', ax=ax1, color=warna_alat, edgecolor='black', width=0.6)
        ax1.set_title(f"TOTAL ALAT BERAT ({data1.sum():,.0f} L)", fontsize=10, fontweight='bold')
        ax1.set_ylabel("")
        for i, v in enumerate(data1):
            ax1.text(v, i, f" {v:,.0f}", va='center', fontsize=8)
    else: 
        ax1.text(0.5, 0.5, "Data Kosong", ha='center'); ax1.set_title("ALAT BERAT")

    if not df_truck.empty:
        df_truck['label'] = df_truck['nama_alat'] + " " + df_truck['no_unit']
        data2 = df_truck.groupby('label')['jumlah_liter'].sum().sort_values()
        data2.plot(kind='barh', ax=ax2, color=warna_truck, edgecolor='black', width=0.6)
        ax2.set_title(f"TOTAL MOBIL & TRUCK ({data2.sum():,.0f} L)", fontsize=10, fontweight='bold')
        ax2.set_ylabel("")
        for i, v in enumerate(data2):
            ax2.text(v, i, f" {v:,.0f}", va='center', fontsize=8)
    else: 
        ax2.text(0.5, 0.5, "Data Kosong", ha='center'); ax2.set_title("MOBIL & TRUCK")

    plt.tight_layout()
    plt.savefig(buffer_img, format='png', dpi=100)
    buffer_img.seek(0)
    plt.close(fig)
    return buffer_img

# --- 1. EXPORT PDF (COLOR BY DATE) ---
def generate_pdf_portrait(df_masuk, df_keluar, nama_lokasi):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=portrait(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elements = []
    styles = getSampleStyleSheet()
    
    tm = df_masuk['jumlah_liter'].sum() if not df_masuk.empty else 0
    tk = df_keluar['jumlah_liter'].sum() if not df_keluar.empty else 0
    sisa = tm - tk

    elements.append(Paragraph(f"LAPORAN PENGGUNAAN BBM", styles['Title']))
    elements.append(Paragraph(f"LOKASI: {nama_lokasi}", styles['Normal']))
    stok_text = f"<b>SISA STOK SAAT INI: {sisa:,.2f} Liter</b> (Masuk: {tm:,.0f} - Keluar: {tk:,.0f})"
    elements.append(Spacer(1, 10))
    elements.append(Paragraph(stok_text, styles['Normal']))
    elements.append(Spacer(1, 10))

    # Definisi Warna
    color_odd = colors.white
    color_even = colors.whitesmoke # Abu-abu sangat muda

    def get_row_colors_by_date(dataframe):
        """Menghasilkan list style background berdasarkan perubahan tanggal"""
        row_styles = []
        if dataframe.empty: return row_styles
        
        last_date = None
        is_grey = False # Mulai dengan Putih
        
        # Row index di ReportLab dimulai dari 1 (karena 0 itu header)
        rl_row_idx = 1 
        
        for _, row in dataframe.iterrows():
            curr_date = row['tanggal']
            if curr_date != last_date:
                is_grey = not is_grey # Ganti warna jika tanggal berubah
                last_date = curr_date
            
            if is_grey:
                # Terapkan warna abu-abu
                row_styles.append(('BACKGROUND', (0, rl_row_idx), (-1, rl_row_idx), color_even))
            else:
                # Terapkan warna putih (opsional, default transparent)
                row_styles.append(('BACKGROUND', (0, rl_row_idx), (-1, rl_row_idx), color_odd))
            
            rl_row_idx += 1
        return row_styles

    base_style = [
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), # Header
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('BACKGROUND', (0,-1), (-1,-1), colors.yellow), # Total
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
    ]

    col_widths = [25, 40, 50, 140, 70, 60, 150] 
    
    # TABEL KELUAR
    elements.append(Paragraph("A. RINCIAN PENGGUNAAN (KELUAR)", styles['Heading3']))
    if not df_keluar.empty:
        data = [['NO', 'HARI', 'TGL', 'NAMA ALAT', 'UNIT', 'LITER', 'KET']]
        for i, row in df_keluar.iterrows():
            data.append([i+1, row['HARI'], row['tanggal'].strftime('%d/%m'), row['nama_alat'], row['no_unit'], f"{row['jumlah_liter']:,.0f}", row['keterangan']])
        data.append(['', '', '', 'TOTAL KELUAR', '', f"{tk:,.0f}", ''])
        
        # Apply dynamic styles
        dynamic_styles = get_row_colors_by_date(df_keluar)
        t = Table(data, colWidths=col_widths)
        t.setStyle(TableStyle(base_style + dynamic_styles))
        elements.append(t)
    else: elements.append(Paragraph("- Data Kosong -", styles['Normal']))

    elements.append(Spacer(1, 15))

    # TABEL MASUK
    elements.append(Paragraph("B. RINCIAN BBM MASUK", styles['Heading3']))
    if not df_masuk.empty:
        data_m = [['NO', 'HARI', 'TGL', 'SUMBER', 'JENIS', 'LITER', 'KET']]
        for i, row in df_masuk.iterrows():
            data_m.append([i+1, row['HARI'], row['tanggal'].strftime('%d/%m'), row['sumber'], row['jenis_bbm'], f"{row['jumlah_liter']:,.0f}", row['keterangan']])
        data_m.append(['', '', '', 'TOTAL MASUK', '', f"{tm:,.0f}", ''])
        
        dynamic_styles_m = get_row_colors_by_date(df_masuk)
        tm_tab = Table(data_m, colWidths=col_widths)
        tm_tab.setStyle(TableStyle(base_style + dynamic_styles_m))
        elements.append(tm_tab)
    else: elements.append(Paragraph("- Data Kosong -", styles['Normal']))

    # REKAP
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("REKAPITULASI & DIAGRAM", styles['Heading2']))

    if not df_keluar.empty and 'kategori' not in df_keluar.columns:
         df_keluar['kategori'] = df_keluar['nama_alat'].apply(cek_kategori)
         
    df_alat = df_keluar[df_keluar['kategori'] == 'ALAT_BERAT'] if not df_keluar.empty else pd.DataFrame()
    df_truck = df_keluar[df_keluar['kategori'] == 'MOBIL_TRUCK'] if not df_keluar.empty else pd.DataFrame()

    if not df_alat.empty:
        rekap_a = df_alat.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
        data_ra = [[Paragraph('<b>TOTAL ALAT BERAT</b>', styles['Normal']), '']]
        for _, row in rekap_a.iterrows(): data_ra.append([f"{row['nama_alat']} {row['no_unit']}", f"{row['jumlah_liter']:,.0f} Liter"])
        data_ra.append(['TOTAL', f"{rekap_a['jumlah_liter'].sum():,.0f} Liter"])
        tr = Table(data_ra, colWidths=[300, 150])
        tr.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black),('BACKGROUND', (0,0), (1,0), colors.orange),('BACKGROUND', (0,-1), (-1,-1), colors.yellow),('ALIGN', (1,0), (1,-1), 'RIGHT')]))
        elements.append(tr); elements.append(Spacer(1, 10))

    if not df_truck.empty:
        rekap_t = df_truck.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
        data_rt = [[Paragraph('<b>TOTAL MOBIL & TRUCK</b>', styles['Normal']), '']]
        for _, row in rekap_t.iterrows(): data_rt.append([f"{row['nama_alat']} {row['no_unit']}", f"{row['jumlah_liter']:,.0f} Liter"])
        data_rt.append(['TOTAL', f"{rekap_t['jumlah_liter'].sum():,.0f} Liter"])
        tt = Table(data_rt, colWidths=[300, 150])
        tt.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black),('BACKGROUND', (0,0), (1,0), colors.skyblue),('BACKGROUND', (0,-1), (-1,-1), colors.yellow),('ALIGN', (1,0), (1,-1), 'RIGHT')]))
        elements.append(tt); elements.append(Spacer(1, 10))

    img_buffer = buat_grafik_final(df_alat, df_truck)
    if img_buffer:
        img = RLImage(img_buffer, width=450, height=500)
        elements.append(img)

    doc.build(elements)
    buffer.seek(0)
    return buffer

# --- 2. EXPORT EXCEL (COLOR BY DATE) ---
def generate_excel_styled(df_masuk, df_keluar, nama_lokasi):
    output = io.BytesIO()
    wb = Workbook()
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    grey_fill = PatternFill(start_color="D3D3D3", fill_type="solid")
    
    # Warna latar baris (Zebra per hari)
    row_grey_fill = PatternFill(start_color="F2F2F2", fill_type="solid") # Abu muda banget
    
    header_font = Font(bold=True, size=11)
    title_font = Font(bold=True, size=14)
    
    tm = df_masuk['jumlah_liter'].sum() if not df_masuk.empty else 0
    tk = df_keluar['jumlah_liter'].sum() if not df_keluar.empty else 0
    
    ws = wb.active; ws.title = "Laporan BBM"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.fitToWidth = 1
    
    ws.merge_cells('A1:G1'); ws['A1'] = f"LAPORAN BBM: {nama_lokasi}"; ws['A1'].font = title_font; ws['A1'].alignment = Alignment(horizontal='center')
    ws.merge_cells('A2:G2'); ws['A2'] = f"SISA STOK SAAT INI: {tm - tk:,.2f} Liter"; ws['A2'].font = Font(bold=True, size=12, color="FF0000"); ws['A2'].alignment = Alignment(horizontal='center')

    current_row = 4
    
    def buat_tabel_excel_colored(ws, start_row, title, headers, dataframe, data_cols, total_label, total_val):
        ws.cell(row=start_row, column=1, value=title).font = header_font
        start_row += 1
        for col_num, h in enumerate(headers, 1):
            c = ws.cell(row=start_row, column=col_num, value=h)
            c.fill = grey_fill; c.border = thin_border; c.font = header_font; c.alignment = Alignment(horizontal='center')
        start_row += 1
        
        if not dataframe.empty:
            last_date = None
            is_grey = False # Mulai putih
            
            for i, row in dataframe.iterrows():
                curr_date = row['tanggal']
                if curr_date != last_date:
                    is_grey = not is_grey
                    last_date = curr_date
                
                # Tentukan warna baris
                current_fill = row_grey_fill if is_grey else None

                ws.cell(row=start_row, column=1, value=i+1).border = thin_border
                if current_fill: ws.cell(row=start_row, column=1).fill = current_fill
                
                col_idx = 2
                for col_name in data_cols:
                    val = row[col_name]
                    if col_name == 'tanggal': val = val.strftime('%d/%m/%Y')
                    c = ws.cell(row=start_row, column=col_idx, value=val)
                    c.border = thin_border
                    if current_fill: c.fill = current_fill
                    if col_name == 'jumlah_liter': c.number_format = '#,##0'
                    col_idx += 1
                start_row += 1
            
            # Total Row
            ws.cell(row=start_row, column=4, value=total_label).font = header_font
            c_tot = ws.cell(row=start_row, column=6, value=total_val)
            c_tot.font = header_font; c_tot.number_format = '#,##0'
            for c in range(1, 8): ws.cell(row=start_row, column=c).fill = yellow_fill; ws.cell(row=start_row, column=c).border = thin_border
            start_row += 2
        else: ws.cell(row=start_row, column=1, value="Data Kosong"); start_row += 2
        return start_row

    current_row = buat_tabel_excel_colored(ws, current_row, "A. PENGGUNAAN BBM (KELUAR)", ['NO', 'HARI', 'TANGGAL', 'NAMA ALAT', 'NO UNIT', 'LITER', 'KETERANGAN'], df_keluar, ['HARI', 'tanggal', 'nama_alat', 'no_unit', 'jumlah_liter', 'keterangan'], "TOTAL KELUAR", tk)
    current_row = buat_tabel_excel_colored(ws, current_row, "B. BBM MASUK", ['NO', 'HARI', 'TANGGAL', 'SUMBER', 'JENIS', 'LITER', 'KETERANGAN'], df_masuk, ['HARI', 'tanggal', 'sumber', 'jenis_bbm', 'jumlah_liter', 'keterangan'], "TOTAL MASUK", tm)

    ws.cell(row=current_row, column=1, value="C. REKAPITULASI UNIT").font = title_font
    current_row += 1
    
    if not df_keluar.empty:
        if 'kategori' not in df_keluar.columns: df_keluar['kategori'] = df_keluar['nama_alat'].apply(cek_kategori)
            
    df_alat = df_keluar[df_keluar['kategori'] == 'ALAT_BERAT'] if not df_keluar.empty else pd.DataFrame()
    df_truck = df_keluar[df_keluar['kategori'] == 'MOBIL_TRUCK'] if not df_keluar.empty else pd.DataFrame()

    def buat_rekap_excel(ws, r_row, title, df_sub, color_fill):
        ws.merge_cells(f'A{r_row}:B{r_row}')
        c = ws.cell(row=r_row, column=1, value=title)
        c.font = header_font; c.fill = color_fill; c.border = thin_border; c.alignment = Alignment(horizontal='center')
        r_row += 1
        if not df_sub.empty:
            rekap = df_sub.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
            for _, row in rekap.iterrows():
                ws.cell(row=r_row, column=1, value=f"{row['nama_alat']} {row['no_unit']}").border = thin_border
                c_val = ws.cell(row=r_row, column=2, value=row['jumlah_liter'])
                c_val.border = thin_border; c_val.number_format = '#,##0'
                r_row += 1
            ws.cell(row=r_row, column=1, value="TOTAL").fill = yellow_fill; ws.cell(row=r_row, column=1).border = thin_border
            c_tot = ws.cell(row=r_row, column=2, value=rekap['jumlah_liter'].sum())
            c_tot.fill = yellow_fill; c_tot.border = thin_border; c_tot.number_format = '#,##0'
            r_row += 2
        return r_row

    row_rekap_start = current_row
    current_row = buat_rekap_excel(ws, current_row, "TOTAL ALAT BERAT", df_alat, PatternFill(start_color="F4B084", fill_type="solid"))
    current_row = buat_rekap_excel(ws, current_row, "TOTAL MOBIL & TRUCK", df_truck, PatternFill(start_color="9BC2E6", fill_type="solid"))

    img_buffer = buat_grafik_final(df_alat, df_truck)
    if img_buffer:
        try: img = XLImage(img_buffer); ws.add_image(img, f'D{row_rekap_start}') 
        except: pass

    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']: ws.column_dimensions[col].width = 15
    ws.column_dimensions['A'].width = 5; ws.column_dimensions['D'].width = 25; ws.column_dimensions['G'].width = 25
    wb.save(output); output.seek(0)
    return output

# --- 3. EXPORT DOCX (COLOR BY DATE) ---
def generate_docx_fixed(df_masuk, df_keluar, nama_lokasi):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(1.27); section.bottom_margin = Cm(1.27)
        section.left_margin = Cm(1.27); section.right_margin = Cm(1.27)

    tm = df_masuk['jumlah_liter'].sum() if not df_masuk.empty else 0
    tk = df_keluar['jumlah_liter'].sum() if not df_keluar.empty else 0

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"LAPORAN BBM: {nama_lokasi}"); r.bold = True; r.font.size = Pt(14)
    p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(f"SISA STOK: {tm - tk:,.2f} Liter"); r2.bold = True; r2.font.color.rgb = RGBColor(255, 0, 0)

    def add_word_table_colored(doc, title, headers, dataframe, data_cols, total_label, total_val):
        doc.add_paragraph(title, style='Heading 3')
        if not dataframe.empty:
            table = doc.add_table(rows=1, cols=len(headers)); table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            for i, h in enumerate(headers): hdr_cells[i].text = h; hdr_cells[i].paragraphs[0].runs[0].bold = True
            
            last_date = None
            is_grey = False
            
            for _, row in dataframe.iterrows():
                curr_date = row['tanggal']
                if curr_date != last_date:
                    is_grey = not is_grey
                    last_date = curr_date
                
                cells = table.add_row().cells
                
                # Jika abu-abu, warnai semua sel di baris ini
                if is_grey:
                    for cell in cells: set_cell_bg(cell, "F2F2F2") # Hex Abu Muda
                
                cells[0].text = str(row['HARI']); cells[1].text = row['tanggal'].strftime('%d/%m')
                col_idx = 2
                for col in data_cols[2:]:
                    val = row[col]; 
                    if col == 'jumlah_liter': val = f"{val:.0f}"
                    cells[col_idx].text = str(val); col_idx += 1
            
            row = table.add_row().cells
            row[2].text = total_label; row[4].text = f"{total_val:.0f}"; row[2].paragraphs[0].runs[0].bold = True
            for cell in row: set_cell_bg(cell, "FFFF00") # Kuning Total

    add_word_table_colored(doc, "A. PENGGUNAAN BBM (KELUAR)", ['HARI', 'TGL', 'ALAT', 'UNIT', 'LITER', 'KET'], df_keluar, ['HARI', 'tanggal', 'nama_alat', 'no_unit', 'jumlah_liter', 'keterangan'], "TOTAL KELUAR", tk)
    add_word_table_colored(doc, "B. BBM MASUK", ['HARI', 'TGL', 'SUMBER', 'JENIS', 'LITER', 'KET'], df_masuk, ['HARI', 'tanggal', 'sumber', 'jenis_bbm', 'jumlah_liter', 'keterangan'], "TOTAL MASUK", tm)

    doc.add_paragraph("C. REKAPITULASI UNIT", style='Heading 2')
    
    if not df_keluar.empty:
        if 'kategori' not in df_keluar.columns: df_keluar['kategori'] = df_keluar['nama_alat'].apply(cek_kategori)
            
    df_alat = df_keluar[df_keluar['kategori'] == 'ALAT_BERAT'] if not df_keluar.empty else pd.DataFrame()
    df_truck = df_keluar[df_keluar['kategori'] == 'MOBIL_TRUCK'] if not df_keluar.empty else pd.DataFrame()

    if not df_alat.empty:
        doc.add_paragraph("TOTAL ALAT BERAT", style='Heading 4')
        table = doc.add_table(rows=1, cols=2); table.style = 'Table Grid'
        table.rows[0].cells[0].text = "NAMA ALAT"; table.rows[0].cells[1].text = "LITER"
        rekap_a = df_alat.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
        for _, row in rekap_a.iterrows():
            cells = table.add_row().cells
            cells[0].text = f"{row['nama_alat']} {row['no_unit']}"; cells[1].text = f"{row['jumlah_liter']:.0f}"
        row = table.add_row().cells
        row[0].text = "TOTAL"; row[1].text = f"{rekap_a['jumlah_liter'].sum():.0f}"; row[0].paragraphs[0].runs[0].bold = True

    if not df_truck.empty:
        doc.add_paragraph("TOTAL MOBIL & TRUCK", style='Heading 4')
        table = doc.add_table(rows=1, cols=2); table.style = 'Table Grid'
        table.rows[0].cells[0].text = "NAMA UNIT"; table.rows[0].cells[1].text = "LITER"
        rekap_t = df_truck.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
        for _, row in rekap_t.iterrows():
            cells = table.add_row().cells
            cells[0].text = f"{row['nama_alat']} {row['no_unit']}"; cells[1].text = f"{row['jumlah_liter']:.0f}"
        row = table.add_row().cells
        row[0].text = "TOTAL"; row[1].text = f"{rekap_t['jumlah_liter'].sum():.0f}"; row[0].paragraphs[0].runs[0].bold = True

    img_buffer = buat_grafik_final(df_alat, df_truck)
    if img_buffer: doc.add_picture(img_buffer, width=Cm(16))
    buffer = io.BytesIO(); doc.save(buffer); buffer.seek(0)
    return buffer

# --- MAIN APP ---
def main():
    st.title("🚜 Sistem Monitoring BBM")
    try: conn = init_connection(); cursor = conn.cursor()
    except: st.error("Database Error"); st.stop()

    # --- LOGIN SEDERHANA ---
    password_rahasia = "123" # Ganti dengan password yang diinginkan

    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if not st.session_state.logged_in:
        pwd = st.text_input("Masukkan Password Admin:", type="password")
        if st.button("Login"):
            if pwd == password_rahasia:
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("Password Salah!")
        st.stop() # Berhenti disini jika belum login

    #--- SIDEBAR LOKASI PROYEK ---
    st.sidebar.header("📍 Lokasi Proyek")
    try: df_lokasi = pd.read_sql("SELECT * FROM lokasi_proyek", conn)
    except: st.error("Tabel belum siap"); st.stop()
    
    lokasi_id, pil = None, None
    if not df_lokasi.empty:
        pil = st.sidebar.selectbox("Pilih Lokasi", df_lokasi['nama_tempat'])
        lokasi_id = int(df_lokasi[df_lokasi['nama_tempat'] == pil]['id'].values[0])
    
    with st.sidebar.expander("➕ Lokasi Baru"):
        nb = st.text_input("Nama"); 
        if st.button("Simpan"): cursor.execute("INSERT INTO lokasi_proyek (nama_tempat) VALUES (%s)", (nb,)); conn.commit(); st.rerun()

    df_masuk = pd.DataFrame(); df_keluar = pd.DataFrame()
    if lokasi_id:
        df_masuk = pd.read_sql(f"SELECT * FROM bbm_masuk WHERE lokasi_id={lokasi_id} ORDER BY tanggal ASC, id ASC", conn)
        df_keluar = pd.read_sql(f"SELECT * FROM bbm_keluar WHERE lokasi_id={lokasi_id} ORDER BY tanggal ASC, id ASC", conn)
        if not df_masuk.empty: df_masuk['HARI'] = pd.to_datetime(df_masuk['tanggal']).apply(get_hari_indonesia)
        if not df_keluar.empty: df_keluar['HARI'] = pd.to_datetime(df_keluar['tanggal']).apply(get_hari_indonesia)

    t1, t2, t3 = st.tabs(["📝 Input Data", "📊 Laporan & Grafik", "🖨️ Export Dokumen"])
    
    with t1:
        if lokasi_id:
            c1, c2 = st.columns(2)
            with c1:
                with st.form("f1"):
                    st.info("MASUK BBM")
                    tg = st.date_input("Tgl Masuk"); sm = st.text_input("Sumber"); jn = st.selectbox("Jenis", ["Dexlite","Solar","Bensin"])
                    jl = st.number_input("Liter", 0.0); kt = st.text_area("Ket")
                    if st.form_submit_button("Simpan Masuk"):
                        cursor.execute("INSERT INTO bbm_masuk (lokasi_id, tanggal, sumber, jenis_bbm, jumlah_liter, keterangan) VALUES (%s,%s,%s,%s,%s,%s)", (lokasi_id, tg, sm, jn, jl, kt))
                        conn.commit(); st.success("Ok"); st.rerun()
            with c2:
                with st.form("f2"):
                    st.warning("PENGGUNAAN BBM")
                    tg_p = st.date_input("Tgl Pakai"); al = st.text_input("Alat"); un = st.text_input("Unit"); jl_p = st.number_input("Liter Pakai", 0.0); kt_p = st.text_area("Ket Penggunaan")
                    if st.form_submit_button("Simpan Pakai"):
                        cursor.execute("INSERT INTO bbm_keluar (lokasi_id, tanggal, nama_alat, no_unit, jumlah_liter, keterangan) VALUES (%s,%s,%s,%s,%s,%s)", (lokasi_id, tg_p, al, un, jl_p, kt_p))
                        conn.commit(); st.success("Ok"); st.rerun()

    with t2:
        if lokasi_id:
            tm, tk = df_masuk['jumlah_liter'].sum(), df_keluar['jumlah_liter'].sum()
            st.markdown(f"""<div style="background-color:#d4edda;padding:15px;border-radius:10px;border:1px solid #c3e6cb;text-align:center;margin-bottom:20px;">
                <h2 style="color:#155724;margin:0;">💰 SISA STOK: {tm-tk:,.2f} Liter</h2>
                <span style="color:#155724;font-weight:bold;">Total Masuk: {tm:,.0f} | Total Keluar: {tk:,.0f}</span></div>""", unsafe_allow_html=True)
            
            with st.expander("🗑️ MENU HAPUS DATA SALAH (Klik disini)", expanded=False):
                col_del1, col_del2 = st.columns(2)
                with col_del1:
                    st.warning("HAPUS DATA MASUK BBM")
                    if not df_masuk.empty:
                        df_m_sorted = df_masuk.sort_values(by=['tanggal', 'id'], ascending=[False, False])
                        pilihan_m = df_m_sorted.apply(lambda x: f"{x['id']} | {x['tanggal']} | {x['sumber']} | {x['jumlah_liter']}L", axis=1)
                        sel_m = st.selectbox("Pilih Data Masuk:", pilihan_m, key="del_m")
                        if st.button("Hapus Data Masuk", key="btn_del_m"):
                            id_del = sel_m.split(" | ")[0]
                            cursor.execute(f"DELETE FROM bbm_masuk WHERE id={id_del}")
                            conn.commit(); st.success("Terhapus!"); st.rerun()
                    else: st.write("Data Kosong")
                with col_del2:
                    st.warning("HAPUS DATA PENGGUNAAN BBM")
                    if not df_keluar.empty:
                        df_k_sorted = df_keluar.sort_values(by=['tanggal', 'id'], ascending=[False, False])
                        pilihan_k = df_k_sorted.apply(lambda x: f"{x['id']} | {x['tanggal']} | {x['nama_alat']} | {x['jumlah_liter']}L", axis=1)
                        sel_k = st.selectbox("Pilih Data Keluar:", pilihan_k, key="del_k")
                        if st.button("Hapus Data Keluar", key="btn_del_k"):
                            id_del = sel_k.split(" | ")[0]
                            cursor.execute(f"DELETE FROM bbm_keluar WHERE id={id_del}")
                            conn.commit(); st.success("Terhapus!"); st.rerun()
                    else: st.write("Data Kosong")

            if not df_keluar.empty:
                df_keluar['kategori'] = df_keluar['nama_alat'].apply(cek_kategori)
                rekap_screen = df_keluar.groupby(['kategori', 'nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
                
                c_rekap1, c_rekap2 = st.columns(2)
                with c_rekap1: 
                    st.write("**Rekap Alat Berat**")
                    st.dataframe(rekap_screen[rekap_screen['kategori']=='ALAT_BERAT'][['nama_alat', 'no_unit', 'jumlah_liter']], hide_index=True)
                with c_rekap2: 
                    st.write("**Rekap Mobil/Truck**")
                    st.dataframe(rekap_screen[rekap_screen['kategori']=='MOBIL_TRUCK'][['nama_alat', 'no_unit', 'jumlah_liter']], hide_index=True)

            st.divider()
            col_a, col_b = st.columns(2)
            with col_a: st.subheader("📋 Data Masuk BBM"); st.dataframe(df_masuk, use_container_width=True)
            with col_b: st.subheader("📋 Data Penggunaan BBM"); st.dataframe(df_keluar, use_container_width=True)

            st.divider(); st.subheader("📊 Monitoring")
            df_alat = df_keluar[df_keluar['kategori'] == 'ALAT_BERAT'] if not df_keluar.empty else pd.DataFrame()
            df_truck = df_keluar[df_keluar['kategori'] == 'MOBIL_TRUCK'] if not df_keluar.empty else pd.DataFrame()
            img_show = buat_grafik_final(df_alat, df_truck)
            st.image(img_show, caption="Diagram Alat Berat & Mobil", use_column_width=False)

    with t3:
        if lokasi_id:
            st.header("Download Laporan")
            c1, c2, c3 = st.columns(3)
            with c1: st.download_button("📕 PDF ", generate_pdf_portrait(df_masuk, df_keluar, pil), f"Laporan_{pil}.pdf", "application/pdf")
            with c2: st.download_button("📗 Excel (Disarankan)", generate_excel_styled(df_masuk, df_keluar, pil), f"Laporan_{pil}.xlsx")
            with c3: st.download_button("📘 Word ", generate_docx_fixed(df_masuk, df_keluar, pil), f"Laporan_{pil}.docx")

if __name__ == "__main__":
    main()