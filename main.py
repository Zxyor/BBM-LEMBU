import streamlit as st
import mysql.connector
import pandas as pd
import io
import datetime
import math
import re
from dateutil.relativedelta import relativedelta

# --- SETUP MATPLOTLIB ---
import matplotlib
matplotlib.use('Agg') 
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg
import matplotlib.ticker as ticker

# --- LIBRARY REPORTING ---
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, portrait
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.drawing.image import Image as XLImage

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
def get_bulan_indonesia(bulan_int):
    nama_bulan = ["", "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", 
                  "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"]
    return nama_bulan[bulan_int]

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

def segregate_data(df, excluded_list):
    if df.empty:
        empty_df = pd.DataFrame(columns=['nama_alat', 'no_unit', 'jumlah_liter', 'kategori'])
        return empty_df, empty_df, empty_df
    
    df = df.copy()
    df['full_name'] = df['nama_alat'].astype(str) + " " + df['no_unit'].astype(str)
    
    df_lainnya = df[df['full_name'].isin(excluded_list)].copy()
    df_main = df[~df['full_name'].isin(excluded_list)].copy()
    
    df_alat = df_main[df_main['kategori'] == 'ALAT_BERAT']
    df_truck = df_main[df_main['kategori'] == 'MOBIL_TRUCK']
    
    return df_alat, df_truck, df_lainnya

def filter_non_consumption(df):
    if df.empty: return df
    mask_donor = df['jumlah_liter'] < 0
    mask_recv = df['keterangan'].str.contains('Transfer|Pinjam', case=False, na=False)
    return df[~(mask_donor | mask_recv)]

def hitung_stok_awal_periode(conn, lokasi_id, start_date):
    cursor = conn.cursor()
    cursor.execute("SELECT stok_awal FROM lokasi_proyek WHERE id = %s", (lokasi_id,))
    res = cursor.fetchone()
    modal_awal = float(res[0]) if res and res[0] is not None else 0.0
    cursor.execute("SELECT COALESCE(SUM(jumlah_liter), 0) FROM bbm_masuk WHERE lokasi_id = %s AND tanggal < %s", (lokasi_id, start_date))
    res_m = cursor.fetchone()
    masuk_prev = float(res_m[0])
    cursor.execute("SELECT COALESCE(SUM(jumlah_liter), 0) FROM bbm_keluar WHERE lokasi_id = %s AND tanggal < %s", (lokasi_id, start_date))
    res_k = cursor.fetchone()
    keluar_prev = float(res_k[0])
    return modal_awal + masuk_prev - keluar_prev

def split_date_range_by_month(start_date, end_date):
    result = []
    current = start_date.replace(day=1)
    while current <= end_date:
        month_start = max(start_date, current)
        next_month = current + relativedelta(months=1)
        month_end = min(end_date, next_month - datetime.timedelta(days=1))
        if month_start <= month_end:
            result.append((month_start, month_end))
        current = next_month
    return result

def safe_text(text, max_chars=25):
    s = str(text) if text else "-"
    if len(s) > max_chars:
        return s[:max_chars] + "..."
    return s

# --- CHART GENERATOR ---
def generate_chart_sidebar(df_alat, df_truck):
    try:
        active_charts = []
        if not df_alat.empty: active_charts.append(("PEMAKAIAN ALAT", df_alat, '#F4B084'))
        if not df_truck.empty: active_charts.append(("PEMAKAIAN MOBIL", df_truck, '#9BC2E6'))
        
        num_charts = len(active_charts)
        if num_charts == 0: return None
        
        # Grafik
        fig = Figure(figsize=(3.5, 2.5 * num_charts), dpi=100) 
        canvas = FigureCanvasAgg(fig)
        axs = fig.subplots(num_charts, 1)
        if num_charts == 1: axs = [axs]
        
        for i, (title, df, color) in enumerate(active_charts):
            ax = axs[i]
            rekap = df.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
            rekap['label'] = rekap.apply(lambda x: safe_text(f"{x['nama_alat']} {x['no_unit']}", 20), axis=1)
            # Top 10 Only
            data = rekap.groupby('label')['jumlah_liter'].sum().sort_values().tail(10)
            
            bars = ax.barh(data.index, data.values, color=color, edgecolor='#555555', height=0.6)
            ax.set_title(title, fontsize=9, fontweight='bold', color='#333333')
            ax.tick_params(labelsize=7, colors='#333333')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            for bar in bars: ax.text(bar.get_width(), bar.get_y()+bar.get_height()/2, f" {bar.get_width():,.0f}", va='center', fontsize=7, color='black')

        fig.tight_layout()
        buf = io.BytesIO()
        canvas.print_png(buf)
        buf.seek(0)
        return buf
    except: return None

def generate_monthly_chart(df_monthly):
    try:
        if df_monthly.empty: return None
        masuk_vals = pd.to_numeric(df_monthly['masuk'], errors='coerce').fillna(0)
        keluar_vals = pd.to_numeric(df_monthly['keluar'], errors='coerce').fillna(0)
        labels = df_monthly['bulan_nama'].astype(str).tolist()

        fig = Figure(figsize=(8, 4), dpi=100)
        canvas = FigureCanvasAgg(fig)
        ax = fig.add_subplot(111)
        x = range(len(labels))
        width = 0.35
        ax.bar([i - width/2 for i in x], masuk_vals, width, label='Masuk', color='#90EE90', edgecolor='black')
        ax.bar([i + width/2 for i in x], keluar_vals, width, label='Keluar', color='#F08080', edgecolor='black')
        
        ax.set_title('GRAFIK MASUK & PENGGUNAAN BBM PER BULAN', fontsize=10, fontweight='bold')
        ax.set_xticks(x)
        ax.set_xticklabels(labels, fontsize=8)
        ax.legend(fontsize=8)
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))
        ax.tick_params(labelsize=8)
        fig.tight_layout()
        buf = io.BytesIO()
        canvas.print_png(buf)
        buf.seek(0)
        return buf
    except: return None

# --- EXPORT PDF
def generate_pdf_portrait(conn, lokasi_id, nama_lokasi, start_date_global, end_date_global, excluded_list):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=portrait(A4), rightMargin=15, leftMargin=15, topMargin=20, bottomMargin=20)
    elements = []
    
    # --- STYLES ---
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(name='ExcelTitle', parent=styles['Heading1'], alignment=TA_CENTER, fontSize=14, fontName='Helvetica-Bold', spaceAfter=2, textColor=colors.HexColor("#2F5496"))
    periode_style = ParagraphStyle(name='ExcelPeriode', parent=styles['Normal'], alignment=TA_CENTER, fontSize=11, spaceAfter=15, textColor=colors.black)
    
    cell_style = ParagraphStyle(name='CellText', parent=styles['Normal'], fontSize=7, leading=8, fontName='Helvetica')
    header_style = ParagraphStyle(name='HeaderTxt', parent=styles['Normal'], fontSize=7, leading=8, fontName='Helvetica-Bold', textColor=colors.white, alignment=TA_CENTER)
    header_black_style = ParagraphStyle(name='HeaderTxtBlk', parent=styles['Normal'], fontSize=7, leading=8, fontName='Helvetica-Bold', textColor=colors.black, alignment=TA_CENTER)
    section_title_style = ParagraphStyle(name='SectionTitle', parent=styles['Normal'], fontSize=8, leading=9, fontName='Helvetica-Bold', textColor=colors.HexColor("#2F5496"))
    
    # Warna
    COLOR_HEADER_BLUE = colors.HexColor("#2F5496") 
    COLOR_ROW_EVEN = colors.HexColor("#F2F2F2")    
    COLOR_ROW_ODD = colors.white
    COLOR_TOTAL_YELLOW = colors.HexColor("#FFD966") 
    COLOR_BORDER = colors.HexColor("#BFBFBF")       

    date_ranges = split_date_range_by_month(start_date_global, end_date_global)
    
    for idx, (start_date, end_date) in enumerate(date_ranges):
        if idx > 0: elements.append(PageBreak())

        # DATA FETCHING
        stok_awal = hitung_stok_awal_periode(conn, lokasi_id, start_date)
        df_masuk = pd.read_sql(f"SELECT * FROM bbm_masuk WHERE lokasi_id={lokasi_id} AND tanggal BETWEEN '{start_date}' AND '{end_date}' ORDER BY tanggal", conn)
        df_keluar = pd.read_sql(f"SELECT * FROM bbm_keluar WHERE lokasi_id={lokasi_id} AND tanggal BETWEEN '{start_date}' AND '{end_date}' ORDER BY tanggal", conn)
        
        if not df_keluar.empty:
            df_keluar['HARI'] = pd.to_datetime(df_keluar['tanggal']).apply(get_hari_indonesia)
            if 'kategori' not in df_keluar.columns: df_keluar['kategori'] = df_keluar['nama_alat'].apply(cek_kategori)
        
        df_keluar_rpt = filter_non_consumption(df_keluar)
        df_alat_g, df_truck_g, df_lain_g = segregate_data(df_keluar_rpt, excluded_list)
        
        tm = float(df_masuk['jumlah_liter'].sum()) if not df_masuk.empty else 0.0
        tk_real = float(df_keluar['jumlah_liter'].sum()) if not df_keluar.empty else 0.0
        tk_rpt = float(df_keluar_rpt['jumlah_liter'].sum()) if not df_keluar_rpt.empty else 0.0
        sisa_akhir = stok_awal + tm - tk_real

        # Header Page
        elements.append(Paragraph(f"LAPORAN BBM: {nama_lokasi}", title_style))
        elements.append(Paragraph(f"PERIODE {get_bulan_indonesia(start_date.month)} {start_date.year}", periode_style))
        
        # --- PREPARE QUEUES ---
        left_queue = []
        left_queue.append({'type': 'title_section', 'val': 'A. PENGGUNAAN BBM (KELUAR)'})
        left_queue.append({'type': 'header_col'}) 
        
        if not df_keluar_rpt.empty:
            for i, r in df_keluar_rpt.iterrows():
                left_queue.append({
                    'type': 'row', 
                    'data': [i+1, r['tanggal'].strftime('%d/%m'), r['nama_alat'], r['no_unit'], f"{r['jumlah_liter']:.0f}", r['keterangan']]
                })
        left_queue.append({'type': 'total_left', 'val': f"{tk_rpt:.0f}"})

        right_queue = []
        
        # 1. BBM MASUK
        right_queue.append({'type': 'title_section', 'val': 'B. BBM MASUK'})
        right_queue.append({'type': 'header_masuk'})
        if not df_masuk.empty:
            for i, r in df_masuk.iterrows():
                right_queue.append({'type': 'row_masuk', 'data': [i+1, r['tanggal'].strftime('%d/%m'), r['sumber'], r['jenis_bbm'], f"{r['jumlah_liter']:.0f}"]})
        else:
            right_queue.append({'type': 'row_masuk', 'data': ['-', '-', 'TIDAK ADA DATA', '-', '0']})
        right_queue.append({'type': 'total_masuk', 'val': f"{tm:.0f}"})
        
        # 2. REKAP
        right_queue.append({'type': 'title_section', 'val': 'RINCIAN PENGGUNAAN BBM'})
        
        def add_rekap(df, title, color, text_is_black=False):
            if title == "LAINNYA":
                 color = "#ED77C4" 
                 text_is_black = False 
            
            right_queue.append({'type': 'sub_rekap', 'title': title, 'bg': color, 'txt_black': text_is_black})
            if not df.empty:
                grp = df.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index().sort_values('jumlah_liter', ascending=False)
                for _, r in grp.iterrows():
                    right_queue.append({'type': 'row_rekap', 'label': f"{r['nama_alat']} {r['no_unit']}", 'val': f"{r['jumlah_liter']:.0f}"})
                right_queue.append({'type': 'total_rekap', 'val': f"{df['jumlah_liter'].sum():.0f}"})
            else:
                 right_queue.append({'type': 'row_rekap', 'label': '-', 'val': '0'})
                 right_queue.append({'type': 'total_rekap', 'val': '0'})

        add_rekap(df_alat_g, "TOTAL ALAT BERAT", "#F4B084", True)
        add_rekap(df_truck_g, "TOTAL MOBIL & TRUCK", "#9BC2E6", True)
        if not df_lain_g.empty: 
            add_rekap(df_lain_g, "LAINNYA", "#ED77C4", False)
        
        # 3. STOK
        right_queue.append({'type': 'title_section', 'val': 'RINCIAN SISA STOK BBM'})
        right_queue.append({'type': 'header_stok', 'label': 'RINGKASAN STOK'})
        right_queue.append({'type': 'row_stok', 'label': 'SISA BULAN LALU', 'val': f"{stok_awal:.0f}"})
        right_queue.append({'type': 'row_stok', 'label': 'TOTAL MASUK', 'val': f"{tm:.0f}"})
        right_queue.append({'type': 'row_stok', 'label': 'TOTAL KELUAR', 'val': f"{tk_real:.0f}"})
        right_queue.append({'type': 'total_stok', 'label': 'SISA AKHIR', 'val': f"{sisa_akhir:.0f}"})
        
        # 4. CHART
        img_buf = generate_chart_sidebar(df_alat_g, df_truck_g)
        if img_buf: 
            num_charts = 0
            if not df_alat_g.empty: num_charts += 1
            if not df_truck_g.empty: num_charts += 1
            span_needed = 12 * num_charts
            right_queue.append({'type': 'chart', 'img': img_buf, 'span': span_needed})

        # --- RENDER TABLE WITH SPANNING ---
        ROWS_PER_PAGE = 40
        ROW_HEIGHT = 15
        
        l_ptr = 0
        r_ptr = 0
        right_occupied_until = -1 
        
        while True: # Infinite loop until data exhausted and spans cleared
            page_data = []
            page_style = [('VALIGN', (0,0), (-1,-1), 'MIDDLE')]
            
            row_idx = 0
            while row_idx < ROWS_PER_PAGE:
                if l_ptr >= len(left_queue) and r_ptr >= len(right_queue) and row_idx > right_occupied_until:
                    break
                
                row_content = [''] * 12 
                
                # --- KIRI ---
                if l_ptr < len(left_queue):
                    item = left_queue[l_ptr]
                    itype = item['type']
                    
                    if itype == 'title_section':
                        row_content[0] = Paragraph(item['val'], section_title_style)
                        page_style.append(('SPAN', (0, row_idx), (5, row_idx)))
                    elif itype == 'header_col':
                        cols = ['NO', 'TGL', 'ALAT', 'UNIT', 'LTR', 'KET']
                        for c, txt in enumerate(cols): row_content[c] = Paragraph(txt, header_style)
                        page_style.append(('BACKGROUND', (0, row_idx), (5, row_idx), COLOR_HEADER_BLUE))
                        page_style.append(('GRID', (0, row_idx), (5, row_idx), 0.5, COLOR_BORDER))
                    elif itype == 'row':
                        d = item['data']
                        row_content[0] = d[0]; row_content[1] = d[1]
                        row_content[2] = Paragraph(safe_text(d[2], 25), cell_style)
                        row_content[3] = Paragraph(safe_text(d[3], 15), cell_style)
                        row_content[4] = d[4]
                        row_content[5] = Paragraph(safe_text(d[5], 25), cell_style)
                        bg = COLOR_ROW_EVEN if (l_ptr % 2 == 0) else COLOR_ROW_ODD
                        page_style.append(('BACKGROUND', (0, row_idx), (5, row_idx), bg))
                        page_style.append(('GRID', (0, row_idx), (5, row_idx), 0.5, COLOR_BORDER))
                        page_style.append(('ALIGN', (4, row_idx), (4, row_idx), 'RIGHT'))
                        page_style.append(('ALIGN', (0, row_idx), (1, row_idx), 'CENTER'))
                    elif itype == 'total_left':
                        row_content[0] = ''; row_content[1] = ''
                        row_content[2] = 'TOTAL PENGGUNAAN'
                        row_content[4] = item['val']
                        page_style.append(('SPAN', (2, row_idx), (3, row_idx)))
                        page_style.append(('BACKGROUND', (0, row_idx), (5, row_idx), COLOR_TOTAL_YELLOW))
                        page_style.append(('FONTNAME', (0, row_idx), (5, row_idx), 'Helvetica-Bold'))
                        page_style.append(('GRID', (0, row_idx), (5, row_idx), 0.5, COLOR_BORDER))
                        page_style.append(('ALIGN', (4, row_idx), (4, row_idx), 'RIGHT'))
                    l_ptr += 1

                # --- KANAN ---
                if row_idx <= right_occupied_until:
                    pass
                elif r_ptr < len(right_queue):
                    item = right_queue[r_ptr]
                    itype = item['type']
                    
                    if itype == 'title_section':
                        row_content[7] = Paragraph(item['val'], section_title_style)
                        page_style.append(('SPAN', (7, row_idx), (11, row_idx)))
                    elif itype == 'header_masuk':
                        cols = ['NO', 'TGL', 'SUMBER', 'JNS', 'LTR']
                        for c, txt in enumerate(cols): row_content[7+c] = Paragraph(txt, header_style)
                        page_style.append(('BACKGROUND', (7, row_idx), (11, row_idx), COLOR_HEADER_BLUE))
                        page_style.append(('GRID', (7, row_idx), (11, row_idx), 0.5, COLOR_BORDER))
                    elif itype == 'row_masuk':
                        d = item['data']
                        row_content[7] = d[0]; row_content[8] = d[1]
                        row_content[9] = Paragraph(safe_text(d[2]), cell_style)
                        row_content[10] = d[3]; row_content[11] = d[4]
                        page_style.append(('GRID', (7, row_idx), (11, row_idx), 0.5, COLOR_BORDER))
                        page_style.append(('ALIGN', (11, row_idx), (11, row_idx), 'RIGHT'))
                    elif itype == 'total_masuk':
                        row_content[7] = 'TOTAL MASUK'; row_content[11] = item['val']
                        page_style.append(('SPAN', (7, row_idx), (10, row_idx)))
                        page_style.append(('BACKGROUND', (7, row_idx), (11, row_idx), COLOR_TOTAL_YELLOW))
                        page_style.append(('FONTNAME', (7, row_idx), (11, row_idx), 'Helvetica-Bold'))
                        page_style.append(('GRID', (7, row_idx), (11, row_idx), 0.5, COLOR_BORDER))
                        page_style.append(('ALIGN', (11, row_idx), (11, row_idx), 'RIGHT'))
                    elif itype == 'sub_rekap':
                        style_to_use = header_black_style if item.get('txt_black') else header_style
                        row_content[7] = Paragraph(item['title'], style_to_use)
                        bg = colors.HexColor(item['bg'])
                        page_style.append(('SPAN', (7, row_idx), (11, row_idx)))
                        page_style.append(('BACKGROUND', (7, row_idx), (11, row_idx), bg))
                        page_style.append(('GRID', (7, row_idx), (11, row_idx), 0.5, COLOR_BORDER))
                    elif itype == 'row_rekap':
                        row_content[7] = Paragraph(safe_text(item['label'], 35), cell_style)
                        row_content[11] = item['val']
                        page_style.append(('SPAN', (7, row_idx), (10, row_idx)))
                        page_style.append(('GRID', (7, row_idx), (11, row_idx), 0.5, COLOR_BORDER))
                        page_style.append(('ALIGN', (11, row_idx), (11, row_idx), 'RIGHT'))
                    elif itype == 'total_rekap':
                        row_content[7] = 'TOTAL'; row_content[11] = item['val']
                        page_style.append(('SPAN', (7, row_idx), (10, row_idx)))
                        page_style.append(('BACKGROUND', (7, row_idx), (11, row_idx), COLOR_TOTAL_YELLOW))
                        page_style.append(('FONTNAME', (7, row_idx), (11, row_idx), 'Helvetica-Bold'))
                        page_style.append(('GRID', (7, row_idx), (11, row_idx), 0.5, COLOR_BORDER))
                        page_style.append(('ALIGN', (11, row_idx), (11, row_idx), 'RIGHT'))
                    elif itype == 'header_stok':
                        row_content[7] = Paragraph(item['label'], header_style)
                        page_style.append(('SPAN', (7, row_idx), (11, row_idx)))
                        page_style.append(('BACKGROUND', (7, row_idx), (11, row_idx), colors.HexColor("#70AD47")))
                        page_style.append(('GRID', (7, row_idx), (11, row_idx), 0.5, COLOR_BORDER))
                    elif itype == 'row_stok':
                        row_content[7] = item['label']; row_content[11] = item['val']
                        page_style.append(('SPAN', (7, row_idx), (10, row_idx)))
                        page_style.append(('GRID', (7, row_idx), (11, row_idx), 0.5, COLOR_BORDER))
                        page_style.append(('ALIGN', (11, row_idx), (11, row_idx), 'RIGHT'))
                    elif itype == 'total_stok':
                        row_content[7] = item['label']; row_content[11] = item['val']
                        page_style.append(('SPAN', (7, row_idx), (10, row_idx)))
                        page_style.append(('BACKGROUND', (7, row_idx), (11, row_idx), colors.HexColor("#70AD47")))
                        page_style.append(('FONTNAME', (7, row_idx), (11, row_idx), 'Helvetica-Bold'))
                        page_style.append(('TEXTCOLOR', (7, row_idx), (11, row_idx), colors.white))
                        page_style.append(('GRID', (7, row_idx), (11, row_idx), 0.5, COLOR_BORDER))
                        page_style.append(('ALIGN', (11, row_idx), (11, row_idx), 'RIGHT'))
                    elif itype == 'chart':
                        # LOGIKA SPAN CHART FIX
                        span_needed = item['span']
                        rows_left = ROWS_PER_PAGE - row_idx
                        if rows_left < 5: 
                            pass # Biarkan loop lanjut
                        else:
                            real_span = min(span_needed, rows_left)
                            img_height = real_span * 14
                            
                            row_content[7] = RLImage(item['img'], width=200, height=img_height)
                            
                            span_end_idx = row_idx + real_span - 1
                            page_style.append(('SPAN', (7, row_idx), (11, span_end_idx)))
                            
                            right_occupied_until = span_end_idx
                            r_ptr += 1 

                    r_ptr += 1
                
                page_data.append(row_content)
                row_idx += 1
            
            if not page_data: break 

            col_widths = [20, 30, 80, 40, 30, 80,  20,  20, 30, 80, 40, 50]
            t = Table(page_data, colWidths=col_widths, rowHeights=[ROW_HEIGHT]*len(page_data))
            t.setStyle(TableStyle(page_style))
            elements.append(t)
            
            if l_ptr >= len(left_queue) and r_ptr >= len(right_queue):
                 break
            
            elements.append(PageBreak())

    # --- HALAMAN REKAP TAHUNAN ---
    elements.append(PageBreak())
    elements.append(Paragraph("LAPORAN BBM PERBULAN", title_style))
    m_data = []
    stok_run = hitung_stok_awal_periode(conn, lokasi_id, start_date_global)
    curr = start_date_global.replace(day=1)
    end_limit = end_date_global.replace(day=1)
    while curr <= end_limit:
        m = curr.month; y = curr.year
        q_in = f"SELECT SUM(jumlah_liter) FROM bbm_masuk WHERE lokasi_id={lokasi_id} AND MONTH(tanggal)={m} AND YEAR(tanggal)={y}"
        q_out = f"SELECT SUM(jumlah_liter) FROM bbm_keluar WHERE lokasi_id={lokasi_id} AND MONTH(tanggal)={m} AND YEAR(tanggal)={y}"
        cursor = conn.cursor()
        cursor.execute(q_in); res_in = cursor.fetchone(); mi = float(res_in[0]) if res_in and res_in[0] else 0.0
        cursor.execute(q_out); res_out = cursor.fetchone(); mo = float(res_out[0]) if res_out and res_out[0] else 0.0
        prev = stok_run
        stok_run = prev + mi - mo
        m_data.append({'bln': curr.strftime("%B %Y"), 'awal': prev, 'masuk': mi, 'keluar': mo, 'sisa': stok_run, 'bulan_nama': get_bulan_indonesia(m)[:3]})
        if m == 12: curr = datetime.date(y+1, 1, 1)
        else: curr = datetime.date(y, m+1, 1)
    
    df_m = pd.DataFrame(m_data)
    if not df_m.empty:
        img_m_buf = generate_monthly_chart(df_m)
        if img_m_buf: elements.append(RLImage(img_m_buf, width=480, height=220)); elements.append(Spacer(1, 15))
    
    # TABLE REKAP BULANAN
    d_m = [['BULAN', 'SISA BULAN LALU', 'MASUK', 'KELUAR', 'SISA']]
    for r in m_data: d_m.append([r['bln'], f"{r['awal']:,.0f}", f"{r['masuk']:,.0f}", f"{r['keluar']:,.0f}", f"{r['sisa']:,.0f}"])
    
    # TAMBAHKAN TOTAL ROW (PDF)
    if m_data:
        t_masuk = sum(x['masuk'] for x in m_data)
        t_keluar = sum(x['keluar'] for x in m_data)
        akhir = m_data[-1]['sisa']
        d_m.append(['TOTAL', '', f"{t_masuk:,.0f}", f"{t_keluar:,.0f}", f"{akhir:,.0f}"])

    t_m = Table(d_m, colWidths=[100, 100, 100, 100, 100])
    
    rekap_style = [
        ('GRID', (0,0), (-1,-1), 0.5, COLOR_BORDER),
        ('BACKGROUND', (0,0), (-1,0), COLOR_HEADER_BLUE),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]
    for i in range(1, len(d_m)):
        bg = COLOR_ROW_EVEN if i % 2 == 0 else COLOR_ROW_ODD
        rekap_style.append(('BACKGROUND', (0, i), (-1, i), bg))
    
    # Style row TOTAL (baris terakhir)
    if m_data:
        rekap_style.append(('BACKGROUND', (0, -1), (-1, -1), COLOR_TOTAL_YELLOW))
        rekap_style.append(('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'))

    t_m.setStyle(TableStyle(rekap_style))
    elements.append(t_m)

    doc.build(elements); buffer.seek(0)
    return buffer

# --- EXPORT EXCEL (RESTORED) ---
def generate_excel_styled(conn, lokasi_id, nama_lokasi, start_date_global, end_date_global, excluded_list):
    output = io.BytesIO(); wb = Workbook(); wb.remove(wb.active)
    thin = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
    date_ranges = split_date_range_by_month(start_date_global, end_date_global)
    for idx, (start_date, end_date) in enumerate(date_ranges):
        sheet_name = get_bulan_indonesia(start_date.month)[:3] + f" {start_date.year}"
        ws = wb.create_sheet(sheet_name)
        ws.column_dimensions['A'].width = 5; ws.column_dimensions['B'].width = 15; ws.column_dimensions['C'].width = 30; ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 15; ws.column_dimensions['F'].width = 35
        ws.column_dimensions['I'].width = 5; ws.column_dimensions['J'].width = 15; ws.column_dimensions['K'].width = 30; ws.column_dimensions['L'].width = 15
        
        stok_awal = hitung_stok_awal_periode(conn, lokasi_id, start_date)
        df_masuk = pd.read_sql(f"SELECT * FROM bbm_masuk WHERE lokasi_id={lokasi_id} AND tanggal BETWEEN '{start_date}' AND '{end_date}' ORDER BY tanggal", conn)
        df_keluar = pd.read_sql(f"SELECT * FROM bbm_keluar WHERE lokasi_id={lokasi_id} AND tanggal BETWEEN '{start_date}' AND '{end_date}' ORDER BY tanggal", conn)
        
        if not df_keluar.empty and 'kategori' not in df_keluar.columns: df_keluar['kategori'] = df_keluar['nama_alat'].apply(cek_kategori)
        df_keluar_rpt = filter_non_consumption(df_keluar)
        
        tm = float(df_masuk['jumlah_liter'].sum()) if not df_masuk.empty else 0.0
        tk_real = float(df_keluar['jumlah_liter'].sum()) if not df_keluar.empty else 0.0
        tk_rpt = float(df_keluar_rpt['jumlah_liter'].sum()) if not df_keluar_rpt.empty else 0.0
        sisa_akhir = stok_awal + tm - tk_real
        
        ws.merge_cells('A1:L1'); ws['A1'] = f"LAPORAN BBM: {nama_lokasi}"; ws['A1'].font = Font(bold=True, size=14); ws['A1'].alignment = Alignment(horizontal='center')
        ws.merge_cells('A2:L2'); ws['A2'] = f"PERIODE {get_bulan_indonesia(start_date.month)} {start_date.year}"; ws['A2'].font = Font(size=12); ws['A2'].alignment = Alignment(horizontal='center')
        
        r = 4; ws.cell(r, 1, "A. PENGGUNAAN (KELUAR)").font = Font(bold=True)
        headers = ['NO', 'TGL', 'ALAT', 'UNIT', 'LTR', 'KET']
        for i, h in enumerate(headers): c=ws.cell(r+1, i+1, h); c.border=thin; c.fill=PatternFill("solid", fgColor="D3D3D3"); c.alignment=Alignment(horizontal='center')
        r += 2
        if not df_keluar_rpt.empty:
            last_date = None; is_grey = False
            for i, row in df_keluar_rpt.iterrows():
                curr_date = row['tanggal']
                if last_date is not None and curr_date != last_date: is_grey = not is_grey
                last_date = curr_date
                fill = PatternFill("solid", fgColor="F2F2F2") if is_grey else None
                vals = [i+1, row['tanggal'].strftime('%d/%m/%Y'), row['nama_alat'], row['no_unit'], float(row['jumlah_liter']), row['keterangan']]
                for j, v in enumerate(vals): c=ws.cell(r, j+1, v); c.border=thin; c.alignment=Alignment(wrap_text=True, vertical='center'); 
                if fill: 
                    for j in range(6): ws.cell(r, j+1).fill = fill
                r += 1
        ws.cell(r, 3, "TOTAL").font=Font(bold=True); c=ws.cell(r, 5, tk_rpt); c.font=Font(bold=True); c.fill=PatternFill("solid", fgColor="FFFF00"); c.border=thin
        
        r_r = 4; ws.cell(r_r, 9, "B. BBM MASUK").font = Font(bold=True)
        headers_m = ['NO', 'TGL', 'SUMBER', 'JNS', 'LTR']
        for i, h in enumerate(headers_m): c=ws.cell(r_r+1, i+9, h); c.border=thin; c.fill=PatternFill("solid", fgColor="D3D3D3"); c.alignment=Alignment(horizontal='center')
        r_r += 2
        if not df_masuk.empty:
            for i, row in df_masuk.iterrows():
                vals = [i+1, row['tanggal'].strftime('%d/%m/%Y'), row['sumber'], row['jenis_bbm'], float(row['jumlah_liter'])]
                for j, v in enumerate(vals): c=ws.cell(r_r, j+9, v); c.border=thin; c.alignment=Alignment(wrap_text=True)
                r_r += 1
        ws.cell(r_r, 11, "TOTAL").font=Font(bold=True); c=ws.cell(r_r, 13, tm); c.font=Font(bold=True); c.fill=PatternFill("solid", fgColor="FFFF00"); c.border=thin
        r_r += 2
        
        ws.cell(r_r, 9, "RINCIAN PENGGUNAAN BBM").font=Font(bold=True); r_r+=1
        df_alat_g, df_truck_g, df_lain_g = segregate_data(df_keluar_rpt, excluded_list)
        
        def write_detail(ws, row, col, title, df, color):
            ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+3)
            c=ws.cell(row, col, title); c.fill=color; c.font=Font(bold=True); c.alignment=Alignment(horizontal='center'); c.border=thin
            row+=1
            if not df.empty and 'jumlah_liter' in df.columns:
                grp = df.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
                for _, x in grp.iterrows():
                    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+2)
                    c1=ws.cell(row, col, f"{x['nama_alat']} {x['no_unit']}"); c1.border=thin
                    c2=ws.cell(row, col+3, float(x['jumlah_liter'])); c2.border=thin
                    row+=1
            total_val = float(df['jumlah_liter'].sum()) if not df.empty and 'jumlah_liter' in df.columns else 0
            ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+2)
            c_tot=ws.cell(row, col, "TOTAL"); c_tot.fill=PatternFill("solid", fgColor="FFFF00"); c_tot.border=thin; c_tot.font=Font(bold=True)
            c_val=ws.cell(row, col+3, total_val); c_val.fill=PatternFill("solid", fgColor="FFFF00"); c_val.border=thin; c_val.font=Font(bold=True)
            return row+2
        
        r_r = write_detail(ws, r_r, 9, "TOTAL PENGGUNAAN ALAT BERAT", df_alat_g, PatternFill("solid", fgColor="F4B084"))
        r_r = write_detail(ws, r_r, 9, "TOTAL PENGGUNAAN MOBIL & TRUCK", df_truck_g, PatternFill("solid", fgColor="9BC2E6"))
        if not df_lain_g.empty: r_r = write_detail(ws, r_r, 9, "TOTAL PENGGUNAAN BBM LAINNYA", df_lain_g, PatternFill("solid", fgColor="FFB6C1"))
            
        ws.cell(r_r, 9, "RINCIAN SISA STOK").font=Font(bold=True); r_r+=1
        data_s = [('SISA BULAN LALU', stok_awal), ('MASUK', tm), ('KELUAR (REAL)', tk_real), ('SISA AKHIR', sisa_akhir)]
        for k, v in data_s:
            ws.merge_cells(start_row=r_r, start_column=9, end_row=r_r, end_column=11)
            c1=ws.cell(r_r, 9, k); c1.border=thin
            c2=ws.cell(r_r, 12, v); c2.border=thin
            if k == 'SISA AKHIR': c1.fill=PatternFill("solid", fgColor="00FF00"); c2.fill=PatternFill("solid", fgColor="00FF00")
            r_r+=1
        r_r+=1
        
        img_buf = generate_chart_sidebar(df_alat_g, df_truck_g)
        if img_buf: img = XLImage(img_buf); img.width = 450; img.height = 450; ws.add_image(img, f'I{r_r}')

    ws2 = wb.create_sheet("Rekap Tahunan"); ws2['A1'] = "LAPORAN BBM PERBULAN"; ws2['A1'].font = Font(bold=True, size=14)
    ws2.column_dimensions['A'].width = 25 
    ws2.column_dimensions['B'].width = 20; ws2.column_dimensions['C'].width = 20
    ws2.column_dimensions['D'].width = 20; ws2.column_dimensions['E'].width = 20
    m_data = []
    stok_run = hitung_stok_awal_periode(conn, lokasi_id, start_date_global)
    curr = start_date_global.replace(day=1)
    end_limit = end_date_global.replace(day=1)
    while curr <= end_limit:
        m = curr.month; y = curr.year
        q_in = f"SELECT SUM(jumlah_liter) FROM bbm_masuk WHERE lokasi_id={lokasi_id} AND MONTH(tanggal)={m} AND YEAR(tanggal)={y}"
        q_out = f"SELECT SUM(jumlah_liter) FROM bbm_keluar WHERE lokasi_id={lokasi_id} AND MONTH(tanggal)={m} AND YEAR(tanggal)={y}"
        cursor = conn.cursor()
        cursor.execute(q_in); res_in = cursor.fetchone(); mi = float(res_in[0]) if res_in and res_in[0] else 0.0
        cursor.execute(q_out); res_out = cursor.fetchone(); mo = float(res_out[0]) if res_out and res_out[0] else 0.0
        prev = stok_run
        stok_run = prev + mi - mo
        m_data.append({'bln': curr.strftime("%B %Y"), 'awal': prev, 'masuk': mi, 'keluar': mo, 'sisa': stok_run, 'bulan_nama': get_bulan_indonesia(m)[:3]})
        if m == 12: curr = datetime.date(y+1, 1, 1)
        else: curr = datetime.date(y, m+1, 1)
    df_m = pd.DataFrame(m_data)
    img_m_buf = generate_monthly_chart(df_m)
    if img_m_buf: img2 = XLImage(img_m_buf); img2.width=500; img2.height=250; ws2.add_image(img2, 'A3')
    r2 = 18
    headers = ['BULAN', 'SISA BULAN LALU', 'MASUK', 'KELUAR', 'SISA']
    for i, h in enumerate(headers): c=ws2.cell(r2, i+1, h); c.border=thin; c.fill=PatternFill("solid", fgColor="D3D3D3")
    r2+=1
    for r in m_data:
        vals = [r['bln'], r['awal'], r['masuk'], r['keluar'], r['sisa']]
        for i, v in enumerate(vals): c=ws2.cell(r2, i+1, v); c.border=thin
        r2+=1
    
    # TAMBAHAN TOTAL ROW EXCEL
    if m_data:
        t_masuk = sum(x['masuk'] for x in m_data)
        t_keluar = sum(x['keluar'] for x in m_data)
        akhir = m_data[-1]['sisa']
        
        ws2.cell(r2, 1, "TOTAL").font = Font(bold=True)
        ws2.cell(r2, 3, t_masuk).font = Font(bold=True)
        ws2.cell(r2, 4, t_keluar).font = Font(bold=True)
        ws2.cell(r2, 5, akhir).font = Font(bold=True)
        
        for i in range(1, 6):
            c = ws2.cell(r2, i)
            c.fill = PatternFill("solid", fgColor="FFD966")
            c.border = thin

    wb.save(output); output.seek(0)
    return output

# --- EXPORT WORD (RESTORED) ---
def generate_docx_fixed(conn, lokasi_id, nama_lokasi, start_date_global, end_date_global, excluded_list):
    doc = Document(); 
    for s in doc.sections: s.left_margin=Cm(1); s.right_margin=Cm(1)
    date_ranges = split_date_range_by_month(start_date_global, end_date_global)

    for idx, (start_date, end_date) in enumerate(date_ranges):
        if idx > 0: doc.add_page_break()
        p = doc.add_paragraph(f"LAPORAN BBM: {nama_lokasi}"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.runs[0].bold = True; p.runs[0].font.size = Pt(14)
        p2 = doc.add_paragraph(f"PERIODE {get_bulan_indonesia(start_date.month)} {start_date.year}"); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER; p2.runs[0].font.size = Pt(11)
        doc.add_paragraph()

        layout_table = doc.add_table(rows=1, cols=2); layout_table.autofit = False; layout_table.allow_autofit = False
        layout_table.columns[0].width = Cm(10); layout_table.columns[1].width = Cm(9)
        cell_left = layout_table.cell(0, 0); cell_right = layout_table.cell(0, 1)

        stok_awal = hitung_stok_awal_periode(conn, lokasi_id, start_date)
        df_keluar = pd.read_sql(f"SELECT * FROM bbm_keluar WHERE lokasi_id={lokasi_id} AND tanggal BETWEEN '{start_date}' AND '{end_date}' ORDER BY tanggal", conn)
        if not df_keluar.empty and 'kategori' not in df_keluar.columns: df_keluar['kategori'] = df_keluar['nama_alat'].apply(cek_kategori)
        df_keluar_rpt = filter_non_consumption(df_keluar)
        
        # TABLE KIRI
        cell_left.add_paragraph("A. PENGGUNAAN BBM (KELUAR)", style='Heading 3')
        tbl_k = cell_left.add_table(rows=1, cols=4); tbl_k.style = 'Table Grid'
        h_k = tbl_k.rows[0].cells; h_k[0].text="TGL"; h_k[1].text="ALAT"; h_k[2].text="UNIT"; h_k[3].text="LTR"
        if not df_keluar_rpt.empty:
            last_date = None; is_grey = False
            for i, r in df_keluar_rpt.iterrows():
                curr_date = r['tanggal']
                if last_date is not None and curr_date != last_date: is_grey = not is_grey
                last_date = curr_date
                row = tbl_k.add_row().cells
                if is_grey:
                    for c in row: set_cell_bg(c, "F2F2F2")
                row[0].text = r['tanggal'].strftime('%d/%m'); row[1].text = r['nama_alat']; row[2].text = r['no_unit']; row[3].text = f"{r['jumlah_liter']:.0f}"
                for c in row: c.paragraphs[0].runs[0].font.size = Pt(8)
        
        # TABLE KANAN
        df_masuk = pd.read_sql(f"SELECT * FROM bbm_masuk WHERE lokasi_id={lokasi_id} AND tanggal BETWEEN '{start_date}' AND '{end_date}' ORDER BY tanggal", conn)
        cell_right.add_paragraph("B. BBM MASUK", style='Heading 3')
        tbl_m = cell_right.add_table(rows=1, cols=3); tbl_m.style='Table Grid'
        h_m = tbl_m.rows[0].cells; h_m[0].text="TGL"; h_m[1].text="SUMBER"; h_m[2].text="LTR"
        if not df_masuk.empty:
            for i, r in df_masuk.iterrows():
                row = tbl_m.add_row().cells; row[0].text = r['tanggal'].strftime('%d/%m'); row[1].text = r['sumber']; row[2].text = f"{r['jumlah_liter']:.0f}"
                for c in row: c.paragraphs[0].runs[0].font.size = Pt(8)
        cell_right.add_paragraph("")

        cell_right.add_paragraph("RINCIAN PENGGUNAAN BBM", style='Heading 4')
        df_alat_g, df_truck_g, df_lain_g = segregate_data(df_keluar_rpt, excluded_list)
        
        def add_detailed_docx(container, title, df_subset, color_hex):
            p = container.add_paragraph(title); p.runs[0].font.bold=True; p.runs[0].font.size=Pt(9)
            t = container.add_table(rows=1, cols=2); t.style='Table Grid'
            total_liter = 0
            if not df_subset.empty and 'jumlah_liter' in df_subset.columns:
                total_liter = df_subset['jumlah_liter'].sum()
                grp = df_subset.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index()
                for _, r in grp.iterrows():
                    row = t.add_row().cells; row[0].text = f"{r['nama_alat']} {r['no_unit']}"; row[1].text = f"{r['jumlah_liter']:.0f}"
                    for c in row: c.paragraphs[0].runs[0].font.size = Pt(8)
            else:
                t.add_row().cells[0].text = "KOSONG"
            rt = t.add_row().cells; rt[0].text="TOTAL"; rt[1].text=f"{total_liter:.0f}"
            for c in rt: set_cell_bg(c, "FFFF00"); c.paragraphs[0].runs[0].bold=True

        add_detailed_docx(cell_right, "TOTAL PENGGUNAAN BBM ALAT BERAT", df_alat_g, "F4B084"); cell_right.add_paragraph("")
        add_detailed_docx(cell_right, "TOTAL PENGGUNAAN BBM MOBIL & TRUCK", df_truck_g, "9BC2E6"); cell_right.add_paragraph("")
        if not df_lain_g.empty: add_detailed_docx(cell_right, "TOTAL PENGGUNAAN BBM LAINNYA", df_lain_g, "FFB6C1"); cell_right.add_paragraph("")

        tm = float(df_masuk['jumlah_liter'].sum()) if not df_masuk.empty else 0.0
        tk_real = float(df_keluar['jumlah_liter'].sum()) if not df_keluar.empty else 0.0
        sisa = stok_awal + tm - tk_real
        
        cell_right.add_paragraph("RINCIAN SISA STOK BBM", style='Heading 4')
        tbl_s = cell_right.add_table(rows=4, cols=2); tbl_s.style='Table Grid'
        tbl_s.cell(0,0).text="SISA BULAN LALU"; tbl_s.cell(0,1).text=f"{stok_awal:.0f}"
        tbl_s.cell(1,0).text="TOTAL MASUK"; tbl_s.cell(1,1).text=f"{tm:.0f}"
        tbl_s.cell(2,0).text="TOTAL KELUAR (REAL)"; tbl_s.cell(2,1).text=f"{tk_real:.0f}"
        tbl_s.cell(3,0).text="SISA AKHIR"; tbl_s.cell(3,1).text=f"{sisa:.0f}"
        cell_right.add_paragraph("")

        img_buf = generate_chart_sidebar(df_alat_g, df_truck_g)
        if img_buf: cell_right.add_paragraph().add_run().add_picture(img_buf, width=Cm(7))

    doc.add_page_break()
    p_title = doc.add_paragraph("LAPORAN BBM PERBULAN"); p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER; p_title.runs[0].bold=True; p_title.runs[0].font.size=Pt(14)
    m_data = []
    stok_run = hitung_stok_awal_periode(conn, lokasi_id, start_date_global)
    curr = start_date_global.replace(day=1)
    end_limit = end_date_global.replace(day=1)
    while curr <= end_limit:
        m = curr.month; y = curr.year
        q_in = f"SELECT SUM(jumlah_liter) FROM bbm_masuk WHERE lokasi_id={lokasi_id} AND MONTH(tanggal)={m} AND YEAR(tanggal)={y}"
        q_out = f"SELECT SUM(jumlah_liter) FROM bbm_keluar WHERE lokasi_id={lokasi_id} AND MONTH(tanggal)={m} AND YEAR(tanggal)={y}"
        cursor = conn.cursor()
        cursor.execute(q_in); res_in = cursor.fetchone(); mi = float(res_in[0]) if res_in and res_in[0] else 0.0
        cursor.execute(q_out); res_out = cursor.fetchone(); mo = float(res_out[0]) if res_out and res_out[0] else 0.0
        prev = stok_run
        stok_run = prev + mi - mo
        m_data.append({'bln': curr.strftime("%B %Y"), 'awal': prev, 'masuk': mi, 'keluar': mo, 'sisa': stok_run, 'bulan_nama': get_bulan_indonesia(m)[:3]})
        if m == 12: curr = datetime.date(y+1, 1, 1)
        else: curr = datetime.date(y, m+1, 1)

    df_m = pd.DataFrame(m_data)
    if not df_m.empty:
        img_m_buf = generate_monthly_chart(df_m)
        if img_m_buf: doc.add_paragraph().add_run().add_picture(img_m_buf, width=Cm(16))

    doc.add_paragraph("RINCIAN MASUK DAN PENGGUNAAN SOLAR PERBULANNYA", style='Heading 4')
    tbl_month = doc.add_table(rows=1, cols=5); tbl_month.style='Table Grid'
    h_month = tbl_month.rows[0].cells
    for i, t in enumerate(['BULAN', 'SISA BULAN LALU', 'MASUK', 'KELUAR', 'SISA']): h_month[i].text=t
    for r in m_data:
        row = tbl_month.add_row().cells; row[0].text=r['bln']; row[1].text=f"{r['awal']:.0f}"; row[2].text=f"{r['masuk']:.0f}"; row[3].text=f"{r['keluar']:.0f}"; row[4].text=f"{r['sisa']:.0f}"

    # TAMBAHAN TOTAL ROW DOCX
    if m_data:
        t_masuk = sum(x['masuk'] for x in m_data)
        t_keluar = sum(x['keluar'] for x in m_data)
        akhir = m_data[-1]['sisa']
        
        row = tbl_month.add_row().cells
        row[0].text = "TOTAL"
        row[2].text = f"{t_masuk:,.0f}"
        row[3].text = f"{t_keluar:,.0f}"
        row[4].text = f"{akhir:,.0f}"
        
        for c in row:
            set_cell_bg(c, "FFD966")
            if len(c.paragraphs) > 0 and len(c.paragraphs[0].runs) > 0:
                c.paragraphs[0].runs[0].font.bold = True
            elif len(c.paragraphs) > 0:
                 c.paragraphs[0].add_run(c.text).font.bold = True

    buffer = io.BytesIO(); doc.save(buffer); buffer.seek(0)
    return buffer

def main():
    try: 
        conn = init_connection(); cursor = conn.cursor(buffered=True) 
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

    password_rahasia = "123" 
    if "logged_in" not in st.session_state: st.session_state.logged_in = False
    if "active_project_id" not in st.session_state: st.session_state.active_project_id = None
    if "active_project_name" not in st.session_state: st.session_state.active_project_name = None

    if not st.session_state.logged_in:
        st.markdown("<h1 style='text-align: center;'>🔐 Login Admin</h1>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns([1,2,1])
        with c2:
            pwd = st.text_input("Masukkan Password Admin:", type="password")
            if st.button("Login", use_container_width=True):
                if pwd == password_rahasia: st.session_state.logged_in = True; st.rerun()
                else: st.error("Password Salah!")
        st.stop() 

    if st.session_state.active_project_id is None:
        st.title("🗂️ Menu Utama")
        st.write("Selamat Datang, Admin. Silakan pilih lokasi proyek atau buat baru.")
        st.divider()
        col_left, col_right = st.columns(2, gap="large")
        with col_left:
            with st.container(border=True):
                st.subheader("📂 Masuk ke Lokasi Proyek")
                try: df_lokasi = pd.read_sql("SELECT * FROM lokasi_proyek", conn)
                except: df_lokasi = pd.DataFrame()
                if not df_lokasi.empty:
                    pilih_nama = st.selectbox("Pilih Lokasi:", df_lokasi['nama_tempat'])
                    input_pass_lokasi = st.text_input("Password Lokasi:", type="password", key="pass_enter")
                    if st.button("Masuk Lokasi", type="primary", use_container_width=True):
                        data_lok = df_lokasi[df_lokasi['nama_tempat'] == pilih_nama].iloc[0]
                        if input_pass_lokasi == data_lok['kunci_lokasi']:
                            st.session_state.active_project_id = int(data_lok['id']); st.session_state.active_project_name = data_lok['nama_tempat']; st.success(f"Berhasil masuk ke {pilih_nama}"); st.rerun()
                        else: st.error("Password Lokasi Salah!")
                else: st.info("Belum ada data lokasi.")
        with col_right:
            with st.container(border=True):
                st.subheader("➕ Buat Lokasi Baru")
                new_name = st.text_input("Nama Lokasi Baru")
                new_pass = st.text_input("Buat Password Lokasi", type="password", key="pass_create")
                st.warning("⚠️ **Password nya diingat baik-baik !**")
                if st.button("Simpan Lokasi Baru", use_container_width=True):
                    if new_name and new_pass:
                        try:
                            cursor.execute("INSERT INTO lokasi_proyek (nama_tempat, kunci_lokasi) VALUES (%s, %s)", (new_name, new_pass)); conn.commit(); st.success("Lokasi berhasil dibuat!"); st.rerun()
                        except Exception as e: st.error(f"Gagal membuat lokasi: {e}")
                    else: st.error("Nama dan Password wajib diisi!")
        st.stop() 

    lokasi_id = st.session_state.active_project_id; nama_proyek = st.session_state.active_project_name
    cursor.execute("SELECT stok_awal FROM lokasi_proyek WHERE id=%s", (lokasi_id,)); stok_awal_modal = cursor.fetchone()[0]
    
    with st.sidebar:
        st.header(f"📍 {nama_proyek}")
        if st.button("⬅️ Kembali ke Menu Utama", use_container_width=True): st.session_state.active_project_id = None; st.session_state.active_project_name = None; st.rerun()
        st.divider(); st.info("ℹ️ Stok Awal dihitung otomatis berdasarkan history transaksi dan filter tanggal.")

    df_masuk_all = pd.read_sql(f"SELECT * FROM bbm_masuk WHERE lokasi_id={lokasi_id}", conn)
    df_keluar_all = pd.read_sql(f"SELECT * FROM bbm_keluar WHERE lokasi_id={lokasi_id}", conn)
    df_log = pd.read_sql(f"SELECT * FROM log_aktivitas WHERE lokasi_id={lokasi_id}", conn)
    df_ex = pd.read_sql(f"SELECT nama_unit_full FROM rekap_exclude WHERE lokasi_id={lokasi_id}", conn)
    excluded_list = df_ex['nama_unit_full'].tolist() if not df_ex.empty else []

    if not df_masuk_all.empty: df_masuk_all['HARI'] = pd.to_datetime(df_masuk_all['tanggal']).apply(get_hari_indonesia)
    if not df_keluar_all.empty: df_keluar_all['HARI'] = pd.to_datetime(df_keluar_all['tanggal']).apply(get_hari_indonesia)

    st.title(f"🚜 Dashboard: {nama_proyek}")
    t1, t2, t3 = st.tabs(["📝 Input & History", "📊 Laporan & Grafik", "🖨️ Export Dokumen"])
    
    with t1:
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
                        cursor.execute("INSERT INTO bbm_masuk (lokasi_id, tanggal, sumber, jenis_bbm, jumlah_liter, keterangan) VALUES (%s,%s,%s,%s,%s,%s)", (lokasi_id, tg, sm, jn, jl, kt)); conn.commit(); st.success("Data Masuk Tersimpan!"); st.rerun()
            elif mode_transaksi == "📤 PENGGUNAAN BBM":
                with st.form("form_keluar"):
                    c1, c2 = st.columns(2)
                    with c1: tg_p = st.date_input("Tanggal Pakai"); al = st.text_input("Nama Alat/Kendaraan"); un = st.text_input("Kode Unit (Ex: DT-01)")
                    with c2: jl_p = st.number_input("Liter Digunakan", 0.0); kt_p = st.text_area("Keterangan / Lokasi Kerja")
                    if st.form_submit_button("Simpan Penggunaan"):
                        cursor.execute("SELECT id FROM bbm_keluar WHERE lokasi_id=%s AND tanggal=%s AND nama_alat=%s AND no_unit=%s AND jumlah_liter=%s", (lokasi_id, tg_p, al, un, jl_p))
                        if cursor.fetchall(): st.warning("⚠️ Data serupa sudah ada!") 
                        cursor.execute("INSERT INTO bbm_keluar (lokasi_id, tanggal, nama_alat, no_unit, jumlah_liter, keterangan) VALUES (%s,%s,%s,%s,%s,%s)", (lokasi_id, tg_p, al, un, jl_p, kt_p)); conn.commit(); st.success("Data Penggunaan Tersimpan!"); st.rerun()
            elif mode_transaksi == "🔄 PINJAM / TRANSFER ANTAR UNIT":
                st.info("ℹ️ Mode ini memindahkan liter dari satu unit ke unit lain.")
                with st.form("form_transfer"):
                    c1, c2 = st.columns(2)
                    with c1: tgl_tf = st.date_input("Tanggal Transfer"); donor_alat = st.text_input("DARI ALAT (Pemberi/Donor)"); donor_unit = st.text_input("No Unit Donor")
                    with c2: liter_tf = st.number_input("Jumlah Liter Dipinjam", min_value=0.0); recv_alat = st.text_input("KE ALAT (Penerima)"); recv_unit = st.text_input("No Unit Penerima"); ket_tf = st.text_area("Keterangan Tambahan")
                    if st.form_submit_button("Proses Transfer"):
                        if liter_tf > 0 and donor_alat and recv_alat:
                            ket_donor = f"Transfer ke {recv_alat} {recv_unit}. {ket_tf}"; cursor.execute("INSERT INTO bbm_keluar (lokasi_id, tanggal, nama_alat, no_unit, jumlah_liter, keterangan) VALUES (%s,%s,%s,%s,%s,%s)", (lokasi_id, tgl_tf, donor_alat, donor_unit, -liter_tf, ket_donor))
                            ket_recv = f"Pinjam dari {donor_alat} {donor_unit}. {ket_tf}"; cursor.execute("INSERT INTO bbm_keluar (lokasi_id, tanggal, nama_alat, no_unit, jumlah_liter, keterangan) VALUES (%s,%s,%s,%s,%s,%s)", (lokasi_id, tgl_tf, recv_alat, recv_unit, liter_tf, ket_recv)); conn.commit(); st.success(f"Berhasil transfer {liter_tf}L"); st.rerun()
                        else: st.error("Mohon lengkapi nama alat dan jumlah liter!")
        st.divider()
        with st.expander("⏳ RIWAYAT INPUT & UNDO (Filter & Sortir)", expanded=True):
            c_f1, c_f2, c_f3 = st.columns(3)
            filter_tipe = c_f1.multiselect("Filter Jenis:", ["MASUK", "PAKAI", "TRANSFER", "KOREKSI"], default=["MASUK", "PAKAI", "TRANSFER", "KOREKSI"])
            filter_sort = c_f2.selectbox("Urutkan:", ["Waktu Input Terbaru (ID)", "Waktu Input Terlama (ID)", "Tanggal Laporan Terbaru", "Tanggal Laporan Terlama"])
            use_date_filter = c_f3.checkbox("Filter Tanggal Tertentu")
            date_val = c_f3.date_input("Pilih Tanggal", value=datetime.date.today(), disabled=not use_date_filter)
            history_data = []
            if not df_masuk_all.empty:
                temp_m = df_masuk_all.copy(); temp_m['Tipe'] = 'MASUK'; temp_m['Kategori_Filter'] = 'MASUK'; temp_m['Detail'] = temp_m['sumber'] + " (" + temp_m['jenis_bbm'] + ")"; temp_m['Label_History'] = "📥 BBM MASUK (Beli)"
                history_data.append(temp_m[['id', 'tanggal', 'Tipe', 'Detail', 'jumlah_liter', 'Label_History', 'keterangan', 'Kategori_Filter']])
            if not df_keluar_all.empty:
                temp_k = df_keluar_all.copy(); temp_k['Tipe'] = 'KELUAR'; temp_k['Detail'] = temp_k['nama_alat'] + " " + temp_k['no_unit']
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
        st.markdown("### 🗓️ Filter Periode Laporan")
        c_p1, c_p2 = st.columns(2)
        today = datetime.date.today(); first_day = today.replace(day=1)
        if 't2_start' not in st.session_state: st.session_state.t2_start = first_day
        if 't2_end' not in st.session_state: st.session_state.t2_end = today
        with c_p1: start_rep = st.date_input("Mulai Tanggal", value=st.session_state.t2_start, key="pick_start_t2"); st.session_state.t2_start = start_rep
        with c_p2: end_rep = st.date_input("Sampai Tanggal", value=st.session_state.t2_end, key="pick_end_t2"); st.session_state.t2_end = end_rep
        
        df_masuk_rep = df_masuk_all[(pd.to_datetime(df_masuk_all['tanggal']).dt.date >= start_rep) & (pd.to_datetime(df_masuk_all['tanggal']).dt.date <= end_rep)]
        df_keluar_rep = df_keluar_all[(pd.to_datetime(df_keluar_all['tanggal']).dt.date >= start_rep) & (pd.to_datetime(df_keluar_all['tanggal']).dt.date <= end_rep)]
        stok_awal_periode_val = hitung_stok_awal_periode(conn, lokasi_id, start_rep)
        tm_rep = float(df_masuk_rep['jumlah_liter'].sum()); tk_rep = float(df_keluar_rep['jumlah_liter'].sum())
        sisa_rep = stok_awal_periode_val + tm_rep - tk_rep

        st.markdown(f"""<div style="background-color:#d4edda;padding:15px;border-radius:10px;border:1px solid #c3e6cb;text-align:center;margin-bottom:20px;margin-top:10px;"><h2 style="color:#155724;margin:0;">💰 SISA STOK PERIODE INI: {sisa_rep:,.2f} Liter</h2><span style="color:#155724;font-weight:bold;">(Sisa Bulan Lalu: {stok_awal_periode_val:,.0f} + Masuk: {tm_rep:,.0f} - Keluar: {tk_rep:,.0f})</span></div>""", unsafe_allow_html=True)
        
        with st.expander("⚙️ ATUR REKAP (Sembunyikan Unit ke 'Lainnya')", expanded=False):
            st.write("Pilih Unit yang ingin digabung menjadi **'Lainnya'** di tabel Rekapitulasi.")
            if not df_keluar_all.empty:
                df_keluar_all['full_name'] = df_keluar_all['nama_alat'].astype(str) + " " + df_keluar_all['no_unit'].astype(str)
                unique_units = sorted(df_keluar_all['full_name'].unique().tolist())
                selected_excludes = st.multiselect("Pilih Unit:", unique_units, default=[u for u in unique_units if u in excluded_list])
                if st.button("Simpan Pengaturan Rekap"):
                    cursor.execute(f"DELETE FROM rekap_exclude WHERE lokasi_id={lokasi_id}")
                    for item in selected_excludes: cursor.execute("INSERT INTO rekap_exclude (lokasi_id, nama_unit_full) VALUES (%s, %s)", (lokasi_id, item))
                    conn.commit(); st.success("Pengaturan Disimpan!"); st.rerun()
            else: st.info("Belum ada data unit keluar.")
        
        with st.expander("🛠️ MENU ADMIN (Koreksi Nama/Unit)", expanded=False):
            t_kor, t_hap = st.tabs(["Koreksi Nama", "Hapus Data (Backup)"])
            with t_kor:
                c1, c2 = st.columns(2)
                with c1:
                    list_alat = sorted(df_keluar_all['nama_alat'].unique().tolist()) if not df_keluar_all.empty else []
                    if list_alat:
                        pilih_lama = st.selectbox("Alat Salah:", list_alat, key="ot"); input_baru = st.text_input("Nama Benar:", key="nt")
                        if st.button("Ganti Nama Alat"): 
                            cursor.execute(f"SELECT id FROM bbm_keluar WHERE nama_alat='{pilih_lama}' AND lokasi_id={lokasi_id}"); ids = [str(r[0]) for r in cursor.fetchall()]; ids_str = ",".join(ids)
                            if ids: cursor.execute("UPDATE bbm_keluar SET nama_alat=%s WHERE nama_alat=%s AND lokasi_id=%s", (input_baru, pilih_lama, lokasi_id)); cursor.execute("INSERT INTO log_aktivitas (lokasi_id, kategori, deskripsi, affected_ids) VALUES (%s, %s, %s, %s)", (lokasi_id, "GANTI NAMA ALAT", f"Mengubah '{pilih_lama}' menjadi '{input_baru}'", ids_str)); conn.commit(); st.success("Nama Diganti & Dicatat!"); st.rerun()
                
                with c2:
                    list_alat_for_unit = sorted(df_keluar_all['nama_alat'].unique().tolist()) if not df_keluar_all.empty else []
                    if list_alat_for_unit:
                        pilih_alat_u = st.selectbox("Pilih Alat utk Ganti Unit:", list_alat_for_unit, key="pilih_alat_unit")
                        units_of_alat = sorted(df_keluar_all[df_keluar_all['nama_alat'] == pilih_alat_u]['no_unit'].unique().tolist())
                        pl_u = st.selectbox("Unit Salah:", units_of_alat, key="ou"); ib_u = st.text_input("Unit Benar:", key="nu")
                        if st.button("Ganti No Unit"):
                            cursor.execute(f"SELECT id FROM bbm_keluar WHERE no_unit='{pl_u}' AND nama_alat='{pilih_alat_u}' AND lokasi_id={lokasi_id}")
                            ids = [str(r[0]) for r in cursor.fetchall()]; ids_str = ",".join(ids)
                            if ids: 
                                cursor.execute("UPDATE bbm_keluar SET no_unit=%s WHERE no_unit=%s AND nama_alat=%s AND lokasi_id=%s", (ib_u, pl_u, pilih_alat_u, lokasi_id))
                                cursor.execute("INSERT INTO log_aktivitas (lokasi_id, kategori, deskripsi, affected_ids) VALUES (%s, %s, %s, %s)", (lokasi_id, "GANTI NO UNIT", f"Mengubah '{pl_u}' menjadi '{ib_u}' pada alat '{pilih_alat_u}'", ids_str)); conn.commit(); st.success("Unit Diganti & Dicatat!"); st.rerun()
                            else: st.warning("Data tidak ditemukan untuk kombinasi Alat dan Unit tersebut.")

            with t_hap:
                c1, c2 = st.columns(2)
                with c1:
                    if not df_masuk_all.empty:
                        m_sel = st.selectbox("Hapus Masuk:", df_masuk_all.apply(lambda x: f"{x['id']}|{x['tanggal']}|{x['sumber']}", axis=1))
                        if st.button("Hapus Masuk"): cursor.execute(f"DELETE FROM bbm_masuk WHERE id={m_sel.split('|')[0]}"); conn.commit(); st.rerun()
                with c2:
                    if not df_keluar_all.empty:
                        k_sel = st.selectbox("Hapus Keluar:", df_keluar_all.apply(lambda x: f"{x['id']}|{x['tanggal']}|{x['nama_alat']}", axis=1))
                        if st.button("Hapus Keluar"): cursor.execute(f"DELETE FROM bbm_keluar WHERE id={k_sel.split('|')[0]}"); conn.commit(); st.rerun()

        df_alat_g = pd.DataFrame(); df_truck_g = pd.DataFrame(); df_lain_g = pd.DataFrame()
        if not df_keluar_rep.empty:
            if 'kategori' not in df_keluar_rep.columns: df_keluar_rep['kategori'] = df_keluar_rep['nama_alat'].apply(cek_kategori)
            df_alat_g, df_truck_g, df_lain_g = segregate_data(df_keluar_rep, excluded_list)
            
            c_rekap1, c_rekap2, c_rekap3 = st.columns(3)
            with c_rekap1: 
                st.write("**Rekap Alat Berat**")
                if not df_alat_g.empty: st.dataframe(df_alat_g.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index(), hide_index=True)
            with c_rekap2: 
                st.write("**Rekap Mobil/Truck**")
                if not df_truck_g.empty: st.dataframe(df_truck_g.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index(), hide_index=True)
            with c_rekap3: 
                st.write("**Rekap Lainnya (Excluded)**")
                if not df_lain_g.empty: st.dataframe(df_lain_g.groupby(['nama_alat', 'no_unit'])['jumlah_liter'].sum().reset_index(), hide_index=True)

        st.divider(); col_a, col_b = st.columns(2)
        with col_a: st.subheader("📋 Data Masuk BBM (Periode Ini)"); st.dataframe(df_masuk_rep.sort_values('tanggal', ascending=False), use_container_width=True)
        with col_b: st.subheader("📋 Data Penggunaan BBM (Periode Ini)"); st.dataframe(df_keluar_rep.sort_values('tanggal', ascending=False), use_container_width=True)
        st.divider(); st.subheader("📊 Monitoring (Periode Ini)"); 
        img_buffer = generate_chart_sidebar(df_alat_g, df_truck_g)
        if img_buffer: st.image(img_buffer, caption="Diagram Alat Berat & Mobil")
        else: st.info("Belum ada data untuk ditampilkan di grafik.")

        st.divider(); st.subheader(f"📅 Monitoring Bulanan ({start_rep.strftime('%b %Y')} - {end_rep.strftime('%b %Y')})")
        stok_run_m = hitung_stok_awal_periode(conn, lokasi_id, start_rep)
        m_data_mon = []
        curr = start_rep.replace(day=1); end_limit = end_rep.replace(day=1)
        while curr <= end_limit:
            m = curr.month; y = curr.year
            q_in = f"SELECT SUM(jumlah_liter) FROM bbm_masuk WHERE lokasi_id={lokasi_id} AND MONTH(tanggal)={m} AND YEAR(tanggal)={y}"
            q_out = f"SELECT SUM(jumlah_liter) FROM bbm_keluar WHERE lokasi_id={lokasi_id} AND MONTH(tanggal)={m} AND YEAR(tanggal)={y}"
            cursor = conn.cursor()
            cursor.execute(q_in); res_in = cursor.fetchone(); mi = float(res_in[0]) if res_in and res_in[0] else 0.0
            cursor.execute(q_out); res_out = cursor.fetchone(); mo = float(res_out[0]) if res_out and res_out[0] else 0.0
            prev = stok_run_m; stok_run_m = prev + mi - mo
            m_data_mon.append({'Bulan': curr.strftime("%B %Y"), 'Sisa Bulan Lalu': prev, 'Masuk': mi, 'Keluar': mo, 'Sisa Akhir': stok_run_m, 'bulan_nama': curr.strftime("%b")})
            if m == 12: curr = datetime.date(y+1, 1, 1)
            else: curr = datetime.date(y, m+1, 1)

        df_m_mon = pd.DataFrame(m_data_mon)
        c_g1, c_g2 = st.columns(2)
        with c_g1: 
            if not df_m_mon.empty:
                df_for_graph = df_m_mon.rename(columns={'Masuk': 'masuk', 'Keluar': 'keluar'})
                img_m = generate_monthly_chart(df_for_graph)
                if img_m: st.image(img_m)
            else: st.info("Tidak ada data dalam range ini.")
        with c_g2: st.dataframe(df_m_mon[['Bulan', 'Sisa Bulan Lalu', 'Masuk', 'Keluar', 'Sisa Akhir']], hide_index=True)

    with t3:
        st.header("🖨️ Export Laporan Periode")
        st.write("Silakan pilih periode laporan. Sistem akan membuat laporan **Bulan demi Bulan** secara otomatis dalam file export.")
        if 't3_start' not in st.session_state: st.session_state.t3_start = first_day
        if 't3_end' not in st.session_state: st.session_state.t3_end = today
        c_d1, c_d2 = st.columns(2)
        with c_d1: start_date_exp = st.date_input("Dari Tanggal", value=st.session_state.t3_start, key="pick_start_t3"); st.session_state.t3_start = start_date_exp
        with c_d2: end_date_exp = st.date_input("Sampai Tanggal", value=st.session_state.t3_end, key="pick_end_t3"); st.session_state.t3_end = end_date_exp

        c1, c2, c3 = st.columns(3)
        if start_date_exp <= end_date_exp:
            with c1: 
                if st.button("📕 Download PDF", use_container_width=True): 
                    pdf = generate_pdf_portrait(conn, lokasi_id, nama_proyek, start_date_exp, end_date_exp, excluded_list)
                    st.download_button("⬇️ Simpan PDF", pdf, f"Laporan_{nama_proyek}_{start_date_exp}_{end_date_exp}.pdf", "application/pdf")
            with c2: 
                if st.button("📗 Download Excel", use_container_width=True): 
                    xl = generate_excel_styled(conn, lokasi_id, nama_proyek, start_date_exp, end_date_exp, excluded_list)
                    st.download_button("⬇️ Simpan Excel", xl, f"Laporan_{nama_proyek}_{start_date_exp}_{end_date_exp}.xlsx")
            with c3: 
                if st.button("📘 Download Word", use_container_width=True): 
                    doc = generate_docx_fixed(conn, lokasi_id, nama_proyek, start_date_exp, end_date_exp, excluded_list)
                    st.download_button("⬇️ Simpan Word", doc, f"Laporan_{nama_proyek}_{start_date_exp}_{end_date_exp}.docx")
        else: st.error("Tanggal Akhir harus lebih besar dari Tanggal Awal")

if __name__ == "__main__":
    main()