import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

def register_fonts():
    # Prefer a bundled font in ./fonts/ (add a font file there, e.g. NotoSansTC-Regular.otf)
    bundled_paths = [
        os.path.join(os.path.dirname(__file__), "fonts", "NotoSansTC-Regular.otf"),
        os.path.join(os.path.dirname(__file__), "fonts", "NotoSansTC-Regular.ttf"),
        os.path.join(os.path.dirname(__file__), "fonts", "msyh.ttc"),  # optionally bundled MSYH
    ]
    for p in bundled_paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont('APP_FONT', p))
                return 'APP_FONT'
            except Exception:
                pass
    # Try system font paths (best-effort) - Windows example
    try:
        pdfmetrics.registerFont(TTFont('MSYH', 'C:/Windows/Fonts/msyh.ttc', subfontIndex=0))
        return 'MSYH'
    except Exception:
        pass
    # Fall back to ReportLab's Helvetica
    return 'Helvetica'

FONT_NAME = register_fonts()

def create_pdf(df, selected_districts):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=1.2*cm,leftMargin=1.2*cm,
                            topMargin=1.5*cm,bottomMargin=2*cm)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title',parent=styles['Heading1'],
                                 fontName=FONT_NAME,fontSize=16,
                                 alignment=1,textColor=colors.HexColor('#13343b'),
                                 spaceAfter=15)
    title_text = "已持有牌照的私營安老院名單 - "+(', '.join(selected_districts) if selected_districts else "全部地區")
    elements.append(Paragraph(title_text, title_style))

    info_style = ParagraphStyle('Info', parent=styles['Normal'],
                                fontName=FONT_NAME,fontSize=10,
                                alignment=1,textColor=colors.HexColor('#626c71'))
    info_text = f"總共有 {len(df)} 間院舍"
    elements.append(Paragraph(info_text, info_style))
    elements.append(Spacer(1,10))

    header_style = ParagraphStyle('Header',parent=styles['Normal'],fontName=FONT_NAME,
                                  fontSize=12,textColor=colors.whitesmoke,alignment=1)
    body_style = ParagraphStyle('Body',parent=styles['Normal'],fontName=FONT_NAME,
                                fontSize=10,leading=12,wordWrap='CJK')

    table_data = [
        [Paragraph("序號", header_style), Paragraph("地區", header_style), Paragraph("院舍名稱", header_style),
         Paragraph("地址", header_style), Paragraph("電話", header_style)]
    ]

    phone_col_candidates = ['電話/\nTelephone No.', '電話／\nTelephone No.']
    phone_key = next((c for c in phone_col_candidates if c in df.columns), None)

    for idx, row in df.iterrows():
        seq = Paragraph(str(len(table_data)), body_style)
        district = Paragraph(str(row.get('地區', '')), body_style)
        home_name = Paragraph(str(row.get('Unnamed: 4', '')), body_style)
        home_addr = Paragraph(str(row.get('Unnamed: 6', '')), body_style)
        phone_raw = str(row.get(phone_key, '')) if phone_key else ''
        phone = '' if phone_raw == 'nan' else phone_raw.split('.')[0] if phone_raw else ''
        phone_par = Paragraph(phone, body_style)
        table_data.append([seq, district, home_name, home_addr, phone_par])

    col_widths = [1 * cm, 2.5 * cm, 6 * cm, 8 * cm, 3 * cm]
    table = Table(table_data, colWidths=col_widths)
    border_width = 0.0283
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.black),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("LINEBEFORE", (0, 0), (-1, -1), border_width, colors.black),
        ("LINEABOVE", (0, 0), (-1, -1), border_width, colors.black),
        ("LINEAFTER", (0, 0), (-1, -1), border_width, colors.black),
        ("LINEBELOW", (0, 0), (-1, -1), border_width, colors.black),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("ALIGN", (0, 1), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
    ]))
    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer

st.title("[translate:院舍PDF導出工具]")

uploaded_file = st.file_uploader("[translate:請上傳Excel檔案]", type=["xls", "xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, skiprows=6, engine='openpyxl')
    df.columns = df.columns.str.strip()
    df = df[df['地區'].notna()]
    districts = sorted(df['地區'].unique())
    selected_districts = st.multiselect("[translate:請選擇地區]", districts, default=districts)

    filtered_df = df[df['地區'].isin(selected_districts)]

    if st.button("[translate:產生PDF]"):
        pdf_bytes = create_pdf(filtered_df, selected_districts)
        st.download_button("[translate:下載PDF]", data=pdf_bytes, file_name="院舍名單.pdf", mime="application/pdf")
