import io
import os
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

def generate_em_tables_page_pdf(context_data):
    """
    Generates a 1-page PDF buffer containing Table 1 and Table 2 using ReportLab Platypus.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=letter,
        leftMargin=30,
        rightMargin=30,
        topMargin=30,
        bottomMargin=30
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'TableTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=10,
        textColor=colors.black,
        spaceAfter=4
    )
    
    cell_hdr_style = ParagraphStyle(
        'HeaderCell',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=6.5,
        leading=8,
        alignment=TA_CENTER,
        textColor=colors.black
    )
    
    cell_body_style = ParagraphStyle(
        'BodyCell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=6,
        leading=7.5,
        alignment=TA_CENTER,
        textColor=colors.black
    )

    cell_body_left = ParagraphStyle(
        'BodyCellLeft',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=6,
        leading=7.5,
        alignment=TA_LEFT,
        textColor=colors.black
    )

    cell_section_hdr = ParagraphStyle(
        'SectionHdr',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=6.5,
        leading=8,
        alignment=TA_LEFT,
        textColor=colors.black
    )

    elements = []

    # --- TABLE 1 ---
    elements.append(Paragraph("Table 1: Read Dates & Incubation Observation", title_style))
    
    t1_headers = [
        Paragraph("ETX Submission<br/>ID", cell_hdr_style),
        Paragraph("Set-up Analyst<br/>& Date", cell_hdr_style),
        Paragraph("Sampling Site &<br/>Location", cell_hdr_style),
        Paragraph("Plate Reading<br/>Analyst (≥ 48H)", cell_hdr_style),
        Paragraph("CFUs Observed after 48 Hour Incubation at 30–35°C (E001031)", cell_hdr_style),
        Paragraph("Plate Reading<br/>Analyst (NLT 5 days)", cell_hdr_style),
        Paragraph("CFUs Observed after NLT 5-day Incubation at 20–25°C (E001034)", cell_hdr_style),
        Paragraph("Microbial Identification", cell_hdr_style)
    ]

    t1_row_vals = [
        Paragraph(context_data.get('etx_id', 'ETX-260216-0348'), cell_body_style),
        Paragraph(f"{context_data.get('setup_analyst_initial', 'GS')}<br/>{context_data.get('test_date', '05 Feb 2026')}", cell_body_style),
        Paragraph(context_data.get('sampling_location', 'Settling Sampling Plate'), cell_body_style),
        Paragraph(context_data.get('reader_48h', 'MC'), cell_body_style),
        Paragraph(context_data.get('cfu_obs_48h', '10 CFU on Settling S2 sampling plate'), cell_body_style),
        Paragraph(context_data.get('reader_5d', 'SMO'), cell_body_style),
        Paragraph(context_data.get('cfu_obs_5d', '10 CFU on Settling S2 sampling plate'), cell_body_style),
        Paragraph(context_data.get('microbial_id', 'Staphylococcus capitis (Gram (+) cocci), Staphylococcus hominis (Gram (+) cocci)'), cell_body_left)
    ]

    t1_data = [t1_headers, t1_row_vals]
    t1_col_widths = [62, 55, 65, 45, 85, 45, 85, 110]
    
    t1 = Table(t1_data, colWidths=t1_col_widths)
    t1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D9D9D9')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
    ]))
    elements.append(t1)
    elements.append(Spacer(1, 8))

    # --- TABLE 2 ---
    elements.append(Paragraph("Table 2: Environmental Monitoring for Analyst & Cleanroom", title_style))
    
    t2_headers = [
        Paragraph("Environmental Monitoring<br/>(EM) Sampling Site", cell_hdr_style),
        Paragraph("Frequency", cell_hdr_style),
        Paragraph("Date<br/>(DDMMMYYYY)", cell_hdr_style),
        Paragraph("Analyst<br/>(Initials)", cell_hdr_style),
        Paragraph("Day /Week(s)", cell_hdr_style),
        Paragraph("Observation", cell_hdr_style),
        Paragraph("Plate<br/>ETX ID", cell_hdr_style),
        Paragraph("Microbial ID", cell_hdr_style),
        Paragraph("Notes", cell_hdr_style)
    ]

    t2_col_widths = [115, 38, 48, 38, 62, 65, 52, 98, 36]
    
    t2_data = [t2_headers]
    
    table_styles = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D9D9D9')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
    ]

    # Helper to add section headers and data rows
    def add_section(title, rows):
        r_idx = len(t2_data)
        # 1-cell spanning row
        hdr_cell = Paragraph(title, cell_section_hdr)
        t2_data.append([hdr_cell] + [''] * 8)
        table_styles.append(('SPAN', (0, r_idx), (-1, r_idx)))
        table_styles.append(('BACKGROUND', (0, r_idx), (-1, r_idx), colors.HexColor('#F2F2F2')))
        table_styles.append(('TOPPADDING', (0, r_idx), (-1, r_idx), 2))
        table_styles.append(('BOTTOMPADDING', (0, r_idx), (-1, r_idx), 2))
        
        for r in rows:
            data_row = []
            for col_i, text in enumerate(r):
                if col_i in [0, 7]:
                    data_row.append(Paragraph(str(text), cell_body_left))
                else:
                    data_row.append(Paragraph(str(text), cell_body_style))
            t2_data.append(data_row)

    # 1. Personnel
    add_section("Personnel EM Bracketing", [
        ["Personal (Left & Right Touch)", "Daily", context_data.get('before_date', '04 Feb 2026'), context_data.get('analyst_initial', 'GS'), "Date Before Testing", context_data.get('pers_obs_before', 'No growth'), context_data.get('pers_etx_before', 'N/A'), context_data.get('pers_id_before', 'N/A'), "None"],
        ["Personal (Left & Right Touch)", "Daily", context_data.get('test_date', '05 Feb 2026'), context_data.get('analyst_initial', 'GS'), "Date of Testing", context_data.get('pers_obs_during', 'No growth'), context_data.get('pers_etx_during', 'N/A'), context_data.get('pers_id_during', 'N/A'), "None"],
        ["Personal (Left & Right Touch)", "Daily", context_data.get('after_date', '06 Feb 2026'), context_data.get('analyst_initial', 'GS'), "Date After Testing", context_data.get('pers_obs_after', 'No growth'), context_data.get('pers_etx_after', 'N/A'), context_data.get('pers_id_after', 'N/A'), "None"],
    ])

    # 2. BSC
    add_section("Biological Safety Cabinet (BSC) EM Bracketing", [
        ["Surface Sampling of ISO 5 (4 loc)", "Daily", context_data.get('before_date', '04 Feb 2026'), context_data.get('bsc_surf_analyst_before', 'GS'), "Date Before Testing", context_data.get('bsc_surf_obs_before', 'No growth'), context_data.get('bsc_surf_etx_before', 'N/A'), context_data.get('bsc_surf_id_before', 'N/A'), "None"],
        ["Surface Sampling of ISO 5 (4 loc)", "Daily", context_data.get('test_date', '05 Feb 2026'), context_data.get('bsc_surf_analyst_during', 'GS'), "Date of Testing", context_data.get('bsc_surf_obs_during', 'No growth'), context_data.get('bsc_surf_etx_during', 'N/A'), context_data.get('bsc_surf_id_during', 'N/A'), "None"],
        ["Surface Sampling of ISO 5 (4 loc)", "Daily", context_data.get('after_date', '06 Feb 2026'), context_data.get('bsc_surf_analyst_after', 'GS'), "Date After Testing", context_data.get('bsc_surf_obs_after', 'No growth'), context_data.get('bsc_surf_etx_after', 'N/A'), context_data.get('bsc_surf_id_after', 'N/A'), "None"],
        ["Settling Sampling of ISO 5 (2 loc)", "Daily", context_data.get('before_date', '04 Feb 2026'), context_data.get('bsc_sett_analyst_before', 'GS'), "Date Before Testing", context_data.get('bsc_sett_obs_before', 'No growth'), context_data.get('bsc_sett_etx_before', 'N/A'), context_data.get('bsc_sett_id_before', 'N/A'), "None"],
        ["Settling Sampling of ISO 5 (2 loc)", "Daily", context_data.get('test_date', '05 Feb 2026'), context_data.get('bsc_sett_analyst_during', 'GS'), "Date of Testing", context_data.get('bsc_sett_obs_during', '10 CFU (Settling S2)'), context_data.get('etx_id', 'ETX-260216-0348'), context_data.get('microbial_id', 'Staphylococcus capitis...'), "None"],
        ["Settling Sampling of ISO 5 (2 loc)", "Daily", context_data.get('after_date', '06 Feb 2026'), context_data.get('bsc_sett_analyst_after', 'GS'), "Date After Testing", context_data.get('bsc_sett_obs_after', 'No growth'), context_data.get('bsc_sett_etx_after', 'N/A'), context_data.get('bsc_sett_id_after', 'N/A'), "None"],
    ])

    # 3. Weekly Air
    add_section("Weekly Active Air Sampling Bracketing", [
        ["Active Air Sampling of Cleanrooms", "Weekly", context_data.get('date_of_weekly_air', '06 Feb 2026'), context_data.get('weekly_air_analyst', 'SMO'), "Week (On or After Date)", context_data.get('air_obs', 'No growth'), context_data.get('air_etx', 'N/A'), context_data.get('air_id', 'N/A'), "None"],
    ])

    # 4. Weekly Surface
    add_section("Surface Sampling of Anteroom & Cleanroom Bracketing", [
        ["Surface Sampling of Cleanrooms", "Weekly", context_data.get('date_of_weekly_surf', '06 Feb 2026'), context_data.get('weekly_surf_analyst', 'SMO'), "Week (On or After Date)", context_data.get('room_surf_obs', 'No growth'), context_data.get('room_surf_etx', 'N/A'), context_data.get('room_surf_id', 'N/A'), "None"],
    ])

    t2 = Table(t2_data, colWidths=t2_col_widths)
    t2.setStyle(TableStyle(table_styles))
    elements.append(t2)

    doc.build(elements)
    buf.seek(0)
    return buf

if __name__ == '__main__':
    dummy_data = {
        'etx_id': 'ETX-260216-0348',
        'setup_analyst_initial': 'GS',
        'test_date': '05 Feb 2026',
        'sampling_location': 'Settling Sampling Plate (BSC 1314)',
        'reader_48h': 'MC',
        'cfu_obs_48h': '10 CFU on Settling S2 sampling plate',
        'reader_5d': 'SMO',
        'cfu_obs_5d': '10 CFU on Settling S2 sampling plate',
        'microbial_id': 'Staphylococcus capitis (Gram (+) cocci), Staphylococcus hominis (Gram (+) cocci), Kocuria indica (Gram (+) cocci), Micrococcus luteus (Gram (+) cocci) and Staphylococcus epidermidis (Gram (+) cocci)'
    }
    buf = generate_em_tables_page_pdf(dummy_data)
    with open('scratch/test_page7.pdf', 'wb') as f:
        f.write(buf.read())
    print('Generated scratch/test_page7.pdf successfully!')
