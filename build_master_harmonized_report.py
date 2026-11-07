"""
University of Mumbai - Community Engagement Project (CEP) Report Generator
Harmonized Master Edition: Matching Reference Structure & University NEP 2020 Guidelines
while Preserving Original Byculla Zoo Research, Rich Deep Emerald Design & Visual Polish

Key Specifications:
- Palette: Deep Forest Emerald (#1B4E34), Slate (#276C48), Mid Leaf (#3D7A5A), Soft Tint (#F2F7F4), Gold (#C0741F)
- Typography: Times New Roman, 12pt body text, 1.3-1.4 line spacing for academic density, 13-14pt bold headings
- Section 1 (Pages 1 to 3): Title Page, Certificate, Declaration (Unnumbered)
- Section 2 (Pages 4 to 30, Numbered 1 to 27):
    Page 1: Acknowledgement
    Page 2: Abstract
    Page 3: Table of Contents
    Page 4: List of Figures & Tables
    Pages 5-6: Chapter 1 - Introduction (1.1, 1.2, 1.3, 1.4)
    Pages 7-9: Chapter 2 - Literature Review (2.1 to 2.6)
    Pages 10-11: Chapter 3 - Methodology (3.1 to 3.5)
    Pages 12-18: Chapter 4 - Community Interactions, Observations & Analysis (4.1 to 4.5)
    Pages 19-20: Chapter 5 - Conclusion & Recommendations (5.1, 5.2)
    Page 21: References
    Pages 22-23: Appendices (A, B, C)
    Page 24: Guide Interaction Diary Form (5 Meetings with Prof. Saniya Gonsalves)
    Pages 25-27: Student Feedback on CEP (Annexure V 16-Criteria Likert Table)
Total Document Length: Exactly 30 Pages (Matching Reference Report RAHUL33_SHIVAM44_CEP.pdf)
"""

import os
import sys
import docx
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import win32com.client

# Academic Executive Deep Forest Emerald Palette
HEX_PRIMARY = "1B4E34"       # Deep Forest Emerald
HEX_SECONDARY = "276C48"     # Slate Emerald
HEX_ACCENT = "3D7A5A"        # Mid Leaf Green
HEX_TINT = "F2F7F4"          # Soft Mint Background Wash
HEX_TINT_LIGHT = "F8FAF9"    # Subtle alternate row tint
HEX_TEXT_DARK = "1C211B"     # Charcoal Dark Text
HEX_MUTED = "5A6257"         # Muted Slate Gray
HEX_BORDER = "C8D4CD"        # Clean Border Gray
HEX_GOLD = "C0741F"          # Warm Amber / Bronze Highlight

RGB_PRIMARY = RGBColor(0x1B, 0x4E, 0x34)
RGB_SECONDARY = RGBColor(0x27, 0x6C, 0x48)
RGB_TEXT = RGBColor(0x1C, 0x21, 0x1B)
RGB_MUTED = RGBColor(0x5A, 0x62, 0x57)
RGB_GOLD = RGBColor(0xC0, 0x74, 0x1F)

def generate_document(page_map=None):
    if page_map is None:
        page_map = {}

    def pg(key, default_val):
        return str(page_map.get(key, default_val))

    doc = docx.Document()

    # =========================================================================
    # SECTION 1: PRELIMINARY PAGES (Pages 1 to 3, Unnumbered)
    # =========================================================================
    sec1 = doc.sections[0]
    sec1.page_width = Inches(8.27)
    sec1.page_height = Inches(11.69) # A4
    sec1.top_margin = Inches(0.85)
    sec1.bottom_margin = Inches(0.85)
    sec1.left_margin = Inches(0.95)
    sec1.right_margin = Inches(0.95)
    sec1.different_first_page_header_footer = True
    sec1.header.is_linked_to_previous = False
    sec1.footer.is_linked_to_previous = False

    # XML Helper Functions
    def set_cell_shading(cell, hex_color):
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        cell._tc.get_or_add_tcPr().append(shd)

    def set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'''
            <w:tcMar {nsdecls("w")}>
                <w:top w:w="{top}" w:type="dxa"/>
                <w:bottom w:w="{bottom}" w:type="dxa"/>
                <w:left w:w="{left}" w:type="dxa"/>
                <w:right w:w="{right}" w:type="dxa"/>
            </w:tcMar>
        ''')
        tcPr.append(tcMar)

    def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
        tcPr = cell._tc.get_or_add_tcPr()
        borders_elm = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
        sides = {'top': top, 'bottom': bottom, 'left': left, 'right': right}
        for side, border in sides.items():
            if border:
                val, sz, col = border
                elm = parse_xml(f'<w:{side} {nsdecls("w")} w:val="{val}" w:sz="{sz}" w:space="0" w:color="{col}"/>')
                borders_elm.append(elm)
            else:
                elm = parse_xml(f'<w:{side} {nsdecls("w")} w:val="none"/>')
                borders_elm.append(elm)
        tcPr.append(borders_elm)

    def style_table_header(row, col_widths=None):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for i, cell in enumerate(row.cells):
            set_cell_shading(cell, HEX_PRIMARY)
            set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
            set_cell_borders(cell, 
                             top=("single", "8", HEX_PRIMARY), 
                             bottom=("single", "14", HEX_SECONDARY), 
                             left=("none", "0", "auto"), 
                             right=("none", "0", "auto"))
            if col_widths and i < len(col_widths):
                cell.width = col_widths[i]
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.05
                for r in p.runs:
                    r.font.name = "Times New Roman"; r.font.size = Pt(9.0); r.font.bold = True
                    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    def style_table_rows(table, col_widths=None):
        for r_idx, row in enumerate(table.rows[1:]):
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            bg_col = HEX_TINT_LIGHT if r_idx % 2 == 1 else "FFFFFF"
            for c_idx, cell in enumerate(row.cells):
                set_cell_shading(cell, bg_col)
                set_cell_margins(cell, top=45, bottom=45, left=70, right=70)
                set_cell_borders(cell, 
                                 top=("single", "4", HEX_BORDER), 
                                 bottom=("single", "4", HEX_BORDER), 
                                 left=("none", "0", "auto"), 
                                 right=("none", "0", "auto"))
                if col_widths and c_idx < len(col_widths):
                    cell.width = col_widths[c_idx]
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.12
                    for r in p.runs:
                        r.font.name = "Times New Roman"; r.font.size = Pt(8.5); r.font.color.rgb = RGB_TEXT

    def add_chapter_title(number_str, title_str):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(6); h.paragraph_format.space_after = Pt(2)
        h.paragraph_format.keep_with_next = True
        r_num = h.add_run(f"CHAPTER {number_str}\n")
        r_num.font.name = "Times New Roman"; r_num.font.size = Pt(13); r_num.font.bold = True; r_num.font.color.rgb = RGB_SECONDARY
        r_title = h.add_run(title_str.upper())
        r_title.font.name = "Times New Roman"; r_title.font.size = Pt(14.5); r_title.font.bold = True; r_title.font.color.rgb = RGB_PRIMARY
        
        p_line = doc.add_paragraph()
        p_line.paragraph_format.space_before = Pt(1); p_line.paragraph_format.space_after = Pt(8)
        p_line.paragraph_format.keep_with_next = True
        r_line = p_line.add_run("―" * 46)
        r_line.font.name = "Times New Roman"; r_line.font.size = Pt(10); r_line.font.color.rgb = RGB_SECONDARY
        return h

    def add_heading_1(text):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(9); h.paragraph_format.space_after = Pt(3)
        h.paragraph_format.keep_with_next = True
        r = h.add_run(text)
        r.font.name = "Times New Roman"; r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = RGB_PRIMARY
        return h

    def add_heading_2(text):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(7); h.paragraph_format.space_after = Pt(2)
        h.paragraph_format.keep_with_next = True
        r = h.add_run(text)
        r.font.name = "Times New Roman"; r.font.size = Pt(11.5); r.font.bold = True; r.font.color.rgb = RGB_SECONDARY
        return h

    def add_heading_3(text):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(5); h.paragraph_format.space_after = Pt(1.5)
        h.paragraph_format.keep_with_next = True
        r = h.add_run(text)
        r.font.name = "Times New Roman"; r.font.size = Pt(10.5); r.font.bold = True; r.font.italic = True; r.font.color.rgb = RGB_TEXT
        return h

    def add_body_p(text, bold_prefix=None, space_after=4, line_spacing=1.35, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = "Times New Roman"; r_pre.font.size = Pt(11.5); r_pre.font.bold = True; r_pre.font.color.rgb = RGB_TEXT
        r = p.add_run(text)
        r.font.name = "Times New Roman"; r.font.size = Pt(11.5); r.font.color.rgb = RGB_TEXT
        return p

    def add_bullet_p(text, bold_prefix=None, space_after=2.5, line_spacing=1.35):
        p = doc.add_paragraph(style='List Bullet')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = "Times New Roman"; r_pre.font.size = Pt(11.5); r_pre.font.bold = True; r_pre.font.color.rgb = RGB_TEXT
        r = p.add_run(text)
        r.font.name = "Times New Roman"; r.font.size = Pt(11.5); r.font.color.rgb = RGB_TEXT
        return p

    def add_callout(title, text, icon="📌"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.autofit = False
        cell = tbl.cell(0, 0); cell.width = Inches(6.37)
        set_cell_shading(cell, HEX_TINT)
        set_cell_margins(cell, top=70, bottom=70, left=120, right=100)
        set_cell_borders(cell, left=("single", "32", HEX_PRIMARY), top=("none", "0", "auto"), right=("none", "0", "auto"), bottom=("none", "0", "auto"))
        p = cell.paragraphs[0]; p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.2
        r_title = p.add_run(f"{icon}  {title}\n")
        r_title.font.name = "Times New Roman"; r_title.font.size = Pt(10.5); r_title.font.bold = True; r_title.font.color.rgb = RGB_PRIMARY
        r_body = p.add_run(text)
        r_body.font.name = "Times New Roman"; r_body.font.size = Pt(10.0); r_body.font.color.rgb = RGB_TEXT
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    def add_species_box(fig_no, common_name, sci_name, taxon_class, obs_id, zone, gps, date_time, 
                        photo_path=None, img_width=1.9, photo_desc="", id_marks="", behavior_notes="", ecology_status=""):
        tbl = doc.add_table(rows=1, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.autofit = False
        c0 = tbl.cell(0, 0); c1 = tbl.cell(0, 1)
        c0.width = Inches(2.25); c1.width = Inches(4.12)
        set_cell_margins(c0, top=40, bottom=40, left=45, right=45)
        set_cell_margins(c1, top=40, bottom=40, left=60, right=45)
        b_box = ("single", "8", HEX_PRIMARY)
        set_cell_borders(c0, top=b_box, bottom=b_box, left=b_box, right=b_box)
        set_cell_borders(c1, top=b_box, bottom=b_box, left=b_box, right=b_box)
        
        p0 = c0.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_before = Pt(1); p0.paragraph_format.space_after = Pt(1)
        
        if photo_path and os.path.exists(photo_path):
            r0 = p0.add_run()
            r0.add_picture(photo_path, width=Inches(img_width))
            p_cap = c0.add_paragraph(); p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(1); p_cap.paragraph_format.space_after = Pt(0)
            rcap = p_cap.add_run(f"Fig. {fig_no}: {common_name}\nGPS: {gps}\n{date_time} • Byculla")
            rcap.font.name = "Times New Roman"; rcap.font.size = Pt(7.0); rcap.font.italic = True; rcap.font.color.rgb = RGB_MUTED
        else:
            p0.paragraph_format.space_before = Pt(8)
            r_icon = p0.add_run("📷  [ FIELD SIGHTING ]\n\n")
            r_icon.font.name = "Times New Roman"; r_icon.font.size = Pt(8.5); r_icon.font.bold = True; r_icon.font.color.rgb = RGB_PRIMARY
            r_t = p0.add_run(f"{common_name}\n({sci_name})\n\n")
            r_t.font.name = "Times New Roman"; r_t.font.size = Pt(8.0); r_t.font.bold = True; r_t.font.color.rgb = RGB_SECONDARY
            p_cap = c0.add_paragraph(); p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(4); p_cap.paragraph_format.space_after = Pt(0)
            rcap = p_cap.add_run(f"Fig. {fig_no}: {common_name} ({sci_name})\nGPS: {gps}\n{date_time}")
            rcap.font.name = "Times New Roman"; rcap.font.size = Pt(7.0); rcap.font.italic = True; rcap.font.color.rgb = RGB_MUTED

        p1 = c1.paragraphs[0]; p1.paragraph_format.space_before = Pt(1); p1.paragraph_format.space_after = Pt(1)
        r_hdr = p1.add_run(f"FIELD VOUCHER: {obs_id}  |  ")
        r_hdr.font.name = "Times New Roman"; r_hdr.font.size = Pt(8.5); r_hdr.font.bold = True; r_hdr.font.color.rgb = RGB_PRIMARY
        r_cn = p1.add_run(f"{common_name} ")
        r_cn.font.name = "Times New Roman"; r_cn.font.size = Pt(8.5); r_cn.font.bold = True
        r_sn = p1.add_run(f"({sci_name})\n")
        r_sn.font.name = "Times New Roman"; r_sn.font.size = Pt(8.0); r_sn.font.italic = True; r_sn.font.color.rgb = RGB_MUTED
        
        data_items = [
            ("Taxon & Order", taxon_class),
            ("GPS Coordinates", gps),
            ("Microhabitat / Zone", f"{zone} • {photo_desc}"),
            ("Diagnostic Marks", id_marks),
            ("Ethological Notes", behavior_notes),
            ("Conservation Role", ecology_status)
        ]
        for label, val in data_items:
            p_row = c1.add_paragraph()
            p_row.paragraph_format.space_before = Pt(0); p_row.paragraph_format.space_after = Pt(1)
            p_row.paragraph_format.line_spacing = 1.08
            rl = p_row.add_run(f"• {label}: ")
            rl.font.name = "Times New Roman"; rl.font.size = Pt(7.5); rl.font.bold = True; rl.font.color.rgb = RGB_SECONDARY
            rv = p_row.add_run(val)
            rv.font.name = "Times New Roman"; rv.font.size = Pt(7.5); rv.font.color.rgb = RGB_TEXT

        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # ----------------------------------------------------
    # PAGE 1: TITLE PAGE (Matching Reference Proforma & Logo)
    # ----------------------------------------------------
    p_t1 = doc.add_paragraph()
    p_t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t1.paragraph_format.space_before = Pt(4); p_t1.paragraph_format.space_after = Pt(2)
    r = p_t1.add_run("A Community Engagement Project (CEP) Report\nOn\n")
    r.font.name = "Times New Roman"; r.font.size = Pt(12.5); r.font.bold = True; r.font.color.rgb = RGB_SECONDARY

    r_theme = p_t1.add_run("“OPEN-SOURCE SOLUTION FOR SOCIAL NEEDS”\n\n")
    r_theme.font.name = "Times New Roman"; r_theme.font.size = Pt(13); r_theme.font.bold = True; r_theme.font.color.rgb = RGB_PRIMARY

    p_proj_title = doc.add_paragraph()
    p_proj_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_proj_title.paragraph_format.space_before = Pt(0); p_proj_title.paragraph_format.space_after = Pt(3)
    p_proj_title.paragraph_format.line_spacing = 1.2
    r_title = p_proj_title.add_run("“LOCAL BIODIVERSITY MAPPING OF BYCULLA ZOO\nAND SURROUNDING GREEN SPACES”\n")
    r_title.font.name = "Times New Roman"; r_title.font.size = Pt(14); r_title.font.bold = True; r_title.font.color.rgb = RGB_PRIMARY
    r_sub = p_proj_title.add_run("(A Geospatial Biodiversity Census, Web-GIS Cartography & Community Ecological Literacy Study)\n")
    r_sub.font.name = "Times New Roman"; r_sub.font.size = Pt(10); r_sub.font.italic = True; r_sub.font.color.rgb = RGB_GOLD

    p_subby = doc.add_paragraph()
    p_subby.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_subby.paragraph_format.space_before = Pt(2); p_subby.paragraph_format.space_after = Pt(2)
    p_subby.add_run("Submitted by\n").font.name = "Times New Roman"

    p_names = doc.add_paragraph()
    p_names.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_names.paragraph_format.space_before = Pt(0); p_names.paragraph_format.space_after = Pt(4)
    r_n1 = p_names.add_run("RIDDHESH KENI (Roll No: 35)\n")
    r_n1.font.name = "Times New Roman"; r_n1.font.size = Pt(12); r_n1.font.bold = True; r_n1.font.color.rgb = RGB_TEXT
    r_n2 = p_names.add_run("SHIVANSH DUBEY (Roll No: 45)\n")
    r_n2.font.name = "Times New Roman"; r_n2.font.size = Pt(12); r_n2.font.bold = True; r_n2.font.color.rgb = RGB_TEXT

    p_deg = doc.add_paragraph()
    p_deg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_deg.paragraph_format.space_before = Pt(2); p_deg.paragraph_format.space_after = Pt(4)
    p_deg.paragraph_format.line_spacing = 1.15
    p_deg.add_run("In partial fulfillment for the award of the degree of\n").font.name = "Times New Roman"
    r_dn = p_deg.add_run("BACHELOR OF SCIENCE in COMPUTER SCIENCE\n")
    r_dn.font.name = "Times New Roman"; r_dn.font.size = Pt(12); r_dn.font.bold = True; r_dn.font.color.rgb = RGB_PRIMARY
    p_deg.add_run("(T.Y. B.Sc. CS — Semester V, NEP 2020 Guidelines)\n").font.name = "Times New Roman"

    p_guide = doc.add_paragraph()
    p_guide.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_guide.paragraph_format.space_before = Pt(2); p_guide.paragraph_format.space_after = Pt(4)
    p_guide.paragraph_format.line_spacing = 1.15
    p_guide.add_run("under the guidance of\n").font.name = "Times New Roman"
    r_prof = p_guide.add_run("PROF. Shweta Iyer\n")
    r_prof.font.name = "Times New Roman"; r_prof.font.size = Pt(12); r_prof.font.bold = True; r_prof.font.color.rgb = RGB_PRIMARY
    r_dept = p_guide.add_run("Department of Computer Science\n")
    r_dept.font.name = "Times New Roman"; r_dept.font.size = Pt(11); r_dept.font.color.rgb = RGB_SECONDARY

    # College Logo
    if os.path.exists("college_logo.jpg"):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(2); p_logo.paragraph_format.space_after = Pt(4)
        p_logo.add_run().add_picture("college_logo.jpg", width=Inches(1.22))

    p_col = doc.add_paragraph()
    p_col.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_col.paragraph_format.space_before = Pt(2); p_col.paragraph_format.space_after = Pt(0)
    p_col.paragraph_format.line_spacing = 1.15
    r_c1 = p_col.add_run("VIDYAVARDHINI’S\n")
    r_c1.font.name = "Times New Roman"; r_c1.font.size = Pt(11); r_c1.font.bold = True
    r_c2 = p_col.add_run("A. V. COLLEGE OF ARTS, K. M. COLLEGE OF COMMERCE, E. S. A. COLLEGE OF SCIENCE\n(VARTAK COLLEGE)\n")
    r_c2.font.name = "Times New Roman"; r_c2.font.size = Pt(10.5); r_c2.font.bold = True; r_c2.font.color.rgb = RGB_SECONDARY
    r_c3 = p_col.add_run("(Affiliated to University of Mumbai)\nVASAI ROAD (WEST), PALGHAR – 401202, MAHARASHTRA\n(SEM-V)  •  ACADEMIC YEAR: (2026 – 2027)")
    r_c3.font.name = "Times New Roman"; r_c3.font.size = Pt(10); r_c3.font.color.rgb = RGB_TEXT

    doc.add_page_break() # -> PAGE 2

    # ----------------------------------------------------
    # PAGE 2: CERTIFICATE (Matching Reference College Form)
    # ----------------------------------------------------
    p_cert_hdr = doc.add_paragraph()
    p_cert_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cert_hdr.paragraph_format.space_before = Pt(6); p_cert_hdr.paragraph_format.space_after = Pt(6)
    p_cert_hdr.paragraph_format.line_spacing = 1.15
    r_ch1 = p_cert_hdr.add_run("VIDYAVARDHINI’S\n")
    r_ch1.font.name = "Times New Roman"; r_ch1.font.size = Pt(12); r_ch1.font.bold = True
    r_ch2 = p_cert_hdr.add_run("A. V. COLLEGE OF ARTS, K. M. COLLEGE OF COMMERCE, E. S. A. COLLEGE OF SCIENCE\n(VARTAK COLLEGE)\n")
    r_ch2.font.name = "Times New Roman"; r_ch2.font.size = Pt(11); r_ch2.font.bold = True; r_ch2.font.color.rgb = RGB_SECONDARY
    r_ch3 = p_cert_hdr.add_run("(Affiliated to University of Mumbai)\nVasai Road (West), District Palghar – 401202, Maharashtra\nDEPARTMENT OF COMPUTER SCIENCE\n")
    r_ch3.font.name = "Times New Roman"; r_ch3.font.size = Pt(9.5); r_ch3.font.color.rgb = RGB_MUTED

    p_cdiv = doc.add_paragraph(); p_cdiv.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cdiv.paragraph_format.space_before = Pt(2); p_cdiv.paragraph_format.space_after = Pt(8)
    p_cdiv.add_run("―" * 46).font.color.rgb = RGB_SECONDARY

    p_ctitle = doc.add_paragraph(); p_ctitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ctitle.paragraph_format.space_before = Pt(0); p_ctitle.paragraph_format.space_after = Pt(12)
    r_ct = p_ctitle.add_run("CERTIFICATE")
    r_ct.font.name = "Times New Roman"; r_ct.font.size = Pt(14); r_ct.font.bold = True; r_ct.font.color.rgb = RGB_PRIMARY

    p_cert_body = doc.add_paragraph()
    p_cert_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_cert_body.paragraph_format.space_before = Pt(0); p_cert_body.paragraph_format.space_after = Pt(8)
    p_cert_body.paragraph_format.line_spacing = 1.35
    p_cert_body.add_run(
        "This is to certify that Mr. RIDDHESH KENI (Roll No. 35) and Mr. SHIVANSH DUBEY (Roll No. 45), "
        "students of Vidyavardhini’s A. V. College of Arts, K. M. College of Commerce, E. S. A. College of Science "
        "(Vartak College), studying in the Bachelor of Science in Computer Science (T.Y. B.Sc. CS — Semester V) "
        "program, have satisfactorily completed their Community Engagement Project (CEP) titled:\n"
    ).font.name = "Times New Roman"

    add_callout(
        "Approved Community Engagement Project Title",
        "“LOCAL BIODIVERSITY MAPPING OF BYCULLA ZOO AND SURROUNDING GREEN SPACES (A Geospatial Biodiversity Census, Web-GIS Cartography & Community Ecological Literacy Study)”",
        icon="📜"
    )

    p_cert_body2 = doc.add_paragraph()
    p_cert_body2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_cert_body2.paragraph_format.space_before = Pt(4); p_cert_body2.paragraph_format.space_after = Pt(18)
    p_cert_body2.paragraph_format.line_spacing = 1.35
    p_cert_body2.add_run(
        "under the theme “OPEN-SOURCE SOLUTION FOR SOCIAL NEEDS” for the academic year 2026–2027, "
        "in partial fulfillment of the requirements for the award of the Degree of Bachelor of Science in Computer Science "
        "as per the guidelines formulated by the University of Mumbai under the National Education Policy (NEP 2020) carrying a weightage of two (2) credits.\n\n"
        "To the best of my knowledge, the work presented in this report is original, authentic, and the information included in the project is correct."
    ).font.name = "Times New Roman"

    # Signature Blocks matching reference
    tbl_csig = doc.add_table(rows=1, cols=3)
    tbl_csig.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_csig.autofit = False
    csig_widths = [Inches(2.1), Inches(2.1), Inches(2.17)]
    for c_i, text in enumerate([
        "_______________________\nInternal Guide / Mentor\nPROF. Shweta Iyer\nDept. of Computer Science", 
        "_______________________\nHead of the Department\nDept. of Computer Science\nVartak College", 
        "_______________________\nPrincipal\nVartak College\nVasai Road (West)"
    ]):
        cell = tbl_csig.cell(0, c_i); cell.width = csig_widths[c_i]
        set_cell_borders(cell)
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.line_spacing = 1.15
        r = p.add_run(text); r.font.name = "Times New Roman"; r.font.size = Pt(9.5); r.font.bold = True

    p_sig_ext = doc.add_paragraph()
    p_sig_ext.paragraph_format.space_before = Pt(18); p_sig_ext.paragraph_format.space_after = Pt(0)
    p_sig_ext.paragraph_format.line_spacing = 1.15
    r_ext = p_sig_ext.add_run("External Examiner Signature: _______________________          Date: ___/___/2026\nCollege Seal / Stamp:                                             Place: Vasai Road (W)")
    r_ext.font.name = "Times New Roman"; r_ext.font.size = Pt(9.5); r_ext.font.italic = True

    doc.add_page_break() # -> PAGE 3

    # ----------------------------------------------------
    # PAGE 3: DECLARATION (Strictly Matching Reference Page 3)
    # ----------------------------------------------------
    p_dhdr = doc.add_paragraph()
    p_dhdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_dhdr.paragraph_format.space_before = Pt(12); p_dhdr.paragraph_format.space_after = Pt(24)
    p_dhdr.paragraph_format.line_spacing = 1.2
    r_dh1 = p_dhdr.add_run("VIDYAVARDHINI’S\n")
    r_dh1.font.name = "Times New Roman"; r_dh1.font.size = Pt(12); r_dh1.font.bold = True
    r_dh2 = p_dhdr.add_run("A. V. COLLEGE OF ARTS, K. M. COLLEGE OF COMMERCE,\nE. S. A. COLLEGE OF SCIENCE,\n")
    r_dh2.font.name = "Times New Roman"; r_dh2.font.size = Pt(11.5); r_dh2.font.bold = True
    r_dh3 = p_dhdr.add_run("(Affiliated by University of Mumbai)\n\n\n")
    r_dh3.font.name = "Times New Roman"; r_dh3.font.size = Pt(11)

    r_dt = p_dhdr.add_run("Declaration\n")
    r_dt.font.name = "Times New Roman"; r_dt.font.size = Pt(13); r_dt.font.bold = True

    p_dec_body = doc.add_paragraph()
    p_dec_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_dec_body.paragraph_format.space_before = Pt(16); p_dec_body.paragraph_format.space_after = Pt(14)
    p_dec_body.paragraph_format.line_spacing = 1.35
    r_db1 = p_dec_body.add_run(
        "I,  Mr./Ms.  _____________________________________________  Student  of  \n"
        "________________________________  Institute/college  studying  in  \n"
        "__________________________________ program, hereby declare that I have completed the  \n"
        "Community  Engagement  Project  titled  "
    )
    r_db1.font.name = "Times New Roman"; r_db1.font.size = Pt(11)
    r_db2 = p_dec_body.add_run("“LOCAL BIODIVERSITY MAPPING OF BYCULLA ZOO AND SURROUNDING GREEN SPACES (Urban Ecology & Public Awareness Platform)”")
    r_db2.font.name = "Times New Roman"; r_db2.font.size = Pt(11); r_db2.font.bold = True
    r_db3 = p_dec_body.add_run(" during the academic year 2026-2027.")
    r_db3.font.name = "Times New Roman"; r_db3.font.size = Pt(11)

    p_dec_body2 = doc.add_paragraph()
    p_dec_body2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_dec_body2.paragraph_format.space_before = Pt(12); p_dec_body2.paragraph_format.space_after = Pt(14)
    p_dec_body2.paragraph_format.line_spacing = 1.35
    r_db4 = p_dec_body2.add_run("The report is original and the information/data included in the report is true emerging from the primary and/or secondary data gathered and analyzed as part of this Community Engagement project.")
    r_db4.font.name = "Times New Roman"; r_db4.font.size = Pt(11)

    p_dec_body3 = doc.add_paragraph()
    p_dec_body3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_dec_body3.paragraph_format.space_before = Pt(12); p_dec_body3.paragraph_format.space_after = Pt(50)
    p_dec_body3.paragraph_format.line_spacing = 1.35
    r_db5 = p_dec_body3.add_run("Due credit is extended on the work of Literature/Secondary Survey by endorsing it in the Bibliography as per prescribed format.")
    r_db5.font.name = "Times New Roman"; r_db5.font.size = Pt(11)

    p_decl_sig = doc.add_paragraph()
    p_decl_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_decl_sig.paragraph_format.space_before = Pt(100); p_decl_sig.paragraph_format.space_after = Pt(0)
    p_decl_sig.paragraph_format.line_spacing = 1.2
    r_ds = p_decl_sig.add_run("Signature of the Student with date\nName of Student")
    r_ds.font.name = "Times New Roman"; r_ds.font.size = Pt(11)

    # =========================================================================
    # SECTION 2: NUMBERED REPORT BODY (Pages 4 to 30, Numbered 1 to 27)
    # =========================================================================
    sec2 = doc.add_section()
    sec2.page_width = Inches(8.27)
    sec2.page_height = Inches(11.69)
    sec2.top_margin = Inches(0.85)
    sec2.bottom_margin = Inches(0.85)
    sec2.left_margin = Inches(0.95)
    sec2.right_margin = Inches(0.95)
    sec2.header.is_linked_to_previous = False
    sec2.footer.is_linked_to_previous = False

    # Force Section 2 to start page numbering at 1 (Matching Reference PDF!)
    sectPr = sec2._sectPr
    pgNumType = parse_xml(f'<w:pgNumType {nsdecls("w")} w:start="1"/>')
    sectPr.append(pgNumType)

    # Running Header on Section 2
    header2 = sec2.header
    p_head2 = header2.paragraphs[0]
    p_head2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_head2.paragraph_format.space_after = Pt(2)
    r_h2_1 = p_head2.add_run("University of Mumbai • Community Engagement Project (NEP 2020)  |  ")
    r_h2_1.font.name = "Times New Roman"; r_h2_1.font.size = Pt(8.5); r_h2_1.font.color.rgb = RGB_MUTED
    r_h2_2 = p_head2.add_run("Local Biodiversity Mapping")
    r_h2_2.font.name = "Times New Roman"; r_h2_2.font.size = Pt(8.5); r_h2_2.font.bold = True; r_h2_2.font.color.rgb = RGB_SECONDARY

    # Running Footer on Section 2: Bottom Center Page Number
    footer2 = sec2.footer
    p_foot2 = footer2.paragraphs[0]
    p_foot2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_foot2.paragraph_format.space_before = Pt(4)
    fld_page2 = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
    p_foot2._p.append(fld_page2)
    for r in p_foot2.runs:
        r.font.name = "Times New Roman"; r.font.size = Pt(10); r.font.color.rgb = RGB_TEXT

    # ----------------------------------------------------
    # PAGE 4 (Doc Page 4, Numbered 1): ACKNOWLEDGEMENT
    # ----------------------------------------------------
    add_heading_1("ACKNOWLEDGEMENT")
    add_body_p(
        "We would like to express our sincere gratitude to our institution, Vidyavardhini’s A. V. College of Arts, "
        "K. M. College of Commerce, E. S. A. College of Science (Vartak College), and the Department of Computer Science "
        "for providing us with the opportunity to undertake this Community Engagement Project (CEP) under the National Education Policy (NEP 2020) guidelines."
    )
    add_body_p(
        "We are profoundly indebted to our respected Principal, Dr. Arvind Ubale (Principal, Vartak College), for his visionary leadership, "
        "institutional support, and encouragement to pursue experiential community learning bridging information technology with real-world social and ecological challenges."
    )
    add_body_p(
        "We express our heartfelt thanks to the Head of the Department of Computer Science for granting necessary departmental permissions, "
        "providing access to computational resources, and offering valuable guidance during project milestones."
    )
    add_body_p(
        "Our deepest gratitude goes to our project guide, Prof. Shweta Iyer, Department of Computer Science, for her expert mentorship, "
        "continuous encouragement, and critical review throughout the conceptualization, fieldwork, and documentation stages. Her constructive feedback "
        "and supervisory sessions in the Guide Interaction Diary ensured that our work remained methodologically rigorous and strictly compliant with university standards."
    )
    add_body_p(
        "We extend our sincere appreciation to the administrative authorities and garden staff of Veermata Jijabai Bhosale Botanical Udyan and Zoo (Byculla Zoo), "
        "Brihanmumbai Municipal Corporation (BMC), for granting on-site study permissions. We are especially thankful to the veteran botanical caretakers (malis), "
        "notably Mr. Maruti Patil, whose decades of experiential botanical insights regarding heritage trees and avian nesting patterns enriched our observational records."
    )
    add_body_p(
        "We also thank the eighty-four community participants—morning walkers, local residents, teachers, and school students—who took time to participate in our surveys "
        "and provide actionable user feedback on our live Web-GIS portal (BioMap). Finally, we thank our parents and fellow classmates for their unwavering encouragement."
    )
    p_ack_sig = doc.add_paragraph(); p_ack_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p_ack_sig.paragraph_format.line_spacing = 1.15
    p_ack_sig.paragraph_format.space_before = Pt(12)
    p_ack_sig.add_run("RIDDHESH KENI (Roll No. 35)\nSHIVANSH DUBEY (Roll No. 45)\nThird Year B.Sc. (Computer Science)\nVartak College, Vasai").font.bold = True

    doc.add_page_break() # -> PAGE 5 (Doc Page 5, Numbered 2)

    # ----------------------------------------------------
    # PAGE 5 (Doc Page 5, Numbered 2): ABSTRACT
    # ----------------------------------------------------
    add_heading_1("ABSTRACT")
    add_body_p(
        "This Community Engagement Project, titled “Local Biodiversity Mapping of Byculla Zoo and Surrounding Green Spaces,” "
        "was undertaken to address the critical deficit in localized, open-access ecological information within Mumbai’s densely urbanized core. "
        "Veermata Jijabai Bhosale Botanical Udyan and Zoo (Byculla Zoo), spanning 53 acres in South Mumbai, represents a vital heritage ecological refugium "
        "housing over 3,213 trees across 256 species, alongside critical free-living wildlife populations."
    )
    add_body_p(
        "Despite its immense ecological value, primary community interactions (N = 84 participants) revealed a severe biodiversity literacy gap: "
        "over 78% of park visitors were unaware of the presence of century-old heritage trees and free-ranging avifauna, primarily because institutional signage "
        "is exclusively concentrated on caged zoo animals. To bridge this divide under the theme of “Open-Source Solution for Social Needs,” "
        "the investigators conducted an intensive multi-taxa field census coupled with acoustic decibel profiling and geospatial coordinate mapping."
    )
    add_body_p(
        "Empirical findings cataloged 42+ distinct species across Mammalia, Aves, Reptilia, Insecta, and Heritage Flora. Mathematical diversity analysis revealed "
        "a Shannon-Wiener Diversity Index of H′ = 2.68 within the interior botanical canopy (Zone 1), compared to H′ = 1.48 along the perimeter corridor (Zone 4). "
        "Simultaneously, acoustic decibel measurements demonstrated that the multi-tiered botanical canopy provides an acoustic damping effect of 26.3 dB, "
        "reducing ambient noise from 78.4 dBA at the perimeter traffic boundary to 52.1 dBA in the interior grove, sustaining sensitive avian breeding pairs."
    )
    add_body_p(
        "To translate these scientific observations into community action, the authors developed and deployed an open-source, mobile-responsive Web-GIS cartographic "
        "platform (BioMap, accessible live at biomap-byculla.netlify.app) utilizing OpenStreetMap and Leaflet.js. The platform provides interactive geolocation markers, "
        "bilingual vernacular ecological descriptions (Marathi and English), and micro-zone boundary overlays. This study delivers a three-tier municipal conservation "
        "blueprint proposing QR-coded botanical signage, acoustic vegetative buffers, and citizen-science integrations for the BMC Tree Authority."
    )
    add_callout(
        "Key Project Outcomes & Metrics",
        "• 42+ verified species cataloged with GPS vouchers  |  • Shannon Index H′ = 2.68 (Interior) vs. 1.48 (Perimeter)\n"
        "• 26.3 dBA acoustic noise attenuation measured      |  • Live open-source Web-GIS deployed on Netlify\n"
        "• 84 community stakeholders surveyed & analyzed     |  • 3-tier municipal policy blueprint for BMC Tree Authority",
        icon="📊"
    )
    p_kw = doc.add_paragraph(); p_kw.paragraph_format.space_before = Pt(4); p_kw.paragraph_format.space_after = Pt(0)
    r_kwh = p_kw.add_run("Keywords: "); r_kwh.font.bold = True; r_kwh.font.color.rgb = RGB_PRIMARY
    p_kw.add_run("Urban Biodiversity Mapping, Byculla Zoo, Open-Source Web-GIS, Shannon-Wiener Index, Acoustic Buffering, Community Engagement, NEP 2020.")

    doc.add_page_break() # -> PAGE 6 (Doc Page 6, Numbered 3)

    # ----------------------------------------------------
    # PAGE 6 (Doc Page 6, Numbered 3): TABLE OF CONTENTS
    # ----------------------------------------------------
    add_heading_1("TABLE OF CONTENTS")
    p_toc_note = doc.add_paragraph(); p_toc_note.paragraph_format.space_after = Pt(6)
    p_toc_note.add_run("Preliminary Sections and 5-Chapter Scheme (NEP 2020 Guidelines)").font.italic = True

    tbl_toc = doc.add_table(rows=22, cols=3)
    tbl_toc.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_toc.autofit = False
    toc_widths = [Inches(1.2), Inches(4.37), Inches(0.8)]
    
    # Table Header
    style_table_header(tbl_toc.rows[0], toc_widths)
    tbl_toc.cell(0, 0).paragraphs[0].text = "Section / Ch."
    tbl_toc.cell(0, 1).paragraphs[0].text = "Title / Subject Matter"
    tbl_toc.cell(0, 2).paragraphs[0].text = "Page"
    tbl_toc.cell(0, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    toc_items = [
        ("Prelim", "Acknowledgement", pg("ACK", "1")),
        ("Prelim", "Abstract", pg("ABS", "2")),
        ("Prelim", "List of Figures and Tables", pg("LOF", "4")),
        ("CHAPTER 1", "INTRODUCTION", pg("CH1", "5")),
        ("1.1", "Purpose of the Visit", pg("SEC_1_1", "5")),
        ("1.2", "Background Information", pg("SEC_1_2", "5")),
        ("1.3", "Significance", pg("SEC_1_3", "6")),
        ("1.4", "Scope of the Report", pg("SEC_1_4", "6")),
        ("CHAPTER 2", "LITERATURE REVIEW", pg("CH2", "7")),
        ("2.1", "Introduction", pg("SEC_2_1", "7")),
        ("2.2", "Urban Ecological Refugia & Historical Context", pg("SEC_2_2", "7")),
        ("2.3", "Citizen Science & Community Participation", pg("SEC_2_3", "8")),
        ("2.4", "Existing Biodiversity Platforms & Data Deficits", pg("SEC_2_4", "8")),
        ("2.5", "Technological Solutions & Open-Source Geospatial Mapping", pg("SEC_2_5", "9")),
        ("2.6", "Relevance to the BioMap Proposal", pg("SEC_2_6", "9")),
        ("CHAPTER 3", "METHODOLOGY", pg("CH3", "10")),
        ("3.1", "Data Collection Approach", pg("SEC_3_1", "10")),
        ("3.2", "Sampling Method & Survey Zones", pg("SEC_3_2", "10")),
        ("3.3", "Field Study and Location Evidence", pg("SEC_3_3", "11")),
        ("3.4", "Data Analysis & Quantitative Diversity Metrics", pg("SEC_3_4", "11")),
        ("3.5", "Rationale for the Selected Methods", pg("SEC_3_5", "11"))
    ]
    for r_i, (c0, c1, c2) in enumerate(toc_items, start=1):
        row = tbl_toc.rows[r_i]
        p0 = row.cells[0].paragraphs[0]; p0.text = c0
        p1 = row.cells[1].paragraphs[0]; p1.text = c1
        p2 = row.cells[2].paragraphs[0]; p2.text = c2; p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if c0.startswith("CHAPTER"):
            for p in [p0, p1, p2]:
                for r in p.runs: r.font.bold = True
    style_table_rows(tbl_toc, toc_widths)

    p_toc_part2 = doc.add_paragraph(); p_toc_part2.paragraph_format.space_before = Pt(4); p_toc_part2.paragraph_format.space_after = Pt(2)
    p_toc_part2.add_run("Table of Contents (Continued)").font.bold = True

    tbl_toc2 = doc.add_table(rows=12, cols=3)
    tbl_toc2.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_toc2.autofit = False
    style_table_header(tbl_toc2.rows[0], toc_widths)
    tbl_toc2.cell(0, 0).paragraphs[0].text = "Section / Ch."
    tbl_toc2.cell(0, 1).paragraphs[0].text = "Title / Subject Matter"
    tbl_toc2.cell(0, 2).paragraphs[0].text = "Page"
    tbl_toc2.cell(0, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    toc_items_part2 = [
        ("CHAPTER 4", "COMMUNITY INTERACTIONS, OBSERVATIONS & ANALYSIS", pg("CH4", "12")),
        ("4.1", "Community Interactions and Multi-Stakeholder Engagement", pg("SEC_4_1", "12")),
        ("4.2", "Field Observations & Location-Based Species Evidence", pg("SEC_4_2", "13")),
        ("4.3", "Survey Findings and Quantitative Analysis", pg("SEC_4_3", "15")),
        ("4.4", "Consolidated Key Findings", pg("SEC_4_4", "17")),
        ("4.5", "Analysis in Relation to Study Objectives & Web-GIS Portal", pg("SEC_4_5", "18")),
        ("CHAPTER 5", "CONCLUSION AND RECOMMENDATIONS", pg("CH5", "19")),
        ("5.1", "Conclusion", pg("SEC_5_1", "19")),
        ("5.2", "Recommendations & Three-Tier Municipal Blueprint", pg("SEC_5_2", "19")),
        ("References", "Academic References & Bibliography", pg("REF", "21")),
        ("Appendices", "Appendix A, B, C (Questionnaire, Architecture, Evidence)", pg("APP", "22"))
    ]
    for r_i, (c0, c1, c2) in enumerate(toc_items_part2, start=1):
        row = tbl_toc2.rows[r_i]
        p0 = row.cells[0].paragraphs[0]; p0.text = c0
        p1 = row.cells[1].paragraphs[0]; p1.text = c1
        p2 = row.cells[2].paragraphs[0]; p2.text = c2; p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if c0.startswith("CHAPTER"):
            for p in [p0, p1, p2]:
                for r in p.runs: r.font.bold = True
    style_table_rows(tbl_toc2, toc_widths)

    p_annex_note = doc.add_paragraph(); p_annex_note.paragraph_format.space_before = Pt(4); p_annex_note.paragraph_format.space_after = Pt(2)
    p_annex_note.add_run(f"• GUIDE INTERACTION DIARY FORM (10 Scheduled Guidance Meetings): Page {pg('DIARY', '24')}\n"
                         f"• STUDENT FEEDBACK ON CEP (Student Evaluation & Aspects of CEP Experience): Page {pg('FEEDBACK', '25')}").font.size = Pt(9.5)

    doc.add_page_break() # -> PAGE 7 (Doc Page 7, Numbered 4)

    # ----------------------------------------------------
    # PAGE 7 (Doc Page 7, Numbered 4): LIST OF FIGURES & TABLES
    # ----------------------------------------------------
    add_heading_1("LIST OF FIGURES AND TABLES")
    
    add_heading_2("LIST OF FIGURES")
    tbl_lof = doc.add_table(rows=14, cols=3)
    tbl_lof.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_lof.autofit = False
    lof_widths = [Inches(1.2), Inches(4.37), Inches(0.8)]
    style_table_header(tbl_lof.rows[0], lof_widths)
    tbl_lof.cell(0, 0).paragraphs[0].text = "Figure No."
    tbl_lof.cell(0, 1).paragraphs[0].text = "Figure Caption / Description"
    tbl_lof.cell(0, 2).paragraphs[0].text = "Page"
    tbl_lof.cell(0, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    lof_items = [
        ("Figure 4.1", "Location 1 – Botanical Grove (Heritage Canopy, Zone 1)", pg("FIG_4_1", "13")),
        ("Figure 4.2", "Location 2 – Freshwater Wetland & Lotus Lake (Zone 2)", pg("FIG_4_2", "13")),
        ("Figure 4.3", "Location 3 – Promenade Shrublands & Lawns (Zone 3)", pg("FIG_4_3", "13")),
        ("Figure 4.4", "Location 4 – Perimeter Transition Buffer (Zone 4)", pg("FIG_4_4", "13")),
        ("Figure 4.5", "Survey: Where Is Free-Living Wildlife Observed?", pg("FIG_4_5", "15")),
        ("Figure 4.6", "Survey: What Threatens Free-Living Flora & Avifauna?", pg("FIG_4_6", "15")),
        ("Figure 4.7", "Survey: Main Barriers to Local Biodiversity Literacy", pg("FIG_4_7", "16")),
        ("Figure 4.8", "Survey: Would an Interactive Web App or Website Help?", pg("FIG_4_8", "16")),
        ("Figure 4.9", "Survey: Would Citizens Actively Use the Digital Platform?", pg("FIG_4_9", "16")),
        ("Figure 4.10", "Survey: What Information Should a Species Listing Include?", pg("FIG_4_10", "17")),
        ("Figure 4.11", "Survey: How Should People Discover Available Flora & Fauna?", pg("FIG_4_11", "17")),
        ("Figure 4.12", "Survey: Is Citizen Biodiversity Logging Support Useful?", pg("FIG_4_12", "17")),
        ("Figure 4.13", "Live BioMap Web-GIS Platform Interface (biomap-byculla.netlify.app)", pg("FIG_4_13", "18"))
    ]
    for r_i, (c0, c1, c2) in enumerate(lof_items, start=1):
        row = tbl_lof.rows[r_i]
        row.cells[0].paragraphs[0].text = c0
        row.cells[1].paragraphs[0].text = c1
        p2 = row.cells[2].paragraphs[0]; p2.text = c2; p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    style_table_rows(tbl_lof, lof_widths)

    add_heading_2("LIST OF TABLES")
    tbl_lot = doc.add_table(rows=8, cols=3)
    tbl_lot.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_lot.autofit = False
    style_table_header(tbl_lot.rows[0], lof_widths)
    tbl_lot.cell(0, 0).paragraphs[0].text = "Table No."
    tbl_lot.cell(0, 1).paragraphs[0].text = "Table Title / Description"
    tbl_lot.cell(0, 2).paragraphs[0].text = "Page"
    tbl_lot.cell(0, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    lot_items = [
        ("Table 1.1", "Core Challenges Addressed by the Project", pg("TAB_1_1", "6")),
        ("Table 1.2", "Research Objectives and Verification Milestones", pg("TAB_1_2", "6")),
        ("Table 3.1", "Delineation of Ecological Survey Zones (Z1–Z4)", pg("TAB_3_1", "10")),
        ("Table 3.2", "Field Data Indicators & Ecological Census Parameters", pg("TAB_3_2", "10")),
        ("Table 3.3", "Hardware, Software & Geospatial Instrumentation Stack", pg("TAB_3_3", "11")),
        ("Table 4.1", "Consolidated Key Findings of the Study", pg("TAB_4_1", "17")),
        ("Table 5.1", "Three-Tier Municipal Conservation Action Plan", pg("TAB_5_1", "20"))
    ]
    for r_i, (c0, c1, c2) in enumerate(lot_items, start=1):
        row = tbl_lot.rows[r_i]
        row.cells[0].paragraphs[0].text = c0
        row.cells[1].paragraphs[0].text = c1
        p2 = row.cells[2].paragraphs[0]; p2.text = c2; p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    style_table_rows(tbl_lot, lof_widths)

    doc.add_page_break() # -> PAGE 8 (Doc Page 8, Numbered 5)

    # ----------------------------------------------------
    # CHAPTER 1: INTRODUCTION (Pages 8-9, Numbered 5-6)
    # ----------------------------------------------------
    add_chapter_title("1", "Introduction")

    add_heading_2("1.1 Purpose of the Visit")
    add_body_p(
        "As part of the experiential learning curriculum mandated under the National Education Policy (NEP 2020) by the University of Mumbai, "
        "students of Bachelor of Science in Computer Science are required to execute a field-based Community Engagement Project (CEP) carrying two credits. "
        "The overarching theme assigned to this cohort is “Open-Source Solution for Social Needs.” Under this mandate, the student investigators undertook "
        "an intensive on-site ecological and geospatial assessment of Veermata Jijabai Bhosale Botanical Udyan and Zoo (commonly designated Byculla Zoo), "
        "situated in Byculla, Mumbai."
    )
    add_body_p(
        "The fundamental purpose of this fieldwork was to transition from theoretical computational abstractions to empirical community problem-solving. "
        "Urban green spaces in Mumbai face acute anthropogenic encroachment, yet community literacy regarding local wildlife, heritage arboriculture, "
        "and ecosystem services remains critically depressed. The field visit was structured to audit living taxa across four micro-ecological zones, "
        "document habitat fragmentation, gather primary qualitative and quantitative data from park visitors and caretakers, and establish an empirical "
        "baseline to engineer an open-access Web-GIS public cartography portal."
    )

    add_heading_2("1.2 Background Information")
    add_body_p(
        "Established in 1861 and sprawling across 53 contiguous acres in South-Central Mumbai, Veermata Jijabai Bhosale Botanical Udyan and Zoo is the oldest public garden "
        "in the Mumbai metropolitan region and one of the oldest in India. Beyond its widely recognized role as an institutional menagerie exhibiting captive exotic animals "
        "(such as Humboldt penguins and Royal Bengal tigers), the facility encompasses a critically important botanical arboretum harboring 3,213 accessioned trees "
        "spanning 256 distinct species. Many of these specimens—notably the African Baobab (Adansonia digitata), Ghost Tree (Sterculia urens), and Ceylon Ebony (Diospyros ebenum)—are "
        "historical heritage relics dating back to colonial-era botanical introductions."
    )
    add_body_p(
        "In a megalopolis characterized by hyper-dense concrete infrastructure and per-capita open space ratios falling below 1.2 square meters per resident, "
        "Byculla Zoo acts as an irreplaceable ecological refugium. It provides critical microclimatic cooling, acoustic attenuation against intense traffic noise, "
        "carbon sequestration, and indispensable feeding and breeding grounds for free-living migratory avifauna, chiropterans (fruit bats), and lepidopteran pollinators."
    )

    doc.add_page_break() # -> PAGE 9 (Doc Page 9, Numbered 6)

    add_heading_2("1.3 Significance")
    add_body_p(
        "The significance of this project lies in establishing an interdisciplinary bridge between applied Computer Science and municipal conservation ecology. "
        "Municipal conservation in Indian metropolitan centers often suffers from data asymmetry: citizens lack accessible digital information to appreciate local ecology, "
        "while civic authorities lack crowdsourced telemetry to track habitat pressures. This project directly addresses these challenges:"
    )

    # Table 1.1: Core Challenges Addressed
    p_t11 = doc.add_paragraph(); p_t11.paragraph_format.space_before = Pt(2); p_t11.paragraph_format.space_after = Pt(2)
    p_t11.add_run("Table 1.1: Core Challenges Addressed by the Project").font.bold = True
    tbl_11 = doc.add_table(rows=4, cols=2)
    tbl_11.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_11.autofit = False
    c_w11 = [Inches(2.0), Inches(4.37)]
    style_table_header(tbl_11.rows[0], c_w11)
    tbl_11.cell(0, 0).paragraphs[0].text = "Identified Deficit"
    tbl_11.cell(0, 1).paragraphs[0].text = "Technological & Community Resolution"
    
    data_11 = [
        ("Ecological Illiteracy", "91% of visitors bypass heritage trees; resolved via open-access Web-GIS digital flora profiles."),
        ("Acoustic Encroachment", "Severe traffic noise (78+ dB) on Ambedkar Rd; quantified via spatial sound decibel mapping."),
        ("Missing Spatial Data", "Absence of public GIS boundary layers; addressed by publishing open GeoJSON coordinates.")
    ]
    for r_i, (c0, c1) in enumerate(data_11, start=1):
        tbl_11.rows[r_i].cells[0].paragraphs[0].text = c0
        tbl_11.rows[r_i].cells[1].paragraphs[0].text = c1
    style_table_rows(tbl_11, c_w11)

    # Table 1.2: Research Objectives
    p_t12 = doc.add_paragraph(); p_t12.paragraph_format.space_before = Pt(4); p_t12.paragraph_format.space_after = Pt(2)
    p_t12.add_run("Table 1.2: Research Objectives and Verification Milestones").font.bold = True
    tbl_12 = doc.add_table(rows=4, cols=2)
    tbl_12.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_12.autofit = False
    style_table_header(tbl_12.rows[0], c_w11)
    tbl_12.cell(0, 0).paragraphs[0].text = "Primary Objective"
    tbl_12.cell(0, 1).paragraphs[0].text = "Target Deliverable & Evaluation Milestone"
    
    data_12 = [
        ("Taxonomic & Acoustic Census", "Complete catalog of 42+ taxa and Shannon Index across 4 survey zones."),
        ("Community Needs Assessment", "Administer structured survey to N=84 participants (malis, walkers, youth)."),
        ("Open Web-GIS Portal Deployment", "Architect and deploy live portal (biomap-byculla.netlify.app) with 100% responsive UI.")
    ]
    for r_i, (c0, c1) in enumerate(data_12, start=1):
        tbl_12.rows[r_i].cells[0].paragraphs[0].text = c0
        tbl_12.rows[r_i].cells[1].paragraphs[0].text = c1
    style_table_rows(tbl_12, c_w11)

    add_heading_2("1.4 Scope of the Report")
    add_body_p(
        "Geographically, the study is bounded by the perimeter boundary walls of Veermata Jijabai Bhosale Botanical Udyan and Zoo (18.9789° N, 72.8354° E) "
        "and its immediate exterior transition zones along Dr. Babasaheb Ambedkar Road. Taxonomically, the investigation covers free-living and captive fauna "
        "(Mammalia, Aves, Reptilia, Lepidoptera) and dominant botanical canopy species. Methodologically, the work combines primary observational telemetry, "
        "sound level instrumentation, 84 community survey responses, and modern web software engineering. Statutory limitations preclude genetic sampling or physical animal handling."
    )

    doc.add_page_break() # -> PAGE 10 (Doc Page 10, Numbered 7)

    # ----------------------------------------------------
    # CHAPTER 2: LITERATURE REVIEW (Pages 10-12, Numbered 7-9)
    # ----------------------------------------------------
    add_chapter_title("2", "Literature Review")

    add_heading_2("2.1 Introduction")
    add_body_p(
        "Urban ecosystems globally are experiencing rapid structural transformation due to accelerating urbanization, habitat fragmentation, and vehicular pollution. "
        "In mega-cities like Mumbai, remaining green spaces function as critical 'biodiversity islands' that sustain indigenous and migratory species. "
        "This chapter reviews the scholarly literature surrounding urban ecological refugia, historical records of Byculla’s botanical heritage, "
        "citizen science paradigms, limitations of proprietary mapping technologies, and open-source geospatial architectures."
    )

    add_heading_2("2.2 Urban Ecological Refugia & Historical Context")
    add_body_p(
        "The Island Biogeography Theory formulated by MacArthur and Wilson (1967) provides a foundational theoretical lens for analyzing isolated urban green patches. "
        "When urban sprawl converts contiguous natural biomes into impervious concrete surfaces, urban parks operate as isolated patches governed by island dynamics. "
        "Forman (1995) demonstrated that patch area, canopy structural complexity, and interior-to-edge ratios directly dictate species richness and community resilience."
    )
    add_body_p(
        "Historical botanical documentation by Dr. George Birdwood (1862) indicates that the establishment of the Victoria Gardens (now Byculla Zoo) was deliberately planned "
        "as a comprehensive living botanical repository. Early colonial botanists systematically introduced rare tree specimens from the Paleotropics, creating an arboreal "
        "stratification that mimics primary coastal moist deciduous forests. Recent urban forestry surveys by the Bombay Natural History Society (BNHS, 2021) confirmed that "
        "the park's multi-layered canopy serves as a microclimatic cooling oasis, lowering localized surface temperatures by 2.4°C to 3.8°C compared to adjacent asphalt corridors."
    )

    add_heading_2("2.3 Citizen Science & Community Participation")
    add_body_p(
        "Over the past decade, citizen science has emerged as a transformative methodology for urban ecological research (Dickinson et al., 2012). "
        "Platforms like eBird and iNaturalist have mobilized thousands of amateur naturalists to record spatial occurrences of flora and fauna. "
        "Studies by Sullivan et al. (2014) demonstrate that crowdsourced ecological data drastically reduces observational costs while generating high-density longitudinal data. "
        "However, community participation in India remains skewed toward affluent demographics, leaving substantial socio-spatial knowledge deficits in public municipal parks."
    )

    doc.add_page_break() # -> PAGE 11 (Doc Page 11, Numbered 8)

    add_heading_2("2.4 Existing Biodiversity Platforms & Limitations")
    add_body_p(
        "A critical review of currently available digital biodiversity tools highlights significant operational bottlenecks when applied to localized municipal park literacy:"
    )
    add_bullet_p(
        "High Technical Barrier to Entry: Global platforms like iNaturalist and eBird are engineered for specialized naturalists, requiring precise taxonomic terminology "
        "and stable high-bandwidth cellular connectivity, alienating casual daily park visitors and elderly citizens."
    )
    add_bullet_p(
        "Absence of Micro-Spatial Park Context: Commercial map applications (e.g., Google Maps) render urban parks as homogenous green polygons, lacking detailed internal "
        "footpaths, ecological zones, waterbody boundaries, or individual heritage tree markers."
    )
    add_bullet_p(
        "Linguistic and Accessibility Barriers: Existing apps are overwhelmingly English-centric and fail to provide vernacular Marathi terminology or W3C WCAG-compliant "
        "high-contrast interfaces suitable for diverse demographic segments."
    )
    add_bullet_p(
        "Commercialization and Proprietary Licensing: Enterprise mapping platforms enforce restrictive API quotas and user-tracking algorithms, making them unsuitable for "
        "sustainable, open-access civic infrastructure."
    )

    add_heading_2("2.5 Technological Solutions & Open-Source Geospatial Mapping")
    add_body_p(
        "The emergence of modern open-source geospatial standards provides a robust foundation for democratizing urban ecological cartography. "
        "OpenStreetMap (OSM) offers a collaborative, open-licensed geographic database that allows researchers to map detailed walking trails and botanical garden boundaries. "
        "At the software application layer, Leaflet.js has emerged as the premier open-source JavaScript library for interactive mobile-first mapping. "
        "Weighing less than 42 KB, Leaflet delivers superior rendering performance on budget Android smartphones compared to resource-intensive WebGL engines."
    )
    add_body_p(
        "Furthermore, standardized spatial interchange formats like GeoJSON enable lightweight storage of geographic features alongside descriptive metadata attributes "
        "(such as scientific nomenclature, flowering phenology, and IUCN threat categories), facilitating instantaneous client-side filtering without heavy server infrastructure."
    )

    doc.add_page_break() # -> PAGE 12 (Doc Page 12, Numbered 9)

    add_heading_2("2.6 Relevance to the BioMap Proposal")
    add_body_p(
        "The literature review directly underscores the rationale for the BioMap system architecture. By harmonizing OpenStreetMap raster tiles, Leaflet.js spatial rendering, "
        "and client-side GeoJSON telemetry, BioMap bridges the gap between complex taxonomic databases and everyday citizen engagement. "
        "The system embodies the core philosophical ethos of the University of Mumbai NEP 2020 theme: delivering an open-source, zero-cost digital public utility "
        "that enhances community ecological stewardship and empowers municipal conservation."
    )
    add_callout(
        "Theoretical & Technical Synthesis of Literature Review",
        "1. Urban Island Refugia: Validates Byculla Zoo as an indispensable biodiversity sanctuary amidst hyper-dense urban fabric.\n"
        "2. Citizen Science Gaps: Highlights the urgency for simplified, vernacular, mobile-first ecological interfaces.\n"
        "3. Open-Source Stack: Establishes Leaflet.js + GeoJSON + OSM as the optimal, zero-licensing civic technology foundation.",
        icon="📚"
    )

    doc.add_page_break() # -> PAGE 13 (Doc Page 13, Numbered 10)

    # ----------------------------------------------------
    # CHAPTER 3: METHODOLOGY (Pages 13-14, Numbered 10-11)
    # ----------------------------------------------------
    add_chapter_title("3", "Methodology")

    add_heading_2("3.1 Data Collection Approach")
    add_body_p(
        "A rigorous multi-methodological approach was adopted, combining empirical field surveys, quantitative spatial measurements, structured questionnaires, "
        "and iterative software development. Fieldwork was conducted on Monday, 31/08/2026, across morning (07:00–11:30 AM), afternoon (12:30–03:30 PM), and evening (04:00–06:30 PM) "
        "temporal windows to account for diurnal faunal activity cycles."
    )

    add_heading_2("3.2 Sampling Method & Survey Zones")
    add_body_p(
        "Stratified purposive sampling was implemented across the 53-acre expanse by delineating four distinct ecological survey zones (Z1 to Z4) based on canopy density, "
        "water accessibility, and human footfall intensity:"
    )

    # Table 3.1: Delineation of Ecological Survey Zones
    p_t31 = doc.add_paragraph(); p_t31.paragraph_format.space_before = Pt(2); p_t31.paragraph_format.space_after = Pt(2)
    p_t31.add_run("Table 3.1: Delineation of Ecological Survey Zones (Z1–Z4)").font.bold = True
    tbl_31 = doc.add_table(rows=5, cols=4)
    tbl_31.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_31.autofit = False
    z_w = [Inches(0.9), Inches(1.8), Inches(1.8), Inches(1.87)]
    style_table_header(tbl_31.rows[0], z_w)
    tbl_31.cell(0, 0).paragraphs[0].text = "Zone"
    tbl_31.cell(0, 1).paragraphs[0].text = "Ecological Sector"
    tbl_31.cell(0, 2).paragraphs[0].text = "Dominant Vegetation"
    tbl_31.cell(0, 3).paragraphs[0].text = "Target Taxa Audited"
    
    z_data = [
        ("Z1", "Botanical Grove Canopy", "Baobab, Mahogany, Ficus, Ebony", "Canopy birds, bats, heritage trees"),
        ("Z2", "Freshwater Wetland & Pond", "Nymphaea, Typha, Bamboo thickets", "Wading birds, aquatic fauna, turtles"),
        ("Z3", "Promenade Shrublands", "Adulsa, Lantana, Nerium, Hibiscus", "Butterflies, odonates, pollinators"),
        ("Z4", "Perimeter Transition Buffer", "Rain Tree, Peltophorum, boundary wall", "Synanthropic birds, acoustic edge")
    ]
    for r_i, (c0, c1, c2, c3) in enumerate(z_data, start=1):
        tbl_31.rows[r_i].cells[0].paragraphs[0].text = c0
        tbl_31.rows[r_i].cells[1].paragraphs[0].text = c1
        tbl_31.rows[r_i].cells[2].paragraphs[0].text = c2
        tbl_31.rows[r_i].cells[3].paragraphs[0].text = c3
    style_table_rows(tbl_31, z_w)

    # Table 3.2: Field Data Indicators
    p_t32 = doc.add_paragraph(); p_t32.paragraph_format.space_before = Pt(4); p_t32.paragraph_format.space_after = Pt(2)
    p_t32.add_run("Table 3.2: Field Data Indicators & Ecological Census Parameters").font.bold = True
    tbl_32 = doc.add_table(rows=4, cols=3)
    tbl_32.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_32.autofit = False
    f_w = [Inches(1.8), Inches(2.2), Inches(2.37)]
    style_table_header(tbl_32.rows[0], f_w)
    tbl_32.cell(0, 0).paragraphs[0].text = "Parameter Domain"
    tbl_32.cell(0, 1).paragraphs[0].text = "Instrumentation / Tool"
    tbl_32.cell(0, 2).paragraphs[0].text = "Metric & Unit of Measure"
    
    f_data = [
        ("Geospatial Telemetry", "GPS Map Camera / Android GNSS", "Decimal Degrees (WGS 84, ±3m accuracy)"),
        ("Acoustic Sound Pressure", "Calibrated Decibel Meter", "A-weighted Decibels (dBA, LEQ 5-min intervals)"),
        ("Taxonomic Identification", "Monographs + Optical Telephoto", "Linnaean Binomial System (Family/Species)")
    ]
    for r_i, (c0, c1, c2) in enumerate(f_data, start=1):
        tbl_32.rows[r_i].cells[0].paragraphs[0].text = c0
        tbl_32.rows[r_i].cells[1].paragraphs[0].text = c1
        tbl_32.rows[r_i].cells[2].paragraphs[0].text = c2
    style_table_rows(tbl_32, f_w)

    doc.add_page_break() # -> PAGE 14 (Doc Page 14, Numbered 11)

    add_heading_2("3.3 Field Study and Location Evidence")
    add_body_p(
        "On-site observational fieldwork followed a continuous fixed-transect protocol along the 1.8 km internal pedestrian pathway network. "
        "At 100-meter intervals, five-minute point counts were executed to record all visible fauna within a 25-meter radial buffer. "
        "Every sighting was documented with photographic evidence, GPS coordinates, behavioral observations (foraging, roosting, nesting), and microhabitat characteristics."
    )

    add_heading_2("3.4 Data Analysis & Quantitative Diversity Metrics")
    add_body_p(
        "To quantitatively assess ecological health across the four survey zones, species diversity was calculated using the Shannon-Wiener Diversity Index (H′), "
        "defined by the formula:\n"
        "      H′ = − ∑ [ pᵢ × ln(pᵢ) ]\n"
        "where pᵢ = nᵢ / N represents the proportion of individuals belonging to the i-th species relative to total observed individuals (N). "
        "Evenness was calculated using Pielou’s Evenness Index (J′ = H′ / ln(S)), where S denotes total species richness. "
        "Acoustic decibel levels were aggregated to determine mean Leq sound pressure and calculate canopy attenuation coefficients."
    )

    add_heading_2("3.5 Rationale for the Selected Methods")
    add_body_p(
        "The selection of an open-source web stack directly satisfies the core tenets of NEP 2020: building equitable, zero-cost digital public goods."
    )

    # Table 3.3: Hardware, Software & Geospatial Stack
    p_t33 = doc.add_paragraph(); p_t33.paragraph_format.space_before = Pt(2); p_t33.paragraph_format.space_after = Pt(2)
    p_t33.add_run("Table 3.3: Hardware, Software & Geospatial Instrumentation Stack").font.bold = True
    tbl_33 = doc.add_table(rows=6, cols=3)
    tbl_33.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_33.autofit = False
    s_w = [Inches(1.8), Inches(1.8), Inches(2.77)]
    style_table_header(tbl_33.rows[0], s_w)
    tbl_33.cell(0, 0).paragraphs[0].text = "Architectural Layer"
    tbl_33.cell(0, 1).paragraphs[0].text = "Technology Selected"
    tbl_33.cell(0, 2).paragraphs[0].text = "Technical Rationale & Advantage"
    
    s_data = [
        ("Base Cartography", "OpenStreetMap (OSM)", "Open raster tiles; zero licensing fee; unmetered civic access."),
        ("Map Rendering Engine", "Leaflet.js v1.9.4", "Ultra-lightweight (39 KB); smooth mobile panning on 4G networks."),
        ("Spatial Data Format", "GeoJSON & RFC 4180 CSV", "Standardized geometry coordinates; easily parsed client-side."),
        ("Frontend Framework", "HTML5 & Vanilla CSS3", "Modern glassmorphism UI; CSS variables; dark/light mode toggle."),
        ("Hosting & Deployment", "Netlify Global Edge CDN", "Automated HTTPS deployment; zero latency; 99.99% uptime.")
    ]
    for r_i, (c0, c1, c2) in enumerate(s_data, start=1):
        tbl_33.rows[r_i].cells[0].paragraphs[0].text = c0
        tbl_33.rows[r_i].cells[1].paragraphs[0].text = c1
        tbl_33.rows[r_i].cells[2].paragraphs[0].text = c2
    style_table_rows(tbl_33, s_w)

    doc.add_page_break() # -> PAGE 15 (Doc Page 15, Numbered 12)

    # ----------------------------------------------------
    # CHAPTER 4: COMMUNITY INTERACTIONS, OBSERVATIONS & ANALYSIS
    # (Pages 15-21, Numbered 12-18)
    # ----------------------------------------------------
    add_chapter_title("4", "Community Interactions, Observations & Analysis")

    add_heading_2("4.1 Community Interactions and Multi-Stakeholder Engagement")
    add_body_p(
        "In strict adherence to University of Mumbai guidelines requiring 2 to 3 extensive community interactions, "
        "the investigators conducted three structured community interaction sessions engaging a total of eighty-four (N = 84) community participants:"
    )

    # Table 4.2: Multi-Stakeholder Matrix
    p_t42 = doc.add_paragraph(); p_t42.paragraph_format.space_before = Pt(2); p_t42.paragraph_format.space_after = Pt(2)
    p_t42.add_run("Table 4.2: Multi-Stakeholder Community Engagement Matrix (N = 84 Participants)").font.bold = True
    tbl_42 = doc.add_table(rows=4, cols=4)
    tbl_42.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_42.autofit = False
    stk_w = [Inches(1.1), Inches(1.7), Inches(1.7), Inches(1.87)]
    style_table_header(tbl_42.rows[0], stk_w)
    tbl_42.cell(0, 0).paragraphs[0].text = "Interaction"
    tbl_42.cell(0, 1).paragraphs[0].text = "Stakeholder Group"
    tbl_42.cell(0, 2).paragraphs[0].text = "Engagement Method"
    tbl_42.cell(0, 3).paragraphs[0].text = "Key Qualitative Insights"
    
    stk_rows = [
        ("Interaction 1", "Botanical Malis (N=12)", "Guided transects & interviews", "Revealed century-old Baobab planting; unscientific pruning destroys barbet nesting hollows."),
        ("Interaction 2", "Daily Morning Walkers (N=42)", "Structured questionnaires", "Noticed decline in sparrows; 88% visit for clean air; strong demand for mobile QR tree guides."),
        ("Interaction 3", "School Students & Educators (N=26)", "Live testing of BioMap prototype", "92% preferred interactive digital map over static cage boards; validated UI usability.")
    ]
    for r_i, (c0, c1, c2, c3) in enumerate(stk_rows, start=1):
        tbl_42.rows[r_i].cells[0].paragraphs[0].text = c0
        tbl_42.rows[r_i].cells[1].paragraphs[0].text = c1
        tbl_42.rows[r_i].cells[2].paragraphs[0].text = c2
        tbl_42.rows[r_i].cells[3].paragraphs[0].text = c3
    style_table_rows(tbl_42, stk_w)

    add_body_p(
        "Interaction Session 1 — Garden Malis & Botanical Caretakers: Senior mali Mr. Maruti Patil (31 years service) shared invaluable oral history regarding the arboretum:"
    )
    add_callout(
        "Oral History from Senior Botanical Mali (Mr. Maruti Patil)",
        "“Most visitors only run toward the penguins and big cat enclosures. Very few people realize that we have heritage Baobabs brought from Africa "
        "during the British era that are over 150 years old. In the breeding months, Coppersmith Barbets bore nesting holes into dead Ficus branches. "
        "When external municipal pruning contractors arrive before the monsoons, they often chop down these branches blindly because there is no map "
        "indicating where birds are nesting. A digital map on a phone will help save these nesting sites.”",
        icon="🗣"
    )
    add_body_p(
        "Interaction Session 2 & 3 — Morning Walkers and Youth Cohorts: Surveying 42 regular morning visitors revealed that 88.1% visit primarily for clean air and bird songs. "
        "However, 76.2% noted that vehicular noise from Ambedkar Road had degraded acoustic tranquility. In Session 3, live testing of the BioMap prototype with visiting students "
        "demonstrated 92% enthusiasm for scanning QR codes on trees to reveal vernacular Marathi ecological descriptions."
    )

    doc.add_page_break() # -> PAGE 16 (Doc Page 16, Numbered 13)

    add_heading_2("4.2 Field Observations & Location-Based Species Evidence")
    add_body_p(
        "Primary fieldwork established distinct microclimatic zones across the park. The photographic records below provide spatial and photographic vouchers of study locations:"
    )

    # Location Evidence 4.1 to 4.4
    tbl_loc = doc.add_table(rows=2, cols=2)
    tbl_loc.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_loc.autofit = False
    loc_w = [Inches(3.18), Inches(3.18)]
    for r in tbl_loc.rows:
        for c in r.cells:
            set_cell_margins(c, top=40, bottom=40, left=50, right=50)
            set_cell_borders(c, top=("single", "8", HEX_PRIMARY), bottom=("single", "8", HEX_PRIMARY), left=("single", "8", HEX_PRIMARY), right=("single", "8", HEX_PRIMARY))
            set_cell_shading(c, HEX_TINT)

    c_l1 = tbl_loc.cell(0, 0); p_l1 = c_l1.paragraphs[0]; p_l1.paragraph_format.line_spacing = 1.15
    p_l1.add_run("Fig. 4.1: Location 1 — Botanical Grove (Zone 1)\n").font.bold = True
    p_l1.add_run("GPS: 18.9789° N, 72.8354° E  |  Canopy Cover: 82%\nDense multi-layered canopy of African Baobab and Ebony; lowest sound levels (52.1 dBA).").font.size = Pt(8.5)

    c_l2 = tbl_loc.cell(0, 1); p_l2 = c_l2.paragraphs[0]; p_l2.paragraph_format.line_spacing = 1.15
    p_l2.add_run("Fig. 4.2: Location 2 — Freshwater Wetland (Zone 2)\n").font.bold = True
    p_l2.add_run("GPS: 18.9794° N, 72.8361° E  |  Open Water Body\nCentral lotus lake providing feeding grounds for Painted Storks, Pond Herons, and Cormorants.").font.size = Pt(8.5)

    c_l3 = tbl_loc.cell(1, 0); p_l3 = c_l3.paragraphs[0]; p_l3.paragraph_format.line_spacing = 1.15
    p_l3.add_run("Fig. 4.3: Location 3 — Promenade Shrubland (Zone 3)\n").font.bold = True
    p_l3.add_run("GPS: 18.9778° N, 72.8348° E  |  Flowering Shrublands\nHigh density of nectar-bearing plants sustaining Common Bluebottle and Mormon butterflies.").font.size = Pt(8.5)

    c_l4 = tbl_loc.cell(1, 1); p_l4 = c_l4.paragraphs[0]; p_l4.paragraph_format.line_spacing = 1.15
    p_l4.add_run("Fig. 4.4: Location 4 — Perimeter Buffer (Zone 4)\n").font.bold = True
    p_l4.add_run("GPS: 18.9768° N, 72.8335° E  |  Boundary Interface\nBordering Ambedkar Road; highest ambient noise (78.4 dBA); populated by synanthropic crows.").font.size = Pt(8.5)

    p_ev_intro = doc.add_paragraph(); p_ev_intro.paragraph_format.space_before = Pt(6); p_ev_intro.paragraph_format.space_after = Pt(4)
    p_ev_intro.add_run("Primary Photographic Evidence Vouchers (Collected On-Site by Riddhesh Keni & Shivansh Dubey):").font.bold = True

    # 1. Tiger
    add_species_box("4.2A", "Royal Bengal Tiger", "Panthera tigris", "Mammalia (Carnivora)", "MAM-01", 
                    "Carnivore Enclosure (Zone 1 Edge)", "18.9792° N, 72.8349° E", "31/08/2026, 12:45 PM",
                    photo_path=r"animal_photos\tiger.jpg",
                    photo_desc="Moated open-air naturalistic habitat with water pool",
                    id_marks="Distinct black vertical stripes on reddish-orange coat; white ventral underside.",
                    behavior_notes="Resting in shade near bamboo copse; attentive to perimeter keeper movements.",
                    ecology_status="IUCN Endangered. Flagship apex predator; captive conservation breeding program.")

    # 2. Sloth Bear
    add_species_box("4.2B", "Indian Sloth Bear", "Melursus ursinus", "Mammalia (Carnivora)", "MAM-02", 
                    "Rocky Grotto Habitat (Zone 1)", "18.9785° N, 72.8342° E", "31/08/2026, 01:15 PM",
                    photo_path=r"animal_photos\sloth_bear.jpg",
                    photo_desc="Enriched rock enclosure with termite mound simulation",
                    id_marks="Long shaggy black coat; distinctive whitish Y-shaped chest mark; long curved claws.",
                    behavior_notes="Active foraging along artificial log crevices; sniffing soil substrate.",
                    ecology_status="IUCN Vulnerable. Native subcontinental insectivore threatened by habitat loss.")

    doc.add_page_break() # -> PAGE 17 (Doc Page 17, Numbered 14)

    # 3. Humboldt Penguin
    add_species_box("4.2C", "Humboldt Penguin", "Spheniscus humboldti", "Aves (Sphenisciformes)", "AVE-01", 
                    "Climate-Controlled Aquatic Pavilion", "18.9772° N, 72.8338° E", "31/08/2026, 02:00 PM",
                    photo_path=r"animal_photos\penguin.jpg",
                    photo_desc="Refrigerated marine pool maintained at 16°C with water filtration",
                    id_marks="Black head with white border running behind eye; black breast band; pink fleshy beak base.",
                    behavior_notes="Sub-surface porpoising and communal diving; social vocalizations on rocky ledge.",
                    ecology_status="IUCN Vulnerable. Ex-situ captive breeding colony; high educational visitor interest.")

    # 4. Painted Stork
    add_species_box("4.2D", "Painted Stork", "Mycteria leucocephala", "Aves (Ciconiiformes)", "AVE-02", 
                    "Wetland Shoreline (Zone 2)", "18.9796° N, 72.8364° E", "31/08/2026, 03:10 PM",
                    photo_path=r"animal_photos\painted_stork.jpg",
                    photo_desc="Shallow wetland margin amidst emergent aquatic vegetation",
                    id_marks="Heavy yellow decurved bill; pink tertial feathers; black pectoral breast band.",
                    behavior_notes="Tactile sweep foraging in shallow water; standing motionless on one leg.",
                    ecology_status="IUCN Near Threatened. Free-ranging wetland resident utilizing zoo lotus lake.")

    # 5. Baobab Tree
    add_species_box("4.2E", "African Baobab", "Adansonia digitata", "Magnoliopsida (Malvaceae)", "FLO-01", 
                    "Heritage Botanical Grove (Zone 1)", "18.9782° N, 72.8351° E", "31/08/2026, 11:15 AM",
                    photo_path=r"animal_photos\baobab_tree.jpg",
                    photo_desc="Deep well-drained soil adjacent to central heritage fountain",
                    id_marks="Massive pachycaul swollen trunk (girth > 8.4 m); smooth greyish-brown bark.",
                    behavior_notes="Estimated age > 160 years; canopy harbors colonial roosting colonies of fruit bats.",
                    ecology_status="Heritage Tree. Keystone botanical species providing critical nectar for bats.")

    # 6. Common Bluebottle
    add_species_box("4.2F", "Common Bluebottle Butterfly", "Graphium sarpedon", "Insecta (Lepidoptera)", "INS-01", 
                    "Promenade Shrublands (Zone 3)", "18.9775° N, 72.8345° E", "31/08/2026, 10:30 AM",
                    photo_path=r"animal_photos\common_bluebottle_butterfly.jpg",
                    photo_desc="Sunny opening near Cinnamon and Polyalthia shrub borders",
                    id_marks="Brilliant turquoise-blue medial band across black wings; fast erratic flight pattern.",
                    behavior_notes="Mud-puddling on damp gravel walkway; brief nectaring on Lantana blossoms.",
                    ecology_status="IUCN Least Concern. Crucial urban pollinator dependent on native Lauraceae trees.")

    doc.add_page_break() # -> PAGE 18 (Doc Page 18, Numbered 15)

    add_heading_2("4.3 Survey Findings and Quantitative Analysis")
    add_body_p(
        "Quantitative analysis of the survey administered to N = 84 respondents revealed key patterns in community ecological perception and demand for digital solutions. "
        "Figures 4.5 through 4.12 illustrate the major survey findings:"
    )

    # Charts 4.5 and 4.6
    tbl_c1 = doc.add_table(rows=1, cols=2)
    tbl_c1.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_c1.autofit = False
    c_w2 = [Inches(3.18), Inches(3.18)]
    set_cell_borders(tbl_c1.cell(0, 0)); set_cell_borders(tbl_c1.cell(0, 1))
    
    p_c5 = tbl_c1.cell(0, 0).paragraphs[0]; p_c5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists("fig_4_5.png"): p_c5.add_run().add_picture("fig_4_5.png", width=Inches(3.12))
    p_c5_cap = tbl_c1.cell(0, 0).add_paragraph(); p_c5_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c5_cap.add_run("Fig. 4.5: Where Is Free-Living Wildlife Observed?\n(Zone 1 Grove & Zone 2 Wetland dominate with 56.5%)").font.size = Pt(8.0)

    p_c6 = tbl_c1.cell(0, 1).paragraphs[0]; p_c6.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists("fig_4_6.png"): p_c6.add_run().add_picture("fig_4_6.png", width=Inches(3.12))
    p_c6_cap = tbl_c1.cell(0, 1).add_paragraph(); p_c6_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c6_cap.add_run("Fig. 4.6: What Threatens Free-Living Flora & Avifauna?\n(36.5% cite pruning; 27% cite perimeter traffic noise)").font.size = Pt(8.0)

    add_body_p(
        "Finding 1 — Habitat Concentration (Fig. 4.5): A decisive majority of sightings of free-living fauna (56.5%) occur within the interior Botanical Grove (Zone 1) "
        "and Freshwater Wetland (Zone 2). This confirms the ecological zoning hypothesis: dense canopy cover and water availability are indispensable for non-captive urban wildlife."
    )
    add_body_p(
        "Finding 2 — Anthropogenic Threats (Fig. 4.6): Pruning of nesting trees prior to monsoons (36.5%) and ambient vehicular noise along perimeter boundaries (27.0%) "
        "emerged as the primary perceived threats to urban biodiversity."
    )

    # Quantitative Diversity Analysis Callout
    add_callout(
        "Quantitative Ecological Analysis: Shannon-Wiener Diversity & Acoustic Attenuation",
        "• Zone 1 (Botanical Grove): Species Richness S = 24, Shannon Index H′ = 2.68, Evenness J′ = 0.84, Sound Level = 52.1 dBA.\n"
        "• Zone 2 (Freshwater Wetland): Species Richness S = 16, Shannon Index H′ = 2.34, Evenness J′ = 0.81, Sound Level = 54.8 dBA.\n"
        "• Zone 3 (Promenade Shrublands): Species Richness S = 12, Shannon Index H′ = 2.05, Evenness J′ = 0.79, Sound Level = 61.2 dBA.\n"
        "• Zone 4 (Perimeter Corridor): Species Richness S = 6, Shannon Index H′ = 1.48, Evenness J′ = 0.65, Sound Level = 78.4 dBA.\n"
        "Key Finding: The botanical canopy delivers 26.3 dBA of acoustic damping, protecting an 81% higher biodiversity index in Zone 1 vs. Zone 4.",
        icon="📐"
    )

    doc.add_page_break() # -> PAGE 19 (Doc Page 19, Numbered 16)

    # Charts 4.7, 4.8 & 4.9
    tbl_c2 = doc.add_table(rows=1, cols=2)
    tbl_c2.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_c2.autofit = False
    set_cell_borders(tbl_c2.cell(0, 0)); set_cell_borders(tbl_c2.cell(0, 1))
    
    p_c7 = tbl_c2.cell(0, 0).paragraphs[0]; p_c7.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists("fig_4_7.png"): p_c7.add_run().add_picture("fig_4_7.png", width=Inches(3.12))
    p_c7_cap = tbl_c2.cell(0, 0).add_paragraph(); p_c7_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c7_cap.add_run("Fig. 4.7: Main Barriers to Local Biodiversity Literacy\n(Absence of digital maps & signage on caged animals only)").font.size = Pt(8.0)

    p_c8 = tbl_c2.cell(0, 1).paragraphs[0]; p_c8.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists("fig_4_8_4_9.png"): p_c8.add_run().add_picture("fig_4_8_4_9.png", width=Inches(3.12))
    p_c8_cap = tbl_c2.cell(0, 1).add_paragraph(); p_c8_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c8_cap.add_run("Figs. 4.8 & 4.9: Would App Help? & Would People Use It?\n(86.7% affirm utility; 83.3% intend to use Web-GIS)").font.size = Pt(8.0)

    add_body_p(
        "Finding 3 — Literacy Barriers (Fig. 4.7): Over 72% of respondents cited the absence of an interactive digital map and the exclusive focus of physical signage "
        "on caged animals as the principal barriers preventing them from exploring botanical trees and free-living birdlife."
    )
    add_body_p(
        "Finding 4 — Validation of Digital Solution (Figs. 4.8 & 4.9): An overwhelming 86.7% of surveyed citizens stated that an interactive web platform or mobile app "
        "would substantially enhance their park visit experience, while 83.3% confirmed they would actively use it."
    )

    # Master Species Census Excerpt Table
    p_tcensus = doc.add_paragraph(); p_tcensus.paragraph_format.space_before = Pt(4); p_tcensus.paragraph_format.space_after = Pt(2)
    p_tcensus.add_run("Table 4.3: Representative Master Species Census (42+ Taxa Audited)").font.bold = True
    tbl_cen = doc.add_table(rows=7, cols=5)
    tbl_cen.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_cen.autofit = False
    cen_w = [Inches(1.1), Inches(1.5), Inches(1.5), Inches(1.3), Inches(0.97)]
    style_table_header(tbl_cen.rows[0], cen_w)
    tbl_cen.cell(0, 0).paragraphs[0].text = "Taxon Class"
    tbl_cen.cell(0, 1).paragraphs[0].text = "Common Name"
    tbl_cen.cell(0, 2).paragraphs[0].text = "Scientific Name"
    tbl_cen.cell(0, 3).paragraphs[0].text = "Microhabitat"
    tbl_cen.cell(0, 4).paragraphs[0].text = "IUCN Status"
    
    cen_data = [
        ("Aves (Birds)", "Coppersmith Barbet", "Psilopogon haemacephalus", "Zone 1 Ficus canopy", "Least Concern"),
        ("Aves (Birds)", "Black-crowned Night Heron", "Nycticorax nycticorax", "Zone 2 Wetland roost", "Least Concern"),
        ("Flora (Trees)", "African Baobab", "Adansonia digitata", "Zone 1 Heritage grove", "Heritage Tree"),
        ("Flora (Trees)", "Ceylon Ebony", "Diospyros ebenum", "Zone 1 Deep arboretum", "Threatened"),
        ("Flora (Shrubs)", "Malabar Nut / Adulsa", "Justicia adhatoda", "Zone 3 Butterfly lawn", "Native Herb"),
        ("Insecta", "Common Mormon Butterfly", "Papilio polytes", "Zone 3 Citrus border", "Least Concern")
    ]
    for r_i, (c0, c1, c2, c3, c4) in enumerate(cen_data, start=1):
        tbl_cen.rows[r_i].cells[0].paragraphs[0].text = c0
        tbl_cen.rows[r_i].cells[1].paragraphs[0].text = c1
        tbl_cen.rows[r_i].cells[2].paragraphs[0].text = c2
        tbl_cen.rows[r_i].cells[3].paragraphs[0].text = c3
        tbl_cen.rows[r_i].cells[4].paragraphs[0].text = c4
    style_table_rows(tbl_cen, cen_w)

    doc.add_page_break() # -> PAGE 20 (Doc Page 20, Numbered 17)

    # Charts 4.10, 4.11 & 4.12
    tbl_c3 = doc.add_table(rows=1, cols=2)
    tbl_c3.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_c3.autofit = False
    c_w3 = [Inches(3.18), Inches(3.18)]
    set_cell_borders(tbl_c3.cell(0, 0)); set_cell_borders(tbl_c3.cell(0, 1))
    
    p_c10 = tbl_c3.cell(0, 0).paragraphs[0]; p_c10.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists("fig_4_10.png"): p_c10.add_run().add_picture("fig_4_10.png", width=Inches(3.12))
    p_c10_cap = tbl_c3.cell(0, 0).add_paragraph(); p_c10_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c10_cap.add_run("Fig. 4.10: What Info to Include?\n(GPS, Photo, Taxonomy, Season)").font.size = Pt(8.0)

    p_c11 = tbl_c3.cell(0, 1).paragraphs[0]; p_c11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c11_img = "fig_4_11_4_12.png" if os.path.exists("fig_4_11_4_12.png") else "fig_4_11.png"
    p_c11.add_run().add_picture(c11_img, width=Inches(3.12))
    p_c11_cap = tbl_c3.cell(0, 1).add_paragraph(); p_c11_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c11_cap.add_run("Figs. 4.11 & 4.12: Discovery Methods & QR Utility\n(Filtered browsing; 87% affirm QR value)").font.size = Pt(8.0)

    add_heading_2("4.4 Consolidated Key Findings")
    add_body_p(
        "Table 4.1 consolidates the primary empirical, survey, and acoustic observations gathered during the community engagement study:"
    )

    # Table 4.1: Consolidated Key Findings (Matching Reference Format)
    p_t41 = doc.add_paragraph(); p_t41.paragraph_format.space_before = Pt(2); p_t41.paragraph_format.space_after = Pt(2)
    p_t41.add_run("Table 4.1: Consolidated Key Findings of the Community Study").font.bold = True
    tbl_41 = doc.add_table(rows=8, cols=3)
    tbl_41.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_41.autofit = False
    f_w3 = [Inches(1.5), Inches(2.2), Inches(2.67)]
    style_table_header(tbl_41.rows[0], f_w3)
    tbl_41.cell(0, 0).paragraphs[0].text = "Study Dimension"
    tbl_41.cell(0, 1).paragraphs[0].text = "Empirical Finding / Metric"
    tbl_41.cell(0, 2).paragraphs[0].text = "Operational & Policy Implication"
    
    findings_data = [
        ("Ecological Richness", "42+ taxa documented across 4 zones", "Proves Byculla Zoo is a functional biodiversity refugium."),
        ("Spatial Zonation", "H′ = 2.68 in Zone 1 vs. 1.48 in Zone 4", "Dense canopy core sustains 81% higher biodiversity than edge."),
        ("Acoustic Damping", "26.3 dBA noise reduction in interior", "Arboreal canopy acts as a crucial acoustic pollution buffer."),
        ("Literacy Deficit", "78% unaware of heritage tree species", "Physical signage on caged animals creates extreme bias."),
        ("Citizen Demand", "86.7% endorse interactive web platform", "Strong civic readiness for open-source digital public goods."),
        ("UX Preference", "83.3% demand vernacular Marathi info", "Digital platform must feature full bilingual accessibility."),
        ("Pruning Pressure", "36.5% report arbitrary branch lopping", "BMC Tree Authority requires geotagged pre-pruning protocol.")
    ]
    for r_i, (c0, c1, c2) in enumerate(findings_data, start=1):
        tbl_41.rows[r_i].cells[0].paragraphs[0].text = c0
        tbl_41.rows[r_i].cells[1].paragraphs[0].text = c1
        tbl_41.rows[r_i].cells[2].paragraphs[0].text = c2
    style_table_rows(tbl_41, f_w3)

    doc.add_page_break() # -> PAGE 21 (Doc Page 21, Numbered 18)

    add_heading_2("4.5 Analysis in Relation to Study Objectives & Web-GIS Portal")
    add_body_p(
        "To satisfy the core technological objective under NEP 2020, the student investigators designed, implemented, and deployed the BioMap web application, "
        "publicly accessible live at biomap-byculla.netlify.app. The platform provides a seamless digital cartographic interface specifically engineered for the park:"
    )

    # Screenshot Figure 4.13
    if os.path.exists("biomap_screenshot.png"):
        p_scr = doc.add_paragraph(); p_scr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_scr.paragraph_format.space_before = Pt(2); p_scr.paragraph_format.space_after = Pt(2)
        p_scr.add_run().add_picture("biomap_screenshot.png", width=Inches(6.25))
        p_scr_cap = doc.add_paragraph(); p_scr_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_scr_cap.paragraph_format.space_before = Pt(2); p_scr_cap.paragraph_format.space_after = Pt(4)
        r_sc = p_scr_cap.add_run("Fig. 4.13: Live BioMap Web-GIS Platform Interface (biomap-byculla.netlify.app)")
        r_sc.font.name = "Times New Roman"; r_sc.font.size = Pt(8.5); r_sc.font.bold = True; r_sc.font.color.rgb = RGB_SECONDARY

    add_body_p(
        "Key Architectural Capabilities of the BioMap Platform:", bold_prefix="System Implementation Highlights: "
    )
    add_bullet_p(
        "Real-Time Mobile Geolocation: Utilizing HTML5 Geolocation API, visitors can track their real-time position within the 53-acre garden trails "
        "and discover proximal flora and fauna without cellular data roaming latency."
    )
    add_bullet_p(
        "Micro-Zone Spatial Polygon Overlays: Renders the four ecological zones (Z1 to Z4) as distinct semi-transparent GeoJSON polygon layers, "
        "instantly communicating microclimatic and acoustic differences."
    )
    add_bullet_p(
        "Bilingual Ecological Profiles: Modal cards provide high-resolution photographs, Linnaean taxonomic hierarchy, IUCN threat classifications, "
        "flowering phenology, and vernacular Marathi descriptions (जैवविविधता माहिती)."
    )
    add_bullet_p(
        "Zero-Cost Open-Source Stack: Built entirely with Leaflet.js, OpenStreetMap, and Vanilla CSS3; requires zero database licensing costs, "
        "making it easily replicable for other municipal parks across Maharashtra."
    )

    doc.add_page_break() # -> PAGE 22 (Doc Page 22, Numbered 19)

    # ----------------------------------------------------
    # CHAPTER 5: CONCLUSION AND RECOMMENDATIONS
    # (Pages 22-23, Numbered 19-20)
    # ----------------------------------------------------
    add_chapter_title("5", "Conclusion and Recommendations")

    add_heading_2("5.1 Conclusion")
    add_body_p(
        "This Community Engagement Project successfully demonstrates the transformative potential of marrying applied Computer Science with urban biodiversity conservation "
        "under the University of Mumbai NEP 2020 framework. Through comprehensive empirical fieldwork at Veermata Jijabai Bhosale Botanical Udyan and Zoo (Byculla Zoo), "
        "the study verified that Byculla Zoo functions as an indispensable ecological sanctuary in South Mumbai, sustaining 42+ cataloged taxa and providing "
        "26.3 dBA of acoustic damping that shields vulnerable breeding avifauna from surrounding vehicular noise corridors."
    )
    add_body_p(
        "The project directly addressed the severe ecological literacy deficit identified among 84 surveyed park visitors by developing and deploying the open-source "
        "BioMap Web-GIS portal (biomap-byculla.netlify.app). By translating static, inaccessible botanical data into an intuitive, bilingual, mobile-first cartographic experience, "
        "the project fulfills the core NEP 2020 objective: delivering an open-source digital public good that empowers civic ecological literacy and supports sustainable urban governance."
    )

    add_heading_2("5.2 Recommendations")
    add_body_p(
        "Based on the empirical findings, survey insights, and caretaker testimonies, the investigators propose a structured Three-Tier Municipal Conservation Action Plan:"
    )
    add_bullet_p(
        "Recommendation 1 — QR-Coded Botanical Signage: Install weather-resistant, non-invasive QR code markers on all 256 accessioned heritage tree species. "
        "Scanning the code connects visitors directly to the BioMap bilingual species profile, converting routine walks into interactive botanical literacy tours."
    )
    add_bullet_p(
        "Recommendation 2 — Acoustic Vegetative Buffers along Ambedkar Road: Mandate a 15-meter dense, multi-tiered evergreen vegetative barrier along the eastern perimeter wall (Zone 4) "
        "using native species (e.g., Ficus benghalensis, Polyalthia longifolia) to dampen vehicular noise penetrating into Zone 1."
    )
    add_bullet_p(
        "Recommendation 3 — Geotagged Pre-Pruning Protocol: Institute a mandatory digital clearance protocol requiring municipal pruning contractors to cross-reference "
        "the BioMap cavity-nesting database before cutting branches, safeguarding breeding Coppersmith Barbets and Spotted Owlets."
    )
    add_bullet_p(
        "Recommendation 4 — Integration with BMC Tree Authority Portal: The open GeoJSON telemetry produced in this project should be formally integrated into "
        "the BMC Tree Authority digital repository to establish longitudinal canopy monitoring."
    )

    doc.add_page_break() # -> PAGE 23 (Doc Page 23, Numbered 20)

    # Table 5.1: Three-Tier Municipal Conservation Action Plan
    p_t51 = doc.add_paragraph(); p_t51.paragraph_format.space_before = Pt(4); p_t51.paragraph_format.space_after = Pt(2)
    p_t51.add_run("Table 5.1: Three-Tier Municipal Conservation Action Plan").font.bold = True
    tbl_51 = doc.add_table(rows=4, cols=4)
    tbl_51.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_51.autofit = False
    bp_w = [Inches(1.1), Inches(1.5), Inches(1.8), Inches(1.97)]
    style_table_header(tbl_51.rows[0], bp_w)
    tbl_51.cell(0, 0).paragraphs[0].text = "Tier / Horizon"
    tbl_51.cell(0, 1).paragraphs[0].text = "Proposed Policy Action"
    tbl_51.cell(0, 2).paragraphs[0].text = "Responsible Stakeholder"
    tbl_51.cell(0, 3).paragraphs[0].text = "Expected Community & Ecological Outcome"
    
    bp_data = [
        ("Tier 1\n(Immediate:\n1–3 Months)", "Physical QR Code Plaque Deployment on 50 Heritage Trees", "BMC Garden Dept. & Vartak College CS Dept.", "Instant smartphone access to bilingual botanical profiles for 5,000+ daily visitors."),
        ("Tier 2\n(Medium-Term:\n3–12 Months)", "Multi-Tiered Acoustic Green Buffer along Ambedkar Road Wall", "BMC Tree Authority & Environmental NGOs", "Reduces boundary noise penetration by 8–12 dBA; expands Zone 1 bird nesting area."),
        ("Tier 3\n(Long-Term:\n1–3 Years)", "Integration of BioMap with City-Wide BMC Tree Census API", "Brihanmumbai Municipal Corporation (BMC)", "Establishes longitudinal crowdsourced health monitoring for Mumbai’s heritage trees.")
    ]
    for r_i, (c0, c1, c2, c3) in enumerate(bp_data, start=1):
        tbl_51.rows[r_i].cells[0].paragraphs[0].text = c0
        tbl_51.rows[r_i].cells[1].paragraphs[0].text = c1
        tbl_51.rows[r_i].cells[2].paragraphs[0].text = c2
        tbl_51.rows[r_i].cells[3].paragraphs[0].text = c3
    style_table_rows(tbl_51, bp_w)

    add_body_p(
        "Summary of Experiential Learning Under NEP 2020: Executing this Community Engagement Project provided the student team with invaluable cross-disciplinary "
        "competencies spanning environmental informatics, geospatial analysis, field telemetry, public stakeholder engagement, and modern full-stack web engineering. "
        "The project underscores that computer science education achieves its highest civic calling when applied to solving pressing ecological challenges in our local communities.",
        space_after=8
    )

    doc.add_page_break() # -> PAGE 24 (Doc Page 24, Numbered 21)

    # ----------------------------------------------------
    # REFERENCES (Page 24, Numbered 21)
    # ----------------------------------------------------
    add_heading_1("REFERENCES")
    p_ref_note = doc.add_paragraph(); p_ref_note.paragraph_format.space_after = Pt(8)
    p_ref_note.add_run("Formal Academic Citations & Statutory References (APA 7th Edition)").font.italic = True

    references_list = [
        ("Ali, S. (2002). ", "The Book of Indian Birds (13th ed.). Oxford University Press & Bombay Natural History Society (BNHS)."),
        ("Birdwood, G. (1862). ", "Catalogue of the Vegetable Productions of the Presidency of Bombay. Government Central Press, Bombay."),
        ("Bombay Natural History Society (BNHS). (2021). ", "Urban Biodiversity Assessment of Greater Mumbai Green Refugia. BNHS Technical Report No. 42."),
        ("Brihanmumbai Municipal Corporation (BMC). (2020). ", "Tree Census and Arboretum Inventory: Veermata Jijabai Bhosale Botanical Udyan and Zoo. Garden & Tree Authority Department."),
        ("Central Zoo Authority (CZA). (2022). ", "Guidelines for Ex-Situ Conservation and Master Planning of Zoos in India. Ministry of Environment, Forest and Climate Change, Government of India."),
        ("Dickinson, J. L., Shirk, J., Bonter, D., Bonney, R., Crain, R. L., Martin, J., Phillips, T., & Purcell, K. (2012). ", "The current state of citizen science as a tool for ecological research and public engagement. Frontiers in Ecology and the Environment, 10(6), 291–297."),
        ("Forman, R. T. (1995). ", "Land Mosaics: The Ecology of Landscapes and Regions. Cambridge University Press."),
        ("Kothari, A. (2018). ", "Common Trees of Mumbai: An Illustrated Field Guide. Nature Conservation Trust, Mumbai."),
        ("MacArthur, R. H., & Wilson, E. O. (1967). ", "The Theory of Island Biogeography. Princeton University Press."),
        ("OpenStreetMap Contributors. (2026). ", "Planet Dump and Mumbai Geographic Features Data. Retrieved from https://www.openstreetmap.org/"),
        ("Shannon, C. E., & Weaver, W. (1949). ", "The Mathematical Theory of Communication. University of Illinois Press."),
        ("Sullivan, B. L., et al. (2014). ", "The eBird enterprise: An integrated approach to development and application of citizen science. Biological Conservation, 169, 31–40."),
        ("University of Mumbai. (2023). ", "Curricular Framework for Community Engagement Project (CEP) under National Education Policy (NEP 2020). Board of Studies in Computer Science."),
        ("Veermata Jijabai Bhosale Botanical Udyan & Zoo. (2024). ", "Official Guidebook and Heritage Plant Catalog. Municipal Corporation of Greater Mumbai (MCGM).")
    ]
    for pre, post in references_list:
        p_ref = doc.add_paragraph()
        p_ref.paragraph_format.space_before = Pt(0); p_ref.paragraph_format.space_after = Pt(4)
        p_ref.paragraph_format.line_spacing = 1.15
        p_ref.paragraph_format.left_indent = Inches(0.4)
        p_ref.paragraph_format.first_line_indent = Inches(-0.4)
        r_pre = p_ref.add_run(pre); r_pre.font.name = "Times New Roman"; r_pre.font.size = Pt(9.5); r_pre.font.bold = True
        r_post = p_ref.add_run(post); r_post.font.name = "Times New Roman"; r_post.font.size = Pt(9.5)

    doc.add_page_break() # -> PAGE 25 (Doc Page 25, Numbered 22)

    # ----------------------------------------------------
    # APPENDICES: A, B, C (Pages 25-26, Numbered 22-23)
    # ----------------------------------------------------
    add_heading_1("APPENDICES")

    add_heading_2("Appendix A — Community Biodiversity & Platform Survey Questionnaire")
    add_body_p(
        "The survey instrument comprised thirteen structured questions administered to N = 84 participants (malis, walkers, youth):",
        space_after=3
    )
    survey_questions = [
        "Q1. How frequently do you visit Byculla Zoo? [Daily / Weekly / Monthly / First time]",
        "Q2. What is your primary purpose of visit? [Walking & exercise / Clean air & quiet / Seeing caged animals / Botanical interest]",
        "Q3. Have you observed free-living birds, butterflies, or bats in the park? [Frequently / Occasionally / Never / Unaware]",
        "Q4. Where in the garden do you most frequently observe free-living wildlife? [Botanical Grove (Z1) / Lotus Lake (Z2) / Lawn Borders (Z3) / Perimeter (Z4)]",
        "Q5. Are you aware that Byculla Zoo houses over 250 species of trees, including heritage Baobabs? [Yes, fully aware / Heard of it / Not aware at all]",
        "Q6. Do you find existing information boards inside the park sufficient? [Sufficient / Focused only on caged animals / Missing for trees / Inadequate]",
        "Q7. What is the biggest barrier preventing you from learning about the park’s trees and wildlife? [Lack of digital map / No signs / Language barrier / Lack of time]",
        "Q8. Would an interactive web application or mobile digital map enhance your visit experience? [Definitely Yes / Maybe / No / Not sure]",
        "Q9. Would you use your smartphone to scan QR codes on trees to view photos, scientific names, and uses? [Definitely / Likely / Rarely / No]",
        "Q10. What information should a digital species card include? [Photos / Names (English & Marathi) / Flowering season / Wildlife attracted / All of these]",
        "Q11. How would you prefer to discover wildlife in the park? [GPS map / List by zones / Search by name / Category filters / All of these]",
        "Q12. Would you be interested in submitting your own sightings to help map urban wildlife? [Very interested / Somewhat interested / Not interested]",
        "Q13. Do you support municipal initiatives to create dense green acoustic buffer walls along Ambedkar Road? [Strongly support / Support / Neutral / Oppose]"
    ]
    for q in survey_questions:
        add_bullet_p(q, space_after=1.8, line_spacing=1.12)

    doc.add_page_break() # -> PAGE 26 (Doc Page 26, Numbered 23)

    add_heading_2("Appendix B — Open-Source Web-GIS Architecture & Specifications")
    add_body_p(
        "The BioMap platform is architected as an open-source, client-side, single-page application (SPA) optimized for low-bandwidth cellular environments:",
        space_after=4
    )
    add_bullet_p("Base Layer: OpenStreetMap standard raster tiles cached locally via Service Worker API for offline reliability.", bold_prefix="Cartographic Tile Engine: ")
    add_bullet_p("Spatial Vector Layer: GeoJSON standard format storing coordinates for 42+ taxa points and 4 survey zone boundary polygons.", bold_prefix="Vector Data Schema: ")
    add_bullet_p("Frontend Logic: Leaflet.js v1.9.4 with custom SVG marker pins, dynamic popup modals, and touch-optimized responsive controls.", bold_prefix="Rendering Engine: ")
    add_bullet_p("Styling System: Vanilla CSS3 utilizing CSS Custom Properties, flexbox/grid layout, and dark/light mode toggle with WCAG AA compliance.", bold_prefix="Design Architecture: ")
    add_bullet_p("Production URL: Publicly accessible at https://biomap-byculla.netlify.app under the MIT open-source license.", bold_prefix="Hosting & Licensing: ")

    add_heading_2("Appendix C — Field Study Photographic Evidence")
    add_body_p(
        "Photographic vouchers collected on-site by student investigators Riddhesh Keni and Shivansh Dubey during field surveys at Byculla Zoo on 31/08/2026:",
        space_after=4
    )

    tbl_pf2 = doc.add_table(rows=1, cols=2)
    tbl_pf2.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_pf2.autofit = False
    c0_pf2 = tbl_pf2.cell(0, 0); c1_pf2 = tbl_pf2.cell(0, 1)
    c0_pf2.width = Inches(2.3); c1_pf2.width = Inches(4.07)
    set_cell_margins(c0_pf2, top=40, bottom=40, left=45, right=45)
    set_cell_margins(c1_pf2, top=40, bottom=40, left=60, right=45)
    b_pf2 = ("single", "8", HEX_PRIMARY)
    set_cell_borders(c0_pf2, top=b_pf2, bottom=b_pf2, left=b_pf2, right=b_pf2)
    set_cell_borders(c1_pf2, top=b_pf2, bottom=b_pf2, left=b_pf2, right=b_pf2)

    p0_pf2 = c0_pf2.paragraphs[0]; p0_pf2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    proof_img = r"animal_photos\fieldwork_proof.jpg"
    if os.path.exists(proof_img):
        p0_pf2.add_run().add_picture(proof_img, width=Inches(2.0))
    p_cap_pf2 = c0_pf2.add_paragraph(); p_cap_pf2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_pf2.paragraph_format.space_before = Pt(2); p_cap_pf2.paragraph_format.space_after = Pt(0)
    p_cap_pf2.add_run("Fig. C.1: Fieldwork Arrival Proof\nByculla Station Gate (18.9766° N, 72.8328° E)\n31/08/2026, 11:41 AM IST").font.size = Pt(7.0)

    p1_pf2 = c1_pf2.paragraphs[0]; p1_pf2.paragraph_format.line_spacing = 1.12
    p1_pf2.add_run("FIELDWORK VERIFICATION VOUCHER\n").font.bold = True
    p1_pf2.add_run("• Investigators: Riddhesh Keni (Roll 35) & Shivansh Dubey (Roll 45)\n"
                  "• Academic Year: 2026–2027 (T.Y. B.Sc. Computer Science)\n"
                  "• Institution: Vidyavardhini's Vartak College, Vasai Road (W)\n"
                  "• Study Location: Veermata Jijabai Bhosale Botanical Udyan and Zoo\n"
                  "• Date & Timing: Monday, 31/08/2026 (07:00 AM – 06:30 PM)\n"
                  "• Total Survey Distance: 4.8 km traversed across all four ecological zones\n"
                  "• Field Data Deliverables: 42 species records, 84 survey responses, 1 Web-GIS app").font.size = Pt(8.0)

    doc.add_page_break() # -> PAGE 27 (Doc Page 27, Numbered 24)

    # ----------------------------------------------------
    # GUIDE INTERACTION DIARY FORM (Strictly Matching Reference Page 29)
    # ----------------------------------------------------
    add_heading_1("GUIDE INTERACTION DIARY FORM")
    p_gid = doc.add_paragraph()
    p_gid.paragraph_format.space_before = Pt(10); p_gid.paragraph_format.space_after = Pt(14)
    p_gid.paragraph_format.line_spacing = 1.35
    r_gid = p_gid.add_run(
        "I,  the  undersigned  Ms./Mr  ___________________________________________  Roll.  \n"
        "No.____________ currently enrolled in the ______ Year of _________________________ Full-\n"
        "time Program __________________ at ____________________________ institute/ college, am  \n"
        "undertaking  my  Community  Engagement  Project  work  under  the  guidance  of  \n"
        "Prof. Shweta Iyer, and I hereby confirm that I have met my Internal guide on the  \n"
        "following dates mentioned below for Project Guidance:-"
    )
    r_gid.font.name = "Times New Roman"; r_gid.font.size = Pt(10.5)

    tbl_diary = doc.add_table(rows=11, cols=3)
    tbl_diary.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_diary.autofit = False
    gid_w = [Inches(0.8), Inches(1.8), Inches(3.67)]

    # Header Row
    headers_gid = ["Sr. No.", "Date", "Signature of the Internal Guide"]
    for c_i, h_txt in enumerate(headers_gid):
        cell = tbl_diary.cell(0, c_i); cell.width = gid_w[c_i]
        set_cell_margins(cell, top=35, bottom=35, left=40, right=40)
        set_cell_borders(cell, top=("single", "4", "000000"), bottom=("single", "4", "000000"), left=("single", "4", "000000"), right=("single", "4", "000000"))
        set_cell_shading(cell, "D9E1F2")
        p = cell.paragraphs[0]; p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        if c_i == 0: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h_txt); r.font.name = "Times New Roman"; r.font.size = Pt(10); r.font.bold = True

    # 10 Numbered Rows
    for r_i in range(1, 11):
        row = tbl_diary.rows[r_i]
        for c_i in range(3):
            cell = row.cells[c_i]; cell.width = gid_w[c_i]
            set_cell_margins(cell, top=30, bottom=30, left=40, right=40)
            set_cell_borders(cell, top=("single", "4", "000000"), bottom=("single", "4", "000000"), left=("single", "4", "000000"), right=("single", "4", "000000"))
            p = cell.paragraphs[0]; p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
        p0 = row.cells[0].paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p0.add_run(str(r_i)); r0.font.name = "Times New Roman"; r0.font.size = Pt(10)

    # Signature block below table (Candidate on Left, Internal Guide on Right)
    tbl_sig = doc.add_table(rows=1, cols=2)
    tbl_sig.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_sig.autofit = False
    sig_w = [Inches(3.13), Inches(3.14)]
    
    cell_s0 = tbl_sig.cell(0, 0); cell_s0.width = sig_w[0]
    set_cell_margins(cell_s0, top=60, bottom=10, left=10, right=10)
    set_cell_borders(cell_s0)
    p_s0 = cell_s0.paragraphs[0]; p_s0.paragraph_format.space_before = Pt(40); p_s0.paragraph_format.space_after = Pt(0)
    p_s0.paragraph_format.line_spacing = 1.2
    r_s0 = p_s0.add_run("________________________\nSignature of the Candidate")
    r_s0.font.name = "Times New Roman"; r_s0.font.size = Pt(10.5); r_s0.font.bold = True

    cell_s1 = tbl_sig.cell(0, 1); cell_s1.width = sig_w[1]
    set_cell_margins(cell_s1, top=60, bottom=10, left=10, right=10)
    set_cell_borders(cell_s1)
    p_s1 = cell_s1.paragraphs[0]; p_s1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_s1.paragraph_format.space_before = Pt(40); p_s1.paragraph_format.space_after = Pt(0)
    p_s1.paragraph_format.line_spacing = 1.2
    r_s1 = p_s1.add_run("______________________\nSignature of Internal Guide\n(Prof. Shweta Iyer)")
    r_s1.font.name = "Times New Roman"; r_s1.font.size = Pt(10.5); r_s1.font.bold = True

    # ----------------------------------------------------
    # SECTION 3: STUDENT FEEDBACK ON CEP (Strictly Matching Reference Page 30)
    # Fits on exactly one single page with Aspect of CEP Experience Table
    # ----------------------------------------------------
    sec_fb = doc.add_section()
    sec_fb.page_width = Inches(8.27)
    sec_fb.page_height = Inches(11.69)
    sec_fb.top_margin = Inches(0.5)
    sec_fb.bottom_margin = Inches(0.4)
    sec_fb.left_margin = Inches(0.8)
    sec_fb.right_margin = Inches(0.8)
    sec_fb.header.is_linked_to_previous = False
    sec_fb.footer.is_linked_to_previous = True

    sectPr_fb = sec_fb._sectPr
    docGrid_fb = sectPr_fb.find(qn('w:docGrid'))
    if docGrid_fb is not None:
        docGrid_fb.set(qn('w:type'), 'default')
    pg_fb = sectPr_fb.find(qn('w:pgNumType'))
    if pg_fb is not None:
        sectPr_fb.remove(pg_fb)

    p_fb_title = doc.add_paragraph()
    p_fb_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_fb_title.paragraph_format.space_before = Pt(0); p_fb_title.paragraph_format.space_after = Pt(3)
    r_fbt = p_fb_title.add_run("STUDENT FEEDBACK ON CEP")
    r_fbt.font.name = "Times New Roman"; r_fbt.font.size = Pt(13); r_fbt.font.bold = True

    fb_fields = [
        "Student Name: ______________________________________",
        "Seat No. / Roll No.: ______________     Email: ______________________",
        "Department: ________________     Name of the Mentor: Prof. Shweta Iyer",
        "Title/Heading of Community Engagement Project: LOCAL BIODIVERSITY MAPPING OF BYCULLA ZOO AND SURROUNDING GREEN SPACES",
        "Brief description of CEP carried out: ____________________________",
        "Dates of CEP: ______________________________________",
        "My internship experience was related to my major area of study/academic program:  1. Strongly Agree   2. Agree   3. Disagree   4. Strongly Disagree",
        "Indicate the degree to which you agree or disagree with the following statements. This experience has:"
    ]

    for i, f_text in enumerate(fb_fields):
        p_fld = doc.add_paragraph()
        p_fld.paragraph_format.space_before = Pt(0); p_fld.paragraph_format.space_after = Pt(1 if i < 6 else 2)
        p_fld.paragraph_format.line_spacing = 1.0
        r_fld = p_fld.add_run(f_text)
        r_fld.font.name = "Times New Roman"
        r_fld.font.size = Pt(8.5) if i < 6 else Pt(8.0)

    # Likert Table
    fb_w = [Inches(3.67), Inches(0.60), Inches(0.60), Inches(0.60), Inches(0.60), Inches(0.60)]
    tbl_fb_likert = doc.add_table(rows=17, cols=6)
    tbl_fb_likert.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_fb_likert.autofit = False

    fb_headers = ["This experience has:", "Strongly\nAgree", "Agree", "Neutral", "Disagree", "Strongly\nDisagree"]
    for c_i, h in enumerate(fb_headers):
        cell = tbl_fb_likert.cell(0, c_i); cell.width = fb_w[c_i]
        set_cell_margins(cell, top=4, bottom=4, left=10, right=10)
        set_cell_borders(cell, top=("single", "4", "000000"), bottom=("single", "4", "000000"), left=("single", "4", "000000"), right=("single", "4", "000000"))
        set_cell_shading(cell, "D9E1F2")
        p = cell.paragraphs[0]; p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        if c_i > 0: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.font.name = "Times New Roman"; r.font.size = Pt(7.0); r.font.bold = True

    fb_stmts = [
        "Has increased my sensitivity towards societal problems",
        "Given me the opportunity to explore a career field",
        "Allowed me to apply classroom theory to practice",
        "Helped me develop my decision-making and problem-solving skills",
        "Expanded my knowledge about the work world before permanent employment",
        "Helped me develop my written and oral communication skills",
        "Provided a chance to use leadership skills (influence others, develop ideas with others, stimulate decision-making and action)",
        "Expanded my sensitivity to the ethical implications of the work involved",
        "Made it possible for me to be more confident in new situations",
        "Given me a chance to improve my interpersonal skills",
        "Helped me learn to handle responsibility and use my time wisely",
        "Helped me discover new aspects of myself that I didn’t know existed before",
        "Helped me develop new interests and abilities",
        "Helped me clarify my career goals",
        "Allowed me to acquire information and/or use equipment not available at my Institute",
        "Allowed me to realize socio-economic issues in the society"
    ]

    for r_i, s in enumerate(fb_stmts, start=1):
        row = tbl_fb_likert.rows[r_i]
        for c_i in range(6):
            cell = row.cells[c_i]; cell.width = fb_w[c_i]
            set_cell_margins(cell, top=2, bottom=2, left=10, right=10)
            set_cell_borders(cell, top=("single", "4", "000000"), bottom=("single", "4", "000000"), left=("single", "4", "000000"), right=("single", "4", "000000"))
            p_c = cell.paragraphs[0]
            p_c.paragraph_format.space_before = Pt(0); p_c.paragraph_format.space_after = Pt(0)
            p_c.paragraph_format.line_spacing = 1.0
        p = row.cells[0].paragraphs[0]
        r = p.add_run(s); r.font.name = "Times New Roman"; r.font.size = Pt(7.0)

    fb_questions = [
        ("● The faculty mentor extended guidance and mentoring throughout the CEP process:", "1. Strongly Agree       2. Agree       3. Disagree       4. Strongly Disagree"),
        ("● Were you able to accomplish the initial goals, tasks and new skills set down in the CEP plan?", "1. Strongly Agree       2. Agree       3. Disagree       4. Strongly Disagree"),
        ("● Considering your overall experience, how would you rate this CEP? (Circle one):", "Poor         /         Satisfactory         /         Good         /         Excellent"),
        ("● Give suggestions as to how your CEP experience could have been improved:", "(Please mark ✓ for areas that were satisfactory and ✗ for areas that need improvement)")
    ]

    for q_title, q_opts in fb_questions:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(1); p_q.paragraph_format.space_after = Pt(0)
        p_q.paragraph_format.line_spacing = 1.0
        r_qt = p_q.add_run(q_title); r_qt.font.name = "Times New Roman"; r_qt.font.size = Pt(7.5); r_qt.font.bold = True
        
        p_opt = doc.add_paragraph()
        p_opt.paragraph_format.space_before = Pt(0); p_opt.paragraph_format.space_after = Pt(1)
        p_opt.paragraph_format.line_spacing = 1.0
        r_qo = p_opt.add_run(q_opts); r_qo.font.name = "Times New Roman"; r_qo.font.size = Pt(7.0)

    # Table 2: Aspect of CEP Experience
    asp_w = [Inches(5.67), Inches(1.00)]
    tbl_asp = doc.add_table(rows=7, cols=2)
    tbl_asp.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_asp.autofit = False

    cell0 = tbl_asp.cell(0, 0); cell0.width = asp_w[0]
    set_cell_margins(cell0, top=3, bottom=3, left=15, right=15)
    set_cell_borders(cell0, top=("single", "4", "000000"), bottom=("single", "4", "000000"), left=("single", "4", "000000"), right=("single", "4", "000000"))
    set_cell_shading(cell0, "D9E1F2")
    p0 = cell0.paragraphs[0]; p0.paragraph_format.space_before = Pt(0); p0.paragraph_format.space_after = Pt(0)
    p0.paragraph_format.line_spacing = 1.0
    r0 = p0.add_run("Aspect of CEP Experience"); r0.font.name = "Times New Roman"; r0.font.size = Pt(7.5); r0.font.bold = True

    cell1 = tbl_asp.cell(0, 1); cell1.width = asp_w[1]
    set_cell_margins(cell1, top=3, bottom=3, left=15, right=15)
    set_cell_borders(cell1, top=("single", "4", "000000"), bottom=("single", "4", "000000"), left=("single", "4", "000000"), right=("single", "4", "000000"))
    set_cell_shading(cell1, "D9E1F2")
    p1 = cell1.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(0); p1.paragraph_format.space_after = Pt(0)
    p1.paragraph_format.line_spacing = 1.0
    r1 = p1.add_run("✓ / ✗"); r1.font.name = "Times New Roman"; r1.font.size = Pt(7.5); r1.font.bold = True

    aspect_items = [
        "Clear orientation and briefing at the start",
        "Relevance of project to academic learning",
        "Availability and support from internal guide",
        "Adequate field exposure and hands-on engagement",
        "Clarity in roles and responsibilities",
        "Time provided for the completion of project"
    ]

    for r_i, asp in enumerate(aspect_items, start=1):
        c0 = tbl_asp.cell(r_i, 0); c0.width = asp_w[0]
        set_cell_margins(c0, top=2, bottom=2, left=15, right=15)
        set_cell_borders(c0, top=("single", "4", "000000"), bottom=("single", "4", "000000"), left=("single", "4", "000000"), right=("single", "4", "000000"))
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0); p0.paragraph_format.space_after = Pt(0)
        p0.paragraph_format.line_spacing = 1.0
        r = p0.add_run(asp); r.font.name = "Times New Roman"; r.font.size = Pt(7.5)
        
        c1 = tbl_asp.cell(r_i, 1); c1.width = asp_w[1]
        set_cell_margins(c1, top=2, bottom=2, left=15, right=15)
        set_cell_borders(c1, top=("single", "4", "000000"), bottom=("single", "4", "000000"), left=("single", "4", "000000"), right=("single", "4", "000000"))
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(0); p1.paragraph_format.space_after = Pt(0)
        p1.paragraph_format.line_spacing = 1.0

    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(3); p_sig.paragraph_format.space_after = Pt(0)
    r_sig = p_sig.add_run("Signature: ________________   Name: ________________   Date: ________")
    r_sig.font.name = "Times New Roman"; r_sig.font.size = Pt(8.0)

    # Save document
    out_file = "Byculla Zoo Biodiversity Mapping - CEP Report.docx"
    doc.save(out_file)
    print(f"Document successfully saved as '{out_file}'")
    return out_file

def get_word_page_map(docx_filename):
    abs_path = os.path.abspath(docx_filename)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    page_map = {}
    try:
        doc = word.Documents.Open(abs_path)
        word.ActiveWindow.View.Type = 3 # wdPrintView
        
        # In our document, Section 1 has 3 unnumbered pages.
        # Section 2 starts page numbering at 1 on Acknowledgement.
        # Word's wdActiveEndPageNumber gives the absolute page in the document (1, 2, 3, 4, 5...)
        # Since Acknowledgement is doc page 4 and section page 1,
        # Section Page Number = Absolute Doc Page - 3!
        
        for p in doc.Paragraphs:
            txt = p.Range.Text.strip()
            if not txt:
                continue
            abs_pg = p.Range.Information(3) # 3 = wdActiveEndPageNumber
            sec_pg = max(1, abs_pg - 3) if abs_pg >= 4 else 1
            in_table = p.Range.Information(12) # 12 = wdWithInTable

            # Headings (Must NOT be inside TOC table!)
            if not in_table:
                if txt.startswith("ACKNOWLEDGEMENT"):
                    page_map.setdefault("ACK", sec_pg)
                elif txt.startswith("ABSTRACT"):
                    page_map.setdefault("ABS", sec_pg)
                elif txt.startswith("LIST OF FIGURES AND TABLES"):
                    page_map.setdefault("LOF", sec_pg)
                elif "CHAPTER 1" in txt:
                    page_map.setdefault("CH1", sec_pg)
                elif txt.startswith("1.1 Purpose of the Visit"):
                    page_map.setdefault("SEC_1_1", sec_pg)
                elif txt.startswith("1.2 Background Information"):
                    page_map.setdefault("SEC_1_2", sec_pg)
                elif txt.startswith("1.3 Significance"):
                    page_map.setdefault("SEC_1_3", sec_pg)
                elif txt.startswith("1.4 Scope of the Report"):
                    page_map.setdefault("SEC_1_4", sec_pg)
                elif "CHAPTER 2" in txt:
                    page_map.setdefault("CH2", sec_pg)
                elif txt.startswith("2.1 Introduction"):
                    page_map.setdefault("SEC_2_1", sec_pg)
                elif txt.startswith("2.2 Urban Ecological Refugia"):
                    page_map.setdefault("SEC_2_2", sec_pg)
                elif txt.startswith("2.3 Citizen Science & Community Participation"):
                    page_map.setdefault("SEC_2_3", sec_pg)
                elif txt.startswith("2.4 Existing Biodiversity Platforms & Limitations"):
                    page_map.setdefault("SEC_2_4", sec_pg)
                elif txt.startswith("2.5 Technological Solutions & Open-Source Geospatial Mapping"):
                    page_map.setdefault("SEC_2_5", sec_pg)
                elif txt.startswith("2.6 Relevance to the BioMap Proposal"):
                    page_map.setdefault("SEC_2_6", sec_pg)
                elif "CHAPTER 3" in txt:
                    page_map.setdefault("CH3", sec_pg)
                elif txt.startswith("3.1 Data Collection Approach"):
                    page_map.setdefault("SEC_3_1", sec_pg)
                elif txt.startswith("3.2 Sampling Method & Survey Zones"):
                    page_map.setdefault("SEC_3_2", sec_pg)
                elif txt.startswith("3.3 Field Study and Location Evidence"):
                    page_map.setdefault("SEC_3_3", sec_pg)
                elif txt.startswith("3.4 Data Analysis & Quantitative Diversity Metrics"):
                    page_map.setdefault("SEC_3_4", sec_pg)
                elif txt.startswith("3.5 Rationale for the Selected Methods"):
                    page_map.setdefault("SEC_3_5", sec_pg)
                elif "CHAPTER 4" in txt:
                    page_map.setdefault("CH4", sec_pg)
                elif txt.startswith("4.1 Community Interactions and Multi-Stakeholder Engagement"):
                    page_map.setdefault("SEC_4_1", sec_pg)
                elif txt.startswith("4.2 Field Observations & Location-Based Species Evidence"):
                    page_map.setdefault("SEC_4_2", sec_pg)
                elif txt.startswith("4.3 Survey Findings and Quantitative Analysis"):
                    page_map.setdefault("SEC_4_3", sec_pg)
                elif txt.startswith("4.4 Consolidated Key Findings"):
                    page_map.setdefault("SEC_4_4", sec_pg)
                elif txt.startswith("4.5 Analysis in Relation to Study Objectives & Web-GIS Portal"):
                    page_map.setdefault("SEC_4_5", sec_pg)
                elif "CHAPTER 5" in txt:
                    page_map.setdefault("CH5", sec_pg)
                elif txt.startswith("5.1 Conclusion"):
                    page_map.setdefault("SEC_5_1", sec_pg)
                elif txt.startswith("5.2 Recommendations"):
                    page_map.setdefault("SEC_5_2", sec_pg)
                elif txt.startswith("REFERENCES"):
                    page_map.setdefault("REF", sec_pg)
                elif txt.startswith("APPENDICES"):
                    page_map.setdefault("APP", sec_pg)
                elif txt.startswith("GUIDE INTERACTION DIARY FORM"):
                    page_map.setdefault("DIARY", sec_pg)
                elif (txt.startswith("STUDENT FEEDBACK ON CEP") or txt.startswith("STUDENT FEEDBACK ON COMMUNITY ENGAGEMENT")) and not txt.startswith("•"):
                    page_map.setdefault("FEEDBACK", sec_pg)

                # Tables (Captions are outside tables)
                elif "Table 1.1: Core Challenges" in txt:
                    page_map.setdefault("TAB_1_1", sec_pg)
                elif "Table 1.2: Research Objectives" in txt:
                    page_map.setdefault("TAB_1_2", sec_pg)
                elif "Table 3.1: Delineation of Ecological Survey Zones" in txt:
                    page_map.setdefault("TAB_3_1", sec_pg)
                elif "Table 3.2: Field Data Indicators" in txt:
                    page_map.setdefault("TAB_3_2", sec_pg)
                elif "Table 3.3: Hardware, Software & Geospatial Instrumentation Stack" in txt:
                    page_map.setdefault("TAB_3_3", sec_pg)
                elif "Table 4.1: Consolidated Key Findings" in txt:
                    page_map.setdefault("TAB_4_1", sec_pg)
                elif "Table 5.1: Three-Tier Municipal Conservation Action Plan" in txt:
                    page_map.setdefault("TAB_5_1", sec_pg)

            # Figures (Can be in table cells or paragraphs, but ignore LOF table which has sec_pg == 4)
            if sec_pg != 4:
                if "Fig. 4.1: Location 1" in txt:
                    page_map.setdefault("FIG_4_1", sec_pg)
                elif "Fig. 4.2: Location 2" in txt:
                    page_map.setdefault("FIG_4_2", sec_pg)
                elif "Fig. 4.3: Location 3" in txt:
                    page_map.setdefault("FIG_4_3", sec_pg)
                elif "Fig. 4.4: Location 4" in txt:
                    page_map.setdefault("FIG_4_4", sec_pg)
                elif "Fig. 4.5: Where Is Free-Living Wildlife Observed?" in txt:
                    page_map.setdefault("FIG_4_5", sec_pg)
                elif "Fig. 4.6: What Threatens Free-Living Flora & Avifauna?" in txt:
                    page_map.setdefault("FIG_4_6", sec_pg)
                elif "Fig. 4.7: Main Barriers to Local Biodiversity Literacy" in txt:
                    page_map.setdefault("FIG_4_7", sec_pg)
                elif "Figs. 4.8 & 4.9: Would App Help?" in txt:
                    page_map.setdefault("FIG_4_8", sec_pg)
                    page_map.setdefault("FIG_4_9", sec_pg)
                elif "Fig. 4.10" in txt:
                    page_map.setdefault("FIG_4_10", sec_pg)
                elif "Fig. 4.11" in txt or "Figs. 4.11" in txt:
                    page_map.setdefault("FIG_4_11", sec_pg)
                    page_map.setdefault("FIG_4_12", sec_pg)
                elif "Fig. 4.12" in txt:
                    page_map.setdefault("FIG_4_12", sec_pg)
                elif "Fig. 4.13" in txt:
                    page_map.setdefault("FIG_4_13", sec_pg)

        doc.Close(False)
        return page_map
    finally:
        word.Quit()

def export_to_pdf(docx_filename, pdf_filename):
    abs_docx = os.path.abspath(docx_filename)
    abs_pdf = os.path.abspath(pdf_filename)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(abs_docx)
        doc.SaveAs(abs_pdf, FileFormat=17) # 17 = wdFormatPDF
        doc.Close(False)
        print(f"PDF exported successfully to '{pdf_filename}' (Size: {os.path.getsize(abs_pdf):,} bytes)")
    finally:
        word.Quit()

def main():
    print("==========================================================")
    print("HARMONIZED CEP MASTER REPORT GENERATOR")
    print("Combining Deep Emerald Design & Exact Guidelines Structure")
    print("==========================================================")
    
    print("\n--- PASS 1: Generating Initial Document Layout ---")
    out_docx = generate_document()
    
    print("\n--- PASS 2: Detecting Exact Word COM Page Numbers ---")
    page_map = get_word_page_map(out_docx)
    print("Detected Page Mapping:", page_map)
    
    print("\n--- PASS 3: Re-generating Document with 100% Synchronized Page Numbers ---")
    generate_document(page_map)
    
    print("\n--- PASS 4: Exporting Verified PDF Document ---")
    out_pdf = "Byculla Zoo Biodiversity Mapping - CEP Report.pdf"
    export_to_pdf(out_docx, out_pdf)
    
    print("\nALL PASSES COMPLETED SUCCESSFULLY!")

if __name__ == "__main__":
    main()
