import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import win32com.client as win32

# ----------------------------------------------------
# Color Palette (Deep Forest Emerald Academic Theme)
# ----------------------------------------------------
HEX_PRIMARY_DARK  = "1B4332"  # Deep Forest Green
HEX_PRIMARY_MED   = "2D6A4F"  # Emerald Green
HEX_PRIMARY_LIGHT = "D8F3DC"  # Mint Tint
HEX_ACCENT_AMBER  = "D97706"  # Warm Amber
HEX_ACCENT_BLUE   = "1E40AF"  # Slate Blue
HEX_BORDER        = "CBD5E1"  # Light Border
HEX_BG_LIGHT      = "F8FAFC"  # Off-white

RGB_PRIMARY_DARK  = RGBColor(27, 67, 50)
RGB_PRIMARY_MED   = RGBColor(45, 106, 79)
RGB_TEXT_DARK     = RGBColor(26, 32, 44)
RGB_TEXT_MUTED    = RGBColor(100, 116, 139)
RGB_AMBER         = RGBColor(217, 119, 6)

def set_cell_shading(cell, color_hex):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side, border in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        el = OxmlElement(f'w:{side}')
        if border:
            val, sz, col = border
            el.set(qn('w:val'), val)
            el.set(qn('w:sz'), str(sz))
            el.set(qn('w:color'), col)
        else:
            el.set(qn('w:val'), 'none')
        tcBorders.append(el)
    tcPr.append(tcBorders)

def build_summary_doc():
    doc = docx.Document()

    # Page Setup: A4 with 0.6" margins
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)

    # Header and Footer
    footer = section.footer
    p_ft = footer.paragraphs[0]
    p_ft.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_ft = p_ft.add_run("Byculla Zoo CEP (BioMap) — Project Overview & Viva Handbook  |  Page ")
    r_ft.font.name = "Calibri"; r_ft.font.size = Pt(8.5); r_ft.font.color.rgb = RGB_TEXT_MUTED
    fld = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
    p_ft._p.append(fld)

    # Header Helper
    def add_sec_title(title, icon=""):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        
        # Bottom border under section title
        pBdr = parse_xml(r'<w:pBdr %s><w:bottom w:val="single" w:sz="12" w:space="4" w:color="%s"/></w:pBdr>' % (nsdecls('w'), HEX_PRIMARY_MED))
        p._p.get_or_add_pPr().append(pBdr)

        r = p.add_run(f"{icon} {title}" if icon else title)
        r.font.name = "Georgia"; r.font.size = Pt(13); r.font.bold = True
        r.font.color.rgb = RGB_PRIMARY_DARK

    def add_callout(title, body, bg_hex=HEX_PRIMARY_LIGHT, border_hex=HEX_PRIMARY_MED, icon="📌"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.autofit = False
        tbl.columns[0].width = Inches(6.67)
        cell = tbl.cell(0, 0)
        set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
        set_cell_borders(cell, left=("single", "24", border_hex), top=None, bottom=None, right=None)
        set_cell_shading(cell, bg_hex)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(2)
        r_t = p.add_run(f"{icon} {title}\n")
        r_t.font.name = "Calibri"; r_t.font.size = Pt(10.5); r_t.font.bold = True
        r_t.font.color.rgb = RGB_PRIMARY_DARK
        
        p_b = cell.add_paragraph()
        p_b.paragraph_format.space_before = Pt(0); p_b.paragraph_format.space_after = Pt(0)
        p_b.paragraph_format.line_spacing = 1.15
        r_b = p_b.add_run(body)
        r_b.font.name = "Calibri"; r_b.font.size = Pt(9.5); r_b.font.color.rgb = RGB_TEXT_DARK

    def add_p(text, bold_prefix=""):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = "Calibri"; r_pre.font.size = Pt(10); r_pre.font.bold = True
            r_pre.font.color.rgb = RGB_PRIMARY_DARK
        r = p.add_run(text)
        r.font.name = "Calibri"; r.font.size = Pt(9.8); r.font.color.rgb = RGB_TEXT_DARK
        return p

    # ----------------------------------------------------
    # DOCUMENT COVER HEADER
    # ----------------------------------------------------
    p_inst = doc.add_paragraph()
    p_inst.paragraph_format.space_before = Pt(0); p_inst.paragraph_format.space_after = Pt(2)
    r_inst = p_inst.add_run("VIDYAVARDHINI’S COLLEGE OF SCIENCE (VARTAK COLLEGE), VASAI ROAD (W)\n")
    r_inst.font.name = "Calibri"; r_inst.font.size = Pt(10.5); r_inst.font.bold = True; r_inst.font.color.rgb = RGB_PRIMARY_MED
    r_sub = p_inst.add_run("Department of Computer Science  •  Affiliated with University of Mumbai  •  Academic Year 2026–2027")
    r_sub.font.name = "Calibri"; r_sub.font.size = Pt(9); r_sub.font.color.rgb = RGB_TEXT_MUTED

    p_doc_title = doc.add_paragraph()
    p_doc_title.paragraph_format.space_before = Pt(8); p_doc_title.paragraph_format.space_after = Pt(3)
    r_dt = p_doc_title.add_run("COMMUNITY ENGAGEMENT PROJECT (NEP 2020)\nPROJECT OVERVIEW & EVALUATION HANDBOOK")
    r_dt.font.name = "Georgia"; r_dt.font.size = Pt(16); r_dt.font.bold = True; r_dt.font.color.rgb = RGB_PRIMARY_DARK

    # Project Metadata Box
    tbl_meta = doc.add_table(rows=2, cols=2)
    tbl_meta.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_meta.autofit = False
    tbl_meta.columns[0].width = Inches(3.33)
    tbl_meta.columns[1].width = Inches(3.34)
    for row in tbl_meta.rows:
        for cell in row.cells:
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
            set_cell_borders(cell, top=("single", "4", "CBD5E1"), bottom=("single", "4", "CBD5E1"), left=("single", "4", "CBD5E1"), right=("single", "4", "CBD5E1"))
            set_cell_shading(cell, "F1F5F9")
    
    c0 = tbl_meta.cell(0, 0).paragraphs[0]; c0.paragraph_format.space_after = Pt(0)
    c0.add_run("PROJECT TITLE:\n").font.size = Pt(8.5); c0.runs[0].font.bold = True; c0.runs[0].font.color.rgb = RGB_PRIMARY_MED
    c0.add_run("Local Biodiversity Mapping of Byculla Zoo & Surrounding Green Spaces").font.size = Pt(9.5); c0.runs[1].font.bold = True

    c1 = tbl_meta.cell(0, 1).paragraphs[0]; c1.paragraph_format.space_after = Pt(0)
    c1.add_run("INTERNAL GUIDE / MENTOR:\n").font.size = Pt(8.5); c1.runs[0].font.bold = True; c1.runs[0].font.color.rgb = RGB_PRIMARY_MED
    c1.add_run("Prof. Shweta Iyer\nDept. of Computer Science").font.size = Pt(9.5); c1.runs[1].font.bold = True

    c2 = tbl_meta.cell(1, 0).paragraphs[0]; c2.paragraph_format.space_after = Pt(0)
    c2.add_run("CLASS & WEIGHTAGE:\n").font.size = Pt(8.5); c2.runs[0].font.bold = True; c2.runs[0].font.color.rgb = RGB_PRIMARY_MED
    c2.add_run("T.Y. B.Sc. CS — Semester V (NEP 2020: 2 Credits)").font.size = Pt(9.5)

    c3 = tbl_meta.cell(1, 1).paragraphs[0]; c3.paragraph_format.space_after = Pt(0)
    c3.add_run("DEPLOYED WEB-GIS APPLICATION:\n").font.size = Pt(8.5); c3.runs[0].font.bold = True; c3.runs[0].font.color.rgb = RGB_PRIMARY_MED
    c3.add_run("https://biomap-byculla.netlify.app").font.size = Pt(9.5); c3.runs[1].font.bold = True; c3.runs[1].font.color.rgb = RGBColor(30, 64, 175)

    # ----------------------------------------------------
    # SECTION 1: THE 30-SECOND ELEVATOR PITCH
    # ----------------------------------------------------
    add_sec_title("1. What is Our Project About? (The Core Summary)", "🎯")
    add_callout(
        "The 30-Second Elevator Pitch",
        "Our project is an empirical ecological audit and digital mapping platform for the free-living, non-captive biodiversity of Mumbai's 53-acre Byculla Zoo (Veermata Jijabai Bhosale Udyan / Ranibagh). Because over 78% of visitors focus exclusively on caged zoo exhibits (like penguins and tigers), the park's rich botanical heritage (280+ tree species) and wild avifauna are unmapped and overlooked.\n\n"
        "To solve this, we collected on-site GPS coordinates, quantified a 26.3 dBA acoustic attenuation shield delivered by the canopy, surveyed 84 community members, and engineered 'BioMap' (https://biomap-byculla.netlify.app) — a zero-cost, open-source mobile Web-GIS platform providing real-time GPS tracking and bilingual (Marathi/English) ecological profiles.",
        bg_hex=HEX_PRIMARY_LIGHT, border_hex=HEX_PRIMARY_DARK, icon="⭐"
    )

    # ----------------------------------------------------
    # SECTION 2: THE PROBLEM STATEMENT
    # ----------------------------------------------------
    add_sec_title("2. The Problem Statement & Civic Motivation", "⚠️")
    add_p(" Established in 1861, Ranibagh is South Mumbai's largest contiguous botanical sanctuary, harboring century-old trees and native birdlife.", "1. The 53-Acre Heritage Sanctuary: ")
    add_p(" Public attention is heavily skewed towards caged animals. Botanical trees and wild avifauna receive zero interpretation on physical placards.", "2. The 'Captive-Species Bias': ")
    add_p(" Municipal contractors routinely lop branches prior to the monsoon, destroying active bird nesting cavities without prior ecological audits.", "3. Arbitrary Pre-Monsoon Tree Lopping: ")
    add_p(" Heavy vehicular traffic along the perimeter (Dr. Ambedkar Road) regularly exceeds 78 dBA, creating severe acoustic stress for vocalizing avifauna.", "4. Acoustic Pollution Vulnerability: ")
    add_p(" No interactive digital map or vernacular mobile guide existed for citizens to explore native flora and fauna on their smartphones.", "5. The Digital Literacy Void: ")

    # ----------------------------------------------------
    # SECTION 3: 4-ZONE METHODOLOGY & INSTRUMENTATION
    # ----------------------------------------------------
    add_sec_title("3. Study Area & 4-Zone Ecological Stratification", "🗺️")
    add_p("We stratified the 53-acre park into 4 distinct micro-habitats to audit ecological variation and sound levels:")

    tbl_z = doc.add_table(rows=5, cols=4)
    tbl_z.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_z.autofit = False
    tbl_z.columns[0].width = Inches(1.2)
    tbl_z.columns[1].width = Inches(1.8)
    tbl_z.columns[2].width = Inches(1.8)
    tbl_z.columns[3].width = Inches(1.87)

    headers_z = ["Zone", "Habitat Character", "GPS Coordinates", "Dominant Species Audited"]
    for c_i, h in enumerate(headers_z):
        cell = tbl_z.cell(0, c_i)
        set_cell_shading(cell, HEX_PRIMARY_DARK)
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.font.name = "Calibri"; r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = RGBColor(255, 255, 255)

    z_rows = [
        ("Zone 1", "Botanical Grove (Heritage Arboretum)", "18°58'38\"N, 72°49'51\"E", "African Baobab, Ceylon Ebony, Coppersmith Barbet (Noise: 52.1 dBA)"),
        ("Zone 2", "Freshwater Wetland & Lotus Lake", "18°58'42\"N, 72°49'56\"E", "Black-crowned Night Heron, Common Moorhen, Odonata (Noise: 54.8 dBA)"),
        ("Zone 3", "Promenade Shrublands & Butterfly Lawns", "18°58'35\"N, 72°49'48\"E", "Common Mormon Butterfly, Bluebottle, Malabar Nut/Adulsa (Noise: 61.2 dBA)"),
        ("Zone 4", "Perimeter Transition Buffer Wall", "18°58'30\"N, 72°49'44\"E", "Black Kite, House Crow, Common Myna (Noise: 78.4 dBA — High Stress)")
    ]
    for r_i, (c0, c1, c2, c3) in enumerate(z_rows, start=1):
        bg = HEX_PRIMARY_LIGHT if r_i == 1 else ("FEE2E2" if r_i == 4 else "FFFFFF")
        for c_i, val in enumerate([c0, c1, c2, c3]):
            cell = tbl_z.cell(r_i, c_i)
            set_cell_shading(cell, bg)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            set_cell_borders(cell, top=("single", "4", "CBD5E1"), bottom=("single", "4", "CBD5E1"), left=("single", "4", "CBD5E1"), right=("single", "4", "CBD5E1"))
            p = cell.paragraphs[0]
            if c_i == 0: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val); r.font.name = "Calibri"; r.font.size = Pt(8.5)
            if c_i == 0: r.font.bold = True

    add_p("Garmin eTrex GPS & Mobile GNSS (±3m accuracy), Digital Sound Level Meter (Type 2, A-weighting, 15-min intervals), Nikon 70-300mm telephoto optics, and fixed 100m transect census walks (06:30–09:30 AM).", "Instrumentation Stack: ")

    # ----------------------------------------------------
    # SECTION 4: KEY SCIENTIFIC DISCOVERIES
    # ----------------------------------------------------
    add_sec_title("4. Key Scientific & Quantitative Discoveries", "🔬")
    
    add_callout(
        "Headline Finding 1: The '26.3 dBA Acoustic Shield' Effect",
        "• Ambient traffic noise at the outer boundary (Zone 4) reaches 78.4 dBA.\n"
        "• Deep inside the Botanical Grove (Zone 1), sound levels plummet to 52.1 dBA.\n"
        "• The multi-tier arboreal tree canopy acts as a massive acoustic sponge, absorbing 26.3 dBA of noise. This creates a critical sonic refuge allowing birds to vocalize, mate, and nest without acoustic interference.",
        bg_hex="FEF3C7", border_hex="D97706", icon="🔊"
    )

    add_callout(
        "Headline Finding 2: 81% Higher Shannon-Wiener Diversity in Core Grove",
        "• We calculated the Shannon-Wiener Diversity Index: H' = - Σ (pi × ln pi) and Pielou's Evenness J' = H' / ln(S).\n"
        "• Zone 1 achieved H' = 2.68 (24 taxa, J' = 0.84), compared to Zone 4 which had H' = 1.48 (only 6 edge-tolerant taxa, J' = 0.65).\n"
        "• Validates that canopy continuity and aquatic features are indispensable for urban ecological resilience.",
        bg_hex=HEX_PRIMARY_LIGHT, border_hex=HEX_PRIMARY_DARK, icon="📐"
    )

    # ----------------------------------------------------
    # SECTION 5: COMMUNITY SURVEY FINDINGS (N = 84)
    # ----------------------------------------------------
    add_sec_title("5. Community Survey & Perception Analysis (N = 84 Citizens)", "👥")
    add_p("In strict alignment with NEP 2020 social engagement criteria, we surveyed 84 diverse stakeholders (morning walkers, school teachers, students, tourists, and BMC caretakers):")
    add_p("56.5% of free-living wildlife encounters occur strictly in the inner Botanical Grove and Lake.", "• Habitat Clustering (Q3): ")
    add_p("36.5% cite pre-monsoon tree trimming, followed by 27.0% citing perimeter traffic noise.", "• Greatest Perceived Threat (Q4): ")
    add_p("72.0% cite the absence of digital maps and signage restricted to cages as their main barrier.", "• Biodiversity Literacy Deficit (Q5): ")
    add_p("86.7% confirmed an interactive mobile web app would enhance their visit, and 83.3% stated they would actively use it.", "• Validation of Digital Solution (Q6 & Q7): ")
    add_p("87.0% consider QR-code placards on heritage trees linking to species profiles highly beneficial.", "• Demand for Tree QR Codes (Q10): ")

    # ----------------------------------------------------
    # SECTION 6: BIOMAP WEB-GIS PLATFORM
    # ----------------------------------------------------
    add_sec_title("6. Technological Solution: BioMap Web-GIS Platform", "💻")
    add_p("To translate our research into a permanent digital public good, we built and deployed BioMap:")
    add_p("Fully deployed and accessible on mobile browsers at https://biomap-byculla.netlify.app with zero download friction.", "• Live URL & Architecture: ")
    add_p("HTML5 Geolocation API displays the visitor's real-time position on 53 acres of garden trails.", "• Real-Time Mobile GPS Tracking: ")
    add_p("Semi-transparent GeoJSON polygon boundaries communicating microclimate, sound levels, and species counts.", "• Interactive 4-Zone Overlays: ")
    add_p("Profiles include Linnaean hierarchy, IUCN status, flowering seasons, and Marathi vernacular descriptions (जैविविधता माहिती).", "• Bilingual Ecological Cards: ")
    add_p("Leaflet.js + OpenStreetMap + Vanilla CSS3 — zero ongoing database or municipal server costs, making it easily replicable for other municipal parks.", "• Zero-Cost Open-Source Stack: ")

    # ----------------------------------------------------
    # SECTION 7: THREE-TIER MUNICIPAL RECOMMENDATIONS
    # ----------------------------------------------------
    add_sec_title("7. Actionable Recommendations (Three-Tier Blueprint)", "🏛️")
    add_p("Laser-etched QR code placards on 42 audited heritage trees; mandatory pre-pruning bird-nest inspections between February and June; floating reed resting islands in Lotus Lake.", "• Tier 1: Park Level (Immediate / 1–3 Months): ")
    add_p("Integrate BioMap's GIS dataset into BMC's statutory People's Biodiversity Register (PBR); plant a 50m dense acoustic green buffer along Dr. Ambedkar Road.", "• Tier 2: BMC Municipal Level (Policy / 6–12 Months): ")
    add_p("Launch university Bio-Blitz weekends for CS/IT and science undergraduates; enable citizen sighting submissions with botanist verification.", "• Tier 3: Citizen Science & Academia (Ongoing): ")

    # ----------------------------------------------------
    # SECTION 8: WHAT EXAMINERS FOCUS ON & VIVA CHEAT SHEET
    # ----------------------------------------------------
    add_sec_title("8. What Evaluators Focus On & Viva Defense Cheat Sheet", "🎓")
    
    add_callout(
        "5 Key Criteria Evaluators Grade in Your Report PDF",
        "1. Academic Structure: Title page, certificate signed by Prof. Shweta Iyer, declaration, 5 NEP chapters, references, diary.\n"
        "2. Interdisciplinary Rigor: Bridging Computer Science (Web-GIS) with municipal urban ecology (not just basic coding).\n"
        "3. Field Authenticity: On-site GPS coordinates, geotagged species photos, and 42+ taxa census.\n"
        "4. Community Evidence: N = 84 citizen surveys with quantitative charts and actionable findings.\n"
        "5. Measurable Findings: 26.3 dBA acoustic damping, Shannon Index (H' = 2.68), and a deployed live website.",
        bg_hex=HEX_PRIMARY_LIGHT, border_hex=HEX_PRIMARY_DARK, icon="✅"
    )

    tbl_viva = doc.add_table(rows=6, cols=2)
    tbl_viva.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_viva.autofit = False
    tbl_viva.columns[0].width = Inches(2.2)
    tbl_viva.columns[1].width = Inches(4.47)

    viva_qna = [
        ("Viva Question", "Winning 1-Minute Model Answer"),
        ("Why did you choose Byculla Zoo for a CS project?",
         "Byculla Zoo is South Mumbai's largest botanical sanctuary (53 acres, 280+ tree species), but public awareness is trapped in a 'captive-species bias' focusing only on caged animals. We applied Computer Science & GIS to map the neglected free-living wildlife and provide an accessible digital public tool."),
        ("What is your project's biggest scientific finding?",
         "The multi-tier botanical canopy acts as an acoustic shield, dampening perimeter city traffic noise from 78.4 dBA down to 52.1 dBA (a 26.3 dBA reduction), preserving an 81% higher Shannon biodiversity index (H' = 2.68 vs. 1.48) in the park center."),
        ("What technology did you use for BioMap?",
         "We used Leaflet.js, OpenStreetMap, HTML5 Geolocation, and GeoJSON polygon layers deployed on Netlify. It requires zero server licensing costs and runs instantly on any smartphone browser without app installation."),
        ("How did you incorporate community engagement?",
         "We surveyed 84 diverse citizens (morning walkers, students, caretakers). 86.7% confirmed the need for a digital guide, and their feedback shaped our bilingual (Marathi/English) interface and QR-code recommendation."),
        ("Who guided your research?",
         "The project was mentored and supervised by Prof. Shweta Iyer, Department of Computer Science, Vartak College, through 10 formal supervisory meetings documented in our Guide Interaction Diary.")
    ]

    for r_i, (q, a) in enumerate(viva_qna):
        bg = HEX_PRIMARY_DARK if r_i == 0 else ("F1F5F9" if r_i % 2 == 1 else "FFFFFF")
        for c_i, txt in enumerate([q, a]):
            cell = tbl_viva.cell(r_i, c_i)
            set_cell_shading(cell, bg)
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            set_cell_borders(cell, top=("single", "4", "CBD5E1"), bottom=("single", "4", "CBD5E1"), left=("single", "4", "CBD5E1"), right=("single", "4", "CBD5E1"))
            p = cell.paragraphs[0]
            r = p.add_run(txt)
            r.font.name = "Calibri"; r.font.size = Pt(8.5 if r_i > 0 else 9.5)
            if r_i == 0:
                r.font.bold = True; r.font.color.rgb = RGBColor(255, 255, 255)
            elif c_i == 0:
                r.font.bold = True; r.font.color.rgb = RGB_PRIMARY_DARK

    out_docx = "Byculla Zoo CEP - Project Overview & Viva Handbook.docx"
    doc.save(out_docx)
    print(f"Word document saved successfully as '{out_docx}'!")
    return out_docx

def export_doc_to_pdf(docx_path, pdf_path):
    abs_docx = os.path.abspath(docx_path)
    abs_pdf = os.path.abspath(pdf_path)
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(abs_docx)
        doc.SaveAs(abs_pdf, FileFormat=17) # 17 = wdFormatPDF
        doc.Close(False)
        print(f"PDF exported successfully as '{pdf_path}'!")
    finally:
        word.Quit()

if __name__ == "__main__":
    docx_file = build_summary_doc()
    pdf_file = "Byculla Zoo CEP - Project Overview & Viva Handbook.pdf"
    try:
        export_doc_to_pdf(docx_file, pdf_file)
    except Exception as e:
        print("Error during Word PDF export:", e)
