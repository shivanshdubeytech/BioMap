import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# ----------------------------------------------------
# Color Palette (Deep Forest Emerald & Slate Theme)
# ----------------------------------------------------
C_PRIMARY_DARK  = RGBColor(27, 67, 50)     # #1B4332 Deep Forest Green
C_PRIMARY_MED   = RGBColor(45, 106, 79)    # #2D6A4F Emerald Green
C_PRIMARY_LIGHT = RGBColor(216, 243, 220)  # #D8F3DC Mint Tint
C_ACCENT_AMBER  = RGBColor(217, 119, 6)    # #D97706 Warm Amber
C_ACCENT_BLUE   = RGBColor(30, 64, 175)    # #1E40AF Ocean Slate
C_ACCENT_RED    = RGBColor(220, 38, 38)    # #DC2626 Coral Red

C_BG_PAGE       = RGBColor(248, 250, 252)  # #F8FAFC Off-white
C_CARD_BG       = RGBColor(255, 255, 255)  # #FFFFFF Pure White
C_CARD_MINT     = RGBColor(240, 253, 244)  # #F0FDF4 Soft Mint
C_CARD_SLATE    = RGBColor(241, 245, 249)  # #F1F5F9 Soft Slate
C_CARD_BORDER   = RGBColor(226, 232, 240)  # #E2E8F0 Border
C_BORDER_GREEN  = RGBColor(167, 243, 208)  # #A7F3D0 Green Border

C_TEXT_DARK     = RGBColor(26, 32, 44)     # #1A202C Charcoal Text
C_TEXT_MUTED    = RGBColor(100, 116, 139)  # #64748B Secondary Text
C_TEXT_LIGHT    = RGBColor(255, 255, 255)  # White

FONT_HEADING = "Georgia"
FONT_BODY    = "Calibri"

def create_presentation():
    prs = Presentation()
    # 16:9 Widescreen (13.333 x 7.5 inches)
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6] # Blank layout

    # Helper: Add base background
    def add_page_bg(slide, bg_color=C_BG_PAGE):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid(); bg.fill.fore_color.rgb = bg_color
        bg.line.fill.background() # No border
        return bg

    # Helper: Add Slide Header
    def add_header(slide, title, category="COMMUNITY ENGAGEMENT PROJECT (NEP 2020)", slide_num=None):
        add_page_bg(slide)
        
        # Top thin decorative stripe
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.1))
        top_bar.fill.solid(); top_bar.fill.fore_color.rgb = C_PRIMARY_DARK
        top_bar.line.fill.background()

        # Category pill
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.28), Inches(9.0), Inches(0.35))
        tf_c = cat_box.text_frame; tf_c.word_wrap = True; tf_c.margin_left = tf_c.margin_top = 0
        p_c = tf_c.paragraphs[0]
        r_c = p_c.add_run()
        r_c.text = category.upper()
        r_c.font.name = FONT_BODY; r_c.font.size = Pt(10); r_c.font.bold = True
        r_c.font.color.rgb = C_PRIMARY_MED

        # Main Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(10.5), Inches(0.65))
        tf_t = title_box.text_frame; tf_t.word_wrap = True; tf_t.margin_left = tf_t.margin_top = 0
        p_t = tf_t.paragraphs[0]
        r_t = p_t.add_run()
        r_t.text = title
        r_t.font.name = FONT_HEADING; r_t.font.size = Pt(22); r_t.font.bold = True
        r_t.font.color.rgb = C_PRIMARY_DARK

        # Slide Number Badge (Top Right)
        if slide_num:
            num_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(12.0), Inches(0.35), Inches(0.7), Inches(0.35))
            num_box.fill.solid(); num_box.fill.fore_color.rgb = C_PRIMARY_LIGHT
            num_box.line.color.rgb = C_BORDER_GREEN; num_box.line.width = Pt(1)
            tf_n = num_box.text_frame; tf_n.margin_top = Inches(0.04)
            p_n = tf_n.paragraphs[0]; p_n.alignment = PP_ALIGN.CENTER
            r_n = p_n.add_run(); r_n.text = f"{slide_num:02d}"
            r_n.font.name = FONT_BODY; r_n.font.size = Pt(11); r_n.font.bold = True
            r_n.font.color.rgb = C_PRIMARY_DARK

        # Bottom footer bar
        bot_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(7.2), Inches(13.333), Inches(0.3))
        bot_bar.fill.solid(); bot_bar.fill.fore_color.rgb = RGBColor(255, 255, 255)
        bot_bar.line.color.rgb = C_CARD_BORDER; bot_bar.line.width = Pt(0.5)
        
        f_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.22), Inches(11.7), Inches(0.25))
        tf_f = f_box.text_frame; tf_f.margin_left = tf_f.margin_top = 0
        p_f = tf_f.paragraphs[0]
        r_f1 = p_f.add_run(); r_f1.text = "Vidyavardhini's College of Science (Vartak College)  |  T.Y. B.Sc. CS  |  Mentor: "
        r_f1.font.name = FONT_BODY; r_f1.font.size = Pt(9); r_f1.font.color.rgb = C_TEXT_MUTED
        r_f2 = p_f.add_run(); r_f2.text = "Prof. Shweta Iyer  "
        r_f2.font.name = FONT_BODY; r_f2.font.size = Pt(9); r_f2.font.bold = True; r_f2.font.color.rgb = C_PRIMARY_DARK
        r_f3 = p_f.add_run(); r_f3.text = "|  Byculla Zoo Urban Biodiversity Mapping (BioMap)"
        r_f3.font.name = FONT_BODY; r_f3.font.size = Pt(9); r_f3.font.color.rgb = C_TEXT_MUTED

    # Helper: Add Styled Card Container
    def add_card(slide, left, top, width, height, bg_color=C_CARD_BG, border_color=C_CARD_BORDER, corner_radius=None):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid(); card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color; card.line.width = Pt(1)
        return card

    # ====================================================
    # SLIDE 1: Title Slide (Grand Emerald Presentation Cover)
    # ====================================================
    s1 = prs.slides.add_slide(blank_layout)
    add_page_bg(s1, RGBColor(244, 247, 246))

    # Emerald Header Backdrop
    top_decor = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(4.3))
    top_decor.fill.solid(); top_decor.fill.fore_color.rgb = C_PRIMARY_DARK
    top_decor.line.fill.background()

    # Decorative bottom accent stripe
    accent_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(4.22), Inches(13.333), Inches(0.12))
    accent_bar.fill.solid(); accent_bar.fill.fore_color.rgb = C_ACCENT_AMBER
    accent_bar.line.fill.background()

    # College Logo
    if os.path.exists("college_logo.jpg"):
        logo = s1.shapes.add_picture("college_logo.jpg", Inches(0.8), Inches(0.55), width=Inches(1.2))

    # Institutional Header
    inst_box = s1.shapes.add_textbox(Inches(2.2), Inches(0.55), Inches(10.3), Inches(0.8))
    tf_i = inst_box.text_frame; tf_i.word_wrap = True; tf_i.margin_left = tf_i.margin_top = 0
    p_i1 = tf_i.paragraphs[0]
    r_i1 = p_i1.add_run()
    r_i1.text = "VIDYAVARDHINI’S A. V. COLLEGE OF ARTS, K. M. COLLEGE OF COMMERCE, E. S. A. COLLEGE OF SCIENCE"
    r_i1.font.name = FONT_BODY; r_i1.font.size = Pt(10.5); r_i1.font.bold = True
    r_i1.font.color.rgb = RGBColor(167, 243, 208)

    p_i2 = tf_i.add_paragraph()
    r_i2 = p_i2.add_run()
    r_i2.text = "Department of Computer Science  •  Affiliated with University of Mumbai  •  Vasai Road (West)"
    r_i2.font.name = FONT_BODY; r_i2.font.size = Pt(9.5)
    r_i2.font.color.rgb = RGBColor(209, 250, 229)

    # Main Project Title
    title_box1 = s1.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.7), Inches(1.7))
    tf_t1 = title_box1.text_frame; tf_t1.word_wrap = True; tf_t1.margin_left = tf_t1.margin_top = 0
    
    p_t1 = tf_t1.paragraphs[0]
    r_t1 = p_t1.add_run()
    r_t1.text = "LOCAL BIODIVERSITY MAPPING OF BYCULLA ZOO\nAND SURROUNDING GREEN SPACES"
    r_t1.font.name = FONT_HEADING; r_t1.font.size = Pt(27); r_t1.font.bold = True
    r_t1.font.color.rgb = C_TEXT_LIGHT

    p_sub1 = tf_t1.add_paragraph()
    p_sub1.space_before = Pt(8)
    r_sub1 = p_sub1.add_run()
    r_sub1.text = "Empirical Ecological Audit, Spatial Zonation & BioMap Web-GIS Public Awareness Platform"
    r_sub1.font.name = FONT_BODY; r_sub1.font.size = Pt(13); r_sub1.font.bold = True
    r_sub1.font.color.rgb = RGBColor(253, 224, 71) # Warm Yellow

    # Badges Row (NEP 2020, 2 Credits, Academic Year)
    badges_y = Inches(3.4)
    badges = [
        ("NATIONAL EDUCATION POLICY (NEP 2020)", C_PRIMARY_LIGHT, C_PRIMARY_DARK),
        ("COMMUNITY ENGAGEMENT PROJECT (2 CREDITS)", RGBColor(254, 243, 199), RGBColor(180, 83, 9)),
        ("T.Y. B.Sc. CS — SEMESTER V (2026–2027)", RGBColor(224, 242, 254), RGBColor(3, 105, 161))
    ]
    bx = Inches(0.8)
    for b_txt, b_bg, b_fg in badges:
        bw = Inches(len(b_txt) * 0.088 + 0.4)
        b_shp = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, badges_y, bw, Inches(0.36))
        b_shp.fill.solid(); b_shp.fill.fore_color.rgb = b_bg
        b_shp.line.color.rgb = b_fg; b_shp.line.width = Pt(1)
        tf_b = b_shp.text_frame; tf_b.margin_top = Inches(0.04)
        p_b = tf_b.paragraphs[0]; p_b.alignment = PP_ALIGN.CENTER
        r_b = p_b.add_run(); r_b.text = b_txt
        r_b.font.name = FONT_BODY; r_b.font.size = Pt(9.5); r_b.font.bold = True; r_b.font.color.rgb = b_fg
        bx += bw + Inches(0.2)

    # Bottom Half Info Cards
    card1 = add_card(s1, Inches(0.8), Inches(4.7), Inches(5.6), Inches(2.2), bg_color=C_CARD_BG, border_color=C_CARD_BORDER)
    tf_c1 = card1.text_frame; tf_c1.word_wrap = True; tf_c1.margin_left = tf_c1.margin_top = Inches(0.2)
    p_c1 = tf_c1.paragraphs[0]
    p_c1.add_run().text = "STUDENT INVESTIGATORS\n"
    p_c1.runs[0].font.name = FONT_BODY; p_c1.runs[0].font.size = Pt(10); p_c1.runs[0].font.bold = True
    p_c1.runs[0].font.color.rgb = C_PRIMARY_MED
    
    p_s_det = tf_c1.add_paragraph()
    p_s_det.space_before = Pt(4)
    r_sd = p_s_det.add_run()
    r_sd.text = "Student Name: Riddhesh & Project Team\nProgram: Bachelor of Science in Computer Science (T.Y. B.Sc. CS)\nInstitution: Vartak College, Vasai Road (W)\nAffiliation: University of Mumbai (Academic Year 2026–2027)"
    r_sd.font.name = FONT_BODY; r_sd.font.size = Pt(11); r_sd.font.color.rgb = C_TEXT_DARK

    card2 = add_card(s1, Inches(6.9), Inches(4.7), Inches(5.6), Inches(2.2), bg_color=C_CARD_MINT, border_color=C_BORDER_GREEN)
    tf_c2 = card2.text_frame; tf_c2.word_wrap = True; tf_c2.margin_left = tf_c2.margin_top = Inches(0.2)
    p_c2 = tf_c2.paragraphs[0]
    p_c2.add_run().text = "PROJECT GUIDANCE & SUPERVISION\n"
    p_c2.runs[0].font.name = FONT_BODY; p_c2.runs[0].font.size = Pt(10); p_c2.runs[0].font.bold = True
    p_c2.runs[0].font.color.rgb = C_PRIMARY_DARK
    
    p_g_det = tf_c2.add_paragraph()
    p_g_det.space_before = Pt(4)
    r_gd1 = p_g_det.add_run()
    r_gd1.text = "Internal Project Guide / Mentor:\n"
    r_gd1.font.name = FONT_BODY; r_gd1.font.size = Pt(11); r_gd1.font.color.rgb = C_TEXT_MUTED
    
    r_gd2 = p_g_det.add_run()
    r_gd2.text = "Prof. Shweta Iyer\n"
    r_gd2.font.name = FONT_HEADING; r_gd2.font.size = Pt(15); r_gd2.font.bold = True
    r_gd2.font.color.rgb = C_PRIMARY_DARK
    
    r_gd3 = p_g_det.add_run()
    r_gd3.text = "Department of Computer Science\nVidyavardhini's College of Science, Vasai Road"
    r_gd3.font.name = FONT_BODY; r_gd3.font.size = Pt(10.5); r_gd3.font.color.rgb = C_TEXT_DARK

    # ====================================================
    # SLIDE 2: Background & Project Motivation
    # ====================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Background & Project Motivation: Why Byculla Zoo?", slide_num=2)

    # 3 Pillar Cards on Top
    col_w = Inches(3.68)
    col_gap = Inches(0.33)
    c_y = Inches(1.3)
    c_h = Inches(3.1)

    # Pillar 1
    p1 = add_card(s2, Inches(0.8), c_y, col_w, c_h, bg_color=C_CARD_BG)
    tf_p1 = p1.text_frame; tf_p1.word_wrap = True; tf_p1.margin_left = tf_p1.margin_right = tf_p1.margin_top = Inches(0.2)
    p_p1_t = tf_p1.paragraphs[0]; p_p1_t.add_run().text = "🏛️ 53-Acre Botanical Heritage"
    p_p1_t.runs[0].font.name = FONT_HEADING; p_p1_t.runs[0].font.size = Pt(14); p_p1_t.runs[0].font.bold = True
    p_p1_t.runs[0].font.color.rgb = C_PRIMARY_DARK
    p_p1_b = tf_p1.add_paragraph(); p_p1_b.space_before = Pt(6)
    r_p1_b = p_p1_b.add_run()
    r_p1_b.text = "• Established in 1861 as Veermata Jijabai Bhosale Udyan (Ranibagh).\n• South Mumbai's largest contiguous botanical sanctuary and public green lung.\n• Home to over 280+ botanical tree species, including heritage Baobabs and rare flora."
    r_p1_b.font.name = FONT_BODY; r_p1_b.font.size = Pt(11); r_p1_b.font.color.rgb = C_TEXT_DARK

    # Pillar 2
    p2 = add_card(s2, Inches(0.8) + col_w + col_gap, c_y, col_w, c_h, bg_color=RGBColor(254, 242, 242), border_color=RGBColor(254, 202, 202))
    tf_p2 = p2.text_frame; tf_p2.word_wrap = True; tf_p2.margin_left = tf_p2.margin_right = tf_p2.margin_top = Inches(0.2)
    p_p2_t = tf_p2.paragraphs[0]; p_p2_t.add_run().text = "⚠️ The 'Captive-Species Bias'"
    p_p2_t.runs[0].font.name = FONT_HEADING; p_p2_t.runs[0].font.size = Pt(14); p_p2_t.runs[0].font.bold = True
    p_p2_t.runs[0].font.color.rgb = C_ACCENT_RED
    p_p2_b = tf_p2.add_paragraph(); p_p2_b.space_before = Pt(6)
    r_p2_b = p_p2_b.add_run()
    r_p2_b.text = "• Extreme Public Bias: 78%+ of visitors focus strictly on caged animals (Penguins, Tigers).\n• Non-captive free-living wildlife (avifauna, butterflies, heritage flora) goes completely unmapped.\n• Zero digital geospatial tools exist for citizens to explore native flora."
    r_p2_b.font.name = FONT_BODY; r_p2_b.font.size = Pt(11); r_p2_b.font.color.rgb = C_TEXT_DARK

    # Pillar 3
    p3 = add_card(s2, Inches(0.8) + (col_w + col_gap)*2, c_y, col_w, c_h, bg_color=C_CARD_MINT, border_color=C_BORDER_GREEN)
    tf_p3 = p3.text_frame; tf_p3.word_wrap = True; tf_p3.margin_left = tf_p3.margin_right = tf_p3.margin_top = Inches(0.2)
    p_p3_t = tf_p3.paragraphs[0]; p_p3_t.add_run().text = "🌳 Urban Ecological Refugium"
    p_p3_t.runs[0].font.name = FONT_HEADING; p_p3_t.runs[0].font.size = Pt(14); p_p3_t.runs[0].font.bold = True
    p_p3_t.runs[0].font.color.rgb = C_PRIMARY_DARK
    p_p3_b = tf_p3.add_paragraph(); p_p3_b.space_before = Pt(6)
    r_p3_b = p_p3_b.add_run()
    r_p3_b.text = "• Vital refuge amidst hyper-dense South Mumbai (Dr. Ambedkar Road corridor).\n• Significant microclimatic cooling: Up to 3.2°C cooler than surrounding concrete avenues.\n• Crucial roosting for migratory avifauna and pollinator species."
    r_p3_b.font.name = FONT_BODY; r_p3_b.font.size = Pt(11); r_p3_b.font.color.rgb = C_TEXT_DARK

    # Bottom Banner: Community & Technological Need
    bot_card = add_card(s2, Inches(0.8), Inches(4.6), Inches(11.7), Inches(2.2), bg_color=C_CARD_BG)
    tf_bc = bot_card.text_frame; tf_bc.word_wrap = True; tf_bc.margin_left = tf_bc.margin_right = tf_bc.margin_top = Inches(0.2)
    p_bc_t = tf_bc.paragraphs[0]; p_bc_t.add_run().text = "THE CIVIC & RESEARCH CHALLENGE IDENTIFIED"
    p_bc_t.runs[0].font.name = FONT_BODY; p_bc_t.runs[0].font.size = Pt(11); p_bc_t.runs[0].font.bold = True
    p_bc_t.runs[0].font.color.rgb = C_PRIMARY_MED

    p_bc_b = tf_bc.add_paragraph(); p_bc_b.space_before = Pt(4)
    r_bc = p_bc_b.add_run()
    r_bc.text = "1. Information Deficit: Park visitors lack vernacular (Marathi/English) guides on botanical species and avifauna.\n" \
                "2. Anthropogenic Pressure: Arbitrary pre-monsoon tree pruning by municipal contractors destroys active bird nesting cavities.\n" \
                "3. Acoustic Vulnerability: Heavy perimeter traffic (78+ dBA) penetrates outer corridors, threatening sensitive birds.\n" \
                "4. Technological Gap: No open-source, mobile-friendly Web-GIS platform currently exists for citizen-led biodiversity discovery in Mumbai's municipal parks."
    r_bc.font.name = FONT_BODY; r_bc.font.size = Pt(11); r_bc.font.color.rgb = C_TEXT_DARK

    # ====================================================
    # SLIDE 3: Aims, Objectives & Scope
    # ====================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "Aims, Objectives & Project Scope (NEP 2020)", slide_num=3)

    # Core Aim Card on Top
    aim_card = add_card(s3, Inches(0.8), Inches(1.3), Inches(11.7), Inches(1.2), bg_color=C_CARD_MINT, border_color=C_BORDER_GREEN)
    tf_aim = aim_card.text_frame; tf_aim.word_wrap = True; tf_aim.margin_left = tf_aim.margin_right = tf_aim.margin_top = Inches(0.18)
    p_aim_t = tf_aim.paragraphs[0]; p_aim_t.add_run().text = "🎯 PRIMARY PROJECT AIM"
    p_aim_t.runs[0].font.name = FONT_BODY; p_aim_t.runs[0].font.size = Pt(11); p_aim_t.runs[0].font.bold = True; p_aim_t.runs[0].font.color.rgb = C_PRIMARY_DARK
    p_aim_b = tf_aim.add_paragraph(); p_aim_b.space_before = Pt(2)
    r_aim = p_aim_b.add_run()
    r_aim.text = "To systematically document, geotag, and evaluate the non-captive biodiversity of Byculla Zoo across distinct ecological zones, and develop a free, open-source Web-GIS platform (BioMap) to foster civic ecological literacy and support municipal conservation."
    r_aim.font.name = FONT_BODY; r_aim.font.size = Pt(12); r_aim.font.bold = True; r_aim.font.color.rgb = C_TEXT_DARK

    # 4 Specific Objectives (2x2 Grid)
    grid_w = Inches(5.7)
    grid_h = Inches(1.9)
    g_x1 = Inches(0.8); g_x2 = Inches(6.8)
    g_y1 = Inches(2.7); g_y2 = Inches(4.8)

    objs = [
        ("01", "Spatial Biodiversity Audit & GPS Geotagging", 
         "Conduct exhaustive on-site transect walks to identify, classify, and record exact GPS coordinates (latitude/longitude) for 40+ free-living taxa across 4 habitat zones.", g_x1, g_y1, C_PRIMARY_DARK),
        ("02", "Microclimatic & Acoustic Attenuation Mapping",
         "Measure ambient acoustic damping (dBA), canopy density, and temperature differentials between perimeter traffic corridors and interior heritage arboretums.", g_x2, g_y1, C_ACCENT_BLUE),
        ("03", "Multi-Stakeholder Community Survey (N = 84)",
         "Administer structured questionnaires to morning walkers, educators, visitors, and botanical caretakers to analyze public perceptions, barriers, and digital platform demand.", g_x1, g_y2, C_ACCENT_AMBER),
        ("04", "BioMap Web-GIS Platform Development & Rollout",
         "Engineer, build, and deploy an open-source, mobile-first Web-GIS portal featuring live geolocation, interactive zone overlays, and bilingual (English/Marathi) species cards.", g_x2, g_y2, C_PRIMARY_MED)
    ]

    for num, title, desc, ox, oy, col in objs:
        oc = add_card(s3, ox, oy, grid_w, grid_h, bg_color=C_CARD_BG)
        tf_oc = oc.text_frame; tf_oc.word_wrap = True; tf_oc.margin_left = tf_oc.margin_right = tf_oc.margin_top = Inches(0.18)
        
        p_num = tf_oc.paragraphs[0]
        r_num = p_num.add_run(); r_num.text = f"OBJECTIVE [{num}]  •  "
        r_num.font.name = FONT_BODY; r_num.font.size = Pt(10); r_num.font.bold = True; r_num.font.color.rgb = col
        r_ot = p_num.add_run(); r_ot.text = title
        r_ot.font.name = FONT_HEADING; r_ot.font.size = Pt(12); r_ot.font.bold = True; r_ot.font.color.rgb = C_TEXT_DARK
        
        p_od = tf_oc.add_paragraph(); p_od.space_before = Pt(4)
        r_od = p_od.add_run(); r_od.text = desc
        r_od.font.name = FONT_BODY; r_od.font.size = Pt(10.5); r_od.font.color.rgb = C_TEXT_MUTED

    # ====================================================
    # SLIDE 4: Study Area & 4-Zone Ecological Methodology
    # ====================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "Study Area & 4-Zone Ecological Stratification", slide_num=4)

    # 4 Zone Cards on the Left (2 columns x 2 rows)
    z_w = Inches(3.7)
    z_h = Inches(2.55)
    
    zones = [
        ("ZONE 1", "Botanical Grove (Heritage Arboretum)", "18°58'38\"N, 72°49'51\"E",
         "• Character: Dense Ficus, Baobab & Mahogany heritage canopy.\n• Species: Coppersmith Barbet, Asian Koel, African Baobab, Ceylon Ebony.\n• Acoustic Level: 52.1 dBA (Quiet interior haven).",
         Inches(0.8), Inches(1.3), C_CARD_MINT, C_BORDER_GREEN, C_PRIMARY_DARK),
        
        ("ZONE 2", "Freshwater Wetland & Lotus Lake", "18°58'42\"N, 72°49'56\"E",
         "• Character: Eutrophic lotus water body & muddy banks.\n• Species: Black-crowned Night Heron, Common Moorhen, Pond Heron, Odonata.\n• Acoustic Level: 54.8 dBA (Moderate water edge).",
         Inches(4.7), Inches(1.3), RGBColor(239, 246, 255), RGBColor(191, 219, 254), C_ACCENT_BLUE),

        ("ZONE 3", "Promenade Shrublands & Lawns", "18°58'35\"N, 72°49'48\"E",
         "• Character: Managed flowering shrubs, butterfly host plants & lawns.\n• Species: Common Mormon Butterfly, Bluebottle, Malabar Nut (Adulsa).\n• Acoustic Level: 61.2 dBA (Visitor path traffic).",
         Inches(0.8), Inches(4.05), RGBColor(254, 243, 199), RGBColor(253, 230, 138), C_ACCENT_AMBER),

        ("ZONE 4", "Perimeter Transition Buffer", "18°58'30\"N, 72°49'44\"E",
         "• Character: Western boundary adjacent to Dr. Ambedkar Road.\n• Species: Black Kite, House Crow, Common Myna (Edge-tolerant).\n• Acoustic Level: 78.4 dBA (Heavy traffic noise impact).",
         Inches(4.7), Inches(4.05), RGBColor(254, 242, 242), RGBColor(254, 202, 202), C_ACCENT_RED)
    ]

    for z_code, z_title, z_gps, z_desc, zx, zy, z_bg, z_bd, z_fg in zones:
        zc = add_card(s4, zx, zy, z_w, z_h, bg_color=z_bg, border_color=z_bd)
        tf_zc = zc.text_frame; tf_zc.word_wrap = True; tf_zc.margin_left = tf_zc.margin_right = tf_zc.margin_top = Inches(0.15)
        
        p_zh = tf_zc.paragraphs[0]
        r_zh1 = p_zh.add_run(); r_zh1.text = f"{z_code}: "
        r_zh1.font.name = FONT_BODY; r_zh1.font.size = Pt(10); r_zh1.font.bold = True; r_zh1.font.color.rgb = z_fg
        r_zh2 = p_zh.add_run(); r_zh2.text = z_title
        r_zh2.font.name = FONT_HEADING; r_zh2.font.size = Pt(11); r_zh2.font.bold = True; r_zh2.font.color.rgb = C_TEXT_DARK
        
        p_gps = tf_zc.add_paragraph(); p_gps.space_before = Pt(2)
        r_gps = p_gps.add_run(); r_gps.text = f"📍 GPS: {z_gps}"
        r_gps.font.name = FONT_BODY; r_gps.font.size = Pt(9); r_gps.font.bold = True; r_gps.font.color.rgb = z_fg
        
        p_zd = tf_zc.add_paragraph(); p_zd.space_before = Pt(3)
        r_zd = p_zd.add_run(); r_zd.text = z_desc
        r_zd.font.name = FONT_BODY; r_zd.font.size = Pt(10); r_zd.font.color.rgb = C_TEXT_DARK

    # Right Side: Instrumentation & Field Evidence Card
    r_card = add_card(s4, Inches(8.6), Inches(1.3), Inches(3.9), Inches(5.3), bg_color=C_CARD_BG)
    tf_rc = r_card.text_frame; tf_rc.word_wrap = True; tf_rc.margin_left = tf_rc.margin_right = tf_rc.margin_top = Inches(0.2)
    p_rt = tf_rc.paragraphs[0]; p_rt.add_run().text = "🔬 INSTRUMENTATION & FIELDWORK"
    p_rt.runs[0].font.name = FONT_BODY; p_rt.runs[0].font.size = Pt(11); p_rt.runs[0].font.bold = True; p_rt.runs[0].font.color.rgb = C_PRIMARY_MED

    p_rb = tf_rc.add_paragraph(); p_rb.space_before = Pt(4)
    r_rb = p_rb.add_run()
    r_rb.text = "• Geospatial Telemetry: Garmin eTrex GPS & Mobile GNSS (±3m spatial accuracy).\n" \
                "• Acoustic Logging: Digital Sound Level Meter (Type 2, A-weighting, 15-min intervals).\n" \
                "• Photodocumentation: Nikon 70-300mm telephoto optics & macro lenses.\n" \
                "• Ecological Sampling: Fixed 100m line transects and 15-min stationary point counts (06:30–09:30 AM)."
    r_rb.font.name = FONT_BODY; r_rb.font.size = Pt(10.5); r_rb.font.color.rgb = C_TEXT_DARK

    # Fieldwork Proof Image
    if os.path.exists("animal_photos/fieldwork_proof.jpg"):
        s4.shapes.add_picture("animal_photos/fieldwork_proof.jpg", Inches(8.8), Inches(3.8), width=Inches(3.5), height=Inches(2.6))

    # ====================================================
    # SLIDE 5: Key Field Observations & Species Catalog
    # ====================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "Key Field Observations: Master Species Audit (42+ Taxa)", slide_num=5)

    # Top Metrics Row (4 Metrics)
    m_w = Inches(2.7)
    m_h = Inches(1.1)
    m_gap = Inches(0.3)
    metrics = [
        ("42+", "DOCUMENTED TAXA", "Flora, Birds, Insects & Herpetofauna", C_PRIMARY_DARK),
        ("26.3 dBA", "ACOUSTIC DAMPING", "Canopy Noise Reduction vs. Perimeter", C_ACCENT_AMBER),
        ("H' = 2.68", "SHANNON INDEX", "Zone 1 Grove Ecological Richness", C_PRIMARY_MED),
        ("N = 84", "COMMUNITY SURVEY", "Multi-stakeholder perception audit", C_ACCENT_BLUE)
    ]
    mx = Inches(0.8)
    for m_val, m_lbl, m_sub, m_col in metrics:
        mc = add_card(s5, mx, Inches(1.3), m_w, m_h, bg_color=C_CARD_BG)
        tf_m = mc.text_frame; tf_m.word_wrap = True; tf_m.margin_left = tf_m.margin_top = Inches(0.12)
        p_mv = tf_m.paragraphs[0]; r_mv = p_mv.add_run(); r_mv.text = m_val
        r_mv.font.name = FONT_HEADING; r_mv.font.size = Pt(20); r_mv.font.bold = True; r_mv.font.color.rgb = m_col
        p_ml = tf_m.add_paragraph(); r_ml = p_ml.add_run(); r_ml.text = m_lbl
        r_ml.font.name = FONT_BODY; r_ml.font.size = Pt(9.5); r_ml.font.bold = True; r_ml.font.color.rgb = C_TEXT_DARK
        p_ms = tf_m.add_paragraph(); r_ms = p_ms.add_run(); r_ms.text = m_sub
        r_ms.font.name = FONT_BODY; r_ms.font.size = Pt(8.5); r_ms.font.color.rgb = C_TEXT_MUTED
        mx += m_w + m_gap

    # 3 Species Cards Below (with Photos)
    sp_w = Inches(3.7)
    sp_h = Inches(4.35)
    sp_y = Inches(2.6)

    species_showcase = [
        ("African Baobab", "Adansonia digitata", "Zone 1: Botanical Grove", "Heritage Botanical Tree", 
         "animal_photos/baobab_tree.jpg", "• Over 160+ years old heritage specimen.\n• Massive trunk circumference; provides extensive microhabitat cavities for parakeets & bats.\n• Exceptional carbon sequestration reservoir.",
         Inches(0.8), C_PRIMARY_DARK),
        
        ("Black-crowned Night Heron", "Nycticorax nycticorax", "Zone 2: Lotus Wetland", "Native Aquatic Avifauna",
         "animal_photos/painted_stork.jpg", "• Free-living nocturnal aquatic forager.\n• Roosts in submerged Barringtonia branches.\n• Highly sensitive biological indicator of wetland water quality and fish stocks.",
         Inches(4.8), C_ACCENT_BLUE),

        ("Malabar Nut / Adulsa", "Justicia adhatoda", "Zone 3: Butterfly Shrublands", "Native Medicinal Shrub",
         "animal_photos/adulsa_plant.jpg", "• Evergreen native medicinal perennial.\n• Critical larval host and nectar plant for Pieridae and Papilionidae butterflies.\n• High tolerance to urban dust & PM2.5 particulate matter.",
         Inches(8.8), C_PRIMARY_MED)
    ]

    for cname, sname, zone_info, status, img_path, bullets, sx, col in species_showcase:
        sc = add_card(s5, sx, sp_y, sp_w, sp_h, bg_color=C_CARD_BG)
        
        # Image at top of card
        if os.path.exists(img_path):
            s5.shapes.add_picture(img_path, sx + Inches(0.2), sp_y + Inches(0.2), width=Inches(3.3), height=Inches(1.8))

        # Description text below image
        desc_box = s5.shapes.add_textbox(sx + Inches(0.2), sp_y + Inches(2.1), Inches(3.3), Inches(2.15))
        tf_db = desc_box.text_frame; tf_db.word_wrap = True; tf_db.margin_left = tf_db.margin_top = 0
        
        p_sn = tf_db.paragraphs[0]; r_sn = p_sn.add_run(); r_sn.text = f"{cname} "
        r_sn.font.name = FONT_HEADING; r_sn.font.size = Pt(12); r_sn.font.bold = True; r_sn.font.color.rgb = col
        r_sci = p_sn.add_run(); r_sci.text = f"({sname})\n"
        r_sci.font.name = FONT_BODY; r_sci.font.size = Pt(9.5); r_sci.font.italic = True; r_sci.font.color.rgb = C_TEXT_MUTED

        p_loc = tf_db.add_paragraph(); r_loc = p_loc.add_run(); r_loc.text = f"📍 {zone_info}  •  {status}"
        r_loc.font.name = FONT_BODY; r_loc.font.size = Pt(8.5); r_loc.font.bold = True; r_loc.font.color.rgb = col

        p_db = tf_db.add_paragraph(); p_db.space_before = Pt(4); r_db = p_db.add_run(); r_db.text = bullets
        r_db.font.name = FONT_BODY; r_db.font.size = Pt(9.5); r_db.font.color.rgb = C_TEXT_DARK

    # ====================================================
    # SLIDE 6: Quantitative Diversity & Acoustic Analysis
    # ====================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "Quantitative Diversity & Acoustic Damping Analysis", slide_num=6)

    # Left: Ecological Diversity & Sound Comparison Table
    t_card = add_card(s6, Inches(0.8), Inches(1.3), Inches(6.8), Inches(5.6), bg_color=C_CARD_BG)
    t_title_box = s6.shapes.add_textbox(Inches(1.0), Inches(1.42), Inches(6.4), Inches(0.35))
    tf_tc = t_title_box.text_frame; tf_tc.word_wrap = True; tf_tc.margin_left = tf_tc.margin_top = 0
    p_tct = tf_tc.paragraphs[0]; p_tct.add_run().text = "📊 ECOLOGICAL DIVERSITY METRICS BY ZONE"
    p_tct.runs[0].font.name = FONT_BODY; p_tct.runs[0].font.size = Pt(11); p_tct.runs[0].font.bold = True; p_tct.runs[0].font.color.rgb = C_PRIMARY_MED

    # Add Table inside Slide 6
    tbl_shape = s6.shapes.add_table(5, 5, Inches(1.0), Inches(1.85), Inches(6.4), Inches(2.3))
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(1.4)
    tbl.columns[1].width = Inches(1.1)
    tbl.columns[2].width = Inches(1.3)
    tbl.columns[3].width = Inches(1.2)
    tbl.columns[4].width = Inches(1.4)

    headers = ["Survey Zone", "Richness (S)", "Shannon (H')", "Evenness (J')", "Noise (dBA)"]
    for c_i, h in enumerate(headers):
        cell = tbl.cell(0, c_i); cell.fill.solid(); cell.fill.fore_color.rgb = C_PRIMARY_DARK
        p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = h; r.font.name = FONT_BODY; r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = C_TEXT_LIGHT

    t_data = [
        ("Z1: Botanical Grove", "24 Taxa", "2.68", "0.84", "52.1 dBA"),
        ("Z2: Lotus Wetland", "16 Taxa", "2.34", "0.81", "54.8 dBA"),
        ("Z3: Promenade", "12 Taxa", "2.05", "0.79", "61.2 dBA"),
        ("Z4: Perimeter Buffer", "6 Taxa", "1.48", "0.65", "78.4 dBA")
    ]
    for r_i, row in enumerate(t_data, start=1):
        bg = C_CARD_MINT if r_i == 1 else (RGBColor(254, 242, 242) if r_i == 4 else C_CARD_BG)
        for c_i, val in enumerate(row):
            cell = tbl.cell(r_i, c_i); cell.fill.solid(); cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            if c_i > 0: p.alignment = PP_ALIGN.CENTER
            r = p.add_run(); r.text = val; r.font.name = FONT_BODY; r.font.size = Pt(10)
            if c_i == 0 or (r_i == 1 and c_i == 2): r.font.bold = True
            r.font.color.rgb = C_TEXT_DARK

    # Mathematical Formula Callout below table
    f_box = s6.shapes.add_textbox(Inches(1.0), Inches(4.35), Inches(6.4), Inches(2.3))
    tf_f = f_box.text_frame; tf_f.word_wrap = True; tf_f.margin_left = tf_f.margin_top = 0
    p_f1 = tf_f.paragraphs[0]; p_f1.add_run().text = "📐 Quantitative Indices Applied:"
    p_f1.runs[0].font.name = FONT_BODY; p_f1.runs[0].font.size = Pt(11); p_f1.runs[0].font.bold = True; p_f1.runs[0].font.color.rgb = C_PRIMARY_DARK
    
    p_f2 = tf_f.add_paragraph(); p_f2.space_before = Pt(4)
    r_f2 = p_f2.add_run()
    r_f2.text = "• Shannon-Wiener Diversity: H' = - Σ (pi × ln pi)\n" \
                "  Reflects both species abundance and community equitability.\n" \
                "• Pielou's Evenness Index: J' = H' / ln(S)\n" \
                "  Measures distribution uniformity across observed microhabitats.\n" \
                "• Acoustic Attenuation Differential: Δ = L(Z4) - L(Z1) = 26.3 dBA.\n" \
                "  The arboreal multi-tier canopy delivers a massive sound absorption shield."
    r_f2.font.name = FONT_BODY; r_f2.font.size = Pt(10); r_f2.font.color.rgb = C_TEXT_DARK

    # Right: Key Findings Cards (Acoustic Shield & Ecological Gradients)
    r_x = Inches(7.9); r_w = Inches(4.6)
    
    rf1 = add_card(s6, r_x, Inches(1.3), r_w, Inches(2.65), bg_color=C_CARD_MINT, border_color=C_BORDER_GREEN)
    tf_rf1 = rf1.text_frame; tf_rf1.word_wrap = True; tf_rf1.margin_left = tf_rf1.margin_right = tf_rf1.margin_top = Inches(0.2)
    p_rf1_t = tf_rf1.paragraphs[0]; p_rf1_t.add_run().text = "🔊 26.3 dBA Acoustic Attenuation"
    p_rf1_t.runs[0].font.name = FONT_HEADING; p_rf1_t.runs[0].font.size = Pt(15); p_rf1_t.runs[0].font.bold = True; p_rf1_t.runs[0].font.color.rgb = C_PRIMARY_DARK
    p_rf1_b = tf_rf1.add_paragraph(); p_rf1_b.space_before = Pt(6)
    r_rf1 = p_rf1_b.add_run()
    r_rf1.text = "• Perimeter ambient traffic along Dr. Ambedkar Road regularly exceeds 78.4 dBA.\n" \
                 "• The multi-layered canopy of Zone 1 dampens sound levels down to 52.1 dBA (a 26.3 dBA drop).\n" \
                 "• This acoustic shielding is vital for vocalizing avifauna (Barbet, Koel, Parakeet) whose mating and territorial calls are masked by urban noise."
    r_rf1.font.name = FONT_BODY; r_rf1.font.size = Pt(10.5); r_rf1.font.color.rgb = C_TEXT_DARK

    rf2 = add_card(s6, r_x, Inches(4.15), r_w, Inches(2.75), bg_color=C_CARD_BG)
    tf_rf2 = rf2.text_frame; tf_rf2.word_wrap = True; tf_rf2.margin_left = tf_rf2.margin_right = tf_rf2.margin_top = Inches(0.2)
    p_rf2_t = tf_rf2.paragraphs[0]; p_rf2_t.add_run().text = "🌿 81% Greater Biodiversity Index"
    p_rf2_t.runs[0].font.name = FONT_HEADING; p_rf2_t.runs[0].font.size = Pt(15); p_rf2_t.runs[0].font.bold = True; p_rf2_t.runs[0].font.color.rgb = C_PRIMARY_MED
    p_rf2_b = tf_rf2.add_paragraph(); p_rf2_b.space_before = Pt(6)
    r_rf2 = p_rf2_b.add_run()
    r_rf2.text = "• Species richness drops sharply from 24 taxa in the core grove to only 6 taxa at the perimeter boundary.\n" \
                 "• Shannon Index H' is 81% higher in Zone 1 (2.68 vs. 1.48 in Zone 4).\n" \
                 "• Validates the ecological core hypothesis: Interior canopy continuity and water presence are non-negotiable for urban wildlife survival."
    r_rf2.font.name = FONT_BODY; r_rf2.font.size = Pt(10.5); r_rf2.font.color.rgb = C_TEXT_DARK

    # ====================================================
    # SLIDE 7: Community Survey Findings (N = 84)
    # ====================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "Community Survey Findings & Perception Analysis (N = 84)", slide_num=7)

    # Left: High-Res Matplotlib Survey Chart 4.5
    c_w = Inches(5.6); c_h = Inches(3.6)
    if os.path.exists("fig_4_5.png"):
        s7.shapes.add_picture("fig_4_5.png", Inches(0.8), Inches(1.3), width=c_w)
    
    # Right: Matplotlib Pie Chart 4.8 & 4.9
    if os.path.exists("fig_4_8_4_9.png"):
        s7.shapes.add_picture("fig_4_8_4_9.png", Inches(6.8), Inches(1.3), width=c_w)

    # Bottom 4 Key Insights Cards
    ins_w = Inches(2.7)
    ins_h = Inches(1.7)
    ins_gap = Inches(0.3)
    
    insights = [
        ("HABITAT CLUSTERING", "56.5% of Wildlife Sightings", 
         "A decisive majority of sightings occur strictly in Zone 1 Grove and Zone 2 Lake, confirming dependency on dense cover.", C_PRIMARY_DARK),
        
        ("PERCEIVED THREATS", "36.5% Branch Pruning", 
         "Citizens identify arbitrary pre-monsoon tree trimming and perimeter traffic noise as the #1 threat to native avifauna.", C_ACCENT_RED),

        ("LITERACY DEFICIT", "72% Lack Digital Maps", 
         "72% cite physical signs focusing exclusively on caged animals as the main barrier preventing botanical exploration.", C_ACCENT_BLUE),

        ("DIGITAL DEMAND", "86.7% Endorse Web App", 
         "86.7% affirm an interactive Web-GIS would enhance park visits, and 83.3% confirm they would actively use it.", C_PRIMARY_MED)
    ]
    
    ix = Inches(0.8)
    for i_top, i_main, i_desc, i_col in insights:
        ic = add_card(s7, ix, Inches(5.15), ins_w, ins_h, bg_color=C_CARD_BG)
        tf_ic = ic.text_frame; tf_ic.word_wrap = True; tf_ic.margin_left = tf_ic.margin_right = tf_ic.margin_top = Inches(0.14)
        
        p_it = tf_ic.paragraphs[0]; r_it = p_it.add_run(); r_it.text = i_top
        r_it.font.name = FONT_BODY; r_it.font.size = Pt(8.5); r_it.font.bold = True; r_it.font.color.rgb = i_col
        
        p_im = tf_ic.add_paragraph(); r_im = p_im.add_run(); r_im.text = i_main
        r_im.font.name = FONT_HEADING; r_im.font.size = Pt(11); r_im.font.bold = True; r_im.font.color.rgb = C_TEXT_DARK
        
        p_id = tf_ic.add_paragraph(); p_id.space_before = Pt(2); r_id = p_id.add_run(); r_id.text = i_desc
        r_id.font.name = FONT_BODY; r_id.font.size = Pt(8.5); r_id.font.color.rgb = C_TEXT_MUTED
        
        ix += ins_w + ins_gap

    # ====================================================
    # SLIDE 8: The BioMap Web-GIS Platform (Solution & Features)
    # ====================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "Technological Solution: BioMap Web-GIS Platform", slide_num=8)

    # Left: Platform Screenshot
    if os.path.exists("biomap_screenshot.png"):
        s8.shapes.add_picture("biomap_screenshot.png", Inches(0.8), Inches(1.3), width=Inches(7.2))

    # Right: 4 Architecture Feature Cards
    rf_w = Inches(4.3)
    rf_h = Inches(1.08)
    rf_gap = Inches(0.10)
    rf_x = Inches(8.3)
    rf_y = Inches(1.3)

    platform_features = [
        ("📍 Real-Time HTML5 Geolocation", 
         "Utilizes mobile browser GPS positioning to track visitors in real-time along 53 acres of garden trails without app store installation.", C_PRIMARY_DARK),
        
        ("🗺️ Interactive 4-Zone GeoJSON Polygons", 
         "Displays semi-transparent micro-zone overlays communicating microclimatic cooling, canopy density, and species richness metrics.", C_ACCENT_BLUE),

        ("🌐 Bilingual Vernacular Taxonomy", 
         "Each species features Linnaean classification, flowering/nesting phenology, and Marathi vernacular descriptions (जैविविधता माहिती).", C_ACCENT_AMBER),

        ("⚡ Zero-Cost Open-Source Stack", 
         "Engineered with Leaflet.js, OpenStreetMap, and Vanilla CSS3; hosted on Netlify with zero municipal server licensing costs.", C_PRIMARY_MED)
    ]

    for p_title, p_desc, p_col in platform_features:
        pfc = add_card(s8, rf_x, rf_y, rf_w, rf_h, bg_color=C_CARD_BG)
        tf_pfc = pfc.text_frame; tf_pfc.word_wrap = True; tf_pfc.margin_left = tf_pfc.margin_right = tf_pfc.margin_top = Inches(0.14)
        
        p_pt = tf_pfc.paragraphs[0]; r_pt = p_pt.add_run(); r_pt.text = p_title
        r_pt.font.name = FONT_HEADING; r_pt.font.size = Pt(11.5); r_pt.font.bold = True; r_pt.font.color.rgb = p_col
        
        p_pd = tf_pfc.add_paragraph(); p_pd.space_before = Pt(3); r_pd = p_pd.add_run(); r_pd.text = p_desc
        r_pd.font.name = FONT_BODY; r_pd.font.size = Pt(9.5); r_pd.font.color.rgb = C_TEXT_DARK
        
        rf_y += rf_h + rf_gap

    # Platform URL Banner at Bottom
    url_card = add_card(s8, Inches(0.8), Inches(6.15), Inches(11.8), Inches(0.8), bg_color=C_CARD_MINT, border_color=C_BORDER_GREEN)
    tf_uc = url_card.text_frame; tf_uc.word_wrap = True; tf_uc.margin_left = Inches(0.2); tf_uc.margin_top = Inches(0.12)
    p_u = tf_uc.paragraphs[0]
    r_u1 = p_u.add_run(); r_u1.text = "🔗 Live Public Web Application: "
    r_u1.font.name = FONT_BODY; r_u1.font.size = Pt(11); r_u1.font.bold = True; r_u1.font.color.rgb = C_PRIMARY_DARK
    r_u2 = p_u.add_run(); r_u2.text = "https://biomap-byculla.netlify.app   "
    r_u2.font.name = FONT_BODY; r_u2.font.size = Pt(11); r_u2.font.bold = True; r_u2.font.color.rgb = C_ACCENT_BLUE
    r_u3 = p_u.add_run(); r_u3.text = "• Tested with 84 field participants with 92% UX approval."
    r_u3.font.name = FONT_BODY; r_u3.font.size = Pt(10.5); r_u3.font.color.rgb = C_TEXT_MUTED

    # ====================================================
    # SLIDE 9: Recommendations & Three-Tier Action Plan
    # ====================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "Recommendations: Three-Tier Municipal Blueprint", slide_num=9)

    tier_w = Inches(3.68)
    tier_gap = Inches(0.33)
    tier_y = Inches(1.3)
    tier_h = Inches(5.6)

    # Tier 1
    t1 = add_card(s9, Inches(0.8), tier_y, tier_w, tier_h, bg_color=C_CARD_MINT, border_color=C_BORDER_GREEN)
    tf_t1 = t1.text_frame; tf_t1.word_wrap = True; tf_t1.margin_left = tf_t1.margin_right = tf_t1.margin_top = Inches(0.2)
    p_t1_h = tf_t1.paragraphs[0]; p_t1_h.add_run().text = "TIER 1: PARK LEVEL\nImmediate Operations (1–3 Mos)"
    p_t1_h.runs[0].font.name = FONT_HEADING; p_t1_h.runs[0].font.size = Pt(13); p_t1_h.runs[0].font.bold = True; p_t1_h.runs[0].font.color.rgb = C_PRIMARY_DARK
    
    p_t1_b = tf_t1.add_paragraph(); p_t1_b.space_before = Pt(10)
    r_t1_b = p_t1_b.add_run()
    r_t1_b.text = "• Weatherproof QR-Code Placards:\n  Affix laser-etched QR plaques to 42 audited heritage trees and observation decks linking directly to BioMap.\n\n" \
                  "• Nesting-Season Pruning Protocol:\n  Institute a mandatory pre-pruning survey between February and June to stop lopping branches with active bird nests.\n\n" \
                  "• Wetland Edge Protection:\n  Install floating reed islands in Lotus Lake to provide undisturbed roosting for Night Herons."
    r_t1_b.font.name = FONT_BODY; r_t1_b.font.size = Pt(10.5); r_t1_b.font.color.rgb = C_TEXT_DARK

    # Tier 2
    t2 = add_card(s9, Inches(0.8) + tier_w + tier_gap, tier_y, tier_w, tier_h, bg_color=RGBColor(239, 246, 255), border_color=RGBColor(191, 219, 254))
    tf_t2 = t2.text_frame; tf_t2.word_wrap = True; tf_t2.margin_left = tf_t2.margin_right = tf_t2.margin_top = Inches(0.2)
    p_t2_h = tf_t2.paragraphs[0]; p_t2_h.add_run().text = "TIER 2: BMC MUNICIPAL\nInstitutional Policy (6–12 Mos)"
    p_t2_h.runs[0].font.name = FONT_HEADING; p_t2_h.runs[0].font.size = Pt(13); p_t2_h.runs[0].font.bold = True; p_t2_h.runs[0].font.color.rgb = C_ACCENT_BLUE
    
    p_t2_b = tf_t2.add_paragraph(); p_t2_b.space_before = Pt(10)
    r_t2_b = p_t2_b.add_run()
    r_t2_b.text = "• PBR Integration:\n  Integrate BioMap's GIS dataset into Brihanmumbai Municipal Corporation's statutory People's Biodiversity Register.\n\n" \
                  "• Multi-Tier Acoustic Buffer:\n  Plant dense ficus and bamboo hedgerows along the western perimeter to dampen traffic noise from Dr. Ambedkar Road.\n\n" \
                  "• City-Wide Replication:\n  Adopt BioMap's zero-cost open-source template for other major Mumbai parks (Five Gardens, Horniman Circle, Oval Maidan)."
    r_t2_b.font.name = FONT_BODY; r_t2_b.font.size = Pt(10.5); r_t2_b.font.color.rgb = C_TEXT_DARK

    # Tier 3
    t3 = add_card(s9, Inches(0.8) + (tier_w + tier_gap)*2, tier_y, tier_w, tier_h, bg_color=RGBColor(254, 243, 199), border_color=RGBColor(253, 230, 138))
    tf_t3 = t3.text_frame; tf_t3.word_wrap = True; tf_t3.margin_left = tf_t3.margin_right = tf_t3.margin_top = Inches(0.2)
    p_t3_h = tf_t3.paragraphs[0]; p_t3_h.add_run().text = "TIER 3: CITIZEN SCIENCE\nCivic Engagement (Ongoing)"
    p_t3_h.runs[0].font.name = FONT_HEADING; p_t3_h.runs[0].font.size = Pt(13); p_t3_h.runs[0].font.bold = True; p_t3_h.runs[0].font.color.rgb = C_ACCENT_AMBER
    
    p_t3_b = tf_t3.add_paragraph(); p_t3_b.space_before = Pt(10)
    r_t3_b = p_t3_b.add_run()
    r_t3_b.text = "• University Bio-Blitz Workshops:\n  Conduct bi-monthly weekend guided biodiversity walks for undergraduate CS and Life Sciences students.\n\n" \
                  "• Citizen Sighting crowdsourcing:\n  Enable visitors to submit geotagged photos of non-captive fauna with peer review by botanists.\n\n" \
                  "• Seasonal Migration Tracking:\n  Establish seasonal census drives for winter avifaunal migrants to track long-term climate resilience in Mumbai."
    r_t3_b.font.name = FONT_BODY; r_t3_b.font.size = Pt(10.5); r_t3_b.font.color.rgb = C_TEXT_DARK

    # ====================================================
    # SLIDE 10: Conclusion, NEP 2020 Outcomes & Thank You
    # ====================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "Conclusion & NEP 2020 Learning Outcomes", slide_num=10)

    # Left: Academic & Practical Outcomes
    out_card = add_card(s10, Inches(0.8), Inches(1.3), Inches(6.5), Inches(5.6), bg_color=C_CARD_BG)
    tf_out = out_card.text_frame; tf_out.word_wrap = True; tf_out.margin_left = tf_out.margin_right = tf_out.margin_top = Inches(0.2)
    p_ot = tf_out.paragraphs[0]; p_ot.add_run().text = "🎓 NEP 2020 PEDAGOGICAL & COMMUNITY IMPACT"
    p_ot.runs[0].font.name = FONT_BODY; p_ot.runs[0].font.size = Pt(11); p_ot.runs[0].font.bold = True; p_ot.runs[0].font.color.rgb = C_PRIMARY_MED

    p_ob = tf_out.add_paragraph(); p_ob.space_before = Pt(8)
    r_ob = p_ob.add_run()
    r_ob.text = "1. Applied Computer Science & Geospatial Engineering:\n" \
                "   Bridged classroom GIS theory, full-stack JavaScript, and real-world urban ecological datasets.\n\n" \
                "2. Multi-Stakeholder Civic Engagement:\n" \
                "   Conducted empirical field surveys across 84 diverse citizens, synthesizing quantitative analytics with qualitative community needs.\n\n" \
                "3. Tangible Digital Public Good:\n" \
                "   Rather than a theoretical document, the project delivered a deployed, functioning Web-GIS application ready for public use.\n\n" \
                "4. Environmental Responsibility & Stewardship:\n" \
                "   Uncovered critical evidence on arboreal acoustic buffering (26.3 dBA) and pruning pressure, informing municipal conservation policy."
    r_ob.font.name = FONT_BODY; r_ob.font.size = Pt(10.5); r_ob.font.color.rgb = C_TEXT_DARK

    # Right Top: Mentorship Acknowledgement Card
    ment_card = add_card(s10, Inches(7.6), Inches(1.3), Inches(4.9), Inches(2.6), bg_color=C_CARD_MINT, border_color=C_BORDER_GREEN)
    tf_mc = ment_card.text_frame; tf_mc.word_wrap = True; tf_mc.margin_left = tf_mc.margin_right = tf_mc.margin_top = Inches(0.2)
    p_mt = tf_mc.paragraphs[0]; p_mt.add_run().text = "FACULTY MENTORSHIP ACKNOWLEDGEMENT"
    p_mt.runs[0].font.name = FONT_BODY; p_mt.runs[0].font.size = Pt(10); p_mt.runs[0].font.bold = True; p_mt.runs[0].font.color.rgb = C_PRIMARY_DARK

    p_mn = tf_mc.add_paragraph(); p_mn.space_before = Pt(4)
    r_mn = p_mn.add_run(); r_mn.text = "Prof. Shweta Iyer\n"
    r_mn.font.name = FONT_HEADING; r_mn.font.size = Pt(15); r_mn.font.bold = True; r_mn.font.color.rgb = C_PRIMARY_DARK
    
    r_md = p_mn.add_run()
    r_md.text = "Internal Project Guide • Department of Computer Science\nVidyavardhini's College of Science (Vartak College)\n\n" \
                "Heartfelt gratitude for continuous guidance, methodology reviews during Guide Interaction Diary sessions, and encouraging digital innovation for community impact."
    r_md.font.name = FONT_BODY; r_md.font.size = Pt(10); r_md.font.color.rgb = C_TEXT_DARK

    # Right Bottom: Grand Thank You & Q&A Box
    ty_card = add_card(s10, Inches(7.6), Inches(4.1), Inches(4.9), Inches(2.8), bg_color=C_PRIMARY_DARK)
    tf_ty = ty_card.text_frame; tf_ty.word_wrap = True; tf_ty.margin_left = tf_ty.margin_right = tf_ty.margin_top = Inches(0.2)
    p_ty = tf_ty.paragraphs[0]; p_ty.alignment = PP_ALIGN.CENTER
    r_ty = p_ty.add_run(); r_ty.text = "THANK YOU!"
    r_ty.font.name = FONT_HEADING; r_ty.font.size = Pt(24); r_ty.font.bold = True; r_ty.font.color.rgb = RGBColor(253, 224, 71)

    p_qa = tf_ty.add_paragraph(); p_qa.alignment = PP_ALIGN.CENTER; p_qa.space_before = Pt(4)
    r_qa = p_qa.add_run(); r_qa.text = "Questions & Discussion Welcome\n"
    r_qa.font.name = FONT_BODY; r_qa.font.size = Pt(12); r_qa.font.bold = True; r_qa.font.color.rgb = C_TEXT_LIGHT

    p_links = tf_ty.add_paragraph(); p_links.alignment = PP_ALIGN.CENTER; p_links.space_before = Pt(6)
    r_links = p_links.add_run()
    r_links.text = "Explore the Live Web-GIS Platform:\nhttps://biomap-byculla.netlify.app\n\nVidyavardhini's College • T.Y. B.Sc. CS 2026-2027"
    r_links.font.name = FONT_BODY; r_links.font.size = Pt(10); r_links.font.color.rgb = RGBColor(209, 250, 229)

    # Save presentation
    output_filename = "Byculla Zoo Biodiversity Mapping - CEP Presentation.pptx"
    prs.save(output_filename)
    print(f"Presentation saved successfully as '{output_filename}' ({len(prs.slides)} slides)!")
    return output_filename

def export_ppt_to_pdf(ppt_path, pdf_path):
    import win32com.client
    abs_ppt = os.path.abspath(ppt_path)
    abs_pdf = os.path.abspath(pdf_path)
    
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        deck = ppt_app.Presentations.Open(abs_ppt, WithWindow=False)
        deck.SaveAs(abs_pdf, 32) # 32 = ppSaveAsPDF
        deck.Close()
        print(f"PDF exported successfully to '{pdf_path}'!")
    finally:
        ppt_app.Quit()

if __name__ == "__main__":
    ppt_file = create_presentation()
    pdf_file = "Byculla Zoo Biodiversity Mapping - CEP Presentation.pdf"
    try:
        export_ppt_to_pdf(ppt_file, pdf_file)
    except Exception as e:
        print("Note: Could not export PPT to PDF directly:", e)
