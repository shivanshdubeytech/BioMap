import os
import win32com.client
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def build_presentation():
    template_path = os.path.abspath('The-hidden-jungle-in-Mumbais-concrete-heart.pptx')
    output_pptx = os.path.abspath('Byculla Zoo Biodiversity Mapping - The Hidden Jungle.pptx')
    output_pdf = os.path.abspath('Byculla Zoo Biodiversity Mapping - The Hidden Jungle.pdf')
    
    prs = Presentation(template_path)
    
    # ----------------------------------------------------
    # SLIDE 1: Title Slide Customization
    # ----------------------------------------------------
    s1 = prs.slides[0]
    
    # Kicker / Institutional Header
    s1_shape0 = s1.shapes[0]
    s1_shape0.text_frame.word_wrap = True
    s1_shape0.left = Inches(0.81)
    s1_shape0.top = Inches(0.65)
    s1_shape0.width = Inches(10.2)
    s1_shape0.height = Inches(0.5)
    
    tf0 = s1_shape0.text_frame
    tf0.clear()
    p0a = tf0.paragraphs[0]
    p0a.text = "VIDYAVARDHINI’S A. V. COLLEGE OF ARTS, K. M. COLLEGE OF COMMERCE & E. S. A. COLLEGE OF SCIENCE"
    p0a.font.name = "Red Hat Mono Medium"
    p0a.font.size = Pt(10.5)
    p0a.font.bold = True
    p0a.font.color.rgb = RGBColor(244, 226, 60) # Gold accent
    
    p0b = tf0.add_paragraph()
    p0b.text = "UNIVERSITY OF MUMBAI  •  COMMUNITY ENGAGEMENT PROJECT (NEP 2020 GUIDELINES)"
    p0b.font.name = "Red Hat Mono Medium"
    p0b.font.size = Pt(9.5)
    p0b.font.color.rgb = RGBColor(200, 225, 210)
    
    # Title: Keep iconic Hedvig Letters Serif, position nicely
    s1_shape1 = s1.shapes[1]
    s1_shape1.top = Inches(1.95)
    s1_shape1.height = Inches(2.2)
    
    # Accent line Shape 2:
    s1_shape2 = s1.shapes[2]
    s1_shape2.top = Inches(4.25)
    s1_shape2.width = Inches(0.9)
    s1_shape2.height = Inches(0.04)
    
    # Subtitle Shape 3:
    s1_shape3 = s1.shapes[3]
    s1_shape3.top = Inches(4.45)
    s1_shape3.height = Inches(0.85)
    s1_shape3.text_frame.word_wrap = True
    p3 = s1_shape3.text_frame.paragraphs[0]
    p3.text = "Local Biodiversity Mapping of Byculla Zoo & Surrounding Green Spaces\nUrban Ecology, 42 Documented Wild Taxa, and the BioMap Web-GIS Platform"
    p3.font.name = "Albert Sans"
    p3.font.size = Pt(17)
    p3.font.color.rgb = RGBColor(235, 245, 235)
    
    # Bottom Academic Metadata Card
    card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.81), Inches(5.5), Inches(11.72), Inches(1.4))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(12, 26, 18) # Deep jungle tone
    card.line.color.rgb = RGBColor(50, 95, 65)
    card.line.width = Pt(1.2)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.35)
    tf.margin_top = Inches(0.18)
    tf.margin_right = Inches(0.35)
    tf.margin_bottom = Inches(0.18)
    
    cp0 = tf.paragraphs[0]
    cp0.text = "ACADEMIC PROJECT CREDENTIALS & VIVA PARTICULARS"
    cp0.font.name = "Red Hat Mono Medium"
    cp0.font.size = Pt(9.5)
    cp0.font.bold = True
    cp0.font.color.rgb = RGBColor(244, 226, 60)
    
    cp1 = tf.add_paragraph()
    cp1.text = "Student Investigator: Riddhesh & Team  |  Course: T.Y. B.Sc. Computer Science (Semester V, NEP 2020)\n" \
               "Project Guide / Mentor: Prof. Shweta Iyer, Department of Computer Science  |  Vartak College, Vasai (W)\n" \
               "Live Web-GIS Platform: https://biomap-byculla.netlify.app  |  Academic Year: 2026–2027"
    cp1.font.name = "Albert Sans"
    cp1.font.size = Pt(11)
    cp1.font.color.rgb = RGBColor(240, 248, 240)
    
    # College Logo in Top Right
    if os.path.exists('college_logo.jpg'):
        logo_box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.35), Inches(0.55), Inches(1.18), Inches(1.0))
        logo_box.fill.solid()
        logo_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        logo_box.line.color.rgb = RGBColor(200, 220, 205)
        s1.shapes.add_picture('college_logo.jpg', Inches(11.42), Inches(0.61), Inches(1.04), Inches(0.88))

    # ----------------------------------------------------
    # SLIDE 8: BioMap Civic Tool & Live Web-GIS Interface
    # ----------------------------------------------------
    s8 = prs.slides[7]
    # Remove wireframe placeholder chips (shapes 13 to 20)
    shapes_to_remove = [s8.shapes[i] for i in range(13, 21)]
    for shp in shapes_to_remove:
        sp = shp._element
        sp.getparent().remove(sp)
        
    # Shape 12 (Shape 7) is our right-hand container card
    s8_card = s8.shapes[12]
    s8_card.left = Inches(8.25)
    s8_card.top = Inches(1.65)
    s8_card.width = Inches(4.45)
    s8_card.height = Inches(5.25)
    s8_card.fill.solid()
    s8_card.fill.fore_color.rgb = RGBColor(242, 248, 243)
    s8_card.line.color.rgb = RGBColor(180, 215, 190)
    
    # Add real BioMap Web-GIS screenshot
    if os.path.exists('biomap_screenshot.png'):
        s8.shapes.add_picture('biomap_screenshot.png', Inches(8.45), Inches(1.85), Inches(4.05), Inches(2.2))
        
    # Add Title Box below screenshot
    t_box = s8.shapes.add_textbox(Inches(8.45), Inches(4.15), Inches(4.05), Inches(0.35))
    tb_tf = t_box.text_frame
    tb_tf.margin_left = tb_tf.margin_top = tb_tf.margin_right = tb_tf.margin_bottom = 0
    tb_p = tb_tf.paragraphs[0]
    tb_p.text = "LIVE WEB-GIS APPLICATION INTERFACE"
    tb_p.font.name = "Red Hat Mono Medium"
    tb_p.font.size = Pt(10.5)
    tb_p.font.bold = True
    tb_p.font.color.rgb = RGBColor(23, 129, 58)
    
    # Add Descriptive Bullets
    b_box = s8.shapes.add_textbox(Inches(8.45), Inches(4.5), Inches(4.05), Inches(1.7))
    bb_tf = b_box.text_frame
    bb_tf.word_wrap = True
    bb_tf.margin_left = bb_tf.margin_top = bb_tf.margin_right = bb_tf.margin_bottom = 0
    
    bullets = [
        "Interactive Leaflet.js engine with custom GPS pins",
        "Taxon filter: Flora, Aves, Lepidoptera, Reptilia",
        "Direct GeoJSON & CSV dataset export (zero paywall)",
        "Fully responsive & WCAG 2.1 AA accessible"
    ]
    for idx, b_text in enumerate(bullets):
        bp = bb_tf.paragraphs[0] if idx == 0 else bb_tf.add_paragraph()
        bp.text = f"•  {b_text}"
        bp.font.name = "Albert Sans"
        bp.font.size = Pt(10.5)
        bp.font.color.rgb = RGBColor(30, 45, 35)
        bp.space_after = Pt(3)
        
    # Direct access URL pill
    url_pill = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.45), Inches(6.25), Inches(4.05), Inches(0.48))
    url_pill.fill.solid()
    url_pill.fill.fore_color.rgb = RGBColor(23, 129, 58)
    url_pill.line.color.rgb = RGBColor(23, 129, 58)
    u_tf = url_pill.text_frame
    u_p = u_tf.paragraphs[0]
    u_p.text = "🌐  biomap-byculla.netlify.app"
    u_p.alignment = PP_ALIGN.CENTER
    u_p.font.name = "Albert Sans Bold"
    u_p.font.size = Pt(11)
    u_p.font.bold = True
    u_p.font.color.rgb = RGBColor(255, 255, 255)

    # ----------------------------------------------------
    # SLIDE 10: Three Action Tiers & Academic Guidance Card
    # ----------------------------------------------------
    s10 = prs.slides[9]
    # Move handoff headers slightly up
    s10.shapes[5].top = Inches(1.8) # THE HANDOFF
    s10.shapes[6].top = Inches(2.15) # Tier 1
    s10.shapes[7].top = Inches(2.7) # Tier 2
    s10.shapes[8].top = Inches(3.25) # Tier 3
    
    # Academic Mentorship Box in Right Column
    conc_box = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.35), Inches(4.05), Inches(3.7), Inches(2.95))
    conc_box.fill.solid()
    conc_box.fill.fore_color.rgb = RGBColor(14, 38, 24) # Dark jungle forest
    conc_box.line.color.rgb = RGBColor(244, 226, 60) # Gold border
    conc_box.line.width = Pt(1.2)
    
    c_tf = conc_box.text_frame
    c_tf.word_wrap = True
    c_tf.margin_left = Inches(0.25)
    c_tf.margin_right = Inches(0.25)
    c_tf.margin_top = Inches(0.2)
    c_tf.margin_bottom = Inches(0.2)
    
    p_c0 = c_tf.paragraphs[0]
    p_c0.text = "PROJECT GUIDANCE & SUBMISSION"
    p_c0.font.name = "Red Hat Mono Medium"
    p_c0.font.size = Pt(9.5)
    p_c0.font.bold = True
    p_c0.font.color.rgb = RGBColor(244, 226, 60)
    p_c0.space_after = Pt(4)
    
    p_c1 = c_tf.add_paragraph()
    p_c1.text = "Project Guide / Mentor:\nProf. Shweta Iyer\nDepartment of Computer Science"
    p_c1.font.name = "Albert Sans Bold"
    p_c1.font.size = Pt(11)
    p_c1.font.bold = True
    p_c1.font.color.rgb = RGBColor(245, 255, 245)
    p_c1.space_after = Pt(4)
    
    p_c2 = c_tf.add_paragraph()
    p_c2.text = "Vidyavardhini’s Vartak College, Vasai (W)\nUniversity of Mumbai  •  NEP 2020 (Sem V)\nLive Web-GIS: biomap-byculla.netlify.app"
    p_c2.font.name = "Albert Sans"
    p_c2.font.size = Pt(9.5)
    p_c2.font.color.rgb = RGBColor(205, 230, 215)
    p_c2.space_after = Pt(8)
    
    p_c3 = c_tf.add_paragraph()
    p_c3.text = "Thank You!  •  Questions & Viva Discussion"
    p_c3.alignment = PP_ALIGN.CENTER
    p_c3.font.name = "Albert Sans Bold"
    p_c3.font.size = Pt(10.5)
    p_c3.font.bold = True
    p_c3.font.color.rgb = RGBColor(244, 226, 60)

    # ----------------------------------------------------
    # WATERMARK COVER / INSTITUTIONAL BRANDING PILL
    # Place a clean footer pill over the bottom-right corner on every slide
    # ----------------------------------------------------
    for s_idx, slide in enumerate(prs.slides):
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.45), Inches(6.95), Inches(1.8), Inches(0.5))
        pill.fill.solid()
        if s_idx == 0:
            pill.fill.fore_color.rgb = RGBColor(12, 26, 18)
            pill.line.color.rgb = RGBColor(50, 95, 65)
            text_color = RGBColor(244, 226, 60)
        elif s_idx == 9: # Slide 10 has pale green column on right
            pill.fill.fore_color.rgb = RGBColor(14, 38, 24)
            pill.line.color.rgb = RGBColor(244, 226, 60)
            text_color = RGBColor(244, 226, 60)
        else:
            pill.fill.fore_color.rgb = RGBColor(240, 245, 242)
            pill.line.color.rgb = RGBColor(200, 215, 205)
            text_color = RGBColor(23, 129, 58)
            
        p_pill = pill.text_frame.paragraphs[0]
        p_pill.text = "VARTAK CS • CEP 2026"
        p_pill.alignment = PP_ALIGN.CENTER
        p_pill.font.name = "Red Hat Mono Medium"
        p_pill.font.size = Pt(8.5)
        p_pill.font.bold = True
        p_pill.font.color.rgb = text_color
        
    # Save Presentation
    prs.save(output_pptx)
    print(f"[OK] Presentation successfully saved: {output_pptx}")
    
    # Export to PDF via PowerPoint COM
    ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        deck = ppt_app.Presentations.Open(output_pptx, WithWindow=False)
        deck.SaveAs(output_pdf, 32)
        deck.Close()
        print(f"[OK] PDF successfully exported: {output_pdf}")
    finally:
        ppt_app.Quit()

if __name__ == '__main__':
    build_presentation()
