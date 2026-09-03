import os
import math
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml import parse_xml
from PIL import Image, ImageDraw

# -----------------------------------------------------------------------------
# DIRECTORY & ASSET SETUP
# -----------------------------------------------------------------------------
OUTPUT_DIR = "output"
ASSETS_DIR = os.path.join(OUTPUT_DIR, "assets")
os.makedirs(ASSETS_DIR, exist_ok=True)

# -----------------------------------------------------------------------------
# PROCEDURAL FLORAL ASSET GENERATOR (Transparent PNGs)
# -----------------------------------------------------------------------------
def create_chrysanthemum_image(path):
    size = (800, 800)
    img = Image.new("RGBA", size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    cx, cy = 400, 400
    
    # Outer petals
    num_petals = 48
    for i in range(num_petals):
        angle = i * (2 * math.pi / num_petals)
        r_out = 330 + 15 * math.sin(i * 3)
        px = cx + r_out * math.cos(angle)
        py = cy + r_out * math.sin(angle)
        draw.line([(cx, cy), (px, py)], fill=(245, 178, 35, 220), width=28)
        
    for i in range(num_petals):
        angle = (i + 0.5) * (2 * math.pi / num_petals)
        r_out = 270 + 10 * math.cos(i * 2)
        px = cx + r_out * math.cos(angle)
        py = cy + r_out * math.sin(angle)
        draw.line([(cx, cy), (px, py)], fill=(255, 206, 68, 240), width=22)

    # Core
    draw.ellipse([cx - 95, cy - 95, cx + 95, cy + 95], fill=(210, 105, 30, 255))
    draw.ellipse([cx - 70, cy - 70, cx + 70, cy + 70], fill=(139, 69, 19, 255))
    draw.ellipse([cx - 40, cy - 40, cx + 40, cy + 40], fill=(85, 40, 10, 255))
    img.save(path, "PNG")


def create_tulip_image(path):
    size = (800, 800)
    img = Image.new("RGBA", size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    cx, cy = 400, 420
    
    # Stem / base leaves
    draw.polygon([(cx - 20, cy + 180), (cx + 20, cy + 180), (cx + 10, cy + 300), (cx - 10, cy + 300)], fill=(96, 138, 70, 240))
    # Left outer petal
    draw.chord([cx - 220, cy - 200, cx + 50, cy + 180], 90, 270, fill=(212, 63, 93, 235))
    # Right outer petal
    draw.chord([cx - 50, cy - 200, cx + 220, cy + 180], 270, 90, fill=(195, 45, 78, 235))
    # Center petal
    draw.ellipse([cx - 95, cy - 220, cx + 95, cy + 150], fill=(235, 90, 120, 250))
    draw.ellipse([cx - 50, cy - 190, cx + 50, cy + 100], fill=(250, 150, 170, 230))
    img.save(path, "PNG")


def create_orchid_image(path):
    size = (800, 800)
    img = Image.new("RGBA", size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    cx, cy = 400, 390
    
    # Lateral petals
    draw.ellipse([cx - 320, cy - 120, cx - 20, cy + 100], fill=(196, 160, 215, 220))
    draw.ellipse([cx + 20, cy - 120, cx + 320, cy + 100], fill=(196, 160, 215, 220))
    # Dorsal sepal
    draw.ellipse([cx - 90, cy - 240, cx + 90, cy - 10], fill=(178, 135, 202, 230))
    # Lower sepals
    draw.ellipse([cx - 140, cy + 10, cx - 10, cy + 210], fill=(170, 125, 195, 220))
    draw.ellipse([cx + 10, cy + 10, cx + 140, cy + 210], fill=(170, 125, 195, 220))
    # Center labellum / lip
    draw.ellipse([cx - 60, cy - 40, cx + 60, cy + 80], fill=(138, 60, 150, 250))
    draw.ellipse([cx - 30, cy, cx + 30, cy + 60], fill=(255, 215, 0, 255))
    img.save(path, "PNG")

flower1_path = os.path.join(ASSETS_DIR, "chrysanthemum.png")
flower2_path = os.path.join(ASSETS_DIR, "tulip.png")
flower3_path = os.path.join(ASSETS_DIR, "orchid.png")

create_chrysanthemum_image(flower1_path)
create_tulip_image(flower2_path)
create_orchid_image(flower3_path)

# -----------------------------------------------------------------------------
# PRESENTATION CONSTANTS & MORPH INJECTION
# -----------------------------------------------------------------------------
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

COLOR_BG_1 = RGBColor(250, 233, 210)   # Pale Peach/Yellow
COLOR_BG_2 = RGBColor(249, 192, 203)   # Soft Pink
COLOR_BG_3 = RGBColor(233, 219, 238)   # Soft Lilac

COLOR_NUM_1 = RGBColor(212, 105, 45)   # Terracotta Orange
COLOR_NUM_2 = RGBColor(196, 45, 80)    # Deep Rose
COLOR_NUM_3 = RGBColor(122, 75, 132)   # Deep Purple

COLOR_TEXT_1 = RGBColor(110, 80, 60)
COLOR_TEXT_2 = RGBColor(115, 55, 75)
COLOR_TEXT_3 = RGBColor(85, 60, 95)

def apply_morph_transition(slide):
    transition_xml = (
        '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="med">\n'
        '  <p14:morph option="byObject"/>\n'
        '</p:transition>'
    )
    slide._element.append(parse_xml(transition_xml))

# -----------------------------------------------------------------------------
# SLIDE BUILDER
# -----------------------------------------------------------------------------
prs = Presentation()
prs.slide_width = SLIDE_WIDTH
prs.slide_height = SLIDE_HEIGHT
blank_layout = prs.slide_layouts[6]

def build_slide(active_card):
    """
    active_card:
      0 -> Overview (all 3 equal width)
      1 -> Card 1 Expanded
      2 -> Card 2 Expanded
      3 -> Card 3 Expanded
    """
    slide = prs.slides.add_slide(blank_layout)
    apply_morph_transition(slide)

    # Geometry Layouts
    if active_card == 0:
        w1, w2, w3 = Inches(4.444), Inches(4.444), Inches(4.445)
        x1, x2, x3 = Inches(0), Inches(4.444), Inches(8.888)
    elif active_card == 1:
        w1, w2, w3 = Inches(9.333), Inches(2.0), Inches(2.0)
        x1, x2, x3 = Inches(0), Inches(9.333), Inches(11.333)
    elif active_card == 2:
        w1, w2, w3 = Inches(2.0), Inches(9.333), Inches(2.0)
        x1, x2, x3 = Inches(0), Inches(2.0), Inches(11.333)
    elif active_card == 3:
        w1, w2, w3 = Inches(2.0), Inches(2.0), Inches(9.333)
        x1, x2, x3 = Inches(0), Inches(2.0), Inches(4.0)

    # CARD 1 RECTANGLE
    rect1 = slide.shapes.add_shape(1, x1, Inches(0), w1, SLIDE_HEIGHT)
    rect1.name = "!!bg_card1"
    rect1.fill.solid()
    rect1.fill.fore_color.rgb = COLOR_BG_1
    rect1.line.fill.background()

    # CARD 2 RECTANGLE
    rect2 = slide.shapes.add_shape(1, x2, Inches(0), w2, SLIDE_HEIGHT)
    rect2.name = "!!bg_card2"
    rect2.fill.solid()
    rect2.fill.fore_color.rgb = COLOR_BG_2
    rect2.line.fill.background()

    # CARD 3 RECTANGLE
    rect3 = slide.shapes.add_shape(1, x3, Inches(0), w3, SLIDE_HEIGHT)
    rect3.name = "!!bg_card3"
    rect3.fill.solid()
    rect3.fill.fore_color.rgb = COLOR_BG_3
    rect3.line.fill.background()

    # ------------------- CARD 1 ELEMENTS -------------------
    if active_card == 1:
        img1 = slide.shapes.add_picture(flower1_path, Inches(0.4), Inches(1.3), Inches(4.6), Inches(4.6))
        
        # Number 1
        num1_box = slide.shapes.add_textbox(Inches(5.0), Inches(0.8), Inches(3.8), Inches(1.8))
        num1_box.name = "!!num_1"
        p = num1_box.text_frame.paragraphs[0]
        p.text = "1"
        p.font.name = "Segoe UI"
        p.font.size = Pt(110)
        p.font.bold = True
        p.font.color.rgb = COLOR_NUM_1

        # Description 1
        desc1_box = slide.shapes.add_textbox(Inches(4.6), Inches(2.8), Inches(4.3), Inches(3.6))
        desc1_box.name = "!!desc_1"
        tf = desc1_box.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = "رفيقة مثالية في اللحظات السعيدة مثل مناسبات الأعراس والفرح بكافة أشكاله."
        p1.font.name = "Calibri"
        p1.font.size = Pt(18)
        p1.font.color.rgb = COLOR_TEXT_1
        p1.alignment = PP_ALIGN.RIGHT
        p1.space_after = Pt(18)
        p2 = tf.add_paragraph()
        p2.text = "ناهيك عن رائحتها العطرية الفواحة التي تشبه رائحة الكافور الأسرة للقلب."
        p2.font.name = "Calibri"
        p2.font.size = Pt(18)
        p2.font.color.rgb = COLOR_TEXT_1
        p2.alignment = PP_ALIGN.RIGHT

        # Label
        lbl1 = slide.shapes.add_textbox(Inches(0.6), Inches(5.9), Inches(4.0), Inches(1.2))
        lbl1.name = "!!label_1"
        tf_l = lbl1.text_frame
        p_l1 = tf_l.paragraphs[0]
        p_l1.text = "Chrysanthemum"
        p_l1.font.name = "Georgia"
        p_l1.font.size = Pt(24)
        p_l1.font.italic = True
        p_l1.font.color.rgb = COLOR_NUM_1
        p_l1.alignment = PP_ALIGN.CENTER
        p_l2 = tf_l.add_paragraph()
        p_l2.text = "الأقحوان"
        p_l2.font.name = "Calibri"
        p_l2.font.size = Pt(16)
        p_l2.font.color.rgb = COLOR_TEXT_1
        p_l2.alignment = PP_ALIGN.CENTER
    else:
        scale = Inches(2.4) if active_card == 0 else Inches(1.4)
        top_pos = Inches(1.8) if active_card == 0 else Inches(2.6)
        img1 = slide.shapes.add_picture(flower1_path, x1 + (w1 - scale) / 2, top_pos, scale, scale)
        lbl1 = slide.shapes.add_textbox(x1 + Inches(0.1), top_pos + scale + Inches(0.4), w1 - Inches(0.2), Inches(1.2))
        lbl1.name = "!!label_1"
        tf_l = lbl1.text_frame
        p_l1 = tf_l.paragraphs[0]
        p_l1.text = "Chrysanthemum" if active_card == 0 else "Chrysan"
        p_l1.font.name = "Georgia"
        p_l1.font.size = Pt(20) if active_card == 0 else Pt(14)
        p_l1.font.italic = True
        p_l1.font.color.rgb = COLOR_NUM_1
        p_l1.alignment = PP_ALIGN.CENTER
        p_l2 = tf_l.add_paragraph()
        p_l2.text = "الأقحوان"
        p_l2.font.name = "Calibri"
        p_l2.font.size = Pt(15) if active_card == 0 else Pt(12)
        p_l2.font.color.rgb = COLOR_TEXT_1
        p_l2.alignment = PP_ALIGN.CENTER
    img1.name = "!!flower_1"

    # ------------------- CARD 2 ELEMENTS -------------------
    if active_card == 2:
        img2 = slide.shapes.add_picture(flower2_path, x2 + Inches(0.5), Inches(1.4), Inches(4.3), Inches(4.3))
        
        num2_box = slide.shapes.add_textbox(x2 + Inches(5.0), Inches(0.8), Inches(3.8), Inches(1.8))
        num2_box.name = "!!num_2"
        p = num2_box.text_frame.paragraphs[0]
        p.text = "2"
        p.font.name = "Segoe UI"
        p.font.size = Pt(110)
        p.font.bold = True
        p.font.color.rgb = COLOR_NUM_2

        desc2_box = slide.shapes.add_textbox(x2 + Inches(4.6), Inches(2.8), Inches(4.3), Inches(3.6))
        desc2_box.name = "!!desc_2"
        tf = desc2_box.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = "من أنواع الزهور طويلة الأمد. جميلة وذات رائحة عطرية فواحة جميلة ومنعشة."
        p1.font.name = "Calibri"
        p1.font.size = Pt(18)
        p1.font.color.rgb = COLOR_TEXT_2
        p1.alignment = PP_ALIGN.RIGHT
        p1.space_after = Pt(18)
        p2 = tf.add_paragraph()
        p2.text = "ناهيك عن ألوانها الزاهية التي تضفي لمسة مميزة للمكان"
        p2.font.name = "Calibri"
        p2.font.size = Pt(18)
        p2.font.color.rgb = COLOR_TEXT_2
        p2.alignment = PP_ALIGN.RIGHT

        lbl2 = slide.shapes.add_textbox(x2 + Inches(0.7), Inches(5.9), Inches(3.8), Inches(1.2))
        lbl2.name = "!!label_2"
        tf_l = lbl2.text_frame
        p_l1 = tf_l.paragraphs[0]
        p_l1.text = "Orchid"
        p_l1.font.name = "Georgia"
        p_l1.font.size = Pt(24)
        p_l1.font.italic = True
        p_l1.font.color.rgb = COLOR_NUM_2
        p_l1.alignment = PP_ALIGN.CENTER
        p_l2 = tf_l.add_paragraph()
        p_l2.text = "الأوركيد"
        p_l2.font.name = "Calibri"
        p_l2.font.size = Pt(16)
        p_l2.font.color.rgb = COLOR_TEXT_2
        p_l2.alignment = PP_ALIGN.CENTER
    else:
        scale = Inches(2.4) if active_card == 0 else Inches(1.4)
        top_pos = Inches(1.8) if active_card == 0 else Inches(2.6)
        img2 = slide.shapes.add_picture(flower2_path, x2 + (w2 - scale) / 2, top_pos, scale, scale)
        lbl2 = slide.shapes.add_textbox(x2 + Inches(0.1), top_pos + scale + Inches(0.4), w2 - Inches(0.2), Inches(1.2))
        lbl2.name = "!!label_2"
        tf_l = lbl2.text_frame
        p_l1 = tf_l.paragraphs[0]
        p_l1.text = "Orchid"
        p_l1.font.name = "Georgia"
        p_l1.font.size = Pt(20) if active_card == 0 else Pt(14)
        p_l1.font.italic = True
        p_l1.font.color.rgb = COLOR_NUM_2
        p_l1.alignment = PP_ALIGN.CENTER
        p_l2 = tf_l.add_paragraph()
        p_l2.text = "الأوركيد"
        p_l2.font.name = "Calibri"
        p_l2.font.size = Pt(15) if active_card == 0 else Pt(12)
        p_l2.font.color.rgb = COLOR_TEXT_2
        p_l2.alignment = PP_ALIGN.CENTER
    img2.name = "!!flower_2"

    # ------------------- CARD 3 ELEMENTS -------------------
    if active_card == 3:
        img3 = slide.shapes.add_picture(flower3_path, x3 + Inches(0.5), Inches(1.4), Inches(4.3), Inches(4.3))
        
        num3_box = slide.shapes.add_textbox(x3 + Inches(5.0), Inches(0.8), Inches(3.8), Inches(1.8))
        num3_box.name = "!!num_3"
        p = num3_box.text_frame.paragraphs[0]
        p.text = "3"
        p.font.name = "Segoe UI"
        p.font.size = Pt(110)
        p.font.bold = True
        p.font.color.rgb = COLOR_NUM_3

        desc3_box = slide.shapes.add_textbox(x3 + Inches(4.6), Inches(2.8), Inches(4.3), Inches(3.6))
        desc3_box.name = "!!desc_3"
        tf = desc3_box.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = "يأسر القلب بمظهره الناعم الأخاذ، فلها جمال راقٍ لا مثيل له يُعبر عن الحب والسعادة"
        p1.font.name = "Calibri"
        p1.font.size = Pt(18)
        p1.font.color.rgb = COLOR_TEXT_3
        p1.alignment = PP_ALIGN.RIGHT

        lbl3 = slide.shapes.add_textbox(x3 + Inches(0.7), Inches(5.9), Inches(3.8), Inches(1.2))
        lbl3.name = "!!label_3"
        tf_l = lbl3.text_frame
        p_l1 = tf_l.paragraphs[0]
        p_l1.text = "Tulip"
        p_l1.font.name = "Georgia"
        p_l1.font.size = Pt(24)
        p_l1.font.italic = True
        p_l1.font.color.rgb = COLOR_NUM_3
        p_l1.alignment = PP_ALIGN.CENTER
        p_l2 = tf_l.add_paragraph()
        p_l2.text = "التوليب"
        p_l2.font.name = "Calibri"
        p_l2.font.size = Pt(16)
        p_l2.font.color.rgb = COLOR_TEXT_3
        p_l2.alignment = PP_ALIGN.CENTER
    else:
        scale = Inches(2.4) if active_card == 0 else Inches(1.4)
        top_pos = Inches(1.8) if active_card == 0 else Inches(2.6)
        img3 = slide.shapes.add_picture(flower3_path, x3 + (w3 - scale) / 2, top_pos, scale, scale)
        lbl3 = slide.shapes.add_textbox(x3 + Inches(0.1), top_pos + scale + Inches(0.4), w3 - Inches(0.2), Inches(1.2))
        lbl3.name = "!!label_3"
        tf_l = lbl3.text_frame
        p_l1 = tf_l.paragraphs[0]
        p_l1.text = "Tulip"
        p_l1.font.name = "Georgia"
        p_l1.font.size = Pt(20) if active_card == 0 else Pt(14)
        p_l1.font.italic = True
        p_l1.font.color.rgb = COLOR_NUM_3
        p_l1.alignment = PP_ALIGN.CENTER
        p_l2 = tf_l.add_paragraph()
        p_l2.text = "التوليب"
        p_l2.font.name = "Calibri"
        p_l2.font.size = Pt(15) if active_card == 0 else Pt(12)
        p_l2.font.color.rgb = COLOR_TEXT_3
        p_l2.alignment = PP_ALIGN.CENTER
    img3.name = "!!flower_3"

# -----------------------------------------------------------------------------
# GENERATE PRESENTATION SLIDES
# -----------------------------------------------------------------------------
# Slide 1: Overview
build_slide(0)
# Slide 2: Card 1 Expanded
build_slide(1)
# Slide 3: Card 2 Expanded
build_slide(2)
# Slide 4: Card 3 Expanded
build_slide(3)
# Slide 5: Loop back to Overview
build_slide(0)

output_pptx_path = os.path.join(OUTPUT_DIR, "reconstructed_presentation.pptx")
prs.save(output_pptx_path)
print(f"[SUCCESS] Presentation generated successfully: {output_pptx_path}")
