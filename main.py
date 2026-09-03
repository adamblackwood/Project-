import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml import parse_xml

# -----------------------------------------------------------------------------
# مسارات الصور ومجلد الإخراج
# -----------------------------------------------------------------------------
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

IMG_HIBISCUS = "images/hibiscus.png"
IMG_LILY = "images/lily.png"
IMG_PEONY = "images/peony.png"

# -----------------------------------------------------------------------------
# لوحة الألوان المتناسقة مع الزهور الجديدة
# -----------------------------------------------------------------------------
COLOR_BG_1 = RGBColor(253, 244, 227)   # أصفر بيج دافئ
COLOR_BG_2 = RGBColor(252, 228, 236)   # وردي باستيل رقيق
COLOR_BG_3 = RGBColor(227, 238, 246)   # أزرق ثلجي هادئ للفاوانيا

COLOR_ACCENT_1 = RGBColor(218, 120, 23)   # برتقالي ذهبي
COLOR_ACCENT_2 = RGBColor(216, 67, 108)   # وردي توتي
COLOR_ACCENT_3 = RGBColor(35, 78, 112)    # أزرق كحلي ملكي

COLOR_TEXT_1 = RGBColor(90, 65, 40)
COLOR_TEXT_2 = RGBColor(95, 45, 65)
COLOR_TEXT_3 = RGBColor(40, 60, 80)

# -----------------------------------------------------------------------------
# إعداد العرض والانتقال التدريجي (Morph)
# -----------------------------------------------------------------------------
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

def apply_morph(slide):
    xml = (
        '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="med">\n'
        '  <p14:morph option="byObject"/>\n'
        '</p:transition>'
    )
    slide._element.append(parse_xml(xml))

def add_flower(slide, img_path, x, y, w, h, name):
    if os.path.exists(img_path):
        pic = slide.shapes.add_picture(img_path, x, y, w, h)
        pic.name = name
        return pic
    return None

# -----------------------------------------------------------------------------
# دالة بناء الشرائح
# -----------------------------------------------------------------------------
def build_slide(active_card):
    slide = prs.slides.add_slide(blank_layout)
    apply_morph(slide)

    # توزيع العرض حسب البطاقة النشطة
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

    # خلفيات البطاقات
    bg1 = slide.shapes.add_shape(1, x1, Inches(0), w1, Inches(7.5))
    bg1.name = "!!bg_card1"
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = COLOR_BG_1
    bg1.line.fill.background()

    bg2 = slide.shapes.add_shape(1, x2, Inches(0), w2, Inches(7.5))
    bg2.name = "!!bg_card2"
    bg2.fill.solid()
    bg2.fill.fore_color.rgb = COLOR_BG_2
    bg2.line.fill.background()

    bg3 = slide.shapes.add_shape(1, x3, Inches(0), w3, Inches(7.5))
    bg3.name = "!!bg_card3"
    bg3.fill.solid()
    bg3.fill.fore_color.rgb = COLOR_BG_3
    bg3.line.fill.background()

    # ==========================
    # عناصر البطاقة 1: الكركديه
    # ==========================
    if active_card == 1:
        add_flower(slide, IMG_HIBISCUS, Inches(0.5), Inches(1.3), Inches(4.5), Inches(4.5), "!!flower_1")
        
        num1 = slide.shapes.add_textbox(Inches(5.0), Inches(0.8), Inches(3.8), Inches(1.8))
        num1.name = "!!num_1"
        p = num1.text_frame.paragraphs[0]
        p.text = "1"
        p.font.name = "Segoe UI"
        p.font.size = Pt(110)
        p.font.bold = True
        p.font.color.rgb = COLOR_ACCENT_1

        desc1 = slide.shapes.add_textbox(Inches(4.7), Inches(2.8), Inches(4.3), Inches(3.6))
        desc1.name = "!!desc_1"
        tf = desc1.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = "رمز للإشراق والبهجة والجمال الاستوائي، تمنح الأجواء طاقة إيجابية وحيوية متجددة."
        p1.font.name = "Calibri"
        p1.font.size = Pt(18)
        p1.font.color.rgb = COLOR_TEXT_1
        p1.alignment = PP_ALIGN.RIGHT
        p1.space_after = Pt(18)

        p2 = tf.add_paragraph()
        p2.text = "تتميز ببتلاتها المتوهجة للشمس، وتعد خياراً مثالياً للاحتفاء بالبدايات السعيدة ولحظات الفرح."
        p2.font.name = "Calibri"
        p2.font.size = Pt(18)
        p2.font.color.rgb = COLOR_TEXT_1
        p2.alignment = PP_ALIGN.RIGHT

        lbl1 = slide.shapes.add_textbox(Inches(0.7), Inches(5.9), Inches(4.1), Inches(1.2))
        lbl1.name = "!!label_1"
        p_l1 = lbl1.text_frame.paragraphs[0]
        p_l1.text = "Yellow Hibiscus"
        p_l1.font.name = "Georgia"
        p_l1.font.size = Pt(24)
        p_l1.font.italic = True
        p_l1.font.color.rgb = COLOR_ACCENT_1
        p_l1.alignment = PP_ALIGN.CENTER
        p_l2 = lbl1.text_frame.add_paragraph()
        p_l2.text = "الكركديه الأصفر"
        p_l2.font.name = "Calibri"
        p_l2.font.size = Pt(16)
        p_l2.font.color.rgb = COLOR_TEXT_1
        p_l2.alignment = PP_ALIGN.CENTER
    else:
        sz = Inches(2.5) if active_card == 0 else Inches(1.5)
        ty = Inches(1.8) if active_card == 0 else Inches(2.6)
        add_flower(slide, IMG_HIBISCUS, x1 + (w1 - sz)/2, ty, sz, sz, "!!flower_1")
        lbl1 = slide.shapes.add_textbox(x1, ty + sz + Inches(0.3), w1, Inches(1.2))
        lbl1.name = "!!label_1"
        p_l1 = lbl1.text_frame.paragraphs[0]
        p_l1.text = "Hibiscus"
        p_l1.font.name = "Georgia"
        p_l1.font.size = Pt(20) if active_card == 0 else Pt(14)
        p_l1.font.italic = True
        p_l1.font.color.rgb = COLOR_ACCENT_1
        p_l1.alignment = PP_ALIGN.CENTER
        p_l2 = lbl1.text_frame.add_paragraph()
        p_l2.text = "الكركديه"
        p_l2.font.name = "Calibri"
        p_l2.font.size = Pt(15) if active_card == 0 else Pt(12)
        p_l2.font.color.rgb = COLOR_TEXT_1
        p_l2.alignment = PP_ALIGN.CENTER

    # ==========================
    # عناصر البطاقة 2: الزنبق
    # ==========================
    if active_card == 2:
        add_flower(slide, IMG_LILY, x2 + Inches(0.5), Inches(1.3), Inches(4.5), Inches(4.5), "!!flower_2")

        num2 = slide.shapes.add_textbox(x2 + Inches(5.0), Inches(0.8), Inches(3.8), Inches(1.8))
        num2.name = "!!num_2"
        p = num2.text_frame.paragraphs[0]
        p.text = "2"
        p.font.name = "Segoe UI"
        p.font.size = Pt(110)
        p.font.bold = True
        p.font.color.rgb = COLOR_ACCENT_2

        desc2 = slide.shapes.add_textbox(x2 + Inches(4.7), Inches(2.8), Inches(4.3), Inches(3.6))
        desc2.name = "!!desc_2"
        tf = desc2.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = "أيقونة الأناقة والنعومة والجمال الهادئ، وتُعبر عن المشاعر الصادقة والمودة الخالصة."
        p1.font.name = "Calibri"
        p1.font.size = Pt(18)
        p1.font.color.rgb = COLOR_TEXT_2
        p1.alignment = PP_ALIGN.RIGHT
        p1.space_after = Pt(18)

        p2 = tf.add_paragraph()
        p2.text = "بحضورها الملكي وعطرها الرقيق، تضفي لمسة ساحرة تلائم أرقى مناسبات التهنئة والمحبة."
        p2.font.name = "Calibri"
        p2.font.size = Pt(18)
        p2.font.color.rgb = COLOR_TEXT_2
        p2.alignment = PP_ALIGN.RIGHT

        lbl2 = slide.shapes.add_textbox(x2 + Inches(0.7), Inches(5.9), Inches(4.1), Inches(1.2))
        lbl2.name = "!!label_2"
        p_l1 = lbl2.text_frame.paragraphs[0]
        p_l1.text = "Pink Lily"
        p_l1.font.name = "Georgia"
        p_l1.font.size = Pt(24)
        p_l1.font.italic = True
        p_l1.font.color.rgb = COLOR_ACCENT_2
        p_l1.alignment = PP_ALIGN.CENTER
        p_l2 = lbl2.text_frame.add_paragraph()
        p_l2.text = "الزنبق الوردي"
        p_l2.font.name = "Calibri"
        p_l2.font.size = Pt(16)
        p_l2.font.color.rgb = COLOR_TEXT_2
        p_l2.alignment = PP_ALIGN.CENTER
    else:
        sz = Inches(2.5) if active_card == 0 else Inches(1.5)
        ty = Inches(1.8) if active_card == 0 else Inches(2.6)
        add_flower(slide, IMG_LILY, x2 + (w2 - sz)/2, ty, sz, sz, "!!flower_2")
        lbl2 = slide.shapes.add_textbox(x2, ty + sz + Inches(0.3), w2, Inches(1.2))
        lbl2.name = "!!label_2"
        p_l1 = lbl2.text_frame.paragraphs[0]
        p_l1.text = "Pink Lily"
        p_l1.font.name = "Georgia"
        p_l1.font.size = Pt(20) if active_card == 0 else Pt(14)
        p_l1.font.italic = True
        p_l1.font.color.rgb = COLOR_ACCENT_2
        p_l1.alignment = PP_ALIGN.CENTER
        p_l2 = lbl2.text_frame.add_paragraph()
        p_l2.text = "الزنبق"
        p_l2.font.name = "Calibri"
        p_l2.font.size = Pt(15) if active_card == 0 else Pt(12)
        p_l2.font.color.rgb = COLOR_TEXT_2
        p_l2.alignment = PP_ALIGN.CENTER

    # ==========================
    # عناصر البطاقة 3: الفاوانيا
    # ==========================
    if active_card == 3:
        add_flower(slide, IMG_PEONY, x3 + Inches(0.5), Inches(1.3), Inches(4.5), Inches(4.5), "!!flower_3")

        num3 = slide.shapes.add_textbox(x3 + Inches(5.0), Inches(0.8), Inches(3.8), Inches(1.8))
        num3.name = "!!num_3"
        p = num3.text_frame.paragraphs[0]
        p.text = "3"
        p.font.name = "Segoe UI"
        p.font.size = Pt(110)
        p.font.bold = True
        p.font.color.rgb = COLOR_ACCENT_3

        desc3 = slide.shapes.add_textbox(x3 + Inches(4.7), Inches(2.8), Inches(4.3), Inches(3.6))
        desc3.name = "!!desc_3"
        tf = desc3.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = "زهرة نادرة تجسد معاني السكينة والعمق والثقة، وتلفت الأنظار بندرتها وسحرها الفريد."
        p1.font.name = "Calibri"
        p1.font.size = Pt(18)
        p1.font.color.rgb = COLOR_TEXT_3
        p1.alignment = PP_ALIGN.RIGHT
        p1.space_after = Pt(18)

        p2 = tf.add_paragraph()
        p2.text = "تأسر القلوب بتموجات بتلاتها المخملية العميقة لتعكس فخامة عصرية لا مثيل لها."
        p2.font.name = "Calibri"
        p2.font.size = Pt(18)
        p2.font.color.rgb = COLOR_TEXT_3
        p2.alignment = PP_ALIGN.RIGHT

        lbl3 = slide.shapes.add_textbox(x3 + Inches(0.7), Inches(5.9), Inches(4.1), Inches(1.2))
        lbl3.name = "!!label_3"
        p_l1 = lbl3.text_frame.paragraphs[0]
        p_l1.text = "Blue Peony"
        p_l1.font.name = "Georgia"
        p_l1.font.size = Pt(24)
        p_l1.font.italic = True
        p_l1.font.color.rgb = COLOR_ACCENT_3
        p_l1.alignment = PP_ALIGN.CENTER
        p_l2 = lbl3.text_frame.add_paragraph()
        p_l2.text = "الفاوانيا الزرقاء"
        p_l2.font.name = "Calibri"
        p_l2.font.size = Pt(16)
        p_l2.font.color.rgb = COLOR_TEXT_3
        p_l2.alignment = PP_ALIGN.CENTER
    else:
        sz = Inches(2.5) if active_card == 0 else Inches(1.5)
        ty = Inches(1.8) if active_card == 0 else Inches(2.6)
        add_flower(slide, IMG_PEONY, x3 + (w3 - sz)/2, ty, sz, sz, "!!flower_3")
        lbl3 = slide.shapes.add_textbox(x3, ty + sz + Inches(0.3), w3, Inches(1.2))
        lbl3.name = "!!label_3"
        p_l1 = lbl3.text_frame.paragraphs[0]
        p_l1.text = "Blue Peony"
        p_l1.font.name = "Georgia"
        p_l1.font.size = Pt(20) if active_card == 0 else Pt(14)
        p_l1.font.italic = True
        p_l1.font.color.rgb = COLOR_ACCENT_3
        p_l1.alignment = PP_ALIGN.CENTER
        p_l2 = lbl3.text_frame.add_paragraph()
        p_l2.text = "الفاوانيا"
        p_l2.font.name = "Calibri"
        p_l2.font.size = Pt(15) if active_card == 0 else Pt(12)
        p_l2.font.color.rgb = COLOR_TEXT_3
        p_l2.alignment = PP_ALIGN.CENTER

# -----------------------------------------------------------------------------
# توليد الشرائح وحفظ العرض
# -----------------------------------------------------------------------------
build_slide(0)  # الشريحة 1: نظرة عامة متساوية
build_slide(1)  # الشريحة 2: تكبير الكركديه
build_slide(2)  # الشريحة 3: تكبير الزنبق
build_slide(3)  # الشريحة 4: تكبير الفاوانيا الزرقاء
build_slide(0)  # الشريحة 5: العودة للوضع المتساوي

output_path = os.path.join(OUTPUT_DIR, "reconstructed_presentation.pptx")
prs.save(output_path)
print(f"[تم بنجاح]: تم إنشاء العرض الجديد في {output_path}")
