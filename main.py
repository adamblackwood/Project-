"""
Cinema-Grade Expanding Accordion Card Slider with Native PowerPoint Morph Transitions.
Fully automated, CI/CD-compatible generation script.
"""

import math
import os
import struct
import wave
from pathlib import Path
from PIL import Image, ImageDraw

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn

# ==============================================================================
# CONFIGURATION & DESIGN SYSTEM CONSTANTS
# ==============================================================================

SLIDE_WIDTH = 13.333
SLIDE_HEIGHT = 7.500

BASE_DIR = Path(__file__).resolve().parent
IMAGES_DIR = BASE_DIR / "images"
OUTPUT_DIR = BASE_DIR / "output"
SOUNDS_DIR = OUTPUT_DIR / "sounds"

CARD_DATA = [
    {
        "id": 1,
        "key": "card_1",
        "name_en": "Yellow Hibiscus",
        "name_ar": "الكركديه الأصفر",
        "num": "1",
        "bg_rgb": RGBColor(253, 244, 227),     # #FDF4E3
        "accent_rgb": RGBColor(218, 120, 23),   # #DA7817
        "text_rgb": RGBColor(90, 65, 40),       # #5A4128
        "image_file": "hibiscus.png",
        "fallback_color": (245, 185, 30, 255),
        "para_1": "رمز للإشراق والبهجة والجمال الاستوائي، تمنح الأجواء طاقة إيجابية وحيوية متجددة",
        "para_2": "تتميز ببتلاتها المتوهجة للشمس، وتعد خياراً مثالياً للاحتفاء بالبدايات السعيدة ولحظات الفرح",
    },
    {
        "id": 2,
        "key": "card_2",
        "name_en": "Pink Lily",
        "name_ar": "الزنبق الوردي",
        "num": "2",
        "bg_rgb": RGBColor(252, 228, 236),     # #FCE4EC
        "accent_rgb": RGBColor(216, 67, 108),   # #D8436C
        "text_rgb": RGBColor(95, 45, 65),       # #5F2D41
        "image_file": "lily.png",
        "fallback_color": (235, 110, 145, 255),
        "para_1": "أيقونة الأناقة والنعومة والجمال الهادئ، وتُعبر عن المشاعر الصادقة والمودة الخالصة",
        "para_2": "بحضورها الملكي وعطرها الرقيق، تضفي لمسة ساحرة تلائم أرقى مناسبات التهنئة والمحبة",
    },
    {
        "id": 3,
        "key": "card_3",
        "name_en": "Blue Peony",
        "name_ar": "الفاوانيا الزرقاء",
        "num": "3",
        "bg_rgb": RGBColor(227, 238, 246),     # #E3EEF6
        "accent_rgb": RGBColor(35, 78, 112),    # #234E70
        "text_rgb": RGBColor(40, 60, 80),       # #283C50
        "image_file": "peony.png",
        "fallback_color": (55, 120, 180, 255),
        "para_1": "زهرة نادرة تجسد معاني السكينة والعمق والثقة، وتلفت الأنظار بندرتها وسحرها الفريد",
        "para_2": "تأسر القلوب بتموجات بتلاتها المخملية العميقة لتعكس فخامة عصرية لا مثيل لها",
    },
]

# ==============================================================================
# PROCEDURAL AUDIO SYNTHESIS
# ==============================================================================

def generate_audio_effects():
    """Synthesizes high-fidelity 44.1 kHz WAV audio files without external assets."""
    SOUNDS_DIR.mkdir(parents=True, exist_ok=True)
    sample_rate = 44100

    # 1. Soft Breeze Transition (1.20 seconds filtered multi-tone swoosh)
    breeze_path = SOUNDS_DIR / "soft_breeze_transition.wav"
    duration_breeze = 1.20
    total_samples_breeze = int(sample_rate * duration_breeze)
    with wave.open(str(breeze_path), "wb") as wav:
        wav.setnchannels(1)
        wav.setsampwidth(2)
        wav.setframerate(sample_rate)
        frames = bytearray()
        for i in range(total_samples_breeze):
            t = i / sample_rate
            # Dynamic envelope: smooth bell curve
            envelope = math.sin(math.pi * (t / duration_breeze)) ** 2
            # Frequency modulation sweeps upward then resolves
            freq = 240.0 + 180.0 * math.sin(math.pi * (t / duration_breeze))
            freq_sub = 120.0 + 80.0 * math.sin(math.pi * (t / duration_breeze))
            sample = (
                0.60 * math.sin(2.0 * math.pi * freq * t)
                + 0.30 * math.sin(2.0 * math.pi * freq_sub * t)
                + 0.10 * math.sin(2.0 * math.pi * (freq * 2.2) * t)
            )
            val = int(sample * envelope * 24000.0)
            val = max(-32767, min(32767, val))
            frames.extend(struct.pack("<h", val))
        wav.writeframes(frames)

    # 2. Floral Ambient Chime (1.80 seconds crystalline harmonic decay)
    chime_path = SOUNDS_DIR / "floral_ambient_chime.wav"
    duration_chime = 1.80
    total_samples_chime = int(sample_rate * duration_chime)
    with wave.open(str(chime_path), "wb") as wav:
        wav.setnchannels(1)
        wav.setsampwidth(2)
        wav.setframerate(sample_rate)
        frames = bytearray()
        partials = [(523.25, 0.40), (659.25, 0.30), (783.99, 0.20), (1046.50, 0.10)]
        for i in range(total_samples_chime):
            t = i / sample_rate
            decay = math.exp(-3.2 * t)
            sample = sum(amp * math.sin(2.0 * math.pi * freq * t) for freq, amp in partials)
            val = int(sample * decay * 28000.0)
            val = max(-32767, min(32767, val))
            frames.extend(struct.pack("<h", val))
        wav.writeframes(frames)

# ==============================================================================
# FALLBACK BOTANICAL GRAPHIC GENERATION
# ==============================================================================

def generate_fallback_botanical_image(filepath: Path, color_rgba: tuple):
    """Draws a high-res, transparent alpha-channel botanical emblem using Pillow."""
    size = (1000, 1000)
    img = Image.new("RGBA", size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    cx, cy = 500, 500

    # Draw a rosette of 8 petals
    num_petals = 8
    petal_len = 320
    petal_wid = 140
    for idx in range(num_petals):
        angle = idx * (2 * math.pi / num_petals)
        px = cx + int(math.cos(angle) * (petal_len * 0.55))
        py = cy + int(math.sin(angle) * (petal_len * 0.55))
        box = [px - petal_wid, py - petal_wid, px + petal_wid, py + petal_wid]
        draw.ellipse(box, fill=color_rgba)

    # Core center disk
    core_radius = 110
    core_color = (255, 255, 255, 240)
    draw.ellipse(
        [cx - core_radius, cy - core_radius, cx + core_radius, cy + core_radius],
        fill=core_color,
    )
    # Inner pistil details
    inner_rad = 60
    draw.ellipse(
        [cx - inner_rad, cy - inner_rad, cx + inner_rad, cy + inner_rad],
        fill=color_rgba,
    )

    filepath.parent.mkdir(parents=True, exist_ok=True)
    img.save(filepath, format="PNG")

def verify_or_create_assets():
    """Ensures botanical transparent assets exist before PPTX assembly."""
    IMAGES_DIR.mkdir(parents=True, exist_ok=True)
    for card in CARD_DATA:
        img_path = IMAGES_DIR / card["image_file"]
        if not img_path.exists():
            generate_fallback_botanical_image(img_path, card["fallback_color"])

# ==============================================================================
# NATIVE PPTX MORPH & RTL XML UTILITIES
# ==============================================================================

def assign_morph_identifier(shape, morph_name: str):
    """Enforces PPTX morph pairing by setting the shape name with the '!!' prefix."""
    shape.name = morph_name

def inject_native_morph_transition(slide):
    """Injects native PowerPoint Morph transition and Breeze sound into slide XML."""
    transition_xml = (
        '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
        'spd="med" advClick="1">\n'
        '  <p14:morph option="byObject"/>\n'
        '  <p:sndAc>\n'
        '    <p:stSnd>\n'
        '      <p:snd name="Breeze"/>\n'
        '    </p:stSnd>\n'
        '  </p:sndAc>\n'
        '</p:transition>'
    )
    slide_elem = slide._element

    # Replace existing p:transition if present, otherwise insert before p:extLst
    for child in list(slide_elem):
        if child.tag.endswith("transition"):
            slide_elem.remove(child)

    trans_elem = parse_xml(transition_xml)
    extLst = slide_elem.find(qn("p:extLst"))
    if extLst is not None:
        slide_elem.insert(slide_elem.index(extLst), trans_elem)
    else:
        slide_elem.append(trans_elem)

def set_paragraph_rtl(paragraph):
    """Configures paragraph text properties to strict Right-to-Left (RTL)."""
    pPr = paragraph._element.get_or_add_pPr()
    pPr.set("rtl", "1")
    paragraph.alignment = PP_ALIGN.RIGHT

# ==============================================================================
# SLIDE BUILDER ENGINE
# ==============================================================================

def compute_layout(active_index: int):
    """
    Computes precise bounding coordinates for State 0, 1, 2, and 3.
    active_index:
       0 -> Overview (State 0: balanced columns)
       1 -> Card 1 Hero (State 1)
       2 -> Card 2 Hero (State 2)
       3 -> Card 3 Hero (State 3)
    """
    layout = {}
    if active_index == 0:
        # State 0: 3 Balanced Cards
        widths = [4.444, 4.444, 4.445]
        xs = [0.000, 4.444, 8.888]
        for i in range(3):
            layout[i + 1] = {
                "card_x": xs[i],
                "card_w": widths[i],
                "is_hero": False,
                "img_x": xs[i] + (widths[i] - 2.40) / 2.0,
                "img_y": 1.40,
                "img_size": 2.40,
                "num_x": xs[i] + 0.35,
                "num_y": 4.05,
                "num_w": widths[i] - 0.70,
                "num_size": 96,
                "num_align": PP_ALIGN.CENTER,
                "title_x": xs[i] + 0.35,
                "title_y": 5.45,
                "title_w": widths[i] - 0.70,
                "title_align": PP_ALIGN.CENTER,
            }
    else:
        # Hero Accordion States
        hero_id = active_index
        cur_x = 0.000
        for cid in range(1, 4):
            if cid == hero_id:
                w = 9.133
                layout[cid] = {
                    "card_x": cur_x,
                    "card_w": w,
                    "is_hero": True,
                    "img_x": cur_x + 0.50,
                    "img_y": 1.25,
                    "img_size": 4.70,
                    "num_x": cur_x + 5.30,
                    "num_y": 0.55,
                    "num_w": 3.40,
                    "num_size": 110,
                    "num_align": PP_ALIGN.RIGHT,
                    "title_x": cur_x + 5.30,
                    "title_y": 2.35,
                    "title_w": 3.40,
                    "title_align": PP_ALIGN.RIGHT,
                    "body_x": cur_x + 5.30,
                    "body_y": 3.75,
                    "body_w": 3.40,
                }
            else:
                w = 2.100
                layout[cid] = {
                    "card_x": cur_x,
                    "card_w": w,
                    "is_hero": False,
                    "img_x": cur_x + (w - 1.50) / 2.0,
                    "img_y": 1.40,
                    "img_size": 1.50,
                    "num_x": cur_x + 0.15,
                    "num_y": 3.40,
                    "num_w": w - 0.30,
                    "num_size": 72,
                    "num_align": PP_ALIGN.CENTER,
                    "title_x": cur_x + 0.15,
                    "title_y": 4.80,
                    "title_w": w - 0.30,
                    "title_align": PP_ALIGN.CENTER,
                }
            cur_x += w
    return layout

def render_slide(prs, blank_layout, state_id: int):
    """Renders an individual keyframe slide with consistent Morph naming."""
    slide = prs.slides.add_slide(blank_layout)
    inject_native_morph_transition(slide)
    layout = compute_layout(state_id)

    # 1. Base Card Shapes
    for card in CARD_DATA:
        cid = card["id"]
        geo = layout[cid]

        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(geo["card_x"]),
            Inches(0.0),
            Inches(geo["card_w"]),
            Inches(SLIDE_HEIGHT),
        )
        assign_morph_identifier(rect, f"!!card_{cid}")
        rect.fill.solid()
        rect.fill.fore_color.rgb = card["bg_rgb"]
        rect.line.fill.background()  # No border

    # 2. Flower Botanical Images
    for card in CARD_DATA:
        cid = card["id"]
        geo = layout[cid]
        img_path = IMAGES_DIR / card["image_file"]

        pic = slide.shapes.add_picture(
            str(img_path),
            Inches(geo["img_x"]),
            Inches(geo["img_y"]),
            width=Inches(geo["img_size"]),
            height=Inches(geo["img_size"]),
        )
        assign_morph_identifier(pic, f"!!flower_{cid}")

    # 3. Numeric Indicators
    for card in CARD_DATA:
        cid = card["id"]
        geo = layout[cid]

        tb = slide.shapes.add_textbox(
            Inches(geo["num_x"]),
            Inches(geo["num_y"]),
            Inches(geo["num_w"]),
            Inches(1.50),
        )
        assign_morph_identifier(tb, f"!!num_{cid}")
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p = tf.paragraphs[0]
        p.text = card["num"]
        p.font.name = "Segoe UI"
        p.font.size = Pt(geo["num_size"])
        p.font.bold = True
        p.font.color.rgb = card["accent_rgb"]
        p.alignment = geo["num_align"]

    # 4. Title Text (English & Arabic)
    for card in CARD_DATA:
        cid = card["id"]
        geo = layout[cid]

        tb = slide.shapes.add_textbox(
            Inches(geo["title_x"]),
            Inches(geo["title_y"]),
            Inches(geo["title_w"]),
            Inches(1.50),
        )
        assign_morph_identifier(tb, f"!!title_{cid}")
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        # English Title
        p_en = tf.paragraphs[0]
        p_en.text = card["name_en"]
        p_en.font.name = "Georgia"
        p_en.font.size = Pt(22 if geo["is_hero"] else 15)
        p_en.font.italic = True
        p_en.font.bold = True
        p_en.font.color.rgb = card["text_rgb"]
        p_en.alignment = geo["title_align"]

        # Arabic Title
        p_ar = tf.add_paragraph()
        p_ar.text = card["name_ar"]
        p_ar.font.name = "Calibri"
        p_ar.font.size = Pt(17 if geo["is_hero"] else 13)
        p_ar.font.bold = True
        p_ar.font.color.rgb = card["accent_rgb"]
        p_ar.alignment = geo["title_align"]
        if geo["title_align"] == PP_ALIGN.RIGHT:
            set_paragraph_rtl(p_ar)

    # 5. Descriptive Arabic Content (Only displayed in Expanded State)
    for card in CARD_DATA:
        cid = card["id"]
        geo = layout[cid]

        # In collapsed state, create an off-screen/zero-opacity paired shape
        # to ensure seamless morph interpolation into the hero state
        if geo["is_hero"]:
            desc_x = Inches(geo["body_x"])
            desc_y = Inches(geo["body_y"])
            desc_w = Inches(geo["body_w"])
            desc_h = Inches(3.00)
        else:
            desc_x = Inches(geo["card_x"] + 0.10)
            desc_y = Inches(7.60)  # Hidden just beyond canvas
            desc_w = Inches(max(0.50, geo["card_w"] - 0.20))
            desc_h = Inches(0.50)

        tb = slide.shapes.add_textbox(desc_x, desc_y, desc_w, desc_h)
        assign_morph_identifier(tb, f"!!desc_{cid}")
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = card["para_1"]
        p1.font.name = "Calibri"
        p1.font.size = Pt(13.5)
        p1.font.color.rgb = card["text_rgb"]
        set_paragraph_rtl(p1)

        p2 = tf.add_paragraph()
        p2.text = card["para_2"]
        p2.font.name = "Calibri"
        p2.font.size = Pt(13.5)
        p2.font.color.rgb = card["text_rgb"]
        p2.space_before = Pt(10)
        set_paragraph_rtl(p2)

# ==============================================================================
# MAIN EXECUTION PIPELINE
# ==============================================================================

def main():
    print("[1/4] Preparing directories and synthesizing procedural audio...")
    generate_audio_effects()

    print("[2/4] Verifying and generating botanical graphics...")
    verify_or_create_assets()

    print("[3/4] Initializing widescreen presentation canvas...")
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH)
    prs.slide_height = Inches(SLIDE_HEIGHT)
    blank_layout = prs.slide_layouts[6]  # Clean blank layout

    print("[4/4] Generating accordion states with native Morph transitions...")
    # 5 Sequential Slides:
    # 0: State 0 (Balanced Overview)
    # 1: State 1 (Card 1 Hero)
    # 2: State 2 (Card 2 Hero)
    # 3: State 3 (Card 3 Hero)
    # 4: State 0 (Loop back to Overview)
    slides_sequence = [0, 1, 2, 3, 0]
    for step_num, state_id in enumerate(slides_sequence, start=1):
        print(f"      -> Building Slide {step_num} (State {state_id})...")
        render_slide(prs, blank_layout, state_id)

    output_pptx = OUTPUT_DIR / "reconstructed_presentation.pptx"
    prs.save(str(output_pptx))
    print(f"\n[SUCCESS] Presentation generated successfully: {output_pptx}")

if __name__ == "__main__":
    main()
