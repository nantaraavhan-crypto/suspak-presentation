from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Colors
GREEN_DARK = RGBColor(0x00, 0x6B, 0x3F)
GREEN_MID = RGBColor(0x00, 0x8C, 0x5A)
GREEN_LIGHT = RGBColor(0x4E, 0xC9, 0x9B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x0A, 0x0A, 0x0A)
GRAY = RGBColor(0x66, 0x66, 0x66)
LIGHT_BG = RGBColor(0xF8, 0xFA, 0xF9)

def add_gradient_bg(slide, color1, color2):
    background = slide.background
    fill = background.fill
    fill.gradient()
    fill.gradient_stops[0].color.rgb = color1
    fill.gradient_stops[1].color.rgb = color2

def add_shape(slide, left, top, width, height, color, shape_type=MSO_SHAPE.RECTANGLE):
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape

def add_rounded_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape

def add_text(slide, left, top, width, height, text, size=18, bold=False, color=BLACK, align=PP_ALIGN.LEFT, font='Calibri'):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font
    p.alignment = align
    return txBox

def add_para(tf, text, size=14, bold=False, color=BLACK, align=PP_ALIGN.LEFT, space=Pt(6)):
    p = tf.add_paragraph()
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = 'Calibri'
    p.alignment = align
    p.space_before = space
    return p

# ==================== SLIDE 1: TITLE ====================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_gradient_bg(slide, RGBColor(0x00, 0x1A, 0x0D), RGBColor(0x00, 0x33, 0x20))

# Decorative circle
circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9), Inches(-1), Inches(5), Inches(5))
circle.fill.solid()
circle.fill.fore_color.rgb = RGBColor(0x00, 0x4D, 0x2C)
circle.line.fill.background()

# Badge
badge = add_rounded_rect(slide, Inches(4.5), Inches(1.5), Inches(4.3), Inches(0.6), RGBColor(0x00, 0x33, 0x20))
add_text(slide, Inches(4.5), Inches(1.55), Inches(4.3), Inches(0.5), "ENACTUS IIT BOMBAY", 12, False, GREEN_LIGHT, PP_ALIGN.CENTER)

# Main Title
add_text(slide, Inches(2), Inches(2.5), Inches(9.3), Inches(1.5), "SUSPAK", 72, True, WHITE, PP_ALIGN.CENTER, 'Space Grotesk')

# Subtitle
add_text(slide, Inches(2), Inches(4), Inches(9.3), Inches(0.6), "PIONEERING A FUTURE WHERE PACKAGING HEALS THE PLANET", 18, False, GREEN_LIGHT, PP_ALIGN.CENTER)

# Tagline
add_text(slide, Inches(3), Inches(5), Inches(7.3), Inches(0.8), "Biodegradable food packaging solutions made from sugarcane bagasse & wood pulp", 14, False, RGBColor(0x99, 0x99, 0x99), PP_ALIGN.CENTER)

# Bottom bar
add_shape(slide, Inches(0), Inches(7), Inches(13.333), Inches(0.5), GREEN_MID)


# ==================== SLIDE 2: PROBLEM ====================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_gradient_bg(slide, WHITE, LIGHT_BG)

# Hero section with green overlay
hero = add_shape(slide, Inches(0), Inches(0), Inches(13.333), Inches(2.5), GREEN_DARK)
add_text(slide, Inches(0.8), Inches(0.8), Inches(11.7), Inches(1), "THE PROBLEM", 42, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')
add_text(slide, Inches(0.8), Inches(1.6), Inches(11.7), Inches(0.5), "Single-use plastic dominates food packaging today", 16, False, RGBColor(0xCC, 0xFF, 0xDD), PP_ALIGN.LEFT)

# Stats row
stats = [
    ("$220 Bn", "Global plastic food\npackaging market"),
    ("60%", "Of global food packaging\nis plastic"),
    ("30-90 Days", "Decomposition time\nfor our solution"),
]
x = 0.8
for num, lbl in stats:
    card = add_rounded_rect(slide, Inches(x), Inches(3), Inches(3.8), Inches(1.5), GREEN_DARK)
    add_text(slide, Inches(x + 0.3), Inches(3.1), Inches(3.2), Inches(0.7), num, 36, True, GREEN_LIGHT, PP_ALIGN.LEFT, 'Space Grotesk')
    add_text(slide, Inches(x + 0.3), Inches(3.8), Inches(3.2), Inches(0.6), lbl, 11, False, RGBColor(0xCC, 0xFF, 0xDD), PP_ALIGN.LEFT)
    x += 4.1

# Problem cards
problems = [
    ("Damages Harvest", "Plastic pollution damages crops and soil quality"),
    ("Soil Degradation", "Microplastics destroy soil microbial health"),
    ("Manufacturing", "Energy-intensive plastic production"),
    ("Health Hazards", "Chemical leaching into food chain"),
]
x = 0.8
for title, desc in problems:
    card = add_rounded_rect(slide, Inches(x), Inches(4.8), Inches(2.9), Inches(1.2), LIGHT_BG)
    add_text(slide, Inches(x + 0.2), Inches(4.9), Inches(2.5), Inches(0.4), title, 13, True, GREEN_DARK, PP_ALIGN.LEFT)
    add_text(slide, Inches(x + 0.2), Inches(5.3), Inches(2.5), Inches(0.6), desc, 10, False, GRAY, PP_ALIGN.LEFT)
    x += 3.1

# Global push
push = [
    ("REGULATIONS", "Extended Producer Responsibility mandates"),
    ("AWARENESS", "66% consumers care about sustainability"),
    ("COMMITMENTS", "100% sustainable packaging by 2030"),
]
x = 0.8
for title, desc in push:
    card = add_rounded_rect(slide, Inches(x), Inches(6.2), Inches(3.8), Inches(0.9), GREEN_MID)
    add_text(slide, Inches(x + 0.2), Inches(6.25), Inches(3.4), Inches(0.3), title, 12, True, WHITE, PP_ALIGN.CENTER)
    add_text(slide, Inches(x + 0.2), Inches(6.55), Inches(3.4), Inches(0.4), desc, 10, False, RGBColor(0xDD, 0xFF, 0xDD), PP_ALIGN.CENTER)
    x += 4.1


# ==================== SLIDE 3: INDUSTRY ====================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_gradient_bg(slide, LIGHT_BG, WHITE)

# Header
header = add_shape(slide, Inches(0.5), Inches(0.5), Inches(12.3), Inches(1), GREEN_DARK)
add_text(slide, Inches(0.8), Inches(0.6), Inches(11.7), Inches(0.8), "INDUSTRY LANDSCAPE", 32, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')

# TAM/SAM/SOM circles
for i, (size, color, label, value) in enumerate([(4, GREEN_LIGHT, "", ""), (3, GREEN_MID, "", ""), (2, GREEN_DARK, "SOM", "$38.27M")]):
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.2 - i*0.3), Inches(2.2 - i*0.3), Inches(size), Inches(size))
    circle.fill.solid()
    circle.fill.fore_color.rgb = color
    circle.line.fill.background()
    if label:
        add_text(slide, Inches(2.5), Inches(3.2), Inches(1.5), Inches(0.3), label, 12, False, RGBColor(0xCC, 0xFF, 0xDD), PP_ALIGN.CENTER)
        add_text(slide, Inches(2.5), Inches(3.5), Inches(1.5), Inches(0.4), value, 18, True, WHITE, PP_ALIGN.CENTER, 'Space Grotesk')

# Justification box
just_box = add_rounded_rect(slide, Inches(5), Inches(2), Inches(3.8), Inches(3.5), WHITE)
add_text(slide, Inches(5.3), Inches(2.1), Inches(3.2), Inches(0.5), "MARKET JUSTIFICATION", 16, True, GREEN_DARK, PP_ALIGN.LEFT, 'Space Grotesk')

just_items = [
    "TAM: Global Food Packaging - $400.9B",
    "SAM: Indian Food Packaging - $38.27B",
    "SOM: 0.1% penetration = $38.27M",
    "First Customer: Mirchi & Mime",
]
txBox = slide.shapes.add_textbox(Inches(5.3), Inches(2.7), Inches(3.2), Inches(2.5))
tf = txBox.text_frame
tf.word_wrap = True
for i, text in enumerate(just_items):
    if i == 0:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    p.text = f"✓  {text}"
    p.font.size = Pt(11)
    p.font.color.rgb = GRAY
    p.font.name = 'Calibri'
    p.space_before = Pt(10)

# Cost comparison
cost_box = add_rounded_rect(slide, Inches(9.2), Inches(2), Inches(3.6), Inches(3.5), WHITE)
add_text(slide, Inches(9.5), Inches(2.1), Inches(3), Inches(0.5), "COST COMPARISON", 16, True, GREEN_DARK, PP_ALIGN.LEFT, 'Space Grotesk')

costs = [("Traditional Plastic", 0.25, RGBColor(0xCC, 0xCC, 0xCC)), ("Molded Fiber", 0.35, GREEN_LIGHT), ("Seaweed-based", 0.65, GREEN_MID), ("Polylactic Acid", 0.85, GREEN_DARK)]
y = 2.7
for name, width, color in costs:
    add_text(slide, Inches(9.5), Inches(y), Inches(3), Inches(0.2), name, 9, False, GRAY, PP_ALIGN.LEFT)
    bar = add_rounded_rect(slide, Inches(9.5), Inches(y + 0.25), Inches(width * 2.8), Inches(0.3), color)
    add_text(slide, Inches(9.5), Inches(y + 0.25), Inches(width * 2.8), Inches(0.3), f"{width/0.25:.1f}x", 8, True, WHITE, PP_ALIGN.CENTER)
    y += 0.65


# ==================== SLIDE 4: SOLUTION ====================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_gradient_bg(slide, WHITE, LIGHT_BG)

# Header
header = add_shape(slide, Inches(0.5), Inches(0.5), Inches(12.3), Inches(1), GREEN_DARK)
add_text(slide, Inches(0.8), Inches(0.6), Inches(11.7), Inches(0.8), "PROPOSED SOLUTION", 32, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')

# Product images section
# Tray image placeholder
tray_box = add_rounded_rect(slide, Inches(0.8), Inches(1.8), Inches(5.8), Inches(2.5), LIGHT_BG)
add_text(slide, Inches(1), Inches(2.8), Inches(5.4), Inches(0.5), "3-Compartment Biodegradable Tray", 16, True, GREEN_DARK, PP_ALIGN.CENTER, 'Space Grotesk')
add_text(slide, Inches(1), Inches(3.3), Inches(5.4), Inches(0.4), "Sugarcane bagasse with wood pulp reinforcement", 11, False, GRAY, PP_ALIGN.CENTER)
add_text(slide, Inches(1.5), Inches(3.8), Inches(4.4), Inches(0.3), "[Insert tray image here]", 10, False, RGBColor(0xBB, 0xBB, 0xBB), PP_ALIGN.CENTER)

# Plate image placeholder
plate_box = add_rounded_rect(slide, Inches(6.8), Inches(1.8), Inches(5.8), Inches(2.5), LIGHT_BG)
add_text(slide, Inches(7), Inches(2.8), Inches(5.4), Inches(0.5), "Round Partitioned Plate", 16, True, GREEN_DARK, PP_ALIGN.CENTER, 'Space Grotesk')
add_text(slide, Inches(7), Inches(3.3), Inches(5.4), Inches(0.4), "Natural fiber composition, fully compostable", 11, False, GRAY, PP_ALIGN.CENTER)
add_text(slide, Inches(7.5), Inches(3.8), Inches(4.4), Inches(0.3), "[Insert plate image here]", 10, False, RGBColor(0xBB, 0xBB, 0xBB), PP_ALIGN.CENTER)

# Features
features = [
    ("01", "COMPOSTABLE", "Decomposes in 30-90 days"),
    ("02", "PEELABLE PLA LINER", "Removable for industrial composting"),
    ("03", "HEAT RESISTANT", "Withstands up to 150°C"),
    ("04", "LIGHTWEIGHT", "Wood pulp ensures durability"),
    ("05", "WATER RESISTANT", "Suitable for wet foods"),
    ("06", "COST-COMPETITIVE", "1.3x-1.5x price of plastic"),
]
x = 0.8
for num, title, desc in features:
    card = add_rounded_rect(slide, Inches(x), Inches(4.6), Inches(1.9), Inches(1.8), WHITE)
    num_circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.6), Inches(4.7), Inches(0.6), Inches(0.6))
    num_circle.fill.solid()
    num_circle.fill.fore_color.rgb = GREEN_DARK
    num_circle.line.fill.background()
    add_text(slide, Inches(x + 0.6), Inches(4.75), Inches(0.6), Inches(0.5), num, 14, True, WHITE, PP_ALIGN.CENTER, 'Space Grotesk')
    add_text(slide, Inches(x + 0.1), Inches(5.4), Inches(1.7), Inches(0.4), title, 10, True, GREEN_DARK, PP_ALIGN.CENTER)
    add_text(slide, Inches(x + 0.1), Inches(5.8), Inches(1.7), Inches(0.5), desc, 9, False, GRAY, PP_ALIGN.CENTER)
    x += 2.05


# ==================== SLIDE 5: COATINGS ====================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_gradient_bg(slide, WHITE, LIGHT_BG)

# Header
header = add_shape(slide, Inches(0.5), Inches(0.5), Inches(12.3), Inches(1), GREEN_DARK)
add_text(slide, Inches(0.8), Inches(0.6), Inches(11.7), Inches(0.8), "EXPLORING COATING MATERIALS", 32, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')

# PLA Card
pla_card = add_rounded_rect(slide, Inches(0.8), Inches(1.8), Inches(5.8), Inches(5.2), GREEN_DARK)
add_text(slide, Inches(1.2), Inches(2), Inches(5), Inches(0.6), "POLYLACTIC ACID (PLA)", 24, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')

pla_features = ["Thermoplastic biopolymer", "Water-resistant surface layer", "Industrial composting compatible", "Continuous film formation"]
txBox = slide.shapes.add_textbox(Inches(1.2), Inches(2.8), Inches(5), Inches(3))
tf = txBox.text_frame
tf.word_wrap = True
for i, feat in enumerate(pla_features):
    if i == 0:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    p.text = f"→  {feat}"
    p.font.size = Pt(14)
    p.font.color.rgb = WHITE
    p.font.name = 'Calibri'
    p.space_before = Pt(14)

# PLA image placeholder
pla_img = add_rounded_rect(slide, Inches(1.2), Inches(5.2), Inches(5), Inches(1.5), RGBColor(0x00, 0x4D, 0x2C))
add_text(slide, Inches(1.5), Inches(5.6), Inches(4.4), Inches(0.5), "[Insert PLA pellets image here]", 11, False, RGBColor(0x99, 0xFF, 0xCC), PP_ALIGN.CENTER)

# Seaweed Card
sw_card = add_rounded_rect(slide, Inches(6.8), Inches(1.8), Inches(5.8), Inches(5.2), GREEN_MID)
add_text(slide, Inches(7.2), Inches(2), Inches(5), Inches(0.6), "SEAWEED-BASED", 24, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')

sw_features = ["Natural polysaccharide coating", "Edible & biodegradable", "Antimicrobial properties", "Renewable ocean resource"]
txBox = slide.shapes.add_textbox(Inches(7.2), Inches(2.8), Inches(5), Inches(3))
tf = txBox.text_frame
tf.word_wrap = True
for i, feat in enumerate(sw_features):
    if i == 0:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    p.text = f"→  {feat}"
    p.font.size = Pt(14)
    p.font.color.rgb = WHITE
    p.font.name = 'Calibri'
    p.space_before = Pt(14)

# Seaweed image placeholder
sw_img = add_rounded_rect(slide, Inches(7.2), Inches(5.2), Inches(5), Inches(1.5), RGBColor(0x00, 0x6B, 0x3F))
add_text(slide, Inches(7.5), Inches(5.6), Inches(4.4), Inches(0.5), "[Insert seaweed image here]", 11, False, RGBColor(0xCC, 0xFF, 0xDD), PP_ALIGN.CENTER)


# ==================== SLIDE 6: PLA PROCESS ====================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_gradient_bg(slide, WHITE, LIGHT_BG)

# Header
header = add_shape(slide, Inches(0.5), Inches(0.5), Inches(12.3), Inches(1), GREEN_DARK)
add_text(slide, Inches(0.8), Inches(0.6), Inches(11.7), Inches(0.8), "PLA COATING PROCESS", 32, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')

# Process diagram placeholder
diag_box = add_rounded_rect(slide, Inches(0.8), Inches(1.8), Inches(5.8), Inches(2.5), LIGHT_BG)
add_text(slide, Inches(1.2), Inches(2.6), Inches(5), Inches(0.5), "Extrusion Coating Process", 16, True, GREEN_DARK, PP_ALIGN.CENTER, 'Space Grotesk')
add_text(slide, Inches(1.5), Inches(3.2), Inches(4.4), Inches(0.4), "[Insert PLA process diagram here]", 10, False, RGBColor(0xBB, 0xBB, 0xBB), PP_ALIGN.CENTER)

# PLA pellets image placeholder
pellets_box = add_rounded_rect(slide, Inches(6.8), Inches(1.8), Inches(5.8), Inches(2.5), LIGHT_BG)
add_text(slide, Inches(7.2), Inches(2.6), Inches(5), Inches(0.5), "PLA Pellets", 16, True, GREEN_DARK, PP_ALIGN.CENTER, 'Space Grotesk')
add_text(slide, Inches(7.5), Inches(3.2), Inches(4.4), Inches(0.4), "[Insert PLA pellets image here]", 10, False, RGBColor(0xBB, 0xBB, 0xBB), PP_ALIGN.CENTER)

# Process steps
steps = [("01", "HEAT", "Industrial dryer\nsoftens PLA pellets"), ("02", "APPLY", "Manual transfer\nonto bagasse"), ("03", "PRESS", "Hot plate applies\npressure & heat"), ("04", "COOL", "Protective layer\nforms")]
x = 1.2
for num, title, desc in steps:
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(4.6), Inches(1.2), Inches(1.2))
    circle.fill.solid()
    circle.fill.fore_color.rgb = GREEN_DARK
    circle.line.fill.background()
    add_text(slide, Inches(x), Inches(4.7), Inches(1.2), Inches(0.5), num, 24, True, WHITE, PP_ALIGN.CENTER, 'Space Grotesk')
    add_text(slide, Inches(x), Inches(5.1), Inches(1.2), Inches(0.3), title, 10, True, WHITE, PP_ALIGN.CENTER)
    add_text(slide, Inches(x - 0.3), Inches(5.9), Inches(1.8), Inches(0.6), desc, 9, False, GRAY, PP_ALIGN.CENTER)
    if num != "04":
        add_text(slide, Inches(x + 1.3), Inches(5), Inches(0.5), Inches(0.5), "→", 24, False, GREEN_MID, PP_ALIGN.CENTER)
    x += 2.8

# Info boxes
why_box = add_rounded_rect(slide, Inches(0.8), Inches(6.6), Inches(5.8), Inches(0.7), LIGHT_BG)
add_text(slide, Inches(1.2), Inches(6.7), Inches(5), Inches(0.5), "WHY PLA? → Thermoplastic nature enables heat-based application", 11, False, GREEN_DARK, PP_ALIGN.LEFT)

setup_box = add_rounded_rect(slide, Inches(6.8), Inches(6.6), Inches(5.8), Inches(0.7), GREEN_MID)
add_text(slide, Inches(7.2), Inches(6.7), Inches(5), Inches(0.5), "IMPROVISED SETUP → Hot plate with manual pressure", 11, False, WHITE, PP_ALIGN.LEFT)


# ==================== SLIDE 7: CHALLENGES ====================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_gradient_bg(slide, LIGHT_BG, WHITE)

# Header
header = add_shape(slide, Inches(0.5), Inches(0.5), Inches(12.3), Inches(1), GREEN_DARK)
add_text(slide, Inches(0.8), Inches(0.6), Inches(11.7), Inches(0.8), "PLA: CHALLENGES & LEARNINGS", 32, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')

# Challenge cards
challenges = [
    ("UNEVEN SPREADING", "Limited temperature uniformity resulted in incomplete contact between PLA and bagasse surface.", "Temperature Control"),
    ("THERMAL MISMATCH", "Difference in thermal behaviour between polymer and substrate caused coating inconsistency.", "Material Compatibility"),
    ("KEY INSIGHT", "Controlled hot-pressing would significantly improve results with proper process control.", "Equipment Needs"),
]
x = 0.8
for title, desc, tag in challenges:
    card = add_rounded_rect(slide, Inches(x), Inches(1.8), Inches(3.8), Inches(3), WHITE)
    card_head = add_shape(slide, Inches(x), Inches(1.8), Inches(3.8), Inches(0.7), GREEN_DARK)
    add_text(slide, Inches(x + 0.3), Inches(1.9), Inches(3.2), Inches(0.5), title, 14, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')
    add_text(slide, Inches(x + 0.3), Inches(2.7), Inches(3.2), Inches(1.5), desc, 12, False, GRAY, PP_ALIGN.LEFT)
    tag_box = add_rounded_rect(slide, Inches(x + 0.3), Inches(4.3), Inches(2), Inches(0.4), GREEN_LIGHT)
    add_text(slide, Inches(x + 0.4), Inches(4.35), Inches(1.8), Inches(0.3), tag, 9, True, WHITE, PP_ALIGN.CENTER)
    x += 4.1

# Takeaway banner
banner = add_shape(slide, Inches(0.8), Inches(5.2), Inches(11.7), Inches(2), GREEN_DARK)
add_text(slide, Inches(1.2), Inches(5.4), Inches(11), Inches(0.5), "WHAT WE LEARNED", 20, True, WHITE, PP_ALIGN.CENTER, 'Space Grotesk')
add_text(slide, Inches(1.5), Inches(5.9), Inches(10.4), Inches(1), "The experiment demonstrated limitations without appropriate equipment, giving practical insight into thermoplastic polymer processing requirements.", 14, False, RGBColor(0xCC, 0xFF, 0xDD), PP_ALIGN.CENTER)


# ==================== SLIDE 8: SEAWEED ====================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_gradient_bg(slide, WHITE, LIGHT_BG)

# Header
header = add_shape(slide, Inches(0.5), Inches(0.5), Inches(12.3), Inches(1), GREEN_MID)
add_text(slide, Inches(0.8), Inches(0.6), Inches(11.7), Inches(0.8), "SEAWEED COATING", 32, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')

# Diagram placeholder
diag_box = add_rounded_rect(slide, Inches(0.8), Inches(1.8), Inches(5.8), Inches(2.5), LIGHT_BG)
add_text(slide, Inches(1.2), Inches(2.6), Inches(5), Inches(0.5), "Seaweed Coating Process", 16, True, GREEN_DARK, PP_ALIGN.CENTER, 'Space Grotesk')
add_text(slide, Inches(1.5), Inches(3.2), Inches(4.4), Inches(0.4), "[Insert seaweed process diagram here]", 10, False, RGBColor(0xBB, 0xBB, 0xBB), PP_ALIGN.CENTER)

# Seaweed types placeholder
types_box = add_rounded_rect(slide, Inches(6.8), Inches(1.8), Inches(5.8), Inches(2.5), LIGHT_BG)
add_text(slide, Inches(7.2), Inches(2.6), Inches(5), Inches(0.5), "Seaweed Types", 16, True, GREEN_DARK, PP_ALIGN.CENTER, 'Space Grotesk')
add_text(slide, Inches(7.5), Inches(3.2), Inches(4.4), Inches(0.4), "[Insert seaweed types image here]", 10, False, RGBColor(0xBB, 0xBB, 0xBB), PP_ALIGN.CENTER)

# Steps
steps = [("01", "DISSOLVE", "Seaweed pellets\ndissolved in water"), ("02", "DIP", "Bagasse plates\nimmersed"), ("03", "DRY", "Ambient laboratory\nconditions"), ("04", "TEST", "Surface coverage\nassessment")]
x = 1.2
for num, title, desc in steps:
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(4.6), Inches(1.2), Inches(1.2))
    circle.fill.solid()
    circle.fill.fore_color.rgb = GREEN_MID
    circle.line.fill.background()
    add_text(slide, Inches(x), Inches(4.7), Inches(1.2), Inches(0.5), num, 24, True, WHITE, PP_ALIGN.CENTER, 'Space Grotesk')
    add_text(slide, Inches(x), Inches(5.1), Inches(1.2), Inches(0.3), title, 10, True, WHITE, PP_ALIGN.CENTER)
    add_text(slide, Inches(x - 0.3), Inches(5.9), Inches(1.8), Inches(0.6), desc, 9, False, GRAY, PP_ALIGN.CENTER)
    if num != "04":
        add_text(slide, Inches(x + 1.3), Inches(5), Inches(0.5), Inches(0.5), "→", 24, False, GREEN_MID, PP_ALIGN.CENTER)
    x += 2.8

# Info boxes
why_box = add_rounded_rect(slide, Inches(0.8), Inches(6.6), Inches(5.8), Inches(0.7), GREEN_MID)
add_text(slide, Inches(1.2), Inches(6.7), Inches(5), Inches(0.5), "WHY SEAWEED? → Naturally abundant, edible, antimicrobial", 11, False, WHITE, PP_ALIGN.LEFT)

dry_box = add_rounded_rect(slide, Inches(6.8), Inches(6.6), Inches(5.8), Inches(0.7), LIGHT_BG)
add_text(slide, Inches(7.2), Inches(6.7), Inches(5), Inches(0.5), "DRYING → Ambient conditions, no controlled equipment", 11, False, GREEN_DARK, PP_ALIGN.LEFT)


# ==================== SLIDE 9: SEAWEED CHALLENGES ====================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_gradient_bg(slide, WHITE, LIGHT_BG)

# Header
header = add_shape(slide, Inches(0.5), Inches(0.5), Inches(12.3), Inches(1), GREEN_MID)
add_text(slide, Inches(0.8), Inches(0.6), Inches(11.7), Inches(0.8), "SEAWEED: CHALLENGES & LEARNINGS", 32, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')

# Main challenge
main_box = add_rounded_rect(slide, Inches(0.8), Inches(1.8), Inches(11.7), Inches(1.5), LIGHT_BG)
add_text(slide, Inches(1.2), Inches(1.9), Inches(11), Inches(0.5), "POOR ADHESION & UNIFORMITY", 22, True, GREEN_DARK, PP_ALIGN.CENTER, 'Space Grotesk')
add_text(slide, Inches(1.5), Inches(2.4), Inches(10.4), Inches(0.8), "The bagasse plates' rough, porous surface absorbed the coating solution, preventing continuous film formation and causing poor mechanical durability.", 14, False, GRAY, PP_ALIGN.CENTER)

# Split cards
# Factors
factors_card = add_rounded_rect(slide, Inches(0.8), Inches(3.6), Inches(5.8), Inches(3.5), GREEN_MID)
add_text(slide, Inches(1.2), Inches(3.8), Inches(5), Inches(0.5), "CONTRIBUTING FACTORS", 18, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')

factors = ["No surface pre-treatment", "Uncontrolled coating conditions", "Solution concentration not optimized", "Viscosity and immersion time issues"]
txBox = slide.shapes.add_textbox(Inches(1.2), Inches(4.4), Inches(5), Inches(2.5))
tf = txBox.text_frame
tf.word_wrap = True
for i, feat in enumerate(factors):
    if i == 0:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    p.text = f"•  {feat}"
    p.font.size = Pt(13)
    p.font.color.rgb = WHITE
    p.font.name = 'Calibri'
    p.space_before = Pt(12)

# Takeaway
take_card = add_rounded_rect(slide, Inches(6.8), Inches(3.6), Inches(5.8), Inches(3.5), GREEN_DARK)
add_text(slide, Inches(7.2), Inches(3.8), Inches(5), Inches(0.5), "KEY TAKEAWAY", 18, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')

takeaways = ["Depositing material alone is insufficient", "Interfacial adhesion is critical", "Film formation requires control", "Surface pre-treatment necessary"]
txBox = slide.shapes.add_textbox(Inches(7.2), Inches(4.4), Inches(5), Inches(2.5))
tf = txBox.text_frame
tf.word_wrap = True
for i, feat in enumerate(takeaways):
    if i == 0:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    p.text = f"•  {feat}"
    p.font.size = Pt(13)
    p.font.color.rgb = WHITE
    p.font.name = 'Calibri'
    p.space_before = Pt(12)


# ==================== SLIDE 10: TESTING ====================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_gradient_bg(slide, LIGHT_BG, WHITE)

# Header
header = add_shape(slide, Inches(0.5), Inches(0.5), Inches(12.3), Inches(1), GREEN_DARK)
add_text(slide, Inches(0.8), Inches(0.6), Inches(11.7), Inches(0.8), "TESTING & VALIDATION", 32, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')

# Photo placeholders
photos = [
    ("H12 Canteen Testing", "[Insert canteen test photo]"),
    ("Lab Production Run", "[Insert lab plates photo]"),
    ("SINE Incubator", "[Insert SINE logo]"),
]
x = 0.8
for title, placeholder in photos:
    photo_box = add_rounded_rect(slide, Inches(x), Inches(1.8), Inches(3.8), Inches(2.2), WHITE)
    photo_inner = add_rounded_rect(slide, Inches(x + 0.2), Inches(2), Inches(3.4), Inches(1.6), LIGHT_BG)
    add_text(slide, Inches(x + 0.5), Inches(2.5), Inches(2.8), Inches(0.5), placeholder, 10, False, RGBColor(0xBB, 0xBB, 0xBB), PP_ALIGN.CENTER)
    add_text(slide, Inches(x + 0.2), Inches(3.7), Inches(3.4), Inches(0.4), title, 12, True, GREEN_DARK, PP_ALIGN.CENTER)
    x += 4.1

# Metrics
metrics_box = add_shape(slide, Inches(0.8), Inches(4.3), Inches(11.7), Inches(2.8), GREEN_DARK)

metrics = [("100%", "Plastic\nReduction"), ("30-90", "Days to\nDecompose"), ("1.3-1.5x", "Cost\nMultiplier"), ("150°C", "Heat\nResistance")]
x = 1.5
for num, lbl in metrics:
    add_text(slide, Inches(x), Inches(4.6), Inches(2.5), Inches(0.8), num, 42, True, GREEN_LIGHT, PP_ALIGN.CENTER, 'Space Grotesk')
    add_text(slide, Inches(x), Inches(5.4), Inches(2.5), Inches(0.8), lbl, 12, False, RGBColor(0xCC, 0xFF, 0xDD), PP_ALIGN.CENTER)
    x += 2.8


# ==================== SLIDE 11: ROADMAP ====================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_gradient_bg(slide, WHITE, LIGHT_BG)

# Header
header = add_shape(slide, Inches(0.5), Inches(0.5), Inches(12.3), Inches(1), GREEN_DARK)
add_text(slide, Inches(0.8), Inches(0.6), Inches(11.7), Inches(0.8), "FUTURE ROADMAP", 32, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')

# Timeline line
line = add_shape(slide, Inches(2), Inches(2.5), Inches(9.3), Inches(0.05), GREEN_LIGHT)

# Timeline items
timeline = [
    ("PHASE 1", "NOW", "Lab Testing", "Prototyping & formulation"),
    ("PHASE 2", "NEXT", "Pilot Production", "Local vendor partnerships"),
    ("PHASE 3", "6 MO", "Commercial Launch", "Restaurant partnerships"),
    ("PHASE 4", "1 YR", "Scale", "Tier 1 & 2 markets"),
]
colors = [GREEN_LIGHT, GREEN_MID, GREEN_DARK, RGBColor(0x00, 0x33, 0x20)]
x = 2.2
for i, (phase, time, title, desc) in enumerate(timeline):
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(2.1), Inches(1), Inches(1))
    circle.fill.solid()
    circle.fill.fore_color.rgb = colors[i]
    circle.line.fill.background()
    add_text(slide, Inches(x), Inches(2.15), Inches(1), Inches(0.3), phase, 7, True, WHITE, PP_ALIGN.CENTER)
    add_text(slide, Inches(x), Inches(2.4), Inches(1), Inches(0.3), time, 8, False, WHITE, PP_ALIGN.CENTER)
    add_text(slide, Inches(x - 0.5), Inches(3.3), Inches(2), Inches(0.4), title, 12, True, GREEN_DARK, PP_ALIGN.CENTER)
    add_text(slide, Inches(x - 0.5), Inches(3.7), Inches(2), Inches(0.4), desc, 10, False, GRAY, PP_ALIGN.CENTER)
    x += 2.5

# Goals
goals = [
    ("SCALABILITY", "Develop scalable manufacturing for mass production"),
    ("PARTNERSHIPS", "Collaborate with McDonald's, Starbucks & local chains"),
    ("IMPACT", "Reduce 1000+ tons of plastic waste annually by Year 3"),
]
x = 0.8
for title, desc in goals:
    goal_card = add_rounded_rect(slide, Inches(x), Inches(4.5), Inches(3.8), Inches(2.5), WHITE)
    goal_head = add_shape(slide, Inches(x), Inches(4.5), Inches(3.8), Inches(0.6), GREEN_DARK)
    add_text(slide, Inches(x + 0.3), Inches(4.55), Inches(3.2), Inches(0.5), title, 13, True, WHITE, PP_ALIGN.LEFT, 'Space Grotesk')
    add_text(slide, Inches(x + 0.3), Inches(5.3), Inches(3.2), Inches(1.5), desc, 12, False, GRAY, PP_ALIGN.LEFT)
    x += 4.1


# ==================== SLIDE 12: THANK YOU ====================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_gradient_bg(slide, RGBColor(0x00, 0x1A, 0x0D), RGBColor(0x00, 0x33, 0x20))

# Decorative circles
for pos in [(10, -1), (-1, 5)]:
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(pos[0]), Inches(pos[1]), Inches(4), Inches(4))
    circle.fill.solid()
    circle.fill.fore_color.rgb = RGBColor(0x00, 0x4D, 0x2C)
    circle.line.fill.background()

# Thank you text
add_text(slide, Inches(2), Inches(2.5), Inches(9.3), Inches(1.5), "THANK YOU", 64, True, WHITE, PP_ALIGN.CENTER, 'Space Grotesk')
add_text(slide, Inches(2), Inches(4), Inches(9.3), Inches(0.5), "PIONEERING A FUTURE WHERE PACKAGING HEALS THE PLANET", 16, False, GREEN_LIGHT, PP_ALIGN.CENTER)
add_text(slide, Inches(2), Inches(5), Inches(9.3), Inches(0.5), "Enactus IIT Bombay  •  SusPak", 14, False, RGBColor(0x99, 0x99, 0x99), PP_ALIGN.CENTER)

# Bottom bar
add_shape(slide, Inches(0), Inches(7), Inches(13.333), Inches(0.5), GREEN_MID)
add_text(slide, Inches(2), Inches(7.05), Inches(9.3), Inches(0.4), "Sustainable Packaging for a Better Tomorrow", 12, False, WHITE, PP_ALIGN.CENTER)


# Save
output = r"C:\Users\Tarun Kumar\Downloads\SusPak-Presentation\SusPak_Premium.pptx"
prs.save(output)
print(f"Saved: {output}")
print(f"Total slides: {len(prs.slides)}")
