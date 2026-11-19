from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
import os
import tempfile


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

WHITE = RGBColor(255, 255, 255)
LIGHT = RGBColor(247, 249, 252)
NAVY = RGBColor(29, 53, 87)
BLUE = RGBColor(70, 130, 180)
TEAL = RGBColor(72, 176, 169)
GOLD = RGBColor(245, 166, 35)
TEXT = RGBColor(52, 58, 64)
MUTED = RGBColor(108, 117, 125)
PALE_BLUE = RGBColor(231, 241, 248)
PALE_TEAL = RGBColor(232, 246, 244)
PALE_GOLD = RGBColor(255, 244, 221)
PALE_RED = RGBColor(251, 235, 235)


def set_bg(slide, color=LIGHT):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def add_text(slide, text, x, y, w, h, size=20, color=TEXT,
             bold=False, align=PP_ALIGN.LEFT, font="Aptos",
             valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.vertical_anchor = valign
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    r = p.runs[0]
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return box


def add_title(slide, title, subtitle=None, number=None):
    add_text(slide, title, 0.7, 0.45, 11.7, 0.55, 28, NAVY, True)
    slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(1.12),
                           Inches(0.65), Inches(0.06)).fill.solid()
    accent = slide.shapes[-1]
    accent.fill.fore_color.rgb = TEAL
    accent.line.fill.background()
    if subtitle:
        add_text(slide, subtitle, 1.55, 1.02, 10.8, 0.32, 12, MUTED)
    if number:
        add_text(slide, f"{number:02}", 12.1, 0.55, 0.55, 0.35, 11, MUTED, True,
                 PP_ALIGN.RIGHT)


def add_footer(slide, text="PERSONALITY DEVELOPMENT"):
    add_text(slide, text, 0.7, 7.12, 4.0, 0.18, 8, MUTED, True)


def add_card(slide, x, y, w, h, fill=WHITE, line=None, radius=True):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    card = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    card.fill.solid()
    card.fill.fore_color.rgb = fill
    card.line.color.rgb = line if line else fill
    card.line.width = Pt(0.8)
    return card


def add_bullet_list(slide, items, x, y, w, h, size=18, color=TEXT,
                    bullet_color=TEAL, spacing=8):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = Inches(0.06)
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.space_after = Pt(spacing)
        p.font.name = "Aptos"
        p.font.size = Pt(size)
        p.font.color.rgb = color
    return box


def add_circle(slide, x, y, d, fill, text, text_color=WHITE, size=20):
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y),
                               Inches(d), Inches(d))
    c.fill.solid()
    c.fill.fore_color.rgb = fill
    c.line.fill.background()
    add_text(slide, text, x, y, d, d, size, text_color, True,
             PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    return c


# 1 Cover
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_card(slide, 8.95, 0.65, 3.65, 5.95, PALE_BLUE, PALE_BLUE)
add_circle(slide, 9.55, 1.35, 2.45, BLUE, "GROW", WHITE, 24)
add_circle(slide, 10.45, 3.65, 1.25, TEAL, "ACT", WHITE, 16)
add_circle(slide, 9.15, 4.45, 0.95, GOLD, "LEARN", NAVY, 12)
add_text(slide, "PERSONALITY\nDEVELOPMENT", 0.8, 1.35, 7.7, 1.55,
         38, NAVY, True)
add_text(slide, "BUILD YOUR CONFIDENT SELF", 0.82, 3.15, 6.6, 0.45,
         16, TEAL, True)
add_text(slide, "A practical guide for students", 0.82, 3.75, 5.8, 0.35,
         20, TEXT)
add_text(slide, "Presentation  |  Student Edition", 0.82, 6.55, 5.5, 0.28,
         11, MUTED)

# 2 Agenda
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_title(slide, "Today’s journey", "A clear path from self-awareness to consistent action", 2)
items = [
    ("01", "Understand", "What personality means"),
    ("02", "Discover", "Your core strengths"),
    ("03", "Develop", "Communication and confidence"),
    ("04", "Practice", "Habits that create impact"),
    ("05", "Plan", "Your 30-day growth roadmap"),
]
for i, (num, heading, desc) in enumerate(items):
    y = 1.55 + i * 0.98
    add_card(slide, 0.75, y, 11.8, 0.78, WHITE)
    add_circle(slide, 1.05, y + 0.14, 0.5, TEAL, num, WHITE, 10)
    add_text(slide, heading, 1.85, y + 0.18, 2.25, 0.34, 17, NAVY, True)
    add_text(slide, desc, 4.15, y + 0.18, 7.6, 0.34, 16, TEXT)
add_footer(slide)

# 3 What is personality
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_title(slide, "What is personality?", "The total pattern of how you think, feel and act", 3)
add_card(slide, 0.75, 1.55, 3.55, 4.75, PALE_BLUE)
add_text(slide, "IT IS", 1.05, 1.92, 1.4, 0.32, 12, BLUE, True)
add_bullet_list(slide, [
    "Distinctive",
    "Observable",
    "Developable",
    "Connected to relationships",
], 1.05, 2.45, 2.75, 3.0, 19, NAVY)
add_card(slide, 4.65, 1.55, 3.55, 4.75, PALE_TEAL)
add_text(slide, "IT IS NOT", 4.95, 1.92, 1.8, 0.32, 12, TEAL, True)
add_bullet_list(slide, [
    "Just appearance",
    "Fixed at birth",
    "Only extroversion",
    "A single perfect type",
], 4.95, 2.45, 2.75, 3.0, 19, NAVY)
add_card(slide, 8.55, 1.55, 4.05, 4.75, PALE_GOLD)
add_text(slide, "KEY IDEA", 8.85, 1.92, 1.8, 0.32, 12, GOLD, True)
add_text(slide, "Personality is revealed\nby repeated behavior.", 8.85, 2.55, 3.1, 1.15,
         24, NAVY, True)
add_text(slide, "Small actions, repeated consistently,\nshape how others experience you.", 8.85, 4.15, 3.1, 1.1,
         15, TEXT)
add_footer(slide)

# 4 Core dimensions
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_title(slide, "The five dimensions of a strong profile", "Balance matters more than perfection", 4)
dimensions = [
    ("01", "Self-awareness", "Know your patterns", PALE_BLUE, BLUE),
    ("02", "Communication", "Express with clarity", PALE_TEAL, TEAL),
    ("03", "Confidence", "Act with purpose", PALE_GOLD, GOLD),
    ("04", "Emotional control", "Respond, don’t react", PALE_RED, RGBColor(202, 103, 103)),
    ("05", "Positive character", "Build trust through values", RGBColor(238, 233, 248), RGBColor(125, 99, 172)),
]
for i, (num, head, desc, fill, accent) in enumerate(dimensions):
    x = 0.75 + i * 2.43
    add_card(slide, x, 1.55, 2.15, 4.65, fill)
    add_circle(slide, x + 0.68, 1.95, 0.78, accent, num, WHITE, 13)
    add_text(slide, head, x + 0.25, 3.05, 1.65, 1.05, 17, NAVY, True,
             PP_ALIGN.CENTER)
    add_text(slide, desc, x + 0.22, 4.45, 1.71, 0.8, 13, TEXT, False,
             PP_ALIGN.CENTER)
add_footer(slide)

# 5 Self-awareness
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_title(slide, "Start with self-awareness", "Use a simple reflection loop", 5)
steps = [
    ("Notice", "What happened?", BLUE),
    ("Name", "What did I feel?", TEAL),
    ("Learn", "What pattern do I see?", GOLD),
    ("Choose", "What will I do next?", RGBColor(125, 99, 172)),
]
for i, (head, desc, accent) in enumerate(steps):
    x = 0.75 + i * 3.05
    add_card(slide, x, 1.75, 2.7, 3.45, WHITE)
    add_circle(slide, x + 0.95, 2.08, 0.8, accent, str(i + 1), WHITE, 16)
    add_text(slide, head, x + 0.3, 3.18, 2.1, 0.4, 20, NAVY, True,
             PP_ALIGN.CENTER)
    add_text(slide, desc, x + 0.25, 3.78, 2.2, 0.65, 14, MUTED, False,
             PP_ALIGN.CENTER)
    if i < 3:
        add_text(slide, "→", x + 2.72, 3.05, 0.5, 0.5, 24, MUTED, True,
                 PP_ALIGN.CENTER)
add_card(slide, 0.75, 5.55, 11.8, 0.75, PALE_BLUE)
add_text(slide, "Reflection prompt:", 1.05, 5.78, 1.75, 0.28, 13, BLUE, True)
add_text(slide, "“Where did I show courage today, and where could I respond better?”",
         2.85, 5.75, 8.95, 0.32, 16, NAVY)
add_footer(slide)

# 6 Communication
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_title(slide, "Communication that connects", "Clear is kind. Listening is powerful.", 6)
add_card(slide, 0.75, 1.55, 7.15, 4.75, WHITE)
add_text(slide, "THE 4C FRAMEWORK", 1.05, 1.92, 3.0, 0.32, 12, BLUE, True)
framework = [
    ("Clear", "Use simple, direct language", BLUE),
    ("Calm", "Control pace and tone", TEAL),
    ("Curious", "Ask before assuming", GOLD),
    ("Considerate", "Respect context and feelings", RGBColor(125, 99, 172)),
]
for i, (head, desc, accent) in enumerate(framework):
    y = 2.55 + i * 0.82
    add_circle(slide, 1.05, y, 0.55, accent, str(i + 1), WHITE, 12)
    add_text(slide, head, 1.82, y + 0.02, 1.45, 0.3, 16, NAVY, True)
    add_text(slide, desc, 3.25, y + 0.03, 4.15, 0.3, 14, TEXT)
add_card(slide, 8.25, 1.55, 4.35, 4.75, PALE_TEAL)
add_text(slide, "ACTIVE LISTENING", 8.55, 1.92, 3.1, 0.32, 12, TEAL, True)
add_bullet_list(slide, [
    "Maintain attentive eye contact",
    "Listen to understand",
    "Summarize what you heard",
    "Ask a thoughtful question",
], 8.55, 2.55, 3.45, 2.8, 16, NAVY, spacing=12)
add_footer(slide)

# 7 Confidence
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_title(slide, "Confidence comes from evidence", "Prepare small wins. Repeat what works.", 7)
add_card(slide, 0.75, 1.55, 11.8, 1.05, PALE_BLUE)
add_text(slide, "CONFIDENCE", 1.05, 1.82, 2.0, 0.28, 13, BLUE, True)
add_text(slide, "Preparation  +  Practice  +  Positive self-talk",
         3.05, 1.78, 8.9, 0.36, 22, NAVY, True)
add_card(slide, 0.75, 2.95, 3.65, 3.35, WHITE)
add_text(slide, "BEFORE", 1.05, 3.28, 1.4, 0.28, 12, MUTED, True)
add_bullet_list(slide, ["Know the goal", "Prepare key points", "Visualize success"],
                1.05, 3.82, 2.85, 1.8, 15)
add_card(slide, 4.85, 2.95, 3.65, 3.35, WHITE)
add_text(slide, "DURING", 5.15, 3.28, 1.4, 0.28, 12, MUTED, True)
add_bullet_list(slide, ["Slow your breathing", "Use a steady voice", "Focus on value"],
                5.15, 3.82, 2.85, 1.8, 15)
add_card(slide, 8.95, 2.95, 3.6, 3.35, PALE_GOLD)
add_text(slide, "AFTER", 9.25, 3.28, 1.4, 0.28, 12, GOLD, True)
add_bullet_list(slide, ["Review honestly", "Celebrate progress", "Choose one improvement"],
                9.25, 3.82, 2.75, 1.8, 15)
add_footer(slide)

# 8 Emotional intelligence
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_title(slide, "Build emotional intelligence", "Feel it. Name it. Choose your response.", 8)
add_card(slide, 0.75, 1.55, 5.25, 4.75, PALE_BLUE)
add_text(slide, "PAUSE", 1.05, 1.92, 1.5, 0.32, 12, BLUE, True)
add_circle(slide, 1.05, 2.55, 3.45, WHITE, "Breathe\nCount to 5", NAVY, 22)
add_card(slide, 6.35, 1.55, 6.2, 4.75, WHITE)
add_text(slide, "RESPOND WITH INTENTION", 6.65, 1.92, 3.7, 0.32, 12, TEAL, True)
steps = [
    ("1", "Identify the emotion"),
    ("2", "Separate fact from assumption"),
    ("3", "Choose a respectful action"),
]
for i, (num, text) in enumerate(steps):
    y = 2.55 + i * 0.95
    add_circle(slide, 6.65, y, 0.58, TEAL if i != 1 else GOLD, num, WHITE if i != 1 else NAVY, 12)
    add_text(slide, text, 7.48, y + 0.1, 4.45, 0.35, 17, NAVY, True)
add_text(slide, "Try: “I felt ___ because ___. Next time, I will ___.”",
         6.65, 5.45, 5.35, 0.45, 13, MUTED)
add_footer(slide)

# 9 Habits
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_title(slide, "Daily habits shape personality", "Make growth visible through routines", 9)
habits = [
    ("MORNING", "Set intention", "1 clear priority", BLUE),
    ("CLASS", "Participate", "Ask or contribute", TEAL),
    ("LEARNING", "Read or listen", "Capture one insight", GOLD),
    ("SOCIAL", "Use 4Cs", "Listen and respond", RGBColor(125, 99, 172)),
    ("EVENING", "Reflect", "Note progress", RGBColor(202, 103, 103)),
]
for i, (phase, action, detail, accent) in enumerate(habits):
    x = 0.75 + i * 2.43
    add_card(slide, x, 1.55, 2.15, 4.75, WHITE)
    add_text(slide, phase, x + 0.25, 1.95, 1.65, 0.28, 11, accent, True,
             PP_ALIGN.CENTER)
    add_circle(slide, x + 0.68, 2.55, 0.78, accent, str(i + 1), WHITE, 15)
    add_text(slide, action, x + 0.25, 3.65, 1.65, 0.55, 17, NAVY, True,
             PP_ALIGN.CENTER)
    add_text(slide, detail, x + 0.22, 4.55, 1.71, 0.55, 13, MUTED, False,
             PP_ALIGN.CENTER)
add_footer(slide)

# 10 30-day plan
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_title(slide, "Your 30-day growth plan", "Four weeks. One meaningful upgrade.", 10)
weeks = [
    ("WEEK 1", "Discover", "Complete a strengths reflection", PALE_BLUE, BLUE),
    ("WEEK 2", "Communicate", "Practice one 4C skill daily", PALE_TEAL, TEAL),
    ("WEEK 3", "Contribute", "Lead or support one class activity", PALE_GOLD, GOLD),
    ("WEEK 4", "Reflect", "Compare results and set next goals", RGBColor(238, 233, 248), RGBColor(125, 99, 172)),
]
for i, (week, head, desc, fill, accent) in enumerate(weeks):
    x = 0.75 + i * 3.05
    add_card(slide, x, 1.55, 2.7, 4.25, fill)
    add_text(slide, week, x + 0.28, 1.92, 2.14, 0.28, 11, accent, True)
    add_text(slide, head, x + 0.28, 2.55, 2.14, 0.52, 21, NAVY, True)
    add_text(slide, desc, x + 0.28, 3.35, 2.12, 1.25, 15, TEXT)
    add_text(slide, f"{(i + 1) * 7} days", x + 0.28, 5.15, 2.14, 0.28, 11, MUTED, True)
add_card(slide, 0.75, 6.15, 11.8, 0.55, NAVY)
add_text(slide, "Measure progress: confidence, consistency, feedback and self-awareness.",
         1.05, 6.29, 11.2, 0.25, 14, WHITE, True, PP_ALIGN.CENTER)
add_footer(slide)

# 11 Closing
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_card(slide, 0.75, 0.75, 11.8, 5.9, PALE_BLUE)
add_text(slide, "Become your\nbest possible self.", 1.25, 1.55, 7.2, 1.55,
         36, NAVY, True)
add_text(slide, "Know yourself. Express clearly. Act intentionally.", 1.25, 3.38,
         8.2, 0.42, 19, TEAL, True)
add_text(slide, "Your personality is not a fixed label—", 1.25, 4.35, 6.8, 0.35,
         16, TEXT)
add_text(slide, "it is a skill you can strengthen every day.", 1.25, 4.82, 7.6, 0.42,
         20, NAVY, True)
add_circle(slide, 9.75, 2.05, 1.8, GOLD, "START", NAVY, 18)
add_text(slide, "THANK YOU", 1.25, 6.35, 3.0, 0.28, 11, MUTED, True)

output = "Personality_Development_for_Students.pptx"
prs.save(output)


def add_slide_animations(pptx_path):
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
    ns = {"p": P_NS}
    fd, temp_path = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    with ZipFile(pptx_path, "r") as source, ZipFile(temp_path, "w", ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename.startswith("ppt/slides/slide") and item.filename.endswith(".xml"):
                root = etree.fromstring(data)
                for old in root.xpath("./p:transition", namespaces=ns):
                    root.remove(old)
                transition = etree.Element("{%s}transition" % P_NS)
                transition.set("spd", "med")
                transition.set("advClick", "1")
                fade = etree.SubElement(transition, "{%s}fade" % P_NS)
                root.insert(1, transition)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                      standalone=True)
            target.writestr(item, data)
    os.replace(temp_path, pptx_path)


add_slide_animations(output)
print(output)
