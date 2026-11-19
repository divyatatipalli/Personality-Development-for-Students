from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
import os


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

NAVY = RGBColor(31, 49, 76)
BLUE = RGBColor(55, 118, 171)
TEAL = RGBColor(58, 148, 140)
GOLD = RGBColor(223, 158, 55)
TEXT = RGBColor(63, 70, 78)
MUTED = RGBColor(112, 123, 132)
WHITE = RGBColor(255, 255, 255)
LIGHT = RGBColor(247, 249, 252)
PALE_BLUE = RGBColor(232, 240, 247)
PALE_TEAL = RGBColor(232, 244, 242)
PALE_GOLD = RGBColor(252, 243, 221)
PALE_RED = RGBColor(248, 237, 237)
PALE_PURPLE = RGBColor(240, 237, 247)


def set_bg(slide, color=LIGHT):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def textbox(slide, text, x, y, w, h, size=18, color=TEXT, bold=False,
            align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    r = p.runs[0]
    r.font.name = "Aptos"
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return box


def card(slide, x, y, w, h, fill=WHITE, line=WHITE):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line
    shape.line.width = Pt(0.8)
    return shape


def circle(slide, x, y, d, fill, text="", text_color=WHITE, size=14):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                   Inches(x), Inches(y), Inches(d), Inches(d))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    if text:
        textbox(slide, text, x, y, d, d, size, text_color, True,
                PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    return shape


def title(slide, heading, subtitle, slide_no):
    textbox(slide, heading, 0.72, 0.38, 11.25, 0.55, 27, NAVY, True)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.72),
                                  Inches(1.06), Inches(0.68), Inches(0.055))
    line.fill.solid()
    line.fill.fore_color.rgb = TEAL
    line.line.fill.background()
    textbox(slide, subtitle, 1.55, 0.96, 10.25, 0.28, 11, MUTED)
    textbox(slide, f"{slide_no:02}", 12.05, 0.48, 0.55, 0.28, 10, MUTED, True,
            PP_ALIGN.RIGHT)


def footer(slide):
    textbox(slide, "PERSONALITY DEVELOPMENT FOR STUDENTS", 0.72, 7.12,
            4.8, 0.16, 8, MUTED, True)


def bullets(slide, items, x, y, w, h, size=15, color=TEXT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = Inches(0.05)
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.space_after = Pt(10)
        p.font.name = "Aptos"
        p.font.size = Pt(size)
        p.font.color.rgb = color


# 1 — Cover
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s)
card(s, 8.75, 0.62, 3.85, 6.25, PALE_BLUE)
circle(s, 9.52, 1.42, 2.35, BLUE, "GROW", WHITE, 24)
circle(s, 10.72, 3.62, 1.18, TEAL, "ACT", WHITE, 15)
circle(s, 9.12, 4.55, 1.02, GOLD, "LEARN", NAVY, 11)
textbox(s, "PERSONALITY\nDEVELOPMENT", 0.82, 1.28, 7.55, 1.55, 38, NAVY, True)
textbox(s, "FROM SELF-AWARENESS TO PROFESSIONAL PRESENCE", 0.84, 3.12,
        7.35, 0.38, 15, TEAL, True)
textbox(s, "A practical learning program for confident communication,\nethical behavior and career readiness.",
        0.84, 3.78, 7.15, 0.9, 20, TEXT)
textbox(s, "STUDENT EDITION  |  WORKSHOP VERSION", 0.84, 6.58, 5.8, 0.28, 10, MUTED, True)

# 2 — Learning outcomes
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s)
title(s, "Learning outcomes", "By the end, students should be able to demonstrate—not merely understand", 2)
outcomes = [
    ("01", "Diagnose", "Recognize personal strengths, habits and triggers."),
    ("02", "Communicate", "Structure clear messages and listen with attention."),
    ("03", "Regulate", "Pause, reflect and respond professionally under pressure."),
    ("04", "Collaborate", "Build trust, manage conflict and contribute consistently."),
]
for i, (num, head, detail) in enumerate(outcomes):
    x = 0.72 + i * 3.05
    card(s, x, 1.58, 2.72, 4.6, WHITE)
    circle(s, x + 0.95, 1.95, 0.82, [BLUE, TEAL, GOLD, RGBColor(124, 101, 166)][i],
           num, WHITE, 13)
    textbox(s, head, x + 0.25, 3.08, 2.22, 0.44, 18, NAVY, True, PP_ALIGN.CENTER)
    textbox(s, detail, x + 0.32, 3.82, 2.08, 1.35, 12, TEXT, False, PP_ALIGN.CENTER)
footer(s)

# 3 — Why it matters
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s)
title(s, "Personality shapes professional outcomes", "Ability creates potential; behavior determines how potential is experienced", 3)
contexts = [
    ("ACADEMIC", "Class participation\nTeam projects\nPresentation quality", PALE_BLUE, BLUE),
    ("PERSONAL", "Relationships\nDecision quality\nStress response", PALE_TEAL, TEAL),
    ("CAREER", "Interview impact\nWorkplace trust\nLeadership readiness", PALE_GOLD, GOLD),
]
for i, (head, text, fill, accent) in enumerate(contexts):
    x = 0.72 + i * 4.05
    card(s, x, 1.58, 3.68, 4.72, fill)
    textbox(s, head, x + 0.32, 1.95, 3.04, 0.32, 12, accent, True)
    textbox(s, text, x + 0.32, 2.78, 3.04, 2.1, 21, NAVY, True)
    textbox(s, ["Competence earns attention.", "Character earns trust.", "Consistency earns opportunity"][i],
            x + 0.32, 5.3, 3.04, 0.55, 13, TEXT)
footer(s)

# 4 — Integrated model
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s)
title(s, "The professional personality model", "Five connected capabilities, each observable in everyday behavior", 4)
parts = [
    ("1", "SELF-AWARENESS", "Values, strengths, triggers", BLUE),
    ("2", "EMOTIONAL CONTROL", "Pause, name, respond", TEAL),
    ("3", "COMMUNICATION", "Clarity, listening, empathy", GOLD),
    ("4", "COLLABORATION", "Reliability, respect, teamwork", RGBColor(124, 101, 166)),
    ("5", "ADAPTABILITY", "Learning, feedback, resilience", RGBColor(196, 102, 102)),
]
for i, (num, head, detail, accent) in enumerate(parts):
    y = 1.55 + i * 0.99
    card(s, 0.72, y, 8.05, 0.78, WHITE)
    circle(s, 0.98, y + 0.14, 0.5, accent, num, WHITE, 10)
    textbox(s, head, 1.72, y + 0.17, 3.15, 0.28, 15, NAVY, True)
    textbox(s, detail, 4.92, y + 0.17, 3.45, 0.28, 13, TEXT)
card(s, 9.15, 1.55, 3.42, 4.95, PALE_BLUE)
textbox(s, "CORE PRINCIPLE", 9.48, 1.92, 2.75, 0.28, 11, BLUE, True)
textbox(s, "Personality is\nnot a fixed label.", 9.48, 2.62, 2.75, 1.1, 25, NAVY, True)
textbox(s, "It is a set of learnable\npatterns shaped through\nfeedback and repetition.",
        9.48, 4.18, 2.75, 1.42, 15, TEXT)
footer(s)

# 5 — Self awareness
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s)
title(s, "Build self-awareness through evidence", "Replace assumptions with observed patterns", 5)
card(s, 0.72, 1.55, 5.72, 4.82, WHITE)
textbox(s, "REFLECTION CANVAS", 1.05, 1.92, 3.2, 0.28, 11, BLUE, True)
canvas = [
    ("STRENGTH", "What consistently works?"),
    ("TRIGGER", "When do I lose focus or composure?"),
    ("PATTERN", "What repeat behavior created this result?"),
    ("CHOICE", "What can I practice differently?"),
]
for i, (head, prompt) in enumerate(canvas):
    y = 2.5 + i * 0.88
    textbox(s, head, 1.05, y, 1.35, 0.26, 11, TEAL, True)
    textbox(s, prompt, 2.48, y - 0.02, 3.55, 0.38, 15, NAVY)
card(s, 6.78, 1.55, 5.78, 4.82, PALE_GOLD)
textbox(s, "VALIDATE WITH FEEDBACK", 7.12, 1.92, 3.7, 0.28, 11, GOLD, True)
textbox(s, "Ask three people:", 7.12, 2.55, 2.4, 0.35, 17, NAVY, True)
bullets(s, [
    "“What do I contribute well?”",
    "“What is a recurring challenge?”",
    "“What one behavior should I change?”",
], 7.12, 3.15, 4.7, 2.2, 16)
footer(s)

# 6 — Strengths
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s)
title(s, "Turn strengths into measurable behaviors", "A strength creates value only when it is visible and used", 6)
rows = [
    ("Reliability", "Meet agreed deadlines", "On-time submissions"),
    ("Curiosity", "Ask purposeful questions", "Evidence-based decisions"),
    ("Initiative", "Act without waiting", "Completed improvement"),
    ("Empathy", "Recognize others’ needs", "Positive team climate"),
]
textbox(s, "STRENGTH", 1.05, 1.62, 2.15, 0.28, 11, MUTED, True)
textbox(s, "BEHAVIOR", 4.18, 1.62, 3.42, 0.28, 11, MUTED, True)
textbox(s, "RESULT", 8.62, 1.62, 3.42, 0.28, 11, MUTED, True)
for i, (head, behavior, result) in enumerate(rows):
    y = 2.12 + i * 0.95
    card(s, 0.72, y, 11.84, 0.72, WHITE)
    textbox(s, head, 1.05, y + 0.19, 2.35, 0.28, 16, NAVY, True)
    textbox(s, behavior, 4.18, y + 0.19, 3.55, 0.28, 14, TEXT)
    textbox(s, result, 8.62, y + 0.19, 3.45, 0.28, 14, TEAL, True)
footer(s)

# 7 — Communication
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s)
title(s, "Professional communication", "Clarity, calibration and respect—before, during and after", 7)
card(s, 0.72, 1.55, 7.18, 4.82, WHITE)
textbox(s, "THE CLEAR FRAMEWORK", 1.05, 1.92, 3.2, 0.28, 11, BLUE, True)
framework = [
    ("C", "Context", "Why this matters now"),
    ("L", "Logic", "What you observed or know"),
    ("E", "Expectation", "What outcome is needed"),
    ("A", "Action", "Your proposed next step"),
    ("R", "Respect", "Invite response and perspective"),
]
for i, (letter, head, detail) in enumerate(framework):
    y = 2.48 + i * 0.68
    circle(s, 1.05, y, 0.48, [BLUE, TEAL, GOLD, RGBColor(124,101,166), RGBColor(196,102,102)][i],
           letter, WHITE, 11)
    textbox(s, head, 1.72, y + 0.06, 1.45, 0.26, 14, NAVY, True)
    textbox(s, detail, 3.25, y + 0.06, 3.95, 0.26, 13, TEXT)
card(s, 8.25, 1.55, 4.31, 4.82, PALE_TEAL)
textbox(s, "ACTIVE LISTENING", 8.58, 1.92, 3.45, 0.28, 11, TEAL, True)
bullets(s, [
    "Attend fully before replying",
    "Paraphrase meaning",
    "Separate facts from interpretations",
    "Ask one clarifying question",
    "Confirm the next action",
], 8.58, 2.55, 3.45, 3.15, 16, NAVY)
footer(s)

# 8 — Emotional intelligence
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s)
title(s, "Respond, do not react", "A repeatable regulation process protects judgment and relationships", 8)
steps = [
    ("PAUSE", "Create a 3–5 second gap", BLUE),
    ("NAME", "Identify the emotion and trigger", TEAL),
    ("ASSESS", "Separate facts from stories", GOLD),
    ("CHOOSE", "Select the respectful next action", RGBColor(124,101,166)),
]
for i, (head, detail, accent) in enumerate(steps):
    x = 0.72 + i * 3.05
    card(s, x, 1.58, 2.72, 3.72, WHITE)
    circle(s, x + 0.95, 1.95, 0.82, accent, str(i + 1), WHITE, 14)
    textbox(s, head, x + 0.25, 3.08, 2.22, 0.38, 17, NAVY, True, PP_ALIGN.CENTER)
    textbox(s, detail, x + 0.28, 3.75, 2.16, 0.82, 14, TEXT, False, PP_ALIGN.CENTER)
card(s, 0.72, 5.68, 11.84, 0.72, PALE_RED)
textbox(s, "SCRIPT", 1.05, 5.93, 1.0, 0.25, 11, RGBColor(196,102,102), True)
textbox(s, "“I want to understand this clearly. Let me restate what I heard, then I will respond.”",
        2.15, 5.88, 9.75, 0.32, 14, NAVY)
footer(s)

# 9 — Confidence
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s)
title(s, "Confidence as preparation plus exposure", "Confidence grows after successful repetition—not before it", 9)
card(s, 0.72, 1.55, 11.84, 1.1, PALE_BLUE)
textbox(s, "CONFIDENCE =", 1.05, 1.91, 2.0, 0.28, 12, BLUE, True)
textbox(s, "Preparation  +  Small practice  +  Accurate feedback",
        3.05, 1.82, 8.65, 0.42, 22, NAVY, True)
phases = [
    ("PREPARE", "Define the goal\nKnow the audience\nPractice the opening", BLUE),
    ("PERFORM", "Use a steady pace\nState evidence\nPause for questions", TEAL),
    ("IMPROVE", "Request specific feedback\nReframe mistakes\nRepeat once", GOLD),
]
for i, (head, text, accent) in enumerate(phases):
    x = 0.72 + i * 4.05
    card(s, x, 3.02, 3.68, 3.28, WHITE)
    textbox(s, head, x + 0.32, 3.38, 3.04, 0.28, 12, accent, True)
    textbox(s, text, x + 0.32, 4.05, 3.04, 1.38, 16, NAVY)
footer(s)

# 10 — Conflict
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s)
title(s, "Resolve disagreement professionally", "Address the issue and preserve the relationship", 10)
card(s, 0.72, 1.55, 11.84, 4.82, WHITE)
steps = [
    ("1", "Define", "Describe the specific behavior or decision"),
    ("2", "Clarify", "Confirm information and intent"),
    ("3", "Explore", "Ask for perspectives and constraints"),
    ("4", "Align", "Agree on standards and next actions"),
    ("5", "Review", "Schedule a checkpoint"),
]
for i, (num, head, detail) in enumerate(steps):
    x = 0.98 + i * 2.32
    circle(s, x + 0.64, 2.05, 0.65, [BLUE, TEAL, GOLD, RGBColor(124,101,166), RGBColor(196,102,102)][i],
           num, WHITE, 12)
    textbox(s, head, x + 0.05, 3.02, 1.82, 0.34, 16, NAVY, True, PP_ALIGN.CENTER)
    textbox(s, detail, x - 0.02, 3.72, 1.96, 1.02, 13, TEXT, False, PP_ALIGN.CENTER)
    if i < 4:
        textbox(s, "→", x + 1.82, 2.15, 0.42, 0.3, 18, MUTED, True, PP_ALIGN.CENTER)
textbox(s, "Avoid blame language. Use: behavior + impact + request.",
        1.05, 5.55, 11.05, 0.38, 16, TEAL, True, PP_ALIGN.CENTER)
footer(s)

# 11 — Digital presence
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s)
title(s, "Manage your digital professional presence", "Your online conduct is part of your personality at scale", 11)
cards = [
    ("PROFILE", "Use a recognizable, consistent identity", PALE_BLUE, BLUE),
    ("CONTENT", "Share useful work; verify before reposting", PALE_TEAL, TEAL),
    ("INTERACTION", "Be constructive, concise and respectful", PALE_GOLD, GOLD),
    ("PRIVACY", "Review permissions and long-term impact", PALE_PURPLE, RGBColor(124,101,166)),
]
for i, (head, detail, fill, accent) in enumerate(cards):
    x = 0.72 + i * 3.05
    card(s, x, 1.58, 2.72, 4.65, fill)
    textbox(s, head, x + 0.25, 2.02, 2.22, 0.32, 12, accent, True, PP_ALIGN.CENTER)
    textbox(s, detail, x + 0.32, 3.05, 2.08, 1.52, 16, NAVY, True, PP_ALIGN.CENTER)
    textbox(s, ["Be intentional", "Be accurate", "Be respectful", "Be cautious"][i],
            x + 0.32, 5.32, 2.08, 0.38, 12, MUTED, True, PP_ALIGN.CENTER)
footer(s)

# 12 — 30-day program
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s)
title(s, "30-day personal development plan", "Use one observable goal, one weekly behavior and one checkpoint", 12)
weeks = [
    ("WEEK 1", "AWARENESS", "Complete strengths and trigger review", PALE_BLUE, BLUE),
    ("WEEK 2", "COMMUNICATION", "Practice active listening daily", PALE_TEAL, TEAL),
    ("WEEK 3", "CONFIDENCE", "Deliver one structured contribution", PALE_GOLD, GOLD),
    ("WEEK 4", "COLLABORATION", "Lead or support one team process", PALE_PURPLE, RGBColor(124,101,166)),
]
for i, (week, head, action, fill, accent) in enumerate(weeks):
    x = 0.72 + i * 3.05
    card(s, x, 1.55, 2.72, 4.45, fill)
    textbox(s, week, x + 0.28, 1.92, 2.16, 0.28, 11, accent, True)
    textbox(s, head, x + 0.28, 2.55, 2.16, 0.56, 18, NAVY, True)
    textbox(s, action, x + 0.28, 3.55, 2.16, 1.18, 14, TEXT)
    textbox(s, f"{(i+1)*7} DAYS", x + 0.28, 5.45, 2.16, 0.28, 10, MUTED, True)
footer(s)

# 13 — Assessment
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s)
title(s, "Measure growth without perfection", "Track evidence, not vague feelings", 13)
card(s, 0.72, 1.55, 11.84, 4.82, WHITE)
metrics = [
    ("Consistency", "How often did I demonstrate the behavior?", "8 / 10"),
    ("Quality", "Did the action achieve the intended result?", "Evidence"),
    ("Feedback", "What did others observe?", "Specific"),
    ("Learning", "What changed in my next attempt?", "Visible"),
]
textbox(s, "METRIC", 1.05, 1.92, 2.2, 0.28, 11, MUTED, True)
textbox(s, "QUESTION", 4.0, 1.92, 5.15, 0.28, 11, MUTED, True)
textbox(s, "TARGET", 9.85, 1.92, 1.65, 0.28, 11, MUTED, True)
for i, (metric, question, target) in enumerate(metrics):
    y = 2.48 + i * 0.82
    textbox(s, metric, 1.05, y, 2.2, 0.28, 15, NAVY, True)
    textbox(s, question, 4.0, y, 5.15, 0.28, 14, TEXT)
    textbox(s, target, 9.85, y, 1.65, 0.28, 14, TEAL, True)
footer(s)

# 14 — Closing
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s)
card(s, 0.72, 0.72, 11.84, 5.95, PALE_BLUE)
textbox(s, "Build character through\nsmall, consistent choices.", 1.22, 1.45, 7.2, 1.38, 34, NAVY, True)
textbox(s, "Know yourself. Listen deliberately. Respond professionally.", 1.22, 3.38,
        8.65, 0.42, 20, TEAL, True)
textbox(s, "YOUR NEXT ACTION", 1.22, 4.45, 2.2, 0.28, 11, BLUE, True)
textbox(s, "Choose one behavior.\nPractice it seven days.\nAsk for evidence.", 1.22, 4.88, 4.1, 1.25, 16, TEXT)
circle(s, 9.68, 2.05, 1.82, GOLD, "ACT", NAVY, 18)
textbox(s, "THANK YOU", 1.22, 6.48, 2.5, 0.28, 11, MUTED, True)


output = "Personality_Development_Professional_Animated_Corrected.pptx"
prs.save(output)


def add_fade_transitions(pptx_path):
    P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    tmp = ".professional_tmp.pptx"
    count = 0
    with ZipFile(pptx_path, "r") as source, ZipFile(tmp, "w", ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename.startswith("ppt/slides/slide") and item.filename.endswith(".xml"):
                root = etree.fromstring(data)
                for old in root.xpath("./p:transition", namespaces={"p": P}):
                    root.remove(old)
                transition = etree.Element("{%s}transition" % P, spd="med", advClick="1")
                etree.SubElement(transition, "{%s}fade" % P)
                root.insert(1, transition)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                count += 1
            target.writestr(item, data)
    os.replace(tmp, pptx_path)
    return count


print(output)
print("fade transitions:", add_fade_transitions(output))
