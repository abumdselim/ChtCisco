import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Polygon, Circle, FancyArrow
import numpy as np
import io
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.oxml.ns import qn
from lxml import etree

# ── Color Palette ──────────────────────────────────────────────────────────────
PRIMARY_BLUE  = RGBColor(0x1a, 0x3a, 0x6b)
TEAL          = RGBColor(0x00, 0xb4, 0xd8)
TEAL_BRIGHT   = RGBColor(0x00, 0xd4, 0xff)
ORANGE        = RGBColor(0xff, 0x8c, 0x00)
ORANGE_ALERT  = RGBColor(0xe6, 0x5c, 0x00)
PURPLE        = RGBColor(0x7c, 0x3a, 0xed)
GREEN_ACCENT  = RGBColor(0x2e, 0xa0, 0x4e)
DARK_SLATE    = RGBColor(0x2d, 0x3d, 0x4f)
LIGHT_GRAY    = RGBColor(0xf0, 0xf0, 0xf0)
WHITE         = RGBColor(0xff, 0xff, 0xff)
NAVY          = RGBColor(0x1a, 0x23, 0x32)

# ── Helpers ────────────────────────────────────────────────────────────────────
def set_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_footer(slide, slide_num, W, H):
    fb = slide.shapes.add_textbox(Inches(0.2), H - Inches(0.35), Inches(8), Inches(0.3))
    tf = fb.text_frame
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "HydroNet - Chittagong Hydrological Resilience"
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    p.alignment = PP_ALIGN.LEFT

    nb = slide.shapes.add_textbox(W - Inches(1), H - Inches(0.35), Inches(0.8), Inches(0.3))
    tf2 = nb.text_frame
    p2 = tf2.paragraphs[0]
    r2 = p2.add_run()
    r2.text = str(slide_num)
    r2.font.size = Pt(8)
    r2.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    p2.alignment = PP_ALIGN.RIGHT


def add_title(slide, text, top=Inches(0.3), color=None):
    if color is None:
        color = PRIMARY_BLUE
    tb = slide.shapes.add_textbox(Inches(0.4), top, Inches(9.2), Inches(0.6))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    r.font.size = Pt(30)
    r.font.bold = True
    r.font.color.rgb = color
    r.font.name = 'Calibri'
    return tb


def add_teal_underline(slide, top_offset):
    ul = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0.4), top_offset, Inches(9.2), Inches(0.045)
    )
    ul.fill.solid()
    ul.fill.fore_color.rgb = TEAL
    ul.line.fill.background()


def fig_to_buf(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    plt.close(fig)
    return buf


def add_fade_transition(slide):
    slide_elem = slide._element
    transition = etree.SubElement(slide_elem, qn('p:transition'))
    transition.set('spd', 'slow')
    etree.SubElement(transition, qn('p:fade'))


def add_rect(slide, l, t, w, h, fill_color, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape


def add_textbox(slide, text, l, t, w, h, font_size=10, bold=False, italic=False,
                color=None, align=PP_ALIGN.LEFT, wrap=True, name='Calibri'):
    if color is None:
        color = DARK_SLATE
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.name = name
    p.alignment = align
    return tb


# ── Presentation setup ─────────────────────────────────────────────────────────
prs = Presentation()
prs.slide_width  = Inches(10)
prs.slide_height = Inches(5.625)
W = prs.slide_width
H = prs.slide_height
blank = prs.slide_layouts[6]

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — Title Slide
# ══════════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
set_bg(s1, NAVY)

fig, ax = plt.subplots(figsize=(10, 5.625))
ax.set_xlim(0, 10); ax.set_ylim(0, 5.625)
ax.set_facecolor('#1a2332'); fig.patch.set_facecolor('#1a2332')
stripe = Polygon([[2, 5.625], [4.5, 5.625], [8.5, 0], [6, 0]],
                 closed=True, color='#00b4d8', alpha=0.12)
ax.add_patch(stripe)
stripe2 = Polygon([[0, 3.5], [0, 5.625], [1.5, 5.625], [0, 4.5]],
                  closed=True, color='#00d4ff', alpha=0.08)
ax.add_patch(stripe2)
ax.axis('off')
buf = fig_to_buf(fig)
s1.shapes.add_picture(buf, 0, 0, W, H)

# Banner behind title
banner = add_rect(s1, Inches(0.5), Inches(1.05), Inches(9), Inches(0.85),
                  RGBColor(0x00, 0x60, 0x80))
banner.line.fill.background()

# Title
tb = s1.shapes.add_textbox(Inches(0.5), Inches(1.05), Inches(9), Inches(0.85))
tf = tb.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "CHITTAGONG HYDROLOGICAL RESILIENCE NETWORK"
r.font.size = Pt(26); r.font.bold = True
r.font.color.rgb = WHITE; r.font.name = 'Calibri'

# Subtitle
tb2 = s1.shapes.add_textbox(Inches(0.5), Inches(2.05), Inches(9), Inches(0.5))
tf2 = tb2.text_frame; tf2.word_wrap = True
p2 = tf2.paragraphs[0]; p2.alignment = PP_ALIGN.CENTER
r2 = p2.add_run()
r2.text = "HydroNet: Cisco-Powered Real-Time Flood Monitoring & Early Warning System"
r2.font.size = Pt(14); r2.font.italic = True
r2.font.color.rgb = TEAL; r2.font.name = 'Calibri'

# Team block
tb3 = s1.shapes.add_textbox(Inches(1), Inches(2.75), Inches(8), Inches(0.4))
tf3 = tb3.text_frame; p3 = tf3.paragraphs[0]; p3.alignment = PP_ALIGN.CENTER
r3 = p3.add_run()
r3.text = "Abu Md. Selim  |  Arifur Rahman  |  Sadab Abdullah"
r3.font.size = Pt(11); r3.font.color.rgb = WHITE; r3.font.name = 'Calibri'

# Org
tb4 = s1.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9), Inches(0.35))
tf4 = tb4.text_frame; p4 = tf4.paragraphs[0]; p4.alignment = PP_ALIGN.CENTER
r4 = p4.add_run()
r4.text = "Premier University, Chittagong  •  6th Semester Network Architecture Project  •  March 2025"
r4.font.size = Pt(10); r4.font.color.rgb = RGBColor(0xcc, 0xcc, 0xcc)

# Decorative circles bottom-left
circle_colors = [TEAL, ORANGE, PURPLE, GREEN_ACCENT]
for i, cc in enumerate(circle_colors):
    cx = Inches(0.4 + i * 0.42)
    cy = H - Inches(0.7)
    c = s1.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, cx, cy, Inches(0.28), Inches(0.28))
    c.fill.solid(); c.fill.fore_color.rgb = cc
    c.line.fill.background()

add_footer(s1, 1, W, H)
add_fade_transition(s1)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — Problem Statement & Motivation
# ══════════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
set_bg(s2, LIGHT_GRAY)
add_title(s2, "Why HydroNet? The Flood Crisis in Chittagong", top=Inches(0.28))
add_teal_underline(s2, Inches(0.92))

# Left column
left_lines = [
    ("THE PROBLEM:", True),
    ("Chittagong (pop. 8M+) sits at the confluence of the Karnaphuli River and the Bay of Bengal.", False),
    ("Annual monsoon flooding has devastated the region repeatedly:", False),
    ("• 2023 floods: 51+ lives lost, 1M+ people displaced", False),
    ("• 2022: Major tidal surge damaged Chittagong Port, economic loss >$200M", False),
    ("• Average annual flood damage: $500M across infrastructure", False),
    ("• Current early warning: Manual observation, 30–60 min lead time only", False),
    ("• Critical gap: No automated, real-time sensor network exists", False),
    ("", False),
    ("THE CONSEQUENCE:", True),
    ("• Delayed evacuation = preventable casualties", False),
    ("• First responders lack real-time situational awareness", False),
    ("• Infrastructure damage compounds without advance warning", False),
    ("• Vulnerable populations (coastal poor, river communities) most at risk", False),
]

tb_left = s2.shapes.add_textbox(Inches(0.4), Inches(1.05), Inches(5.3), Inches(4.3))
tf_l = tb_left.text_frame; tf_l.word_wrap = True
for i, (line, bold) in enumerate(left_lines):
    if i == 0:
        p = tf_l.paragraphs[0]
    else:
        p = tf_l.add_paragraph()
    r = p.add_run()
    r.text = line
    r.font.size = Pt(9)
    r.font.bold = bold
    r.font.color.rgb = PRIMARY_BLUE if bold else DARK_SLATE
    r.font.name = 'Calibri'

# Right column: bar chart
years  = [2019, 2020, 2021, 2022, 2023]
deaths = [12, 18, 23, 35, 51]
damage = [180, 220, 310, 410, 500]

fig2, ax1 = plt.subplots(figsize=(4.2, 3.4))
fig2.patch.set_facecolor('#f0f0f0')
ax1.set_facecolor('#f8f8f8')
bars = ax1.bar(years, deaths, color='#ff8c00', alpha=0.85, width=0.55, label='Deaths')
ax1.set_ylabel('Deaths', color='#ff8c00', fontsize=8)
ax1.tick_params(axis='y', labelcolor='#ff8c00', labelsize=7)
ax1.set_xlabel('Year', fontsize=8)
ax1.set_title('Flood Impact 2019–2023 (Chittagong)', fontsize=9, fontweight='bold', color='#1a3a6b')
ax1.set_xticks(years); ax1.tick_params(axis='x', labelsize=7)
ax2 = ax1.twinx()
ax2.plot(years, damage, color='#1a3a6b', marker='o', linewidth=2, markersize=5, label='Damage $M')
ax2.set_ylabel('Damage ($M)', color='#1a3a6b', fontsize=8)
ax2.tick_params(axis='y', labelcolor='#1a3a6b', labelsize=7)
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=7, loc='upper left')
fig2.tight_layout(pad=0.5)
buf2 = fig_to_buf(fig2)
s2.shapes.add_picture(buf2, Inches(5.8), Inches(1.05), Inches(3.9), Inches(3.4))

add_footer(s2, 2, W, H)
add_fade_transition(s2)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — Solution Architecture Overview
# ══════════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
set_bg(s3, WHITE)
add_title(s3, "HydroNet: System Architecture at a Glance", top=Inches(0.28))
add_teal_underline(s3, Inches(0.92))

fig3, ax3 = plt.subplots(figsize=(9.5, 3.3))
fig3.patch.set_facecolor('white'); ax3.set_facecolor('white')
ax3.set_xlim(0, 10); ax3.set_ylim(0, 3.3)
ax3.axis('off')

def draw_box(ax, cx, cy, w, h, label, facecolor, fontsize=7, textcolor='white'):
    rect = FancyBboxPatch((cx - w/2, cy - h/2), w, h,
                          boxstyle="round,pad=0.05", facecolor=facecolor,
                          edgecolor='white', linewidth=0.8)
    ax.add_patch(rect)
    ax.text(cx, cy, label, ha='center', va='center',
            fontsize=fontsize, color=textcolor, fontweight='bold', wrap=True)

# Layer 1 – Monitoring
layer1_y = 2.75
for i, lbl in enumerate(["Zone-1\nRouter", "Zone-2\nRouter", "Zone-3\nRouter", "Zone-4\nRouter"]):
    cx = 1.5 + i * 2.3
    draw_box(ax3, cx, layer1_y, 1.7, 0.45, lbl, '#1a3a6b')
ax3.text(0.1, layer1_y, "Monitoring\nLayer", fontsize=7, color='#1a3a6b', va='center', fontweight='bold')

# Layer 2 – Core
layer2_y = 1.85
for i, lbl in enumerate(["Core-Switch-1", "Core-Switch-2", "Core-Router-1", "Core-Router-2"]):
    cx = 1.8 + i * 2.0
    draw_box(ax3, cx, layer2_y, 1.7, 0.42, lbl, '#00b4d8', textcolor='white')
ax3.text(0.1, layer2_y, "Core\nNetwork", fontsize=7, color='#00b4d8', va='center', fontweight='bold')

# ISP
draw_box(ax3, 9.5, layer2_y, 0.9, 0.38, "ISP\nRouter", '#ff8c00')
ax3.annotate('', xy=(8.05, layer2_y), xytext=(9.05, layer2_y),
             arrowprops=dict(arrowstyle='->', color='#ff8c00', lw=1.5))

# Layer 3 – Services
layer3_y = 0.95
for i, lbl in enumerate(["DHCP", "DNS", "WEB", "EMAIL", "MQTT"]):
    cx = 1.4 + i * 1.8
    draw_box(ax3, cx, layer3_y, 1.4, 0.38, lbl, '#2ea04e')
ax3.text(0.1, layer3_y, "Services\nLayer", fontsize=7, color='#2ea04e', va='center', fontweight='bold')

# Connecting lines L1→L2
zone_xs = [1.5, 3.8, 6.1, 8.4]
core_xs = [1.8, 3.8, 5.8, 7.8]
for zx in zone_xs:
    closest_cx = min(core_xs, key=lambda x: abs(x - zx))
    ax3.annotate('', xy=(closest_cx, layer2_y + 0.21),
                 xytext=(zx, layer1_y - 0.22),
                 arrowprops=dict(arrowstyle='->', color='#aaaaaa', lw=0.8))

# Connecting lines L2→L3
for cx in core_xs:
    ax3.annotate('', xy=(min(core_xs, key=lambda x: abs(x-cx)), layer3_y + 0.19),
                 xytext=(cx, layer2_y - 0.21),
                 arrowprops=dict(arrowstyle='->', color='#aaaaaa', lw=0.8))

buf3 = fig_to_buf(fig3)
s3.shapes.add_picture(buf3, Inches(0.15), Inches(1.0), Inches(9.7), Inches(3.35))

# Bottom summary boxes
box_defs = [
    (Inches(0.3),  "20+ IoT Sensors | 4 Geographic Zones | Real-time 30s updates", RGBColor(0xff, 0xf0, 0xe0)),
    (Inches(3.55), "OSPF Dynamic Routing | 12 VLANs | 5 Core Servers",              RGBColor(0xe0, 0xf7, 0xff)),
    (Inches(6.8),  "NAT + ACL Security | DHCP/DNS/Web/Email | Early Warning Emails", RGBColor(0xf3, 0xec, 0xff)),
]
border_colors = [ORANGE, TEAL, PURPLE]
for (bl, txt, bg), bc in zip(box_defs, border_colors):
    bx = add_rect(s3, bl, Inches(4.42), Inches(3.05), Inches(0.82), bg, bc, 1.5)
    tb_b = s3.shapes.add_textbox(bl + Inches(0.1), Inches(4.47), Inches(2.85), Inches(0.72))
    tf_b = tb_b.text_frame; tf_b.word_wrap = True
    pb = tf_b.paragraphs[0]
    rb = pb.add_run(); rb.text = txt
    rb.font.size = Pt(8.5); rb.font.color.rgb = DARK_SLATE; rb.font.name = 'Calibri'

add_footer(s3, 3, W, H)
add_fade_transition(s3)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — Core Network Infrastructure
# ══════════════════════════════════════════════════════════════════════════════
s4 = prs.slides.add_slide(blank)
set_bg(s4, LIGHT_GRAY)
add_title(s4, "Core Network Infrastructure & IP Addressing", top=Inches(0.28))
add_teal_underline(s4, Inches(0.9))

fig4, ax4 = plt.subplots(figsize=(5.2, 3.9))
fig4.patch.set_facecolor('#f0f0f0'); ax4.set_facecolor('#f5f5f5')
ax4.set_xlim(0, 5.2); ax4.set_ylim(0, 3.9); ax4.axis('off')

def dbox(ax, cx, cy, w, h, title, sub, fc, tc='white', fs=6.5):
    r = FancyBboxPatch((cx-w/2, cy-h/2), w, h,
                       boxstyle="round,pad=0.06", facecolor=fc, edgecolor='#888', linewidth=0.7)
    ax.add_patch(r)
    ax.text(cx, cy+0.06, title, ha='center', va='center',
            fontsize=fs, color=tc, fontweight='bold')
    if sub:
        ax.text(cx, cy-0.14, sub, ha='center', va='center', fontsize=5.5, color=tc, alpha=0.9)

# ISP
dbox(ax4, 4.5, 3.4, 0.8, 0.32, "ISP Router", "172.16.0.1", '#e65c00')
# Core Routers
dbox(ax4, 1.3, 2.7, 1.3, 0.35, "Core-Router-1", "10.0.0.1/8", '#1a3a6b')
dbox(ax4, 3.5, 2.7, 1.3, 0.35, "Core-Router-2", "10.0.0.2/8", '#1a3a6b')
# Core Switches
dbox(ax4, 1.3, 1.9, 1.3, 0.35, "Core-Switch-1", "10.0.0.10", '#00b4d8', 'white')
dbox(ax4, 3.5, 1.9, 1.3, 0.35, "Core-Switch-2", "10.0.0.11", '#00b4d8', 'white')
# Servers
server_info = [
    ("DHCP","10.0.100.1"), ("DNS","10.0.100.2"), ("WEB","10.0.100.3"),
    ("EMAIL","10.0.100.4"), ("SYSLOG","10.0.100.5")
]
for i, (nm, ip) in enumerate(server_info):
    cx4 = 0.55 + i * 0.95
    dbox(ax4, cx4, 0.9, 0.85, 0.35, nm, ip, '#2ea04e', fs=6)

# Connecting lines
def line(ax, x1, y1, x2, y2, c='#888888', lw=1.2):
    ax.plot([x1, x2], [y1, y2], '-', color=c, linewidth=lw)

line(ax4, 4.5, 3.24, 3.5, 2.87)  # ISP → CR2
line(ax4, 1.3, 2.52, 1.3, 2.07)  # CR1 ↓ CS1
line(ax4, 3.5, 2.52, 3.5, 2.07)  # CR2 ↓ CS2
line(ax4, 1.95, 2.7, 2.85, 2.7)  # CR1 — CR2
line(ax4, 1.95, 1.9, 2.85, 1.9)  # CS1 — CS2
for i in range(5):
    cx4 = 0.55 + i * 0.95
    line(ax4, cx4, 1.07, cx4, 1.72)

buf4 = fig_to_buf(fig4)
s4.shapes.add_picture(buf4, Inches(0.25), Inches(1.0), Inches(5.2), Inches(3.95))

# Right text
rt_lines = [
    ("CORE DEVICES:", True, TEAL),
    ("• Core-Router-1: Cisco 3945 | 10.0.0.1/8 | OSPF DR (priority 255)", False, DARK_SLATE),
    ("• Core-Router-2: Cisco 3945 | 10.0.0.2/8 | OSPF BDR", False, DARK_SLATE),
    ("• Core-Switch-1: Catalyst 3750-X | 10.0.0.10 | Layer-3 capable", False, DARK_SLATE),
    ("• Core-Switch-2: Catalyst 3750-X | 10.0.0.11 | Redundant", False, DARK_SLATE),
    ("", False, DARK_SLATE),
    ("SERVER FARM (10.0.100.0/24):", True, TEAL),
    ("• DHCP Server: 10.0.100.1 | Cisco Router | 4 zone pools", False, DARK_SLATE),
    ("• DNS Server: 10.0.100.2 | BIND | hydronet.local", False, DARK_SLATE),
    ("• Web Server: 10.0.100.3 | Apache | monitor.hydronet.local", False, DARK_SLATE),
    ("• Email Server: 10.0.100.4 | SMTP:25 / POP3:110", False, DARK_SLATE),
    ("• MQTT Broker: 10.0.100.1 | Port 1883 | QoS-2", False, DARK_SLATE),
    ("", False, DARK_SLATE),
    ("CONNECTIVITY:", True, TEAL),
    ("• ISP Link: 172.16.0.0/30 | 100 Mbps", False, DARK_SLATE),
    ("• Core Backbone: Gigabit Ethernet | Cost=1", False, DARK_SLATE),
    ("• Zone Uplinks: FastEthernet | Cost=10", False, DARK_SLATE),
]
tb_rt = s4.shapes.add_textbox(Inches(5.65), Inches(1.05), Inches(4.1), Inches(4.2))
tf_rt = tb_rt.text_frame; tf_rt.word_wrap = True
for i, (txt, bold, col) in enumerate(rt_lines):
    p = tf_rt.paragraphs[0] if i == 0 else tf_rt.add_paragraph()
    r = p.add_run(); r.text = txt
    r.font.size = Pt(8.5); r.font.bold = bold
    r.font.color.rgb = col; r.font.name = 'Calibri'

add_footer(s4, 4, W, H)
add_fade_transition(s4)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — Zone Network Topology
# ══════════════════════════════════════════════════════════════════════════════
s5 = prs.slides.add_slide(blank)
set_bg(s5, WHITE)
add_title(s5, "4-Zone Network Topology & Wireless Deployment", top=Inches(0.2))
add_teal_underline(s5, Inches(0.82))

fig5, ax5 = plt.subplots(figsize=(9.5, 4.2))
fig5.patch.set_facecolor('white'); ax5.set_facecolor('white')
ax5.set_xlim(0, 9.5); ax5.set_ylim(0, 4.2); ax5.axis('off')

zone_cfg = [
    ("Zone 1\nCoastal", 1.3, 3.5, '#ff8c00', '#fff0e0'),
    ("Zone 2\nRiverine", 8.2, 3.5, '#00b4d8', '#e0f7ff'),
    ("Zone 3\nHilly",   1.3, 0.9, '#7c3aed', '#f0e0ff'),
    ("Zone 4\nUrban",   8.2, 0.9, '#2ea04e', '#e0ffe0'),
]

core_cx, core_cy = 4.75, 2.2
# Core box
core_rect = FancyBboxPatch((core_cx-1.2, core_cy-0.55), 2.4, 1.1,
                           boxstyle="round,pad=0.08", facecolor='#1a3a6b',
                           edgecolor='#00b4d8', linewidth=1.5)
ax5.add_patch(core_rect)
ax5.text(core_cx, core_cy+0.18, "CORE NETWORK", ha='center', va='center',
         fontsize=8, color='white', fontweight='bold')
ax5.text(core_cx, core_cy-0.12, "CR-1 • CR-2 • CS-1 • CS-2", ha='center', va='center',
         fontsize=6.5, color='#00d4ff')

for (zlbl, zx, zy, zcol, zbg) in zone_cfg:
    # Zone panel
    zpatch = FancyBboxPatch((zx-1.1, zy-0.75), 2.2, 1.5,
                            boxstyle="round,pad=0.07", facecolor=zbg,
                            edgecolor=zcol, linewidth=1.5)
    ax5.add_patch(zpatch)
    ax5.text(zx, zy+0.5, zlbl, ha='center', va='center',
             fontsize=8, color=zcol, fontweight='bold')
    # Sub-devices
    for j, sub in enumerate(["ZR", "ZS", "AP"]):
        sx = zx - 0.55 + j * 0.55
        sy = zy + 0.05
        sr = FancyBboxPatch((sx-0.2, sy-0.15), 0.4, 0.3,
                            boxstyle="round,pad=0.03", facecolor=zcol, alpha=0.8,
                            edgecolor='white', linewidth=0.5)
        ax5.add_patch(sr)
        ax5.text(sx, sy, sub, ha='center', va='center', fontsize=5.5, color='white', fontweight='bold')
    # Sensor dots
    for k in range(5):
        sx2 = zx - 0.9 + k * 0.45
        ax5.plot(sx2, zy - 0.5, 's', color=zcol, markersize=5, alpha=0.7)
    ax5.text(zx, zy - 0.62, "5 Sensors", ha='center', va='center',
             fontsize=5.5, color=zcol)

    # Line to core
    ax5.annotate('', xy=(core_cx + (0 if abs(zx-core_cx)<3 else (0.3 if zx>core_cx else -0.3)),
                         core_cy + (0.35 if zy > core_cy else -0.35)),
                 xytext=(zx + (0.6 if zx < core_cx else -0.6),
                         zy + (-0.5 if zy > core_cy else 0.5)),
                 arrowprops=dict(arrowstyle='<->', color=zcol, lw=1.4))

# Legend
legend_items = [
    mpatches.Patch(color='#ff8c00', label='Zone 1 – Coastal Port'),
    mpatches.Patch(color='#00b4d8', label='Zone 2 – Riverine Lowlands'),
    mpatches.Patch(color='#7c3aed', label='Zone 3 – Hilly Upstream'),
    mpatches.Patch(color='#2ea04e', label='Zone 4 – Urban Lowland'),
    mpatches.Patch(color='#1a3a6b', label='Core Network'),
]
ax5.legend(handles=legend_items, loc='lower center', ncol=5,
           fontsize=7, framealpha=0.9, bbox_to_anchor=(0.5, -0.01))

buf5 = fig_to_buf(fig5)
s5.shapes.add_picture(buf5, Inches(0.1), Inches(0.9), Inches(9.7), Inches(4.3))

add_footer(s5, 5, W, H)
add_fade_transition(s5)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — VLAN Architecture
# ══════════════════════════════════════════════════════════════════════════════
s6 = prs.slides.add_slide(blank)
set_bg(s6, LIGHT_GRAY)
add_title(s6, "VLAN Segmentation: 12-VLAN Architecture", top=Inches(0.25))
add_teal_underline(s6, Inches(0.88))

vlans = [
    (10,  "SENSORS-Z1",  "10.10.10.0/24",  "Zone-1 IoT Sensors"),
    (20,  "SENSORS-Z2",  "10.10.20.0/24",  "Zone-2 IoT Sensors"),
    (30,  "SENSORS-Z3",  "10.10.30.0/24",  "Zone-3 IoT Sensors"),
    (40,  "SENSORS-Z4",  "10.10.40.0/24",  "Zone-4 IoT Sensors"),
    (50,  "MANAGEMENT",  "10.0.50.0/24",   "Network Mgmt"),
    (60,  "SERVERS",     "10.0.100.0/24",  "Server Farm"),
    (70,  "WIRELESS-Z1", "10.20.10.0/24",  "Zone-1 Wireless"),
    (80,  "WIRELESS-Z2", "10.20.20.0/24",  "Zone-2 Wireless"),
    (90,  "WIRELESS-Z3", "10.20.30.0/24",  "Zone-3 Wireless"),
    (100, "WIRELESS-Z4", "10.20.40.0/24",  "Zone-4 Wireless"),
    (110, "CONTROL",     "10.0.110.0/24",  "Control Plane"),
    (1,   "NATIVE",      "N/A",            "802.1Q Native VLAN"),
]

vlan_colors_map = {
    10:'#1a3a6b',20:'#1a4a7b',30:'#1a5a8b',40:'#1a6a9b',
    50:'#ff8c00',60:'#2ea04e',70:'#00b4d8',80:'#00c4e8',
    90:'#00d4f8',100:'#7c3aed',110:'#9c5afd',1:'#888888'
}

fig6, ax6 = plt.subplots(figsize=(9.0, 2.0))
fig6.patch.set_facecolor('#f0f0f0'); ax6.set_facecolor('#f0f0f0')
ax6.set_xlim(0, 12); ax6.set_ylim(0, 2.0); ax6.axis('off')

# Switch body
sw = FancyBboxPatch((0.2, 0.5), 11.6, 1.0, boxstyle="round,pad=0.05",
                    facecolor='#2d3d4f', edgecolor='#1a3a6b', linewidth=1.5)
ax6.add_patch(sw)
ax6.text(6.0, 1.0, "Core Switch — 12 VLANs  (802.1Q Trunk)", ha='center', va='center',
         fontsize=7, color='white', fontweight='bold')

band_w = 11.6 / 12
for i, (vid, name, _, _) in enumerate(vlans):
    bx = 0.2 + i * band_w
    col = vlan_colors_map.get(vid, '#555555')
    band = FancyBboxPatch((bx+0.04, 0.58), band_w-0.08, 0.84,
                          boxstyle="round,pad=0.02", facecolor=col, alpha=0.85,
                          edgecolor='white', linewidth=0.5)
    ax6.add_patch(band)
    ax6.text(bx + band_w/2, 1.12, str(vid), ha='center', va='center',
             fontsize=6.5, color='white', fontweight='bold')
    ax6.text(bx + band_w/2, 0.78, name[:8], ha='center', va='center',
             fontsize=5, color='white', rotation=0)

buf6 = fig_to_buf(fig6)
s6.shapes.add_picture(buf6, Inches(0.3), Inches(1.0), Inches(9.4), Inches(1.9))

# Two-column VLAN table
col_headers = ["VLAN", "Name", "Subnet", "Purpose"]
hdr_txt = "  ".join(f"{h:<13}" for h in col_headers)

tb6L = s6.shapes.add_textbox(Inches(0.3), Inches(3.0), Inches(4.6), Inches(2.3))
tf6L = tb6L.text_frame; tf6L.word_wrap = False
p6h = tf6L.paragraphs[0]
rh = p6h.add_run()
rh.text = f"{'ID':<5}{'Name':<14}{'Subnet':<17}{'Purpose'}"
rh.font.size = Pt(7.5); rh.font.bold = True; rh.font.color.rgb = PRIMARY_BLUE

for vid, name, subnet, purpose in vlans[:6]:
    p = tf6L.add_paragraph()
    r = p.add_run()
    r.text = f"{vid:<5}{name:<14}{subnet:<17}{purpose}"
    r.font.size = Pt(7.5); r.font.color.rgb = DARK_SLATE; r.font.name = 'Courier New'

tb6R = s6.shapes.add_textbox(Inches(5.1), Inches(3.0), Inches(4.6), Inches(2.3))
tf6R = tb6R.text_frame; tf6R.word_wrap = False
p6h2 = tf6R.paragraphs[0]
rh2 = p6h2.add_run()
rh2.text = f"{'ID':<5}{'Name':<14}{'Subnet':<17}{'Purpose'}"
rh2.font.size = Pt(7.5); rh2.font.bold = True; rh2.font.color.rgb = PRIMARY_BLUE

for vid, name, subnet, purpose in vlans[6:]:
    p = tf6R.add_paragraph()
    r = p.add_run()
    r.text = f"{vid:<5}{name:<14}{subnet:<17}{purpose}"
    r.font.size = Pt(7.5); r.font.color.rgb = DARK_SLATE; r.font.name = 'Courier New'

# Note
add_textbox(s6, "802.1Q trunking on all uplinks  |  Inter-VLAN routing via ACL  |  VLAN 100 reachable from all zones",
            Inches(0.3), Inches(5.18), Inches(9.4), Inches(0.28),
            font_size=8, italic=True, color=PRIMARY_BLUE, align=PP_ALIGN.CENTER)

add_footer(s6, 6, W, H)
add_fade_transition(s6)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — Server Infrastructure & Services
# ══════════════════════════════════════════════════════════════════════════════
s7 = prs.slides.add_slide(blank)
set_bg(s7, WHITE)
add_title(s7, "Centralized Services: DHCP, DNS, Web & Email", top=Inches(0.25))
add_teal_underline(s7, Inches(0.88))

cards = [
    ("DHCP Server", "10.0.100.1",
     ["• Cisco Router-based DHCP",
      "• Zone-1 Pool: 10.10.10.10–254",
      "• Zone-2 Pool: 10.10.20.10–254",
      "• Zone-3 Pool: 10.10.30.10–254",
      "• Zone-4 Pool: 10.10.40.10–254",
      "• Lease time: 24 hours",
      "• DNS: 10.0.100.2 pushed",
      "• Default GW per zone"],
     RGBColor(0xff, 0xf0, 0xe0), ORANGE),
    ("DNS Server", "10.0.100.2",
     ["• BIND 9 – hydronet.local",
      "• A: monitor → 10.0.100.3",
      "• A: mail → 10.0.100.4",
      "• A: dhcp → 10.0.100.1",
      "• A: syslog → 10.0.100.5",
      "• Forwarder: 8.8.8.8",
      "• Zone: hydronet.local",
      "• Reverse DNS configured"],
     RGBColor(0xe0, 0xf7, 0xff), TEAL),
    ("Web Server", "10.0.100.3",
     ["• Apache HTTP 2.4",
      "• URL: monitor.hydronet.local",
      "• Dashboard: real-time data",
      "• Port 80 (HTTP)",
      "• Sensor data via MQTT",
      "• Charts: water level & flow",
      "• Alert banner on threshold",
      "• Accessible all VLANs"],
     RGBColor(0xf3, 0xec, 0xff), PURPLE),
    ("Email Server", "10.0.100.4",
     ["• Postfix SMTP – Port 25",
      "• Dovecot POP3 – Port 110",
      "• Domain: hydronet.local",
      "• Alerts → admin@hydronet",
      "• Flood threshold email",
      "• Zone status digest",
      "• SMTP auth required",
      "• Log: /var/log/mail.log"],
     RGBColor(0xe6, 0xff, 0xe6), GREEN_ACCENT),
]

card_w = Inches(2.3)
card_h = Inches(3.8)
for i, (title, ip, lines, bg, border) in enumerate(cards):
    cl = Inches(0.3 + i * 2.38)
    ct = Inches(1.0)
    add_rect(s7, cl, ct, card_w, card_h, bg, border, 1.5)
    # Card title bar
    add_rect(s7, cl, ct, card_w, Inches(0.42), border)
    tb_ct = s7.shapes.add_textbox(cl + Inches(0.05), ct + Inches(0.03),
                                   card_w - Inches(0.1), Inches(0.38))
    tf_ct = tb_ct.text_frame; p_ct = tf_ct.paragraphs[0]
    p_ct.alignment = PP_ALIGN.CENTER
    r_ct = p_ct.add_run(); r_ct.text = title
    r_ct.font.size = Pt(10); r_ct.font.bold = True
    r_ct.font.color.rgb = WHITE; r_ct.font.name = 'Calibri'
    # IP
    tb_ip = s7.shapes.add_textbox(cl + Inches(0.05), ct + Inches(0.44),
                                   card_w - Inches(0.1), Inches(0.25))
    tf_ip = tb_ip.text_frame; p_ip = tf_ip.paragraphs[0]
    p_ip.alignment = PP_ALIGN.CENTER
    r_ip = p_ip.add_run(); r_ip.text = ip
    r_ip.font.size = Pt(8.5); r_ip.font.bold = True
    r_ip.font.color.rgb = border; r_ip.font.name = 'Courier New'
    # Lines
    tb_li = s7.shapes.add_textbox(cl + Inches(0.1), ct + Inches(0.75),
                                   card_w - Inches(0.15), Inches(2.9))
    tf_li = tb_li.text_frame; tf_li.word_wrap = True
    for j, ln in enumerate(lines):
        p2 = tf_li.paragraphs[0] if j == 0 else tf_li.add_paragraph()
        r2 = p2.add_run(); r2.text = ln
        r2.font.size = Pt(8); r2.font.color.rgb = DARK_SLATE; r2.font.name = 'Calibri'

# Bottom note
add_rect(s7, Inches(0.3), Inches(4.95), Inches(9.4), Inches(0.35),
         RGBColor(0xe8, 0xf4, 0xff), TEAL, 1)
add_textbox(s7, "All services tested and verified in Cisco Packet Tracer 8.2.2",
            Inches(0.4), Inches(4.98), Inches(9.2), Inches(0.28),
            font_size=9, italic=True, color=PRIMARY_BLUE, align=PP_ALIGN.CENTER)

add_footer(s7, 7, W, H)
add_fade_transition(s7)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — 4 Zones Overview (2×2 Grid)
# ══════════════════════════════════════════════════════════════════════════════
s8 = prs.slides.add_slide(blank)
set_bg(s8, LIGHT_GRAY)
add_title(s8, "Zone Architecture: 4 Strategic Monitoring Locations", top=Inches(0.22))
add_teal_underline(s8, Inches(0.85))

zones_data = [
    ("Zone 1: Coastal Port Area",
     RGBColor(0xff, 0xf0, 0xe6), ORANGE,
     ["Router: Zone1-Router (Cisco 2901) | 10.10.10.1",
      "Switch: Zone1-Switch (Catalyst 2960-X)",
      "AP: Zone1-AP (Cisco Aironet 1840)",
      "Sensors (5): Water Level, Tidal Gauge,",
      "  Rain Gauge, Salinity, Current Meter",
      "VLAN: 10 (Sensors), 70 (Wireless)",
      "Coverage: Port district, coastal belt"]),
    ("Zone 2: Riverine Lowlands",
     RGBColor(0xe6, 0xf7, 0xff), TEAL,
     ["Router: Zone2-Router (Cisco 2901) | 10.10.20.1",
      "Switch: Zone2-Switch (Catalyst 2960-X)",
      "AP: Zone2-AP (Cisco Aironet 1840)",
      "Sensors (5): Water Level, Flow Velocity,",
      "  Turbidity, Soil Moisture, Rain Gauge",
      "VLAN: 20 (Sensors), 80 (Wireless)",
      "Coverage: Karnaphuli floodplains"]),
    ("Zone 3: Hilly Upstream",
     RGBColor(0xf0, 0xe6, 0xff), PURPLE,
     ["Router: Zone3-Router (Cisco 2901) | 10.10.30.1",
      "Switch: Zone3-Switch (Catalyst 2960-X)",
      "AP: Zone3-AP (Cisco Aironet 1840)",
      "Sensors (5): Water Level, Flow Rate,",
      "  Landslide, Rain Gauge, Temperature",
      "VLAN: 30 (Sensors), 90 (Wireless)",
      "Coverage: Sitakund hills, upstream catchment"]),
    ("Zone 4: Urban Lowland",
     RGBColor(0xe6, 0xff, 0xe6), GREEN_ACCENT,
     ["Router: Zone4-Router (Cisco 2901) | 10.10.40.1",
      "Switch: Zone4-Switch (Catalyst 2960-X)",
      "AP: Zone4-AP (Cisco Aironet 1840)",
      "Sensors (5): Street Flood, Stormwater,",
      "  Rain Gauge, Humidity, Water Quality",
      "VLAN: 40 (Sensors), 100 (Wireless)",
      "Coverage: GEC, Agrabad, Nasirabad"]),
]

positions = [(0, 0), (1, 0), (0, 1), (1, 1)]
qw = Inches(4.75); qh = Inches(1.95)
for (col, row), (zname, zbg, zborder, zlines) in zip(positions, zones_data):
    ql = Inches(0.2 + col * 4.85)
    qt = Inches(1.0 + row * 2.05)
    add_rect(s8, ql, qt, qw, qh, zbg, zborder, 1.8)
    # Title bar
    add_rect(s8, ql, qt, qw, Inches(0.32), zborder)
    tb_zt = s8.shapes.add_textbox(ql + Inches(0.08), qt + Inches(0.03),
                                   qw - Inches(0.1), Inches(0.28))
    tf_zt = tb_zt.text_frame; p_zt = tf_zt.paragraphs[0]
    r_zt = p_zt.add_run(); r_zt.text = zname
    r_zt.font.size = Pt(9); r_zt.font.bold = True
    r_zt.font.color.rgb = WHITE; r_zt.font.name = 'Calibri'
    # Content
    tb_zc = s8.shapes.add_textbox(ql + Inches(0.08), qt + Inches(0.36),
                                   qw - Inches(0.1), Inches(1.52))
    tf_zc = tb_zc.text_frame; tf_zc.word_wrap = True
    for j, zl in enumerate(zlines):
        p2 = tf_zc.paragraphs[0] if j == 0 else tf_zc.add_paragraph()
        r2 = p2.add_run(); r2.text = zl
        r2.font.size = Pt(7.5); r2.font.color.rgb = DARK_SLATE; r2.font.name = 'Calibri'

# Legend strip
legend_strip = add_rect(s8, Inches(0.2), Inches(5.08), Inches(9.6), Inches(0.28),
                        RGBColor(0xe0, 0xe0, 0xe0))
legend_items_txt = [
    ("■ Water Level", TEAL), ("■ Rain Gauge", PRIMARY_BLUE),
    ("■ Tidal/Flow", ORANGE), ("■ Soil/Landslide", PURPLE),
    ("■ AP Wireless", GREEN_ACCENT)
]
leg_l = Inches(0.4)
for txt, col in legend_items_txt:
    tb_lg = s8.shapes.add_textbox(leg_l, Inches(5.1), Inches(1.8), Inches(0.25))
    tf_lg = tb_lg.text_frame
    r_lg = tf_lg.paragraphs[0].add_run(); r_lg.text = txt
    r_lg.font.size = Pt(7.5); r_lg.font.color.rgb = col
    leg_l += Inches(1.85)

add_footer(s8, 8, W, H)
add_fade_transition(s8)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — IoT Devices & Sensor Integration
# ══════════════════════════════════════════════════════════════════════════════
s9 = prs.slides.add_slide(blank)
set_bg(s9, WHITE)
add_title(s9, "IoT Sensor Deployment & Device Specifications", top=Inches(0.25))
add_teal_underline(s9, Inches(0.88))

sensor_specs = [
    ("Water Level Sensor",    "Ultrasonic + Pressure", "±2mm accuracy", "MQTT/TCP, 30s interval", "10.10.x.10-30"),
    ("Rain Gauge",            "Tipping bucket 0.2mm",  "±5% accuracy",  "MQTT/UDP, 60s interval", "10.10.x.31-50"),
    ("Flow Velocity Meter",   "Acoustic Doppler",      "0.001–10 m/s",  "MQTT/TCP, 30s interval", "10.10.x.51-70"),
    ("Soil Moisture Sensor",  "Capacitive probe",      "±3% VWC",       "MQTT/UDP, 120s interval","10.10.x.71-90"),
    ("Tidal Gauge",           "Pressure transducer",   "±1cm accuracy", "MQTT/TCP, 15s interval", "10.10.x.91-110"),
    ("Water Quality Probe",   "Multi-parameter",       "pH/Turbid/Cond","MQTT/TCP, 60s interval", "10.10.x.111-130"),
]

tb9 = s9.shapes.add_textbox(Inches(0.3), Inches(1.05), Inches(5.8), Inches(4.2))
tf9 = tb9.text_frame; tf9.word_wrap = True
p9h = tf9.paragraphs[0]
rh9 = p9h.add_run()
rh9.text = "SENSOR SPECIFICATIONS:"
rh9.font.size = Pt(9.5); rh9.font.bold = True; rh9.font.color.rgb = TEAL

for sname, stype, saccuracy, sprotocol, sip in sensor_specs:
    p9 = tf9.add_paragraph()
    r9 = p9.add_run()
    r9.text = f"▶ {sname}"
    r9.font.size = Pt(9); r9.font.bold = True; r9.font.color.rgb = PRIMARY_BLUE

    for detail in [f"  Type: {stype}", f"  Accuracy: {saccuracy}",
                   f"  Protocol: {sprotocol}", f"  IP Range: {sip}"]:
        pd = tf9.add_paragraph()
        rd = pd.add_run(); rd.text = detail
        rd.font.size = Pt(7.5); rd.font.color.rgb = DARK_SLATE; rd.font.name = 'Calibri'

    tf9.add_paragraph()

# Device count + protocol stack
add_rect(s9, Inches(0.3), Inches(4.7), Inches(5.8), Inches(0.55),
         RGBColor(0xe8, 0xf4, 0xff), TEAL, 1)
tb9b = s9.shapes.add_textbox(Inches(0.4), Inches(4.72), Inches(5.6), Inches(0.5))
tf9b = tb9b.text_frame; p9b = tf9b.paragraphs[0]
r9b = p9b.add_run()
r9b.text = "Devices: 20 sensors × 4 zones = 80 total  |  Protocol: MQTT over Wi-Fi (802.11n)  |  QoS Level 2"
r9b.font.size = Pt(8); r9b.font.bold = True; r9b.font.color.rgb = PRIMARY_BLUE

# Right: map diagram
fig9, ax9 = plt.subplots(figsize=(3.8, 4.0))
fig9.patch.set_facecolor('white'); ax9.set_facecolor('#e8f4ff')
ax9.set_xlim(0, 4); ax9.set_ylim(0, 4.5)
ax9.set_title("Chittagong Sensor Deployment", fontsize=8, fontweight='bold', color='#1a3a6b')
ax9.set_xticks([]); ax9.set_yticks([])

# Simplified map: Bay of Bengal area
bay = Polygon([[2.5, 0], [4, 0], [4, 2.5], [3, 2]],
              closed=True, color='#a8d8ea', alpha=0.5)
ax9.add_patch(bay)
ax9.text(3.2, 0.8, "Bay of\nBengal", ha='center', fontsize=7, color='#1a3a6b', alpha=0.8)

river = Polygon([[1.5, 0.5], [2.5, 0.5], [3.0, 2.0], [2.5, 2.5], [2.0, 2.0], [1.0, 1.5]],
                closed=True, color='#90caf9', alpha=0.6)
ax9.add_patch(river)
ax9.text(1.9, 1.3, "Karnaphuli\nRiver", ha='center', fontsize=6.5, color='#1a3a6b')

zone_map = [
    ("Z1\nCoastal", 2.7, 1.2, '#ff8c00'),
    ("Z2\nRiverine", 1.8, 2.0, '#00b4d8'),
    ("Z3\nHilly",   0.7, 3.5, '#7c3aed'),
    ("Z4\nUrban",   1.5, 3.2, '#2ea04e'),
]
for zlbl9, zx9, zy9, zcol9 in zone_map:
    circ9 = Circle((zx9, zy9), 0.45, color=zcol9, alpha=0.2)
    ax9.add_patch(circ9)
    ax9.plot(zx9, zy9, 'o', color=zcol9, markersize=12, alpha=0.85)
    ax9.text(zx9, zy9, zlbl9, ha='center', va='center',
             fontsize=6, color='white', fontweight='bold')
    for s in range(5):
        angle9 = s * 72 * np.pi / 180
        sx9 = zx9 + 0.35 * np.cos(angle9)
        sy9 = zy9 + 0.35 * np.sin(angle9)
        ax9.plot(sx9, sy9, '^', color=zcol9, markersize=4)

leg9 = [mpatches.Patch(color=c, label=l) for l, _, _, c in zone_map]
ax9.legend(handles=leg9, fontsize=6, loc='lower left', framealpha=0.9)

buf9 = fig_to_buf(fig9)
s9.shapes.add_picture(buf9, Inches(6.2), Inches(0.95), Inches(3.55), Inches(3.85))

add_footer(s9, 9, W, H)
add_fade_transition(s9)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 10 — OSPF Routing & Security Architecture
# ══════════════════════════════════════════════════════════════════════════════
s10 = prs.slides.add_slide(blank)
set_bg(s10, LIGHT_GRAY)
add_title(s10, "Dynamic Routing (OSPF) and Security Implementation", top=Inches(0.22))
add_teal_underline(s10, Inches(0.85))

# OSPF diagram
fig10, ax10 = plt.subplots(figsize=(9.0, 2.3))
fig10.patch.set_facecolor('#f0f0f0'); ax10.set_facecolor('#f5f5f5')
ax10.set_xlim(0, 9); ax10.set_ylim(0, 2.3); ax10.axis('off')

# Center hexagon (backbone)
from matplotlib.patches import RegularPolygon
hex10 = RegularPolygon((4.5, 1.15), numVertices=6, radius=0.7,
                       orientation=0, facecolor='#1a3a6b', edgecolor='#00b4d8', linewidth=2)
ax10.add_patch(hex10)
ax10.text(4.5, 1.22, "OSPF Area 0", ha='center', va='center', fontsize=7.5,
          color='white', fontweight='bold')
ax10.text(4.5, 0.92, "(Backbone)", ha='center', va='center', fontsize=6.5, color='#00d4ff')

ospf_areas = [
    ("Area 1\nZone 1", 1.2, 1.9, '#ff8c00', "10.10.10.0/24"),
    ("Area 2\nZone 2", 7.8, 1.9, '#00b4d8', "10.10.20.0/24"),
    ("Area 3\nZone 3", 1.2, 0.4, '#7c3aed', "10.10.30.0/24"),
    ("Area 4\nZone 4", 7.8, 0.4, '#2ea04e', "10.10.40.0/24"),
]
for (albl, ax_, ay_, acol, anet) in ospf_areas:
    arect = FancyBboxPatch((ax_-0.7, ay_-0.28), 1.4, 0.56,
                           boxstyle="round,pad=0.05", facecolor=acol, alpha=0.85,
                           edgecolor='white', linewidth=0.8)
    ax10.add_patch(arect)
    ax10.text(ax_, ay_+0.06, albl, ha='center', va='center',
              fontsize=6.5, color='white', fontweight='bold')
    ax10.text(ax_, ay_-0.16, anet, ha='center', va='center', fontsize=5.5, color='white')
    dx = 4.5 - ax_; dy = 1.15 - ay_
    length = (dx**2 + dy**2)**0.5
    ax10.annotate('', xy=(ax_ + dx/length*0.72, ay_ + dy/length*0.28),
                  xytext=(ax_, ay_),
                  arrowprops=dict(arrowstyle='->', color=acol, lw=1.3))

# ISP
isp10 = FancyBboxPatch((4.0, 1.95), 1.0, 0.32,
                       boxstyle="round,pad=0.04", facecolor='#e65c00',
                       edgecolor='white', linewidth=0.8)
ax10.add_patch(isp10)
ax10.text(4.5, 2.11, "ISP / Internet", ha='center', va='center',
          fontsize=6.5, color='white', fontweight='bold')
ax10.annotate('', xy=(4.5, 1.85), xytext=(4.5, 1.97),
             arrowprops=dict(arrowstyle='<->', color='#e65c00', lw=1.3))

ax10.text(4.5, 0.05, "OSPF Hello: 10s | Dead: 40s | Process-ID: 1 | Router-ID auto | SPF runs on topology change",
          ha='center', va='center', fontsize=6, color='#555555', style='italic')

buf10 = fig_to_buf(fig10)
s10.shapes.add_picture(buf10, Inches(0.2), Inches(1.0), Inches(9.5), Inches(2.35))

# Lower half: three columns
lower_data = [
    ("NAT Configuration", ORANGE, [
        "ip nat inside source list 1",
        "  interface GigEth0/0 overload",
        "access-list 1 permit",
        "  10.0.0.0 0.255.255.255",
        "• PAT (overload) for all zones",
        "• Inside: all VLANs",
        "• Outside: ISP interface",
        "• Static NAT: Web server",
        "  10.0.100.3 → 172.16.0.10",
    ]),
    ("ACL Security Rules", TEAL, [
        "ACL 100 – Inbound ISP:",
        "  permit tcp any host 172.16.0.10 eq 80",
        "  permit tcp any host 172.16.0.10 eq 443",
        "  deny   ip any 10.0.50.0 0.0.0.255",
        "  deny   ip any 10.0.110.0 0.0.0.255",
        "ACL 101 – Zone isolation:",
        "  permit tcp 10.10.x.0 host 10.0.100.x",
        "  deny   ip 10.10.x.0 10.10.y.0",
        "  permit ip any 10.0.100.0 0.0.0.255",
    ]),
    ("Security Zones", PURPLE, [
        "DMZ: Web + Email servers",
        "  → Accessible from internet",
        "Internal: DHCP + DNS + MQTT",
        "  → No direct internet",
        "Management: VLAN 50",
        "  → Admin access only",
        "Sensor VLANs 10–40:",
        "  → MQTT only to broker",
        "  → No cross-zone traffic",
    ]),
]

col_w = Inches(3.05)
for ci, (ctitle, ccol, clines) in enumerate(lower_data):
    cl = Inches(0.3 + ci * 3.2)
    ct = Inches(3.42)
    add_rect(s10, cl, ct, col_w, Inches(1.88), RGBColor(0xff, 0xff, 0xff), ccol, 1)
    # col header
    add_rect(s10, cl, ct, col_w, Inches(0.28), ccol)
    tb_ch = s10.shapes.add_textbox(cl + Inches(0.05), ct + Inches(0.03),
                                    col_w - Inches(0.1), Inches(0.24))
    p_ch = tb_ch.text_frame.paragraphs[0]
    r_ch = p_ch.add_run(); r_ch.text = ctitle
    r_ch.font.size = Pt(8.5); r_ch.font.bold = True
    r_ch.font.color.rgb = WHITE

    tb_cl = s10.shapes.add_textbox(cl + Inches(0.08), ct + Inches(0.32),
                                    col_w - Inches(0.12), Inches(1.5))
    tf_cl = tb_cl.text_frame; tf_cl.word_wrap = True
    for j, ln in enumerate(clines):
        p2 = tf_cl.paragraphs[0] if j == 0 else tf_cl.add_paragraph()
        r2 = p2.add_run(); r2.text = ln
        r2.font.size = Pt(7); r2.font.color.rgb = DARK_SLATE
        r2.font.name = 'Courier New' if 'permit' in ln or 'deny' in ln or 'ip nat' in ln else 'Calibri'

add_footer(s10, 10, W, H)
add_fade_transition(s10)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 11 — Key Technical Highlights
# ══════════════════════════════════════════════════════════════════════════════
s11 = prs.slides.add_slide(blank)
set_bg(s11, WHITE)
add_title(s11, "Implementation Highlights: What We Configured", top=Inches(0.25))
add_teal_underline(s11, Inches(0.88))

highlight_cards = [
    ("VLAN Segmentation", RGBColor(0xff, 0xf0, 0xe0), ORANGE,
     ["12 VLANs configured on all switches",
      "802.1Q trunk links between all switches and routers",
      "Inter-VLAN routing via Layer-3 switch",
      "VLAN 50 (Mgmt) isolated with ACL",
      "Native VLAN 1 – untagged frames only"]),
    ("DHCP Services", RGBColor(0xe0, 0xf7, 0xff), TEAL,
     ["4 DHCP pools: one per zone VLAN",
      "Excluded addresses: .1–.9 per pool",
      "DNS option: 10.0.100.2 pushed",
      "Default gateway: zone router IP",
      "Lease: 24h; tested across VLANs"]),
    ("DNS & Web Monitoring", RGBColor(0xf3, 0xec, 0xff), PURPLE,
     ["BIND DNS: hydronet.local zone",
      "A records for all servers resolved",
      "Apache web: monitor.hydronet.local",
      "Real-time sensor dashboard page",
      "HTTP tested from all zone hosts"]),
    ("Email & Alert System", RGBColor(0xe6, 0xff, 0xe6), GREEN_ACCENT,
     ["SMTP server: port 25 configured",
      "POP3: port 110 for retrieval",
      "Alert emails on flood threshold",
      "Recipients: admin + zone officers",
      "Tested: email received in PT sim"]),
    ("OSPF Dynamic Routing", RGBColor(0xf0, 0xf0, 0xf0), DARK_SLATE,
     ["Single Area 0 (backbone) design",
      "OSPF enabled on all routers",
      "Priority 255 on Core-Router-1 (DR)",
      "All zone subnets redistributed",
      "Convergence tested: <5 seconds"]),
    ("Security (NAT + ACL)", RGBColor(0xff, 0xee, 0xee), ORANGE_ALERT,
     ["PAT (overload) on ISP interface",
      "Static NAT for web server access",
      "ACL 100: filters ISP inbound",
      "ACL 101: zone-to-zone isolation",
      "Management VLAN: admin-only ACL"]),
]

card_w11 = Inches(2.98); card_h11 = Inches(1.8)
for i, (ctitle, cbg, cborder, clines) in enumerate(highlight_cards):
    col = i % 3; row = i // 3
    cl = Inches(0.3 + col * 3.15)
    ct = Inches(1.0 + row * 1.95)
    add_rect(s11, cl, ct, card_w11, card_h11, cbg, cborder, 1.5)
    add_rect(s11, cl, ct, card_w11, Inches(0.32), cborder)
    tb_h = s11.shapes.add_textbox(cl + Inches(0.08), ct + Inches(0.04),
                                   card_w11 - Inches(0.1), Inches(0.25))
    p_h = tb_h.text_frame.paragraphs[0]
    r_h = p_h.add_run(); r_h.text = ctitle
    r_h.font.size = Pt(9); r_h.font.bold = True
    r_h.font.color.rgb = WHITE; r_h.font.name = 'Calibri'

    tb_c = s11.shapes.add_textbox(cl + Inches(0.1), ct + Inches(0.38),
                                   card_w11 - Inches(0.15), Inches(1.35))
    tf_c = tb_c.text_frame; tf_c.word_wrap = True
    for j, ln in enumerate(clines):
        p2 = tf_c.paragraphs[0] if j == 0 else tf_c.add_paragraph()
        r2 = p2.add_run(); r2.text = "• " + ln
        r2.font.size = Pt(7.5); r2.font.color.rgb = DARK_SLATE; r2.font.name = 'Calibri'

# Bottom summary box
add_rect(s11, Inches(0.3), Inches(4.92), Inches(9.4), Inches(0.38),
         RGBColor(0xe0, 0xf7, 0xff), TEAL, 1)
add_textbox(s11,
    "Complete end-to-end network deployed in Cisco Packet Tracer 8.2.2 — "
    "all services operational, routing converged, security policies enforced, "
    "and sensor data flowing to the central dashboard.",
    Inches(0.4), Inches(4.95), Inches(9.2), Inches(0.33),
    font_size=8, italic=True, color=PRIMARY_BLUE, align=PP_ALIGN.CENTER)

add_footer(s11, 11, W, H)
add_fade_transition(s11)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 12 — Limitations & Future Directions
# ══════════════════════════════════════════════════════════════════════════════
s12 = prs.slides.add_slide(blank)
set_bg(s12, WHITE)
add_title(s12, "Limitations of Current Implementation & Future Roadmap", top=Inches(0.22))
add_teal_underline(s12, Inches(0.85))

limitations = [
    ("Simulation Environment",
     ["Cisco Packet Tracer lacks full IoT physics",
      "Real sensor latency not modeled",
      "No packet loss or interference simulation"]),
    ("Physical Layer",
     ["No real cable/wireless deployment",
      "Power supply and PoE not considered",
      "Environmental sensor enclosures absent"]),
    ("Security Depth",
     ["No VPN or encryption (HTTPS/TLS)",
      "No intrusion detection system (IDS)",
      "Password policies not fully enforced"]),
    ("Scalability",
     ["Current design supports 4 zones only",
      "Adding zones needs full redesign",
      "No load balancing on core routers"]),
    ("Redundancy",
     ["No HSRP/VRRP for gateway failover",
      "Single ISP uplink — no redundancy",
      "Spanning Tree not fully optimized"]),
    ("Data Analytics",
     ["No ML-based flood prediction",
      "No historical data storage (no DB)",
      "Alert thresholds are static values"]),
]

roadmap = [
    ("Phase 1 (0–6 mo): Deployment",
     ["Physical sensor procurement & install",
      "Fiber backbone for core links",
      "PoE switches for sensor power"]),
    ("Phase 2 (6–12 mo): Security",
     ["TLS encryption on MQTT",
      "VPN tunnels for remote zones",
      "Implement IDS/IPS (Snort)"]),
    ("Phase 3 (12–18 mo): Redundancy",
     ["HSRP on all zone gateways",
      "Dual ISP with BGP failover",
      "UPS & solar backup for sensors"]),
    ("Phase 4 (18–24 mo): Intelligence",
     ["ML flood prediction model",
      "Time-series DB (InfluxDB)",
      "Dynamic alert thresholds"]),
    ("Phase 5 (24–30 mo): Scale",
     ["Expand to 12 zones city-wide",
      "Mobile app for public alerts",
      "API for govt. disaster portal"]),
    ("Phase 6 (30–36 mo): Integration",
     ["Integrate Bangladesh Met. Dept.",
      "Cross-border data sharing",
      "ISO/IEC 27001 certification"]),
]

add_rect(s12, Inches(0.2), Inches(1.0), Inches(4.65), Inches(3.85),
         RGBColor(0xff, 0xf5, 0xf5), ORANGE_ALERT, 1.2)
add_rect(s12, Inches(0.2), Inches(1.0), Inches(4.65), Inches(0.3), ORANGE_ALERT)
tb_lh = s12.shapes.add_textbox(Inches(0.3), Inches(1.02), Inches(4.45), Inches(0.26))
p_lh = tb_lh.text_frame.paragraphs[0]
r_lh = p_lh.add_run(); r_lh.text = "⚠  Current Limitations"
r_lh.font.size = Pt(10); r_lh.font.bold = True; r_lh.font.color.rgb = WHITE

tb_lim = s12.shapes.add_textbox(Inches(0.3), Inches(1.35), Inches(4.45), Inches(3.45))
tf_lim = tb_lim.text_frame; tf_lim.word_wrap = True
first = True
for cat, items in limitations:
    p = tf_lim.paragraphs[0] if first else tf_lim.add_paragraph()
    first = False
    r = p.add_run(); r.text = cat
    r.font.size = Pt(8.5); r.font.bold = True; r.font.color.rgb = ORANGE_ALERT
    for item in items:
        pd = tf_lim.add_paragraph()
        rd = pd.add_run(); rd.text = "  • " + item
        rd.font.size = Pt(7.5); rd.font.color.rgb = DARK_SLATE

add_rect(s12, Inches(5.15), Inches(1.0), Inches(4.65), Inches(3.85),
         RGBColor(0xf0, 0xff, 0xf0), GREEN_ACCENT, 1.2)
add_rect(s12, Inches(5.15), Inches(1.0), Inches(4.65), Inches(0.3), GREEN_ACCENT)
tb_rh = s12.shapes.add_textbox(Inches(5.25), Inches(1.02), Inches(4.45), Inches(0.26))
p_rh = tb_rh.text_frame.paragraphs[0]
r_rh = p_rh.add_run(); r_rh.text = "🚀  6-Phase Roadmap (3 Years)"
r_rh.font.size = Pt(10); r_rh.font.bold = True; r_rh.font.color.rgb = WHITE

tb_rm = s12.shapes.add_textbox(Inches(5.25), Inches(1.35), Inches(4.45), Inches(3.45))
tf_rm = tb_rm.text_frame; tf_rm.word_wrap = True
first2 = True
for phase, items in roadmap:
    p = tf_rm.paragraphs[0] if first2 else tf_rm.add_paragraph()
    first2 = False
    r = p.add_run(); r.text = phase
    r.font.size = Pt(8.5); r.font.bold = True; r.font.color.rgb = GREEN_ACCENT
    for item in items:
        pd = tf_rm.add_paragraph()
        rd = pd.add_run(); rd.text = "  • " + item
        rd.font.size = Pt(7.5); rd.font.color.rgb = DARK_SLATE

# Impact projection
add_rect(s12, Inches(0.2), Inches(4.92), Inches(9.6), Inches(0.38),
         RGBColor(0xe8, 0xf4, 0xff), TEAL, 1)
add_textbox(s12,
    "Impact Projection: Full deployment expected to reduce flood casualties by 60–70%, "
    "provide 2–4 hour advance warning, and protect $300M+ in annual economic value.",
    Inches(0.3), Inches(4.95), Inches(9.4), Inches(0.33),
    font_size=8.5, italic=True, color=PRIMARY_BLUE, align=PP_ALIGN.CENTER)

add_footer(s12, 12, W, H)
add_fade_transition(s12)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 13 — Conclusion
# ══════════════════════════════════════════════════════════════════════════════
s13 = prs.slides.add_slide(blank)
set_bg(s13, LIGHT_GRAY)
add_title(s13, "Conclusion: Networking for Resilience", top=Inches(0.25))
add_teal_underline(s13, Inches(0.88))

# Left body
tb13 = s13.shapes.add_textbox(Inches(0.3), Inches(1.05), Inches(6.6), Inches(4.15))
tf13 = tb13.text_frame; tf13.word_wrap = True

conclusion_lines = [
    ("ACHIEVEMENT SUMMARY", True, TEAL),
    ("HydroNet successfully demonstrates a complete, production-ready network architecture for Chittagong's flood resilience challenge.", False, DARK_SLATE),
    ("", False, DARK_SLATE),
    ("KEY ACCOMPLISHMENTS:", True, PRIMARY_BLUE),
    ("✔  4-zone hierarchical network with 12 VLANs fully operational", False, DARK_SLATE),
    ("✔  20+ IoT sensors across Coastal, Riverine, Hilly, and Urban zones", False, DARK_SLATE),
    ("✔  OSPF dynamic routing with sub-5s convergence", False, DARK_SLATE),
    ("✔  Full server suite: DHCP, DNS, Web, Email, MQTT broker", False, DARK_SLATE),
    ("✔  NAT + ACL security with zone isolation", False, DARK_SLATE),
    ("✔  Real-time monitoring dashboard at monitor.hydronet.local", False, DARK_SLATE),
    ("✔  Automated flood alert emails to zone officers", False, DARK_SLATE),
    ("", False, DARK_SLATE),
    ("TECHNICAL EXCELLENCE:", True, PRIMARY_BLUE),
    ("✔  Industry-standard Cisco equipment (3945, 2901, 3750-X, 2960-X)", False, DARK_SLATE),
    ("✔  Designed to scale to 12 zones with phased roadmap", False, DARK_SLATE),
    ("✔  Fully verified in Cisco Packet Tracer 8.2.2", False, DARK_SLATE),
    ("", False, DARK_SLATE),
    ("IMPACT: Potential to save hundreds of lives annually and protect $300M+ in infrastructure.", False, PRIMARY_BLUE),
]

for i, (txt, bold, col) in enumerate(conclusion_lines):
    p = tf13.paragraphs[0] if i == 0 else tf13.add_paragraph()
    r = p.add_run(); r.text = txt
    r.font.size = Pt(8.5) if not bold else Pt(9)
    r.font.bold = bold; r.font.color.rgb = col; r.font.name = 'Calibri'

# Right accent: matplotlib circle with mission text
fig13, ax13 = plt.subplots(figsize=(2.9, 3.6))
fig13.patch.set_facecolor('#f0f0f0'); ax13.set_facecolor('#f0f0f0')
ax13.set_xlim(-1.5, 1.5); ax13.set_ylim(-1.8, 1.8); ax13.axis('off')

for r_, alpha_, color_ in [(1.4, 0.15, '#1a3a6b'), (1.2, 0.2, '#00b4d8'), (1.0, 1.0, '#1a3a6b')]:
    c13 = Circle((0, 0.1), r_, color=color_, alpha=alpha_)
    ax13.add_patch(c13)

ax13.text(0, 0.75, "4 ZONES", ha='center', va='center',
          fontsize=9, color='white', fontweight='bold')
ax13.text(0, 0.42, "20+ SENSORS", ha='center', va='center',
          fontsize=8.5, color='#00d4ff', fontweight='bold')
ax13.text(0, 0.12, "5 SERVERS", ha='center', va='center',
          fontsize=8.5, color='white', fontweight='bold')
ax13.text(0, -0.2, "1 MISSION:", ha='center', va='center',
          fontsize=8, color='#ffcc00', fontweight='bold')
ax13.text(0, -0.52, "SAVE LIVES", ha='center', va='center',
          fontsize=10, color='#ff8c00', fontweight='bold')

ax13.text(0, -1.3, '"Resilience through\nconnectivity"',
          ha='center', va='center', fontsize=7.5,
          color='#1a3a6b', style='italic')

buf13 = fig_to_buf(fig13)
s13.shapes.add_picture(buf13, Inches(7.0), Inches(0.95), Inches(2.8), Inches(3.6))

add_footer(s13, 13, W, H)
add_fade_transition(s13)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 14 — Thank You & Team Credits
# ══════════════════════════════════════════════════════════════════════════════
s14 = prs.slides.add_slide(blank)
set_bg(s14, NAVY)

# Split background: left NAVY, right TEAL_BRIGHT
add_rect(s14, 0, 0, Inches(5.3), H, NAVY)
add_rect(s14, Inches(5.3), 0, Inches(4.7), H, TEAL_BRIGHT)

# Left side — Thank you
tb_ty = s14.shapes.add_textbox(Inches(0.3), Inches(0.55), Inches(4.8), Inches(0.9))
tf_ty = tb_ty.text_frame; p_ty = tf_ty.paragraphs[0]
p_ty.alignment = PP_ALIGN.LEFT
r_ty = p_ty.add_run(); r_ty.text = "THANK YOU"
r_ty.font.size = Pt(44); r_ty.font.bold = True
r_ty.font.color.rgb = WHITE; r_ty.font.name = 'Calibri'

tb_fl = s14.shapes.add_textbox(Inches(0.3), Inches(1.55), Inches(4.8), Inches(0.5))
tf_fl = tb_fl.text_frame; p_fl = tf_fl.paragraphs[0]
r_fl = p_fl.add_run(); r_fl.text = "for Listening"
r_fl.font.size = Pt(24); r_fl.font.italic = True
r_fl.font.color.rgb = TEAL; r_fl.font.name = 'Calibri'

tb_q = s14.shapes.add_textbox(Inches(0.3), Inches(2.3), Inches(4.8), Inches(0.5))
tf_q = tb_q.text_frame; p_q = tf_q.paragraphs[0]
r_q = p_q.add_run(); r_q.text = "Questions & Discussion Welcome"
r_q.font.size = Pt(11); r_q.font.color.rgb = RGBColor(0xcc, 0xcc, 0xcc)

# White bars
for bi, bw in enumerate([Inches(3.8), Inches(2.9), Inches(2.0)]):
    bar14 = add_rect(s14, Inches(0.3), Inches(3.15 + bi * 0.28), bw, Inches(0.12),
                     WHITE)

# Right side — Team info
tb_team = s14.shapes.add_textbox(Inches(5.5), Inches(0.4), Inches(4.2), Inches(0.42))
tf_team = tb_team.text_frame; p_team = tf_team.paragraphs[0]
r_team = p_team.add_run(); r_team.text = "TEAM HydroNet Trio"
r_team.font.size = Pt(17); r_team.font.bold = True
r_team.font.color.rgb = NAVY; r_team.font.name = 'Calibri'

members = [
    ("Abu Md. Selim",    "ID: 0242210005048", "Team Lead / Network Architect"),
    ("Arifur Rahman",    "ID: 0242210005051", "IoT Integration & Routing"),
    ("Sadab Abdullah",   "ID: 0242210005055", "Security & Server Config"),
]
for mi, (mname, mid, mrole) in enumerate(members):
    mb_l = Inches(5.5); mb_t = Inches(0.97 + mi * 0.88)
    add_rect(s14, mb_l, mb_t, Inches(4.2), Inches(0.75),
             RGBColor(0xff, 0xff, 0xff), NAVY, 1)
    tb_mn = s14.shapes.add_textbox(mb_l + Inches(0.1), mb_t + Inches(0.06),
                                    Inches(4.0), Inches(0.28))
    tf_mn = tb_mn.text_frame; p_mn = tf_mn.paragraphs[0]
    r_mn = p_mn.add_run(); r_mn.text = mname
    r_mn.font.size = Pt(11); r_mn.font.bold = True
    r_mn.font.color.rgb = PRIMARY_BLUE; r_mn.font.name = 'Calibri'

    tb_mid = s14.shapes.add_textbox(mb_l + Inches(0.1), mb_t + Inches(0.35),
                                     Inches(4.0), Inches(0.35))
    tf_mid = tb_mid.text_frame; tf_mid.word_wrap = True
    for j, mtxt in enumerate([mid, mrole]):
        p2 = tf_mid.paragraphs[0] if j == 0 else tf_mid.add_paragraph()
        r2 = p2.add_run(); r2.text = mtxt
        r2.font.size = Pt(8.5); r2.font.color.rgb = DARK_SLATE

# University info
tb_uni = s14.shapes.add_textbox(Inches(5.5), Inches(3.7), Inches(4.2), Inches(0.6))
tf_uni = tb_uni.text_frame; tf_uni.word_wrap = True
uni_lines = [
    ("Premier University, Chittagong", True),
    ("Dept. of Computer Science & Engineering", False),
    ("6th Semester — Network Architecture — March 2025", False),
]
for j, (ul, ub) in enumerate(uni_lines):
    p2 = tf_uni.paragraphs[0] if j == 0 else tf_uni.add_paragraph()
    r2 = p2.add_run(); r2.text = ul
    r2.font.size = Pt(8.5 if ub else 8); r2.font.bold = ub
    r2.font.color.rgb = NAVY; r2.font.name = 'Calibri'

# Bottom full-width bar
add_rect(s14, 0, H - Inches(0.55), W, Inches(0.55), PRIMARY_BLUE)
tb_bot = s14.shapes.add_textbox(Inches(0.2), H - Inches(0.52), W - Inches(0.4), Inches(0.42))
tf_bot = tb_bot.text_frame; p_bot = tf_bot.paragraphs[0]
p_bot.alignment = PP_ALIGN.CENTER
r_bot = p_bot.add_run()
r_bot.text = ("monitor.hydronet.local  |  admin@hydronet.local  |  "
              "Premier University, Chittagong  |  HydroNet 2025  |  "
              "Cisco Packet Tracer 8.2.2")
r_bot.font.size = Pt(7.5); r_bot.font.color.rgb = WHITE; r_bot.font.name = 'Calibri'

add_footer(s14, 14, W, H)
add_fade_transition(s14)

# ── Save ───────────────────────────────────────────────────────────────────────
prs.save("HydroNet_Chittagong_Resilience_Final.pptx")
print("Generated: HydroNet_Chittagong_Resilience_Final.pptx (14 slides)")
