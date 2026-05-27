"""Generate a 5-slide academic presentation summarizing the GT2 midterm exam."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ---------- Theme palette (Academic / Navy) ----------
NAVY = RGBColor(0x1F, 0x38, 0x64)
NAVY_DARK = RGBColor(0x14, 0x26, 0x45)
GOLD = RGBColor(0xE8, 0xB3, 0x3E)
LIGHT_BG = RGBColor(0xF5, 0xF7, 0xFA)
GREY = RGBColor(0x6B, 0x72, 0x80)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1F, 0x29, 0x37)
ACCENT_BLUE = RGBColor(0x3B, 0x82, 0xF6)
GREEN = RGBColor(0x10, 0xB9, 0x81)
RED = RGBColor(0xEF, 0x44, 0x44)

# 16:9 widescreen
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height

BLANK = prs.slide_layouts[6]


def set_solid_fill(shape, rgb):
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb
    shape.line.fill.background()


def add_rect(slide, x, y, w, h, rgb, line=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb
    if line:
        shp.line.color.rgb = rgb
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_text(slide, x, y, w, h, text, *, size=18, bold=False, color=DARK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font="Calibri",
             italic=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return tb


def set_slide_bg(slide, rgb):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = rgb


def add_footer(slide, page_num, total=5):
    # Bottom strip
    add_rect(slide, Emu(0), SH - Inches(0.35), SW, Inches(0.35), NAVY_DARK)
    add_text(slide, Inches(0.4), SH - Inches(0.33), Inches(8), Inches(0.3),
             "Giải tích 2 — MT1005 | HK242 2024-2025",
             size=10, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, SW - Inches(1.4), SH - Inches(0.33), Inches(1), Inches(0.3),
             f"{page_num} / {total}",
             size=10, color=GOLD, align=PP_ALIGN.RIGHT, bold=True,
             anchor=MSO_ANCHOR.MIDDLE)


# ============================================================
# SLIDE 1 — TITLE
# ============================================================
s1 = prs.slides.add_slide(BLANK)
set_slide_bg(s1, NAVY)

# Decorative diagonal stripe
stripe = s1.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM,
                             Inches(-2), Inches(5.2),
                             Inches(18), Inches(2))
set_solid_fill(stripe, NAVY_DARK)

# Gold accent bar
add_rect(s1, Inches(1.2), Inches(2.3), Inches(0.15), Inches(0.6), GOLD)

# Subtitle (eyebrow)
add_text(s1, Inches(1.5), Inches(2.25), Inches(8), Inches(0.5),
         "ĐẠI HỌC BÁCH KHOA — ĐHQG TP.HCM",
         size=14, bold=True, color=GOLD, font="Calibri")

# Main title
add_text(s1, Inches(1.2), Inches(2.9), Inches(11), Inches(1.4),
         "GIẢI TÍCH 2",
         size=72, bold=True, color=WHITE, font="Calibri")

# Sub
add_text(s1, Inches(1.2), Inches(4.1), Inches(11), Inches(0.8),
         "Phân tích đề thi giữa kỳ — Mã đề 2425",
         size=28, color=WHITE, font="Calibri")

# Meta info bottom
add_text(s1, Inches(1.2), Inches(5.6), Inches(11), Inches(0.5),
         "MT1005   •   HK242 (2024–2025)   •   12/04/2025   •   50 phút   •   16 câu trắc nghiệm",
         size=14, color=GOLD, font="Calibri")

# Mathematical decoration: integral symbol on right
deco = add_text(s1, Inches(10.5), Inches(0.6), Inches(2.5), Inches(2),
                "∬", size=200, color=RGBColor(0x2A, 0x4A, 0x80),
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                font="Cambria Math")

# ============================================================
# SLIDE 2 — THÔNG TIN ĐỀ THI
# ============================================================
s2 = prs.slides.add_slide(BLANK)
set_slide_bg(s2, LIGHT_BG)

# Header bar
add_rect(s2, Emu(0), Emu(0), SW, Inches(0.9), NAVY)
add_rect(s2, Inches(0.4), Inches(0.25), Inches(0.08), Inches(0.4), GOLD)
add_text(s2, Inches(0.65), Inches(0.2), Inches(11), Inches(0.5),
         "01  •  Tổng quan đề thi",
         size=24, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_text(s2, SW - Inches(2.5), Inches(0.2), Inches(2), Inches(0.5),
         "Slide 02",
         size=12, color=GOLD, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# Section title
add_text(s2, Inches(0.65), Inches(1.15), Inches(12), Inches(0.6),
         "Cấu trúc đề & quy tắc chấm điểm",
         size=28, bold=True, color=NAVY)
add_text(s2, Inches(0.65), Inches(1.75), Inches(12), Inches(0.4),
         "16 câu trắc nghiệm  •  50 phút  •  thang điểm 10",
         size=14, color=GREY, italic=True)

# 4 stat cards
cards = [
    ("16", "câu hỏi", NAVY),
    ("50", "phút", ACCENT_BLUE),
    ("+0.625", "điểm / câu đúng", GREEN),
    ("−0.125", "điểm / câu sai", RED),
]
card_y = Inches(2.5)
card_w = Inches(2.85)
card_h = Inches(1.6)
gap = Inches(0.2)
start_x = Inches(0.65)
for i, (val, lbl, col) in enumerate(cards):
    x = start_x + i * (card_w + gap)
    # card body
    add_rect(s2, x, card_y, card_w, card_h, WHITE)
    # left color stripe
    add_rect(s2, x, card_y, Inches(0.12), card_h, col)
    add_text(s2, x + Inches(0.3), card_y + Inches(0.15), card_w - Inches(0.3),
             Inches(0.9), val, size=40, bold=True, color=col,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s2, x + Inches(0.3), card_y + Inches(1.0), card_w - Inches(0.3),
             Inches(0.5), lbl, size=14, color=GREY, anchor=MSO_ANCHOR.MIDDLE)

# Info row (giảng viên / ngày...)
info_y = Inches(4.4)
add_rect(s2, Inches(0.65), info_y, Inches(12.05), Inches(2.4), WHITE)
add_rect(s2, Inches(0.65), info_y, Inches(12.05), Inches(0.45), NAVY)
add_text(s2, Inches(0.85), info_y + Inches(0.05), Inches(11), Inches(0.4),
         "THÔNG TIN HÀNH CHÍNH", size=13, bold=True, color=WHITE,
         anchor=MSO_ANCHOR.MIDDLE)

rows = [
    ("Giảng viên ra đề", "Nguyễn Thị Xuân Anh"),
    ("Người phê duyệt", "Trần Ngọc Diễm"),
    ("Khoa", "Khoa học Ứng dụng"),
    ("Ngày thi", "12/04/2025"),
]
ry = info_y + Inches(0.6)
for i, (k, v) in enumerate(rows):
    col = i % 2
    row = i // 2
    rx = Inches(0.85) + Inches(5.85) * col
    rry = ry + Inches(0.85) * row
    add_text(s2, rx, rry, Inches(2.5), Inches(0.4),
             k.upper(), size=10, bold=True, color=GOLD)
    add_text(s2, rx, rry + Inches(0.3), Inches(5.5), Inches(0.5),
             v, size=16, bold=True, color=NAVY)

add_footer(s2, 2)

# ============================================================
# SLIDE 3 — PHÂN LOẠI CHỦ ĐỀ
# ============================================================
s3 = prs.slides.add_slide(BLANK)
set_slide_bg(s3, LIGHT_BG)

add_rect(s3, Emu(0), Emu(0), SW, Inches(0.9), NAVY)
add_rect(s3, Inches(0.4), Inches(0.25), Inches(0.08), Inches(0.4), GOLD)
add_text(s3, Inches(0.65), Inches(0.2), Inches(11), Inches(0.5),
         "02  •  Phân loại chủ đề",
         size=24, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_text(s3, SW - Inches(2.5), Inches(0.2), Inches(2), Inches(0.5),
         "Slide 03",
         size=12, color=GOLD, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

add_text(s3, Inches(0.65), Inches(1.15), Inches(12), Inches(0.6),
         "Các chủ đề được kiểm tra trong đề",
         size=28, bold=True, color=NAVY)
add_text(s3, Inches(0.65), Inches(1.75), Inches(12), Inches(0.4),
         "16 câu được phân thành 6 nhóm chủ đề chính",
         size=14, color=GREY, italic=True)

# Topic cards (2 rows x 3 cols)
topics = [
    ("Hàm nhiều biến", "Câu 1, 2, 3", "Miền xác định, đường mức, tiếp tuyến", NAVY),
    ("Mặt bậc hai", "Câu 4", "Phân loại quadric surfaces", ACCENT_BLUE),
    ("Đạo hàm & vi phân", "Câu 5, 6, 7, 8", "Đạo hàm theo hướng, hàm ẩn, xấp xỉ tuyến tính", GREEN),
    ("Cực trị", "Câu 9, 10, 16", "Lagrange, cực trị địa phương", GOLD),
    ("Tích phân kép", "Câu 11, 12, 13, 15", "Miền phẳng, đổi biến, ứng dụng", RED),
    ("Tọa độ cực", "Câu 14", "Mô tả đường cong", RGBColor(0x8B, 0x5C, 0xF6)),
]
cw = Inches(4.0)
ch = Inches(2.05)
gx = Inches(0.15)
gy = Inches(0.2)
sx = Inches(0.65)
sy = Inches(2.4)
for idx, (name, qs, desc, col) in enumerate(topics):
    r = idx // 3
    c = idx % 3
    x = sx + c * (cw + gx)
    y = sy + r * (ch + gy)
    add_rect(s3, x, y, cw, ch, WHITE)
    # top stripe
    add_rect(s3, x, y, cw, Inches(0.08), col)
    # number badge
    badge = s3.shapes.add_shape(MSO_SHAPE.OVAL,
                                x + cw - Inches(0.85), y + Inches(0.25),
                                Inches(0.55), Inches(0.55))
    set_solid_fill(badge, col)
    tf = badge.text_frame
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    rn = p.add_run()
    rn.text = str(idx + 1)
    rn.font.size = Pt(20)
    rn.font.bold = True
    rn.font.color.rgb = WHITE

    add_text(s3, x + Inches(0.25), y + Inches(0.3), cw - Inches(1.1), Inches(0.5),
             name, size=18, bold=True, color=NAVY)
    add_text(s3, x + Inches(0.25), y + Inches(0.85), cw - Inches(0.4), Inches(0.4),
             qs, size=12, bold=True, color=col)
    add_text(s3, x + Inches(0.25), y + Inches(1.25), cw - Inches(0.4), Inches(0.7),
             desc, size=12, color=GREY)

add_footer(s3, 3)

# ============================================================
# SLIDE 4 — ĐÁP ÁN
# ============================================================
s4 = prs.slides.add_slide(BLANK)
set_slide_bg(s4, LIGHT_BG)

add_rect(s4, Emu(0), Emu(0), SW, Inches(0.9), NAVY)
add_rect(s4, Inches(0.4), Inches(0.25), Inches(0.08), Inches(0.4), GOLD)
add_text(s4, Inches(0.65), Inches(0.2), Inches(11), Inches(0.5),
         "03  •  Bảng đáp án",
         size=24, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_text(s4, SW - Inches(2.5), Inches(0.2), Inches(2), Inches(0.5),
         "Slide 04",
         size=12, color=GOLD, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

add_text(s4, Inches(0.65), Inches(1.15), Inches(12), Inches(0.6),
         "Đáp án 16 câu trắc nghiệm",
         size=28, bold=True, color=NAVY)
add_text(s4, Inches(0.65), Inches(1.75), Inches(12), Inches(0.4),
         "Tham chiếu: trang 5 của đề thi",
         size=14, color=GREY, italic=True)

# Answers grid 4 cols x 4 rows
answers = [
    ("1", "A"), ("2", "A"), ("3", "A"), ("4", "A"),
    ("5", "D"), ("6", "D"), ("7", "A"), ("8", "D"),
    ("9", "A"), ("10", "E"), ("11", "D"), ("12", "D"),
    ("13", "E"), ("14", "A"), ("15", "C"), ("16", "E"),
]
gw = Inches(2.0)
gh = Inches(1.05)
ggx = Inches(0.18)
ggy = Inches(0.18)
gsx = Inches(1.55)
gsy = Inches(2.5)
for i, (q, a) in enumerate(answers):
    rr = i // 4
    cc = i % 4
    x = gsx + cc * (gw + ggx)
    y = gsy + rr * (gh + ggy)
    add_rect(s4, x, y, gw, gh, WHITE)
    add_rect(s4, x, y, Inches(0.5), gh, NAVY)
    add_text(s4, x, y, Inches(0.5), gh,
             f"Q{q}", size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s4, x + Inches(0.5), y, gw - Inches(0.5), gh,
             a, size=36, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Score note bottom
note_y = Inches(7.05) - Inches(0.35) - Inches(0.55)  # above footer
add_rect(s4, Inches(0.65), Inches(6.45), Inches(12.05), Inches(0.55),
         NAVY_DARK)
add_text(s4, Inches(0.85), Inches(6.45), Inches(12), Inches(0.55),
         "Tip: Trả lời đúng ≥ 13/16 câu  →  ≥ 8.0 điểm.  Trả lời mò bị trừ điểm — bỏ qua câu không chắc!",
         size=12, color=GOLD, italic=True, anchor=MSO_ANCHOR.MIDDLE)

add_footer(s4, 4)

# ============================================================
# SLIDE 5 — CÔNG THỨC TRỌNG TÂM
# ============================================================
s5 = prs.slides.add_slide(BLANK)
set_slide_bg(s5, LIGHT_BG)

add_rect(s5, Emu(0), Emu(0), SW, Inches(0.9), NAVY)
add_rect(s5, Inches(0.4), Inches(0.25), Inches(0.08), Inches(0.4), GOLD)
add_text(s5, Inches(0.65), Inches(0.2), Inches(11), Inches(0.5),
         "04  •  Công thức trọng tâm",
         size=24, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_text(s5, SW - Inches(2.5), Inches(0.2), Inches(2), Inches(0.5),
         "Slide 05",
         size=12, color=GOLD, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

add_text(s5, Inches(0.65), Inches(1.15), Inches(12), Inches(0.6),
         "Tổng kết — những công thức cần nhớ",
         size=28, bold=True, color=NAVY)

# Two columns of formula cards
formulas = [
    ("Đạo hàm theo hướng",
     "D_u f = ∇f · û",
     "với û là vector đơn vị theo hướng u"),
    ("Xấp xỉ tuyến tính",
     "f(x,y) ≈ f(a,b) + f_x·Δx + f_y·Δy",
     "tại điểm (a,b) đã biết"),
    ("Đạo hàm hàm ẩn",
     "∂z/∂x = − F_x / F_z",
     "với F(x,y,z) = 0"),
    ("Nhân tử Lagrange",
     "∇f = λ·∇g    ,    g(x,y) = 0",
     "tìm cực trị có ràng buộc"),
    ("Tích phân kép — tọa độ cực",
     "∬ f dA = ∫∫ f(r,φ)·r dr dφ",
     "Jacobian = r"),
    ("Khối lượng tấm phẳng",
     "m = ∬_D ρ(x,y) dA",
     "ρ là mật độ khối lượng"),
]
fw = Inches(6.0)
fh = Inches(1.55)
fgx = Inches(0.2)
fgy = Inches(0.18)
fsx = Inches(0.65)
fsy = Inches(2.05)
for i, (title, formula, note) in enumerate(formulas):
    r = i // 2
    c = i % 2
    x = fsx + c * (fw + fgx)
    y = fsy + r * (fh + fgy)
    add_rect(s5, x, y, fw, fh, WHITE)
    add_rect(s5, x, y, Inches(0.1), fh, GOLD)
    add_text(s5, x + Inches(0.3), y + Inches(0.12), fw - Inches(0.5), Inches(0.4),
             title, size=14, bold=True, color=NAVY)
    add_text(s5, x + Inches(0.3), y + Inches(0.55), fw - Inches(0.5), Inches(0.55),
             formula, size=18, bold=True, color=DARK,
             font="Cambria Math")
    add_text(s5, x + Inches(0.3), y + Inches(1.12), fw - Inches(0.5), Inches(0.4),
             note, size=11, italic=True, color=GREY)

add_footer(s5, 5)

# Save
out = "/projects/sandbox/Siu/docs/exam_2425/slides_GT2_summary.pptx"
prs.save(out)
print(f"Saved: {out}")
