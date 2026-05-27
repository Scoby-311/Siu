"""Render a faithful preview PDF of the 5-slide GT2 deck (16:9, A4-ish landscape)."""
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# ---------- 16:9 page (landscape) ----------
W = 13.333 * inch
H = 7.5 * inch

# Colors
NAVY = HexColor("#1F3864")
NAVY_DARK = HexColor("#142645")
GOLD = HexColor("#E8B33E")
LIGHT_BG = HexColor("#F5F7FA")
GREY = HexColor("#6B7280")
WHITE = HexColor("#FFFFFF")
DARK = HexColor("#1F2937")
ACCENT = HexColor("#3B82F6")
GREEN = HexColor("#10B981")
RED = HexColor("#EF4444")
PURPLE = HexColor("#8B5CF6")

FONT = "Helvetica"
FONT_B = "Helvetica-Bold"
FONT_I = "Helvetica-Oblique"

# Register Noto Sans (supports Vietnamese diacritics)
FONT_DIR = "/projects/sandbox/.fonts"
try:
    pdfmetrics.registerFont(TTFont("Noto", f"{FONT_DIR}/NotoSans-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Noto-Bold", f"{FONT_DIR}/NotoSans-Bold.ttf"))
    pdfmetrics.registerFont(TTFont("Noto-Italic", f"{FONT_DIR}/NotoSans-Italic.ttf"))
    FONT = "Noto"
    FONT_B = "Noto-Bold"
    FONT_I = "Noto-Italic"
except Exception as e:
    print("Font registration warning:", e)

print(f"Using fonts: {FONT}, {FONT_B}, {FONT_I}")


def rect(c, x, y, w, h, color, stroke=0):
    c.setFillColor(color)
    c.rect(x, H - y - h, w, h, stroke=stroke, fill=1)


def text(c, x, y, txt, size=14, bold=False, italic=False, color=DARK,
         align="left", w=None):
    fnt = FONT_B if bold else (FONT_I if italic else FONT)
    c.setFont(fnt, size)
    c.setFillColor(color)
    if align == "center" and w is not None:
        tw = c.stringWidth(txt, fnt, size)
        c.drawString(x + (w - tw) / 2, H - y - size * 0.85, txt)
    elif align == "right" and w is not None:
        tw = c.stringWidth(txt, fnt, size)
        c.drawString(x + w - tw, H - y - size * 0.85, txt)
    else:
        c.drawString(x, H - y - size * 0.85, txt)


def circle(c, cx, cy, r, color):
    c.setFillColor(color)
    c.circle(cx, H - cy, r, stroke=0, fill=1)


def parallelogram(c, x, y, w, h, skew, color):
    c.setFillColor(color)
    p = c.beginPath()
    p.moveTo(x + skew, H - y)
    p.lineTo(x + w, H - y)
    p.lineTo(x + w - skew, H - y - h)
    p.lineTo(x, H - y - h)
    p.close()
    c.drawPath(p, stroke=0, fill=1)


def header(c, num_str, page_str, title):
    rect(c, 0, 0, W, 0.9 * inch, NAVY)
    rect(c, 0.4 * inch, 0.25 * inch, 0.08 * inch, 0.4 * inch, GOLD)
    text(c, 0.65 * inch, 0.32 * inch, f"{num_str}  •  {title}",
         size=22, bold=True, color=WHITE)
    text(c, W - 1.5 * inch, 0.36 * inch, page_str, size=11, color=GOLD,
         align="right", w=1 * inch)


def footer(c, page, total=5):
    rect(c, 0, H - 0.35 * inch, W, 0.35 * inch, NAVY_DARK)
    text(c, 0.4 * inch, H - 0.27 * inch,
         "Giải tích 2 — MT1005 | HK242 2024-2025",
         size=9, color=WHITE)
    text(c, W - 1.4 * inch, H - 0.27 * inch, f"{page} / {total}",
         size=10, bold=True, color=GOLD, align="right", w=1 * inch)


# ============ slide 1 ============
def slide1(c):
    rect(c, 0, 0, W, H, NAVY)
    parallelogram(c, -2 * inch, 5.2 * inch, 18 * inch, 2 * inch, 1.5 * inch,
                  NAVY_DARK)
    # decorative integral
    c.setFont(FONT_B, 200)
    c.setFillColor(HexColor("#2A4A80"))
    c.drawString(10.7 * inch, H - 2.4 * inch, "∬")

    rect(c, 1.2 * inch, 2.3 * inch, 0.15 * inch, 0.6 * inch, GOLD)
    text(c, 1.5 * inch, 2.45 * inch,
         "ĐẠI HỌC BÁCH KHOA — ĐHQG TP.HCM",
         size=14, bold=True, color=GOLD)
    text(c, 1.2 * inch, 3.0 * inch, "GIẢI TÍCH 2",
         size=68, bold=True, color=WHITE)
    text(c, 1.2 * inch, 4.3 * inch,
         "Phân tích đề thi giữa kỳ — Mã đề 2425",
         size=26, color=WHITE)
    text(c, 1.2 * inch, 5.7 * inch,
         "MT1005   •   HK242 (2024–2025)   •   12/04/2025   •   50 phút   •   16 câu trắc nghiệm",
         size=13, color=GOLD)
    c.showPage()


# ============ slide 2 ============
def slide2(c):
    rect(c, 0, 0, W, H, LIGHT_BG)
    header(c, "01", "Slide 02", "Tổng quan đề thi")
    text(c, 0.65 * inch, 1.2 * inch, "Cấu trúc đề & quy tắc chấm điểm",
         size=26, bold=True, color=NAVY)
    text(c, 0.65 * inch, 1.85 * inch,
         "16 câu trắc nghiệm  •  50 phút  •  thang điểm 10",
         size=13, italic=True, color=GREY)

    cards = [
        ("16", "câu hỏi", NAVY),
        ("50", "phút", ACCENT),
        ("+0.625", "điểm / câu đúng", GREEN),
        ("−0.125", "điểm / câu sai", RED),
    ]
    cw = 2.85 * inch
    ch = 1.6 * inch
    sx = 0.65 * inch
    sy = 2.5 * inch
    gap = 0.2 * inch
    for i, (val, lbl, col) in enumerate(cards):
        x = sx + i * (cw + gap)
        rect(c, x, sy, cw, ch, WHITE)
        rect(c, x, sy, 0.12 * inch, ch, col)
        text(c, x + 0.3 * inch, sy + 0.4 * inch, val, size=36, bold=True, color=col)
        text(c, x + 0.3 * inch, sy + 1.2 * inch, lbl, size=13, color=GREY)

    # info box
    iy = 4.4 * inch
    rect(c, 0.65 * inch, iy, 12.05 * inch, 2.4 * inch, WHITE)
    rect(c, 0.65 * inch, iy, 12.05 * inch, 0.45 * inch, NAVY)
    text(c, 0.85 * inch, iy + 0.13 * inch, "THÔNG TIN HÀNH CHÍNH",
         size=12, bold=True, color=WHITE)

    rows = [
        ("GIẢNG VIÊN RA ĐỀ", "Nguyễn Thị Xuân Anh"),
        ("NGƯỜI PHÊ DUYỆT", "Trần Ngọc Diễm"),
        ("KHOA", "Khoa học Ứng dụng"),
        ("NGÀY THI", "12/04/2025"),
    ]
    ry = iy + 0.75 * inch
    for i, (k, v) in enumerate(rows):
        col = i % 2
        row = i // 2
        rx = 0.85 * inch + 5.85 * inch * col
        rry = ry + 0.85 * inch * row
        text(c, rx, rry, k, size=10, bold=True, color=GOLD)
        text(c, rx, rry + 0.32 * inch, v, size=15, bold=True, color=NAVY)

    footer(c, 2)
    c.showPage()


# ============ slide 3 ============
def slide3(c):
    rect(c, 0, 0, W, H, LIGHT_BG)
    header(c, "02", "Slide 03", "Phân loại chủ đề")
    text(c, 0.65 * inch, 1.2 * inch, "Các chủ đề được kiểm tra trong đề",
         size=26, bold=True, color=NAVY)
    text(c, 0.65 * inch, 1.85 * inch,
         "16 câu được phân thành 6 nhóm chủ đề chính",
         size=13, italic=True, color=GREY)

    topics = [
        ("Hàm nhiều biến", "Câu 1, 2, 3",
         "Miền xác định, đường mức, tiếp tuyến", NAVY),
        ("Mặt bậc hai", "Câu 4",
         "Phân loại quadric surfaces", ACCENT),
        ("Đạo hàm & vi phân", "Câu 5, 6, 7, 8",
         "Đạo hàm hướng, hàm ẩn, xấp xỉ tuyến tính", GREEN),
        ("Cực trị", "Câu 9, 10, 16",
         "Lagrange, cực trị địa phương", GOLD),
        ("Tích phân kép", "Câu 11, 12, 13, 15",
         "Miền phẳng, đổi biến, ứng dụng", RED),
        ("Tọa độ cực", "Câu 14",
         "Mô tả đường cong", PURPLE),
    ]
    cw = 4.0 * inch
    ch = 2.05 * inch
    gx = 0.15 * inch
    gy = 0.2 * inch
    sx = 0.65 * inch
    sy = 2.5 * inch
    for idx, (name, qs, desc, col) in enumerate(topics):
        r = idx // 3
        cc = idx % 3
        x = sx + cc * (cw + gx)
        y = sy + r * (ch + gy)
        rect(c, x, y, cw, ch, WHITE)
        rect(c, x, y, cw, 0.08 * inch, col)
        # number badge
        bx = x + cw - 0.6 * inch
        by = y + 0.5 * inch
        circle(c, bx, by, 0.27 * inch, col)
        text(c, bx - 0.27 * inch, by - 0.18 * inch, str(idx + 1),
             size=18, bold=True, color=WHITE,
             align="center", w=0.54 * inch)

        text(c, x + 0.25 * inch, y + 0.4 * inch, name,
             size=17, bold=True, color=NAVY)
        text(c, x + 0.25 * inch, y + 0.95 * inch, qs,
             size=12, bold=True, color=col)
        text(c, x + 0.25 * inch, y + 1.4 * inch, desc,
             size=11, color=GREY)

    footer(c, 3)
    c.showPage()


# ============ slide 4 ============
def slide4(c):
    rect(c, 0, 0, W, H, LIGHT_BG)
    header(c, "03", "Slide 04", "Bảng đáp án")
    text(c, 0.65 * inch, 1.2 * inch, "Đáp án 16 câu trắc nghiệm",
         size=26, bold=True, color=NAVY)
    text(c, 0.65 * inch, 1.85 * inch, "Tham chiếu: trang 5 của đề thi",
         size=13, italic=True, color=GREY)

    answers = [
        ("1", "A"), ("2", "A"), ("3", "A"), ("4", "A"),
        ("5", "D"), ("6", "D"), ("7", "A"), ("8", "D"),
        ("9", "A"), ("10", "E"), ("11", "D"), ("12", "D"),
        ("13", "E"), ("14", "A"), ("15", "C"), ("16", "E"),
    ]
    gw = 2.0 * inch
    gh = 1.05 * inch
    ggx = 0.18 * inch
    ggy = 0.18 * inch
    gsx = 1.55 * inch
    gsy = 2.5 * inch
    for i, (q, a) in enumerate(answers):
        rr = i // 4
        cc = i % 4
        x = gsx + cc * (gw + ggx)
        y = gsy + rr * (gh + ggy)
        rect(c, x, y, gw, gh, WHITE)
        rect(c, x, y, 0.5 * inch, gh, NAVY)
        text(c, x, y + 0.4 * inch, f"Q{q}", size=13, bold=True, color=WHITE,
             align="center", w=0.5 * inch)
        text(c, x + 0.5 * inch, y + 0.3 * inch, a, size=32, bold=True,
             color=GOLD, align="center", w=gw - 0.5 * inch)

    # tip box
    rect(c, 0.65 * inch, 6.45 * inch, 12.05 * inch, 0.55 * inch, NAVY_DARK)
    text(c, 0.85 * inch, 6.62 * inch,
         "Tip: Trả lời đúng ≥ 13/16 câu  →  ≥ 8.0 điểm.   "
         "Trả lời mò bị trừ điểm — bỏ qua câu không chắc!",
         size=11, italic=True, color=GOLD)

    footer(c, 4)
    c.showPage()


# ============ slide 5 ============
def slide5(c):
    rect(c, 0, 0, W, H, LIGHT_BG)
    header(c, "04", "Slide 05", "Công thức trọng tâm")
    text(c, 0.65 * inch, 1.2 * inch, "Tổng kết — những công thức cần nhớ",
         size=26, bold=True, color=NAVY)

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
    fw = 6.0 * inch
    fh = 1.55 * inch
    fgx = 0.2 * inch
    fgy = 0.18 * inch
    fsx = 0.65 * inch
    fsy = 2.05 * inch
    for i, (title, formula, note) in enumerate(formulas):
        r = i // 2
        cc = i % 2
        x = fsx + cc * (fw + fgx)
        y = fsy + r * (fh + fgy)
        rect(c, x, y, fw, fh, WHITE)
        rect(c, x, y, 0.1 * inch, fh, GOLD)
        text(c, x + 0.3 * inch, y + 0.2 * inch, title,
             size=13, bold=True, color=NAVY)
        text(c, x + 0.3 * inch, y + 0.7 * inch, formula,
             size=17, bold=True, color=DARK)
        text(c, x + 0.3 * inch, y + 1.25 * inch, note,
             size=10, italic=True, color=GREY)

    footer(c, 5)
    c.showPage()


# ============ build ============
out = "/projects/sandbox/Siu/docs/exam_2425/slides_GT2_summary_preview.pdf"
c = canvas.Canvas(out, pagesize=(W, H))
slide1(c)
slide2(c)
slide3(c)
slide4(c)
slide5(c)
c.save()
print(f"Saved: {out}")
