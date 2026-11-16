# -*- coding: utf-8 -*-
"""
OrcaX Office Suite - Ribbon Icon Generator
==========================================
Sinh bo icon PNG duotone thuong hieu (navy #1E2D82 + teal #0FA3B1) cho Ribbon.
Ve o canvas 128px (khong gian toa do 32 don vi, scale x4) roi xuong thu 32px & 16px.

Cach dung: python generate_icons.py            -> sinh toan bo PNG vao ./png/
           python generate_icons.py bolt hub   -> chi sinh 1 vai icon
"""
import os
import sys
from PIL import Image, ImageDraw

# ---- Brand palette ------------------------------------------------------
NAVY = (30, 45, 130, 255)        # #1E2D82 - mau thuong hieu chinh
NAVY_SOFT = (70, 90, 190, 255)   # sac nhat cua navy
TEAL = (15, 163, 177, 255)       # #0FA3B1 - mau nhan manh
TEAL_SOFT = (180, 230, 236, 255) # teal nhat (dien)
SLATE = (51, 65, 85, 255)        # #334155 - chi tiet dam
LIGHT = (226, 232, 240, 255)     # #E2E8F0 - dien sang
WHITE = (255, 255, 255, 255)
CLEAR = (0, 0, 0, 0)

SCALE = 4          # 32 units * 4 = 128px canvas
CANVAS = 32
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "png")


class Pen:
    """Ve trong khong gian 32 don vi, tu dong nhan scale."""

    def __init__(self, draw):
        self.d = draw

    def rrect(self, x, y, w, h, r=2, fill=None, outline=None, width=2):
        self.d.rounded_rectangle([x * SCALE, y * SCALE, (x + w) * SCALE, (y + h) * SCALE],
                                 radius=r * SCALE, fill=fill, outline=outline, width=int(width * SCALE))

    def rect(self, x, y, w, h, fill=None, outline=None, width=2):
        self.d.rectangle([x * SCALE, y * SCALE, (x + w) * SCALE, (y + h) * SCALE],
                         fill=fill, outline=outline, width=int(width * SCALE))

    def line(self, pts, color=NAVY, width=2):
        flat = []
        for c in pts:
            if isinstance(c, (tuple, list)):
                flat.extend(c)
            else:
                flat.append(c)
        scaled = [c * SCALE for c in flat]
        self.d.line(scaled, fill=color, width=int(width * SCALE))

    def poly(self, pts, fill=None, outline=None, width=2):
        scaled = [(x * SCALE, y * SCALE) for x, y in pts]
        self.d.polygon(scaled, fill=fill, outline=outline, width=int(width * SCALE))

    def circle(self, cx, cy, r, fill=None, outline=None, width=2):
        self.d.ellipse([(cx - r) * SCALE, (cy - r) * SCALE, (cx + r) * SCALE, (cy + r) * SCALE],
                       fill=fill, outline=outline, width=int(width * SCALE))

    def arc(self, cx, cy, r, a0, a1, color=NAVY, width=2):
        self.d.arc([(cx - r) * SCALE, (cy - r) * SCALE, (cx + r) * SCALE, (cy + r) * SCALE],
                   start=a0, end=a1, fill=color, width=int(width * SCALE))


# ---- Icon library (moi ham ve trong khong 32x32) ------------------------

def i_bolt(p):  # Chuan hoa 1-Click - tia chop
    p.rrect(4, 3, 24, 26, 3, fill=LIGHT)
    p.rrect(4, 3, 24, 26, 3, outline=NAVY, width=2)
    p.poly([(17, 6), (10, 17), (14, 17), (12, 26), (20, 14), (16, 14), (19, 6)], fill=TEAL)


def i_hub(p):  # Trung tam dieu hanh - dashboard
    p.rrect(3, 4, 26, 24, 3, fill=LIGHT, outline=NAVY, width=2)
    p.rect(6, 8, 8, 6, fill=TEAL)
    p.rect(17, 8, 9, 6, fill=NAVY_SOFT)
    p.rect(6, 17, 12, 8, fill=NAVY)
    p.rect(21, 17, 5, 8, fill=TEAL)


def i_search_doc(p):  # Phan tich van ban
    p.rrect(6, 3, 17, 24, 2, fill=LIGHT, outline=NAVY, width=2)
    p.line([9, 8, 20, 8], color=NAVY_SOFT, width=1.6)
    p.line([9, 12, 20, 12], color=NAVY_SOFT, width=1.6)
    p.line([9, 16, 16, 16], color=NAVY_SOFT, width=1.6)
    p.circle(21, 21, 6, outline=TEAL, width=2.6)
    p.line([25, 25, 29, 29], color=TEAL, width=3)


def i_undo(p):  # Khoi phuc
    p.arc(16, 17, 10, 200, 90, color=NAVY, width=2.6)
    p.poly([(8, 13), (4, 20), (13, 20)], fill=NAVY)
    p.line([16, 10, 22, 10], color=TEAL, width=2.6)


def i_compass(p):  # Quy chuan
    p.circle(16, 16, 12, fill=LIGHT, outline=NAVY, width=2.4)
    p.circle(16, 16, 8, outline=TEAL_SOFT, width=2)
    p.poly([(16, 9), (19, 16), (16, 23), (13, 16)], fill=TEAL)
    p.circle(16, 16, 2.2, fill=NAVY)


def i_layers(p):  # Cap bac
    p.poly([(16, 4), (28, 10), (16, 16), (4, 10)], fill=TEAL)
    p.poly([(4, 15), (16, 21), (28, 15), (28, 19), (16, 25), (4, 19)], fill=NAVY_SOFT)
    p.poly([(4, 23), (16, 29), (28, 23)], outline=NAVY, width=2)


def i_doc_type(p):  # Loai van ban
    p.rrect(7, 3, 18, 26, 2, fill=LIGHT, outline=NAVY, width=2)
    p.rect(10, 7, 12, 7, fill=TEAL)
    p.line([10, 17, 22, 17], color=NAVY, width=1.6)
    p.line([10, 20, 22, 20], color=NAVY_SOFT, width=1.6)
    p.line([10, 23, 18, 23], color=NAVY_SOFT, width=1.6)


def i_id_card(p):  # Thong tin van ban (metadata)
    p.rrect(3, 7, 26, 18, 2, fill=LIGHT, outline=NAVY, width=2)
    p.rect(6, 10, 8, 10, fill=TEAL)
    p.line([17, 11, 26, 11], color=NAVY, width=1.6)
    p.line([17, 15, 26, 15], color=NAVY_SOFT, width=1.6)
    p.line([17, 19, 23, 19], color=NAVY_SOFT, width=1.6)
    p.circle(10, 13.5, 2.4, fill=WHITE)


def i_heading(p):  # Khung tieu de
    p.rect(4, 5, 24, 8, fill=NAVY)
    p.line([4, 17, 28, 17], color=SLATE, width=1.6)
    p.line([4, 21, 22, 21], color=NAVY_SOFT, width=1.6)
    p.line([4, 25, 25, 25], color=NAVY_SOFT, width=1.6)
    p.rect(4, 5, 24, 8, outline=NAVY, width=0.5)


def i_text_center(p):  # Ten loai & trich yeu
    p.line([8, 6, 24, 6], color=NAVY, width=2.6)
    p.line([10, 11, 22, 11], color=TEAL, width=2.2)
    p.line([6, 17, 26, 17], color=SLATE, width=1.6)
    p.line([8, 21, 24, 21], color=NAVY_SOFT, width=1.6)
    p.line([11, 25, 21, 25], color=NAVY_SOFT, width=1.6)


def i_justify(p):  # Than bai can deu
    p.line([4, 6, 28, 6], color=NAVY, width=2)
    p.line([4, 11, 28, 11], color=SLATE, width=1.8)
    p.line([4, 16, 28, 16], color=SLATE, width=1.8)
    p.line([4, 21, 28, 21], color=NAVY_SOFT, width=1.8)
    p.line([4, 26, 20, 26], color=TEAL, width=2)


def i_pen(p):  # Chu ky
    p.rrect(5, 5, 22, 22, 2, fill=LIGHT, outline=NAVY_SOFT, width=1.6)
    p.poly([(26, 4), (30, 8), (17, 21), (12, 22), (13, 17)], fill=TEAL)
    p.line([(27.4, 5.6), (24, 9)], color=NAVY, width=1.2)


def i_checklist(p):  # Ra soat the thuc
    p.rrect(6, 3, 20, 26, 2, fill=LIGHT, outline=NAVY, width=2)
    p.circle(10, 9, 1.6, fill=TEAL)
    p.line([13, 9, 23, 9], color=NAVY, width=1.6)
    p.circle(10, 15, 1.6, fill=TEAL)
    p.line([13, 15, 23, 15], color=NAVY, width=1.6)
    p.circle(10, 21, 1.6, fill=TEAL)
    p.line([13, 21, 20, 21], color=NAVY, width=1.6)


def i_grammar(p):  # Van phong cong vu
    p.rrect(3, 8, 26, 16, 2, fill=LIGHT, outline=NAVY, width=2)
    p.line([6, 13, 20, 13], color=SLATE, width=1.8)
    p.line([6, 17.5, 16, 17.5], color=NAVY_SOFT, width=1.8)
    p.poly([(22, 16), (24.5, 19), (29, 13.5)], outline=TEAL, width=2.6)


def i_compare(p):  # So sanh doi chieu
    p.rrect(3, 5, 11, 22, 2, fill=LIGHT, outline=NAVY, width=2)
    p.rrect(18, 5, 11, 22, 2, fill=LIGHT, outline=TEAL, width=2)
    p.line([15.5, 16, 16.5, 16], color=SLATE, width=0.5)
    p.line([6, 10, 11, 10], color=NAVY, width=1.4)
    p.line([6, 14, 11, 14], color=NAVY_SOFT, width=1.4)
    p.line([21, 10, 26, 10], color=TEAL, width=1.4)
    p.line([21, 14, 26, 14], color=NAVY_SOFT, width=1.4)


def i_report(p):  # Xuat bao cao
    p.rrect(6, 3, 20, 26, 2, fill=LIGHT, outline=NAVY, width=2)
    p.line([10, 22, 10, 16], color=TEAL, width=2.6)
    p.line([14, 22, 14, 10], color=NAVY, width=2.6)
    p.line([18, 22, 18, 14], color=TEAL, width=2.6)
    p.line([22, 22, 22, 8], color=NAVY, width=2.6)


def i_building(p):  # Co quan
    p.rect(6, 12, 20, 17, fill=LIGHT, outline=NAVY, width=2)
    p.poly([(16, 3), (27, 12), (5, 12)], fill=TEAL)
    p.rect(10, 16, 3, 3, fill=NAVY)
    p.rect(19, 16, 3, 3, fill=NAVY)
    p.rect(14, 22, 4, 7, fill=NAVY)


def i_user_pen(p):  # Nguoi ky
    p.circle(11, 9, 4.5, fill=NAVY)
    p.arc(11, 25, 8.5, 180, 360, color=NAVY, width=2.6)
    p.poly([(22, 12), (29, 19), (19, 29), (14, 28), (15, 23)], fill=TEAL)


def i_inbox(p):  # Noi nhan
    p.rrect(4, 8, 24, 18, 2, fill=LIGHT, outline=NAVY, width=2)
    p.poly([(4, 16), (11, 16), (13, 20), (19, 20), (21, 16), (28, 16)], fill=TEAL)
    p.line([8, 8, 12, 12], color=NAVY_SOFT, width=1.6)
    p.line([24, 8, 20, 12], color=NAVY_SOFT, width=1.6)


def i_shield(p):  # Do khan / Do mat
    p.poly([(16, 3), (28, 8), (28, 17), (16, 29), (4, 17), (4, 8)], fill=LIGHT, outline=NAVY, width=2)
    p.line([16, 8, 16, 16], color=NAVY, width=2)
    p.arc(16, 12, 4.5, 270, 90, color=TEAL, width=2)


def i_ruler_page(p):  # A4 can le
    p.rrect(8, 3, 16, 26, 2, fill=LIGHT, outline=NAVY, width=2)
    p.line([11, 6, 11, 26], color=TEAL, width=1.6)
    p.line([21, 6, 21, 26], color=TEAL, width=1.6)
    p.line([11, 9, 13, 9], color=NAVY, width=1.2)
    p.line([21, 9, 19, 9], color=NAVY, width=1.2)
    p.line([11, 14, 13, 14], color=NAVY, width=1.2)
    p.line([21, 14, 19, 14], color=NAVY, width=1.2)


def i_landscape(p):  # Trang ngang
    p.rrect(3, 8, 26, 16, 2, fill=LIGHT, outline=NAVY, width=2)
    p.line([7, 13, 25, 13], color=NAVY_SOFT, width=1.4)
    p.line([7, 17, 22, 17], color=TEAL, width=1.4)


def i_portrait(p):  # Trang doc
    p.rrect(9, 3, 14, 26, 2, fill=LIGHT, outline=NAVY, width=2)
    p.line([12, 8, 20, 8], color=NAVY_SOFT, width=1.4)
    p.line([12, 12, 20, 12], color=TEAL, width=1.4)


def i_page_num(p):  # Danh so trang
    p.rrect(8, 3, 16, 26, 2, fill=LIGHT, outline=NAVY, width=2)
    p.rect(12, 12, 8, 7, fill=TEAL)
    p.line([9, 15.5, 12, 15.5], color=NAVY, width=1.2)
    p.line([20, 15.5, 23, 15.5], color=NAVY, width=1.2)


def i_broom(p):  # Xoa trang trang
    p.line([20, 5, 12, 15], color=NAVY, width=2.6)
    p.poly([(12, 14), (17, 17), (10, 27), (3, 22)], fill=TEAL)
    p.line([21, 9, 27, 9], color=NAVY_SOFT, width=1.6)
    p.line([22, 13, 28, 14], color=NAVY_SOFT, width=1.6)


def i_grid(p):  # Thu vien bieu mau
    p.rrect(3, 3, 26, 26, 2, outline=NAVY, width=2, fill=LIGHT)
    p.line([16, 3, 16, 29], color=NAVY_SOFT, width=1.4)
    p.line([3, 16, 29, 16], color=NAVY_SOFT, width=1.4)
    p.rect(6, 6, 7, 7, fill=TEAL)
    p.rect(19, 19, 7, 7, fill=NAVY)


def i_gear(p):  # He thong OrcaX
    p.circle(16, 16, 6, fill=TEAL)
    p.circle(16, 16, 3, fill=LIGHT)
    for ang in range(0, 360, 45):
        import math
        rad = math.radians(ang)
        x1 = 16 + 8.5 * math.cos(rad)
        y1 = 16 + 8.5 * math.sin(rad)
        x2 = 16 + 12.5 * math.cos(rad)
        y2 = 16 + 12.5 * math.sin(rad)
        p.line([x1, y1, x2, y2], color=NAVY, width=3)


def i_rocket(p):  # Xu ly hang loat
    p.poly([(16, 2), (22, 10), (22, 20), (16, 17), (10, 20), (10, 10)], fill=LIGHT, outline=NAVY, width=2)
    p.circle(16, 11, 3, fill=TEAL)
    p.poly([(10, 20), (5, 27), (11, 25)], fill=NAVY)
    p.poly([(22, 20), (27, 27), (21, 25)], fill=NAVY)
    p.line([16, 22, 16, 29], color=TEAL, width=2)


def i_key(p):  # Ban quyen
    p.circle(10, 12, 6.5, outline=NAVY, width=2.6)
    p.circle(10, 12, 2.2, fill=TEAL)
    p.line([14.5, 16.5, 27, 29], color=NAVY, width=2.6)
    p.line([22, 24, 25.5, 20.5], color=NAVY, width=2.6)


def i_toc(p):  # Muc luc tu dong
    p.line([5, 6, 13, 6], color=NAVY, width=2.2)
    p.line([5, 12, 11, 12], color=TEAL, width=2.2)
    p.line([8, 18, 13, 18], color=NAVY_SOFT, width=2.2)
    p.line([5, 24, 11, 24], color=NAVY_SOFT, width=2.2)
    p.line([17, 6, 27, 6], color=SLATE, width=1.4)
    p.line([17, 12, 27, 12], color=SLATE, width=1.4)
    p.line([17, 18, 27, 18], color=SLATE, width=1.4)
    p.line([17, 24, 23, 24], color=SLATE, width=1.4)


def i_star(p):  # Mau chuan / uu tien
    p.poly([(16, 3), (20, 12), (29, 12.5), (22, 18.5), (24.5, 28), (16, 22.5),
            (7.5, 28), (10, 18.5), (3, 12.5), (12, 12)], fill=TEAL)


def i_whale(p):  # Logo Orca
    p.poly([(4, 20), (10, 12), (20, 10), (28, 15), (28, 21), (22, 25), (10, 25)], fill=NAVY)
    p.poly([(28, 15), (29, 9), (24, 13)], fill=TEAL)
    p.circle(24.5, 16.5, 1.2, fill=WHITE)
    p.poly([(10, 25), (12, 29), (16, 25)], fill=TEAL)




def i_doc_plain(p):  # Van ban don gian
    p.rrect(8, 3, 16, 26, 2, fill=LIGHT, outline=NAVY, width=2)
    p.line([11, 9, 21, 9], color=NAVY_SOFT, width=1.4)
    p.line([11, 13, 21, 13], color=NAVY_SOFT, width=1.4)
    p.line([11, 17, 18, 17], color=TEAL, width=1.4)


def i_calendar(p):  # Ke hoach
    p.rrect(4, 6, 24, 22, 2, fill=LIGHT, outline=NAVY, width=2)
    p.rect(4, 6, 24, 6, fill=NAVY)
    p.line([10, 3, 10, 8], color=SLATE, width=2)
    p.line([22, 3, 22, 8], color=SLATE, width=2)
    p.rect(8, 15, 5, 4, fill=TEAL)
    p.rect(14, 15, 5, 4, fill=NAVY_SOFT)
    p.rect(20, 15, 4, 4, fill=TEAL_SOFT)


def i_chart_bar(p):  # Bao cao / so lieu
    p.rrect(4, 4, 24, 24, 2, fill=LIGHT, outline=NAVY, width=2)
    p.line([8, 23, 8, 16], color=TEAL, width=3)
    p.line([13, 23, 13, 11], color=NAVY, width=3)
    p.line([18, 23, 18, 18], color=TEAL, width=3)
    p.line([23, 23, 23, 9], color=NAVY, width=3)
    p.line([6, 25, 26, 25], color=SLATE, width=1.6)


def i_people(p):  # Hoi hop / nhieu nguoi
    p.circle(10, 10, 3.5, fill=NAVY)
    p.circle(21, 10, 3.5, fill=TEAL)
    p.arc(10, 23, 6, 180, 360, color=NAVY, width=2.2)
    p.arc(21, 23, 6, 180, 360, color=TEAL, width=2.2)
    p.line([14, 17, 17, 17], color=SLATE, width=1.4)


def i_mail(p):  # Cong van / giay moi
    p.rrect(3, 8, 26, 17, 2, fill=LIGHT, outline=NAVY, width=2)
    p.poly([(3, 8), (16, 18), (29, 8)], outline=NAVY, width=1.8)
    p.circle(24, 12, 2, fill=TEAL)


def i_stamp(p):  # Quyet dinh
    p.rrect(5, 20, 22, 6, 1, fill=NAVY)
    p.rrect(12, 10, 8, 8, 1, fill=TEAL, outline=NAVY, width=1.6)
    p.line([14, 10, 14, 6], color=SLATE, width=2)
    p.line([18, 10, 18, 6], color=SLATE, width=2)
    p.arc(16, 6, 2.4, 180, 360, color=SLATE, width=2)


def i_badge(p):  # Giay gioi thieu / giay tu cach
    p.rrect(7, 4, 18, 24, 2, fill=LIGHT, outline=NAVY, width=2)
    p.circle(16, 11, 3.5, fill=TEAL)
    p.line([10, 18, 22, 18], color=NAVY, width=1.6)
    p.line([10, 22, 19, 22], color=NAVY_SOFT, width=1.6)


def i_copy_doc(p):  # Ban sao van ban
    p.rrect(4, 3, 14, 18, 2, fill=LIGHT, outline=NAVY_SOFT, width=2)
    p.rrect(12, 11, 14, 18, 2, fill=TEAL_SOFT, outline=NAVY, width=2)
    p.line([15, 17, 23, 17], color=NAVY, width=1.4)
    p.line([15, 21, 23, 21], color=NAVY, width=1.4)


def i_folder(p):  # Thu muc
    p.poly([(3, 8), (12, 8), (14, 11), (28, 11), (28, 25), (3, 25)], fill=TEAL_SOFT, outline=NAVY, width=2)
    p.line([3, 14, 28, 14], color=NAVY_SOFT, width=1.4)


ICONS = {
    "bolt": i_bolt, "hub": i_hub, "search_doc": i_search_doc, "undo": i_undo,
    "compass": i_compass, "layers": i_layers, "doc_type": i_doc_type, "id_card": i_id_card,
    "heading": i_heading, "text_center": i_text_center, "justify": i_justify, "pen": i_pen,
    "checklist": i_checklist, "grammar": i_grammar, "compare": i_compare, "report": i_report,
    "building": i_building, "user_pen": i_user_pen, "inbox": i_inbox, "shield": i_shield,
    "ruler_page": i_ruler_page, "landscape": i_landscape, "portrait": i_portrait,
    "page_num": i_page_num, "broom": i_broom, "grid": i_grid, "gear": i_gear,
    "rocket": i_rocket, "key": i_key, "toc": i_toc, "star": i_star, "whale": i_whale,
    "doc_plain": i_doc_plain, "calendar": i_calendar, "chart_bar": i_chart_bar,
    "people": i_people, "mail": i_mail, "stamp": i_stamp, "badge": i_badge,
    "copy_doc": i_copy_doc, "folder": i_folder,
}


def render(name, func, size):
    img = Image.new("RGBA", (CANVAS * SCALE, CANVAS * SCALE), CLEAR)
    pen = Pen(ImageDraw.Draw(img))
    func(pen)
    return img.resize((size, size), Image.LANCZOS)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    only = set(sys.argv[1:])
    made = []
    for name, func in ICONS.items():
        if only and name not in only:
            continue
        for size in (32, 16):
            out = os.path.join(OUT_DIR, f"ico_{name}_{size}.png")
            render(name, func, size).save(out)
        made.append(name)
    print(f"OK: generated {len(made)} icons x 2 sizes -> {OUT_DIR}")
    print(" ".join(made))


if __name__ == "__main__":
    main()
