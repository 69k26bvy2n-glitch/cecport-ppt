# -*- coding: utf-8 -*-
# 样例：三角色泳道 + 双层信息流/物流流程图（P11 VMI 库存管理服务，2026-10-15 定稿）
# 打法说明见技能 references/page-patterns.md §8.3；组件全部来自 cec_lib。
# 运行：把本文件与 scripts/cec_lib.py、assets/ 复制到同一工程目录后 python3 运行，
# 不要直接在技能目录内跑（依赖同级 cec_lib.py 与 assets/）。复用前改文末 out= 输出路径。
import sys, os, math
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cec_lib import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# 流程组件全部来自 cec_lib：set_dash/set_arrow、rr/R_CARD/R_BOX/card_shadow、
# dashed_box/lane_title、vlabel/line_pill（VL_W）。详见技能 references/page-patterns.md §8.3。

# ============================================================
prs = new_deck()
s = new_page(prs)
header(s, "VMI 库存管理服务",
       "02 · 以上 7 种类型的示意图 · Vendor Managed Inventory")
footer(s, 11)

# ============ 左列：服务介绍卡 ============
IX, IY, IW, IH = ML, 1.28, 3.35, 1.62
card(s, IX, IY, IW, IH)
badge(s, IX + 0.30, IY + 0.24, 0.74, "02", fill=ORANGE, tsize=20)
add_text(s, IX + 1.18, IY + 0.22, IW - 1.40, 0.34,
         [{"runs": [R("VMI 库存管理服务", 14.5, True, NAVY)]}],
         anchor=MSO_ANCHOR.MIDDLE)
add_text(s, IX + 0.30, IY + 1.06, IW - 0.58, 0.50,
         [{"runs": [R("提供安全库存代管、按需供货，降本提效、保障交付",
                      9.8, False, SUB)], "line_spacing": 1.25}])

# ============ 左列：核心能力卡 ============
CY0, CW0, CH0 = 3.06, IW, 3.84
card(s, IX, CY0, CW0, CH0)
icon_title(s, IX + 0.28, CY0 + 0.20, CW0 - 0.5, "核心能力", "target",
           isize=0.27, size=13.5)
caps = [
    ("多源物料统一集散", "package"),
    ("动态库存精准管控", "chart"),
    ("按需齐套精准配送", "truck"),
    ("全链路信息实时协同", "digital"),
]
ry0, pitch = 3.86, 0.72
for i, (t, ic) in enumerate(caps):
    y = ry0 + i * pitch
    if i:
        add_line(s, IX + 0.28, y - 0.06, CW0 - 0.56, color=HAIR, weight=0.75)
    icon_chip(s, ic, IX + 0.57, y + 0.30, d=0.54)
    add_text(s, IX + 0.99, y, CW0 - 1.2, 0.60,
             [{"runs": [R(t, 12, True, INK)]}], anchor=MSO_ANCHOR.MIDDLE)

# ============ 右：流程图大卡 ============
FX, FY, FW, FH = 4.08, 1.28, 8.70, 5.62
card(s, FX, FY, FW, FH)
icon_title(s, FX + 0.24, FY + 0.16, 3.0, "VMI 业务流程示意图", "sitemap",
           isize=0.26, size=13.5)
# 图例（说明线型语义）
l1 = add_line_xy(s, FX + FW - 2.42, FY + 0.22, FX + FW - 2.10, FY + 0.22,
                 color=BLUE, weight=1.25); set_dash(l1)
add_text(s, FX + FW - 2.04, FY + 0.095, 0.62, 0.24,
         [{"runs": [R("信息流", 8.5, False, MUTE)]}], anchor=MSO_ANCHOR.MIDDLE)
l2 = add_line_xy(s, FX + FW - 1.36, FY + 0.22, FX + FW - 1.04, FY + 0.22,
                 color=ORANGE, weight=2.25); set_arrow(l2, tail=True)
add_text(s, FX + FW - 0.98, FY + 0.095, 0.74, 0.24,
         [{"runs": [R("物流 / 作业流", 8.5, False, MUTE)]}],
         anchor=MSO_ANCHOR.MIDDLE)

# 三条泳道
INL, INR = FX + 0.20, FX + FW - 0.20          # 4.28 / 12.58
GAP = 0.62
LW3 = (INR - INL - 2 * GAP) / 3               # 2.49
L1X = INL
L2X = INL + LW3 + GAP                          # 7.18
L3X = INL + 2 * (LW3 + GAP)                    # 10.08
L3W = INR - L3X
LY, LHh = 2.62, 3.90
dashed_box(s, L1X, LY, LW3, LHh)
dashed_box(s, L2X, LY, LW3, LHh)
dashed_box(s, L3X, LY, L3W, LHh)

# 泳道标题
lane_title(s, L1X + LW3 / 2, LY + 0.12, 1.50, "供应商", NAVY)
lane_title(s, L2X + LW3 / 2, LY + 0.10, 1.90, "客户 VMI HUB", ORANGE)
lane_title(s, L3X + L3W / 2, LY + 0.12, 1.86, "客户生产基地", NAVY)

# ---------- 泳道一：供应商 ----------
RGX, RGW = L1X + 0.13, 0.62                    # 区域标签列
ROWX, ROWW = RGX + RGW + 0.10, L1X + LW3 - 0.14 - (RGX + RGW + 0.10)
rp_y0, rp_h, rp_pitch = 3.34, 0.62, 0.76
regions = [("特殊区域", rp_y0, rp_h),
           ("境内区外", rp_y0 + rp_pitch, rp_h + rp_pitch - 0.14),
           ("境外", rp_y0 + 3 * rp_pitch, rp_h)]
for txt, ry, rh in regions:
    add_rect(s, RGX, ry, RGW, rh, fill=TINT, radius=rr(RGW, rh))
    add_text(s, RGX, ry, RGW, rh,
             [{"runs": [R(txt, 8.8, True, BLUE)], "align": PP_ALIGN.CENTER,
              "line_spacing": 1.08}], anchor=MSO_ANCHOR.MIDDLE)
for i, name in enumerate(["供应商 1", "供应商 2", "供应商 3", "供应商 N"]):
    y = rp_y0 + i * rp_pitch
    add_rect(s, ROWX, y, ROWW, rp_h, fill=CARD, line=HAIR, line_w=0.75,
             radius=rr(ROWW, rp_h))
    add_icon(s, "warehouse", ROWX + 0.14, y + (rp_h - 0.28) / 2, 0.28)
    add_text(s, ROWX + 0.50, y, ROWW - 0.56, rp_h,
             [{"runs": [R(name, 10.5, True, NAVY)]}],
             anchor=MSO_ANCHOR.MIDDLE)

# ---------- 泳道二：VMI HUB 九宫格服务 ----------
GX0, GX1 = L2X + 0.13, L2X + LW3 - 0.13
svc = ["报关", "物流", "仓储", "装卸", "分拣", "重工",
       "制标", "贴标", "退换"]
sic = ["file-text", "truck", "warehouse", "truck-loading", "pick", "gear",
       "tag", "tags", "arrows-exchange"]
gpitch_x = (GX1 - GX0) / 3
gy0, gpitch_y = 3.20, 0.70
for i, (txt, ic) in enumerate(zip(svc, sic)):
    col, row = i % 3, i // 3
    cx = GX0 + gpitch_x * (col + 0.5)
    cy = gy0 + gpitch_y * row
    icon_chip(s, ic, cx, cy, d=0.44)
    add_text(s, cx - 0.40, cy + 0.28, 0.80, 0.20,
             [{"runs": [R(txt, 8.8, True, NAVY)], "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)
# 增值服务全文（原文一字不动）
SBX, SBY, SBW, SBH = L2X + 0.13, 5.14, LW3 - 0.26, 1.24
add_rect(s, SBX, SBY, SBW, SBH, fill=TINT_OR, radius=rr(SBW, SBH))
add_text(s, SBX + 0.14, SBY, SBW - 0.28, SBH,
         [{"runs": [R("报关、物流、仓储、装卸、个性化分拣、重工、制标、贴标、不良品退换",
                      8.8, True, ORANGE2)], "align": PP_ALIGN.CENTER,
              "line_spacing": 1.32}], anchor=MSO_ANCHOR.MIDDLE)

# ---------- 泳道三：客户生产基地 ----------
bx0, bw0 = L3X + 0.14, L3W - 0.28
by0, bh0, bpitch = 3.34, 0.94, 1.02
bcols = [BLUE, SKY, MIST]
for i, (name, col) in enumerate(zip(["生产基地A", "生产基地2", "生产基地N"], bcols)):
    y = by0 + i * bpitch
    add_rect(s, bx0, y, bw0, bh0, fill=CARD, line=HAIR, line_w=0.75,
             radius=rr(bw0, bh0))
    add_rect(s, bx0 + 0.16, y + (bh0 - 0.56) / 2, 0.56, 0.56,
             fill=col, radius=0.22)
    add_icon(s, "building-factory", bx0 + 0.27, y + (bh0 - 0.34) / 2, 0.34,
             white=True)
    add_text(s, bx0 + 0.86, y, bw0 - 1.0, bh0,
             [{"runs": [R(name, 11, True, NAVY)]}],
             anchor=MSO_ANCHOR.MIDDLE)

# ---------- 顶部：客户 ----------
C2CX = L2X + LW3 / 2
CBW, CBH = 1.94, 0.56
CBX, CBY = C2CX - CBW / 2, 1.86
cust = add_rect(s, CBX, CBY, CBW, CBH, fill=TINT, line=SKY, line_w=1.1,
                radius=rr(CBW, CBH, R_CARD))
add_icon(s, "building", CBX + 0.30, CBY + (CBH - 0.30) / 2, 0.30)
add_text(s, CBX + 0.68, CBY, CBW - 0.85, CBH,
         [{"runs": [R("客户", 14, True, NAVY)]}], anchor=MSO_ANCHOR.MIDDLE)

# ---------- 信息流（蓝色虚线） ----------
# 1) 客户 → 供应商：订单信息（折线，箭头在下端）
h1 = add_line_xy(s, CBX, CBY + CBH / 2, L1X + LW3 / 2, CBY + CBH / 2,
                 color=BLUE, weight=1.25); set_dash(h1)
v1 = add_line_xy(s, L1X + LW3 / 2, CBY + CBH / 2, L1X + LW3 / 2, LY + 0.10,
                 color=BLUE, weight=1.25); set_dash(v1); set_arrow(v1, tail=True)
# 标签直接坐在线上：白色圆角胶囊（描边随线色）压住贯通虚线断线，文字色随线色
_lab1_cx = (CBX + L1X + LW3 / 2) / 2 + 0.025
line_pill(s, _lab1_cx, CBY + CBH / 2, 0.69, "订单信息", 8.8)
# 2) 客户 ↔ 生产基地：物料需求信息（双向；标签坐水平虚线上，虚线两段切入胶囊）
_h2y = CBY + CBH / 2
_x2r = L3X + L3W / 2
_lab2_cx = (CBX + CBW + _x2r) / 2
_w2 = 0.96
h2a = add_line_xy(s, CBX + CBW, _h2y, _lab2_cx - _w2 / 2 + 0.02, _h2y,
                  color=BLUE, weight=1.25); set_dash(h2a)
h2b = add_line_xy(s, _lab2_cx + _w2 / 2 - 0.02, _h2y, _x2r, _h2y,
                  color=BLUE, weight=1.25); set_dash(h2b)
v2 = add_line_xy(s, _x2r, _h2y, _x2r, LY + 0.10,
                 color=BLUE, weight=1.25); set_dash(v2)
set_arrow(v2, head=True, tail=True)
line_pill(s, _lab2_cx, _h2y, _w2, "物料需求信息", 8.8)
# 3) 客户 → 生产基地：生产订单（斜向虚线；标签坐斜线中点，斜线按胶囊边界切成两段）
_d3x0, _d3y0 = CBX + CBW - 0.06, CBY + CBH - 0.12
_d3x1, _d3y1 = L3X + 0.62, LY + 0.06
_cx3, _cy3 = (_d3x0 + _d3x1) / 2, (_d3y0 + _d3y1) / 2
_w3 = 0.68
_dxl, _dyl = _d3x1 - _d3x0, _d3y1 - _d3y0
_ang3 = math.degrees(math.atan2(_dyl, _dxl))   # 胶囊与文字沿线走向旋转
_L3 = math.hypot(_dxl, _dyl)
_ux, _uy = _dxl / _L3, _dyl / _L3
# 旋转后斜线从胶囊短边穿出：沿线方向半宽 w/2，端点内收 0.02in
_t3 = _w3 / 2 - 0.02
d3a = add_line_xy(s, _d3x0, _d3y0, _cx3 - _t3 * _ux, _cy3 - _t3 * _uy,
                  color=BLUE, weight=1.25); set_dash(d3a)
d3b = add_line_xy(s, _cx3 + _t3 * _ux, _cy3 + _t3 * _uy, _d3x1, _d3y1,
                  color=BLUE, weight=1.25); set_dash(d3b); set_arrow(d3b, tail=True)
line_pill(s, _cx3, _cy3, _w3, "生产订单", 8.5, rot=_ang3)
# 4) 客户 ↔ HUB：收发货信息（短竖线双向 + 竖排标签）
v4 = add_line_xy(s, C2CX, CBY + CBH, C2CX, LY + 0.08,
                 color=BLUE, weight=1.25); set_dash(v4)
set_arrow(v4, head=True, tail=True)
add_text(s, C2CX + 0.12, 2.43, 0.72, 0.20,
         [{"runs": [R("收发货信息", 7.5, True, BLUE)]}],
         anchor=MSO_ANCHOR.MIDDLE)
# 跨泳道连线统一打法（用户规则）：间隙放不下横排标签 → 竖排白底标签居中；
# 连线贯通、白框压住中段断线（直角、无边框），文字色随线色；
# 虚线用两段（端点切入白框，避免白框边外露出虚线碎头）。
# 双向流箭头两端外张（◄─字─►）；单向流只在目的泳道边带入向箭头。

# 5) 供应商 ↔ HUB：收发货信息（蓝色虚线两段，端点切入白框避免虚线碎头，两端外张箭头）
y5 = 4.42
cx5 = L1X + LW3 + GAP / 2
_eL5, _eR5 = cx5 - VL_W / 2 + 0.02, cx5 + VL_W / 2 - 0.02
a5 = add_line_xy(s, _eL5, y5, L1X + LW3, y5,
                 color=BLUE, weight=1.25); set_dash(a5); set_arrow(a5, tail=True)
b5 = add_line_xy(s, _eR5, y5, L2X, y5,
                 color=BLUE, weight=1.25); set_dash(b5); set_arrow(b5, tail=True)
vlabel(s, cx5, y5, "收发货信息", color=BLUE, size=8)

# ---------- 物流 / 作业流（橙色实线粗箭头） ----------
# 6) 供应商 → HUB：质检及入库（橙线贯通，箭头落在 HUB 泳道边）
y6 = 6.12
cx6 = L1X + LW3 + GAP / 2
l6 = add_line_xy(s, L1X + LW3, y6, L2X, y6,
                 color=ORANGE, weight=2.25); set_arrow(l6, tail=True)
vlabel(s, cx6, y6, "质检及入库", color=ORANGE, size=8.2)
# 7) HUB → 生产基地：齐套及出货（橙线贯通，箭头落在生产基地泳道边）
y7 = 5.30
cx7 = L2X + LW3 + GAP / 2
l7 = add_line_xy(s, L2X + LW3, y7, L3X, y7,
                 color=ORANGE, weight=2.25); set_arrow(l7, tail=True)
vlabel(s, cx7, y7, "齐套及出货", color=ORANGE, size=8.2)

out = "/Users/le.wen/Desktop/中电港PPT优化/P11_VMI库存管理服务_中电港风格.pptx"
prs.save(out)
print("saved", out)
