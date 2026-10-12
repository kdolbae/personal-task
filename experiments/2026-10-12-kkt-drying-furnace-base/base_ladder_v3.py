import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.patches import Rectangle
from mpl_toolkits.mplot3d.art3d import Poly3DCollection

for f in fm.findSystemFonts():
    if "wqy" in f.lower():
        fm.fontManager.addfont(f)
plt.rcParams["font.family"] = "WenQuanYi Zen Hei"
plt.rcParams["axes.unicode_minus"] = False

# ---------------- geometry (mm) ----------------
L, W, H = 10000, 5000, 800
RAIL_Y = [300, 1400, 3600, 4700]
LANE = (4700, 5300)                         # weld lane across the width (gap between two rungs)
STA = [150, 2400, 4625, 5375, 7600, 9850]   # rung / column stations (none inside the lane)

TOP_D, TOP_B, TOP_TW, TOP_TF = 250, 125, 6, 9      # top rails  H-250x125x6x9
BOT_D, BOT_B, BOT_TW, BOT_TF = 200, 100, 5.5, 8    # bottom rails H-200x100x5.5x8
RUNG_D, RUNG_B, RUNG_TW, RUNG_TF = 150, 75, 5, 7   # rungs H-150x75x5x7
COL = 125                                          # columns H-125x125x6.5x9

Z_BOT_TOP = BOT_D                 # 200
Z_TOP_BOT = H - TOP_D             # 550  (clear height under top rails)
COL_H = Z_TOP_BOT - Z_BOT_TOP     # 350

BOT_SEG = [(0, LANE[0]), (LANE[1], L)]            # bottom rails are split at the lane
KG = {"top": 29.6, "bot": 21.3, "rung": 14.0, "col": 23.8}

# ---------------- BOM ----------------
bays = [RAIL_Y[1] - RAIL_Y[0] - TOP_B, RAIL_Y[2] - RAIL_Y[1] - TOP_B, RAIL_Y[3] - RAIL_Y[2] - TOP_B]
rung_len_per_sta = sum(bays) / 1000       # m
n_sta = len(STA)
bom = []
m = len(RAIL_Y) * L / 1000
bom.append(("상부 사다리 레일 H-250×125×6×9", f"{len(RAIL_Y)}EA × 10.0m", m, m * KG["top"]))
m = n_sta * rung_len_per_sta
bom.append(("상부 사다리 가로대 H-150×75×5×7 (상면 일치)", f"{n_sta*3}EA / {n_sta}개소", m, m * KG["rung"]))
seg_len = sum(b - a for a, b in BOT_SEG) / 1000
m = len(RAIL_Y) * seg_len
bom.append(("하부 사다리 레일 H-200×100×5.5×8", f"{len(RAIL_Y)*2}EA (통로부 분할)", m, m * KG["bot"]))
m = n_sta * rung_len_per_sta
bom.append(("하부 사다리 가로대 H-150×75×5×7", f"{n_sta*3}EA / {n_sta}개소", m, m * KG["rung"]))
n_col = len(RAIL_Y) * n_sta
m = n_col * COL_H / 1000
bom.append(("기둥 H-125×125×6.5×9", f"{n_col}EA × {COL_H/1000:.2f}m", m, m * KG["col"]))
st = n_col * (2 * 120 * 230 * 9 + 2 * 90 * 184 * 9) * 7.85e-6
bom.append(("보강 스티프너 PL-9 (레일 웹 내부)", f"{n_col*4}EA", None, st))
rb = 2 * 300 * 400 * 25 * 7.85e-6 + 4 * 100 * 300 * 16 * 7.85e-6
bom.append(("푸시 반력 브래킷 (탈착식, 예시)", "1식", None, rb))
total = sum(b[3] for b in bom)
contact = (len(RAIL_Y) * seg_len * BOT_B / 1000 + n_sta * rung_len_per_sta * RUNG_B / 1000)   # m2
print("TOTAL", round(total), "loss5", round(total * 1.05), "contact m2", round(contact, 2), "tf/m2", round(90 / contact, 1))
for b in bom:
    print(b[0], round(b[3]))

ORANGE, STEEL, DARK, RED, BLUE = "#e8742a", "#8a97a6", "#3c4655", "#c0392b", "#3b78b5"

def box(ax, x0, x1, y0, y1, z0, z1, color):
    v = [(x0, y0, z0), (x1, y0, z0), (x1, y1, z0), (x0, y1, z0),
         (x0, y0, z1), (x1, y0, z1), (x1, y1, z1), (x0, y1, z1)]
    faces = [[v[i] for i in f] for f in
             [(0, 1, 2, 3), (4, 5, 6, 7), (0, 1, 5, 4), (2, 3, 7, 6), (1, 2, 6, 5), (0, 3, 7, 4)]]
    ax.add_collection3d(Poly3DCollection(faces, facecolors=color, edgecolors=DARK, linewidths=0.3))

def ibeam_x(ax, x0, x1, yc, z0, z1, b, tf, tw, color):     # beam running along X
    box(ax, x0, x1, yc - b / 2, yc + b / 2, z0, z0 + tf, color)
    box(ax, x0, x1, yc - b / 2, yc + b / 2, z1 - tf, z1, color)
    box(ax, x0, x1, yc - tw / 2, yc + tw / 2, z0 + tf, z1 - tf, color)

def ibeam_y(ax, xc, y0, y1, z0, z1, b, tf, tw, color):     # beam running along Y
    box(ax, xc - b / 2, xc + b / 2, y0, y1, z0, z0 + tf, color)
    box(ax, xc - b / 2, xc + b / 2, y0, y1, z1 - tf, z1, color)
    box(ax, xc - tw / 2, xc + tw / 2, y0, y1, z0 + tf, z1 - tf, color)

def bottom_ladder(ax, dz=0):
    for yc in RAIL_Y:
        for a, b in BOT_SEG:
            ibeam_x(ax, a, b, yc, dz, dz + BOT_D, BOT_B, BOT_TF, BOT_TW, BLUE)
    for xc in STA:
        for k in range(3):
            y0 = RAIL_Y[k] + (BOT_B / 2 if k != 1 else TOP_B / 2 + 0)
            y0 = RAIL_Y[k] + BOT_B / 2
            y1 = RAIL_Y[k + 1] - BOT_B / 2
            ibeam_y(ax, xc, y0, y1, dz, dz + RUNG_D, RUNG_B, RUNG_TF, RUNG_TW, "#7fa6cf")

def columns(ax, dz=0):
    for xc in STA:
        for yc in RAIL_Y:
            box(ax, xc - COL / 2, xc + COL / 2, yc - COL / 2, yc + COL / 2, dz + Z_BOT_TOP, dz + Z_TOP_BOT, STEEL)

def top_ladder(ax, dz=0):
    for yc in RAIL_Y:
        ibeam_x(ax, 0, L, yc, dz + Z_TOP_BOT, dz + H, TOP_B, TOP_TF, TOP_TW, ORANGE)
    for xc in STA:
        for k in range(3):
            y0 = RAIL_Y[k] + TOP_B / 2
            y1 = RAIL_Y[k + 1] - TOP_B / 2
            ibeam_y(ax, xc, y0, y1, dz + H - RUNG_D, dz + H, RUNG_B, RUNG_TF, RUNG_TW, "#f3b48a")

def setup3d(ax, zmax, elev, azim, zoom=1.0):
    ax.set_xlim(0, L); ax.set_ylim(0, W); ax.set_zlim(0, zmax)
    ax.set_box_aspect((L, W, zmax * 1.6), zoom=zoom)
    ax.view_init(elev=elev, azim=azim)
    ax.set_axis_off()

fig = plt.figure(figsize=(18, 12.5), facecolor="white")
fig.suptitle("진공건조로 설치용 사다리형 일체 베이스 검토안 (하부 사다리 + 기둥 + 상부 사다리)", fontsize=22, fontweight="bold", x=0.02, ha="left", y=0.986)
fig.text(0.02, 0.948, "10,000 × 5,000 × H800 · 목표하중 약 90톤(45t+38t) · 개략 검토안(제작도 아님): 단면·받침간격·탱크 접촉부·지반 확인 후 구조검토 필수",
         fontsize=12.5, color=RED)

# ---- ① assembled ----
ax = fig.add_axes([-0.04, 0.50, 0.44, 0.43], projection="3d")
bottom_ladder(ax); columns(ax); top_ladder(ax)
ax.add_collection3d(Poly3DCollection([[(LANE[0], 0, 4), (LANE[1], 0, 4), (LANE[1], W, 4), (LANE[0], W, 4)]],
                                     facecolors="#f6b7ae", edgecolors=RED, linewidths=1.2, alpha=0.9))
for _y in (1400, 3600):
    box(ax, 20, 90, _y - 150, _y + 150, H, H + 300, RED)
setup3d(ax, 1300, 27, -57, 1.0)
ax.text2D(0.04, 0.93, "① 조립 완성", transform=ax.transAxes, fontsize=14, fontweight="bold")
ax.text2D(0.52, 0.06, "빨간 띠 = 용접 통로 (가로대 사이 빈 공간)", transform=ax.transAxes, fontsize=10.5, color=RED, ha="center")

# ---- ② exploded ----
ax_e = fig.add_axes([0.33, 0.50, 0.32, 0.43], projection="3d")
bottom_ladder(ax_e, 0); columns(ax_e, 450); top_ladder(ax_e, 900)
setup3d(ax_e, 2000, 24, -57, 1.0)
ax_e.text2D(0.04, 0.93, "② 제작 순서 (분해)", transform=ax_e.transAxes, fontsize=14, fontweight="bold")
ax_e.text2D(0.02, 0.18, "1) 하부 사다리 바닥에서 조립", transform=ax_e.transAxes, fontsize=10.5, color=BLUE)
ax_e.text2D(0.02, 0.12, "2) 기둥 세우기 (레일 교차점 24곳)", transform=ax_e.transAxes, fontsize=10.5, color="#555")
ax_e.text2D(0.02, 0.06, "3) 상부 사다리 덮기 (레일 상면 평탄)", transform=ax_e.transAxes, fontsize=10.5, color=ORANGE)

# ---- ③ BOM ----
ax4 = fig.add_axes([0.66, 0.50, 0.33, 0.40]); ax4.axis("off")
ax4.set_title("③ 자재 산출 (개략)", loc="left", fontsize=14, fontweight="bold")
rows = [[b[0], b[1], "-" if b[2] is None else f"{b[2]:.1f}", f"{b[3]:,.0f}"] for b in bom]
rows.append(["합계", "", "", f"{total:,.0f} kg"])
rows.append(["손실 5% 포함", "", "", f"{total*1.05:,.0f} kg (약 {total*1.05/1000:.1f}톤)"])
tb = ax4.table(cellText=rows, colLabels=["품명·규격", "수량", "길이(m)", "중량(kg)"], loc="upper center",
               cellLoc="left", colWidths=[0.50, 0.19, 0.10, 0.21])
tb.auto_set_font_size(False); tb.set_fontsize(8.8); tb.scale(1, 1.6)
for (r, c), cell in tb.get_celld().items():
    cell.set_edgecolor("#bbb")
    if r == 0:
        cell.set_facecolor("#2c3a4a"); cell.get_text().set_color("white")
    elif r >= len(rows) - 1:
        cell.set_facecolor("#fdf0e6"); cell.get_text().set_fontweight("bold")

# ---- plans ----
def plan(rect, title, kind):
    a = fig.add_axes(rect)
    a.set_title(title, loc="left", fontsize=13, fontweight="bold")
    a.add_patch(Rectangle((LANE[0], 0), LANE[1] - LANE[0], W, fc="#fde8e4", ec=RED, lw=1.1, ls="--"))
    if kind == "top":
        for yc in RAIL_Y:
            a.add_patch(Rectangle((0, yc - TOP_B / 2), L, TOP_B, fc=ORANGE, ec=DARK, lw=0.5))
    else:
        for yc in RAIL_Y:
            for s, e in BOT_SEG:
                a.add_patch(Rectangle((s, yc - BOT_B / 2), e - s, BOT_B, fc=BLUE, ec=DARK, lw=0.5))
    for xc in STA:
        for k in range(3):
            y0 = RAIL_Y[k] + 60; y1 = RAIL_Y[k + 1] - 60
            a.add_patch(Rectangle((xc - RUNG_B / 2, y0), RUNG_B, y1 - y0,
                                  fc="#f3b48a" if kind == "top" else "#7fa6cf", ec=DARK, lw=0.4))
        for yc in RAIL_Y:
            a.add_patch(Rectangle((xc - COL / 2, yc - COL / 2), COL, COL, fc="none", ec=DARK, lw=0.5, ls=":"))
    if kind == "top":
        for _y in (1400, 3600):
            a.add_patch(Rectangle((20, _y - 150), 70, 300, fc=RED, ec=RED))
        a.annotate("", xy=(3400, 2500), xytext=(1400, 2500), arrowprops=dict(arrowstyle="-|>", color=RED, lw=2))
        a.text(2400, 2750, "푸시(밀착) 방향", ha="center", fontsize=10, color=RED)
    a.text(5000, W + 280, "용접 통로 W600", ha="center", fontsize=10.5, color=RED)
    a.annotate("", xy=(0, -330), xytext=(L, -330), arrowprops=dict(arrowstyle="<->"))
    a.text(L / 2, -700, "10,000", ha="center", fontsize=10.5)
    a.set_xlim(-250, L + 250); a.set_ylim(-950, W + 650)
    a.set_aspect("equal"); a.axis("off")

plan([0.03, 0.27, 0.31, 0.19], "④ 상부 사다리 평면", "top")
plan([0.36, 0.27, 0.31, 0.19], "⑤ 하부 사다리 평면 (통로부 레일 분할)", "bot")

# ---- elevation ----
a3 = fig.add_axes([0.03, 0.035, 0.62, 0.20])
a3.set_title("⑥ 측면도 (길이방향 = 밀착 방향, 단위 mm)", loc="left", fontsize=13, fontweight="bold")
a3.add_patch(Rectangle((0, Z_TOP_BOT), L, TOP_D, fc=ORANGE, ec=DARK, lw=0.8))
a3.text(1300, Z_TOP_BOT + TOP_D / 2, "상부 레일 H-250 (10m 연속, 상면 평탄)", va="center", fontsize=9.5, color="white")
for s, e in BOT_SEG:
    a3.add_patch(Rectangle((s, 0), e - s, BOT_D, fc=BLUE, ec=DARK, lw=0.8))
a3.text(1300, BOT_D / 2, "하부 레일 H-200", va="center", fontsize=9, color="white")
for xc in STA:
    a3.add_patch(Rectangle((xc - COL / 2, Z_BOT_TOP), COL, COL_H, fc=STEEL, ec=DARK, lw=0.6))
a3.add_patch(Rectangle((LANE[0], 0), LANE[1] - LANE[0], Z_TOP_BOT, fc="#fde8e4", ec=RED, lw=1.2, ls="--"))
a3.text(5000, Z_TOP_BOT / 2, "통로\n600", ha="center", va="center", fontsize=10, color=RED)
a3.annotate("", xy=(5900, 0), xytext=(5900, Z_TOP_BOT), arrowprops=dict(arrowstyle="<->"))
a3.text(6000, Z_TOP_BOT / 2, "통로부 순높이 550", va="center", fontsize=10)
a3.annotate("", xy=(-250, 0), xytext=(-250, H), arrowprops=dict(arrowstyle="<->"))
a3.text(-330, H / 2, "800", rotation=90, ha="right", va="center", fontsize=10)
a3.add_patch(Rectangle((200, H), LANE[0] - 250, 650, fc="#e6e6e6", ec="#888", lw=0.8, ls="--"))
a3.add_patch(Rectangle((LANE[1] + 50, H), L - LANE[1] - 250, 650, fc="#e6e6e6", ec="#888", lw=0.8, ls="--"))
a3.text(2400, H + 330, "1호기 45t(도어)", ha="center", fontsize=10, color="#555")
a3.text(7650, H + 330, "2호기 38t", ha="center", fontsize=10, color="#555")
a3.add_patch(Rectangle((20, H), 50, 420, fc=RED, ec=RED))
a3.annotate("", xy=(4000, H + 180), xytext=(3000, H + 180), arrowprops=dict(arrowstyle="-|>", color=RED, lw=2))
a3.set_xlim(-700, L + 400); a3.set_ylim(-120, H + 800)
a3.set_aspect("equal"); a3.axis("off")

# ---- notes ----
a5 = fig.add_axes([0.66, 0.035, 0.33, 0.40]); a5.axis("off")
a5.set_title("⑦ 설계 포인트", loc="left", fontsize=14, fontweight="bold")
a5.text(0.0, 0.96,
        "· 하부 사다리가 바닥에 면으로 하중 분산\n"
        f"   (접지면적 약 {contact:.1f}㎡ → 약 {90/contact:.0f} tf/㎡, 이전 바닥판 안 약 23 tf/㎡)\n"
        "· 용접 통로 = 가로대 사이 빈 칸 (너비방향 전폭 개방)\n"
        "   통로부는 하부 레일도 분할 → 바닥이 평평, 기어서 작업 가능\n"
        "· 상부 레일 4줄 10m 연속 → 통로 위를 600mm만 통과\n"
        "· 상부 가로대는 레일 상면과 높이 일치(걸림 없음)\n"
        "   스티프너는 레일 웹 안쪽, 용접 비드는 그라인딩\n"
        "· 반력 브래킷은 후단에 탈착식 (밀착 후 제거)\n"
        "· 영상(크레인 100t·60t)의 베이스는 H600으로 표기됨\n"
        "   H-250 레일을 쓰면 통로부 순높이가 350 → 800 이상 필요\n"
        "· 하중 90톤: 레일 휨·웹 크리플링·지반 지지력은\n"
        "   구조검토 후 확정(규격은 예시)",
        fontsize=10.2, va="top", color="#222", linespacing=1.5)

out = __import__('os').path.join(__import__('os').path.dirname(__import__('os').path.abspath(__file__)), 'base_ladder_v3.png')
fig.savefig(out, dpi=105, facecolor="white")
print("saved", out)
