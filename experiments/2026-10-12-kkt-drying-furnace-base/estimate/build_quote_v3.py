"""엔엠솔루션 → 금강변압기 견적서 v3.
v2와 달라진 점 (사용자 피드백 2026-10-12):
- 기업이윤을 끼워 맞추는 대신 원가계산 방식(직접공사비 → 법정경비 → 일반관리비 → 이윤)으로 계산.
- 산재·고용보험은 2026 고시 요율로 각각 계상 (3.3% 일괄 삭제).
- 용접사 단가 상향, 고급기술자는 제관사 수준, 식대·출장 교통비 추가, 안전펜스 추가.
- 베이스는 임시 구조물: 방청·도장 삭제, H빔은 회수조건 단가 적용.
공종 목표(사용자 지정, 대략): 도비 2,000만 / 용접 2,500만 / 베이스 1,800만. 이윤율은 전 공종 12%로 같게 두고,
나머지 금액은 단가·요율에서 나온 값 그대로 둔다."""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "견적서_금강변압기_진공건조로_회전안착_용접_v3.xlsx")
# 디자인 규격 (엔엠솔루션 견적서 양식, kevin 2026-10-13): Pretendard 하나, Navy/Orange + 본문·보조·구분선·표머리 색만 사용
FN = "Pretendard"
FN_SB = "Pretendard SemiBold"
C_NAVY, C_ORANGE, C_TEXT, C_SUB, C_LINE, C_HEAD = "0F2A44", "E8742A", "1D2430", "5D6878", "DDE3EA", "F4F6F9"
WON = '#,##0;-#,##0;"-"'
thin = Side(style="thin", color=C_LINE)
navy2 = Side(style="medium", color=C_NAVY)   # 표 위 굵은 선 2pt
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
HEADBOX = Border(left=thin, right=thin, top=navy2, bottom=thin)
GRAY = PatternFill("solid", fgColor=C_HEAD)   # 표 머리·소계 배경
NAVY = PatternFill("solid", fgColor=C_NAVY)   # 합계 줄
YEL = PatternFill(fill_type=None)             # 규격상 노란 강조 미사용
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")

QNO = "NM-2610-01"
QDATE = "2026-10-19"
SUP = {"name": "엔엠솔루션", "en": "NM SOLUTION", "ceo": "김성배",
       "addr": "경기도 평택시 고덕동 도시지원로 121 지식공장아이타워 906호",
       "tel": "1800-5901", "mobile": "010-2369-3691", "web": "nm-solution.co.kr", "bizno": ""}
SUP_LINE = f"{SUP['name']}  |  {SUP['addr']}  |  대표전화 {SUP['tel']}  |  직통 {SUP['mobile']}  |  {SUP['web']}"
_CMAP = {"000000": C_TEXT, "0000FF": C_TEXT, "008000": C_TEXT, "555555": C_SUB, "777777": C_SUB,
         "999999": C_SUB, "AAAAAA": C_SUB, "0563C1": C_NAVY}

def font(b=False, sz=10, color="000000", sb=False):
    return Font(name=FN_SB if sb else FN, bold=b, size=sz, color=_CMAP.get(color, color))


def head_cell(x):
    x.font = font(sz=10, sb=True); x.fill = GRAY; x.border = HEADBOX; x.alignment = CEN

# ---------------------------------------------------------------- 출처
S_INS = "고용노동부 2026년 보험료율 고시 (대한전문건설협회 안내자료)"
S_SAFE = "고용노동부 건설업 산업안전보건관리비 계상 및 사용기준 (5억 미만 건축공사 3.11%)"
S_RULE = "국가계약법 시행규칙 제8조 (일반관리비 6%·이윤 15% 상한, 민간공사 참고)"
U = "사용자 제시"
SH = "서해테크 도비 견적서(2026-10-10) 단가"
VQ = "공급처 H빔 견적"
MK = "현장 실거래 수준 적용 (사용자 지시: 용접사 상향, 고급기술자=제관사 수준)"
AS = "가정값 (업체 견적으로 확인 필요)"
URLS = [
    ("산재·고용보험 요율", "https://www.kosca21.or.kr/km_board/upload/km_b_notice/1767346180_02.pdf"),
    ("산업안전보건관리비", "https://biz.newdaily.co.kr/site/data/html/2026/05/07/2026050700116.html"),
    ("원가계산 제비율", "https://04noti.com/construction-cost-ratio-guide/"),
    ("2026 4대보험(연금·건강)", "https://calctools.co.kr/blog/daily-worker-insurance"),
    ("화재감시자 기준", "산업안전보건기준에 관한 규칙 제241조의2"),
]

# ---------------------------------------------------------------- 단가 (적용)
R_WELD = 500_000      # 용접사 (사용자 지정 2026-10-12, 플랜트용접공 노임 × 약 1.71)
R_BOIL = 350_000      # 고급기술자 = 제관사 수준 (사용자 지정 2026-10-12, 플랜트제관공 노임 × 약 1.32)
R_GRIND = 300_000     # 사상 기술자 (사용자 지정)
R_HELP = 200_000      # 조공 (사용자 지정)
R_FIRE = 250_000      # 화재감시자 (사용자 지정)
R_SIGNAL = 260_000    # 신호수
R_SAFE = 320_000      # 안전담당 관리감독자
MEAL = 20_000         # 1인 1일 식대 (중식 12,000 + 간식·음료 8,000)
CAR = 85_000          # 출장 차량 1대 1일 (유류·통행료)

EXP = 600_000         # 공종별 경비 1식 (간식·음료·교통비 등, 사용자 지정 2026-10-12)

EXPD = 50_000         # 경비 인당 1일 (식대·교통비·부대 잡비, 사용자 지정 2026-10-12, 도비 공종 제외)

def expense_pd(md, note):
    return ("경비", "경비", "식대·교통비·부대 잡비", "식", 1, md * EXPD, note, "E", f"{md}인·일 기준 1식", U, None)

def expense():  # 사용자 지시(2026-10-12): 간식·음료·교통비 등을 하나로 묶어 경비 60만원
    return ("경비", "경비", "간식·음료·교통비 등", "식", 1, EXP, "", "E", "1식 600,000원", U, None)

# (구분, 항목, 규격, 단위, 수량, 단가, 비고, 비목 M/L/E, 산출식, 근거, 노임참고)
DOBI = [
    ("노무비", "도비 전문기공", "회전·안착·밀착·완성품 이동", "인·일", 21, 300_000, "안착일 6명 + 밀착 6명·3명 + 이동일 6명 (6, 6, 3, 6)", "L",
     "(6 + 6 + 3 + 6)명 × 300,000원", "인원 " + U + " (2026-10-19), 단가 " + SH, None),
    ("노무비", "신호수", "크레인 인양 신호", "인·일", 4, R_SIGNAL, "2명 × 2일 (안착일 + 이동일)", "L",
     "2명 × 2일 × 260,000원", AS, None),
    ("노무비", "안전담당 (관리감독자)", "작업 전 TBM·인양구역 통제", "인·일", 4, R_SAFE, "1명 × 4일", "L",
     "1명 × 4일 × 320,000원", AS, None),
    # 2026-10-19 kevin: 250톤 1대 단독 작업(다른 크레인 없음). 250t 1일 단가 5,000,000원은 기본값(업체 견적 없음)
    ("장비", "크레인", "250TON", "대·일", 2, 5_000_000, "안착일 1일 + 이동일 1일", "E", "1대 × 2일 × 5,000,000원", U, None),
    ("장비", "아웃리거 받침 철판", "1,500×1,500×25t, 8매", "식", 1, 640_000, "아스팔트 지반 보호", "E", "8매 × 80,000원", AS, None),
    ("장비", "밀착 장비", "유압잭·윈치·스키드", "SET", 1, 2_000_000, "밀착 2일", "E", "1식 (2일)", AS, None),
    ("장비", "공구사용손료", "도비 전문장비·줄걸이구", "SET", 1, 3_000_000, "", "E", "1식", SH, None),
    ("재료비", "소모품·잡자재", "목재 받침·와이어 보호대 등", "식", 1, 1_000_000, "", "M", "1식", SH, None),
    ("안전시설", "안전펜스 임대", "이동식 메쉬펜스 H1.8×W2.4m", "매·일", 22 * 12, 5_000, "22매(둘레 50m + 출입문) × 12일", "E",
     "작업구역 10×15m, 둘레 50m ÷ 2.4m = 21매 + 출입문 1매 → 22매 × 12일 × 5,000원", "수량 " + AS + ", 단가 " + U, None),
    ("안전시설", "펜스 설치·해체", "설치 인건비", "인·일", 2, 250_000, "설치 2명 × 0.5일 + 해체 2명 × 0.5일", "L",
     "2명 × 1일 × 250,000원", U, None),
    ("안전시설", "펜스 운반", "1톤 화물 왕복", "회", 2, 260_000, "반입 1 + 반출 1", "E", "2회 × 260,000원", AS, None),
    ("안전시설", "출입통제 안전장비", "", "식", 1, 380_000, "", "M", "1식", AS, None),
    expense(),
]
WELD = [
    ("노무비", "용접사", "CO2 용접 (접합부)", "인·일", 9, R_WELD, "3명 × 3일", "L",
     "3명 × 3일 × 500,000원", "인원·일수 " + U + ", 단가 " + MK, None),
    ("노무비", "고급기술자", "제관사 수준 (취부·치수 관리)", "인·일", 6, R_BOIL, "2명 × 3일", "L",
     "2명 × 3일 × 350,000원", "인원·일수 " + U + ", 단가 " + MK, None),
    ("노무비", "사상 기술자", "접합면 그라인딩", "인·일", 4, R_GRIND, "2명 × 2일", "L",
     "2명 × 2일 × 300,000원", "인원·일수 " + U, None),
    ("노무비", "조공", "사상·용접 보조", "인·일", 4, R_HELP, "2명 × 2일", "L",
     "2명 × 2일 × 200,000원", "인원·일수 " + U, None),
    ("노무비", "화재감시자", "용접·사상 화기작업 전담", "인·일", 5, R_FIRE, "1명 × 5일 (사상 2 + 용접 3)", "L",
     "1명 × 5일 × 250,000원", "산업안전보건기준에 관한 규칙 제241조의2, 단가 " + AS, None),
    ("장비", "CO2 용접기 대여", "500A급", "대·일", 9, 400_000, "3대 × 3일", "E", "3대 × 3일 × 400,000원", U, None),
    ("장비", "그라인더·환기·조명", "송풍기, 투광등 포함", "SET", 1, 800_000, "", "E", "1식", AS, None),
    ("재료비", "용접 소모품", "CO2 가스·와이어·디스크·케이블·잡자재", "식", 1, 3_680_000, "", "M", "1식 (용접재 3,000,000 + 전원 케이블·잡자재 680,000)", AS, None),
    ("장비", "작업발판 장비 임대", "", "식", 1, 1_000_000, "", "E", "1식", U, None),
    ("장비", "가스절단기 임대", "산소·LPG, 2대 공정기간 상시", "대·일", 24, 100_000, "2대 × 12일", "E", "2대 × 12일 × 100,000원", U, None),
    expense_pd(9 + 6 + 4 + 4 + 5, "용접 9 + 고급기술자 6 + 사상 4 + 조공 4 + 화재감시 5"),
]
BASE = [   # 2026-10-16 kevin: 재료비·운반비·노무비(3명 × 1일)만, 법정경비·일반관리비·이윤 미적용
    # 2026-10-19 kevin: 기존 철판 조방 위 H빔 10m 2본(길이 방향) + 양 끝 3m 빔 직각 용접. 금액은 그대로
    ("재료비", "H빔 (회수조건)", "H-250×250×9×14, 10m 2EA·3m 2EA", "식", 1, 1_800_000, "공사 후 회수 조건", "M", "1식 (1,882kg)", None, None),
    ("운반비", "H빔 운반", "반입·회수", "식", 1, 160_000, "", "E", "1식", None, None),
    ("노무비", "시공자", "베이스 배치·용접·회수", "인·일", 3, 350_000, "3명 × 1일", "L", "3명 × 1일 × 350,000원", None, None),
]
NO_OH = {"베이스"}   # 법정경비·일반관리비·이윤을 붙이지 않는 공종

# 요율 (요율근거 시트의 입력 칸과 같은 값)
RATES = [  # key, 항목, 값, 기준, 출처
    ("SAN", "산재보험료", 0.0356, "노무비 ×", S_INS + ": 건설업 3.56% (출퇴근재해 0.06% 포함)"),
    ("GOY", "고용보험료", 0.0115, "노무비 ×", S_INS + ": 실업급여 사업주 0.9% + 고용안정·직업능력개발 0.25% (150인 미만)"),
    ("SAFE", "산업안전보건관리비", 0.0311, "(재료비 + 노무비) ×", S_SAFE + ". 총공사금액 2천만원 이상 의무 계상"),
    ("GEN", "일반관리비", 0.05, "순공사원가 ×", S_RULE + ". 6% 이내에서 5% 적용"),
]
PROFIT = {"도비": 0.15, "용접": 0.15, "베이스": 0.15}   # 이윤율 전 공종 동일 15% (사용자 지정 2026-10-12)
WORKS = [("도비", DOBI, "진공건조로 2기 회전·베이스 안착(1일), 용접면 밀착(2일), 완성품 이동(1일), 작업구역 안전펜스"),
         ("용접", WELD, "접합면 사상(2일) 및 접합부 CO2 용접(3일)"),
         ("베이스", BASE, "기존 철판 조방 위 H빔 임시 베이스 설치 및 공사 후 회수, H-250×250")]
TARGETS = {"도비": 20_000_000, "용접": 25_000_000, "베이스": 18_000_000}

# ---------------------------------------------------------------- B안 (대안): 기존 조방 7×3m 개조 + 전체 90° 회전 후 하부 용접
# 200t 크레인 단가는 업체 견적이 없어 사용자 제시 100t(240만)·60t(170만)에서 톤당 증가분(1.75만/t)으로 외삽한 추정값.
R_PIPE = 350_000      # 플랜트배관공 (노임 320,480 × 약 1.09)
B_DOBI = [
    ("노무비", "도비 전문기공", "안착·밀착·90° 회전·회전 후 정리", "인·일", 21, 300_000, "A안 (6, 6, 3) + 회전일 6명 × 1일", "L",
     "(6 + 6 + 3 + 6)명 × 300,000원", "인원 " + U + " (2026-10-12: A안 + 6명 1일), 단가 " + SH + "", None),
    ("노무비", "신호수", "크레인 인양 신호", "인·일", 4, R_SIGNAL, "2명 × 2일 (안착일 + 회전일)", "L", "2명 × 2일 × 260,000원", AS, None),
    ("노무비", "안전담당 (관리감독자)", "인양구역 통제", "인·일", 3, R_SAFE, "1명 × 3일", "L", "1명 × 3일 × 320,000원", AS, None),
    ("장비", "크레인", "100TON (안착)", "대·일", 1, 2_400_000, "안착일", "E", "1대 × 1일 × 2,400,000원", U, None),
    ("장비", "크레인", "60TON (안착)", "대·일", 1, 1_700_000, "안착일", "E", "1대 × 1일 × 1,700,000원", U, None),
    ("장비", "크레인", "200TON급 (90° 회전 주 크레인)", "대·일", 1, 3_500_000, "회전일, 일체 약 86t", "E",
     "1대 × 1일 × 3,500,000원", U + " (사용자 제시 2026-10-12: 350만/일)", None),
    ("장비", "200TON 회송·웨이트 운반", "카운터웨이트 트레일러 왕복", "식", 1, 1_200_000, "대형 크레인 부대비", "E", "1식", AS, None),
    ("장비", "회전용 대형 줄걸이", "100t급 샤클·슬링·스프레더 손료", "식", 1, 800_000, "200t 단독 회전", "E", "1식", AS, None),
    ("장비", "아웃리거 받침 철판", "1,500×1,500×25t, 8매", "회", 2, 640_000, "안착일 + 회전일", "E", "2회 × 640,000원", AS, None),
    ("장비", "공구사용손료", "도비 전문장비·줄걸이구", "SET", 1, 3_000_000, "", "E", "1식", SH, None),
    ("재료비", "회전 착지 깔목·고무판", "목재 받침 300×300, 고무판", "식", 1, 900_000, "회전 후 착지면 보호", "M", "1식", AS, None),
    ("재료비", "소모품·잡자재", "와이어 보호대 등", "식", 1, 1_000_000, "", "M", "1식", SH, None),
    ("안전시설", "안전펜스 임대", "이동식 메쉬펜스 H1.8×W2.4m", "매·일", 22 * 11, 5_000, "22매 × 11일", "E", "A안과 같음", "수량 " + AS + ", 단가 " + U, None),
    ("안전시설", "펜스 설치·해체", "설치 인건비", "인·일", 2, 250_000, "", "L", "2명 × 1일 × 250,000원", U, None),
    ("안전시설", "펜스 운반", "1톤 화물 왕복", "회", 2, 260_000, "", "E", "2회 × 260,000원", AS, None),
    ("안전시설", "출입통제 안전장비", "", "식", 1, 380_000, "", "M", "1식", AS, None),
    expense(),
]
B_WELD = list(WELD)   # 용접 범위·인원 같음 (회전 후 하부 용접 포함)
B_JIG = [   # 사용자 제시(2026-10-12): 조방 개조안 약 800만원으로 베이스(약 1,800만원) 대체, 대신 200t급 크레인 1대 추가(B_도비)
    ("재료비", "보강보", "H-200급, 약 25m (현장 보유 강재 사용)", "식", 1, 0, "현장 보유분 사용, 자재비 미계상", "M", "현장 보유 강재 사용", U, None),
    ("재료비", "용접 소모품", "용접봉·가스·디스크", "식", 1, 420_000, "", "M", "1식", AS, None),
    ("재료비", "배관 복구 자재", "가스켓·볼트·너트", "식", 1, 380_000, "제거한 배관 재조립용", "M", "1식", AS, None),
    ("노무비", "제관사", "조방 보강보 취부", "인·일", 2, R_BOIL, "2명 × 1일", "L", "2명 × 1일 × 350,000원", AS, None),
    ("노무비", "용접사", "보강보 용접", "인·일", 2, R_WELD, "2명 × 1일", "L", "2명 × 1일 × 500,000원", AS, None),
    ("노무비", "조공", "보조", "인·일", 2, R_HELP, "2명 × 1일", "L", "2명 × 1일 × 200,000원", AS, None),
    ("노무비", "화재감시자", "화기작업 전담", "인·일", 1, R_FIRE, "1명 × 1일", "L", "1명 × 1일 × 250,000원", AS, None),
    ("노무비", "제관사", "바닥 접촉면 기구물·배관 제거·복원", "인·일", 4, R_BOIL, "2명 × 2일 (제거 1 + 복원 1)", "L",
     "2명 × 2일 × 350,000원", "사용자 지시 2026-10-12 (눕혀 작업하기 위한 기구물 제거·복원), 단가 " + AS, None),
    expense_pd(2 + 2 + 2 + 1 + 4, "개조 1일 7명 + 제거·복원 2명 × 2일 (용접기·절단기 임대는 용접 공종에 계상)"),
]
WORKS_B = [("B_도비", B_DOBI, "B안: 조방 안착(100t·60t) + 전체 90° 회전(200t 단독), 안전펜스"),
           ("B_용접", B_WELD, "B안: 사상, 3면 용접 → 회전 후 나머지 1면 용접"),
           ("B_조방", B_JIG, "B안: 기존 조방 7×3m 개조(현장 보유 보강보), 바닥 접촉면 기구물·배관 제거·복원")]
PROFIT.update({"B_도비": 0.15, "B_용접": 0.15, "B_조방": 0.15})
ALLWORKS = WORKS   # 2026-10-16 kevin: A안 단독 최종 견적. B안 데이터(WORKS_B)는 남겨 두되 싣지 않음


def basis(row):
    """산출근거 시트의 고객용 근거 문구."""
    g, item = row[0], row[1]
    if g == "노무비" or (g == "안전시설" and row[7] == "L"):
        if item == "화재감시자":
            return "%s, 단가는 숙련 기능인력 시장 단가" % '산업안전보건기준에 관한 규칙 제241조의2 (화기작업 화재감시자 배치)'
        if item == "도비 전문기공":
            return "인원·일수: 작업 범위·공정표 기준 / 단가: 중량물 도비 전문기공 시장 단가"
        if item == "신호수" or item.startswith("안전담당"):
            return "크레인 인양 작업 안전관리 인력, 단가: 시장 단가"
        return "인원·일수: 작업 범위·공정표 기준 / 단가: 플랜트 숙련 기능인력 시장 단가"
    if g == "장비":
        if item == "크레인":
            return "이동식 크레인 1일 임대 시장 단가 (운전원 포함)"
        if "손료" in item or "줄걸이" in item:
            return "중량물 작업 전문장비 사용 손료"
        return "장비 임대 시장 단가, 투입 수량·일수는 공정표 기준"
    if g == "재료비":
        if item.startswith("H빔") and "운임" not in item:
            return "H빔 1,882kg (베이스자재 시트), 공사 후 회수 조건 1식"
        if item == "보강보":
            return "현장 보유 강재 사용, 자재비 미계상"
        if "운임" in item:
            return "화물 운반 시장 단가"
        return "자재 시장 단가, 작업 범위 기준 1식"
    if g == "안전시설":
        if "임대" in item:
            return "작업구역 둘레 기준 수량 × 임대 시장 단가"
        if "운반" in item:
            return "화물 운반 시장 단가"
        return "작업구역 출입통제 자재, 1식"
    if g == "운반비":
        return "화물 운반 시장 단가 (반입·회수)"
    if g == "경비":
        return "투입 인·일 기준 현장 부대비" if row[4] == 1 and "인·일" in row[8] else "공종 현장 부대비 1식"
    return ""


def calc(rows, pr, no_oh=False):
    M = sum(r[4] * r[5] for r in rows if r[7] == "M")
    L = sum(r[4] * r[5] for r in rows if r[7] == "L")
    E = sum(r[4] * r[5] for r in rows if r[7] == "E")
    rt = {k: v for k, _, v, _, _ in RATES}
    if no_oh: rt = {k: 0 for k in rt}; pr = 0
    san, goy, safe = L * rt["SAN"], L * rt["GOY"], (M + L) * rt["SAFE"]
    net = M + L + E + san + goy + safe
    gen = net * rt["GEN"]
    prof = (L + E + san + goy + safe + gen) * pr
    pre = net + gen + prof
    return dict(M=M, L=L, E=E, san=san, goy=goy, safe=safe, net=net, gen=gen, prof=prof, pre=pre, tot=int(pre // 1000 * 1000))


def korean_won(n):
    digits = "영일이삼사오육칠팔구"
    units = ["", "십", "백", "천"]
    big = ["", "만", "억", "조"]
    out, i = "", 0
    while n > 0:
        chunk, n = n % 10000, n // 10000
        if chunk:
            s = ""
            for p in range(3, -1, -1):
                d = chunk // (10 ** p) % 10
                if d:
                    s += digits[d] + units[p]
            out = s + big[i] + out
        i += 1
    return out


wb = Workbook()
cover = wb.active
cover.title = "견적서"


def border_range(ws, r1, r2, c1, c2, fill=None):
    for rr in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            x = ws.cell(rr, c)
            x.border = BOX
            if fill: x.fill = fill


def setw(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def page(ws, landscape=False):
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.63   # 16mm
    ws.page_margins.top = ws.page_margins.bottom = 0.71    # 18mm
    ws.page_margins.header = ws.page_margins.footer = 0.35
    ws.sheet_view.showGridLines = False
    ws.oddFooter.left.text = f"&\"{FN}\"&8 " + SUP_LINE
    ws.oddFooter.right.text = f"&\"{FN}\"&8 &P / &N"


# ---------------------------------------------------------------- 적용 요율 (원가계산서 하단 표, 먼저 만들어 셀 주소 확보)
cs = wb.create_sheet("원가계산서", 1)
RATE_ROW = 25   # 원가계산서 본표(6~20행)·주석(22행) 아래
RATE_SRC = {"SAN": "2026년 고용노동부 보험료율 고시 (건설업)", "GOY": "2026년 고용노동부 보험료율 고시 (150인 미만)",
            "SAFE": "고용노동부 산업안전보건관리비 계상기준", "GEN": "6% 이내 적용 (국가계약법 시행규칙 제8조 기준)"}
cs.cell(RATE_ROW - 1, 1, "적용 요율").font = font(sz=11, sb=True)
for c, h in ((1, "항목"), (2, "적용 기준"), (4, "요율"), (5, "근거")):
    cs.cell(RATE_ROW, c, h)
cs.merge_cells(start_row=RATE_ROW, start_column=2, end_row=RATE_ROW, end_column=3)
cs.merge_cells(start_row=RATE_ROW, start_column=5, end_row=RATE_ROW, end_column=6)
for c in range(1, 7): head_cell(cs.cell(RATE_ROW, c))
RC = {}
r = RATE_ROW + 1
for key, nm, v, base, _ in RATES + [("PROF", "이윤", PROFIT["도비"], "(노무비 + 경비 + 일반관리비) ×", None)]:
    cs.cell(r, 1, nm); cs.cell(r, 2, base); x = cs.cell(r, 4, v); x.number_format = "0.00%"
    cs.cell(r, 5, RATE_SRC.get(key, "15% 이내 적용 (국가계약법 시행규칙 제8조 기준)"))
    cs.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    cs.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
    RC[key] = f"원가계산서!$D${r}"
    for c in range(1, 7):
        x = cs.cell(r, c); x.border = BOX; x.font = font(sz=9 if c == 5 else 10); x.alignment = CEN if c == 4 else LEFT
    cs.row_dimensions[r].height = 30
    r += 1
assert len(set(PROFIT.values())) == 1   # 이윤율은 전 공종 같은 값 한 칸
for nm, _, _ in ALLWORKS:
    RC["P_" + nm] = RC["PROF"]


# ---------------------------------------------------------------- 공종 시트
def header(ws, work, ncol=9):
    last = "ABCDEFGHI"[ncol - 1]
    ws.merge_cells(f"A1:{last}2"); ws["A1"] = "공 종 별   내 역 서"; ws["A1"].font = font(True, 16, C_NAVY); ws["A1"].alignment = CEN
    ws.merge_cells(f"A3:{last}3"); ws["A3"] = f"견적번호 {QNO}   |   견적일 {QDATE}   |   주식회사 금강변압기 귀중"
    ws["A3"].font = font(sz=9, color="555555"); ws["A3"].alignment = CEN
    ws.merge_cells(f"A5:{last}5"); ws["A5"] = "공 종 : " + work; ws["A5"].font = font(sz=12, sb=True); ws["A5"].alignment = LEFT


def table_head(ws, r):
    for c, h in zip(range(1, 10), ("구분", "품명", "규격", "단위", "수량", "단가", "금액", "비고", "비목")):
        x = ws.cell(r, c, h); head_cell(x)


ROWMAP, TOTREF, CALC = {}, {}, {}


def detail(name, rows, work):
    ws = wb.create_sheet(name)
    header(ws, work)
    table_head(ws, 7)
    r = first = 8
    groups = []
    for row in rows:
        g, item, spec, unit, q, p, note, cat = row[:8]
        if groups and groups[-1][0] == g: groups[-1][2] = r
        else: groups.append([g, r, r])
        ws.cell(r, 2, item); ws.cell(r, 3, spec); ws.cell(r, 4, unit)
        ws.cell(r, 5, q).font = font(color="0000FF"); ws.cell(r, 6, p).font = font(color="0000FF")
        ws.cell(r, 7, f"=E{r}*F{r}"); ws.cell(r, 8, note); ws.cell(r, 9, {"M": "재료비", "L": "노무비", "E": "경비"}[cat])
        ROWMAP.setdefault(name, []).append((r, row))
        r += 1
    last = r - 1
    for g, a, b in groups:
        if b > a: ws.merge_cells(start_row=a, start_column=1, end_row=b, end_column=1)
        ws.cell(a, 1, g)
    rng = f"$I${first}:$I${last}", f"$G${first}:$G${last}"
    pr = RC["P_" + name]
    summ = [
        ("재료비 계", f'=SUMIF({rng[0]},"재료비",{rng[1]})', "M"),
        ("노무비 계", f'=SUMIF({rng[0]},"노무비",{rng[1]})', "L"),
        ("경비 계", f'=SUMIF({rng[0]},"경비",{rng[1]})', "E"),
        ("산재보험료", "=G{L}*" + RC["SAN"], "노무비 × 3.56%"),
        ("고용보험료", "=G{L}*" + RC["GOY"], "노무비 × 1.15%"),
        ("산업안전보건관리비", "=(G{M}+G{L})*" + RC["SAFE"], "(재료비 + 노무비) × 3.11%"),
        ("순공사원가", "=SUM(G{M}:G{SAFE})", ""),
        ("일반관리비", "=G{NET}*" + RC["GEN"], "순공사원가 × 5%"),
        ("이윤", "=(G{L}+G{E}+G{SAN}+G{GOY}+G{SAFE}+G{GEN})*" + pr, f"(노무비 + 경비 + 일반관리비) × {PROFIT[name]*100:.1f}%"),
        ("단수정리", "=-MOD(G{NET}+G{GEN}+G{PROF},1000)", "천원 미만 절사"),
        ("공종 합계", "=G{NET}+G{GEN}+G{PROF}+G{RND}", "부가세 별도"),
    ]
    if name in NO_OH:
        for i in (3, 4, 5, 7, 8):
            summ[i] = (summ[i][0], "=0", "미적용")
    keys = ["M", "L", "E", "SAN", "GOY", "SAFE", "NET", "GEN", "PROF", "RND", "TOT"]
    pos = {k: r + i for i, k in enumerate(keys)}
    for i, (label, f, note) in enumerate(summ):
        rr = r + i
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=6)
        ws.cell(rr, 1, label); ws.cell(rr, 7, f.format(**pos)); ws.cell(rr, 8, note if len(note) > 1 else "")
        ws.merge_cells(start_row=rr, start_column=8, end_row=rr, end_column=9)
    tot = pos["TOT"]
    for rr in range(first, tot + 1):
        for c in range(1, 10):
            x = ws.cell(rr, c); x.border = BOX
            if x.font.name != FN: x.font = font()
            x.alignment = LEFT if c in (3, 8) else CEN
            if c in (5, 6, 7): x.number_format = WON; x.alignment = RIGHT
        ws.row_dimensions[rr].height = 20
    for rr in range(last + 1, tot + 1):
        for c in range(1, 10): ws.cell(rr, c).fill = GRAY
        ws.cell(rr, 1).font = font(rr in (pos["NET"], tot))
        ws.cell(rr, 7).font = font(rr in (pos["NET"], tot))
    for c in range(1, 10):
        ws.cell(tot, c).fill = NAVY; ws.cell(tot, c).font = font(True, color="FFFFFF")
    ws.cell(tot, 1).font = font(True, 11, "FFFFFF"); ws.cell(tot, 7).font = font(True, 11, "FFFFFF")
    fr = tot + 2
    for t in ["※ 수량·단가의 근거는 '산출근거' 시트에 정리함.",
              "※ 연장·야간·휴일 작업은 공정에 따라 당사가 수행함 (추가 정산 없음)."]:
        ws.merge_cells(start_row=fr, start_column=1, end_row=fr, end_column=9)
        ws.cell(fr, 1, t).font = font(sz=9, color="555555"); fr += 1
    setw(ws, {"A": 9, "B": 20, "C": 30, "D": 7, "E": 8, "F": 11, "G": 13, "H": 38, "I": 7})
    page(ws)
    ws.print_title_rows = "7:7"
    TOTREF[name] = (f"{name}!$G$", pos)
    CALC[name] = calc(rows, PROFIT[name], name in NO_OH)


for nm, rows, work in ALLWORKS:
    detail(nm, rows, work)

# 협의 조정 (2026-10-19 kevin: 최종 견적은 VAT 포함 70,000,000원 그대로)
GRAND_TARGET = 70_000_000
SUBTOTAL = sum(CALC[n]["tot"] for n, _, _ in WORKS)
def supply_for_grand(t):
    s_ = round(t / 1.1)
    while s_ + round(s_ / 10) > t: s_ -= 1
    while s_ + 1 + round((s_ + 1) / 10) <= t: s_ += 1
    return s_
ADJ = supply_for_grand(GRAND_TARGET) - SUBTOTAL
ADJ_LABEL = "협의 조정"
ADJ_NOTE = "발주처 협의 금액 반영"

# ---------------------------------------------------------------- 원가계산서
cs.merge_cells("A1:F2"); cs["A1"] = "공 사 원 가 계 산 서"; cs["A1"].font = font(True, 16, C_NAVY); cs["A1"].alignment = CEN
cs.merge_cells("A3:F3"); cs["A3"] = "공사명 : 진공건조로 회전·안착 및 접합부 용접 공사   |   " + QNO; cs["A3"].font = font(sz=9, color="555555"); cs["A3"].alignment = CEN
for c, h in enumerate(["비목", "도비", "용접", "베이스", "합계", "구성비"], 1):
    x = cs.cell(5, c, h); head_cell(x)
lines = [("재료비", "M"), ("노무비", "L"), ("경비", "E"), ("산재보험료", "SAN"), ("고용보험료", "GOY"),
         ("산업안전보건관리비", "SAFE"), ("순공사원가", "NET"), ("일반관리비", "GEN"), ("이윤", "PROF"),
         ("단수정리", "RND"), ("공종 합계", "TOT")]
r = 6
for label, k in lines:
    cs.cell(r, 1, label)
    for j, (nm, _, _) in enumerate(WORKS):
        pre, pos = TOTREF[nm]
        x = cs.cell(r, 2 + j, f"={pre}{pos[k]}"); x.font = font(color="008000")
    cs.cell(r, 5, f"=SUM(B{r}:D{r})")
    cs.cell(r, 6, f"=E{r}/$E${6 + len(lines) - 1}")
    for c in range(1, 7):
        x = cs.cell(r, c); x.border = BOX
        if c >= 2: x.number_format = WON if c < 6 else "0.0%"; x.alignment = RIGHT
        else: x.alignment = CEN
        if k in ("NET", "TOT"): x.fill = GRAY; x.font = font(True, color=x.font.color.rgb if c in (2, 3, 4) else "000000")
    cs.row_dimensions[r].height = 22
    r += 1
SUMROW = r - 1
cs.cell(r, 1, ADJ_LABEL); cs.cell(r, 5, ADJ).font = font(color="0000FF"); ADJROW = r; r += 1
cs.cell(r, 1, "공급가액"); cs.cell(r, 5, f"=E{SUMROW}+E{ADJROW}"); SUPROW = r; r += 1
cs.cell(r, 1, "부가가치세 (10%)"); cs.cell(r, 5, f"=ROUND(E{SUPROW}*0.1,0)"); VATROW = r; r += 1
cs.cell(r, 1, "총 공사금액"); cs.cell(r, 5, f"=E{SUPROW}+E{VATROW}"); GRAND = r
for rr in (ADJROW, SUPROW, VATROW, GRAND):
    for c in range(1, 7):
        x = cs.cell(rr, c); x.border = BOX; x.font = font(rr == GRAND); x.alignment = CEN if c == 1 else RIGHT
        if c >= 2: x.number_format = WON
        if rr == SUPROW: x.fill = GRAY; x.font = font(True)
        if rr == GRAND: x.fill = NAVY; x.font = font(True, color="FFFFFF")
cs.cell(ADJROW, 6, ADJ_NOTE).font = font(sz=8, color="555555"); cs.cell(ADJROW, 6).alignment = LEFT
cs.cell(GRAND + 2, 1, "요율: 산재 3.56%·고용 1.15% (노무비 기준), 산업안전보건관리비 3.11% ((재료비 + 노무비) 기준), 일반관리비 5% (순공사원가 기준), 이윤 15%. 근거는 아래 적용 요율 표. 베이스 공종은 회수조건 임시 구조물로 재료비·운반비·노무비만 계상함.").font = font(sz=9, color="555555")
cs.merge_cells(start_row=GRAND + 2, start_column=1, end_row=GRAND + 2, end_column=6)
cs.cell(GRAND + 2, 1).alignment = LEFT; cs.row_dimensions[GRAND + 2].height = 30
setw(cs, {"A": 20, "B": 15, "C": 15, "D": 15, "E": 16, "F": 9})
page(cs)

# ---------------------------------------------------------------- 견적서 (엔엠솔루션 양식)
sup_val = SUBTOTAL + ADJ
grand_val = sup_val + round(sup_val / 10)
WORKNAME = "진공건조로 2분할체 회전·안착, 임시 베이스 제작, 접합부 사상·용접 공사"
c = cover
setw(c, {"A": 9, "B": 13, "C": 13, "D": 22, "E": 7, "F": 6, "G": 12, "H": 15, "I": 24})


def merge_put(ws, r, c1, c2, v, f=None, al=LEFT, fill=None, border=None, fmt=None):
    if c2 > c1: ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    x = ws.cell(r, c1, v); x.font = f or font(); x.alignment = al
    for cc in range(c1, c2 + 1):
        y = ws.cell(r, cc)
        if fill: y.fill = fill
        if border: y.border = border
    if fmt: x.number_format = fmt
    return x


# 1. 머리
merge_put(c, 1, 1, 4, "N M   S O L U T I O N", font(sz=9, color=C_ORANGE, sb=True))
merge_put(c, 2, 1, 4, "견  적  서", font(True, 22, C_NAVY)); c.merge_cells("A2:D3"); c.row_dimensions[2].height = 20; c.row_dimensions[3].height = 20
c.cell(1, 8, "견적일자").font = font(color=C_SUB); c.cell(1, 8).alignment = RIGHT
c.cell(1, 9, QDATE).font = font(); c.cell(1, 9).alignment = RIGHT
c.cell(2, 8, "견적번호").font = font(color=C_SUB); c.cell(2, 8).alignment = RIGHT
c.cell(2, 9, QNO).font = font(); c.cell(2, 9).alignment = RIGHT
for cc in range(1, 10): c.cell(4, cc).border = Border(bottom=thin)
c.row_dimensions[4].height = 8
# 2. 수신 · 공급자 2단
c.cell(5, 1, "수신").font = font(sz=12, sb=True); c.cell(5, 6, "공급자").font = font(sz=12, sb=True); c.row_dimensions[5].height = 22
recv = [("업체명", "주식회사 금강변압기"), ("담당자", ""), ("연락처", ""),
        ("현장 주소", "경기 화성시 마도면 마도공단로1길 214 (사내 마당)"), ("공사명", WORKNAME), ("", "")]
supl = [("상호", f"{SUP['name']} ({SUP['en']})"), ("대표", SUP["ceo"]), ("주소", SUP["addr"]),
        ("전화", f"대표 {SUP['tel']}  /  직통 {SUP['mobile']}"), ("홈페이지", SUP["web"]), ("사업자등록번호", SUP["bizno"])]
for i, ((lk, lv), (rk, rv)) in enumerate(zip(recv, supl)):
    r = 6 + i
    c.cell(r, 1, lk).font = font(color=C_SUB); c.cell(r, 1).alignment = LEFT
    merge_put(c, r, 2, 4, lv)
    merge_put(c, r, 6, 7, rk, font(color=C_SUB))
    merge_put(c, r, 8, 9, rv)
    c.row_dimensions[r].height = 30 if (len(lv) > 24 or len(rv) > 24) else 20
    for cc in (1, 2, 3, 4, 6, 7, 8, 9): c.cell(r, cc).border = Border(bottom=thin)
# 3. 합계 박스
r = 13
merge_put(c, r, 1, 9, f"견적금액  일금 {korean_won(sup_val)}원정  (₩ {sup_val:,}, VAT 별도)",
          font(True, 13, C_NAVY), CEN, border=Border(top=navy2, bottom=navy2))
c.row_dimensions[r].height = 36
AMT_TEXT_ROW = r
# 4. 견적 내역
r = 15
c.cell(r, 1, "견적 내역").font = font(sz=12, sb=True); c.row_dimensions[r].height = 22; r += 1
hdr = [(1, 1, "번호"), (2, 3, "품명"), (4, 4, "규격"), (5, 5, "수량"), (6, 6, "단위"), (7, 7, "단가"), (8, 8, "금액"), (9, 9, "비고")]
for a, b_, h in hdr:
    if b_ > a: c.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b_)
    c.cell(r, a, h)
for cc in range(1, 10): head_cell(c.cell(r, cc))
c.row_dimensions[r].height = 22
HEADROW = r; r += 1
groups = [("도비", "도비 작업: 회전·베이스 안착, 밀착, 완성품 이동 (250t 1대), 안전펜스"),
          ("용접", "용접 작업: 접합면 사상, CO2 용접, 화재감시"),
          ("베이스", "베이스 작업: 기존 철판 조방 위 H빔 베이스 설치·회수")]
sub_rows = []
no = 0
for key, gtitle in groups:
    merge_put(c, r, 2, 9, gtitle, font(sb=True), LEFT, border=BOX); c.cell(r, 1).border = BOX; c.row_dimensions[r].height = 20; r += 1
    pre, pos = TOTREF[key]
    for dr, row in ROWMAP[key]:
        no += 1
        c.cell(r, 1, no)
        merge_put(c, r, 2, 3, row[1]); c.cell(r, 4, row[2])
        c.cell(r, 5, f"={key}!E{dr}"); c.cell(r, 6, f"={key}!D{dr}")
        c.cell(r, 7, f"={key}!F{dr}"); c.cell(r, 8, f"={key}!G{dr}"); c.cell(r, 9, row[6])
        r += 1
    if key not in NO_OH:
        no += 1
        c.cell(r, 1, no); merge_put(c, r, 2, 3, "법정경비·일반관리비·이윤")
        c.cell(r, 4, "산재·고용보험, 산업안전보건관리비, 일반관리비, 이윤, 단수정리")
        c.cell(r, 5, 1); c.cell(r, 6, "식"); c.cell(r, 7, f"=H{r}")
        c.cell(r, 8, f"={pre}{pos['TOT']}-{pre}{pos['M']}-{pre}{pos['L']}-{pre}{pos['E']}"); c.cell(r, 9, "원가계산서 참조"); r += 1
    merge_put(c, r, 1, 7, f"{key} 소계", font(sb=True), RIGHT, fill=GRAY)
    x = c.cell(r, 8, f"={pre}{pos['TOT']}"); x.font = font(sb=True); c.cell(r, 9, f"{key} 시트 참조").font = font(color=C_SUB)
    for cc in range(1, 10): c.cell(r, cc).fill = GRAY
    sub_rows.append(r); r += 1
first_item = HEADROW + 1
tot_rows = [("공종 합계", "=" + "+".join(f"H{x}" for x in sub_rows)), (ADJ_LABEL, ADJ), ("공급가액", f"=H{r}+H{r+1}"),
            ("부가세 (10%)", f"=ROUND(H{r+2}*0.1,0)"), ("합계 (VAT 포함)", f"=H{r+2}+H{r+3}")]
for j, (k, f) in enumerate(tot_rows):
    last = j == len(tot_rows) - 1
    merge_put(c, r, 1, 7, k, font(True, 13 if last else 10, "FFFFFF" if last else C_TEXT), RIGHT, fill=NAVY if last else GRAY)
    x = c.cell(r, 8, f); x.font = font(True, 13 if last else 10, "FFFFFF" if last else C_TEXT)
    for cc in range(1, 10): c.cell(r, cc).fill = NAVY if last else GRAY
    if k == ADJ_LABEL: c.cell(r, 9, ADJ_NOTE)
    c.row_dimensions[r].height = 26 if last else 20
    r += 1
SUPROW_C, VATROW_C, GRANDROW_C = r - 3, r - 2, r - 1
for rr in range(first_item, r):
    for cc in range(1, 10):
        x = c.cell(rr, cc); x.border = BOX
        if x.font.name not in (FN, FN_SB): x.font = font()
        if cc in (5, 7, 8): x.number_format = WON; x.alignment = RIGHT
        elif cc in (1, 6): x.alignment = CEN
        elif x.alignment.horizontal is None: x.alignment = LEFT
    if c.row_dimensions[rr].height is None: c.row_dimensions[rr].height = 20
# 5. 특기사항
r += 1
c.cell(r, 1, "특기사항").font = font(sz=12, sb=True); c.row_dimensions[r].height = 22; r += 1
NOTES = [
    "1. 공사 기간: 2026.10.21 ~ 2026.11.03 (가정 일정, 작업일 12일, 공정표 시트 참조). 날짜는 가정이며 착공일은 발주처 협의로 확정함.",
    "2. 결제 조건: 선수금 30% (계약 시), 중도금 30% (회전·안착 완료 시), 잔금 40% (완성품 이동 완료 후). VAT 포함 기준 21,000,000원 / 21,000,000원 / 28,000,000원.",
    "3. 견적 유효기간: 발행일로부터 30일.",
    "4. 포함: 임시 베이스 설치·회수, 이동식 크레인(250t 1대)·도비, 접합면 사상, CO2 용접, 화재감시, 완성품 지정 장소 이동, 안전펜스·출입통제.",
    "5. 제외: 전기 수전, 배관·닥트 해체·복구, 보온, 도장, 비파괴검사(PT/UT), 시운전. 비파괴검사는 발주처 요구 시 별도 정산함.",
    "6. 크레인 사양·배치는 크레인 업체와 협의해 확정하며, 크레인 변경에 따른 추가비용은 없음.",
    "7. 연장·야간·휴일 작업은 공정에 따라 당사가 수행하며 추가 정산하지 않음. 기상(우천 시 용접 불가)·발주처 사유에 따른 대기는 별도 협의함.",
    "8. 작업구역(약 10×15m)은 이동식 안전펜스로 통제하며, 발주처 안전관리 규정과 출입 절차를 따름.",
]
for t in NOTES:
    merge_put(c, r, 1, 9, t); c.row_dimensions[r].height = 30 if len(t) > 60 else 20; r += 1
# 6. 바닥
r += 1
merge_put(c, r, 1, 9, "위와 같이 견적합니다.", font(), CEN); c.row_dimensions[r].height = 22; r += 1
_y, _m, _d = QDATE.split("-")
merge_put(c, r, 1, 9, f"{_y}년 {int(_m)}월 {int(_d)}일", font(), CEN); c.row_dimensions[r].height = 22; r += 1
merge_put(c, r, 1, 9, f"{SUP['name']}   대표 {SUP['ceo']}   (인)", font(sz=12, sb=True), CEN); c.row_dimensions[r].height = 26; r += 2
merge_put(c, r, 1, 9, SUP_LINE, font(sz=8, color=C_SUB), CEN, border=Border(top=thin)); c.row_dimensions[r].height = 18
page(c)
c.print_title_rows = f"{HEADROW}:{HEADROW}"

# ---------------------------------------------------------------- 산출근거
gs = wb.create_sheet("산출근거", 2)
gs["A1"] = "견적 산출근거"; gs["A1"].font = font(True, 14, C_NAVY)
gs["A2"] = "금액은 공종 시트에서 연동됨. 인원·일수는 작업 범위와 공정표 기준, 단가는 동종 플랜트 현장 시장 단가 기준으로 산정함."; gs["A2"].font = font(sz=9, color="555555")
for cc, h in enumerate(["공종", "품명", "규격", "산출식", "금액(원)", "근거"], 1):
    x = gs.cell(4, cc, h); head_cell(x)
r = 5
for nm, _, _ in ALLWORKS:
    for dr, row in ROWMAP[nm]:
        item, spec, f_ = row[1], row[2], row[8]
        src = basis(row)
        gs.cell(r, 1, nm); gs.cell(r, 2, item); gs.cell(r, 3, spec); gs.cell(r, 4, f_)
        x = gs.cell(r, 5, f"={nm}!G{dr}"); x.font = font(color="008000"); x.number_format = WON
        gs.cell(r, 6, src)
        r += 1
    pre, pos = TOTREF[nm]
    for label, k in ((("법정경비·일반관리비·이윤", None),) if nm not in NO_OH else ()) + (("공종 합계", "TOT"),):
        gs.cell(r, 2, label).font = font(True)
        if k: x = gs.cell(r, 5, f"={pre}{pos[k]}")
        else: x = gs.cell(r, 5, f"={pre}{pos['TOT']}-SUM({pre}{pos['M']}:$G${pos['E']})")
        x.font = font(True); x.number_format = WON
        gs.cell(r, 6, "법정 요율·일반관리비·이윤 (원가계산서 하단 참조)" if not k else ("부가세 별도" if nm in TARGETS else "B안 (대안), 부가세 별도"))
        for cc in range(1, 7): gs.cell(r, cc).fill = GRAY
        r += 1
    r += 1
for rr in range(5, r):
    if gs.cell(rr, 1).value is None and gs.cell(rr, 2).value is None: continue
    for cc in range(1, 7):
        x = gs.cell(rr, cc); x.border = BOX
        if x.font.name != FN: x.font = font(x.font.b)
        x.alignment = RIGHT if cc == 5 else LEFT
        x.alignment = Alignment(horizontal=x.alignment.horizontal, vertical="center", wrap_text=True)
setw(gs, {"A": 8, "B": 18, "C": 30, "D": 44, "E": 14, "F": 52})
page(gs, landscape=True)
gs.print_title_rows = "4:4"

# ---------------------------------------------------------------- 베이스 자재
bs = wb.create_sheet("베이스자재")
bs["A1"] = "베이스 자재: H-250×250×9×14 (회수조건)"; bs["A1"].font = font(True, 14, C_NAVY)
for cc, h in enumerate(["길이(m)", "수량(EA)", "총 길이(m)", "단위중량(kg/m)", "중량(kg)"], 1):
    x = bs.cell(3, cc, h); head_cell(x)
for i, (ln, q) in enumerate([(10, 2), (3, 2)]):
    rr = 4 + i
    bs.cell(rr, 1, ln).font = font(color="0000FF"); bs.cell(rr, 2, q).font = font(color="0000FF")
    bs.cell(rr, 3, f"=A{rr}*B{rr}"); bs.cell(rr, 4, 72.4).font = font(color="0000FF"); bs.cell(rr, 5, f"=C{rr}*D{rr}")
    bs.cell(rr, 5).number_format = '#,##0'
bs.cell(6, 2, "합계").font = font(True); bs.cell(6, 3, "=SUM(C4:C5)"); bs.cell(6, 5, "=SUM(E4:E5)"); bs.cell(6, 5).number_format = '#,##0'
border_range(bs, 4, 6, 1, 5)
for cc in range(1, 6): bs.cell(6, cc).fill = GRAY
for cc, h in enumerate(["구분", "중량(kg)", "수량", "금액(원)", "비고"], 1):
    x = bs.cell(9, cc, h); head_cell(x)
bs.cell(10, 1, "H빔 (회수조건)"); bs.cell(10, 2, "=E6").number_format = '#,##0'
bs.cell(10, 3, "1식")
bs.cell(10, 4, "=베이스!G8").number_format = WON; bs.cell(10, 5, "공사 후 회수, 사용 기간 손료 개념")
border_range(bs, 9, 10, 1, 5)
for cc in range(1, 6):
    x = bs.cell(10, cc)
    if x.font.name != FN: x.font = font()
    x.alignment = LEFT if cc in (1, 5) else RIGHT
bs.cell(12, 1, "기존 철판 조방 위에 10m H빔 2본을 길이 방향으로 놓고, 양 끝에 3m H빔을 직각으로 용접함. 단위중량 72.4 kg/m는 KS H-250×250×9×14 공칭값. 임시 구조물로 방청·도장 미적용.").font = font(sz=9, color="555555")
setw(bs, {"A": 18, "B": 16, "C": 16, "D": 16, "E": 30})
page(bs)

# ---------------------------------------------------------------- 공정표
ps = wb.create_sheet("공정표")
import datetime as _dt
WDAYS = []
_d = _dt.date(2026, 10, 21)
while len(WDAYS) < 12:
    if _d.weekday() != 6: WDAYS.append(f"{_d.month}/{_d.day}\n" + "월화수목금토일"[_d.weekday()])
    _d += _dt.timedelta(days=1)
ps["A1"] = "공정표 (가정 일정 2026.10.21 ~ 11.03, 작업일 12일)"; ps["A1"].font = font(True, 14, C_NAVY)
steps = [("안전펜스 설치·작업구역 통제", "조공 2명", 1, 1), ("베이스 설치 (H빔 배치·용접)", "시공자 4명", 1, 3),
         ("회전·베이스 안착 (250t)", "도비 6명, 신호수 2, 안전담당", 4, 4),
         ("접합면 사상", "사상 2, 조공 2, 화재감시 1", 5, 6), ("용접면 밀착", "도비 6명·3명, 안전담당", 7, 8),
         ("접합부 용접", "용접사 3, 고급기술자 2, 화재감시 1", 9, 11),
         ("완성품 이동 (250t)·베이스 회수", "도비 6명, 신호수 2, 안전담당", 12, 12), ("펜스 해체·정리", "조공 2명", 12, 12)]
ps.cell(3, 1, "공정").font = font(True); ps.cell(3, 2, "투입 인원").font = font(True)
for d in range(1, 13):
    x = ps.cell(3, 2 + d, WDAYS[d - 1]); x.font = font(True, sz=9); x.alignment = CEN
    ps.column_dimensions[x.column_letter].width = 5.5
BAR = PatternFill("solid", fgColor="E8742A")
for i, (nm, crew, a, b) in enumerate(steps):
    rr = 4 + i
    ps.cell(rr, 1, nm).font = font(); ps.cell(rr, 2, crew).font = font()
    for d in range(1, 13):
        if a <= d <= b: ps.cell(rr, 2 + d).fill = BAR
border_range(ps, 3, 3 + len(steps), 1, 14)
for cc in range(1, 15): ps.cell(3, cc).fill = GRAY; ps.cell(3, cc).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ps.row_dimensions[3].height = 30
setw(ps, {"A": 30, "B": 32})
ps.cell(5 + len(steps), 1, "가정 일정: 2026-10-21(수) 착공 ~ 11-03(화) 완료, 일요일 휴무 기준 작업일 12일. 날짜는 가정이며 착공일은 발주처 협의로 확정함.").font = font(sz=9, color="555555")
page(ps, landscape=True)


# A·B 비교 시트는 2026-10-16 kevin 지시(A안 단독 최종 견적)로 뺌

wb.save(OUT)
for n, _, _ in ALLWORKS:
    k = CALC[n]
    print(n, {a: round(b) for a, b in k.items()})
print("supply", sup_val, "grand", grand_val, korean_won(grand_val))
print(OUT)
