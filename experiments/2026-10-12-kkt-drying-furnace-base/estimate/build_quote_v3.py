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
FN = "맑은 고딕"
WON = '#,##0;-#,##0;"-"'
thin = Side(style="thin", color="000000")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
GRAY = PatternFill("solid", fgColor="F2F2F2")
NAVY = PatternFill("solid", fgColor="1F3864")
YEL = PatternFill("solid", fgColor="FFF2CC")
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")

QNO = "NMS-2610-KKT-01"
QDATE = "2026-10-12"

def font(b=False, sz=10, color="000000"):
    return Font(name=FN, bold=b, size=sz, color=color)

# ---------------------------------------------------------------- 출처
S_WAGE1 = "대한건설협회 2026년 상반기 적용 건설업 임금실태조사 보고서(2026.1.1 적용)"
S_WAGE2 = "대한건설협회 2026년 하반기 적용 조사(2026.9.1 적용) 132개 직종 평균 282,737원, 직전 대비 +0.98%"
S_INS = "고용노동부 2026년 보험료율 고시 (대한전문건설협회 안내자료)"
S_SAFE = "고용노동부 건설업 산업안전보건관리비 계상 및 사용기준 (5억 미만 건축공사 3.11%)"
S_RULE = "국가계약법 시행규칙 제8조 (일반관리비 6%·이윤 15% 상한, 민간공사 참고)"
U = "사용자 제시"
SH = "서해테크 도비 견적서(2026-10-10) 단가"
VQ = "공급처 H빔 견적"
MK = "현장 실거래 수준 적용 (사용자 지시: 용접사 상향, 고급기술자=제관사 수준)"
AS = "가정값 (업체 견적으로 확인 필요)"
URLS = [
    ("시중노임 상반기", "https://www.kmcca.or.kr/images/pegs/down_file/2026%EB%85%84%20%EC%83%81%EB%B0%98%EA%B8%B0%EC%A0%81%EC%9A%A9%20%EA%B1%B4%EC%84%A4%EC%97%85%20%EC%9E%84%EA%B8%88%EC%8B%A4%ED%83%9C%EC%A1%B0%EC%82%AC%20%EB%B3%B4%EA%B3%A0%EC%84%9C.pdf"),
    ("시중노임 하반기", "http://www.conslove.co.kr/news/articleView.html?idxno=89479"),
    ("산재·고용보험 요율", "https://www.kosca21.or.kr/km_board/upload/km_b_notice/1767346180_02.pdf"),
    ("산업안전보건관리비", "https://biz.newdaily.co.kr/site/data/html/2026/05/07/2026050700116.html"),
    ("원가계산 제비율", "https://04noti.com/construction-cost-ratio-guide/"),
    ("2026 4대보험(연금·건강)", "https://calctools.co.kr/blog/daily-worker-insurance"),
    ("화재감시자 기준", "산업안전보건기준에 관한 규칙 제241조의2"),
]

# 시중노임 (원/일, 2026.1.1 적용)
WAGE = {"용접공": 282_536, "플랜트용접공": 292_640, "플랜트제관공": 266_029, "특별인부": 226_122,
        "보통인부": 172_068, "비계공": 281_939, "철골공": 252_641}

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
    ("노무비", "도비 전문기공", "회전·안착·밀착 이동", "인·일", 15, 300_000, "안착일 6명 + 밀착 이동 6명·3명 (6, 6, 3)", "L",
     "(6 + 6 + 3)명 × 300,000원", "인원 " + U + " (2026-10-12), 단가 " + SH, ("비계공", WAGE["비계공"])),
    ("노무비", "신호수", "크레인 인양 신호", "인·일", 2, R_SIGNAL, "2명 × 1일 (안착일)", "L",
     "2명 × 1일 × 260,000원", AS, ("특별인부", WAGE["특별인부"])),
    ("노무비", "안전담당 (관리감독자)", "작업 전 TBM·인양구역 통제", "인·일", 3, R_SAFE, "1명 × 3일", "L",
     "1명 × 3일 × 320,000원", AS, None),
    ("장비", "크레인", "100TON", "대·일", 1, 2_400_000, "안착일 1일", "E", "1대 × 1일 × 2,400,000원", U, None),
    ("장비", "크레인", "60TON", "대·일", 1, 1_700_000, "안착일 1일", "E", "1대 × 1일 × 1,700,000원", U, None),
    ("장비", "아웃리거 받침 철판", "1,500×1,500×25t, 8매", "식", 1, 640_000, "아스팔트 지반 보호", "E", "8매 × 80,000원", AS, None),
    ("장비", "밀착 이동 장비", "유압잭·윈치·스키드", "SET", 1, 2_000_000, "밀착 이동 2일", "E", "1식 (2일)", AS, None),
    ("장비", "공구사용손료", "도비 전문장비·줄걸이구", "SET", 1, 3_000_000, "", "E", "1식", SH, None),
    ("재료비", "소모품·잡자재", "목재 받침·와이어 보호대 등", "식", 1, 1_000_000, "", "M", "1식", SH, None),
    ("안전시설", "안전펜스 임대", "이동식 메쉬펜스 H1.8×W2.4m", "매·일", 22 * 11, 5_000, "22매(둘레 50m + 출입문) × 11일", "E",
     "작업구역 10×15m, 둘레 50m ÷ 2.4m = 21매 + 출입문 1매 → 22매 × 11일 × 5,000원", "수량 " + AS + ", 단가 " + U, None),
    ("안전시설", "펜스 설치·해체", "설치 인건비", "인·일", 2, 250_000, "설치 2명 × 0.5일 + 해체 2명 × 0.5일", "L",
     "2명 × 1일 × 250,000원", U, ("보통인부", WAGE["보통인부"])),
    ("안전시설", "펜스 운반", "1톤 화물 왕복", "회", 2, 260_000, "반입 1 + 반출 1", "E", "2회 × 260,000원", AS, None),
    ("안전시설", "출입통제 안전장비", "", "식", 1, 380_000, "", "M", "1식", AS, None),
    expense(),
]
WELD = [
    ("노무비", "용접사", "CO2 용접 (접합부)", "인·일", 9, R_WELD, "3명 × 3일", "L",
     "3명 × 3일 × 500,000원", "인원·일수 " + U + ", 단가 " + MK, ("플랜트용접공", WAGE["플랜트용접공"])),
    ("노무비", "고급기술자", "제관사 수준 (취부·치수 관리)", "인·일", 6, R_BOIL, "2명 × 3일", "L",
     "2명 × 3일 × 350,000원", "인원·일수 " + U + ", 단가 " + MK, ("플랜트제관공", WAGE["플랜트제관공"])),
    ("노무비", "사상 기술자", "접합면 그라인딩", "인·일", 4, R_GRIND, "2명 × 2일", "L",
     "2명 × 2일 × 300,000원", "인원·일수 " + U, ("특별인부", WAGE["특별인부"])),
    ("노무비", "조공", "사상·용접 보조", "인·일", 4, R_HELP, "2명 × 2일", "L",
     "2명 × 2일 × 200,000원", "인원·일수 " + U, ("보통인부", WAGE["보통인부"])),
    ("노무비", "화재감시자", "용접·사상 화기작업 전담", "인·일", 5, R_FIRE, "1명 × 5일 (사상 2 + 용접 3)", "L",
     "1명 × 5일 × 250,000원", "산업안전보건기준에 관한 규칙 제241조의2, 단가 " + AS, ("특별인부", WAGE["특별인부"])),
    ("장비", "CO2 용접기 대여", "500A급", "대·일", 9, 400_000, "3대 × 3일", "E", "3대 × 3일 × 400,000원", U, None),
    ("장비", "그라인더·환기·조명", "송풍기, 투광등 포함", "SET", 1, 800_000, "", "E", "1식", AS, None),
    ("재료비", "용접 소모품", "CO2 가스·와이어·디스크·케이블·잡자재", "식", 1, 3_680_000, "", "M", "1식 (용접재 3,000,000 + 전원 케이블·잡자재 680,000)", AS, None),
    ("장비", "작업발판 장비 임대", "", "식", 1, 1_000_000, "", "E", "1식 (사용자 제시 약 100만원)", U, None),
    ("장비", "가스절단기 임대", "산소·LPG, 2대 공정기간 상시", "대·일", 22, 100_000, "2대 × 11일", "E", "2대 × 11일 × 100,000원", U, None),
    expense_pd(9 + 6 + 4 + 4 + 5, "용접 9 + 고급기술자 6 + 사상 4 + 조공 4 + 화재감시 5"),
]
BASE = [
    ("재료비", "H빔 (회수조건)", "H-250×250×9×14, 10m 4·5m 16·0.6m 30EA", "kg", 9_991, 700, "공사 후 회수, 중고 구매가 9,577,783원", "M",
     "9,991kg × 700원 (중고 구매 9,577,783원 − 회수 고철가치 약 400만원 + 회수 손료)", VQ + " + " + U + " (회수조건 약 700만원)", None),
    ("재료비", "H빔 반입 운임", "", "식", 1, 160_000, "", "M", "1식", VQ, None),
    ("재료비", "회수 반출 운임", "", "식", 1, 160_000, "반입 운임과 같게 적용", "M", "1식", AS, None),
    ("재료비", "용접 소모품", "용접봉·가스·디스크", "식", 1, 1_000_000, "", "M", "1식", AS, None),
    ("재료비", "레벨 조정재", "라이너·심 플레이트", "식", 1, 420_000, "베이스 수평 ±2mm", "M", "1식", AS, None),
    ("노무비", "제관사", "절단·취부", "인·일", 6, R_BOIL, "2명 × 3일", "L",
     "2명 × 3일 × 350,000원", "인원 " + U + "(4명), 직종 구성 가정", ("플랜트제관공", WAGE["플랜트제관공"])),
    ("노무비", "용접사", "베이스 조립 용접", "인·일", 3, R_WELD, "1명 × 3일", "L",
     "1명 × 3일 × 500,000원", "인원 " + U + "(4명), 직종 구성 가정", ("플랜트용접공", WAGE["플랜트용접공"])),
    ("노무비", "조공", "자재 정리·보조", "인·일", 3, R_HELP, "1명 × 3일", "L",
     "1명 × 3일 × 200,000원", "인원 " + U + "(4명), 직종 구성 가정", ("보통인부", WAGE["보통인부"])),
    ("노무비", "화재감시자", "용접·절단 화기작업 전담", "인·일", 3, R_FIRE, "1명 × 3일", "L",
     "1명 × 3일 × 250,000원", "산업안전보건기준에 관한 규칙 제241조의2, 단가 " + AS, ("특별인부", WAGE["특별인부"])),
    ("장비", "지게차", "자재 하역·조립", "일", 2, 300_000, "", "E", "2일 × 300,000원", U, None),
    ("장비", "지게차", "회수 시 해체·상차", "일", 1, 300_000, "", "E", "1일 × 300,000원", U + " 단가 준용", None),
    ("장비", "CO2 용접기 대여", "500A급", "대·일", 3, 400_000, "1대 × 3일", "E", "1대 × 3일 × 400,000원", U + " 단가 준용", None),
    expense_pd(6 + 3 + 3 + 3, "제관 6 + 용접 3 + 조공 3 + 화재감시 3 (가스절단기는 용접 공종에 계상)"),
]

# 요율 (요율근거 시트의 입력 칸과 같은 값)
RATES = [  # key, 항목, 값, 기준, 출처
    ("SAN", "산재보험료", 0.0356, "노무비 ×", S_INS + ": 건설업 3.56% (출퇴근재해 0.06% 포함)"),
    ("GOY", "고용보험료", 0.0115, "노무비 ×", S_INS + ": 실업급여 사업주 0.9% + 고용안정·직업능력개발 0.25% (150인 미만)"),
    ("SAFE", "산업안전보건관리비", 0.0311, "(재료비 + 노무비) ×", S_SAFE + ". 총공사금액 2천만원 이상 의무 계상"),
    ("GEN", "일반관리비", 0.05, "순공사원가 ×", S_RULE + ". 6% 이내에서 5% 적용"),
]
PROFIT = {"도비": 0.15, "용접": 0.15, "베이스": 0.15}   # 이윤율 전 공종 동일 15% (사용자 지정 2026-10-12)
WORKS = [("도비", DOBI, "진공건조로 2기 회전·베이스 안착(1일), 용접면 밀착 이동(2일), 작업구역 안전펜스"),
         ("용접", WELD, "접합면 사상(2일) 및 접합부 CO2 용접(3일)"),
         ("베이스", BASE, "사다리형 일체 임시 베이스 제작·설치 및 공사 후 회수, H-250×250")]
TARGETS = {"도비": 20_000_000, "용접": 25_000_000, "베이스": 18_000_000}

# ---------------------------------------------------------------- B안 (대안): 기존 조방 7×3m 개조 + 전체 90° 회전 후 하부 용접
# 200t 크레인 단가는 업체 견적이 없어 사용자 제시 100t(240만)·60t(170만)에서 톤당 증가분(1.75만/t)으로 외삽한 추정값.
R_PIPE = 350_000      # 플랜트배관공 (노임 320,480 × 약 1.09)
B_DOBI = [
    ("노무비", "도비 전문기공", "안착·밀착·90° 회전·회전 후 정리", "인·일", 21, 300_000, "A안 (6, 6, 3) + 회전일 6명 × 1일", "L",
     "(6 + 6 + 3 + 6)명 × 300,000원", "인원 " + U + " (2026-10-12: A안 + 6명 1일), 단가 " + SH + "", ("비계공", WAGE["비계공"])),
    ("노무비", "신호수", "크레인 인양 신호", "인·일", 4, R_SIGNAL, "2명 × 2일 (안착일 + 회전일)", "L", "2명 × 2일 × 260,000원", AS, ("특별인부", WAGE["특별인부"])),
    ("노무비", "안전담당 (관리감독자)", "인양구역 통제", "인·일", 3, R_SAFE, "1명 × 3일", "L", "1명 × 3일 × 320,000원", AS, None),
    ("장비", "크레인", "100TON (안착)", "대·일", 1, 2_400_000, "안착일", "E", "1대 × 1일 × 2,400,000원", U, None),
    ("장비", "크레인", "60TON (안착)", "대·일", 1, 1_700_000, "안착일", "E", "1대 × 1일 × 1,700,000원", U, None),
    ("장비", "크레인", "200TON급 (90° 회전 주 크레인)", "대·일", 1, 3_500_000, "회전일, 일체 약 83t", "E",
     "1대 × 1일 × 3,500,000원", U + " (사용자 제시 2026-10-12: 350만/일)", None),
    ("장비", "200TON 회송·웨이트 운반", "카운터웨이트 트레일러 왕복", "식", 1, 1_200_000, "대형 크레인 부대비", "E", "1식", AS, None),
    ("장비", "회전용 대형 줄걸이", "100t급 샤클·슬링·스프레더 손료", "식", 1, 800_000, "200t 단독 회전", "E", "1식", AS, None),
    ("장비", "아웃리거 받침 철판", "1,500×1,500×25t, 8매", "회", 2, 640_000, "안착일 + 회전일", "E", "2회 × 640,000원", AS, None),
    ("장비", "공구사용손료", "도비 전문장비·줄걸이구", "SET", 1, 3_000_000, "", "E", "1식", SH, None),
    ("재료비", "회전 착지 깔목·고무판", "목재 받침 300×300, 고무판", "식", 1, 900_000, "회전 후 착지면 보호", "M", "1식", AS, None),
    ("재료비", "소모품·잡자재", "와이어 보호대 등", "식", 1, 1_000_000, "", "M", "1식", SH, None),
    ("안전시설", "안전펜스 임대", "이동식 메쉬펜스 H1.8×W2.4m", "매·일", 22 * 11, 5_000, "22매 × 11일", "E", "A안과 같음", "수량 " + AS + ", 단가 " + U, None),
    ("안전시설", "펜스 설치·해체", "설치 인건비", "인·일", 2, 250_000, "", "L", "2명 × 1일 × 250,000원", U, ("보통인부", WAGE["보통인부"])),
    ("안전시설", "펜스 운반", "1톤 화물 왕복", "회", 2, 260_000, "", "E", "2회 × 260,000원", AS, None),
    ("안전시설", "출입통제 안전장비", "", "식", 1, 380_000, "", "M", "1식", AS, None),
    expense(),
]
B_WELD = list(WELD)   # 용접 범위·인원 같음 (회전 후 하부 용접 포함)
B_JIG = [   # 사용자 제시(2026-10-12): 조방 개조안 약 800만원으로 베이스(약 1,800만원) 대체, 대신 200t급 크레인 1대 추가(B_도비)
    ("재료비", "보강보", "H-200급, 약 25m (현장 보유 강재 사용)", "식", 1, 0, "현장 보유분 사용, 자재비 미계상", "M", "사용자 지시 2026-10-12: 현장에 있는 강재 사용", U, None),
    ("재료비", "용접 소모품", "용접봉·가스·디스크", "식", 1, 420_000, "", "M", "1식", AS, None),
    ("재료비", "배관 복구 자재", "가스켓·볼트·너트", "식", 1, 380_000, "제거한 배관 재조립용", "M", "1식", AS, None),
    ("노무비", "제관사", "조방 보강보 취부", "인·일", 2, R_BOIL, "2명 × 1일", "L", "2명 × 1일 × 350,000원", AS, ("플랜트제관공", WAGE["플랜트제관공"])),
    ("노무비", "용접사", "보강보 용접", "인·일", 2, R_WELD, "2명 × 1일", "L", "2명 × 1일 × 500,000원", AS, ("플랜트용접공", WAGE["플랜트용접공"])),
    ("노무비", "조공", "보조", "인·일", 2, R_HELP, "2명 × 1일", "L", "2명 × 1일 × 200,000원", AS, ("보통인부", WAGE["보통인부"])),
    ("노무비", "화재감시자", "화기작업 전담", "인·일", 1, R_FIRE, "1명 × 1일", "L", "1명 × 1일 × 250,000원", AS, ("특별인부", WAGE["특별인부"])),
    ("노무비", "제관사", "바닥 접촉면 기구물·배관 제거·복원", "인·일", 4, R_BOIL, "2명 × 2일 (제거 1 + 복원 1), 사용자 지정 2명", "L",
     "2명 × 2일 × 350,000원", "사용자 지시 2026-10-12 (눕혀 작업하기 위한 기구물 제거·복원), 단가 " + AS, ("플랜트제관공", WAGE["플랜트제관공"])),
    expense_pd(2 + 2 + 2 + 1 + 4, "개조 1일 7명 + 제거·복원 2명 × 2일 (용접기·절단기 임대는 용접 공종에 계상)"),
]
WORKS_B = [("B_도비", B_DOBI, "B안: 조방 안착(100t·60t) + 전체 90° 회전(200t 단독), 안전펜스"),
           ("B_용접", B_WELD, "B안: 사상, 3면 용접 → 회전 후 나머지 1면 용접"),
           ("B_조방", B_JIG, "B안: 기존 조방 7×3m 개조(현장 보유 보강보), 바닥 접촉면 기구물·배관 제거·복원")]
PROFIT.update({"B_도비": 0.15, "B_용접": 0.15, "B_조방": 0.15})
ALLWORKS = WORKS + WORKS_B


def calc(rows, pr):
    M = sum(r[4] * r[5] for r in rows if r[7] == "M")
    L = sum(r[4] * r[5] for r in rows if r[7] == "L")
    E = sum(r[4] * r[5] for r in rows if r[7] == "E")
    rt = {k: v for k, _, v, _, _ in RATES}
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
    ws.page_margins.left = ws.page_margins.right = 0.5
    ws.page_margins.top = ws.page_margins.bottom = 0.6
    ws.sheet_view.showGridLines = False
    ws.oddFooter.center.text = "엔엠솔루션  |  " + QNO + "  |  &P / &N"
    ws.oddFooter.center.size = 8


# ---------------------------------------------------------------- 요율근거 (먼저 만들어 셀 주소 확보)
rs = wb.create_sheet("요율근거")
rs["A1"] = "요율 및 단가 근거"; rs["A1"].font = font(True, 14)
rs["A2"] = "파란 글씨는 입력값임. 변경 시 모든 공종 시트와 견적서가 다시 계산됨."; rs["A2"].font = font(sz=9, color="555555")
for c, h in enumerate(["항목", "적용 기준", "요율", "출처·비고"], 1):
    x = rs.cell(4, c, h); x.font = font(True, color="FFFFFF"); x.fill = NAVY; x.alignment = CEN
RC = {}
r = 5
for key, nm, v, base, src in RATES:
    rs.cell(r, 1, nm); rs.cell(r, 2, base); x = rs.cell(r, 3, v); x.font = font(color="0000FF"); x.number_format = "0.00%"
    rs.cell(r, 4, src); RC[key] = f"요율근거!$C${r}"; r += 1
for nm, _, _ in ALLWORKS:
    rs.cell(r, 1, f"이윤 ({nm})"); rs.cell(r, 2, "(노무비 + 경비 + 일반관리비) ×")
    x = rs.cell(r, 3, PROFIT[nm]); x.font = font(color="0000FF"); x.number_format = "0.0%"
    rs.cell(r, 4, S_RULE + ". 15% 이내"); RC["P_" + nm] = f"요율근거!$C${r}"; r += 1
rs.cell(r, 1, "국민연금·건강·장기요양"); rs.cell(r, 2, "미계상"); rs.cell(r, 3, "-")
rs.cell(r, 4, "1인 투입 3~5일로 건설일용 사업장가입 기준(월 8일 이상 또는 월 220만원 이상)에 미달하는 것으로 봄. 인원 운용이 바뀌면 국민연금 4.75%, 건강 3.545%, 장기요양(건강보험료의 12.95%) 추가")
r += 1
border_range(rs, 4, r - 1, 1, 4)
for rr in range(5, r):
    for c in range(1, 5):
        x = rs.cell(rr, c); x.alignment = LEFT if c in (1, 2, 4) else CEN
        if x.font.name != FN: x.font = font()
    rs.row_dimensions[rr].height = 30
r += 1
rs.cell(r, 1, "시중노임 대비 적용 단가").font = font(True, 12); r += 1
for c, h in enumerate(["직종 (견적)", "시중노임 직종", "시중노임(원/일)", "적용 단가(원/일)", "배율"], 1):
    x = rs.cell(r, c, h); x.font = font(True, color="FFFFFF"); x.fill = NAVY; x.alignment = CEN
h0 = r; r += 1
for job, ref, rate in [("용접사", "플랜트용접공", R_WELD), ("고급기술자·제관사", "플랜트제관공", R_BOIL),
                       ("사상 기술자", "특별인부", R_GRIND), ("조공", "보통인부", R_HELP),
                       ("화재감시자", "특별인부", R_FIRE), ("신호수", "특별인부", R_SIGNAL),
                       ("도비 전문기공", "비계공", 300_000)]:
    rs.cell(r, 1, job); rs.cell(r, 2, ref)
    x = rs.cell(r, 3, WAGE[ref]); x.number_format = WON; x.font = font(color="0000FF")
    x = rs.cell(r, 4, rate); x.number_format = WON; x.font = font(color="0000FF")
    x = rs.cell(r, 5, f"=D{r}/C{r}"); x.number_format = "0.00"
    r += 1
border_range(rs, h0, r - 1, 1, 5)
for rr in range(h0 + 1, r):
    for c in range(1, 6):
        x = rs.cell(rr, c)
        if x.font.name != FN: x.font = font()
        x.alignment = RIGHT if c >= 3 else LEFT
rs.cell(r, 1, "시중노임은 " + S_WAGE1 + " 기준. " + S_WAGE2 + ". 적용 단가는 숙련 기능공 현장 실거래 수준(사용자 지시)으로 시중노임보다 높게 적용함.").font = font(sz=9, color="555555")
rs.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5); rs.cell(r, 1).alignment = LEFT; rs.row_dimensions[r].height = 30
r += 2
rs.cell(r, 1, "출처").font = font(True, 12); r += 1
for nm, url in URLS:
    rs.cell(r, 1, nm).font = font(); x = rs.cell(r, 2, url); x.font = font(sz=9, color="0563C1")
    if url.startswith("http"): x.hyperlink = url
    rs.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5); r += 1
rs.cell(r + 1, 1, "조사일 2026-10-12. 요율은 해마다 바뀌므로 계약 시점 고시로 재확인 필요함.").font = font(sz=9, color="555555")
setw(rs, {"A": 22, "B": 26, "C": 14, "D": 60, "E": 10})
page(rs, landscape=True)


# ---------------------------------------------------------------- 공종 시트
def header(ws, work, ncol=9):
    last = "ABCDEFGHI"[ncol - 1]
    ws.merge_cells(f"A1:{last}2"); ws["A1"] = "공 종 별   내 역 서"; ws["A1"].font = font(True, 18); ws["A1"].alignment = CEN
    ws.merge_cells(f"A3:{last}3"); ws["A3"] = f"견적번호 {QNO}   |   견적일 {QDATE}   |   금강변압기 유한회사 귀중"
    ws["A3"].font = font(sz=9, color="555555"); ws["A3"].alignment = CEN
    ws.merge_cells(f"A5:{last}5"); ws["A5"] = "공 종 : " + work; ws["A5"].font = font(True, 11); ws["A5"].alignment = LEFT


def table_head(ws, r):
    for c, h in zip(range(1, 10), ("구분", "품명", "규격", "단위", "수량", "단가", "금액", "비고", "비목")):
        x = ws.cell(r, c, h); x.font = font(True, color="FFFFFF"); x.alignment = CEN; x.fill = NAVY; x.border = BOX


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
        ws.cell(tot, c).fill = YEL
    ws.cell(tot, 1).font = font(True, 11); ws.cell(tot, 7).font = font(True, 11)
    fr = tot + 2
    for t in ["※ 파란 글씨는 수량·단가 입력값이며, 근거는 '산출근거'·'요율근거' 시트에 정리함.",
              "※ 연장·야간·휴일 작업은 공정에 따라 당사가 수행함 (추가 정산 없음)."]:
        ws.merge_cells(start_row=fr, start_column=1, end_row=fr, end_column=9)
        ws.cell(fr, 1, t).font = font(sz=9, color="555555"); fr += 1
    setw(ws, {"A": 9, "B": 20, "C": 30, "D": 7, "E": 8, "F": 11, "G": 13, "H": 38, "I": 7})
    page(ws)
    ws.print_title_rows = "7:7"
    TOTREF[name] = (f"{name}!$G$", pos)
    CALC[name] = calc(rows, PROFIT[name])


for nm, rows, work in ALLWORKS:
    detail(nm, rows, work)

# ---------------------------------------------------------------- 원가계산서
cs = wb.create_sheet("원가계산서", 1)
cs.merge_cells("A1:F2"); cs["A1"] = "공 사 원 가 계 산 서"; cs["A1"].font = font(True, 18); cs["A1"].alignment = CEN
cs.merge_cells("A3:F3"); cs["A3"] = "공사명 : 진공건조로 회전·안착 및 접합부 용접 공사   |   " + QNO; cs["A3"].font = font(sz=9, color="555555"); cs["A3"].alignment = CEN
for c, h in enumerate(["비목", "도비", "용접", "베이스", "합계", "구성비"], 1):
    x = cs.cell(5, c, h); x.font = font(True, color="FFFFFF"); x.fill = NAVY; x.alignment = CEN; x.border = BOX
lines = [("재료비", "M"), ("노무비", "L"), ("경비", "E"), ("산재보험료", "SAN"), ("고용보험료", "GOY"),
         ("산업안전보건관리비", "SAFE"), ("순공사원가", "NET"), ("일반관리비", "GEN"), ("이윤", "PROF"),
         ("단수정리", "RND"), ("공급가액", "TOT")]
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
SUPROW = r - 1
cs.cell(r, 1, "부가가치세 (10%)"); cs.cell(r, 5, f"=E{SUPROW}*0.1"); VATROW = r; r += 1
cs.cell(r, 1, "총 공사금액"); cs.cell(r, 5, f"=E{SUPROW}+E{VATROW}"); GRAND = r
for rr in (VATROW, GRAND):
    for c in range(1, 7):
        x = cs.cell(rr, c); x.border = BOX; x.font = font(rr == GRAND); x.alignment = CEN if c == 1 else RIGHT
        if c >= 2: x.number_format = WON
        if rr == GRAND: x.fill = YEL
cs.cell(GRAND + 2, 1, "요율: 산재 3.56%·고용 1.15% (노무비 기준), 산업안전보건관리비 3.11% ((재료비 + 노무비) 기준), 일반관리비 5% (순공사원가 기준), 이윤은 공종별 15% 이내. 근거는 '요율근거' 시트.").font = font(sz=9, color="555555")
cs.merge_cells(start_row=GRAND + 2, start_column=1, end_row=GRAND + 2, end_column=6)
cs.cell(GRAND + 2, 1).alignment = LEFT; cs.row_dimensions[GRAND + 2].height = 30
setw(cs, {"A": 20, "B": 15, "C": 15, "D": 15, "E": 16, "F": 9})
page(cs)

# ---------------------------------------------------------------- 견적서 표지
sup_val = sum(CALC[n]["tot"] for n, _, _ in WORKS)
grand_val = sup_val + sup_val // 10
WORKNAME = "진공건조로 2분할체 회전·안착, 임시 베이스 제작, 접합부 사상·용접 공사"
c = cover
c.merge_cells("A1:I2"); c["A1"] = "견          적          서"; c["A1"].font = font(True, 24); c["A1"].alignment = CEN
c.merge_cells("A3:I3"); c["A3"] = "QUOTATION"; c["A3"].font = font(sz=10, color="777777"); c["A3"].alignment = CEN
# 왼쪽: 수신
info_l = [("견적번호", QNO), ("견 적 일", QDATE), ("수    신", "금강변압기 유한회사 (KUMKANG TRANSFORMER) 귀중"),
          ("참    조", ""), ("공 사 명", WORKNAME),
          ("공사 장소", "경기 화성시 마도면 마도공단로1길 214 사내 마당"),
          ("공사 기간", "2026.10.21 ~ 2026.11.02 (가계획, 작업일 11일, 공정표 시트 참조)"),
          ("결제 조건", "협의 (예: 착수 30%, 준공 후 70%)"),
          ("유효 기간", "견적일로부터 30일")]
r = 5
for k, v in info_l:
    c.cell(r, 1, k).font = font(True); c.cell(r, 1).fill = GRAY; c.cell(r, 1).alignment = CEN
    c.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5); c.cell(r, 2, v).font = font(); c.cell(r, 2).alignment = LEFT
    if k in ("참    조", "결제 조건"): c.cell(r, 2).fill = YEL
    border_range(c, r, r, 1, 5); c.row_dimensions[r].height = 24
    r += 1
# 오른쪽: 공급자
c.merge_cells("G5:G9"); c["G5"] = "공\n급\n자"; c["G5"].font = font(True); c["G5"].alignment = CEN; c["G5"].fill = GRAY
sup = [("상    호", "엔엠솔루션"), ("대 표 자", "                    (인)"), ("사업자번호", ""), ("주    소", ""), ("연 락 처", "")]
for i, (k, v) in enumerate(sup):
    rr = 5 + i
    c.cell(rr, 8, k).font = font(True); c.cell(rr, 8).alignment = CEN
    c.cell(rr, 9, v).font = font(i == 0); c.cell(rr, 9).alignment = LEFT
    if i: c.cell(rr, 9).fill = YEL
border_range(c, 5, 9, 7, 9)
c.merge_cells("G10:I13"); c["G10"] = "(직인)"; c["G10"].font = font(sz=9, color="AAAAAA"); c["G10"].alignment = CEN
border_range(c, 10, 13, 7, 9)
r = 15
c.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
c.cell(r, 1, "아래와 같이 견적함.").font = font(sz=10); r += 1
c.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2); c.cell(r, 1, "합계 금액\n(부가세 포함)").font = font(True)
c.cell(r, 1).alignment = CEN; c.cell(r, 1).fill = GRAY
c.merge_cells(start_row=r, start_column=3, end_row=r, end_column=9)
c.cell(r, 3, f"=원가계산서!E{GRAND}").font = font(True, 14); c.cell(r, 3).number_format = '"₩"#,##0'; c.cell(r, 3).alignment = CEN
border_range(c, r, r, 1, 9); c.row_dimensions[r].height = 40
WORDROW = r
r += 2
for col, h in zip(range(1, 10), ("No", "공종", "", "작업 내용", "단위", "수량", "공급가액", "", "비고")):
    pass
hdr = [(1, 1, "No"), (2, 3, "공  종"), (4, 6, "작업 내용"), (7, 7, "단위"), (8, 8, "공급가액"), (9, 9, "비고")]
for a, b, h in hdr:
    if b > a: c.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)
    x = c.cell(r, a, h); x.font = font(True, color="FFFFFF"); x.alignment = CEN
for cc in range(1, 10): c.cell(r, cc).fill = NAVY; c.cell(r, cc).border = BOX
r += 1
items = [("1", "도비 작업", "회전·베이스 안착(100t·60t), 밀착 이동, 안전펜스", "도비"),
         ("2", "용접 작업", "접합면 사상, CO2 용접, 화재감시", "용접"),
         ("3", "베이스 작업", "임시 사다리형 베이스 제작·설치·회수", "베이스")]
i0 = r
for no, nm, desc, key in items:
    c.cell(r, 1, no); c.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3); c.cell(r, 2, nm)
    c.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6); c.cell(r, 4, desc)
    c.cell(r, 7, "식"); pre, pos = TOTREF[key]; x = c.cell(r, 8, f"={pre}{pos['TOT']}"); x.font = font(color="008000")
    c.cell(r, 9, f"{key} 시트 참조")
    r += 1
rows_tot = [("공 급 가 액", f"=SUM(H{i0}:H{r-1})"), ("부 가 세 (10%)", f"=H{r}*0.1"), ("합      계", f"=H{r}+H{r+1}")]
for j, (k, f) in enumerate(rows_tot):
    c.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7); c.cell(r, 1, k).font = font(True)
    c.cell(r, 8, f).font = font(True); c.cell(r, 8).number_format = WON
    for cc in range(1, 10): c.cell(r, cc).fill = GRAY if j < 2 else YEL
    r += 1
for rr in range(i0, r):
    for cc in range(1, 10):
        x = c.cell(rr, cc); x.border = BOX
        if x.font.name != FN: x.font = font()
        if x.alignment.horizontal is None: x.alignment = LEFT if cc in (4, 9) else CEN
        if cc == 8: x.number_format = WON; x.alignment = RIGHT
    c.row_dimensions[rr].height = 24
r += 1
c.cell(r, 1, "특기사항").font = font(True, 11); r += 1
NOTES = [
    "1. 전기 수전, 배관·닥트 해체·복구, 보온, 도장, 시운전은 포함하지 않음.",
    "2. 크레인 사양·배치는 인양계획서와 지반 검토 후 확정하며, 크레인 변경에 따른 추가비용은 없음.",
    "3. 비파괴검사(PT/UT)는 포함하지 않으며, 발주처 요구 시 별도 정산함.",
    "4. 연장·야간·휴일 작업은 공정에 따라 당사가 수행하며 추가 정산하지 않음. 기상(우천 시 용접 불가)·발주처 사유에 따른 대기는 별도 협의함.",
    "5. 작업구역(약 10×15m)은 이동식 안전펜스로 통제하며, 발주처 안전관리 규정과 출입 절차를 따름.",
    "6. 대안(B안: 기존 조방 개조 + 전체 90° 회전)의 금액은 AB비교 시트에 정리함.",
]
for t in NOTES:
    c.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    c.cell(r, 1, t).font = font(sz=9); c.cell(r, 1).alignment = LEFT; c.row_dimensions[r].height = 18; r += 1
c.cell(r + 1, 1, "노란 칸은 공급자 정보·협의 항목으로 채워야 하는 칸임.").font = font(sz=8, color="999999")
setw(c, {"A": 11, "B": 9, "C": 9, "D": 12, "E": 12, "F": 8, "G": 8, "H": 15, "I": 26})
page(c)

# ---------------------------------------------------------------- 산출근거
gs = wb.create_sheet("산출근거", 2)
gs["A1"] = "견적 산출근거"; gs["A1"].font = font(True, 14)
gs["A2"] = "금액은 공종 시트에서 연동됨(초록). 노란 칸은 업체 견적으로 확인할 가정값임."; gs["A2"].font = font(sz=9, color="555555")
for cc, h in enumerate(["공종", "품명", "규격", "산출식", "금액(원)", "근거", "시중노임 참고"], 1):
    x = gs.cell(4, cc, h); x.font = font(True, color="FFFFFF"); x.fill = NAVY; x.border = BOX; x.alignment = CEN
r = 5
for nm, _, _ in ALLWORKS:
    for dr, row in ROWMAP[nm]:
        item, spec, f_, src, wref = row[1], row[2], row[8], row[9], row[10]
        gs.cell(r, 1, nm); gs.cell(r, 2, item); gs.cell(r, 3, spec); gs.cell(r, 4, f_)
        x = gs.cell(r, 5, f"={nm}!G{dr}"); x.font = font(color="008000"); x.number_format = WON
        gs.cell(r, 6, src)
        if wref: gs.cell(r, 7, f"{wref[0]} {wref[1]:,}원")
        if "가정" in src: gs.cell(r, 6).fill = YEL
        r += 1
    pre, pos = TOTREF[nm]
    for label, k in (("법정경비·일반관리비·이윤", None), ("공종 합계", "TOT")):
        gs.cell(r, 2, label).font = font(True)
        if k: x = gs.cell(r, 5, f"={pre}{pos[k]}")
        else: x = gs.cell(r, 5, f"={pre}{pos['TOT']}-SUM({pre}{pos['M']}:{pre}{pos['E']})")
        x.font = font(True); x.number_format = WON
        gs.cell(r, 6, "요율근거 시트" if not k else (f"목표 약 {TARGETS[nm]:,}원 (사용자 제시)" if nm in TARGETS else "B안 (대안)"))
        for cc in range(1, 8): gs.cell(r, cc).fill = GRAY
        r += 1
    r += 1
for rr in range(5, r):
    if gs.cell(rr, 1).value is None and gs.cell(rr, 2).value is None: continue
    for cc in range(1, 8):
        x = gs.cell(rr, cc); x.border = BOX
        if x.font.name != FN: x.font = font(x.font.b)
        x.alignment = RIGHT if cc == 5 else LEFT
        x.alignment = Alignment(horizontal=x.alignment.horizontal, vertical="center", wrap_text=True)
gs.cell(r, 1, "근거 구분: 사용자 제시 = 받은 단가·인원·일수 / 서해테크 = 도비 협력사 견적서(2026-10-10) / 공급처 = H빔 견적 / 가정값 = 업체 견적 전 임시값").font = font(sz=9, color="555555")
setw(gs, {"A": 8, "B": 18, "C": 28, "D": 40, "E": 13, "F": 44, "G": 20})
page(gs, landscape=True)
gs.print_title_rows = "4:4"

# ---------------------------------------------------------------- 베이스 자재
bs = wb.create_sheet("베이스자재")
bs["A1"] = "베이스 자재: H-250×250×9×14 (공급처 견적, 부가세 미포함)"; bs["A1"].font = font(True, 13)
for cc, h in enumerate(["길이(m)", "수량(EA)", "총 길이(m)", "단위중량(kg/m)", "중량(kg)"], 1):
    x = bs.cell(3, cc, h); x.font = font(True, color="FFFFFF"); x.fill = NAVY; x.border = BOX; x.alignment = CEN
for i, (ln, q) in enumerate([(10, 4), (5, 16), (0.6, 30)]):
    rr = 4 + i
    bs.cell(rr, 1, ln).font = font(color="0000FF"); bs.cell(rr, 2, q).font = font(color="0000FF")
    bs.cell(rr, 3, f"=A{rr}*B{rr}"); bs.cell(rr, 4, 72.4).font = font(color="0000FF"); bs.cell(rr, 5, f"=C{rr}*D{rr}")
    bs.cell(rr, 5).number_format = '#,##0'
bs.cell(7, 2, "합계").font = font(True); bs.cell(7, 3, "=SUM(C4:C6)"); bs.cell(7, 5, "=SUM(E4:E6)"); bs.cell(7, 5).number_format = '#,##0'
border_range(bs, 4, 7, 1, 5)
for cc in range(1, 6): bs.cell(7, cc).fill = GRAY
for cc, h in enumerate(["구분", "금액(원)", "kg당(원)", "비고"], 1):
    x = bs.cell(9, cc, h); x.font = font(True, color="FFFFFF"); x.fill = NAVY; x.alignment = CEN
vals = [("신품 구매", 13_575_440, "공급처 견적"), ("중고 구매", 9_577_783, "공급처 견적"),
        ("고철 매각 예상", 4_000_000, "사용자 제시 (약 400만원)"),
        ("회수조건 적용", "=E7*700", "kg당 700원, 견적 반영"),
        ("회수 손료 (차액)", "=B13-(B11-B12)", "중고 구매가 − 고철가치 대비 추가분")]
for i, (k, v, note) in enumerate(vals):
    rr = 10 + i
    bs.cell(rr, 1, k); x = bs.cell(rr, 2, v); x.number_format = WON
    if isinstance(v, int): x.font = font(color="0000FF")
    bs.cell(rr, 3, f"=B{rr}/$E$7"); bs.cell(rr, 3).number_format = '#,##0'; bs.cell(rr, 4, note)
border_range(bs, 9, 14, 1, 4)
bs.cell(16, 1, "단위중량 72.4 kg/m는 KS H-250×250×9×14 공칭값임. 베이스는 임시 구조물로 방청·도장 미적용.").font = font(sz=9, color="555555")
setw(bs, {"A": 18, "B": 16, "C": 14, "D": 34, "E": 14})
page(bs)

# ---------------------------------------------------------------- 공정표
ps = wb.create_sheet("공정표")
import datetime as _dt
WDAYS = []
_d = _dt.date(2026, 10, 21)
while len(WDAYS) < 11:
    if _d.weekday() != 6: WDAYS.append(f"{_d.month}/{_d.day}\n" + "월화수목금토일"[_d.weekday()])
    _d += _dt.timedelta(days=1)
ps["A1"] = "공정표 (가계획 2026.10.21 ~ 11.02, 작업일 11일)"; ps["A1"].font = font(True, 14)
steps = [("안전펜스 설치·작업구역 통제", "조공 2명", 1, 1), ("베이스 제작·설치", "제관 2, 용접 1, 조공 1, 화재감시 1", 1, 3),
         ("회전·베이스 안착 (100t + 60t)", "도비 6명, 신호수 2, 안전담당", 4, 4),
         ("접합면 사상", "사상 2, 조공 2, 화재감시 1", 5, 6), ("용접면 밀착 이동", "도비 6명·3명, 안전담당", 7, 8),
         ("접합부 용접", "용접사 3, 고급기술자 2, 화재감시 1", 9, 11), ("펜스 해체·정리", "조공 2명", 11, 11)]
ps.cell(3, 1, "공정").font = font(True); ps.cell(3, 2, "투입 인원").font = font(True)
for d in range(1, 12):
    x = ps.cell(3, 2 + d, WDAYS[d - 1]); x.font = font(True, sz=9); x.alignment = CEN
    ps.column_dimensions[x.column_letter].width = 5.5
BAR = PatternFill("solid", fgColor="E8742A")
for i, (nm, crew, a, b) in enumerate(steps):
    rr = 4 + i
    ps.cell(rr, 1, nm).font = font(); ps.cell(rr, 2, crew).font = font()
    for d in range(1, 12):
        if a <= d <= b: ps.cell(rr, 2 + d).fill = BAR
border_range(ps, 3, 3 + len(steps), 1, 13)
for cc in range(1, 14): ps.cell(3, cc).fill = GRAY; ps.cell(3, cc).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ps.row_dimensions[3].height = 30
setw(ps, {"A": 30, "B": 32})
ps.cell(5 + len(steps), 1, "가계획: 2026-10-21(수) 착공 ~ 11-02(월) 완료, 일요일 휴무 기준 작업일 11일. 착공일은 발주처 협의로 확정함.").font = font(sz=9, color="555555")
page(ps, landscape=True)


# ---------------------------------------------------------------- A·B 비교
xs = wb.create_sheet("AB비교", 2)
xs.merge_cells("A1:E1"); xs["A1"] = "A안 · B안 금액 비교 (공급가액, 부가세 별도)"; xs["A1"].font = font(True, 14)
for cc, h in enumerate(["구분", "A안 (임시 베이스 + 밀착 이동)", "B안 (조방 개조 + 90° 회전)", "차액 (B − A)", "비고"], 1):
    x = xs.cell(3, cc, h); x.font = font(True, color="FFFFFF"); x.fill = NAVY; x.alignment = CEN
pairs = [("도비", "도비", "B_도비", "B안: 200t급 크레인 1대 추가(350만/일, 단독 90° 회전), 도비 6명 1일 추가"),
         ("용접", "용접", "B_용접", "같은 인원·장비로 상부·측면 용접 후 회전해서 하부 용접"),
         ("베이스 / 조방", "베이스", "B_조방", "B안: 신규 베이스 대신 조방 개조(현장 보강보), 기구물·배관 제거·복원 2명 × 2일")]
r = 4
for lab, a, b, note in pairs:
    xs.cell(r, 1, lab)
    xs.cell(r, 2, f"={TOTREF[a][0]}{TOTREF[a][1]['TOT']}").font = font(color="008000")
    xs.cell(r, 3, f"={TOTREF[b][0]}{TOTREF[b][1]['TOT']}").font = font(color="008000")
    xs.cell(r, 4, f"=C{r}-B{r}"); xs.cell(r, 5, note); r += 1
xs.cell(r, 1, "공급가액 계").font = font(True)
for cc, col in ((2, "B"), (3, "C"), (4, "D")):
    xs.cell(r, cc, f"=SUM({col}4:{col}{r-1})").font = font(True)
border_range(xs, 3, r, 1, 5)
for rr in range(4, r + 1):
    for cc in range(1, 6):
        x = xs.cell(rr, cc)
        if x.font.name != FN: x.font = font()
        x.alignment = RIGHT if cc in (2, 3, 4) else LEFT
        if cc in (2, 3, 4): x.number_format = WON
    xs.row_dimensions[rr].height = 30
for cc in range(1, 6): xs.cell(r, cc).fill = YEL
r += 2
for t in ["※ B안 200t급 크레인은 사용자 제시 단가 350만/일 적용. 회송·웨이트 운반은 별도 줄로 넣음. B안 용접은 3면 용접 → 전체 90° 회전 → 나머지 1면 용접 순서로, 용접 1일이 회전 후로 이동함(금액 동일).",
          "※ B안 회전은 일체 약 83t(45t + 38t)을 모서리 기준으로 세우는 작업으로, 실제 크레인 사양은 인양계획서·구조검토로 확정 필요함.",
          "※ 회전 후 원래 자세로 되돌리는 회전이나 영구 받침이 필요하면 별도임."]:
    xs.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    xs.cell(r, 1, t).font = font(sz=9, color="555555"); xs.cell(r, 1).alignment = LEFT; xs.row_dimensions[r].height = 28; r += 1
setw(xs, {"A": 16, "B": 22, "C": 22, "D": 16, "E": 52})
page(xs, landscape=True)

wb.move_sheet("요율근거", offset=len(wb.sheetnames))
wb.save(OUT)
for n, _, _ in ALLWORKS:
    k = CALC[n]
    print(n, {a: round(b) for a, b in k.items()})
print("supply", sup_val, "grand", grand_val, korean_won(grand_val))
print(OUT)
