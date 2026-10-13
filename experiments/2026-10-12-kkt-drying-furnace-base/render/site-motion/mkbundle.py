"""표지 + 견적서 + 작업계획서를 한 PDF로 묶고 북마크를 붙임.
사용: python3 mkbundle.py cover.pdf 작업계획서.pdf 견적서.pdf out.pdf  (제본 순서: 표지 → 계획서 → 견적서)"""
import sys
from pypdf import PdfReader, PdfWriter

cover, plan, quote, out = sys.argv[1:5]
w = PdfWriter()
rc, rq, rp = PdfReader(cover), PdfReader(quote), PdfReader(plan)
for r in (rc, rp, rq):
    for p in r.pages:
        w.add_page(p)
nc, np_ = len(rc.pages), len(rp.pages)

w.add_outline_item("표지", 0)
w.add_outline_item("문서 구성", 1)
p_root = w.add_outline_item("Ⅰ. 작업계획서", nc)
PLAN_SECTIONS = [("1. 공사 개요", 1), ("2. A안 작업 순서", 2), ("3. 인양계획", 6),
                 ("4. A안 타임테이블", 7), ("5. B안 대안", 7), ("6. A안·B안 비교", 13),
                 ("7. 안전관리 계획", 14), ("8. 확인이 필요한 사항", 15)]
for title, pg in PLAN_SECTIONS:
    w.add_outline_item(title, nc + pg - 1, parent=p_root)
q_root = w.add_outline_item("Ⅱ. 견적서", nc + np_)
for o in rq.outline:
    if not isinstance(o, list):
        w.add_outline_item(o.title, nc + np_ + rq.get_destination_page_number(o), parent=q_root)
w.add_metadata({"/Title": "금강변압기 진공건조로 회전·안착 및 접합 용접 공사 - 견적서·작업계획서", "/Author": "엔엠솔루션"})
with open(out, "wb") as f:
    w.write(f)
print("pages", len(w.pages))
