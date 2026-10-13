"""표지 + 견적서 + 작업계획서를 한 PDF로 묶고 북마크를 붙임.
사용: python3 mkbundle.py cover.pdf 견적서.pdf 작업계획서.pdf out.pdf"""
import sys
from pypdf import PdfReader, PdfWriter

cover, quote, plan, out = sys.argv[1:5]
w = PdfWriter()
rc, rq, rp = PdfReader(cover), PdfReader(quote), PdfReader(plan)
for r in (rc, rq, rp):
    for p in r.pages:
        w.add_page(p)
nc, nq = len(rc.pages), len(rq.pages)

w.add_outline_item("표지", 0)
w.add_outline_item("문서 구성", 1)
q_root = w.add_outline_item("Ⅰ. 견적서", nc)
for o in rq.outline:
    if not isinstance(o, list):
        w.add_outline_item(o.title, nc + rq.get_destination_page_number(o), parent=q_root)
p_root = w.add_outline_item("Ⅱ. 작업계획서", nc + nq)
PLAN_SECTIONS = [("1. 공사 개요", 1), ("2. A안 작업 순서", 2), ("3. 인양계획", 6),
                 ("4. A안 타임테이블", 7), ("5. B안 대안", 7), ("6. A안·B안 비교", 13),
                 ("7. 안전관리 계획", 14), ("8. 확인이 필요한 사항", 15)]
for title, pg in PLAN_SECTIONS:
    w.add_outline_item(title, nc + nq + pg - 1, parent=p_root)
w.add_metadata({"/Title": "금강변압기 진공건조로 회전·안착 및 접합 용접 공사 - 견적서·작업계획서", "/Author": "엔엠솔루션"})
with open(out, "wb") as f:
    w.write(f)
print("pages", len(w.pages))
