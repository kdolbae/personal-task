import sys, subprocess, re, os, shutil
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from pypdf import PdfReader, PdfWriter
SRC=sys.argv[1]; OUT=sys.argv[2]; W='/tmp/claude-0/-home-claude-personal-task/cc458d43-e61f-5be7-bef2-0fa5339e22e4/scratchpad/pdfwork'
shutil.rmtree(W,ignore_errors=True); os.makedirs(W)
def prep(wb, index=None):
    names=[ws.title for ws in wb.worksheets]
    for i,ws in enumerate(wb.worksheets):
        ws.page_setup.paperSize=ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
        ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.print_options.horizontalCentered=True
        ws.page_margins.left=ws.page_margins.right=0.5; ws.page_margins.top=0.7; ws.page_margins.bottom=0.6
        ws.oddHeader.left.text=f"&\"Arial\"&8 금강변압기 진공건조로 회전·안착·용접 견적서"
        ws.oddHeader.right.text=f"&\"Arial\"&8 §{i+1}. &A"
        ws.oddFooter.center.text="&\"Arial\"&8 - &P / &N -"
    if index is not None:
        ix=wb.create_sheet("목차",0)
        ix.page_setup.paperSize=ix.PAPERSIZE_A4; ix.page_setup.fitToWidth=1; ix.page_setup.fitToHeight=0
        ix.sheet_properties.pageSetUpPr.fitToPage=True
        ix.oddFooter.center.text="&\"Arial\"&8 - &P / &N -"
        ix.column_dimensions['A'].width=8; ix.column_dimensions['B'].width=34; ix.column_dimensions['C'].width=46; ix.column_dimensions['D'].width=12
        ix['A1']="견적서 목차"; ix['A1'].font=Font(name='Arial',bold=True,size=16)
        ix['A2']="금강변압기 진공건조로 회전·안착 및 접합 용접 공사 (공급자: 엔엠솔루션)"; ix['A2'].font=Font(name='Arial',size=10,color='555555')
        hdr=["No","시트","내용","페이지"]
        desc={"견적서":"표지·공종 합계·특기사항","원가계산서":"비목별 원가 집계 (A안)","AB비교":"A안·B안 금액 비교","산출근거":"항목별 산출식·근거",
              "도비":"A안 도비 내역","용접":"A안 용접 내역","베이스":"A안 베이스 내역","B_도비":"B안 도비 내역","B_용접":"B안 용접 내역","B_조방":"B안 조방 개조 내역",
              "베이스자재":"H빔 자재 산출","공정표":"공정 가계획 (10/21~11/2)","요율근거":"보험·관리비·이윤 요율 근거"}
        thin=Side(style='thin',color='999999')
        for c,h in enumerate(hdr,1):
            x=ix.cell(4,c,h); x.font=Font(name='Arial',bold=True,color='FFFFFF'); x.fill=PatternFill('solid',fgColor='1F3864'); x.alignment=Alignment(horizontal='center')
            x.border=Border(top=thin,bottom=thin,left=thin,right=thin)
        r=5
        for i,n in enumerate(names):
            vals=[f"§{i+1}",n,desc.get(n,""),f"p.{index[n][0]}" if index[n][0]==index[n][1] else f"p.{index[n][0]}~{index[n][1]}"]
            for c,v in enumerate(vals,1):
                x=ix.cell(r,c,v); x.font=Font(name='Arial',size=10); x.border=Border(top=thin,bottom=thin,left=thin,right=thin)
                x.alignment=Alignment(horizontal='center' if c in (1,4) else 'left',vertical='center')
            ix.row_dimensions[r].height=20; r+=1
    return names
def convert(xlsx):
    subprocess.run(["soffice","--headless","--convert-to","pdf","--outdir",W,xlsx],check=True,capture_output=True,timeout=180)
    return os.path.join(W,os.path.splitext(os.path.basename(xlsx))[0]+".pdf")
# pass 1
wb=load_workbook(SRC); names=prep(wb); p1=os.path.join(W,"pass1.xlsx"); wb.save(p1); pdf1=convert(p1)
rd=PdfReader(pdf1); pages={}
for pi,pg in enumerate(rd.pages,1):
    t=pg.extract_text() or ""
    m=re.search(r"§(\d+)\.",t)
    if m:
        n=names[int(m.group(1))-1]; pages.setdefault(n,[pi,pi]); pages[n][1]=pi
    else:
        print("no header on page",pi,repr(t[:60]))
print(pages)
# pass 2 with index (index page adds 1 page; assume index is 1 page)
index={n:[a+1,b+1] for n,(a,b) in pages.items()}
wb=load_workbook(SRC); prep(wb,index); p2=os.path.join(W,"pass2.xlsx"); wb.save(p2); pdf2=convert(p2)
rd=PdfReader(pdf2); w=PdfWriter()
for pg in rd.pages: w.add_page(pg)
w.add_outline_item("목차",0)
for i,n in enumerate(names):
    w.add_outline_item(f"§{i+1}. {n}",index[n][0]-1)
with open(OUT,"wb") as f: w.write(f)
print("pages",len(rd.pages),"->",OUT)
