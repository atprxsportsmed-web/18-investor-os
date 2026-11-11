# Last text/structure pass: removes the exit and returned-capital scenarios, rebuilds Investor returns
# (cash only), merges the ramp-loss and cushion rows, drops staged funding, and strips explanatory text.
import openpyxl
from copy import copy
wb=openpyxl.load_workbook("stage3b.xlsx")

# ---------- 1. Investor returns, rebuilt: cash only, no exit, no returned reserve ----------
old=wb["Investor returns"]; idx=wb.sheetnames.index("Investor returns")
st={k:copy(old[a]._style) for k,a in dict(title="A1",band="A4",label="A5",link="B5",pct="B6",blue="B10",cur="C32",mult="C21",hdr="B14",yr="C31",sub="A35").items()}
wb.remove(old); R=wb.create_sheet("Investor returns",idx)
for k,w in dict(A=58,B=15,C=15,D=15,E=15,F=15,G=15).items(): R.column_dimensions[k].width=w
def put(a,v,s=None,fmt=None):
    c=R[a]; c.value=v
    if s: c._style=copy(st[s])
    if fmt: c.number_format=fmt
CUR='$#,##0;($#,##0);-'; MULT='0.00"x"'; PCT='0%'
put("A1","Investor returns: five years, cash only","title")
put("A3","Inputs","band")
for col in "BCDEFG": put(f"{col}3",None,"band")
put("A4","Investment ($)","label");put("B4","=Assumptions!B54","link",CUR)
put("A5","Equity offered to the investor (after payback)","label");put("B5","=Assumptions!B55","pct",PCT)
put("A6","Share of operating profit paid out as cash","label");put("B6","=Assumptions!B56","pct",PCT)
put("A7","Investor's share of cash until paid back / payback multiple","label");put("B7","=Assumptions!B57","pct",PCT);put("C7","=Assumptions!B58","link",MULT)
put("A9","Summary","band")
for col in "BCDEFG": put(f"{col}9",None,"band")
for col,t in zip("BCD",["Conservative","Base","Upside"]): put(f"{col}10",t,"hdr")
put("A10","Metric","label")
for i,t in enumerate(["Cash returned over 5 years, multiple of capital","Year 5 operating profit","Year the investor is paid back"]): put(f"A{11+i}",t,"label")
cases=[("Conservative","B",15),("Base","C",30),("Upside","D",45)]
for name,scol,r0 in cases:
    p,d=r0+2,r0+3
    put(f"A{r0}",f"{name} case","band")
    for col in "BCDEFG": put(f"{col}{r0}",None,"band")
    put(f"A{r0+1}","Year","label")
    for j,col in enumerate("BCDEFG"): put(f"{col}{r0+1}",j,"yr")
    put(f"A{p}","Operating profit after the owner's salary","label"); put(f"A{d}","Distributable cash","label")
    for j,col in enumerate("CDEFG"):
        a=openpyxl.utils.get_column_letter(2+12*j); b=openpyxl.utils.get_column_letter(13+12*j)
        put(f"{col}{p}",f"=SUM({name}!{a}30:{b}30)","cur",CUR)
        put(f"{col}{d}",f"=MAX(0,{col}{p})*$B$6","cur",CUR)
    put(f"A{r0+4}","Priority payback, then equity","sub")
    o,n,c,e=r0+5,r0+6,r0+7,r0+8
    put(f"A{o}","Priority amount still owed, start of year","label");put(f"A{n}","Distributions needed to clear it this year","label")
    put(f"A{c}","Investor cash from operations","label");put(f"A{e}","Priority amount still owed, end of year","label")
    for j,col in enumerate("CDEFG"):
        prev=openpyxl.utils.get_column_letter(2+j)
        put(f"{col}{o}","=$B$4*$C$7" if j==0 else f"={prev}{e}","cur",CUR)
        put(f"{col}{n}",f"=IF($B$7<=0,0,MIN({col}{d},{col}{o}/$B$7))","cur",CUR)
        put(f"{col}{c}",f"=$B$7*{col}{n}+$B$5*({col}{d}-{col}{n})","cur",CUR)
        put(f"{col}{e}",f"=MAX(0,{col}{o}-$B$7*{col}{n})","cur",CUR)
    put(f"{scol}11",f"=SUM(C{c}:G{c})/$B$4","cur",MULT)
    put(f"{scol}12",f"=G{p}","cur",CUR)
    put(f"{scol}13",f'=IFERROR(INDEX($C${r0+1}:$G${r0+1},MATCH(0,C{e}:G{e},0)),"Not within 5 years")',"cur","0")
    R[f"{scol}13"].alignment=openpyxl.styles.Alignment(horizontal="right")

# ---------- 2. Sources and uses: merge ramp + cushion, no staged funding ----------
S=wb["Sources and uses"]
snap={r:[(copy(S.cell(r,c)._style),S.cell(r,c).value) for c in range(1,6)] for r in range(15,37)}
for r in range(15,37):
    for c in range(1,6): S.cell(r,c).value=None
def sty(r,src):
    for c in range(1,6): S.cell(r,c)._style=copy(snap[src][c-1][0])
sty(15,15); S["A15"]="Operating losses until break-even, plus safety cushion"
S["B15"]="=Startup!B39+Startup!B40"; S["D15"]="=Startup!C39+Startup!C40"; S["C15"]="=IF($B$16=0,0,B15/$B$16)"
sty(16,17); S["A16"]="Total uses"; S["B16"]="=SUM(B5:B15)"; S["C16"]="=SUM(C5:C15)"; S["D16"]="=SUM(D5:D15)"
sty(17,18); S["A17"]="Check: equals Startup total capital required (should be 0)"; S["B17"]="=ROUND(B16-Startup!B41,2)"; S["D17"]="=ROUND(D16-Startup!C41,2)"
sty(19,20); S["A19"]="Sources"; S["B19"]="Amount"
sty(20,21); S["A20"]="Investor raise"; S["B20"]="=Assumptions!B54"
sty(21,24); S["A21"]="Total sources"; S["B21"]="=SUM(B20:B20)"
sty(22,25); S["A22"]="Buffer above the Conservative-sized uses"; S["B22"]="=B21-B16"
sty(23,26); S["A23"]="Buffer above the Base-case uses"; S["B23"]="=B21-D16"
for r in range(5,15): S[f"C{r}"]=f"=IF($B$16=0,0,B{r}/$B$16)"
S["A1"]="Sources and uses of funds"
SC=wb["Station capacity"]
SC["A17"]="Crowding buffer (share of peak capacity left open)"; SC["B17"]=0.1; SC["B17"]._style=copy(SC["B9"]._style); SC["A17"]._style=copy(SC["A9"]._style); SC["B17"].number_format="0%"
SC["A26"]="Member cap that keeps the buffer (rounded down)"; SC["B26"]="=ROUNDDOWN(B21*(1-B17),0)"; SC["A26"]._style=copy(SC["A25"]._style); SC["B26"]._style=copy(SC["B24"]._style)
# ---------- 3. Labels ----------
wb["Summary"]["A11"]="Buffer above capital required"
wb["Startup"]["A43"]="Buffer above capital required"
# ---------- 4. Ownership: drop exit-value rows; Assumptions exit multiple ----------
O=wb["Ownership"]
for r in list(range(15,22))+list(range(28,34)):
    for c in range(1,6): O.cell(r,c).value=None
A=wb["Assumptions"]; A["A59"]=None; A["B59"]=None
# contractor take-home rows no longer apply (dues only)
Sm=wb["Summary"]
for r in range(31,38):
    for c in range(1,6): Sm.cell(r,c).value=None
# ---------- 4b. Remove the TrackMan expansion logic: five hitting stations only ----------
for n in ("Base","Conservative","Upside"):
    ws=wb[n]
    for c in range(2,ws.max_column+1):
        L=openpyxl.utils.get_column_letter(c)
        ws[f"{L}5"]="=Assumptions!$B$17"
        ws[f"{L}22"]="=Assumptions!$B$41+Assumptions!$B$122"
        for r in range(45,51): ws[f"{L}{r}"].value=None
    for r in range(45,51): ws[f"A{r}"].value=None
    ws["A22"]="Software + billing + access control"
for n,col in (("Conservative","B"),("Base","C"),("Upside","D")): wb["Startup"][f"{col}39"]=f"={n}!B36"
St=wb["Startup"]
for r in (6,8):
    for c in range(1,6): St.cell(r,c).value=None
for r in (39,40,41):
    for c in range(1,6): wb["Summary"].cell(r,c).value=None
for r in (18,110,111,112,113,118,119,120):
    for c in range(1,3): A.cell(r,c).value=None
A["A117"]="Five hitting stations, 24/7 access, performance onboarding"
for row in wb["Dues only"].iter_rows():
    for c in row:
        if isinstance(c.value,str) and "Assumptions!$B$18" in c.value: c.value=c.value.replace("Assumptions!$B$18","Assumptions!$B$17")
# ---------- 5. Delete explanatory and scenario sheets ----------
for n in ("Guide","Changes and flags","Ownership scenarios"):
    if n in wb.sheetnames: wb.remove(wb[n])
# ---------- 6. Strip explanatory text ----------
notes={"Summary":("E",None),"Sources and uses":("E",None),"Unit economics":("E",None),"Station capacity":("H",None),
 "Space plan":("E",None),"Sensitivity":("K",None),"Startup":("CE",None),"Equipment list":("K",None),"Ownership":("E",27),
 "Assumptions":("E",None),"Dues only":("C",None)}
def isfill(c): return c.fill is not None and c.fill.fill_type not in (None,"none")
for sh,(cols,maxr) in notes.items():
    ws=wb[sh]
    for col in cols:
        for r in range(1,(maxr or ws.max_row)+1):
            c=ws[f"{col}{r}"]
            if isinstance(c.value,str) and not c.value.startswith("="): c.value=None
for ws in wb:
    if ws.title in ("Risks and checklist","Founding members","Scripts"): continue
    for r in range(2,ws.max_row+1):
        c=ws[f"A{r}"]; v=c.value
        if not isinstance(v,str) or isfill(c): continue
        prose=(len(v)>60 and not v.startswith("=")) or v.startswith('="') or v.startswith("•")
        if prose and all(ws.cell(r,k).value is None for k in range(2,ws.max_column+1)): c.value=None
wb.save("stage3b.xlsx"); print("simplified")
