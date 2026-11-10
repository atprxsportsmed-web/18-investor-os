import re
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L
wb=load_workbook("stage3.xlsx")
F=lambda **k: Font(name="Arial",size=k.pop("size",10),**k)
BLK=F(); BOLD=F(bold=True); GREY=F(size=9,color="FF555555"); GREYI=F(size=9,italic=True,color="FF555555")
BLUE=F(color="FF0000FF"); GREEN=F(color="FF008000"); TITLE=F(size=14,bold=True); WH=F(bold=True,color="FFFFFFFF")
HDR=PatternFill("solid",fgColor="FFD9E1F2"); YEL=PatternFill("solid",fgColor="FFFFFF00"); DARK=PatternFill("solid",fgColor="FF1F3864")
CUR='$#,##0;($#,##0);-'; PCT='0.0%;(0.0%);-'; MULT='0.00"x"'
O=wb.create_sheet("Ownership scenarios",index=wb.sheetnames.index("Ownership")+1)
for k,v in {"A":58,"B":16,"C":16,"D":16,"E":16,"F":16,"G":16,"H":14,"I":14}.items(): O.column_dimensions[k].width=v
def put(ref,val,font=BLK,fmt=None,fill=None,wrap=False,align=None):
    c=O[ref]; c.value=val; c.font=font
    if fmt: c.number_format=fmt
    if fill: c.fill=fill
    if wrap or align: c.alignment=Alignment(wrap_text=wrap,vertical="top",horizontal=align)
def band(r,t,fill=HDR,font=BOLD):
    for c in range(1,10): O.cell(r,c).fill=fill
    O.cell(r,1,t).font=font
IR="'Investor returns'!"
put("A1","Ownership scenarios: owner pay and investor return at different stakes",TITLE)
put("A2","Tests a structure where the investor holds the larger share at the start and the owner buys shares back over time. Ownership percent and decision rights are separate: control must be set in the operating agreement.",GREYI)
band(4,"Inputs")
inp=[(5,"Raise ($)","=Assumptions!B54",CUR,GREEN,"From Assumptions B54"),
(6,"Owner salary after opening ($/yr)","=Assumptions!B32",CUR,GREEN,"Paid regardless of ownership percent"),
(7,"Share of profit distributed","=Assumptions!B56",PCT,GREEN,"Remainder stays in the business"),
(8,"Buyout price multiple (x annual operating profit)","=Assumptions!B59",'0.0"x"',GREEN,"Price of 100% of the company = multiple x that year's operating profit"),
(9,"Case shown in the buyout schedule","Base",None,BLUE,"Conservative, Base or Upside"),
(10,"Owner starting stake for the buyout schedule",0.3,PCT,BLUE,"Investor holds the remainder"),
(11,"Stake bought back each year (% of company)",0.05,PCT,BLUE,"Bought from the investor at the formula price"),
(12,"First buyout at the end of year",2,"0",BLUE,"Books are closed and profit is known at year end"),
(13,"Tax rate on owner distributions",0.25,PCT,BLUE,"Estimate. Distributions are taxable even when used for buyouts; confirm with a CPA")]
for r,a,v,fm,fo,n in inp:
    put(f"A{r}",a); put(f"B{r}",v,fo,fm); put(f"C{r}",n,GREY)
    if fo is BLUE: O[f"B{r}"].fill=YEL
dv=DataValidation(type="list",formula1='"Conservative,Base,Upside"',allow_blank=False); O.add_data_validation(dv); dv.add("B9")

band(15,"Year 3 owner pay by starting stake, no buyouts")
put("A16","Owner starting stake",BOLD)
for col,v,fo in (("B",0.2,BLUE),("C",0.3,BLUE),("D",0.4,BLUE),("E","=1-Assumptions!B55",GREEN)):
    put(f"{col}16",v,fo,PCT)
    if fo is BLUE: O[f"{col}16"].fill=YEL
put("F16","Year 3 operating profit",BOLD,align="right"); O["F16"].alignment=Alignment(wrap_text=True,horizontal="right")
put("A17","Column E is the structure in the model as built (investor holds the equity offered)",GREYI)
put("A18","Investor stake")
put("A19","Implied company value after the investment")
put("A20","Implied company value before the investment")
for col in "BCDE":
    put(f"{col}18",f"=1-{col}16",BLK,PCT); put(f"{col}19",f"=$B$5/{col}18",BLK,CUR); put(f"{col}20",f"={col}19-$B$5",BLK,CUR)
put("A21","Owner distribution, Year 3",BOLD)
put("A25","Total owner pay, Year 3 (salary plus distribution)",BOLD)
prow={"Conservative":32,"Base":57,"Upside":82}
for i,(n,pr) in enumerate(prow.items()):
    put(f"A{22+i}",n); put(f"A{26+i}",n)
    put(f"F{22+i}",f"={IR}E{pr}",GREEN,CUR)
    for col in "BCDE":
        put(f"{col}{22+i}",f"={col}$16*MAX(0,$F{22+i})*$B$7",BLK,CUR)
        put(f"{col}{26+i}",f"=$B$6+{col}{22+i}",BLK,CUR)
put("A29","The owner's pay is mostly salary at low stakes. The value of a low starting stake comes from the buyouts below.",GREYI)

band(31,"Buyout schedule")
put("A32","Case and starting stake shown",BOLD); put("B32",'=B9&", "&TEXT(B10,"0%")&" owner stake"',BOLD)
put("A33","Year",BOLD)
for k in range(6): put(f"{L(2+k)}33",k,BOLD,"0",align="right")
def sel(cols_rows):  # (cons,base,up) row refs in IR for col
    return lambda col: f'IF($B$9="Conservative",{IR}{col}{cols_rows[0]},IF($B$9="Base",{IR}{col}{cols_rows[1]},{IR}{col}{cols_rows[2]}))'
profit=sel((32,57,82)); res=sel((53,78,103)); exitv=sel((34,59,84))
lab={34:"Operating profit",35:"Distributable cash",36:"Owner stake, start of year",37:"Owner distribution",38:"Owner salary",39:"Total owner pay",
40:"Stake bought back at year end",41:"Buyout cost (stake x multiple x that year's profit)",42:"Owner distribution after tax",43:"Cash left after paying for the buyout",44:"Buyout covered by distributions?",45:"Owner stake, end of year"}
for r,t in lab.items(): put(f"A{r}",t,BOLD if r in (39,45) else BLK)
for k in range(5):
    c=L(3+k); p=L(2+k)
    put(f"{c}34",f"={profit(c)}",GREEN,CUR)
    put(f"{c}35",f"=MAX(0,{c}34)*$B$7",BLK,CUR)
    put(f"{c}36","=$B$10" if k==0 else f"={p}45",BLK,PCT)
    put(f"{c}37",f"={c}36*{c}35",BLK,CUR)
    put(f"{c}38","=$B$6",BLK,CUR)
    put(f"{c}39",f"={c}37+{c}38",BOLD,CUR)
    put(f"{c}40",f"=IF({k+1}>=$B$12,MIN($B$11,1-{c}36),0)",BLK,PCT)
    put(f"{c}41",f"={c}40*$B$8*MAX(0,{c}34)",BLK,CUR)
    put(f"{c}42",f"={c}37*(1-$B$13)",BLK,CUR)
    put(f"{c}43",f"={c}42-{c}41",BLK,CUR)
    put(f"{c}44",f'=IF({c}41=0,"n/a",IF({c}43>=0,"Yes","Shortfall"))',BLK,align="right")
    put(f"{c}45",f"={c}36+{c}40",BOLD,PCT)
put("A46","Five-year totals",BOLD)
put("A47","Owner pay over five years"); put("B47","=SUM(C39:G39)",BLK,CUR)
put("A48","Owner stake at the end of Year 5"); put("B48","=G45",BLK,PCT)
put("A49","Value of owner stake at the Year 5 exit value"); put("B49",f"=G45*{exitv('G')}",BLK,CUR)

band(51,"Investor result in the same case")
put("A52","Year",BOLD)
for k in range(6): put(f"{L(2+k)}52",k,BOLD,"0",align="right")
put("A53","Investor stake, start of year"); put("A54","Investor cash from distributions"); put("A55","Investor cash from selling shares to the owner"); put("A56","Unspent raise returned"); put("A57","Exit value at the end of Year 5"); put("A58","Investor cash flows, with exit",BOLD)
put("B58","=-$B$5",BOLD,CUR)
for k in range(5):
    c=L(3+k)
    put(f"{c}53",f"=1-{c}36",BLK,PCT); put(f"{c}54",f"={c}53*{c}35",BLK,CUR); put(f"{c}55",f"={c}41",BLK,CUR); put(f"{c}56",f"={res(c)}",GREEN,CUR)
    if k==4: put("G57",f"={exitv('G')}",GREEN,CUR)
    put(f"{c}58",f"={c}54+{c}55+{c}56"+("+(1-G45)*G57" if k==4 else ""),BOLD,CUR)
put("A59","Investor IRR with exit"); put("B59",'=IFERROR(IRR(B58:G58),"n/a")',BLK,PCT)
put("A60","Investor multiple of capital returned with exit"); put("B60","=SUM(C58:G58)/$B$5",BLK,MULT)

band(62,"Investor result by starting owner stake (same case, no buyouts)")
put("A63","Year",BOLD)
for k in range(6): put(f"{L(2+k)}63",k,BOLD,"0",align="right")
put("H63","IRR",BOLD,align="right"); put("I63","Multiple",BOLD,align="right")
for i,col in enumerate("BCDE"):
    r=64+i
    put(f"A{r}",f'="Owner "&TEXT({col}16,"0%")&", investor "&TEXT({col}18,"0%")')
    put(f"B{r}","=-$B$5",BLK,CUR)
    for k in range(5):
        c=L(3+k)
        put(f"{c}{r}",f"={col}$18*{c}$35+{c}$56"+(f"+{col}$18*$G$57" if k==4 else ""),BLK,CUR)
    put(f"H{r}",f'=IFERROR(IRR(B{r}:G{r}),"n/a")',BLK,PCT); put(f"I{r}",f"=SUM(C{r}:G{r})/$B$5",BLK,MULT)
put("A69","A smaller investor stake lowers the investor's return for the same money. The stake at which the investor stops accepting the deal is the practical ceiling on how low the owner stake can go.",GREYI)

band(71,"Investor stake you could concede: owner keeps 51% or more (selected case, no buyouts)")
put("A72","Year",BOLD)
for k in range(6): put(f"{L(2+k)}72",k,BOLD,"0",align="right")
put("H72","IRR (exit)",BOLD,align="right"); put("I72","Multiple (exit)",BOLD,align="right"); put("J72","Multiple (cash only)",BOLD,align="right"); put("K72","Investor stake (input)",BOLD,align="right")
for i,sv in enumerate([0.20,0.25,0.30,0.40,0.49]):
    r=73+i
    put(f"K{r}",sv,BLUE,"0%")
    put(f"A{r}",f'="Investor "&TEXT(K{r},"0%")&", owner keeps "&TEXT(1-K{r},"0%")')
    put(f"B{r}","=-$B$5",BLK,CUR)
    for k in range(5):
        c=L(3+k)
        put(f"{c}{r}",f"=$K{r}*{c}$35+{c}$56"+(f"+$K{r}*$G$57" if k==4 else ""),BLK,CUR)
    put(f"H{r}",f'=IFERROR(IRR(B{r}:G{r}),"n/a")',BLK,PCT); put(f"I{r}",f"=SUM(C{r}:G{r})/$B$5",BLK,MULT)
    put(f"J{r}",f"=(SUM(C{r}:G{r})-$K{r}*$G$57)/$B$5",BLK,MULT)
put("A79","Negotiating frame: open around 20 to 25%, work around 30%, ceiling about 40%, and keep at least 51% as owner. Protective rights (vetoes on a sale, new debt or owner pay) matter as much as the percentage. Securities counsel should review all terms.",GREYI)

band(82,"Ownership after the investor and the strength coach (selected case)")
put("A83","Strength coach equity (Assumptions B206)"); put("B83","=Assumptions!$B$206",GREEN,"0.0%")
put("A84","Investor stake",BOLD); put("B84","Strength coach",BOLD,align="right"); put("C84","Owner keeps",BOLD,align="right"); put("D84","Owner at 51% or more?",BOLD,align="right")
for i in range(5):
    r=85+i
    put(f"A{r}",f"=K{73+i}",BLK,"0%"); put(f"B{r}","=$B$83",BLK,"0.0%"); put(f"C{r}",f"=1-A{r}-B{r}",BLK,"0.0%")
    put(f"D{r}",f'=IF(C{r}>=0.51,"Yes","No")',BOLD,align="right")
put("A90","Winston's equity is assumed to come out of the owner's side, so investor cash flows above do not change. Equity should be earned and vest over time (Assumptions B207) and return if he leaves. Securities counsel should draft it.",GREYI)
O.freeze_panes="B4"
import sys; sys.path.insert(0,"."); import clean
clean.run(wb)
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value,str) and "Driven" in c.value:
                c.value=c.value.replace("Driven Golf Performance","Eighteen Athletic Club").replace("Driven Golf","Eighteen Athletic Club")
wb.properties.title="Eighteen Athletic Club financial model"
RK=wb["Risks and checklist"]
from copy import copy
for col,txt in zip("ABCD",["Bay capacity at peak","Two bays carry a limited number of members at peak times. The member cap can exceed what the bays can serve, and members who cannot book leave.","Peak pricing or booking limits, off-peak incentives, a lower advertised cap, or a third bay. Replace the peak-demand estimates with booking data after opening.","Bay utilization tab, Assumptions B17 and B118"]):
    c=RK[f"{col}16"]; c.value=txt; src=RK[f"{col}15"]
    c.font=copy(src.font); c.alignment=copy(src.alignment); c.border=copy(src.border); c.fill=copy(src.fill)
RK.row_dimensions[16].height=RK.row_dimensions[15].height
wb.save("stage3b.xlsx")
