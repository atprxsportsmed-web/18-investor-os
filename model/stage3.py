import json, math
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter as L
wb=load_workbook("stage1.xlsx")
F=lambda **k: Font(name="Arial",size=k.pop("size",10),**k)
BLK=F(); BOLD=F(bold=True); GREY=F(size=9,color="FF555555"); GREYI=F(size=9,italic=True,color="FF555555")
BLUE=F(color="FF0000FF"); GREEN=F(color="FF008000"); TITLE=F(size=14,bold=True)
HDR=PatternFill("solid",fgColor="FFD9E1F2"); YEL=PatternFill("solid",fgColor="FFFFFF00"); DARK=PatternFill("solid",fgColor="FF1F3864")
CUR='$#,##0;($#,##0);-'; PCT='0.0%;(0.0%);-'; NUM='#,##0.0;(#,##0.0);-'; INT='#,##0;(#,##0);-'; MULT='0.00"x"'
def est_h(text,width,base=13.5):
    if not text: return base
    lines=sum(max(1,math.ceil(len(p)/max(1,width*1.05))) for p in str(text).split("\n"))
    return max(base,base*lines)
def put(ws,ref,val,font=BLK,fmt=None,fill=None,wrap=False,align=None):
    c=ws[ref]; c.value=val; c.font=font
    if fmt: c.number_format=fmt
    if fill: c.fill=fill
    if wrap or align: c.alignment=Alignment(wrap_text=wrap,vertical="top",horizontal=align)
    return c
def band(ws,row,ncols,text=None,fill=HDR,font=BOLD):
    for c in range(1,ncols+1): ws.cell(row,c).fill=fill
    if text is not None: ws.cell(row,1,text).font=font
def widths(ws,ws_w):
    for k,v in ws_w.items(): ws.column_dimensions[k].width=v

# =============== INVESTOR RETURNS ===============
R=wb.create_sheet("Investor returns")
widths(R,{"A":58,"B":15,"C":15,"D":15,"E":15,"F":15,"G":15,"H":15,"I":70})
put(R,"A1","Investor returns: five-year view, three ways to pay the investor",TITLE)
put(R,"A2","Uses the monthly sheets extended to month 60. Distributable cash is a share of each year's positive operating profit (after the owner's $60K salary and the GM's $45K base). A loss year pays nothing and is not clawed back. Structure A is straight equity: the investor receives the equity share of every distribution and of the exit value.",GREYI)
band(R,4,9,"Inputs"); 
rows=[("Investment ($)","=Assumptions!B54",CUR,"You said: $450K, non-loan"),
("Equity offered to the investor","=Assumptions!B55",PCT,"You said: 20%"),
("Share of operating profit paid out as cash","=Assumptions!B56",PCT,"My assumption: the rest stays in the business for reserves and equipment"),
("Exit value as a multiple of Year 5 operating profit","=Assumptions!B59",'0.0"x"',"My assumption: investors will question this. Small businesses commonly sell for low single-digit multiples of profit; verify with a broker")]
for i,(a,f,fmt,n) in enumerate(rows):
    r=5+i; put(R,f"A{r}",a); put(R,f"B{r}",f,GREEN,fmt); put(R,f"I{r}",n,GREY)
put(R,"A9","Structure B: investor's priority share of cash / payback multiple"); put(R,"B9","=Assumptions!B57",GREEN,PCT); put(R,"C9","=Assumptions!B58",GREEN,'0.0"x"'); put(R,"I9","Investor gets the priority share of distributions until the multiple of capital is returned, then only the equity share. Set on Assumptions B57 and B58",GREY)
put(R,"A10","Structure C: investor's priority share of cash / payback multiple"); put(R,"B10",0.5,BLUE,PCT,YEL); put(R,"C10",1.5,BLUE,'0.0"x"',YEL); put(R,"I10","Same idea at a higher multiple. Edit here to test other terms",GREY)
put(R,"A11","Share of unspent raise returned to the investor"); put(R,"B11",1,BLUE,PCT,YEL); put(R,"I11","Raise minus the capital that case actually needs (Summary row 11). Returned as a return of capital; their equity stays the same. Terms to agree in the documents",GREY)
put(R,"A12","Year the reserve is returned (1 to 5)"); put(R,"B12",2,BLUE,"0",YEL); put(R,"I12","Paid at the end of that year. A later year protects you if the ramp is slow; an earlier year is better for the investor",GREY)
# summary table placeholders (filled after blocks)
band(R,13,9,"Summary of investor returns (includes any reserve returned)"); 
for j,h in enumerate(["Conservative","Base","Upside"]): put(R,f"{L(2+j)}14",h,F(bold=True,color="FFFFFFFF"),fill=DARK,align="center")
put(R,"A14","Metric",F(bold=True,color="FFFFFFFF"),fill=DARK)
sums=["IRR with exit: A, straight equity","IRR with exit: B, priority payback","IRR with exit: C, priority to higher multiple",
"Multiple of capital returned with exit: A","Multiple of capital returned with exit: B","Multiple of capital returned with exit: C",
"Multiple of capital returned from cash only (no exit): A","Multiple of capital returned from cash only (no exit): B","Multiple of capital returned from cash only (no exit): C",
"Year 5 operating profit","Year 5 exit value (profit x multiple)"]
for i,s in enumerate(sums): put(R,f"A{15+i}",s)
put(R,"A27","IRR is the yearly return that makes the investor's cash flows net to zero. It is only meaningful when the exit is real: a buyer, or a buyback you have agreed. 'n/a' means the cash flows never turn positive.",GREYI)
blocks={}
scen=[("Conservative","Conservative"),("Base","Base"),("Upside","Upside")]
yr=[("B","M"),("N","Y"),("Z","AK"),("AL","AW"),("AX","BI")]
r0=30
for si,(title,sn) in enumerate(scen):
    b={}
    band(R,r0,9,f"{title} case")
    put(R,f"A{r0+1}","Year",BOLD)
    for k in range(6): put(R,f"{L(2+k)}{r0+1}",k,BOLD,"0",align="right")
    b["profit"]=r0+2; b["dist"]=r0+3; b["exit"]=r0+4; b["res"]=r0+23
    put(R,f"A{b['res']}","Unspent raise returned to investor (return of capital)")
    for k in range(5):
        col=L(3+k); put(R,f"{col}{b['res']}",f"=IF({k+1}=$B$12,MAX(0,Summary!{L(2+si)}11)*$B$11,0)",GREEN,CUR)
    put(R,f"A{b['profit']}","Operating profit after your salary")
    for k,(c1,c2) in enumerate(yr): put(R,f"{L(3+k)}{b['profit']}",f"=SUM({sn}!{c1}30:{c2}30)",GREEN,CUR)
    put(R,f"A{b['dist']}","Distributable cash")
    for k in range(5): col=L(3+k); put(R,f"{col}{b['dist']}",f"=MAX(0,{col}{b['profit']})*$B$7",BLK,CUR)
    put(R,f"A{b['exit']}","Exit value at the end of Year 5"); put(R,f"G{b['exit']}",f"=MAX(0,G{b['profit']})*$B$8",BLK,CUR)
    # structure A
    put(R,f"A{r0+5}","Structure A: straight equity",BOLD)
    b["Acash"]=r0+6; b["Aops"]=r0+7; b["Aex"]=r0+8
    put(R,f"A{b['Acash']}","Investor cash from operations")
    for k in range(5): col=L(3+k); put(R,f"{col}{b['Acash']}",f"=$B$6*{col}{b['dist']}",BLK,CUR)
    put(R,f"A{b['Aops']}","Investor cash flows, cash only"); put(R,f"B{b['Aops']}","=-$B$5",BLK,CUR)
    put(R,f"A{b['Aex']}","Investor cash flows, with exit"); put(R,f"B{b['Aex']}","=-$B$5",BLK,CUR)
    for k in range(5):
        col=L(3+k); put(R,f"{col}{b['Aops']}",f"={col}{b['Acash']}+{col}{b['res']}",BLK,CUR)
        put(R,f"{col}{b['Aex']}",f"={col}{b['Acash']}+{col}{b['res']}"+(f"+$B$6*G{b['exit']}" if k==4 else ""),BLK,CUR)
    # structures B, C
    for tag,start,pc,mc,label in (("B",r0+9,"$B$9","$C$9","Structure B: priority payback, then equity"),("C",r0+16,"$B$10","$C$10","Structure C: priority to the higher multiple, then equity")):
        put(R,f"A{start}",label,BOLD)
        rows_={"owed":start+1,"thr":start+2,"cash":start+3,"end":start+4,"ops":start+5,"ex":start+6}
        b[tag]=rows_
        put(R,f"A{rows_['owed']}","Priority amount still owed, start of year")
        put(R,f"A{rows_['thr']}","Distributions needed to clear it this year")
        put(R,f"A{rows_['cash']}","Investor cash from operations")
        put(R,f"A{rows_['end']}","Priority amount still owed, end of year")
        put(R,f"A{rows_['ops']}","Investor cash flows, cash only"); put(R,f"B{rows_['ops']}","=-$B$5",BLK,CUR)
        put(R,f"A{rows_['ex']}","Investor cash flows, with exit"); put(R,f"B{rows_['ex']}","=-$B$5",BLK,CUR)
        for k in range(5):
            col=L(3+k); prev=L(2+k)
            put(R,f"{col}{rows_['owed']}",(f"=$B$5*{mc}" if k==0 else f"={prev}{rows_['end']}"),BLK,CUR)
            put(R,f"{col}{rows_['thr']}",f"=IF({pc}<=0,0,MIN({col}{b['dist']},{col}{rows_['owed']}/{pc}))",BLK,CUR)
            put(R,f"{col}{rows_['cash']}",f"={pc}*{col}{rows_['thr']}+$B$6*({col}{b['dist']}-{col}{rows_['thr']})",BLK,CUR)
            put(R,f"{col}{rows_['end']}",f"=MAX(0,{col}{rows_['owed']}-{pc}*{col}{rows_['thr']}-{col}{b['res']})",BLK,CUR)
            put(R,f"{col}{rows_['ops']}",f"={col}{rows_['cash']}+{col}{b['res']}",BLK,CUR)
            put(R,f"{col}{rows_['ex']}",f"={col}{rows_['cash']}+{col}{b['res']}"+(f"+$B$6*G{b['exit']}" if k==4 else ""),BLK,CUR)
    blocks[sn]=b
    # summary links
    col=L(2+si)
    def irr(row): return f'=IFERROR(IRR(B{row}:G{row}),"n/a")'
    put(R,f"{col}15",irr(b["Aex"]),BLK,PCT); put(R,f"{col}16",irr(b["B"]["ex"]),BLK,PCT); put(R,f"{col}17",irr(b["C"]["ex"]),BLK,PCT)
    put(R,f"{col}18",f"=SUM(C{b['Aex']}:G{b['Aex']})/$B$5",BLK,MULT); put(R,f"{col}19",f"=SUM(C{b['B']['ex']}:G{b['B']['ex']})/$B$5",BLK,MULT); put(R,f"{col}20",f"=SUM(C{b['C']['ex']}:G{b['C']['ex']})/$B$5",BLK,MULT)
    put(R,f"{col}21",f"=SUM(C{b['Aops']}:G{b['Aops']})/$B$5",BLK,MULT); put(R,f"{col}22",f"=SUM(C{b['B']['ops']}:G{b['B']['ops']})/$B$5",BLK,MULT); put(R,f"{col}23",f"=SUM(C{b['C']['ops']}:G{b['C']['ops']})/$B$5",BLK,MULT)
    put(R,f"{col}24",f"=G{b['profit']}",BLK,CUR); put(R,f"{col}25",f"=G{b['exit']}",BLK,CUR)
    r0+=25
R.freeze_panes="B5"

# =============== SOURCES AND USES ===============
U=wb.create_sheet("Sources and uses")
widths(U,{"A":58,"B":18,"C":12,"D":18,"E":80})
put(U,"A1","Sources and uses of funds",TITLE)
put(U,"A2","Two views of the same uses: sized to the Conservative case (a slow ramp) and to the Base case (a normal ramp). The raise is compared with both in rows 25 and 26.",GREYI)
band(U,4,5); 
for j,h in enumerate(["Uses","Conservative-sized","% of total","Base case","Notes"]): put(U,f"{L(1+j)}4",h,F(bold=True,color="FFFFFFFF"),fill=DARK,align=None if j==0 else "center")
uses=[("Golf area: TrackMan bays, hitting stations, net and partition, flooring","=Startup!B6+Startup!B7+Startup!B8+Startup!B9","Startup rows 6 to 9"),
("Gym floor equipment and storage","=Startup!B11+Startup!B12+Startup!B13+Startup!B14","Startup rows 11 to 14, itemized on Equipment list. Planning number about $38K"),
("Testing, treatment and recovery","=Startup!B16+Startup!B17+Startup!B18","Startup rows 16 to 18"),
("Space build-out","=Startup!B20","Startup row 20. The biggest swing: get a contractor quote"),
("Lobby furniture and pro shop stock","=Startup!B21","Startup row 21"),
("Start-up contingency","=Startup!B23","Startup row 23: 10% of equipment and build-out"),
("Lease deposit, first month and any pre-opening rent","=Startup!B25","Startup row 25"),
("Legal, permits and insurance down payments","=Startup!B26+Startup!B27","Startup rows 26 and 27. Add securities counsel for the raise"),
("Founding-member campaign, brand, website, booking and access control","=Startup!B28+Startup!B29","Startup rows 28 and 29"),
("Your pay during the build","=Startup!B31+Startup!B32+Startup!B33","Startup rows 31 to 33. $55K a year pro-rated over the paid build months"),
("Operating losses until break-even (ramp)","=Startup!B39","From the monthly model: the deepest cumulative loss"),
("Safety cushion","=Startup!B40","Months of operating costs held in reserve")]
basecells=["=Startup!B6+Startup!B7+Startup!B8+Startup!B9","=Startup!B11+Startup!B12+Startup!B13+Startup!B14","=Startup!B16+Startup!B17+Startup!B18","=Startup!B20","=Startup!B21","=Startup!B23","=Startup!B25","=Startup!B26+Startup!B27","=Startup!B28+Startup!B29","=Startup!B31+Startup!B32+Startup!B33","=Startup!C39","=Startup!C40"]
for i,((a,f,n),bf) in enumerate(zip(uses,basecells)):
    r=5+i; put(U,f"A{r}",a); put(U,f"B{r}",f,GREEN,CUR); put(U,f"C{r}",f"=IF($B$17=0,0,B{r}/$B$17)",BLK,PCT); put(U,f"D{r}",bf,GREEN,CUR); put(U,f"E{r}",n,GREY)
put(U,"A17","Total uses",BOLD); put(U,"B17","=SUM(B5:B16)",BOLD,CUR); put(U,"C17","=SUM(C5:C16)",BOLD,PCT); put(U,"D17","=SUM(D5:D16)",BOLD,CUR)
put(U,"A18","Check: equals Startup total capital required (should be 0)",GREY); put(U,"B18","=ROUND(B17-Startup!B41,2)",GREY,CUR); put(U,"D18","=ROUND(D17-Startup!C41,2)",GREY,CUR)
band(U,20,5)
for j,h in enumerate(["Sources","Amount","","",""]): put(U,f"{L(1+j)}20",h,F(bold=True,color="FFFFFFFF"),fill=DARK)
put(U,"A21","Investor raise"); put(U,"B21","=Assumptions!B54",GREEN,CUR); put(U,"E21","You said: $450K, non-loan",GREY)
put(U,"A22","Your cash contribution"); put(U,"B22",0,BLUE,CUR,YEL); put(U,"E22","Your rule so far: the business funds itself. Enter an amount only if that changes",GREY)
put(U,"A23","Founding-member money available before opening"); put(U,"B23",0,BLUE,CUR,YEL); put(U,"E23","Refundable deposits are a liability, not income, so they are not counted by default. Enter only money you can spend",GREY)
put(U,"A24","Total sources",BOLD); put(U,"B24","=SUM(B21:B23)",BOLD,CUR)
put(U,"A25","Surplus / (shortfall) against the Conservative-sized uses",BOLD); put(U,"B25","=B24-B17",BOLD,CUR)
put(U,"A26","Surplus / (shortfall) against the Base-case uses",BOLD); put(U,"B26","=B24-D17",BOLD,CUR)
band(U,28,5)
for j,h in enumerate(["Funding in stages (suggested)","Amount needed","% of raise","Cumulative","Release trigger and what it covers"]): put(U,f"{L(1+j)}28",h,F(bold=True,color="FFFFFFFF"),fill=DARK,align=None if j==0 else "center")
tr=[("Tranche 1: at close","=Startup!B26+Startup!B27+Startup!B28+Startup!B29+Startup!B31+Startup!B32+Startup!B33","Trigger: entity, DBA and operating agreement filed, securities counsel engaged. Covers legal, permits, brand, the founding-member campaign and your build pay"),
("Tranche 2: lease and deposits","=Startup!B25+Startup!B20+Startup!B9","Trigger: lease signed and at least 15 founding-member commitments (signed intent, no money collected). Covers the lease deposit, build-out and flooring"),
("Tranche 3: equipment and opening","=Startup!B6+Startup!B7+Startup!B8+Startup!B11+Startup!B12+Startup!B13+Startup!B14+Startup!B16+Startup!B17+Startup!B18+Startup!B21+Startup!B23","Trigger: build-out complete and all 25 founding slots committed. Covers TrackMan bays, gym, testing, recovery, lobby and contingency"),
("Reserve: raise less tranches 1 to 3","=$B$24-SUM(B29:B31)","Held in the account and drawn monthly as needed for ramp losses. Whatever the ramp does not use above the capital needed is returned to the investor (Investor returns B11 and B12)")]
for i,(a,f,n) in enumerate(tr):
    r=29+i; put(U,f"A{r}",a); put(U,f"B{r}",f,GREEN,CUR); put(U,f"C{r}",f"=IF($B$33=0,0,B{r}/$B$33)",BLK,PCT)
    put(U,f"D{r}",(f"=B{r}" if i==0 else f"=D{r-1}+B{r}"),BLK,CUR); put(U,f"E{r}",n,GREY,wrap=True); U.row_dimensions[r].height=est_h(n,80)
put(U,"A33","Total",BOLD); put(U,"B33","=SUM(B29:B32)",BOLD,CUR); put(U,"C33","=SUM(C29:C32)",BOLD,PCT)
put(U,"A34","Check: equals total sources (should be 0)",GREY); put(U,"B34","=ROUND(B33-B24,2)",GREY,CUR)
put(U,"A36","Staged amounts split the actual raise. The Conservative case needs more than the raise (see row 25), so a slow ramp would need the amount in row 25 from another source. Staging lowers the investor's risk because money is released against proof, and it can support a better valuation. Triggers are suggestions: agree real ones with counsel.",GREYI)

# =============== UNIT ECONOMICS ===============
E=wb.create_sheet("Unit economics")
widths(E,{"A":64,"B":16,"C":16,"D":16,"E":84})
WH=F(bold=True,color="FFFFFFFF")
put(E,"A1","Unit economics: what one member is worth and costs to win",TITLE)
put(E,"A2","Per regular member (not founding). Conservative framing: acquisition cost includes launch spend and selling time, and lifetime value is capped at 12 and 24 months.",GREYI)
band(E,4,5,fill=DARK)
put(E,"A4","Metric",WH,fill=DARK)
for j,h in enumerate(["Conservative","Base","Upside"]): put(E,f"{L(2+j)}4",h,WH,fill=DARK,align="center")
put(E,"E4","Notes",WH,fill=DARK)
sc=["B","C","D"]
def ue_rows(rows):
    for r,a,f,fmt,n,kind in rows:
        if f is None:
            band(E,r,5,a); continue
        put(E,f"A{r}",a,BOLD if kind=="key" else BLK)
        if kind=="input":
            put(E,f"B{r}",f,BLUE,fmt,YEL)
            for j in (1,2): put(E,f"{L(2+j)}{r}",f"=$B${r}",BLK,fmt)
        else:
            for j in range(3):
                c=L(2+j); ff=f.replace("{c}",c).replace("{s}",sc[j])
                put(E,f"{c}{r}",ff,GREEN if (ff.startswith("=Assumptions") or ff.startswith("=Summary")) else BLK,fmt)
        put(E,f"E{r}",n,GREY)
X='0.0"x"'
ue_rows([
(5,"1. Contribution per regular member",None,None,None,None),
(6,"Blended dues per regular member ($/month)","=Assumptions!$B$14",CUR,"Mix of base, premium and junior tiers",""),
(7,"Facility share of paid add-ons per regular member ($/month)","=Assumptions!$B$107",CUR,"The facility's 20% of Winston's and the pro's billings, weighted for juniors",""),
(8,"Revenue per regular member ($/month)","={c}6+{c}7",CUR,"",""),
(9,"Card fees and follow-up coaching ($/month)","=-({c}8*Assumptions!$B$45+Assumptions!$B$191)",CUR,"3% card fees plus follow-up coaching",""),
(10,"Contribution per regular member ($/month)","={c}8+{c}9",CUR,"Before fixed costs","key"),
(12,"2. Cost to acquire a member (fully loaded)",None,None,None,None),
(13,"Launch spend before opening ($)",12000,CUR,"Estimate: about $12K for campaign, brand and website. Startup rows 28 and 29 hold $18K including access control and cameras","input"),
(14,"Members at opening","=Assumptions!${s}$21",NUM,"Founding members and others signed before the doors open",""),
(15,"Launch cost per opening member","=IF({c}14<=0,\"n/a\",{c}13/{c}14)",CUR,"",""),
(16,"Ongoing marketing per month","=Assumptions!$B$42",CUR,"Assumptions row 42",""),
(17,"New members per month after opening","=Assumptions!${s}$22",NUM,"",""),
(18,"Marketing cost per new member","=IF({c}17<=0,\"n/a\",{c}16/{c}17)",CUR,"",""),
(19,"Selling hours per new member",3,NUM,"Estimate: tour, trial session, follow-up and onboarding. Track real hours after opening","input"),
(20,"Paid hours per year",2080,INT,"Full-time year, used to convert salary to an hourly cost","input"),
(21,"Owner hourly cost, loaded ($/hr)","=Assumptions!$B$32*(1+Assumptions!$B$35)/{c}20",CUR,"Owner salary with payroll load / paid hours. The owner's selling time is not free",""),
(22,"Selling time cost per new member","={c}19*{c}21",CUR,"",""),
(23,"Fully loaded CAC after opening","=IF({c}18=\"n/a\",\"n/a\",{c}18+{c}22)",CUR,"Marketing per new member plus selling time","key"),
(24,"Blended CAC, first 12 months (launch + 12 months)","=({c}13+12*{c}16+12*{c}17*{c}22)/({c}14+12*{c}17)",CUR,"All launch and marketing spend and selling time over everyone signed in the first year. Use this one with investors","key"),
(26,"3. How long members stay",None,None,None,None),
(27,"Monthly churn","=Assumptions!${s}$23",PCT,"Assumption, not data. Replace with tracked retention after the first 90 days",""),
(28,"Expected membership life, uncapped (months)","=IF({c}27<=0,\"n/a\",1/{c}27)",NUM,"1 divided by churn. Reference only",""),
(29,"Expected paying months within the first 12","=IF({c}27<=0,12,(1-(1-{c}27)^12)/{c}27)",NUM,"",""),
(30,"Expected paying months within the first 24","=IF({c}27<=0,24,(1-(1-{c}27)^24)/{c}27)",NUM,"",""),
(31,"Paying months with a 12-month commitment, capped at 24","=12+IF({c}27<=0,12,(1-{c}27)*(1-(1-{c}27)^12)/{c}27)",NUM,"Assumes the 12 months are collected in full, then normal churn for 12 more months",""),
(33,"4. Lifetime value (conservative)",None,None,None,None),
(34,"LTV, capped at 12 months","={c}10*{c}29",CUR,"Contribution x paying months",""),
(35,"LTV, capped at 24 months","={c}10*{c}30",CUR,"","key"),
(36,"LTV, 12-month commitment, capped at 24 months","={c}10*{c}31",CUR,"Depends on enforcing the commitment",""),
(37,"LTV, uncapped (reference only)","=IF({c}27<=0,\"n/a\",{c}10/{c}27)",CUR,"Contribution / churn. Overstates value if churn is lower than real",""),
(39,"5. Ratios and payback",None,None,None,None),
(40,"LTV (24 months) / blended first-year CAC","=IF({c}24<=0,\"n/a\",{c}35/{c}24)",X,"Healthy is 3x or more. Above about 5x usually means CAC is understated or sign-ups cannot be bought with marketing. Treat as an upper bound","key"),
(41,"LTV (12 months) / blended first-year CAC","=IF({c}24<=0,\"n/a\",{c}34/{c}24)",X,"",""),
(42,"LTV (12-month commitment) / blended first-year CAC","=IF({c}24<=0,\"n/a\",{c}36/{c}24)",X,"",""),
(43,"LTV (24 months) / fully loaded CAC after opening","=IF({c}23=\"n/a\",\"n/a\",{c}35/{c}23)",X,"",""),
(44,"Months to earn back blended first-year CAC","={c}24/{c}10",NUM,"Healthy is under 12 months","key"),
(45,"Months to earn back fully loaded CAC after opening","=IF({c}23=\"n/a\",\"n/a\",{c}23/{c}10)",NUM,"",""),
(46,"Uncapped LTV / fully loaded CAC (reference only)","=IF(OR({c}37=\"n/a\",{c}23=\"n/a\"),\"n/a\",{c}37/{c}23)",X,"Not for investors: it assumes churn stays at the input forever",""),
(48,"6. Stress tests on the 24-month ratio",None,None,None,None),
(49,"LTV (24 months) / blended CAC if churn is 1.5x","={c}10*(1-(1-{c}27*1.5)^24)/({c}27*1.5)/{c}24",X,"",""),
(50,"LTV (24 months) / blended CAC if churn is 2x","={c}10*(1-(1-{c}27*2)^24)/({c}27*2)/{c}24",X,"",""),
(51,"LTV (24 months) / blended CAC if CAC is 3x","={c}35/({c}24*3)",X,"Tests what happens if real acquisition is much harder than the model's sign-up assumptions",""),
(52,"7. Premium and capacity",None,None,None,None),
(53,"Founding member contribution ($/month)","=Assumptions!$B$9*(1-Assumptions!$B$45)-Assumptions!$B$191",CUR,"$350 less card fees. The first 25 members",""),
(54,"Premium: extra revenue per premium member vs. base adult ($/month)","=Assumptions!$B$6-Assumptions!$B$5",CUR,"$800 less $399",""),
(55,"Owner hours per premium member per month","=Assumptions!$B$89",NUM,"Follow-up, retest and adjustments. The one-time assessment and routine are extra",""),
(56,"Extra revenue per hour of owner time ($)","=IF({c}55<=0,\"n/a\",{c}54/{c}55)",CUR,"What an hour of premium follow-up earns compared with the owner's hourly cost",""),
(57,"Fixed monthly cost per member at month 12","=Summary!$B$25/Summary!{s}13",CUR,"Fixed costs before contingency / members at month 12",""),
(58,"Fixed monthly cost per member at month 36","=Summary!$B$25/Summary!{s}15",CUR,"Fixed costs before contingency / members at month 36",""),
])
E.freeze_panes="B5"

# =============== BAY UTILIZATION ===============
BU=wb.create_sheet("Bay utilization")
widths(BU,{"A":64,"B":16,"C":14,"D":18,"E":16,"F":16,"G":18,"H":80})
put(BU,"A1","Bay utilization: can the bays carry the membership?",TITLE)
put(BU,"A2","Compares bay-hours members want with bay-hours available. Bookings, not the member cap, set the real limit. Every blue input is an estimate until booking data exists.",GREYI)
band(BU,4,8,"Inputs")
bi=[(5,"TrackMan bays at opening","=Assumptions!$B$118",INT,"Assumptions row 118",GREEN,None),
(6,"Bookable hours per bay per day",18,NUM,"Estimate: 5am to 11pm. Key-fob access is 24/7 but overnight sessions are rare",BLUE,YEL),
(7,"Bookable days per week",7,NUM,"",BLUE,None),
(8,"Weeks per month","=Assumptions!$B$97",NUM,"Assumptions row 97",GREEN,None),
(9,"Share of members who use the bays",0.7,PCT,"Estimate: some members use only the gym or recovery space",BLUE,YEL),
(10,"Bay sessions per bay user per month",4,NUM,"Estimate: about one a week",BLUE,YEL),
(11,"Session length (hours)",1,NUM,"",BLUE,None),
(12,"Lesson hours per member per month","=Assumptions!$B$76*Assumptions!$B$77",NUM,"Lessons also use a bay: share taking lessons x lessons a month x 1 hour",GREEN,None),
(13,"Bay hours per member per month","=B9*B10*B11+B12",NUM,"Self-practice plus lessons",BLK,None),
(14,"Peak hours per week (all bays)",32,NUM,"Estimate: weekdays 5pm to 9pm (20) plus weekend 8am to 2pm (12)",BLUE,YEL),
(15,"Share of bay demand that lands in peak hours",0.55,PCT,"Estimate: replace with booking data. The most sensitive input on this tab",BLUE,YEL),
(16,"Peak utilization ceiling before members wait",0.85,PCT,"Above this, members cannot find a time when they want one",BLUE,None)]
for r,a,v,fmt,n,fo,fi in bi:
    put(BU,f"A{r}",a); put(BU,f"B{r}",v,fo,fmt,fi); put(BU,f"H{r}",n,GREY)
band(BU,18,8,"Capacity")
bc=[(19,"Bay hours available per month","=B5*B6*B7*B8",INT,""),
(20,"Peak bay hours available per month","=B5*B14*B8",INT,"Peak hours per week x bays x weeks"),
(21,"Members the bays carry at the peak ceiling","=B20*B16/(B13*B15)",INT,"The binding limit"),
(22,"Members the bays carry across all hours at the ceiling","=B19*B16/B13",INT,"Only reachable if demand spreads evenly, which it will not"),
(23,"Member cap","=Assumptions!$B$17",INT,"Assumptions row 17"),
(24,"Members above what the bays carry at peak, at the cap","=MAX(0,B23-B21)",INT,"If above 0, the cap is not deliverable without booking rules, peak pricing or a third bay"),
(25,"Highest peak share of demand that still fits at the cap","=MIN(1,B20*B16/(B23*B13))",PCT,"Compare with the peak share in B15. Booking rules need to move demand below this")]
for r,a,f,fmt,n in bc:
    put(BU,f"A{r}",a,BOLD if r in (21,24) else BLK); put(BU,f"B{r}",f,GREEN if "Assumptions" in f else BLK,fmt); put(BU,f"H{r}",n,GREY)
band(BU,27,8,"Demand by case and month")
for j,h in enumerate(["Case","Month","Members","Bay hours demanded / month","Average utilization","Peak utilization","Status"]):
    put(BU,f"{L(1+j)}28",h,WH,fill=DARK,align=None if j==0 else "center")
put(BU,"H28","",WH,fill=DARK)
rr=29
for ci,(cn,col) in enumerate([("Conservative","B"),("Base","C"),("Upside","D")]):
    for mi,(mo,srow) in enumerate([(12,13),(24,14),(36,15)]):
        put(BU,f"A{rr}",cn); put(BU,f"B{rr}",mo,BLK,INT)
        put(BU,f"C{rr}",f"=Summary!{col}{srow}",GREEN,INT)
        put(BU,f"D{rr}",f"=C{rr}*$B$13",BLK,INT)
        put(BU,f"E{rr}",f"=D{rr}/$B$19",BLK,PCT)
        put(BU,f"F{rr}",f"=D{rr}*$B$15/$B$20",BLK,PCT)
        put(BU,f"G{rr}",f'=IF(F{rr}>1,"Over capacity",IF(F{rr}>$B$16,"Waits at peak","OK"))',BOLD)
        rr+=1
band(BU,39,8,"First month each case passes what the bays carry at peak")
for i,n in enumerate(["Conservative","Base","Upside"]):
    r=40+i
    put(BU,f"A{r}",n)
    put(BU,f"B{r}",f'=IF(COUNTIF({n}!$B$7:$BI$7,">"&$B$21)=0,"Not within 60 months",61-COUNTIF({n}!$B$7:$BI$7,">"&$B$21))',BLK,INT)
put(BU,"H40","Assumes members rise through the ramp, which they do in all three cases",GREY)
put(BU,"A44",'="Two bays carry about "&TEXT(B21,"0")&" members at peak. The cap is "&B23&"."',BOLD)
put(BU,"A46","Ways to add capacity",BOLD)
for i,t in enumerate(["Peak pricing or a peak-hour booking limit per member (for example 2 peak sessions a month).","Booking windows that reserve off-peak hours for lessons and juniors.","Incentives to shift sessions to mornings and midday (lower price for off-peak use).","A third bay, or a second bay pair in a later phase. Test this on Assumptions row 118 and row 111.","A lower advertised cap until booking data shows where demand lands."]):
    put(BU,f"A{47+i}","• "+t)


# =============== SENSITIVITY (static) ===============
S2=wb.create_sheet("Sensitivity")
widths(S2,{"A":56,"B":16,"C":16,"D":14,"E":16,"F":14,"G":16,"H":14,"I":16,"J":16,"K":58})
put(S2,"A1","Sensitivity: what happens when an assumption is wrong",TITLE)
put(S2,"A2","SNAPSHOT. These values were produced by changing the inputs in column K on a copy of this model and recalculating. They do not update when you change inputs. Rerun a case by making the same change on Assumptions and reading the Summary.",F(size=9,italic=True,color="FFC00000"))
S2["A2"].alignment=Alignment(wrap_text=True,vertical="top"); S2.row_dimensions[2].height=28
band(S2,4,11)
hd=["Case","Base: capital required","Base: surplus vs. raise","Base: first profitable month","Base: Year 3 profit","Base: members at month 36","Conservative: capital required","Conservative: first profitable month","Conservative: Year 3 profit","Base: years for investor to recoup (straight equity)","What was changed"]
for j,h in enumerate(hd):
    c=S2.cell(5,1+j,h); c.font=F(bold=True,color="FFFFFFFF"); c.fill=DARK; c.alignment=Alignment(wrap_text=True,vertical="center",horizontal="center" if j else "left")
S2.row_dimensions[5].height=58
chg={"As built":"No change",
"Sign-ups 25% slower":"Assumptions B22, C22, D22 multiplied by 0.75",
"Monthly churn doubled":"Assumptions B23, C23, D23 multiplied by 2",
"All-in price $450 (sign-ups 12% faster)":"B5 and B6 set to 450, founding price 399, sign-ups x1.125",
"All-in price $600 (sign-ups 19% slower)":"B5 and B6 set to 600, founding price 529, sign-ups x0.8125",
"All dues 10% lower":"B5, B6, B7, B9 and B124 multiplied by 0.9",
"Paid add-on take-up halved (S&C, lessons, junior packages)":"B70, B76, B94, B95, B96 multiplied by 0.5",
"Rent 50% higher ($18/sq ft)":"B27 set to 18",
"Build-out cost doubled ($40/sq ft)":"B29 set to 40",
"Insurance doubled ($1,600/month)":"B40 set to 1,600",
"Add hourly floor staff (one 7-hour shift, 7 days)":"B157 (staffed days per week) set to 7",
"No part-time hire (owner covers the floor for all 60 months)":"B161 set to 0",
"Part-time hire works 30 hours a week":"B162 set to 30",
"Tiered pricing on ($400 / $450 / $500, slower sign-ups at each step)":"B175 set to 1",
"12-month commitments: renewal rates 60% / 70% / 80%, no early churn":"B186 set to 2 (renewal rates in B24:D24)",
"Owner delivers the follow-ups (no strength coach pay)":"B189 set to 0",
"Rent starts at lease signing (6 months paid before opening)":"B194 set to 6",
"Landlord allowance of $20K toward the build-out":"B195 set to 20,000",
"Larger space: 4,800 sq ft (rent and build-out scale with size)":"B26 (facility size) set to 4,800",
"Winston trains 10 non-member athletes a month":"B198 set to 10",
"GM salary $60K instead of $45K":"B33 set to 60,000",
"Seasonal sign-ups (winter 0.7, spring 1.3 to 1.4)":"B137 to B148 set to 0.7, 0.7, 1.3, 1.4, 1.4, 1.1, 1, 1, 1, 1, 0.7, 0.7 (illustrative)",
"Winston also needs a $40K base salary":"B34 set to 40,000",
"Stress: sign-ups 25% slower, churn 1.5x":"B22:D22 x 0.75, B23:D23 x 1.5"}
res=json.load(open("sens_results.json"))
for i,d in enumerate(res):
    r=6+i
    vals=[d["name"],d["base_cap"],d["base_surplus"],d["base_first"],d["base_y3"],d["base_m36"],d["cons_cap"],d["cons_first"],d["cons_y3"],d["base_pay"],chg[d["name"]]]
    fmts=[None,CUR,CUR,INT,CUR,NUM,CUR,INT,CUR,NUM,None]
    for j,(v,fm) in enumerate(zip(vals,fmts)):
        c=S2.cell(r,1+j,v); c.font=BOLD if i==0 else BLK
        if fm: c.number_format=fm
        if j in (3,7,9): c.alignment=Alignment(horizontal="right")
    S2.cell(r,11).font=GREY
    S2.cell(r,11).alignment=Alignment(wrap_text=True,vertical="top"); S2.row_dimensions[r].height=est_h(vals[10],56)
n=6+len(res)+1
put(S2,f"A{n}","How to read it",BOLD)
notes=["Churn is the most dangerous assumption: doubling it cuts Base Year 3 profit by more than half and leaves the Conservative case unprofitable within 36 months.",
"A base salary for Winston is a further risk. His no-base deal depends on him accepting the ramp risk, so confirm it before relying on it. The GM is modeled with a $45K base.",
"Floor coverage is the largest lever: adding hourly staff back (12 hours a day, six days a week) raises Conservative capital by about $209K and cuts Base Year 3 profit by about $72K. Phase coverage with membership.",
"Build-out cost swings the capital needed by about $100K. Get the contractor quote first.",
"Premium, dues and add-on take-up each move profit by $40K to $75K a year, but none of them break the case alone. They matter most together, and in the Conservative case.",
"'Not within 36 mo' means the monthly P&L never turns positive in the model's three years."]
for i,t in enumerate(notes):
    put(S2,f"A{n+1+i}",t,BLK,wrap=True); S2.merge_cells(f"A{n+1+i}:K{n+1+i}"); S2.row_dimensions[n+1+i].height=18
S2.freeze_panes="B6"

# =============== RISKS AND CHECKLIST ===============
K=wb.create_sheet("Risks and checklist")
widths(K,{"A":38,"B":52,"C":58,"D":40})
put(K,"A1","Risks, legal checklist, market and data room",TITLE)
def table(ws,start,heads,rows,fills=None):
    band(ws,start,len(heads))
    for j,h in enumerate(heads):
        c=ws.cell(start,1+j,h); c.font=F(bold=True,color="FFFFFFFF"); c.fill=DARK
    for i,row in enumerate(rows):
        r=start+1+i; h=13.5
        for j,v in enumerate(row):
            c=ws.cell(r,1+j,v); c.font=BOLD if j==0 else BLK; c.alignment=Alignment(wrap_text=True,vertical="top")
            h=max(h,est_h(v,ws.column_dimensions[L(1+j)].width))
        ws.row_dimensions[r].height=h
    return start+len(rows)+2
risks=[("Slow ramp","Sign-ups below plan stretch the losses. Churn is the most sensitive input in the model.","Presell the 25 founding slots with signed commitments (no deposits taken for now). Collect letters of intent. Fund in stages and size the raise to the Conservative case.","Assumptions rows 21 to 23"),
("Key person: you","You are the GM, the clinician and the TPI provider, on site about 35 hours a week. If you are unavailable, the business has no operator.","Written procedures, a trained backup, key-person insurance, an advisor. Say how premium service hours will be covered if you are ill.","Assumptions rows 89 to 92, 129 to 132"),
("24/7 unstaffed access","Injury, theft, damage or misuse with nobody present. Insurers may price it high or decline.","Entry control, cameras, waivers, an emergency procedure and call point, an AED, clear rules, and an insurance quote that names 24/7 access.","Assumptions B40, B121, B122"),
("Minors on site","Supervision, background checks, SafeSport and insurance for minors. Unsupervised minors can void cover.","Parent-signed waivers, limited-time windows aligned to staffed hours, ratios, background checks, insurer confirmation in writing.","Assumptions rows 7, 12, 93 to 108"),
("1099 classification","Winston and the pro may be treated as employees if the facility controls hours, methods and equipment.","Written agreements, their own schedule, pricing and clients, other income sources, and a CPA or attorney review before offers.","Assumptions B33, B34, B73, B79"),
("Contractor volume and price","The model assumes 35% to 40% of adults buy $150 to $200 sessions on top of dues. Unproven.","Ask Winston and the pro what their current clients spend. Pilot at a partner site before opening.","Assumptions rows 70 to 79"),
("Pricing against competitors","$399 to $800 is above TeeBox's reported $200 to $399.","Presale evidence at your price. Show the clinical and S&C access they do not offer.","Assumptions B5, B6"),
("Build-out and equipment cost","The biggest swing in capital. Quotes are not in yet.","Contractor quote, landlord tenant-improvement allowance, 10% contingency (already in).","Assumptions B29, Startup"),
("Seasonality","Golf slows in winter. Indoor helps, but sign-ups may follow the weather.","Open in the shoulder season, lean on juniors and college teams, show a seasonality curve in the deck.","Assumptions rows 136 to 149"),
("Lease and location","Downtown rent is likely above the industrial-flex rate used.","Get real quotes; negotiate exit and expansion options.","Assumptions B26, B27"),
("Equipment replacement","No replacement reserve or depreciation is in the model.","Add a monthly reserve once equipment quotes are final.","Not modeled"),
("Securities compliance","Raising money from investors is a regulated offering. Public promotion of the raise can cause problems.","Securities counsel before outreach; keep the reel series separate from the raise; verify accredited investors.","n/a")]
nxt=table(K,3,["Risk","Why it matters","What reduces it","Model input to test"],risks)
legal=[("Entity and name","Parent ATPRX, LLC with the facility as a brand or subsidiary; DBA for 'Driven Golf Performance' (check TN and GA filings and handles first)","Not started"),
("Operating agreement","Roles, your 75% to 80% control, investor rights, distributions and exit","Not started"),
("Securities counsel","Exemption choice, offering documents, accredited-investor verification, investor reporting","Not started"),
("Contractor agreements","Winston and the pro: scope, pay split, autonomy, non-solicit, insurance, IP","Not started"),
("Lease review","Term, exit, expansion rights for bay 2, tenant-improvement allowance, use clause for 24/7 access","Not started"),
("Insurance","General and professional liability, 24/7 access, minors, equipment, key-person","Not started"),
("Waivers and policies","Adult and minor waivers, emergency plan, facility rules, photo and data consent","Not started"),
("Scope of practice","Confirm what TPI screening, retests and routines you may deliver for a fee in TN and GA","Not started"),
("IP","Protect your testing battery and programme documents before sharing outside","Not started"),
("Payments and billing","Membership billing, refundable deposits handling, sales tax","Not started")]
nxt=table(K,nxt,["Legal and structure item","What to settle","Status (edit)"],[(a,b,c) for a,b,c in legal])
for r in range(nxt-len(legal)-1,nxt-1):
    K.cell(r,3).font=BLUE
comp=[("TeeBox (franchise)","Indoor golf simulator membership","Reported: about $200/month basic, $249 to $399 performance; founding rate near $249. FDD figures reported: $49,500 franchise fee, $496K to $798K investment, 8% royalty","Source was an AI search relayed by your investor friend, and he doubted the revenue figure. Pull the actual FDD before using any number."),
("D-BAT (franchise)","Indoor baseball and softball, a different sport","Pull the FDD (Item 19)","Useful comparison for revenue per facility and mix"),
("Wampler Sports Performance","Ooltewah, TPI-certified, one-person operation","Not checked","Competitor and possible referral partner. Visit in person."),
("Dalton Golf Lab","Fittings and precision equipment","Not checked","Possible partner. Your plan to partner with them is not yet shared with the owner."),
("Strike Zone Chattanooga","Local indoor facility","Little found","Research pricing and membership")]
nxt=table(K,nxt,["Competitor or comparison","What it is","Price and terms","What to check"],comp)
band(K,nxt,4,"Market size: fill these in (blue cells)")
msz=[("Competitive golfers aged 35 to 55 within 20 minutes of the site","Source: club memberships, state golf association, rounds data"),
("Private golf clubs within 20 minutes","Source: your own list"),
("High-school golf programs nearby","Source: school lists"),
("College golf programs nearby","Source: school lists"),
("Junior golfers aged 13 to 18 in those programs","Source: coaches")]
r=nxt+1
put(K,f"A{r}","Item",BOLD); put(K,f"B{r}","Number",BOLD); put(K,f"C{r}","Source or note",BOLD)
for i,(a,b) in enumerate(msz):
    rr=r+1+i; put(K,f"A{rr}",a,BLK,wrap=True); put(K,f"B{rr}",None,BLUE,INT,YEL); put(K,f"C{rr}",b,GREY,wrap=True); K.row_dimensions[rr].height=est_h(a,38)
rr=r+1+len(msz)
put(K,f"A{rr}","Members at the 200 cap as a share of the golfer market",BOLD,wrap=True); put(K,f"B{rr}",f'=IF(B{r+1}="","",200/B{r+1})',BLK,PCT); put(K,f"C{rr}","Fills in once you enter the first number above. A small share is a stronger story.",GREY)
nxt=rr+2
band(K,nxt,4,"Proof of demand: targets and progress (edit)")
dem=[("Serious golfers interviewed about $399 and $800",12),("Founding-member commitments (signed, no money taken)",40),("Letters of intent from clubs or programs",3)]
put(K,f"A{nxt+1}","Item",BOLD); put(K,f"B{nxt+1}","Target",BOLD); put(K,f"C{nxt+1}","Done so far",BOLD); put(K,f"D{nxt+1}","Progress",BOLD)
for i,(a,t) in enumerate(dem):
    rr=nxt+2+i; put(K,f"A{rr}",a,BLK,wrap=True); put(K,f"B{rr}",t,BLUE,INT); put(K,f"C{rr}",(0 if i==2 else ("='Founding members'!B10" if i==0 else "='Founding members'!B12")),(BLUE if i==2 else GREEN),INT,(YEL if i==2 else None)); put(K,f"D{rr}",f"=IF(B{rr}=0,\"\",C{rr}/B{rr})",BLK,PCT); K.row_dimensions[rr].height=est_h(a,38)
nxt=nxt+2+len(dem)+1
dr=[("Pitch deck","10 to 12 slides: problem, offer, market, model, returns, team, risks, ask"),("This model","With the Guide, Summary, Investor returns, Sources and uses and Sensitivity tabs"),("Equipment and build-out quotes","Replaces the placeholders"),("Lease letter of intent","Term, rate, expansion rights"),("Entity documents and operating agreement",""),("Contractor agreements","Winston and the pro"),("Insurance quotes","24/7 access and minors named"),("Founding-member evidence","Deposits, interview notes, letters of intent"),("Bios and references","You, Winston, the pro, advisors")]
nxt=table(K,nxt,["Data-room item","Detail"],dr)

# =============== GUIDE ===============
G=wb.create_sheet("Guide")
widths(G,{"A":30,"B":78,"C":62})
put(G,"A1","Driven Golf Performance: how to read this model and learn it",TITLE)
put(G,"A2","Working name. Planning estimates only: not a promised return. Have an attorney review before any investor sees projections.",GREYI)
def sect(r,t): band(G,r,3,t); return r+1
r=4
r=sect(r,"What this workbook does")
G[f"A{r}"]="It estimates what the facility costs to open, how members build over five years, when it turns profitable, what it pays an investor, and what could go wrong. It is built so you can change any input and see the effect everywhere."
G.merge_cells(f"A{r}:C{r}"); G[f"A{r}"].font=BLK; G[f"A{r}"].alignment=Alignment(wrap_text=True,vertical="top"); G.row_dimensions[r].height=30; r+=2
r=sect(r,"Colour and cell rules")
for a,b in [("Blue text","An input you can change."),("Black text","A formula. Do not type over it."),("Green text","A number pulled from another tab."),("Yellow fill","A key assumption that moves the results a lot."),("Column E notes on Assumptions","Each line says 'You said' (your figure) or 'My assumption' (a placeholder to replace).")]:
    put(G,f"A{r}",a,BOLD); put(G,f"B{r}",b,BLK,wrap=True); r+=1
r+=1
r=sect(r,"Where the numbers flow")
flow=["Assumptions is the control panel. Almost every input lives there.","Startup adds up the one-time costs. Base, Conservative and Upside each run 60 months. The three cases differ only in members at opening, sign-ups per month and churn.","Summary, Ownership, Investor returns, Sources and uses, Unit economics and Dues only all read from those sheets.","Sensitivity is a pasted snapshot. It does not update."]
for t in flow:
    put(G,f"A{r}","•",BLK,align="right"); put(G,f"B{r}",t,BLK,wrap=True); G.row_dimensions[r].height=est_h(t,78); r+=1
r+=1
r=sect(r,"Tab by tab")
put(G,f"A{r}","Tab",BOLD); put(G,f"B{r}","What it answers",BOLD); put(G,f"C{r}","Where to look first",BOLD); r+=1
tabs=[("Summary","How much capital, when profit starts, and how big the profit is, for three cases","Rows 6 to 11 (capital), 16 (first profitable month), 21 to 23 (profit)"),
("Investor returns","What an investor gets back under three structures, with IRR and multiple of capital","Summary table at the top"),
("Sources and uses","Where the raise goes, and the staged release of funds","The check lines should read 0"),
("Unit economics","What a member is worth and costs to win, using full acquisition cost and 12 and 24 month value","Rows 24 and 35, then the ratios in rows 40 to 44"),
("Bay utilization","Whether two bays can carry the membership at peak times","Row 21 (members the bays carry) against the cap in row 23"),
("Sensitivity","Which wrong assumptions hurt most","Compare each row to 'As built'"),
("Dues only","Can membership dues alone cover costs, at different premium shares?","The table from row 13"),
("Assumptions","Every input, with its source","Column E notes; the yellow cells"),
("Startup","The one-time cost build-up and total capital required","Rows 34 and 41"),
("Ownership","The ask: valuation, payback, owner pay","Rows 8 to 26"),
("Ownership scenarios","Owner pay at 20%, 30% and 40% starting stakes, a yearly buyout schedule, and the investor result at each stake","Year 3 pay table, then the buyout schedule"),
("Base / Conservative / Upside","Month-by-month members, revenue, costs and profit for 60 months","Column B (month 1) read downward"),
("Risks and checklist","Risks, legal list, competitors, market-size blanks, data-room list","Risk table, then the legal list"),
("Equipment list","Itemized gym floor and storage with a quote column and a buy-now or defer switch","Column E (quotes) and the totals under the list"),
("Founding members","Tracker for interviews and commitments, with the funnel and which opening case the signed letters support","The funnel block at the top"),
("Scripts","The interview script and a non-binding founding-member commitment letter","Read before the first conversation"),
("Changes and flags","Everything changed since your first file, and what to confirm","Read top to bottom once")]
for a,b,c in tabs:
    put(G,f"A{r}",a,BOLD,wrap=True); put(G,f"B{r}",b,BLK,wrap=True); put(G,f"C{r}",c,GREY,wrap=True); G.row_dimensions[r].height=max(est_h(b,78),est_h(c,62)); r+=1
r+=1
r=sect(r,"Numbers to know (live)")
for a,f,fm in [("Raise","=Assumptions!B54",CUR),("Total start-up cost","=Summary!C6",CUR),("Capital required: Conservative","=Summary!B9",CUR),("Capital required: Base","=Summary!C9",CUR),("Break-even members","=Summary!B27",NUM),("Base: first profitable month","=Summary!C16",INT),("Base: Year 3 operating profit","=Summary!C23",CUR),("Post-money valuation at the ask","=Ownership!C8",CUR)]:
    put(G,f"A{r}",a,BLK); put(G,f"B{r}",f,GREEN,fm); G[f"B{r}"].alignment=Alignment(horizontal="left"); r+=1
r+=1
r=sect(r,"A path to learn the model (about 2 to 3 hours)")
put(G,f"A{r}","Step",BOLD); put(G,f"B{r}","What to do",BOLD); put(G,f"C{r}","Try this",BOLD); r+=1
steps=[("1. Read the Summary","Read the Summary tab top to bottom. Note how much capital each case needs and what the profit looks like in Years 1 to 3.","Which case needs the most capital? Is it more than the raise?"),
("2. Sort the assumptions","On Assumptions read column E. List every 'My assumption' you can replace with a quote or a conversation.","Replace the biggest first: build-out cost (B29), rent (B27), add-on take-up (rows 70 and 76), sign-ups (row 22), churn (row 23)."),
("3. Follow the start-up cost","On Startup, rows 6 to 21 are equipment and build-out, 23 is contingency, 25 to 33 are opening costs and build pay. Rows 38 to 44 add ramp losses and a cushion.","Change build-out cost per sq ft (Assumptions B29) from 20 to 40 and watch Summary row 9."),
("4. Trace one month by hand","On Base, read column B downward: members (row 7), revenue (rows 12 to 15), costs (rows 18 to 28), profit (row 30). Redo month 1 on a calculator.","Then read column N (month 13) and see how sign-ups and churn change members."),
("5. Compare the three cases","Conservative, Base and Upside are identical except for Assumptions rows 21 to 23.","Change one input in one case and watch only that case move."),
("6. Understand break-even","Summary rows 24 to 29 and the Dues only tab show how many members cover costs, with and without add-ons.","Change the premium share (Assumptions B11, B125). Notice it moves profit more than break-even."),
("7. Learn the investor math","On Ownership and Investor returns compare the three structures. IRR is the yearly return; the multiple of capital is cash back divided by cash in.","Change equity offered (B55) or the priority multiple (B58) and see which structure favours the investor."),
("8. See where the money goes","Sources and uses lists every dollar and the staged release of funds.","Check that the check lines read 0 and read the shortfall line."),
("9. Learn the unit economics","Unit economics shows what one member is worth, how long they stay and what they cost to win.","Halve churn (Assumptions row 23) and watch lifetime value double."),
("10. Test what breaks it","Read the Sensitivity snapshot. To run your own case, change the input on Assumptions and read the Summary.","Pick the three cases that hurt most and write what you would do about each."),
("11. Close the open items","Read Risks and checklist and Changes and flags. Work the legal list and the proof-of-demand targets.","Do not show an investor until quotes, counsel and demand evidence are in.")]
for a,b,c in steps:
    put(G,f"A{r}",a,BOLD,wrap=True); put(G,f"B{r}",b,BLK,wrap=True); put(G,f"C{r}",c,GREY,wrap=True); G.row_dimensions[r].height=max(est_h(b,78),est_h(c,62)); r+=1
r+=1
r=sect(r,"Glossary")
gl=[("Capital required","Start-up cost, plus the deepest cumulative loss before profit, plus a safety cushion."),("Operating profit","Revenue minus operating costs, including your salary. Before depreciation, debt and taxes."),("Break-even members","Members needed for monthly revenue to cover monthly costs."),("Churn","Share of members who leave each month. 3% means an average stay of about 33 months."),("Distributable cash","The share of profit paid out. The rest stays in the business."),("Post-money valuation","Raise divided by the equity offered: $500K for 30% implies about $1.67M."),("Priority payback","The investor takes a larger share of cash until a set multiple of their money is back, then drops to their equity share."),("Multiple of capital (MOIC)","Total cash back divided by cash invested."),("IRR","The yearly return that makes the cash flows net to zero. Higher is better and it rewards cash back sooner."),("LTV and CAC","Lifetime contribution of one member, and the cost to win one."),("Contribution","Revenue per member less card fees, before fixed costs."),("1099 contractor","Sets their own schedule and pays their own taxes. Not an employee.")]
for a,b in gl:
    put(G,f"A{r}",a,BOLD,wrap=True); put(G,f"B{r}",b,BLK,wrap=True); G.row_dimensions[r].height=est_h(b,78); r+=1
r+=1
r=sect(r,"What the model leaves out")
lo=["Equipment replacement, depreciation, debt service and taxes.","Your own TPI follow-up and clinical revenue through ATPRX, LLC.","Lost sign-ups while membership is full, and any waitlist.","Seasonality is built in but switched off (factors at 1.0).","Damage or theft cost in a 24/7 unstaffed building beyond the insurance line.","Paid services bought by non-members.","Founding-member deposits as usable cash before opening.","Personal taxes on distributions."]
for t in lo:
    put(G,f"A{r}","•",BLK,align="right"); put(G,f"B{r}",t,BLK,wrap=True); r+=1
G.sheet_view.showGridLines=False

# ---- Changes sheet addition
N=wb["Changes and flags"]
rr=N.max_row+2
N[f"A{rr}"]="New tabs added"; N[f"A{rr}"].font=BOLD
for t in ["Guide, Investor returns, Sources and uses, Unit economics, Sensitivity and Risks and checklist are new. The monthly sheets now run 60 months. Seasonality factors (Assumptions rows 135 to 149) are added and switched off.",
 "Sensitivity is a static snapshot of 24 cases and does not update. Everything else is live."]:
    rr+=1; N[f"A{rr}"]=t; N.merge_cells(f"A{rr}:D{rr}"); N[f"A{rr}"].font=BLK; N[f"A{rr}"].alignment=Alignment(wrap_text=True,vertical="top"); N.row_dimensions[rr].height=34

# =============== FOUNDING MEMBERS TRACKER ===============
FM=wb.create_sheet("Founding members")
from openpyxl.worksheet.datavalidation import DataValidation
widths(FM,{"A":34,"B":18,"C":13,"D":16,"E":16,"F":18,"G":50,"H":14,"I":14})
put(FM,"A1","Founding members: interviews and commitments",TITLE)
put(FM,"A2","Log every conversation in the table below. Circles: 1 = people already treated or known, 2 = golf community, 3 = introductions, 4 = cold local. Commitments are non-binding and no money is taken.",GREYI)
band(FM,4,9,"Funnel")
fm=[(5,"Target signed commitments",40,INT,BLUE,"Owner plan: 30 to 50 signed commitments"),
(6,"Interview-to-commitment conversion",0.33,PCT,BLUE,"Estimate: about one in three interviews"),
(7,"Commitment-to-real-member conversion",0.67,PCT,BLUE,"Estimate: about two in three signed commitments become members"),
(8,"Conversations needed to reach the target","=B5/B6",INT,BLK,"Formula"),
(9,"Expected real members from the target","=B5*B7",INT,BLK,"Formula"),
(10,"Interviews logged","=COUNTA(A17:A216)",INT,BLK,""),
(11,"Committed or signed","=COUNTIF(H17:H216,\"Committed\")+COUNTIF(H17:H216,\"Signed\")",INT,BLK,""),
(12,"Signed commitment letters","=COUNTIF(H17:H216,\"Signed\")",INT,BLK,"Feeds Risks and checklist"),
(13,"Expected real members from signed letters","=B12*B7",NUM,BLK,""),
(14,"Opening-member case the signed letters support","=IF(B13>=Assumptions!D21,\"Upside\",IF(B13>=Assumptions!C21,\"Base\",IF(B13>=Assumptions!B21,\"Conservative\",\"Below Conservative\")))",None,BOLD,"Compares expected real members with opening members in Assumptions B21:D21")]
for r,lab,v,fmt,fo,nt in fm:
    put(FM,f"A{r}",lab); put(FM,f"B{r}",v,fo,fmt,YEL if fo is BLUE else None); put(FM,f"C{r}",nt,GREY)
for j,h in enumerate(["Name","Source (circle)","Date","Tier they would take","Price accepted ($/mo)","Would take premium at $800?","Objections, in their words","Status","Follow-up date"]):
    put(FM,f"{L(1+j)}16",h,F(bold=True,color="FFFFFFFF"),fill=DARK,wrap=True)
FM.row_dimensions[16].height=30
for r in range(17,217):
    for c,fmt in ((3,"mm/dd/yyyy"),(5,CUR),(9,"mm/dd/yyyy")):
        FM.cell(r,c).number_format=fmt
    for c in range(1,10): FM.cell(r,c).font=BLUE
def dvl(col,lst):
    d=DataValidation(type="list",formula1='"'+lst+'"',allow_blank=True); FM.add_data_validation(d); d.add(f"{col}17:{col}216")
dvl("B","Circle 1,Circle 2,Circle 3,Circle 4"); dvl("D","Founding,Member,Premium,Junior"); dvl("F","Yes,No,Maybe"); dvl("H","Interviewed,Committed,Signed,Declined")
FM.freeze_panes="A17"

# =============== SCRIPTS ===============
SC=wb.create_sheet("Scripts")
widths(SC,{"A":120})
put(SC,"A1","Scripts: founding-member interview and commitment letter",TITLE)
lines=[("Interview script (about 15 minutes)",True),
("1. Open with curiosity: \"I'm planning a golf performance facility downtown and I'd love your honest reaction.\"",False),
("2. Ask about them first: how often they play, what they spend now on lessons, fitting or training, what frustrates them, and any aches or limits that hold them back.",False),
("3. Describe the offer in two sentences: a members-only golf performance facility with two TrackMan bays, a performance gym and a clinical room, open 24/7 to members. Founding members pay $350 a month with a 12-month commitment, regular membership starts at $399, and premium (a TPI assessment and monthly follow-up) is $800.",False),
("4. Name the price, then stop talking. The silence after the price is where you learn the most.",False),
("5. Ask three things: Would you join at $350 or $400 a month? Would the $800 premium tier be worth it, and what would it need to include? What would stop you from joining?",False),
("6. Write objections down word for word in the tracker.",False),
("7. Soft ask: \"If this opens, would you sign a non-binding founding-member commitment?\" If yes, send the letter below the same day.",False),
("8. Follow up within a week. Ask for two referrals: who else plays seriously?",False),
("Rules: take no money. Make no promises about results or injury prevention. Do not mention the investor raise to members.",False),
("",False),
("Founding-member commitment letter (non-binding). Have an attorney review it before use, including Tennessee and Georgia consumer and auto-renewal rules.",True),
("Eighteen Athletic Club: Founding Member Letter of Intent",False),
("Date: [____]    Name: [____]",False),
("I am interested in becoming a founding member of Eighteen Athletic Club, planned for [location to be confirmed] with a target opening of [late 2027].",False),
("Founding membership as currently planned: [$350] per month, a 12-month commitment from the opening date, 24/7 access, the gym floor and both TrackMan bays. [Premium add-on selected: yes / no.]",False),
("This letter is an expression of interest only. It is not a contract. I am not required to join, and no payment is requested or taken now. Eighteen Athletic Club is not obligated to open or to offer membership on these terms, and terms may change before opening. If the facility opens, I will be offered a membership agreement, which will be the only binding document.",False),
("Signature: [____]    Phone: [____]    Email: [____]",False)]
for i,(tx,hd) in enumerate(lines):
    r=3+i; put(SC,f"A{r}",tx,BOLD if hd else BLK,wrap=True)
    SC.row_dimensions[r].height=est_h(tx,118)

# ---- order
order=["Guide","Summary","Investor returns","Sources and uses","Unit economics","Bay utilization","Sensitivity","Dues only","Assumptions","Startup","Equipment list","Ownership","Base","Conservative","Upside","Risks and checklist","Founding members","Scripts","Changes and flags"]
wb._sheets=[wb[n] for n in order]
wb.active=0
wb.save("stage3.xlsx")
