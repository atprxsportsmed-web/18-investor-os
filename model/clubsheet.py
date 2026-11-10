import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter as L
wb=openpyxl.load_workbook("stage3b.xlsx")
if "Club scenario" in wb.sheetnames: del wb["Club scenario"]
ws=wb.create_sheet("Club scenario")
BLUE=Font(color="FF0000FF"); GRN=Font(color="FF008000"); BOLD=Font(bold=True)
HDR=PatternFill("solid",fgColor="FF1F3A4D"); HF=Font(bold=True,color="FFFFFFFF"); YEL=PatternFill("solid",fgColor="FFFFF2CC")
W=Alignment(wrap_text=True,vertical="top")
ws.column_dimensions["A"].width=62
for c in "BCDEFGH": ws.column_dimensions[c].width=17
def hdr(r,labels):
    for j,h in enumerate(labels):
        c=ws.cell(r,1+j,h); c.font=HF; c.fill=HDR; c.alignment=Alignment(wrap_text=True,vertical="center")
ws["A1"]="Club model: what changed from the earlier plan, and what the plan costs and earns now"; ws["A1"].font=Font(bold=True,size=14)
ws["A2"]=("The model now describes a coached performance club: $500 all in for every member (founders $449), a cap of 100, and onboarding in every membership (interview, testing, a Tonal routine, a retest every 3 months). Golf tech is five hitting stations ($82,197.10 in total) and no TrackMan bay (none planned). "
 "Extra one-on-one training, rehab and PGA instruction are bought per session or in packages. The left column below is the earlier plan, kept as a static snapshot; the right columns are live.")
ws["A2"].alignment=W; ws.merge_cells("A2:H2"); ws.row_dimensions[2].height=62
r=4; hdr(r,["Comparison","Earlier plan (static)","Club model: Conservative","Club model: Base","Club model: Upside"]); ws.row_dimensions[r].height=34; r+=1
earlier={"start":266120,"cap":(406701,347424,329373),"rev":(283106,520985,723600),"prof":(36456,245318,439624),"first":(20,7,2),"m36":(54,98,125)}
rows=[("Total start-up cost","start","=Summary!{c}6","$#,##0"),
("Capital required","cap","=Summary!{c}9","$#,##0"),
("Members at month 36","m36","=Summary!{c}15","0"),
("First profitable month","first","=Summary!{c}16","0"),
("Year 3 revenue","rev","=Summary!{c}20","$#,##0"),
("Year 3 operating profit","prof","=Summary!{c}23","$#,##0"),
("Year 3 revenue per sq ft",None,"=Summary!{c}20/Assumptions!$B$26","$#,##0")]
for lab,k,fm,nf in rows:
    ws.cell(r,1,lab)
    if k=="start": ws.cell(r,2,earlier[k]).font=BLUE; ws.cell(r,2).number_format=nf
    elif k: ws.cell(r,2,earlier[k][1]).font=BLUE; ws.cell(r,2).number_format=nf
    else: ws.cell(r,2,earlier["rev"][1]/3000).font=BLUE; ws.cell(r,2).number_format=nf
    for j,c in enumerate("BCD"):
        x=ws.cell(r,3+j,fm.format(c=c)); x.font=GRN; x.number_format=nf
    r+=1
ws.cell(r,1,"Earlier plan column shows its Base case: $399 dues (founders $350) plus an optional performance block, cap 125, two TrackMan bays at $10,000 each, five Rapsodo stations at $12,500 total. Break-even there was about 45 members.").alignment=W
ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8); ws.row_dimensions[r].height=32; r+=2
ws.cell(r,1,"Read this before relying on it").font=BOLD; r+=1
notes=["Price is the big unknown. Sign-ups per month are scaled for the $500 price (x0.80 of the earlier pace) and that is my guess, not data. The Sensitivity tab shows $450 and $600. Only your founding-member conversations can test it.",
"Coach and owner time. Every member is onboarded and retested every 3 months. The strength coach is paid about $22.50 per member per month for follow-up, and the model carries 3 hours of your time per new member (interview, testing, routine build). Onboarding the first 25 to 30 members before opening is about 90 hours and is not scheduled in the model.",
"Gym is a $45,000 placeholder: 3 Tonals ($15,000), the functional rig with storage ($15,000), plus placeholders for cardio, free weights and a velocity device.",
"Add-on revenue (one-on-one training, rehab, PGA instruction) uses the earlier take-up estimates: 35% of members train with the strength coach and 40% take lessons. Rehab billed by you is not in the model; it may belong in ATPRX, LLC, not the facility.",
"The $82,197.10 for the five stations and the $15,000 clinical budget are your figures. Station software subscriptions are carried at +$300 a month as a placeholder.",
"Capacity: the Station capacity tab checks that five stations can carry the cap at peak. The layout has not been drawn; ceiling height and net clearance are unverified."]
for t in notes:
    c=ws.cell(r,1,"• "+t); c.alignment=W; ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8); ws.row_dimensions[r].height=48; r+=1
ws.freeze_panes="A4"
names=wb.sheetnames; names.remove("Club scenario"); names.insert(names.index("Space plan")+1,"Club scenario")
wb._sheets=[wb[n] for n in names]
wb.save("stage3b.xlsx")
