import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment
wb=openpyxl.load_workbook("stage3b.xlsx")
if "Space plan" in wb.sheetnames: del wb["Space plan"]
ws=wb.create_sheet("Space plan")
BLUE=Font(color="FF0000FF"); BLK=Font(color="FF000000"); GRN=Font(color="FF008000"); BOLD=Font(bold=True)
HDR=PatternFill("solid",fgColor="FF1F3A4D"); HF=Font(bold=True,color="FFFFFFFF")
ws.column_dimensions["A"].width=58
for c in "BCD": ws.column_dimensions[c].width=16
ws.column_dimensions["E"].width=70
ws["A1"]="Space plan: how 3,000 square feet is used, and what each square foot earns"; ws["A1"].font=Font(bold=True,size=14)
ws["A2"]="Square footage is the rent and build-out driver, so layout is a financial decision. The zone sizes below are estimates to test with the contractor and the net requirements; the model does not check that the layout fits."
ws["A2"].alignment=Alignment(wrap_text=True,vertical="top"); ws.merge_cells("A2:E2"); ws.row_dimensions[2].height=34
for j,h in enumerate(["Zone","Estimate, sq ft","% of space","","Notes"]):
    c=ws.cell(4,1+j,h); c.font=HF; c.fill=HDR
zones=[("Hitting stations (5, in a row facing a net wall)",900,"Estimate: about 9 ft wide by 20 ft deep each, with the net-return wall and logo wall behind. Ceiling height and net clearance are the first things to confirm"),
("Gym floor: 3 digital stations, functional rig with storage, ski erg, bike, treadmill",600,"Estimate. The storage center sits in the middle of the floor so nothing lives on the edges"),
("Turf warm-up, plyo and recovery zone",250,"Estimate. One zone with two jobs: warm-up and recovery"),
("Interview, testing and clinical room",200,"Estimate: enclosed, with force plates, treatment table and a desk for the intake interview"),
("Entry, lounge and tour area",150,"Estimate: compact, doubles as the tour space"),
("Restrooms (no showers)",150,"Estimate"),
("Storage, mechanical and office",150,"Estimate"),
("Spare floor (flex or gym growth)",0,"Estimate: unallocated reserve. Any space left over after the layout is drawn goes here")]
r=5
for n,v,t in zones:
    ws.cell(r,1,n); c=ws.cell(r,2,v); c.font=BLUE; c.number_format="#,##0"
    ws.cell(r,3,f"=B{r}/$B$14").number_format="0%"; ws.cell(r,5,t).alignment=Alignment(wrap_text=True,vertical="top"); r+=1
ws.cell(13,1,"Circulation and clear floor (remainder)"); c=ws.cell(13,2,"=B14-SUM(B5:B12)"); c.number_format="#,##0"
ws.cell(13,3,"=B13/$B$14").number_format="0%"; ws.cell(13,5,"Formula: facility size minus the zones above. If this falls much below about 15% the layout is too tight").alignment=Alignment(wrap_text=True,vertical="top")
ws.cell(14,1,"Facility size (Assumptions B26)").font=BOLD; c=ws.cell(14,2,"=Assumptions!$B$26"); c.font=Font(bold=True,color="FF008000"); c.number_format="#,##0"
ws.cell(15,1,"Square feet per member at the cap"); c=ws.cell(15,2,"=B14/Assumptions!$B$17"); c.number_format="0.0"; ws.cell(15,5,"Facility size divided by the member cap (Assumptions B17)")
ws.cell(16,1,"Comparison size for the per-square-foot rows (sq ft)"); c=ws.cell(16,2,4800); c.font=BLUE; c.number_format="#,##0"; ws.cell(16,5,"The earlier 4,800 sq ft layout")
for j,h in enumerate(["Per square foot, per year","Conservative","Base","Upside","Notes"]):
    c=ws.cell(18,1+j,h); c.font=HF; c.fill=HDR
rows=[("Year 1 revenue per sq ft","=Summary!{c}18/$B$14","$#,##0",""),
("Year 2 revenue per sq ft","=Summary!{c}19/$B$14","$#,##0",""),
("Year 3 revenue per sq ft","=Summary!{c}20/$B$14","$#,##0","The number to push up"),
("Year 3 operating profit per sq ft","=Summary!{c}23/$B$14","$#,##0","After the owner's salary"),
("Rent as a share of Year 3 revenue","=Assumptions!$B$28*12/Summary!{c}20","0%","Lower is better"),
("Members at month 36 per 1,000 sq ft","=Summary!{c}15/$B$14*1000","0.0","Crowding check against the cap"),
("Year 3 revenue per sq ft at the comparison size","=Summary!{c}20/$B$16","$#,##0","Same revenue spread over the larger space")]
r=19
for lab,fm,nf,nt in rows:
    ws.cell(r,1,lab)
    for j,c in enumerate("BCD"):
        x=ws.cell(r,2+j,fm.format(c=c)); x.number_format=nf; x.font=GRN if "Summary" in fm else BLK
    ws.cell(r,5,nt); r+=1
ws.cell(27,1,"What raises revenue per square foot").font=BOLD
tips=["Every square foot does two jobs: the recovery zone is also the warm-up and plyo space, the lounge is also the tour room, and the storage center joins the rig so no gear sits on the floor.",
"Store up, not out: wall-mounted and rack-integrated storage, wall-mounted screens and boards, and a clear floor.",
"The stations are the highest-value square feet. Keep them busy off-peak (see the Bay utilization tab) and sell lessons, one-on-one training and rehab, which earn more per square foot than dues.",
"At the 100 cap this space gives each member 30 sq ft. Watch peak crowding and tighten booking rules before the cap, not after.",
"No showers, no front desk and a compact lounge are already in the design and save the most space.",
"Test the layout on paper with the contractor and the net and ceiling requirements before signing the lease."]
for i,t in enumerate(tips):
    c=ws.cell(28+i,1,"• "+t); c.alignment=Alignment(wrap_text=True,vertical="top"); ws.merge_cells(start_row=28+i,start_column=1,end_row=28+i,end_column=5); ws.row_dimensions[28+i].height=32
for rr in range(5,17): ws.row_dimensions[rr].height=32 if rr<=13 else 18
ws.freeze_panes="A5"
names=wb.sheetnames; names.remove("Space plan"); names.insert(names.index("Bay utilization")+1,"Space plan")
wb._sheets=[wb[n] for n in names]
cf=wb["Changes and flags"]
for row in cf.iter_rows():
    for c in row:
        if isinstance(c.value,str) and "Equipment list, Sensitivity" in c.value and "Space plan" not in c.value:
            c.value=c.value.replace("Equipment list, Sensitivity","Equipment list, Space plan, Sensitivity")
wb.save("stage3b.xlsx")
