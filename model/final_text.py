import openpyxl,re
wb=openpyxl.load_workbook("stage3b.xlsx")
D={("Unit economics","A9"):"Card fees and follow-up coaching ($/month)",
("Unit economics","E9"):"3% card fees ($16.98) plus $22.50 of strength-coach follow-up (0.75 hour at $30)",
("Unit economics","E10"):"Before fixed costs",
("Unit economics","A53"):"Contribution from dues alone ($/month)",
("Unit economics","E53"):"$500 less card fees and follow-up coaching; the add-on share is on top",
("Equipment list","A6"):"Functional training tower (cables, pull-up and bodyweight stations), two units, joined by the storage center",
("Equipment list","A11"):"Free weights and implements (dumbbells, kettlebells, med balls, sled, bands; no barbells)",
("Equipment list","A34"):"The planning number is a placeholder until itemized quotes arrive. Deferred items must be paid from operating cash in the early months. The digital stations' recurring subscription is not priced here: add it to Assumptions B41 once quoted. Ask the insurer about free weights and the functional tower used alone during unstaffed hours.",
("Space plan","A6"):"Gym floor: 3 digital stations, functional training tower with storage, ski erg, bike, treadmill",
("Dues only","A9"):"Junior share of regular members",
("Dues only","C9"):"Not used: no junior tier at opening",
("Dues only","A21"):"The member cap in the model is 100 (Assumptions B17). Profit at 125, 150 and 200 members is a steady-state view of dues only, shown for reference above the cap.",
("Dues only","A20"):"The member-count formula assumes the 25 founding slots are full. The second-tier share has no effect because there is one all-in price. Edit the blue percentages to test other mixes.",
("Changes and flags","A23"):"The PGA professional is not yet named. The model assumes an independent 1099 pro on an 80% lesson split; the $45K salary belongs to the GM. Have a CPA or attorney confirm the contractor setup.",
("Station capacity","H6"):"Estimate: 24/7 keycard or app access. There is no front desk. Station use is limited by demand, not by hours open",
("Assumptions","E8"):"Formula. There is no founding discount: every member pays $500 on a 12-month commitment",
("Assumptions","E9"):"No founding rate: set equal to the $500 price",
("Changes and flags","A48"):"Membership: $500/mo all in for every member, capped at 100, on a 12-month commitment. The performance onboarding is in every membership. One-on-one training, rehab and PGA instruction are bought per session or in packages. Juniors buy session packages only.",
("Changes and flags","A54"):"There is no premium tier and no founding discount. Every member pays $500 and is onboarded.",
("Changes and flags","A68"):"Members at opening: 25 conservative, 30 base, 40 upside (stated figures). All pay $500.",
("Dues only","A7"):"Founding members (paying $500)",
("Risks and checklist","C6"):"Keycard or app entry, cameras with remote access, waivers, an emergency procedure and call point, an AED, clear rules, and an insurance quote that names 24/7 unstaffed access."}
R=[("Golf pro (Wyatt)","PGA professional (to be named)"),("(Wyatt, 1099 golf pro)","(PGA professional, 1099, to be named)"),("Wyatt is the golf pro (1099), not the GM","The PGA professional (to be named) is a 1099 contractor, not the GM"),("Winston and Wyatt","Winston and the PGA professional"),("Wyatt: 1099 terms","PGA professional: 1099 terms"),("Owner, Winston, Wyatt, GM","Owner, Winston, PGA professional, GM"),("access 4:30 am to 10 pm","24/7 access"),("open 4:30 am to 10 pm on a keycard","open 24/7 on a keycard or app"),("4:30 am to 10 pm keycard access","24/7 keycard or app access"),
("keycard access 4:30 am to 10 pm","24/7 keycard or app access"),("members enter 4:30 am to 10 pm with a keycard","members enter 24/7 with a keycard or app"),
("(4:30 am to 10 pm)","(24/7)"),("4:30 am to 10 pm","24/7"),
("Founding members pay $449 a month with a 12-month commitment; regular membership is $500 all in.","Every member pays $500 all in on a 12-month commitment."),
("[$449] per month","$500 per month"),(" (founders $449)",""),("Founding membership as currently planned","Membership as currently planned"),
("(founders $449)","")]
n=0
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if not isinstance(c.value,str) or c.value.startswith("="): continue
            v=D.get((ws.title,c.coordinate),c.value)
            for a,b in R: v=v.replace(a,b)
            if v!=c.value: c.value=v; n+=1
# Audit note on the Changes and flags sheet (idempotent)
import copy
AUDIT="Nov 10 audit fix: strength-coach follow-up pay ($22.50 per member per month) now applies to all members. Before, it was charged only to members past the first 25. Break-even now includes it (about 38 members), contribution per member is about $527, and Base Year 3 profit is about $215K. Capital required: Base $385K, Conservative $433K."
ws=wb["Changes and flags"]
if not any(ws.cell(r,1).value==AUDIT for r in range(1,ws.max_row+1)):
    last=ws.max_row
    while ws.cell(last,1).value is None: last-=1
    r=last+1; ws.cell(r,1).value=AUDIT
    for c in range(1,ws.max_column+1):
        s=ws.cell(last,c); t2=ws.cell(r,c)
        t2.font=copy.copy(s.font); t2.alignment=copy.copy(s.alignment); t2.fill=copy.copy(s.fill); t2.border=copy.copy(s.border)
wb.save("stage3b.xlsx"); print(n)
