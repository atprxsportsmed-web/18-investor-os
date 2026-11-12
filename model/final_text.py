import openpyxl,re
wb=openpyxl.load_workbook("stage3b.xlsx")
D={("Sources and uses","E6"):"Startup rows 11 to 14. $55K: 3 Tonals ($33K with 5-year warranty), tower, weights and storage ($15K), treadmill and assault bike ($7K)",
("Space plan","A6"):"Gym floor: 3 digital stations, functional training tower with storage, treadmill, assault bike",
("Club scenario","A17"):"\u2022 Gym is $40,390: 3 Tonal 2 units at $6,130 each due at signing (equipment, delivery and year 1 of the 5-year service term, $18,390), the functional tower, free weights and storage ($15,000), a treadmill ($4,000) and an assault bike ($3,000). Service for years 2 to 5 is $1,445 per unit a year, carried in monthly software and service costs.",
("Equipment list","A24"):"Planning number for the gym floor and storage ($): $40,390: 3 Tonals ($18,390 due at signing, quoted), tower with weights and storage ($15K), treadmill ($4K) and assault bike ($3K)",
("Changes and flags","A79"):"Still placeholders: the tower, weights and cardio figures are owner estimates pending quotes; station software, if any beyond the Tonal warranty; rent at $12 per sq ft (no flex industrial figure yet); onboarding of the first 25 to 30 members before opening (about 90 hours).",
("Equipment list","A2"):"Tonal units are quoted (Tonal 2, 5-year service term: $6,130 per unit due at signing, then $1,445 per unit a year). The tower, weights and storage ($15K), treadmill ($4K) and assault bike ($3K) are the owner's figures. Replace with itemized quotes as they arrive.",
("Startup","C11"):"Formula: total bought at opening from the Equipment list: 3 Tonals, the functional tower with weights and storage, a treadmill and an assault bike",
("Assumptions","A40"):"Insurance (liability for 24/7 unstaffed use; no building yet)",
("Assumptions","E40"):"Owner estimate: $185 a month for liability coverage of a 24/7 unstaffed facility, without a building. Property coverage is not included. Confirm with a formal quote",
("Unit economics","A9"):"Card fees and follow-up coaching ($/month)",
("Unit economics","E9"):"3% card fees ($16.98) plus $22.50 of strength-coach follow-up (0.75 hour at $30)",
("Unit economics","E10"):"Before fixed costs",
("Unit economics","A53"):"Contribution from dues alone ($/month)",
("Unit economics","E53"):"$500 less card fees and follow-up coaching; the add-on share is on top",
("Equipment list","A34"):"The gym is priced from the owner's figures: Tonal 2 quote on the 5-year service option ($6,130 per unit due at signing, $1,445 per unit a year after), $15K for the tower, free weights and storage, and $4K and $3K for the treadmill and assault bike. The Tonal service payments for years 2 to 5 ($4,335 a year for 3 units) are in Assumptions B41, spread monthly. Ask the insurer about the free weights and the functional tower used alone during unstaffed hours.",
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
("(founders $449)",""),("three Tonals at $5,000 each plus a functional rig with storage, $15,000 for clinical and testing, and a $45,000 gym placeholder","three Tonals at $11,000 each (Nov 10 quote with a 5-year warranty), a functional tower with weights and storage, $15,000 for clinical and testing, and a $55,000 gym")]
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
AUDITS=["Nov 10 inputs: Tonal quote $11K each with a 5-year insurance and service warranty paid upfront; gym now $55K (tower, weights and storage $15K, treadmill $4K, assault bike $3K, no velocity device or ski erg). Insurance $185 a month for liability on a 24/7 unstaffed facility, without a building. Rent is still the $12 per sq ft placeholder.",
"Nov 10 audit fix: strength-coach follow-up pay ($22.50 per member per month) now applies to all members. Before, it was charged only to members past the first 25. Break-even now includes it (about 38 members), contribution per member is about $527, and Base Year 3 profit is about $215K. Capital required: Base $385K, Conservative $433K."]
ws=wb["Changes and flags"]
for AUDIT in AUDITS:
    if any(ws.cell(r,1).value==AUDIT for r in range(1,ws.max_row+1)): continue
    last=ws.max_row
    while ws.cell(last,1).value is None: last-=1
    r=last+1; ws.cell(r,1).value=AUDIT
    for c in range(1,ws.max_column+1):
        s=ws.cell(last,c); t2=ws.cell(r,c)
        t2.font=copy.copy(s.font); t2.alignment=copy.copy(s.alignment); t2.fill=copy.copy(s.fill); t2.border=copy.copy(s.border)
SET={("Equipment list","J5"):"Quoted",("Equipment list","K5"):"$6,130 per unit due at signing; $1,445 per unit a year for years 2 to 5. Taxes not included",
("Equipment list","J6"):"Owner figure",("Equipment list","K6"):"Functional training tower, dumbbells, kettlebells, med balls, sled, bands and the storage center",
("Equipment list","J7"):"Owner figure",("Equipment list","K7"):"Assault bike",("Equipment list","J8"):"Owner figure",("Equipment list","K8"):"Commercial treadmill"}
for r in range(9,13):
    SET[("Equipment list",f"J{r}")]=None; SET[("Equipment list",f"K{r}")]=None
for (sh,a),val in SET.items(): wb[sh][a].value=val
wb.save("stage3b.xlsx"); print(n)
