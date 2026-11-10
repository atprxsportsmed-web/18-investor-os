import openpyxl,re
from openpyxl.styles import Font,Alignment
wb=openpyxl.load_workbook("stage3b.xlsx")
EX={
("Assumptions","A6"):"Second price point ($/month): same as base (one all-in price)",
("Assumptions","E5"):"You said: $500/mo all in after the first 25. Includes the performance onboarding (interview, ROM, strength, power and injury-risk testing, a custom Tonal routine, a retest every 3 months), 4:30 am to 10 pm keycard access, the gym floor and the hitting stations",
("Assumptions","E6"):"Same as B5. Every member gets the performance onboarding, so there is no separate premium price",
("Assumptions","E8"):"Formula. Founding members pay $449 against $500 on a 12-month commitment; a lifetime price lock is under consideration",
("Assumptions","E9"):"You said: $449/mo for the first 25 founding members",
("Assumptions","A10"):"Mix of non-founding members: base tier (not offered)",
("Assumptions","E10"):"Set to 0. There is one all-in price",
("Assumptions","A11"):"Mix of non-founding members: all-in membership",
("Assumptions","E11"):"Set to 100%. Every non-founding member pays B6 and carries the follow-up cost in B191",
("Assumptions","E17"):"You said: 100. Hard cap, never raised to fill a month. A coached club works with fewer members. The Station capacity tab checks what five stations carry. Waitlist once full",
("Assumptions","A18"):"Member cap once a later TrackMan bay is live (not used at opening)",
("Assumptions","E18"):"Kept at 100 to match B17. There is no TrackMan bay at opening",
("Assumptions","E21"):"25 conservative, 30 base, 40 upside. The first 25 are the founding members; anything above 25 pays $500",
("Assumptions","E22"):"My assumption: the earlier pace (2.5 / 4 / 6) x 0.80 for the $500 price. Replace with what the founding-member conversations show",
("Assumptions","A41"):"Software, station subscriptions, billing platform",
("Assumptions","E41"):"Estimate: $400 plus about $300 for station software and subscriptions. Not quoted",
("Assumptions","A88"):"Onboarding and retest time (the owner's hours)",
("Assumptions","A89"):"Monthly follow-up and retest (hours per member)",
("Assumptions","A90"):"Custom routine build (hours per new member)",
("Assumptions","A91"):"Owner onboarding hours per week at month 12 (Base case)",
("Assumptions","E91"):"Formula: (new members x (testing + routine hours) + members x follow-up hours not delivered by the coach) / 4.33. Founding members are excluded",
("Assumptions","A92"):"Owner onboarding hours per week at month 36 (Base case)",
("Assumptions","E92"):"Formula. Onboarding is the one product that depends on the owner's time",
("Assumptions","A110"):"Later TrackMan bay (not planned)",
("Assumptions","A111"):"Consecutive months at the cap before a later bay is ordered (not used)",
("Assumptions","E111"):"Not used. No TrackMan bay is planned",
("Assumptions","A112"):"TrackMan bay cost, all-in ($)",
("Assumptions","E112"):"Formula. Not used at opening",
("Assumptions","A113"):"Added monthly cost once a bay is live ($)",
("Assumptions","E113"):"Estimate: software, utilities, upkeep. Not used at opening",
("Assumptions","A117"):"Five performance stations, no TrackMan bay, access 4:30 am to 10 pm, performance onboarding",
("Assumptions","A118"):"Bay switch (technical setting, keep at 2: no bay is bought)",
("Assumptions","E118"):"Leave at 2 with B119 at 0. Setting it to 1 would switch on the old staged-bay logic. No TrackMan bay is planned",
("Assumptions","A119"):"TrackMan bay cost per bay ($): none at opening",
("Assumptions","E119"):"0. There is no TrackMan bay in this plan. The five stations are $82,197.10 in total, all in, and are priced on Startup row 7",
("Assumptions","A120"):"Extra enclosure, screen, turf and lighting per added bay ($)",
("Assumptions","E120"):"Set to 0. Not used",
("Assumptions","E123"):"2 hours: the intake interview and baseline testing. Included in every membership",
("Assumptions","A124"):"Junior second tier ($/month): not used",
("Assumptions","A125"):"Share of juniors on the second tier: not used",
("Assumptions","A126"):"Share of regular members on the second price point",
("Assumptions","E126"):"Formula. Set to 100%: every regular member is on the all-in price",
("Assumptions","A132"):"Onboarding hours as a share of the owner's on-site hours (Base case, month 36)",
("Assumptions","E132"):"Formula: how much of the owner's time on site onboarding uses. The rest is admin, marketing, sales and tours",
("Assumptions","A183"):"Only the base adult tier steps up when tiered pricing is on (off in this plan)",
("Assumptions","A188"):"Follow-up pay for the strength coach",
("Assumptions","A189"):"Share of follow-up hours delivered by the strength coach",
("Assumptions","A191"):"Follow-up cost per member per month",
("Assumptions","E191"):"Formula: follow-up hours per member x rate x share delivered by the coach. Applies to every member",
("Summary","A39"):"TrackMan bay (none at opening)",
("Summary","A40"):"Month a later bay goes live (not used)",
("Summary","E40"):"Not used. There is no TrackMan bay at opening",
("Summary","A41"):"Peak cumulative deficit",
("Summary","A43"):"Your onboarding hours per week",
("Startup","A6"):"TrackMan bays: none at opening",
("Startup","C6"):"No TrackMan bay in this plan and none planned",
("Startup","A7"):"5 performance stations, all in (launch monitors, software, hardware, net, footing, iPad)",
("Startup","C7"):"You said: $82,197.10 for all five stations, all in. Get quotes",
("Startup","A8"):"Extra net or partition (not used)",
("Startup","C8"):"Set to 0. The station price includes the net and footing",
("Startup","A9"):"Flooring: gym and warm-up zone only",
("Startup","C9"):"Estimate. The stations' footing is in their $13,000",
("Startup","C11"):"Formula: total bought at opening from the Equipment list. Planning number $45K (placeholder); replace estimates with quotes there",
("Startup","A16"):"Testing equipment: force plates, dynamometry, ROM tools (athlete hub not priced)",
("Startup","C16"):"You said: $15,000 budget for the clinical and testing area, including the next two lines. Replace with quotes",
("Startup","A39"):"Operating losses until break-even (peak cumulative deficit)",
("Startup","E39"):"From each monthly model sheet, row 50",
("Unit economics","E6"):"One all-in price for every non-founding member",
("Unit economics","E21"):"Shows what an hour of follow-up earns compared with the owner's salary",
("Unit economics","A52"):"7. Founding members and capacity",
("Unit economics","E53"):"$449 less card fees. The first 25 members",
("Unit economics","A54"):"Extra revenue per member from a second tier (none: one all-in price)",
("Unit economics","E54"):"$500 less $500",
("Unit economics","A55"):"Owner onboarding hours per member per month",
("Unit economics","E56"):"What an hour of follow-up earns compared with the owner's hourly cost",
("Dues only","A7"):"Founding members (paying $449)",
("Dues only","A13"):"Second-tier share of regular members (no effect: one price)",
("Dues only","F13"):"Monthly profit at the member cap",
("Dues only","A20"):"The member-count formula assumes the 25 founding slots are full. The second-tier share has no effect because there is one all-in price. Edit the blue percentages to test other mixes.",
("Sources and uses","A5"):"Golf area: 5 performance stations (hardware, software, net, footing) and gym-zone flooring",
("Sources and uses","E31"):"Trigger: build-out complete and all 25 founding slots committed. Covers stations, gym, testing, recovery, lobby and contingency",
("Sources and uses","E6"):"Startup rows 11 to 14. $45K gym placeholder",
("Risks and checklist","C5"):"Written procedures, a trained backup, key-person insurance, an advisor. Plan how onboarding and retest hours are covered during illness.",
("Risks and checklist","B10"):"$500 all in is above TeeBox's reported $200 to $399, so it has to be justified by the coached program and the results.",
("Risks and checklist","A16"):"Station capacity at peak",
("Risks and checklist","B16"):"Five stations carry a limited number of members at peak times. The member cap can exceed what the stations can serve, and members who cannot book leave.",
("Risks and checklist","C16"):"Peak pricing or booking limits, off-peak incentives, a lower advertised cap, or adding a sixth station. Replace the peak-demand estimates with booking data after opening.",
("Risks and checklist","D16"):"Station capacity tab, Assumptions B17",
("Risks and checklist","B22"):"Term, exit, expansion rights, tenant-improvement allowance, a use clause for 4:30 am to 10 pm access, and enough ceiling height and depth for full swings into a net wall",
("Risks and checklist","A47"):"Serious golfers interviewed about $500 all in and the performance onboarding",
("Guide","B26"):"Whether the five stations can carry the membership at peak times",
("Guide","C26"):"Row 21 (members the stations carry) against the cap in row 23",
("Guide","B28"):"Can membership dues alone cover costs, at different member counts?",
("Guide","C57"):"Is the 24-month ratio above 5x? If so, ask whether CAC is understated. Then see how many members the five stations carry at peak versus the cap.",
("Guide","B80"):"Class, event or private-rental bookings that take stations away from members.",
("Scripts","A6"):"3. Describe the offer in two sentences: a members-only golf performance club with five data-driven hitting stations, a performance gym and a clinical testing room, open 4:30 am to 10 pm on a keycard with no front desk. Every member starts with an interview, testing and a custom routine, retested every 3 months. Founding members pay $449 a month with a 12-month commitment; regular membership is $500 all in.",
("Scripts","A8"):"5. Ask three things: Would you join at $450 or $500 a month with the onboarding and retests included? Would you add one-on-one training, rehab or lessons, and how often? What would stop you from joining?",
("Scripts","A18"):"Founding membership as currently planned: [$449] per month, a 12-month commitment from the opening date, access 4:30 am to 10 pm, the gym floor and the hitting stations, with the onboarding and a retest every 3 months included.",
("Sensitivity","A36"):"Dues price, add-on take-up and sign-up pace each move profit by $40K to $75K a year, but none of them break the case alone. They matter most together, and in the Conservative case.",
("Changes and flags","A38"):"No TrackMan bay at opening and none planned. The member cap is 100 from month 1 (Assumptions B17). The old staged-bay logic is switched off by keeping B118 at 2 with the bay cost at 0 (B119).",
("Changes and flags","A39"):"If a TrackMan bay is ever wanted, see Assumptions rows 110 to 120; the old staged-bay logic would need switching on.",
("Changes and flags","A40"):"Not modeled: lost sign-ups if a cap is advertised, and whether the layout fits five stations in a row. Confirm the floor plan, ceiling height and net clearance before the deck says five stations.",
("Changes and flags","A48"):"Membership: $449/mo for the first 25 founding members, then $500/mo all in, capped at 100. The performance onboarding is in every membership. One-on-one training, rehab and PGA instruction are bought per session or in packages. Juniors buy session packages only.",
("Changes and flags","A50"):"Onboarding for every member: an intake interview, ROM, strength, power and injury-risk testing, a custom Tonal routine, and a retest every 3 months. The strength coach is paid about $22.50 per member per month for follow-up; the owner's onboarding time is Assumptions rows 89 to 92.",
("Changes and flags","A51"):"Check the cap against the Station capacity tab, which shows how many members five stations carry at peak hours. Gym floor and downtown parking are other likely limits. Use booking rules rather than headcount alone as the real limit.",
("Changes and flags","A53"):"One price point",
("Changes and flags","A54"):"There is no premium tier. Every regular member pays B6 and carries the follow-up cost. Founding members pay $449 and are also onboarded.",
("Changes and flags","A55"):"Owner time is the limit on onboarding. Check Summary rows 44 and 45 against the owner's 10 hours/week building cap. If Winston runs the routines, he will expect to be paid out of dues; the follow-up line covers that.",
("Changes and flags","A65"):"A cap of 100 keeps the club coached rather than crowded. The stations book up before the floor does. Gym equipment, station and recovery-space prices are still estimates until quotes arrive.",
("Changes and flags","A68"):"Members at opening: 25 conservative, 30 base, 40 upside (stated figures). The first 25 pay the $449 founding price.",
}
# Monthly sheet row labels
for n in ("Base","Conservative","Upside"):
    EX.update({(n,"A22"):"Software + billing + access control (+ added bay once live)",(n,"A46"):"Bay switch (technical: 1 = no bay is bought)",(n,"A47"):"Bay cost if bought after opening (none)",
     (n,"A49"):"Month a later bay goes live (not used)",(n,"A50"):"Peak cumulative deficit",(n,"A52"):"New members this month (interview, testing and routine build)",
     (n,"A53"):"Members (follow-up load)",(n,"A54"):"Owner onboarding hours per week (interviews, testing, routines, and any follow-ups not delivered by the strength coach)",(n,"A63"):"Onboarding hours as a share of owner on-site hours"})
EX.update({("Assumptions","E78"):"Strength sessions and lessons are both priced at $150 per session",
("Risks and checklist","B9"):"The model assumes 35% to 40% of members buy $150 sessions on top of dues. Unproven.",
("Changes and flags","A43"):"Strength sessions and lessons are both $150 per session (stated). Take-up of 35% and 40% on top of $500 dues is unproven. Test lower take-up on Assumptions rows 70 and 76."})
EX.update({("Unit economics","A2"):"Per member at the flat $500 rate. Conservative framing: acquisition cost includes launch spend and selling time, and lifetime value is capped at 12 and 24 months.",
("Unit economics","A6"):"Dues per member ($/month)",
("Unit economics","E6"):"One flat all-in price for every member",
("Unit economics","A52"):"7. Capacity",
("Unit economics","A53"):"Contribution from dues alone ($/month)",
("Unit economics","E53"):"$500 less card fees; add-on share is on top",
("Unit economics","E14"):"Pre-opening sign-ups before the doors open",
("Sources and uses","E31"):"Trigger: build-out complete and 25 pre-opening commitments in hand. Covers stations, gym, testing, recovery, lobby and contingency",
("Sources and uses","E30"):"Trigger: lease signed and at least 15 pre-opening commitments (signed intent, no money collected). Covers the lease deposit, build-out and flooring"})
for (sh,c),t in EX.items():
    if sh in wb.sheetnames: wb[sh][c].value=t
# Bay utilization -> Station capacity
B=wb["Bay utilization"]
BL=Font(color="FF0000FF")
P={"A1":"Station capacity: can the five stations carry the membership?","A2":"Compares station-hours members want with station-hours available. Bookings, not the member cap, set the real limit. Every blue input is an estimate until booking data exists.",
"A5":"Hitting stations at opening","H5":"Five stations ($82,197.10 in total)","A6":"Bookable hours per station per day","A9":"Share of members who use the stations","H9":"Estimate: a coached club, so most members use them",
"A10":"Station sessions per user per month","H10":"Estimate: about two a week","A12":"Lesson hours per member per month","H12":"Lessons also use a station: share taking lessons x lessons a month x 1 hour",
"A13":"Station hours per member per month","A14":"Peak hours per week (all stations)","A19":"Station hours available per month","A20":"Peak station hours available per month","H20":"Peak hours per week x stations x weeks",
"A21":"Members the stations carry at the peak ceiling","A22":"Members the stations carry across all hours at the ceiling","A24":"Members above what the stations carry at peak, at the cap",
"H24":"If above 0, the cap is not deliverable without booking rules, peak pricing or another station","D28":"Station hours demanded / month",
"A39":"First month each case passes what the stations carry at peak","A44":'="Five stations carry about "&TEXT(B21,"0")&" members at peak. The cap is "&B23&"."'}
for c,t in P.items(): B[c].value=t
B["B5"].value=5; B["B5"].font=BL; B["B9"].value=0.9; B["B10"].value=8
tips=["Peak pricing or a peak-hour booking limit per member (for example 2 peak sessions a month).","Booking windows that reserve off-peak hours for lessons and juniors.","Incentives to shift sessions to mornings and midday.","Add a sixth station if booking data shows peak demand.","A lower advertised cap until booking data shows where demand lands."]
for i,t in enumerate(tips): B[f"A{47+i}"].value="• "+t
# sensitivity A24 etc fine. Append Changes entry
cf=wb["Changes and flags"]; r=cf.max_row+2
cf.cell(r,1,"Club model (November 3, 2026)").font=Font(bold=True)
items=["Switched the main model to the club structure: $500 all in (founders $449), cap 100, performance onboarding in every membership, five stations at $82,197.10 in total, no TrackMan bay (none planned), three Tonals at $5,000 each plus a functional rig with storage, $15,000 for clinical and testing, and a $45,000 gym placeholder.",
"The earlier $399 plan with two TrackMan bays is kept as a static comparison on the Club scenario tab. Sign-up pace is scaled to 0.80 for the higher price; that is an assumption.",
"The Bay utilization tab is now Station capacity. Space plan was updated to match.",
"Still placeholders: cardio, free weights and the velocity device inside the $45,000 gym placeholder; station software (+$300 a month); onboarding of the first 25 to 30 members before opening (about 90 hours)."]
for i,t in enumerate(items):
    c=cf.cell(r+1+i,1,t); c.alignment=Alignment(wrap_text=True,vertical="top")
# global substitutions
SUB=[("Eighteen Athletic Club","18 Golf Performance Club"),("Bay utilization","Station capacity"),("Premium and capacity","Founding members and capacity")]
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value,str) and not c.value.startswith("="):
                v=c.value
                for a,b in SUB: v=v.replace(a,b)
                c.value=v
wb["Bay utilization"].title="Station capacity"
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value,str) and c.value.startswith("=") and "Bay utilization" in c.value: print("FORMULA REF",ws.title,c.coordinate)
wb.save("stage3b.xlsx")
wb=openpyxl.load_workbook("stage3b.xlsx")
wb["Station capacity"]["A15"].value="Share of station demand that lands in peak hours"
wb["Risks and checklist"]["A43"].value="Members at the cap as a share of the golfer market"
wb["Changes and flags"]["A60"].value="New sheet 'Dues only' answers whether dues alone sustain the facility at different member counts."
wb.save("stage3b.xlsx")
