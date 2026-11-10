import openpyxl
wb=openpyxl.load_workbook("stage3b.xlsx")
ACC="Keycard access 4:30 am to 10 pm, no front desk and no standard staffed hours; the owner runs tours on set days"
SUB=[
("building that is unstaffed outside 11 am to 6 pm","building with no front desk"),
(" (staffed 11 am to 6 pm)"," (no front desk)"),
(", staffed 11 am to 6 pm",", no front desk"),
("; the building is staffed 11 am to 6 pm and closed overnight","; there is no front desk and the building is closed overnight"),
("the building is staffed 11 am to 6 pm and closed overnight","there is no front desk and the building is closed overnight"),
("TPI assessment hours per new premium member","Baseline testing hours per new performance-block member"),
("Custom routine build after the TPI assessment","Custom routine build after baseline testing"),
("(TPI assessment + routine build)","(baseline testing + routine build)"),
("Two TrackMan bays, access 4:30 am to 10 pm, TPI assessment","Two TrackMan bays, access 4:30 am to 10 pm, baseline testing"),
("Force plates (TPI / Y-Balance kit not priced)","Force plates (dynamometry, Y-Balance kit and athlete hub not priced)"),
("No TPI assessment: that is premium only","Baseline testing is part of the paid performance block"),
("what TPI screening, retests and routines may be delivered","what testing, retests and routines may be delivered"),
("The owner is the clinician, TPI provider and operator","The owner is the clinician, testing lead and operator"),
("Mix of base, premium and junior tiers","Mix of base and premium tiers"),
("The facility's 20% of Winston's and the pro's billings, weighted for juniors","The facility's 20% of Winston's and the pro's billings"),
]
FLOORNOTE='Floor coverage is the largest lever: adding a daily 7-hour hourly shift raises Conservative capital by about $139K and cuts Base Year 3 profit by about $47K. Phase coverage with membership.'
EX={
("Assumptions","A33"):"GM base salary (to be hired) ($/yr)",
("Assumptions","E33"):"Owner plan: a full-time general manager at $45K runs member experience and day-to-day operations. To be hired. Wyatt is the golf pro (1099), not the GM",
("Assumptions","A75"):"Golf pro (Wyatt): lesson revenue, 1099, keeps 80%",
("Assumptions","A93"):"Junior packages (juniors are session buyers, not members)",
("Assumptions","E64"):"Optional block: one 7-hour shift. Default is off; the owner runs tours on set days and a part-time hire covers the floor during lessons",
("Assumptions","E155"):"Optional block: 7 hours",
("Assumptions","E157"):"Default 0: no standard staffed hours. Set to 7 to test a daily 7-hour shift",
("Assumptions","E162"):"Estimate: about 20 hours a week. First job: cover the floor during the golf pro's lessons",
("Dues only","A20"):"The member-count formula assumes the 25 founding slots are full. Premium share applies to all dues-paying members. Edit the blue percentages to test other shares.",
("Dues only","A2"):"Steady-state view. Revenue is dues only: no pro or strength-coach share, no junior packages, no performance blocks, no add-ons. Costs are the model's fixed costs, card fees and contingency, including the owner's $60K and GM's $45K salaries + load.",
("Bay utilization","H6"):"Estimate: 4:30 am to 10 pm. Members enter with a keycard in that window; there is no front desk and the building is closed overnight",
("Risks and checklist","B5"):"The owner is the clinician, testing lead and operator: about 40 working hours a week, about 35 on site at first. A salaried GM runs member experience and operations. If the owner is unavailable, performance blocks and clinical care stop and the GM carries the building alone.",
("Risks and checklist","B6"):"Injury, theft, damage or misuse with nobody present. Insurers may price it high or decline. There is no front desk and no standard staffed hours, so entry control, cameras and a tablet check-in carry the building. A part-time hire covers the floor during the golf pro's lessons once profit supports it.",
("Risks and checklist","C6"):"Keycard entry, cameras with remote access, waivers, an emergency procedure and call point, an AED, clear rules, and an insurance quote that names the 4:30 am to 10 pm access window. Closed 10 pm to 4:30 am.",
("Risks and checklist","B7"):"Supervision, background checks, SafeSport and insurance for minors. Unsupervised minors can void cover. Minors must be accompanied by an adult; whether a waiver can cover other cases is an open question for counsel.",
("Risks and checklist","C7"):"Parent-signed waivers, minors only for booked sessions or with a parent or coach, ratios, background checks, insurer confirmation in writing.",
("Risks and checklist","B8"):"Winston may be treated as an employee if the facility controls hours, methods and equipment; his own independent business and other clients help. The golf pro is also a 1099 contractor paid per lesson, so the same test applies. The GM is an employee.",
("Risks and checklist","B21"):"Winston: scope, 80/20 fee split, use of the space, non-member athletes, ownership of programs and testing tools, equity terms. Wyatt: 1099 terms, 80/20 lesson split, role on the floor during lessons. GM: employment terms, salary, duties.",
("Risks and checklist","B57"):"Winston and Wyatt agreements; GM offer",
("Risks and checklist","B60"):"Owner, Winston, Wyatt, GM, advisors",
("Changes and flags","A15"):"The GM is a full-time employee at a $45K base (Assumptions B33), to be hired, paid from the first month of operation and 2 months before opening. Wyatt is the golf pro, a 1099 contractor who keeps 80% of his lesson revenue. Winston is a $0 base input and is paid from his share of billings through his own independent business. The owner is the clinician and operator.",
("Changes and flags","A23"):"Wyatt currently keeps ~70% of lessons. This version gives him the 80% lesson split as an independent pro; the $45K salary belongs to the GM. Have a CPA or attorney confirm the contractor setup.",
("Changes and flags","A29"):"Juniors are session packages, not members",
("Changes and flags","A30"):"Juniors buy lesson and strength session packages and are in the building only for booked sessions or with a parent or coach. They do not pay dues and there is no free-time access. The contractors keep 80% of package revenue, same as adults. Package prices are to be set with the pro and the strength coach.",
("Changes and flags","A31"):"Junior share of dues-paying members is 0%. No junior revenue is counted until package prices are set; the non-member athlete line (Assumptions rows 198 to 203) can be used to test it.",
("Changes and flags","A34"):"Juniors have no access windows of their own: they come for booked sessions or with a parent or coach. Entry control (keycard) and cameras are in the cost lines.",
("Changes and flags","A35"):"Removing junior members assumes the same sign-up pace is filled by adults. If juniors were part of that pace, use the sign-up sensitivity cases to see the cost.",
("Changes and flags","A48"):"Membership: $350/mo for the first 25 founding members, then $399/mo, capped at 125. Adults add a performance block (about $120/mo blended, $519 total). Juniors buy session packages only.",
("Changes and flags","A49"):"Hybrid access: members enter 4:30 am to 10 pm with a keycard. There is no front desk and no standard staffed hours; the owner runs tours on set days, cameras can be viewed remotely, and a part-time hire covers the floor during lessons once profit supports it.",
("Changes and flags","A50"):"Premium ($519 blended) is $399 dues plus a performance block: baseline testing, a prescribed 12-week routine, and a retest at the end of each block. Testing and blocks are offered, not required. All of this is the owner's time (Assumptions rows 89 to 92), not cash.",
("Changes and flags","A54"):"Premium share is 20% of dues-paying members. Founding members stay at $350 and are not modeled on premium. All shares are estimates: test them on Assumptions rows 10 to 12.",
("Changes and flags","A17"):"Floor staff: with no front desk, the owner covers the floor first. A part-time hire (Assumptions rows 160 to 172) starts when residual profit after the investor's distribution could cover two hires for several months in a row; then one is hired. Owner on-site hours then step down. An optional daily 7-hour hourly block (rows 153 to 158) is built and switched off (B157 = 0).",
("Changes and flags","A63"):"Without hourly staff, the owner carries cleaning and the floor. Turning the hourly block on (B157 = 7, one 7-hour shift a day) adds about $3,700 a month and raises Conservative capital by about $139K: the Sensitivity tab shows it. The owner is also unavailable for other work during on-site hours: check this against the owner's Chattanooga State schedule.",
("Changes and flags","A2"):"Same structure and formulas as the original. Changed: GM base salary, optional hourly floor coverage (off), contractor logic, pricing and tiers, keycard access 4:30 am to 10 pm with no front desk, two TrackMan bays, 125 cap.",
("Unit economics","E54"):"$519 less $399",
("Assumptions","E124"):"Not used: juniors are session buyers, not members",
("Base","A40"):"Gross lesson billings (Wyatt, 1099 golf pro)",
("Conservative","A40"):"Gross lesson billings (Wyatt, 1099 golf pro)",
("Upside","A40"):"Gross lesson billings (Wyatt, 1099 golf pro)",
("Risks and checklist","B10"):"$399, or about $520 a month with a performance block, is above TeeBox's reported $200 to $399.",
("Risks and checklist","A47"):"Serious golfers interviewed about $399 and a performance block",
("Founding members","F16"):"Would add a performance block (about $425 per 12 weeks)?",
("Scripts","A8"):"5. Ask three things: Would you join at $350 or $400 a month? Would a 12-week performance block at about $425 be worth it, and what would it need to include? What would stop you from joining?",
("Sources and uses","E6"):"Equipment list tab. Planning number about $42.7K: digital stations and racks with storage are the owner's figures, cardio and the rest are placeholders",
("Assumptions","E26"):"Target about 3,000 sq ft (the realtor is searching near that size). The 4,800 sq ft layout is the Sensitivity comparison",
("Scripts","A6"):"3. Describe the offer in two sentences: a members-only golf performance facility with two TrackMan bays, a performance gym and a clinical room, open 4:30 am to 10 pm to members with no front desk. Founding members pay $350 a month with a 12-month commitment, regular membership is $399, and a 12-week performance block (testing, a prescribed routine and a retest) is an optional add-on at about $425.",
}
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            v=c.value
            if isinstance(v,str) and not v.startswith("="):
                for a,b in SUB: v=v.replace(a,b)
                c.value=v
for (s,k),t in EX.items():
    if t: wb[s][k]=t
sn=wb["Sensitivity"]
for r in range(1,60):
    v=sn.cell(r,1).value
    if isinstance(v,str) and v.startswith("A base salary for Winston"):
        sn.cell(r,1).value="A base salary for Winston is a further risk. His no-base deal depends on him accepting the ramp risk, so confirm it before relying on it. The GM is modeled with a $45K base and is a separate hire from the golf pro."
    if isinstance(v,str) and v.startswith("Floor coverage is the largest lever"):
        sn.cell(r,1).value=FLOORNOTE
wb.save("stage3b.xlsx")
