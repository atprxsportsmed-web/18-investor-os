import re
MANUAL={
("Sources and uses","E23"):"Refundable deposits are a liability, not income, so they are not counted by default. Enter only money available to spend",
("Unit economics","E15"):"Marketing spend / new members. Likely understated until real sources are tracked",
("Assumptions","E18"):"Kept at 200. Only matters when bays at opening (B118) is set to 1 to test the staged version",
("Assumptions","E23"):"Estimate. Real operator churn is private data; replace after calling operators",
("Risks and checklist","C43"):"Fills in once the first number above is entered. A small share is a stronger story.",
("Guide","B49"):"On Assumptions read column E. List every input marked 'Estimate' that can be replaced with a quote or data.",
("Guide","B57"):"Read the Sensitivity snapshot. To run a new case, change the input on Assumptions and read the Summary.",
("Guide","C57"):"Pick the three cases that hurt most and write the response to each.",
("Assumptions","E91"):"Formula: (new premium members x (assessment + routine hours) + premium members x follow-up hours) / 4.33. Founding members are excluded. The cap while building is 10 hrs/week total",
("Assumptions","E92"):"Formula. This is the one product that depends on the owner's time, so it caps how many premium members can be served",
("Risks and checklist","B25"):"Confirm what TPI screening, retests and routines may be delivered for a fee in TN and GA",
("Changes and flags","A54"):"Premium share is 20% of regular members (adults) plus 25% of juniors, about 25% of regular members overall. Founding members stay at $350 and are not modeled on premium. All shares are estimates: test them on Assumptions rows 11, 12 and 125.",
("Changes and flags","A64"):"Being on site is not the same as supervising. If juniors come in during the owner's hours while the owner is in an appointment, an insurer may still treat them as unsupervised. Ask the insurer how supervision is defined before setting junior windows.",
("Changes and flags","A65"):"Guidance is that 150 to 200 members is where a facility starts to feel crowded, so a 200 cap is the top of that range. The bays book up before the floor does. Gym equipment and recovery-space prices are still placeholders until quotes arrive.",
("Changes and flags","A69"):"Equipment: gym floor $30,000 all-in including VBT, force plates $3,000, recovery zone $3,000. The HHD is already owned. These replace placeholders for the rig, cable tower, cardio, free weights, testing kit and recovery zone. Quotes are still pending.",
("Changes and flags","A33"):"Rent: the model still uses $12/sq ft all-in, based on industrial flex listings. If the space is downtown, expect more. Update Assumptions B27 once a real quote is available.",
("Changes and flags","A45"):"Owner pay during the build is $55K (Assumptions B115); after opening it is $75K. The model assumes 4 paid months before opening (B49). If the build runs closer to 12 months, that line grows by about $5,100 per month.",
("Changes and flags","A72"):"Guide, Investor returns, Sources and uses, Unit economics, Sensitivity, Ownership scenarios and Risks and checklist are new. The monthly sheets now run 60 months. Seasonality factors (Assumptions rows 135 to 149) are added and switched off. A $10,000 putting and chipping green is added to the build-out (Assumptions B151).",
("Base","A2"):"60 months from opening. All formulas pull from the Assumptions sheet.",
("Conservative","A2"):"60 months from opening. All formulas pull from the Assumptions sheet.",
("Upside","A2"):"60 months from opening. All formulas pull from the Assumptions sheet.",
("Guide","B12"):"Notes in column E of Assumptions give the source or basis of each input. 'Estimate' marks a placeholder to replace with a quote or data.",
("Guide","A5"):"It estimates what the facility costs to open, how members build over five years, when it turns profitable, what it pays an investor, and what could go wrong. Every input can be changed and the effect flows through all tabs.",
("Guide","B33"):"Everything changed from the original file, and what to confirm",
("Guide","B8"):"An input that can be changed.",
("Summary","E28"):"Change this to test any member count (about 50 is the survival target discussed)",
("Investor returns","I5"):"Non-loan investment",
("Investor returns","I6"):"Equity offered for the investment",
("Investor returns","I7"):"Estimate: the rest stays in the business for reserves and equipment",
("Investor returns","I8"):"Estimate: investors will question this. Small businesses commonly sell for low single-digit multiples of profit; verify with a broker",
("Investor returns","I12"):"Paid at the end of that year. A later year protects the business if the ramp is slow; an earlier year favors the investor",
("Investor returns","A27"):"IRR is the yearly return that makes the investor's cash flows net to zero. It is only meaningful when the exit is real: a buyer, or an agreed buyback. 'n/a' means the cash flows never turn positive.",
("Sources and uses","E6"):"Startup rows 11 to 14. $30K budget",
("Sources and uses","E21"):"Non-loan investment",
("Sources and uses","E22"):"Base plan: the business funds itself. Enter an amount only if the owner contributes cash",
("Sources and uses","A36"):"Amounts are Conservative-sized. Staging lowers the investor's risk because money is released against proof, and it can support a better valuation. Triggers are suggestions to be agreed with counsel.",
("Unit economics","E21"):"Shows what an hour of premium follow-up earns compared with the owner's salary",
("Unit economics","E16"):"Looks very high because CAC counts only the marketing line, not selling time or the pre-opening campaign. Not suitable to quote to investors as is",
("Sensitivity","A2"):"SNAPSHOT. These values were produced by changing the inputs in column K on a copy of this model and recalculating. They do not update when inputs change. To rerun a case, make the same change on Assumptions and read the Summary.",
("Assumptions","A2"):"Blue = input. Black = formula. Yellow fill = key lever. Notes marked 'Estimate' are placeholders to verify; other notes give the stated figure or the basis.",
("Assumptions","E8"):"Formula. Founding members pay $350 against $399 on a 12-month commitment; a lifetime price lock is under consideration",
("Assumptions","E17"):"200. Guidance is that 150 to 200 members is where a facility starts to feel crowded, so 200 is the top of that range: two TrackMan bays, 5 Rapsodo stations, the gym and the recovery space. The bays book up before the floor does. Revisit once the floor plan is final",
("Assumptions","E27"):"Chattanooga industrial flex listings run about $8 to $10.50/sf/yr (Crexi, CBRE, CommercialCafe, 2026), plus about $1.50 for NNN/CAM",
("Assumptions","E29"):"Estimate and the biggest swing in the plan. Replace with the contractor quote",
("Assumptions","E32"):"$75K salary, not including equity or distributions. Build-period pay is on row 115. The 12% payroll load on row 35 is applied on top; confirm with a CPA how an owner is paid in an LLC",
("Assumptions","E33"):"1099, no base pay. Enter 38000 here to test the $38K to $40K base that was floated",
("Assumptions","E40"):"Estimate: raised from $500 for 24/7 unstaffed access and minors. Get a real quote and disclose both to the insurer",
("Assumptions","E43"):"Estimate: raised from $800 because there is no staff to clean. Budget an outside cleaning service plus supplies and repairs. Get a quote",
("Assumptions","E44"):"Estimate (see the CPA quote on file)",
("Assumptions","E49"):"Estimate: 4 months. Target opening is Oct/Nov 2027, so a longer paid build is likely; change this and the cost updates",
("Assumptions","E57"):"Suggested structure, negotiable",
("Assumptions","E59"):"Estimate: small-business multiples commonly run low single digits; verify with a broker or investor",
("Assumptions","E62"):"$15/hr, intern level. Cleaning, restocking, open/close checks. There are no official staffed hours",
("Assumptions","E63"):"The owner is on site about 7 hrs/day on weekdays, so there are no hourly staff. Enter days here only to test paying staff",
("Assumptions","E73"):"80/20 on in-house training, read as the contractor keeping 80% and the facility 20%. Reverse here if that is wrong",
("Assumptions","E79"):"80/20 split, read the same way as Winston: the pro keeps 80%. The pro currently keeps 70%",
("Assumptions","E90"):"Estimate. If Winston builds the routines, he will expect to be paid from premium dues: not modeled",
("Assumptions","A109"):"Juniors have limited-time access windows set by the owner and no staffed hours. Training and teaching are by appointment only.",
("Assumptions","E115"):"At least $55K during full business-building and advertising mode. Salary after opening stays on row 32. Each extra month of build costs this / 12 x (1 + payroll load), about $5,100",
("Assumptions","E129"):"Weekdays. Saturday and Sunday are by appointment only, with no standing hours",
("Assumptions","E136"):"Target opening Oct/Nov 2027. 11 = November. Month 1 of the model falls in this calendar month",
("Assumptions","E130"):"About 7",
("Startup","A2"):"Costs in blue are inputs; notes marked 'Estimate' are placeholders. Get quotes before the deck is final.",
("Startup","C16"):"Force plates about $3,000; VBT is in the gym line; HHD already owned. Add a Y-Balance kit if one needs to be bought",
("Startup","C29"):"Estimate: $4,000 plus access control and cameras (Assumptions B121)",
("Startup","C11"):"Full gym floor equipment budgeted at about $30,000, VBT included. Replace with real quotes",
("Startup","C18"):"Recovery zone about $3,000",
("Ownership","E24"):"After payback the investor keeps the offered equity % and the owner's share goes up to the rest",
("Risks and checklist","A5"):"Key person: owner",
("Risks and checklist","B5"):"The owner is the GM, the clinician and the TPI provider, on site about 35 hours a week. If the owner is unavailable, the business has no operator.",
("Risks and checklist","C5"):"Written procedures, a trained backup, key-person insurance, an advisor. Plan how premium service hours are covered during illness.",
("Risks and checklist","C10"):"Presale evidence at the stated price. Show the clinical and S&C access competitors do not offer.",
("Risks and checklist","B19"):"Roles, 75% to 80% control, investor rights, distributions and exit (revisit if the owner stake is lowered)",
("Risks and checklist","B26"):"Protect the testing battery and program documents before sharing outside",
("Risks and checklist","D30"):"Source was an AI search relayed by an investor, who doubted the revenue figure. Pull the actual FDD before using any number.",
("Risks and checklist","D33"):"Possible partner. The partnership plan has not yet been shared with the owner.",
("Risks and checklist","C39"):"Source: own list",
("Risks and checklist","B60"):"Owner, Winston, the pro, advisors",
("Changes and flags","A1"):"Changes from the original file, and things to check",
("Changes and flags","A2"):"Same structure and formulas as the original. Changed: staffing and contractor logic, pricing and tiers, 24/7 access, two TrackMan bays, on-site time, 200 cap.",
("Changes and flags","B4"):"Original file",
("Changes and flags","A12"):"'Original file' column is typed in from the workbook received (blue). 'This version' links to the live model (green).",
("Changes and flags","A22"):"Split direction: 80/20 is read as the contractor keeping 80%. If it is the reverse, change rows 73 and 79 on Assumptions.",
("Changes and flags","A31"):"Junior share of regular members is 20% (was 10%) because junior golf is seen as a major downtown Chattanooga opportunity. Estimate: replace with real interest from clubs and programs.",
("Changes and flags","A39"):"The second bay now costs $10,000 all-in (TrackMan plus full build-out) and adds $250/month in running costs. It is paid from operating cash, and the peak-deficit line includes it.",
("Changes and flags","A40"):"Not modeled: lost sign-ups if a cap is advertised, and whether the layout fits two bays. Confirm the floor plan before the deck says two bays.",
("Changes and flags","A44"):"Junior package prices ($45 S&C, $75 teaching) are still group-rate placeholders. Confirm that Winston and Wyatt will teach juniors at those rates, or raise them.",
("Changes and flags","A26"):"Owner clinical revenue (ATPRX housed in the facility) is not in this model. It is upside, and it needs its own treatment if ATPRX, LLC is the parent.",
("Changes and flags","A15"):"The owner is the GM: Wyatt's $65K GM salary and Winston's $55K salary are now $0 base inputs. Both are 1099 and paid only from their share of billings.",
("Changes and flags","A17"):"Floor staff: $0. The owner staffs the facility about 7 hours a day on weekdays, and weekends are by appointment only (Assumptions rows 63, 64 and 129 to 132).",
("Changes and flags","A58"):"Owner salary is $75K after opening (not including equity or distributions) and is the only salary in the model. Winston and Wyatt have no base. There are no hourly staff.",
}
def generic(v):
    v=re.sub(r"^You said:\s*","",v)
    v=re.sub(r"^My assumption\s*\(([^)]*)\)\s*$",r"Estimate (\1)",v)
    v=re.sub(r"^My assumption:\s*","Estimate: ",v)
    v=re.sub(r"^My assumption\.\s*","Estimate. ",v)
    v=re.sub(r"^My assumption\s*$","Estimate; verify",v)
    v=re.sub(r"\bmy placeholders?\b","placeholders",v,flags=re.I)
    v=re.sub(r"\bYour own\b","The owner's own",v)
    v=re.sub(r"^You: ","",v)
    v=re.sub(r"^Your ","Owner ",v)
    v=re.sub(r"\byour CPA\b","a CPA",v); v=re.sub(r"\byour attorney\b","an attorney",v); v=re.sub(r"\byour insurer\b","the insurer",v)
    v=re.sub(r"\byour investor friend\b","an investor",v)
    v=re.sub(r"\byour figures\b","stated figures",v)
    v=re.sub(r"\byour\b","the owner's",v)
    v=re.sub(r"\bYou are\b","The owner is",v); v=re.sub(r"\byou are\b","the owner is",v)
    v=re.sub(r"\bAfter your\b","After the owner's",v)
    if v and v[0].islower(): v=v[0].upper()+v[1:]
    return v
def run(wb):
    changed=[]
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                v=c.value
                if not isinstance(v,str) or v.startswith("="): continue
                if ws.title=="Ownership scenarios": continue
                key=(ws.title,c.coordinate)
                nv=MANUAL.get(key)
                if nv is None:
                    nv=generic(v)
                if nv!=v: c.value=nv; changed.append(key)
    return changed

MANUAL.update({
('Guide', 'B79'):'Damage or theft cost in a 24/7 building beyond the insurance line, including outside staffed hours (Mon-Sat 12 hrs).',
('Assumptions', 'E7'):'$200/mo with limited-time access windows aligned to staffed hours. Packages are paid additions',
('Assumptions', 'A36'):'Floor staff ($/month, loaded)',
('Assumptions', 'E40'):'Estimate: raised from $500 for 24/7 access and minors. Get a real quote and disclose both to the insurer',
('Assumptions', 'E43'):'Estimate: raised from $800 for an outside cleaning service plus supplies and repairs, in addition to staff cleaning at crossover. Get a quote',
('Assumptions', 'E62'):'$15/hr. Cleaning, restocking, front desk, open/close checks during staffed hours',
('Assumptions', 'A109'):'Juniors have limited-time access windows set by the owner, aligned to staffed hours. Training and teaching are by appointment.',
('Assumptions', 'E108'):'Formula. Juniors use the facility only in staffed or supervised hours: waivers and insurance need to cover minors (see Changes and flags)',
('Risks and checklist', 'A6'):'24/7 access outside staffed hours',
('Risks and checklist', 'B6'):'Injury, theft, damage or misuse with nobody present outside the 12 staffed hours. Insurers may price it high or decline.',
('Changes and flags', 'A2'):'Same structure and formulas as the original. Changed: hourly floor-staff coverage, contractor logic, pricing and tiers, 24/7 access, two TrackMan bays, 200 cap.',
('Changes and flags', 'A17'):'Floor staff: $15/hr, 2 shifts of 6 hours with a 30 minute crossover (12.5 paid hrs/day), 6 days/week, 5% buffer, 12% payroll load: about $68,800/yr (Assumptions rows 62 to 67 and 153 to 158). This frees the owner to work as clinician and operator.',
('Changes and flags', 'A32'):'Juniors under 18: waivers signed by a parent, background checks for anyone who coaches or supervises minors, and insurance that covers minors need to be confirmed with the insurer and attorney. Outside staffed hours minors may be on site with no staff present: insurers often treat that as a major issue.',
('Changes and flags', 'A49'):'Full members have 24/7 access; staff are on site Mon-Sat for 12 hours a day, with no staff overnight or on Sundays. Clinical and strength services are by appointment. Access control, cameras and a higher insurance line are added for unstaffed hours.',
('Changes and flags', 'A57'):'Hourly floor coverage and one salary',
('Changes and flags', 'A58'):'Owner salary is $75K after opening (not including equity or distributions). Winston and Wyatt have no base. Hourly floor staff are a separate loaded line of about $5,700/month.',
('Changes and flags', 'A63'):'Hourly staff cover 12 hours a day, 6 days a week, so the owner is not needed for cleaning or front desk. Fixed costs rise to about $23,500/month and break-even to about 53 members. Staffing is the largest new fixed cost: the Sensitivity tab shows the effect of removing it.',
('Changes and flags', 'A72'):'Guide, Investor returns, Sources and uses, Unit economics, Sensitivity, Ownership scenarios and Risks and checklist are new. The monthly sheets now run 60 months. Seasonality factors (Assumptions rows 135 to 149) are added and switched off. A $10,000 putting and chipping green is added to the build-out (Assumptions B151). Hourly staff coverage is added (Assumptions rows 153 to 158).',
('Risks and checklist', 'C7'):'Parent-signed waivers, limited-time windows aligned to staffed hours, ratios, background checks, insurer confirmation in writing.',
})
MANUAL.update({
("Assumptions","E63"):"Formula from the hourly coverage block (B157). Set B157 to 0 to remove hourly staff",
("Assumptions","E64"):"Formula from the hourly coverage block (B158): two 6-hour shifts plus a 30-minute crossover",
("Assumptions","E36"):"Formula from the floor-staff block below: two shifts a day, six days a week, so the owner can work as clinician and operator",
("Assumptions","E121"):"Estimate: keyfob or app entry, interior cameras, door sensors, emergency call point. Not linked: Startup B29 is a typed $10,000 quote (Pass Factory + cameras)",
})

MANUAL.update({
("Assumptions","A32"):"Owner salary after opening ($/yr)",
("Assumptions","E32"):"$60K salary, not including equity or distributions. Build-period pay is on row 115. The 12% payroll load on row 35 is applied on top: confirm with a CPA how an owner is paid in an LLC",
("Assumptions","A33"):"GM / head golf pro: base salary ($/yr)",
("Assumptions","E33"):"$45K base for a salaried GM who runs the floor day to day. Check whether the GM also keeps the 80% lesson split (rows 73 and 79), which would double count",
("Assumptions","E62"):"$15/hr. Used only if staffed days (B157) are above 0. Set to 0 days in this version: no hourly floor staff",
("Assumptions","E63"):"Formula from the hourly coverage block (B157). Currently 0: no hourly floor staff",
("Assumptions","E36"):"Formula from the floor-staff block below. Zero while staffed days (B157) is 0",
("Assumptions","A36"):"Floor staff ($/month, loaded)",
("Changes and flags","A2"):"Same structure and formulas as the original. Changed: GM base salary, optional hourly floor coverage (off), contractor logic, pricing and tiers, 24/7 access, two TrackMan bays, 200 cap.",
("Changes and flags","A17"):"Floor staff: none in this version. An optional hourly block (Assumptions rows 153 to 158: $15/hr, two 6-hour shifts, 30 minute crossover, 6 days) is built and switched off (B157 = 0). The GM covers the floor.",
("Changes and flags","A45"):"Owner pay during the build is $55K (Assumptions B115); after opening it is $60K. The GM base is $45K, paid for 2 months before opening. The model assumes 4 paid owner months before opening (B49). If the build runs closer to 12 months, that line grows by about $5,100 per month.",
("Changes and flags","A49"):"Full members have 24/7 access; there are no hourly staff. The owner and GM are on site for appointments and set hours. Access control, cameras and a higher insurance line are added for unstaffed hours.",
("Changes and flags","A57"):"Owner salary, GM salary, no hourly staff",
("Changes and flags","A58"):"Owner salary is $60K after opening (not including equity or distributions) and the GM has a $45K base. Winston has no base. There are no hourly floor staff.",
("Changes and flags","A63"):"Without hourly staff, the owner and GM carry cleaning, front desk and open/close. Turning the hourly block on (B157 = 6) adds about $5,700 a month and raises Conservative capital by about $209K: the Sensitivity tab shows it. The owner is also unavailable for other work during on-site hours: check this against the owner's Chattanooga State schedule.",
("Changes and flags","A72"):"Guide, Investor returns, Sources and uses, Unit economics, Sensitivity, Ownership scenarios and Risks and checklist are new. The monthly sheets now run 60 months. Seasonality factors (Assumptions rows 135 to 149) are added and switched off. A $10,000 putting and chipping green is added to the build-out (Assumptions B151). An hourly staff block is added and switched off (Assumptions rows 153 to 158).",
("Guide","B79"):"Damage or theft cost in a 24/7 unstaffed building beyond the insurance line.",
("Risks and checklist","A6"):"24/7 unstaffed access",
("Risks and checklist","B6"):"Injury, theft, damage or misuse with nobody present. Insurers may price it high or decline.",
("Dues only","A2"):"Steady-state view. Revenue is dues only: no pro or strength-coach share, no junior packages, no add-ons. Costs are the model's fixed costs, card fees and contingency, including the owner's $60K and GM's $45K salaries + load.",
("Assumptions","E43"):"Estimate: raised from $800 because there is no staff to clean. Budget an outside cleaning service plus supplies and repairs. Get a quote",
("Assumptions","E7"):"$200/mo with limited-time access windows and no hourly staff. Packages are paid additions",
("Assumptions","A109"):"Juniors have limited-time access windows set by the owner. Training and teaching are by appointment only.",
("Assumptions","E108"):"Formula. Juniors use the facility only in supervised hours: waivers and insurance need to cover minors (see Changes and flags)",
("Risks and checklist","C7"):"Parent-signed waivers, limited-time windows when an adult is on site, ratios, background checks, insurer confirmation in writing.",
("Changes and flags","A32"):"Juniors under 18: waivers signed by a parent, background checks for anyone who coaches or supervises minors, and insurance that covers minors need to be confirmed with the insurer and attorney. With no hourly staff, minors may be on site with no staff present: insurers often treat that as a major issue.",
("Risks and checklist","C6"):"Entry control, cameras, waivers, an emergency procedure and call point, an AED, clear rules, and an insurance quote that names 24/7 access.",
})
MANUAL.update({
("Changes and flags","A15"):"GM base is $45K (Assumptions B33), paid from the first month of operation and 2 months before opening. The original $65K GM salary and Winston's $55K salary are replaced: Winston is a $0 base input and is 1099, paid only from his share of billings. The owner is the clinician and operator.",
("Summary","E35"):"Compare Year 1 with what each would need to commit. The contractor carries all ramp risk, since there is no base salary",
("Startup","A32"):"GM / golf pro base salary during build (0 if 1099)",
})

MANUAL.update({
("Guide","B17"):"Summary, Ownership, Investor returns, Sources and uses, Unit economics, Bay utilization and Dues only all read from those sheets.",
("Guide","B33"):"Month-by-month members, revenue, costs and profit for 60 months",
("Guide","B50"):"On Assumptions read column E. List every input marked 'Estimate' that can be replaced with a quote or a conversation.",
("Guide","B57"):"Read Unit economics sections 2 to 5: launch spend and selling time are in the cost to acquire a member, and value is capped at 12 and 24 months. Then read Bay utilization.",
("Guide","C57"):"Is the 24-month ratio above 5x? If so, ask whether CAC is understated. Then see how many members the two bays carry at peak versus the cap.",
("Guide","B71"):"Lifetime contribution of one member (capped at 12 or 24 months here), and the full cost to win one: marketing, launch spend and selling time.",
("Guide","B79"):"Damage or theft cost in a 24/7 unstaffed building beyond the insurance line.",
("Guide","B80"):"Class, event or private-rental bookings that take bays away from members.",
("Guide","B82"):"Founding-member deposits as usable cash before opening. Deposits are not being collected for now.",
("Changes and flags","A23"):"Wyatt: he currently keeps ~70% of lessons and earlier talked about a $38K to $40K base. This version pays the GM a $45K base (Assumptions B33) and still credits the 80% lesson split to the pro. Decide whether the GM is salary, split, or a blend, so the model does not double count.",
("Changes and flags","A51"):"Check that the 24/7 model works at the 200 cap: the Bay utilization tab shows how many members the two bays carry at peak hours. Gym floor and downtown parking are other likely limits. Use booking rules rather than headcount alone as the real limit.",
("Changes and flags","A72"):"Guide, Investor returns, Sources and uses, Unit economics, Bay utilization, Sensitivity, Ownership scenarios and Risks and checklist are new. Unit economics now counts launch spend and selling time in CAC and caps lifetime value at 12 and 24 months. The monthly sheets now run 60 months. Seasonality factors (Assumptions rows 135 to 149) are added and switched off. A $10,000 putting and chipping green is added to the build-out (Assumptions B151). An hourly staff block is added and switched off (Assumptions rows 153 to 158).",
("Risks and checklist","C4"):"Presell the 25 founding slots with signed commitments (no deposits taken for now). Collect letters of intent. Fund in stages. The raise covers the Base case with a cushion; the Conservative case needs more (Summary row 11).",
("Risks and checklist","B5"):"The owner is the clinician, TPI provider and operator, on site about 35 hours a week, with a salaried GM on the floor. If the owner is unavailable, premium service stops and the GM carries the building alone.",
("Risks and checklist","B8"):"Winston may be treated as an employee if the facility controls hours, methods and equipment. The GM is salaried and should be set up as an employee.",
("Risks and checklist","B21"):"Winston: scope, pay split, autonomy, non-solicit, insurance, IP. GM: employment terms, salary, duties, whether the lesson split also applies",
("Risks and checklist","B27"):"Membership billing, sales tax. Deposits are not being collected for now",
("Risks and checklist","B59"):"Signed commitments, interview notes, letters of intent",
("Risks and checklist","B57"):"Winston, and the GM employment terms",
("Risks and checklist","B60"):"Owner, Winston, the GM, advisors",
})
MANUAL.update({
("Sources and uses","A36"):"Staged amounts split the actual raise. The Conservative case needs more than the raise (see row 25), so a slow ramp would need that amount from another source. Staging lowers the investor's risk because money is released against proof. Triggers are suggestions to be agreed with counsel.",
("Sources and uses","A2"):"Two views of the same uses: sized to the Conservative case (a slow ramp) and to the Base case (a normal ramp). The raise is compared with both in rows 25 and 26.",
})
MANUAL.update({
("Assumptions","A128"):"Owner on-site time (before the part-time hire)",
("Changes and flags","A17"):"Floor staff: the owner covers the floor first. A part-time hire (Assumptions rows 160 to 172) starts when residual profit after the investor's distribution covers the hire cost with a cushion for several months in a row; owner on-site hours then step down. A full-coverage hourly block (rows 153 to 158) is built and switched off (B157 = 0).",
("Changes and flags","A57"):"Owner salary, GM salary, phased part-time help",
("Changes and flags","A58"):"Owner salary is $60K after opening for about 40 hours a week (not including equity or distributions) and the GM has a $45K base. Winston has no base. Part-time floor help is added by the trigger in Assumptions rows 160 to 172.",
("Changes and flags","A63"):"At first the owner and GM carry cleaning, front desk and open/close (about 35 owner on-site hours a week). Part-time help starts when profit after the investor's distribution covers 1.5x its cost for 3 months in a row, then owner on-site hours step down to about 20. Monthly sheet rows 56 to 63 show the test; Summary rows 47 to 50 show when it fires in each case. The Sensitivity tab shows no hire at all and a 30-hour hire.",
("Changes and flags","A49"):"Full members have 24/7 access; the owner and GM are on site for appointments and set hours at first, with part-time help added later. Access control, cameras and a higher insurance line are added for unstaffed hours.",
("Risks and checklist","B5"):"The owner is the clinician, TPI provider and operator: about 40 working hours a week, about 35 on site at first. A salaried GM covers the floor. If the owner is unavailable, premium service stops and the GM carries the building alone.",
("Guide","C33"):"Column B (month 1) read downward. Rows 56 to 63 show when part-time help starts",
("Summary","E28"):"Surviving on about 50 members was an owner target; change this to test other counts",
("Assumptions","E43"):"Estimate: raised from $800 because staff are not cleaning at first. Budget an outside cleaning service plus supplies and repairs. Get a quote",
})

MANUAL.update({
("Changes and flags","A45"):"Owner pay during the build is $55K a year (Assumptions B115); after opening it is $60K. The GM base is $45K, paid for 2 months before opening. The model assumes 6 paid owner months before opening (B49), a typical 6 to 9 month tenant build-out. Each extra month adds about $5,100. Rent before opening is not modeled beyond the first month and deposit: negotiate a rent-free build period.",
("Changes and flags","A72"):"Guide, Investor returns, Sources and uses, Unit economics, Bay utilization, Sensitivity, Ownership scenarios and Risks and checklist are new. Unit economics counts launch spend and selling time in CAC and caps lifetime value at 12 and 24 months. The monthly sheets run 60 months. Seasonality factors (Assumptions rows 135 to 149) are added and switched off. The $10,000 putting and chipping green was added and then removed (Assumptions B151 = 0). Added and switched off or optional: full hourly staff (rows 153 to 158), tiered pricing (rows 174 to 183), a 12-month renewal-rate view (row 24 and B186). Added and on: a phased part-time hire (rows 160 to 172) and strength coach follow-up pay (rows 188 to 191).",
("Guide","C33"):"Column B (month 1) read downward. Rows 56 to 67 show the part-time trigger, tiered pricing and the renewal view",
("Assumptions","E90"):"Estimate. The strength coach delivers follow-ups (rows 188 to 191); routine builds remain owner time",
})

MANUAL.update({
("Assumptions","A33"):"GM base salary (Wyatt, GM and head golf pro) ($/yr)",
("Assumptions","E33"):"Owner plan: Wyatt is the GM on a $45K salary and also keeps 80% of his lesson revenue (rows 76 to 79). This is intended, not a double count. A salaried GM with set duties is likely an employee: have a CPA or attorney confirm how the lesson work is structured",
("Assumptions","E73"):"80/20 fee split: Winston keeps 80% and the facility keeps 20%, covering use of the space. Non-member athletes are on rows 198 to 203",
("Changes and flags","A15"):"Wyatt is the GM on a $45K base (Assumptions B33), paid from the first month of operation and 2 months before opening, and he also keeps 80% of his lesson revenue. Winston is a $0 base input and is paid from his share of billings through his own independent business. The owner is the clinician and operator.",
("Changes and flags","A23"):"Wyatt currently keeps ~70% of lessons and earlier talked about a $38K to $40K base. This version pays him a $45K GM salary and the 80% lesson split. Check the total package, and have a CPA or attorney confirm how the salaried GM role and the lesson work are structured.",
("Changes and flags","A44"):"Junior package prices ($45 S&C, $75 teaching) are still group-rate placeholders. Confirm that Winston and Wyatt will teach juniors at those rates, or raise them.",
("Changes and flags","A72"):"Guide, Investor returns, Sources and uses, Unit economics, Bay utilization, Equipment list, Sensitivity, Ownership scenarios, Founding members, Scripts and Risks and checklist are new. Unit economics counts launch spend and selling time in CAC and caps lifetime value at 12 and 24 months. The monthly sheets run 60 months. Seasonality factors (Assumptions rows 135 to 149) are added and switched off. The putting and chipping green was removed (B151 = 0). The gym floor and storage are now itemized on Equipment list at a $38K planning number. Added and switched off or optional: full hourly staff (rows 153 to 158), tiered pricing (rows 174 to 183), a 12-month renewal-rate view (row 24 and B186), pre-opening rent and a landlord allowance (rows 194 and 195), and non-member athletes for Winston's business (rows 198 to 203). Added and on: a phased part-time hire (rows 160 to 172) and strength coach follow-up pay (rows 188 to 191).",
("Risks and checklist","B8"):"Winston may be treated as an employee if the facility controls hours, methods and equipment; his own independent business and other clients help. Wyatt is salaried as GM, which is likely an employee role.",
("Risks and checklist","B21"):"Winston: scope, 80/20 fee split, use of the space, non-member athletes, ownership of programs and testing tools, equity terms. Wyatt: GM employment terms, salary, duties, and how the lesson revenue is handled",
("Risks and checklist","B57"):"Winston and Wyatt agreements",
("Risks and checklist","B60"):"Owner, Winston, Wyatt, advisors",
("Sources and uses","E11"):"Startup rows 11 to 14, itemized on Equipment list. Planning number about $38K including storage",
})
MANUAL.update({("Guide","B53"):"On Assumptions read column E. List every input marked 'Estimate' that can be replaced with a quote or a conversation."})
MANUAL.update({
("Sources and uses","E6"):"Equipment list tab. Planning number about $38K including custom storage",
("Sources and uses","E11"):"Startup rows 25 to 27: deposit, first month and any pre-opening rent months",
("Sources and uses","A36"):"Staged amounts split the actual raise. A slow ramp (Conservative) is covered by the reserve if row 25 is positive; whatever the ramp does not use above the capital needed is returned to the investor. Triggers are suggestions to be agreed with counsel.",
("Assumptions","A193"):"Lease terms that lower the capital required",
("Changes and flags","D5"):"Start-up + ramp losses + cushion. Compare to the $500K raise",
("Guide","B70"):"Raise divided by the equity offered: $500K for 30% implies about $1.67M.",
})
