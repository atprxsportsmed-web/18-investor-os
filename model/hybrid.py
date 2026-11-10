import openpyxl,re
from copy import copy
wb=openpyxl.load_workbook("stage3b.xlsx")
# late edits carried over from the 500K/75% step
o=wb["Ownership scenarios"]
o["A79"]="Negotiating frame: anchor at 30% with priority payback and return of unused capital, ceiling about 40%, and keep at least 51% as owner. Protective rights (vetoes on a sale, new debt or owner pay) matter as much as the percentage. Securities counsel should review all terms."
r=wb["Investor returns"]
r["A28"]='="Base case payback: structure B (1.0x) still owes "&TEXT(E68,"$#,##0")&" at the end of Year 3 and "&TEXT(F68,"$#,##0")&" at the end of Year 4. Structure C (1.5x) still owes "&TEXT(G75,"$#,##0")&" at the end of Year 5."'
r["A28"].font=copy(r["A27"].font); r["A28"].alignment=copy(r["A27"].alignment)
# numbers
b=wb["Bay utilization"]; b["B6"]=17.5
b["H6"]="Estimate: 4:30 am to 10 pm. Members enter with a key fob in that window; the building is staffed 11 am to 6 pm and closed overnight"
WIN="4:30 am to 10 pm"
SUB=[
("open 24/7 to members","open "+WIN+" to members, staffed 11 am to 6 pm"),
("24/7 unstaffed building","building that is unstaffed outside 11 am to 6 pm"),
("Includes 24/7 access","Includes access from "+WIN+" (staffed 11 am to 6 pm)"),
("raised from $500 for 24/7 access and minors","raised from $500 for extended unstaffed hours ("+WIN+") and minors"),
("Two TrackMan bays, 24/7 access","Two TrackMan bays, access "+WIN),
("during 24/7 access","during unstaffed hours"),
("24/7 unstaffed access","Extended unstaffed access"),
("names 24/7 access","names the "+WIN+" access window and the 11 am to 6 pm staffed window"),
("use clause for 24/7 access","use clause for "+WIN+" access"),
("24/7 access, gym","access "+WIN+", gym"),
("24/7 access and minors named","Extended hours and minors named"),
("General and professional liability, 24/7 access","General and professional liability, extended hours"),
("24/7 access","access "+WIN),
("24/7 model","extended-hours model"),
]
EX={
("Changes and flags","A2"):"Same structure and formulas as the original. Changed: GM base salary, optional hourly floor coverage (off), contractor logic, pricing and tiers, hybrid access (4:30 am to 10 pm, staffed 11 am to 6 pm), two TrackMan bays, 200 cap.",
("Changes and flags","A49"):"Hybrid access: members enter 4:30 am to 10 pm with a key fob; the building is staffed 11 am to 6 pm and closed overnight. The owner and GM cover the staffed window at first (about 49 hours a week), with part-time help added later. Access control, cameras and a higher insurance line cover the unstaffed hours.",
("Changes and flags","A63"):"Without hourly staff, the owner and GM carry cleaning, front desk and the staffed window. Turning the hourly block on (B157 = 7, one 7-hour shift from 11 am to 6 pm) adds about $3,700 a month and raises Conservative capital by about $129K: the Sensitivity tab shows it. The owner is also unavailable for other work during on-site hours: check this against the owner's Chattanooga State schedule.",
("Changes and flags","A17"):"Floor staff: the owner and GM cover the 11 am to 6 pm staffed window first. A part-time hire (Assumptions rows 160 to 172) starts when residual profit after the investor's distribution covers the hire cost with a cushion for several months in a row; owner on-site hours then step down. A full-coverage hourly block (rows 153 to 158: one 7-hour shift, 7 days) is built and switched off (B157 = 0).",
("Assumptions","E64"):"Formula from the hourly coverage block (B158): one 7-hour shift covering the 11 am to 6 pm window",
("Assumptions","A154"):"Shifts per staffed day", ("Assumptions","A156"):"Crossover between shifts, cleaning handoff (hours)",
("Assumptions","E157"):"Estimate: set to 7 for the 11 am to 6 pm window every day. Default 0: the owner and GM cover the window",
}
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            v=c.value
            if isinstance(v,str) and not v.startswith("="):
                for a,bb in SUB: v=v.replace(a,bb)
                c.value=v
for (s,k),t in EX.items(): wb[s][k]=t
rk=wb["Risks and checklist"]
rk["B6"]=rk["B6"].value.rstrip()+" The staffed window is 11 am to 6 pm, about 49 hours a week, so the owner and GM must cover it even when one is teaching or seeing a client."
rk["C6"]=rk["C6"].value.rstrip()+" Closed 10 pm to 4:30 am. Part-time help covers the window when profit supports it."
wb.save("stage3b.xlsx")
