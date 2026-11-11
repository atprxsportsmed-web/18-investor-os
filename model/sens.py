import subprocess, json, time, sys
from openpyxl import load_workbook
def scale(A,cells,k):
    for c in cells: A[c]=A[c].value*k
def mix_to_base(A):  # premium to 0 keep mix 100%
    A["B10"]=A["B10"].value+A["B11"].value; A["B11"]=0; A["B125"]=0
def half_premium(A):
    A["B10"]=A["B10"].value+A["B11"].value*0.5; A["B11"]=A["B11"].value*0.5; A["B125"]=A["B125"].value*0.5
cases=[
 ("As built",lambda A:None),
 ("Sign-ups 25% slower",lambda A:scale(A,["B22","C22","D22"],0.75)),
 ("Quarterly churn doubled",lambda A:scale(A,["B24","C24","D24"],2)),
 ("All-in price $450 (sign-ups 12% faster)",lambda A:(A.__setitem__("B5",450),A.__setitem__("B6",450),A.__setitem__("B9",399),scale(A,["B22","C22","D22"],1.125))),
 ("All-in price $600 (sign-ups 19% slower)",lambda A:(A.__setitem__("B5",600),A.__setitem__("B6",600),A.__setitem__("B9",529),scale(A,["B22","C22","D22"],0.8125))),
 ("All dues 10% lower",lambda A:scale(A,["B5","B6","B7","B124","B9"],0.9)),
 ("Follow-up coaching paid at $60/hour",lambda A:A.__setitem__("B190",60)),
 ("Rent 50% higher ($18/sq ft)",lambda A:A.__setitem__("B27",18)),
 ("Build-out cost doubled ($40/sq ft)",lambda A:A.__setitem__("B29",40)),
 ("Insurance doubled ($370/month)",lambda A:A.__setitem__("B40",370)),
 ("Add hourly floor staff (one 7-hour shift, 7 days)",lambda A:A.__setitem__("B157",7)),
 ("No part-time hire (owner covers the floor for all 60 months)",lambda A:A.__setitem__("B161",0)),
 ("Part-time hire works 30 hours a week",lambda A:A.__setitem__("B162",30)),
 ("Tiered pricing on ($400 / $450 / $500, slower sign-ups at each step)",lambda A:A.__setitem__("B175",1)),
 ("30% of members on longer terms (15% 6-month, 15% 12-month)",lambda A:(A.__setitem__("B217",0.15),A.__setitem__("B218",0.15))),
 ("Churn at 3% per quarter in the Base case (lower than planned)",lambda A:A.__setitem__("C24",0.03)),
 ("Owner delivers the follow-ups (no strength coach pay)",lambda A:A.__setitem__("B189",0)),
 ("Rent starts at lease signing (6 months paid before opening)",lambda A:A.__setitem__("B194",6)),
 ("Landlord allowance of $20K toward the build-out",lambda A:A.__setitem__("B195",20000)),
("Larger space: 4,800 sq ft (rent and build-out scale with size)",lambda A:A.__setitem__("B26",4800)),
 ("Follow-up coaching paid at $120/hour (80% of $150)",lambda A:A.__setitem__("B190",120)),
 ("GM salary $60K instead of $45K",lambda A:A.__setitem__("B33",60000)),
 ("No seasonality: flat sign-ups all year",lambda A:[A.__setitem__(f"B{137+k}",1.0) for k in range(12)]),
 ("Winston also needs a $40K base salary",lambda A:A.__setitem__("B34",40000)),
 ("Stress: sign-ups 25% slower, churn 1.5x",lambda A:(scale(A,["B22","C22","D22"],0.75),scale(A,["B24","C24","D24"],1.5))),
]
out=[]
for i,(name,fn) in enumerate(cases):
    t=time.time()
    wb=load_workbook("stage1.xlsx"); fn(wb["Assumptions"])
    p=f"sens/c{i}.xlsx"; wb.save(p)
    res=json.loads(subprocess.run(["python3","/mnt/skills/public/xlsx/scripts/recalc.py",p,"120"],capture_output=True,text=True).stdout)
    v=load_workbook(p,data_only=True); S=v["Summary"]; O=v["Ownership"]
    rec=dict(name=name,errors=res.get("total_errors"),
      base_cap=S["C9"].value,base_surplus=S["C11"].value,base_first=S["C16"].value,base_y3=S["C23"].value,base_m36=S["C15"].value,
      cons_cap=S["B9"].value,cons_first=S["B16"].value,cons_y3=S["B23"].value,base_pay=O["C19"].value)
    out.append(rec); print(i,name,round(time.time()-t,1),"s",rec["errors"],flush=True)
json.dump(out,open("sens_results.json","w"))
