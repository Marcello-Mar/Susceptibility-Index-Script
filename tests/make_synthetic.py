"""Generate the synthetic demo/validation dataset. Not real patient data."""
import numpy as np, pandas as pd, sys
rng = np.random.default_rng(20261122)
orgs = [  # name, weight, gram
 ("Escherichia coli",22,'n'),("Klebsiella pneumoniae",9,'n'),("Pseudomonas aeruginosa",8,'n'),("Enterobacter cloacae",4,'n'),
 ("Klebsiella oxytoca",2,'n'),("Proteus mirabilis",2,'n'),("Serratia marcescens",2,'n'),("Stenotrophomonas maltophilia",2,'n'),("Acinetobacter baumannii",1,'n'),
 ("Staphylococcus aureus",14,'p'),("Staphylococcus epidermidis",12,'p'),("Staphylococcus hominis",3,'p'),("Enterococcus faecalis",6,'p'),("Enterococcus faecium",4,'p'),
 ("Streptococcus mitis",3,'p'),("Streptococcus pneumoniae",1,'p'),("Micrococcus luteus",1,'p'),("Candida albicans",1,'f')]
# base susceptibility per (gram, agent); None = rarely tested in that group
agents = {
 "Ampicillin":      {'n':0.45,'p':0.55},
 "Amoxicillin/clavulanate": {'n':0.70,'p':0.60},
 "Piperacillin/tazobactam": {'n':0.82,'p':None},
 "Oxacillin":       {'n':None,'p':0.45},
 "Cefotaxime":      {'n':0.80,'p':None},
 "Ceftazidime":     {'n':0.78,'p':None},
 "Cefepime":        {'n':0.84,'p':None},
 "Meropenem":       {'n':0.95,'p':None},
 "Gentamicin":      {'n':0.88,'p':0.75},
 "Amikacin":        {'n':0.95,'p':None},
 "Ciprofloxacin":   {'n':0.76,'p':0.70},
 "Vancomycin":      {'n':None,'p':0.93},
 "Teicoplanin":     {'n':None,'p':0.88},
 "Daptomycin":      {'n':None,'p':0.97},
 "Linezolid":       {'n':None,'p':0.99},
 "Clindamycin":     {'n':None,'p':0.62},
 "Trimethoprim/sulfamethoxazole": {'n':0.68,'p':0.80},
 "Colistin":        {'n':0.93,'p':None},
}
# EUCAST expected resistant phenotypes (reported R when tested) and pairs labs do not report
INTRINSIC_R = set()
for org, drugs in {
    "Pseudomonas aeruginosa": ["Ampicillin", "Amoxicillin/clavulanate", "Cefotaxime", "Trimethoprim/sulfamethoxazole"],
    "Stenotrophomonas maltophilia": ["Ampicillin", "Amoxicillin/clavulanate", "Piperacillin/tazobactam", "Cefotaxime", "Meropenem", "Gentamicin", "Amikacin"],
    "Klebsiella pneumoniae": ["Ampicillin"], "Klebsiella oxytoca": ["Ampicillin"],
    "Enterobacter cloacae": ["Ampicillin", "Amoxicillin/clavulanate"],
    "Serratia marcescens": ["Ampicillin", "Amoxicillin/clavulanate", "Colistin"],
    "Proteus mirabilis": ["Colistin"],
    "Acinetobacter baumannii": ["Ampicillin", "Amoxicillin/clavulanate", "Cefotaxime"],
    "Enterococcus faecalis": ["Clindamycin"], "Enterococcus faecium": ["Clindamycin"],
}.items():
    INTRINSIC_R |= {(org, d) for d in drugs}
NOT_TESTED = {(o, "Oxacillin") for o in ("Enterococcus faecalis", "Enterococcus faecium", "Streptococcus mitis", "Streptococcus pneumoniae")}
NOT_TESTED |= {(o, "Trimethoprim/sulfamethoxazole") for o in ("Enterococcus faecalis", "Enterococcus faecium")}
NOT_TESTED |= {("Stenotrophomonas maltophilia", d) for d in ("Ceftazidime", "Cefepime", "Colistin")}

N = 640
w = np.array([o[1] for o in orgs], float); w /= w.sum()
rows = []
pid = 0
for i in range(N):
    o = orgs[rng.choice(len(orgs), p=w)]
    if rng.random() < 0.18 and rows:
        prev = rows[rng.integers(len(rows))]
        patient = prev["Patient ID"]
        if rng.random() < 0.6: o = next(x for x in orgs if x[0]==prev["Microorganism"])
    else:
        pid += 1; patient = f"P{pid:04d}"
    day = pd.Timestamp("2024-01-01") + pd.Timedelta(days=int(rng.integers(0, 730)))
    row = {"Patient ID": patient, "Collection date": day.date().isoformat(), "Microorganism": o[0]}
    g = o[2]
    # organism-specific resistance shifts
    shift = {"Enterococcus faecium":-0.35,"Staphylococcus epidermidis":-0.2,"Pseudomonas aeruginosa":-0.05,"Klebsiella pneumoniae":-0.1,"Stenotrophomonas maltophilia":-0.5}.get(o[0],0)
    # correlated resistance phenotype (e.g. ESBL / MRSA-like)
    mdr = rng.random() < (0.22 if g=='n' else 0.35)
    for a, base in agents.items():
        p = base.get(g) if g in base else None
        if g == 'f': row[a] = None; continue
        if p is None or (o[0], a) in NOT_TESTED: row[a] = None; continue
        if (o[0], a) in INTRINSIC_R:
            row[a] = None if rng.random() < 0.5 else "R"; continue
        if rng.random() < 0.12: row[a] = None; continue
        pp = min(0.99, max(0.02, p + shift - (0.35 if mdr and a not in ("Meropenem","Amikacin","Vancomycin","Linezolid","Daptomycin","Colistin") else 0)))
        u = rng.random()
        row[a] = "S" if u < pp else ("I" if u < pp + 0.03 else "R")
    rows.append(row)
df = pd.DataFrame(rows)
# a few messy cells to exercise normalisation
df.loc[5, "Gentamicin"] = " s "; df.loc[17, "Ciprofloxacin"] = "nt"; df.loc[40, "Ampicillin"] = "r"
out = sys.argv[1] if len(sys.argv) > 1 else "synthetic.xlsx"
df.to_excel(out, index=False)
df.to_csv(out.replace(".xlsx", ".csv"), index=False)
print(df.shape, df["Microorganism"].value_counts().head())
