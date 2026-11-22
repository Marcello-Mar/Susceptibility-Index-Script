"""Compare the JS engine with the v1 Python logic on the synthetic dataset (All isolates)."""
import json, subprocess, sys, types
import pandas as pd, numpy as np
nb = json.load(open(sys.argv[1]))
src = "".join(nb["cells"][0]["source"]).replace("from google.colab import files", "files=None")
src = src.split("def main():")[0]
ns = {}; exec(src, ns)
df = pd.read_excel("data/synthetic_demo.xlsx")
df = df.drop(columns=["Patient ID", "Collection date"])
df = df.replace({"nt": np.nan, " s ": "S", "r": "R"})
df.to_excel("/tmp/v1_input.xlsx", index=False)
pd.DataFrame({"Microorganism": [], "Gram_stain": []}).to_excel("/tmp/gram.xlsx", index=False)
A = ns["AntibioticAnalyzer"]
A._get_gram_stain = lambda self, o: "other"
results = {}
for conf in (0.95, 0.90):
  for thr in (0, 50):
    an = A("/tmp/v1_input.xlsx", "/tmp/gram.xlsx", None, confidence_level=conf)
    res = an.analyze(thr)["All Isolates"]
    results[f"{conf}-{thr}"] = {
      "singles": {r.name: [r.s_percentage, r.lower_ci, r.upper_ci, r.global_effectiveness, r.lower_ci_global, r.upper_ci_global] for r in res["Singles"]},
      "combos": {f"{c.combination[0]}|{c.combination[1]}": [c.effectiveness, c.lower_ci, c.upper_ci, c.global_effectiveness, c.lower_ci_global, c.upper_ci_global] for c in res["Combinations"]},
    }
js = subprocess.run(["node", "tests/run_engine.js", "/tmp/v1_input.xlsx"], capture_output=True, text=True, check=True)
jsr = json.loads(js.stdout)
worst = 0; count = 0
for key, ref in results.items():
  got = jsr[key]
  assert set(ref["singles"]) == set(got["singles"]), (key, set(ref["singles"]) ^ set(got["singles"]))
  assert set(ref["combos"]) == set(got["combos"]), key
  for kind in ("singles", "combos"):
    for name, vals in ref[kind].items():
      for a, b in zip(vals, got[kind][name]):
        worst = max(worst, abs(a - b)); count += 1
print(f"compared {count} values across {len(results)} parameter sets; max abs diff = {worst:.2e} percentage points")
assert worst < 1e-6
