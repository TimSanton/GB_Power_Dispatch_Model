"""
Hand-checkable cases for the coupled dispatch.

    python test_dispatch.py

No network access — exercises dispatch() against a small synthetic stack.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import run_model as m

STACK = [("Wind", None, 0.0, "renewable"),
         ("Nuclear", None, 4.0, "nuclear"),
         ("CCGT A", 10.0, 70.0, "gas"),
         ("CCGT B", 10.0, 90.0, "gas")]

def run(demand, wind, nuc, links):
    return m.dispatch(demand, STACK, wind, 0.0, 0.0, nuc, 0.0, links)

fails = []
def check(label, got, want):
    ok = got == want
    if not ok:
        fails.append(label)
    print(f"  {'PASS' if ok else 'FAIL'}  {label}\n        got {got}  want {want}")

print("1. No links — pure domestic clearing")
check("demand 12, wind 5, nuc 3 -> CCGT A at 70",
      run(12, 5, 3, []), ("CCGT A", 70.0, 0.0))

print("\n2. Cheap link undercuts gas -> full import, link marginal only if it clears")
# demand 12, wind 5, nuc 3 => 4 residual. Link 2GW at 50 is cheaper than CCGT A.
# supply at 50: 5+3+2=10 < 12. supply at 70: +10 => 20 >= 12. CCGT A marginal.
check("2GW @ 50 -> imports 2, CCGT A marginal",
      run(12, 5, 3, [("L", 2.0, 50.0)]), ("CCGT A", 70.0, 2.0))

print("\n3. Dear link -> GB exports at full capacity, price rises")
# demand 12, wind 5, nuc 3. Link 2GW at 150 (dearer than all GB gas).
# at p=70: supply 5+3+10=18 ; demand 12+2(export)=14 -> clears at 70, exports 2
check("2GW @ 150 -> exports 2, CCGT A marginal",
      run(12, 5, 3, [("L", 2.0, 150.0)]), ("CCGT A", 70.0, -2.0))

print("\n4. Export pushes GB into the dearer tranche")
# demand 18, wind 5, nuc 3 => domestic cheap = 8. CCGT A 10 => 18 at p=70.
# with 2GW export, demand at 70 = 20 > 18 supply -> must go to CCGT B at 90.
check("demand 18 + 2GW export -> CCGT B at 90",
      run(18, 5, 3, [("L", 2.0, 150.0)]), ("CCGT B", 90.0, -2.0))
print("        (same case with no link stays on CCGT A:)")
check("demand 18, no link -> CCGT A at 70",
      run(18, 5, 3, []), ("CCGT A", 70.0, 0.0))

print("\n5. Coupling: GB price converges onto the link price, link partly loaded")
# demand 19, wind 5, nuc 3, CCGT A 10 => 18 domestic below 80.
# link 5GW at 80. at p=70: supply 18, demand 19+5=24 -> no.
# at p=80: supply 18+5=23 >= demand 19 -> clears at 80, link marginal, partial.
check("link @ 80 becomes marginal, imports 1 of 5GW",
      run(19, 5, 3, [("L", 5.0, 80.0)]), ("L", 80.0, 1.0))

print("\n6. Mixed: cheap link imports while dear link exports simultaneously")
# demand 12, wind 5, nuc 3. cheap 2GW@50, dear 3GW@150
# at p=70: supply 5+3+2+10=20 ; demand 12+3=15 -> clears 70. net = 2-3 = -1
check("import 2, export 3 -> net -1",
      run(12, 5, 3, [("C", 2.0, 50.0), ("D", 3.0, 150.0)]),
      ("CCGT A", 70.0, -1.0))

print("\n7. Negative continental price is taken before free wind")
# demand 12, wind 5, nuc 3, link 2GW @ -5
# at p=-5: supply 2 < 12. at 0: 2+5=7. at 4: +3=10. at 70: 20>=12 -> CCGT A
check("link @ -5 imports fully",
      run(12, 5, 3, [("L", 2.0, -5.0)]), ("CCGT A", 70.0, 2.0))

print("\n8. Surplus renewables set the price at zero, no unserved")
check("wind 30 > demand 12 -> Wind marginal at 0",
      run(12, 30, 3, []), ("Wind", 0.0, 0.0))

print("\n9. Unserved when demand exceeds everything")
check("demand 100 -> Unserved",
      run(100, 5, 3, []), ("Unserved", 0.0, 0.0))

print("\n10. Monotonicity: supply-demand crosses zero once (random probe)")
import random
random.seed(7)
bad = 0
for _ in range(3000):
    d = random.uniform(0, 40)
    ls = [(f"L{i}", random.uniform(0.1, 4), random.uniform(-10, 200))
          for i in range(random.randint(0, 4))]
    w, n = random.uniform(0, 15), random.uniform(0, 6)
    units = [(nm, c, mc, t) for nm, c, mc, t in
             [("Wind", w, 0.0, "r"), ("Nuclear", n, 4.0, "n"),
              ("CCGT A", 10.0, 70.0, "g"), ("CCGT B", 10.0, 90.0, "g")] if c > 0]
    cands = sorted({mc for _, _, mc, _ in units} | {p for _, _, p in ls})
    sign = [(sum(c for _, c, mc, _ in units if mc <= p)
             + sum(c for _, c, pc in ls if pc <= p))
            - (d + sum(c for _, c, pc in ls if pc > p)) for p in cands]
    # must be non-decreasing
    if any(sign[i] > sign[i+1] + 1e-9 for i in range(len(sign)-1)):
        bad += 1
print(f"  {'PASS' if bad == 0 else 'FAIL'}  balance non-decreasing in price "
      f"in {3000-bad}/3000 random cases")
if bad: fails.append("monotonicity")

print("\n" + ("ALL PASS" if not fails else f"FAILURES: {fails}"))
sys.exit(1 if fails else 0)
