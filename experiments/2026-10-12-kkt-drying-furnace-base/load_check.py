import math
g = 9.81
# H-250x125x6x9 (KS) approx properties
Z = 324e3      # mm3 (Zx)
I = 4050e4     # mm4
Aw = 250*6     # mm2 web area (approx)
tw, tf, d = 6, 9, 250
Fy, E = 275.0, 205000.0
phi = 0.9
FAC = 1.5      # 1.2 dead x ~1.25 handling/placing impact
phiMn = phi*Fy*Z/1e6   # kNm
print("phi*Mn per rail = %.1f kNm" % phiMn)

def unit(name, tons, Lm, n_rails_eff, mode):
    W = tons*g                       # kN
    per = W/n_rails_eff              # kN per effective rail
    if mode == "uniform":
        w = per/Lm
        span = 2.25
        M = w*span**2/10             # continuous beam approx
    else:                            # concentrated on 2 skids crossing the rails
        P = per/2                    # each skid
        span = 2.25
        M = P*span/4                 # midspan point load, simple span (conservative)
    Mu = M*FAC
    return per, M, Mu, Mu/phiMn

L45, L38 = 4.7, 4.7
cases = [
 ("A 균등, 4줄 분담, 45t", 45, L45, 4, "uniform"),
 ("B 스키드 2개, 4줄 분담, 45t", 45, L45, 4, "point"),
 ("C 스키드 2개, 2줄만 분담, 45t", 45, L45, 2, "point"),
 ("D 스키드 2개, 3줄 분담, 45t", 45, L45, 3, "point"),
]
for c in cases:
    per, M, Mu, u = unit(*c)
    print("%-28s per-rail %.0f kN  M=%.1f  Mu=%.1f kNm  util=%.2f" % (c[0], per, M, Mu, u))

# reaction / web crippling
R = 1.1*(45*g/4/L45)*2.25*FAC
N = 125
Rn = 0.8*tw**2*(1+3*(N/d)*(tw/tf)**1.5)*math.sqrt(E*Fy*tf/tw)/1e3
print("column reaction (A, factored) %.0f kN ; web crippling phi*Rn = %.0f kN" % (R, 0.75*Rn))
# column
Ac = 3031
print("column phi*Fy*A (no buckling) = %.0f kN" % (0.9*Fy*Ac/1e3))
# shear
V = 0.6*(45*g/4/L45)*2.25*FAC
print("shear V=%.0f kN ; phi*Vn=%.0f kN" % (V, 0.9*0.6*Fy*Aw/1e3))
# deflection, case B per rail
P = 45*g/4/2
dl = P*1e3*2250**3/(48*E*I)
print("defl (B) %.2f mm  (L/%.0f)" % (dl, 2250/dl))
# ground
contact = 5.57
tot = 90 + 3.2
print("ground avg = %.0f kPa (%.1f tf/m2) incl. self weight" % (tot*g/contact, tot/contact))
