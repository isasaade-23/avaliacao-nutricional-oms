# -*- coding: utf-8 -*-
"""Valida zscore_ref.py contra who_zscore.py (validado vs R) e asserts oficiais."""
import sys, os
sys.path.insert(0, r"C:\Users\Renato\OneDrive\Documentos\anthroplus\who-zscore-python")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from who_zscore import calc_who_zscore           # 5-19, validado vs R anthroplus
from zscore_ref import _lookup, _z, _z_extremo


def zbmi(X, meses, sexo):
    lms = _lookup("imc_idade_5a19", sexo, meses)
    if not lms: return None
    z = _z(X, *lms); return round(_z_extremo(X, *lms, z), 2)


def zhfa(X, meses, sexo):
    lms = _lookup("est_idade_5a19", sexo, meses)
    if not lms: return None
    return round(_z(X, *lms), 2)


def zwa_d(X, dias, sexo):
    lms = _lookup("peso_idade_0a5", sexo, dias)
    if not lms: return None
    z = _z(X, *lms); return round(_z_extremo(X, *lms, z), 2)


# ---- 1) Paridade com who_zscore.py (5-19), varrendo idades/sexos/medidas ----
print("1) Paridade ref vs who_zscore.py (5-19):")
ndiff = ntot = 0
for sexo in (1, 2):
    for meses in range(61, 229, 1):
        for X in (14, 16, 18, 22, 28, 35):           # IMC
            ntot += 1
            a = zbmi(X, meses, sexo); b = calc_who_zscore(X, meses, sexo, "bfa")
            if a is None or abs(a - b) > 0.011:
                ndiff += 1
                if ndiff <= 5: print(f"   DIFF bfa sexo={sexo} m={meses} X={X}: ref={a} who={b}")
        for H in (110, 130, 150, 170):               # estatura
            ntot += 1
            a = zhfa(H, meses, sexo); b = calc_who_zscore(H, meses, sexo, "hfa")
            if a is None or abs(a - b) > 0.011:
                ndiff += 1
                if ndiff <= 5: print(f"   DIFF hfa sexo={sexo} m={meses} H={H}: ref={a} who={b}")
print(f"   -> {ntot} casos, {ndiff} divergencias\n")

# ---- 2) Assert oficial anthro (0-5, peso/idade) ----
print("2) Assert oficial anthro (peso/idade 0-5):")
got = zwa_d(17, 1522, 2)
print(f"   peso=17kg, 1522 dias, menina -> z={got}  (esperado 0.24)  {'OK' if abs(got-0.24)<0.011 else 'ERRO'}")

# ---- 3) Conferencia tabela estatura dia 44 menino (test anthro usa M=56.4833?) ----
lms = _lookup("est_idade_0a5", 1, 44)
print(f"\n3) est_idade_0a5 sexo=1 dia=44: L={lms[0]} M={lms[1]} S={lms[2]}")
print(f"   z(50cm) = {round(_z(50,*lms),4)}")
