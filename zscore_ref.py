# -*- coding: utf-8 -*-
"""
Implementacao de REFERENCIA (Python) dos escores-z OMS 0-19 anos.

Esta logica e' a "fonte unica da verdade": as formulas do Excel em
build_planilha.py replicam EXATAMENTE o que esta' aqui. Validamos este modulo
e depois conferimos Excel == este modulo.

Regras (espelham WHO Anthro / AnthroPlus):
- Idade: meses completos (escolhe tabela) e dias (chave 0-5).
- 0-5 anos (WHO 2006): chave = idade em DIAS (0-1826). Para 1827-1856 dias
  (ainda "60 meses") usa-se a ultima linha (1826) -> equivalente ao CORRESP
  aproximado (maior chave <= valor) que o Excel faz.
- 5-19 anos (WHO 2007): chave = idade em MESES (61-228). Peso/idade so' 61-120.
- Ajuste deitado/em pe (+-0,7 cm) para estatura/idade e peso/estatura.
- Ajuste de escore-z extremo da OMS (|z|>3) para indices baseados em PESO
  (IMC/idade, peso/idade, peso/estatura) -- NAO para estatura/idade.

sexo: 1=masculino, 2=feminino
"""
import os
import math
import bisect
import pandas as pd

DIR = os.path.dirname(os.path.abspath(__file__))
TAB = os.path.join(DIR, "tabelas_oms")

# ---- carga das tabelas como listas de chaves k ordenadas + dict k->(l,m,s) ----
_TABELAS = {}


def _load(name):
    if name in _TABELAS:
        return _TABELAS[name]
    df = pd.read_csv(os.path.join(TAB, name + ".csv"))
    ks = df["k"].tolist()
    lms = {int(r.k): (float(r.l), float(r.m), float(r.s)) for r in df.itertuples()}
    chave_min = {1: None, 2: None}
    chave_max = {1: None, 2: None}
    for s in (1, 2):
        sub = df[df.sexo == s]["chave"]
        if len(sub):
            chave_min[s] = int(sub.min())
            chave_max[s] = int(sub.max())
    _TABELAS[name] = (ks, lms, chave_min, chave_max)
    return _TABELAS[name]


def _lookup(name, sexo, chave):
    """Replica CORRESP(k;coluna_k;1): maior k <= alvo, validando faixa/sexo.
    Retorna (L,M,S) ou None se fora da faixa da tabela para aquele sexo."""
    if sexo not in (1, 2):
        return None
    ks, lms, cmin, cmax = _load(name)
    if cmin[sexo] is None or chave < cmin[sexo]:
        return None
    # clamp ao maximo (1827-1856 dias -> 1826), como o CORRESP aproximado
    if chave > cmax[sexo]:
        chave = cmax[sexo]
    alvo = sexo * 1_000_000 + chave
    i = bisect.bisect_right(ks, alvo) - 1
    if i < 0:
        return None
    return lms[ks[i]]


def _z(X, L, M, S):
    if L != 0:
        return (((X / M) ** L) - 1.0) / (L * S)
    return math.log(X / M) / S


def _z_extremo(X, L, M, S, z):
    """Ajuste OMS para |z|>3 (indices baseados em peso)."""
    if abs(z) <= 3:
        return z
    if L != 0:
        sd3p = M * (1 + L * S * 3) ** (1 / L)
        sd3n = M * (1 + L * S * -3) ** (1 / L)
        sd2p = M * (1 + L * S * 2) ** (1 / L)
        sd2n = M * (1 + L * S * -2) ** (1 / L)
    else:
        sd3p = M * math.exp(S * 3); sd3n = M * math.exp(S * -3)
        sd2p = M * math.exp(S * 2); sd2n = M * math.exp(S * -2)
    if z > 3:
        return 3 + (X - sd3p) / (sd3p - sd2p)
    return -3 + (X - sd3n) / (sd2n - sd3n)


def idade_meses(dn, data):
    """meses completos (=DATADIF m)."""
    m = (data.year - dn.year) * 12 + (data.month - dn.month)
    if data.day < dn.day:
        m -= 1
    return m


def idade_dias(dn, data):
    return (data - dn).days


def _ajuste_lh(valor_cm, meses, medida_em):
    """Converte para comprimento(<24m) ou estatura(>=24m) com regra +-0,7."""
    if valor_cm is None:
        return None
    me = (medida_em or "").strip().lower()
    if meses < 24:  # padrao: comprimento (deitado)
        if me.startswith("em p"):  # medido em pe -> +0,7 -> comprimento
            return valor_cm + 0.7
        return valor_cm
    else:           # padrao: estatura (em pe)
        if me.startswith("deit"):  # medido deitado -> -0,7 -> estatura
            return valor_cm - 0.7
        return valor_cm


def calc(sexo, dn, data, peso_kg, altura_cm, medida_em=None):
    """Retorna dict com idade, IMC e os 4 escores-z (ou None se fora de faixa)."""
    out = {"meses": None, "dias": None, "imc": None,
           "z_imc": None, "z_est": None, "z_peso": None, "z_pe": None}
    if None in (sexo, dn, data, peso_kg, altura_cm) or peso_kg <= 0 or altura_cm <= 0:
        return out
    meses = idade_meses(dn, data)
    dias = idade_dias(dn, data)
    out["meses"] = meses
    out["dias"] = dias
    if meses < 0 or meses > 228:
        return out
    imc = peso_kg / (altura_cm / 100.0) ** 2
    out["imc"] = round(imc, 2)
    lh = _ajuste_lh(altura_cm, meses, medida_em)

    usar_0a5 = meses <= 60

    # IMC/idade (0-228)
    lms = _lookup("imc_idade_0a5" if usar_0a5 else "imc_idade_5a19",
                  sexo, dias if usar_0a5 else meses)
    if lms:
        z = _z(imc, *lms)
        out["z_imc"] = round(_z_extremo(imc, *lms, z), 2)

    # Estatura/idade (0-228) -- sem ajuste de extremo
    lms = _lookup("est_idade_0a5" if usar_0a5 else "est_idade_5a19",
                  sexo, dias if usar_0a5 else meses)
    if lms:
        out["z_est"] = round(_z(lh, *lms), 2)

    # Peso/idade (0-120 meses)
    if meses <= 120:
        lms = _lookup("peso_idade_0a5" if usar_0a5 else "peso_idade_5a10",
                      sexo, dias if usar_0a5 else meses)
        if lms:
            z = _z(peso_kg, *lms)
            out["z_peso"] = round(_z_extremo(peso_kg, *lms, z), 2)

    # Peso/estatura (0-60 meses)
    if meses <= 60 and lh is not None:
        if meses < 24:
            cm10 = round(lh * 10)
            lms = _lookup("peso_comprimento", sexo, cm10)
        else:
            cm10 = round(lh * 10)
            lms = _lookup("peso_altura", sexo, cm10)
        if lms:
            z = _z(peso_kg, *lms)
            out["z_pe"] = round(_z_extremo(peso_kg, *lms, z), 2)

    return out


# ---- Classificacoes SISVAN/MS ----
def classif_imc(z, meses):
    if z is None:
        return ""
    if meses <= 60:  # <=60 meses usa categorias de menores de 5 anos (consistente com a tabela do escore-z)
        if z < -3: return "Magreza acentuada"
        if z < -2: return "Magreza"
        if z <= 1: return "Eutrofia"
        if z <= 2: return "Risco de sobrepeso"
        if z <= 3: return "Sobrepeso"
        return "Obesidade"
    else:
        if z < -3: return "Magreza acentuada"
        if z < -2: return "Magreza"
        if z <= 1: return "Eutrofia"
        if z <= 2: return "Sobrepeso"
        if z <= 3: return "Obesidade"
        return "Obesidade grave"


def classif_est(z):
    if z is None: return ""
    if z < -3: return "Muito baixa"
    if z < -2: return "Baixa"
    return "Adequada"


def classif_peso(z):
    if z is None: return ""
    if z < -3: return "Muito baixo"
    if z < -2: return "Baixo"
    if z <= 2: return "Adequado"
    return "Elevado"


def classif_pe(z):  # peso/estatura = mesmos cortes de IMC <5
    if z is None: return ""
    if z < -3: return "Magreza acentuada"
    if z < -2: return "Magreza"
    if z <= 1: return "Eutrofia"
    if z <= 2: return "Risco de sobrepeso"
    if z <= 3: return "Sobrepeso"
    return "Obesidade"
