# -*- coding: utf-8 -*-
"""
Monta as tabelas de referencia LMS da OMS (limpas) a partir dos arquivos-texto
oficiais baixados dos repositorios WorldHealthOrganization/anthro (0-5, 2006) e
/anthroplus (5-19, 2007).

Saida: tabelas_oms/*.csv com colunas [k, sexo, chave, l, m, s]
  k = sexo*1_000_000 + chave_inteira  (chave de busca binaria, ordenada asc.)
  chave = idade_em_dias (0-5) | idade_em_meses (5-19) | cm*10 (peso/estatura)

Tambem roda checagens-ancora contra valores publicados da OMS.

Uso: python montar_tabelas.py
"""
import os
import pandas as pd

DIR = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join(DIR, "tabelas_oms", "raw")
OUT = os.path.join(DIR, "tabelas_oms")


def _read(fn):
    return pd.read_csv(os.path.join(RAW, fn), sep="\t")


def _save(df, name, keycol):
    """df precisa ter colunas: sex, <keycol>, l, m, s."""
    out = pd.DataFrame({
        "sexo": df["sex"].astype(int),
        "chave": df[keycol].round(0).astype(int) if keycol != "cm10" else df["cm10"].astype(int),
        "l": df["l"].astype(float),
        "m": df["m"].astype(float),
        "s": df["s"].astype(float),
    })
    out["k"] = out["sexo"] * 1_000_000 + out["chave"]
    out = out[["k", "sexo", "chave", "l", "m", "s"]].sort_values("k").reset_index(drop=True)
    path = os.path.join(OUT, name)
    out.to_csv(path, index=False)
    rng = f"{out['chave'].min()}-{out['chave'].max()}"
    print(f"  {name:28s} {len(out):5d} linhas  chave {rng:>12s}  sexos {sorted(out['sexo'].unique())}")
    return out


def main():
    print("Montando tabelas de referencia LMS da OMS...\n")

    # ---- 0-5 anos (WHO 2006), idade em DIAS ----
    wei = _read("weianthro.txt")          # sex age l m s
    leng = _read("lenanthro.txt")         # sex age l m s loh
    bmi = _read("bmianthro.txt")          # sex age l m s loh
    wfl = _read("wflanthro.txt")          # sex length l m s lorh
    wfh = _read("wfhanthro.txt")          # sex height l m s lorh

    # ---- 5-19 anos (WHO 2007), idade em MESES ----
    wfa = _read("wfawho2007.txt")         # sex age l m s   (60-120)
    bfa = _read("bfawho2007.txt")         # sex age l m s   (60-228)
    hfa = _read("hfawho2007.txt")         # sex age l m s   (60-228)

    print("Tabelas geradas (chave = dias | meses | cm*10):")

    t = {}
    # peso/idade
    t["peso_idade_0a5"] = _save(wei.rename(columns={"age": "age"}), "peso_idade_0a5.csv", "age")
    t["peso_idade_5a10"] = _save(wfa[wfa["age"] >= 61].copy(), "peso_idade_5a10.csv", "age")
    # estatura/idade
    t["est_idade_0a5"] = _save(leng[["sex", "age", "l", "m", "s"]].copy(), "est_idade_0a5.csv", "age")
    t["est_idade_5a19"] = _save(hfa[(hfa["age"] >= 61) & (hfa["age"] <= 228)].copy(), "est_idade_5a19.csv", "age")
    # IMC/idade
    t["imc_idade_0a5"] = _save(bmi[["sex", "age", "l", "m", "s"]].copy(), "imc_idade_0a5.csv", "age")
    t["imc_idade_5a19"] = _save(bfa[(bfa["age"] >= 61) & (bfa["age"] <= 228)].copy(), "imc_idade_5a19.csv", "age")
    # peso/comprimento e peso/altura (cm*10)
    w = wfl.rename(columns={"length": "cmval"}).copy()
    w["cm10"] = (w["cmval"] * 10).round(0).astype(int)
    t["peso_comprimento"] = _save(w, "peso_comprimento.csv", "cm10")
    h = wfh.rename(columns={"height": "cmval"}).copy()
    h["cm10"] = (h["cmval"] * 10).round(0).astype(int)
    t["peso_altura"] = _save(h, "peso_altura.csv", "cm10")

    # ---- Checagens-ancora (valores publicados pela OMS) ----
    print("\nChecagens-ancora (valor calculado vs OMS publicado):")
    checks = [
        ("peso_idade_0a5", 1, 0, "m", 3.3464),   # menino 0 dias, mediana peso
        ("peso_idade_0a5", 2, 0, "m", 3.2322),   # menina 0 dias
        ("est_idade_0a5", 1, 0, "m", 49.8842),   # menino 0 dias, mediana comprimento
        ("est_idade_0a5", 2, 0, "m", 49.1477),   # menina 0 dias
        ("imc_idade_0a5", 1, 0, "m", 13.4069),   # menino 0 dias, mediana IMC
        ("peso_comprimento", 1, 450, "m", 2.4410),  # menino 45,0 cm
        ("peso_altura", 1, 650, "m", 7.4327),       # menino 65,0 cm
    ]
    ok = True
    for tab, sexo, chave, col, esperado in checks:
        df = t[tab]
        val = float(df[(df.sexo == sexo) & (df.chave == chave)][col].iloc[0])
        match = abs(val - esperado) < 0.001
        ok = ok and match
        print(f"  {tab:20s} sexo={sexo} chave={chave:4d} {col}={val:9.4f}  esperado={esperado:9.4f}  {'OK' if match else 'ERRO'}")
    print(f"\n  {'TODAS AS ANCORAS OK' if ok else 'FALHA EM ALGUMA ANCORA'}")


if __name__ == "__main__":
    main()
