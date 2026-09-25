# -*- coding: utf-8 -*-
"""
Gera os dois arquivos Excel da Avaliacao Nutricional Infantil (0-19 anos):
  - Calculadora_OMS.xlsx  (motor: tabelas OMS + formulas; PROCV/INDICE-CORRESP)
  - BD_AvaliacaoNutricional.xlsx (banco leve, so dados digitados)

So' usa funcoes CLASSICAS do Excel (INDICE/CORRESP/SE/CONT.SE/AGREGAR-livre)
para maxima compatibilidade. A logica espelha zscore_ref.py (validado).

Uso: python build_planilha.py
"""
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import LineChart, BarChart, Reference

DIR = os.path.dirname(os.path.abspath(__file__))
TAB = os.path.join(DIR, "tabelas_oms")
N = 1500                      # linhas de dados na aba Calculo
DASH = "—"              # em dash para "fora de faixa"

# ---------- estilos / cores ----------
AZUL = "1F4E79"; AZUL2 = "2E75B6"; CINZA = "808080"; CINZAF = "F2F2F2"
VERDE_F = "C6EFCE"; VERDE_T = "006100"
VERM_F = "FFC7CE"; VERM_T = "9C0006"
AMAR_F = "FFEB9C"; AMAR_T = "9C6500"
HEAD_FILL = PatternFill("solid", fgColor=AZUL)
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")   # amarelo claro = digitar aqui
AUTO_FILL = PatternFill("solid", fgColor="DDEBF7")    # azul claro = vem do banco (Power Query)
HEAD_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(color=AZUL, bold=True, size=16)
SUB_FONT = Font(color=CINZA, italic=True, size=10)
BORDER = Border(*[Side(style="thin", color="D9D9D9")] * 4)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")

# categorias -> cor (para formatacao condicional)
CAT_CORES = {
    VERDE_F: ["Eutrofia", "Adequada", "Adequado"],
    AMAR_F: ["Magreza", "Baixa", "Baixo", "Risco de sobrepeso", "Sobrepeso", "Elevado"],
    VERM_F: ["Magreza acentuada", "Muito baixa", "Muito baixo", "Obesidade", "Obesidade grave"],
}
CAT_TXT = {VERDE_F: VERDE_T, AMAR_F: AMAR_T, VERM_F: VERM_T}


# ============================================================
# 1) Tabelas de referencia (abas ocultas) + nomes definidos
# ============================================================
REFS = [
    ("imc_idade_0a5",  "ref_imc_0a5",  "imc05"),
    ("imc_idade_5a19", "ref_imc_5a19", "imc519"),
    ("est_idade_0a5",  "ref_est_0a5",  "est05"),
    ("est_idade_5a19", "ref_est_5a19", "est519"),
    ("peso_idade_0a5", "ref_peso_0a5", "peso05"),
    ("peso_idade_5a10","ref_peso_5a10","peso510"),
    ("peso_comprimento","ref_peso_comp","pcomp"),
    ("peso_altura",    "ref_peso_alt", "palt"),
]


def build_refs(wb):
    for csv, sheet, pref in REFS:
        df = pd.read_csv(os.path.join(TAB, csv + ".csv"))
        ws = wb.create_sheet(sheet)
        ws.sheet_state = "hidden"
        ws.append(["k", "sexo", "chave", "l", "m", "s"])
        for _, row in df.iterrows():
            ws.append([int(row.k), int(row.sexo), int(row.chave),
                       float(row.l), float(row.m), float(row.s)])
        last = len(df) + 1
        for nm, colL in [("k", "A"), ("l", "D"), ("m", "E"), ("s", "F")]:
            wb.defined_names.add(DefinedName(
                f"{pref}_{nm}", attr_text=f"{sheet}!${colL}$2:${colL}${last}"))


# ============================================================
# 2) Aba Calculo  (motor)
# ============================================================
# ordem das colunas (visiveis primeiro, depois auxiliares ocultas)
KEYS_VIS = ["ID", "Nome", "Sexo", "DN", "Data", "Peso", "Altura", "Medida",
            "Idade", "IMC", "zIMC", "dIMC", "zEst", "dEst", "zPeso", "dPeso",
            "zPE", "dPE"]
KEYS_AUX = ["sexo_n", "meses", "dias", "comp", "k_age", "k_pe",
            "Limc", "Mimc", "Simc", "zrimc", "sd2pi", "sd3pi", "sd2ni", "sd3ni",
            "Lest", "Mest", "Sest",
            "Lpeso", "Mpeso", "Speso", "zrpeso", "sd2pp", "sd3pp", "sd2np", "sd3np",
            "Lpe", "Mpe", "Spe", "zrpe", "sd2pe", "sd3pe", "sd2ne", "sd3ne",
            "fichaseq"]
KEYS = KEYS_VIS + KEYS_AUX
COL = {k: get_column_letter(i + 1) for i, k in enumerate(KEYS)}

HEADERS = {
    "ID": "ID / Prontuário", "Nome": "Nome", "Sexo": "Sexo",
    "DN": "Data nasc.", "Data": "Data da medida", "Peso": "Peso (kg)",
    "Altura": "Altura (cm)", "Medida": "Medida",
    "Idade": "Idade", "IMC": "IMC",
    "zIMC": "Z IMC/idade", "dIMC": "Diagnóstico IMC/idade",
    "zEst": "Z Estatura/idade", "dEst": "Diagnóstico Estatura/idade",
    "zPeso": "Z Peso/idade", "dPeso": "Diagnóstico Peso/idade",
    "zPE": "Z Peso/estatura", "dPE": "Diagnóstico Peso/estatura",
}


def C(k):  # coluna absoluta sem linha
    return "$" + COL[k]


def templates(r):
    """Retorna dict key->formula (str) para a linha r."""
    g = lambda k: f"${COL[k]}{r}"  # ref absoluta na coluna, linha r
    t = {}
    # --- entradas: puxadas da aba Banco (preenchida pelo Power Query), 1:1 por linha ---
    # Banco tem as colunas na MESMA ordem/letra (A=ID..H=Medida).
    # Usa INDICE(coluna_inteira;LIN()) -> imune ao deslocamento quando o Power Query
    # insere/remove linhas no refresh. Calculo linha r <-> Banco linha r.
    for k in ("ID", "Nome", "Sexo", "DN", "Data", "Peso", "Altura", "Medida"):
        ix = f'INDEX(Banco!${COL[k]}:${COL[k]},ROW())'
        t[k] = f'=IF({ix}="","",{ix})'
    # --- auxiliares base ---
    t["sexo_n"] = (f'=IF({g("Sexo")}="","",IF(UPPER(LEFT({g("Sexo")},1))="M",1,'
                   f'IF(UPPER(LEFT({g("Sexo")},1))="F",2,"")))')
    t["meses"] = f'=IF(OR({g("DN")}="",{g("Data")}=""),"",DATEDIF({g("DN")},{g("Data")},"m"))'
    t["dias"] = f'=IF(OR({g("DN")}="",{g("Data")}=""),"",{g("Data")}-{g("DN")})'
    t["comp"] = (f'=IF(OR({g("Altura")}="",{g("meses")}=""),"",'
                 f'IF({g("meses")}<24,IF(LEFT(LOWER({g("Medida")}),4)="em p",{g("Altura")}+0.7,{g("Altura")}),'
                 f'IF(LEFT(LOWER({g("Medida")}),4)="deit",{g("Altura")}-0.7,{g("Altura")})))')
    t["k_age"] = (f'=IF(OR({g("sexo_n")}="",{g("meses")}=""),"",'
                  f'{g("sexo_n")}*1000000+IF({g("meses")}<=60,{g("dias")},{g("meses")}))')
    t["k_pe"] = (f'=IF(OR({g("sexo_n")}="",{g("comp")}="",{g("meses")}>60,{g("meses")}=""),"",'
                 f'{g("sexo_n")}*1000000+ROUND({g("comp")}*10,0))')
    t["Idade"] = f'=IF({g("meses")}="","",INT({g("meses")}/12)&"a "&MOD({g("meses")},12)&"m")'
    t["IMC"] = f'=IF(OR({g("Peso")}="",{g("Altura")}=""),"",ROUND({g("Peso")}/({g("Altura")}/100)^2,2))'

    # --- lookups L,M,S (INDICE + CORRESP tipo 1) ---
    def lms(prefL, k_used, cond, n5, n19, suffix):
        # n5/n19 = nomes definidos (ex: imc05, imc519)
        rng = lambda col: f'IF({cond},{n5}_{col},{n19}_{col})'
        return (f'=IF({g(k_used)}="","",IFERROR(INDEX({rng(suffix)},'
                f'MATCH({g(k_used)},{rng("k")},1)),""))')

    cond_age = f'{g("meses")}<=60'
    t["Limc"] = lms("Limc", "k_age", cond_age, "imc05", "imc519", "l")
    t["Mimc"] = lms("Mimc", "k_age", cond_age, "imc05", "imc519", "m")
    t["Simc"] = lms("Simc", "k_age", cond_age, "imc05", "imc519", "s")
    t["Lest"] = lms("Lest", "k_age", cond_age, "est05", "est519", "l")
    t["Mest"] = lms("Mest", "k_age", cond_age, "est05", "est519", "m")
    t["Sest"] = lms("Sest", "k_age", cond_age, "est05", "est519", "s")
    t["Lpeso"] = lms("Lpeso", "k_age", cond_age, "peso05", "peso510", "l")
    t["Mpeso"] = lms("Mpeso", "k_age", cond_age, "peso05", "peso510", "m")
    t["Speso"] = lms("Speso", "k_age", cond_age, "peso05", "peso510", "s")
    cond_pe = f'{g("meses")}<24'
    # peso/estatura so' 0-60 meses
    def lms_pe(suffix):
        rng = lambda col: f'IF({cond_pe},pcomp_{col},palt_{col})'
        return (f'=IF(OR({g("k_pe")}="",{g("meses")}>60),"",IFERROR(INDEX({rng(suffix)},'
                f'MATCH({g("k_pe")},{rng("k")},1)),""))')
    t["Lpe"] = lms_pe("l"); t["Mpe"] = lms_pe("m"); t["Spe"] = lms_pe("s")

    # --- z bruto + cortes SD (para ajuste de extremos) ---
    def zraw(L, M, S, X):
        return (f'=IF({g(L)}="","",IF({g(L)}=0,LN({g(X)}/{g(M)})/{g(S)},'
                f'((({g(X)}/{g(M)})^{g(L)})-1)/({g(L)}*{g(S)})))')

    def sd(L, M, S, mult):
        return (f'=IF({g(L)}="","",IF({g(L)}=0,{g(M)}*EXP({g(S)}*{mult}),'
                f'{g(M)}*(1+{g(L)}*{g(S)}*{mult})^(1/{g(L)})))')

    t["zrimc"] = zraw("Limc", "Mimc", "Simc", "IMC")
    t["sd2pi"] = sd("Limc", "Mimc", "Simc", 2); t["sd3pi"] = sd("Limc", "Mimc", "Simc", 3)
    t["sd2ni"] = sd("Limc", "Mimc", "Simc", -2); t["sd3ni"] = sd("Limc", "Mimc", "Simc", -3)
    t["zrpeso"] = zraw("Lpeso", "Mpeso", "Speso", "Peso")
    t["sd2pp"] = sd("Lpeso", "Mpeso", "Speso", 2); t["sd3pp"] = sd("Lpeso", "Mpeso", "Speso", 3)
    t["sd2np"] = sd("Lpeso", "Mpeso", "Speso", -2); t["sd3np"] = sd("Lpeso", "Mpeso", "Speso", -3)
    t["zrpe"] = zraw("Lpe", "Mpe", "Spe", "Peso")
    t["sd2pe"] = sd("Lpe", "Mpe", "Spe", 2); t["sd3pe"] = sd("Lpe", "Mpe", "Spe", 3)
    t["sd2ne"] = sd("Lpe", "Mpe", "Spe", -2); t["sd3ne"] = sd("Lpe", "Mpe", "Spe", -3)

    # --- z final (com ajuste OMS para |z|>3 em indices de peso) ---
    def zfin_peso(zr, X, sd2p, sd3p, sd2n, sd3n, guard_max, Lh):
        return (f'=IF(OR({g("sexo_n")}="",{g(X)}="",{g("meses")}="",{g("meses")}<0,'
                f'{g("meses")}>{guard_max},{g(Lh)}=""),"{DASH}",'
                f'ROUND(IF(ABS({g(zr)})<=3,{g(zr)},'
                f'IF({g(zr)}>3,3+({g(X)}-{g(sd3p)})/({g(sd3p)}-{g(sd2p)}),'
                f'-3+({g(X)}-{g(sd3n)})/({g(sd2n)}-{g(sd3n)}))),2))')

    t["zIMC"] = zfin_peso("zrimc", "IMC", "sd2pi", "sd3pi", "sd2ni", "sd3ni", 228, "Limc")
    t["zPeso"] = zfin_peso("zrpeso", "Peso", "sd2pp", "sd3pp", "sd2np", "sd3np", 120, "Lpeso")
    t["zPE"] = zfin_peso("zrpe", "Peso", "sd2pe", "sd3pe", "sd2ne", "sd3ne", 60, "Lpe")
    # estatura/idade: sem ajuste de extremos
    t["zEst"] = (f'=IF(OR({g("sexo_n")}="",{g("comp")}="",{g("meses")}="",{g("meses")}<0,'
                 f'{g("meses")}>228,{g("Lest")}=""),"{DASH}",'
                 f'ROUND(IF({g("Lest")}=0,LN({g("comp")}/{g("Mest")})/{g("Sest")},'
                 f'((({g("comp")}/{g("Mest")})^{g("Lest")})-1)/({g("Lest")}*{g("Sest")})),2))')

    # --- diagnosticos (SISVAN/MS) ---
    z = lambda k: g(k)
    def ge(k):  # guarda: vazio ou fora de faixa
        return f'OR({g(k)}="",{g(k)}="{DASH}")'
    t["dIMC"] = (f'=IF({ge("zIMC")},"",IF({g("meses")}<=60,'
                 f'IF({z("zIMC")}<-3,"Magreza acentuada",IF({z("zIMC")}<-2,"Magreza",'
                 f'IF({z("zIMC")}<=1,"Eutrofia",IF({z("zIMC")}<=2,"Risco de sobrepeso",'
                 f'IF({z("zIMC")}<=3,"Sobrepeso","Obesidade"))))),'
                 f'IF({z("zIMC")}<-3,"Magreza acentuada",IF({z("zIMC")}<-2,"Magreza",'
                 f'IF({z("zIMC")}<=1,"Eutrofia",IF({z("zIMC")}<=2,"Sobrepeso",'
                 f'IF({z("zIMC")}<=3,"Obesidade","Obesidade grave")))))))')
    t["dEst"] = (f'=IF({ge("zEst")},"",IF({z("zEst")}<-3,"Muito baixa",'
                 f'IF({z("zEst")}<-2,"Baixa","Adequada")))')
    t["dPeso"] = (f'=IF({ge("zPeso")},"",IF({z("zPeso")}<-3,"Muito baixo",'
                  f'IF({z("zPeso")}<-2,"Baixo",IF({z("zPeso")}<=2,"Adequado","Elevado"))))')
    t["dPE"] = (f'=IF({ge("zPE")},"",IF({z("zPE")}<-3,"Magreza acentuada",'
                f'IF({z("zPE")}<-2,"Magreza",IF({z("zPE")}<=1,"Eutrofia",'
                f'IF({z("zPE")}<=2,"Risco de sobrepeso",IF({z("zPE")}<=3,"Sobrepeso","Obesidade"))))))')

    # --- sequencia p/ ficha (k-esima medida da crianca selecionada) ---
    t["fichaseq"] = (f'=IF(OR({g("ID")}="",{g("ID")}<>Ficha!$C$4),"",'
                     f'COUNTIF($A$2:$A{r},Ficha!$C$4))')
    return t


def build_calculo(wb):
    ws = wb.create_sheet("Calculo")
    # cabecalho
    for k in KEYS:
        c = ws.cell(row=1, column=KEYS.index(k) + 1, value=HEADERS.get(k, k))
        c.fill = HEAD_FILL; c.font = HEAD_FONT; c.alignment = CENTER; c.border = BORDER
    # formulas
    for r in range(2, N + 2):
        t = templates(r)
        for k in KEYS:
            col = KEYS.index(k) + 1
            cell = ws.cell(row=r, column=col)
            if k in t:
                cell.value = t[k]
            if k in ("DN", "Data"):
                cell.number_format = "dd/mm/yyyy"
            if k in ("Peso", "Altura", "IMC", "zIMC", "zEst", "zPeso", "zPE"):
                cell.number_format = "0.00" if k != "IMC" else "0.0"
            if k in KEYS_VIS[:8]:
                cell.fill = AUTO_FILL
    # larguras
    larg = {"ID": 14, "Nome": 22, "Sexo": 11, "DN": 12, "Data": 12, "Peso": 9,
            "Altura": 10, "Medida": 10, "Idade": 9, "IMC": 7,
            "zIMC": 8, "dIMC": 19, "zEst": 9, "dEst": 19, "zPeso": 8, "dPeso": 18,
            "zPE": 9, "dPE": 19}
    for k, w in larg.items():
        ws.column_dimensions[COL[k]].width = w
    # ocultar auxiliares
    for k in KEYS_AUX:
        ws.column_dimensions[COL[k]].hidden = True
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{COL['dPE']}{N+1}"

    # formatacao condicional nos diagnosticos
    for dcol in ["dIMC", "dEst", "dPeso", "dPE"]:
        rng = f"{COL[dcol]}2:{COL[dcol]}{N+1}"
        for fill_color, cats in CAT_CORES.items():
            for cat in cats:
                ws.conditional_formatting.add(rng, CellIsRule(
                    operator="equal", formula=[f'"{cat}"'],
                    fill=PatternFill("solid", fgColor=fill_color),
                    font=Font(color=CAT_TXT[fill_color])))

    # nomes definidos para Ficha/Painel
    nm = {"calc_ID": "ID", "calc_Nome": "Nome", "calc_Data": "Data",
          "calc_zIMC": "zIMC", "calc_dIMC": "dIMC", "calc_zEst": "zEst",
          "calc_dEst": "dEst", "calc_zPeso": "zPeso", "calc_dPeso": "dPeso",
          "calc_zPE": "zPE", "calc_dPE": "dPE", "calc_seq": "fichaseq",
          "calc_meses": "meses"}
    for name, k in nm.items():
        wb.defined_names.add(DefinedName(
            name, attr_text=f"Calculo!${COL[k]}$2:${COL[k]}${N+1}"))
    return ws


# ============================================================
# 3) Aba Ficha (acompanhamento individual)
# ============================================================
NF = 30  # linhas no historico da ficha


def build_ficha(wb):
    ws = wb.create_sheet("Ficha")
    ws.sheet_view.showGridLines = False
    ws["B2"] = "Ficha de Acompanhamento Individual"; ws["B2"].font = TITLE_FONT
    ws.merge_cells("B3:H3")
    ws["B3"] = "Selecione o ID/Prontuário para ver o histórico e a curva de evolução:"
    ws["B3"].font = SUB_FONT
    ws["B4"] = "ID:"; ws["B4"].font = Font(bold=True)
    # dropdown com a lista de IDs (celula de selecao C4)
    dv = DataValidation(type="list", formula1="calc_ID", allow_blank=True)
    ws.add_data_validation(dv); dv.add("C4")
    ws["C4"].fill = INPUT_FILL; ws["C4"].border = BORDER
    ws["B5"] = "Nome:"; ws["B5"].font = Font(bold=True)
    ws["C5"] = '=IFERROR(INDEX(calc_Nome,MATCH(C4,calc_ID,0)),"")'

    # situacao atual (ultima medicao), para nao precisar rolar a tabela
    ws["B6"] = "Última medição:"; ws["B6"].font = Font(bold=True)
    ws["C6"] = ('=IFERROR(TEXT(INDEX(calc_Data,MATCH(MAX(calc_seq),calc_seq,0)),"dd/mm/yyyy"),"")')
    ws["B7"] = "Diagnóstico atual (IMC/idade):"; ws["B7"].font = Font(bold=True)
    ws.merge_cells("C7:E7")
    ws["C7"] = '=IFERROR(INDEX(calc_dIMC,MATCH(MAX(calc_seq),calc_seq,0)),"")'
    ws["C7"].font = Font(bold=True, size=12)
    ws["C7"].alignment = LEFT
    for fill_color, cats in CAT_CORES.items():
        for cat in cats:
            ws.conditional_formatting.add("C7", CellIsRule(
                operator="equal", formula=[f'"{cat}"'],
                fill=PatternFill("solid", fgColor=fill_color),
                font=Font(color=CAT_TXT[fill_color], bold=True, size=12)))

    # tabela historico
    htxt = ["#", "Data", "Z IMC/id", "Diagnóstico IMC/id", "Z Est/id", "Z Peso/id", "Z Peso/est"]
    hrow = 10
    for j, h in enumerate(htxt):
        c = ws.cell(row=hrow, column=2 + j, value=h)
        c.fill = HEAD_FILL; c.font = HEAD_FONT; c.alignment = CENTER; c.border = BORDER
    for i in range(NF):
        r = hrow + 1 + i
        k = i + 1
        pos = f'MATCH({k},calc_seq,0)'
        ws.cell(row=r, column=2, value=f'=IF(ISNUMBER({pos}),{k},"")')
        ws.cell(row=r, column=3, value=f'=IFERROR(INDEX(calc_Data,{pos}),"")').number_format = "dd/mm/yyyy"
        ws.cell(row=r, column=4, value=f'=IFERROR(INDEX(calc_zIMC,{pos}),"")').number_format = "0.00"
        ws.cell(row=r, column=5, value=f'=IFERROR(INDEX(calc_dIMC,{pos}),"")')
        ws.cell(row=r, column=6, value=f'=IFERROR(INDEX(calc_zEst,{pos}),"")').number_format = "0.00"
        ws.cell(row=r, column=7, value=f'=IFERROR(INDEX(calc_zPeso,{pos}),"")').number_format = "0.00"
        ws.cell(row=r, column=8, value=f'=IFERROR(INDEX(calc_zPE,{pos}),"")').number_format = "0.00"
        for col in range(2, 9):
            ws.cell(row=r, column=col).border = BORDER
            ws.cell(row=r, column=col).alignment = CENTER
    for col, w in zip("BCDEFGH", [6, 12, 9, 18, 9, 9, 10]):
        ws.column_dimensions[col].width = w

    # grafico de evolucao dos escores-z
    chart = LineChart()
    chart.title = "Evolucao dos escores-z"
    chart.y_axis.title = "escore-z"; chart.x_axis.title = "Data"
    chart.height = 8; chart.width = 16
    data = Reference(ws, min_col=4, max_col=4, min_row=hrow, max_row=hrow + NF)  # zIMC
    data2 = Reference(ws, min_col=6, max_col=7, min_row=hrow, max_row=hrow + NF)  # zEst, zPeso
    cats = Reference(ws, min_col=3, max_col=3, min_row=hrow + 1, max_row=hrow + NF)
    chart.add_data(data, titles_from_data=True)
    chart.add_data(data2, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, f"J{hrow}")
    return ws


# ============================================================
# 4) Aba Consulta (calculo avulso de 1 crianca)
# ============================================================
def build_consulta(wb):
    ws = wb.create_sheet("Consulta")
    ws.sheet_view.showGridLines = False
    ws["B2"] = "Consulta Rápida (1 criança)"; ws["B2"].font = TITLE_FONT
    rotulos = [("Sexo", "Masculino/Feminino"), ("Data de nascimento", "dd/mm/aaaa"),
               ("Data da medida", "dd/mm/aaaa"), ("Peso (kg)", ""),
               ("Altura (cm)", ""), ("Medida", "Deitado/Em pé (opcional)")]
    base = 4
    for i, (rot, dica) in enumerate(rotulos):
        r = base + i
        ws.cell(row=r, column=2, value=rot).font = Font(bold=True)
        ic = ws.cell(row=r, column=3); ic.fill = INPUT_FILL; ic.border = BORDER
        ws.cell(row=r, column=4, value=dica).font = SUB_FONT
        if "nasc" in rot.lower() or "medida" in rot.lower():
            ic.number_format = "dd/mm/yyyy"
    # celulas de entrada: C4=Sexo C5=DN C6=Data C7=Peso C8=Altura C9=Medida
    dvs = DataValidation(type="list", formula1='"Masculino,Feminino"', allow_blank=True)
    dvm = DataValidation(type="list", formula1='"Deitado,Em pe"', allow_blank=True)
    ws.add_data_validation(dvs); ws.add_data_validation(dvm)
    dvs.add("C4"); dvm.add("C9")

    # helpers ocultos (col Z em diante) reaproveitando a MESMA logica
    # mapeamento de entrada -> celulas
    M = {"Sexo": "$C$4", "DN": "$C$5", "Data": "$C$6", "Peso": "$C$7",
         "Altura": "$C$8", "Medida": "$C$9"}
    H = {}  # key -> celula auxiliar (coluna Z+, linha 4)
    aux_keys = ["sexo_n", "meses", "dias", "comp", "k_age", "k_pe",
                "IMC", "Limc", "Mimc", "Simc", "zrimc", "sd2pi", "sd3pi", "sd2ni", "sd3ni",
                "Lest", "Mest", "Sest",
                "Lpeso", "Mpeso", "Speso", "zrpeso", "sd2pp", "sd3pp", "sd2np", "sd3np",
                "Lpe", "Mpe", "Spe", "zrpe", "sd2pe", "sd3pe", "sd2ne", "sd3ne",
                "zIMC", "zEst", "zPeso", "zPE", "dIMC", "dEst", "dPeso", "dPE"]
    for i, k in enumerate(aux_keys):
        H[k] = f"${get_column_letter(26 + i)}$4"

    def gg(k):
        if k in M: return M[k]
        if k in H: return H[k]
        return M.get(k, H.get(k))

    # constroi formulas de consulta espelhando templates(), trocando refs
    def buildf():
        f = {}
        f["sexo_n"] = (f'=IF({gg("Sexo")}="","",IF(UPPER(LEFT({gg("Sexo")},1))="M",1,'
                       f'IF(UPPER(LEFT({gg("Sexo")},1))="F",2,"")))')
        f["meses"] = f'=IF(OR({gg("DN")}="",{gg("Data")}=""),"",DATEDIF({gg("DN")},{gg("Data")},"m"))'
        f["dias"] = f'=IF(OR({gg("DN")}="",{gg("Data")}=""),"",{gg("Data")}-{gg("DN")})'
        f["comp"] = (f'=IF(OR({gg("Altura")}="",{gg("meses")}=""),"",'
                     f'IF({gg("meses")}<24,IF(LEFT(LOWER({gg("Medida")}),4)="em p",{gg("Altura")}+0.7,{gg("Altura")}),'
                     f'IF(LEFT(LOWER({gg("Medida")}),4)="deit",{gg("Altura")}-0.7,{gg("Altura")})))')
        f["k_age"] = (f'=IF(OR({gg("sexo_n")}="",{gg("meses")}=""),"",'
                      f'{gg("sexo_n")}*1000000+IF({gg("meses")}<=60,{gg("dias")},{gg("meses")}))')
        f["k_pe"] = (f'=IF(OR({gg("sexo_n")}="",{gg("comp")}="",{gg("meses")}>60,{gg("meses")}=""),"",'
                     f'{gg("sexo_n")}*1000000+ROUND({gg("comp")}*10,0))')
        f["IMC"] = f'=IF(OR({gg("Peso")}="",{gg("Altura")}=""),"",ROUND({gg("Peso")}/({gg("Altura")}/100)^2,2))'
        ca = f'{gg("meses")}<=60'
        def lms(n5, n19, s):
            rng = lambda c: f'IF({ca},{n5}_{c},{n19}_{c})'
            return f'=IF({gg("k_age")}="","",IFERROR(INDEX({rng(s)},MATCH({gg("k_age")},{rng("k")},1)),""))'
        for s in ("l", "m", "s"):
            f[{"l": "Limc", "m": "Mimc", "s": "Simc"}[s]] = lms("imc05", "imc519", s)
            f[{"l": "Lest", "m": "Mest", "s": "Sest"}[s]] = lms("est05", "est519", s)
            f[{"l": "Lpeso", "m": "Mpeso", "s": "Speso"}[s]] = lms("peso05", "peso510", s)
        cpe = f'{gg("meses")}<24'
        def lmspe(s):
            rng = lambda c: f'IF({cpe},pcomp_{c},palt_{c})'
            return (f'=IF(OR({gg("k_pe")}="",{gg("meses")}>60),"",IFERROR(INDEX({rng(s)},'
                    f'MATCH({gg("k_pe")},{rng("k")},1)),""))')
        f["Lpe"], f["Mpe"], f["Spe"] = lmspe("l"), lmspe("m"), lmspe("s")

        def zraw(L, Mk, S, X):
            return (f'=IF({gg(L)}="","",IF({gg(L)}=0,LN({gg(X)}/{gg(Mk)})/{gg(S)},'
                    f'((({gg(X)}/{gg(Mk)})^{gg(L)})-1)/({gg(L)}*{gg(S)})))')
        def sd(L, Mk, S, mult):
            return (f'=IF({gg(L)}="","",IF({gg(L)}=0,{gg(Mk)}*EXP({gg(S)}*{mult}),'
                    f'{gg(Mk)}*(1+{gg(L)}*{gg(S)}*{mult})^(1/{gg(L)})))')
        f["zrimc"] = zraw("Limc", "Mimc", "Simc", "IMC")
        f["sd2pi"] = sd("Limc", "Mimc", "Simc", 2); f["sd3pi"] = sd("Limc", "Mimc", "Simc", 3)
        f["sd2ni"] = sd("Limc", "Mimc", "Simc", -2); f["sd3ni"] = sd("Limc", "Mimc", "Simc", -3)
        f["zrpeso"] = zraw("Lpeso", "Mpeso", "Speso", "Peso")
        f["sd2pp"] = sd("Lpeso", "Mpeso", "Speso", 2); f["sd3pp"] = sd("Lpeso", "Mpeso", "Speso", 3)
        f["sd2np"] = sd("Lpeso", "Mpeso", "Speso", -2); f["sd3np"] = sd("Lpeso", "Mpeso", "Speso", -3)
        f["zrpe"] = zraw("Lpe", "Mpe", "Spe", "Peso")
        f["sd2pe"] = sd("Lpe", "Mpe", "Spe", 2); f["sd3pe"] = sd("Lpe", "Mpe", "Spe", 3)
        f["sd2ne"] = sd("Lpe", "Mpe", "Spe", -2); f["sd3ne"] = sd("Lpe", "Mpe", "Spe", -3)

        def zfin(zr, X, s2p, s3p, s2n, s3n, gmax, Lh):
            return (f'=IF(OR({gg("sexo_n")}="",{gg(X)}="",{gg("meses")}="",{gg("meses")}<0,'
                    f'{gg("meses")}>{gmax},{gg(Lh)}=""),"{DASH}",'
                    f'ROUND(IF(ABS({gg(zr)})<=3,{gg(zr)},'
                    f'IF({gg(zr)}>3,3+({gg(X)}-{gg(s3p)})/({gg(s3p)}-{gg(s2p)}),'
                    f'-3+({gg(X)}-{gg(s3n)})/({gg(s2n)}-{gg(s3n)}))),2))')
        f["zIMC"] = zfin("zrimc", "IMC", "sd2pi", "sd3pi", "sd2ni", "sd3ni", 228, "Limc")
        f["zPeso"] = zfin("zrpeso", "Peso", "sd2pp", "sd3pp", "sd2np", "sd3np", 120, "Lpeso")
        f["zPE"] = zfin("zrpe", "Peso", "sd2pe", "sd3pe", "sd2ne", "sd3ne", 60, "Lpe")
        f["zEst"] = (f'=IF(OR({gg("sexo_n")}="",{gg("comp")}="",{gg("meses")}="",{gg("meses")}<0,'
                     f'{gg("meses")}>228,{gg("Lest")}=""),"{DASH}",'
                     f'ROUND(IF({gg("Lest")}=0,LN({gg("comp")}/{gg("Mest")})/{gg("Sest")},'
                     f'((({gg("comp")}/{gg("Mest")})^{gg("Lest")})-1)/({gg("Lest")}*{gg("Sest")})),2))')
        ge = lambda k: f'OR({gg(k)}="",{gg(k)}="{DASH}")'
        f["dIMC"] = (f'=IF({ge("zIMC")},"",IF({gg("meses")}<=60,'
                     f'IF({gg("zIMC")}<-3,"Magreza acentuada",IF({gg("zIMC")}<-2,"Magreza",'
                     f'IF({gg("zIMC")}<=1,"Eutrofia",IF({gg("zIMC")}<=2,"Risco de sobrepeso",'
                     f'IF({gg("zIMC")}<=3,"Sobrepeso","Obesidade"))))),'
                     f'IF({gg("zIMC")}<-3,"Magreza acentuada",IF({gg("zIMC")}<-2,"Magreza",'
                     f'IF({gg("zIMC")}<=1,"Eutrofia",IF({gg("zIMC")}<=2,"Sobrepeso",'
                     f'IF({gg("zIMC")}<=3,"Obesidade","Obesidade grave")))))))')
        f["dEst"] = (f'=IF({ge("zEst")},"",IF({gg("zEst")}<-3,"Muito baixa",'
                     f'IF({gg("zEst")}<-2,"Baixa","Adequada")))')
        f["dPeso"] = (f'=IF({ge("zPeso")},"",IF({gg("zPeso")}<-3,"Muito baixo",'
                      f'IF({gg("zPeso")}<-2,"Baixo",IF({gg("zPeso")}<=2,"Adequado","Elevado"))))')
        f["dPE"] = (f'=IF({ge("zPE")},"",IF({gg("zPE")}<-3,"Magreza acentuada",'
                    f'IF({gg("zPE")}<-2,"Magreza",IF({gg("zPE")}<=1,"Eutrofia",'
                    f'IF({gg("zPE")}<=2,"Risco de sobrepeso",IF({gg("zPE")}<=3,"Sobrepeso","Obesidade"))))))')
        return f

    f = buildf()
    for k, cellref in H.items():
        col = cellref.split("$")[1]; row = int(cellref.split("$")[2])
        ws[f"{col}{row}"] = f[k]
        ws.column_dimensions[col].hidden = True

    # painel de resultados visivel
    res = [("IMC", "IMC", None), ("IMC/idade", "zIMC", "dIMC"),
           ("Estatura/idade", "zEst", "dEst"), ("Peso/idade", "zPeso", "dPeso"),
           ("Peso/estatura", "zPE", "dPE")]
    rr = 4
    for i, (rot, zk, dk) in enumerate(res):
        r = rr + i
        destaque = (rot == "IMC/idade")  # indice mais usado: reforcar visualmente
        ws.cell(row=r, column=6, value=rot).font = Font(bold=True, size=12 if destaque else 11)
        ws.cell(row=r, column=7, value=f"={H[zk]}").number_format = "0.00"
        if dk:
            dc = ws.cell(row=r, column=8, value=f"={H[dk]}")
            if destaque:
                dc.font = Font(bold=True, size=13)
    ws.cell(row=3, column=6, value="Resultado").font = Font(bold=True, color=AZUL, size=13)
    ws.cell(row=3, column=7, value="escore-z").font = SUB_FONT
    ws.cell(row=3, column=8, value="diagnostico").font = SUB_FONT
    for col, w in zip("BCDEFGH", [12, 14, 6, 6, 16, 11, 20]):
        ws.column_dimensions[col].width = max(ws.column_dimensions[col].width or 0, w)
    # CF nos diagnosticos da consulta
    for r in range(rr, rr + len(res)):
        cellr = f"H{r}"
        for fill_color, cats in CAT_CORES.items():
            for cat in cats:
                ws.conditional_formatting.add(cellr, CellIsRule(
                    operator="equal", formula=[f'"{cat}"'],
                    fill=PatternFill("solid", fgColor=fill_color),
                    font=Font(color=CAT_TXT[fill_color])))
    return ws


# ============================================================
# 5) Aba Painel (resumo)
# ============================================================
def build_painel(wb):
    ws = wb.create_sheet("Painel")
    ws.sheet_view.showGridLines = False
    ws["B2"] = "Painel — Resumo da turma / unidade"; ws["B2"].font = TITLE_FONT
    ws["B4"] = "Total de crianças avaliadas:"; ws["B4"].font = Font(bold=True)
    ws["E4"] = '=SUMPRODUCT(--(calc_ID<>""))'
    ws["B5"] = "Com IMC/idade calculado:"; ws["B5"].font = Font(bold=True)
    ws["E5"] = '=SUMPRODUCT(--(calc_dIMC<>""))'

    cats_imc = ["Magreza acentuada", "Magreza", "Eutrofia", "Risco de sobrepeso",
                "Sobrepeso", "Obesidade", "Obesidade grave"]
    ws["B7"] = "Distribuição IMC/idade"; ws["B7"].font = Font(bold=True, color=AZUL, size=13)
    ws.cell(row=8, column=2, value="Categoria").font = HEAD_FONT
    ws.cell(row=8, column=2).fill = HEAD_FILL
    ws.cell(row=8, column=3, value="N").font = HEAD_FONT; ws.cell(row=8, column=3).fill = HEAD_FILL
    ws.cell(row=8, column=4, value="%").font = HEAD_FONT; ws.cell(row=8, column=4).fill = HEAD_FILL
    for i, cat in enumerate(cats_imc):
        r = 9 + i
        ws.cell(row=r, column=2, value=cat)
        ws.cell(row=r, column=3, value=f'=COUNTIF(calc_dIMC,"{cat}")')
        ws.cell(row=r, column=4, value=f'=IFERROR(C{r}/$E$5,0)').number_format = "0.0%"
    nlast = 9 + len(cats_imc) - 1

    # indicadores-chave (com sinalizacao por cor: verde/amarelo/vermelho por faixa
    # de referencia aproximada, para leitura imediata sem precisar interpretar %)
    ws["B18"] = "Indicadores-chave"; ws["B18"].font = Font(bold=True, color=AZUL, size=13)
    ind = [
        ("Excesso de peso (sobrepeso+obesidade) %",
         '=IFERROR((COUNTIF(calc_dIMC,"Sobrepeso")+COUNTIF(calc_dIMC,"Obesidade")+COUNTIF(calc_dIMC,"Obesidade grave"))/$E$5,0)',
         0.10, 0.20),
        ("Magreza (magreza+magreza acentuada) %",
         '=IFERROR((COUNTIF(calc_dIMC,"Magreza")+COUNTIF(calc_dIMC,"Magreza acentuada"))/$E$5,0)',
         0.05, 0.10),
        ("Baixa estatura para idade %",
         '=IFERROR((COUNTIF(calc_dEst,"Baixa")+COUNTIF(calc_dEst,"Muito baixa"))/SUMPRODUCT(--(calc_dEst<>"")),0)',
         0.05, 0.10),
    ]
    for i, (rot, fml, corte_verde, corte_amarelo) in enumerate(ind):
        r = 19 + i
        ws.cell(row=r, column=2, value=rot)
        cell = ws.cell(row=r, column=5, value=fml)
        cell.number_format = "0.0%"
        rng = f"E{r}"
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="lessThan", formula=[str(corte_verde)],
            fill=PatternFill("solid", fgColor=VERDE_F), font=Font(color=VERDE_T)))
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="between", formula=[str(corte_verde), str(corte_amarelo)],
            fill=PatternFill("solid", fgColor=AMAR_F), font=Font(color=AMAR_T)))
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="greaterThan", formula=[str(corte_amarelo)],
            fill=PatternFill("solid", fgColor=VERM_F), font=Font(color=VERM_T)))
    ws.cell(row=19 + len(ind) + 1, column=2,
            value="Cores: referência aproximada para triagem rápida (verde/amarelo/vermelho), "
                  "não substitui avaliação clínica.").font = SUB_FONT

    for col, w in zip("BCDE", [42, 8, 8, 10]):
        ws.column_dimensions[col].width = w

    # grafico de barras
    chart = BarChart(); chart.title = "IMC/idade (N por categoria)"
    chart.height = 8; chart.width = 14; chart.legend = None
    data = Reference(ws, min_col=3, max_col=3, min_row=8, max_row=nlast)
    cats = Reference(ws, min_col=2, max_col=2, min_row=9, max_row=nlast)
    chart.add_data(data, titles_from_data=True); chart.set_categories(cats)
    ws.add_chart(chart, "G8")
    return ws


# ============================================================
# 6) Aba Instrucoes
# ============================================================
def build_instr(wb):
    ws = wb.create_sheet("Instrucoes")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["B"].width = 110
    linhas = [
        ("Avaliação Nutricional Infantil (0 a 19 anos) — OMS / SISVAN", "t"),
        ("Calculadora de escores-z e diagnóstico nutricional (padrão OMS 2006/2007).", "s"),
        ("", ""),
        ("COMO USAR — 3 PASSOS", "h"),
        ("1. Digite as medições", "n"),
        ("No arquivo BD_AvaliacaoNutricional.xlsx, aba Dados: uma linha por medição. "
         "Use sempre o mesmo ID/Prontuário para acompanhar a mesma criança ao longo do tempo.", "p"),
        ("2. Abra a Calculadora", "n"),
        ("Ela puxa o banco sozinha. Na primeira vez, clique em 'Habilitar Conteúdo'. "
         "Depois de mudar o banco, atualize em Dados > Atualizar Tudo.", "p"),
        ("3. Veja os resultados", "n"),
        ("CÁLCULO traz tudo calculado e colorido · FICHA mostra o histórico de uma criança · "
         "CONSULTA faz um cálculo avulso · PAINEL resume a turma/unidade.", "p"),
        ("", ""),
        ("DICAS IMPORTANTES:", "h"),
        ("• Mantenha os dois arquivos na MESMA PASTA local do computador (evite OneDrive/Drive).", "p"),
        ("• Preencha Data de nascimento e Data da medida — a idade é calculada sozinha.", "p"),
        ("• Você só digita no BANCO. Na Calculadora, não edite as colunas cinza (são cálculos).", "p"),
        ("• Peso/idade só até 10 anos e Peso/estatura só até 5 anos (padrão OMS); "
         "fora dessas faixas aparece o símbolo —.", "p"),
        ("Outros detalhes (medida deitado/em pé, mais de 1500 medições, etc.) estão no "
         "Guia rápido — Avaliação Nutricional.docx.", "s"),
        ("", ""),
        ("CLASSIFICAÇÃO (SISVAN / Ministério da Saúde):", "h"),
        ("IMC/idade < 5 anos:  < -3 Magreza acentuada | -3 a < -2 Magreza | -2 a +1 Eutrofia | "
         "> +1 a +2 Risco de sobrepeso | > +2 a +3 Sobrepeso | > +3 Obesidade", "p"),
        ("IMC/idade 5 a 19 anos:  < -3 Magreza acentuada | -3 a < -2 Magreza | -2 a +1 Eutrofia | "
         "> +1 a +2 Sobrepeso | > +2 a +3 Obesidade | > +3 Obesidade grave", "p"),
        ("Estatura/idade:  < -3 Muito baixa | -3 a < -2 Baixa | ≥ -2 Adequada", "p"),
        ("Peso/idade (0 a 10 anos):  < -3 Muito baixo | -3 a < -2 Baixo | -2 a +2 Adequado | > +2 Elevado", "p"),
        ("Peso/estatura (0 a 5 anos): mesmos cortes do IMC/idade < 5 anos.", "p"),
        ("", ""),
        ("Fontes: WHO Child Growth Standards (2006, 0–5 anos) e WHO Growth Reference (2007, 5–19 anos). "
         "Escores-z pelo método LMS, com o ajuste da OMS para |z| > 3 nos índices baseados em peso.", "s"),
    ]
    r = 2
    for txt, kind in linhas:
        c = ws.cell(row=r, column=2, value=txt)
        if kind == "t": c.font = TITLE_FONT
        elif kind == "s": c.font = SUB_FONT
        elif kind == "h": c.font = Font(bold=True, color=AZUL, size=12)
        elif kind == "n": c.font = Font(bold=True, color=AZUL2, size=12)
        else: c.font = Font(size=11)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30 if len(txt) > 90 else 16
        r += 1

    # celula com o caminho do banco (mesma pasta) -> usada pelo Power Query
    r += 1
    ws.cell(row=r, column=2, value="Caminho do banco de dados (detectado automaticamente):").font = SUB_FONT
    vc = ws.cell(row=r, column=3, value=(
        '=IFERROR(LEFT(CELL("filename",$A$1),FIND("[",CELL("filename",$A$1))-1)'
        '&"BD_AvaliacaoNutricional.xlsx","(salve este arquivo na mesma pasta do banco)")'))
    vc.font = SUB_FONT
    ws.parent.defined_names.add(DefinedName("caminho_bd", attr_text=f"Instrucoes!$C${r}"))
    return ws


# ============================================================
# Banco de Dados (arquivo separado, leve)
# ============================================================
def build_bd():
    wb = Workbook(); ws = wb.active; ws.title = "Dados"
    cols = ["ID / Prontuário", "Nome", "Sexo", "Data nasc.", "Data da medida",
            "Peso (kg)", "Altura (cm)", "Medida", "Observações"]
    for j, h in enumerate(cols):
        c = ws.cell(row=1, column=j + 1, value=h)
        c.fill = HEAD_FILL; c.font = HEAD_FONT; c.alignment = CENTER; c.border = BORDER
    ws.freeze_panes = "A2"
    larg = [16, 24, 12, 12, 12, 9, 10, 11, 30]
    for j, w in enumerate(larg):
        ws.column_dimensions[get_column_letter(j + 1)].width = w
    NB = 5000
    for r in range(2, NB + 2):
        ws.cell(row=r, column=4).number_format = "dd/mm/yyyy"
        ws.cell(row=r, column=5).number_format = "dd/mm/yyyy"
    dvs = DataValidation(type="list", formula1='"Masculino,Feminino"', allow_blank=True)
    dvm = DataValidation(type="list", formula1='"Deitado,Em pe"', allow_blank=True)
    ws.add_data_validation(dvs); ws.add_data_validation(dvm)
    dvs.add(f"C2:C{NB+1}"); dvm.add(f"H2:H{NB+1}")
    ws.auto_filter.ref = f"A1:I{NB+1}"
    # aba de ajuda
    h = wb.create_sheet("LEIA-ME")
    h.sheet_view.showGridLines = False
    h.column_dimensions["B"].width = 100
    msg = [
        "BANCO DE DADOS — Avaliação Nutricional Infantil",
        "",
        "Este arquivo guarda as medições (é leve: só dados, sem cálculos).",
        "Cada linha = uma medição de uma criança, na aba Dados.",
        "Para ACOMPANHAR uma criança ao longo do tempo, use sempre o MESMO ID/Prontuário",
        "e acrescente uma nova linha a cada nova medida (com a data daquele dia).",
        "",
        "Para ver os escores-z e diagnósticos: abra a Calculadora_OMS.xlsx,",
        "copie as linhas da aba Dados (colunas A até H) e cole na aba CÁLCULO (célula A2).",
        "",
        "Mantenha os dois arquivos na mesma pasta.",
    ]
    for i, m in enumerate(msg):
        c = h.cell(row=i + 2, column=2, value=m)
        c.font = TITLE_FONT if i == 0 else Font(size=11)
    return wb


# ============================================================
# main
# ============================================================
def main():
    wb = Workbook()
    wb.remove(wb.active)
    build_refs(wb)
    build_instr(wb)
    build_calculo(wb)
    build_ficha(wb)
    build_consulta(wb)
    build_painel(wb)
    # aba Banco (vazia) — sera' preenchida pelo Power Query (setup via COM)
    banco = wb.create_sheet("Banco")
    banco["A1"] = "ID"; banco.sheet_state = "hidden"
    # ordem das abas visiveis
    order = ["Instrucoes", "Calculo", "Ficha", "Consulta", "Painel"]
    sheets = [wb[n] for n in order] + [s for s in wb.worksheets if s.title not in order]
    wb._sheets = sheets
    wb.active = wb["Instrucoes"]
    try:
        wb.calculation.fullCalcOnLoad = True   # recalcula ao abrir
    except Exception:
        pass
    out1 = os.path.join(DIR, "Calculadora_OMS.xlsx")
    wb.save(out1)
    print(f"OK -> {out1}")

    bd = build_bd()
    out2 = os.path.join(DIR, "BD_AvaliacaoNutricional.xlsx")
    bd.save(out2)
    print(f"OK -> {out2}")


if __name__ == "__main__":
    main()
