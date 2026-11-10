<div align="center">

# 🌱 Child Nutritional Assessment — WHO / SISVAN

**A single macro-free spreadsheet that computes WHO z-scores and the nutritional diagnosis of children and adolescents (0–19 years) — to support growth monitoring in daycares, schools and health services.**

🌐 **Official site: [isasaade-23.github.io/avaliacao-nutricional-oms](https://isasaade-23.github.io/avaliacao-nutricional-oms/)** (Excel, LibreOffice and Google Sheets versions)

[🇧🇷 Português](README.md) · 🇺🇸 **English** · [🇪🇸 Español](README.es.md)

[![CC BY-NC-SA 4.0 License](https://img.shields.io/badge/license-CC%20BY--NC--SA%204.0-green)](https://creativecommons.org/licenses/by-nc-sa/4.0/)
![Made with Python](https://img.shields.io/badge/made%20with-Python-blue)
![WHO standard](https://img.shields.io/badge/standard-WHO%20Anthro%20%2F%20AnthroPlus-orange)
![No macros](https://img.shields.io/badge/Excel-no%20macros-lightgrey)
[![Sponsor](https://img.shields.io/badge/support-GitHub%20Sponsors-EA4AAA?logo=github-sponsors)](https://github.com/sponsors/isasaade-23)
[![Ko-fi](https://img.shields.io/badge/support-Ko--fi-FF5E5B?logo=ko-fi&logoColor=white)](https://ko-fi.com/ailuciola)

</div>

---

## 🎯 Purpose

Tracking a child's **nutritional status over time** is tedious: it means looking up WHO tables, computing z-scores by hand and classifying them per national guidelines (Brazil's SISVAN). In daycares, schools and health units this is often skipped — or done with errors.

This project delivers a **simple, reliable tool** for anyone who cares for children:

- You **type weight, height and dates** into the **Banco** tab;
- The **same workbook automatically returns** the z-scores and the nutritional diagnosis, colour-coded;
- You can **follow the same child across months** (longitudinal growth) and see the trend curve.

The goal is to **improve nutritional surveillance** — catching wasting, stunting, overweight and obesity early — in services that see children but rarely have a system for it.

> **No macros. In the Excel and LibreOffice versions no data ever leaves your computer.**

---

## ✨ Features

- 📏 **Four WHO anthropometric indices**: BMI-for-age, Height-for-age, Weight-for-age and Weight-for-height.
- 🎨 **Automatic diagnosis** using SISVAN / Ministry of Health categories, colour-coded (green = adequate, yellow = attention, red = deficit/excess).
- 📈 **Longitudinal follow-up**: pick a child and see the history and z-score curve.
- 👥 **Group dashboard**: nutritional-status distribution for the class/unit.
- 📄 **One single file**: data entry and calculation in the same workbook, nothing to link.
- ✅ **Validated**: Excel formulas checked against a Python reference implementation and the official WHO R package.

---

## 🚀 Getting started

The calculator is a **single file**, in three versions with the same content. Pick the one that matches your software:

| Version | File | Download |
|---|---|---|
| Excel | `excel_Calculadora_OMS.xlsx` | [download](https://github.com/isasaade-23/avaliacao-nutricional-oms/raw/main/excel_Calculadora_OMS.xlsx) |
| LibreOffice | `libreoffice_Calculadora_OMS.ods` | [download](https://github.com/isasaade-23/avaliacao-nutricional-oms/raw/main/libreoffice_Calculadora_OMS.ods) |
| Google Sheets | `gsheet_Calculadora_OMS` | [make a copy](https://docs.google.com/spreadsheets/d/1ZrYO0hEXhPV054k1nwWsasHnFxFsZ1CQfY2td47YB90/copy) |

In Google Sheets the link opens the "Make a copy" window. The copy lands in your own Drive, and that is where you type the data.

1. **Enter measurements** in the **Banco** tab — one row per measurement. To follow a child, always reuse the **same ID** and add a new row at each visit.
2. **Read the results** in the **Calculo** tab, colour-coded. Use **Ficha** for a child's history and **Painel** for the group view. The **Instrucoes** tab has the step-by-step inside the workbook.

Data and calculations live in the same file, so there is nothing to link. The bundled user guide (`.pdf`/`.docx`, in Portuguese) walks through the steps in plain language.

---

## 📊 Indices and classification (SISVAN / Ministry of Health)

| Index | Age range | Categories (by z-score) |
|---|---|---|
| **BMI-for-age** (< 5 y) | 0–60 months | <−3 Severe thinness · −3 to <−2 Thinness · −2 to +1 Normal · >+1 to +2 Risk of overweight · >+2 to +3 Overweight · >+3 Obesity |
| **BMI-for-age** (5–19 y) | 61–228 months | <−3 Severe thinness · −3 to <−2 Thinness · −2 to +1 Normal · >+1 to +2 Overweight · >+2 to +3 Obesity · >+3 Severe obesity |
| **Height-for-age** | 0–228 months | <−3 Severely stunted · −3 to <−2 Stunted · ≥−2 Normal |
| **Weight-for-age** | 0–120 months (≤ 10 y) | <−3 Severely underweight · −3 to <−2 Underweight · −2 to +2 Normal · >+2 High |
| **Weight-for-height** | 0–60 months (≤ 5 y) | same cut-offs as BMI-for-age < 5 y |

Outside these ranges the result is "—" (Weight-for-age is undefined above 10 y and Weight-for-height above 5 y in the WHO standard).

---

## 🔬 Methodology

Z-scores follow the **LMS method**, the WHO standard:

$$z = \frac{(X/M)^L - 1}{L \cdot S} \quad \text{(or } z = \frac{\ln(X/M)}{S} \text{ when } L = 0\text{)}$$

where **L, M, S** are parameters by sex and age (or length/height) taken from the official WHO tables.

Details that faithfully reproduce WHO Anthro / AnthroPlus:

- **Age:** 0–5 years computed **in days** (WHO Anthro 2006); 5–19 years **in months** (WHO Growth Reference 2007).
- **WHO adjustment for |z| > 3** on weight-based indices (BMI-for-age, Weight-for-age, Weight-for-height), avoiding extrapolation beyond observed data.
- **Lying/standing adjustment (±0.7 cm):** applied when the measurement mode differs from the age standard (< 2 y lying; ≥ 2 y standing).
- **Only classic Excel functions** (INDEX/MATCH/IF) → runs on old versions too, no macros.

### Data sources
- **WHO Child Growth Standards (2006)** — 0–5 years — official [`anthro`](https://github.com/worldhealthorganization/anthro) package.
- **WHO Growth Reference (2007)** — 5–19 years — official [`anthroplus`](https://github.com/worldhealthorganization/anthroplus) package.

Original WHO source files are versioned under `tabelas_oms/raw/` (provenance).

### Validation
- The reference logic (`zscore_ref.py`) is **identical** to an implementation validated against the official WHO **R** package across thousands of cases.
- **16 test cases** were recomputed in **real Excel** and match the reference, including extremes (|z|>3) and age boundaries (0, 24, 60, 61, 120 and 228 months).

---

## 🗂️ Project structure

| File | Description |
|---|---|
| `excel_Calculadora_OMS.xlsx` | **Deliverable** — the calculator for Excel (WHO tables + formulas + data entry) |
| `libreoffice_Calculadora_OMS.ods` | **Deliverable** — the same calculator for LibreOffice |
| `Guia rápido — Avaliação Nutricional.docx` | Plain-language user guide (Portuguese) |
| `tabelas_oms/*.csv` | Clean LMS tables used to build the workbook |
| `tabelas_oms/raw/*.txt` | Original WHO source files (provenance) |
| `montar_tabelas.py` | Builds the LMS tables from the WHO files |
| `zscore_ref.py` | Reference implementation (single source of truth) |
| `build_planilha.py` | Generates the `.xlsx` file |
| `validate_ref.py` · `verify_*.py` | Validation (reference vs R; Excel vs reference) |

### Rebuild from scratch
```bash
python montar_tabelas.py     # build the LMS tables
python build_planilha.py     # generate the .xlsx file
```

---

## 📖 Citation

If this project is useful in your work, a citation is very welcome 🙏 (see the **"Cite this repository"** button on GitHub, or the [`CITATION.cff`](CITATION.cff) file):

> Saade, I. (2026). *Child Nutritional Assessment — WHO / SISVAN* [software]. https://github.com/isasaade-23/avaliacao-nutricional-oms

If you'd like to support the project's upkeep, it's entirely optional — via [GitHub Sponsors](https://github.com/sponsors/isasaade-23), [Ko-fi](https://ko-fi.com/ailuciola), or Pix (Brazil, key: isasaade23@gmail.com).

---

## 🔒 Privacy

This is **children's health data**. Use the **ID/record number** as the key, keep the **Name** optional, and store the file in a **restricted, controlled location**. The Excel and LibreOffice versions run **entirely offline**. In the Google Sheets version the data lives in Google's cloud: make a copy into your own Drive, never type into the original template, and restrict sharing of your copy.

---

## 📄 License

Released under **[CC BY-NC-SA 4.0](LICENSE)**: you may **use, copy, modify and distribute** freely, including in daycare centers, schools and health services. Conditions: **keep the attribution notice**, **do not sell** it (original or modified versions), and if you adapt the project, **share it under the same license**. If you can, **cite the project** — it helps a lot. 💚
