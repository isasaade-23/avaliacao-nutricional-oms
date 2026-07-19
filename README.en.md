<div align="center">

# 🌱 Child Nutritional Assessment — WHO / SISVAN

**Two macro-free Excel spreadsheets that compute WHO z-scores and the nutritional diagnosis of children and adolescents (0–19 years) — to support growth monitoring in daycares, schools and health services.**

[🇧🇷 Português](README.md) · 🇺🇸 **English** · [🇪🇸 Español](README.es.md)

![MIT License](https://img.shields.io/badge/license-MIT-green)
![Made with Python](https://img.shields.io/badge/made%20with-Python-blue)
![WHO standard](https://img.shields.io/badge/standard-WHO%20Anthro%20%2F%20AnthroPlus-orange)
![No macros](https://img.shields.io/badge/Excel-no%20macros-lightgrey)

</div>

---

## 🎯 Purpose

Tracking a child's **nutritional status over time** is tedious: it means looking up WHO tables, computing z-scores by hand and classifying them per national guidelines (Brazil's SISVAN). In daycares, schools and health units this is often skipped — or done with errors.

This project delivers a **simple, reliable tool** for anyone who cares for children:

- You **type weight, height and dates** into a small "logbook" (the database);
- The **calculator automatically returns** the z-scores and the nutritional diagnosis, colour-coded;
- You can **follow the same child across months** (longitudinal growth) and see the trend curve.

The goal is to **improve nutritional surveillance** — catching wasting, stunting, overweight and obesity early — in services that see children but rarely have a system for it.

> **Nothing to install beyond Excel. No macros. No data ever leaves your computer.**

---

## ✨ Features

- 📏 **Four WHO anthropometric indices**: BMI-for-age, Height-for-age, Weight-for-age and Weight-for-height.
- 🎨 **Automatic diagnosis** using SISVAN / Ministry of Health categories, colour-coded (green = adequate, yellow = attention, red = deficit/excess).
- 📈 **Longitudinal follow-up**: pick a child and see the history and z-score curve.
- 👥 **Group dashboard**: nutritional-status distribution for the class/unit.
- 🔗 **Automatic link** database → calculator via **Power Query** (a native feature, not a macro).
- ✅ **Validated**: Excel formulas checked against a Python reference implementation and the official WHO R package.

---

## 🚀 Getting started

1. **Download both files** and keep them **in the same local folder** (avoid OneDrive/Drive — see note below):
   - `BD_AvaliacaoNutricional.xlsx` — the **database** (where you type);
   - `Calculadora_OMS.xlsx` — the **calculator** (computation engine).

2. **Enter measurements** in the database, **Dados** tab — one row per measurement. To follow a child, always reuse the **same ID** and add a new row at each visit. Save and close.

3. **Open the calculator.** It pulls the database on its own. The first time, click **"Enable Content"** on the yellow bar. Results appear on the **CÁLCULO** tab; use the **FICHA** tab for a child's history and **PAINEL** for the group view.

> 💡 **Tip (important):** Excel can only link the two files when they sit in a **local folder** (e.g. `C:\AvaliacaoNutricional`). In cloud-synced folders (OneDrive/Google Drive) the link may fail with a "file path" error. The bundled user guide (`.docx`, in Portuguese) walks through this in plain language.

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
| `Calculadora_OMS.xlsx` | **Deliverable** — the calculator (WHO tables + formulas) |
| `BD_AvaliacaoNutricional.xlsx` | **Deliverable** — the database (empty template) |
| `Guia rápido — Avaliação Nutricional.docx` | Plain-language user guide (Portuguese) |
| `tabelas_oms/*.csv` | Clean LMS tables used to build the workbook |
| `tabelas_oms/raw/*.txt` | Original WHO source files (provenance) |
| `montar_tabelas.py` | Builds the LMS tables from the WHO files |
| `zscore_ref.py` | Reference implementation (single source of truth) |
| `build_planilha.py` | Generates the two `.xlsx` files |
| `configurar_powerquery.ps1` | Sets up the automatic database → calculator link |
| `validate_ref.py` · `verify_*.py` | Validation (reference vs R; Excel vs reference) |

### Rebuild from scratch
```bash
python montar_tabelas.py     # build the LMS tables
python build_planilha.py     # generate the two .xlsx files
powershell -ExecutionPolicy Bypass -File configurar_powerquery.ps1   # wire up Power Query
```

---

## 📖 Citation

If this project is useful in your work, a citation is very welcome 🙏 (see the **"Cite this repository"** button on GitHub, or the [`CITATION.cff`](CITATION.cff) file):

> Saade, I. (2026). *Child Nutritional Assessment — WHO / SISVAN* [software]. https://github.com/isasaade-23/avaliacao-nutricional-oms

---

## 🔒 Privacy

This is **children's health data**. Use the **ID/record number** as the key, keep the **Name** optional, and store the database in a **restricted, controlled location**. The tool runs **entirely offline** — nothing is sent to the internet.

---

## 📄 License

Released under the **[MIT License](LICENSE)**: you may **use, copy, modify, distribute and sell**, including commercially, for free. I only ask that you **keep the copyright notice** and, if you can, **cite the project** — it helps a lot. 💚
