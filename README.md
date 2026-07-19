<div align="center">

# 🌱 Avaliação Nutricional Infantil — OMS / SISVAN

**Duas planilhas em Excel (sem macros) que calculam escores-z da OMS e o diagnóstico nutricional de crianças e adolescentes (0 a 19 anos) — para apoiar o acompanhamento em creches, escolas e serviços de saúde.**

🇧🇷 **Português** · [🇺🇸 English](README.en.md) · [🇪🇸 Español](README.es.md)

![Licença MIT](https://img.shields.io/badge/licen%C3%A7a-MIT-green)
![Feito com Python](https://img.shields.io/badge/feito%20com-Python-blue)
![Padrão OMS](https://img.shields.io/badge/padr%C3%A3o-WHO%20Anthro%20%2F%20AnthroPlus-orange)
![Sem macros](https://img.shields.io/badge/Excel-sem%20macros-lightgrey)

</div>

---

## 🎯 Para que serve

Acompanhar o **estado nutricional de crianças ao longo do tempo** costuma ser trabalhoso: exige consultar as tabelas da OMS, calcular escores-z à mão e classificar segundo o SISVAN. Em creches, escolas e unidades de saúde, isso muitas vezes não é feito — ou é feito com erro.

Este projeto entrega uma ferramenta **simples e confiável** para quem cuida de crianças:

- **Digita-se peso, altura e datas** em um "caderninho" (o banco de dados);
- A **calculadora devolve automaticamente** os escores-z e o diagnóstico nutricional, com cores;
- É possível **acompanhar a mesma criança ao longo dos meses** (crescimento longitudinal), enxergando a curva de evolução.

O objetivo é **melhorar a vigilância nutricional** — detectar cedo magreza, déficit de estatura, sobrepeso e obesidade — em serviços que atendem crianças mas nem sempre têm um sistema para isso.

> **Sem instalar nada além do Excel. Sem macros. Sem enviar dados para a internet.**

---

## ✨ Recursos

- 📏 **4 índices antropométricos** da OMS: IMC/idade, Estatura/idade, Peso/idade e Peso/estatura.
- 🎨 **Diagnóstico automático** pelas categorias do SISVAN / Ministério da Saúde, com cores (verde = adequado, amarelo = atenção, vermelho = déficit/excesso).
- 📈 **Acompanhamento longitudinal**: escolha uma criança e veja o histórico e a curva de escores-z.
- 👥 **Painel da turma/unidade**: distribuição do estado nutricional do grupo.
- 🔗 **Ligação automática** banco → calculadora via **Power Query** (recurso nativo, não é macro).
- ✅ **Validado**: fórmulas do Excel conferidas contra uma implementação de referência em Python e o pacote R oficial da OMS.

---

## 🚀 Como usar

1. **Baixe os dois arquivos** e mantenha-os **na mesma pasta local** (evite o OneDrive/Drive — veja a nota abaixo):
   - `BD_AvaliacaoNutricional.xlsx` — o **banco** (onde você digita);
   - `Calculadora_OMS.xlsx` — a **calculadora** (motor de cálculo).

2. **Digite as medições** no banco, aba **Dados** — uma linha por medição. Para acompanhar uma criança, use sempre o **mesmo ID/Prontuário** e acrescente uma linha a cada retorno. Salve e feche.

3. **Abra a calculadora.** Ela puxa o banco sozinha. Na primeira vez, clique em **"Habilitar Conteúdo"** na barra amarela. Os resultados aparecem na aba **CÁLCULO**; use a aba **FICHA** para o histórico de uma criança e o **PAINEL** para a visão do grupo.

> 💡 **Dica (importante):** o Excel só consegue ligar as duas planilhas se elas estiverem em uma **pasta local do computador** (ex.: `C:\AvaliacaoNutricional`). Em pastas sincronizadas na nuvem (OneDrive/Google Drive), a ligação pode falhar com um erro de "caminho de arquivo". O guia de uso (`.docx`) explica isso passo a passo, em linguagem simples.

---

## 📊 Índices e classificação (SISVAN / Ministério da Saúde)

| Índice | Faixa etária | Categorias (por escore-z) |
|---|---|---|
| **IMC/idade** (< 5 anos) | 0–60 meses | <−3 Magreza acentuada · −3 a <−2 Magreza · −2 a +1 Eutrofia · >+1 a +2 Risco de sobrepeso · >+2 a +3 Sobrepeso · >+3 Obesidade |
| **IMC/idade** (5–19 anos) | 61–228 meses | <−3 Magreza acentuada · −3 a <−2 Magreza · −2 a +1 Eutrofia · >+1 a +2 Sobrepeso · >+2 a +3 Obesidade · >+3 Obesidade grave |
| **Estatura/idade** | 0–228 meses | <−3 Muito baixa · −3 a <−2 Baixa · ≥−2 Adequada |
| **Peso/idade** | 0–120 meses (até 10 a) | <−3 Muito baixo · −3 a <−2 Baixo · −2 a +2 Adequado · >+2 Elevado |
| **Peso/estatura** | 0–60 meses (até 5 a) | mesmos cortes do IMC/idade < 5 anos |

Fora dessas faixas o resultado é "—" (Peso/idade não existe acima de 10 anos e Peso/estatura acima de 5 anos no padrão da OMS).

---

## 🔬 Base metodológica

Os escores-z seguem o **método LMS**, oficial da Organização Mundial da Saúde:

$$z = \frac{(X/M)^L - 1}{L \cdot S} \quad \text{(ou } z = \frac{\ln(X/M)}{S} \text{ quando } L = 0\text{)}$$

onde **L, M, S** são parâmetros por sexo e idade (ou comprimento/estatura), extraídos das tabelas oficiais da OMS.

Detalhes que reproduzem fielmente o WHO Anthro / AnthroPlus:

- **Idade:** de 0 a 5 anos, calculada **em dias** (WHO Anthro 2006); de 5 a 19 anos, **em meses** (WHO Growth Reference 2007).
- **Ajuste da OMS para |z| > 3** nos índices baseados em peso (IMC/idade, Peso/idade, Peso/estatura), evitando extrapolar além dos dados observados.
- **Ajuste deitado/em pé (±0,7 cm):** aplicado quando a forma de medir diverge do padrão para a idade (< 2 anos deitado; ≥ 2 anos em pé).
- **Só funções clássicas do Excel** (ÍNDICE/CORRESP/SE) → funciona também em versões antigas, sem macros.

### Fontes dos dados
- **WHO Child Growth Standards (2006)** — 0 a 5 anos — pacote oficial [`anthro`](https://github.com/worldhealthorganization/anthro).
- **WHO Growth Reference (2007)** — 5 a 19 anos — pacote oficial [`anthroplus`](https://github.com/worldhealthorganization/anthroplus).

Os arquivos-fonte originais da OMS estão versionados em `tabelas_oms/raw/` (procedência).

### Validação
- A lógica de referência (`zscore_ref.py`) é **idêntica** a uma implementação validada contra o pacote **R oficial** da OMS em milhares de casos.
- **16 casos-teste** foram recalculados no **Excel real** e batem com a referência, incluindo extremos (|z|>3) e bordas de idade (0, 24, 60, 61, 120 e 228 meses).

---

## 🗂️ Estrutura do projeto

| Arquivo | Descrição |
|---|---|
| `Calculadora_OMS.xlsx` | **Entregável** — a calculadora (tabelas da OMS + fórmulas) |
| `BD_AvaliacaoNutricional.xlsx` | **Entregável** — o banco de dados (vazio, template) |
| `Guia rápido — Avaliação Nutricional.docx` | Guia de uso em linguagem simples para a equipe |
| `tabelas_oms/*.csv` | Tabelas LMS limpas usadas na construção |
| `tabelas_oms/raw/*.txt` | Arquivos-fonte originais da OMS (procedência) |
| `montar_tabelas.py` | Monta as tabelas LMS a partir dos arquivos da OMS |
| `zscore_ref.py` | Implementação de referência (fonte única da lógica) |
| `build_planilha.py` | Gera os dois arquivos `.xlsx` |
| `configurar_powerquery.ps1` | Configura a ligação automática banco → calculadora |
| `validate_ref.py` · `verify_*.py` | Validação (referência vs R; Excel vs referência) |

### Reconstruir do zero
```bash
python montar_tabelas.py     # monta as tabelas LMS
python build_planilha.py     # gera os dois .xlsx
powershell -ExecutionPolicy Bypass -File configurar_powerquery.ps1   # liga o Power Query
```

---

## 📖 Como citar

Se este projeto for útil no seu trabalho, uma citação é muito bem-vinda 🙏 (veja o botão **"Cite this repository"** no GitHub, ou o arquivo [`CITATION.cff`](CITATION.cff)):

> Saade, I. (2026). *Avaliação Nutricional Infantil — OMS / SISVAN* [software]. https://github.com/isasaade-23/avaliacao-nutricional-oms

---

## 🔒 Privacidade (LGPD)

São **dados de saúde de menores**. Recomenda-se usar o **ID/Prontuário** como chave, tratar o **Nome** como opcional e guardar o banco em **local restrito e controlado**. A ferramenta funciona **inteiramente offline** — nada é enviado para a internet.

---

## 📄 Licença

Distribuído sob a **[Licença MIT](LICENSE)**: você pode **usar, copiar, modificar, distribuir e vender**, inclusive comercialmente, gratuitamente. Só peço que **mantenha o aviso de autoria** e, se puder, **cite o projeto** — vai me ajudar muito. 💚
