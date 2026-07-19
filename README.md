# Avaliação Nutricional Infantil (0 a 19 anos) — OMS / SISVAN

Duas planilhas em Excel (só fórmulas, **sem macros**) para avaliação antropométrica
de crianças e adolescentes em escolas e serviços de saúde:

- **`BD_AvaliacaoNutricional.xlsx`** — Banco de Dados (leve, só dados digitados).
- **`Calculadora_OMS.xlsx`** — Calculadora (tabelas da OMS + fórmulas; escores-z e diagnósticos).

> Mantenha os **dois arquivos na mesma pasta**.

---

## Como usar

1. **Digite as medições** no `BD_AvaliacaoNutricional.xlsx`, aba **Dados**.
   Uma linha por medição. Colunas: ID/Prontuário, Nome, Sexo, Data de nascimento,
   Data da medida, Peso (kg), Altura (cm), Medida (Deitado/Em pé — opcional), Observações.
   Para **acompanhar** uma criança, use sempre o **mesmo ID** e acrescente uma nova
   linha a cada retorno. Salve e feche o banco.

2. **Abra a `Calculadora_OMS.xlsx`.** Ela **puxa o banco automaticamente** (Power Query).
   - Na **primeira vez**, o Excel mostra uma barra amarela: clique em **"Habilitar Conteúdo"**
     (e, se aparecer, em **"Habilitar"** as conexões de dados). Isso é pedido só uma vez.
   - Depois de mudar o banco, atualize em **Dados ▸ Atualizar Tudo** (ou feche e reabra).

   Os escores-z e diagnósticos aparecem na aba **CÁLCULO**, com cores
   (verde = adequado, amarelo = atenção, vermelho = déficit/excesso, "—" = fora da faixa válida).

3. **Aba FICHA:** escolha o ID para ver o histórico e a **curva de evolução** dos escores-z.

4. **Aba CONSULTA:** cálculo avulso de uma criança, sem mexer no banco.

5. **Aba PAINEL:** resumo da turma/unidade (distribuição de IMC/idade e indicadores-chave).

### Por que dois arquivos?
O banco fica **leve** (só dados, cresce com o tempo) e a calculadora (com as tabelas da OMS
e as fórmulas) fica separada. A calculadora **puxa o banco sozinha** via Power Query — você
só digita no banco. A Calculadora acha o banco automaticamente **se os dois estiverem na
mesma pasta** (detecta o caminho sozinha); se precisar mover, **mova os dois juntos**.
A calculadora vem pronta para **1500** medições; para mais, selecione a última linha de
fórmulas da aba Cálculo e arraste para baixo.

> Se um dia mudar o **nome** do arquivo do banco, ajuste a consulta em
> Dados ▸ Consultas e Conexões ▸ (clique direito em *Banco*) ▸ Editar.

---

## Índices e classificação (SISVAN / Ministério da Saúde)

| Índice | Faixa | Categorias (por escore-z) |
|---|---|---|
| IMC/idade < 5 anos | 0–60 m | <−3 Magreza acentuada · −3 a <−2 Magreza · −2 a +1 Eutrofia · >+1 a +2 Risco de sobrepeso · >+2 a +3 Sobrepeso · >+3 Obesidade |
| IMC/idade 5–19 anos | 61–228 m | <−3 Magreza acentuada · −3 a <−2 Magreza · −2 a +1 Eutrofia · >+1 a +2 Sobrepeso · >+2 a +3 Obesidade · >+3 Obesidade grave |
| Estatura/idade | 0–228 m | <−3 Muito baixa · −3 a <−2 Baixa · ≥−2 Adequada |
| Peso/idade | 0–120 m (até 10 a) | <−3 Muito baixo · −3 a <−2 Baixo · −2 a +2 Adequado · >+2 Elevado |
| Peso/estatura | 0–60 m (até 5 a) | mesmos cortes do IMC/idade < 5 anos |

Fora dessas faixas o resultado é "—" (Peso/idade não existe acima de 10 anos e
Peso/estatura acima de 5 anos no padrão OMS).

---

## Metodologia (resumo técnico)

- **Idade:** calculada de Data de nascimento + Data da medida. Para 0–5 anos usa-se a
  idade **em dias** (padrão WHO Anthro); para 5–19 anos, **em meses**.
- **Escore-z (método LMS):** `z = ((X/M)^L − 1) / (L·S)` (ou `ln(X/M)/S` quando L=0),
  com L, M, S das tabelas da OMS por sexo e idade/comprimento.
- **Ajuste da OMS para |z| > 3** nos índices baseados em peso (IMC/idade, Peso/idade,
  Peso/estatura), evitando extrapolação além dos dados observados.
- **Ajuste deitado/em pé (±0,7 cm):** aplicado quando a forma de medida diverge do padrão
  para a idade (<2 anos deitado; ≥2 anos em pé).
- **Só funções clássicas do Excel** (ÍNDICE/CORRESP/SE) nos cálculos → funciona em versões antigas.
- **Sem macros.** A ligação banco→calculadora usa **Power Query** (recurso nativo do Excel,
  não é macro): a aba oculta `Banco` importa o banco e a aba Cálculo a referencia linha a linha.

### Fontes dos dados
- WHO Child Growth Standards (2006) — 0 a 5 anos — pacote oficial `anthro`.
- WHO Growth Reference (2007) — 5 a 19 anos — pacote oficial `anthroplus`.

---

## Arquivos do projeto

| Arquivo | Descrição |
|---|---|
| `Calculadora_OMS.xlsx` | **Entregável** — a calculadora |
| `BD_AvaliacaoNutricional.xlsx` | **Entregável** — o banco de dados |
| `tabelas_oms/*.csv` | Tabelas LMS limpas usadas na construção |
| `tabelas_oms/raw/*.txt` | Arquivos-fonte originais da OMS (provenance) |
| `montar_tabelas.py` | Monta as tabelas LMS a partir dos arquivos da OMS |
| `zscore_ref.py` | Implementação de referência (fonte única da lógica) |
| `build_planilha.py` | Gera os dois arquivos `.xlsx` |
| `configurar_powerquery.ps1` | Configura o Power Query (ligação automática banco→calculadora) |
| `validate_ref.py` | Valida `zscore_ref` vs pacote R oficial / asserts da OMS |
| `verify_prep.py` / `verify_check.py` | Conferem Excel == referência (16 casos-teste) |

### Reconstruir do zero
```
python montar_tabelas.py     # monta as tabelas LMS
python build_planilha.py     # gera os dois .xlsx
powershell -ExecutionPolicy Bypass -File configurar_powerquery.ps1   # liga o Power Query
```
> O `configurar_powerquery.ps1` precisa ser rodado **depois** do `build_planilha.py`
> (que gera o arquivo do zero) e com o **Excel fechado**.

### Validação realizada
- Tabelas LMS conferidas contra valores-âncora publicados pela OMS.
- `zscore_ref` idêntico ao `who_zscore.py` (validado vs R `anthroplus`) em 3360 casos
  (5–19 anos) e às asserções oficiais do pacote `anthro` (0–5 anos).
- 16 casos-teste recalculados no **Excel real**: escores-z e diagnósticos **idênticos**
  à referência, incluindo extremos (|z|>3), bordas de idade (0, 24, 60, 61, 120, 228 meses)
  e ajuste deitado/em pé.

---

## Observações (LGPD)
São dados de saúde de menores. Recomenda-se usar o **ID/Prontuário** como chave, tratar o
**Nome** como opcional e guardar o banco em local restrito/controlado.
