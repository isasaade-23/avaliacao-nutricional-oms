<div align="center">

# Avaliação Nutricional Infantil, padrão OMS e SISVAN

**Duas planilhas em Excel, sem macros, que calculam os escores-z da OMS e o diagnóstico nutricional de crianças e adolescentes de 0 a 19 anos.** Apoiam o acompanhamento em creches, escolas e serviços de saúde.

🇧🇷 **Português** · [🇺🇸 English](README.en.md) · [🇪🇸 Español](README.es.md)

[![Licença CC BY-NC-SA 4.0](https://img.shields.io/badge/licen%C3%A7a-CC%20BY--NC--SA%204.0-green)](https://creativecommons.org/licenses/by-nc-sa/4.0/)
![Feito com Python](https://img.shields.io/badge/feito%20com-Python-blue)
![Padrão OMS](https://img.shields.io/badge/padr%C3%A3o-WHO%20Anthro%20%2F%20AnthroPlus-orange)
![Sem macros](https://img.shields.io/badge/Excel-sem%20macros-lightgrey)
[![Sponsor](https://img.shields.io/badge/apoiar-GitHub%20Sponsors-EA4AAA?logo=github-sponsors)](https://github.com/sponsors/isasaade-23)
[![Ko-fi](https://img.shields.io/badge/apoiar-Ko--fi-FF5E5B?logo=ko-fi&logoColor=white)](https://ko-fi.com/ailuciola)

[![Baixar tudo em ZIP](https://img.shields.io/badge/baixar-arquivo%20ZIP-2ea44f?style=for-the-badge&logo=github)](https://github.com/isasaade-23/avaliacao-nutricional-oms/archive/refs/heads/main.zip)

</div>

---

## Sumário

- [Para que serve](#para-que-serve)
- [Como baixar](#como-baixar)
- [Como usar](#como-usar)
- [Recursos](#recursos)
- [Índices e classificação](#índices-e-classificação)
- [Perguntas frequentes](#perguntas-frequentes)
- [Como citar](#como-citar)
- [Privacidade](#privacidade)
- [Licença](#licença)
- [Detalhes técnicos](#detalhes-técnicos)
- [Quem produziu](#quem-produziu)
- [Tem alguma sugestão?](#tem-alguma-sugestão)

---

## Para que serve

Acompanhar o estado nutricional de crianças ao longo do tempo costuma ser trabalhoso, pois exige consultar as tabelas da OMS, calcular escores-z à mão e classificar o resultado segundo o SISVAN. Por isso, muitas vezes esse acompanhamento não é feito, ou é feito com erro.

Nesta ferramenta, você digita peso, altura e datas em um banco de dados, e a calculadora devolve os escores-z e o diagnóstico nutricional na hora, já com cores. Também dá para acompanhar a mesma criança ao longo dos meses e ver a curva de crescimento, além de ver o resumo da turma ou da unidade.

Não é preciso instalar nada além do Excel. A planilha não usa macros e não envia nenhum dado para a internet.

---

## Como baixar

Esta seção explica como baixar os arquivos, mesmo para quem nunca usou o GitHub antes.

Clique no botão verde "baixar arquivo ZIP" no topo desta página. O navegador salva um arquivo chamado `avaliacao-nutricional-oms-main.zip` na pasta Downloads do computador.

Esse arquivo vem compactado, então é preciso extraí-lo antes de usar. No Windows, clique com o botão direito sobre o arquivo baixado e escolha "Extrair tudo". No Mac, basta dar dois cliques sobre o arquivo, e o sistema extrai sozinho.

Dentro da pasta extraída, só dois arquivos interessam para o uso diário. São eles `BD_AvaliacaoNutricional.xlsx` e `Calculadora_OMS.xlsx`. Guarde os dois juntos, como explica a seção seguinte.

---

## Como usar

Guarde os dois arquivos juntos em uma pasta local do computador, como `C:\AvaliacaoNutricional`. Evite o OneDrive e o Google Drive, porque o Excel só consegue ligar as duas planilhas quando elas estão em uma pasta local.

Abra o `BD_AvaliacaoNutricional.xlsx` e digite as medições na aba Dados, uma linha para cada medição. Use sempre o mesmo ID ou número de prontuário para acompanhar a mesma criança ao longo do tempo. Depois salve e feche o arquivo.

Abra o `Calculadora_OMS.xlsx`. Na primeira vez, clique em "Habilitar Conteúdo" na barra amarela do topo. Os resultados aparecem na aba CÁLCULO, já com cores. A aba FICHA mostra o histórico de uma criança, e a aba PAINEL mostra o resumo da turma.

---

## Recursos

- Quatro índices da OMS, que são IMC por idade, Estatura por idade, Peso por idade e Peso por estatura.
- Diagnóstico automático com cores, seguindo o SISVAN e o Ministério da Saúde. O verde indica estado adequado, o amarelo indica atenção e o vermelho indica déficit ou excesso.
- Histórico de cada criança, com a curva de escores-z ao longo do tempo.
- Painel com o resumo da turma ou da unidade.
- Ligação automática entre o banco e a calculadora, sem precisar copiar dados na mão.
- Fórmulas conferidas contra o pacote R oficial da OMS.

---

## Índices e classificação

As categorias seguem os cortes do SISVAN e do Ministério da Saúde.

| Índice | Faixa etária | Categorias (por escore-z) |
|---|---|---|
| IMC/idade, para menores de 5 anos | 0 a 60 meses (0 a 5 anos) | < -3 magreza acentuada · -3 a < -2 magreza · -2 a +1 eutrofia · > +1 a +2 risco de sobrepeso · > +2 a +3 sobrepeso · > +3 obesidade |
| IMC/idade, de 5 a 19 anos | 61 a 228 meses (5 a 19 anos) | < -3 magreza acentuada · -3 a < -2 magreza · -2 a +1 eutrofia · > +1 a +2 sobrepeso · > +2 a +3 obesidade · > +3 obesidade grave |
| Estatura/idade | 0 a 228 meses (0 a 19 anos) | < -3 muito baixa · -3 a < -2 baixa · ≥ -2 adequada |
| Peso/idade | 0 a 120 meses (0 a 10 anos) | < -3 muito baixo · -3 a < -2 baixo · -2 a +2 adequado · > +2 elevado |
| Peso/estatura | 0 a 60 meses (0 a 5 anos) | mesmos cortes do IMC/idade para menores de 5 anos |

Fora dessas faixas, a planilha mostra um traço no lugar do resultado, porque o Peso/idade não existe acima de 10 anos e o Peso/estatura não existe acima de 5 anos no padrão da OMS.

---

## Perguntas frequentes

<details>
<summary>Não sei extrair um arquivo ZIP, e agora</summary>

No Windows, clique com o botão direito sobre o arquivo baixado e escolha "Extrair tudo". No Mac, basta dar dois cliques sobre o arquivo, e o sistema extrai sozinho. Depois é só abrir a pasta extraída e localizar os dois arquivos `.xlsx`.
</details>

<details>
<summary>A calculadora não encontra o banco de dados</summary>

Isso costuma acontecer quando os dois arquivos não estão na mesma pasta, ou quando essa pasta está sincronizada com o OneDrive ou o Google Drive. Mova os dois arquivos juntos para uma pasta local, feche os dois arquivos e abra a calculadora de novo. Em seguida, na aba Dados do Excel, clique em "Atualizar Tudo".
</details>

<details>
<summary>Não tenho o Microsoft Excel instalado</summary>

A planilha usa recursos do Excel, como o Power Query, que podem não funcionar em outros programas de planilha. Recomendamos o Microsoft Excel, mesmo em uma versão antiga, para garantir que tudo funcione como esperado.
</details>

<details>
<summary>O Excel mostra uma barra amarela de aviso ao abrir a calculadora</summary>

Essa barra aparece porque o arquivo se conecta a outro arquivo do computador. Clique em "Habilitar Conteúdo" para permitir a conexão. O arquivo não usa macros nem se conecta à internet.
</details>

---

## Como citar

Se este projeto ajudar no seu trabalho, uma citação é bem-vinda. Use o botão "Cite this repository" no GitHub, ou o arquivo [`CITATION.cff`](CITATION.cff).

> Saade, I. (2026). *Avaliação Nutricional Infantil, OMS e SISVAN* [software]. https://github.com/isasaade-23/avaliacao-nutricional-oms

Apoiar o projeto é opcional, pelo [GitHub Sponsors](https://github.com/sponsors/isasaade-23), pelo [Ko-fi](https://ko-fi.com/ailuciola) ou por Pix, com a chave isasaade23@gmail.com.

---

## Privacidade

São dados de saúde de menores. Use o ID ou o número de prontuário como chave, trate o nome como um campo opcional e guarde o banco em um local restrito. A ferramenta funciona inteiramente offline e não envia nenhum dado para a internet.

---

## Licença

Este projeto é distribuído sob a licença [CC BY-NC-SA 4.0](LICENSE). Você pode usar, copiar, modificar e distribuir o material livremente, inclusive em creches, escolas e serviços de saúde, desde que mantenha o aviso de autoria e não venda o material. Se você adaptar o projeto, redistribua a adaptação sob a mesma licença. 💚

---

## Detalhes técnicos

<details>
<summary>Base metodológica, fontes dos dados e validação</summary>

Os escores-z seguem o método LMS, que é o método oficial da Organização Mundial da Saúde.

$$z = \frac{(X/M)^L - 1}{L \cdot S} \quad \text{(ou } z = \frac{\ln(X/M)}{S} \text{ quando } L = 0\text{)}$$

Nessa fórmula, L, M e S são parâmetros que variam por sexo e idade, ou por sexo e comprimento ou estatura, e vêm das tabelas oficiais da OMS.

A planilha reproduz fielmente o WHO Anthro e o WHO AnthroPlus nos pontos a seguir. A idade é calculada em dias entre 0 e 5 anos, seguindo o WHO Anthro de 2006, e em meses entre 5 e 19 anos, seguindo o WHO Growth Reference de 2007. Os índices baseados em peso, que são IMC por idade, Peso por idade e Peso por estatura, recebem o ajuste da OMS quando o escore-z ultrapassa 3 em módulo, para evitar extrapolar além dos dados observados. Quando a forma de medir diverge do padrão para a idade, a planilha aplica o ajuste de 0,7 cm da OMS, usado quando uma criança com menos de 2 anos é medida em pé, ou quando uma criança com 2 anos ou mais é medida deitada. Por fim, a planilha usa apenas funções clássicas do Excel, como ÍNDICE, CORRESP e SE, e por isso funciona também em versões antigas, sem macros.

Os dados de 0 a 5 anos vêm do WHO Child Growth Standards de 2006, publicado no pacote oficial [`anthro`](https://github.com/worldhealthorganization/anthro). Os dados de 5 a 19 anos vêm do WHO Growth Reference de 2007, publicado no pacote oficial [`anthroplus`](https://github.com/worldhealthorganization/anthroplus). Os arquivos originais da OMS ficam versionados na pasta `tabelas_oms/raw`, para preservar a procedência dos dados.

A lógica de referência, no arquivo `zscore_ref.py`, é idêntica a uma implementação já validada contra o pacote R oficial da OMS em milhares de casos. Além disso, 16 casos de teste foram recalculados no Excel real e bateram com a referência, incluindo valores extremos e as bordas de idade em 0, 24, 60, 61, 120 e 228 meses.
</summary>
</details>

<details>
<summary>Estrutura do projeto e como reconstruir os arquivos</summary>

| Arquivo | Descrição |
|---|---|
| `Calculadora_OMS.xlsx` | Arquivo entregável, com as tabelas da OMS e as fórmulas |
| `BD_AvaliacaoNutricional.xlsx` | Arquivo entregável, o banco de dados vazio, pronto para uso |
| `Guia rápido — Avaliação Nutricional.docx` | Guia de uso em linguagem simples para a equipe |
| `tabelas_oms/*.csv` | Tabelas LMS já limpas, usadas na construção da planilha |
| `tabelas_oms/raw/*.txt` | Arquivos originais da OMS, mantidos para procedência |
| `montar_tabelas.py` | Monta as tabelas LMS a partir dos arquivos da OMS |
| `zscore_ref.py` | Implementação de referência, fonte única da lógica de cálculo |
| `build_planilha.py` | Gera os dois arquivos `.xlsx` |
| `configurar_powerquery.ps1` | Configura a ligação automática entre o banco e a calculadora |
| `validate_ref.py` e `verify_*.py` | Scripts de validação, comparam a referência com o R e o Excel com a referência |

```bash
python montar_tabelas.py     # monta as tabelas LMS
python build_planilha.py     # gera os dois .xlsx
powershell -ExecutionPolicy Bypass -File configurar_powerquery.ps1   # liga o Power Query
```
</details>

---

## Quem produziu

Este projeto foi desenvolvido por Isabela Venancio, com auxílio do modelo Sonnet da Anthropic para parte do código e da documentação.

- [Currículo Lattes](http://lattes.cnpq.br/7006765766090773)
- [LinkedIn](https://www.linkedin.com/in/isabela-venancio-67530a260/)

## Tem alguma sugestão?

Me mande um e-mail. [isasaade23@gmail.com](mailto:isasaade23@gmail.com)
