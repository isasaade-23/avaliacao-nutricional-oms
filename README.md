<div align="center">

# Avaliação Nutricional Infantil, padrão OMS e SISVAN

**Ferramenta para avaliação e acompanhamento nutricional de crianças e adolescentes de 0 a 19 anos, segundo os padrões da OMS e os critérios do SISVAN.**

Um arquivo só, com os dados e os cálculos juntos. Funciona no Excel, no LibreOffice e no Google Sheets, sem macros. É gratuita.

🌐 **Endereço oficial: [isasaade-23.github.io/avaliacao-nutricional-oms](https://isasaade-23.github.io/avaliacao-nutricional-oms/)**

🇧🇷 **Português** · [🇺🇸 English](README.en.md) · [🇪🇸 Español](README.es.md)

[![Licença CC BY-NC-SA 4.0](https://img.shields.io/badge/licen%C3%A7a-CC%20BY--NC--SA%204.0-green)](https://creativecommons.org/licenses/by-nc-sa/4.0/)
![Feito com Python](https://img.shields.io/badge/feito%20com-Python-blue)
![Padrão OMS](https://img.shields.io/badge/padr%C3%A3o-WHO%20Anthro%20%2F%20AnthroPlus-orange)
![Sem macros](https://img.shields.io/badge/planilha-sem%20macros-lightgrey)
[![Sponsor](https://img.shields.io/badge/apoiar-GitHub%20Sponsors-EA4AAA?logo=github-sponsors)](https://github.com/sponsors/isasaade-23)
[![Ko-fi](https://img.shields.io/badge/apoiar-Ko--fi-FF5E5B?logo=ko-fi&logoColor=white)](https://ko-fi.com/ailuciola)

[![Baixar tudo em ZIP](https://img.shields.io/badge/baixar-arquivo%20ZIP-2ea44f?style=for-the-badge&logo=github)](https://github.com/isasaade-23/avaliacao-nutricional-oms/archive/refs/heads/main.zip)

</div>

---

## Comece aqui

Se você só quer usar a ferramenta, não precisa entender GitHub nem programação. Escolha a sua versão, baixe o arquivo e digite as medições na aba Banco. Os resultados aparecem sozinhos nas outras abas. Há também o [Guia rápido em PDF](Guia%20r%C3%A1pido%20%E2%80%94%20Avalia%C3%A7%C3%A3o%20Nutricional.pdf), com o passo a passo em linguagem simples, pronto para imprimir. As seções [Como baixar](#como-baixar) e [Como usar](#como-usar) explicam cada passo em detalhe.

---

## Sumário

- [Para que serve](#para-que-serve)
- [Como baixar](#como-baixar)
- [Como usar](#como-usar)
- [Índices e classificação](#índices-e-classificação)
- [Perguntas frequentes](#perguntas-frequentes)
- [Como citar](#como-citar)
- [Privacidade e licença](#privacidade-e-licença)
- [Detalhes técnicos](#detalhes-técnicos)
- [Quem produziu e contato](#quem-produziu-e-contato)

---

## Para que serve

Acompanhar o estado nutricional de crianças ao longo do tempo costuma ser trabalhoso, pois exige consultar as tabelas da OMS, calcular escores-z à mão e classificar o resultado segundo o SISVAN. Por isso, muitas vezes esse acompanhamento não é feito, ou é feito com erro.

Nesta ferramenta, você digita peso, altura e datas na aba Banco, e a mesma planilha devolve os escores-z e a classificação do estado nutricional na hora, já com cores. Também dá para acompanhar a mesma criança ao longo dos meses, ver a curva de crescimento e ver o resumo da turma ou da unidade.

---

## Como baixar

A calculadora é um arquivo único, em três versões com o mesmo conteúdo. Escolha a que combina com o programa que você usa.

| Versão | Arquivo | Onde baixar |
|---|---|---|
| Excel | `excel_Calculadora_OMS.xlsx` | [baixar](https://github.com/isasaade-23/avaliacao-nutricional-oms/raw/main/excel_Calculadora_OMS.xlsx) |
| LibreOffice | `libreoffice_Calculadora_OMS.ods` | [baixar](https://github.com/isasaade-23/avaliacao-nutricional-oms/raw/main/libreoffice_Calculadora_OMS.ods) |
| Google Sheets | `gsheet_Calculadora_OMS` | [criar minha cópia](https://docs.google.com/spreadsheets/d/1ZrYO0hEXhPV054k1nwWsasHnFxFsZ1CQfY2td47YB90/copy) · [ver sem entrar](https://docs.google.com/spreadsheets/d/1ZrYO0hEXhPV054k1nwWsasHnFxFsZ1CQfY2td47YB90/preview) |

No Google Sheets, você pode ver o modelo sem entrar em uma conta. Para preencher e salvar, clique em "criar minha cópia" e entre em uma conta Google, se solicitado. Confirme em "Fazer uma cópia": ela fica no seu Drive, e é nela que você digita os dados. Não é preciso solicitar permissão para editar o modelo original. Prefere usar sem conta Google? Baixe a versão Excel ou LibreOffice pelos links acima.

Também dá para baixar o projeto inteiro pelo botão verde "baixar arquivo ZIP" no topo desta página. Nesse caso o arquivo vem compactado e é preciso extraí-lo antes de usar. No Windows, clique com o botão direito sobre ele e escolha "Extrair tudo". No Mac, basta dar dois cliques.

---

## Como usar

Abra a planilha e vá para a aba **Banco**. Digite as medições, uma linha para cada medição, com ID, sexo, data de nascimento, data da medida, peso e altura. Use sempre o mesmo ID ou número de prontuário para acompanhar a mesma criança ao longo do tempo.

Os resultados aparecem na aba **Cálculo**, já com cores. A aba **Ficha** mostra o histórico de uma criança, e a aba **Painel** mostra o resumo da turma. A aba **Instruções** traz o passo a passo dentro da própria planilha.

Como os dados e os cálculos ficam no mesmo arquivo, não há nada para conectar. Guarde a planilha onde preferir e faça cópias de segurança dela como faria com qualquer documento. No Excel, se aparecer o aviso de modo protegido, clique em "Habilitar edição".

---

## Índices e classificação

A ferramenta calcula quatro índices antropométricos da OMS, que são IMC por idade, Estatura por idade, Peso por idade e Peso por estatura. As categorias de cada um seguem os cortes do SISVAN e do Ministério da Saúde, e aparecem com cor no resultado. O verde indica estado adequado, o amarelo indica atenção e o vermelho indica déficit ou excesso.

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

No Windows, clique com o botão direito sobre o arquivo baixado e escolha "Extrair tudo". No Mac, basta dar dois cliques sobre o arquivo, e o sistema extrai sozinho. Depois é só abrir a pasta extraída e localizar o arquivo da sua versão.
</details>

<details>
<summary>Não tenho o Microsoft Excel instalado</summary>

Use a versão para LibreOffice (`libreoffice_Calculadora_OMS.ods`), que é gratuito, ou [crie sua cópia no Google Sheets](https://docs.google.com/spreadsheets/d/1ZrYO0hEXhPV054k1nwWsasHnFxFsZ1CQfY2td47YB90/copy), entrando em uma conta Google. O conteúdo é o mesmo nas três versões.
</details>

---

## Como citar

Se este projeto ajudar no seu trabalho, uma citação é bem-vinda. Use o botão "Cite this repository" no GitHub, ou o arquivo [`CITATION.cff`](CITATION.cff).

> Venancio da Silva, I. (2026). *Avaliação Nutricional Infantil, OMS e SISVAN* [software]. https://github.com/isasaade-23/avaliacao-nutricional-oms

---

## Privacidade e licença

São dados de saúde de menores. Use o ID ou o número de prontuário como chave, trate o nome como um campo opcional e guarde o arquivo em um local restrito. As versões Excel e LibreOffice funcionam inteiramente offline e não enviam nenhum dado para a internet. Na versão Google Sheets, os dados ficam na nuvem do Google. Faça uma cópia para o seu Drive, não digite nada no modelo original e restrinja o compartilhamento da sua cópia.

Este projeto é distribuído sob a licença [CC BY-NC-SA 4.0](LICENSE). Você pode usar, copiar, modificar e distribuir o material livremente, inclusive em creches, escolas e serviços de saúde, desde que mantenha o aviso de autoria e não venda o material. Se você adaptar o projeto, redistribua a adaptação sob a mesma licença.

---

## Detalhes técnicos

<details>
<summary>Base metodológica, fontes dos dados e validação</summary>

Os escores-z seguem o método LMS, que é o método oficial da Organização Mundial da Saúde.

$$z = \frac{(X/M)^L - 1}{L \cdot S} \quad \text{(ou } z = \frac{\ln(X/M)}{S} \text{ quando } L = 0\text{)}$$

Nessa fórmula, L, M e S são parâmetros que variam por sexo e idade, ou por sexo e comprimento ou estatura, e vêm das tabelas oficiais da OMS.

A planilha reproduz fielmente o WHO Anthro e o WHO AnthroPlus nos pontos a seguir. A idade é calculada em dias entre 0 e 5 anos, seguindo o WHO Anthro de 2006, e em meses entre 5 e 19 anos, seguindo o WHO Growth Reference de 2007. Os índices baseados em peso, que são IMC por idade, Peso por idade e Peso por estatura, recebem o ajuste da OMS quando o escore-z ultrapassa 3 em módulo, para evitar extrapolar além dos dados observados. Quando a forma de medir diverge do padrão para a idade, a planilha aplica o ajuste de 0,7 cm da OMS, usado quando uma criança com menos de 2 anos é medida em pé, ou quando uma criança com 2 anos ou mais é medida deitada. Por fim, a planilha usa apenas funções clássicas, como ÍNDICE, CORRESP e SE, e por isso funciona também em versões antigas do Excel, no LibreOffice e no Google Sheets, sem macros.

Os dados de 0 a 5 anos vêm do WHO Child Growth Standards de 2006, publicado no pacote oficial [`anthro`](https://github.com/worldhealthorganization/anthro). Os dados de 5 a 19 anos vêm do WHO Growth Reference de 2007, publicado no pacote oficial [`anthroplus`](https://github.com/worldhealthorganization/anthroplus). Os arquivos originais da OMS foram mantidos à parte, para preservar a procedência dos dados.

A lógica de referência foi validada contra o pacote R oficial da OMS em milhares de casos. Além disso, 16 casos de teste foram recalculados no Excel real e bateram com a referência, incluindo valores extremos e as bordas de idade em 0, 24, 60, 61, 120 e 228 meses. Os scripts de construção e de validação não ficam neste repositório, mas quem quiser auditar o cálculo pode escrever para o e-mail da seção [Quem produziu e contato](#quem-produziu-e-contato).
</summary>
</details>

---

## Quem produziu e contato

Este projeto foi desenvolvido por Isabela Venancio da Silva, com auxílio do modelo Sonnet da Anthropic para parte do código e da documentação.

- [Currículo Lattes](http://lattes.cnpq.br/7006765766090773)
- [LinkedIn](https://www.linkedin.com/in/isabela-venancio-67530a260/)

Encontrou um problema ou tem uma sugestão? Me mande um e-mail. [isasaade23@gmail.com](mailto:isasaade23@gmail.com)

Apoiar a manutenção do projeto é opcional, pelo [GitHub Sponsors](https://github.com/sponsors/isasaade-23), pelo [Ko-fi](https://ko-fi.com/ailuciola) ou por Pix, com a chave isasaade23@gmail.com.
