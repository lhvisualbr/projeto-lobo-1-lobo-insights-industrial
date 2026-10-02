# Limitações — V0.1 a V0.2

> Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.

Ler este documento antes de citar qualquer número do projeto.

## Sobre os dados

1. **Dados sintéticos.** Tudo foi gerado por script (`scripts/gerar_dados.py`, semente fixa) a partir de pesos e sorteios simples. Nenhuma operação real foi observada; portanto nenhum resultado descreve uma empresa real.
2. **Período de apenas três meses** (2026-06-01 a 2026-08-31). Não há sazonalidade, ciclo anual nem base para afirmar tendência. Diferenças entre os meses (por exemplo, agosto tem menos movimentações que junho e julho) vêm do **sorteio aleatório**, não de um fenômeno operacional; não devem ser interpretadas como queda de demanda.
3. **Custos fictícios.** `Custo_Unitario` e a moeda são inventados; o "custo estimado consumido" é apenas Quantidade x Custo_Unitario e não reflete preço real, frete, imposto ou desconto.
4. **Estoque como snapshot simplificado.** `estoque_ficticio.csv` é uma foto única (considerada ao final do período, mas a data não é um campo). Não há histórico de estoque, entradas, saídas reconciliadas nem relação contábil entre o consumo dos três meses e o saldo atual.
5. **Ausência de compras reais.** Não existem pedidos, notas fiscais, recebimentos ou preços negociados. `Lead_Time_Dias` é um valor fictício por item.
6. **Ausência de contratos reais.** Fornecedores são nomes genéricos ("Fornecedor Alfa"…) sem contrato, prazo acordado ou desempenho.
7. **Movimentações somente de Consumo.** O tipo "Ajuste" é permitido pela regra, mas não há nenhum registro dele na V0.1; os KPIs já filtram `Tipo_Movimentacao = "Consumo"`.
8. **Dados repetidos por desenho.** Nome, categoria, unidade e custo aparecem em mais de uma tabela (exigência do enunciado). Isso foi mantido e é conferido pelas verificações Q14–Q23; em um sistema real, seria normalizado.

## Sobre a análise

9. **Ausência de previsão.** Não há forecast, modelo estatístico, série temporal ou cálculo de reposição preditiva. O status de estoque (REPOR / ATENÇÃO / OK) compara apenas o **estado atual** com o mínimo. Essas questões pertencem ao Projeto 2 (Lobo Forecast AI Industrial).
10. **Regra ATENÇÃO é uma convenção do estudo:** Atual maior que o Mínimo e até 25% acima dele (parâmetro editável). Não é norma industrial.
11. **"Itens críticos" = Criticidade Alta e Estoque_Atual menor ou igual ao Mínimo** (somente REPOR). Itens de criticidade alta em ATENÇÃO aparecem como indicador informativo à parte.

## Sobre a ferramenta

12. **Tabelas dinâmicas não foram criadas.** Os resumos usam fórmulas SOMASES/CONT.SES.
13. **Testado com openpyxl e LibreOffice, não com Microsoft Excel.** Ver a seção "NÃO TESTADO" em `TEST_REPORT.md`.
14. **CSV com ponto decimal e separador `;`.** Em Excel com configuração regional brasileira, importar pelo assistente de texto/CSV escolhendo a localidade correta; o XLSX já vem pronto.

## Uso

15. **Nenhum uso operacional real.** O material serve a estudo e portfólio. Não deve orientar compra, liberação de ferramenta ou qualquer decisão em uma operação verdadeira.

## Sobre o ambiente de execução (adicionado na V0.1.1)

16. **Ambiente verificado:** Linux x86_64, Python 3.12.3, LibreOffice 24.2. Windows, macOS e outras versões de Python e das bibliotecas não foram testados. Reconstruir o XLSX e rodar os testes de Excel por fórmulas exige o LibreOffice (`soffice` no `PATH` ou `LOBO_SOFFICE`); sem ele, os CSVs, as verificações de qualidade em Python e a maior parte da suíte funcionam, e os testes dependentes ficam SKIP.

## Sobre a V0.2 — Analytics & Dashboard (adicionado na V0.2)

17. **Não testado no Microsoft Excel real.** O XLSX V0.2 foi gerado com openpyxl e recalculado/renderizado com LibreOffice. Aparência e comportamento dos 9 gráficos no Excel (fontes, espaçamentos, rótulos, eixo secundário do Pareto) não foram verificados; a inspeção visual foi feita sobre a renderização do LibreOffice e é parcial (`docs/manual/inspecao_visual_v0_2.md`).
18. **Sem filtros nem segmentações no DASHBOARD.** Foram avaliados e descartados por critério técnico (`docs/V0_2_NOTES.md`, seção 3, item 9). As tabelas de dados têm filtro de coluna; três parâmetros (limite do Pareto, Top N, margem de ATENÇÃO) são editáveis.
19. **Ainda não há previsão.** A V0.2 não projeta demanda, não calcula tendência futura nem quantidade de compra. Variações mensais, razões Atual ÷ Mínimo, Pareto e a fila de atenção descrevem o passado e o estado atual dos dados sintéticos; **nenhum resultado é recomendação operacional real**.
20. **Diferenças entre meses vêm do sorteio dos dados** (limitação 2), não de fenômeno operacional. As variações "vs mês anterior" mostradas no painel não indicam tendência.
21. **Pareto e concentração são descritivos.** O limite de 80% é uma convenção editável do estudo, não uma regra de gestão; "Núcleo" só significa "conjunto de materiais que atinge o limite neste período".
22. **Unidades de medida heterogêneas.** "Unidades consumidas" soma quantidades de itens em UN, CX, PCT e PAR (definição da V0.1); médias como "unidades por movimentação" misturam essas unidades.
23. **Estrutura fixa.** O layout assume 3 meses, 6 categorias, 5 centros e 20 materiais; acrescentar itens exige regenerar o arquivo. O painel de estoque exibe 10 itens e informa quantos ficaram fora.
24. **Localidade e pré-visualização.** Fórmulas em inglês, traduzidas pelo Excel; textos de leitura usam `ROUND`/`FIXED` para respeitar o separador decimal, mas isso só foi verificado no LibreOffice em inglês. Os valores em cache (o que um visualizador mostra antes de recalcular) usam ponto decimal, e os gráficos não têm cache próprio: visualizadores que não recalculam podem omiti-los.
25. **Testes de renderização** dependem de LibreOffice e poppler (`pdftoppm`, `pdfinfo`); sem eles ficam SKIP, nunca PASS.

## Sobre a V0.2.1 — Packaging & Reproducibility Fix (adicionado na V0.2.1)

26. **Arquivos regeneráveis.** `docs/TEST_REPORT.md` e `images/dashboard_v0_2_pagina_1.png` a `_3.png` são reescritos por uma execução normal da suíte e variam com o ambiente (versões, LibreOffice, fontes); por isso ficam fora do `SHA256SUMS.txt`. O snapshot `RELEASE_SHA256SUMS.txt` só confere integralmente no pacote recém-extraído. A variação de bytes no ambiente do auditor não foi observada nesta versão (`docs/V0_2_1_NOTES.md`, seção 9).
