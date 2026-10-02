# Notas da V0.2 — Analytics & Dashboard

> Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.

## 1. Objetivo

Transformar a base técnica validada (V0.1.1, **baseline congelada**) em uma camada analítica e uma experiência de leitura executiva no Excel, para portfólio, demonstração em processo seletivo, apresentação e estudo de análise de dados aplicada ao contexto industrial. A pergunta central continua sendo: **"O que aconteceu na operação e o que precisa de atenção?"**

**Dentro do escopo:** KPIs em cartões, análises operacionais (mês, categoria, centro, material, estoque), Pareto de custo e de consumo, situação de estoque e criticidade, aba `DASHBOARD`, identidade visual, auditabilidade, testes de regressão, documentação.
**Fora do escopo (não feito, de propósito):** previsão de qualquer tipo, tendência futura, IA/LLM, classificador, Power BI, n8n, API, banco de dados, aplicação web, publicação no GitHub e qualquer item do Projeto 2 (Lobo Forecast AI). Nenhuma meta ou percentual sem base nos dados foi inventado: todo percentual do painel é participação sobre o total observado.

## 2. Recursos adicionados

Arquivo novo: `excel/lobo_insights_industrial_v0_2.xlsx` (o XLSX histórico `lobo_insights_industrial_v0_1.xlsx` permanece, sem alteração). Abas, na ordem:

| Aba | Camada | Conteúdo |
|---|---|---|
| **DASHBOARD** (nova, primeira e ativa) | Visualização | Cabeçalho com título, subtítulo e período; faixa de links; 6 cartões de KPI; 9 gráficos; 6 cartões de concentração; fila de atenção de estoque; matriz criticidade × status; leituras geradas por fórmula; rodapé de rastreabilidade e conferências. Impressão em 3 páginas A4 paisagem. |
| LEIA_ME (reescrita) | Documentação | Camadas, como ler o painel, decisões, legenda, SHA-256 dos CSVs. |
| Indicadores | Indicadores | Os 6 KPIs oficiais e resumos da V0.1.1 — fórmulas **idênticas**; só o texto de A1 e A3 mudou. |
| **Analises** (nova) | Cálculo | Evolução mensal com variação; categorias e centros (base + ordenados por unidades e por custo); base de materiais; **Pareto de custo** e **Pareto de consumo**; concentração; estoque (base de prioridade, fila, matriz, painel); **18 conferências** de fechamento. |
| Materiais, Estoque, Consumo | Dados | Idênticas à V0.1.1 (célula a célula). |
| Qualidade | Validação | As 36 verificações da V0.1.1, inalteradas. |
| Parametros | Parâmetros | V0.1.1 (A1:F18) inalterada + 3 parâmetros novos em A20:C24: `Limite_Pareto` (80%), `Top_N_A` (3), `Top_N_B` (5). |

**Cadeia de rastreabilidade (só fórmulas):** `tbConsumo / tbEstoque / tbMateriais → Indicadores → Analises → DASHBOARD`. A aba Analises **não recalcula** somas por material/categoria/mês: lê os resumos de `Indicadores` e apenas ordena, acumula e compara — evita duplicar lógica. Todo cartão do DASHBOARD é uma referência a `Indicadores` (KPIs) ou a `Analises` (demais).

**Onde cada pergunta do enunciado é respondida**

| Pergunta | Onde |
|---|---|
| Maior consumo por material | DASHBOARD §4, Pareto de consumo (e `Analises` 5b) |
| Maior custo por material | DASHBOARD §4, Pareto de custo (e `Analises` 5a) |
| Evolução em três meses | DASHBOARD §1 (unidades e custo por mês, com variação vs mês anterior) |
| Categorias: volume e custo | DASHBOARD §2 (dois gráficos ordenados) |
| Centros: consumo e custo | DASHBOARD §3 (dois gráficos ordenados) |
| Materiais em REPOR / ATENÇÃO | DASHBOARD §5, fila de atenção e gráfico Atual × Mínimo |
| Críticos que demandam atenção | KPI 6, texto vermelho na fila e matriz (Alta × REPOR) |
| Concentração em poucos materiais | DASHBOARD §4, seis cartões (nº de materiais para o limite; participação dos Top N) |

### Dicionário da camada nova (aba Analises)

| Bloco | Colunas principais | Regra |
|---|---|---|
| 1. Mensal | Mês, Rótulo (`Jun/2026`), Movimentações, Unidades, Custo, Var. unidades, Var. custo | Lê `Indicadores` seção 2; variação = valor ÷ valor do mês anterior − 1; `n/d` se o anterior for 0; 1º mês sem variação. |
| 2–3. Categoria / Centro | Base (Mov., Unid., Custo, %, posição por unidades e por custo) e duas tabelas ordenadas | Posição = `RANK` + `COUNTIF` das linhas acima com o mesmo valor (**desempate = ordem original**). |
| 4. Materiais | Material_ID, Material, Categoria, Criticidade, Mov., Unid., Custo, posições | Lê `Indicadores` seção 5. |
| 5a/5b. Pareto | Posição, ID, Material, valor, % do total, % acumulado, limite, Faixa (`Núcleo`/`Cauda`), colunas de série | % sobre o KPI oficial; `Núcleo` = % acumulado **antes** do item < limite (o item que cruza o limite entra). |
| 6. Concentração | Nº de materiais para o limite (e % dos movimentados); participação dos Top N; maior categoria/centro | Lê o Pareto e as tabelas ordenadas; `Top N` e limite vêm de `Parametros`. |
| 7a–7d. Estoque | Base com chave de prioridade; fila; matriz criticidade × status; painel (10 itens) | Status e `Critico_Em_Risco` vêm de `tbEstoque` (regras da V0.1.1). Ordem da fila: REPOR → ATENÇÃO → OK; dentro do status, criticidade Alta → Média → Baixa; depois menor razão Atual ÷ Mínimo. |
| 8. Conferências | 18 linhas "Valor A = Valor B" | Somas dos resumos = KPIs; Pareto fecha em 100%; posições únicas; matriz = KPIs; fila sem inversões; cartões do DASHBOARD = KPIs. |

## 3. Decisões analíticas

1. **KPIs preservados.** Os 6 KPIs oficiais e seus significados não mudaram (`Indicadores`). Os cartões só os exibem, com o título oficial.
2. **Ordenações sem `SORT`/`FILTER`/`XLOOKUP`.** Essas funções não são portáveis (LibreOffice, Excel antigo) e/ou dependem de matrizes dinâmicas. Foi usado `RANK` + `COUNTIF` + `INDEX/MATCH` (funções que já eram usadas na V0.1.1). Um teste garante que nenhuma função dinâmica entrou no arquivo.
3. **Empates reais tratados.** Há empates nas unidades por material (MAT011, MAT012 e MAT020 = 11; MAT006 e MAT018 = 5). A regra de desempate (ordem do catálogo) é determinística e testada contra um cálculo independente.
4. **Pareto.** Percentuais sobre o total **oficial** (KPI), não sobre a soma da própria tabela; o fechamento em 100% é uma das conferências. O limite (80%) é **parâmetro editável**, apresentado como convenção do estudo, e não como regra de gestão.
5. **Concentração = números, sem juízo.** O painel informa, por exemplo, quantos materiais somam o limite e a participação dos Top N. Não afirma que a concentração seja "alta" ou "problemática".
6. **Variação mensal é descritiva.** Comparar mês com mês anterior é permitido; **não há tendência, extrapolação nem previsão**. O painel repete o aviso de que diferenças entre meses vêm do sorteio dos dados sintéticos (já registrado em `LIMITATIONS.md`, item 2).
7. **Fila de estoque é ordenação de apresentação.** A ordem de exibição (status, criticidade, razão Atual ÷ Mínimo) organiza a leitura; **não** é prioridade de compra, não sugere quantidade e não prevê reposição. A razão Atual ÷ Mínimo é apenas a razão observada.
8. **Leituras por fórmula.** Textos como "Maior consumo: Jul/2026 (691 un.)" são fórmulas que citam o dado; nada é escrito à mão. Para o separador decimal seguir a localidade do Excel, usam-se `ROUND` e `FIXED` (não `TEXT` com máscara numérica).
9. **Interatividade (filtros/segmentações) — avaliada e NÃO implementada.**
   - *Segmentações nativas* dependem de tabelas dinâmicas ou de XML que o openpyxl não gera, e não puderam ser verificadas sem o Excel real.
   - *Listas suspensas com fórmulas* mudariam os painéis sem mudar os 6 KPIs (ou mudariam os KPIs, quebrando a regressão "1.647 / 24.606,20…"), e impressões/PDFs mostrariam um estado escolhido, não o oficial.
   - Com 3 meses, 6 categorias, 5 centros e 20 materiais, todas as dimensões cabem lado a lado, sem filtro. O ganho analítico não compensava o risco. Em vez disso: cada dimensão aparece em unidades **e** custo; há tabelas ordenadas em `Analises`; as tabelas de dados têm filtro por coluna; e três parâmetros (limite do Pareto, Top N, margem de ATENÇÃO) são editáveis.
10. **Dados, regras e fórmulas da V0.1.1 intocados.** As novas colunas/blocos vivem só na aba nova; `lobo_common.py`, `gerar_dados.py`, `validar_dados.py`, `construir_excel.py` e `executar_testes.py` estão byte a byte iguais (verificação por hash).

## 4. Decisões visuais

- **Identidade:** azul-marinho (herdado da V0.1) com filete **cobre** como acento industrial; cinza-azulado para texto secundário; fundo cinza muito claro com painéis brancos; **Arial** em todo o painel. Sem gradientes, sem 3D, sem sombras, sem ícones decorativos.
- **Semântica de cor consistente:** azul = unidades; cobre = custo; cinza = referência (estoque mínimo, "demais" no Pareto); vermelho/âmbar/verde **somente** para status de estoque e alertas; vermelho tracejado = limite do Pareto.
- **Hierarquia:** cabeçalho → 6 KPIs → seções numeradas (1 a 5). Cada seção compara unidades × custo lado a lado; cada gráfico tem título em célula (controlável e imprimível), rótulos de dados e uma linha de leitura por fórmula.
- **Grade:** 12 colunas de largura igual (largura útil ≈ 1.180 px a 100% de zoom, adequada a notebook); cartões de 2 colunas; painéis duplos de 6 colunas. Gráficos ancorados às células (`twoCellAnchor`), então acompanham a grade.
- **Ordem da narrativa:** o que aconteceu (mensal, categoria, centro, material) → o que precisa de atenção (estoque e criticidade). Os cartões 5 e 6 (REPOR e críticos) ficam em vermelho quando > 0, já no topo.
- **Navegação e impressão:** faixa de hiperlinks **internos**; sem congelamento de painéis no DASHBOARD (rolagem livre — em notebook, cabeçalho congelado consumiria altura útil); 3 páginas A4 paisagem com quebras entre seções; rodapé com aviso de dados sintéticos.
- **Sem valores decorativos:** nenhuma meta, semáforo de desempenho ou percentual sem base.

## 5. Limitações

Além das limitações da V0.1/V0.1.1 (`docs/LIMITATIONS.md`, itens 1–16):

1. **Não testado no Microsoft Excel real.** Gerado com openpyxl, recalculado e renderizado com LibreOffice. Aparência dos gráficos no Excel (espaçamentos, fontes, rótulos, eixo secundário) não foi verificada.
2. **Excel em pt-BR e outras localidades não verificado.** Fórmulas são gravadas em inglês; o valor em cache (o que um visualizador mostra antes de recalcular) usa ponto decimal.
3. **Gráficos sem cache interno.** O arquivo grava valores em cache nas células, mas não dentro do XML dos gráficos; visualizadores que não recalculam podem mostrar os números e omitir os gráficos.
4. **Estrutura fixa.** O layout da aba Analises e do DASHBOARD assume 3 meses, 6 categorias, 5 centros e 20 materiais. Acrescentar materiais/categorias exige regenerar o arquivo (`construir_excel_v0_2.py`) e ajustar a estrutura, como já ocorria na V0.1.1 para as linhas do resumo por material. O painel de estoque exibe 10 itens e informa quantos ficaram de fora.
5. **Unidades de medida heterogêneas.** "Unidades consumidas" soma `Quantidade` de itens com unidades UN, CX, PCT e PAR (definição herdada da V0.1). Médias como "unidades por movimentação" misturam essas unidades e são apenas descritivas.
6. **Sem filtros/segmentações** (ver decisão 9 da seção 3).
7. **Sem previsão.** Variações mensais e razões Atual ÷ Mínimo descrevem o passado/estado atual; não devem ser extrapoladas nem lidas como recomendação operacional.
8. **Ferramentas de teste:** os testes de renderização usam `pdftoppm`/`pdfinfo` (poppler) e o LibreOffice; sem eles ficam **SKIP** (nunca PASS). Não foi adicionada nenhuma biblioteca Python nova (`requirements.txt` inalterado).
9. **Inspeção visual parcial** (ver `docs/manual/inspecao_visual_v0_2.md`).
10. **Ambiente verificado:** Linux x86_64, Python 3.12.3, LibreOffice 24.2 (mesmas limitações de ambiente da V0.1.1).
11. **O XLSX não é byte a byte reproduzível** quando reconstruído (metadados variam); a reprodução é comprovada por comparação de fórmulas e valores de todas as células (teste "Reconstrução / mutação").

## 6. Testes executados

Executor: `python scripts/executar_testes_v0_2.py --verificar`. Ele importa e roda **inteira e sem alteração** a suíte da V0.1.1 (T01–T103) e acrescenta 71 testes (T104–T174). Resultado da última execução: **174 testes — 174 PASS, 0 FAIL, 0 SKIP**; 36/36 verificações de qualidade dos CSVs (Python e Excel); 36/36 testes de detecção por injeção de erros da V0.1.1. Relatório completo: `docs/TEST_REPORT.md`.

Testes novos, por área: regressão contra a baseline (23 arquivos protegidos, CSVs, XLSX histórico); os 6 KPIs (referência × V0.1.1 × Indicadores × cartão × CSV independente); estrutura do XLSX (abas, tabelas, nomes, sem vínculos externos, sem caminhos absolutos, hiperlinks internos); preservação célula a célula das abas de dados; fórmulas (erros, cache, ausência de números digitados, ausência de funções de previsão e de funções não portáveis, cadeia de leitura entre abas); **valores de todas as tabelas da Analises contra um cálculo independente em Python (Decimal)**; textos de leitura; 9 gráficos (tipo, séries, pontos, intervalos, ausência de `#REF!`); layout e impressão; **reconstrução idêntica** do XLSX; **teste de mutação** (dados alterados ⇒ tudo recalcula e continua batendo); renderização automatizada; documentação.

**Controles positivos** (verificação manual, cópia descartada): cinco defeitos injetados um a um (cache adulterado, cartão fixo, CSV alterado, gráfico sem `autoTitleDeleted`, referência de gráfico quebrada) — a suíte reprovou o teste correspondente em todos.

## 7. Evidências de regressão

| KPI oficial | Referência do enunciado | XLSX V0.1.1 | Indicadores (V0.2) | Cartão do DASHBOARD | CSV (cálculo independente) |
|---|---|---|---|---|---|
| Unidades consumidas | 1.647 | 1.647 | 1.647 | 1.647 | 1.647 |
| Custo estimado consumido | 24.606,20 | 24.606,20 | 24.606,20 | 24.606,20 | 24.606,20 |
| Movimentações | 180 | 180 | 180 | 180 | 180 |
| Materiais distintos movimentados | 20 | 20 | 20 | 20 | 20 |
| Itens abaixo ou iguais ao mínimo | 5 | 5 | 5 | 5 | 5 |
| Itens críticos abaixo ou iguais ao mínimo | 2 | 2 | 2 | 2 | 2 |

- **Dados:** os 3 CSVs têm os mesmos SHA-256 da baseline (`materiais` `eca98597…231b`, `estoque` `9f90cd52…a8d0`, `consumo` `b5b82497…233c`).
- **Arquivos protegidos:** 23 arquivos da V0.1.1 (lista em `docs/V0_1_1_BASELINE_SHA256.txt`) seguem byte a byte iguais — inclui o XLSX histórico, `executar_testes.py` e `construir_excel.py`. O relatório da V0.1.1 foi preservado como `docs/TEST_REPORT_V0_1_1.md` (mesmo hash do `TEST_REPORT.md` original).
- **Abas de dados:** Materiais, Estoque, Consumo e Qualidade têm 0 diferenças (constantes e fórmulas), mesmas validações de dados e formatação condicional; valores calculados iguais em todas as células. Indicadores difere só em A1 e A3 (texto); Parametros só ganhou A20:C24.
- **Regressões encontradas:** nenhuma.

## 8. Arquivos

| Situação | Arquivos |
|---|---|
| **Novos** | `excel/lobo_insights_industrial_v0_2.xlsx`; `scripts/construir_excel_v0_2.py`; `scripts/executar_testes_v0_2.py`; `docs/V0_2_NOTES.md`; `docs/V0_1_1_BASELINE_SHA256.txt`; `docs/TEST_REPORT_V0_1_1.md` (cópia idêntica do relatório da V0.1.1); `docs/manual/inspecao_visual_v0_2.md`; `docs/manual/observacoes_execucao_v0_2.md`; `images/dashboard_v0_2_pagina_1.png`, `_2.png`, `_3.png` |
| **Modificados** | `README.md`; `CHANGELOG.md`; `docs/LIMITATIONS.md` (itens 17 em diante); `docs/TEST_REPORT.md` (regenerado para a V0.2); `SHA256SUMS.txt` |
| **Idênticos à V0.1.1** | 3 CSVs; `excel/lobo_insights_industrial_v0_1.xlsx`; `lobo_common.py`; `gerar_dados.py`; `validar_dados.py`; `construir_excel.py`; `executar_testes.py`; `requirements.txt`; `DATA_DICTIONARY.md`; `V0_1_NOTES.md`; `V0_1_1_NOTES.md`; `TEST_REPORT_V0_1.md`; `V0_1_BASELINE_SHA256.txt`; `manual/inspecao_visual.md`; `manual/observacoes_execucao.md`; 5 `.gitkeep` |

Nenhum arquivo foi removido.

## 9. O que o autor precisa saber explicar

- **Por que a V0.2 importa `construir()` da V0.1.1 em vez de editá-lo:** a baseline é congelada e o teste de reconstrução da V0.1.1 depende dela; assim a mudança é aditiva e auditável.
- **Por que `RANK` + `COUNTIF` em vez de `SORT`:** portabilidade e compatibilidade; e como o desempate garante posições únicas.
- **O que é um Pareto e o que ele NÃO diz:** descreve a concentração do período observado; o 80% é convenção parametrizável; não prevê nem recomenda.
- **Por que não há filtros:** critério técnico (seção 3, item 9), não omissão.
- **Como se prova que nada é fixo:** teste de mutação (dados alterados, painel recalcula e bate com um cálculo independente) e teste de ausência de constantes numéricas.
- **Por que um teste que sempre passa é suspeito:** controles positivos (injetar o defeito e ver a suíte reprovar).
- **Limites do que foi verificado:** LibreOffice ≠ Excel; ver seção 5.
