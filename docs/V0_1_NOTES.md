# Notas da V0.1 — Foundation

> Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.

Este documento explica **o que foi feito, por quê, como validar e o que você precisa saber explicar em uma entrevista**. Linguagem técnica, mas pensada para quem está em transição para TI e análise de dados.

## 1. Escopo desta versão

Entregue: pastas, dicionário de dados, catálogo (20 materiais), estoque (20 posições), consumo (180 movimentações em 3 meses), Excel estruturado, 36 verificações de qualidade, 6 KPIs, documentação, testes e pacote reproduzível.
**Não** entregue (de propósito): dashboard (V0.2), IA, classificador, Power BI, n8n, GitHub, qualquer previsão.

## 2. Resultados de referência

| Indicador | Valor |
|---|---|
| Unidades consumidas | 1.647 |
| Custo estimado consumido | 24.606,20 |
| Movimentações | 180 |
| Materiais distintos movimentados | 20 |
| Itens abaixo ou iguais ao mínimo | 5 |
| Itens críticos abaixo ou iguais ao mínimo | 2 |

Estoque atual: 5 itens REPOR, 4 ATENÇÃO, 11 OK. Estes números foram **calculados**, não digitados: saem das fórmulas do Excel e foram reproduzidos por um cálculo independente em Python (ver `TEST_REPORT.md`, seção 5).

## 3. Decisões técnicas

### D1 — Começar pelos dados, não pelo dashboard
- **O QUE FOI FEITO:** primeiro CSVs e verificações; só depois Excel e indicadores; nenhum gráfico.
- **POR QUE FOI FEITO:** um gráfico bonito sobre dado errado é pior que nenhum gráfico: ele passa confiança falsa. Corrigir erro na origem custa pouco; corrigir depois de publicado custa credibilidade.
- **COMO VALIDAR:** `python scripts/validar_dados.py` deve mostrar 36 PASS antes de qualquer KPI ser lido.
- **O QUE PRECISO SABER EXPLICAR:** "dados corretos primeiro" e qual foi a ordem: gerar, validar, montar o Excel, reconciliar, testar.

### D2 — Formato dos CSVs (`;`, UTF-8 com BOM, ponto decimal)
- **O QUE FOI FEITO:** separador `;` (pedido pelo PDF), UTF-8 com BOM, custo como `9.80`, datas `aaaa-mm-dd`, finais de linha LF.
- **POR QUE FOI FEITO:** `;` evita conflito com vírgulas dentro de nomes ("Disco Flap 4,5 pol"). O BOM faz o Excel do Windows exibir acentos corretamente. Ponto decimal segue o exemplo do PDF (`12.50`) e é o formato que pandas, SQL e Power BI leem sem ajuste. Data ISO não é ambígua.
- **COMO VALIDAR:** testes de "Formato CSV" no `TEST_REPORT.md` (cabeçalho, número de campos, BOM, LF, 2 casas decimais).
- **O QUE PRECISO SABER EXPLICAR:** que separador, codificação e decimal são decisões explícitas; em Excel brasileiro, abrir CSV com ponto decimal exige escolher a localidade na importação.

### D3 — Chaves e integridade referencial
- **O QUE FOI FEITO:** `Material_ID` é chave primária do catálogo; `estoque` e `consumo` a usam como chave estrangeira; `Movimento_ID` é a chave do consumo.
- **POR QUE FOI FEITO:** sem chave única não há como contar, juntar tabelas ou apontar "qual linha está errada".
- **COMO VALIDAR:** verificações Q01–Q03 (duplicidade), Q11–Q13 (órfãos) e Q14–Q23 (divergência de atributos repetidos).
- **O QUE PRECISO SABER EXPLICAR:** chave = coluna que identifica uma linha de forma única; integridade referencial = todo valor de chave estrangeira existe na tabela de origem (um consumo de `MAT999` sem `MAT999` no catálogo violaria a regra).

### D4 — Dados sintéticos gerados por script com semente fixa
- **O QUE FOI FEITO:** `gerar_dados.py` sorteia materiais, datas, centros e quantidades com pesos; a semente (20260601) foi fixada **antes** de olhar o resultado; a distribuição é desigual (máx/mín de movimentações por material >= 3) com dias úteis pesando mais, sábados pouco e nenhum domingo. Nenhuma conclusão foi forçada.
- **POR QUE FOI FEITO:** dados sintéticos eliminam risco de vazar informação real e permitem publicar o projeto. Semente fixa torna o resultado reproduzível: qualquer pessoa gera os mesmos bytes.
- **COMO VALIDAR:** teste "gerar_dados.py reproduz os 3 CSVs byte a byte" e a varredura de PII no `TEST_REPORT.md`.
- **O QUE PRECISO SABER EXPLICAR:** por que sintético (segurança, ética, publicação) e que "padrões" do dado vêm de sorteio; por isso a queda de agosto **não** é um achado de negócio (ver `LIMITATIONS.md`).

### D5 — Catálogo e estoque desenhados para testar regras
- **O QUE FOI FEITO:** 20 materiais em 6 categorias e 3 níveis de criticidade; 5 itens REPOR (um exatamente igual ao mínimo), 4 ATENÇÃO e 11 OK; 2 itens de criticidade Alta em REPOR.
- **POR QUE FOI FEITO:** o estoque é um **cenário de teste**: precisa ter casos em cada faixa para que o status e os KPIs 5 e 6 sejam realmente exercitados (o PDF pede ao menos 5 itens em ATENÇÃO/REPOR). O consumo, ao contrário, não foi desenhado para provar nada.
- **COMO VALIDAR:** testes "Estoque" e "Detecção de erros" (limites Atual = Mín, Mín x 1,25 e logo acima).
- **O QUE PRECISO SABER EXPLICAR:** a regra de status é simples e transparente; **não é previsão**: compara o hoje com o mínimo.

### D6 — Excel com tabelas estruturadas e fórmulas reais
- **O QUE FOI FEITO:** `tbMateriais`, `tbEstoque`, `tbConsumo`; colunas calculadas por fórmula; resumos com SOMASES/CONT.SES; `SE`/`SEERRO` no status e nas buscas; ÍNDICE+CORRESP como alternativa compatível ao PROCX; validação de dados por listas nomeadas; formatação condicional para REPOR/ATENÇÃO; parâmetros na aba `Parametros`.
- **POR QUE FOI FEITO:** tabela estruturada dá fórmulas legíveis (`tbConsumo[Quantidade]`) e cresce sozinha; fórmula real deixa o número **rastreável**; parâmetros evitam "número mágico" dentro de fórmula. Foi usado `ANO&"-"&TEXTO(MÊS,"00")` em vez de `TEXTO(data,"aaaa-mm")` porque este último depende do idioma do Excel.
- **COMO VALIDAR:** testes "Excel – estrutura", "Excel – fórmulas" (0 erros, 0 números digitados na aba Indicadores) e "Excel – resumos".
- **O QUE PRECISO SABER EXPLICAR:** diferença entre valor colado e fórmula; o que é tabela estruturada; por que ÍNDICE+CORRESP funciona onde não há PROCX. Tabelas dinâmicas **não** foram criadas nesta versão.

### D7 — Qualidade em dois níveis independentes
- **O QUE FOI FEITO:** as mesmas 36 verificações (Q01–Q36) existem em Python (`validar_dados.py`, lê o CSV como texto) e em fórmulas na aba `Qualidade`.
- **POR QUE FOI FEITO:** se duas implementações independentes concordam, a chance de um erro passar despercebido cai muito; e o Excel passa a ser autoverificável quando alguém editar a base.
- **COMO VALIDAR:** teste "vetor de erros do Excel = vetor do Python".
- **O QUE PRECISO SABER EXPLICAR:** o que cada verificação protege (duplicidade, nulos, valores impossíveis, chave órfã, divergência entre tabelas, domínio, datas).

### D8 — Testar os testes (injeção de erros)
- **O QUE FOI FEITO:** um dataset em memória com erros plantados (ID repetido, MAT999, quantidade 0/-3/2,5, custo 0, data 2026-13-45 etc.) passa pelo Python e por um XLSX temporário; o número de erros detectado é comparado ao esperado, teste a teste.
- **POR QUE FOI FEITO:** "0 erros" só significa algo se a verificação for capaz de achar erro. Um teste que nunca falha não prova nada.
- **COMO VALIDAR:** seção 6 do `TEST_REPORT.md`.
- **O QUE PRECISO SABER EXPLICAR:** essa técnica se chama teste de mutação/injeção de falhas; sem ela, um PASS pode ser só uma verificação quebrada.

### D9 — Valores calculados gravados dentro do XLSX
- **O QUE FOI FEITO:** o openpyxl escreve as fórmulas; o LibreOffice recalcula uma cópia; os resultados são gravados como cache no arquivo final (as fórmulas continuam sendo fórmulas; o Excel ainda recalcula ao abrir).
- **POR QUE FOI FEITO:** sem cache, visualizadores e celulares mostram células vazias.
- **COMO VALIDAR:** testes "células de fórmula sem valor em cache = 0" e "células com erro = 0".
- **O QUE PRECISO SABER EXPLICAR:** que o XLSX não é byte a byte reproduzível (metadados variam), mas os CSVs são; o que garante o XLSX é o SHA-256 registrado e os testes.

### D10 — Pasta `scripts/` adicionada
- **O QUE FOI FEITO:** cinco scripts em `scripts/`.
- **POR QUE FOI FEITO:** o critério "pacote reproduzível" exige o código que gerou e testou os dados; a estrutura do enunciado não previa onde colocá-lo. Documentado como desvio consciente, não silencioso.
- **COMO VALIDAR:** teste de reprodutibilidade dos CSVs; `executar_testes.py --verificar` roda dentro do ZIP.
- **O QUE PRECISO SABER EXPLICAR:** por que o código faz parte da evidência.

## 4. Inconsistências e ambiguidades encontradas (e a interpretação adotada)

Nada foi alterado em silêncio. Em cada ponto, usei a leitura mais conservadora.

| # | Ponto | PDF | Enunciado | Interpretação adotada |
|---|---|---|---|---|
| 1 | Base de solicitações | 4 arquivos, incluindo `solicitacoes_ficticias.csv` | V0.1 pede 3 CSVs | Solicitações ficam para a V0.4 |
| 2 | `Fornecedor` no catálogo | Está em `materiais_ficticios.csv` | Campos mínimos do catálogo não o citam; o estoque o exige | Incluído nas duas tabelas e conferido (Q23, Q36) |
| 3 | Campos do estoque e do consumo | Estoque com 5 campos; consumo sem nome, categoria, unidade e custo | Lista campos adicionais | Usados os campos do enunciado (superconjunto do PDF) |
| 4 | Status ATENÇÃO | "Se próximo do mínimo", sem critério | "Se criar, documentar o critério" | Criado: Atual > Mín e até Mín x 1,25 (parâmetro editável) |
| 5 | "Crítico em risco" | Alta e status diferente de OK (inclui ATENÇÃO) | KPI: críticos abaixo ou iguais ao mínimo | KPI usa só REPOR; críticos em ATENÇÃO viram indicador informativo |
| 6 | Tipo de movimentação | "Consumo ou ajuste fictício" | "Tipo principal: Consumo" | Só Consumo nos dados; "Ajuste" permitido pela regra; KPIs filtram Consumo |
| 7 | Decimal do CSV | Exemplo `12.50` e CSV com `;` | Pede conferir separadores | Ponto decimal, documentado |
| 8 | Nome do arquivo Excel | `lobo_insights_industrial.xlsx` | `lobo_insights_industrial_v0_1.xlsx` | Seguido o enunciado |
| 9 | Tabelas dinâmicas | Listadas nas habilidades | "Quando fizer sentido" | Não criadas; resumos por fórmula; a criar no Excel desktop se desejado na V0.2 |
| 10 | Notas finais | `PORTFOLIO_NOTES` no repositório final | `V0_1_NOTES.md` agora | Seguido o enunciado |

## 5. O que o autor precisa saber explicar

1. **Por que começar pelos dados antes do dashboard.** Porque o dashboard só repete o que a base diz. Se a base tem ID duplicado ou custo negativo, todo número bonito fica errado, e ninguém percebe. Validar primeiro é barato; retratar depois é caro.
2. **O que é chave.** Uma coluna (ou conjunto) que identifica cada linha de forma única. Aqui: `Material_ID` no catálogo e `Movimento_ID` no consumo.
3. **O que é integridade referencial.** Regra de que todo valor de chave estrangeira exista na tabela de origem. Se `consumo_ficticio.csv` tivesse `MAT999` e o catálogo não, seria uma linha órfã: não teria custo, categoria ou criticidade confiáveis para análise.
4. **Diferença entre estoque atual e consumo.** Estoque é uma **foto** (quanto existe agora); consumo é um **fluxo** (quanto saiu ao longo do tempo). São tabelas diferentes e não se somam nem se subtraem diretamente: o estoque de hoje não é "o que sobrou do consumo dos três meses".
5. **Como calcular o custo estimado consumido.** Por movimentação: `Quantidade x Custo_Unitario`. Exemplo real: `MOV0027` consumiu 2 CX de "Eletrodo Revestido 2,5 mm" a 96,00 = **192,00**. Depois soma-se tudo do tipo Consumo (SOMASES): 24.606,20. É "estimado" porque o custo é fictício e fixo por material.
6. **Diferença entre movimentações e quantidade consumida.** Movimentação = número de **linhas** (180). Quantidade consumida = soma da coluna Quantidade (1.647 unidades). Uma movimentação pode ter 1 ou 60 unidades, e as unidades misturam itens (caixa, par, unidade), então é preciso cuidado ao comparar. Como ilustração (não é conclusão de negócio): "Corte" tem 940 unidades e 43 movimentações, enquanto "Soldagem" tem 57 unidades e 28 movimentações, mas custo estimado de 4.869,00 contra 7.565,30; volume e custo contam histórias diferentes.
7. **Por que todo KPI precisa ser reconciliado.** Reconciliar = provar que o número do relatório bate com a origem por outro caminho. Aqui, cada KPI do Excel foi recalculado em Python direto dos CSVs, e cada resumo (mês, categoria, centro, material) precisa fechar com o KPI total. Se não bater, investiga-se a causa; não se ajusta o número.
8. **Por que usar dados sintéticos.** Porque é seguro (nada real vaza), publicável (GitHub) e controlável (dá para plantar erros para testar as verificações). O custo é que não se pode inferir nada sobre uma operação real.
9. **Quais são as limitações.** Três meses de dados, custos fictícios, estoque como foto única, sem compras nem contratos, sem previsão, sem uso operacional; diferenças entre meses são fruto de sorteio. Ver `LIMITATIONS.md`.
10. **O que será feito na V0.2.** Transformar os números da V0.1 em um painel Excel legível em menos de dois minutos: KPIs no topo, evolução mensal, Top 10 materiais, consumo por categoria e por centro, tabela destacada de REPOR/ATENÇÃO, filtros e (se fizer sentido) tabelas dinâmicas. Continua sem previsão.

## 6. Como reproduzir

```bash
python scripts/gerar_dados.py        # recria os 3 CSVs (mesma semente = mesmos bytes)
python scripts/validar_dados.py      # 36 verificações
python scripts/construir_excel.py    # recria o XLSX (requer LibreOffice)
python scripts/executar_testes.py    # roda os testes e reescreve docs/TEST_REPORT.md
```
