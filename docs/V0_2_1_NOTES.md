# Notas da V0.2.1 — Packaging & Reproducibility Fix

> Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.

**Natureza da versão.** Microversão técnica de *release engineering*. Corrige **somente** (1) a nomenclatura dos entregáveis principais e (2) a estabilidade do manifesto de hashes. **Nenhuma** funcionalidade, análise, dado, fórmula, KPI, gráfico, layout ou resultado foi alterado. Não há V0.3, previsão, IA, Power BI ou automação aqui.

- ZIP da release: `Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.zip`
- XLSX atual: `Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx` (dentro do ZIP, em `excel/`)

## 1. Resultado da auditoria da V0.2

Resultado informado pela auditoria independente (e reconfirmado nesta versão pelas execuções descritas na seção 6): a camada funcional e analítica da V0.2 foi **aprovada** — 174 PASS, 0 FAIL, 0 SKIP (103 testes herdados da V0.1.1 + 71 da V0.2); 36/36 verificações de qualidade dos CSVs; 3 CSVs oficiais preservados; 2.417 fórmulas no XLSX, 0 erros de fórmula, 0 vínculos externos; 9 gráficos no DASHBOARD; 6 KPIs oficiais preservados; suíte executada a partir de outro diretório com o mesmo resultado.

A auditoria apontou dois achados de empacotamento, ambos tratados aqui:

1. **Nomenclatura:** o ZIP e o XLSX da V0.2 não tinham o prefixo obrigatório `Projeto_Lobo_1_`.
2. **Manifesto de hashes:** depois de uma execução legítima de `python scripts/executar_testes_v0_2.py`, quatro arquivos regenerados mudaram de hash e `sha256sum -c SHA256SUMS.txt` deixou de passar.

## 2. Ausência de alteração funcional

Provas (todas executadas; ver seção 6):

- `excel/Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx` é **byte a byte igual** ao XLSX auditado da V0.2 (SHA-256 `81f51886dc3d68408073a9e0bf54859ae100fe418c2c30060b8d92fb9662b846`, mesmo do `excel/lobo_insights_industrial_v0_2.xlsx`, que permanece com o nome histórico). Logo dados, fórmulas, KPIs, rankings, Pareto, parâmetros, dashboard, gráficos, layout e resultados são os mesmos.
- 30 arquivos da V0.2 auditada (lista em `docs/V0_2_AUDITED_SHA256.txt`) seguem byte a byte iguais: os 3 CSVs, os 2 XLSX históricos, **todos** os scripts e testes da V0.1.1 e da V0.2 (`construir_excel_v0_2.py` e `executar_testes_v0_2.py` não foram tocados), `V0_2_NOTES.md` e os documentos históricos.
- A bateria de testes da V0.2 (71 testes, incluindo cálculo independente de todas as tabelas da aba Analises e o teste de mutação) foi executada **sem alteração e agora sobre o XLSX com o nome novo**: tudo PASS.
- O único código novo é de empacotamento: `scripts/lobo_release.py`, `scripts/construir_excel_v0_2_1.py` e `scripts/executar_testes_v0_2_1.py`. Nenhum contém lógica analítica.

**Marca de versão interna do XLSX.** O texto interno do workbook (aba LEIA_ME, rodapé de impressão) continua dizendo "V0.2". É deliberado: alterar qualquer string interna mudaria o arquivo auditado. V0.2 é a versão *analítica* do conteúdo; V0.2.1 é a versão de *empacotamento*.

## 3. Correção de nomenclatura

| Entregável | Antes (V0.2) | Agora (V0.2.1) |
|---|---|---|
| ZIP | `Lobo_Insights_Industrial_V0_2.zip` | `Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.zip` |
| XLSX atual | `lobo_insights_industrial_v0_2.xlsx` | `Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx` |

- **Preservado (históricos, não renomeados):** `excel/lobo_insights_industrial_v0_1.xlsx` e `excel/lobo_insights_industrial_v0_2.xlsx`, todos os scripts e documentos das versões anteriores. O arquivo histórico da V0.2 é mantido para que `docs/V0_2_NOTES.md` e a bateria de testes da V0.2 continuem verdadeiros; ele é idêntico ao XLSX atual (um teste garante).
- **Genéricos mantidos:** `README.md`, `CHANGELOG.md`, `SHA256SUMS.txt`, `docs/TEST_REPORT.md`.
- **Pasta-raiz dentro do ZIP:** continua `lobo-insights-industrial/` (identidade do projeto, igual nas versões anteriores; a versão vai no nome do ZIP).
- **Constantes:** o nome obrigatório vive em `scripts/lobo_release.py` (`NOME_ZIP`, `NOME_XLSX`, `PREFIXO`); `scripts/construir_excel_v0_2_1.py` grava o XLSX nesse nome; `python scripts/lobo_release.py empacotar` gera o ZIP com o nome exato.
- **Testes que impedem regressão:** o XLSX e o ZIP têm exatamente os nomes obrigatórios (literais independentes no teste); começam com `Projeto_Lobo_1_`; a regra de validação **rejeita** nomes ambíguos (`Lobo_…`, `lobo_…`, prefixos do Projeto 2, extensão errada — 7 exemplos negativos); `excel/` só contém o XLSX atual e os dois nomes históricos; nenhum outro `.xlsx` no projeto; README, CHANGELOG e estas notas citam os nomes exatos.

## 4. Estratégia de hashes

**Causa.** Uma execução normal da suíte reescreve `docs/TEST_REPORT.md` (registra versões de Python, bibliotecas, LibreOffice e plataforma) e `images/dashboard_v0_2_pagina_1.png` a `_3.png` (o desenho depende da versão do LibreOffice, das fontes e do ambiente gráfico). Esses quatro arquivos estavam no `SHA256SUMS.txt`, então variavam de forma legítima e invalidavam o manifesto. (Neste ambiente as regenerações saem byte a byte idênticas, por isso a falha não reproduziu aqui; a causa foi confirmada pela comparação de árvores antes/depois de uma execução normal — só esses 4 arquivos são reescritos — e pelo conteúdo do relatório. Detalhes em `docs/manual/observacoes_execucao_v0_2_1.md`.)

**Solução: dois manifestos com papéis distintos.**

| | `SHA256SUMS.txt` | `RELEASE_SHA256SUMS.txt` |
|---|---|---|
| Papel | Manifesto **operacional** | **Snapshot** da release |
| Conteúdo | Todos os arquivos **estáveis** (dados, scripts, XLSX, documentação-fonte, linhas de base) | **Todos** os arquivos, inclusive `SHA256SUMS.txt` e os 4 regeneráveis, no instante do empacotamento |
| Fora dele | Os 4 regeneráveis e os dois manifestos | Somente ele mesmo |
| Quando confere | **Sempre**, inclusive depois de executar os testes | Só no pacote **recém-extraído** (depois de uma execução normal, os 4 regeneráveis divergem — por desenho) |
| Serve para | Detectar alteração real em dados, scripts, XLSX e fontes | Provar o conteúdo exato entregue |

Formato idêntico ao do `sha256sum` (`<hash>  <caminho>`, sem comentários), ordem bytewise, sem linhas duplicadas.

**Por que isso não mascara alterações indevidas.**

- Regeneráveis são uma lista **fechada e fixa** (4 caminhos) em `scripts/lobo_release.py`; um teste exige que a lista seja exatamente essa. Acrescentar um arquivo a ela exige mudar o código, o teste e esta nota.
- Um teste de cobertura exige que **todo** arquivo do projeto que não seja regenerável nem manifesto esteja em `SHA256SUMS.txt` — arquivo novo não registrado reprova.
- Controles em cópia descartável (executados a cada rodada): alterar os 4 regeneráveis, ou apagar um, **não** invalida `SHA256SUMS.txt` (e o snapshot acusa exatamente os 4); alterar 1 byte de um CSV, de um script, do XLSX ou de uma nota **reprova** só aquele arquivo; apagar um arquivo estável é acusado; criar arquivo não registrado é acusado.
- `SHA256SUMS.txt` é reproduzível: um teste recalcula o manifesto a partir da árvore e exige igualdade com o arquivo gravado.

**Comandos** (a partir da raiz do projeto):

```
sha256sum -c SHA256SUMS.txt                        # operacional: passa antes e depois de executar os testes
python scripts/lobo_release.py verificar           # o mesmo, em Python (portável)
sha256sum -c RELEASE_SHA256SUMS.txt                # snapshot: só no pacote recém-extraído
python scripts/lobo_release.py manifestos          # regrava SHA256SUMS.txt e, em seguida, RELEASE_SHA256SUMS.txt
python scripts/lobo_release.py empacotar           # cria o ZIP com o nome exato (ordem e datas fixas)
```

**Procedimento de release usado:** (1) executar os testes em modo normal (regenera os 4 arquivos); (2) `manifestos`; (3) `empacotar`; (4) extrair numa pasta limpa: `sha256sum -c SHA256SUMS.txt` e `sha256sum -c RELEASE_SHA256SUMS.txt`; (5) executar os testes em modo normal na cópia extraída e repetir `sha256sum -c SHA256SUMS.txt`; (6) executar `--verificar` a partir de outro diretório.

## 5. Arquivos regeneráveis

| Arquivo | Por que varia | Quem o reescreve |
|---|---|---|
| `docs/TEST_REPORT.md` | Registra versões de Python/pandas/openpyxl/LibreOffice e a plataforma | Suíte de testes em modo normal (sem `--verificar`) |
| `images/dashboard_v0_2_pagina_1.png` | Renderização do DASHBOARD pelo LibreOffice (versão, fontes, ambiente gráfico) | idem |
| `images/dashboard_v0_2_pagina_2.png` | idem | idem |
| `images/dashboard_v0_2_pagina_3.png` | idem | idem |

Observação: o executor histórico `executar_testes_v0_2.py` (não modificado) também reescreve esses mesmos 4 arquivos em modo normal — com o texto do relatório da V0.2. Isso não afeta o manifesto operacional; para obter o relatório da V0.2.1 use `executar_testes_v0_2_1.py`.

## 6. Testes executados

Executor: `python scripts/executar_testes_v0_2_1.py --verificar`. Ele executa, **sem alteração**, a suíte da V0.1.1 (T01–T103) e a bateria da V0.2 (T104–T174, agora sobre o XLSX com o nome novo) e acrescenta 27 testes (T175–T201): nomenclatura (9), regressão contra a V0.2 auditada (4), manifesto SHA-256 e estratégia de hashes (10) e documentação (4).

Resultado da última execução completa: **201 testes — 201 PASS, 0 FAIL, 0 SKIP**; 36/36 verificações de qualidade dos CSVs; 36/36 testes de detecção por injeção de erros (V0.1.1). Relatório: `docs/TEST_REPORT.md` (regenerável). Execuções adicionais registradas em `docs/manual/observacoes_execucao_v0_2_1.md`: suíte a partir de outro diretório, execução normal seguida de `sha256sum -c SHA256SUMS.txt`, e conferência do snapshot no pacote recém-extraído.

## 7. Confirmação de regressão

| Verificação exigida | Resultado |
|---|---|
| Suíte legada (V0.1.1) | 103 PASS |
| Bateria V0.2 sobre o XLSX atual | 71 PASS |
| Testes novos da V0.2.1 | 27 PASS |
| 6 KPIs | conferem (seção 8) |
| 36 verificações de qualidade dos CSVs | 36/36 PASS (Python) |
| Integridade dos 3 CSVs | SHA-256 idênticos aos da baseline (`materiais` `eca98597…231b`, `estoque` `9f90cd52…a8d0`, `consumo` `b5b82497…233c`) |
| Nenhum dado mudou | 30 arquivos da V0.2 auditada e os CSVs byte a byte iguais |
| Nenhum resultado analítico mudou | XLSX atual byte a byte igual ao auditado; bateria V0.2 (cálculo independente + mutação) PASS |
| 9 gráficos presentes | 9 (lidos do XML), todos com referências válidas |
| Erros de fórmula | 0, em 2.417 fórmulas |
| Vínculos externos | 0 (nem relações externas nem caminhos absolutos) |
| Execução a partir de outro diretório | PASS (ver observações de execução) |
| `sha256sum -c SHA256SUMS.txt` após executar os testes | PASS (ver observações de execução) |
| **Regressões** | **nenhuma** |

## 8. Confirmação dos KPIs

| KPI oficial | Valor |
|---|---|
| Unidades consumidas | 1.647 |
| Custo estimado consumido | 24.606,20 |
| Movimentações | 180 |
| Materiais distintos movimentados | 20 |
| Itens abaixo ou iguais ao mínimo | 5 |
| Itens críticos abaixo ou iguais ao mínimo | 2 |

Cada valor confere entre a referência do enunciado, o XLSX V0.1.1, a aba Indicadores do XLSX atual, o cartão do DASHBOARD e o cálculo independente dos CSVs (teste "KPIs oficiais (regressão)").

## 9. Limitações desta versão

- Continuam valendo todas as limitações da V0.1/V0.1.1/V0.2 (`docs/LIMITATIONS.md`): dados sintéticos, sem previsão, sem recomendação operacional, não testado no Microsoft Excel real, inspeção visual parcial.
- A falha de manifesto relatada pela auditoria **não foi reproduzida** no ambiente desta versão (as regenerações saem idênticas aqui). A estratégia foi validada por simulação e por execução completa; a variação real do ambiente do auditor não pôde ser observada.
- O ZIP não contém a prova do próprio hash: o SHA-256 do ZIP é informado fora dele (no relatório final).
- O snapshot `RELEASE_SHA256SUMS.txt` não é reproduzível depois de executar os testes (por desenho): use `SHA256SUMS.txt` para verificação contínua.
- O XLSX não é byte a byte reproduzível quando reconstruído (metadados variam); o entregue é o mesmo arquivo auditado da V0.2, e o gerador da V0.2.1 é testado por igualdade de conteúdo (fórmulas, valores, gráficos).

## 10. Arquivos

| Situação | Arquivos |
|---|---|
| **Novos** | `excel/Projeto_Lobo_1_Lobo_Insights_Industrial_V0_2_1.xlsx`; `scripts/lobo_release.py`; `scripts/construir_excel_v0_2_1.py`; `scripts/executar_testes_v0_2_1.py`; `docs/V0_2_1_NOTES.md`; `docs/V0_2_AUDITED_SHA256.txt`; `docs/manual/observacoes_execucao_v0_2_1.md`; `RELEASE_SHA256SUMS.txt` |
| **Modificados** | `README.md`; `CHANGELOG.md`; `docs/LIMITATIONS.md` (item 26); `SHA256SUMS.txt` (agora só os arquivos estáveis) |
| **Regenerados** (não entram no manifesto operacional) | `docs/TEST_REPORT.md` (agora do V0.2.1); `images/dashboard_v0_2_pagina_1.png`, `_2.png`, `_3.png` |
| **Idênticos à V0.2 auditada** | os 30 arquivos de `docs/V0_2_AUDITED_SHA256.txt` |

Nenhum arquivo foi removido.
