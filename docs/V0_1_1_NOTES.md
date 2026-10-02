# Notas da V0.1.1 — Portable Baseline

> Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.

## 1. Objetivo

Tornar a V0.1 **portátil e reproduzível** em outro ambiente, sem mudar nada do que ela faz: mesmos dados, mesmos KPIs, mesmas fórmulas, mesma planilha. É uma revisão técnica de estabilização, **não** é a V0.2.

## 2. Escopo

**Dentro:** eliminar a dependência de um script fora do projeto (`recalc.py`, em pasta privada do ambiente original); corrigir outros pontos de portabilidade encontrados na auditoria da infraestrutura; testes de regressão contra a V0.1; documentação da versão.
**Fora (não feito, de propósito):** novas funcionalidades, novos KPIs, dashboard, Power BI, IA, previsão, machine learning, qualquer conteúdo do Projeto 2.
**A V0.1 permanece congelada:** o relatório dela está preservado sem alteração em `TEST_REPORT_V0_1.md`, e seus hashes em `V0_1_BASELINE_SHA256.txt`.

## 3. Base utilizada e estado antes da alteração

- Base: `Lobo_Insights_Industrial_V0_1.zip`, SHA-256 `ecf568bb4b4cfe250cbd37ccbe64b5591e4233850d01d5414a07e7df9591d690`. O ZIP citado como anexo não estava na pasta de uploads da sessão; foi usado o ZIP entregue na etapa anterior. O hash coincide com o registrado e os 21 arquivos conferem com o `SHA256SUMS.txt` da V0.1.
- Suíte da V0.1 reexecutada **antes de qualquer edição**: 90 PASS / 0 FAIL, 36/36 verificações de qualidade, 36/36 testes de detecção.
- **Defeito reproduzido:** com o `recalc.py` externo indisponível, a suíte da V0.1 **abortava inteira** (`FileNotFoundError`), e não apenas os testes que dependem dele. A dependência aparece em `construir_excel.py` (constante `RECALC`, com caminho absoluto fora do projeto, e o `subprocess` que a chamava). Esse recálculo é usado pela construção do XLSX e pelos testes T71 a T74 (detecção de erros no Excel).
- Isto **não** é erro de dados nem falha funcional da V0.1.

## 4. Alteração realizada

### 4.1 Correção principal — recálculo portátil (`scripts/construir_excel.py`)
- **O QUE FOI FEITO:** `recalcular()` chama o LibreOffice diretamente (`soffice --headless --convert-to xlsx`) com um perfil temporário isolado que força "recalcular sempre ao abrir" (`OOXMLRecalcMode = 0`). O executável é localizado por `localizar_soffice()`: variável `LOBO_SOFFICE`, senão `PATH`. Sem LibreOffice, lança `LibreOfficeIndisponivel` com instrução clara. A assinatura e o formato de retorno de `recalcular()` foram mantidos.
- **POR QUE FOI FEITO:** o `recalc.py` original dependia de um pacote auxiliar (`office`) e de um ajuste de sandbox existentes só no ambiente de criação. Copiar esses arquivos para o projeto traria a complexidade do sandbox junto; trocar por outro caminho absoluto só moveria o problema. Usar o `soffice` pelo `PATH` é o padrão portátil e é bem menos código.
- **COMO VALIDAR:** T94 reconstrói o XLSX a partir dos CSVs e compara com o entregue; T91 e T92 verificam a ausência de caminhos do ambiente original e o comportamento sem LibreOffice.
- **O QUE PRECISO SABER EXPLICAR:** o que é dependência implícita de ambiente; por que "fallback controlado" (erro claro ou SKIP) é melhor que quebrar ou fingir que passou.

### 4.2 Testes sem LibreOffice: SKIP, nunca PASS (`scripts/executar_testes.py`)
Se o LibreOffice não existir, os testes que dependem dele viram **SKIP** com o motivo, a suíte continua, e o código de saída é 2 (0 = tudo executado e aprovado, 1 = há FAIL).

### 4.3 Achados adicionais de portabilidade (auditoria estática)
| Achado | Correção |
|---|---|
| O gerador do relatório lia as seções 7 e 8 de uma pasta **fora do pacote**; em outro ambiente o relatório sairia sem elas | Notas movidas para `docs/manual/` (dentro do pacote) |
| f-string com barra invertida no campo (sintaxe válida só no Python 3.12+) | Constantes `BOM` e `CRLF` no topo do módulo |
| Import não usado (`copy`) | Removido; também `json` e `sys` deixaram de ser usados em `construir_excel.py` e foram removidos |
| Dois `open()` sem `with` | Substituídos por `ler_csv()`, função que já existia |
| Bibliotecas não listadas em arquivo | `requirements.txt` (versões testadas) |

### 4.4 Testes novos (T91 a T103) — T01 a T90 são os mesmos da V0.1
Caminhos absolutos (T91); LibreOffice ausente (T92); LibreOffice encontrado (T93); reconstrução idêntica do XLSX (T94); execução a partir de outro diretório (T95); regressão byte a byte contra a base V0.1 (T96, T97); arquivos novos e changelog da V0.1.1 (T98, T99); `requirements.txt` (T100); notas manuais dentro do pacote (T101); imports não usados (T102); sintaxe compatível com Python anterior a 3.12 (T103).

## 5. Arquivos

| Situação | Arquivos |
|---|---|
| **Modificados (7)** | `scripts/construir_excel.py`, `scripts/executar_testes.py`, `CHANGELOG.md`, `README.md`, `docs/LIMITATIONS.md` (nova limitação de ambiente, item 16), `docs/TEST_REPORT.md` (regenerado nesta versão), `SHA256SUMS.txt` (regenerado) |
| **Novos (6)** | `requirements.txt`, `docs/V0_1_1_NOTES.md`, `docs/V0_1_BASELINE_SHA256.txt`, `docs/TEST_REPORT_V0_1.md` (cópia idêntica do relatório da V0.1), `docs/manual/inspecao_visual.md`, `docs/manual/observacoes_execucao.md` |
| **Idênticos à V0.1 (14)** | 3 CSVs, o XLSX, `DATA_DICTIONARY.md`, `V0_1_NOTES.md`, `lobo_common.py`, `gerar_dados.py`, `validar_dados.py` e os 5 `.gitkeep` |

Nenhum arquivo foi removido. Total no pacote: 27 arquivos (21 da V0.1 + 6 novos).

## 6. Evidências e resultados

### 6.1 Suíte completa (neste ambiente: Linux x86_64, Python 3.12.3, LibreOffice 24.2)
- **103 testes: 103 PASS, 0 FAIL, 0 SKIP.** As 90 primeiras linhas (T01 a T90) são **idênticas, linha a linha**, às do relatório da V0.1 (comparação feita entre os dois relatórios); T91 a T103 são novos.
- **36/36** verificações de qualidade dos CSVs (Python); **36/36** testes de detecção por injeção de erros (Python e Excel).
- Detalhes em `TEST_REPORT.md`.

### 6.2 Cenário sem LibreOffice (`LOBO_SOFFICE` apontando para caminho inexistente)
Suíte concluída sem travar: **97 PASS, 0 FAIL, 6 SKIP** (T71 a T74, T93 e T94) e código de saída 2; a tabela de detecção mostra o lado Excel como "NÃO EXECUTADO". `construir_excel.py` encerra com mensagem clara e **não altera** o XLSX existente.

### 6.3 Reconstrução do Excel
Reconstruído a partir dos CSVs com o recálculo novo e comparado com o XLSX entregue: **3.278 células, 857 fórmulas, 0 diferenças** em fórmulas/constantes e em valores calculados, 0 erros de fórmula (T94). Verificação manual adicional: com um valor em cache propositalmente errado (999), `recalcular()` devolveu 1.647, ou seja, o recálculo é forçado de fato.

### 6.4 Controles positivos (as novas verificações detectam o problema?)
- Aplicadas ao código da V0.1: o detector de caminhos acusou `/mnt/`, `recalc.py` e `LOBO_RECALC`; o de imports, `copy`; o de f-string, 2 ocorrências na linha 265.
- Alterando **um único caractere** em um CSV, numa cópia do pacote, a suíte reprovou 4 testes (T40, T83, T94 e T96). Verificação manual; a cópia foi descartada.

### 6.5 Regressão: V0.1 × V0.1.1
| KPI | V0.1 (XLSX) | V0.1.1 (XLSX) | CSVs, cálculo independente (V0.1.1) |
|---|---|---|---|
| Unidades consumidas | 1.647 | 1.647 | 1.647 |
| Custo estimado consumido | 24.606,20 | 24.606,20 | 24.606,20 |
| Movimentações | 180 | 180 | 180 |
| Materiais distintos movimentados | 20 | 20 | 20 |
| Itens abaixo ou iguais ao mínimo | 5 | 5 | 5 |
| Itens críticos abaixo ou iguais ao mínimo | 2 | 2 | 2 |

- **Dados:** os 3 CSVs têm o mesmo SHA-256 da V0.1 (`materiais` `eca98597…231b`, `estoque` `9f90cd52…a8d0`, `consumo` `b5b82497…233c`).
- **Excel:** o XLSX é o mesmo arquivo da V0.1, byte a byte (SHA-256 `cfdd4c65e4e53537fc4ff948d2651cb9ac1cfaeba3798e1919291d15e2068964`); por isso nome, abas, tabelas, fórmulas e estrutura são idênticos. O arquivo não foi regenerado: a reconstrução foi feita só em pasta temporária para comparação.
- **Regras de negócio e nomes de campos:** `lobo_common.py`, `gerar_dados.py` e `validar_dados.py` idênticos aos da V0.1.
- **Regressões encontradas:** nenhuma.

## 7. Limitações remanescentes

1. **Ambiente:** só foram verificados Linux x86_64 e Python 3.12.3. Windows, macOS e outras versões de Python e de bibliotecas não foram testados. A correção de sintaxe visa Python 3.11 (mínimo do pandas 3), mas isso não foi verificado.
2. **Instalação limpa:** `pip install -r requirements.txt` em ambiente virtual novo não foi testado (sem acesso à rede neste ambiente).
3. **LibreOffice continua necessário** para reconstruir o XLSX e para 6 testes; sem ele os testes ficam SKIP. `LOBO_SOFFICE` com caminho válido em local não padrão não foi testado.
4. **XLSX não é byte a byte reproduzível** quando reconstruído (metadados variam), como já registrado na V0.1; os CSVs são.
5. Permanecem as limitações da V0.1: não testado no Microsoft Excel real, sem tabelas dinâmicas, inspeção visual parcial (`LIMITATIONS.md` e `TEST_REPORT_V0_1.md`).

## 8. O que o autor precisa saber explicar

- **Por que uma dependência de caminho externo quebra a reprodutibilidade:** quem clona o projeto não tem aquela pasta; o projeto precisa carregar consigo ou declarar tudo de que depende.
- **Por que não trocar por outro caminho absoluto:** só mudaria o endereço do problema. Resolver por `PATH`/variável de ambiente ou por caminho relativo ao projeto funciona em qualquer máquina.
- **Diferença entre PASS, FAIL e SKIP:** SKIP é "não executado", com motivo. Contá-lo como PASS seria mentir sobre a cobertura.
- **Como se prova "não mudou nada":** comparando hashes dos arquivos protegidos, KPIs lado a lado e as linhas de teste T01 a T90 entre versões; e testando o próprio teste (controle positivo).
- **Por que a V0.1 continua congelada:** rastreabilidade. Cada versão precisa poder ser auditada como foi entregue.
