**Método.** O XLSX V0.2 foi convertido em PDF pelo LibreOffice (modo headless), as páginas foram rasterizadas com `pdftoppm` e **as imagens foram abertas e examinadas visualmente pelo executor (Claude)**. Isso mostra como o LibreOffice desenha o arquivo; **não** é uma inspeção no Microsoft Excel.

**Páginas inspecionadas (com resultado):**

| Aba / página | Resultado da inspeção |
|---|---|
| DASHBOARD, página 1 (cabeçalho, KPIs, evolução mensal, categorias) | Faixa superior em azul-marinho com filete cobre; 6 cartões legíveis, números em destaque, leitura rápida sob cada valor; gráficos de colunas (mensal) e barras horizontais (categorias) com rótulos de dados; sem sobreposição de texto. Ajustado durante a inspeção: altura do rótulo dos cartões (títulos de 2 linhas estavam cortados) e margem à direita das barras (o rótulo do maior valor encostava na borda). |
| DASHBOARD, página 2 (centros, Pareto de custo e de consumo, cartões de concentração) | Barras dos centros legíveis; os dois Paretos mostram colunas do "núcleo" em cor forte, "demais" em tom claro, linha de % acumulado com marcadores e linha tracejada do limite; eixo secundário 0–100%. Nomes dos 20 materiais rotacionados em 90°, legíveis. Ajustado: altura do gráfico (a área de plotagem estava baixa) e quebra de página (a página 2 ficou só com "Centros" na primeira tentativa). |
| DASHBOARD, página 3 (estoque e criticidade) | Fila de atenção com REPOR em vermelho-claro e ATENÇÃO em âmbar; itens críticos em risco (Alta em REPOR) com texto vermelho; matriz criticidade × status com cores; gráfico Atual × Mínimo legível; rodapé com rastreabilidade, conferências e aviso de dados sintéticos. Ajustado: título da matriz estava truncado. |
| LEIA_ME | Texto legível; camadas, decisões (sem filtros) e SHA-256 dos CSVs presentes. |
| Analises, página 1 | Tabelas com cabeçalho azul-marinho, legíveis. Na impressão a página fica pequena (a aba é de cálculo e foi pensada para a tela). |
| Indicadores, página 1 | Idêntica à V0.1.1, exceto o título e a nota das linhas 1 e 3. |
| Parametros | Bloco novo (A20:C24) no mesmo padrão visual da V0.1; a primeira versão (colunas H:J) colidia com o texto longo da coluna C e foi movida. |

**NÃO inspecionado:** abas Consumo (páginas 13 a 16 do PDF), Qualidade, Materiais e Estoque (dados/tabelas idênticos aos da V0.1.1, comparados célula a célula por teste automatizado, mas não revistos visualmente nesta versão); páginas 8 em diante da aba Analises; **qualquer visualização no Microsoft Excel real**.

**Limites do método.** O LibreOffice difere do Excel em fontes, espaçamento interno dos gráficos, posição de rótulos e tratamento do eixo secundário. As proporções dos gráficos e a quebra de linha dos textos de leitura podem parecer diferentes no Excel. Além da inspeção visual, há verificações automatizadas de renderização (testes de "Inspeção visual (automatizada)" no relatório: texto de cada página do PDF e contagem de pixels das cores das séries).

**Prévias salvas:** `images/dashboard_v0_2_pagina_1.png`, `_2.png` e `_3.png` (geradas pela suíte de testes a partir do XLSX entregue).
