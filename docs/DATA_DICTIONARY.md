# Dicionário de dados — V0.1

> Dados 100% sintéticos utilizados exclusivamente para estudo e portfólio.

Convenções: CSV com separador `;`, UTF-8 com BOM, decimal com ponto. A coluna **Exemplo** traz um valor que existe de fato no arquivo (conferido por teste automatizado). Valores monetários são fictícios.

## Colunas dos CSVs

| Tabela | Coluna | Tipo | Descrição | Regra | Exemplo |
|---|---|---|---|---|---|
| materiais_ficticios | Material_ID | Texto (chave primária) | Identificador do material no catálogo | Único; formato MATnnn; MAT001 a MAT020 | MAT001 |
| materiais_ficticios | Material | Texto | Nome fictício padronizado do material | Não vazio; igual ao usado em estoque e consumo | Disco de Desbaste 7 pol |
| materiais_ficticios | Categoria | Texto | Família do material | Abrasivos, Soldagem, Corte, Elétrica, EPI ou Ferramentas | Abrasivos |
| materiais_ficticios | Unidade | Texto | Unidade de medida do item | UN, CX, PCT ou PAR | UN |
| materiais_ficticios | Custo_Unitario | Decimal (2 casas) | Custo por unidade, valor fictício | Maior que zero; ponto decimal | 9.80 |
| materiais_ficticios | Criticidade | Texto | Importância do item para a operação fictícia | Baixa, Média ou Alta | Média |
| materiais_ficticios | Fornecedor | Texto | Fornecedor fictício | Começa com "Fornecedor "; nomes fictícios | Fornecedor Beta |
| estoque_ficticio | Material_ID | Texto (chave estrangeira) | Material a que a posição se refere | Existe no catálogo; um registro por material | MAT001 |
| estoque_ficticio | Material | Texto | Nome do material | Igual ao catálogo | Disco de Desbaste 7 pol |
| estoque_ficticio | Categoria | Texto | Categoria do material | Igual ao catálogo | Abrasivos |
| estoque_ficticio | Unidade | Texto | Unidade do material | Igual ao catálogo | UN |
| estoque_ficticio | Estoque_Atual | Inteiro | Quantidade em estoque no snapshot fictício | Maior ou igual a zero | 120 |
| estoque_ficticio | Estoque_Minimo | Inteiro | Quantidade mínima desejada | Maior ou igual a zero | 40 |
| estoque_ficticio | Custo_Unitario | Decimal (2 casas) | Custo por unidade | Maior que zero; igual ao catálogo | 9.80 |
| estoque_ficticio | Lead_Time_Dias | Inteiro | Prazo fictício de reposição em dias | Inteiro maior que zero | 7 |
| estoque_ficticio | Fornecedor | Texto | Fornecedor fictício | Igual ao catálogo | Fornecedor Beta |
| estoque_ficticio | Criticidade | Texto | Criticidade do item | Baixa, Média ou Alta; igual ao catálogo | Média |
| consumo_ficticio | Movimento_ID | Texto (chave primária) | Identificador da movimentação | Único; formato MOVnnnn | MOV0001 |
| consumo_ficticio | Data | Data (aaaa-mm-dd) | Dia da movimentação | Válida; entre 2026-06-01 e 2026-08-31 | 2026-06-01 |
| consumo_ficticio | Material_ID | Texto (chave estrangeira) | Material movimentado | Existe no catálogo | MAT004 |
| consumo_ficticio | Material | Texto | Nome do material | Igual ao catálogo | Bico de Contato MIG 1,0 mm |
| consumo_ficticio | Categoria | Texto | Categoria do material | Igual ao catálogo | Soldagem |
| consumo_ficticio | Quantidade | Inteiro | Unidades movimentadas | Inteiro maior que zero | 2 |
| consumo_ficticio | Unidade | Texto | Unidade do material | Igual ao catálogo | PCT |
| consumo_ficticio | Custo_Unitario | Decimal (2 casas) | Custo por unidade | Maior que zero; igual ao catálogo | 45.00 |
| consumo_ficticio | Centro_Trabalho | Texto | Centro de trabalho fictício que consumiu | MANUTENCAO, SOLDAGEM, FABRICACAO, MECANICA ou ELETRICA | SOLDAGEM |
| consumo_ficticio | Tipo_Movimentacao | Texto | Tipo do registro | Consumo ou Ajuste (todos são Consumo na V0.1) | Consumo |

## Colunas calculadas (somente no Excel)

Não existem nos CSVs; são fórmulas nas tabelas estruturadas do Excel (cabeçalho verde).

| Tabela | Coluna | Tipo | Descrição | Regra | Exemplo |
|---|---|---|---|---|---|
| tbEstoque | Status_Estoque | Texto (fórmula) | Estado atual do estoque | REPOR se Atual menor ou igual ao Mínimo; ATENÇÃO se maior que o Mínimo e até Mínimo x 1,25; senão OK | REPOR |
| tbEstoque | Critico_Em_Risco | Texto (fórmula) | Item de alta criticidade abaixo ou igual ao mínimo | Sim se Criticidade Alta e Status REPOR; senão Não | Sim |
| tbConsumo | Custo_Movimentado | Decimal (fórmula) | Custo estimado da movimentação | Quantidade x Custo_Unitario, arredondado a 2 casas | 90.00 |
| tbConsumo | Mes | Texto (fórmula) | Ano e mês da data | Formato aaaa-mm | 2026-06 |

## Relacionamentos

- `estoque_ficticio.Material_ID` e `consumo_ficticio.Material_ID` referenciam `materiais_ficticios.Material_ID` (integridade referencial).
- Nome, categoria, unidade, custo, criticidade e fornecedor aparecem repetidos entre tabelas por exigência do enunciado; por isso são conferidos contra o catálogo (verificações Q14 a Q23).

## Parâmetros (aba Parametros do Excel)

| Parâmetro | Valor | Uso |
|---|---|---|
| Margem_ATENCAO | 25% | Define a faixa ATENÇÃO do status de estoque |
| Inicio_Periodo / Fim_Periodo | 2026-06-01 / 2026-08-31 | Período válido de datas e meses do resumo |
