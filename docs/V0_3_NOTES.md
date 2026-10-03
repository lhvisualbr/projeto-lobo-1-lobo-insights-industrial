# Lobo Insights Industrial — V0.3.0

## Executive Intelligence Foundation

**Status:** RELEASED
**Tag oficial:** `v0.3.0`
**Data de fechamento:** 03/10/2026
**Baseline histórica preservada:** `v0.2.1`

> Este documento descreve a release V0.3.0 já publicada.
> Atualizações documentais posteriores na branch `main` não alteram a tag, o ZIP ou o SHA-256 da release oficial.

---

# 1. Visão geral

A V0.3.0 estabelece a primeira camada estruturada de inteligência executiva do **Lobo Insights Industrial**.

O objetivo desta versão é transformar dados industriais sintéticos em sinais:

- determinísticos;
- explicáveis;
- priorizados;
- rastreáveis;
- reproduzíveis;
- sustentados por evidências.

A versão não depende de IA generativa externa para executar o núcleo analítico.

O motor funciona localmente e produz resultados estruturados que podem futuramente alimentar dashboards, APIs, automações, aplicações web ou modelos de linguagem.

---

# 2. Princípio da versão

> **Inteligência verificável antes de inteligência generativa.**

Os indicadores, riscos, classificações e recomendações-base são calculados por código.

Uma futura camada de linguagem natural poderá apresentar os fatos produzidos pelo motor, mas não deverá inventar:

- números;
- materiais;
- custos;
- consumos;
- fornecedores;
- níveis de estoque;
- conclusões sem evidência.

---

# 3. Baseline histórica

A V0.2.1 permanece preservada como baseline histórica.

Tag:

`v0.2.1`

Commit de importação da baseline:

`0e8ccf6 — chore: import official V0.2.1 baseline`

SHA-256 do pacote oficial V0.2.1:

`c5e14aba5c436cdfc2d2f62ff3a5d0c6be301184331ec8c29fc8c827e35a9117`

A implementação da V0.3.0 não substituiu retroativamente:

- scripts históricos;
- dados históricos;
- arquivos Excel históricos;
- manifestos anteriores;
- documentação das releases anteriores.

---

# 4. Executive Intelligence Engine

Arquivo principal:

`src/intelligence/engine.py`

Fluxo consolidado:

```text
dados
→ features analíticas
→ regras determinísticas
→ sinais estruturados
→ scoring
→ priorização
→ recomendações
→ indicadores
→ resumo executivo
→ relatório