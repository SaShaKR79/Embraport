# 🧭 CHECKPOINT — Projeto Another House (Embracon × CNP)

> **Ponto de ancoragem do projeto.** Leia este arquivo primeiro ao iniciar uma nova sessão/chat.
> Ele resume o caso, o que já foi feito, onde está cada coisa e o que falta.
> Última atualização: **28.08.2026** (sessão de análise completa + dashboard).

---

## 1. O caso em 10 linhas

- **Operação**: a Embracon Adm. de Consórcio S.A. **incorpora a CNP Consórcio S.A.** (incorporação plena, art. 227 LSA — versão da totalidade do patrimônio com extinção da CNP; aprovada pelo BACEN, Ofício 31070/2026 de 17.08.2026). As Partes CNP recebem **13,8%** da Embracon (relação de troca: EV CNP R$ 821,9 mi × EV Embracon R$ 5.132,9 mi, data-base 31.12.2024) e, em ato seguinte, **compram ~26,2%** das holdings Savian/JVFJ por **R$ 1,2 bi + correções** (earn-out de até R$ 680 mi), fechando **60% famílias / 40% CNP**. Acordo assinado em **20.10.2025**; closing pendente.
- **Controvérsia central**: caracterização contábil da incorporação — **CPC 15 (combinação de negócios)**, posição do contrato e da Stocche Forbes, × **CPC 19 (joint operation)**, posição da Deloitte (auditora, que decide) baseada em "co-controle" pelos vetos qualificados da CNP. A CNP "precisava que fosse joint operation para reconhecer o resultado" (Felipe Argemi/DTT).
- **Por que importa**: o enquadramento define a natureza da mais-valia/AVJ, a contrapartida no PL (ORA × reserva), o passivo fiscal diferido, o reflexo MEP nas holdings e, sobretudo, **o custo fiscal na alienação futura** ("custo gordão" nunca tributado — tese WTorre/art. 33 §2º DL 1.598 — × "custo magrinho" com diferimento carimbado a 34% em subconta).

## 2. Estado do projeto (o que já foi feito)

| Etapa | Status | Onde está |
|---|---|---|
| Extração de texto de todos os documentos | ✅ | (temporário na sessão; fontes no raiz do repo) |
| Análise dos calls (transcrição 26.08 + resumo) | ✅ | `docs/analise/02_analise_calls.md` |
| Controvérsia contábil (call Deloitte 17.08 + e-mails Malu 18–21.08 + PPT Reflexo MV) | ✅ | `docs/analise/03_analise_controversia_contabil.md` |
| Acordo de Investimento (20.10.2025, 989 p.) | ✅ | `docs/analise/04_analise_acordo_investimento.md` |
| Estrutura societária (deck SF 13.05.2025 + Ofício BACEN) | ✅ | `docs/analise/05_analise_estrutura_societaria.md` |
| Planilha do Molina (simulação tributária 21.05.2025) | ✅ | `docs/analise/06_analise_planilha_molina.md` |
| Razonetes da Malu (3 PNGs — cenário CPC 19) | ✅ | `docs/analise/07_analise_razonetes_malu.md` |
| Artigo doutrinário ágio/mais-valia Lei 12.973 (59 p., lido visualmente) | ✅ | `docs/analise/08…` e `09_analise_artigo_agio_p2.md` |
| **Síntese consolidada do caso** | ✅ | `docs/analise/01_sintese_caso.md` ⭐ |
| Spec de design do dashboard modelo (Grupo Primo) | ✅ | `docs/analise/10_design_spec_modelo_dashboard.md` |
| **Dashboard da operação** | ✅ | `dashboard/dashboard_embracon_cnp.html` |
| **Planilha expositiva CPC 15** (2 hipóteses de custo: "magrinho" × "gordinho") | ✅ | `planilhas/Exercicio_Incorporacao_CPC15_v3.xlsx` — **versão vigente** (aba "Contabilização" com seção do acervo vertido) (v1 mantida como histórico; geradores `build_planilha_cpc15*.py`) |

## 3. Números-chave (conferidos nas fontes)

| Item | Valor | Fonte |
|---|---|---|
| EV Embracon (data-base 31.12.2024) | R$ 5.132.886.600 | Acordo |
| EV CNP Consórcio (data-base 31.12.2024) | R$ 821.924.400 | Acordo |
| Participação CNP pós-incorporação | 13,8% (Savian 43,10% + JVFJ 43,10%) | Acordo/calls |
| Preço de Aquisição (secundária ~26,2%) | R$ 1,2 bi + IPCA desde 31.03.2025 + correção sobre R$ 360 mi | Acordo |
| Earn-out (Parcela Contingente) | até R$ 680 mi (EBITDA + Taxa Futura Bruta, janelas 3+2 anos) | Acordo |
| Resultado final | 60% famílias (30/30) × 40% CNP | Acordo |
| PL Embracon 12/2024 | R$ 178.389.573,07 | Planilha Molina |
| Prejuízo fiscal Embracon (IRPJ = CSLL) | R$ 32.350.933,88 | Planilha Molina |
| Planilha Molina (versão 21.05.2025 — estrutura antiga: venda de 20% por R$ 1,5 bi) | GK holdings R$ 497,9 mi × GK PFs R$ 323,5 mi; ajuste de preço R$ 132,1 mi | Planilha Molina |
| Ilustrativos Deloitte (call 17.08) | Invest. CNP 300; Embracon 178; 40% da cia = 2.355; AVJ 3.426 em resultado; 14,7% ações + 25,3% caixa | Call Deloitte |
| Razonetes Malu (ilustrativo CPC 19) | Valuation 100% = 5.888; JO comprador 2.355 (DTL 155); vendedor: remensuração 60% → ganho 4.955 em resultado (DTL 1.165) | Razonetes |

⚠️ Os números da planilha do Molina e do deck SF refletem **estruturas anteriores** (venda de 20%/25% pelas PFs) superadas pelo Acordo assinado. Os ilustrativos da Deloitte e da Malu usam escalas diferentes (não são os valores contratuais).

## 4. Pendências do caso (substância)

1. **Enquadramento CPC 15 × CPC 19** — decide o auditor da Embracon (Deloitte); SF quer discutir a premissa de "co-controle" com o Ricardo/DTT, advogados da CNP e o accounting da companhia. Não trava o closing.
2. **Conta de contrapartida no PL** (ORA × reserva de capital) — Deloitte nunca respondeu.
3. **Aprimorar a planilha do Molina** (pedido do Renato, e-mail 19.08): efeitos numéricos contábeis e fiscais segregando **ganho por variação de % puro** × **efeito reflexo questionável**; 2 cenários de alienação futura (custo "gordão" × "magrinho"); sensibilidade de preço para evidenciar o AVJ.
4. **Rendimentos das aplicações financeiras do acervo incorporado** — natureza/tributação **não tratada em nenhum documento até aqui** (ponto citado pelo usuário como parte do escopo).
5. Deck visual do fluxo da operação (action item do call 26.08) — o dashboard deste repo atende parcialmente.
6. Fundamentação técnica da Deloitte para a JO (prometida, não recebida); planilha de referência do Gabriel/"Cheng"; pesquisa no "Jet".

## 5. Estrutura do repositório

```
├── docs/analise/          ← base de conhecimento (esta pasta) — 00 a 10
├── dashboard/             ← dashboard HTML (padrão visual Grupo Primo)
├── planilhas/             ← planilha expositiva CPC 15 (v1) + draft MLA fonte + script gerador
├── *.pdf, *.pptx, *.png, *.txt, *.xlsx  ← documentos-fonte do caso (raiz)
```

### Nota sobre a planilha CPC 15 (sessão 31.08.2026)
- **v3 (vigente)**: 3 abas — **Premissas** (somente valores do draft MLA: 13,8%; 40%; 34%; dividendos 120; preço 600/holding; segregação DTT VJ 866/MV 300) · **Contabilização** (fluxo sequencial em aba única: 0. balanços-base 30.06.26 → 1. pré-incorporação → 2. incorporação com reflexo e segregação de risco em 3 camadas → **3. situação após a incorporação: acervo vertido/cindido linha a linha (839,6 com check de 0,84), balanço combinado da Embracon, balanços das holdings após e prova da somatória (148,8 CNP / 929,4 S+J / 464,7 por holding / diluição individual 74,4)** → 4. cash out com as 2 hipóteses → 5. alienação futura com preço editável, cada bloco com quadro conclusivo) · **Conclusão** (síntese A × B + leitura de riscos). 156 fórmulas, recalculada sem erros. Formatação: fontes pretas; vermelho só em barras de título e negativos; amarelo = inputs editáveis.
- v1 (8 abas, mais analítica) e v2 mantidas como histórico em `planilhas/`.
- Base de valores: draft da MLA (`Exercicio_Incorporaçao_draft1`), aba **"base 30 06 26com dividendos"** (Embracon PL 239,488 após dividendos de 120; CNP PL 838,735).
- Números-síntese: parcela incorporada 344,97/holding (689,94 total); custo gordinho 464,71 × magrinho 119,74; IR cash out 383,25 (A) × 311,95 (B); benefício total da tese do custo gordinho = **234,58 = 34% × 689,94** (71,30 no cash out + até 163,28 na alienação futura).
- **Critério de rateio do custo no cash out**: proporcional à fração alienada (13,1/43,1 = 30,39%) — o draft multiplicava o custo por 13,1% direto; ajuste mantido na v2.
- Diferenças de fechamento dos balanços do draft (Embracon −0,34; CNP +0,84) não afetam os cálculos (v2 parte das composições de PL).

Branch de trabalho: `claude/asset-incorporation-analysis-p46li6`.

## 6. Como retomar em nova sessão

1. Leia este checkpoint e a **síntese** (`01_sintese_caso.md`).
2. Para aprofundar um tema, vá ao arquivo temático correspondente (02–09).
3. O dashboard (`dashboard/dashboard_embracon_cnp.html`) consolida a visão executiva; o spec de design (10) permite mantê-lo fiel ao modelo.
4. Tarefas naturais da próxima sessão: item 3 e 4 das pendências acima (planilha aprimorada com cenários numéricos CPC 15 × CPC 19 e análise dos rendimentos do acervo).

## 7. Glossário mínimo

**Partes CNP**: CNP Assurances Participações, CNP Assurances S.A., CNP Latam Holding · **Holdings/famílias**: Cia. Savian (família Savian — Guido 90%) e Cia. JVFJ (família Silva/Dutra — Juarez usufrutuário de 100%) · **Molina**: José Luzia Molina, consultor contábil da Embracon (autor da planilha) · **Malu**: Maria Luiza Assad (SF, autora da tabela CPC 15×19 e dos razonetes) · **RC**: Renato Coelho (SF, sócio) · **DTT**: Deloitte (Mayara, Pammela, Felipe Argemi, "Ricardo") · **AVJ**: avaliação a valor justo · **MV**: mais-valia · **ORA**: outros resultados abrangentes · **DTL**: passivo fiscal diferido · **JO**: joint operation · **Passo F**: futura aquisição de controle pela CNP (secundária, sem preço definido) · **WTorre**: Ac. CARF 1402-004.537/2020 (neutralidade definitiva) · **CARF 2024**: CPFL 1101-001.409, Litela 1101-001.404, Litel 1401-007.055 (diferimento; tributa ao migrar ORA→DRE).
