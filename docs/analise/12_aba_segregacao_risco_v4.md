# 12 — Aba "Segregação de Risco" na planilha-entregável (v4, 04.09.2026)

> Reconstrução da aba 6 (antes em `Exercicio_Incorporacao_CPC15_v1.1.xlsx`) dentro da planilha que o usuário
> está usando como entregável: `planilhas/Embracon_simulacao_tributaria_04.09.26_v4.xlsx`
> (fonte do usuário: `..._v3 (fonte usuario).xlsx`). Gerador: `planilhas/build_segregacao_sim.py`
> (+ `inject_values.py`), que edita o pacote OOXML diretamente — preserva logo EMF, abas ocultas e
> formatação; só muda `sheet3.xml`, `styles.xml` (acréscimos), `workbook.xml` (fullCalcOnLoad) e remove `calcChain`.

## 1. Premissas (fonte: aba "Cálculos da Operação", por vínculo)

| Item | Valor | Célula-fonte |
|---|---|---|
| PL Embracon = custo original das Holdings (127 + 127) | 254 | C19 / F47 |
| PL contábil CNP | 800 | C23 |
| Valor de mercado CNP (valor incorporado no CPC 15 — PPA) | 822 | C24 |
| Mais-valia CNP (VM − PL) | 22 | calculado |
| Relação de troca / alvo CNP / famílias | 13,8% / 40% / 60% | I15 / I16 |
| Fração alienada no cash-out (13,1 ÷ 43,1) | 30,39% | D69 |
| IRPJ/CSLL | 34% | I17 |
| **Premissas locais da aba** | VJ Embracon 5.133 (EV do Acordo); preços 1.000 / VJ / 5.000; PF 22,5%; custo PF 18 (ilustrativo MLA) | inputs cinza |

Conferências embutidas (mostram "ok"): custo incrementado 927,5 = E53; custo remanescente 645,6 = J89;
baixa no cash-out 281,9 = J82; incremento em discussão 468,8 = K89.

## 2. Resultados (R$ mi, Savian + JVFJ)

**Bloco 1 — CPC 15, camadas do custo** (100% → 86,2% pós-incorporação → 60% após cash-out):
custo original 254 → 254 → 176,8 · diluição — → (35,1) → (24,4) · reflexo PL contábil CNP 800 → 689,6 → 480,0 ·
mais-valia 22 → 19,0 → 13,2 · **total 1.076 → 927,5 → 645,6**. Leitura após o cash-out: incontroverso 176,8 (cenário A);
em discussão 468,8 = reflexo PL líquido da diluição 455,6 (IR 154,9) + mais-valia 13,2 (IR 4,5); IR potencial total 159,4.

**Bloco 2 — CPC 19**: VJ JO 5.955 (60% = 3.573) − base contábil 1.054 (632,4) = AVJ 4.901 (2.940,6) → carimbo 34% = 999,8.
Custo após cash-out: CPC 15 645,6 × CPC 19 3.573 (Δ 2.927,4).

**Bloco 3 — venda dos 60%** (CPC 15 PJ / PF · CPC 19 PJ / PF):
1.000 → 120,5 / 221,0 · 999,8 / 999,8 · 3.573 → 995,3 / 799,9 · 999,8 / 999,8 · 5.000 → 1.480,5 / 1.121,0 · 1.485,0 / 1.320,9.
Δ CPC 19 − CPC 15 (PJ): 879,3 / 4,5 / 4,5 (= 34% × mais-valia 13,2 nos preços ≥ VJ).
Piso do carimbo (venda a 1.000): baixa (3.573) + preço 1.000 = (2.573); carimbo 999,8; crédito teórico (874,8);
líquido teórico 125,0 (= RISCO 2); adotado 999,8 (presumido / trava de 30%).

**Bloco 4 — sensibilidade CPC 15**: custos 645,6 (integral) / 632,4 (RISCO 2) / 176,8 (RISCO 1);
IR a 1.000 / VJ / 5.000: 120,5 · 995,3 · 1.480,5 | 125,0 · 999,8 · 1.485,0 | 279,9 · 1.154,7 · 1.639,9;
degraus fixos 4,5 (RISCO 2) e 159,4 (RISCO 1). RISCO 2 = CPC 19 a partir do VJ; RISCO 1 > CPC 19 a partir do VJ (154,9).

**Bloco 5 — conclusão executiva**: tabela CPC 15 × CPC 19 × Δ com leitura por linha + 3 bullets + ressalvas
(presumido/LR, ITCMD, custo PF ilustrativo, cash-out no CPC 19 não tratado).

## 3. Diferenças em relação à versão anterior (v1.1 / MLA)

- Mais-valia da CNP passou a ser **VM − PL = 22** (antes: 34,6% do PL, proporção DTT 300/866). Consequência: a camada
  "mais exposta" ficou residual (13,2 → IR 4,5) e o risco relevante do CPC 15 migrou para a neutralidade do reflexo do PL
  contábil (455,6 → IR 154,9). Se o PPA identificar mais-valia maior, a exposição cresce 34% × 60% × mais-valia.
- Critério de baixa no cash-out: **fração 30,39%** (o mesmo da aba Cálculos), não mais 13,1% × custo (MLA).
- Reflexo do PL contábil não é mais residual: calculado diretamente (PL × 86,2% / 60%); a diluição aparece em linha própria.
- Premissas: bloco mínimo ("Informações utilizadas", vinculado) + premissas locais em cada bloco (pedido do usuário).
- Explicações por bloco (parágrafo introdutório + "Leitura" por tabela/linha) para leitura do cliente.

## 4. Observações sobre a planilha do usuário (não alteradas)

- `Cálculos da Operação`!G5 ainda diz "21 de maio de 2025" (data do modelo antigo).
- A CNP é incorporada pelo valor de mercado (822) e não pelo PL (800): coerente com o CPC 15 (PPA); a diferença é a mais-valia.
- O esboço original da aba trazia "PL Embracon | 30.06.2026" (rótulo trocado por "Data-base dos balanços" na nova aba).
- Totais da aba Cálculos conferidos: economia no cash-out K86 = 69,6 = 2 × 34,8; Δ custo remanescente K89 = 468,8.
