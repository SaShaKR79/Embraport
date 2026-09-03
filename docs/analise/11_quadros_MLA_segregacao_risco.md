# Quadros da MLA — Segregação de risco do custo (CPC 15 × CPC 19) e venda futura

> Explicação compilada dos blocos em vermelho do draft2 da MLA (`Exercicio_Incorporaçao_draft2_MLA_31.08`, aba "base 30 06 26com dividendos", região N76:AA102) e do PPT conclusivo da SF ("Implicações Fiscais da Contabilização da Transação", minuta 31.08.2026). Reproduzidos por fórmula na aba **"6. Segregação de Risco"** da planilha `Exercicio_Incorporacao_CPC15_v1.1.xlsx`. Valores em R$ milhões, base 30.06.26 (PL Embracon 239,5; PL CNP 838,7; VJ JO 5.888 do exercício DTT).

## 1. O que cada quadro calcula

| Quadro (draft) | Pergunta que responde | Mecânica | Resultado |
|---|---|---|---|
| **MLA – Segregação risco custo do investimento** (CPC 15) | De que é feito o custo "gordinho" da posição final de 60% das holdings, e quanto dele é seguro? | TOTAL = PL combinado (239,5 + 548,2 + 290,6 = 1.078,2), aberto em 3 camadas com a proporção DTT (CNP: 65,4% contábil / 34,6% AVJ). DEPOIS = 60% × 1.078,2 = **646,9**: custo antigo remanescente 239,5 − 31,4 (custo baixado na secundária, "perda diluição") = **208,1**; AVJ = 60% × 290,6 = **174,3**; reflexo PL = residual = **264,5** (draft T85). | Δ total +407,4. Camadas de risco crescente: 208,1 (incontroverso) → 264,5 (reflexo PL, neutralidade defensável) → 174,3 (reflexo AVJ, maior risco). |
| **Segregação Deloitte – proporção** | Qual parte do PL da CNP é "AVJ"? | Exemplo DTT: VJ 866 = 566 contábil + 300 MV → 34,6% / 65,4%, aplicados ao PL contábil da CNP (838,7). | 548,2 contábil / 290,6 AVJ. |
| **VJ JO / VJ CNP / VJ Embracon** | Valor justo do negócio combinado | 5.888 (DTT) − 866 = 5.022. | Base do cenário CPC 19. |
| **MLA – Segregação AVJ reflexo holdings** (CPC 19) | Como fica o investimento das holdings se for joint operation? | Inv JO contábil = 239,5 + 548,2 = 787,7 (60% = **472,6**); AVJ JO = 5.888 − 787,7 = 5.100,3 (60% = **3.060,2**); Novo Embracon a VJ = 5.888 (60% = **3.532,8**). | Carimbo = 34% × 3.060,2 = **1.040,5** (passivo diferido / subconta), devido na realização. |
| **VENDA PJ / VENDA PF** | Quanto de imposto na venda futura dos 60%, em 3 preços (1.000 / 3.533 = VJ / 5.000)? | CPC 15 PJ = (preço − 646,9) × 34%. CPC 19 PJ = máx[carimbo 1.040,5; (preço − 472,6) × 34%]. CPC 15 PF = (preço − 18) × 22,5%. CPC 19 PF = carimbo + 22,5% × excedente sobre o VJ. | PJ: 120/981/1.480 (CPC 15) × 1.040/1.040/1.539 (CPC 19). PF: 221/791/1.121 × 1.040/1.040/1.371. |
| **VENDA PJ – RISCO 2** | E se o fisco negar a camada AVJ (custo cai para 472,6)? | (preço − 472,6) × 34%. | 179/1.040/1.539 — coincide com o CPC 19: degrau de +59,3 (= 34% × 174,3). |
| **VENDA PJ – RISCO 1** | E se negar também o reflexo PL (custo cai para 208,1)? | (preço − 208,1) × 34%. | 269/1.130/1.629 — degrau adicional de +89,9 (= 34% × 264,5). |
| **Baixa / Entra / Resultado** (auxiliar CPC 19) | Por que a carga do CPC 19 não cai quando o preço é menor que o VJ? | Venda a 1.000 do investimento a 3.532,8: prejuízo contábil 2.532,8 → crédito teórico 861,2; carimbo 1.040,5 devido → líquido 179,3 (= 34% × (1.000 − 472,6)) **só se o prejuízo compensasse**; leitura gravosa: 179,3 + 1.040,5 = 1.219,8. | O draft adota o piso de 1.040 (perda não compensa de imediato — prejuízo fiscal limitado a 30%/ano). |

## 2. Conclusões (PPT SF 31.08 + quadros)

1. **CPC 15**: só o PL da CNP é reavaliado; o reflexo nas holdings (MEP) tem tese de neutralidade; o risco fica limitado ao reflexo do VJ da CNP — e, dentro dele, concentra-se na camada AVJ (174,3 nos 60%).
2. **CPC 19**: reavaliação de CNP **e** Embracon; ganho de AVJ diferido e carimbado a 34% (1.040,5 nos 60%), com subconta e passivo diferido, tributável na realização "independentemente da entidade que alienar" — a PF não escapa.
3. **No preço = VJ**, CPC 19 custa 59,3 a mais que CPC 15 (PJ) — exatamente o imposto sobre a camada AVJ. Acima do VJ, CPC 19 tributa 34% do excedente sobre a base contábil (472,6) e CPC 15 sobre o custo gordinho (646,9).
4. **Venda pela PF** é vantajosa apenas no CPC 15 (791 × 981 no VJ) e pressupõe "subir pra física"; no CPC 19 o carimbo permanece na holding.
5. **Ressalvas do PPT**: holdings no presumido (no lucro real com prejuízos, a carga futura pode cair); ITCMD fora do escopo, mas a JO tende a elevar a base; números ilustrativos, não auditados.

## 3. Atualização com as premissas de 31.07 (draft2 v3)

PL Embracon 254,0; VJ acordado 5.954,8 (Embracon 5.132,9 / CNP 821,9); custo baixado 66,6 (26,2%): custo 60% = 655,7 (187,5 + 293,9 + 174,3); JO contábil 481,3; AVJ 3.091,6; carimbo 1.051,1; venda ao VJ (3.572,9): CPC 15 991,9 × CPC 19 1.051,1 (PJ), 799,8 × 1.051,1 (PF); riscos: +59,3 (AVJ) e +99,9 (PL).

## 4. Critérios e cuidados

- **Custo baixado na secundária**: o draft usa custo × p.p. vendidos (13,1% por holding = 31,4 no total; na versão 31.07, 26,2% × 254,0 = 66,6). A aba 3 da nossa planilha usa rateio proporcional (13,1 ÷ 43,1). A aba 6 reproduz o critério do draft para bater com os quadros; a premissa é uma célula.
- **Reflexo PL como residual**: no draft o reflexo do PL contábil é o "plug" para fechar 60% × PL combinado (não 60% × 548,2). Mantido como no draft, com memo explicando.
- **Piso do CPC 19**: o draft "trava" 1.040 nos preços ≤ VJ; na aba 6 isso é a função MÁXIMO.
- **"18" da venda PF**: custo das PFs no draft (≈ custo Molina de 17,8 por família); mantido como input.
