# Análise da Planilha de Simulação Tributária (Molina, 21.05.2025)

**Arquivo:** `Embracon_simulação tributária 21MAI2025.xlsx` — elaborada pelo consultor "Molina"
**Projeto:** "Projeto Another House" (Capa!B11) — grupo Embracon x grupo CNP Assurances
**Data do documento:** 21 de maio de 2025 (Capa!B15)

---

## 1. Visão geral

A planilha tem por objeto os **cálculos tributários da parcela de aquisição secundária** da operação: a venda, pelos acionistas do grupo Embracon, de **20% da Embracon Administradora de Consórcio** por **R$ 1,5 bilhão** (mais earn-out ainda indefinido). Ela compara dois caminhos de venda — (1) venda pelas holdings Savian e JVFJ (pessoas jurídicas) e (2) venda diretamente pelas pessoas físicas — quantifica a economia fiscal entre eles (R$ 174,35 MM) e deriva um **ajuste de preço negocial** (R$ 132,09 MM com gross-up). Complementarmente, consolida balanços de 12/2024 e informações de LALUR das principais empresas do grupo Embracon, além de dados públicos da CNP Seguros.

**Estrutura das abas (10):**

| Aba | Conteúdo |
|---|---|
| Capa | Título "Projeto Another House / Cálculos tributários / 21 de maio de 2025" |
| Organograma | Sem células; contém as imagens dos organogramas societários |
| Cálculo GK | Premissas econômicas + Cenário 1 (Holdings) + Cenário 2 (PF) + ajuste de preço |
| Embracon Adm. de Consórcio | Balanço 12/2024, prejuízos fiscais, itens LALUR, tabela de incentivos fiscais |
| Dourada C. e A. SA | Balanço 12/2024, prejuízos fiscais, incentivos fiscais |
| Embracon Franchising | Balanço 12/2024 (lucro presumido) |
| Embrafisa Consul. | Balanço 12/2024 (lucro presumido) + DRE 2024/2023 (imagem) |
| Savian | Balanço 12/2024 da Companhia Savian de Participações |
| GSJ | Balanço 12/2024 da Companhia GSJ de Participações |
| CNP Seguros | Apenas nota "Informação coletada da internet" + imagem do Balanço Patrimonial publicado |

Fontes citadas pelo Molina nas abas: "Excel: LALUR", "Excel LALUR 2024", "Balanço", "Balanço Savian 12_2024", "GSJ Balanço", "Informação coletada da internet" (CNP).

---

## 2. Organogramas e imagens embutidas

**image1.emf** — arquivo EMF (Windows Enhanced Metafile), não legível na extração; existe, mas seu conteúdo não pôde ser inspecionado (provavelmente outro diagrama/logotipo vetorial da aba Organograma).

**image2.png — Logotipo "Consórcio Embracon".** Marca da companhia (fundo vermelho, mão estilizada, slogan "Porque sonhar não tem limites"). Sem conteúdo societário.

**image3.png — "Organograma de Composição e Participação Societária na Embracon"** (estrutura restrita ao alvo da transação):
- **Lado esquerdo (família Savian):** Guido Savian Júnior e os filhos Daniel, Gabriel, Lucas, André e Matheus Gelli Savian → **Companhia Savian de Participações**. Tabela de composição: Guido 100,00% ON / — PN / **90,0% total**; cada filho — ON / 20,00% PN / **2,00% total** (soma 100%).
- **Lado direito (família Silva/Dutra):** Juarez Antonio da Silva e os filhos Juarez Dutra da Silva, Vanessa Dutra da Silva Rigolin e Fernanda Dutra da Silva → **Companhia JVFJ de Participações**. Composição: Juarez — ON / 83,33% PN / **6,25% total**; Dutra, Vanessa e Fernanda: 33,33% ON / 5,56% PN / **31,25% total** cada.
- **Centro:** Embracon Administradora de Consórcio Ltda detida **49,99% pela Savian + 49,99% pela JVFJ + 0,01%** (setas partindo de Guido e Juarez pessoas físicas — participação direta residual).
- **Observação no rodapé direito:** "Juarez Antonio da Silva é usufrutuário de 100% das ações (ON+PN)" da JVFJ.

**image4.png — "Organograma de Composição e Participação Societária: Grupo Econômico Embracon"** (grupo completo):
- **Família Savian:** Guido (100% ON) e filhos (100% PN) na **Companhia Savian de Participações**; Guido/filhos também na **Companhia GSJ de Participações** (rótulos "100% (ON)" e "10% (PN)").
- **Família Silva/Dutra:** Juarez e filhos na **Companhia JUVAFE de Participações** (rótulos 93,75% e 6,25%) e na **Companhia JVFJ de Participações** (mesma composição da image3).
- **Operacionais:**
  - **Embracon Administradora de Consórcio Ltda:** 49,99% Savian + 49,99% JVFJ + 0,01% direto;
  - **Embrafisa Consultoria Empresarial, Corretagem de Seguros e Empreendimentos Imobiliários Ltda:** 50% GSJ + 50% JUVAFE;
  - **Embracon Corretora de Seguros Ltda:** 50% / 50% (GSJ / JUVAFE);
  - **Dourada Comercial e Agropecuária S/A:** 50% / 50% (GSJ / JUVAFE).
- Mesma obs. sobre o usufruto de Juarez. Ou seja: o grupo tem **dois eixos de holdings por família** — Savian/JVFJ (que detêm a Embracon, alvo da venda) e GSJ/JUVAFE (que detêm Embrafisa, Corretora e Dourada, fora do perímetro do cálculo de GK).

**image5.png — DRE da Embrafisa (2024 x 2023)**, em R$: Receita operacional líquida 77.764.180,22 (2023: 42.175.692,71); Lucro bruto 76.342.567,02; Despesas operacionais 17.262.123,00; Resultado operacional 59.080.444,02; Resultado financeiro líquido +1.757.665,02; Resultado antes dos impostos 60.791.423,35; IR/CSLL 7.970.652,55 (CSLL 2.166.717,20 + IRPJ 5.803.935,35 — carga efetiva ~13,1%, coerente com lucro presumido); **Lucro líquido 2024: 52.820.770,80** (2023: 26.243.874,53).

**image6.png — Balanço Patrimonial da CNP Seguros** (publicação oficial, em R$ mil, colunas Controladora e Consolidado em 31/12/2023, 31/12/2022 reapresentado e 01/01/2022): Total do ativo controladora **5.476.451** / consolidado **10.437.415**; Patrimônio líquido controladora **5.156.127** / consolidado **5.200.500** (controladores 5.120.352 + não controladores 80.148); Capital social 2.204.000; Investimentos em controladas e coligadas (controladora) 4.655.446; passivos de operações de seguros, provisões para contingências consolidadas 3.920.506, etc. É a única informação do lado CNP na planilha — **coletada da internet**, não de data room.

---

## 3. Premissas econômicas (aba "Cálculo GK", B2:F38)

1. **Preço da parcela de aquisição secundária:** R$ 1.500.000.000 (D4).
2. **Earn-Out:** R$ 0 (D5), com nota **"[A confirmar - entre 300MM a 800MM]"** (E5) — ou seja, o preço final pode chegar a R$ 1,8–2,3 bi; a simulação usa só o preço fixo.
3. **Preço final:** R$ 1.500.000.000 (D6 `=SUM(D4:D5)`).
4. **Percentual da Embracon a ser adquirido (secundária): 20%** (D8), com nota **"[A confirmar]"** (E8). Implica avaliação de 100% da Embracon em **R$ 7,5 bi** (~42x o PL contábil).
5. **Alocação do preço 50/50:** Savian R$ 750.000.000 (D11 `=D6/2`) e JVFJ R$ 750.000.000 (D12 `=D6/2`).
6. **PL da Embracon (12/2024): R$ 178.389.573,07** (D14, vinculado a `'Embracon Adm. de Consórcio'!G13`).
7. **Valor contábil registrado por cada holding na Embracon: R$ 89.194.786,535** (D16 e D17 `=D14/2` — metade do PL, i.e., equivalência patrimonial de 50%). Nota em E17: *"Embora não tenhamos recebido documentação contábil da companhia [JVFJ], estamos considerando que o valor seria equivalente ao da Savian"*.
8. **Composição societária Savian (C19:F27):** Guido 100% ON = 90% total; Daniel, Gabriel, Lucas, André e Matheus 20% PN cada = 2% total cada; soma 100%.
9. **Composição societária JVFJ (C29:F35):** Juarez — ON / 83,33% PN / 6,25% total; Dutra, Vanessa e Fernanda 33,33% ON / 5,56% PN / 31,25% total cada. Somas com arredondamento imperfeito: ON 0,9999 e PN 1,0001 (F35 total = 100%).
10. **Usufruto (C37):** *"Juarez é usufrutuário de 100% das ações e, portanto, consideramos que ele será o único beneficiário da operação"* — no Cenário 2, 100% do preço da JVFJ é atribuído a Juarez; os filhos aparecem com "-".
11. **Custo de aquisição das PF (C38):** *"Conforme fomos informados, estamos considerando que o custo de aquisição das pessoas físicas é igual ao valor contábil da participação registrado pelas Holdings."* — premissa essencial ao Cenário 2 (compatível com devolução de capital a valor contábil, art. 22 da Lei 9.249/95, ainda que a planilha não descreva esse passo).

---

## 4. Cenário 1 — Venda pelas Holdings (B41:D58)

**Ganho de capital SAVIAN (B43:D48):**
- Preço da operação: **R$ 750.000.000,00** (D45 `=D11`)
- Valor contábil da Embracon (proporcional a 20%): **R$ 17.838.957,307** (D46 `=D16*D8` = 89.194.786,535 × 20%)
- Ganho de capital: **R$ 732.161.042,693** (D47 `=D45-D46`)
- IRPJ e CSLL (34%): **R$ 248.934.754,52** (D48 `=D47*34%`)

**Ganho de capital JVFJ (B51:D56):** cálculo idêntico → GK R$ 732.161.042,693; tributo **R$ 248.934.754,52** (D56 `=34%*D55`).

**TOTAL DE TRIBUTOS (Savian + JVFJ): R$ 497.869.509,03** (D58 `=D56+D48`).

**Lógica tributária:** na PJ, o ganho de capital na alienação de participação societária permanente compõe integralmente a base do IRPJ (15% + adicional de 10%) e da CSLL (9%) — carga combinada de **34%**. Isso vale tanto no lucro real quanto no **lucro presumido** (regime das holdings, conforme aba Savian): no presumido o ganho de capital é acrescido diretamente à base trimestral, sem percentual de presunção (art. 25, Lei 9.430/96 / art. 595 RIR/18). A planilha aplica 34% "cheio", desprezando a franquia do adicional (R$ 60 mil/trimestre — imaterial) e assumindo, corretamente, que não há PIS/COFINS sobre a venda de investimento permanente. Não considera compensação de prejuízos fiscais das holdings (não indicados nas abas Savian/GSJ) nem custos de transação dedutíveis.

---

## 5. Cenário 2 — Venda pelas Pessoas Físicas (B61:D83)

**Base legal das alíquotas:** tabela progressiva do ganho de capital de PF, art. 21 da Lei 8.981/95 com redação da **Lei 13.259/2016** (vigente desde 01/01/2017):
- **15%** sobre a parcela do ganho até R$ 5 MM;
- **17,5%** sobre a parcela entre R$ 5 MM e R$ 10 MM;
- **20%** sobre a parcela entre R$ 10 MM e R$ 30 MM;
- **22,5%** sobre a parcela que exceder R$ 30 MM.

A fórmula da coluna G (ex.: G66) implementa exatamente esses degraus por somatório de IFs encadeados: `=IF(F66>5000000;5000000*15%;...) + IF(F66>10000000;5000000*17,5%;...) + IF(F66>30000000;20000000*20%;...) + IF(F66>30000000;(F66-30000000)*22,5%;0)`. Nas três primeiras faixas o imposto "cheio" é R$ 750.000 + R$ 875.000 + R$ 4.000.000 = **R$ 5.625.000 sobre os primeiros R$ 30 MM**.

**Família Savian (C63:G72)** — preço rateado pelo % total de cada sócio (`D66=$D$72*F21` etc.); custo idem sobre o custo total de R$ 17.838.957,307 (E72 `=D16*20%`):

| Sócio | Preço | Custo | Ganho de capital | IRPF |
|---|---:|---:|---:|---:|
| Guido (90%) | 675.000.000,00 | 16.055.061,58 | 658.944.938,42 | **147.137.611,15** |
| Daniel (2%) | 15.000.000,00 | 356.779,15 | 14.643.220,85 | 2.553.644,17 |
| Gabriel (2%) | 15.000.000,00 | 356.779,15 | 14.643.220,85 | 2.553.644,17 |
| Lucas (2%) | 15.000.000,00 | 356.779,15 | 14.643.220,85 | 2.553.644,17 |
| André (2%) | 15.000.000,00 | 356.779,15 | 14.643.220,85 | 2.553.644,17 |
| Matheus (2%) | 15.000.000,00 | 356.779,15 | 14.643.220,85 | 2.553.644,17 |
| **TOTAL** | **750.000.000,00** | **17.838.957,31** | — | **159.905.832,00** (G72) |

Conferência: Guido = 5.625.000 + 22,5% × (658.944.938,42 − 30.000.000) = 147.137.611,15 (alíquota efetiva 22,33%). Cada filho = 750.000 + 875.000 + 20% × 4.643.220,85 = 2.553.644,17 (efetiva 17,44%).

**Família Juarez (B74:G81)** — pelo usufruto, 100% atribuído a Juarez (D77 `=D81`; E77 `=F35*E81`; filhos com "-"):

| Sócio | Preço | Custo | GK | IRPF |
|---|---:|---:|---:|---:|
| Juarez | 750.000.000,00 | 17.838.957,31 | 732.161.042,69 | **163.611.234,61** |
| Dutra / Vanessa / Fernanda | – | – | – | – |
| **TOTAL** | **750.000.000,00** | **17.838.957,31** | — | **163.611.234,61** (G81) |

Conferência: 5.625.000 + 22,5% × 702.161.042,69 = 163.611.234,61 (efetiva 22,35%).

**TOTAL DE TRIBUTOS (Savian + Juarez): R$ 323.517.066,61** (D83 `=G81+G72`). Carga efetiva global do Cenário 2 ≈ **22,09%** do ganho (vs. 34% no Cenário 1).

---

## 6. Cálculo do ajuste de preço (B85:D91)

| Item | Célula/Fórmula | Valor |
|---|---|---:|
| Tributação — venda pelas Holdings | D87 `=D58` | R$ 497.869.509,03 |
| Tributação — venda pelas PF | D88 `=D83` | R$ 323.517.066,61 |
| **Diferença (economia fiscal)** | D89 `=D87-D88` | **R$ 174.352.442,43** |
| 50% da diferença | D90 `=D89/2` | R$ 87.176.221,21 |
| **Ajuste com gross-up (34%)** | D91 `=D90/(1-34%)` | **R$ 132.085.183,66** |

**Lógica negocial:** migrar a venda para as pessoas físicas gera uma economia tributária de ~R$ 174,35 MM (redução de 35% da carga). Como essa eficiência depende de cooperação/aceitação das duas partes (e a estrutura beneficia o vendedor), a praxe de M&A é **compartilhar a sinergia fiscal 50/50** entre comprador (CNP) e vendedores — daí os R$ 87,18 MM. O **gross-up pela alíquota de 34%** (divisão por 0,66) reconstitui o valor bruto necessário para que, após a incidência de tributo de 34% sobre a própria parcela de ajuste, reste líquido exatamente os R$ 87,18 MM: 132.085.183,66 × 66% = 87.176.221,21. A planilha não explicita o sentido do fluxo (aumento ou redução de preço, nem quem paga a quem); o uso de 34% (alíquota de PJ) sugere que a parcela de ajuste transitaria por uma pessoa jurídica tributada a 34% (ex.: ajuste pago/recebido em contexto tributável na PJ), premissa a validar no SPA — se o ajuste fosse recebido por PF (GK a ~22,5%), o gross-up correto seria outro.

---

## 7. Balanços por empresa (12/2024, salvo CNP)

### 7.1 Embracon Adm. de Consórcio ("EMBRACON ADMINISTRATORA DE CONSÓRCIO LTDA" — sic)
- **Regime: Lucro Real** (C3). Data-base 12/2024. Fonte: "Excel: LALUR".
- **Ativos R$ 896.783.712,31** (D4): circulante 847.878.260,36 — destaque para **"Despesa do exercício seguinte" R$ 666.446.778,73** (D12 — ativo de comissões diferidas/CPC 47), aplicações interfinanceiras 96.754.396,60, taxa de administração a receber 42.031.358,23, impostos a compensar 16.585.526,04; não circulante 48.905.451,95 (imobilizado 16.484.284,62 + intangível 32.421.167,33).
- **Passivos:** circulante 718.394.139,24 — destaque para **"Provisões" R$ 617.638.088,48** (G12 — essencialmente a receita diferida CPC 47), obrigações trabalhistas 54.107.743,68, encargos 20.770.727,46.
- **PL R$ 178.389.573,07** (G13): capital social 85.000.000 + reserva legal 5.988.406,16 + retenção de lucros 87.401.166,91. Célula auxiliar G19 `=G13/2` = 89.194.786,535 (valor por holding usado no Cálculo GK).
- **Prejuízos fiscais:** IRPJ **R$ 32.350.933,88** (K3) e CSLL **R$ 32.350.933,88** (K4); K5 soma ambos (R$ 64.701.867,76) — soma conceitualmente imprópria (são bases paralelas, não aditivas; valor potencial ≈ 32,35 MM × 34% ≈ R$ 11 MM, sujeito à trava de 30%).
- **Itens LALUR — diferenças temporárias (adições/exclusões), total K13 = R$ 1.216.164.312,35:**
  - Outras add/exclusões (aquisição carteira PAN/AGIBANK): 6.279.263,19
  - **Receita de prestação de serviço diferida — CPC 47: 533.696.093,06**
  - **Despesas de comissões diferidas — CPC 47: 666.446.778,73**
  - PDD: 4.640.281,59; Provisão Serv. VW: 1.293.840,13; Provisão perda esperada (risco de crédito): 2.999.249,88; Provisão desp. juros passivo: 808.805,77
- **Provisões não dedutíveis, total K18 = R$ 18.872.556,36:** ações trabalhistas 3.049.268,31; serviços Renault 8.059.932,87; ações contrárias 4.308.410,29; despesa com comissão 3.454.944,89.
- **Tabela de incentivos fiscais** (B20:E26): PAT (4% do IRPJ antes do adicional), FIA-Criança/Adolescente (1%), Audiovisual (3%), Rouanet (4%), Fundo do Idoso (1%), Lei de Incentivo ao Esporte (1%) — relevantes por ser lucro real.
- **Leitura tributária:** o par CPC 47 (receita diferida adicionada / comissão diferida excluída, ou vice-versa) gera enorme estoque de diferenças temporárias (~R$ 1,2 bi brutos; efeito líquido receita−comissão = −132.750.685,67, comissões diferidas excedem a receita diferida) — ponto central para IR diferido, purchase accounting e para o valor dos atributos fiscais na incorporação.

### 7.2 Dourada Comercial e Agropecuária S/A
- **Regime: Lucro Real**. Fonte "Excel LALUR 2024".
- Ativos **R$ 10.218.399,67**; imobilizado 8.254.345,43; ativos circulantes 1.964.054,24 (caixa de apenas R$ 2,00).
- PL **R$ 10.053.653,40**: capital 14.139.130 + reserva de reavaliação 4.562.148,38 + reserva legal 48.664,65 **− prejuízos acumulados 8.696.289,63**.
- **Prejuízo fiscal IRPJ 956.469,605 e CSLL 956.469,605** (J3/J4; J5 soma 1.912.939,21 — mesma ressalva da soma). PDD indedutível 34.739,57.
- Mesma tabela de incentivos fiscais da Embracon. Empresa fora do perímetro da venda (detida por GSJ/JUVAFE).

### 7.3 Embracon Franchising Ltda
- **Regime: Lucro Presumido.** Micro balanço: ativos **R$ 1.167.425,47** (TVM 317.739,29; realizável LP 784.017,14); PL **R$ 1.149.700,60** (capital 1.150.000 − a integralizar 200.000 + reserva de lucros 199.700,60). Sem prejuízos/LALUR relevantes.

### 7.4 Embrafisa Consul. Empre. Corre. de Segu. e Empreend. Imob. Ltda
- **Regime: Lucro Presumido.** Ativos **R$ 122.371.162,57**: investimentos 46.794.804,11 + imobilizado 61.134.021,55 no não circulante (107.928.825,66); circulante 14.442.336,91.
- Passivo circulante 21.461.591,92 (contas a pagar 17.920.181,50). **PL R$ 100.909.570,65** (capital 86.545.880 + reserva de lucros 14.363.690,65).
- DRE embutida (image5): **lucro líquido 2024 de R$ 52,82 MM** sobre ROL de R$ 77,76 MM — margem altíssima; presumido muito vantajoso (carga efetiva ~13%). Detida 50/50 por GSJ/JUVAFE — fora do perímetro do GK.

### 7.5 Companhia Savian de Participações
- **Regime: Lucro Presumido.** Fonte "Balanço Savian 12_2024".
- Ativos **R$ 90.240.194,89**: **participações societárias R$ 89.176.947,58** (praticamente só a Embracon) + circulante 1.063.247,31 (impostos a compensar 986.294,79).
- Passivo circulante 898.578,42 (remuneração do capital a pagar 893.596,05). **PL R$ 89.341.616,47** (capital 56.200.000 + reserva de lucros a realizar 29.034.177,88 + reserva legal 4.107.438,59).
- Obs.: o investimento contábil (89.176.947,58) difere em **R$ 17.838,96** (dezessete mil reais) da premissa usada no Cálculo GK (89.194.786,54 = PL Embracon/2) — diferença irrelevante, mas mostra que a premissa foi tomada por MEP teórico (metade do PL) e não pelo razão contábil da holding.

### 7.6 Companhia GSJ de Participações
- **Regime: Lucro Presumido.** Ativos **R$ 109.379.942,84**: participações societárias 64.315.927,33 + outros investimentos 21.372.769,94 + imobilizado 22.870.860,66.
- **PL R$ 108.460.364,76** (capital 79.875.997,74 − a integralizar 2.603.500 + reserva legal 1.800.035,86 + reserva de lucros 7.330.593,32 + lucros acumulados 22.057.237,84). Holding "irmã" (Embrafisa/Corretora/Dourada) — fora do perímetro; a JUVAFE (par da JVFJ) **não tem aba**.

### 7.7 CNP Seguros
- Aba quase vazia: "CNP SEGUROS / **Informação coletada da internet**" + imagem do balanço publicado (31/12/2023, R$ mil): ativo total controladora 5.476.451 / consolidado 10.437.415; **PL controladora 5.156.127** / consolidado 5.200.500; capital social 2.204.000. Nenhum dado do data room da CNP Consórcio S.A. especificamente — o balanço é do grupo segurador, não da administradora de consórcio alvo da incorporação.

---

## 8. Pontos de atenção

1. **Premissas "a confirmar" explícitas:** (i) earn-out entre R$ 300 e 800 MM zerado na conta — se confirmado, os tributos e o ajuste de preço mudam materialmente (no Cenário 2, o GK em recebimentos parcelados é tributado proporcionalmente, mantendo a alíquota apurada sobre o ganho total — art. 21, §3º); (ii) o percentual de 20% adquirido; (iii) o valor contábil da JVFJ (sem documentação — assumido igual ao da Savian); (iv) o custo de aquisição das PF igual ao valor contábil nas holdings ("conforme fomos informados").
2. **O Cenário 2 pressupõe passo prévio não modelado:** as PF hoje não detêm a Embracon diretamente (detêm as holdings). Para vender como PF seria necessária **redução de capital das holdings com entrega das ações da Embracon a valor contábil (art. 22, Lei 9.249/95)** — passo omitido na planilha, que é a origem da premissa C38. Risco de requalificação pelo Fisco se feito às vésperas da venda (jurisprudência CARF majoritariamente favorável quando há propósito negocial, mas é o principal risco fiscal da estrutura, não quantificado).
3. **Usufruto de Juarez:** tratar Juarez como "único beneficiário" de 100% do preço da JVFJ é juridicamente discutível — o usufruto dá direito aos frutos, não necessariamente ao produto da alienação da nua-propriedade; a alocação 100%/0% aumenta o IRPF em ~R$ 3,4 MM vs. a venda pelos 4 (perda de 3 conjuntos de faixas progressivas), e exigiria anuência/estruturação específica (extinção do usufruto, sub-rogação). Premissa a validar juridicamente.
4. **Participação de 0,01% direta** (organograma) ignorada no cálculo (holdings tratadas como 50%/50%).
5. **Arredondamentos JVFJ:** ON soma 99,99% e PN 100,01% (D35/E35) — inofensivo, mas denota tabela montada com percentuais truncados.
6. **34% cheio nas holdings:** desconsidera a franquia do adicional de IRPJ e eventuais atributos fiscais das holdings; conservador (a favor do argumento da economia).
7. **Soma de prejuízo fiscal IRPJ + CSLL** (Embracon K5 = 64,7 MM; Dourada J5 = 1,9 MM) é tecnicamente imprópria — são bases, não créditos aditivos; o valor econômico é ~34% de 32,35 MM (com trava de 30%).
8. **O que a planilha NÃO cobre:** ela é **anterior à estruturação final por incorporação de acervo líquido** — simula apenas a **venda secundária** de 20% por caixa. Não trata: (i) da incorporação do acervo da CNP Consórcio pela Embracon (relação de troca, CNP ~40% / holdings ~60%); (ii) da discussão contábil **CPC 15 (combinação de negócios/aquisição reversa, com ágio/mais-valia e PPA) vs. CPC 19/instrumento de capital** e seus efeitos fiscais (amortização de ágio, arts. 20-22 Lei 12.973/14); (iii) tributação do lado CNP; (iv) PIS/COFINS, ITBI/ITCMD, planejamento sucessório; (v) IR diferido sobre o estoque CPC 47 de ~R$ 1,2 bi de diferenças temporárias na Embracon (relevantíssimo para o PPA e para o preço); (vi) laudos de avaliação. Os dados da CNP são meramente públicos (balanço 2023 do grupo segurador).
9. **Nomenclaturas:** a aba chama a Embracon de "Ltda" (contexto societário atual fala em S.A.) e "ADMINISTRATORA" (sic); GSJ aparece como "Companhia GSJ de Part." — conferir denominações atuais nos atos societários.
10. **Data-base:** balanços de 12/2024; para o fechamento, custo/PL e prejuízos deverão ser atualizados (a economia de R$ 174 MM é sensível ao custo, mas pouco — o custo é ~2,4% do preço).

---

## 9. Tabela-resumo dos números-chave

| Item | Valor (R$) | Fonte |
|---|---:|---|
| Preço parcela secundária (20% Embracon) | 1.500.000.000,00 | Cálculo GK!D4 |
| Earn-out (assumido; faixa 300–800 MM a confirmar) | 0,00 | D5/E5 |
| Valuation implícito 100% Embracon | 7.500.000.000,00 | D4/D8 |
| Alocação por holding (Savian / JVFJ) | 750.000.000,00 cada | D11/D12 |
| PL Embracon 12/2024 | 178.389.573,07 | D14 |
| Valor contábil por holding (50% PL) | 89.194.786,54 | D16/D17 |
| Custo proporcional aos 20% vendidos (por holding) | 17.838.957,31 | D46/D54 |
| Ganho de capital por holding | 732.161.042,69 | D47/D55 |
| Tributo Cenário 1 por holding (34%) | 248.934.754,52 | D48/D56 |
| **Tributo total Cenário 1 (Holdings)** | **497.869.509,03** | D58 |
| IRPF Guido (GK 658.944.938,42) | 147.137.611,15 | G66 |
| IRPF por filho Savian (GK 14.643.220,85 cada, x5) | 2.553.644,17 | G67:G71 |
| Subtotal família Savian | 159.905.832,00 | G72 |
| IRPF Juarez (GK 732.161.042,69) | 163.611.234,61 | G77 |
| **Tributo total Cenário 2 (Pessoas Físicas)** | **323.517.066,61** | D83 |
| **Economia fiscal (Cenário 1 − Cenário 2)** | **174.352.442,43** | D89 |
| 50% da economia | 87.176.221,21 | D90 |
| **Ajuste de preço com gross-up 34%** | **132.085.183,66** | D91 |
| Carga efetiva Cenário 1 / Cenário 2 sobre o GK | 34,00% / ~22,09% | derivado |
| Prejuízo fiscal Embracon (IRPJ = CSLL) | 32.350.933,88 cada | Embracon!K3:K4 |
| Diferenças temporárias LALUR Embracon (soma) | 1.216.164.312,35 | Embracon!K13 |
| — Receita diferida CPC 47 | 533.696.093,06 | K7 |
| — Comissões diferidas CPC 47 | 666.446.778,73 | K8 |
| Provisões não dedutíveis Embracon | 18.872.556,36 | K18 |
| PL Savian / GSJ / Embrafisa / Franchising / Dourada | 89.341.616,47 / 108.460.364,76 / 100.909.570,65 / 1.149.700,60 / 10.053.653,40 | abas respectivas |
| Prejuízo fiscal Dourada (IRPJ = CSLL) | 956.469,61 cada | Dourada!J3:J4 |
| Lucro líquido Embrafisa 2024 | 52.820.770,80 | image5 (DRE) |
| PL CNP Seguros 31/12/2023 (controladora, R$ mil) | 5.156.127 | image6 |

---

## Dados para dashboard

```
projeto: Projeto Another House
planilha: Embracon_simulação tributária 21MAI2025.xlsx (consultor Molina)
data_planilha: 2025-05-21
data_base_balancos: 2024-12

preco_secundaria: 1500000000
earn_out_assumido: 0
earn_out_faixa_min: 300000000
earn_out_faixa_max: 800000000
percentual_adquirido: 0.20
valuation_implicito_100pct: 7500000000
alocacao_savian: 750000000
alocacao_jvfj: 750000000

pl_embracon_dez2024: 178389573.07
valor_contabil_por_holding: 89194786.535
custo_20pct_por_holding: 17838957.307
gk_por_holding: 732161042.693

cenario1_tributo_savian: 248934754.52
cenario1_tributo_jvfj: 248934754.52
cenario1_total: 497869509.03
cenario1_aliquota: 0.34

cenario2_irpf_guido: 147137611.15
cenario2_irpf_filho_savian_cada: 2553644.17
cenario2_subtotal_savian: 159905832.00
cenario2_irpf_juarez: 163611234.61
cenario2_total: 323517066.61
cenario2_aliquota_efetiva: 0.2209
faixas_irpf_gk: 15% até 5MM | 17,5% 5–10MM | 20% 10–30MM | 22,5% >30MM (Lei 13.259/2016)

economia_fiscal: 174352442.43
economia_50pct: 87176221.21
ajuste_preco_grossup_34pct: 132085183.66

embracon_prejuizo_fiscal_irpj: 32350933.88
embracon_prejuizo_fiscal_csll: 32350933.88
embracon_dif_temporarias_lalur_total: 1216164312.35
embracon_cpc47_receita_diferida: 533696093.06
embracon_cpc47_comissoes_diferidas: 666446778.73
embracon_provisoes_indedutiveis: 18872556.36
embracon_regime: Lucro Real

pl_savian: 89341616.47
pl_gsj: 108460364.76
pl_embrafisa: 100909570.65
pl_franchising: 1149700.60
pl_dourada: 10053653.40
dourada_prejuizo_fiscal_irpj: 956469.605
dourada_prejuizo_fiscal_csll: 956469.605
embrafisa_lucro_liquido_2024: 52820770.80
regimes: Savian/GSJ/Embrafisa/Franchising = Lucro Presumido; Embracon/Dourada = Lucro Real

cnp_seguros_pl_controladora_2023_rmil: 5156127
cnp_seguros_ativo_consolidado_2023_rmil: 10437415
cnp_fonte: informação pública (internet), balanço 31/12/2023

participacao_embracon: Savian 49,99% + JVFJ 49,99% + 0,01% direto (PF)
composicao_savian: Guido 90% (100% ON); Daniel/Gabriel/Lucas/André/Matheus 2% cada (20% PN cada)
composicao_jvfj: Juarez 6,25% (83,33% PN); Dutra/Vanessa/Fernanda 31,25% cada — Juarez usufrutuário de 100% das ações
escopo: planilha cobre APENAS a venda secundária; não cobre incorporação de acervo CNP, CPC 15 x CPC 19, ágio/PPA, IR diferido CPC 47, lado CNP
```
