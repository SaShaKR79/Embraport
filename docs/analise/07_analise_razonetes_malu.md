# Análise dos razonetes modelo (Malu) — CNP Group / Embracon Holdings / Joint Operation

**Fonte:** 3 capturas de tela do arquivo Excel `CNP - Draft razonetes.xlsx` (marcado "Confidential", aba ativa **Razonetes**; demais abas: **SPA** e **Joint Operation**; faixa de opções mostra o suplemento "Omnia 4.0 EMA", sugerindo ambiente de auditoria).

**Contexto geral inferido da operação:**
- A **Embracon Holdings** detém 100% de um investimento (custo histórico **178**) cujo valor justo (AVJ) de 100% é **5.888** ("Valuation – carteira de clientes").
- O **CNP Group** adquire **40%** desse negócio em duas tranches: **14,7% por troca de ações (FMV)** e **25,3% em espécie (caixa de 1.600)**. A Embracon **retém 60%**, e o negócio passa a ser tratado como **Joint Operation (operação em conjunto — CPC 19/IFRS 11)** — daí a conta "Joint Op." nos dois lados.
- Há um **earn-out** cujo valor ainda não foi definido (marcado como "**x**" em laranja nos razonetes): a pagar no CNP, a receber na Embracon.
- Todas as consistências aritméticas fecham com arredondamento em milhares: 40% × 5.888 = 2.355; 14,7% × 5.888 = 866; 25,3% × 5.888 = 1.490; 60% × 5.888 = 3.533; 40% × 178 = 71; 60% × 178 = 107; 25,3% × 178 = 45; 14,7% × 178 ≈ 26.

---

## Imagem 1 — `razonentes modelo (Malu).png` (visão geral da aba "Razonetes")

### 1. Título/contexto
Tela completa do Excel com **dois blocos de razonetes**: "**CNP Group**" (comprador, linhas 2–11) e "**Embracon Holdings**" (vendedor, linhas 13–27), cada um dividido em colunas "Ativo", "Passivo e PL" e "Resultado". À direita (colunas T–Z), um **quadro-memória** com a mensuração das tranches (Custo × AVJ) e observações fiscais.

### 2 e 3. Contas, valores e lançamentos

#### Bloco CNP Group (comprador)
| Conta | Débitos | Créditos | Saldo final |
|---|---|---|---|
| **Cash** (Ativo) | — | (1.600) | (1.600) |
| **Invest.** (Ativo) | 300 | (300) | – (zera) |
| **Joint Op.** (Ativo) | 866; 1.490 | — | **2.355** (destacado em verde) |
| **PL** (Passivo e PL) | — | (300) | (300) |
| **Earn-out a pagar** (Passivo) | — | **(x)** (célula laranja) | – |
| **Resultado (ganho/perda)** | 110; **x** (laranja) | (566) (destaque azul) | soma 110 / (566); líquido **(455)** |

Lançamentos inferidos (D/C):
1. **Saldo inicial/integralização do investimento entregue:** D – Invest. 300 / C – PL 300.
2. **Aquisição de 14,7% por troca de ações (FMV):** D – Joint Op. 866 / C – Invest. (300) / C – Resultado (566) → **ganho de 566** na entrega do ativo (custo 300, valor justo 866).
3. **Aquisição de 25,3% em espécie:** D – Joint Op. 1.490 / D – Resultado 110 (perda) / C – Cash (1.600) → pagou 1.600 por participação com AVJ 1.490; o excedente de **110 vai a resultado como perda** (não se reconhece ágio/goodwill — tratamento de aquisição de ativos/JO fora do CPC 15).
4. **Earn-out:** D – Resultado x / C – Earn-out a pagar (x) → provisão de valor ainda indefinido, **como despesa** (não como contraprestação contingente).

Resultado líquido do CNP: **ganho de ~455** (566 − 110, com arredondamento).

#### Bloco Embracon Holdings (vendedor)
| Conta | Débitos | Créditos | Saldo final |
|---|---|---|---|
| **Cash** (Ativo) | 1.600 | — | 1.600 |
| **Invest.** (Ativo) | 178 | (178) | – (zera) |
| **Joint Op.** (Ativo) | 3.533 | — | **3.533** (verde) |
| **Earn-out a receber** (Ativo) | **x** (laranja) | — | – |
| **PL** | — | (178) | (178) |
| **Resultado (ganho/perda)** | 71 | (1.600); (3.426) (azul); **(x)** (laranja) | soma 71 / (5.026); líquido **(4.955)** |

Nas linhas 26–27 aparece o início do **quadro de abertura/prova** (detalhado na Imagem 3): Cash 1.600; Invest. (71); Joint Op. 755 / 107; Resultado (2.355) / 71.

### 4. Anotações e destaques
- Quadro à direita (CNP): "a) Aquisição 14,7% – troca ações (FMV): Custo 300 | AVJ 866 | 566"; "b) Aquisição 25,3% – espécie: 1.600 | 1.490 | (110)"; "c) Reavaliação Ativos (AVJ): **455**" (destaque azul); "d) Diferido Passivo: (155)".
- Nota manuscrita/digitada do CNP: "**Earn-out: Permanente – não cpc 15**" e "**AVP (subconta e neutraliza) + juros (indedutível)**".
- Quadro à direita (Embracon): "a) Venda 14,7% – troca ações (FMV)"; "b) Venda 25,3% – espécie: 45 | 1.600 | 1.555"; "c) Ganho líquido venda ações: **1.555**" (destaque rosa); "d) Reavaliação Ativos (AVJ) – 60%: 100% → 178 | 5.888 (verde); 60% → 107 | 3.533 | **3.426** (azul)"; "e) Diferido Passivo: (1.165)".
- Nota da Embracon: "**Earn-out: tributar como 'complemento preço' – GK**" e "**AVP (subconta e neutraliza) + juros (tributar – complemento preço)**". ("GK" aparenta ser referência a assessor/memorando tributário.)
- Células "x"/"(x)" em laranja = valor do earn-out ainda a definir (placeholder).

### 5. Interpretação contábil
A Malu monta, **lado a lado, a contabilização espelhada do comprador e do vendedor** na formação da joint operation:
- **CNP (comprador):** reconhece sua parcela de 40% da JO **a valor justo (2.355)**; a diferença entre o custo dos ativos entregues (300 + 1.600 = 1.900) e o AVJ recebido vai a **resultado** (ganho 566 na troca de ações; perda 110 no sobrepreço em caixa). **Não há goodwill** — coerente com aquisição de participação em operação conjunta tratada fora do CPC 15 (aquisição de ativos), com a mais-valia alocada à **carteira de clientes** (455) e respectivo **IR diferido passivo de 155** (34% × 455).
- **Embracon (vendedor):** baixa o investimento (178), reconhece o caixa recebido (1.600) e **remensura a valor justo os 60% retidos** na JO (3.533), com ganho total de **4.955** no resultado (ganho na venda + ganho de remensuração), além do earn-out a receber (x). IR diferido passivo de **1.165** sobre a diferença entre base contábil (3.533) e fiscal (107).

---

## Imagem 2 — `razonentes modelo (Malu) 2.png` (zoom no quadro-memória de mensuração e efeitos fiscais)

### 1. Título/contexto
Ampliação do quadro lateral (colunas T–Z da aba Razonetes), com os **cálculos de suporte** dos razonetes: colunas **Custo | AVJ | (diferença) | BC | BF | DTA/(L)**.

### 2 e 3. Conteúdo linha a linha

**Bloco superior (CNP Group – comprador):**
| Linha | Custo | AVJ | Dif. | BC | BF | Dif. BC−BF | IR diferido |
|---|---|---|---|---|---|---|---|
| a) Aquisição 14,7% – troca ações (FMV) | 300 | **866** (célula selecionada) | 566 | **2.355** | **1.900** | **(455)** | **(155)** |
| b) Aquisição 25,3% – espécie | 1.600 | 1.490 | (110) | | | | |
| c) Reavaliação Ativos (AVJ) | | | **455** (azul) — anotação: "**Valuation – carteira clientes**" (com ícone de comentário) | | | | |
| d) Diferido Passivo | | | (155) | | | | |

Notas: "Earn-out: Permanente – não cpc 15" / "AVP (subconta e neutraliza) + juros (indedutível)".

**Bloco inferior (Embracon Holdings – vendedor):**
| Linha | Custo | AVJ/FMV | Ganho | BC | BF | Dif. | IR diferido |
|---|---|---|---|---|---|---|---|
| a) Venda 14,7% – troca ações (FMV) | (sem valores exibidos) | | | | | | |
| b) Venda 25,3% – espécie | 45 | 1.600 | 1.555 | | | | |
| c) Ganho líquido venda ações | | | **1.555** (rosa) | | | | |
| d) Reavaliação Ativos (AVJ) – 60%: linha 100% | 178 | **5.888** (verde) | | | | | |
| d) linha 60% | 107 | 3.533 | **3.426** (azul) | **3.533** | **107** | **(3.426)** | **(1.165)** |
| e) Diferido Passivo | | | (1.165) | | | | |

Notas: "Earn-out: tributar como 'complemento preço' – GK" / "AVP (subconta e neutraliza) + juros (tributar – complemento preço)".

### 4. Anotações/destaques
- Célula AVJ 866 selecionada (borda verde) e 1.490 sombreada — indicam edição/ponto de atenção.
- Destaques de cor: 455 e 3.426 em azul (mais-valia/remensuração), 1.555 em rosa (ganho de venda), 5.888 em verde (valuation de 100%).
- Ícone de comentário ao lado do 455 com o texto "Valuation – carteira clientes".

### 5. Interpretação contábil
É o **PPA simplificado + memória de tributos diferidos**:
- **Comprador:** base contábil (BC) da participação na JO = 2.355 (AVJ); base fiscal (BF) = 1.900 (custo dos ativos entregues: 300 + 1.600). Diferença temporária tributável de **455** — exatamente a **mais-valia alocada à carteira de clientes** — gera **passivo fiscal diferido de 155** (455 × 34%).
- **Vendedor:** BC dos 60% retidos = 3.533 (AVJ); BF = 107 (custo remanescente). Diferença de **3.426** gera **passivo fiscal diferido de 1.165** (3.426 × 34%).
- O ganho contábil do vendedor na tranche em espécie (1.555 = 1.600 − 45) é destacado como **ganho líquido tributável na venda das ações**; o earn-out, quando realizado, será tributado como **complemento de preço** (ganho de capital), e o **AVP do earn-out fica em subconta (Lei 12.973/14), neutralizado fiscalmente**, com os juros do AVP **indedutíveis no comprador** e **tributáveis como complemento de preço no vendedor**.

---

## Imagem 3 — `razonentes modelo (Malu) 3.png` (zoom no bloco Embracon Holdings + quadro de prova)

### 1. Título/contexto
Ampliação do bloco "**Embracon Holdings**" da aba Razonetes: razonetes do vendedor, quadro lateral (idêntico ao bloco inferior da Imagem 2) e, embaixo, um **quadro de decomposição/prova** dos saldos de Joint Op. e Resultado.

### 2 e 3. Contas, valores e lançamentos

**Razonetes (mesmos da Imagem 1, agora legíveis por completo):**
- **Cash:** D 1.600 | totais 1.600 / – | saldo **1.600**.
- **Invest.:** D 178, C (178) | saldo **–**.
- **Joint Op.:** D 3.533 | saldo **3.533** (verde).
- **Earn-out a receber:** D **x** (laranja) | saldo –.
- **PL:** C (178) | saldo **(178)**.
- **Resultado (ganho/perda):** D 71; C (1.600), (3.426) (azul), (x) (laranja) | totais 71 / (5.026) | saldo líquido **(4.955)**.

Conferência de partidas dobradas: débitos 1.600 + 3.533 + 71 = 5.204 = créditos 178 + 1.600 + 3.426 (fora PL inicial e earn-out "x").

**Quadro de decomposição (parte inferior):**
| Cash | Invest. | Joint Op. | Resultado | |
|---|---|---|---|---|
| 1.600 | | 755 | (2.355) | – |
| | (71) | | 71 | – |
| | (107) | 107 | – | – |
| | – | 2.671 | (2.671) | – |
| | | **3.533** | **71 / (5.026)** | |

Leitura das linhas:
1. **Venda dos 40% a valor justo total de 2.355**: recebe caixa 1.600 (25,3%) + participação na JO de **755** (contrapartida da troca de ações dos 14,7%; 755 = 2.355 − 1.600) contra ganho bruto (2.355).
2. **Baixa do custo dos 40% vendidos**: C – Invest. (71) / D – Resultado 71 (71 = 40% × 178).
3. **Reclassificação dos 60% retidos**: C – Invest. (107) / D – Joint Op. 107 (ao custo).
4. **Remensuração dos 60% retidos a valor justo**: D – Joint Op. 2.671 / C – Resultado (2.671), levando a JO a 3.533 (= 107 + 755 + 2.671).

Prova econômica do ganho total de 4.955: ganho na venda de 25,3% (1.600 − 45 = **1.555**) + ganho na troca de 14,7% (755 − 26 ≈ **729**) + ganho de remensuração dos 60% (**2.671**) = **4.955**. Nota: no razonete, o crédito (3.426) agrega 755 + 2.671 (e coincide, por identidade aritmética, com a "Reavaliação Ativos 60%" do quadro: 3.533 − 107 = 3.426).

### 4. Anotações/destaques
- Quadro lateral repetido: a) Venda 14,7% – troca ações (FMV); b) Venda 25,3% – espécie 45 | 1.600 | 1.555; c) Ganho líquido venda ações 1.555 (rosa); d) Reavaliação AVJ 60% (100%: 178 | 5.888 verde; 60%: 107 | 3.533 | 3.426 azul); e) Diferido Passivo (1.165).
- Notas: "Earn-out: tributar como 'complemento preço' – GK" e "AVP (subconta e neutraliza) + juros (tributar – complemento preço)".
- Cores: verde = saldos finais de Joint Op./valuation; azul = remensuração AVJ; laranja = earn-out pendente ("x"); rosa = ganho de venda.

### 5. Interpretação contábil
É a demonstração completa da **perda de controle com retenção de participação** (lógica do CPC 36/IFRS 10 aplicada à formação da joint operation): o vendedor (i) baixa o investimento pelo custo, (ii) reconhece a contraprestação recebida a valor justo (caixa + participação recebida na troca + earn-out), e (iii) **remensura a valor justo a participação retida de 60%**, levando todo o efeito (4.955 + x) ao **resultado**, e não ao PL/reserva. O IR diferido passivo de 1.165 sobre o step-up dos 60% é calculado no quadro lateral (ainda não lançado nos razonetes, que estão "pré-tributos diferidos").

---

## Síntese técnica (CPC 15 / CPC 19)
1. O modelo trata a operação como **formação de joint operation (CPC 19)**, não como combinação de negócios com goodwill: o comprador reconhece sua parcela dos ativos **a valor justo**, expensa o sobrepreço (110) e **não registra ágio**; a mais-valia identificada (455) é alocada a **intangível de carteira de clientes**, com DTL de 155.
2. O **earn-out é explicitamente excluído do CPC 15** ("Permanente – não cpc 15"): no comprador é despesa permanente (indedutível, inclusive juros do AVP); no vendedor é receita a tributar como **complemento de preço**, com **AVP controlado em subconta e neutralizado** (Lei 12.973/14).
3. O vendedor aplica a mecânica de **remensuração da participação retida** com ganho integral em **resultado** (não em reserva), e o quadro BC × BF documenta as **diferenças temporárias** e os **tributos diferidos (34%)** de cada lado.
4. Valores-chave: valuation 100% = **5.888**; contraprestação 40% = **2.355** (866 troca de ações + 1.490 caixa, pago 1.600); custo histórico 100% = **178**; JO no comprador = **2.355**; JO no vendedor (60%) = **3.533**; resultado comprador ≈ **+455**; resultado vendedor = **+4.955** (+ earn-out "x"); DTL comprador **155**; DTL vendedor **1.165**.
