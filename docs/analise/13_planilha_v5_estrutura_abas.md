# 13 — Planilha-entregável v5 (04.09.2026): estrutura de abas e vínculos

> Base: v4 editada pelo usuário (`planilhas/Embracon_simulacao_tributaria_04.09.26_v4 (fonte usuario).xlsx`).
> Entregável: `planilhas/Embracon_simulacao_tributaria_04.09.26_v5.xlsx`. Gerador: `planilhas/build_v5.py`
> (+ `inject_values_v5.py`) — cirurgia OOXML: preserva logo EMF, abas ocultas, sharedStrings e formatação;
> altera só `workbook.xml`, rels, `[Content_Types]`, `app.xml`, `styles.xml` (acréscimos), `sheet2.xml` (links de
> premissas + remoção do título órfão B96), `sheet3.xml` (reconstruída) e cria `sheet8/9/10.xml` + drawings 5-7.

## Ordem das abas e função

| # | Aba | Conteúdo | Fonte dos números |
|---|---|---|---|
| 1 | Organograma | imagens (inalterada; sem logo — as imagens ocupam a área) | — |
| 2 | **Premissas** (nova, aba inicial) | fatos da operação; **inputs** (cinza; custo PF em amarelo); valores derivados; premissas fiscais e ressalvas do PDF SF 31.08 | inputs |
| 3 | Cálculos da Operação | inalterada em valores; as células de premissa (C13, C14, C16, C17, C19, C23, C24, I13:I17, D68) passaram a **=Premissas!…** | Premissas |
| 4 | **CPC 19** (nova) | 1) VJ da JO (100%); 2) investimento das Holdings (60%), base contábil, AVJ, carimbo 34%; 3) lançamentos contábeis (partidas dobradas, conferência); 4) custo após cash-out CPC 15 × CPC 19; leitura | Premissas + Cálculos!J93 |
| 5 | Alienação Futura (PJ x PF) | 1) premissas da venda (links); 2) venda PJ×PF (tabela do usuário, Δ em vermelho) + fórmulas + leitura; cenário alternativo (perda); análise hipotética RISCO 1/2; 3) conclusão CPC 15 × CPC 19 (tabela do usuário) | Premissas, Cálculos (J93, I93, K93), CPC 19 (D22:D26) |
| 6 | **Conclusão** (nova) | quadros do PPT/MLA: custo do investimento (CPC 15 em camadas PPA/Antes/Depois/Δ/IR/Risco com marcadores coloridos; CPC 19 100%/60%/IR/Risco) e venda futura lado a lado; conclusão executiva + ressalvas | Premissas, Cálculos, CPC 19, Alienação |
| 7-10 | Dourada / Franchising / Embrafisa / GSJ | ocultas, inalteradas | — |

## Referências-chave (para manutenção)

- `Cálculos da Operação`: D71 fração alienada (30,39%); I93 custo remanescente cenário A (176,8); J93 cenário B (645,6); K93 Δ (468,8).
- `Premissas`: C19 PL Embracon; C20/C21 investimentos; C22 PL CNP; C23 VM CNP; C24 VJ Embracon; C25 13,8%; C26 40%; C27 cash-out; C28 34%; C29 22,5%; C30 custo PF; C31/C32/C33 preços (C32 = 'CPC 19'!D22); C36 famílias 60%; C37 mais-valia; C38 26,2%; C39 fração.
- `CPC 19`: C14 VJ 100% (5.955); C15 base (1.054); C16 AVJ (4.901); D22 VJ 60% (3.573); D23 base 60% (632,4); D24 AVJ 60% (2.940,6); D25 carimbo (999,8); D26 custo CPC 19.
- `Alienação`: C13 custo CPC 15 (J93); C14 cenário A (I93); C15 mais-valia 60%; C16:C19 CPC 19; C20:C22 preços; tabela venda D30:I32; RISCO D50:H53; conclusão C59:E66.
- `Conclusão`: E13 = residual (J93 − I93 − mais-valia×60%) com conferência C16 = "ok"; venda B28:H30 vinculada à Alienação.

## Conformidade visual (item v)

Logo: mesma âncora da aba Cálculos (coluna B + 15240 EMU, linha 3 + 145908 EMU, 2828149 × 648190) em Premissas, CPC 19,
Alienação (reancorada) e Conclusão. Título: linha 4, itálico 20, coluna C (abas largas) ≈ mesma posição horizontal de
`Cálculos`!E4; data na linha 5 (altura 26 para o logo não encostar na barra vermelha).

## O que mudou de valor/estrutura (informado ao usuário)

- Nenhum valor/resultado da aba Cálculos mudou (verificado célula a célula contra a v4); mudou a origem (links).
- Título órfão "CPC 19 — Joint Operation" na linha 96 da aba Cálculos removido (conteúdo foi para a aba CPC 19).
- Alienação: blocos "1) Custo do Investimento…" (CPC 15 em camadas e CPC 19) saíram da aba (camadas → Conclusão; CPC 19 → aba própria); seções renumeradas 1/2/3; título com acento ("Análise"); 54 células da venda/conclusão conferidas idênticas à v4.
- Correção de bug no build: preço "igual ao valor justo" apontava para 'CPC 19'!D21 (493,2); corrigido para D22 (3.573).
