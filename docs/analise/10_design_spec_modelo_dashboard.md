# Especificação de Design — Dashboard modelo "Grupo Primo — Fluxos Intercompany" (17AGO2026)

> Documento de referência para replicar FIELMENTE o visual/estilo do dashboard modelo em um novo
> dashboard sobre outro tema. Todos os valores (hex, px, famílias, raios, sombras) foram copiados
> literalmente do arquivo original — nada foi inventado.
>
> Arquivo-fonte analisado:
> `/tmp/claude-0/-home-user-Embraport/59f689c0-4eb3-5e38-9571-3e4f90356cb1/scratchpad/uploads/44b2257b-Grupo_Primo__Fluxos_Intercompany_e_Radiografia_Financeira__17AGO2026.html`
> (820 linhas, ~117KB, arquivo único autocontido — funciona offline).

---

## 1. Estrutura geral da página

### 1.1 Arquitetura do arquivo

É **um único arquivo HTML autocontido**, sem nenhuma dependência externa (sem CDN, sem Google
Fonts, sem biblioteca de gráficos). A estrutura é:

```
<!DOCTYPE html>
<html lang="pt-BR" data-theme="light">
<head>
  <meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
  <title>…</title>
  <style> …TODO o CSS inline (~210 linhas)… </style>
</head>
<body>
  <header>  ← logo + título + botão de tema + <nav> de abas
  <div class="wrap">  ← container central
    <section id="t-XXX"> vazias (uma por aba; todas menos a 1ª com class="hidden")
    <p class="foot" id="foot">  ← rodapé de fontes/ressalvas
  </div>
  <script id="p" type="application/json"> …TODOS os dados em JSON… </script>
  <script> …renderização por template literals + navegação de abas + toggle de tema… </script>
</body></html>
```

**Padrão-chave de arquitetura**: separação dados/apresentação dentro do próprio arquivo.
Todo o conteúdo (números, textos, listas) vive num bloco `<script id="p" type="application/json">`;
o JS lê esse JSON (`const D=JSON.parse(document.getElementById('p').textContent);`) e cada aba tem
uma função `renderXxx()` que monta o HTML com template literals e injeta via `innerHTML` na
`<section>` correspondente. Para atualizar valores basta editar o bloco JSON.

### 1.2 Header

```html
<header>
 <div class="hd">
  <img class="logo" src="data:image/png;base64,…">   <!-- logo embutido em base64 -->
  <div class="hdiv"></div>                            <!-- divisor vertical de 1px -->
  <div>
    <p class="br">GRUPO PRIMO</p>                     <!-- "eyebrow" da marca, caps, cor brand -->
    <h1>Fluxos intercompany e radiografia financeira</h1>
    <p class="sb">Posição 17.08.2026 · balancetes de 30.06.2026 · …</p>  <!-- subtítulo com metadados separados por " · " -->
  </div>
  <div class="rt"><button class="tg" id="tg" type="button">Modo escuro</button></div>  <!-- toggle de tema à direita -->
 </div>
 <nav id="nav" role="tablist"></nav>                  <!-- abas geradas por JS -->
</header>
```

CSS do header (literal):

```css
header{background:var(--surface);color:var(--ink);border-bottom:1px solid var(--edge)}
.hd{max-width:1240px;margin:0 auto;padding:22px 26px 0;display:flex;gap:20px;align-items:center;flex-wrap:wrap}
.logo{height:38px;width:auto;display:block;flex:0 0 auto}
:root[data-theme=dark] .logo{filter:brightness(1.12)}
.hdiv{width:1px;align-self:stretch;background:var(--line);margin:3px 0}
.hd .br{font-size:10px;letter-spacing:.18em;color:var(--brand);font-weight:700;margin:0 0 3px}
.hd h1{margin:0;font-size:20px;font-weight:700;letter-spacing:-.015em;color:var(--ink)}
.hd .sb{margin:4px 0 0;font-size:12.5px;color:var(--mut)}
.hd .rt{margin-left:auto}
.tg{background:transparent;color:var(--ink2);border:1px solid var(--edge);border-radius:7px;
 padding:6px 11px;font:inherit;font-size:12px;cursor:pointer}
.tg:hover{background:var(--plane);color:var(--ink)}
```

### 1.3 Navegação por abas (tabs)

- As abas ficam num `<nav id="nav" role="tablist">` dentro do header (abaixo do título).
- Botões sem fundo, com **sublinhado inferior de 2.5px na cor da marca** quando ativos
  (padrão "underline tabs", não "pill tabs").

```css
nav{max-width:1240px;margin:0 auto;padding:16px 26px 0;display:flex;gap:4px;flex-wrap:wrap}
nav button{background:transparent;border:0;border-bottom:2.5px solid transparent;color:var(--mut);
 padding:10px 16px;font:inherit;font-size:13.5px;font-weight:600;cursor:pointer}
nav button:hover{color:var(--ink)}
nav button[aria-selected=true]{color:var(--ink);border-bottom-color:var(--brand)}
```

Mecânica JS (literal, é o coração da navegação — render **lazy**: cada aba só é renderizada na
primeira vez que é aberta, com cache em `done{}`):

```js
const T=[['org','1. Organograma'],['emp','2. Radiografia 1S2026'],['op','3. Oportunidades (v3)'],
 ['fin','4. Efeito financeiro'],['trib','5. Panorama tributário'],['ctr','6. Contratos'],['src','7. Fontes e pendências']];
const R={org:renderOrg,emp:renderEmp,op:renderOp,fin:renderFin,trib:renderTrib,ctr:renderCtr,src:renderSrc}, done={};
function show(k){T.forEach(([id])=>{document.getElementById('t-'+id).classList.toggle('hidden',id!==k);
 document.getElementById('b-'+id).setAttribute('aria-selected',id===k);});
 if(!done[k]){R[k]();done[k]=1;} window.scrollTo({top:0,behavior:'instant'});}
document.getElementById('nav').innerHTML=T.map(([id,l])=>
 `<button id="b-${id}" role="tab" aria-selected="false" type="button">${l}</button>`).join('');
T.forEach(([id])=>document.getElementById('b-'+id).onclick=()=>show(id));
show('org');   // aba inicial
```

Convenções de id: seção `id="t-<chave>"`, botão `id="b-<chave>"`. Ocultação por classe utilitária
`.hidden{display:none!important}`. Cada troca de aba faz scroll para o topo (instant).

Os títulos das abas são **numerados** ("1. Organograma", "2. Radiografia 1S2026", …) e o `<h2>`
dentro de cada seção repete o mesmo número/título.

### 1.4 Corpo e rodapé

```css
.wrap{max-width:1240px;margin:0 auto;padding:0 26px 90px}
section{padding-top:36px}
.foot{margin-top:44px;padding-top:18px;border-top:1px solid var(--line);font-size:12px;color:var(--mut)}
.foot b{color:var(--ink2)}
```

O rodapé (`p.foot#foot`) é preenchido por JS e é **compartilhado entre todas as abas** (fica fora
das sections): contém "**Fontes.** …" (fonte de cada aba) e "**Ressalvas.** …" (disclaimers),
separados por `<br><br>`.

### 1.5 Utilitários JS de formatação (padrão do modelo)

```js
const E=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));  // escape HTML
const F0=n=>n==null?'—':n.toLocaleString('pt-BR',{maximumFractionDigits:0});                                  // inteiro pt-BR
const FM=n=>n==null?'—':(Math.abs(n)>=1e6?(n/1e6).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+' mi'
                                          :(n/1e3).toLocaleString('pt-BR',{maximumFractionDigits:0})+' mil'); // "12,3 mi" / "800 mil"
```

Valores ausentes são sempre exibidos como travessão `—`. Negativos usam o sinal `−` (U+2212), não hífen.

---

## 2. Paleta de cores completa

Todas as cores são **variáveis CSS em `:root`**, com um segundo bloco `:root[data-theme="dark"]`
que as sobrescreve. Copiado literalmente:

```css
:root{--surface:#fcfcfb;--plane:#f6f6f3;--ink:#0b0b0b;--ink2:#54534f;--mut:#8d8b85;
 --line:#e4e3dc;--edge:#c9c8bf;--ring:rgba(11,11,11,.09);
 --real:#2a78d6;--pres:#eb6834;--nd:#8d8b85;
 --ok:#0d8a2f;--aval:#c98500;--rever:#c0392b;--esp:#8d8b85;
 --brand:#e10913;--shell:#211f1e}
:root[data-theme="dark"]{--surface:#1a1a19;--plane:#111110;--ink:#fff;--ink2:#c3c2b7;--mut:#8d8b85;
 --line:#2c2c2a;--edge:#3a3a37;--ring:rgba(255,255,255,.11);
 --real:#3987e5;--pres:#e0703f;--ok:#2faa4c;--aval:#d19a1e;--rever:#e07168;--shell:#000}
```

### Papel de cada variável

| Variável | Light | Dark | Papel |
|---|---|---|---|
| `--surface` | `#fcfcfb` | `#1a1a19` | Fundo de **cartões/superfícies elevadas** (cards, tabelas, notas, header) |
| `--plane` | `#f6f6f3` | `#111110` | **Fundo da página** (body) e fundo de sub-áreas dentro de cards (headers de card, trilhas de barra, hover de linha, th de planilha) |
| `--ink` | `#0b0b0b` | `#fff` | Texto primário (títulos, valores, ênfase) |
| `--ink2` | `#54534f` | `#c3c2b7` | Texto secundário (parágrafos, células descritivas) |
| `--mut` | `#8d8b85` | `#8d8b85` (igual) | Texto atenuado (labels, subtítulos, cabeçalhos de tabela, metadados) |
| `--line` | `#e4e3dc` | `#2c2c2a` | Bordas internas leves (divisores de linhas de tabela, separadores) |
| `--edge` | `#c9c8bf` | `#3a3a37` | Bordas mais fortes (bordas de nós, cards de contrato, cabeçalhos de planilha, conectores) |
| `--ring` | `rgba(11,11,11,.09)` | `rgba(255,255,255,.11)` | Borda ultrassutil translúcida de cartões grandes (kpi, tw, note, org, bars, base) e de tags |
| `--brand` | `#e10913` | (não muda) | **Cor da marca** (vermelho): eyebrow do header, sublinhado da aba ativa, destaque do nó "apex" do organograma, numeração do checklist, negritos especiais |
| `--shell` | `#211f1e` | `#000` | Fundo escuro do "pill" de percentual nos nós do organograma |
| `--real` | `#2a78d6` | `#3987e5` | **Cor categórica 1** (azul) — no modelo: regime "Lucro Real"; também usada como cor das barras de receita e da nota `.note.info` |
| `--pres` | `#eb6834` | `#e0703f` | **Cor categórica 2** (laranja) — regime "Lucro Presumido" |
| `--nd` | `#8d8b85` | (igual) | Categoria "não definido / a levantar" (cinza) |
| `--ok` | `#0d8a2f` | `#2faa4c` | **Semântica positiva** (verde): valores positivos, status "Possível", saldo que economiza |
| `--aval` | `#c98500` | `#d19a1e` | **Semântica de atenção/alerta médio** (âmbar): status "Aprofundar", borda padrão da `.note` |
| `--rever` | `#c0392b` | `#e07168` | **Semântica negativa** (vermelho terroso): valores negativos, alertas graves, "deficitária" |
| `--esp` | `#8d8b85` | — | Reservada (cinza, igual a `--mut`; declarada mas pouco usada) |

Cores fixas fora de variáveis: `#fff` (texto do pill `.pp` sobre `--shell`), `#31312d`
(fundo do `.pp` no dark: `:root[data-theme=dark] .node .pp{background:#31312d}`).

**Caráter da paleta**: neutros **quentes/off-white** (tons de areia/greige: `#fcfcfb`, `#f6f6f3`,
`#e4e3dc`, `#c9c8bf`), nunca cinza puro. Acento de marca vermelho usado com muita parcimônia.
Cores semânticas dessaturadas/escurecidas (verde floresta, âmbar, tijolo) — nada neon.
No dark theme os acentos clareiam um pouco para manter contraste.

---

## 3. Tipografia

**Não há import de fonte externa** (sem Google Fonts). Tudo usa a pilha de sistema:

```css
body{margin:0;background:var(--plane);color:var(--ink);
 font-family:system-ui,-apple-system,"Segoe UI",sans-serif;font-size:14.5px;line-height:1.6}
```

- Base: `14.5px / 1.6`.
- **Não usa fonte monoespaçada**: números tabulares são obtidos com
  `font-variant-numeric:tabular-nums` (aplicado em toda célula/valor numérico). A classe `.mono`
  na verdade é sans com tabular-nums: `.mono{font-variant-numeric:tabular-nums;font-size:12px;color:var(--ink2)}`.
- Botões e inputs herdam a fonte com `font:inherit`.
- Pesos usados: 400 (corrido), 600, 650, 700, 750, 800 (o modelo usa pesos intermediários
  como 650/750 — em fontes variáveis do sistema funcionam; degradam bem para 600/700).
- Títulos com letter-spacing **negativo**; labels/eyebrows em CAPS com letter-spacing **positivo largo**.

### Hierarquia tipográfica (valores literais)

| Elemento | CSS |
|---|---|
| H1 (header) | `font-size:20px;font-weight:700;letter-spacing:-.015em` |
| Eyebrow marca `.br` | `font-size:10px;letter-spacing:.18em;color:var(--brand);font-weight:700` (texto em CAIXA ALTA) |
| Subtítulo header `.sb` | `font-size:12.5px;color:var(--mut)` |
| H2 (título da aba) | `font-size:21px;margin:0 0 8px;letter-spacing:-.02em;font-weight:700` |
| H3 de seção `.sec` | `font-size:16px;margin:38px 0 6px;letter-spacing:-.015em` |
| Lead `p.lead` | `margin:0 0 26px;color:var(--ink2);max-width:860px` (herda 14.5px) |
| Label de KPI `.kl` | `font-size:10px;text-transform:uppercase;letter-spacing:.08em;font-weight:800;color:var(--mut)` |
| Valor de KPI `.kv` | `font-size:22px;font-weight:750;letter-spacing:-.02em;font-variant-numeric:tabular-nums` |
| Sub-KPI `.ks` | `font-size:11px;color:var(--mut);line-height:1.4` |
| Cabeçalho de tabela `thead th` | `font-size:10.5px;text-transform:uppercase;letter-spacing:.06em;color:var(--mut);font-weight:700` |
| Corpo de tabela | `font-size:13px` (tabela simples) / `12.5px` (planilha) / `12px` (tabela de card) |
| Tag/badge | `font-size:11px;font-weight:650` |
| Nav | `font-size:13.5px;font-weight:600` |
| Rodapé | `font-size:12px;color:var(--mut)` |

**Padrão de eyebrow**: micro-label em caps (9–10.5px, letter-spacing .05–.18em, weight 700–800,
cor `--mut` ou `--brand`) acima de conteúdo — usado no header, KPIs, timeline (`.sy`), cards (`.it`).

---

## 4. Componentes (CSS literal)

### 4.1 Reset e utilitários

```css
*{box-sizing:border-box}
.hidden{display:none!important}
```

### 4.2 Badges / tags (pills com pontinho colorido)

Estrutura HTML: `<span class="tag t-real"><i></i>Lucro Real</span>` — o `<i>` é o dot.

```css
.tag{display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:650;
 border-radius:999px;padding:2px 9px;border:1px solid var(--ring);white-space:nowrap}
.tag i{width:7px;height:7px;border-radius:50%;display:block}
.t-real{color:var(--real)} .t-real i{background:var(--real)}
.t-pres{color:var(--pres)} .t-pres i{background:var(--pres)}
.t-nd{color:var(--nd)} .t-nd i{background:var(--nd)}
.t-ok{color:var(--ok)} .t-ok i{background:var(--ok)}
.t-aval{color:var(--aval)} .t-aval i{background:var(--aval)}
.t-rever{color:var(--rever)} .t-rever i{background:var(--rever)}
```

Padrão: pill de fundo transparente, borda `--ring`, texto e dot na cor da categoria/semântica.
Helper JS que gera a tag de categoria:

```js
const RG=r=>r?`<span class="tag ${r.indexOf('Real')>=0?'t-real':'t-pres'}"><i></i>${E(r)}</span>`
              :`<span class="tag t-nd"><i></i>Regime a levantar</span>`;
```

### 4.3 Cartões KPI (stat tiles)

HTML: `<div class="kpis"><div class="kpi"><div class="kl">LABEL</div><div class="kv">R$ 12,3 mi<small>/ano</small></div><div class="ks">nota</div></div>…</div>`

```css
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:10px;margin:20px 0}
.kpi{background:var(--surface);border:1px solid var(--ring);border-radius:11px;padding:14px 16px}
.kpi .kl{font-size:10px;text-transform:uppercase;letter-spacing:.08em;font-weight:800;color:var(--mut)}
.kpi .kv{font-size:22px;font-weight:750;letter-spacing:-.02em;font-variant-numeric:tabular-nums;margin-top:4px}
.kpi .kv small{font-size:13px;font-weight:650;color:var(--ink2)}
.kv.pos{color:var(--ok)} .kv.neg{color:var(--rever)}
.kpi .ks{font-size:11px;color:var(--mut);margin-top:3px;line-height:1.4}
```

Uso: sempre uma faixa de **4 KPIs** logo após o `p.lead` da aba. Valores positivos/negativos
colorem com `.pos`/`.neg`; sufixos de unidade (`/ano`) em `<small>`.

### 4.4 Tabela simples (wrapper `.tw`)

HTML: `<div class="tw"><table><thead>…<tbody>…</table></div>`

```css
.tw{background:var(--surface);border:1px solid var(--ring);border-radius:12px;overflow:hidden;margin:22px 0;overflow-x:auto}
table{border-collapse:collapse;width:100%;font-size:13px}
th,td{padding:11px 15px;text-align:left;border-bottom:1px solid var(--line);vertical-align:top}
thead th{font-size:10.5px;text-transform:uppercase;letter-spacing:.06em;color:var(--mut);font-weight:700;
 border-bottom:1.5px solid var(--edge);white-space:nowrap}
tbody tr:last-child td{border-bottom:0}
.mono{font-variant-numeric:tabular-nums;font-size:12px;color:var(--ink2)}
.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
```

Convenções de conteúdo: 1ª coluna com `<strong>` (nome), colunas de código/data com `.mono`,
numéricas com `.num`, colunas descritivas com `style="color:var(--ink2)"`.

### 4.5 Caixas de nota/alerta (`.note`)

```css
.note{background:var(--surface);border:1px solid var(--ring);border-left:3px solid var(--aval);
 border-radius:9px;padding:14px 18px;margin:20px 0;font-size:13px;color:var(--ink2)}
.note strong{color:var(--ink)}
.note.info{border-left-color:var(--real)}
.note.pos{border-left-color:var(--ok)}
.note.neg{border-left-color:var(--rever)}
```

Padrão: caixa com **borda esquerda de 3px** colorida por semântica. Default (sem modificador) =
âmbar/atenção. Sempre inicia com `<strong>Título.</strong>` seguido do texto corrido.
Usada para: leituras centrais, alertas estruturais, contradições, "o que mudou".

### 4.6 Diagrama de organograma (100% divs + flexbox — sem SVG)

O organograma é construído apenas com divs: linhas horizontais de nós (`.row`), conectores
verticais (`.stem`), uma linha horizontal (`.spine`, com largura calculada em JS) e a fileira de
filhos (`.kids`). Sequência: `row(topo) → stem → row(apex) → stem → spine → stem → kids`.

```css
.org{background:var(--surface);border:1px solid var(--ring);border-radius:14px;padding:34px 26px 30px;overflow-x:auto}
.row{display:flex;justify-content:center;gap:14px}
.node{border:1px solid var(--edge);border-radius:10px;background:var(--plane);padding:13px 14px;text-align:center}
.node .nm{font-weight:700;font-size:13px;line-height:1.3}
.node .cn{font-size:10.5px;color:var(--mut);margin-top:3px;font-variant-numeric:tabular-nums}
.node .pp{display:inline-block;margin-top:8px;font-size:13px;font-weight:700;font-variant-numeric:tabular-nums;
 background:var(--shell);color:#fff;border-radius:6px;padding:2px 9px}
:root[data-theme=dark] .node .pp{background:#31312d}
.node .rg{margin-top:8px}
.node .rs{font-size:9.5px;color:var(--mut);margin-top:4px;line-height:1.25;font-weight:600}
.node.apex .rs{font-size:11px}
.node.top{min-width:238px;max-width:262px}
.node.apex{border:2px solid var(--brand);background:var(--surface);min-width:320px;padding:16px 18px}
.node.apex .nm{font-size:15px;color:var(--brand);letter-spacing:-.01em}
.node.dash{border-style:dashed}
.stem{width:0;border-left:1.5px solid var(--edge);height:26px;margin:0 auto}
.spine{height:1.5px;background:var(--edge);margin:0 auto}
.kids{display:flex;justify-content:center;gap:7px;flex-wrap:nowrap;min-width:min-content;align-items:stretch}
.kids .node{flex:0 0 112px;width:112px;padding:12px 6px;display:flex;flex-direction:column;align-items:center}
.kids .node .nm{font-size:11px;height:3.4em;display:flex;align-items:center;justify-content:center}
.kids .node .cn{font-size:9.5px;height:2.3em;display:flex;align-items:center}
.kids .node .rg{margin-top:auto;padding-top:8px}
.kids .node .tag{font-size:9.5px;padding:2px 6px;gap:4px}
.kids .node .tag i{width:6px;height:6px}
.kids .node .pp{font-size:12px;padding:1px 7px}
```

Semântica visual: nó destacado `.apex` (borda 2px na cor da marca, nome na cor da marca);
`.dash` (borda tracejada) = dado a confirmar; `.pp` = pill escuro com percentual.
A largura da spine é calculada em JS: `W = nFilhos*112 + (nFilhos-1)*7` e aplicada inline
(`style="width:${W}px;max-width:100%"`).

Legenda do diagrama (swatches quadrados):

```css
.leg{display:flex;gap:18px;flex-wrap:wrap;margin-top:26px;padding-top:18px;border-top:1px solid var(--line);
 font-size:12px;color:var(--ink2);align-items:center}
.sw{width:13px;height:13px;border-radius:3px;border:1px solid var(--ring);display:inline-block;
 vertical-align:-2px;margin-right:6px}
```

Uso: `<span><span class="sw" style="background:var(--real)"></span>Lucro Real — 4 entidades</span>`.

### 4.7 Gráficos de barras horizontais (100% divs — sem biblioteca, sem SVG/canvas)

Barras horizontais construídas com grid de 3 colunas (label | trilha | valor). A largura do
preenchimento é % do máximo, calculada em JS e aplicada inline.

```css
.bars{background:var(--surface);border:1px solid var(--ring);border-radius:12px;padding:18px 20px 14px;margin:14px 0}
.bars .bt{font-size:12px;font-weight:750;margin-bottom:2px}
.bars .bs{font-size:11px;color:var(--mut);margin-bottom:12px}
.brow{display:grid;grid-template-columns:150px 1fr 110px;gap:10px;align-items:center;margin:0 0 7px}
.brow .bl{font-size:11.5px;font-weight:650;color:var(--ink);text-align:right;line-height:1.25}
.brow .tr{background:var(--plane);border-radius:4px;height:14px;position:relative;overflow:hidden}
.brow .fl{position:absolute;top:2px;bottom:2px;border-radius:3px}
.brow .bv{font-size:11.5px;font-variant-numeric:tabular-nums;color:var(--ink2);white-space:nowrap}
.brow.neg .bl{color:var(--rever)}
@media(max-width:700px){.brow{grid-template-columns:110px 1fr 90px}}
```

Gerador JS (literal — note o mínimo de 1.2% para barras não sumirem e vermelho para negativos):

```js
function bars(title,sub,rows,color,fmt){
 const mx=Math.max(...rows.map(r=>Math.abs(r.v)));
 return `<div class="bars"><div class="bt">${title}</div><div class="bs">${sub}</div>
 ${rows.map(r=>`<div class="brow${r.v<0?' neg':''}"><div class="bl">${E(r.n)}</div>
  <div class="tr"><div class="fl" style="left:0;width:${Math.max(1.2,Math.abs(r.v)/mx*100)}%;background:${r.v<0?'var(--rever)':color}"></div></div>
  <div class="bv">${fmt(r.v)}</div></div>`).join('')}</div>`;}
```

**Não há nenhuma biblioteca de gráficos, nenhum SVG, nenhum canvas em todo o dashboard.**

### 4.8 "Planilha" — tabela densa estilo spreadsheet (`.sheet` / `table.sh`)

Variante de tabela mais densa, com grade completa (bordas verticais), `table-layout:fixed` +
`<colgroup>` com larguras percentuais, hover de linha, linha de total e **codificação semântica
por borda esquerda de 3px na primeira célula**:

```css
.sheet{background:var(--surface);border:1px solid var(--edge);border-radius:8px;overflow:hidden;overflow-x:auto}
table.sh{border-collapse:collapse;width:100%;table-layout:fixed;font-size:12.5px;min-width:980px}
table.sh th{background:var(--plane);font-size:10px;text-transform:uppercase;letter-spacing:.07em;
 color:var(--mut);font-weight:700;text-align:left;padding:9px 11px;border-bottom:1px solid var(--edge);
 border-right:1px solid var(--line);white-space:normal;line-height:1.35;vertical-align:bottom}
table.sh th.a-r{text-align:right}
table.sh th:last-child{border-right:0}
table.sh td{padding:10px 11px;border-bottom:1px solid var(--line);border-right:1px solid var(--line);
 vertical-align:top;color:var(--ink2)}
table.sh td:last-child{border-right:0}
table.sh tbody tr:last-child td{border-bottom:0}
table.sh tbody tr:hover td{background:var(--plane)}
table.sh tr.tot td{background:var(--plane);font-weight:750;color:var(--ink);border-top:1.5px solid var(--edge)}
td.c-op{font-weight:700;color:var(--ink);border-left:3px solid transparent;line-height:1.4}
tr.k-pos td.c-op{border-left-color:var(--ok)}
tr.k-apr td.c-op{border-left-color:var(--aval)}
tr.k-neg td.c-op{border-left-color:var(--rever)}
tr.k-nd td.c-op{border-left-color:var(--edge)}
td.c-op .nn{display:block;font-size:10.5px;font-weight:700;color:var(--mut);
 font-variant-numeric:tabular-nums;margin-bottom:2px}
td.c-en{color:var(--ink);font-weight:600}
td.c-rg{font-size:11.5px;font-weight:650}
.r-real{color:var(--real)} .r-pres{color:var(--pres)} .r-nd{color:var(--mut);font-weight:400}
td.c-rg .df{display:block;font-size:10.5px;font-weight:600;color:var(--mut)}
td.c-es{text-align:right;font-variant-numeric:tabular-nums;font-weight:700;color:var(--ink);
 font-size:11.5px;white-space:nowrap}
td.c-es .np{color:var(--mut);font-weight:400;font-size:11.5px}
td.c-es.pos{color:var(--ok)} td.c-es.neg{color:var(--rever)}
td.c-cm{line-height:1.62;padding-right:16px}
td.c-cm b.lb{font-weight:800}
td.c-cm b.l-pos{color:var(--ok)}
td.c-cm b.l-apr{color:var(--aval)}
td.c-cm b.l-neg{color:var(--rever)}
td.c-cm .qq{display:block;margin-top:8px;padding-top:7px;border-top:1px dotted var(--edge);font-size:11.8px}
td.c-cm .qq b{color:var(--brand);font-weight:700;font-style:normal}
@media(max-width:1000px){table.sh{table-layout:auto;font-size:12px}}
```

Padrões de célula: `.c-op` (operação: micro-label `.nn` numerada acima do nome), `.c-en` (entidade),
`.c-rg` (categoria colorida com sub-rótulo `.df`), `.c-es` (valor à direita, `.pos`/`.neg`),
`.c-cm` (comentário longo iniciado por `<b class="lb l-pos">Rótulo.</b>` e complemento `.qq`
separado por borda pontilhada).

### 4.9 Legenda de planilha (`.shleg`)

Legenda com tracinhos (11×3px) em vez de quadrados — usada sob as planilhas:

```css
.shleg{display:flex;gap:18px;flex-wrap:wrap;font-size:11.5px;color:var(--mut);margin:11px 2px 0;
 align-items:center}
.shleg .k{display:inline-flex;align-items:center;gap:6px}
.shleg .k i{width:11px;height:3px;border-radius:2px;display:block}
```

Uso: `<span class="k"><i style="background:var(--ok)"></i>Possível — 7</span>`.

### 4.10 Cards de entidade/empresa (`.egrid` / `.ecard`)

Card com header (nome + código + badges), mini-tabela de métricas (label à esquerda com sub-linha
`.sub`; valor à direita) e rodapé de observações com eyebrow `.it`:

```css
.egrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(360px,1fr));gap:12px;margin:18px 0}
.ecard{background:var(--surface);border:1px solid var(--edge);border-radius:11px;overflow:hidden;display:flex;flex-direction:column}
.ecard .eh{padding:12px 14px 10px;border-bottom:1px solid var(--line);background:var(--plane);display:flex;gap:8px;align-items:baseline;flex-wrap:wrap}
.ecard .en{font-size:13.5px;font-weight:750;letter-spacing:-.01em}
.ecard .ec{font-size:10px;color:var(--mut);font-variant-numeric:tabular-nums}
.ecard .eb{margin-left:auto;display:flex;gap:5px;flex-wrap:wrap}
.ecard table.et{font-size:12px;width:100%;border-collapse:collapse}
.ecard table.et td{padding:6.5px 14px;border-bottom:1px solid var(--line)}
.ecard table.et td:last-child{text-align:right;font-variant-numeric:tabular-nums;font-weight:650;color:var(--ink);white-space:nowrap}
.ecard table.et tr:last-child td{border-bottom:0}
.ecard table.et td .sub{font-size:10px;color:var(--mut);display:block;font-weight:400}
.ecard .ei{padding:10px 14px 12px;border-top:1px solid var(--line);font-size:11.3px;color:var(--ink2);line-height:1.55}
.ecard .ei b{color:var(--ink)}
.ecard .ei .it{font-size:9px;text-transform:uppercase;letter-spacing:.09em;font-weight:800;color:var(--mut);display:block;margin-bottom:4px}
.v-pos{color:var(--ok)!important} .v-neg{color:var(--rever)!important}
```

### 4.11 Timeline horizontal (`.tl`)

Etapas em flex com **setas em chevron feitas de pseudo-elemento** (borda rotacionada 45°):

```css
.tl{display:flex;gap:0;margin:18px 0;overflow-x:auto}
.tl .st{flex:1;min-width:190px;position:relative;padding:0 14px 0 0}
.tl .st:not(:last-child):after{content:"";position:absolute;right:2px;top:17px;width:10px;height:10px;
 border-top:2px solid var(--edge);border-right:2px solid var(--edge);transform:rotate(45deg)}
.tl .stb{background:var(--surface);border:1px solid var(--edge);border-radius:10px;padding:12px 14px;height:100%}
.tl .sy{font-size:10px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:var(--mut)}
.tl .sv{font-size:13px;font-weight:750;margin:3px 0 4px}
.tl .sd{font-size:11.3px;color:var(--ink2);line-height:1.5}
```

HTML: `<div class="tl"><div class="st"><div class="stb"><div class="sy">2025</div><div class="sv">Título</div><div class="sd">descrição</div></div></div>…</div>`

### 4.12 Cards de contrato/ficha (`.cgrid` / `.ccard` com `<dl>`)

Ficha "campo: valor" com definition list em grid de 2 colunas:

```css
.cgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(430px,1fr));gap:12px;margin:18px 0}
@media(max-width:960px){.cgrid{grid-template-columns:1fr}}
.ccard{background:var(--surface);border:1px solid var(--edge);border-radius:11px;overflow:hidden}
.ccard .ch{padding:12px 16px;background:var(--plane);border-bottom:1px solid var(--line)}
.ccard .ct{font-size:13.5px;font-weight:750}
.ccard .cs{font-size:11px;color:var(--mut);margin-top:2px}
.ccard dl{margin:0;display:grid;grid-template-columns:132px 1fr;gap:0 14px;font-size:12px;padding:6px 16px 12px}
.ccard dt{color:var(--mut);font-weight:700;padding:6px 0;border-bottom:1px solid var(--line);font-size:10.5px;text-transform:uppercase;letter-spacing:.05em}
.ccard dd{margin:0;color:var(--ink2);padding:6px 0;border-bottom:1px solid var(--line)}
.ccard dt:last-of-type,.ccard dd:last-of-type{border-bottom:0}
.ccard dd b{color:var(--ink)}
```

### 4.13 Caixa "base" (premissas / metadados / continuidade)

Painel com título + subtítulo + `<dl>` de 2 colunas com divisores no topo:

```css
.base{background:var(--surface);border:1px solid var(--ring);border-radius:13px;padding:18px 22px 20px;margin:26px 0 0}
.base h3{margin:0 0 4px;font-size:14px;letter-spacing:-.01em}
.base .sub{font-size:11.5px;color:var(--mut);margin:0 0 14px}
.base dl{margin:0;display:grid;grid-template-columns:196px 1fr;gap:0 22px;font-size:12.5px}
.base dt{color:var(--ink);font-weight:700;padding:7px 0;border-top:1px solid var(--line)}
.base dd{margin:0;color:var(--ink2);padding:7px 0;border-top:1px solid var(--line)}
.base dt:first-of-type,.base dd:first-of-type{border-top:0}
@media(max-width:760px){.base dl{grid-template-columns:1fr}.base dd{padding-top:0;border-top:0}}
```

### 4.14 Lista ordenada de pendências

```css
ol.pend{margin:8px 0 0;padding-left:20px;font-size:13px;color:var(--ink2)}
ol.pend li{margin-bottom:8px;line-height:1.55}
ol.pend b{color:var(--ink)}
```

### 4.15 Checklist numerado (variação de tabela)

Tabela `.tw` sem thead onde a 1ª célula é o número formatado `01`, `02`… em cor de marca:

```html
<tr><td class="mono" style="width:44px;font-weight:700;color:var(--brand)">01</td>
    <td style="color:var(--ink2)">texto do item…</td></tr>
```

### 4.16 Componentes que NÃO existem no modelo

Sem tooltips, sem seções colapsáveis/accordion, sem sticky headers, sem animações/transições,
sem sombras (box-shadow **não é usado em lugar nenhum** — a elevação é feita só com bordas e
contraste `--surface` sobre `--plane`), sem ícones (exceto dots/chevrons em CSS puro),
sem imagens além do logo base64.

---

## 5. Layout

- **Container**: `max-width:1240px; margin:0 auto; padding:0 26px 90px` (`.wrap`). Header e nav
  repetem `max-width:1240px` + `padding … 26px` internamente para alinhar.
- **Grids responsivos com auto-fit/auto-fill + minmax** (nunca colunas fixas):
  - KPIs: `repeat(auto-fit,minmax(210px,1fr))`, gap 10px
  - Cards de empresa: `repeat(auto-fill,minmax(360px,1fr))`, gap 12px
  - Cards de contrato: `repeat(auto-fit,minmax(430px,1fr))`, gap 12px
  - Linha de barra: `grid-template-columns:150px 1fr 110px`
  - `<dl>` fichas: `132px 1fr` (ccard), `196px 1fr` (base)
- **Border-radius** (escala usada): 3px (swatch), 4px (trilha barra), 6px (pill pp), 7px (botão),
  8px (sheet), 9px (note), 10px (node, timeline), 11px (kpi, ecard, ccard), 12px (tw, bars),
  13px (base), 14px (org), 999px (tag pill).
- **Sombras: nenhuma.** Elevação = borda (`--ring` sutil ou `--edge` forte) + fundo `--surface`
  sobre página `--plane`.
- **Espaçamentos recorrentes**: seções `padding-top:36px`; h3.sec `margin:38px 0 6px`;
  componentes com `margin` vertical 14–26px; padding interno de cards 12–18px.
- **Overflow**: todo componente largo tem `overflow-x:auto` no próprio wrapper
  (`.org`, `.tw`, `.sheet`, `.tl`) — a página nunca rola horizontalmente.
- **Media queries** (todas as 5 do arquivo):
  - `@media(max-width:700px){.brow{grid-template-columns:110px 1fr 90px}}`
  - `@media(max-width:1000px){table.sh{table-layout:auto;font-size:12px}}`
  - `@media(max-width:960px){.cgrid{grid-template-columns:1fr}}`
  - `@media(max-width:760px){.base dl{grid-template-columns:1fr}.base dd{padding-top:0;border-top:0}}`
  - (implícita) grids auto-fit/auto-fill colapsam sozinhos em telas estreitas.
- Header usa `flex-wrap:wrap` para quebrar em telas pequenas; nav idem.

---

## 6. Tema (light/dark)

- **Default: light** (`<html lang="pt-BR" data-theme="light">`). Fundo off-white quente.
- Dark theme por **atributo `data-theme="dark"` no `<html>`**, que redefine as variáveis
  (bloco `:root[data-theme="dark"]{…}` — ver §2).
- **Auto-detecção** de preferência do sistema no load + **toggle manual** no header:

```js
const tg=document.getElementById('tg');
tg.onclick=()=>{const d=document.documentElement.getAttribute('data-theme')==='dark';
 document.documentElement.setAttribute('data-theme',d?'light':'dark');tg.textContent=d?'Modo escuro':'Modo claro';};
if(window.matchMedia&&window.matchMedia('(prefers-color-scheme: dark)').matches){
 document.documentElement.setAttribute('data-theme','dark');tg.textContent='Modo claro';}
```

- O botão alterna o próprio rótulo: "Modo escuro" ↔ "Modo claro". Não persiste em localStorage.
- Ajustes específicos de dark: `:root[data-theme=dark] .logo{filter:brightness(1.12)}` e
  `:root[data-theme=dark] .node .pp{background:#31312d}`.
- Como TUDO usa variáveis, nenhum outro seletor precisa de regra dark.

---

## 7. Padrões de conteúdo / organização da informação

Padrão editorial repetido em todas as abas:

1. **`<h2>` numerado** ("N. Título") ecoando o rótulo da aba.
2. **`p.lead`** logo abaixo: parágrafo introdutório dizendo fonte dos dados, o que a aba mostra e
   como ler (max-width 860px, cor secundária).
3. **Faixa de 4 KPIs** (nas abas quantitativas) com os agregados-chave — cada KPI com label caps,
   valor grande e nota `.ks` de contexto/ressalva.
4. **Corpo**: componentes intercalados — gráficos de barras, planilha densa, grid de cards,
   timeline, tabelas — sempre seguidos de **legenda** (`.shleg`/`.leg`) quando há codificação
   por cor.
5. **Caixas `.note`** pontuando o corpo com leitura interpretativa ("Leitura central…",
   "Onde o desenho ganha/perde dinheiro", "O que mudou da v2 para a v3") — sempre com
   `<strong>Título.</strong>` inicial e semântica de cor na borda esquerda.
6. **Sub-seções** com `h3.sec`.
7. **Premissas explícitas** em painel `.base` (dl "Premissa 1 / Premissa 2…") antes de tabelas
   de cálculo.
8. **Última aba = rastreabilidade**: fontes (tabela documento/data/o-que-sustenta), contradições
   entre fontes (série de `.note`), pendências (ol numerada) e painel de continuidade.
9. **Rodapé global** (`.foot`): "**Fontes.** …" por aba + "**Ressalvas.** …" (disclaimers,
   limitações metodológicas, "não substitui X").
10. Números sempre em formato pt-BR com tabular-nums; ausência = "—"; abreviação "mi"/"mil";
    metadados no título separados por " · ".
11. Semântica de cor consistente em TODO o dashboard: verde=positivo/ok, âmbar=atenção/aprofundar,
    vermelho terroso=negativo/risco, azul/laranja=categorias (dupla categórica), cinza=indefinido,
    tracejado=dado a confirmar.

### Conteúdo das abas do modelo (referência de organização)

| # | id | Rótulo da aba | Conteúdo |
|---|---|---|---|
| 1 | `org` | 1. Organograma | Diagrama de organograma (2 sócios topo → holding apex → 9 controladas), legenda de regimes, note info "mudança relevante", tabela "Entidades fora do organograma", tabela "Outras participações" |
| 2 | `emp` | 2. Radiografia 1S2026 | 4 KPIs agregados; 2 gráficos de barras (receita e resultado por empresa); note de leitura central; grid de 10 cards de empresa (header+badges, mini-tabela de 6 métricas, rodapé "Sinais intercompany"); note de fluxos societários |
| 3 | `op` | 3. Oportunidades (v3) | Planilha densa `.sh` de 16 fluxos (7 colunas: operação, quem fatura, regime, quem paga, regime, estimativa, parecer/racional/status) com linha TOTAL, legenda de status, note "o que mudou v2→v3" |
| 4 | `fin` | 4. Efeito financeiro | 4 KPIs; painel `.base` de premissas; planilha de cálculo por fluxo com subtotais e TOTAL; legenda; notes pos/neg ("onde ganha/perde dinheiro"); h3 + tabela de materialidade por empresa; note info comparativo com alternativa |
| 5 | `trib` | 5. Panorama tributário | 4 KPIs; timeline horizontal de 4 etapas (caso Portfel); tabela regime por empresa (2025×2026×evidência); série de 8 notes "Alertas estruturais" com semântica de cor |
| 6 | `ctr` | 6. Contratos | Grid de 2 cards-ficha de contrato (dl: objeto, preço, faturamento, vigência, efeito, cláusulas) com tag de status no header; note de contraste; checklist numerado em tabela |
| 7 | `src` | 7. Fontes e pendências | Tabela de fontes (documento/data/o que sustenta); notes de contradições; ol.pend de pendências; painel `.base` "Continuidade do trabalho" |

---

## 8. Snippet base (esqueleto fiel ao modelo)

Ponto de partida mínimo copiando o CSS central literal (reset, variáveis, header, tabs, KPI,
tabela, note, barras, legenda) e a mecânica JS de abas/tema/dados:

```html
<!DOCTYPE html>
<html lang="pt-BR" data-theme="light">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>TÍTULO — Subtítulo do dashboard</title>
<style>
:root{--surface:#fcfcfb;--plane:#f6f6f3;--ink:#0b0b0b;--ink2:#54534f;--mut:#8d8b85;
 --line:#e4e3dc;--edge:#c9c8bf;--ring:rgba(11,11,11,.09);
 --real:#2a78d6;--pres:#eb6834;--nd:#8d8b85;
 --ok:#0d8a2f;--aval:#c98500;--rever:#c0392b;--esp:#8d8b85;
 --brand:#e10913;--shell:#211f1e}
:root[data-theme="dark"]{--surface:#1a1a19;--plane:#111110;--ink:#fff;--ink2:#c3c2b7;--mut:#8d8b85;
 --line:#2c2c2a;--edge:#3a3a37;--ring:rgba(255,255,255,.11);
 --real:#3987e5;--pres:#e0703f;--ok:#2faa4c;--aval:#d19a1e;--rever:#e07168;--shell:#000}
*{box-sizing:border-box}
body{margin:0;background:var(--plane);color:var(--ink);
 font-family:system-ui,-apple-system,"Segoe UI",sans-serif;font-size:14.5px;line-height:1.6}
.wrap{max-width:1240px;margin:0 auto;padding:0 26px 90px}
header{background:var(--surface);color:var(--ink);border-bottom:1px solid var(--edge)}
.hd{max-width:1240px;margin:0 auto;padding:22px 26px 0;display:flex;gap:20px;align-items:center;flex-wrap:wrap}
.hdiv{width:1px;align-self:stretch;background:var(--line);margin:3px 0}
.hd .br{font-size:10px;letter-spacing:.18em;color:var(--brand);font-weight:700;margin:0 0 3px}
.hd h1{margin:0;font-size:20px;font-weight:700;letter-spacing:-.015em;color:var(--ink)}
.hd .sb{margin:4px 0 0;font-size:12.5px;color:var(--mut)}
.hd .rt{margin-left:auto}
.tg{background:transparent;color:var(--ink2);border:1px solid var(--edge);border-radius:7px;
 padding:6px 11px;font:inherit;font-size:12px;cursor:pointer}
.tg:hover{background:var(--plane);color:var(--ink)}
nav{max-width:1240px;margin:0 auto;padding:16px 26px 0;display:flex;gap:4px;flex-wrap:wrap}
nav button{background:transparent;border:0;border-bottom:2.5px solid transparent;color:var(--mut);
 padding:10px 16px;font:inherit;font-size:13.5px;font-weight:600;cursor:pointer}
nav button:hover{color:var(--ink)}
nav button[aria-selected=true]{color:var(--ink);border-bottom-color:var(--brand)}
section{padding-top:36px}
h2{font-size:21px;margin:0 0 8px;letter-spacing:-.02em;font-weight:700}
h3.sec{font-size:16px;margin:38px 0 6px;letter-spacing:-.015em}
p.lead{margin:0 0 26px;color:var(--ink2);max-width:860px}
.hidden{display:none!important}
/* badges */
.tag{display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:650;
 border-radius:999px;padding:2px 9px;border:1px solid var(--ring);white-space:nowrap}
.tag i{width:7px;height:7px;border-radius:50%;display:block}
.t-real{color:var(--real)} .t-real i{background:var(--real)}
.t-pres{color:var(--pres)} .t-pres i{background:var(--pres)}
.t-nd{color:var(--nd)} .t-nd i{background:var(--nd)}
.t-ok{color:var(--ok)} .t-ok i{background:var(--ok)}
.t-aval{color:var(--aval)} .t-aval i{background:var(--aval)}
.t-rever{color:var(--rever)} .t-rever i{background:var(--rever)}
/* tabela simples */
.tw{background:var(--surface);border:1px solid var(--ring);border-radius:12px;overflow:hidden;margin:22px 0;overflow-x:auto}
table{border-collapse:collapse;width:100%;font-size:13px}
th,td{padding:11px 15px;text-align:left;border-bottom:1px solid var(--line);vertical-align:top}
thead th{font-size:10.5px;text-transform:uppercase;letter-spacing:.06em;color:var(--mut);font-weight:700;
 border-bottom:1.5px solid var(--edge);white-space:nowrap}
tbody tr:last-child td{border-bottom:0}
.mono{font-variant-numeric:tabular-nums;font-size:12px;color:var(--ink2)}
.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
/* notas */
.note{background:var(--surface);border:1px solid var(--ring);border-left:3px solid var(--aval);
 border-radius:9px;padding:14px 18px;margin:20px 0;font-size:13px;color:var(--ink2)}
.note strong{color:var(--ink)}
.note.info{border-left-color:var(--real)}
.note.pos{border-left-color:var(--ok)}
.note.neg{border-left-color:var(--rever)}
/* KPIs */
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:10px;margin:20px 0}
.kpi{background:var(--surface);border:1px solid var(--ring);border-radius:11px;padding:14px 16px}
.kpi .kl{font-size:10px;text-transform:uppercase;letter-spacing:.08em;font-weight:800;color:var(--mut)}
.kpi .kv{font-size:22px;font-weight:750;letter-spacing:-.02em;font-variant-numeric:tabular-nums;margin-top:4px}
.kpi .kv small{font-size:13px;font-weight:650;color:var(--ink2)}
.kpi .ks{font-size:11px;color:var(--mut);margin-top:3px;line-height:1.4}
.kv.pos{color:var(--ok)} .kv.neg{color:var(--rever)}
/* barras */
.bars{background:var(--surface);border:1px solid var(--ring);border-radius:12px;padding:18px 20px 14px;margin:14px 0}
.bars .bt{font-size:12px;font-weight:750;margin-bottom:2px}
.bars .bs{font-size:11px;color:var(--mut);margin-bottom:12px}
.brow{display:grid;grid-template-columns:150px 1fr 110px;gap:10px;align-items:center;margin:0 0 7px}
.brow .bl{font-size:11.5px;font-weight:650;color:var(--ink);text-align:right;line-height:1.25}
.brow .tr{background:var(--plane);border-radius:4px;height:14px;position:relative;overflow:hidden}
.brow .fl{position:absolute;top:2px;bottom:2px;border-radius:3px}
.brow .bv{font-size:11.5px;font-variant-numeric:tabular-nums;color:var(--ink2);white-space:nowrap}
.brow.neg .bl{color:var(--rever)}
@media(max-width:700px){.brow{grid-template-columns:110px 1fr 90px}}
/* legenda */
.shleg{display:flex;gap:18px;flex-wrap:wrap;font-size:11.5px;color:var(--mut);margin:11px 2px 0;
 align-items:center}
.shleg .k{display:inline-flex;align-items:center;gap:6px}
.shleg .k i{width:11px;height:3px;border-radius:2px;display:block}
/* rodapé */
.foot{margin-top:44px;padding-top:18px;border-top:1px solid var(--line);font-size:12px;color:var(--mut)}
.foot b{color:var(--ink2)}
</style></head>
<body>
<header>
 <div class="hd">
  <!-- <img class="logo" src="data:image/png;base64,..."> opcional -->
  <div class="hdiv"></div>
  <div><p class="br">NOME DA MARCA</p><h1>Título do dashboard</h1>
   <p class="sb">Posição DD.MM.AAAA · fonte A · fonte B</p></div>
  <div class="rt"><button class="tg" id="tg" type="button">Modo escuro</button></div>
 </div>
 <nav id="nav" role="tablist"></nav>
</header>
<div class="wrap">
 <section id="t-a"></section>
 <section id="t-b" class="hidden"></section>
 <p class="foot" id="foot"></p>
</div>
<script id="p" type="application/json">
{"exemplo":{"kpis":[{"l":"Métrica X","v":12345678,"s":"nota de contexto"}]}}
</script>
<script>
const D=JSON.parse(document.getElementById('p').textContent);
const E=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const F0=n=>n==null?'—':n.toLocaleString('pt-BR',{maximumFractionDigits:0});
const FM=n=>n==null?'—':(Math.abs(n)>=1e6?(n/1e6).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+' mi':(n/1e3).toLocaleString('pt-BR',{maximumFractionDigits:0})+' mil');
function bars(title,sub,rows,color,fmt){
 const mx=Math.max(...rows.map(r=>Math.abs(r.v)));
 return `<div class="bars"><div class="bt">${title}</div><div class="bs">${sub}</div>
 ${rows.map(r=>`<div class="brow${r.v<0?' neg':''}"><div class="bl">${E(r.n)}</div>
  <div class="tr"><div class="fl" style="left:0;width:${Math.max(1.2,Math.abs(r.v)/mx*100)}%;background:${r.v<0?'var(--rever)':color}"></div></div>
  <div class="bv">${fmt(r.v)}</div></div>`).join('')}</div>`;}
function renderA(){document.getElementById('t-a').innerHTML=`
 <h2>1. Primeira aba</h2>
 <p class="lead">Texto introdutório da aba: fonte dos dados e como ler.</p>
 <div class="kpis">${D.exemplo.kpis.map(k=>`<div class="kpi"><div class="kl">${E(k.l)}</div>
  <div class="kv">R$ ${FM(k.v)}</div><div class="ks">${E(k.s)}</div></div>`).join('')}</div>
 <div class="note info"><strong>Leitura.</strong> Texto interpretativo da nota.</div>`;}
function renderB(){document.getElementById('t-b').innerHTML=`
 <h2>2. Segunda aba</h2><p class="lead">…</p>`;}
/* nav */
const T=[['a','1. Primeira aba'],['b','2. Segunda aba']];
const R={a:renderA,b:renderB}, done={};
function show(k){T.forEach(([id])=>{document.getElementById('t-'+id).classList.toggle('hidden',id!==k);
 document.getElementById('b-'+id).setAttribute('aria-selected',id===k);});
 if(!done[k]){R[k]();done[k]=1;} window.scrollTo({top:0,behavior:'instant'});}
document.getElementById('nav').innerHTML=T.map(([id,l])=>
 `<button id="b-${id}" role="tab" aria-selected="false" type="button">${l}</button>`).join('');
T.forEach(([id])=>document.getElementById('b-'+id).onclick=()=>show(id));
document.getElementById('foot').innerHTML='<b>Fontes.</b> … <br><br><b>Ressalvas.</b> …';
/* tema */
const tg=document.getElementById('tg');
tg.onclick=()=>{const d=document.documentElement.getAttribute('data-theme')==='dark';
 document.documentElement.setAttribute('data-theme',d?'light':'dark');tg.textContent=d?'Modo escuro':'Modo claro';};
if(window.matchMedia&&window.matchMedia('(prefers-color-scheme: dark)').matches){
 document.documentElement.setAttribute('data-theme','dark');tg.textContent='Modo claro';}
show('a');
</script></body></html>
```

Para os componentes adicionais (organograma, planilha `.sh`, cards `.ecard`, timeline `.tl`,
fichas `.ccard`, painel `.base`, `ol.pend`, `.leg`/`.sw`), copiar os blocos CSS literais do §4.

---

## 9. Checklist de fidelidade para o novo dashboard

- [ ] Arquivo único, sem dependências externas; dados em `<script id="p" type="application/json">`.
- [ ] Variáveis de cor exatamente as do §2 (light + dark); trocar apenas `--brand` se a nova marca exigir.
- [ ] Fonte system-ui 14.5px/1.6; `tabular-nums` em todo número; sem Google Fonts.
- [ ] Tabs sublinhadas na cor da marca com render lazy e `aria-selected`.
- [ ] Zero box-shadow; elevação por borda `--ring`/`--edge` e contraste `--surface`/`--plane`.
- [ ] Cada aba: h2 numerado → p.lead → (KPIs) → componentes + legendas → notes interpretativas.
- [ ] Última aba de fontes/contradições/pendências + rodapé global "Fontes/Ressalvas".
- [ ] Diagramas e gráficos em divs puros (flex/grid + larguras % inline), nunca biblioteca.
- [ ] Números pt-BR, "—" para ausente, "−" (U+2212) para negativo, "mi"/"mil" abreviados.
