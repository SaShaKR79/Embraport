#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Reconstrói a aba "Segregação de Risco" da planilha "Embracon_simulação tributária
04.09.26 v3" (arquivo do usuário) por cirurgia direta no pacote OOXML:
  - substitui xl/worksheets/sheet3.xml (a aba "Segregação de Risco");
  - acrescenta 2 formatos numéricos e alguns cellXfs em xl/styles.xml
    (reaproveitando fontes/preenchimentos/bordas já existentes);
  - remove calcChain e liga fullCalcOnLoad para o Excel recalcular ao abrir.
Todo o restante do pacote (logo EMF, abas ocultas, formatação das demais abas,
sharedStrings) fica byte a byte igual ao original.

Uso: python3 build_segregacao_sim.py <original.xlsx> <saida.xlsx>
"""
import sys, re, zipfile, math, copy
from xml.sax.saxutils import escape
from lxml import etree

SRC = sys.argv[1] if len(sys.argv) > 1 else "uploads/simulacao_0409_v3.xlsx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "Embracon_simulacao_tributaria_04.09.26_v4.xlsx"
CALC = "'Cálculos da Operação'"          # aba-fonte (referências externas)

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NSMAP = {"m": NS}

zin = zipfile.ZipFile(SRC)
files = {n: zin.read(n) for n in zin.namelist()}

# --------------------------------------------------------------------------
# 1. styles.xml — descobrir índices existentes e acrescentar o que falta
# --------------------------------------------------------------------------
styles = etree.fromstring(files["xl/styles.xml"])
fonts = styles.find("m:fonts", NSMAP)
fills = styles.find("m:fills", NSMAP)
borders = styles.find("m:borders", NSMAP)
numfmts = styles.find("m:numFmts", NSMAP)
cellxfs = styles.find("m:cellXfs", NSMAP)
xfs = list(cellxfs)

def border_sig(b):
    out = []
    for side in ("left", "right", "top", "bottom"):
        e = b.find("m:" + side, NSMAP)
        out.append(e.get("style") if e is not None and e.get("style") else None)
    return tuple(out)

border_idx = {}
for i, b in enumerate(borders):
    border_idx.setdefault(border_sig(b), i)

def get_border(left=None, right=None, top=None, bottom=None):
    sig = (left, right, top, bottom)
    if sig in border_idx:
        return border_idx[sig]
    b = etree.SubElement(borders, "{%s}border" % NS)
    for side, st in zip(("left", "right", "top", "bottom"), sig):
        e = etree.SubElement(b, "{%s}%s" % (NS, side))
        if st:
            e.set("style", st)
            c = etree.SubElement(e, "{%s}color" % NS); c.set("indexed", "64")
    etree.SubElement(b, "{%s}diagonal" % NS)
    borders.set("count", str(len(borders)))
    border_idx[sig] = len(borders) - 1
    return border_idx[sig]

# fontes existentes (ordem verificada no arquivo-fonte)
F_REG, F_IT, F_B, F_BI, F_IT20, F_B14W = 0, 12, 15, 16, 13, 19
# preenchimentos existentes
FILL_NONE, FILL_GRAY, FILL_RED = 0, 6, 9
# bordas
B_NONE = get_border()
B_BOT = get_border(bottom="thin")
B_TOP = get_border(top="thin")
B_TB = get_border(top="thin", bottom="thin")
B_LR = get_border(left="thin", right="thin")

# formatos numéricos novos
existing_ids = {int(n.get("numFmtId")) for n in numfmts}
def add_numfmt(code):
    for n in numfmts:
        if n.get("formatCode") == code:
            return int(n.get("numFmtId"))
    nid = max(existing_ids | {176}) + 1
    existing_ids.add(nid)
    e = etree.SubElement(numfmts, "{%s}numFmt" % NS)
    e.set("numFmtId", str(nid)); e.set("formatCode", code)
    numfmts.set("count", str(len(numfmts)))
    return nid

NF_NUM = add_numfmt('#,##0.0;[Red]\\(#,##0.0\\);\\-')
NF_CHK = add_numfmt('0.0;[Red]\\-0.0;"ok"')
NF_INT = 3      # #,##0
NF_PCT = 168    # 0.0%
NF_GEN = 0

xf_cache = {}
def xf(font=F_REG, fill=FILL_NONE, border=B_NONE, nf=NF_GEN, h=None, v=None, wrap=False, indent=0):
    key = (font, fill, border, nf, h, v, wrap, indent)
    if key in xf_cache:
        return xf_cache[key]
    e = etree.SubElement(cellxfs, "{%s}xf" % NS)
    e.set("numFmtId", str(nf)); e.set("fontId", str(font)); e.set("fillId", str(fill))
    e.set("borderId", str(border)); e.set("xfId", "0")
    if nf: e.set("applyNumberFormat", "1")
    e.set("applyFont", "1")
    if fill: e.set("applyFill", "1")
    if border: e.set("applyBorder", "1")
    if h or v or wrap or indent:
        e.set("applyAlignment", "1")
        a = etree.SubElement(e, "{%s}alignment" % NS)
        if h: a.set("horizontal", h)
        if v: a.set("vertical", v)
        if wrap: a.set("wrapText", "1")
        if indent: a.set("indent", str(indent))
    cellxfs.set("count", str(len(cellxfs)))
    xf_cache[key] = len(cellxfs) - 1
    return xf_cache[key]

# --- estilos usados na aba (reaproveitando índices já existentes quando idênticos)
S = {}
S["title20"]  = 103                                   # itálico 20 (G4 da aba Cálculos)
S["date"]     = 102                                   # itálico 11
S["redbar"]   = 206                                   # barra vermelha, negrito 14 branco
S["sub"]      = xf(F_B, border=B_BOT, h="left", v="center")
S["para"]     = xf(F_IT, h="left", v="top", wrap=True)
S["para_reg"] = xf(F_REG, h="left", v="top", wrap=True)
S["lab_it"]   = xf(F_IT, h="left", v="center", wrap=True)
S["lab_reg"]  = xf(F_REG, h="left", v="center", wrap=True)
S["lab_b"]    = xf(F_B, h="left", v="center", wrap=True)
S["lab_b_top"]= xf(F_B, border=B_TOP, h="left", v="center", wrap=True)
S["lab_bi"]   = xf(F_BI, h="left", v="center", wrap=True)
S["note"]     = xf(F_IT, h="left", v="center", wrap=True)
S["hdr_l"]    = xf(F_B, fill=FILL_GRAY, border=B_BOT, h="left", v="center", wrap=True)
S["hdr_c"]    = xf(F_B, fill=FILL_GRAY, border=B_BOT, h="center", v="center", wrap=True)
S["num"]      = xf(F_REG, nf=NF_NUM, h="center", v="center")
S["num_it"]   = xf(F_IT, nf=NF_NUM, h="center", v="center")
S["num_b"]    = xf(F_B, nf=NF_NUM, h="center", v="center")
S["num_b_top"]= xf(F_B, border=B_TOP, nf=NF_NUM, h="center", v="center")
S["num_bi"]   = xf(F_BI, nf=NF_NUM, h="center", v="center")
S["pct"]      = xf(F_REG, nf=NF_PCT, h="center", v="center")
S["pct_it"]   = xf(F_IT, nf=NF_PCT, h="center", v="center")
S["pct_b"]    = xf(F_B, nf=NF_PCT, h="center", v="center")
S["pct_b_top"]= xf(F_B, border=B_TOP, nf=NF_PCT, h="center", v="center")
S["txt_c"]    = xf(F_REG, h="center", v="center")
S["txt_c_it"] = xf(F_IT, h="center", v="center")
S["dash"]     = xf(F_IT, h="center", v="center")
S["dash_top"] = xf(F_IT, border=B_TOP, h="center", v="center")
S["chk"]      = xf(F_IT, nf=NF_CHK, h="center", v="center")
S["inp_lab"]  = xf(F_B, fill=FILL_GRAY, h="left", v="center", wrap=True)
S["inp_int"]  = xf(F_B, fill=FILL_GRAY, nf=NF_INT, h="center", v="center")
S["inp_num"]  = xf(F_B, fill=FILL_GRAY, nf=NF_NUM, h="center", v="center")
S["inp_pct"]  = xf(F_B, fill=FILL_GRAY, nf=NF_PCT, h="center", v="center")
S["cell_it"]  = xf(F_IT, h="left", v="center", wrap=True)
S["cell_top"] = xf(F_IT, border=B_TOP, h="left", v="center", wrap=True)
S["blank_top"]= xf(F_REG, border=B_TOP)

files["xl/styles.xml"] = etree.tostring(styles, xml_declaration=True, encoding="UTF-8", standalone=True)

# --------------------------------------------------------------------------
# 2. conteúdo da aba
# --------------------------------------------------------------------------
cells = {}      # (row, col) -> dict(kind, value, style)
merges = []
row_h = {}
COLS = "ABCDEFGHIJK"
def col_i(c): return COLS.index(c) + 1

def put(ref, value=None, style=None, kind=None):
    col = re.match(r"[A-Z]+", ref).group(0); row = int(ref[len(col):])
    if kind is None:
        if value is None: kind = "blank"
        elif isinstance(value, str) and value.startswith("="): kind = "f"
        elif isinstance(value, str): kind = "s"
        else: kind = "n"
    cells[(row, col_i(col))] = dict(kind=kind, value=value, style=style)

def merge(rng, style=None):
    merges.append(rng)
    a, b = rng.split(":")
    ca, ra = re.match(r"([A-Z]+)(\d+)", a).groups(); cb, rb = re.match(r"([A-Z]+)(\d+)", b).groups()
    for r in range(int(ra), int(rb) + 1):
        for c in range(col_i(ca), col_i(cb) + 1):
            if (r, c) not in cells:
                cells[(r, c)] = dict(kind="blank", value=None, style=style)
            elif style is not None and cells[(r, c)]["style"] is None:
                cells[(r, c)]["style"] = style

def para(row, text, style="para", cols="B:H", chars_per_line=135, first_last=None):
    a, b = cols.split(":")
    put(f"{a}{row}", text, S[style]); merge(f"{a}{row}:{b}{row}", S[style])
    lines = max(1, math.ceil(len(text) / chars_per_line))
    row_h[row] = 15.0 * lines + 4

def redbar(row, text):
    put(f"B{row}", text, S["redbar"]); merge(f"B{row}:K{row}", S["redbar"]); row_h[row] = 18.5

def sub(row, text, cols="B:H"):
    a, b = cols.split(":")
    put(f"{a}{row}", text, S["sub"]); merge(f"{a}{row}:{b}{row}", S["sub"]); row_h[row] = 16

def header(row, items, first_left=True, height=32):
    for i, (col, text) in enumerate(items):
        put(f"{col}{row}", text, S["hdr_l"] if (i == 0 and first_left) else S["hdr_c"])
    row_h[row] = height

# ---- cabeçalho da aba (espelha a aba "Cálculos da Operação": logo nas linhas 2-5, título em G4)
put("G4", "Segregação de risco", S["title20"]); row_h[4] = 26
put("G5", "4 de setembro de 2026", S["date"])
redbar(8, "Segregação de Risco — Custo do Investimento das Holdings (CPC 15 × CPC 19)")

# ---- 0. O que esta aba analisa
sub(10, "O que esta aba analisa")
para(11, "Esta aba complementa a aba 'Cálculos da Operação'. Lá se apura o custo do investimento das Holdings "
         "(Savian e JVFJ) na Embracon após a incorporação da CNP e o imposto do cash-out, comparando o custo "
         "original (cenário A) com o custo incrementado pelo reflexo da incorporação (cenário B). Aqui respondemos "
         "a três perguntas que decorrem daquele resultado: (1) de que camadas é formado o custo incrementado e qual "
         "o risco fiscal de cada uma, no enquadramento da operação como combinação de negócios (CPC 15); (2) o que "
         "muda se a operação for enquadrada como operação em conjunto (CPC 19), em que as Holdings reconhecem 60% do "
         "valor justo de toda a JO e um imposto diferido 'carimbado' de 34%; e (3) quanto de imposto seria devido em "
         "uma venda futura dos 60%, pela Holding (PJ) ou pelas pessoas físicas (PF), em cada cenário. Os valores "
         "partem das premissas da aba 'Cálculos da Operação' (relacionadas abaixo) e estão em R$ milhões; as "
         "premissas específicas de cada bloco são indicadas no próprio bloco.")

# ---- Informações utilizadas
sub(13, "Informações utilizadas (fonte: aba 'Cálculos da Operação', salvo indicação em contrário)")
header(14, [("B", "Informação"), ("C", "Valor")], height=16)
put("D14", "Origem / observação", S["hdr_l"]); merge("D14:H14", S["hdr_l"])
info = [
    (15, "Data-base dos balanços",                                   "30.06.2026", "txt_c_it", "Balanços das companhias utilizados na aba 'Cálculos da Operação'"),
    (16, "Unidade",                                                  "R$ milhões", "txt_c_it", ""),
    (17, "Participação das Partes CNP na incorporação (relação de troca)", f"={CALC}!I15", "pct_it", "Ações novas emitidas pela Embracon às Partes CNP na incorporação"),
    (18, "Participação-alvo final das Partes CNP",                   f"={CALC}!I16", "pct_it", "Após o cash-out (compra de 26,2% das Holdings)"),
    (19, "Participação final das famílias (Savian + JVFJ)",          "=1-C18",      "pct_it", "Posição remanescente das Holdings, objeto da venda futura simulada no bloco 3"),
    (20, "Fração do investimento das Holdings alienada no cash-out", f"={CALC}!D69", "pct_it", "13,1% vendidos ÷ 43,1% detidos por Holding — mesmo critério de alocação de custo da aba 'Cálculos da Operação'"),
    (21, "IRPJ/CSLL — ganho de capital PJ",                          f"={CALC}!I17", "pct_it", "Alíquota combinada aplicada às Holdings"),
    (22, "PL contábil da Embracon (antes da incorporação)",          f"={CALC}!C19", "num_it", "PL da Embracon; corresponde ao custo original das Holdings"),
    (23, "Custo original do investimento das Holdings (Savian + JVFJ)", f"={CALC}!F47", "num_it", "Investimento original em Embracon: 127 de cada Holding"),
    (24, "PL contábil da CNP",                                       f"={CALC}!C23", "num_it", "Patrimônio líquido contábil da CNP incorporada"),
    (25, "Valor de mercado da CNP (valor pelo qual o acervo é incorporado no CPC 15 — PPA)", f"={CALC}!C24", "num_it", "Valor de mercado adotado na aba 'Cálculos da Operação'"),
    (26, "Mais-valia da CNP (valor de mercado − PL contábil)",       "=C25-C24",    "num_it", "Camada de avaliação a valor justo (AVJ) que o CPC 15 leva ao custo das Holdings via MEP"),
]
for r, lab, val, st, note in info:
    put(f"B{r}", lab, S["lab_it"]); put(f"C{r}", val, S[st])
    if note: put(f"D{r}", note, S["note"]); merge(f"D{r}:H{r}", S["note"])
    row_h[r] = 15 if len(lab) < 42 else 30

# ---- 1) CPC 15
redbar(28, "1) CPC 15 — Combinação de Negócios: de que é feito o custo do investimento das Holdings e onde está o risco")
para(30, "Na combinação de negócios (CPC 15) a Embracon adquire o controle do negócio da CNP: o acervo da CNP entra pelo "
         "valor de mercado apurado em laudo (PPA) e o patrimônio da própria Embracon não é reavaliado. Nas Holdings, o "
         "investimento sobe pelo método de equivalência patrimonial (MEP): ao custo original soma-se o reflexo do acervo "
         "incorporado, deduzida a perda pela diluição de 13,8%. O quadro abre esse custo em camadas porque cada camada tem "
         "um risco fiscal diferente. O custo original é incontroverso. O reflexo do PL contábil da CNP depende da tese de "
         "que o resultado de equivalência é neutro para fins fiscais também na alienação futura (art. 33, §2º, do DL "
         "1.598/77; precedente WTorre). O reflexo da mais-valia (valor de mercado acima do PL contábil) é a parcela mais "
         "exposta, por corresponder a uma avaliação a valor justo que nunca foi tributada. As colunas mostram o valor a "
         "100% (PL da Embracon após a incorporação), a parcela das Holdings logo após a incorporação (86,2%) e a parcela "
         "remanescente após o cash-out (60%).")
sub(32, "1.1  Composição do custo do investimento das Holdings, por camada")
header(33, [("B", "Camada do custo"), ("C", "PL Embracon pós-incorporação (100%)"),
            ("D", "Holdings pós-incorporação (86,2%)"), ("E", "Holdings após o cash-out (60%)"),
            ("F", "Baixa no cash-out"), ("G", "IR potencial (34%) sobre a camada"), ("H", "Leitura de risco")], height=48)
layers = [
    (34, "Custo original (PL contábil da Embracon)", "=C22", "=C23", "=D34*(1-$C$20)", None,
         "Incontroverso: custo histórico das ações; é o custo do cenário A da aba 'Cálculos da Operação'."),
    (35, "(−) Perda por diluição de 13,8% sobre o custo original", None, "=-D34*$C$17", "=D35*(1-$C$20)", "=E35*$C$21",
         "Perda de MEP por variação de percentual: fiscalmente neutra (art. 33, §2º, DL 1.598/77); reduz o reflexo contábil (e o IR potencial da camada seguinte), não o custo fiscal original."),
    (36, "Reflexo do PL contábil da CNP", "=C24", "=C36*(1-$C$17)", "=C36*$C$19", "=E36*$C$21",
         "Risco limitado: reflexo MEP do PL incorporado a valor contábil; sustenta-se na tese de neutralidade do MEP (WTorre), com precedentes CARF 2024 em sentido contrário."),
    (37, "Reflexo da mais-valia da CNP (AVJ do PPA)", "=C26", "=C37*(1-$C$17)", "=C37*$C$19", "=E37*$C$21",
         "Camada mais exposta: valor de mercado acima do PL contábil, nunca tributado; é a parcela que o fisco tenderia a atacar primeiro."),
]
for r, lab, c, d, e, g, h in layers:
    put(f"B{r}", lab, S["lab_it"])
    put(f"C{r}", c if c else "–", S["num_it"] if c else S["dash"])
    put(f"D{r}", d, S["num_it"]); put(f"E{r}", e, S["num_it"]); put(f"F{r}", f"=E{r}-D{r}", S["num_it"])
    put(f"G{r}", g if g else "–", S["num_it"] if g else S["dash"])
    put(f"H{r}", h, S["cell_it"]); row_h[r] = 60
put("B38", "(=) Custo do investimento das Holdings", S["lab_b_top"])
for col in "CDEF": put(f"{col}38", f"=SUM({col}34:{col}37)", S["num_b_top"])
put("G38", "=SUM(G35:G37)", S["num_b_top"])
put("H38", "Total das camadas: após a incorporação = custo incrementado; após o cash-out = custo remanescente do cenário B (aba 'Cálculos da Operação').", S["cell_top"]); row_h[38] = 45
put("B39", "Conferência com a aba 'Cálculos da Operação' (diferença; 'ok' = zero)", S["lab_it"])
put("C39", "–", S["dash"])
put("D39", f"=ROUND(D38-{CALC}!E53,6)", S["chk"])
put("E39", f"=ROUND(E38-{CALC}!J89,6)", S["chk"])
put("F39", f"=ROUND(F38+{CALC}!J82,6)", S["chk"])
put("G39", "–", S["dash"])
put("H39", "D = custo incrementado (E53) · E = custo remanescente do cenário B (J89) · F = custo alocado no cash-out (J82)", S["cell_it"]); row_h[39] = 45

sub(41, "1.2  Leitura: quanto do custo remanescente é seguro e quanto está em discussão (após o cash-out)")
header(42, [("B", "Parcela do custo após o cash-out"), ("C", "Valor"), ("D", "% do custo"), ("E", "IR potencial (34%)")], height=32)
put("F42", "Comentário", S["hdr_l"]); merge("F42:H42", S["hdr_l"])
lead = [
    (43, "Custo fiscal incontroverso (custo original remanescente = cenário A)", "=E34", None,
         "Custo aceito em qualquer leitura; é o cenário mais conservador da aba 'Cálculos da Operação'."),
    (44, "Reflexo do PL contábil da CNP, líquido da perda por diluição (risco limitado)", "=E36+E35", "=C44*$C$21",
         "Depende da tese de neutralidade do reflexo de MEP; o IR só seria devido se o fisco tributar o reflexo na alienação."),
    (45, "Reflexo da mais-valia da CNP (maior exposição)", "=E37", "=C45*$C$21",
         "Parcela do PPA levada ao custo; com as premissas atuais é pequena, porque o valor de mercado da CNP está próximo do seu PL contábil."),
]
for r, lab, c, e, note in lead:
    put(f"B{r}", lab, S["lab_it"]); put(f"C{r}", c, S["num_it"]); put(f"D{r}", f"=C{r}/$C$46", S["pct_it"])
    put(f"E{r}", e if e else "–", S["num_it"] if e else S["dash"])
    put(f"F{r}", note, S["cell_it"]); merge(f"F{r}:H{r}", S["cell_it"]); row_h[r] = 30
put("B46", "(=) Custo após o cash-out (cenário B) — incremento em discussão = linhas 2 + 3", S["lab_b_top"])
put("C46", "=SUM(C43:C45)", S["num_b_top"]); put("D46", "=SUM(D43:D45)", S["pct_b_top"]); put("E46", "=SUM(E44:E45)", S["num_b_top"])
put("F46", "O incremento em discussão (C44 + C45) é exatamente a diferença de custo remanescente entre os cenários B e A da aba 'Cálculos da Operação'.", S["cell_top"]); merge("F46:H46", S["cell_top"]); row_h[46] = 30
put("B47", "Conferência: incremento em discussão − Δ custo remanescente da aba 'Cálculos da Operação' (K89)", S["lab_it"])
put("C47", f"=ROUND(C44+C45-{CALC}!K89,6)", S["chk"]); row_h[47] = 30

# ---- 2) CPC 19
redbar(49, "2) CPC 19 — Joint Operation: reconhecimento de 60% do valor justo da JO e imposto diferido 'carimbado'")
para(51, "Na operação em conjunto (CPC 19), leitura sustentada pela Deloitte a partir dos vetos qualificados da CNP, as "
         "Holdings deixam de avaliar o investimento pelo MEP e passam a reconhecer diretamente a sua parcela (60%) dos "
         "ativos e passivos da JO — Embracon e CNP — a valor justo. Isso inclui o valor justo da própria Embracon, que no "
         "CPC 15 não é reavaliada. O ganho de avaliação a valor justo (AVJ) não é tributado no reconhecimento, desde que "
         "controlado em subconta (arts. 13 e 14 da Lei 12.973/14), mas fica 'carimbado': o IRPJ/CSLL de 34% sobre o AVJ é "
         "devido quando o investimento for realizado (venda), qualquer que seja o preço e independentemente de quem venda. "
         "O bloco usa como valor justo da Embracon o enterprise value do Acordo de Investimento (premissa editável abaixo) "
         "e, para a CNP, o valor de mercado já adotado na aba 'Cálculos da Operação'.")
sub(53, "2.1  Premissa adicional deste bloco")
put("B54", "Valor justo da Embracon — enterprise value do Acordo de Investimento (data-base 31.12.2024)", S["inp_lab"])
put("C54", 5133, S["inp_int"])
put("D54", "Premissa editável. No CPC 15 a Embracon não é reavaliada; no CPC 19 o seu valor justo entra integralmente no investimento das Holdings.", S["note"]); merge("D54:H54", S["note"]); row_h[54] = 30
sub(56, "2.2  Investimento das Holdings na JO (após o cash-out, 60%)")
header(57, [("B", "Componente"), ("C", "100% da JO"), ("D", "60% (Holdings)")], height=32)
put("E57", "Leitura", S["hdr_l"]); merge("E57:H57", S["hdr_l"])
jo = [
    (58, "Valor justo da Embracon", "=C54", "=C58*$C$19", "lab_it", "num_it", "Parcela que não existe no CPC 15: a Embracon passa a ser reavaliada."),
    (59, "Valor justo da CNP (= valor de mercado)", "=C25", "=C59*$C$19", "lab_it", "num_it", "Mesmo valor incorporado no CPC 15."),
    (60, "(=) Valor justo total da JO — custo do investimento das Holdings no CPC 19", "=C58+C59", "=D58+D59", "lab_b_top", "num_b_top", "Investimento reconhecido pelas Holdings: 60% do valor justo de Embracon + CNP."),
    (61, "(−) Base contábil (PL da Embracon + PL contábil da CNP)", "=C22+C24", "=C61*$C$19", "lab_it", "num_it", "Coincide com o custo do CPC 15 sem a camada da mais-valia (bloco 1: custo original + diluição + reflexo do PL contábil)."),
    (62, "(=) Ganho de AVJ com tributação diferida (subconta — arts. 13 e 14, Lei 12.973/14)", "=C60-C61", "=D60-D61", "lab_b", "num_b", "Ganho reconhecido no resultado das Holdings e neutralizado fiscalmente apenas enquanto controlado em subconta."),
    (63, "Passivo fiscal diferido 'carimbado' (34% × AVJ)", "=C62*$C$21", "=D62*$C$21", "lab_bi", "num_bi", "Devido na realização do investimento (venda), independentemente do preço; é o piso de imposto do CPC 19 nos blocos 3 e 5."),
]
for r, lab, c, d, ls, ns, note in jo:
    put(f"B{r}", lab, S[ls]); put(f"C{r}", c, S[ns]); put(f"D{r}", d, S[ns])
    put(f"E{r}", note, S["cell_top"] if ls.endswith("top") else S["cell_it"]); merge(f"E{r}:H{r}", S["cell_top"] if ls.endswith("top") else S["cell_it"]); row_h[r] = 30
sub(65, "2.3  Custo do investimento das Holdings após o cash-out — CPC 15 × CPC 19")
header(66, [("B", "Cenário"), ("C", "Custo (60%)"), ("D", "Δ vs. CPC 15")], height=16)
put("E66", "Leitura", S["hdr_l"]); merge("E66:H66", S["hdr_l"])
put("B67", "CPC 15 — custo incrementado remanescente (bloco 1)", S["lab_it"]); put("C67", "=E38", S["num_it"]); put("D67", "–", S["dash"])
put("E67", "Custo formado pelo custo original mais o reflexo da CNP a valor de mercado; sem passivo fiscal diferido.", S["cell_it"]); merge("E67:H67", S["cell_it"]); row_h[67] = 30
put("B68", "CPC 19 — 60% do valor justo da JO (bloco 2)", S["lab_it"]); put("C68", "=D60", S["num_it"]); put("D68", "=C68-C67", S["num_it"])
put("E68", "A diferença é o AVJ de toda a JO (60%) menos a mais-valia da CNP que o CPC 15 já leva ao custo; sobre o AVJ há IR diferido de 34%.", S["cell_it"]); merge("E68:H68", S["cell_it"]); row_h[68] = 30

# ---- 3) Venda futura
redbar(70, "3) Venda futura dos 60% — IRPJ/CSLL em cada cenário, pela Holding (PJ) ou pelas pessoas físicas (PF)")
para(72, "Simulamos a venda dos 60% remanescentes em três preços — abaixo, igual e acima do valor justo (60% do valor "
         "justo da JO). No CPC 15, o imposto incide sobre a diferença entre o preço e o custo fiscal apurado no bloco 1 "
         "(custo incrementado remanescente). No CPC 19, o imposto diferido de 34% sobre o AVJ é devido de qualquer forma e "
         "funciona como piso; se o preço superar o valor justo, o excedente também é tributado. Na venda pelas pessoas "
         "físicas (após 'subir' as ações das Holdings para as PFs), aplica-se a alíquota simplificada de 22,5% sobre o "
         "ganho; porém o custo das PFs é o custo das quotas das Holdings, muito inferior ao custo da Holding, e no CPC 19 o "
         "imposto carimbado permanece na Holding mesmo que a venda seja feita pela PF.")
sub(74, "3.1  Premissas deste bloco")
inputs3 = [
    (75, "Preço de venda abaixo do valor justo", 1000, "inp_int", "Premissa editável (cenário i do material SF)."),
    (76, "Preço de venda igual ao valor justo (60% do valor justo da JO)", "=D60", "num_b", "Calculado a partir do bloco 2."),
    (77, "Preço de venda acima do valor justo", 5000, "inp_int", "Premissa editável (cenário iii do material SF)."),
    (78, "Alíquota de IR da pessoa física — ganho de capital (simplificada)", 0.225, "inp_pct", "Alíquota progressiva de 15% a 22,5%; adotada a máxima, por simplificação."),
    (79, "Custo das quotas das Holdings nas pessoas físicas", 18, "inp_int", "Premissa ilustrativa do exercício da MLA, a confirmar com o custo real das quotas."),
]
for r, lab, val, st, note in inputs3:
    put(f"B{r}", lab, S["inp_lab"] if st.startswith("inp") else S["lab_it"]); put(f"C{r}", val, S[st])
    put(f"D{r}", note, S["note"]); merge(f"D{r}:H{r}", S["note"]); row_h[r] = 15 if len(lab) < 42 else 30
sub(81, "3.2  Imposto na venda dos 60% (IRPJ/CSLL de 34% na Holding; IR de 22,5% na PF)")
header(82, [("B", "Cenário de preço"), ("C", "Preço de venda"), ("D", "CPC 15 — venda pela Holding"),
            ("E", "CPC 15 — venda pela PF"), ("F", "CPC 19 — venda pela Holding"), ("G", "CPC 19 — venda pela PF"),
            ("H", "Δ CPC 19 − CPC 15 (venda pela Holding)")], height=48)
prices = [(83, "Abaixo do valor justo", "=C75"), (84, "Igual ao valor justo", "=C76"), (85, "Acima do valor justo", "=C77")]
for r, lab, p in prices:
    put(f"B{r}", lab, S["lab_it"]); put(f"C{r}", p, S["num_b"])
    put(f"D{r}", f"=(C{r}-$E$38)*$C$21", S["num"])
    put(f"E{r}", f"=(C{r}-$C$79)*$C$78", S["num"])
    put(f"F{r}", f"=MAX($D$63,(C{r}-$D$61)*$C$21)", S["num"])
    put(f"G{r}", f"=$D$63+MAX(0,C{r}-$D$60)*$C$78", S["num"])
    put(f"H{r}", f"=F{r}-D{r}", S["num_b"]); row_h[r] = 16
put("B86", "Fórmulas: CPC 15 PJ = (preço − custo do bloco 1) × 34% · CPC 15 PF = (preço − custo das quotas) × 22,5% · "
           "CPC 19 PJ = maior entre o imposto carimbado e (preço − base contábil) × 34% · CPC 19 PF = imposto carimbado na Holding + 22,5% sobre o excedente ao valor justo.", S["note"])
merge("B86:H86", S["note"]); row_h[86] = 45
para(87, "Leitura: em preços iguais ou superiores ao valor justo, a diferença entre os regimes na venda pela Holding é "
         "pequena e corresponde a 34% da mais-valia da CNP — a única camada que o CPC 15 leva ao custo e o CPC 19 tributa. "
         "Abaixo do valor justo, o CPC 19 é substancialmente mais oneroso, porque o imposto diferido não diminui com o preço. "
         "A venda pela PF só é vantajosa no CPC 15, e apenas a partir de preços em que a alíquota menor compensa o custo menor "
         "das quotas; no CPC 19 ela não elimina o imposto carimbado, que permanece na Holding.")
sub(89, "3.3  Por que o imposto diferido do CPC 19 funciona como piso — venda abaixo do valor justo")
header(90, [("B", "Movimento na Holding (CPC 19)"), ("C", "Valor")], height=16)
put("D90", "Comentário", S["hdr_l"]); merge("D90:H90", S["hdr_l"])
floor = [
    (91, "Baixa do investimento (60% do valor justo da JO)", "=-D60", "lab_it", "num_it", "Valor contábil do investimento reconhecido no CPC 19."),
    (92, "Preço recebido", "=C75", "lab_it", "num_it", ""),
    (93, "(=) Resultado contábil da venda", "=C91+C92", "lab_b", "num_b", "Prejuízo contábil, porque o preço é inferior ao valor justo reconhecido."),
    (94, "IR diferido 'carimbado' devido na realização", "=D63", "lab_it", "num_it", "Devido integralmente: a subconta é baixada com a venda."),
    (95, "Crédito teórico sobre a perda (34%)", "=C93*$C$21", "lab_it", "num_it", "No lucro presumido a perda não é aproveitável; no lucro real, a compensação é limitada à trava de 30% e depende de lucros futuros."),
    (96, "Carga líquida teórica (se a perda fosse integralmente aproveitada)", "=C94+C95", "lab_it", "num_it", "Coincide com o RISCO 2 do bloco 4 (mesma base contábil)."),
    (97, "Carga adotada nos quadros (piso = imposto carimbado)", "=F83", "lab_b_top", "num_b_top", "Leitura conservadora adotada no material: o imposto carimbado é pago e a perda não gera economia."),
]
for r, lab, c, ls, ns, note in floor:
    put(f"B{r}", lab, S[ls]); put(f"C{r}", c, S[ns])
    st = S["cell_top"] if ls.endswith("top") else S["cell_it"]
    put(f"D{r}", note, st); merge(f"D{r}:H{r}", st); row_h[r] = 30 if (len(note) > 70 or len(lab) > 42) else 15

# ---- 4) Sensibilidade
redbar(99, "4) Sensibilidade do CPC 15 — e se o fisco não aceitar as camadas do custo?")
para(101, "O quadro testa o cenário CPC 15 sob duas hipóteses adversas. RISCO 2: o fisco não aceita a camada da mais-valia "
          "— o custo fiscal recua para a base contábil (custo original + reflexo do PL contábil da CNP). RISCO 1: o fisco não "
          "aceita nenhuma camada do reflexo — resta apenas o custo original remanescente (cenário A da aba 'Cálculos da "
          "Operação'). Cada degrau de risco tem valor fixo, igual a 34% da camada negada, independentemente do preço de venda.")
header(103, [("B", "Cenário de preço"), ("C", "Preço de venda"), ("D", "CPC 15 — custo integral (bloco 1)"),
             ("E", "RISCO 2 — fisco nega a mais-valia"), ("F", "RISCO 1 — fisco nega todo o reflexo"),
             ("G", "Δ RISCO 2 = 34% × mais-valia"), ("H", "Δ RISCO 1 = 34% × incremento em discussão")], height=48)
put("B104", "Custo fiscal considerado em cada hipótese", S["lab_bi"]); put("C104", "–", S["dash"])
put("D104", "=E38", S["num_bi"]); put("E104", "=E38-E37", S["num_bi"]); put("F104", "=E34", S["num_bi"])
put("G104", "=D104-E104", S["num_bi"]); put("H104", "=D104-F104", S["num_bi"]); row_h[104] = 16
for r, (rp, lab, _) in zip((105, 106, 107), prices):
    put(f"B{r}", lab, S["lab_it"]); put(f"C{r}", f"=C{rp}", S["num_b"])
    put(f"D{r}", f"=(C{r}-$D$104)*$C$21", S["num"]); put(f"E{r}", f"=(C{r}-$E$104)*$C$21", S["num"])
    put(f"F{r}", f"=(C{r}-$F$104)*$C$21", S["num"]); put(f"G{r}", f"=E{r}-D{r}", S["num_b"]); put(f"H{r}", f"=F{r}-D{r}", S["num_b"]); row_h[r] = 16
para(108, "Leitura: o RISCO 2 coincide com o CPC 19 nas vendas ao valor justo ou acima (mesma base contábil); abaixo do valor "
          "justo o CPC 19 continua pior, pelo piso do imposto carimbado. Já o RISCO 1 supera o CPC 19 em qualquer preço a partir "
          "do valor justo: se a tese de neutralidade do reflexo do PL contábil da CNP fosse integralmente afastada, o CPC 15 "
          "deixaria de ser o cenário mais favorável na venda futura — daí a importância de documentar a tese e o laudo (PPA).")

# ---- 5) Conclusão executiva
redbar(110, "5) Conclusão executiva — CPC 15 × CPC 19")
header(112, [("B", "Ponto de comparação"), ("C", "CPC 15"), ("D", "CPC 19"), ("E", "Δ (CPC 19 − CPC 15)")], height=32)
put("F112", "Leitura", S["hdr_l"]); merge("F112:H112", S["hdr_l"])
concl = [
    (113, "Custo do investimento das Holdings após o cash-out", "=E38", "=D60",
          "No CPC 19 o custo inclui 60% do AVJ de toda a JO; no CPC 15 apenas o reflexo da CNP a valor de mercado."),
    (114, "Parcela do custo com IR diferido (CPC 19) ou em discussão com o fisco (CPC 15)", "=C44+C45", "=D62",
          "CPC 15: reflexo cuja neutralidade depende de tese; CPC 19: AVJ carimbado, sem discussão."),
    (115, "IR potencial sobre essa parcela (34%)", "=E46", "=D63",
          "CPC 15: devido só se o fisco negar o custo e prevalecer; CPC 19: devido na venda, qualquer que seja o preço."),
    (116, "IR na venda pela Holding — preço abaixo do valor justo", "=D83", "=F83",
          "O piso do imposto carimbado torna o CPC 19 muito mais oneroso em vendas abaixo do valor justo."),
    (117, "IR na venda pela Holding — preço igual ao valor justo", "=D84", "=F84",
          "Diferença = 34% da mais-valia da CNP: nos preços a partir do valor justo os regimes se aproximam."),
    (118, "IR na venda pela Holding — preço acima do valor justo", "=D85", "=F85",
          "Mesma diferença fixa; o excedente ao valor justo é tributado a 34% nos dois regimes."),
    (119, "IR na venda pelas pessoas físicas — preço igual ao valor justo", "=E84", "=G84",
          "Só o CPC 15 permite capturar a alíquota de 22,5%; no CPC 19 o imposto carimbado fica na Holding."),
    (120, "CPC 15 em cenário adverso (RISCO 1) × CPC 19 — preço igual ao valor justo", "=F106", "=F84",
          "Se a tese de neutralidade do reflexo for integralmente afastada, o CPC 15 passa a ser mais oneroso que o CPC 19."),
]
for r, lab, c, d, note in concl:
    put(f"B{r}", lab, S["lab_it"]); put(f"C{r}", c, S["num"]); put(f"D{r}", d, S["num"]); put(f"E{r}", f"=D{r}-C{r}", S["num_b"])
    put(f"F{r}", note, S["cell_it"]); merge(f"F{r}:H{r}", S["cell_it"]); row_h[r] = 30
para(122, "• CPC 15 é o cenário mais favorável às Holdings: o custo do investimento incorpora o reflexo da incorporação sem "
          "registro de passivo fiscal diferido, o imposto na venda futura acompanha o preço e existe a alternativa de venda "
          "pelas pessoas físicas à alíquota de 22,5%. O risco concentra-se na neutralidade do reflexo do PL contábil da CNP "
          "(bloco 1); com as premissas atuais a mais-valia é pequena e o imposto potencial sobre ela é residual.", "para_reg")
para(123, "• CPC 19 elimina a discussão sobre o custo, mas ao preço de reconhecer ganho de AVJ sobre toda a JO — inclusive a "
          "Embracon, que no CPC 15 não é reavaliada — com IRPJ/CSLL de 34% carimbado e devido na venda em qualquer cenário de "
          "preço, sem a alternativa da PF. Além disso, o valor patrimonial das Holdings sobe ao valor justo, com impacto direto "
          "na base de cálculo do ITCMD (tema não tratado neste material).", "para_reg")
para(124, "• Na venda ao valor justo ou acima, a carga dos dois regimes é próxima (diferença = 34% da mais-valia da CNP); abaixo "
          "do valor justo, o CPC 19 é substancialmente mais oneroso. Se a tese de neutralidade do reflexo fosse integralmente "
          "afastada (RISCO 1), o CPC 15 passaria a ser mais oneroso que o CPC 19 nas vendas a partir do valor justo — daí a "
          "importância de documentar a tese e o laudo de avaliação (PPA).", "para_reg")
para(126, "Ressalvas: valores ilustrativos em R$ milhões, data-base 30.06.2026, com as premissas da aba 'Cálculos da Operação' "
          "e valor justo da Embracon pelo EV do Acordo de Investimento. As Holdings estão no lucro presumido; na venda futura, "
          "a carga efetiva na PJ poderia ser reduzida com a opção pelo lucro real e a existência de prejuízos compensáveis. O "
          "custo das quotas nas PFs é premissa ilustrativa a confirmar. O material não avalia o ITCMD nem o cash-out no "
          "CPC 19, tratado na aba 'Cálculos da Operação' com o custo original (cenário A) e o incrementado (cenário B).")

# --------------------------------------------------------------------------
# 3. serializar sheet3.xml
# --------------------------------------------------------------------------
col_widths = {1: 7.54296875, 2: 46, 3: 13, 4: 13, 5: 13, 6: 13, 7: 13, 8: 48, 9: 9, 10: 9, 11: 9}
def cell_xml(r, c, d):
    ref = f"{COLS[c-1]}{r}"
    s = f' s="{d["style"]}"' if d["style"] is not None else ""
    k, v = d["kind"], d["value"]
    if k == "blank": return f'<c r="{ref}"{s}/>'
    if k == "s": return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{escape(v)}</t></is></c>'
    if k == "f": return f'<c r="{ref}"{s}><f>{escape(v[1:])}</f></c>'
    return f'<c r="{ref}"{s}><v>{repr(float(v)) if isinstance(v, float) else v}</v></c>'

rows = sorted({r for r, _ in cells})
max_row = max(rows)
parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
         f'<worksheet xmlns="{NS}" xmlns:r="{NSR}">',
         '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>',
         f'<dimension ref="B4:K{max_row}"/>',
         '<sheetViews><sheetView showGridLines="0" zoomScale="90" zoomScaleNormal="90" workbookViewId="0"><selection activeCell="B8" sqref="B8"/></sheetView></sheetViews>',
         '<sheetFormatPr defaultColWidth="8.90625" defaultRowHeight="14.5"/>',
         '<cols>' + ''.join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in col_widths.items()) + '</cols>',
         '<sheetData>']
for r in rows:
    ht = f' ht="{row_h[r]}" customHeight="1"' if r in row_h else ""
    parts.append(f'<row r="{r}"{ht}>' + ''.join(cell_xml(r, c, cells[(r, c)]) for c in sorted(c for rr, c in cells if rr == r)) + '</row>')
parts.append('</sheetData>')
parts.append(f'<mergeCells count="{len(merges)}">' + ''.join(f'<mergeCell ref="{m}"/>' for m in merges) + '</mergeCells>')
parts.append('<pageMargins left="0.511811024" right="0.511811024" top="0.78740157499999996" bottom="0.78740157499999996" header="0.31496062000000002" footer="0.31496062000000002"/>')
parts.append('<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>')
parts.append('<drawing r:id="rId1"/>')
parts.append('</worksheet>')
files["xl/worksheets/sheet3.xml"] = ''.join(parts).encode("utf-8")

# --------------------------------------------------------------------------
# 4. calcChain fora, fullCalcOnLoad ligado
# --------------------------------------------------------------------------
files.pop("xl/calcChain.xml", None)
ct = files["[Content_Types].xml"].decode("utf-8")
ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', '', ct)
files["[Content_Types].xml"] = ct.encode("utf-8")
rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
rels = re.sub(r'<Relationship Id="[^"]+" Type="[^"]+/calcChain" Target="calcChain.xml"/>', '', rels)
files["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
wbx = files["xl/workbook.xml"].decode("utf-8")
wbx = re.sub(r'<calcPr calcId="(\d+)"/>', r'<calcPr calcId="\1" fullCalcOnLoad="1"/>', wbx)
files["xl/workbook.xml"] = wbx.encode("utf-8")

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    order = ["[Content_Types].xml"] + [n for n in zin.namelist() if n != "[Content_Types].xml" and n in files]
    for n in order:
        z.writestr(n, files[n])
print("saved", OUT, "| células:", len(cells), "| fórmulas:", sum(1 for d in cells.values() if d["kind"] == "f"),
      "| xfs:", len(cellxfs), "| numFmts:", NF_NUM, NF_CHK)
