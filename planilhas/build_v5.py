#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v5 da planilha "Embracon_Simulacao_Tributaria_04.09.26" (base: v4 editada pelo usuário).
Cirurgia direta no pacote OOXML — preserva logo EMF, abas ocultas, sharedStrings e a
formatação existente. Mudanças:
  (iv) nova aba "Premissas" (após Organograma) — inputs centralizados (cinza) + fatos + ressalvas
       do material SF de 31.08; as células de premissa da aba "Cálculos da Operação" passam a
       apontar para ela (valores idênticos);
  (i)  nova aba "CPC 19" (após Cálculos) — bloco CPC 19 transferido/ampliado (lançamentos);
  (ii) aba "Alienação Futura (PJ x PF)" reconstruída só com premissas da venda, venda PJ×PF,
       cenário alternativo, análise hipotética e conclusão;
  (iii) nova aba "Conclusão" (após Alienação) — quadros do PPT/MLA (custo do investimento e
       venda futura) + conclusão executiva;
  (v)  logo + título em todas as abas com a mesma âncora/estilo da aba "Cálculos da Operação";
  calcChain removido e fullCalcOnLoad ligado.
Uso: python3 build_v5.py <base.xlsx> <saida.xlsx>
"""
import sys, re, zipfile, math
from xml.sax.saxutils import escape
from lxml import etree

SRC = sys.argv[1]; OUT = sys.argv[2]
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NSMAP = {"m": NS}
CALC = "'Cálculos da Operação'"; PREM = "Premissas"; CPC19 = "'CPC 19'"; ALI = "'Alienação Futura (PJ x PF)'"

zin = zipfile.ZipFile(SRC)
files = {n: zin.read(n) for n in zin.namelist()}

# ============================================================ styles
styles = etree.fromstring(files["xl/styles.xml"])
fonts = styles.find("m:fonts", NSMAP); fills = styles.find("m:fills", NSMAP)
borders = styles.find("m:borders", NSMAP); numfmts = styles.find("m:numFmts", NSMAP)
cellxfs = styles.find("m:cellXfs", NSMAP)

def font_sig(f):
    d = {"b": False, "i": False, "sz": None, "color": None, "name": None}
    for ch in f:
        t = etree.QName(ch).localname
        if t == "b": d["b"] = True
        elif t == "i": d["i"] = True
        elif t == "sz": d["sz"] = ch.get("val")
        elif t == "color": d["color"] = ch.get("rgb") or ("theme" + str(ch.get("theme")))
        elif t == "name": d["name"] = ch.get("val")
    return d
def find_font(b=False, i=False, sz="11", color="theme1", name="Aptos Narrow"):
    for idx, f in enumerate(fonts):
        s = font_sig(f)
        if s["b"] == b and s["i"] == i and s["sz"] == sz and s["color"] == color and s["name"] == name:
            return idx
    f = etree.SubElement(fonts, "{%s}font" % NS)
    if b: etree.SubElement(f, "{%s}b" % NS)
    if i: etree.SubElement(f, "{%s}i" % NS)
    etree.SubElement(f, "{%s}sz" % NS).set("val", sz)
    c = etree.SubElement(f, "{%s}color" % NS)
    if color.startswith("theme"): c.set("theme", color[5:])
    else: c.set("rgb", color)
    etree.SubElement(f, "{%s}name" % NS).set("val", name)
    etree.SubElement(f, "{%s}family" % NS).set("val", "2")
    etree.SubElement(f, "{%s}scheme" % NS).set("val", "minor")
    fonts.set("count", str(len(fonts)))
    return len(fonts) - 1

F_REG = find_font(); F_IT = find_font(i=True); F_B = find_font(b=True); F_BI = find_font(b=True, i=True)
F_IT20 = find_font(i=True, sz="20"); F_B14W = find_font(b=True, sz="14", color="theme0")
F_IT14 = find_font(i=True, sz="14"); F_BI14 = find_font(b=True, i=True, sz="14")
F_BI_RED = find_font(b=True, i=True, color="FFC00000"); F_B_RED = find_font(b=True, color="FFC00000")
F_B_BLUE = find_font(b=True, color="FF1F3864"); F_B_GREEN = find_font(b=True, color="FF00703C")
F_IT_RED = find_font(i=True, color="FFFF0000")

def fill_sig(f):
    p = f.find("m:patternFill", NSMAP)
    if p is None or p.get("patternType") in (None, "none"): return ("none", None, None, None)
    fg = p.find("m:fgColor", NSMAP)
    if fg is None: return (p.get("patternType"), None, None, None)
    return (p.get("patternType"), fg.get("rgb"), fg.get("theme"), fg.get("tint"))
FILL_NONE = 0
FILL_GRAY = next(i for i, f in enumerate(fills) if fill_sig(f)[0] == "solid" and fill_sig(f)[2] == "0" and fill_sig(f)[3] and fill_sig(f)[3].startswith("-4.99"))
FILL_RED = next(i for i, f in enumerate(fills) if fill_sig(f)[1] == "FFC00000")
FILL_YEL = next(i for i, f in enumerate(fills) if fill_sig(f)[1] == "FFFFFF00")
FILL_WHITE = next(i for i, f in enumerate(fills) if fill_sig(f)[0] == "solid" and fill_sig(f)[2] == "0" and not fill_sig(f)[3])

def border_sig(b):
    return tuple((b.find("m:" + s, NSMAP).get("style") if b.find("m:" + s, NSMAP) is not None else None) for s in ("left", "right", "top", "bottom"))
border_idx = {}
for i, b in enumerate(borders): border_idx.setdefault(border_sig(b), i)
def get_border(left=None, right=None, top=None, bottom=None):
    sig = (left, right, top, bottom)
    if sig in border_idx: return border_idx[sig]
    b = etree.SubElement(borders, "{%s}border" % NS)
    for side, st in zip(("left", "right", "top", "bottom"), sig):
        e = etree.SubElement(b, "{%s}%s" % (NS, side))
        if st: e.set("style", st); etree.SubElement(e, "{%s}color" % NS).set("indexed", "64")
    etree.SubElement(b, "{%s}diagonal" % NS)
    borders.set("count", str(len(borders))); border_idx[sig] = len(borders) - 1
    return border_idx[sig]
B_NONE = get_border(); B_BOT = get_border(bottom="thin"); B_TOP = get_border(top="thin")
B_TB = get_border(top="thin", bottom="thin"); B_L = get_border(left="thin"); B_R = get_border(right="thin")
B_LR = get_border(left="thin", right="thin"); B_BL = get_border(left="thin", bottom="thin"); B_BR = get_border(right="thin", bottom="thin")
B_BLR = get_border(left="thin", right="thin", bottom="thin"); B_TBLR = get_border("thin", "thin", "thin", "thin")
B_TBL = get_border(left="thin", top="thin", bottom="thin"); B_TBR = get_border(right="thin", top="thin", bottom="thin")

existing_ids = {int(n.get("numFmtId")) for n in numfmts}
def add_numfmt(code):
    for n in numfmts:
        if n.get("formatCode") == code: return int(n.get("numFmtId"))
    nid = max(existing_ids | {176}) + 1; existing_ids.add(nid)
    e = etree.SubElement(numfmts, "{%s}numFmt" % NS); e.set("numFmtId", str(nid)); e.set("formatCode", code)
    numfmts.set("count", str(len(numfmts))); return nid
NF_NUM = add_numfmt('#,##0.0;[Red]\\(#,##0.0\\);\\-')
NF_INT = add_numfmt('#,##0;[Red]\\(#,##0\\);\\-')
NF_CHK = add_numfmt('0.0;[Red]\\-0.0;"ok"')
NF_DELTA = add_numfmt('\\+#,##0;[Red]\\-#,##0;\\-')
NF_PCT = add_numfmt('0.0%'); NF_PCT0 = 9; NF_GEN = 0; NF_INT0 = 3

xf_cache = {}
def xf(font=F_REG, fill=FILL_NONE, border=B_NONE, nf=NF_GEN, h=None, v=None, wrap=False, indent=0):
    key = (font, fill, border, nf, h, v, wrap, indent)
    if key in xf_cache: return xf_cache[key]
    e = etree.SubElement(cellxfs, "{%s}xf" % NS)
    e.set("numFmtId", str(nf)); e.set("fontId", str(font)); e.set("fillId", str(fill)); e.set("borderId", str(border)); e.set("xfId", "0")
    if nf: e.set("applyNumberFormat", "1")
    e.set("applyFont", "1")
    if fill: e.set("applyFill", "1")
    if border: e.set("applyBorder", "1")
    if h or v or wrap or indent:
        e.set("applyAlignment", "1"); a = etree.SubElement(e, "{%s}alignment" % NS)
        if h: a.set("horizontal", h)
        if v: a.set("vertical", v)
        if wrap: a.set("wrapText", "1")
        if indent: a.set("indent", str(indent))
    cellxfs.set("count", str(len(cellxfs))); xf_cache[key] = len(cellxfs) - 1
    return xf_cache[key]

S = {}
S["title"]     = xf(F_IT20, h="left", v="bottom")
S["redbar"]    = xf(F_B14W, fill=FILL_RED, border=B_TB, h="left", v="center")
S["h14"]       = xf(F_BI14, border=B_BOT, h="left", v="center")          # sub-cabeçalho itálico 14 (padrão do usuário)
S["h14_red"]   = xf(find_font(b=True, i=True, sz="14", color="FFC00000"), border=B_BOT, h="left", v="center")
S["h14_blue"]  = xf(find_font(b=True, i=True, sz="14", color="FF1F3864"), border=B_BOT, h="left", v="center")
S["sub"]       = xf(F_B, border=B_BOT, h="center", v="center")            # título de tabela centralizado (padrão do usuário)
S["sub_l"]     = xf(F_B, border=B_BOT, h="left", v="center")
S["para"]      = xf(F_IT, h="left", v="top", wrap=True)
S["para_reg"]  = xf(F_REG, h="left", v="top", wrap=True)
S["lab_it"]    = xf(F_IT, h="left", v="center", wrap=True)
S["lab_reg"]   = xf(F_REG, h="left", v="center", wrap=True)
S["lab_b"]     = xf(F_B, h="left", v="center", wrap=True)
S["lab_bi"]    = xf(F_BI, h="left", v="center", wrap=True)
S["lab_b_top"] = xf(F_B, border=B_TOP, h="left", v="center", wrap=True)
S["lab_bi_top"]= xf(F_BI, border=B_TOP, h="left", v="center", wrap=True)
S["lab_it_bot"]= xf(F_IT, border=B_BOT, h="left", v="center", wrap=True)
S["note"]      = xf(F_IT, h="left", v="center", wrap=True)
S["note_top"]  = xf(F_IT, border=B_TOP, h="left", v="center", wrap=True)
S["hdr_l"]     = xf(F_B, fill=FILL_GRAY, border=B_BOT, h="left", v="center", wrap=True)
S["hdr_c"]     = xf(F_B, fill=FILL_GRAY, border=B_BOT, h="center", v="center", wrap=True)
S["hdr_c_bl"]  = xf(F_B, fill=FILL_GRAY, border=B_BL, h="center", v="center", wrap=True)
S["hdr_c_br"]  = xf(F_B, fill=FILL_GRAY, border=B_BR, h="center", v="center", wrap=True)
S["hdr_c_blr"] = xf(F_B, fill=FILL_GRAY, border=B_BLR, h="center", v="center", wrap=True)
S["hdr_red"]   = xf(find_font(b=True, color="theme0"), fill=FILL_RED, border=B_TBLR, h="center", v="center", wrap=True)
S["hdr_red_l"] = xf(find_font(b=True, color="theme0"), fill=FILL_RED, border=B_TBLR, h="left", v="center", wrap=True)
S["num"]       = xf(F_REG, nf=NF_NUM, h="center", v="center")
S["num_it"]    = xf(F_IT, nf=NF_NUM, h="center", v="center")
S["num_b"]     = xf(F_B, nf=NF_NUM, h="center", v="center")
S["num_bi"]    = xf(F_BI, nf=NF_NUM, h="center", v="center")
S["num_b_top"] = xf(F_B, border=B_TOP, nf=NF_NUM, h="center", v="center")
S["num_bi_red"]= xf(F_BI_RED, fill=FILL_GRAY, nf=NF_NUM, h="center", v="center")
S["int"]       = xf(F_REG, nf=NF_INT, h="center", v="center")
S["int_it"]    = xf(F_IT, nf=NF_INT, h="center", v="center")
S["int_b"]     = xf(F_B, nf=NF_INT, h="center", v="center")
S["int_bi"]    = xf(F_BI, nf=NF_INT, h="center", v="center")
S["int_b_top"] = xf(F_B, border=B_TOP, nf=NF_INT, h="center", v="center")
S["int_l"]     = xf(F_REG, border=B_L, nf=NF_INT, h="center", v="center")
S["int_r"]     = xf(F_REG, border=B_R, nf=NF_INT, h="center", v="center")
S["int_bl"]    = xf(F_REG, border=B_BL, nf=NF_INT, h="center", v="center")
S["int_br"]    = xf(F_REG, border=B_BR, nf=NF_INT, h="center", v="center")
S["int_b_lr"]  = xf(F_B, border=B_LR, nf=NF_INT, h="center", v="center")
S["int_b_blr"] = xf(F_B, border=B_BLR, nf=NF_INT, h="center", v="center")
S["int_gray_b"]= xf(F_B, fill=FILL_GRAY, border=B_TBLR, nf=NF_INT, h="center", v="center")
S["int_dark"]  = xf(find_font(b=True, color="theme0"), fill=find_font and next(i for i, f in enumerate(fills) if fill_sig(f)[0] == "solid" and fill_sig(f)[2] == "0" and fill_sig(f)[3] and fill_sig(f)[3].startswith("-0.3499")), border=B_TBLR, nf=NF_INT, h="center", v="center")
S["delta"]     = xf(F_B, nf=NF_DELTA, h="center", v="center")
S["delta_top"] = xf(F_B, border=B_TOP, nf=NF_DELTA, h="center", v="center")
S["pct"]       = xf(F_REG, nf=NF_PCT, h="center", v="center")
S["pct_it"]    = xf(F_IT, nf=NF_PCT, h="center", v="center")
S["pct_b"]     = xf(F_B, nf=NF_PCT, h="center", v="center")
S["txt_c_it"]  = xf(F_IT, h="center", v="center")
S["txt_c_b"]   = xf(F_B, h="center", v="center")
S["dash"]      = xf(F_IT, h="center", v="center")
S["dash_top"]  = xf(F_IT, border=B_TOP, h="center", v="center")
S["chk"]       = xf(F_IT, nf=NF_CHK, h="center", v="center")
S["inp_lab"]   = xf(F_B, fill=FILL_GRAY, h="left", v="center", wrap=True)
S["inp_int"]   = xf(F_B, fill=FILL_GRAY, nf=NF_INT0, h="center", v="center")
S["inp_num"]   = xf(F_B, fill=FILL_GRAY, nf=NF_NUM, h="center", v="center")
S["inp_pct"]   = xf(F_B, fill=FILL_GRAY, nf=NF_PCT, h="center", v="center")
S["inp_txt"]   = xf(F_B, fill=FILL_GRAY, h="center", v="center")
S["inp_lab_y"] = xf(F_B, fill=FILL_YEL, h="left", v="center", wrap=True)
S["inp_int_y"] = xf(F_B, fill=FILL_YEL, nf=NF_INT0, h="center", v="center")
S["cell_it"]   = xf(F_IT, h="left", v="center", wrap=True)
S["cell_top"]  = xf(F_IT, border=B_TOP, h="left", v="center", wrap=True)
S["cell_l"]    = xf(F_IT, border=B_L, h="left", v="center", wrap=True)
S["cell_reg"]  = xf(F_REG, h="left", v="center", wrap=True)
S["cell_reg_top"] = xf(F_REG, border=B_TOP, h="left", v="center", wrap=True)
S["lab_reg_l"] = xf(F_REG, h="left", v="center")
S["bul_lab"]   = xf(F_B, h="left", v="center", wrap=True)
S["hdr_dark_l"]= xf(find_font(b=True, color="theme0"), fill=next(i for i, f in enumerate(fills) if fill_sig(f)[0] == "solid" and fill_sig(f)[2] == "0" and fill_sig(f)[3] and fill_sig(f)[3].startswith("-0.3499")), border=B_TBLR, h="left", v="center", wrap=True)
S["gray_lab"]  = xf(F_B, fill=FILL_GRAY, border=B_TBLR, h="left", v="center", wrap=True)
S["gray_lab_it"] = xf(F_IT, fill=FILL_GRAY, border=B_TBLR, h="left", v="center", wrap=True)

# ============================================================ sheet model
COLS = [chr(ord("A") + i) for i in range(17)]  # A..Q
def col_i(c): return COLS.index(c) + 1

class Sheet:
    def __init__(self, name, widths):
        self.name = name; self.widths = widths; self.cells = {}; self.merges = []; self.row_h = {}
    def put(self, ref, value=None, style=None, kind=None):
        col = re.match(r"[A-Z]+", ref).group(0); row = int(ref[len(col):])
        if kind is None:
            if value is None: kind = "blank"
            elif isinstance(value, list): kind = "rich"
            elif isinstance(value, str) and value.startswith("="): kind = "f"
            elif isinstance(value, str): kind = "s"
            else: kind = "n"
        self.cells[(row, col_i(col))] = dict(kind=kind, value=value, style=style)
    def merge(self, rng, style=None):
        self.merges.append(rng); a, b = rng.split(":")
        ca, ra = re.match(r"([A-Z]+)(\d+)", a).groups(); cb, rb = re.match(r"([A-Z]+)(\d+)", b).groups()
        for r in range(int(ra), int(rb) + 1):
            for c in range(col_i(ca), col_i(cb) + 1):
                if (r, c) not in self.cells: self.cells[(r, c)] = dict(kind="blank", value=None, style=style)
                elif style is not None and self.cells[(r, c)]["style"] is None: self.cells[(r, c)]["style"] = style
    def width_units(self, a, b):
        return sum(self.widths.get(col_i(c), 8.9) for c in COLS[col_i(a) - 1:col_i(b)])
    def para(self, row, text, style="para", cols="B:K", cpl=None):
        a, b = cols.split(":")
        self.put(f"{a}{row}", text, S[style]); self.merge(f"{a}{row}:{b}{row}", S[style])
        cpl = cpl or max(40, int(self.width_units(a, b) * 1.15))
        lines = max(1, math.ceil(len(text) / cpl)); self.row_h[row] = 15.0 * lines + 4
    def text_h(self, row, text, cols):
        a, b = cols.split(":"); cpl = max(20, int(self.width_units(a, b) * 1.15))
        lines = max(1, math.ceil(len(text) / cpl)); self.row_h[row] = max(self.row_h.get(row, 0), 15.0 * lines + 2)
    def redbar(self, row, text):
        self.put(f"B{row}", text, S["redbar"]); self.merge(f"B{row}:K{row}", S["redbar"]); self.row_h[row] = 18.5
    def h14(self, row, text, style="h14", cols="B:K"):
        a, b = cols.split(":"); self.put(f"{a}{row}", text, S[style]); self.merge(f"{a}{row}:{b}{row}", S[style]); self.row_h[row] = 19
    def sub(self, row, text, cols="B:E", style="sub"):
        a, b = cols.split(":"); self.put(f"{a}{row}", text, S[style]); self.merge(f"{a}{row}:{b}{row}", S[style]); self.row_h[row] = 16
    def header(self, row, items, height=16):
        for col, text, st in items: self.put(f"{col}{row}", text, S[st])
        self.row_h[row] = height
    def head(self, title, date="4 de setembro de 2026", title_col="C"):
        self.put(f"{title_col}4", title, S["title"]); self.row_h[4] = 26
        self.put(f"{title_col}5", date, xf(F_IT, h="left", v="top")); self.row_h[5] = 44
    def xml(self, rid_drawing="rId1", tab_selected=False):
        rows = sorted({r for r, _ in self.cells}); max_row = max(rows)
        def cell_xml(r, c, d):
            ref = f"{COLS[c-1]}{r}"; s = f' s="{d["style"]}"' if d["style"] is not None else ""; k, v = d["kind"], d["value"]
            if k == "blank": return f'<c r="{ref}"{s}/>'
            if k == "s": return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{escape(v)}</t></is></c>'
            if k == "rich":
                runs = ""
                for text, color, bold in v:
                    col_xml = ('<color rgb="%s"/>' % color) if color else '<color theme="1"/>'
                    rpr = '<rPr>' + ('<b/>' if bold else '') + '<sz val="11"/>' + col_xml + '<rFont val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></rPr>'
                    runs += f'<r>{rpr}<t xml:space="preserve">{escape(text)}</t></r>'
                return f'<c r="{ref}"{s} t="inlineStr"><is>{runs}</is></c>'
            if k == "f": return f'<c r="{ref}"{s}><f>{escape(v[1:])}</f></c>'
            return f'<c r="{ref}"{s}><v>{repr(float(v)) if isinstance(v, float) else v}</v></c>'
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n', f'<worksheet xmlns="{NS}" xmlns:r="{NSR}">',
                 '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>', f'<dimension ref="B3:Q{max_row}"/>',
                 '<sheetViews><sheetView showGridLines="0"' + (' tabSelected="1"' if tab_selected else '') + ' zoomScale="90" zoomScaleNormal="90" workbookViewId="0"><selection activeCell="B6" sqref="B6"/></sheetView></sheetViews>',
                 '<sheetFormatPr defaultColWidth="8.90625" defaultRowHeight="14.5"/>',
                 '<cols>' + ''.join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in sorted(self.widths.items())) + '</cols>', '<sheetData>']
        for r in rows:
            ht = f' ht="{self.row_h[r]}" customHeight="1"' if r in self.row_h else ""
            parts.append(f'<row r="{r}"{ht}>' + ''.join(cell_xml(r, c, self.cells[(r, c)]) for c in sorted(c for rr, c in self.cells if rr == r)) + '</row>')
        parts.append('</sheetData>')
        if self.merges: parts.append(f'<mergeCells count="{len(self.merges)}">' + ''.join(f'<mergeCell ref="{m}"/>' for m in self.merges) + '</mergeCells>')
        parts.append('<pageMargins left="0.511811024" right="0.511811024" top="0.78740157499999996" bottom="0.78740157499999996" header="0.31496062000000002" footer="0.31496062000000002"/>')
        parts.append('<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>')
        parts.append(f'<drawing r:id="{rid_drawing}"/>'); parts.append('</worksheet>')
        return ''.join(parts).encode("utf-8")

W_WIDE = {1: 7.54296875, 2: 46.1796875, 3: 21.7265625, 4: 16.54296875, 5: 18, 6: 16.26953125, 7: 11.54296875, 8: 10.54296875, 9: 9, 10: 9, 11: 9}

def bullet(color, text):  # rich text: quadrado colorido + texto
    return [("■ ", color, True), (text, None, False)]
GREEN, AMBER, RED = "FF00B050", "FFFFC000", "FFFF0000"

# ============================================================ PREMISSAS
P = Sheet("Premissas", W_WIDE); P.head("Premissas e Fatos da Operação")
P.redbar(6, "Fatos da operação (Acordo de Investimento, Ofício BACEN 31070/2026, calls Deloitte/SF e material SF de 31.08.2026)")
facts = [
    "• Incorporação: a Embracon incorpora a CNP (incorporação plena, aprovada pelo BACEN); o acervo da CNP entra pelo valor de mercado apurado em laudo (PPA) e as Partes CNP recebem 13,8% do capital da Embracon (relação de troca).",
    "• Cash-out simultâneo: as Partes CNP adquirem das Holdings 26,2% da Embracon (13,1% de cada uma) por R$ 1.200 milhões, desconsiderando correção pelo IPCA, ajustes de caixa e eventual earn-out de até R$ 680 milhões. A secundária é entendida como concomitante à incorporação (operação única).",
    "• Posição final: famílias 60% (Savian 30% e JVFJ 30%) e Grupo CNP 40%.",
    "• Qualificação contábil em discussão: combinação de negócios (CPC 15 — leitura do contrato e da Stocche Forbes; o PL da Embracon não é reavaliado) ou operação em conjunto (CPC 19 — leitura da Deloitte a partir dos vetos qualificados da CNP; Embracon e CNP reavaliadas a valor justo, com passivo fiscal diferido).",
    "• Objetivo do material: ilustrar, para Savian e JVFJ, os impactos de cada qualificação sobre o custo do investimento, o imposto no cash-out e o imposto na venda futura dos 60% (pela Holding ou pelas pessoas físicas).",
]
r = 8
for t in facts: P.para(r, t, "para_reg"); r += 1
P.redbar(14, "Premissas numéricas — inputs do modelo (células cinza; alterações aqui propagam para todas as abas)")
P.header(16, [("B", "Premissa", "hdr_l"), ("C", "Valor", "hdr_c"), ("D", "Origem / observação", "hdr_l")]); P.merge("D16:K16", S["hdr_l"])
prem = [
    # row, label, value, style, note, registry key
    (17, "Data-base dos balanços", "30.06.2026", "inp_txt", "Balanços esboçados pela Embracon (agosto/2026), não definitivos", "data"),
    (18, "Unidade monetária", "R$ milhões", "inp_txt", "", "unid"),
    (19, "PL contábil da Embracon (antes da incorporação)", 254, "inp_int", "Corresponde ao custo original do investimento das Holdings (Savian + JVFJ)", "pl_e"),
    (20, "Investimento original — Savian (50% do PL da Embracon)", "=C19/2", "int_it", "Derivado: as Holdings são simétricas no modelo (50/50); a aba 'Cálculos da Operação' dobra a coluna Savian para obter os totais", "inv_s"),
    (21, "Investimento original — JVFJ (50% do PL da Embracon)", "=C20", "int_it", "Derivado: igual ao custo da Savian (simetria das Holdings)", "inv_j"),
    (22, "PL contábil da CNP", 800, "inp_int", "Esboço da Embracon; o material SF de 31.08 usava 839", "pl_c"),
    (23, "Valor de mercado da CNP (valor incorporado no CPC 15 — PPA)", 822, "inp_int", "≈ EV da CNP no Acordo de Investimento (R$ 821,9 mi, data-base 31.12.2024); a diferença para o PL contábil é a mais-valia", "vm_c"),
    (24, "Valor justo da Embracon (EV — Acordo de Investimento, data-base 31.12.2024)", 5133, "inp_int", "Usado apenas no cenário CPC 19 (reavaliação da Embracon); o material SF de 31.08 usava o EV total DTT de 5.588", "vj_e"),
    (25, "Participação das Partes CNP na incorporação (relação de troca)", 0.138, "inp_pct", "Ações novas emitidas pela Embracon às Partes CNP", "p_inc"),
    (26, "Participação-alvo final das Partes CNP", 0.40, "inp_pct", "Após o cash-out", "p_cnp"),
    (27, "Preço do cash-out (compra de 26,2% das Holdings)", 1200, "inp_int", "Sem IPCA, ajustes de caixa e earn-out (até 680)", "cashout"),
    (28, "IRPJ/CSLL — ganho de capital das Holdings", 0.34, "inp_pct", "Alíquota combinada", "ir"),
    (29, "IR — ganho de capital das pessoas físicas (simplificada)", 0.225, "inp_pct", "Adotada a alíquota máxima da tabela progressiva, por simplificação", "ir_pf"),
    (30, "Custo das quotas das Holdings nas pessoas físicas", 18, "inp_int_y", "Premissa ilustrativa (exercício MLA), a confirmar com o custo real das quotas", "custo_pf"),
    (31, "Preço de venda futura — abaixo do valor justo", 1000, "inp_int", "Cenário (i) do material SF", "p_low"),
    (32, "Preço de venda futura — igual ao valor justo (60% do valor justo da JO)", f"={CPC19}!D22", "num_b", "Cenário (ii): calculado na aba CPC 19 (quadro 2, valor justo total × 60%)", "p_vj"),
    (33, "Preço de venda futura — acima do valor justo", 5000, "inp_int", "Cenário (iii) do material SF", "p_high"),
]
PR = {}
for row, lab, val, st, note, key in prem:
    P.put(f"B{row}", lab, S["inp_lab_y"] if st.endswith("_y") else (S["inp_lab"] if st.startswith("inp") else S["lab_it"]))
    P.put(f"C{row}", val, S["num_it"] if st == "int_it" else S[st]); PR[key] = f"C{row}"
    if note: P.put(f"D{row}", note, S["note"]); P.merge(f"D{row}:K{row}", S["note"])
    P.row_h[row] = 30 if (len(lab) > 44 or len(note) > 95) else 15
P.sub(35, "Valores derivados das premissas", cols="B:C", style="sub_l")
derived = [
    (36, "Participação final das famílias (Savian + JVFJ)", f"=1-{PR['p_cnp']}", "pct_it", "fam"),
    (37, "Mais-valia da CNP (valor de mercado − PL contábil)", f"={PR['vm_c']}-{PR['pl_c']}", "num_it", "mv"),
    (38, "Percentual vendido no cash-out (alvo CNP − relação de troca)", f"={PR['p_cnp']}-{PR['p_inc']}", "pct_it", "p_venda"),
    (39, "Fração do investimento de cada Holding alienada no cash-out (13,1% ÷ 43,1%)", f"={CALC}!D71", "pct_it", "fr"),
]
for row, lab, f, st, key in derived:
    P.put(f"B{row}", lab, S["lab_it"]); P.put(f"C{row}", f, S[st]); PR[key] = f"C{row}"; P.row_h[row] = 30 if len(lab) > 44 else 15
P.put("D39", "Critério de alocação de custo da aba 'Cálculos da Operação'; o material SF de 31.08 alocava 26,2% do custo (13,1% por Holding)", S["note"]); P.merge("D39:K39", S["note"]); P.row_h[39] = 30
P.redbar(41, "Premissas fiscais e ressalvas (material SF de 31.08.2026 e critérios adotados neste modelo)")
ress = [
    "• Custo alocado no cash-out (critério deste modelo, herdado da aba 'Cálculos da Operação'; o material SF de 31.08 alocava 26,2% do custo, 13,1% de cada Holding): a secundária é tratada como concomitante à incorporação e o custo de cada Holding é alocado proporcionalmente à fração alienada (13,1% ÷ 43,1%). O imposto do cash-out é apurado na aba 'Cálculos da Operação' nos cenários A (custo original) e B (custo incrementado).",
    "• Regime das Holdings: lucro presumido — premissa relevante para a tributação do cash-out. Na futura alienação dos 60% remanescentes, a carga poderá ser reduzida se, à época, as Holdings estiverem no lucro real e dispuserem de prejuízos fiscais e bases negativas de CSLL compensáveis, observadas as limitações aplicáveis.",
    "• Segregação de risco no CPC 15: o custo das Holdings é aberto em custo original (incontroverso), reflexo do PL contábil da CNP (neutralidade do MEP — art. 33, §2º, do DL 1.598/77; precedente WTorre, com decisões CARF de 2024 em sentido contrário — fundamentos discutidos nas calls com a SF) e reflexo da mais-valia (parcela mais exposta). No material de 31.08 a mais-valia era estimada em 34,6% do PL da CNP (proporção DTT 300/866); aqui é o valor de mercado menos o PL contábil.",
    "• CPC 19: o ganho de avaliação a valor justo (AVJ) tem tributação diferida, controlado em subconta (arts. 13 e 14 da Lei 12.973/14), fica 'carimbado' a 34% e é devido na realização do investimento, independentemente da entidade que aliene a participação.",
    "• ITCMD: não analisado neste material. Em caráter preliminar, o cenário de operação em conjunto pode elevar o valor patrimonial das Holdings (reconhecimento a valor justo da CNP e da Embracon), com possível repercussão na base do imposto, a depender do critério legal aplicável.",
    "• Caráter ilustrativo: os exemplos foram elaborados com informações estimadas, de diversas fontes, não auditadas; não substituem a análise dos documentos definitivos da Transação e da contabilização efetivamente adotada, nem devem ser entendidos como estimativa de custo ou contingência tributária.",
]
r = 43
for t in ress: P.para(r, t, "para_reg"); r += 1

# ============================================================ CPC 19
C = Sheet("CPC 19", W_WIDE); C.head("CPC 19 — Joint Operation")
fam = f"{PREM}!{PR['fam']}"; ir = f"{PREM}!{PR['ir']}"
C.redbar(6, "CPC 19 — Joint Operation: reconhecimento de 60% do valor justo da JO e imposto diferido 'carimbado'")
C.para(8, "Na operação em conjunto (CPC 19), leitura sustentada pela Deloitte a partir dos vetos qualificados da CNP, as Holdings deixam de avaliar o investimento "
          "pelo método de equivalência patrimonial e passam a reconhecer diretamente a sua parcela (60%) dos ativos e passivos da JO — Embracon e CNP — a valor justo. "
          "Isso inclui o valor justo da própria Embracon, que no CPC 15 não é reavaliada. O ganho de avaliação a valor justo (AVJ) não é tributado no reconhecimento, "
          "desde que controlado em subconta (arts. 13 e 14 da Lei 12.973/14), mas fica 'carimbado': o IRPJ/CSLL de 34% sobre o AVJ é devido quando o investimento for "
          "realizado (venda), qualquer que seja o preço e independentemente de quem venda. As premissas (valor justo da Embracon pelo EV do Acordo de Investimento e "
          "valor de mercado da CNP) vêm da aba 'Premissas'.")
C.sub(10, "1) Avaliação a valor justo da JO (100%)", cols="B:C")
C.header(11, [("B", "Componente", "hdr_l"), ("C", "Valor", "hdr_c"), ("D", "Origem / leitura", "hdr_l")]); C.merge("D11:K11", S["hdr_l"])
rows19 = [
    (12, "Valor justo da Embracon (EV — Acordo de Investimento)", f"={PREM}!{PR['vj_e']}", "lab_it", "num_it", "Aba Premissas. No CPC 15 a Embracon não é reavaliada; no CPC 19 entra integralmente a valor justo."),
    (13, "Valor justo da CNP (= valor de mercado incorporado)", f"={PREM}!{PR['vm_c']}", "lab_it", "num_it", "Aba Premissas; mesmo valor incorporado no CPC 15."),
    (14, "(=) Valor justo total da JO", "=C12+C13", "lab_b_top", "num_b_top", "Embracon + CNP a valor justo."),
    (15, "(−) Base contábil (PL da Embracon + PL contábil da CNP)", f"={PREM}!{PR['pl_e']}+{PREM}!{PR['pl_c']}", "lab_it", "num_it", "Patrimônio a valor contábil das duas companhias."),
    (16, "(=) Ganho de AVJ total da JO", "=C14-C15", "lab_b", "num_b", "Diferença entre valor justo e base contábil; tributação diferida (subconta)."),
]
for row, lab, f, ls, ns, note in rows19:
    C.put(f"B{row}", lab, S[ls]); C.put(f"C{row}", f, S[ns]); st = S["cell_top"] if ls.endswith("top") else S["cell_it"]
    C.put(f"D{row}", note, st); C.merge(f"D{row}:K{row}", st); C.row_h[row] = 30 if len(note) > 90 else 15
C.sub(18, "2) Investimento das Holdings na JO (60% — após o cash-out)", cols="B:D")
C.header(19, [("B", "Componente", "hdr_l"), ("C", "100% da JO", "hdr_c"), ("D", "60% (Holdings)", "hdr_c"), ("E", "Leitura", "hdr_l")], height=16); C.merge("E19:K19", S["hdr_l"])
rows60 = [
    (20, "Valor justo da Embracon", "=C12", f"=C20*{fam}", "lab_it", "num_it", "Parcela que não existe no CPC 15: a Embracon passa a ser reavaliada."),
    (21, "Valor justo da CNP", "=C13", f"=C21*{fam}", "lab_it", "num_it", "Mesmo valor incorporado no CPC 15."),
    (22, "(=) Valor justo total — custo do investimento das Holdings no CPC 19", "=C20+C21", "=D20+D21", "lab_b_top", "num_b_top", "Investimento reconhecido pelas Holdings: 60% do valor justo de Embracon + CNP."),
    (23, "(−) Base contábil (PL da Embracon + PL contábil da CNP)", "=C15", f"=C23*{fam}", "lab_it", "num_it", "Coincide com o custo do CPC 15 sem a camada da mais-valia (custo original + reflexo do PL contábil da CNP)."),
    (24, "(=) Ganho de AVJ com tributação diferida (subconta — arts. 13 e 14, Lei 12.973/14)", "=C22-C23", "=D22-D23", "lab_b", "num_b", "Ganho reconhecido no resultado das Holdings e neutralizado fiscalmente apenas enquanto controlado em subconta."),
    (25, "Passivo fiscal diferido 'carimbado' (34% × AVJ)", f"=C24*{ir}", f"=D24*{ir}", "lab_bi", "num_bi_red", "Devido na realização do investimento (venda), independentemente do preço; é o piso de imposto do CPC 19 na venda futura."),
    (26, "Custo do investimento das Holdings — CPC 19 (60% do valor justo)", "=C22", "=D22", "lab_bi_top", "num_bi", "Valor levado às abas 'Alienação Futura' e 'Conclusão'."),
]
for row, lab, c, d, ls, ns, note in rows60:
    C.put(f"B{row}", lab, S[ls]); C.put(f"C{row}", c, S[ns if ns != "num_bi_red" else "num_bi_red"]); C.put(f"D{row}", d, S[ns])
    if ls.endswith("top") and ns == "num_bi": C.cells[(row, 3)]["style"] = S["num_b_top"]; C.cells[(row, 4)]["style"] = S["num_b_top"]
    st = S["cell_top"] if ls.endswith("top") else S["cell_it"]; C.put(f"E{row}", note, st); C.merge(f"E{row}:K{row}", st)
    C.row_h[row] = 30 if (len(note) > 80 or len(lab) > 44) else 15
# registro de referências usadas por outras abas
C19 = {"vj100": "C14", "base100": "C15", "avj100": "C16", "vj60": "D22", "base60": "D23", "avj60": "D24", "carimbo": "D25", "custo": "D26", "vj_e60": "D20", "vj_c60": "D21"}
C.sub(28, "3) Lançamentos contábeis nas Holdings — CPC 19 (60%)", cols="B:D")
C.header(29, [("B", "Lançamento", "hdr_l"), ("C", "Débito", "hdr_c"), ("D", "Crédito", "hdr_c"), ("E", "Comentário", "hdr_l")], height=16); C.merge("E29:K29", S["hdr_l"])
lanc = [
    (30, "Investimento na JO (60% do valor justo de Embracon + CNP)", "=D22", None, "Reconhecimento direto da parcela das Holdings nos ativos e passivos da JO."),
    (31, "Investimento em Embracon avaliado pelo MEP (base contábil, 60%)", None, "=D23", "Baixa do investimento a valor contábil (custo original + reflexo do PL contábil da CNP)."),
    (32, "Ganho de avaliação a valor justo — resultado (controlado em subconta)", None, "=D24", "Neutralizado fiscalmente enquanto em subconta (arts. 13 e 14 da Lei 12.973/14)."),
    (33, "Despesa de IRPJ/CSLL diferidos — resultado", "=D25", None, "Reconhecimento do imposto sobre o AVJ."),
    (34, "Passivo fiscal diferido (34% × AVJ, 'carimbado')", None, "=D25", "Realizado (pago) na venda do investimento, qualquer que seja o preço."),
]
for row, lab, deb, cred, note in lanc:
    C.put(f"B{row}", lab, S["lab_it"]); C.put(f"C{row}", deb if deb else "–", S["num_it"] if deb else S["dash"]); C.put(f"D{row}", cred if cred else "–", S["num_it"] if cred else S["dash"])
    C.put(f"E{row}", note, S["cell_it"]); C.merge(f"E{row}:K{row}", S["cell_it"]); C.row_h[row] = 30 if (len(note) > 80 or len(lab) > 44) else 15
C.put("B35", "(=) Totais (débitos = créditos)", S["lab_b_top"]); C.put("C35", "=SUM(C30:C34)", S["num_b_top"]); C.put("D35", "=SUM(D30:D34)", S["num_b_top"])
C.put("E35", "=ROUND(C35-D35,6)", S["chk"]); C.put("F35", "← conferência ('ok' = partidas dobradas fechadas)", S["note_top"]); C.merge("F35:K35", S["note_top"]); C.row_h[35] = 16
C.sub(37, "4) Custo do investimento das Holdings após o cash-out — CPC 15 × CPC 19", cols="B:D")
C.header(38, [("B", "Cenário", "hdr_l"), ("C", "Custo (60%)", "hdr_c"), ("D", "Δ vs. CPC 15", "hdr_c"), ("E", "Leitura", "hdr_l")], height=16); C.merge("E38:K38", S["hdr_l"])
C.put("B39", "CPC 15 — custo incrementado remanescente (aba 'Cálculos da Operação', cenário B)", S["lab_it"]); C.put("C39", f"={CALC}!J93", S["num_it"]); C.put("D39", "–", S["dash"])
C.put("E39", "Custo original + reflexo da CNP a valor de mercado, líquido da parcela baixada no cash-out; sem passivo fiscal diferido.", S["cell_it"]); C.merge("E39:K39", S["cell_it"]); C.row_h[39] = 30
C.put("B40", "CPC 19 — 60% do valor justo da JO (quadro 2)", S["lab_it"]); C.put("C40", "=D22", S["num_it"]); C.put("D40", "=C40-C39", S["num_it"])
C.put("E40", "A diferença é o AVJ de toda a JO (60%) menos a mais-valia da CNP que o CPC 15 já leva ao custo; sobre o AVJ há IR diferido de 34%.", S["cell_it"]); C.merge("E40:K40", S["cell_it"]); C.row_h[40] = 30
C.put("B41", "Parcela do custo CPC 19 com IR diferido (AVJ, 60%) e imposto 'carimbado'", S["lab_bi_top"]); C.put("C41", "=D24", S["num_b_top"]); C.put("D41", "=D25", S["num_b_top"])
C.put("E41", "Coluna Δ = imposto carimbado (34% × AVJ), devido na venda em qualquer cenário de preço.", S["cell_top"]); C.merge("E41:K41", S["cell_top"]); C.row_h[41] = 30
C.para(43, "Leitura: no CPC 19 o custo das Holdings sobe para 60% do valor justo de toda a JO, mas o acréscimo é integralmente 'carimbado' — o IR diferido de 34% "
           "sobre o AVJ é devido na venda, qualquer que seja o preço, e permanece na Holding mesmo que a venda seja feita pelas pessoas físicas. Elimina-se a discussão "
           "sobre o custo, ao preço de reconhecer ganho sobre a própria Embracon (não reavaliada no CPC 15) e de elevar o valor patrimonial das Holdings, com possível "
           "impacto na base do ITCMD. A comparação com o CPC 15 na venda futura está na aba 'Alienação Futura (PJ x PF)'.")

# ============================================================ ALIENAÇÃO FUTURA (rebuild)
A = Sheet("Alienação Futura (PJ x PF)", W_WIDE); A.head("Alienação Futura — Análise dos Cenários")
A.redbar(6, "1) Premissas utilizadas na simulação da venda futura dos 60%")
A.header(8, [("B", "Informação", "hdr_l"), ("C", "Valor", "hdr_c"), ("D", "Origem", "hdr_l")], height=16); A.merge("D8:K8", S["hdr_l"])
ali = [
    (9,  "Participação final das famílias (Savian + JVFJ)", f"={PREM}!{PR['fam']}", "pct_it", "Aba Premissas", "fam"),
    (10, "IRPJ/CSLL — ganho de capital das Holdings", f"={PREM}!{PR['ir']}", "pct_it", "Aba Premissas", "ir"),
    (11, "IR — ganho de capital das pessoas físicas (simplificada)", f"={PREM}!{PR['ir_pf']}", "pct_it", "Aba Premissas", "ir_pf"),
    (12, "Custo das quotas das Holdings nas pessoas físicas", f"={PREM}!{PR['custo_pf']}", "num_it", "Aba Premissas — premissa ilustrativa, a confirmar", "custo_pf"),
    (13, "Custo fiscal do investimento — CPC 15 (custo incrementado remanescente após o cash-out)", f"={CALC}!J93", "num_it", "Aba 'Cálculos da Operação', cenário B (custo remanescente)", "c15"),
    (14, "Custo fiscal original remanescente — CPC 15 (cenário A)", f"={CALC}!I93", "num_it", "Aba 'Cálculos da Operação', cenário A — usado na análise hipotética (RISCO 1)", "c15a"),
    (15, "Mais-valia da CNP contida no custo CPC 15 (60%)", f"={PREM}!{PR['mv']}*{PREM}!{PR['fam']}", "num_it", "Aba Premissas (valor de mercado − PL contábil) × 60% — usada na análise hipotética (RISCO 2)", "mv60"),
    (16, "Base contábil — CPC 19 (60%)", f"={CPC19}!{C19['base60']}", "num_it", "Aba CPC 19", "base"),
    (17, "Ganho de AVJ com tributação diferida — CPC 19 (60%)", f"={CPC19}!{C19['avj60']}", "num_it", "Aba CPC 19", "avj"),
    (18, "IR diferido 'carimbado' — CPC 19 (34% × AVJ)", f"={CPC19}!{C19['carimbo']}", "num_it", "Aba CPC 19", "car"),
    (19, "Custo do investimento — CPC 19 (60% do valor justo da JO)", f"={CPC19}!{C19['custo']}", "num_it", "Aba CPC 19", "c19"),
    (20, "Preço de venda — abaixo do valor justo", f"={PREM}!{PR['p_low']}", "int_it", "Aba Premissas", "p_low"),
    (21, "Preço de venda — igual ao valor justo (60% do valor justo da JO)", f"={PREM}!{PR['p_vj']}", "int_it", "Aba Premissas (= aba CPC 19)", "p_vj"),
    (22, "Preço de venda — acima do valor justo", f"={PREM}!{PR['p_high']}", "int_it", "Aba Premissas", "p_high"),
]
AR = {}
for row, lab, f, st, note, key in ali:
    A.put(f"B{row}", lab, S["lab_it"]); A.put(f"C{row}", f, S[st]); AR[key] = f"$C${row}"
    A.put(f"D{row}", note, S["note"]); A.merge(f"D{row}:K{row}", S["note"]); A.row_h[row] = 30 if len(lab) > 44 else 15
A.redbar(24, "2) Venda futura dos 60% — PJ x PF")
A.para(26, "Simulamos a venda dos 60% remanescentes em três preços — abaixo, igual e acima do valor justo (60% do valor justo da JO). No CPC 15, o imposto incide sobre a "
           "diferença entre o preço e o custo fiscal incrementado remanescente (aba 'Cálculos da Operação'). No CPC 19, o imposto diferido de 34% sobre o AVJ é devido de "
           "qualquer forma e funciona como piso; se o preço superar o valor justo, o excedente também é tributado. Na venda pelas pessoas físicas (após 'subir' as ações das "
           "Holdings para as PFs), aplica-se a alíquota simplificada de 22,5% sobre o ganho; porém o custo das PFs é o custo das quotas das Holdings, muito inferior ao custo "
           "da Holding, e no CPC 19 o imposto carimbado permanece na Holding mesmo que a venda seja feita pela PF.")
A.put("B28", "Imposto na venda dos 60%", S["sub_l"]); A.put("C28", None, S["sub_l"])
A.put("D28", "CPC 15", S["hdr_c_blr"]); A.merge("D28:E28", S["hdr_c_blr"]); A.put("F28", "CPC 19", S["hdr_c_blr"]); A.merge("F28:G28", S["hdr_c_blr"])
A.put("H28", "Δ (CPC 19 − CPC 15)", S["hdr_red"]); A.merge("H28:I28", S["hdr_red"]); A.row_h[28] = 16
A.header(29, [("B", "Cenário de preço", "hdr_l"), ("C", "Preço de venda", "hdr_c"), ("D", "Venda PJ", "hdr_c_bl"), ("E", "Venda PF", "hdr_c_br"),
              ("F", "Venda PJ", "hdr_c_bl"), ("G", "Venda PF", "hdr_c_br"), ("H", "Venda PJ", "hdr_red"), ("I", "Venda PF", "hdr_red")], height=16)
venda_rows = [(30, "Abaixo do valor justo", AR["p_low"]), (31, "Igual ao valor justo", AR["p_vj"]), (32, "Acima do valor justo", AR["p_high"])]
for row, lab, p in venda_rows:
    last = row == 32
    A.put(f"B{row}", lab, S["lab_it_bot"] if last else S["lab_it"]); A.put(f"C{row}", f"={p}", S["int_b"])
    A.put(f"D{row}", f"=(C{row}-{AR['c15']})*{AR['ir']}", S["int_bl"] if last else S["int_l"])
    A.put(f"E{row}", f"=(C{row}-{AR['custo_pf']})*{AR['ir_pf']}", S["int_br"] if last else S["int_r"])
    A.put(f"F{row}", f"=MAX({AR['car']},(C{row}-{AR['base']})*{AR['ir']})", S["int_bl"] if last else S["int_l"])
    A.put(f"G{row}", f"={AR['car']}+MAX(0,C{row}-{AR['c19']})*{AR['ir_pf']}", S["int_br"] if last else S["int_r"])
    A.put(f"H{row}", f"=F{row}-D{row}", S["int_b_blr"] if last else S["int_b_lr"]); A.put(f"I{row}", f"=G{row}-E{row}", S["int_b_blr"] if last else S["int_b_lr"]); A.row_h[row] = 16
A.put("B33", "Fórmulas: CPC 15 PJ = (preço − custo incrementado remanescente) × 34% · CPC 15 PF = (preço − custo das quotas) × 22,5% · CPC 19 PJ = maior entre o imposto "
             "carimbado e (preço − base contábil) × 34% · CPC 19 PF = imposto carimbado na Holding + 22,5% sobre o excedente ao valor justo.", S["note"]); A.merge("B33:K33", S["note"]); A.row_h[33] = 30
A.para(34, "Leitura: em preços iguais ou superiores ao valor justo, a diferença entre os regimes na venda pela Holding é pequena e corresponde a 34% da mais-valia da CNP — "
           "a única camada que o CPC 15 leva ao custo e o CPC 19 tributa. Abaixo do valor justo, o CPC 19 é substancialmente mais oneroso, porque o imposto diferido não "
           "diminui com o preço. A venda pela PF só é vantajosa no CPC 15, e apenas a partir de preços em que a alíquota menor compensa o custo menor das quotas; no CPC 19 "
           "ela não elimina o imposto carimbado, que permanece na Holding.")
A.sub(36, "Cenário alternativo — venda abaixo do valor justo e utilização da perda (CPC 19)", cols="B:E", style="sub_l")
A.header(37, [("B", "Movimento na Holding (CPC 19)", "hdr_l"), ("C", "Valor", "hdr_c"), ("D", "Comentário", "hdr_l")], height=16); A.merge("D37:K37", S["hdr_l"])
alt = [
    (38, "Baixa do investimento (60% do valor justo da JO)", f"=-{AR['c19']}", "lab_it", "int_it", "Valor contábil do investimento reconhecido no CPC 19."),
    (39, "Preço recebido", f"={AR['p_low']}", "lab_it", "int_it", ""),
    (40, "(=) Resultado contábil da venda", "=C38+C39", "lab_b", "int_b", "Prejuízo contábil, porque o preço é inferior ao valor justo reconhecido."),
    (41, "IR diferido 'carimbado' devido na realização", f"={AR['car']}", "lab_it", "int_it", "Devido integralmente: a subconta é baixada com a venda."),
    (42, "Crédito teórico sobre a perda (34%)", f"=C40*{AR['ir']}", "lab_it", "int_it", "No lucro presumido a perda não é aproveitável; no lucro real, a compensação é limitada à trava de 30% e depende de lucros futuros."),
    (43, "Imposto devido (se a perda fosse integralmente aproveitada)", "=C41+C42", "lab_b_top", "int_b_top", "Coincide com o RISCO 2 da análise hipotética (mesma base contábil); nos quadros adota-se o piso = imposto carimbado."),
]
for row, lab, f, ls, ns, note in alt:
    A.put(f"B{row}", lab, S[ls]); A.put(f"C{row}", f, S[ns]); st = S["cell_top"] if ls.endswith("top") else S["cell_it"]
    A.put(f"D{row}", note, st); A.merge(f"D{row}:K{row}", st); A.row_h[row] = 30 if (len(note) > 80 or len(lab) > 44) else 15
A.sub(45, "Análise hipotética — risco CPC 15 (fisco não aceita camadas do custo)", cols="B:E", style="sub_l")
A.put("B46", "RISCO 1: o fisco não aceita nenhuma camada do reflexo — resta apenas o custo original remanescente (cenário A da aba 'Cálculos da Operação').", S["cell_reg"]); A.merge("B46:K46", S["cell_reg"]); A.row_h[46] = 15
A.put("B47", "RISCO 2: o fisco não aceita a camada da mais-valia — o custo fiscal recua para a base contábil (custo original + reflexo do PL contábil da CNP), a mesma do CPC 19.", S["cell_reg"]); A.merge("B47:K47", S["cell_reg"]); A.row_h[47] = 15
A.header(49, [("B", "Cenário de preço", "hdr_l"), ("C", "Preço de venda", "hdr_c"), ("D", "CPC 15 — custo integral", "hdr_c"), ("E", "RISCO 2", "hdr_c"), ("F", "RISCO 1", "hdr_c"), ("G", "Δ RISCO 2", "hdr_c"), ("H", "Δ RISCO 1", "hdr_c")], height=30)
A.put("B50", "Custo fiscal considerado em cada hipótese", S["lab_bi"]); A.put("C50", "–", S["dash"])
A.put("D50", f"={AR['c15']}", S["int_bi"]); A.put("E50", f"={AR['c15']}-{AR['mv60']}", S["int_bi"]); A.put("F50", f"={AR['c15a']}", S["int_bi"])
A.put("G50", "=D50-E50", S["int_bi"]); A.put("H50", "=D50-F50", S["int_bi"]); A.row_h[50] = 16
for row, (vr, lab, p) in zip((51, 52, 53), venda_rows):
    A.put(f"B{row}", lab, S["lab_it"]); A.put(f"C{row}", f"=C{vr}", S["int_b"])
    A.put(f"D{row}", f"=(C{row}-$D$50)*{AR['ir']}", S["int"]); A.put(f"E{row}", f"=(C{row}-$E$50)*{AR['ir']}", S["int"]); A.put(f"F{row}", f"=(C{row}-$F$50)*{AR['ir']}", S["int"])
    A.put(f"G{row}", f"=E{row}-D{row}", S["int_b"]); A.put(f"H{row}", f"=F{row}-D{row}", S["int_b"]); A.row_h[row] = 16
A.para(54, "Leitura: cada degrau de risco tem valor fixo, igual a 34% da camada negada, independentemente do preço. O RISCO 2 coincide com o CPC 19 nas vendas ao valor "
           "justo ou acima (mesma base contábil); abaixo do valor justo o CPC 19 continua pior, pelo piso do imposto carimbado. O RISCO 1 supera o CPC 19 em qualquer preço a "
           "partir do valor justo: se a tese de neutralidade do reflexo do PL contábil da CNP fosse integralmente afastada, o CPC 15 deixaria de ser o cenário mais favorável "
           "na venda futura — daí a importância de documentar a tese e o laudo de avaliação (PPA).")
A.redbar(56, "3) Conclusão CPC 15 × CPC 19")
A.header(58, [("B", "Ponto de comparação", "hdr_l"), ("C", "CPC 15", "hdr_c"), ("D", "CPC 19", "hdr_c"), ("E", "Δ (CPC 19 − CPC 15)", "hdr_c"), ("F", "Comentários", "hdr_c_bl")], height=30); A.merge("F58:K58", S["hdr_c_bl"])
concl = [
    (59, "Custo do investimento das Holdings após o cash-out", f"={AR['c15']}", f"={AR['c19']}", "CPC 19: o custo inclui 60% do AVJ de toda a JO; CPC 15: apenas o reflexo da CNP a valor de mercado."),
    (60, "Parcela do custo com IR diferido (CPC 19) ou reflexo neutro (CPC 15)", f"={CALC}!K93", f"={AR['avj']}", "CPC 15: reflexo cuja neutralidade depende de tese; CPC 19: AVJ carimbado, sem discussão."),
    (61, "IR potencial sobre essa parcela (34%)", f"={CALC}!K93*{AR['ir']}", f"={AR['car']}", "CPC 15: devido só se o fisco negar o custo e prevalecer; CPC 19: devido na venda, qualquer que seja o preço."),
    (62, "IR na venda pela Holding — preço abaixo do valor justo", "=D30", "=F30", "O piso do imposto carimbado torna o CPC 19 muito mais oneroso em vendas abaixo do valor justo."),
    (63, "IR na venda pela Holding — preço igual ao valor justo", "=D31", "=F31", "Diferença = 34% da mais-valia da CNP: nos preços a partir do valor justo os regimes se aproximam."),
    (64, "IR na venda pela Holding — preço acima do valor justo", "=D32", "=F32", "Mesma diferença fixa; o excedente ao valor justo é tributado a 34% nos dois regimes."),
    (65, "IR na venda pelas pessoas físicas — preço igual ao valor justo", "=E31", "=G31", "Só o CPC 15 permite capturar a alíquota de 22,5%; no CPC 19 o imposto diferido fica na Holding."),
    (66, "CPC 15 em cenário adverso (RISCO 1) × CPC 19 — preço igual ao valor justo", "=F52", "=F31", "Se a tese de neutralidade do reflexo for integralmente afastada, o CPC 15 passa a ser mais oneroso que o CPC 19."),
]
for row, lab, c, d, note in concl:
    A.put(f"B{row}", lab, S["lab_it"]); A.put(f"C{row}", c, S["num"]); A.put(f"D{row}", d, S["num"]); A.put(f"E{row}", f"=D{row}-C{row}", S["num_b"])
    A.put(f"F{row}", note, S["cell_l"]); A.merge(f"F{row}:K{row}", S["cell_l"]); A.row_h[row] = 30
ALR = {"venda_c15_pj": ["D30", "D31", "D32"], "venda_c15_pf": ["E30", "E31", "E32"], "venda_c19_pj": ["F30", "F31", "F32"], "venda_c19_pf": ["G30", "G31", "G32"], "precos": ["C30", "C31", "C32"], "risco1_vj": "F52"}

# ============================================================ CONCLUSÃO (quadros do PPT / MLA)
K = Sheet("Conclusão", W_WIDE); K.head("Conclusão — CPC 15 × CPC 19")
K.redbar(6, "Custo do investimento para Savian e JVFJ")
K.para(8, "O objetivo destes exemplos é ilustrar o impacto que a classificação da Transação como Combinação de Negócios (CPC 15) ou como Joint Operation (CPC 19) pode gerar "
          "no custo do investimento das Holdings e no caso de venda futura da participação remanescente de 60% da Embracon. Para isso, simulamos a venda com (i) preço "
          "inferior ao valor justo; (ii) preço idêntico ao valor justo; e (iii) preço superior ao valor justo — pela própria Holding ou pelas pessoas físicas (alíquota "
          "simplificada de 22,5% sobre o ganho de capital). Os números são meramente estimados e não devem ser utilizados para ilustrar custos e contingências reais. "
          "Todos os valores são vinculados às abas 'Premissas', 'Cálculos da Operação', 'CPC 19' e 'Alienação Futura (PJ x PF)'.")
K.h14(10, "CPC 15 | Combinação de Negócios", "h14_red")
K.header(11, [("B", "", "hdr_l"), ("C", "PPA CNP", "hdr_c"), ("D", "Antes da Transação", "hdr_c"), ("E", "Depois do Cash-Out", "hdr_c"), ("F", "Δ", "hdr_c"), ("G", "IR 34% se negado", "hdr_c"), ("H", "Risco — 34% diferido", "hdr_l")], height=30); K.merge("H11:K11", S["hdr_l"])
K.put("B12", "Embracon (PL)", S["gray_lab"]); K.put("C12", "–", S["dash"]); K.put("D12", f"={PREM}!{PR['pl_e']}", S["int_b"]); K.put("E12", f"={CALC}!I93", S["int_b"]); K.put("F12", "=E12-D12", S["delta"]); K.put("G12", "–", S["dash"])
K.put("H12", bullet(GREEN, "Custo original remanescente — incontroverso (Δ = parcela baixada no cash-out)."), S["cell_it"]); K.merge("H12:K12", S["cell_it"])
K.put("B13", "CNP (PL)", S["gray_lab"]); K.put("C13", f"={PREM}!{PR['pl_c']}", S["int_it"]); K.put("D13", 0, S["int_b"]); K.put("E13", "=E15-E12-E14", S["int_b"]); K.put("F13", "=E13-D13", S["delta"]); K.put("G13", f"=E13*{PREM}!{PR['ir']}", S["int"])
K.put("H13", bullet(AMBER, "Reflexo do PL contábil da CNP (líquido da diluição) — neutralidade do MEP, risco limitado."), S["cell_it"]); K.merge("H13:K13", S["cell_it"])
K.put("B14", "CNP (AVJ)", S["gray_lab"]); K.put("C14", f"={PREM}!{PR['mv']}", S["int_it"]); K.put("D14", 0, S["int_b"]); K.put("E14", f"={PREM}!{PR['mv']}*{PREM}!{PR['fam']}", S["int_b"]); K.put("F14", "=E14-D14", S["delta"]); K.put("G14", f"=E14*{PREM}!{PR['ir']}", S["int"])
K.put("H14", bullet(RED, "Reflexo do valor justo (mais-valia) do PL da CNP — parcela mais exposta."), S["cell_it"]); K.merge("H14:K14", S["cell_it"])
K.put("B15", "Inv. Embracon", S["hdr_dark_l"]); K.put("C15", "=C13+C14", S["int_dark"]); K.put("D15", "=SUM(D12:D14)", S["int_dark"]); K.put("E15", f"={CALC}!J93", S["int_dark"]); K.put("F15", "=E15-D15", xf(find_font(b=True, color="theme0"), fill=next(i for i, f in enumerate(fills) if fill_sig(f)[0] == "solid" and fill_sig(f)[2] == "0" and fill_sig(f)[3] and fill_sig(f)[3].startswith("-0.3499")), border=B_TBLR, nf=NF_DELTA, h="center", v="center")); K.put("G15", "=SUM(G13:G14)", S["int_dark"])
K.put("H15", "Total = custo incrementado remanescente (aba 'Cálculos', cenário B); PPA = valor de mercado incorporado.", S["cell_it"]); K.merge("H15:K15", S["cell_it"])
for r in (12, 13, 14, 15): K.row_h[r] = 45
K.put("B16", "Conferência: reflexo do PL contábil (linha CNP PL) × cálculo direto [PL CNP × 60% − diluição de 13,8% sobre o custo original × (1 − fração vendida)]", S["lab_it"])
K.put("C16", f"=ROUND(E13-({PREM}!{PR['pl_c']}*{PREM}!{PR['fam']}-{PREM}!{PR['pl_e']}*{PREM}!{PR['p_inc']}*(1-{PREM}!{PR['fr']})),6)", S["chk"]); K.row_h[16] = 45
K.h14(18, "CPC 19 | Joint Operation", "h14_blue")
K.header(19, [("B", "", "hdr_l"), ("C", "100% JO", "hdr_c"), ("D", "60% JO", "hdr_c"), ("E", "IR 34% diferido", "hdr_c"), ("F", "Risco — 34% diferido", "hdr_l")], height=30); K.merge("F19:K19", S["hdr_l"])
K.put("B20", "JO (PL)", S["gray_lab"]); K.put("C20", f"={CPC19}!{C19['base100']}", S["int_it"]); K.put("D20", f"={CPC19}!{C19['base60']}", S["int_b"]); K.put("E20", "–", S["dash"])
K.put("F20", bullet(GREEN, "Reflexo patrimonial (base contábil: PL da Embracon + PL contábil da CNP)."), S["cell_it"]); K.merge("F20:K20", S["cell_it"])
K.put("B21", "AVJ JO", S["gray_lab"]); K.put("C21", f"={CPC19}!{C19['avj100']}", S["int_it"]); K.put("D21", f"={CPC19}!{C19['avj60']}", S["int_b"]); K.put("E21", f"={CPC19}!{C19['carimbo']}", S["int"])
K.put("F21", bullet(RED, "AVJ de Embracon + CNP — IR de 34% 'carimbado', devido na venda em qualquer preço."), S["cell_it"]); K.merge("F21:K21", S["cell_it"])
K.put("B22", "Inv. JO", S["hdr_dark_l"]); K.put("C22", f"={CPC19}!{C19['vj100']}", S["int_dark"]); K.put("D22", f"={CPC19}!{C19['custo']}", S["int_dark"]); K.put("E22", "=E21", S["int_dark"])
K.put("F22", "Total = 60% do valor justo da JO (aba 'CPC 19').", S["cell_it"]); K.merge("F22:K22", S["cell_it"])
for r in (20, 21, 22): K.row_h[r] = 32
K.redbar(24, "Venda futura dos 60%")
K.h14(26, "CPC 15 | Combinação de Negócios", "h14_red", cols="C:E"); K.h14(26, "CPC 19 | Joint Operation", "h14_blue", cols="G:I")
K.header(27, [("C", "Preço", "hdr_l"), ("D", "Venda pela PJ", "hdr_c"), ("E", "Venda pela PF", "hdr_c"), ("G", "Preço", "hdr_l"), ("H", "Venda pela PJ", "hdr_c"), ("I", "Venda pela PF", "hdr_c")], height=30)
for i, row in enumerate((28, 29, 30)):
    K.put(f"C{row}", f"={ALI}!{ALR['precos'][i]}", xf(F_B, fill=FILL_GRAY, border=B_TBLR, nf=NF_INT, h="left", v="center"))
    K.put(f"D{row}", f"={ALI}!{ALR['venda_c15_pj'][i]}", S["int_b"]); K.put(f"E{row}", f"={ALI}!{ALR['venda_c15_pf'][i]}", S["int_b"])
    K.put(f"G{row}", f"={ALI}!{ALR['precos'][i]}", xf(F_B, fill=FILL_GRAY, border=B_TBLR, nf=NF_INT, h="left", v="center"))
    K.put(f"H{row}", f"={ALI}!{ALR['venda_c19_pj'][i]}", S["int_b"]); K.put(f"I{row}", f"={ALI}!{ALR['venda_c19_pf'][i]}", S["int_b"]); K.row_h[row] = 18
K.put("B31", "Imposto (IRPJ/CSLL de 34% na Holding; IR de 22,5% na PF) na venda dos 60%, em R$ milhões — valores da aba 'Alienação Futura (PJ x PF)'.", S["note"]); K.merge("B31:K31", S["note"]); K.row_h[31] = 16
K.redbar(33, "Conclusão executiva")
concl_txt = [
    "• CPC 15 é o cenário mais favorável às Holdings: o custo do investimento incorpora o reflexo da incorporação sem registro de passivo fiscal diferido, o imposto na venda futura acompanha o preço e existe a alternativa de venda pelas pessoas físicas à alíquota de 22,5%. O risco concentra-se na neutralidade do reflexo do PL contábil da CNP; com as premissas atuais a mais-valia é pequena e o imposto potencial sobre ela é residual.",
    "• CPC 19 elimina a discussão sobre o custo, mas ao preço de reconhecer ganho de AVJ sobre toda a JO — inclusive a Embracon, que no CPC 15 não é reavaliada — com IRPJ/CSLL de 34% carimbado e devido na venda em qualquer cenário de preço, sem a alternativa da PF. Além disso, o valor patrimonial das Holdings poderá subir ao valor justo (reconhecimento da CNP e da Embracon), com possível repercussão na base de cálculo do ITCMD, a depender do critério legal aplicável (tema não analisado neste material).",
    "• Na venda ao valor justo ou acima, a carga dos dois regimes é próxima (diferença = 34% da mais-valia da CNP); abaixo do valor justo, o CPC 19 é substancialmente mais oneroso. Se a tese de neutralidade do reflexo fosse integralmente afastada (RISCO 1), o CPC 15 passaria a ser mais oneroso que o CPC 19 nas vendas a partir do valor justo — daí a importância de documentar a tese e o laudo de avaliação (PPA).",
    "Ressalvas: valores ilustrativos em R$ milhões, data-base 30.06.2026, com as premissas da aba 'Premissas'. As Holdings estão no lucro presumido; na venda futura, a carga efetiva na PJ poderia ser reduzida com a opção pelo lucro real e a existência de prejuízos compensáveis. O custo das quotas nas PFs é premissa ilustrativa a confirmar. O material não avalia o ITCMD.",
]
r = 35
for t in concl_txt: K.para(r, t, "para_reg" if t.startswith("•") else "para"); r += 1

# ============================================================ package assembly
sheets_new = [("Premissas", P, "sheet8.xml", "drawing5.xml", 18), ("CPC 19", C, "sheet9.xml", "drawing6.xml", 19), ("Conclusão", K, "sheet10.xml", "drawing7.xml", 20)]
drawing_tpl = files["xl/drawings/drawing2.xml"].decode("utf-8")   # logo da aba Cálculos (âncora B3)
drawing_rels = files["xl/drawings/_rels/drawing2.xml.rels"]
for name, sh, sxml, dxml, sid in sheets_new:
    files[f"xl/worksheets/{sxml}"] = sh.xml(tab_selected=(name == "Premissas"))
    files[f"xl/worksheets/_rels/{sxml}.rels"] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                                 f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/{dxml}"/></Relationships>').encode()
    files[f"xl/drawings/{dxml}"] = drawing_tpl.encode("utf-8"); files[f"xl/drawings/_rels/{dxml}.rels"] = drawing_rels
# Alienação (sheet3) reconstruída; âncora do logo igualada à da aba Cálculos
files["xl/worksheets/sheet3.xml"] = A.xml()
d3 = files["xl/drawings/drawing3.xml"].decode("utf-8")
d3 = re.sub(r'<xdr:from>.*?</xdr:from>', '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>15240</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>145908</xdr:rowOff></xdr:from>', d3, count=1, flags=re.S)
files["xl/drawings/drawing3.xml"] = d3.encode("utf-8")

# Organograma: logo + título (mesma âncora/estilo); as duas imagens do organograma descem 7 linhas para abrir espaço
d1 = files["xl/drawings/drawing1.xml"].decode("utf-8")
d1 = re.sub(r"<xdr:row>(\d+)</xdr:row>", lambda m: f"<xdr:row>{int(m.group(1)) + 7}</xdr:row>", d1)
def _to_one_cell(m):
    body = m.group(0)
    ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', body)
    body = re.sub(r"<xdr:to>.*?</xdr:to>", f'<xdr:ext cx="{ext.group(1)}" cy="{ext.group(2)}"/>', body, count=1, flags=re.S)
    return body.replace('<xdr:twoCellAnchor editAs="oneCell">', "<xdr:oneCellAnchor>").replace("</xdr:twoCellAnchor>", "</xdr:oneCellAnchor>")
d1 = re.sub(r'<xdr:twoCellAnchor editAs="oneCell">.*?</xdr:twoCellAnchor>', _to_one_cell, d1, flags=re.S)
logo_anchor = re.search(r"<xdr:oneCellAnchor>.*?</xdr:oneCellAnchor>", drawing_tpl, re.S).group(0)
logo_anchor = logo_anchor.replace('r:embed="rId1"', 'r:embed="rId3"').replace('id="2" name="Imagem 1"', 'id="4" name="Logo SF"')
d1 = d1.replace("</xdr:wsDr>", logo_anchor + "</xdr:wsDr>")
files["xl/drawings/drawing1.xml"] = d1.encode("utf-8")
r1 = files["xl/drawings/_rels/drawing1.xml.rels"].decode("utf-8")
r1 = r1.replace("</Relationships>", '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image3.emf"/></Relationships>')
files["xl/drawings/_rels/drawing1.xml.rels"] = r1.encode("utf-8")
s1 = files["xl/worksheets/sheet1.xml"].decode("utf-8")
date_xf = xf(F_IT, h="left", v="top")
s1 = re.sub(r"<sheetData>.*?</sheetData>", f'<sheetData><row r="2"><c r="C2" s="6"/></row><row r="4" ht="26" customHeight="1"><c r="E4" s="{S["title"]}" t="inlineStr"><is><t>Organograma</t></is></c></row><row r="5" ht="26" customHeight="1"><c r="E5" s="{date_xf}" t="inlineStr"><is><t>4 de setembro de 2026</t></is></c></row></sheetData>', s1, flags=re.S)
s1 = s1.replace('<dimension ref="C2"/>', '<dimension ref="C2:E5"/>')
s1 = re.sub(r"<cols>.*?</cols>", '<cols><col min="1" max="1" width="7.54296875" customWidth="1"/><col min="3" max="3" width="16.453125" customWidth="1"/><col min="4" max="4" width="20" customWidth="1"/></cols>', s1, count=1, flags=re.S)
files["xl/worksheets/sheet1.xml"] = s1.encode("utf-8")

# Cálculos da Operação: premissas passam a apontar para a aba Premissas (valores idênticos); remove título órfão "CPC 19" (linha 96)
s2 = files["xl/worksheets/sheet2.xml"].decode("utf-8")
links = {"C13": ("n", PR["inv_s"]), "C14": ("n", PR["inv_s"]), "C16": ("n", PR["inv_j"]), "C17": ("n", PR["inv_j"]), "C19": ("n", PR["pl_e"]),
         "C23": ("n", PR["pl_c"]), "C24": ("n", PR["vm_c"]), "I15": ("n", PR["p_inc"]), "I16": ("n", PR["p_cnp"]), "I17": ("n", PR["ir"]), "D68": ("n", PR["cashout"]),
         "I13": ("s", PR["data"]), "I14": ("s", PR["unid"])}
for ref, (kind, target) in links.items():
    m = re.search(rf'<c r="{ref}"([^>]*)>(.*?)</c>', s2)
    assert m, ref
    attrs = re.sub(r'\s+t="[^"]*"', '', m.group(1))
    if kind == "n":
        v = re.search(r'<v>(.*?)</v>', m.group(2)).group(1)
        new = f'<c r="{ref}"{attrs}><f>{PREM}!{target}</f><v>{v}</v></c>'
    else:
        txt = {"I13": "30.06.2026", "I14": "R$ milhões"}[ref]
        new = f'<c r="{ref}"{attrs} t="str"><f>{PREM}!{target}</f><v>{txt}</v></c>'
    s2 = s2.replace(m.group(0), new)
s2 = re.sub(r'<row r="96".*?</row>', '', s2, count=1, flags=re.S)
s2 = re.sub(r'<row r="5"[^>]*>.*?</row>', f'<row r="5" spans="2:11" ht="26" customHeight="1"><c r="E5" s="{xf(F_IT, h="left", v="top")}" t="inlineStr"><is><t>4 de setembro de 2026</t></is></c></row>', s2, count=1, flags=re.S)
for ref, tgt in (("C84", "D69"), ("E84", "D70")):
    m = re.search(rf'<c r="{ref}"([^>]*)><v>600</v></c>', s2); assert m, ref
    s2 = s2.replace(m.group(0), f'<c r="{ref}"{m.group(1)}><f>{tgt}</f><v>600</v></c>')
files["xl/worksheets/sheet2.xml"] = s2.encode("utf-8")

# workbook.xml: ordem das abas, aba ativa, fullCalcOnLoad
wbx = files["xl/workbook.xml"].decode("utf-8")
wbx = wbx.replace('<sheet name="Organograma" sheetId="3" r:id="rId1"/>', '<sheet name="Organograma" sheetId="3" r:id="rId1"/><sheet name="Premissas" sheetId="18" r:id="rId12"/>')
wbx = wbx.replace('<sheet name="Cálculos da Operação" sheetId="15" r:id="rId2"/>', '<sheet name="Cálculos da Operação" sheetId="15" r:id="rId2"/><sheet name="CPC 19" sheetId="19" r:id="rId13"/>')
wbx = wbx.replace('<sheet name="Alienação Futura (PJ x PF)" sheetId="17" r:id="rId3"/>', '<sheet name="Alienação Futura (PJ x PF)" sheetId="17" r:id="rId3"/><sheet name="Conclusão" sheetId="20" r:id="rId14"/>')
assert wbx.count("<sheet ") == 10, wbx.count("<sheet ")
if re.search(r'<workbookView [^>]*activeTab="\d+"', wbx): wbx = re.sub(r'activeTab="\d+"', 'activeTab="1"', wbx, count=1)
else: wbx = re.sub(r'<workbookView ', '<workbookView activeTab="1" ', wbx, count=1)
wbx = re.sub(r'<calcPr calcId="(\d+)"[^>]*/>', r'<calcPr calcId="\1" fullCalcOnLoad="1"/>', wbx)
files["xl/workbook.xml"] = wbx.encode("utf-8")
rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
rels = re.sub(r'<Relationship Id="[^"]+" Type="[^"]+/calcChain" Target="calcChain.xml"/>', '', rels)
rels = rels.replace('</Relationships>', ''.join(f'<Relationship Id="rId{12+i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{s}"/>' for i, s in enumerate(["sheet8.xml", "sheet9.xml", "sheet10.xml"])) + '</Relationships>')
files["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
files.pop("xl/calcChain.xml", None)
ct = files["[Content_Types].xml"].decode("utf-8")
ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', '', ct)
ct = ct.replace('</Types>', ''.join(f'<Override PartName="/xl/worksheets/{s}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for s in ["sheet8.xml", "sheet9.xml", "sheet10.xml"])
                + ''.join(f'<Override PartName="/xl/drawings/{d}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' for d in ["drawing5.xml", "drawing6.xml", "drawing7.xml"]) + '</Types>')
files["[Content_Types].xml"] = ct.encode("utf-8")
app = files["docProps/app.xml"].decode("utf-8")
app = app.replace('<vt:i4>7</vt:i4>', '<vt:i4>10</vt:i4>').replace('<vt:vector size="7" baseType="lpstr"><vt:lpstr>Organograma</vt:lpstr><vt:lpstr>Cálculos da Operação</vt:lpstr><vt:lpstr>Alienação Futura (PJ x PF)</vt:lpstr>',
    '<vt:vector size="10" baseType="lpstr"><vt:lpstr>Organograma</vt:lpstr><vt:lpstr>Premissas</vt:lpstr><vt:lpstr>Cálculos da Operação</vt:lpstr><vt:lpstr>CPC 19</vt:lpstr><vt:lpstr>Alienação Futura (PJ x PF)</vt:lpstr><vt:lpstr>Conclusão</vt:lpstr>')
files["docProps/app.xml"] = app.encode("utf-8")
# a aba Alienação deixa de ser a selecionada (evita agrupamento de abas)
files["xl/worksheets/sheet3.xml"] = files["xl/worksheets/sheet3.xml"].replace(b' tabSelected="1"', b'')

files["xl/styles.xml"] = etree.tostring(styles, xml_declaration=True, encoding="UTF-8", standalone=True)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    order = ["[Content_Types].xml"] + [n for n in zin.namelist() if n != "[Content_Types].xml" and n in files] + [n for n in files if n not in zin.namelist()]
    for n in order: z.writestr(n, files[n])
nf = sum(1 for sh in (P, C, A, K) for d in sh.cells.values() if d["kind"] == "f")
print("saved", OUT, "| fórmulas novas:", nf, "| xfs:", len(cellxfs), "| fontes:", len(fonts))
