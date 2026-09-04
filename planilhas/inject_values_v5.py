#!/usr/bin/env python3
"""Injeta valores em cache (<v>) nas células de fórmula das abas novas/reconstruídas, a partir da cópia recalculada."""
import sys, zipfile, re, openpyxl, warnings
warnings.filterwarnings("ignore")
final, verified, out = sys.argv[1:4]
parts_map = {"Premissas": "xl/worksheets/sheet8.xml", "CPC 19": "xl/worksheets/sheet9.xml", "Alienação Futura (PJ x PF)": "xl/worksheets/sheet3.xml", "Conclusão": "xl/worksheets/sheet10.xml"}
wb = openpyxl.load_workbook(verified, data_only=True)
z = zipfile.ZipFile(final); parts = {n: z.read(n) for n in z.namelist()}
total = 0
for sheet, part in parts_map.items():
    ws = wb[sheet]; vals = {c.coordinate: c.value for row in ws.iter_rows() for c in row if c.value is not None}
    xml = parts[part].decode("utf-8"); n = 0
    def repl(m):
        global n
        ref, attrs, f = m.group(1), m.group(2), m.group(3); v = vals.get(ref)
        if v is None: return m.group(0)
        n += 1
        if isinstance(v, str): return f'<c r="{ref}"{attrs} t="str"><f>{f}</f><v>{v}</v></c>'
        return f'<c r="{ref}"{attrs}><f>{f}</f><v>{repr(float(v)) if isinstance(v, float) else v}</v></c>'
    xml = re.sub(r'<c r="([A-Z]+\d+)"([^>]*)><f>(.*?)</f></c>', repl, xml)
    parts[part] = xml.encode("utf-8"); total += n; print(sheet, "cached:", n)
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
    for name in z.namelist(): zo.writestr(name, parts[name])
print("total cached values injected:", total)
