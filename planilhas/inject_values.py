#!/usr/bin/env python3
"""Copia os valores calculados (LibreOffice) da cópia verificada para as células de fórmula
da aba 'Segregação de Risco' do arquivo final, como valores em cache (<v>), sem tocar no resto."""
import sys, zipfile, re, openpyxl, warnings
warnings.filterwarnings("ignore")
final, verified, out = sys.argv[1:4]
wb = openpyxl.load_workbook(verified, data_only=True); ws = wb["Segregação de Risco"]
vals = {c.coordinate: c.value for row in ws.iter_rows() for c in row if c.value is not None}
z = zipfile.ZipFile(final); parts = {n: z.read(n) for n in z.namelist()}
xml = parts["xl/worksheets/sheet3.xml"].decode("utf-8")
n = 0
def repl(m):
    global n
    ref, attrs, f = m.group(1), m.group(2), m.group(3)
    v = vals.get(ref)
    if v is None: return m.group(0)
    n += 1
    if isinstance(v, str):
        return f'<c r="{ref}"{attrs} t="str"><f>{f}</f><v>{v}</v></c>'
    return f'<c r="{ref}"{attrs}><f>{f}</f><v>{repr(float(v)) if isinstance(v, float) else v}</v></c>'
xml2 = re.sub(r'<c r="([A-Z]+\d+)"([^>]*)><f>(.*?)</f></c>', repl, xml)
parts["xl/worksheets/sheet3.xml"] = xml2.encode("utf-8")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
    for name in z.namelist(): zo.writestr(name, parts[name])
print("cached values injected:", n)
