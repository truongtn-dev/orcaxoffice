import zipfile
import re
from oletools import olevba

p = r'distribution/orcax-commercial-pack/OrcaX_Addin.dotm'
print(f"Checking {p}...")

with zipfile.ZipFile(p, 'r') as z:
    xml = z.read('customUI/customUI14.xml').decode('utf-8')

# Get all callbacks in XML
callbacks = re.findall(r'(?:onAction|getSelectedItemIndex|onLoad)\s*=\s*"([^"]+)"', xml)
callbacks = sorted(list(set(callbacks)))

# Get all Subs in VBA
parser = olevba.VBA_Parser(p)
vba_subs = {}
for (filename, stream_path, vba_filename, vba_code) in parser.extract_macros():
    for line in vba_code.splitlines():
        line = line.strip()
        m = re.match(r'^(?:Public\s+)?Sub\s+([A-Za-z0-9_]+)\s*\((.*?)\)', line, re.IGNORECASE)
        if m:
            vba_subs[m.group(1)] = (vba_filename, m.group(2))

print(f"Total Ribbon callbacks: {len(callbacks)}")
print(f"Total VBA Subs: {len(vba_subs)}")

missing = []
for cb in callbacks:
    if cb in vba_subs:
        print(f"  [MATCH] {cb} -> {vba_subs[cb][0]} : ({vba_subs[cb][1]})")
    else:
        missing.append(cb)
        print(f"  [MISSING] {cb} NOT FOUND IN VBA!")

print(f"Total Missing: {len(missing)}")
