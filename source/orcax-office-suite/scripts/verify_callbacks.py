import zipfile
import re
from oletools import olevba

with zipfile.ZipFile('distribution/orcax-commercial-pack/OrcaX_Addin.dotm', 'r') as z:
    xml = z.read('customUI/customUI14.xml').decode('utf-8', errors='ignore')

# Extract all callback attributes
callbacks = re.findall(r'(?:onAction|getItemCount|getItemLabel|getSelectedItemIndex|getSelectedItemID|getText|getEnabled|onLoad)\s*=\s*"([^"]+)"', xml)
callbacks = sorted(list(set(callbacks)))
print(f"Total callbacks referenced in Ribbon XML: {len(callbacks)}")

vbaparser = olevba.VBA_Parser('distribution/orcax-commercial-pack/OrcaX_Addin.dotm')
all_subs = {}
for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
    for line in vba_code.splitlines():
        line_clean = line.strip()
        m = re.match(r'^(?:Public\s+|Private\s+)?(?:Sub|Function)\s+([A-Za-z0-9_]+)\s*\((.*?)\)', line_clean, re.IGNORECASE)
        if m:
            sub_name = m.group(1)
            all_subs[sub_name] = (vba_filename, m.group(2))

print(f"Total subs/functions in VBA: {len(all_subs)}")

missing = []
found = []
for cb in callbacks:
    if cb in all_subs:
        found.append((cb, all_subs[cb]))
    else:
        missing.append(cb)

print(f"MATCHED: {len(found)}")
print(f"MISSING ({len(missing)}):")
for m in missing:
    print(f"  MISSING: {m}")
