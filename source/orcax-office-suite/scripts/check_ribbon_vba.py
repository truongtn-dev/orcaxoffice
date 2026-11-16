import zipfile
from oletools import olevba

dotm_path = 'distribution/orcax-commercial-pack/OrcaX_Addin.dotm'
vbaparser = olevba.VBA_Parser(dotm_path)

for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
    if 'VFormat_Ribbon' in vba_filename:
        print(f"=== {vba_filename} ===")
        print(f"Total lines: {len(vba_code.splitlines())}")
        print("First 40 lines:")
        for line in vba_code.splitlines()[:40]:
            print("  ", line)
