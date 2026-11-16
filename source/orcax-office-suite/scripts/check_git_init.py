import subprocess, zipfile, io, re

out = subprocess.check_output(['git', 'show', '65c7353:distribution/orcax-commercial-pack/OrcaX_Addin.dotm'])
z = zipfile.ZipFile(io.BytesIO(out))
xml = z.read('customUI/customUI14.xml').decode('utf-8')

actions = re.findall(r'onAction="([^"]+)"', xml)
print('onAction sample in 65c7353:', actions[:25])
tabs = re.findall(r'<tab\s+[^>]+>', xml)
print('Tabs in 65c7353:', tabs)
