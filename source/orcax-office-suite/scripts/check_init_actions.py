import zipfile
import re

with zipfile.ZipFile('dotm_initial.dotm', 'r') as z:
    xml = z.read('customUI/customUI14.xml').decode('utf-8')

actions = re.findall(r'onAction="([^"]+)"', xml)
print('onAction sample in dotm_initial:', actions[:20])

# Check onLoad
onload = re.findall(r'onLoad="([^"]+)"', xml)
print('onLoad:', onload)

# Check tab id and label
tabs = re.findall(r'<tab\s+[^>]+>', xml)
print('Tabs:', tabs)
