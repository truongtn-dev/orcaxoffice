import win32com.client
import os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0

doc = word.Documents.Add()

syntaxes = [
    "Ribbon_AutoStandardize",
    "VFormat_Ribbon.Ribbon_AutoStandardize",
    "Project.VFormat_Ribbon.Ribbon_AutoStandardize",
    "OrcaX_Addin.dotm!Ribbon_AutoStandardize",
    "OrcaX_Addin.dotm!VFormat_Ribbon.Ribbon_AutoStandardize",
    "OrcaX_Addin!Ribbon_AutoStandardize",
    "OrcaX_Addin!VFormat_Ribbon.Ribbon_AutoStandardize",
]

for s in syntaxes:
    try:
        print(f"Testing syntax: {s}")
        # We don't actually run it if it hangs, let's test a simple sub or check if it's found
        word.Run(s, None)
        print(f"  -> FOUND AND EXECUTED with syntax: {s}")
        break
    except Exception as e:
        print(f"  -> Not found / Error: {e}")

doc.Close(False)
word.Quit()
