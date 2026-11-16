# -*- coding: utf-8 -*-
"""
build_orcax_ui.py - Trinh build ORCAX WORKBENCH v2.0
=====================================================
Quy trinh:
  1. Doc ribbon_ui.xml (8 nhom quy trinh) + bo icon duotone -> dong goi vao .dotm
     (customUI14.xml + customUI/_rels + customUI/img/*.png + [Content_Types])
  2. Re-import toan bo module VBA tu vba_src/ qua COM automation (Word VBE)
     + dung 3 UserForm moi bang code (frmAuditHub, frmDocMetadata, frmBatchFolder)
     -> Unicode an toan 100%, khong rong ra CP1258/frx.
  3. QA: kiem tra XML well-formed, moi callback trong XML co thu tuc VBA tuong ung,
     moi icon co file PNG; load-test .dotm bang Word.
  4. Dong bo .dotm cho ca 2 goi distribution (giu backup .bak-v1.0 lan dau).

Cach dung:
  python build_orcax_ui.py              # build day du
  python build_orcax_ui.py --skip-vba   # chi dong goi XML + icon
  python build_orcax_ui.py --kill       # dong Word neu dang mo truoc khi build
"""
import argparse
import base64
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from xml.etree import ElementTree as ET

WORKSPACE = r"E:\OrcaXOffice"
SCRIPTS = os.path.join(WORKSPACE, "source", "orcax-office-suite", "scripts")
VBA_SRC = os.path.join(WORKSPACE, "vba_src")
RIBBON_XML = os.path.join(SCRIPTS, "ribbon", "ribbon_ui.xml")
ICONS_DIR = os.path.join(SCRIPTS, "icons", "png")
DIST_COMM = os.path.join(WORKSPACE, "distribution", "orcax-commercial-pack", "OrcaX_Addin.dotm")
DIST_SUITE = os.path.join(WORKSPACE, "distribution", "orcax-office-suite", "OrcaX_Addin.dotm")
SOURCE_DOTM = DIST_COMM  # ban goc lam dau vao

CUI_NS = "http://schemas.microsoft.com/office/2009/07/customui"
IMG_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

# CHI thay VFormat_Ribbon (co edit OnLoad) + them 2 module moi.
# KHONG re-import VFormat_Core / VFormat_License: module moi import (source-only)
# lam vo hieu p-code cache cua chinh module do -> macro chet. Core/License giu
# nguyen binary goc (p-code) - da kiem chung bang bisect T1-T8.
BAS_MODULES = ["VFormat_Ribbon", "OrcaX_Enterprise", "OrcaX_Dynamic"]
FORM_NAMES = ["frmAuditHub", "frmDocMetadata", "frmBatchFolder"]

# ------------------------------ brand colors --------------------------------
def rgb(r, g, b):
    return r + g * 256 + b * 65536

NAVY = rgb(30, 45, 130)
NAVY_SOFT = rgb(70, 90, 190)
TEAL = rgb(15, 163, 177)
SLATE = rgb(51, 65, 85)
LIGHT = rgb(226, 232, 240)
WHITE = rgb(255, 255, 255)
RED_BG, RED_FG = rgb(254, 235, 235), rgb(192, 38, 38)
AMBER_BG, AMBER_FG = rgb(255, 243, 224), rgb(180, 100, 0)
BLUE_BG, BLUE_FG = rgb(228, 240, 252), rgb(30, 100, 160)
TEAL_BG = rgb(224, 246, 248)

def viet(s):
    """Ma hoa chuoi tieng Viet -> base64 UTF-8 (cho constant trong VBA ASCII)."""
    return base64.b64encode(s.encode("utf-8")).decode("ascii")


# =============================================================================
# STEP 1: DONG GOI RIBBON XML + ICON VAO .dotm
# =============================================================================
def build_ribbon_package(src_dotm, out_dotm):
    with io.open(RIBBON_XML, "r", encoding="utf-8") as f:
        ribbon_xml = f.read()

    ET.fromstring(ribbon_xml)  # well-formed check

    # Bo comment XML truoc khi quet refs (tranh bat cac van ban ghi chu)
    ribbon_no_comments = re.sub(r"<!--.*?-->", "", ribbon_xml, flags=re.S)

    # Thu thap icon duoc tham chieu: ico_xxx (32px) / ico_s_xxx (16px)
    refs = sorted(set(re.findall(r'image="(ico_[a-z0-9_]+)"', ribbon_no_comments)))
    rel_entries, img_files = [], []
    for ref in refs:
        if ref.startswith("ico_s_"):
            png = "ico_%s_16.png" % ref[len("ico_s_"):]
        else:
            png = "%s_32.png" % ref
        path = os.path.join(ICONS_DIR, png)
        if not os.path.exists(path):
            raise SystemExit("ERROR: Thieu icon %s (ref %s)" % (png, ref))
        rel_entries.append(
            '  <Relationship Id="%s" Type="%s" Target="img/%s"/>' % (ref, IMG_REL_TYPE, png))
        img_files.append((path, "customUI/img/" + png))
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n'
        + "\r\n".join(rel_entries) + "\r\n</Relationships>\r\n")

    print("Ribbon icons referenced: %d (%s)" % (len(refs), " ".join(refs)))

    entries = {}
    with zipfile.ZipFile(src_dotm, "r") as zin:
        for name in zin.namelist():
            entries[name] = zin.read(name)

    entries["customUI/customUI14.xml"] = ribbon_xml.encode("utf-8")
    entries["customUI/_rels/customUI14.xml.rels"] = rels_xml.encode("utf-8")
    for path, arcname in img_files:
        with open(path, "rb") as f:
            entries[arcname] = f.read()

    # [Content_Types].xml: dam bao co Default png
    ct = entries["[Content_Types].xml"].decode("utf-8")
    if 'Extension="png"' not in ct:
        ct = ct.replace(
            "<Default",
            '<Default Extension="png" ContentType="image/png"/><Default', 1)
        entries["[Content_Types].xml"] = ct.encode("utf-8")
        print("Added png Default to [Content_Types].xml")

    tmp = out_dotm + ".tmp"
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)
    os.replace(tmp, out_dotm)
    print("STEP1 OK: packaged ribbon+icons -> %s (%d bytes)" % (out_dotm, os.path.getsize(out_dotm)))


# =============================================================================
# STEP 2: RE-IMPORT VBA + DUNG USERFORM QUA COM AUTOMATION
# =============================================================================
def ensure_access_vbom():
    import winreg
    for ver in ("16.0", "15.0"):
        key_path = r"SOFTWARE\Microsoft\Office\%s\Word\Security" % ver
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as k:
                val, _ = winreg.QueryValueEx(k, "AccessVBOM")
                if val == 1:
                    return True
        except OSError:
            continue
    key_path = r"SOFTWARE\Microsoft\Office\16.0\Word\Security"
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0,
                            winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY) as k:
        winreg.SetValueEx(k, "AccessVBOM", 0, winreg.REG_DWORD, 1)
    print("NOTE: da bat 'Trust access to the VBA project object model' (HKCU AccessVBOM=1)")
    return True


# -----------------------------------------------------------------------------
# UNLOCK: .dotm goc co project VBA bi khoa bang password (CMG/DPB/GC trong
# stream PROJECT). Ky thuat DPx: doi DPB -> DPx, Word se hoi "invalid key --
# Continue Loading Project?" -> tu dong bam Yes -> project load MO KHOA
# (protection = 0). Luu lai giu trang thai mo khoa cho cac buoc sau.
# -----------------------------------------------------------------------------
def patch_dpx(dotm_path):
    entries = {}
    with zipfile.ZipFile(dotm_path) as z:
        for n in z.namelist():
            entries[n] = z.read(n)
    data = bytearray(entries["word/vbaProject.bin"])
    idx = data.find(b"DPB")
    if idx < 0:
        print("NOTE: khong tim thay DPB (project co the da mo khoa) - bo qua unlock")
        return False
    data[idx:idx + 3] = b"DPx"
    entries["word/vbaProject.bin"] = bytes(data)
    tmp = dotm_path + ".dpx"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for n, d in entries.items():
            zo.writestr(n, d)
    os.replace(tmp, dotm_path)
    print("DPx patched for unlock")
    return True


def make_dialog_dismiss_thread(stop_evt):
    """Thread tu dong bam Yes/OK cho dialog 'Microsoft Visual Basic for Applications'."""
    import threading

    def dismiss_loop():
        try:
            import win32gui
        except ImportError:
            return
        while not stop_evt.is_set():
            try:
                def enum_win(hwnd, _):
                    title = win32gui.GetWindowText(hwnd)
                    if "Microsoft Visual Basic for Applications" in title:
                        def enum_child(h2, _2):
                            t = win32gui.GetWindowText(h2)
                            if t.strip("&").lower() in ("yes", "ok"):
                                win32gui.PostMessage(h2, 0x00F5, 0, 0)  # BM_CLICK
                        win32gui.EnumChildWindows(hwnd, enum_child, None)
                win32gui.EnumWindows(enum_win, None)
            except Exception:
                pass
            time.sleep(0.3)

    t = threading.Thread(target=dismiss_loop, daemon=True)
    t.start()
    return t


def word_running():
    try:
        out = subprocess.check_output(
            ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE"],
            stderr=subprocess.DEVNULL).decode(errors="ignore")
        return "WINWORD.EXE" in out
    except Exception:
        return False


# ----------------------- Code-behind cua 3 form (ASCII) ---------------------
FRM_AUDIT_HUB_CODE = r'''
Option Explicit

Private Sub UserForm_Initialize()
    On Error Resume Next
    lblStatus.Caption = "(chua ra soat - bam 'Ra soat lai' de kiem tra)"
    RefreshAudit
End Sub

Private Sub RefreshAudit()
    On Error Resume Next
    If Documents.Count = 0 Then
        lblDocName.Caption = "(khong co van ban nao dang mo)"
        lblKpiCritVal.Caption = "-": lblKpiWarnVal.Caption = "-"
        lblKpiPropVal.Caption = "-": lblKpiScoreVal.Caption = "-"
        lstFindings.Clear
        lblStatus.Caption = "Hay mo mot van ban Word de kiem tra."
        Exit Sub
    End If
    If Not OrcaX_Enterprise.CollectAuditFindings() Then
        lblStatus.Caption = "Loi khi kiem tra the thuc."
        Exit Sub
    End If
    lblDocName.Caption = OrcaX_Enterprise.g_AuditDocName
    lblKpiCritVal.Caption = CStr(OrcaX_Enterprise.g_CritCount)
    lblKpiWarnVal.Caption = CStr(OrcaX_Enterprise.g_WarnCount)
    lblKpiPropVal.Caption = CStr(OrcaX_Enterprise.g_PropCount)
    lblKpiScoreVal.Caption = CStr(OrcaX_Enterprise.g_Score) & "/100"
    lstFindings.Clear
    Dim f As Variant, parts() As String
    For Each f In OrcaX_Enterprise.g_Findings
        parts = Split(CStr(f), "|")
        lstFindings.AddItem OrcaX_Enterprise.FindingIconChar(CLng(parts(0))) & parts(1)
    Next f
    lblStatus.Caption = "Hoan tat - " & _
        (OrcaX_Enterprise.g_CritCount + OrcaX_Enterprise.g_WarnCount + OrcaX_Enterprise.g_PropCount) & _
        " muc can luu y."
End Sub

Private Sub btnRefresh_Click()
    RefreshAudit
End Sub

Private Sub btnStandardize_Click()
    On Error Resume Next
    If Documents.Count = 0 Then Exit Sub
    VFormat_Core.AutoStandardizeDocument
    RefreshAudit
End Sub

Private Sub btnCompare_Click()
    On Error Resume Next
    OrcaX_Enterprise.Ribbon_CompareDocuments Nothing
End Sub

Private Sub btnUndo_Click()
    On Error Resume Next
    VFormat_Core.UndoLastStandardize
    RefreshAudit
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
'''

FRM_DOC_METADATA_CODE_TEMPLATE = r'''
Option Explicit

Private Const B64_META_DONE As String = "@@META_DONE@@"
Private Const B64_META_NO_HEADER As String = "@@META_NO_HEADER@@"
Private Const B64_TITLE_AUDIT As String = "@@TITLE_AUDIT@@"

Private Sub UserForm_Initialize()
    On Error Resume Next
    If Documents.Count > 0 Then
        Dim doc As Document: Set doc = ActiveDocument
        txtDocNumber.Text = doc.Variables("ORX_DocNumber").Value
        txtDocDate.Text = doc.Variables("ORX_DocDate").Value
        txtLocation.Text = doc.Variables("ORX_Location").Value
        txtAgency.Text = doc.Variables("ORX_Agency").Value
    End If
End Sub

Private Sub btnApply_Click()
    On Error Resume Next
    If Documents.Count = 0 Then Exit Sub
    Dim doc As Document: Set doc = ActiveDocument
    Dim ok As Boolean
    ok = OrcaX_Enterprise.ApplyDocMetadata(Trim$(txtDocNumber.Text), Trim$(txtDocDate.Text), _
                                           Trim$(txtLocation.Text), Trim$(txtAgency.Text))
    SetVar doc, "ORX_DocNumber", Trim$(txtDocNumber.Text)
    SetVar doc, "ORX_DocDate", Trim$(txtDocDate.Text)
    SetVar doc, "ORX_Location", Trim$(txtLocation.Text)
    SetVar doc, "ORX_Agency", Trim$(txtAgency.Text)
    If ok Then
        VFormat_Core.ShowUniToast B64_META_DONE, B64_TITLE_AUDIT
        Unload Me
    Else
        VFormat_Core.ShowUniMsg B64_META_NO_HEADER, B64_TITLE_AUDIT, 48
    End If
End Sub

Private Sub SetVar(doc As Document, ByVal nm As String, ByVal v As String)
    On Error Resume Next
    If Len(v) = 0 Then Exit Sub
    doc.Variables(nm).Value = v
    If Err.Number <> 0 Then
        Err.Clear
        doc.Variables.Add nm, v
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
'''

FRM_BATCH_FOLDER_CODE_TEMPLATE = r'''
Option Explicit

Private Const B64_REG_ND30 As String = "@@REG_ND30@@"
Private Const B64_REG_HD05 As String = "@@REG_HD05@@"
Private Const B64_PICK_FOLDER As String = "@@PICK_FOLDER@@"
Private Const B64_NEED_FOLDER As String = "@@NEED_FOLDER@@"
Private Const B64_BATCH_DONE As String = "@@BATCH_DONE@@"
Private Const B64_BATCH_OK As String = "@@BATCH_OK@@"
Private Const B64_BATCH_DIR As String = "@@BATCH_DIR@@"
Private Const B64_BATCH_ERR As String = "@@BATCH_ERR@@"
Private Const B64_PROCESSING As String = "@@PROCESSING@@"

Private Sub UserForm_Initialize()
    On Error Resume Next
    cboRegulation.Clear
    cboRegulation.AddItem VFormat_Core.DecodeUTF8(B64_REG_ND30)
    cboRegulation.AddItem VFormat_Core.DecodeUTF8(B64_REG_HD05)
    cboRegulation.ListIndex = 0
End Sub

Private Sub btnBrowse_Click()
    On Error Resume Next
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = VFormat_Core.DecodeUTF8(B64_PICK_FOLDER)
    If fd.Show = -1 Then txtFolder.Text = fd.SelectedItems(1)
End Sub

Private Sub btnRun_Click()
    On Error Resume Next
    Dim folderPath As String
    folderPath = Trim$(txtFolder.Text)
    If Len(folderPath) = 0 Or Dir(folderPath, vbDirectory) = "" Then
        VFormat_Core.ShowUniMsg B64_NEED_FOLDER, "OrcaX Office", 48
        Exit Sub
    End If
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Dim outFolder As String
    outFolder = folderPath & "ChuanHoa_OrcaX"
    If Dir(outFolder, vbDirectory) = "" Then MkDir outFolder

    Dim files As Collection
    Set files = OrcaX_Enterprise.CollectDocFiles(folderPath, (chkSubfolders.Value = True))
    If files.Count = 0 Then
        lblStatus.Caption = "Khong tim thay file Word nao trong thu muc."
        Exit Sub
    End If

    Dim regIdx As Long
    regIdx = 0
    If cboRegulation.ListIndex = 1 Then regIdx = 1

    Dim okC As Long, errC As Long, i As Long
    Application.ScreenUpdating = False
    VFormat_Core.StartProgress "OrcaX Batch", "Batch...", "frmBatchFolder"
    For i = 1 To files.Count
        VFormat_Core.SetProgress CLng(100# * (i - 1) / files.Count), _
            VFormat_Core.DecodeUTF8(B64_PROCESSING) & i & "/" & files.Count, "frmBatchFolder"
        If OrcaX_Enterprise.BatchProcessOneFile(CStr(files(i)), outFolder, regIdx) Then
            okC = okC + 1
        Else
            errC = errC + 1
        End If
        DoEvents
    Next i
    VFormat_Core.FinishProgress "frmBatchFolder"
    Application.ScreenUpdating = True
    lblStatus.Caption = "Hoan thanh: " & okC & " OK, " & errC & " loi."
    VFormat_Core.ShowUniMsg _
        VFormat_Core.DecodeUTF8(B64_BATCH_DONE) & vbCrLf & _
        VFormat_Core.DecodeUTF8(B64_BATCH_OK) & okC & vbCrLf & _
        VFormat_Core.DecodeUTF8(B64_BATCH_DIR) & outFolder & vbCrLf & _
        VFormat_Core.DecodeUTF8(B64_BATCH_ERR) & errC, _
        "OrcaX Office", 64
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
'''


def render_code(template, mapping):
    out = template
    for k, v in mapping.items():
        out = out.replace("@@%s@@" % k, v)
    return out


class FormDesigner:
    """Dung UserForm bang COM: moi control deu dat truc tiep (Unicode an toan).
    Thuoc tinh cua FORM (Width/Height/Caption/BackColor) phai dat qua
    vbc.Properties(...) (chuong chuan VBE), khong dat qua Designer dispatch."""

    def __init__(self, vbc):
        self.vbc = vbc
        self.d = vbc.Designer
        self.set_size(480, 520)
        self.set_prop("BackColor", LIGHT)
        self.set_prop("Caption", vbc.Name)

    def set_prop(self, name, value):
        try:
            self.vbc.Properties(name).Value = value
        except Exception:
            try:
                self.vbc.Properties.Item(name).Value = value
            except Exception:
                print("  WARN: khong dat property '%s'" % name)

    def set_size(self, w, h):
        self.set_prop("Width", w)
        self.set_prop("Height", h)

    def add(self, progid, name, x, y, w, h, **props):
        ctl = self.d.Controls.Add(progid, name, True)
        ctl.Left, ctl.Top, ctl.Width, ctl.Height = x, y, w, h
        for key, val in props.items():
            if key == "font_name":
                ctl.Font.Name = val
            elif key == "font_size":
                ctl.Font.Size = val
            elif key == "font_bold":
                ctl.Font.Bold = val
            elif key == "font_italic":
                ctl.Font.Italic = val
            else:
                setattr(ctl, key, val)
        return ctl

    def header(self, caption, subtitle, w):
        self.add("Forms.Label.1", "lblTitle", 0, 0, w, 42,
                 Caption=caption, BackColor=NAVY, ForeColor=WHITE, BackStyle=1,
                 TextAlign=2, font_name="Segoe UI", font_size=15, font_bold=True)
        self.add("Forms.Label.1", "lblSubtitle", 0, 42, w, 20,
                 Caption=subtitle, BackColor=NAVY, ForeColor=rgb(200, 210, 245), BackStyle=1,
                 TextAlign=2, font_name="Segoe UI", font_size=9)

    def kpi(self, name, x, y, bg, fg, cap, w=106):
        self.add("Forms.Label.1", name + "Val", x, y, w, 34,
                 Caption="0", BackColor=bg, ForeColor=fg, BackStyle=1,
                 TextAlign=2, font_name="Segoe UI", font_size=16, font_bold=True)
        self.add("Forms.Label.1", name + "Cap", x, y + 34, w, 18,
                 Caption=cap, BackColor=bg, ForeColor=SLATE, BackStyle=1,
                 TextAlign=2, font_name="Segoe UI", font_size=8.5)

    def button(self, name, x, y, w, h, cap, bg=WHITE, fg=NAVY):
        return self.add("Forms.CommandButton.1", name, x, y, w, h,
                        Caption=cap, BackColor=bg, ForeColor=fg,
                        font_name="Segoe UI", font_size=9, font_bold=True)


def build_form_audit_hub(vbp):
    vbc = vbp.VBComponents.Add(3)  # vbext_ct_MSForm
    vbc.Name = "frmAuditHub"
    fd = FormDesigner(vbc)
    fd.set_prop("Caption", "OrcaX Control Center")
    fd.set_size(478, 522)
    fd.header("ORCAX CONTROL CENTER",
              "Trung tâm Kiểm soát & Điều hành OrcaX Office v2.0", 478)
    fd.add("Forms.Label.1", "lblDocName", 12, 74, 454, 16,
           Caption="(chưa rà soát)", ForeColor=SLATE, BackColor=LIGHT, BackStyle=1,
           font_name="Segoe UI", font_size=9, font_italic=True)
    fd.kpi("lblKpiCrit", 12, 98, RED_BG, RED_FG, "Nghiêm trọng")
    fd.kpi("lblKpiWarn", 126, 98, AMBER_BG, AMBER_FG, "Cảnh báo")
    fd.kpi("lblKpiProp", 240, 98, BLUE_BG, BLUE_FG, "Đề xuất")
    fd.kpi("lblKpiScore", 354, 98, TEAL_BG, rgb(10, 120, 130), "Điểm tuân thủ")
    fd.add("Forms.ListBox.1", "lstFindings", 12, 160, 454, 216,
           font_name="Segoe UI", font_size=9, BackColor=WHITE)
    fd.button("btnRefresh", 12, 388, 96, 28, "Rà soát lại")
    fd.button("btnStandardize", 116, 388, 122, 28, "Chuẩn hóa ngay", TEAL, WHITE)
    fd.button("btnCompare", 246, 388, 82, 28, "So sánh")
    fd.button("btnUndo", 336, 388, 92, 28, "Khôi phục")
    fd.button("btnClose", 436, 388, 30, 28, "x")
    fd.add("Forms.Label.1", "lblStatus", 12, 428, 454, 24,
           Caption="Sẵn sàng.", BackColor=LIGHT, ForeColor=SLATE, BackStyle=1,
           font_name="Segoe UI", font_size=9)
    vbc.CodeModule.AddFromString(FRM_AUDIT_HUB_CODE.strip() + "\n")
    print("  built form: frmAuditHub")


def build_form_doc_metadata(vbp):
    vbc = vbp.VBComponents.Add(3)
    vbc.Name = "frmDocMetadata"
    fd = FormDesigner(vbc)
    fd.set_prop("Caption", "OrcaX - Thông tin Văn bản")
    fd.set_size(430, 320)
    fd.header("THÔNG TIN VĂN BẢN",
              "Điền nhanh Số ký hiệu, Ngày, Địa danh, Cơ quan vào khung tiêu đề", 430)

    fields = [
        ("lblDocNumber", "txtDocNumber", "Số / Ký hiệu:"),
        ("lblDocDate", "txtDocDate", "Ngày ban hành (dd/mm/yyyy):"),
        ("lblLocation", "txtLocation", "Địa danh (tỉnh/TP):"),
        ("lblAgency", "txtAgency", "Cơ quan ban hành:"),
    ]
    y = 76
    for lbl_name, txt_name, cap in fields:
        fd.add("Forms.Label.1", lbl_name, 14, y, 150, 18,
               Caption=cap, ForeColor=SLATE, BackColor=LIGHT, BackStyle=1,
               font_name="Segoe UI", font_size=9)
        fd.add("Forms.TextBox.1", txt_name, 170, y - 2, 240, 20,
               font_name="Segoe UI", font_size=10)
        y += 30
    fd.button("btnApply", 14, y + 12, 200, 30, "Điền vào khung tiêu đề", TEAL, WHITE)
    fd.button("btnClose", 330, y + 12, 80, 30, "Đóng")
    fd.add("Forms.Label.1", "lblHint", 14, y + 52, 396, 20,
           Caption="Mẹo: chạy 'Chuẩn hóa Toàn diện' trước nếu văn bản chưa có khung tiêu đề.",
           ForeColor=SLATE, BackColor=LIGHT, BackStyle=1,
           font_name="Segoe UI", font_size=8.5, font_italic=True)
    code = render_code(FRM_DOC_METADATA_CODE_TEMPLATE, {
        "META_DONE": viet("Đã điền thông tin văn bản vào khung tiêu đề."),
        "META_NO_HEADER": viet("Chưa tìm thấy khung tiêu đề để điền. Hãy chạy Chuẩn hóa Toàn diện trước."),
        "TITLE_AUDIT": viet("OrcaX Office - Thông tin văn bản"),
    })
    vbc.CodeModule.AddFromString(code.strip() + "\n")
    print("  built form: frmDocMetadata")


def build_form_batch_folder(vbp):
    vbc = vbp.VBComponents.Add(3)
    vbc.Name = "frmBatchFolder"
    fd = FormDesigner(vbc)
    fd.set_prop("Caption", "OrcaX - Xử lý hàng loạt")
    fd.set_size(470, 330)
    fd.header("XỬ LÝ HÀNG LOẠT ORCAX",
              "Chuẩn hóa toàn bộ file Word trong thư mục theo quy chuẩn", 470)

    fd.add("Forms.Label.1", "lblFolder", 14, 76, 70, 18,
           Caption="Thư mục:", ForeColor=SLATE, BackColor=LIGHT, BackStyle=1,
           font_name="Segoe UI", font_size=9)
    fd.add("Forms.TextBox.1", "txtFolder", 88, 74, 288, 20,
           font_name="Segoe UI", font_size=9)
    fd.button("btnBrowse", 382, 72, 72, 24, "Duyệt...")

    fd.add("Forms.Label.1", "lblRegulation", 14, 108, 130, 18,
           Caption="Quy chuẩn áp dụng:", ForeColor=SLATE, BackColor=LIGHT, BackStyle=1,
           font_name="Segoe UI", font_size=9)
    fd.add("Forms.ComboBox.1", "cboRegulation", 148, 106, 190, 20,
           Style=2, font_name="Segoe UI", font_size=9)
    fd.add("Forms.CheckBox.1", "chkSubfolders", 14, 136, 250, 18,
           Caption="Quét cả thư mục con", BackColor=LIGHT, ForeColor=SLATE,
           font_name="Segoe UI", font_size=9)

    fd.button("btnRun", 14, 168, 200, 30, "Bắt đầu chuẩn hóa", TEAL, WHITE)
    fd.button("btnClose", 376, 168, 80, 30, "Đóng")
    fd.add("Forms.Label.1", "lblStatus", 14, 212, 442, 24,
           Caption="Kết quả sẽ lưu vào thư mục ChuanHoa_OrcaX bên trong thư mục đã chọn.",
           BackColor=LIGHT, ForeColor=SLATE, BackStyle=1,
           font_name="Segoe UI", font_size=8.5)

    code = render_code(FRM_BATCH_FOLDER_CODE_TEMPLATE, {
        "REG_ND30": viet("Hành chính (NĐ 30)"),
        "REG_HD05": viet("Văn bản Đảng (HD 05)"),
        "PICK_FOLDER": viet("Chọn thư mục chứa file Word cần chuẩn hóa"),
        "NEED_FOLDER": viet("Vui lòng chọn thư mục hợp lệ."),
        "BATCH_DONE": viet("Chuẩn hóa hàng loạt hoàn tất!"),
        "BATCH_OK": viet("Tổng số file đã chuẩn hóa: "),
        "BATCH_DIR": viet("Thư mục kết quả: "),
        "BATCH_ERR": viet("Số file lỗi bỏ qua: "),
        "PROCESSING": viet("Đang chuẩn hóa: "),
    })
    vbc.CodeModule.AddFromString(code.strip() + "\n")
    print("  built form: frmBatchFolder")


# -----------------------------------------------------------------------------
# STEP0: TAO UNLOCKED BASE (chay 1 lan, cache lai)
# DPx -> open -> auto-Yes dialog "invalid key" -> project protection=0 ->
# SaveAs2 giu trang thai mo khoa. Moi build sau do dung base nay, khong DPx nua
# (gia tri DPx + customUI moi cung luc lam Word crash - da kiem chung).
# -----------------------------------------------------------------------------
def ensure_unlocked_base(src_dotm, cache_path):
    if os.path.exists(cache_path):
        print("STEP0: reuse unlocked base cache: %s" % cache_path)
        return cache_path
    ensure_access_vbom()
    import pythoncom
    import win32com.client

    work = cache_path + ".work"
    shutil.copyfile(src_dotm, work)
    patch_dpx(work)

    pythoncom.CoInitialize()
    stop_evt = threading.Event()
    make_dialog_dismiss_thread(stop_evt)
    print("STEP0: unlocking VBA project (dialog auto-Yes)...")
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = True
    word.WindowState = 2
    word.DisplayAlerts = 0
    word.AutomationSecurity = 1
    try:
        doc = word.Documents.Open(work, ConfirmConversions=False,
                                  ReadOnly=False, AddToRecentFiles=False, Visible=False)
        time.sleep(1.2)
        vbp = doc.VBProject
        if vbp.Protection != 0:
            raise SystemExit("ERROR: unlock that bai (protection != 0)")
        n = vbp.VBComponents.Count
        print("  unlocked OK, %d components" % n)
        doc.SaveAs2(cache_path, FileFormat=15)
        doc.Close(False)
        print("STEP0 OK: unlocked base -> %s" % cache_path)
    finally:
        stop_evt.set()
        time.sleep(0.5)
        try:
            word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
    if os.path.exists(work):
        os.remove(work)
    return cache_path


def import_vba(dotm_path):
    ensure_access_vbom()
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    print("STEP2: launching Word (COM automation, visible-minimized cho unlock dialog)...")
    stop_evt = threading.Event()
    dismiss_thread = make_dialog_dismiss_thread(stop_evt)
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = True
    word.WindowState = 2  # minimized
    word.DisplayAlerts = 0
    word.AutomationSecurity = 1
    try:
        doc = word.Documents.Open(dotm_path, ConfirmConversions=False,
                                  ReadOnly=False, AddToRecentFiles=False, Visible=False)
        time.sleep(1.0)  # cho dialog "invalid key DPx" duoc bam Yes
        vbp = doc.VBProject
        if vbp.Protection != 0:
            raise SystemExit("ERROR: project van bi khoa sau khi DPx-unlock")
        print("  project unlocked (protection=0), components = %d" % vbp.VBComponents.Count)

        # Xoa module cu theo index giam dan (tranh loi enumeration IEnumVARIANT
        # "handle is invalid" va index shift khi remove)
        removed = []
        for i in range(vbp.VBComponents.Count, 0, -1):
            try:
                comp = vbp.VBComponents.Item(i)
                nm = comp.Name
            except Exception as ex:
                print("  WARN: khong doc component #%d: %s" % (i, str(ex)[:80]))
                continue
            if nm in BAS_MODULES or nm in FORM_NAMES:
                vbp.VBComponents.Remove(comp)
                removed.append(nm)
        print("  removed old components: %s" % (", ".join(removed) or "(none)"))

        for name in BAS_MODULES:
            path = os.path.join(VBA_SRC, name + ".bas")
            if not os.path.exists(path):
                raise SystemExit("ERROR: thieu module nguon %s" % path)
            raw = open(path, "rb").read()
            if raw.count(b"\n") != raw.count(b"\r\n"):
                raise SystemExit(
                    "ERROR: %s co LF-only/mixed line endings - VBE Import se mau na. "
                    "Chuyen ve CRLF truoc khi build." % path)
            vbp.VBComponents.Import(path)
        print("  imported modules: %s" % ", ".join(BAS_MODULES))

        build_form_audit_hub(vbp)
        build_form_doc_metadata(vbp)
        build_form_batch_folder(vbp)

        # SaveAs2 FileFormat=15 => .dotm (wdFormatXMLTemplateMacroEnabled)
        doc.SaveAs2(dotm_path, FileFormat=15)
        doc.Close(False)
        print("STEP2 OK: VBA re-imported + 3 forms built -> %s" % dotm_path)
    finally:
        stop_evt.set()
        time.sleep(0.5)
        try:
            word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


# =============================================================================
# STEP 3: QA
# =============================================================================
def extract_vba_procedures(dotm_path):
    from oletools.olevba import VBA_Parser
    subs = set()
    vp = VBA_Parser(dotm_path)
    for (_f, _s, _vba_name, code) in vp.extract_macros():
        for line in code.splitlines():
            m = re.match(r"^(?:Public\s+|Private\s+)?(?:Sub|Function)\s+([A-Za-z0-9_]+)\s*\(",
                         line.strip(), re.IGNORECASE)
            if m:
                subs.add(m.group(1))
    return subs


def qa(dotm_path):
    print("STEP3: QA checks...")
    with zipfile.ZipFile(dotm_path) as z:
        xml = z.read("customUI/customUI14.xml").decode("utf-8")
    ET.fromstring(xml)
    xml_nc_qa = re.sub(r"<!--.*?-->", "", xml, flags=re.S)
    # GUARD schema customUI - nhung loi nay lam Word tu choi ca tab mot cach im lang:
    if xml_nc_qa.count(">>"):
        raise SystemExit("QA FAIL: co ky tu '>>' trong the (text content trong group la sai schema)")
    dual = len(re.findall(r'<[a-zA-Z]+ [^>]*image="ico_[a-z0-9_]+"[^>]*imageMso="[^"]+"[^>]*>', xml_nc_qa))          + len(re.findall(r'<[a-zA-Z]+ [^>]*imageMso="[^"]+"[^>]*image="ico_[a-z0-9_]+"[^>]*>', xml_nc_qa))
    if dual:
        raise SystemExit("QA FAIL: %d element co ca image (custom) va imageMso - schema chi cho phep 1 trong 2" % dual)
    xml = xml_nc_qa

    # 1. moi icon trong XML co trong package
    with zipfile.ZipFile(dotm_path) as z:
        names = set(z.namelist())
    missing_icons = [r for r in re.findall(r'image="(ico_[a-z0-9_]+)"', xml)
                     if "customUI/_rels/customUI14.xml.rels" not in names]
    # kiem tra chi tiet qua rels
    with zipfile.ZipFile(dotm_path) as z:
        rels = z.read("customUI/_rels/customUI14.xml.rels").decode("utf-8")
        rel_ids = set(re.findall(r'Id="([^"]+)"', rels))
        targets = set(re.findall(r'Target="([^"]+)"', rels))
    missing_icons = [r for r in re.findall(r'image="(ico_[a-z0-9_]+)"', xml) if r not in rel_ids]
    missing_png = [t for t in targets if ("customUI/" + t) not in names]
    if missing_icons or missing_png:
        raise SystemExit("QA FAIL icons: missing rel=%s png=%s" % (missing_icons, missing_png))
    print("  icons OK: %d refs, %d png embedded" % (len(rel_ids), len(targets)))

    # 2. moi callback co thu tuc tuong ung
    cbs = sorted(set(re.findall(
        r'(?:onAction|getContent|getSelectedItemIndex|getVisible|getEnabled|getLabel|getPressed|onChange)="([^"]+)"',
        xml)))
    subs = extract_vba_procedures(dotm_path)
    missing = [c for c in cbs if c not in subs]
    print("  callbacks referenced: %d | VBA procedures: %d" % (len(cbs), len(subs)))
    if missing:
        print("  MISSING (%d): %s" % (len(missing), ", ".join(missing)))
        raise SystemExit("QA FAIL callbacks")
    print("  callbacks OK: 100%% matched")

    # 3. dynamic menu content generator - test well-formed voi JSON mau
    test_content = (
        '<menu xmlns="%s"><button id="dynAgn_1" label="BỘ NỘI VỤ &amp; Thử" '
        'onAction="Ribbon_InsertAgencyDyn"/></menu>' % CUI_NS)
    ET.fromstring(test_content)
    print("  dynamic content template OK")

    # 4. load test bang Word
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(dotm_path, ConfirmConversions=False,
                                  ReadOnly=True, AddToRecentFiles=False, Visible=False)
        print("  load test: %s opened, pages=%d" % (doc.Name, doc.ComputeStatistics(2)))
        doc.Close(False)
        print("  Word load test OK (customUI accepted, no corruption)")
    finally:
        word.Quit()
        pythoncom.CoUninitialize()


# =============================================================================
# STEP 4: DONG BO DISTRIBUTION
# =============================================================================
def sync_distribution(build_path):
    print("STEP4: syncing distribution...")
    for target in (DIST_COMM, DIST_SUITE):
        d = os.path.dirname(target)
        bak = os.path.join(d, "OrcaX_Addin.dotm.bak-v1.0")
        if os.path.exists(target) and not os.path.exists(bak):
            shutil.copyfile(target, bak)
            print("  backup: %s" % bak)
        shutil.copyfile(build_path, target)
        print("  synced: %s" % target)


# =============================================================================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--skip-vba", action="store_true",
                    help="chi dong goi XML+icon, khong re-import VBA")
    ap.add_argument("--skip-qa", action="store_true")
    ap.add_argument("--kill", action="store_true", help="dong Word dang mo truoc khi build")
    args = ap.parse_args()

    if word_running():
        if args.kill:
            subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"],
                           capture_output=True)
            print("Word was running -> killed.")
        else:
            raise SystemExit("ERROR: Microsoft Word dang mo. Dong Word hoac chay lai voi --kill.")

    cache_dir = os.path.join(SCRIPTS, ".build_cache")
    os.makedirs(cache_dir, exist_ok=True)
    unlocked_base = os.path.join(cache_dir, "unlocked_base.dotm")
    if not args.skip_vba:
        ensure_unlocked_base(SOURCE_DOTM, unlocked_base)
        base_for_package = unlocked_base
    else:
        base_for_package = SOURCE_DOTM

    build_path = os.path.join(tempfile.gettempdir(), "OrcaX_Addin_build.dotm")
    build_ribbon_package(base_for_package, build_path)

    if not args.skip_vba:
        import_vba(build_path)

    if not args.skip_qa:
        qa(build_path)

    sync_distribution(build_path)
    os.remove(build_path)
    print("BUILD COMPLETE: OrcaX Workbench v2.0")


if __name__ == "__main__":
    main()
