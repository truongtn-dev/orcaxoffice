import os
import shutil
import zipfile
import io
import olefile
import tempfile
import unicodedata

# -------------------------------------------------------------
# CP1258 Encoding Helper
# -------------------------------------------------------------
def to_cp1258(text):
    res = bytearray()
    i = 0
    while i < len(text):
        ch = text[i]
        try:
            b = ch.encode('cp1258')
            res.extend(b)
        except:
            decomp = unicodedata.normalize('NFD', ch)
            base = ""
            tone = ""
            for c in decomp:
                if c in ['\u0300', '\u0301', '\u0303', '\u0309', '\u0323']:
                    tone = c
                else:
                    base += c
            base_nfc = unicodedata.normalize('NFC', base)
            try:
                b_base = base_nfc.encode('cp1258')
                b_tone = tone.encode('cp1258')
                res.extend(b_base + b_tone)
            except Exception as ex:
                res.append(ord('?'))
        i += 1
    return bytes(res)

# -------------------------------------------------------------
# MSForms OLE Stream Block Parser & Builder
# -------------------------------------------------------------
def parse_stream_to_blocks(data):
    pos = 0
    blocks = []
    while pos < len(data):
        tag = data[pos:pos+2]
        bsize = int.from_bytes(data[pos+2:pos+4], 'little')
        bdata = data[pos+4 : pos+4+bsize]
        blocks.append((tag, bytearray(bdata)))
        pos += 4 + bsize
    return blocks

def rebuild_stream_from_blocks(blocks):
    out = bytearray()
    for tag, bdata in blocks:
        out.extend(tag)
        out.extend(len(bdata).to_bytes(2, 'little'))
        out.extend(bdata)
    return bytes(out)

def inspect_ansi_block(bdata):
    for j in range(len(bdata)-4):
        val = int.from_bytes(bdata[j:j+4], 'little')
        if val & 0x80000000:
            slen = val & 0x7FFFFFFF
            if 1 <= slen <= 250 and j+4+slen <= len(bdata):
                raw = bdata[j+4:j+4+slen]
                pad_len = (4 - (slen % 4)) % 4
                suffix = bdata[j+4+slen+pad_len:]
                prefix = bdata[:j]
                return prefix, val, raw, pad_len, suffix
    return None

def set_ansi_block_text(bdata, text_cp1258):
    res = inspect_ansi_block(bdata)
    assert res is not None, "ANSI string not found in block"
    prefix, val, raw, pad_len, suffix = res
    new_len = len(text_cp1258)
    new_pad = (4 - (new_len % 4)) % 4
    new_val = 0x80000000 | new_len
    return bytearray(prefix + new_val.to_bytes(4, 'little') + text_cp1258 + (b'\x00' * new_pad) + suffix)

def set_utf16_block_text(bdata, text_utf16):
    prefix = bdata[:12]
    raw = text_utf16.encode('utf-16le')
    slen = len(raw)
    pad = (4 - (slen % 4)) % 4
    suffix = bdata[-8:]
    return bytearray(prefix + slen.to_bytes(4, 'little') + raw + (b'\x00' * pad) + suffix)

def build_patched_streams(ole_obj):
    # 1. frmPersonalize/o (772 bytes)
    orig_o = ole_obj.openstream(["frmPersonalize", "o"]).read()
    b_o = parse_stream_to_blocks(orig_o)
    b_o[0] = (b_o[0][0], set_utf16_block_text(b_o[0][1], "CÀI ĐẶT CÁ NHÂN HÓA DỮ LIỆU ORCAX"))
    b_o[2] = (b_o[2][0], set_utf16_block_text(b_o[2][1], "Thiết lập thông tin Cơ quan, Người ký & Nơi nhận mặc định để chèn"))
    b_o[4] = (b_o[4][0], set_ansi_block_text(b_o[4][1], to_cp1258("Hành chính (NĐ 30)")))
    b_o[6] = (b_o[6][0], set_ansi_block_text(b_o[6][1], to_cp1258("Đảng bộ (HD 05)")))
    b_o[8] = (b_o[8][0], set_ansi_block_text(b_o[8][1], to_cp1258("Người ký")))
    b_o[10] = (b_o[10][0], set_ansi_block_text(b_o[10][1], to_cp1258("Nơi nhận")))
    b_o[12] = (b_o[12][0], set_ansi_block_text(b_o[12][1], to_cp1258("LƯU DỮ LIỆU")))
    b_o[14] = (b_o[14][0], set_ansi_block_text(b_o[14][1], to_cp1258("Khôi phục mặc định")))
    b_o[16] = (b_o[16][0], set_ansi_block_text(b_o[16][1], to_cp1258("Đóng")))
    rebuilt_o = rebuild_stream_from_blocks(b_o)
    assert len(rebuilt_o) == len(orig_o), f"frmPersonalize/o size mismatch: {len(rebuilt_o)} vs {len(orig_o)}"

    # 2. frmPersonalize/i07/o (964 bytes)
    orig_i07 = ole_obj.openstream(["frmPersonalize", "i07", "o"]).read()
    b_i07 = parse_stream_to_blocks(orig_i07)
    b_i07[0] = (b_i07[0][0], set_ansi_block_text(b_i07[0][1], to_cp1258("Chọn mẫu lưu:")))
    b_i07[4] = (b_i07[4][0], set_ansi_block_text(b_i07[4][1], to_cp1258("Tên gợi nhớ (Menu):")))
    b_i07[8] = (b_i07[8][0], set_ansi_block_text(b_i07[8][1], to_cp1258("Cơ quan ban hành:")))
    b_i07[12] = (b_i07[12][0], set_ansi_block_text(b_i07[12][1], to_cp1258("Cơ quan cấp trên:")))
    b_i07[16] = (b_i07[16][0], set_ansi_block_text(b_i07[16][1], to_cp1258("Địa danh hành chính:")))
    b_i07[20] = (b_i07[20][0], set_ansi_block_text(b_i07[20][1], to_cp1258("Ký hiệu số văn bản:")))
    note_admin = "Ghi chú NĐ 30: Cấp Bộ không có cấp trên. Cấp Cục: Cấp trên là BỘ, ban hành CỤC... Phía dưới có đường kẻ 1/3-1/2. Tên gợi nhớ hiện trên Ribbon."
    b_i07[24] = (b_i07[24][0], set_ansi_block_text(b_i07[24][1], to_cp1258(note_admin)))
    rebuilt_i07 = rebuild_stream_from_blocks(b_i07)
    assert len(rebuilt_i07) == len(orig_i07), f"i07/o size mismatch: {len(rebuilt_i07)} vs {len(orig_i07)}"

    # 3. frmPersonalize/i21/o (904 bytes)
    orig_i21 = ole_obj.openstream(["frmPersonalize", "i21", "o"]).read()
    b_i21 = parse_stream_to_blocks(orig_i21)
    b_i21[0] = (b_i21[0][0], set_ansi_block_text(b_i21[0][1], to_cp1258("Chọn mẫu Đảng:")))
    b_i21[4] = (b_i21[4][0], set_ansi_block_text(b_i21[4][1], to_cp1258("Tên gợi nhớ (Menu):")))
    b_i21[8] = (b_i21[8][0], set_ansi_block_text(b_i21[8][1], to_cp1258("Cơ quan Đảng ban hành:")))
    b_i21[12] = (b_i21[12][0], set_ansi_block_text(b_i21[12][1], to_cp1258("Cấp ủy cấp trên:")))
    b_i21[16] = (b_i21[16][0], set_ansi_block_text(b_i21[16][1], to_cp1258("Địa danh hành chính:")))
    b_i21[20] = (b_i21[20][0], set_ansi_block_text(b_i21[20][1], to_cp1258("Ký hiệu số Đảng:")))
    note_party = "Ghi chú: Thể thức theo Hướng dẫn 05-HD/TW. Ký hiệu thường gặp: NQ/ĐU, NQ/TU, NQ/CB."
    b_i21[24] = (b_i21[24][0], set_ansi_block_text(b_i21[24][1], to_cp1258(note_party)))
    rebuilt_i21 = rebuild_stream_from_blocks(b_i21)
    assert len(rebuilt_i21) == len(orig_i21), f"i21/o size mismatch: {len(rebuilt_i21)} vs {len(orig_i21)}"

    # 4. frmPersonalize/i35/o (912 bytes)
    orig_i35 = ole_obj.openstream(["frmPersonalize", "i35", "o"]).read()
    b_i35 = parse_stream_to_blocks(orig_i35)
    b_i35[0] = (b_i35[0][0], set_ansi_block_text(b_i35[0][1], to_cp1258("Chọn mẫu người ký:")))
    b_i35[4] = (b_i35[4][0], set_ansi_block_text(b_i35[4][1], to_cp1258("Tên gợi nhớ (Menu):")))
    b_i35[8] = (b_i35[8][0], set_ansi_block_text(b_i35[8][1], to_cp1258("Thẩm quyền ký (dòng trên):")))
    b_i35[12] = (b_i35[12][0], set_ansi_block_text(b_i35[12][1], to_cp1258("Chức vụ cụ thể:")))
    b_i35[16] = (b_i35[16][0], set_ansi_block_text(b_i35[16][1], to_cp1258("Họ và tên người ký:")))
    b_i35[20] = (b_i35[20][0], set_ansi_block_text(b_i35[20][1], to_cp1258("Loại thẩm quyền:")))
    note_signer = "Ghi chú: Thẩm quyền ký nhập 2 dòng (Ví dụ: TM. UBND, CHỦ TỊCH, GIÁM ĐỐC)."
    b_i35[24] = (b_i35[24][0], set_ansi_block_text(b_i35[24][1], to_cp1258(note_signer)))
    rebuilt_i35 = rebuild_stream_from_blocks(b_i35)
    assert len(rebuilt_i35) == len(orig_i35), f"i35/o size mismatch: {len(rebuilt_i35)} vs {len(orig_i35)}"

    # 5. frmPersonalize/i49/o (404 bytes)
    orig_i49 = ole_obj.openstream(["frmPersonalize", "i49", "o"]).read()
    b_i49 = parse_stream_to_blocks(orig_i49)
    b_i49[0] = (b_i49[0][0], set_ansi_block_text(b_i49[0][1], to_cp1258("Chọn mẫu nơi nhận:")))
    b_i49[4] = (b_i49[4][0], set_ansi_block_text(b_i49[4][1], to_cp1258("Tiêu đề danh mục:")))
    b_i49[8] = (b_i49[8][0], set_ansi_block_text(b_i49[8][1], to_cp1258("Danh sách nơi nhận (mỗi dòng 1 mục):")))
    rebuilt_i49 = rebuild_stream_from_blocks(b_i49)
    assert len(rebuilt_i49) == len(orig_i49), f"i49/o size mismatch: {len(rebuilt_i49)} vs {len(orig_i49)}"

    return {
        "frmPersonalize/o": rebuilt_o,
        "frmPersonalize/i07/o": rebuilt_i07,
        "frmPersonalize/i21/o": rebuilt_i21,
        "frmPersonalize/i35/o": rebuilt_i35,
        "frmPersonalize/i49/o": rebuilt_i49,
    }

def patch_dotm_file(dotm_path):
    print(f"\nPatching dotm: {dotm_path}")
    with zipfile.ZipFile(dotm_path, 'r') as z:
        file_dict = {name: z.read(name) for name in z.namelist()}

    vba_data = file_dict["word/vbaProject.bin"]
    
    # Write to a temp file for olefile in-place write_stream
    with tempfile.NamedTemporaryFile(delete=False, suffix=".bin") as tf:
        temp_vba_path = tf.name
        tf.write(vba_data)

    ole = olefile.OleFileIO(temp_vba_path, write_mode=True)
    patched_streams = build_patched_streams(ole)

    for sname, sbytes in patched_streams.items():
        ole.write_stream(sname, sbytes)
        print(f"  [OK] Wrote stream {sname} ({len(sbytes)} bytes)")
    ole.close()

    # Read back modified vbaProject.bin
    with open(temp_vba_path, "rb") as tf:
        new_vba_data = tf.read()
    try:
        os.remove(temp_vba_path)
    except:
        pass

    file_dict["word/vbaProject.bin"] = new_vba_data

    # Re-write .dotm
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in file_dict.items():
            zout.writestr(name, data)

    with open(dotm_path, "wb") as f:
        f.write(out_buf.getvalue())
    print(f"  [DONE] Saved patched dotm -> {dotm_path} ({os.path.getsize(dotm_path)} bytes)")

def main():
    workspace = r"e:\OrcaXOffice"
    src_presets = os.path.join(workspace, "templates", "user_presets.json")
    
    # 1. Update user_presets.json across all locations
    preset_targets = [
        os.path.join(workspace, "distribution", "orcax-commercial-pack", "templates", "user_presets.json"),
        os.path.join(workspace, "distribution", "orcax-office-suite", "templates", "user_presets.json"),
        os.path.expandvars(r"%APPDATA%\OrcaX\user_presets.json"),
        os.path.expandvars(r"%APPDATA%\VFormat\user_presets.json"),
    ]
    for pt in preset_targets:
        try:
            os.makedirs(os.path.dirname(pt), exist_ok=True)
            shutil.copyfile(src_presets, pt)
            print(f"[OK] Synced user_presets.json -> {pt}")
        except Exception as e:
            print(f"[SKIP/WARN] Could not sync user_presets.json to {pt}: {e}")

    # 2. Patch dotm files
    dotm_targets = [
        os.path.join(workspace, "distribution", "orcax-commercial-pack", "OrcaX_Addin.dotm"),
        os.path.join(workspace, "distribution", "orcax-office-suite", "OrcaX_Addin.dotm"),
        os.path.expandvars(r"%APPDATA%\Microsoft\Word\STARTUP\OrcaX_Addin.dotm")
    ]
    for dt in dotm_targets:
        try:
            if os.path.exists(dt):
                patch_dotm_file(dt)
        except Exception as e:
            print(f"[SKIP/WARN] Could not patch {dt}: {e}")

    print("\n[ALL TASKS COMPLETED SUCCESSFULLY!]")

if __name__ == "__main__":
    main()
