import os
import oletools.olevba as vba

# ------------------------------------------------------------------
# 定数
OUT_DIR = "./out"
"""出力するフォルダ"""

# ------------------------------------------------------------------
# 日本語出力用の設定
def _bytes2str(bytes_string: bytes, encoding="utf-8"):
    # ShiftJis でデコードする
    return bytes_string.decode("shift_jis", errors="replace")

# VBA を文字列に変換する関数を置き換える
vba.bytes2str = _bytes2str

# ------------------------------------------------------------------
# 処理
vba_parser = vba.VBA_Parser("./assets/sample.xlsm")

vba_modules = vba_parser.extract_all_macros() if vba_parser.detect_vba_macros() else []

for _, _, filename, contents in vba_modules:
    file = os.path.join(OUT_DIR, filename + ".vb")
    os.makedirs(os.path.dirname(file), exist_ok=True)
    with open(file, mode="w", encoding="utf-8") as fp:
        fp.write(vba.filter_vba(contents))
