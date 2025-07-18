import win32com.client as win32
import re
import zipfile
from pathlib import Path

# ========== 配置 ==========
SRC_DIR = Path(__file__).parent / 'Source' / 'VBAProject'
WORD_FILE = SRC_DIR / 'temp.docm'
PROJECT_FILE = SRC_DIR / 'vbaProject.bin'
MACRO_FILE = SRC_DIR / 'CUITMacro.bas'
VERSION_FILE = Path(__file__).parent / 'Version.txt'
VALID_EXT = {'.bas', '.frm'} # .frx 随 .frm 一起导入


# ========== 主流程 ==========
def build_word_file():
    word = win32.Dispatch('Word.Application')
    word.Visible = False          # 调试时可设为 True
    try:
        doc = word.Documents.Add()
        prj = doc.VBProject
    except Exception as e:
        print(f'发生错误，检查是否 Word 或 WPS 正在运行：{e}')
        exit(1)

    for f in SRC_DIR.iterdir():
        if f.suffix.lower() in VALID_EXT:
            prj.VBComponents.Import(str(f))
            print(f'导入 {f.name} 文件')

    this_file = SRC_DIR / 'ThisDocument.cls'
    if this_file.exists():
        with open(this_file, encoding='utf-8') as fp:
            lines = fp.readlines()

        for idx, line in enumerate(lines):
            if line.strip().lower() == 'option explicit':
                start = idx
                break

        clean_code = ''.join(lines[start:])  # 保留 Option Explicit 及之后

        mod = prj.VBComponents('ThisDocument').CodeModule
        mod.DeleteLines(1, mod.CountOfLines)  # 可选：清空旧代码
        mod.AddFromString(clean_code)
        print('将 ThisDocument.cls 文件中的内容合并到 Word 文档的 ThisDocument 模块中')

    # 如果文件存在先删除，再生成新文件
    if WORD_FILE.exists():
        WORD_FILE.unlink()
    doc.SaveAs2(str(WORD_FILE), FileFormat=13)  # 13 = wdFormatXMLDocumentMacroEnabled
    doc.Close(SaveChanges=False)
    word.Quit()
    print(f'保存临时 Word 文件 {WORD_FILE}')


def extract_vbaproject():
    """从 OUT_FILE 提取 vbaProject.bin"""
    with zipfile.ZipFile(WORD_FILE) as zf:
        # 通用路径：word/vbaProject.bin
        data = zf.read("word/vbaProject.bin")
        if PROJECT_FILE.exists():
            PROJECT_FILE.unlink()
        PROJECT_FILE.write_bytes(data)
        print(f"从临时 Word 文件中析出 vbaProject.bin")

    WORD_FILE.unlink()


def update_version():
    # 1. 读取版本号
    new_ver = VERSION_FILE.read_text(encoding='utf-8').strip()

    # 2. 正则替换 Const Version = "..."
    pattern = re.compile(r'^(\s*Const Version\s*=\s*"v).*(")', re.IGNORECASE)

    def repl(m):
        return f'{m.group(1)}{new_ver}{m.group(2)}'

    old_lines = MACRO_FILE.read_text(encoding='gbk').splitlines(keepends=True)
    new_lines = [pattern.sub(repl, ln) for ln in old_lines]

    # 3. 写回文件
    MACRO_FILE.write_text(''.join(new_lines), encoding='gbk', newline='')
    print(f'已更新 {MACRO_FILE} -> {new_ver}')


if __name__ == '__main__':
    update_version()
    build_word_file()
    extract_vbaproject()
