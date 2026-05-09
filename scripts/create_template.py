"""
create_template.py - 自动创建 FormulaConverter.docm 模板

将 src/ 下的 .bas 文件导入 Word，生成启用宏的 .docm 模板。

使用方法:
    pip install pywin32
    python scripts/create_template.py

输出: word-template/FormulaConverter.docm
"""

import os
import sys
import shutil

try:
    import win32com.client
except ImportError:
    print("请先安装 pywin32: pip install pywin32")
    sys.exit(1)

# 路径配置
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(PROJECT_ROOT, "src")
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "word-template")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "FormulaConverter.docm")

# 需要导入的模块（按依赖顺序）
MODULES = [
    "FormulaMatch.cls",
    "FormulaFinder.bas",
    "FormulaValidator.bas",
    "LaTeXTranslator.bas",
    "EquationBuilder.bas",
    "FormulaConverter.bas",
]


def create_template():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print("启动 Microsoft Word...")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    try:
        doc = word.Documents.Add()

        # 访问 VBA 项目
        vb_project = doc.VBProject

        for module_file in MODULES:
            module_path = os.path.join(SRC_DIR, module_file)
            if not os.path.exists(module_path):
                print(f"  警告: 找不到 {module_path}, 跳过")
                continue

            module_name = os.path.splitext(module_file)[0]
            print(f"  导入: {module_file} → {module_name}")

            # 如果模块已存在则删除
            for comp in vb_project.VBComponents:
                if comp.Name == module_name:
                    vb_project.VBComponents.Remove(comp)
                    break

            # 导入 .bas 文件
            vb_project.VBComponents.Import(module_path)

        # 另存为 .docm
        print(f"保存模板: {OUTPUT_FILE}")
        doc.SaveAs2(OUTPUT_FILE, FileFormat=13)  # 13 = wdFormatXMLDocumentMacroEnabled
        doc.Close()

        print("完成! 模板已生成: " + OUTPUT_FILE)

    finally:
        word.Quit()


if __name__ == "__main__":
    create_template()
