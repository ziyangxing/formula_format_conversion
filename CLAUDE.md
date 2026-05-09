# Formula Format Conversion

将AI对话（ChatGPT、DeepSeek、豆包等）输出的纯文本公式自动转换为 Word 专业公式格式。

## 使用方式

1. 在 Word 中打开 `word-template/FormulaConverter.docm`
2. 粘贴 AI 对话内容
3. 全选（Ctrl+A），点击工具栏"转换公式"按钮
4. 所有 `$...$` 和 `$$...$$` 公式自动转为专业格式

## 项目结构

```
├── CLAUDE.md                          # 本文件
├── README.md                          # 用户使用说明
├── dist/
│   └── FormulaConverter.exe           # ✨ 打包好的 .exe（直接可用）
├── scripts/
│   ├── gui.py                         # GUI 界面（tkinter）
│   ├── convert.py                     # 主转换逻辑
│   ├── latex2omml.py                  # LaTeX → OMML Professional 编译器
│   ├── create_template.py             # 创建 .docm 模板（仅MS Word）
│   └── build.bat                      # PyInstaller 打包脚本
├── src/                               # VBA源码（用于MS Word内置转换）
│   ├── FormulaConverter.bas           # 主入口模块
│   ├── FormulaFinder.bas              # 公式查找
│   ├── FormulaValidator.bas           # 公式验证
│   ├── LaTeXTranslator.bas            # LaTeX → Word线性格式翻译
│   ├── EquationBuilder.bas            # OMath创建与BuildUp
│   └── FormulaMatch.cls               # 公式匹配信息类
├── word-template/
│   └── test.docx                      # 测试文件（物理公式）
└── tests/
    └── sample_input.txt               # 测试样本
```

## 技术要点

### Python 方案 (主要)
- 使用 `python-docx` + `lxml` 直接操作文档 XML，不依赖 COM
- 兼容 WPS 和 Microsoft Word
- 递归下降解析器将 LaTeX 编译为 OMML Professional XML
- 生成结构化元素：m:f(分数)、m:sSup(上标)、m:sSub(下标)、m:rad(根号)、m:nary(求和/积分)、m:acc(重音符)
- 公式扫描：先块级($$, \\[\\])后行内($, \\(\\))，逐字符遍历避免歧义
- 验证层过滤：货币金额、代码块、转义符

### VBA 方案 (备选，仅MS Word)
- 使用 `VBScript_RegExp_55.RegExp` 作辅助正则
- `OMath.BuildUp` 转换线性格式为专业格式

## 注意事项

- **WPS/Word**: 需要 2019 或更高版本（支持 OMML 方程）
- 处理后的公式以 OMML Professional XML 存储，打开即显示为专业数学排版
- 转换生成新文件，原文件不变
