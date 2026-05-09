# AI公式 → Word/WPS 专业公式 转换工具

将 AI 对话（ChatGPT、DeepSeek、豆包等）输出的纯文本公式，转换为 Word/WPS 专业公式格式。

## 效果预览

转换前（纯文本）：

> 根据质能方程 \(E=mc^2\)，动能为 \(E_k = \dfrac{1}{2}mv^2\)
> 牛顿第二定律 \(F = ma\)，万有引力 \(F = G\dfrac{Mm}{r^2}\)

转换后：公式以专业数学排版显示在 Word/WPS 中，分数、上下标、希腊字母正确渲染。

## 系统要求

- **Python 3.8+** + `python-docx` + `lxml`
- **WPS Office** 或 **Microsoft Word 2019+**

## 快速开始

### 方式一：使用打包好的 .exe（推荐，无需安装任何东西）

1. 下载 `FormulaConverter.exe`
2. 双击运行
3. 添加 Word 文件（.docx/.docm）或文件夹
4. 选择输出目录
5. 点击"开始转换"

**支持功能：**
- ✅ 单个/批量转换
- ✅ 添加文件夹（含子文件夹）
- ✅ 自定义输出目录
- ✅ 进度显示

### 方式二：命令行

```bash
pip install python-docx lxml
python scripts/convert.py 你的文档.docx
```

输出文件自动命名为 `你的文档_公式转换.docx`，也可以指定输出文件名：

```bash
python scripts/convert.py 你的文档.docx 输出文件.docx
```

### 打包为 .exe

如需自行打包：

```bash
pip install pyinstaller
scripts/build.bat
```

## 支持的公式格式

| 格式 | 说明 | 示例 |
|------|------|------|
| `\(...\)` | 行内公式 | `\(E=mc^2\)` |
| `\[...\]` | 块级公式 | `\[\frac{a}{b}\]` |
| `$...$` | 行内公式 | `$E=mc^2$` |
| `$$...$$` | 块级公式（居中） | `$$\sum_{i=1}^n x_i$$` |

## 支持的 LaTeX 命令

- **分数**: `\frac{a}{b}`, `\dfrac{a}{b}`
- **根号**: `\sqrt{x}`, `\sqrt[n]{x}`
- **上下标**: `x^2`, `a_i`
- **希腊字母**: `\alpha`, `\beta`, `\gamma`, `\pi`, `\mu`, `\theta`, `\omega` 等
- **求和/积分**: `\sum`, `\int`, `\prod`, `\lim`
- **矩阵**: `\begin{pmatrix}...\end{pmatrix}`, `\begin{bmatrix}...\end{bmatrix}`
- **分段函数**: `\begin{cases}...\end{cases}`
- **对齐公式**: `\begin{aligned}...\end{aligned}`
- **括号**: `\left(...\right)`, `\left\{...\right.`

## 不会转换的内容（智能跳过）

- **货币金额**: `$100`, `$50.00`
- **代码块内公式**: ` ``` ``` ` 和 `` ` `` 包裹的内容
- **转义符号**: `\$变量`

## 在 WPS 中使用

WPS 支持 OMML (Office Math Markup Language) 方程格式。转换后的文档在 WPS 中打开时，公式会显示为线性格式。如需专业格式渲染，在 WPS 中双击公式 → 选择"专业型"即可。

## 项目文件结构

```
formula_format_conversion/
├── CLAUDE.md                          # 项目开发文档
├── README.md                          # 本文件
├── scripts/
│   ├── convert.py                     # ✨ 主转换脚本（Python, 兼容WPS+Word）
│   └── create_template.py             # 自动创建 .docm 模板（仅MS Word）
├── src/                               # VBA源码（用于MS Word内置转换）
│   ├── FormulaConverter.bas           # 主入口
│   ├── FormulaFinder.bas              # 公式查找
│   ├── FormulaValidator.bas           # 公式验证
│   ├── LaTeXTranslator.bas            # LaTeX翻译
│   ├── EquationBuilder.bas            # 公式构建
│   └── FormulaMatch.cls               # 匹配信息类
├── word-template/
│   └── test.docm                      # 测试文件
└── tests/
    └── sample_input.txt               # 测试样本
```

## 常见问题

**Q: WPS 转换后公式显示为文本而非数学格式？**
A: WPS 默认以线性格式显示 OMML 方程。双击公式 → 在设计选项卡中选择"专业型"即可切换。

**Q: 转换后公式不显示？**
A: 确保使用 WPS 2019+ 或 MS Word 2019+，旧版本对 OMML 支持有限。

**Q: 某些复杂公式转换不正确？**
A: 复杂 LaTeX 环境（如 `\begin{align}`、`\begin{array}`）可能转换不完整。建议简化公式或拆分为更小的表达式。

**Q: 如何撤销转换？**
A: 转换会生成新文件，原文件不变。如需撤销，直接使用原文件即可。

**Q: 如果使用 MS Word 想要一键转换？**
A: 请参考下面的"MS Word 宏模板"章节，使用 VBA 宏在 Word 内直接转换。

---

## MS Word 宏模板（可选，Word 用户）

### 安装步骤

1. 打开 **Microsoft Word**
2. 按 **Alt+F11** 打开 VBA 编辑器
3. 依次导入 `src` 目录下的 6 个模块文件：
   - `FormulaMatch.cls`
   - `FormulaFinder.bas`
   - `FormulaValidator.bas`
   - `LaTeXTranslator.bas`
   - `EquationBuilder.bas`
   - `FormulaConverter.bas`
4. 关闭 VBA 编辑器
5. **文件 → 另存为**，文件类型选择 **启用宏的Word文档 (.docm)**

### 使用方法

1. 粘贴 AI 对话内容到文档
2. **Ctrl+A** 全选
3. 运行宏 `ConvertFormulasInSelection`
4. 所有公式自动转为专业格式
