# PDF 英文新闻提取方案

这个项目用于把版式复杂的英文 PDF 新闻杂志整理成结构化结果，并输出到 Word 文档。

## 方案设计

核心目标：

- 正确识别每篇新闻的标题
- 正确拼接正文段落
- 处理左/中/右多栏版面，不把不同栏内容串乱
- 输出可复核的 `json` 和可直接查看的 `docx`

实现思路：

1. 使用 `PyMuPDF` 读取 PDF 的文本块、文本行、坐标、字号信息。
2. 用“标题区 + `By ...` 作者行”的版式特征识别真正的文章起始页。
3. 对首篇 `EDITORIAL` 单独做兼容处理。
4. 在正文页中过滤页眉、页脚、图片说明、`Photo by` 等非正文内容。
5. 按列坐标重建阅读顺序：
   - 同页多栏时按左栏 -> 中栏 -> 右栏排序
   - 保留栏内自上而下顺序
6. 按文本行和段首缩进重建段落。
7. 生成：
   - `output/newschina170_articles.json`
   - `output/newschina170_articles.docx`

## 当前适用范围

当前版本优先覆盖：

- `EDITORIAL`
- `INTERNATIONAL`
- `POLITICS`
- `ENVIRONMENT`
- `COVER STORY`
- `SOCIETY`
- `ECONOMY`
- `CULTURE`
- `INTERVIEW`
- `HISTORY`
- `OUTSIDE IN`
- `ESSAY`
- `FLAVOR OF THE MONTH / REAL CHINESE`

说明：

- 前部的 `MEDIA FOCUS`、`WHAT THEY SAY`、`NETIZEN WATCH` 这类“多短新闻拼版页”版式与长文不同，当前脚本没有作为第一版重点处理对象。
- 如果后续你需要，我可以在这个项目基础上继续加“短新闻拼版识别”规则。

## 运行方式

在 PowerShell 中执行：

```powershell
python .\extract_news.py `
  --pdf "E:\news\newschina170 no.pdf" `
  --json "E:\news\pdf_extract_project\output\newschina170_articles.json" `
  --docx "E:\news\pdf_extract_project\output\newschina170_articles.docx"
```

也可以直接使用一键脚本：

```powershell
cd E:\news\pdf_extract_project

.\run_extract.ps1 `
  -PdfPath "E:\news\你的PDF.pdf" `
  -OutputName "你的输出文件名"
```

例如：

```powershell
.\run_extract.ps1 `
  -PdfPath "E:\news\sample.pdf" `
  -OutputName "sample_result"
```

执行后会生成：

- `output\sample_result.json`
- `output\sample_result.docx`

## 输出说明

- `json` 适合后续做程序二次处理或人工比对
- `docx` 适合直接阅读和业务整理

每篇文章包含：

- 文章标题
- 栏目
- 作者行
- PDF 页码范围
- 正文段落列表
