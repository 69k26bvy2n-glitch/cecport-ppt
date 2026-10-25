# 工作流、输入解析、渲染与交付

## 1. 接单

1. 确认输入文件与交付用途（对外宣传 / 汇报 / 投标）、是否有已定稿参考页。
2. 非 16:9 一律统一到 13.333×7.5，提前告知。
3. 工程目录建在**用户工作目录**下（如 `<工作目录>/重排工程`），不要只放 /tmp（临时目录会被回收，曾出现文件“被删”罗生门）。中间产物（解包、dump、qa 图）可放工程子目录或 /tmp，不堆进交付目录。
4. 初始化：复制 cec_lib.py 与 assets 到工程（见 SKILL.md）。

## 2. 输入解析（把各类材料变成“事实点清单”）

目标产物：每页/每节的文字事实点、数字、专名、原图素材。

### .pptx

- python-pptx 遍历 shapes 提文字；表格行单独提；备注页别忘了。
- `unzip -o x.pptx -d x_dir` 拿 ppt/media/ 全部原图；ppt/slides/slideN.xml 可查烤位图位置。
- 源稿无原生 chart/table（图表全烤位图）是常态：图中的数字以周边文字或图注为准提取。
- 先做页型扫描表：页码 × 内容摘要 × 页型 × 可复用媒体 × 风险（字多/整页位图）。

### .docx

- 优先 python-docx；样式丢失时 macOS `textutil -convert txt 文件.docx -stdout` 兜底。
- 图片在 docx 包 word/media/；表格用 doc.tables 提全。
- 按“标题样式”切段，映射为页；长文档每 3–6 个二级标题组织一章。

### .pdf

- `pdftotext -layout` 提文字（保版式线索）；`pdfimages -all` 抽图；扫描件走视觉读图。
- PDF 宣传稿常有跨页大图，按页截图作为原图素材入卡。

### 图片 / 长图 / 微信截图

- 直接视觉读图记录事实点与数据；含图表的照片只取数据，图一律重画（不截图入稿）。
- 多张素材图过 PIL 统一裁切圆角后入卡。

### .key / 旧模板 / 网页

- .key 先 Keynote 导出 pptx/pdf；网页正文复制后按事实点重组；均不得整页截图当内容。

## 3. 文案保真自检（每页必做）

- 生成后把新页文字 dump 出来，与源事实点清单逐条对：数字（含单位/小数位）、地名、公司名、产品名、时间、承诺语必须一字不差。
- 允许：合并同义重复句、归纳小标题、调顺序、删语气词。不允许：改数字、改专名、加源稿没有的卖点/承诺。
- 原文确实有语病/错字时标出来问用户，不擅自改。
- 源稿固有的“文字过多会溢出”问题重排要消化（提层级/拆分），不保留为已知瑕疵。

## 4. 渲染 QA

### macOS（scripts/render.sh，强制 PowerPoint）

```bash
bash scripts/render.sh 成品.pptx qa_out 1600
# → qa_out/slides/slides.001.png ...
```

链路与坑（已固化在脚本里，改动时保持）：

1. AppleScript 让 PowerPoint open 文件 → `save ... in ... as save as PDF`；AppleScript 原生“save as PNG”在此版 PPT 静默失败。
2. **PowerPoint 是沙盒应用，对任何非自身容器目录写 PDF 都弹授权框导致卡死**；唯一不弹的路径是它自己的容器 `~/Library/Containers/com.microsoft.Powerpoint/Data/tmp/qa-render.pdf`（固定名覆盖），shell 只读出来再转 PNG。
3. PDF→PNG 用 pdf2png.js（JXA + PDFKit 零依赖白底）。注意 JXA 的 null 桥接成 NSNull：目录由 shell 先 mkdir；representationUsingTypeProperties 第二参传 NSDictionary.dictionary；图形状态用 CGContextSaveGState。
4. 成功判据：PDF mtime 晚于本次开始时间，防旧 PDF 残留误判。
5. 所有渲染相关 shell/osascript 调用必须关闭沙箱执行。
6. 偶发卡死：`osascript -e 'tell application "Microsoft PowerPoint" to close every presentation saving no'`（或 quit）后重跑。
7. 前提：装了 PowerPoint、完成首次启动向导、终端/Trae 有辅助功能权限。
8. **Keynote 不作为终审通道**（字体度量、中文避头尾、阴影与 PowerPoint 有差异）。

### Windows（scripts/render.ps1）

```powershell
powershell -ExecutionPolicy Bypass -File scripts\render.ps1 成品.pptx qa_out 1600
```

- COM 一步直接逐页导 PNG，无需 PDF 中转；PowerPoint 不存在自动回退 WPS（KWPP.Application）。
- 自动 Unblock-File 解网络下载标记；输出命名与 macOS 一致。
- 首次运行需手工走完 Office 激活/向导。

### “修复弹窗”检测（XML 结构问题的终审）

渲染后若发现 PDF 异常，用 osascript 打开文件并让 System Events 遍历窗口按钮，检查是否出现“修复/修复内容”；出现即 XML 结构错误（最常见原因：kinsoku 字符表、命名空间漏 /2006/main、spPr/txPr 元素顺序错），修 XML 后重渲，绝不带修复框交付。

## 5. 目检清单（逐页）

1. 文案：对照事实点清单，无丢字/改数/增义。
2. 信息分区：眯眼 3 问（分区可数 / 组内成块 / 组间距>组内距），见 design-spec §4.4。
3. 元素不压页脚带（y>7.0）、不出卡片边、文字不重叠、图片不变形。
4. 颜色只出现规范色板；相邻并列元素不同蓝；橙只有一个主角；无发灰蓝。
5. 渐变只在装饰线或全图统一语言，数据块无纯色/渐变混用。
6. logo/页脚为官方原图且位置尺寸正确、无效果；页码连续。
7. 中文无标点落行首、无孤字成行（可通过加宽度或调措辞解决，措辞不改义）。
8. 图表数据与源稿一致；极小值处理有注释；单位齐全。
9. 整本无色差：所有页用同一 cec_lib 常量，不硬编码旧色值。
10. **高密度区必须放大目检，不能只看 1600px 全页小图**（小尺寸重叠肉眼会漏）：用 PIL 按坐标把标签穿线、图标网格、多行列表、图文交界面切成 4–6 块、放大 2–4x 逐块看；低密度区看全页即可。排布网格时先算单元最小高（图标 d + 间隙 + 标签 h），行距 pitch 必须大于它。
11. 几何疑问用 numpy 量测而非肉眼：渲染图 sc=宽/13.333，按色板阈值（如蓝线 B 通道高、橙线 R 高）提墨线像素，量虚线/竖排文字列的真实中心坐标，确认与设计中心误差 ≤0.01in。

### 用户回改 PPTX 的迭代（重要）

用户常直接在成品 PPT 里手改样式作为"规则示范"。读取时：

- 用 python-pptx dump 形状几何/填充/描边/字号色值，再渲染放大看实际效果，两者对上才算读懂；注意手绘形状常带 PowerPoint 默认样式（如默认蓝 `0070C0` 描边），落到生成器时统一回品牌色板。
- **文件里某个元素缺失，不代表用户有意删除**（可能是挪动时误删）；只有渲染图+XML 双重印证、或用户明说的改动才同步进生成器，拿不准先问一句再动手。
- 规则一旦被用户确认（如"所有标签都这么处理"），全页同类元素一次性统一，不要只改示范的那一个。

## 6. 烤色位图清单

以下内容脚本改不了，必须列入交付说明：

- 整页背景/示意图里烤死的旧 logo、旧配色、旧 slogan；
- 截图里的第三方系统界面、客户数据后台；
- 流程图位图内部的文字。

处理优先级：能用原生重画就重画（地图点位、流程、图表）；重画成本高的原图装卡保留并在清单中注明，必要时请客户提供矢量原图。

## 7. 交付

- 成品与必要的渲染图放用户工作目录；用 `computer://` 链接给 pptx（正文里附关键页 png 链接便于速览）。
- 命名建议：`<原名>_中电港风格.pptx` 或章节名；**绝不覆盖原文件**。
- 交付说明包含：页面结构调整说明、保真声明（数字/专名未动）、烤色位图清单、MiSans 字体提示（未装会被替换，可附字体文件）、Windows 端复核建议。
- 不主动导 PDF（客户要再给）。
