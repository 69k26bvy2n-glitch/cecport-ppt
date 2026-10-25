---
name: cecport-ppt
description: 把任意 PPT/Word/PDF/图片等材料重排制作为中电港 CECport 品牌风格的原生可编辑 PPTX（咨询报告信息图风、MiSans、蓝白配色加橙色强调）。当用户要求把 PPT 或其他文件做成中电港风格、中电港汇报/宣传/方案演示美化重排时使用。不适用于与中电港品牌无关的通用美化，也不是只改配色不动版式的机械换肤。
---

# 中电港风格 PPT 制作

把杂乱材料**重排**为中电港品牌演示稿：不是解包换色，而是用 python-pptx **从零直建原生可编辑 PPTX**（形状/文本/真图表，客户双击即改；装饰用 PNG 图标）。视觉为咨询报告信息图风——浅灰白底、白卡片、高饱和蓝阶、橙色点睛。

## 三条铁律

1. **文案保真**：不删减、不增义、不造事实；数字、专有名词、承诺表述一字不动。允许归纳小标题、重组呈现顺序。每页生成后做文本自检（原文事实点全部可在新页找到）。
2. **原生可编辑**：正文/形状/流程图一律 python-pptx 对象；只有图标和不可重建的原图用图片。禁止整页位图、禁止 HTML 转 PPTX 当交付。
3. **每页必过 PowerPoint 实开渲染 QA**，不能只看脚本不报错。终审只信 PowerPoint 通道（Keynote 字体度量与避头尾有差异）。

## 参考文档（按需读，勿凭记忆）

- 设计规范与版面铁律：[references/design-spec.md](references/design-spec.md)
- 页型打法与图表/XML 技术细节：[references/page-patterns.md](references/page-patterns.md)
- 接单流程、多格式输入解析、渲染坑、QA 与交付：[references/workflow-qa.md](references/workflow-qa.md)
- 完整可跑样例在 `examples/`：`p11_vmi_swimlane/` 是三角色泳道+双层信息流/物流流程图（穿线标签、斜线旋转标签、原生竖排），做流程页先照它改。
- 品牌素材在 `assets/`：logo、官方页脚、封面/封底装饰、Tabler 图标库（深蓝版 icons2/、白版 icons2w/）。
  - **封面底图首选 `cover-clean.jpg`**（2026-10-18 用户提供）：32:9→16:9 官方弧线装饰版，左上自带 logo 与股票代码，**左侧大片留白且无烤死标题**，可直接在左侧排标题。旧 `cover-bg.jpeg` 烤死了英文主标题与合作伙伴 logo，套标题会压字，仅在需要官方原版封面时用。
  - 封底可用同一张 `cover-clean.jpg` 作书档呼应；`closing-bg.png` 为淡电路纹理（近白底，直接铺在 BG 上会有色差，需先抠白转 alpha）。

## 环境

- python3 + `python-pptx`（1.0.2）+ Pillow；macOS 渲染需 Microsoft PowerPoint；Windows 端用 PowerPoint 或 WPS。
- 字体 MiSans（客户机未装时随交付附字体文件并提示）。
- 所有 osascript/PowerPoint 渲染调用在沙箱环境内都需关闭沙箱执行。

## 工程初始化

每个任务建独立工程目录（优先用户工作目录下，不要只放 /tmp——临时目录会被系统回收）：

```bash
SKILL_DIR="<本技能绝对路径>"
WORK="<用户工作目录>/重排工程"        # 或 /tmp/cec-job-<名>
mkdir -p "$WORK"
cp "$SKILL_DIR/scripts/cec_lib.py" "$WORK/"
cp -R "$SKILL_DIR/assets" "$WORK/"
```

写 `build_<页名>.py`，`import cec_lib as C` 即可。cec_lib 按「同级 assets → ../assets → cwd/assets」自动定位素材。

## 组件库速查（scripts/cec_lib.py）

- 页面：`new_deck()` / `new_page(prs)`（含底色与左下浅蓝、右下浅橙装饰弧）/ `header(slide,标题,kicker)` / `footer(slide,页码)`
- 容器：`card(x,y,w,h)` 白底圆角发丝边轻投影；`card_title(...)` / `icon_title(...,icon名)`
- 文字：`add_text(slide,x,y,w,h, paras, anchor=)`；段落 dict 支持 `runs / align / line_spacing`；run 用 `R(文本, 字号, 粗体, 颜色)`
- 形状：`add_rect(..., fill, line, shape=MSO_SHAPE.XXX, radius, shadow)`；`add_line` / `add_line_xy`；`add_icon(slide, 名字, x, y, 尺寸, white=True)`
- 装饰：`set_gradient(形状, c1, c2, angle)`（只用于蓝→橙装饰细线）；`soft_shadow(...)`；`card_shadow(形状)`（外卡专用轻投影，卡内元素禁用）；`fill_alpha(形状, 百分比)`；`tag(...)` 保税/完税类胶囊；`badge` / `chip` / `icon_chip`
- 圆角：`rr(w,h,r)` 把绝对英寸换算成 adjustments 比例（必须换算！）；常量 `R_CARD=0.10 / R_BOX=0.08`，细则见 design-spec §6
- 流程图：`set_dash(线)` / `set_arrow(线, head=, tail=)`（tail 在画线终点）；泳道 `dashed_box(...)` / `lane_title(...)`；穿线标签 `line_pill(cx,cy,w,文字,号,rot=)`（白胶囊+同色描边，可沿线旋转）、`vlabel(cx,cy,文字,color=)`（原生 eaVert 竖排+白底框压线），打法见 page-patterns §8.3
- 色板常量：`NAVY BLUE SKY MIST`（4 级蓝）、`ORANGE ORANGE2 TINT_OR`、`INK SUB MUTE`、`TINT TINT2 BG CARD HAIR WHITE`
- 画布 13.333×7.5 in，左右页边距 ML=MR=0.55

## 工作流

1. **接单解析**：拿到输入先全文提取文字与媒体（方法见 workflow-qa）。列清单——总页数、每页事实点、数字、页型初判（封面/过渡/长文/卡片/数据/流程/地图/图文/封底）、烤在图里改不了的位图清单。
2. **规划页型**：按 page-patterns 对每页选打法；素材缺图标时从 Tabler 补（方法见该文档）。
3. **建工程写脚本**：可一个 build_all.py 出全本，也可分批 build；同一套 cec_lib 保证全本无色差。
4. **渲染 QA**：`bash scripts/render.sh 成品.pptx qa_out 1600`（macOS）或 `powershell -File scripts/render.ps1 ...`（Windows）。逐页目检，重点查信息分区、溢出、压页脚、修复弹窗（检查方法见 workflow-qa）。
5. **交付**：成品放用户工作目录，用 `computer://` 链接；附渲染图、烤色位图清单、字体提示；不覆盖原文件；中间产物不堆进交付目录。

## 关键红线（详见 references）

- 蓝色只用 4 级 `#0057C8 / #1478E0 / #4599F0 / #6AAEF3`，最浅也不许发灰；大面积强调色只有橙 `#E86F0B`，橙不压暗。
- 形状一律纯色；同一图的数据块不得混“深浅纯色色块 + 渐变色块”；渐变只允许蓝→橙装饰细线或全图统一语言。
- **关联项间距 < 非关联项间距**；多条记录必须做视觉分区（分隔线/贯穿锚条/卡片）。横向条形图标签右对齐贴柱基线，不加序号圆。
- 中文避头尾只走段落 `eaLnBrk/hangingPunct` + run 标 zh-CN；**绝不写 presentation 级 kinsoku 表**（PowerPoint 弹修复框）。
- logo/页脚用官方 PNG 原图，无效果；页脚坐标固定，任何元素不得压它。
