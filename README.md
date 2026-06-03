# PPT 图片转可编辑 PPT 使用说明

这个Skill可以把一组从 PPT 导出的图片，重新生成一个 `.pptx` 文件。

它适合这些情况：

- 你只有一张张 PPT 截图或导出的图片；
- 原始 PPT 文件找不到了；
- 想把这些图片重新做成一个可以继续编辑的 PPT；
- 希望图片里的文字尽量变成可编辑文字。

## 先说清楚：它能做到什么？

这个工具会把每张图片变成一页 PPT，并尽量识别图片里的文字。

生成后的 PPT 默认会尽量做到：

1. **保留页面底图**
   - 保证页面整体看起来接近原图。
   - 如果识别到了文字，工具会先尝试把底图里对应位置的原始文字擦掉。

2. **放上可编辑文字层**
   - 识别成功的文字会变成 PPT 里的文本框。
   - 这些新生成的文字框可以修改、复制、删除、调整字体。

简单理解：

> 底图负责“保留版面”，可编辑文字框负责“真正改文字”。默认不会故意保留底图里的同一份文字，以避免重影。

## 需要准备什么？

准备一个文件夹，里面放好 PPT 导出的图片。

例如：

```text
slide-images/
├── 001.png
├── 002.png
├── 003.png
└── 004.png
```

图片顺序最好按页码命名，比如：

```text
001.png
002.png
003.png
```

## 最简单的使用方式

在 Codex 里告诉它：

```text
请使用这个 Skill，把 slide-images 文件夹里的图片生成 editable.pptx。
如果本地没有 OCR 环境，请让我选择本地安装、AI 视觉识别、专用 OCR API，或者跳过文字识别。
```

如果需要手动运行，可以使用：

```bash
python scripts/rebuild_pptx_from_images.py \
  --input ./slide-images \
  --output ./editable.pptx \
  --ocr auto
```

运行完成后，会生成：

```text
editable.pptx
```

这个文件就是新的 PPT。

## OCR、AI 视觉识别、专用 OCR API 有什么区别？

### 本地 OCR

本地 OCR 是在你的电脑或 Codex 环境里安装 OCR 工具，例如 Tesseract，然后用它识别图片里的文字。

优点：不用联网到第三方识别服务，批量处理成本低。  
缺点：安装可能麻烦，复杂版式和小字识别效果一般。

### AI 视觉识别

AI 视觉识别是把图片发给支持看图的 AI 模型，让模型读出文字并估计文字位置。

这个工具支持：

- OpenAI 官方 API；
- 第三方 OpenAI 兼容接口。

优点：对复杂页面、中文、自然排版通常更灵活。  
缺点：需要 API Key、联网，可能产生费用；它不是传统 OCR 服务。

### 专用 OCR API

专用 OCR API 是云服务商专门提供的文字识别接口，不是通用 AI 聊天/视觉模型。

当前这个工具已支持：

- OCR.space；
- Azure AI Vision Read OCR；
- Google Cloud Vision OCR。

优点：这是更传统、更明确的 OCR 服务路线。  
缺点：不同服务商需要的 Key、Endpoint、语言代码不一样，识别效果和价格也不同。

注意：百度 OCR、腾讯 OCR、阿里 OCR 这类接口格式不同，目前还没有内置适配器。如果要支持，需要后续单独增加。

## 如果电脑没有本地 OCR，会发生什么？

工具会询问你想用哪种方式：

```text
当前没有检测到本地 OCR 环境。
请选择：
1. 自动安装本地 OCR
2. 配置 AI 视觉识别：OpenAI 或第三方 OpenAI 兼容 API
3. 配置专用 OCR API：OCR.space、Azure Vision 或 Google Vision
4. 不识别文字，只生成图片版 PPT
```

你不用改代码，只需要根据提示选择和填写信息。

## 四种方式怎么选？

### 方式一：自动安装本地 OCR

适合你：

- 不想配置 API；
- 想在本机处理；
- 图片数量比较多；
- 可以接受识别效果一般。

运行：

```bash
python scripts/rebuild_pptx_from_images.py \
  --input ./slide-images \
  --output ./editable.pptx \
  --ocr local \
  --auto-install-local
```

### 方式二：AI 视觉识别

适合你：

- 不想安装 OCR；
- 希望复杂页面识别效果更好；
- 有 OpenAI API Key，或有第三方 OpenAI 兼容接口；
- 可以接受联网和 API 费用。

运行配置向导：

```bash
python scripts/rebuild_pptx_from_images.py \
  --input ./slide-images \
  --output ./editable.pptx \
  --ocr api \
  --api-config-wizard
```

它会让你选择：

```text
1. OpenAI 官方 API
2. 第三方 OpenAI 兼容 API / 网关 / 代理
```

如果选择第三方 OpenAI 兼容接口，你需要准备：

```text
API Base URL
API Key
模型名 Model Name
```

### 方式三：专用 OCR API

适合你：

- 你已经有专门的 OCR 服务；
- 你不想用通用 AI 视觉模型；
- 你希望走传统 OCR 服务路线。

运行配置向导：

```bash
python scripts/rebuild_pptx_from_images.py \
  --input ./slide-images \
  --output ./editable.pptx \
  --ocr ocr-api \
  --ocr-api-config-wizard
```

它会让你选择：

```text
1. OCR.space API
2. Azure AI Vision Read OCR
3. Google Cloud Vision OCR
```

不同服务需要的信息：

| 服务 | 需要准备什么 |
|---|---|
| OCR.space | API Key，语言代码可选 |
| Azure AI Vision | Endpoint URL、API Key，语言提示可选 |
| Google Cloud Vision | API Key，语言提示可选 |

如果你不知道选哪个，建议先从 `OCR.space` 开始，因为填写项最少。

### 方式四：不识别文字，只生成图片版 PPT

适合你：

- 只想快速把图片变成 PPT；
- 暂时没有 API Key；
- 后续愿意手动添加文字框。

运行：

```bash
python scripts/rebuild_pptx_from_images.py \
  --input ./slide-images \
  --output ./editable.pptx \
  --ocr none
```

## API Key 要不要保存？

工具可能会问：

```text
是否保存 API Key 到本地 .env 文件？
```

建议：

- 如果这是你自己的电脑，可以选择保存；
- 如果是公共电脑、公司共享环境、临时服务器，不建议保存；
- 不确定时，选择不保存，下一次再输入即可。

API Key 相当于账号钥匙，不要发到微信群、公开文档或公开仓库。

## 常用命令

### 自动模式

推荐大多数人使用：

```bash
python scripts/rebuild_pptx_from_images.py \
  --input ./slide-images \
  --output ./editable.pptx \
  --ocr auto
```

它会自动判断环境，并在需要时询问你。

### 本地 OCR

```bash
python scripts/rebuild_pptx_from_images.py \
  --input ./slide-images \
  --output ./editable.pptx \
  --ocr local
```

### AI 视觉识别

```bash
python scripts/rebuild_pptx_from_images.py \
  --input ./slide-images \
  --output ./editable.pptx \
  --ocr api \
  --api-config-wizard
```

### 专用 OCR API

```bash
python scripts/rebuild_pptx_from_images.py \
  --input ./slide-images \
  --output ./editable.pptx \
  --ocr ocr-api \
  --ocr-api-config-wizard
```

### 只生成图片版 PPT

```bash
python scripts/rebuild_pptx_from_images.py \
  --input ./slide-images \
  --output ./editable.pptx \
  --ocr none
```

## 生成的 PPT 文字为什么有时不准？

因为识别是从图片里猜文字，不是从原始 PPT 里读取。

这些情况可能导致识别不准：

- 图片不清晰；
- 字太小；
- 背景太复杂；
- 文字和背景颜色太接近；
- 有艺术字、阴影、渐变；
- 表格和图表里的文字很多；
- 页面里有很多装饰元素。

建议生成后重点检查：

- 标题；
- 数字；
- 专有名词；
- 人名、公司名；
- 表格里的内容；
- 页脚和备注。

## 底图里的原始文字会不会还在？

默认情况下，如果识别到了文字，工具会尝试把底图中对应位置的原始文字擦掉，然后再放上可编辑文字框。

也就是说，默认目标是：

```text
擦掉底图原文字 + 添加可编辑文字框
```

而不是：

```text
保留底图原文字 + 再叠一层文字框
```

这样可以减少“双层文字”“文字重影”的问题。

但要注意：如果原图背景很复杂，比如照片、渐变、纹理、图表底色，自动擦文字可能会留下浅色块或修补痕迹。这时可以切换成保留原图模式：

```bash
python scripts/rebuild_pptx_from_images.py \
  --input ./slide-images \
  --output ./editable.pptx \
  --ocr auto \
  --background-text-mode preserve
```

两种模式可以这样理解：

| 模式 | 效果 | 适合情况 |
|---|---|---|
| `erase` 默认 | 尝试擦掉底图原文字，再放可编辑文字 | 想减少重影，想让文字更像真正可编辑 |
| `preserve` | 完整保留原图，再叠加可编辑文字 | 背景复杂，担心擦除留下痕迹 |

## 生成后应该怎么编辑？

打开 `editable.pptx` 后，你可以：

1. 点击识别出来的文字框；
2. 修改文字；
3. 调整字体、字号和颜色；
4. 删除识别错误的文本框；
5. 手动添加缺失的文字框。

默认生成时已经会尽量清理识别到的底图文字。如果发现背景被擦得不自然，可以重新运行并加上 `--background-text-mode preserve`，改用完整底图参考。

## 常见问题

### 为什么我点不到文字？

可能有两个原因：

1. 当前页面只有图片，没有识别出文字；
2. 你点到的是底层图片，不是上面的文字框。

可以尝试在 PPT 里打开“选择窗格”，查看这一页是否有文本框。

### 为什么文字位置有点偏？

这是正常情况。

工具会尽量根据识别结果放置文字框，但图片识别出来的位置可能会有误差。生成后的 PPT 仍然建议人工检查和微调。

### 表格可以编辑吗？

一般不能完整自动还原。

工具可能会识别表格里的文字，但表格线、单元格、合并单元格等结构不一定能自动重建。复杂表格建议人工重新制作。

### 图表可以编辑吗？

一般不能。

如果图片里有柱状图、折线图、饼图，工具通常只能保留为图片。图表中的文字可能被识别成文本框，但图表本身不会自动变成 PowerPoint 里的可编辑图表。

## 最佳实践

为了得到更好的效果，建议：

1. 使用清晰图片，最好是 PNG；
2. 图片按页码命名；
3. 一次先测试 3 到 5 页；
4. 确认效果可以接受后，再处理全部页面；
5. 生成后人工检查文字；
6. 如果出现文字重影，优先确认是否使用了默认的 `--background-text-mode erase`；如果擦除痕迹明显，再改用 `preserve`。

## 最后提醒

这个工具的目标是帮你快速得到一个“可继续编辑的 PPT 初稿”。

它不能保证 100% 还原原始 PPT，因为原始 PPT 的图层、字体、动画、图表数据、形状结构在导出成图片后已经丢失。

最推荐的使用方式是：

> 先用工具生成初稿，再人工检查和微调。
