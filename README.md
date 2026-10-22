# SD WebUI（Aki）提示词图集生成器

扫描指定文件夹里的图片，读出 sd-webui-aki（A1111 系）写进图片的提示词，
然后把「图片 + 正向提示词 + 负向提示词 + 生成参数」整理成一个 Word 文档。

## 文件说明

| 文件 | 作用 |
| --- | --- |
| `setup_env.bat` | 一键准备环境：检查/安装 Python、创建 `.venv`、设置**临时**加速源并安装依赖 |
| `run_extract.bat` | 日常使用入口：把图片文件夹拖上去，或用命令行带参数运行 |
| `sd_aki_prompt_doc.py` | 主脚本：识别元数据、提取提示词、生成 docx |
| `README.md` | 本说明 |

## 快速开始

1. 双击 `setup_env.bat`
   - 第一次会询问用哪个 pip 加速源，选 1 即可（失败会自动换阿里云 / 官方源重试）。
   - 如果电脑没有 Python，会询问是否自动下载安装 Python 3.11.9（仅当前用户，自动加 PATH）。
   - 依赖只装进脚本目录下的 `.venv`，不会污染系统 Python。
2. 把存放图片的文件夹**直接拖到 `run_extract.bat` 上**。
3. 文档默认生成在图片文件夹里，文件名为 `SD提示词图集_年月日_时分秒.docx`。

也可以双击 `run_extract.bat`，按提示粘贴文件夹路径。

## 命令行用法

```bat
run_extract.bat "D:\画图输出"
run_extract.bat "D:\画图输出" -o "D:\提示词合集.docx" --page-break
run_extract.bat "D:\画图输出" --only-aki --max-side 0
```

等价于直接调用 Python：

```bat
.venv\Scripts\python.exe sd_aki_prompt_doc.py "D:\画图输出" --page-break
```

### 参数

| 参数 | 说明 |
| --- | --- |
| `folder` | 图片文件夹（必填，支持拖拽） |
| `-o, --output` | 输出 docx 路径，默认放在图片文件夹内 |
| `--no-recursive` | 只扫描当前层，不进子文件夹 |
| `--only-aki` | 只保留 Aki/A1111 格式的图片 |
| `--max-width-cm` | 文档里图片宽度，默认 15 厘米 |
| `--max-side` | 嵌入前把图片长边缩到多少像素，默认 1600；`0` 表示原图不缩放 |
| `--page-break` | 每张图片单独一页 |
| `--no-summary-table` | 不生成“扫描结果一览”表格 |
| `--list-only` | 只在窗口里列出识别结果，不生成文档 |
| `--help` | 查看帮助 |

## 能识别什么

- **PNG**：读取 `tEXt` 块里的 `parameters`（Aki/A1111 默认写法），也兼容 `Comment` / `Description`。
- **JPG / JPEG / WebP**：读取 EXIF 的 `UserComment`（A1111 在 jpg/webp 里就是这样存的）。
- **ComfyUI**：能识别出 `prompt` 工作流，但会标注“非 Aki 格式”，不会把 JSON 当提示词写进正文。

Aki 的 `parameters` 会被拆成三块：

```
正向提示词……
Negative prompt: 负向提示词……
Steps: 28, Sampler: ..., CFG scale: 7, Seed: ..., Size: ..., Model: ...
```

生成的文档包含：

1. 标题、扫描统计（总图片数 / 含提示词数 / Aki 格式数 / 读取失败数）。
2. 扫描结果一览表：每张图是否有提示词、格式和备注。
3. 提示词正文：每张图一段，依次是图片、正向提示词、负向提示词、生成参数（提示词用等宽字体并在浅色底纹里，方便复制）。
4. 末尾列出没有提示词的图片和读取失败的文件。

## 常见问题

**为什么有些图识别不到提示词？**
提示词是写进图片文件的元数据里的。以下情况会丢：

- 用微信 / QQ / 微博等转发图片（通常会被重新压缩）；
- 截图、另存为、用看图软件“导出”；
- 经过图片压缩工具、图床压缩；
- 从网页另存为的图片。

原始出图目录里的 png 一般都没问题。

**为什么文档里图片变小了？**
默认会把长边超过 1600 像素的图缩小后再嵌入，避免 docx 过大、打开卡顿。
想保留原图画质：加 `--max-side 0`。

**重装 / 重置环境**
删掉脚本目录下的 `.venv`，重新运行 `setup_env.bat` 即可。

**想换加速源**
双击运行时按菜单选；或用参数直接指定：

```bat
setup_env.bat aliyun
setup_env.bat tencent
setup_env.bat official
```

加速源通过 `PIP_INDEX_URL` / `PIP_TRUSTED_HOST` 环境变量设置，**只在当前 bat 窗口生效**，
不会写进全局 pip 配置；窗口关掉就恢复原样。

## 环境要求

- Windows 10 / 11
- Python 3.9+（没有的话 `setup_env.bat` 可以帮你装）
- 依赖：`pillow`、`python-docx`（由 `setup_env.bat` 自动安装）
