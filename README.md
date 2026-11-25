# 文件夹批量转 Markdown（文件名自动加当天日期）

给一个文件夹路径，把里面的文件批量转成 `.md`，输出文件名前自动加上运行当天的日期，例如：

```
资料/周报.txt          →  资料/md_2026-11-25/2026-11-25_周报.md
资料/报表.xlsx        →  资料/md_2026-11-25/2026-11-25_报表.md
资料/2026-11-20_旧.txt →  资料/md_2026-11-25/2026-11-20_旧.md   （已有日期前缀，不重复加）
```

**默认不删除、不修改任何原文件**，结果统一放在新目录里，看不到就先找 `md_<日期>` 文件夹。

---

## 一、最快上手（macOS 双击）

1. 双击本文件夹里的 **`启动转换.command`**（首次若被系统拦截：右键 → 打开，或到「系统设置 → 隐私与安全性」点「仍要打开」）。
2. 终端窗口出现提示后，**把要处理的文件夹直接拖进窗口**，回车。
3. 等转换完成，结果文件夹会自动在访达中打开；按回车关闭窗口。

也可以把参数直接传给启动器，例如：

```bash
./启动转换.command "/Users/你的用户名/Documents/资料" -r
```

> 需要 Python 3.9+。没有的话终端会给出安装提示，或执行 `brew install python3`。
> 启动器会打印实际使用的 Python 版本与路径，排查问题时可以先看这一行。
> Windows / Linux 用户直接用下面的命令行方式即可。

## 二、命令行用法

```bash
python3 files_to_md.py "/Users/你的用户名/Documents/资料"
```

不写路径也可以，脚本会提示你把文件夹拖进终端：

```bash
python3 files_to_md.py
```

常用组合：

```bash
# 连同子文件夹一起处理
python3 files_to_md.py ~/Documents/资料 -r

# 结果直接放在原文件夹里（而不是 md_<日期> 子目录）
python3 files_to_md.py ~/Documents/资料 --in-place

# 指定输出目录 + 日期改成 20261125_ 这种格式
python3 files_to_md.py ~/Documents/资料 -o ~/Desktop/转换结果 --date-format "%Y%m%d"

# 先预览会生成哪些文件，不写任何内容
python3 files_to_md.py ~/Documents/资料 --dry-run

# 转换完自动在访达打开结果目录
python3 files_to_md.py ~/Documents/资料 --open-result
```

### 参数一览

| 参数 | 作用 | 默认值 |
| --- | --- | --- |
| `folder` | 要处理的文件夹路径（可省略，运行时输入） | — |
| `-r`, `--recursive` | 包含子文件夹里的文件 | 只处理第一层 |
| `-o`, `--output` | 指定输出目录 | `源文件夹/md_<日期>/` |
| `--in-place` | 结果直接放在源文件夹里 | 关闭（与 `-o` 互斥） |
| `--date-format` | 日期前缀格式，遵循 `strftime` | `%Y-%m-%d` → `2026-11-25_` |
| `--overwrite` | 同名时覆盖 | 关闭，自动加 `-2`、`-3` |
| `--include-hidden` | 处理隐藏文件 | 关闭 |
| `--no-media-links` | 不为图片/音视频生成引用 md | 关闭，即默认生成 |
| `--dry-run` | 只列计划，不写文件 | 关闭 |
| `--open-result` | 完成后在访达打开结果目录 | 关闭（启动器默认开启） |
| `-h`, `--help` | 查看完整帮助 | — |

---

## 三、支持哪些文件

| 类别 | 扩展名 | 转换结果 |
| --- | --- | --- |
| 纯文本 | `.txt` `.md` `.log` `.srt` `.rst` `.ini` 等 | 原样写入 md（自动识别 UTF-8 / GBK / Big5 / UTF-16） |
| 代码 / 数据 | `.py` `.js` `.ts` `.java` `.sh` `.sql`、`.json` `.yaml` `.xml` 等 | 包成带语言标记的代码块，JSON 自动缩进美化 |
| 表格数据 | `.csv` `.tsv` | Markdown 表格（自动转义竖线、换行） |
| 网页 | `.html` `.htm` `.xhtml` | 标题、段落、列表、链接、图片、表格、代码块转成 md |
| Word | `.docx` `.doc` `.rtf` | 标题层级、段落、项目符号、表格 |
| Excel | `.xlsx` `.xlsm` | 每个工作表一节，转成 Markdown 表格 |
| PPT | `.pptx` | 每页一节，提取文本框文字 |
| PDF | `.pdf` | 按页分节提取文字（见下方 FAQ） |
| 电子书 / 开放文档 | `.odt` `.epub` | 提取正文 |
| 图片 / 音视频 | `.png` `.jpg` `.mp4` `.mp3` 等 | 生成一个引用原文件的 md（相对路径，移动后仍可用） |
| 其他二进制 | `.zip` `.exe` `.psd` `.xls` `.ppt` 等 | 跳过，并在结尾「跳过明细」中列出原因 |

即使扩展名被改掉，脚本也会按文件头识别 PDF 和 Office（OOXML）格式。

---

## 四、输出规则

- 文件名：`<当天日期>_<原文件名（去扩展名）>.md`；已有日期前缀（如 `2026-11-20_`、`20261120-`）不会重复添加。
- 输出目录：默认 `源文件夹/md_<当天日期>/`，可用 `-o` 或 `--in-place` 改。
- 同名不覆盖：默认生成 `...-2.md`；**重复运行且已有内容完全相同的 md 会直接跳过**，不会越堆越多。
- 转换结果目录位于源文件夹内时，重复运行会跳过该目录，不会把生成的 md 再转一遍。
- 隐藏文件、`.DS_Store`、Office 临时文件 `~$...` 默认跳过。
- 每个文件的处理结果会逐行打印，结尾给出「成功 / 跳过 / 失败」统计与失败明细；只要没有失败，退出码就是 0。

---

## 五、常见问题

**Q：双击 `.command` 没反应，或提示无法打开？**
macOS 可能拦截未签名的脚本：在访达中右键该文件 → 「打开」→「打开」；或到「系统设置 → 隐私与安全性」允许。若提示权限不足，执行一次 `chmod +x 启动转换.command`。

**Q：PDF 转出来是空的 / 提示需要 OCR？**
扫描版、图片版 PDF 本身没有文字层，需要先做 OCR（macOS 可用「预览」+ 实况文本，或 OCR 工具），再重新转换。文字版 PDF 会正常提取。

**Q：转 PDF 需要额外装工具吗？**
不需要。本地这份已内置纯 Python 的 `vendor/pypdf`（含它在 Python 3.9 下需要的 `typing_extensions`），双击就能用，不需要 pip 安装，也不需要系统编译器。只有当内置 pypdf 读某个 PDF 失败时，才会自动降级尝试 `PyMuPDF` / `pdfminer.six` / `pdftotext` / macOS PDFKit。从 GitHub 克隆的版本不含 `vendor/`，处理方式见第六节。其他可选增强：`pandoc`、`openpyxl`，装了会自动使用，**不装也全部能跑**。

**Q：所有文件都提示失败，怎么排查？**
先看终端第一行的「当前 Python」版本，再看结尾的「失败明细」——每个文件的具体报错都会逐条列出。需要 Python 3.9+；若版本过老会给出安装提示。把失败明细里的报错文字发出来即可定位。

**Q：老的 `.doc`、`.xls`、`.ppt` 怎么办？**
`.doc` / `.rtf` 可直接转；`.xls`、`.ppt` 建议先用 Office/WPS/Numbers 另存为 `.xlsx`、`.pptx` 再处理。

**Q：想换日期格式？**
`--date-format "%Y%m%d"` → `20261125_文件名.md`；`--date-format "%Y年%m月%d日"` 也可以。

---

## 六、从 GitHub 克隆后使用

仓库里只提交脚本本身，PDF 依赖用 `requirements.txt` 管理；`vendor/` 与 `__pycache__/` 不提交（你本地的这份已内置 `vendor/`，双击即可零依赖运行）。

- **macOS**：直接双击 `启动转换.command`。没装 pypdf 时脚本会自动降级使用系统自带的 PDFKit，无需安装任何东西。
- **Windows / Linux，或希望统一用纯 Python 解析 PDF**：

  ```bash
  python3 -m pip install -r requirements.txt
  python3 files_to_md.py "/你的文件夹路径"
  ```

- 想让克隆下来的版本也完全零依赖，把本地的 `vendor/` 目录一起复制进仓库目录即可，脚本会自动识别。

## 七、文件清单

| 文件 | 说明 |
| --- | --- |
| `files_to_md.py` | 主脚本，Python 3.9+ |
| `启动转换.command` | macOS 双击启动器，自动调用主脚本并打开结果目录 |
| `README.md` | 本文档 |
| `requirements.txt` | 可选的 PDF 解析依赖（macOS 可不装） |
| `.gitignore` | 忽略输出目录、缓存与工作区配置 |
| `vendor/` | 本地内置的 `pypdf` 与 `typing_extensions`（含许可证，仅本地保留，不提交仓库） |

运行前提只有：Python 3.9+（已在 macOS 自带的 Python 3.9.6 与 Python 3.13 下完整验证）。本地版 PDF 依赖已内置；克隆版在 macOS 上用系统 PDFKit，其他平台按上一节安装。
