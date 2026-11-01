# 金种子研究工具

保存调研抓取、数据整理、Word 报告、工作簿及文档排版工具，并包含重建现有报告所需的自编文本模板和结构化公开研究输入。下载网页、原始参考文档、生成的 DOCX/XLSX/PDF、预览截图和缓存不入库。

## 重建报告

Python 3.12：`python -m pip install -r requirements.lock.txt`。依次运行：

```sh
python build_report.py
python research/interactive-teacher-20261025/build_report.py
python research/gold-study-20261025/build_study.py
python research/gold-reality-20261025/build_report.py
python research/candidate-plans-20261025/build_report.py
python research/teacher-alumni-20261026/build_documents.py
python research/teacher-alumni-20261026/build_brief.py
python research/teacher-alumni-20261026/fix_brief_mobile.py
```

结果在 `outputs/`。结构化输入固定本次已使用的公开资料时点，不需要重新抓取可能变化的网页。需要更新研究时才运行抓取/提取脚本；它们需要公开网络及对应原始材料。

工作簿脚本需要 Node.js 和 `@oai/artifact-tool` 2.8.106，该库由 Codex 文档运行环境提供。在该环境中让 `node_modules` 指向其依赖目录后运行 `node build_workbook.mjs`；普通电脑需先取得相同运行库，不将其第三方程序包复制进源码仓库。

布局字体由宿主机提供；生成文档可复建，但 PDF 渲染、字体和创建时间会影响二进制摘要。源码同步清单见 [SOURCE_STATE.md](SOURCE_STATE.md)。
