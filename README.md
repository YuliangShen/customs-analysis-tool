# 海关数据 AI 智能分析系统

> 外贸业务场景下的真实工具：上传海关原始 Excel，自动生成多维度分析报告 + AI 文字洞察。

## 背景

做外贸时每次分析采购商数据，都要手动整理多个国家的 Excel Sheet，再用透视表逐个做排名，重复且耗时。这个工具把整个流程自动化——上传文件，几秒内输出带图表的完整报告，并调用 AI 生成中文分析结论。

## 功能

- 支持多 Sheet 合并分析（按国家分 Sheet 的海关数据格式）
- 四个维度：采购商排名 / 竞争对手分析 / 月度趋势 / HS 编码分布
- 接入千问 API，自动输出 AI 分析结论
- 输出带图表的 Excel 报告，支持 CSV 下载

## 三个版本

| 版本 | 文件 | 适用场景 |
|------|------|----------|
| 本地脚本版 | `customs_analyzer_charts.py` | 有 Python 环境，处理大量数据 |
| 网页版 | `customs_web.py` | 浏览器上传，团队共用 |
| 单文件版 | `customs_analysis.html` | 双击即用，无需安装任何环境 |

## 快速开始

```bash
pip install pandas openpyxl openai matplotlib streamlit
```

将代码中的 `your_api_key_here` 替换为千问 API Key，运行即可。

## 技术栈

Python · Pandas · Matplotlib · Streamlit · 千问 API（兼容 OpenAI 格式）· HTML / JS
