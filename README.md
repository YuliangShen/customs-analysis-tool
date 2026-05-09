# 🤖 AI 自动化工具集

> 基于 Python + 千问 API 构建的两款 AI 自动化工具，覆盖外贸数据分析与谷歌广告文案生成场景。

## 📦 项目一：海关数据 AI 智能分析系统

上传海关原始 Excel 数据，自动生成多维度分析报告，支持 AI 洞察与一键下载。

### 功能
- 支持多 Sheet（按国家分 Sheet），自动合并分析
- 4 个维度：采购商排名 / 竞争对手分析 / 月度趋势 / HS 编码分布
- 接入千问 API，自动生成 AI 分析结论
- 输出带图表的 Excel 报告

### 三个版本

| 版本 | 文件 | 使用方式 |
|------|------|---------|
| 本地脚本版 | `customs_analyzer_charts.py` | `python customs_analyzer_charts.py` |
| 网页版 | `customs_web.py` | `streamlit run customs_web.py` |
| 单文件版 | `customs_analysis.html` | 双击浏览器打开，无需安装 |

### 快速开始

```bash
pip install pandas openpyxl openai matplotlib streamlit
```

在脚本中将 `your_api_key_here` 替换为真实的千问 API Key，运行即可。

---

## 📢 项目二：Google Ads 文案批量生成器

输入产品信息，AI 批量生成符合 Google Ads 字符限制的广告标题和描述，一键导出 CSV 直接导入 Google Ads Editor。

### 功能
- 输入产品名称、卖点、目标受众，AI 自动生成多组广告文案
- 实时字符计数，标红超过限制的标题（标题 ≤30字符，描述 ≤90字符）
- Google 广告预览效果展示
- 导出标准 CSV，兼容 Google Ads Editor 批量导入格式
- 展示 AI 返回的原始 JSON，可学习结构化输出原理

### 使用方式

直接双击 `google_ads_generator.html` 用浏览器打开，填入千问 API Key 即可使用，无需安装任何环境。

---

## 🛠 技术栈

Python · Pandas · Matplotlib · Streamlit · 千问 API（兼容 OpenAI 格式）· HTML/JS · SheetJS · Chart.js

## ⚠️ 注意

API Key 请勿硬写在代码中，使用前将 `your_api_key_here` 替换为真实 Key（本地运行），GitHub 上保持占位符。
