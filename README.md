# 🌟 AI 智能人才培养与精准就业平台 

![Python Version](https://img.shields.io/badge/Python-3.9%2B-blue.svg)
![Streamlit](https://img.shields.io/badge/Streamlit-1.32.0-FF4B4B.svg)
![LLM Engine](https://img.shields.io/badge/LLM-润道星算大模型-7B3F00.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

> **基于「润道星算大模型服务平台」构建的下一代双端 (B2C) 人才精准流转与能力赋能引擎。**

---

## 📖 项目简介

在当今瞬息万变的科技职场中，传统的“海投-盲选”招聘模式已无法满足企业的高效用人需求，学生也往往因为“能力断层”而错失良机。

本项目致力于打破这一信息壁垒，利用大语言模型（LLM）的深度推理与自然语言理解能力，打造了一个**连接高校学生（C端）与招聘企业（B端）的智能数据飞轮**：
- **对企业**：提供从“人才画像提取”、“人岗精准匹配算法”到“入职表现反馈落盘”的完整人才流转闭环。
- **对学生**：提供结构化能力诊断、自动生成多维雷达图，并基于短板（Gaps）结合本地教研智库生成个性化、可导出的《4 周专属成长规划报告》。

## ✨ 核心技术创新点

### 1. 岗位标准驱动的动态能力画像
摒弃静态题库，将真实企业岗位模型动态注入 Prompt。根据不同岗位的核心特征，动态生成候选人四维能力分数、核心差距与 TOP5 岗位推荐。
- 动态感知应用状态（主页展示/数据阅读），平滑切换背景透明度与文字高光。
- **双端智能自适应 (Responsive)**：通过 Echarts 的 `media` 查询，雷达图能够根据 PC 大屏和移动端窄屏自动调整半径 (Radius) 与换行策略，解决移动端文字越界、截断等恶性 UI 漏洞，实现像素级完美呈现。
### 2. 双端交互驱动的数据飞轮闭环
构建 B2C 双端业务架构。企业端 HR 录入的真实用人反馈（Feedback Logs）自动落盘，转化为对齐大模型的“偏好样本”，反向修正能力基准。
### 3. 面向大模型流式输出的结构化治理与异常兜底机制
针对大语言模型在实际调用中易出现的 Markdown 包裹、字段缺失或网络超时问题，在 AI 引擎层设计了严格的输出治理机制。通过字符串清洗、`json.loads` 反序列化异常捕获与网络异常兜底返回等策略，将不稳定的自然语言转化为前端组件可稳定消费的数据，大幅提升系统鲁棒性。
### 4.诊断—规划—导出一体化交付
深度集成 Pandas 与 python-docx。实现从线上流式诊断、4 周阶段式定制学习路径规划，到离线结构化报表（Word/Excel）一键导出的完整业务闭环。
### 5.隐私前置脱敏与轻量化数据治理
在文档解析后、模型调用前设置隐私前置脱敏环节，自动识别并替换简历数据中常见的手机号、身份证号、邮箱等敏感信息。同时，采用轻量化 JSON 文件管理岗位模型和反馈日志，在保证极低部署成本的前提下，坚守“脱敏特征上云端，隐私明文留本地”的风控底线。

---

## 🛠️ 系统核心功能架构

### 🏢 B 端（企业人才管理门户）
- **动态岗位模型库**：HR 可自定义增删改岗位能力四维基准（持久化至本地 JSON）。
- **批量智能清洗与排位**：支持批量拖入数百份 PDF 简历，大模型并发解析，最终输出包含综合匹配度、核心优势与首要短板的顺位清单。
- **业务表现反馈追踪**：录入员工真实表现数据，反哺 AI 诊断大模型。

### 🎓 C 端（学生成长诊断平台）
- **多模态简历提炼**：支持 PDF 文档一键解析与纯文本手动双向核对。
- **流式结构化报表 (SSE)**：打字机效果实时呈现 AI 推理过程，生成包含雷达图的综合匹配诊断书。
- **智能教务路径规划**：基于大模型提炼的能力 Gap，结合本地教研智库（如 MDN、Oracle 教程等），生成详细到每周的 Markdown 学习规划。
- **全格式数据导出**：支持一键导出包含样式排版的 `.docx` 规划报告与包含数据指标的 `.xlsx` 画像报表。

---

## 💻 技术栈 (Tech Stack)
- **核心引擎**：Python 3.9+
- **前端与全栈框架**：Streamlit 1.32+
- **AI / LLM 接入**：OpenAI 兼容接口 (基于润道星算大模型 Qwen3.6-27B)
- **数据可视化**：Apache ECharts (streamlit-echarts)
- **文档处理**：PyPDF2, python-docx, pandas, openpyxl
---

## 🚀 本地快速运行与部署指南

### 1. 克隆项目
```bash
git clone [https://github.com/llwh890-lang/ai-talent-platform.git](https://github.com/YourUsername/ai-talent-platform.git)
cd ai-talent-platform
```
### 2. 创建并激活虚拟环境
```bash
conda create -n ai-talent python=3.10
conda activate ai-talent
```
### 3. 安装环境依赖
```bash
pip install -r requirements.txt
```
### 4. 配置核心环境变量
```Ini, TOML
RUNDAO_API_KEY="您的API_KEY"
RUNDAO_BASE_URL="您的模型网关地址"
```
### 5. 启动系统
```bash
streamlit run app.py
```
## 📁 核心目录结构
```Plaintext
├── app.py                  # Streamlit 前端视图层与全局状态管理 (多端自适应渲染)
├── core_engine.py          # LLM 通信引擎、Prompt 构建与防御性 JSON 提取
├── requirements.txt        # 生产环境依赖清单
├── .env                    # 环境变量配置文件 (部署时请勿提交至公网)
└── data/
    ├── job_models.json     # 企业动态岗位基准库 (持久化层)
    └── feedback_logs.json  # 业务评价反馈落盘日志 (数据飞轮层)
```
## 平台部署
部署链接：http://47.97.244.9:8501
<img width="536" height="88" alt="image" src="https://github.com/user-attachments/assets/4497495e-1823-4fb8-b850-87739d51132d" />
