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

## ✨ 三大核心技术创新点

### 1. 🔄 融入 RLHF/DPO 思想的真实数据闭环 (Data Flywheel)
打破传统 AI 工具“单向输出”的瓶颈。在企业端（B端）引入了**员工入职表现与反馈落盘引擎**。HR 的真实评价数据会进行 JSON 格式的持久化加密存储，为后续大模型进行 RAG（检索增强生成）特征对齐或 DPO（直接偏好优化）微调提供高质量的真实人类反馈样本。

### 2. 🛡️ 高鲁棒性结构化解析与容错机制
自研了底层 `parse_profile_json` 解析器与多层异常捕获机制。
- **防幻觉**：智能剥离 LLM 偶尔生成的 Markdown 标记与冗余寒暄语，确保业务引擎 100% 拿到合法 JSON。
- **防崩溃**：针对大模型网关波动，加入了 `timeout=30` 的超时控制与 SSE 流式断点续传保护；针对用户上传的纯图片扫描版 PDF 简历，加入了少字数拦截防御，防止空数据击穿后端。

### 3. 🎨 商业级动态 UI 与沉浸式多端适配 (Glassmorphism UX)
抛弃传统的“能跑就行”的粗糙 Demo UI，采用 **“全局动态毛玻璃特效 (Glassmorphism)”** 与底层 CSS 注入技术：
- 动态感知应用状态（主页展示/数据阅读），平滑切换背景透明度与文字高光。
- **双端智能自适应 (Responsive)**：通过 Echarts 的 `media` 查询，雷达图能够根据 PC 大屏和移动端窄屏自动调整半径 (Radius) 与换行策略，解决移动端文字越界、截断等恶性 UI 漏洞，实现像素级完美呈现。

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
