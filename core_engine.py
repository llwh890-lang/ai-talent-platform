"""
Core Engine 模块
负责处理与底层大模型的通信、提示词工程 (Prompt Engineering) 
以及 JSON 数据的结构化清洗和防御性解析。
"""

import os
import json
import re
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI

# ==========================================
# 基础环境与鉴权配置 (兼容本地环境与云端部署)
# ==========================================
load_dotenv()

# 优先读取 Streamlit Cloud 的 Secrets，若不存在则回退读取本地 .env 环境变量
try:
    api_key = st.secrets["RUNDAO_API_KEY"]
    base_url = st.secrets["RUNDAO_BASE_URL"]
except Exception:
    api_key = os.getenv("RUNDAO_API_KEY")
    base_url = os.getenv("RUNDAO_BASE_URL")

# 初始化全局模型客户端
client = OpenAI(
    api_key=api_key,
    base_url=base_url
)


def parse_profile_json(result_text):
    """
    暴力 JSON 解析器 (防御模型幻觉)

    截获并清洗大模型可能附带的 Markdown 标记 (如 ```json ... 
```)，
    确保系统能够稳定提取并反序列化合法的 JSON 对象。

    Args:
        result_text (str): 大模型返回的原始文本内容。

    Returns:
        dict | None: 解析成功返回字典，失败则返回 None。
    """
    if result_text.startswith("```json"):
        result_text = result_text[7:]
    elif result_text.startswith(" ```"):
       result_text = result_text[3:]
    if result_text.endswith("```"):
        result_text = result_text[:-3]

    result_text = result_text.strip()
    try:
        return json.loads(result_text)
    except json.JSONDecodeError:
        return None


def generate_student_profile(resume_text, job_text):
    """
    [静态生成] 候选人能力画像生成器 (供 B 端企业批量处理使用)

    通过大模型对简历和目标岗位进行硬核对比，动态提取能力维度并打分。
    内置 30 秒超时控制，避免批量处理时发生线程阻塞。

    Args:
        resume_text (str): 解析出的候选人简历纯文本。
        job_text (str): 企业端设定的岗位能力基准。

    Returns:
        dict | None: 包含综合评分、雷达图分值、差距和推荐岗位的结构化字典。
    """
    system_prompt = """
    你是一个资深的技术 HR 和数据分析专家。
    你的任务是根据用户提供的【学生简历】和【目标岗位要求】，进行深度对比分析。

    【硬性输出要求】
    你必须严格输出一个 JSON 格式的数据，绝对不要包含任何 Markdown 标记（如 
```json）、不要包含任何解释性文字。

    必须包含以下字段：
    {
      "name": "学生姓名",
      "match_score": 0-100的整数,
      "radar_scores": {
        "请根据目标岗位要求，动态提取最核心的4个能力维度作为key(例如: 算法基础、模型优化等)": 0-100的整数作为value
      },
      "gaps": ["差距点1", "差距点2"],
      "top_positions": [
        {"role": "推荐岗位1", "reason": "一段简短的推荐理由"},
        {"role": "推荐岗位2", "reason": "一段简短的推荐理由"},
        {"role": "推荐岗位3", "reason": "一段简短的推荐理由"},
        {"role": "推荐岗位4", "reason": "一段简短的推荐理由"},
        {"role": "推荐岗位5", "reason": "一段简短的推荐理由"}
      ]
    }
    """
    try:
        response = client.chat.completions.create(
            model="public/qwen3.6-27b",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"【学生简历】\n{resume_text}\n\n【目标岗位要求】\n{job_text}"}
            ],
            temperature=0.1,
            timeout=30.0  # 增加网络波动容错
        )
        return parse_profile_json(response.choices[0].message.content.strip())
    except Exception as e:
        print(f"B端批量解析发生网络或超时错误: {e}")
        return None


def generate_student_profile_stream(resume_text, job_text):
    """
    [流式生成] 候选人能力画像生成器 (供 C 端打字机动画使用)

    使用 Server-Sent Events (SSE) 技术将大模型的推理过程实时推送到前端。

    Args:
        resume_text (str): 学生上传的简历文本。
        job_text (str): 学生选择的目标岗位要求。

    Yields:
        str: 模型推理的增量字符串切片 (Chunks)。
    """
    system_prompt = """
    你是一个资深的技术 HR 和数据分析专家。
    你的任务是根据用户提供的【学生简历】和【目标岗位要求】，进行深度对比分析。

    【硬性要求】
    你必须且只能输出一个合法的 JSON 字典，绝对不要包含任何开头或结尾的寒暄语！
    必须包含以下键：name, match_score, radar_scores(请根据目标岗位要求，动态归纳出最核心的4个能力维度作为key，评估分数为value), gaps, top_positions(必须严格生成5个推荐岗位，内含：role, reason)。
    """
    try:
        response = client.chat.completions.create(
            model="public/qwen3.6-27b",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"【学生简历】\n{resume_text}\n\n【目标岗位要求】\n{job_text}"}
            ],
            temperature=0.1,
            stream=True,
            timeout=30.0  # 增加网络波动容错
        )
        for chunk in response:
            # 安全防御：严格校验 choices 列表是否为空，防止由于结束符分片导致程序崩溃
            if chunk.choices and len(chunk.choices) > 0:
                delta = chunk.choices[0].delta
                if hasattr(delta, "content") and delta.content is not None:
                    yield delta.content
    except Exception as e:
        # 优雅降级：网络异常时向前端抛出友好的 JSON 字符串替代报错
        yield '{"name": "网络超时", "match_score": 0, "radar_scores": {"网络异常": 0}, "gaps": ["服务器连接超时，请重试"], "top_positions": []}'


def get_local_resources():
    """获取本地教研资源库 (Mock 数据库)"""
    return [
        {"skill": "HTML/CSS", "url": "https://developer.mozilla.org/zh-CN/docs/Learn", "desc": "MDN Web Docs"},
        {"skill": "Vue", "url": "https://cn.vuejs.org/guide/introduction.html", "desc": "Vue.js Official Guide"},
        {"skill": "Git", "url": "https://git-scm.com/book/zh/v2", "desc": "Pro Git Book"},
        {"skill": "Java", "url": "https://docs.oracle.com/javase/tutorial/", "desc": "Oracle Java Tutorials"},
        {"skill": "SQL", "url": "https://www.w3schools.com/sql/", "desc": "W3Schools SQL Quiz"}
    ]


def generate_learning_plan_stream(gaps, target_job):
    """
    [流式生成] 个性化成长路径规划器

    结合诊断阶段发现的能力断层 (Gaps) 和系统内置的本地资源库，
    为学生提供结构化的四周成长路径和面试预测。

    Args:
        gaps (str): 诊断模块生成的候选人核心差距。
        target_job (str): 目标岗位。

    Yields:
        str: Markdown 格式的增量字符串切片。
    """
    local_resources = get_local_resources()
    resources_str = json.dumps(local_resources, ensure_ascii=False)

    system_prompt = """
    你是一个负责技术培训和就业指导的教务专家。
    任务：根据学生的【能力差距】和【目标岗位】，制定一份详尽的报告。

    必须严格按 Markdown 结构输出：
    # 第一部分：智能职业发展建议与路径规划
    ## 阶段目标
    ## 4周详细成长路径
    ### 第1周：基础筑基
    - 学习任务...
    - 资源链接... (从本地资源库挑选)
    - 达标标准...
    (以此类推写满4周)
    # 第二部分：精准就业指导与面试备战提示
    ## 核心简历优化建议
    ## 目标岗位面试真题预测
    """
    try:
        response = client.chat.completions.create(
            model="public/qwen3.6-27b",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user",
                 "content": f"【目标岗位】\n{target_job}\n\n【能力差距】\n{gaps}\n\n【本地资源库】\n{resources_str}"}
            ],
            temperature=0.3,
            stream=True,
            timeout=30.0
        )
        for chunk in response:
            if chunk.choices and len(chunk.choices) > 0:
                delta = chunk.choices[0].delta
                if hasattr(delta, "content") and delta.content is not None:
                    yield delta.content
    except Exception as e:
        yield f"\n\n⚠️ **报告生成失败（网络连接异常），请重新点击生成按钮。**\n<!-- 错误详情: {str(e)} -->"