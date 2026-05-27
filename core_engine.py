import os
import json
import re
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI

# 基础配置加载 (兼容本地和云端)
load_dotenv()

# 优先使用云端 Secrets，如果找不到再使用本地环境变量
try:
    api_key = st.secrets["RUNDAO_API_KEY"]
    base_url = st.secrets["RUNDAO_BASE_URL"]
except Exception:
    api_key = os.getenv("RUNDAO_API_KEY")
    base_url = os.getenv("RUNDAO_BASE_URL")

client = OpenAI(
    api_key=api_key,
    base_url=base_url
)


def parse_profile_json(result_text):
    """解析 AI 输出的 JSON 字符串"""
    if result_text.startswith("```json"):
        result_text = result_text[7:]
    elif result_text.startswith("```"):
        result_text = result_text[3:]
    if result_text.endswith("```"):
        result_text = result_text[:-3]

    result_text = result_text.strip()
    try:
        return json.loads(result_text)
    except json.JSONDecodeError:
        return None


def generate_student_profile(resume_text, job_text):
    """静态生成，供 B 端企业批量处理使用"""
    system_prompt = """
    你是一个资深的技术 HR 和数据分析专家。
    你的任务是根据用户提供的【学生简历】和【目标岗位要求】，进行深度对比分析。

    【硬性输出要求】
    你必须严格输出一个 JSON 格式的数据，绝对不要包含 any Markdown 标记（如 ```json）、不要包含任何解释性文字。

    必须包含以下字段：
    {
      "name": "学生姓名",
      "match_score": 0-100的整数,
      "radar_scores": {
        "前端基础": 0-100的整数,
        "框架应用": 0-100的整数,
        "工程化能力": 0-100的整数,
        "计算机基础": 0-100的整数
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
            temperature=0.1
        )
        return parse_profile_json(response.choices[0].message.content.strip())
    except Exception as e:
        print(f"Error: {e}")
        return None


def generate_student_profile_stream(resume_text, job_text):
    """流式生成器，供 C 端打字机动画使用"""
    system_prompt = """
    你是一个资深的技术 HR 和数据分析专家。
    你的任务是根据用户提供的【学生简历】和【目标岗位要求】，进行深度对比分析。
    必须严格输出 JSON 格式，包含：name, match_score, radar_scores(前端基础/框架应用/工程化能力/计算机基础), gaps, top_positions(包含 role 和 reason)。
    """
    response = client.chat.completions.create(
        model="public/qwen3.6-27b",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"【学生简历】\n{resume_text}\n\n【目标岗位要求】\n{job_text}"}
        ],
        temperature=0.1,
        stream=True
    )
    for chunk in response:
        # 安全防御：严格校验 choices 列表是否为空，防止由于结束符分片导致程序崩溃
        if chunk.choices and len(chunk.choices) > 0:
            delta = chunk.choices[0].delta
            if hasattr(delta, "content") and delta.content is not None:
                yield delta.content


def get_local_resources():
    return [
        {"skill": "HTML/CSS", "url": "https://developer.mozilla.org/zh-CN/docs/Learn", "desc": "MDN Web Docs"},
        {"skill": "Vue", "url": "https://cn.vuejs.org/guide/introduction.html", "desc": "Vue.js Official Guide"},
        {"skill": "Git", "url": "https://git-scm.com/book/zh/v2", "desc": "Pro Git Book"},
        {"skill": "Java", "url": "https://docs.oracle.com/javase/tutorial/", "desc": "Oracle Java Tutorials"},
        {"skill": "SQL", "url": "https://www.w3schools.com/sql/", "desc": "W3Schools SQL Quiz"}
    ]


def generate_learning_plan_stream(gaps, target_job):
    """流式计划生成器"""
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
    response = client.chat.completions.create(
        model="public/qwen3.6-27b",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user",
             "content": f"【目标岗位】\n{target_job}\n\n【能力差距】\n{gaps}\n\n【本地资源库】\n{resources_str}"}
        ],
        temperature=0.3,
        stream=True
    )
    for chunk in response:
        # 安全防御：严格校验分片结构
        if chunk.choices and len(chunk.choices) > 0:
            delta = chunk.choices[0].delta
            if hasattr(delta, "content") and delta.content is not None:
                yield delta.content