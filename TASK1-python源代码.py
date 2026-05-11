# pdf读取
import re
import sys
import fitz  # PyMuPDF
 
 
def extract_text_from_pdf(pdf_path: str) -> str:
    """
    使用 PyMuPDF 从 PDF 文件中提取并清洗文本。
 
    Args:
        pdf_path: PDF 文件的路径
 
    Returns:
        处理好的完整文本字符串
    """
    all_text = []
 
    # 打开 PDF 文件
    pdf_path = r"C:\Users\15239\Desktop\Medical Task 1pdf\A case of portal vein recanalization and symptomatic heart failure.pdf"
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise FileNotFoundError(f"无法打开 PDF 文件: {pdf_path}\n错误信息: {e}")
 
    print(f"[INFO] 成功打开 PDF，共 {doc.page_count} 页")
 
    # 遍历每一页，提取文本
    for page_num in range(doc.page_count):
        page = doc[page_num]
        # 使用 "text" 模式提取，保留段落结构
        page_text = page.get_text("text")
        if page_text.strip():
            all_text.append(page_text)
 
    doc.close()
 
    # 合并所有页面文本
    raw_text = "\n".join(all_text)
 
    # ── 文本清洗 ──────────────────────────────────────────────
    # 1. 修复常见乱码符号（PDF 中常见的替换字符）
    cleaned = raw_text.replace("\ufffd", "")       # 替换 Unicode 替换字符 ?
    cleaned = cleaned.replace("\x00", "")          # 移除空字节
 
    # 2. 将多个连续空行压缩为最多一个空行
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
 
    # 3. 移除行内多余的空格（保留段落换行）
    cleaned = re.sub(r"[ \t]+", " ", cleaned)
 
    # 4. 去除每行首尾多余空白
    lines = [line.strip() for line in cleaned.splitlines()]
    cleaned = "\n".join(lines)
 
    # 5. 去除整体首尾非法空白
    cleaned = cleaned.strip()
 
    return cleaned
 
 
if __name__ == "__main__":
    # 默认路径，可通过命令行参数覆盖
    default_path = "A case of portal vein recanalization and symptomatic heart failure.pdf"
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else default_path
 
    print(f"[INFO] 正在读取: {pdf_path}")
    text = extract_text_from_pdf(pdf_path)
 
    print(f"[INFO] 提取完成，总字符数: {len(text)}")
    print("\n" + "=" * 60)
    print("【文本预览：前 500 个字符】")
    print("=" * 60)
    print(text[:500])
    print("=" * 60)

#调用API

import json
import fitz                        # PyMuPDF，用于读取 PDF
from openai import OpenAI          # openai SDK（兼容阿里云百炼接口）
 
# ── 配置区：填入你的 API Key ──────────────────────────────────────
API_KEY  = "在此处填入你的API_KEY"   # 阿里云百炼 API Key
BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"  # 百炼兼容端点
MODEL    = "qwen-plus"                     # 可选: qwen-turbo / qwen-plus / qwen-max
# ──────────────────────────────────────────────────────────────────
 
# System Prompt：要求模型只返回合法 JSON，不输出任何多余内容
SYSTEM_PROMPT = (
    "你是一位资深医疗信息提取专家。"
    "请从患者病例中提取：患者基本信息、主要症状、既往史、诊断结果、治疗方案。"
    "请仅返回合法的 JSON，不要输出任何 markdown 标记或解释文字。"
)

# —— Step 1: 清洗文本已完成
# ── Step 2：调用大模型 API，提取结构化医学实体 ─────────────────────
def extract_medical_data(text: str) -> dict:
    client = OpenAI(api_key=API_KEY, base_url=BASE_URL)    # 初始化客户端
 
    print("[INFO] 正在调用大模型 API，请稍候...")
 
    try:
        response = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},   # 系统角色设定
                {"role": "user",   "content": text},             # 传入病例全文
            ],
            temperature=0.1,       # 低随机性，保证输出格式稳定
            max_tokens=2048,       # 最大返回 token 数
        )
    except Exception as e:
        # API 调用本身失败（网络错误、Key 无效等）
        print(f"[ERROR] API 调用失败: {e}")
        raise
 
    raw = response.choices[0].message.content              # 取出模型返回的字符串
    print("[INFO] API 返回成功，正在解析 JSON...")
 
    # 清理模型可能误加的 markdown 代码块标记
    raw = raw.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
 
    try:
        result = json.loads(raw)                           # 解析为 Python 字典
    except json.JSONDecodeError as e:
        print(f"[ERROR] JSON 解析失败: {e}")
        print(f"[DEBUG] 模型原始返回内容:\n{raw}")         # 打印原文便于排查
        raise
 
    return result
 
 
# ── Step 3：将结构化结果保存为 JSON 文件 ──────────────────────────
def save_to_json(data: dict, output_path: str = "result.json") -> None:
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)   # 中文不转义，格式化缩进
    print(f"[INFO] 结果已保存至: {output_path}")
 
 
# ── Main：串联完整流程 ─────────────────────────────────────────────
def main():
    PDF_PATH = "A case of portal vein recanalization and symptomatic heart failure.pdf"
 
    # 1. 读取并清洗 PDF 文本
    print(f"[INFO] 正在读取 PDF: {PDF_PATH}")
    text = extract_text_from_pdf(PDF_PATH)
    print(f"[INFO] 文本提取完成，共 {len(text)} 个字符")
    print("── 文本预览（前 300 字）──")
    print(text[:300])
    print("────────────────────────\n")
 
    # 2. 调用大模型提取医学实体
    medical_data = extract_medical_data(text)
 
    # 3. 保存结果
    save_to_json(medical_data, "result.json")
 
    # 4. 控制台预览结果
    print("\n── 提取结果预览 ──")
    print(json.dumps(medical_data, ensure_ascii=False, indent=2))
 
 
if __name__ == "__main__":
    main()
 
