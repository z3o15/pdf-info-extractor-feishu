import os
import sys
import re
import pdfplumber

def fix_text_format(text: str) -> str:
    """修复PDF文本格式问题：断行、连字符、空格、作者分隔符"""
    if not text:
        return ""

    # 修复连字符断行 (neuro-\nscience → neuroscience)
    text = re.sub(r'(\w)-\s*\n\s*(\w)', r'\1\2', text)

    # 修复作者分隔符（空格 + | + 空格 → 中文逗号）
    text = re.sub(r'\s*\|\s*', '，', text)

    # 按行处理未完句合并
    lines = text.splitlines()
    merged_lines = []
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue  # 跳过空行

        if (i + 1 < len(lines)):
            next_line = lines[i + 1].strip()
            # 如果行尾不是句号、冒号、分号、问号、感叹号或大写字母结尾，则合并
            if not re.search(r'[.:;?!A-Z]$', line):
                line = f"{line} {next_line}"
                lines[i + 1] = ""  # 标记下一行为空
        merged_lines.append(line)

    text = "\n".join(l for l in merged_lines if l.strip())

    # 清理多余空格
    text = re.sub(r'\s+', ' ', text)

    # 恢复段落换行（句号或问号后加换行以便阅读）
    text = re.sub(r'([。！？!?])\s*', r'\1\n', text)

    # 清理首尾空格
    return text.strip()


def extract_pdf_pages_direct(pdf_path, pages_to_extract=[1]):
    """直接从PDF提取指定页面的文本，不生成中间文件"""
    try:
        if not os.path.exists(pdf_path):
            print(f"文件不存在: {pdf_path}")
            return None

        extracted_text = ""
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            for page_num in pages_to_extract:
                if 1 <= page_num <= total_pages:
                    page = pdf.pages[page_num - 1]
                    page_text = page.extract_text()
                    if page_text:
                        extracted_text += f"\n\n===== 第 {page_num} 页 =====\n\n"
                        extracted_text += page_text
                else:
                    print(f"页面 {page_num} 超出范围，跳过")

        return extracted_text

    except Exception as e:
        print(f"处理PDF文件时出错: {e}")
        return None


def extract_pdf_info(pdf_path):
    """提取PDF信息（正文+摘要+自动修复格式）"""
    try:
        # 直接提取前两页文本，不生成中间文件
        text_content = extract_pdf_pages_direct(pdf_path, [1])
        
        if not text_content:
            print(f"无法提取PDF内容: {pdf_path}")
            return None

        # ====== 1. 截取从头到页尾网址 ======
        end_match = re.search(r'https?://[A-Za-z0-9\-\.]+\.org\.cn/?', text_content)
        if end_match:
            text_content = text_content[:end_match.end()]

        # ====== 2. 删除页码行 ======
        text_content = re.sub(r'^=+\s*第\s*\d+\s*页\s*=+$', '', text_content, flags=re.MULTILINE)
        text_content = re.sub(r'^=+\s*Page\s*\d+\s*=+$', '', text_content, flags=re.MULTILINE)

        # ====== 3. 提取摘要 ======
        abstract_pattern = re.compile(
            r'(?:ABSTRACT|Abstract|摘要)\s*(.*?)\s*(?:Key\s*words|Keywords|KEYWORDS|关键词)',
            re.DOTALL
        )
        abstract_match = re.search(abstract_pattern, text_content)
        abstract_content = abstract_match.group(1).strip() if abstract_match else ''

        # ====== 4. 移除 Abstract ~ Correspondence 段 ======
        remove_pattern = re.compile(
            r'(ABSTRACT|Abstract|摘要)[\s\S]*?(\*?Correspondence:|Correspondence)',
            re.MULTILINE
        )
        cleaned_text = re.sub(remove_pattern, r'\2', text_content)

        # ====== 5. 分别修复摘要与正文格式 ======
        abstract_content = fix_text_format(abstract_content)
        cleaned_text = fix_text_format(cleaned_text)

        # ====== 6. 返回结果 ======
        return {
            '文件名': os.path.basename(pdf_path),  # 添加文件名字段
            '简介': cleaned_text,  # 合并后的内容作为简介
            '摘要': abstract_content
        }

    except Exception as e:
        print(f"提取PDF信息时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return None