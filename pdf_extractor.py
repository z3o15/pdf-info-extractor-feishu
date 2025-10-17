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
    """提取PDF信息（按照新规则提取简介和摘要）"""
    try:
        # 直接提取前两页文本，不生成中间文件
        text_content = extract_pdf_pages_direct(pdf_path, [1])
        
        if not text_content:
            print(f"无法提取PDF内容: {pdf_path}")
            return None

        # ====== 1. 删除页码行 ======
        text_content = re.sub(r'^=+\s*第\s*\d+\s*页\s*=+$', '', text_content, flags=re.MULTILINE)
        text_content = re.sub(r'^=+\s*Page\s*\d+\s*=+$', '', text_content, flags=re.MULTILINE)

        # ====== 2. 提取简介（按照规则一和规则二） ======
        intro_content = ""
        
        # 检查是否以"Integrate Medicine"开头
        if text_content.strip().startswith("Integrate Medicine"):
            # 规则二：从内容起始处开始，到"Abstract"为止
            abstract_start = text_content.find("Abstract")
            if abstract_start != -1:
                first_part = text_content[:abstract_start]
                
                # 第二次提取：从"Article history:"开始，到期刊网址结束
                article_history_start = text_content.find("Article history:")
                if article_history_start != -1:
                    # 查找期刊网址
                    url_pattern = r'https?://[A-Za-z0-9\-\.]+\.org\.cn/?'
                    url_match = re.search(url_pattern, text_content[article_history_start:])
                    if url_match:
                        second_part = text_content[article_history_start:article_history_start + url_match.end()]
                        intro_content = first_part + "\n\n" + second_part
                    else:
                        intro_content = first_part + "\n\n" + text_content[article_history_start:]
                else:
                    intro_content = first_part
        else:
            # 规则一：从内容起始处开始，到"Abstract"为止
            abstract_start = text_content.find("Abstract")
            if abstract_start != -1:
                first_part = text_content[:abstract_start]
                
                # 第二次提取：从"*Correspondence"开始，到期刊网址结束
                correspondence_start = text_content.find("*Correspondence")
                if correspondence_start == -1:
                    correspondence_start = text_content.find("Correspondence")
                
                if correspondence_start != -1:
                    # 查找期刊网址
                    url_pattern = r'https?://[A-Za-z0-9\-\.]+\.org\.cn/?'
                    url_match = re.search(url_pattern, text_content[correspondence_start:])
                    if url_match:
                        second_part = text_content[correspondence_start:correspondence_start + url_match.end()]
                        intro_content = first_part + "\n\n" + second_part
                    else:
                        intro_content = first_part + "\n\n" + text_content[correspondence_start:]
                else:
                    intro_content = first_part

        # ====== 3. 提取摘要 ======
        abstract_content = ""
        abstract_start = text_content.find("Abstract")
        if abstract_start != -1:
            # 从"Abstract"开始，到"Key words"结束
            keywords_start = text_content.find("Key words", abstract_start)
            if keywords_start == -1:
                keywords_start = text_content.find("Keywords", abstract_start)
            if keywords_start == -1:
                keywords_start = text_content.find("KEYWORDS", abstract_start)
            if keywords_start == -1:
                keywords_start = text_content.find("关键词", abstract_start)
            
            if keywords_start != -1:
                abstract_content = text_content[abstract_start:keywords_start]
            else:
                # 如果没有找到Key words，则提取到Abstract后的合理长度
                abstract_content = text_content[abstract_start:abstract_start + 2000]
        
        # 清理摘要内容（移除Abstract标签本身及可能的冒号）
        if abstract_content.startswith("Abstract"):
            # 移除"Abstract"（8个字符）和可能跟随的冒号、空格等
            abstract_content = abstract_content[8:].strip()
            # 如果开头是冒号，继续移除
            if abstract_content.startswith(":"):
                abstract_content = abstract_content[1:].strip()
            # 如果开头是空格，继续移除
            if abstract_content.startswith(" "):
                abstract_content = abstract_content[1:].strip()

        # ====== 4. 分别修复简介与摘要格式 ======
        intro_content = fix_text_format(intro_content)
        abstract_content = fix_text_format(abstract_content)

        # ====== 5. 返回结果 ======
        return {
            '简介': intro_content,
            '摘要': abstract_content
        }

    except Exception as e:
        print(f"提取PDF信息时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return None