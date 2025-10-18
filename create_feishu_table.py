import os
import pandas as pd
from tkinter import filedialog, Tk, messagebox
import json
from datetime import datetime

# 导入拆分的模块
from pdf_extractor import extract_pdf_info
from word_extractor import extract_word_info
from feishu_uploader import (
    get_tenant_access_token,
    create_new_bitable,
    create_bitable_table,
    create_table_fields,
    add_records_to_wiki_table,
    add_records_to_bitable,
    get_existing_tables
)

def get_file_extractor(file_path):
    """根据文件扩展名返回对应的解析器"""
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext == '.pdf':
        return extract_pdf_info
    elif file_ext in ['.docx', '.doc']:
        return extract_word_info
    else:
        return None

def main():
    """主函数 - 程序入口点"""
    
    # 检查配置文件
    config_file = "feishu_config.json"
    if not os.path.exists(config_file):
        print("❌ 配置文件不存在，请先配置飞书应用信息")
        print("请创建 feishu_config.json 文件并填写以下内容：")
        print("""
{
    "app_id": "your_app_id",
    "app_secret": "your_app_secret"
}
""")
        return
    
    # 读取配置
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        app_id = config.get('app_id')
        app_secret = config.get('app_secret')
        
        if not app_id or not app_secret:
            print("❌ 配置文件缺少 app_id 或 app_secret")
            return
    except Exception as e:
        print(f"❌ 读取配置文件出错: {e}")
        return

    # 选择文件或文件夹
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    
    choice = messagebox.askquestion(
        "选择处理方式", 
        "请选择处理方式：\n\n"
        "是(Y) - 选择文件夹（批量处理所有PDF和Word文件）\n"
        "否(N) - 选择单个或多个文件",
        icon='question'
    )
    
    files_to_process = []
    
    if choice == 'yes':  # 批量处理文件夹
        folder_path = filedialog.askdirectory(title="选择包含PDF和Word文件的文件夹")
        if folder_path:
            files_to_process = [
                os.path.join(folder_path, f)
                for f in os.listdir(folder_path)
                if f.lower().endswith(('.pdf', '.docx', '.doc'))
            ]
            if files_to_process:
                print(f"📁 找到 {len(files_to_process)} 个文件（PDF和Word）")
            else:
                print("❌ 文件夹中没有PDF或Word文件")
                root.destroy()
                return
    else:  # 处理单个或多个文件
        messagebox.showinfo("选择文件", "请选择要处理的PDF或Word文件")
        files_to_process = filedialog.askopenfilenames(
            title="选择PDF或Word文件",
            filetypes=[
                ("PDF文件", "*.pdf"), 
                ("Word文件", "*.docx"), 
                ("Word文件", "*.doc"),
                ("所有文件", "*.*")
            ]
        )
    
    root.destroy()
    
    if not files_to_process:
        print("❌ 未选择任何文件")
        return

    # 处理每个文件
    results = []
    
    for file_path in files_to_process:
        print(f"\n📄 正在处理: {os.path.basename(file_path)}")
        
        # 根据文件类型选择解析器
        extractor = get_file_extractor(file_path)
        if not extractor:
            print(f"❌ 不支持的文件类型: {os.path.splitext(file_path)[1]}")
            continue
        
        # 提取文件信息
        try:
            file_info = extractor(file_path)
            if file_info:
                results.append(file_info)
                print(f"✅ 成功提取信息")
            else:
                print(f"❌ 无法提取信息")
        except Exception as e:
            print(f"❌ 处理文件时出错: {str(e)}")
    
    if not results:
        print("❌ 没有成功处理任何文件")
        return

    # 保存到CSV文件
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_filename = f"PDF提取结果_{timestamp}.csv"
    
    df = pd.DataFrame(results)
    df.to_csv(csv_filename, index=False, encoding='utf-8-sig')
    print(f"\n💾 结果已保存到: {csv_filename}")
    
    # 获取访问令牌
    print("\n🔑 获取飞书访问令牌...")
    token = get_tenant_access_token(app_id, app_secret)
    if not token:
        print("❌ 获取访问令牌失败")
        return
    
    # 选择上传方式
    root = Tk()
    root.withdraw()
    
    upload_choice = messagebox.askquestion(
        "上传方式",
        "请选择上传方式：\n\n"
        "是(Y) - 上传到飞书多维表格\n"
        "否(N) - 上传到飞书知识库表格",
        icon='question'
    )
    
    root.destroy()
    
    if upload_choice == 'yes':
        # 上传到多维表格
        print("\n📊 上传到飞书多维表格...")
        
        # 创建新的多维表格
        bitable_info = create_new_bitable(token, "PDF信息提取结果")
        if not bitable_info:
            print("❌ 创建多维表格失败")
            return
        
        app_token = bitable_info['app_token']
        table_id = bitable_info['table_id']
        
        # 创建表格字段
        if not create_table_fields(token, app_token, table_id):
            print("❌ 创建表格字段失败")
            return
        
        # 添加记录
        if add_records_to_bitable(token, app_token, table_id, results):
            print("✅ 数据上传成功")
        else:
            print("❌ 数据上传失败")
    
    else:
        # 上传到知识库表格
        print("\n📚 上传到飞书知识库表格...")
        
        # 获取现有表格列表
        tables = get_existing_tables(token)
        if not tables:
            print("❌ 获取知识库表格列表失败")
            return
        
        # 选择表格
        print("\n📋 可用表格列表：")
        for i, table in enumerate(tables, 1):
            print(f"{i}. {table['name']}")
        
        try:
            choice = int(input("\n请选择要上传的表格编号: ")) - 1
            if 0 <= choice < len(tables):
                selected_table = tables[choice]
                
                # 添加记录到知识库表格
                if add_records_to_wiki_table(token, selected_table['token'], results):
                    print("✅ 数据上传成功")
                else:
                    print("❌ 数据上传失败")
            else:
                print("❌ 无效的选择")
        except ValueError:
            print("❌ 请输入有效的数字")

if __name__ == "__main__":
    main()