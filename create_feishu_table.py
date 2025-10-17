import os
import pandas as pd
from tkinter import filedialog, Tk, messagebox
import json
from datetime import datetime

# 导入拆分的模块
from pdf_extractor import extract_pdf_info
from feishu_uploader import (
    get_tenant_access_token,
    create_new_bitable,
    create_bitable_table,
    create_table_fields,
    add_records_to_wiki_table,
    add_records_to_bitable,
    get_existing_tables
)

def main():
    """主函数 - 程序入口点"""
    # 读取配置文件
    config_path = "feishu_config.json"
    if not os.path.exists(config_path):
        print(f"❌ 配置文件 {config_path} 不存在")
        return
    
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    app_id = config.get("app_id")
    app_secret = config.get("app_secret")
    
    if not app_id or not app_secret:
        print("❌ 配置文件中缺少 app_id 或 app_secret")
        return
    
    # 获取租户访问令牌
    tenant_access_token = get_tenant_access_token(app_id, app_secret)
    if not tenant_access_token:
        print("❌ 获取租户访问令牌失败")
        return
    
    # 选择PDF文件或文件夹
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    
    choice = messagebox.askquestion(
        "选择处理方式", 
        "请选择处理方式：\n\n"
        "是(Y) - 选择文件夹（批量处理所有PDF文件）\n"
        "否(N) - 选择单个或多个PDF文件",
        icon='question'
    )
    
    pdf_files = []
    
    if choice == 'yes':  # 批量处理文件夹
        folder_path = filedialog.askdirectory(title="选择包含PDF文件的文件夹")
        if folder_path:
            pdf_files = [
                os.path.join(folder_path, f)
                for f in os.listdir(folder_path)
                if f.lower().endswith('.pdf')
            ]
            if pdf_files:
                print(f"📁 找到 {len(pdf_files)} 个PDF文件")
            else:
                print("❌ 文件夹中没有PDF文件")
                root.destroy()
                return
    else:  # 处理单个或多个文件
        messagebox.showinfo("选择文件", "请选择要处理的PDF文件")
        pdf_files = filedialog.askopenfilenames(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
    
    root.destroy()
    
    if not pdf_files:
        print("❌ 未选择任何PDF文件")
        return
    
    # 提取 PDF 信息（不保留文件名）
    results = []
    for pdf_file in pdf_files:
        print(f"🔍 正在处理: {os.path.basename(pdf_file)}")
        result = extract_pdf_info(pdf_file)
        if result:
            results.append(result)
    
    if not results:
        print("❌ 没有成功处理任何文件")
        return
    
    # 保存结果到 CSV
    csv_file = "PDF提取结果.csv"
    try:
        if os.path.exists(csv_file):
            os.remove(csv_file)
        df = pd.DataFrame(results)
        df.to_csv(csv_file, index=False, encoding='utf-8-sig')
        print(f"✅ 结果已保存到 {csv_file}")
    except Exception as e:
        print(f"⚠️ 保存CSV文件出错: {e}")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_file = f"PDF提取结果_{timestamp}.csv"
        pd.DataFrame(results).to_csv(csv_file, index=False, encoding='utf-8-sig')
        print(f"✅ 结果已保存到 {csv_file}")
    
    # 上传到飞书多维表格（如启用）
    if config.get("upload_to_feishu", False):
        app_token = config.get("app_token")
        table_id = config.get("table_id")
        if not app_token or not table_id:
            print("⚠️ 缺少表格配置 app_token 或 table_id，跳过上传")
            return
        
        print(f"📋 上传至飞书多维表格（App Token: {app_token}, Table ID: {table_id}）")
        success_count = add_records_to_bitable(app_token, table_id, tenant_access_token, results)
        
        if success_count > 0:
            print(f"✅ 成功上传 {success_count} 条记录")
        else:
            print("❌ 上传失败，请检查表格字段配置")

if __name__ == "__main__":
    main()