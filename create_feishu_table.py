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
    
    # 询问用户选择方式
    choice = messagebox.askquestion("选择处理方式", 
                                   "请选择处理方式：\n\n"
                                   "是(Y) - 选择文件夹（批量处理所有PDF文件）\n"
                                   "否(N) - 选择单个或多个PDF文件",
                                   icon='question')
    
    pdf_files = []
    
    if choice == 'yes':  # 选择文件夹
        folder_path = filedialog.askdirectory(
            title="选择包含PDF文件的文件夹"
        )
        
        if folder_path:
            # 查找文件夹中的所有PDF文件
            for file in os.listdir(folder_path):
                if file.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(folder_path, file))
            
            if pdf_files:
                print(f"📁 找到文件夹中的PDF文件: {len(pdf_files)} 个")
                for pdf_file in pdf_files:
                    print(f"   - {os.path.basename(pdf_file)}")
            else:
                print("❌ 选择的文件夹中没有找到PDF文件")
                root.destroy()
                return
    else:  # 选择单个或多个文件
        # 默认选择当前目录下的PDF文件
        current_dir = os.getcwd()
        default_pdf = None
        
        # 查找当前目录下的PDF文件
        for file in os.listdir(current_dir):
            if file.lower().endswith('.pdf'):
                default_pdf = os.path.join(current_dir, file)
                break
        
        # 如果找到默认PDF文件，直接使用它
        if default_pdf:
            pdf_files = [default_pdf]
            print(f"🔍 找到默认PDF文件: {os.path.basename(default_pdf)}")
        else:
            # 否则打开文件选择对话框
            messagebox.showinfo("选择文件", "请选择要处理的PDF文件")
            pdf_files = filedialog.askopenfilenames(
                title="选择PDF文件",
                filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
            )
    
    root.destroy()
    
    if not pdf_files:
        print("❌ 未选择任何PDF文件")
        return
    
    # 处理PDF文件 - 使用导入的函数
    results = []
    for pdf_file in pdf_files:
        print(f"🔍 正在处理文件: {os.path.basename(pdf_file)}")
        result = extract_pdf_info(pdf_file)
        if result:
            results.append(result)
    
    if not results:
        print("❌ 没有成功处理任何文件")
        return
    
    # 保存结果到CSV文件
    csv_file = "PDF提取结果.csv"
    try:
        # 如果文件已存在，先尝试删除
        if os.path.exists(csv_file):
            os.remove(csv_file)
        
        df = pd.DataFrame(results)
        df.to_csv(csv_file, index=False, encoding='utf-8-sig')
        print(f"结果已保存到 {csv_file}")
    except Exception as e:
        print(f"保存CSV文件时出错: {e}")
        # 尝试使用不同的文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_file = f"PDF提取结果_{timestamp}.csv"
        df = pd.DataFrame(results)
        df.to_csv(csv_file, index=False, encoding='utf-8-sig')
        print(f"结果已保存到 {csv_file}")
    
    # 检查是否上传到飞书
    upload_to_feishu = config.get("upload_to_feishu", False)
    
    if upload_to_feishu:
        # 使用租户访问令牌（不需要用户访问令牌）
        user_access_token = tenant_access_token
        
        # 获取配置中的表格信息
        app_token = config.get("app_token")
        table_id = config.get("table_id")
        
        # 检查表格是否存在
        table_exists = False
        if app_token and table_id:
            # 测试表格是否存在
            import requests
            url = f'https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables'
            headers = {'Authorization': f'Bearer {tenant_access_token}'}
            
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                result = response.json()
                if result.get('code') == 0:
                    tables = result.get('data', {}).get('items', [])
                    table_ids = [table.get('table_id') for table in tables]
                    if table_id in table_ids:
                        table_exists = True
        
        # 使用正确的多维表格API访问知识库中的表格
        print(f"📋 使用多维表格API访问表格:")
        print(f"   App Token: {app_token}")
        print(f"   Table ID: {table_id}")
        
        # 直接使用多维表格API上传数据
        success_count = add_records_to_bitable(app_token, table_id, tenant_access_token, results)
        
        if success_count > 0:
            print(f"✅ 成功上传 {success_count} 条记录到飞书多维表格")
        else:
            print("❌ 数据上传失败，请检查表格字段配置")

if __name__ == "__main__":
    main()