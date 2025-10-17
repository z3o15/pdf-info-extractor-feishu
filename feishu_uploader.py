import requests
import json
from datetime import datetime

def get_tenant_access_token(app_id: str, app_secret: str) -> str:
    """获取租户访问令牌"""
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    headers = {"Content-Type": "application/json; charset=utf-8"}
    data = {"app_id": app_id, "app_secret": app_secret}
    
    try:
        response = requests.post(url, headers=headers, json=data)
        response.raise_for_status()
        result = response.json()
        
        if result.get("code") == 0:
            return result.get("tenant_access_token")
        else:
            print(f"❌ 获取租户访问令牌失败: {result.get('msg', 'Unknown error')}")
            return ""
    except Exception as e:
        print(f"❌ 请求租户访问令牌出错: {e}")
        return ""

def create_new_bitable(app_token: str, tenant_access_token: str, name: str = "PDF信息提取结果") -> str:
    """创建新的多维表格"""
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables"
    headers = {
        "Authorization": f"Bearer {tenant_access_token}",
        "Content-Type": "application/json; charset=utf-8"
    }
    data = {"table": {"name": name}}
    
    try:
        response = requests.post(url, headers=headers, json=data)
        response.raise_for_status()
        result = response.json()
        
        if result.get("code") == 0:
            table_id = result.get("data", {}).get("table", {}).get("table_id")
            print(f"✅ 创建多维表格成功，Table ID: {table_id}")
            return table_id
        else:
            print(f"❌ 创建多维表格失败: {result.get('msg', 'Unknown error')}")
            return ""
    except Exception as e:
        print(f"❌ 创建多维表格出错: {e}")
        return ""

def create_bitable_table(app_token: str, table_id: str, tenant_access_token: str) -> bool:
    """创建多维表格的数据表"""
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/fields"
    headers = {
        "Authorization": f"Bearer {tenant_access_token}",
        "Content-Type": "application/json; charset=utf-8"
    }
    
    # 定义表格字段（仅保留简介和摘要）
    fields = [
        {
            "field_name": "简介",
            "type": 1,  # 文本类型
            "property": {"formatter": "text"}
        },
        {
            "field_name": "摘要",
            "type": 1,  # 文本类型
            "property": {"formatter": "text"}
        }
    ]
    
    success_count = 0
    for field in fields:
        try:
            response = requests.post(url, headers=headers, json=field)
            response.raise_for_status()
            result = response.json()
            
            if result.get("code") == 0:
                success_count += 1
                print(f"✅ 创建字段 '{field['field_name']}' 成功")
            else:
                print(f"❌ 创建字段 '{field['field_name']}' 失败: {result.get('msg', 'Unknown error')}")
        except Exception as e:
            print(f"❌ 创建字段 '{field['field_name']}' 出错: {e}")
    
    return success_count == len(fields)

def create_table_fields(app_token: str, table_id: str, tenant_access_token: str) -> bool:
    """创建表格字段（兼容性函数）"""
    return create_bitable_table(app_token, table_id, tenant_access_token)

def add_records_to_wiki_table(app_token: str, table_id: str, tenant_access_token: str, records: list) -> int:
    """添加记录到知识库表格（兼容性函数）"""
    return add_records_to_bitable(app_token, table_id, tenant_access_token, records)

def add_records_to_bitable(app_token: str, table_id: str, tenant_access_token: str, records: list) -> int:
    """添加记录到多维表格"""
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_create"
    headers = {
        "Authorization": f"Bearer {tenant_access_token}",
        "Content-Type": "application/json; charset=utf-8"
    }
    
    # 构建记录数据（仅包含简介和摘要）
    records_data = []
    for result in records:
        record = {}
        
        # 添加简介字段
        if "简介" in result:
            record["简介"] = result["简介"]
        
        # 添加摘要字段
        if "摘要" in result:
            record["摘要"] = result["摘要"]
        
        if record:
            records_data.append({"fields": record})
    
    if not records_data:
        print("⚠️ 没有有效记录可上传")
        return 0
    
    data = {"records": records_data}
    
    try:
        response = requests.post(url, headers=headers, json=data)
        response.raise_for_status()
        result = response.json()
        
        if result.get("code") == 0:
            success_count = len(result.get("data", {}).get("records", []))
            print(f"✅ 成功上传 {success_count} 条记录")
            return success_count
        else:
            print(f"❌ 上传记录失败: {result.get('msg', 'Unknown error')}")
            return 0
    except Exception as e:
        print(f"❌ 上传记录出错: {e}")
        return 0

def get_existing_tables(app_token: str, tenant_access_token: str) -> list:
    """获取现有的表格列表"""
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables"
    headers = {"Authorization": f"Bearer {tenant_access_token}"}
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        result = response.json()
        
        if result.get("code") == 0:
            return result.get("data", {}).get("items", [])
        else:
            print(f"❌ 获取表格列表失败: {result.get('msg', 'Unknown error')}")
            return []
    except Exception as e:
        print(f"❌ 获取表格列表出错: {e}")
        return []