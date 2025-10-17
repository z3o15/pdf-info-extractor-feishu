import requests
import json
import time

def get_tenant_access_token(app_id, app_secret):
    """
    获取租户访问令牌
    
    Args:
        app_id (str): 应用ID
        app_secret (str): 应用密钥
        
    Returns:
        str: 租户访问令牌，失败返回None
    """
    url = 'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal'
    headers = {'Content-Type': 'application/json'}
    data = {
        'app_id': app_id,
        'app_secret': app_secret
    }
    
    try:
        response = requests.post(url, headers=headers, json=data)
        result = response.json()
        
        print(f"获取应用访问令牌 - 响应内容: {json.dumps(result, ensure_ascii=False)}")
        
        if result.get('code') == 0:
            return result.get('tenant_access_token')
        else:
            print(f"获取访问令牌失败: {result.get('msg')}")
            return None
    except Exception as e:
        print(f"获取访问令牌时发生异常: {e}")
        return None

def get_table_fields(app_token, table_id, tenant_access_token):
    """
    获取表格的字段信息
    
    Args:
        app_token (str): 多维表格的App Token
        table_id (str): 表格ID
        tenant_access_token (str): 租户访问令牌
        
    Returns:
        list: 字段信息列表，失败返回空列表
    """
    url = f'https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/fields'
    headers = {'Authorization': f'Bearer {tenant_access_token}'}
    
    try:
        response = requests.get(url, headers=headers)
        result = response.json()
        
        if result.get('code') == 0:
            fields = result.get('data', {}).get('items', [])
            field_names = [field.get('field_name') for field in fields]
            print(f"表格中的实际字段: {field_names}")
            return fields
        else:
            print(f"获取表格字段失败: {result.get('msg')}")
            return []
    except Exception as e:
        print(f"获取表格字段时发生异常: {e}")
        return []

def add_records_to_bitable(app_token, table_id, tenant_access_token, records, batch_size=20):
    """
    向多维表格添加记录
    
    Args:
        app_token (str): 多维表格的App Token
        table_id (str): 表格ID
        tenant_access_token (str): 租户访问令牌
        records (list): 要添加的记录列表
        batch_size (int): 每批次处理的记录数量
        
    Returns:
        int: 成功上传的记录数量
    """
    if not records:
        print("❌ 没有记录需要上传")
        return 0
    
    # 获取表格的实际字段结构
    fields = get_table_fields(app_token, table_id, tenant_access_token)
    if not fields:
        print("❌ 无法获取表格字段信息")
        return 0
    
    # 提取表格中存在的字段名称
    existing_field_names = [field.get('field_name') for field in fields]
    
    # 初始化成功计数
    success_count = 0
    
    # 分批处理记录
    total_batches = (len(records) + batch_size - 1) // batch_size
    
    for batch_num in range(total_batches):
        start_idx = batch_num * batch_size
        end_idx = min(start_idx + batch_size, len(records))
        batch_records = records[start_idx:end_idx]
        
        # 构建上传数据，只包含表格中存在的字段
        records_data = []
        for record in batch_records:
            fields_data = {}
            
            # 只添加表格中存在的字段
            for field_name in existing_field_names:
                if field_name in record:
                    fields_data[field_name] = record[field_name]
            
            if fields_data:
                records_data.append({"fields": fields_data})
        
        if not records_data:
            print(f"❌ 批次 {batch_num + 1}/{total_batches} 没有有效数据")
            continue
        
        # 构建请求数据
        data = {
            "records": records_data
        }
        
        url = f'https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_create'
        headers = {
            'Authorization': f'Bearer {tenant_access_token}',
            'Content-Type': 'application/json'
        }
        
        try:
            response = requests.post(url, headers=headers, json=data)
            result = response.json()
            
            print(f"响应状态码: {response.status_code}")
            
            if response.status_code == 200 and result.get('code') == 0:
                batch_success = len(records_data)
                success_count += batch_success
                print(f"✅ 批次 {batch_num + 1}/{total_batches} 上传成功，共 {batch_success} 条记录")
            else:
                error_msg = result.get('msg', '未知错误')
                print(f"❌ 批次 {batch_num + 1}/{total_batches} 上传失败: {error_msg}")
                
        except Exception as e:
            print(f"❌ 批次 {batch_num + 1}/{total_batches} 上传时发生异常: {e}")
        
        # 添加延迟避免API限制
        if batch_num < total_batches - 1:
            time.sleep(0.5)
    
    print(f"数据上传完成，成功上传 {success_count}/{len(records)} 条记录")
    return success_count

def add_records_to_wiki_table(wiki_token, table_id, tenant_access_token, records):
    """
    向知识库中的多维表格添加记录（已弃用，使用add_records_to_bitable代替）
    """
    print("⚠️  此函数已弃用，请使用 add_records_to_bitable")
    return 0

def create_new_bitable(tenant_access_token, space_id=None):
    """创建新的多维表格"""
    # 实现创建多维表格的逻辑
    pass

def create_bitable_table(tenant_access_token, app_token):
    """在多维表格中创建新表"""
    # 实现创建表格的逻辑
    pass

def create_table_fields(tenant_access_token, app_token, table_id, fields):
    """创建表格字段"""
    # 实现创建字段的逻辑
    pass

def get_existing_tables(tenant_access_token, app_token):
    """获取现有的表格列表"""
    # 实现获取表格列表的逻辑
    pass

if __name__ == "__main__":
    # 测试函数
    print("飞书上传器模块加载成功")