"""
测试后端 API 的简单脚本
"""

import requests
import os

def test_backend_api():
    """测试后端 API 是否正常工作"""
    base_url = "http://localhost:8001"
    
    print("测试后端 API...")
    
    # 测试根路径
    try:
        response = requests.get(f"{base_url}/")
        if response.status_code == 200:
            print("✅ 根路径测试通过")
            print(f"   响应: {response.json()}")
        else:
            print(f"❌ 根路径测试失败: {response.status_code}")
    except requests.exceptions.ConnectionError:
        print("❌ 无法连接到后端服务，请确保后端服务已启动")
        return False
    
    # 测试健康检查
    try:
        response = requests.get(f"{base_url}/health")
        if response.status_code == 200:
            print("✅ 健康检查通过")
        else:
            print(f"❌ 健康检查失败: {response.status_code}")
    except Exception as e:
        print(f"❌ 健康检查异常: {e}")
    
    # 测试 Excel API 信息
    try:
        response = requests.get(f"{base_url}/api/excel/info")
        if response.status_code == 200:
            print("✅ Excel API 信息获取成功")
            print(f"   API 信息: {response.json()}")
        else:
            print(f"❌ Excel API 信息获取失败: {response.status_code}")
    except Exception as e:
        print(f"❌ Excel API 信息获取异常: {e}")
    
    # 测试文件上传 API（如果有示例文件）
    if os.path.exists("sample-data.xlsx"):
        try:
            with open("sample-data.xlsx", "rb") as f:
                files = {"file": f}
                data = {"columns": "3,5"}
                response = requests.post(f"{base_url}/api/excel/delete-columns", files=files, data=data)
                
                if response.status_code == 200:
                    print("✅ Excel 文件处理测试通过")
                    print(f"   响应大小: {len(response.content)} 字节")
                    
                    # 保存测试结果
                    with open("test-output.xlsx", "wb") as output_file:
                        output_file.write(response.content)
                    print("   测试输出已保存为 test-output.xlsx")
                else:
                    print(f"❌ Excel 文件处理测试失败: {response.status_code}")
                    print(f"   错误信息: {response.text}")
        except Exception as e:
            print(f"❌ Excel 文件处理测试异常: {e}")
    else:
        print("⚠️  未找到示例文件 sample-data.xlsx，跳过文件处理测试")
    
    print("\n测试完成！")
    return True

if __name__ == "__main__":
    test_backend_api()