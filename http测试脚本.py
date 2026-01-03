# -*- coding: utf-8 -*-
"""
test_http_service.py - NX HTTP 服务测试脚本
===========================================
用途: 模拟 MCP 客户端，测试所有 API 端点
"""

import json
import time
import urllib.error
import urllib.request
from datetime import datetime

# ==================== 配置 ====================

NX_SERVICE_URL = "http://127.0.0.1:8765"
TEST_PART_PATH = r"C:/Projects/NC/DIE-03-_dwg.prt"  # ⚠️ 修改为实际路径

# ==================== HTTP 客户端 ====================

class TestClient:
    """测试客户端"""
    
    def __init__(self, base_url: str):
        self.base_url = base_url
        self.test_count = 0
        self.success_count = 0
        self.fail_count = 0
    
    def print_header(self, text: str):
        """打印标题"""
        print("\n" + "="*70)
        print(f"  {text}")
        print("="*70)
    
    def print_step(self, step_num: int, text: str):
        """打印步骤"""
        print(f"\n[步骤 {step_num}] {text}")
        print("-" * 70)
    
    def print_result(self, success: bool, result: dict):
        """打印结果"""
        if success:
            print("✅ 成功")
            if "data" in result:
                self._print_data(result["data"])
        else:
            print(f"❌ 失败: {result.get('error', '未知错误')}")
            if "hint" in result:
                print(f"💡 提示: {result['hint']}")
    
    def _print_data(self, data: dict, indent: int = 1):
        """递归打印数据"""
        prefix = "  " * indent
        for key, value in data.items():
            if isinstance(value, dict):
                print(f"{prefix}• {key}:")
                self._print_data(value, indent + 1)
            elif isinstance(value, list):
                if len(value) > 5:
                    print(f"{prefix}• {key}: {value[:5]}... (共{len(value)}个)")
                else:
                    print(f"{prefix}• {key}: {value}")
            else:
                print(f"{prefix}• {key}: {value}")
    
    def call_api(self, endpoint: str, params: dict = None, description: str = "") -> dict:
        """调用 API"""
        self.test_count += 1
        
        if description:
            print(f"📡 {description}...")
        
        url = f"{self.base_url}{endpoint}"
        data = json.dumps(params or {}).encode('utf-8')
        
        try:
            req = urllib.request.Request(
                url, 
                data=data, 
                headers={"Content-Type": "application/json"}, 
                method="POST"
            )
            
            with urllib.request.urlopen(req, timeout=30) as response:
                result = json.loads(response.read().decode('utf-8'))
                
                success = result.get("success", False)
                if success:
                    self.success_count += 1
                else:
                    self.fail_count += 1
                
                self.print_result(success, result)
                return result
        
        except urllib.error.URLError as e:
            self.fail_count += 1
            error_result = {"success": False, "error": f"网络错误: {str(e)}"}
            self.print_result(False, error_result)
            return error_result
        
        except Exception as e:
            self.fail_count += 1
            error_result = {"success": False, "error": f"请求失败: {str(e)}"}
            self.print_result(False, error_result)
            return error_result
    
    def check_service(self) -> bool:
        """检查服务状态"""
        print("🔍 检查服务状态...")
        
        try:
            with urllib.request.urlopen(f"{self.base_url}/health", timeout=2) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                print("✅ 服务运行正常")
                print(f"  • 状态: {data.get('status')}")
                print(f"  • 时间: {data.get('timestamp')}")
                print(f"  • Toolbox: {'已就绪' if data.get('toolbox_ready') else '未初始化'}")
                return True
        except Exception as e:
            print(f"❌ 服务未运行: {e}")
            return False
    
    def get_service_info(self):
        """获取服务信息"""
        print("📊 获取服务信息...")
        
        try:
            with urllib.request.urlopen(f"{self.base_url}/", timeout=2) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                print("✅ 服务信息:")
                print(f"  • 服务名称: {data.get('service')}")
                print(f"  • 版本: {data.get('version')}")
                print(f"  • 运行时间: {data.get('uptime_seconds', 0):.1f}秒")
                print(f"  • 请求次数: {data.get('request_count', 0)}")
                print(f"  • 当前部件: {data.get('current_part', '无')}")
                print(f"  • Toolbox: {'已就绪' if data.get('toolbox_ready') else '未初始化'}")
        except Exception as e:
            print(f"❌ 获取信息失败: {e}")
    
    def print_summary(self):
        """打印测试摘要"""
        self.print_header("测试摘要")
        print(f"总测试数: {self.test_count}")
        print(f"成功: {self.success_count} ✅")
        print(f"失败: {self.fail_count} ❌")
        
        if self.test_count > 0:
            success_rate = (self.success_count / self.test_count) * 100
            print(f"成功率: {success_rate:.1f}%")
        
        print("="*70)


# ==================== 测试用例 ====================

def test_basic_workflow(client: TestClient, part_path: str):
    """测试基础工作流"""
    
    client.print_header("测试 1: 基础工作流")
    
    # 步骤 1: 打开部件
    client.print_step(1, "打开部件文件")
    result = client.call_api(
        "/api/open_part",
        {"file_path": part_path},
        "打开部件"
    )
    
    if not result.get("success"):
        print("\n⚠️  打开部件失败，后续测试可能无法进行")
        return
    
    time.sleep(0.5)
    
    # 步骤 2: 切换到加工环境
    client.print_step(2, "切换到加工环境")
    client.call_api(
        "/api/switch_to_manufacturing",
        {},
        "切换环境"
    )
    
    time.sleep(0.5)
    
    # 步骤 3: 查找面 Tag
    client.print_step(3, "查找面 Tag")
    tag_result = client.call_api(
        "/api/find_tag_by_id",
        {"face_ids": ["mian1"]},
        "查找 mian1"
    )
    
    # 提取 face_tags
    face_tags = []
    if tag_result.get("success"):
        face_tags = tag_result.get("data", {}).get("face_tags", [])
        print(f"📌 获取到的 Tags: {face_tags}")
    
    time.sleep(0.5)
    
    # 步骤 4: 创建工序（如果找到了 tags）
    if face_tags:
        client.print_step(4, "创建 mian1 工序")
        client.call_api(
            "/api/create_mian1_operation",
            {
                "face_tags": face_tags,
                "tool_name": "10R0.5"
            },
            "创建平面铣削工序"
        )
    else:
        print("\n⚠️  未找到面 Tag，跳过创建工序")
    
    time.sleep(0.5)
    
    # 步骤 5: 保存部件
    client.print_step(5, "保存部件")
    client.call_api(
        "/api/save_part",
        {"close_after_save": False},
        "保存部件"
    )


def test_all_cam_operations(client: TestClient):
    """测试所有 CAM 工序创建"""
    
    client.print_header("测试 2: 所有 CAM 工序")
    
    # 测试数据：假设的 face_tags
    test_tags = [100, 101]  # ⚠️ 这些是模拟的 tags，实际需要从 find_tag_by_id 获取
    
    operations = [
        {
            "name": "D4 螺旋深度轮廓铣",
            "endpoint": "/api/create_d4_helical_operation",
            "params": {"face_tags": test_tags, "target_tool_name": "10R0.5"}
        },
        {
            "name": "清角工序",
            "endpoint": "/api/create_corner_clearing_operation",
            "params": {"face_tags": test_tags, "tool_name": "D4"}
        },
        {
            "name": "封闭加刀补",
            "endpoint": "/api/create_sealed_with_cutter_compensation_operation",
            "params": {"face_tags": test_tags, "tool_name": "10R0.5"}
        },
        {
            "name": "开放加刀补",
            "endpoint": "/api/create_open_with_cutter_compensation_operation",
            "params": {"face_tags": test_tags, "tool_name": "10R0.5"}
        },
        {
            "name": "往复等高",
            "endpoint": "/api/create_reciprocating_zlevel_operation",
            "params": {"face_tags": test_tags, "tool_name": "D4"}
        },
        {
            "name": "行腔 D4",
            "endpoint": "/api/create_cavity_milling_d4_operation",
            "params": {"face_tags": test_tags, "tool_name": "10R0.5"}
        },
        {
            "name": "爬面工序",
            "endpoint": "/api/create_surface_contour_operation",
            "params": {"face_tags": test_tags, "tool_name": "10R0.5"}
        }
    ]
    
    for i, op in enumerate(operations, 1):
        client.print_step(i, f"测试 {op['name']}")
        print(f"⚠️  这是模拟测试，使用假 Tags: {test_tags}")
        print("   实际使用时需要先用 find_tag_by_id 获取真实 Tags")
        print("   跳过此测试...")
        # client.call_api(op['endpoint'], op['params'], op['name'])
        # time.sleep(0.3)


def test_find_multiple_tags(client: TestClient):
    """测试查找多个面 Tag"""
    
    client.print_header("测试 3: 查找多个面 Tag")
    
    face_id_groups = [
        ["mian1"],
        ["lx_1", "lx_2"],
        ["qj_1", "qj_2"],
        ["xq_1"],
    ]
    
    for i, face_ids in enumerate(face_id_groups, 1):
        client.print_step(i, f"查找 {', '.join(face_ids)}")
        client.call_api(
            "/api/find_tag_by_id",
            {"face_ids": face_ids},
            f"查找面 {face_ids}"
        )
        time.sleep(0.3)




# ==================== 主程序 ====================

def main():
    """主测试流程"""
    
    print("="*70)
    print("  NX HTTP 服务测试脚本")
    print("  版本: 1.0")
    print("  时间:", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    print("="*70)
    
    # 创建测试客户端
    client = TestClient(NX_SERVICE_URL)
    
    # 检查服务
    client.print_header("前置检查")
    if not client.check_service():
        print("\n❌ 服务未运行，请先启动 HTTP 服务")
        print("启动命令: python nx_http_service.py")
        return
    
    print()
    client.get_service_info()

    
    # 询问部件路径
    part_path = input(f"\n请输入部件路径 (回车使用默认: {TEST_PART_PATH}): ").strip()
    if not part_path:
        part_path = TEST_PART_PATH

    
    test_basic_workflow(client, part_path)

    # 打印摘要
    client.print_summary()
    

if __name__ == "__main__":
    main()