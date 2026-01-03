# -*- coding: utf-8 -*-
"""
mcp_server.py - NX HTTP 服务的 MCP 接口（极简版）
===========================================
作者: 鲁统林
版本: 11.0 - 对应极简版 HTTP 服务
"""

import json
import time
import urllib.error
import urllib.request

from fastmcp import FastMCP

# ==================== 配置 ====================

NX_SERVICE_URL = "http://127.0.0.1:8765"
MCP_SERVICE_NAME = "nx-cam-service"
MCP_SERVICE_VERSION = "11.0"

API_TIMEOUT = 120  # 默认超时时间（秒）
API_RETRY_COUNT = 3  # 重试次数
API_RETRY_DELAY = 1  # 重试延迟（秒）

# ==================== 初始化 MCP ====================

mcp = FastMCP(MCP_SERVICE_NAME)


# ==================== HTTP 客户端 ====================

class NXServiceClient:
    """NX 服务客户端（带重试和缓存）"""
    
    def __init__(self, base_url: str, timeout: int = API_TIMEOUT):
        self.base_url = base_url
        self.timeout = timeout
        self._last_check = 0
        self._check_interval = 5  # 缓存5秒
        self._is_running = False
    
    def is_running(self) -> bool:
        """检查服务状态（带缓存）"""
        now = time.time()
        if now - self._last_check < self._check_interval:
            return self._is_running
        
        try:
            with urllib.request.urlopen(f"{self.base_url}/health", timeout=2) as resp:
                self._is_running = True
                self._last_check = now
                return True
        except:
            self._is_running = False
            self._last_check = now
            return False
    
    def call(self, endpoint: str, params: dict = None, timeout: int = None) -> dict:
        """调用 API（带重试机制）"""
        if not self.is_running():
            return {
                "success": False,
                "error": "NX 服务未运行",
                "hint": "请先在 NX 中启动服务: python http_server.py"
            }
        
        url = f"{self.base_url}{endpoint}"
        data = json.dumps(params or {}).encode('utf-8')
        actual_timeout = timeout or self.timeout
        
        for attempt in range(API_RETRY_COUNT):
            try:
                req = urllib.request.Request(
                    url, 
                    data=data, 
                    headers={"Content-Type": "application/json"}, 
                    method="POST"
                )
                with urllib.request.urlopen(req, timeout=actual_timeout) as response:
                    return json.loads(response.read().decode('utf-8'))
            
            except urllib.error.URLError as e:
                if attempt < API_RETRY_COUNT - 1:
                    time.sleep(API_RETRY_DELAY)
                    continue
                return {"success": False, "error": f"网络错误: {str(e)}"}
            
            except Exception as e:
                return {"success": False, "error": f"请求失败: {str(e)}"}
        
        return {"success": False, "error": "未知错误"}


# 全局客户端实例
nx_client = NXServiceClient(NX_SERVICE_URL)


# ==================== 结果格式化 ====================

def format_result(result: dict) -> str:
    """格式化结果输出"""
    if not result.get("success", False):
        lines = ["❌ 操作失败", ""]
        lines.append(f"错误: {result.get('error', '未知错误')}")
        
        if "hint" in result:
            lines.append(f"💡 提示: {result['hint']}")
        
        if "details" in result:
            lines.append(f"详情: {result['details']}")
        
        return "\n".join(lines)
    
    lines = ["✅ 操作成功", ""]
    
    # 获取数据
    data = result.get("data", {})
    
    # 显示主要消息
    if "message" in data:
        lines.append(f"📝 {data['message']}")
    elif "_message_info" in result:
        lines.append(f"📝 {result['_message_info']}")
    
    # 显示详细信息
    if data and len(data) > 1 or (len(data) == 1 and "message" not in data):
        lines.append("")
        lines.append("📊 详细信息:")
        
        # 优先显示的关键字段
        priority_keys = [
            "part_name", "file_path", "unit", "is_modified",
            "operation_name", "tool_name", "environment",
            "craft_result", "workflow_result", "saved"
        ]
        
        # 显示优先字段
        for key in priority_keys:
            if key in data and key != "message":
                value = data[key]
                lines.append(f"  • {key}: {_format_value(value)}")
        
        # 显示其他字段
        for key, value in data.items():
            if key not in priority_keys and key != "message":
                lines.append(f"  • {key}: {_format_value(value)}")
    
    return "\n".join(lines)


def _format_value(value):
    """格式化单个值"""
    if isinstance(value, list):
        if len(value) > 5:
            return f"{value[:5]}... (共 {len(value)} 个)"
        return value
    elif isinstance(value, dict):
        return f"{{...}} (共 {len(value)} 项)"
    elif isinstance(value, bool):
        return "是" if value else "否"
    return value


# ==================== 服务管理 ====================

@mcp.tool()
def check_service() -> str:
    """
    检查 NX 服务状态
    
    返回服务运行状态、版本信息、当前部件等
    """
    if not nx_client.is_running():
        return "❌ 服务未运行\n\n💡 请先在 NX 中启动服务:\n  python http_server.py"
    
    result = nx_client.call("/")
    if not result:
        return "⚠️ 服务响应异常"
    
    lines = ["✅ NX 服务运行中", ""]
    lines.append(f"📌 版本: {result.get('version', 'N/A')}")
    lines.append(f"⏱️ 运行时长: {result.get('uptime_seconds', 0):.1f} 秒")
    lines.append(f"📊 请求计数: {result.get('request_count', 0)}")
    lines.append(f"🔧 Session: {'✓ 就绪' if result.get('session_ready') else '✗ 未就绪'}")
    
    if result.get('current_part'):
        lines.append(f"📁 当前部件: {result['current_part']}")
    else:
        lines.append("📁 当前部件: 未打开")
    
    if result.get('last_error'):
        lines.append("")
        lines.append(f"⚠️ 上次错误: {result['last_error']}")
    
    return "\n".join(lines)


@mcp.tool()
def get_all_endpoints() -> str:
    """
    获取所有可用的 API 端点
    
    返回完整的端点列表和参数说明
    """
    result = nx_client.call("/api/endpoints")
    
    if not result.get("success"):
        return format_result(result)
    
    endpoints = result.get("endpoints", [])
    lines = [f"📋 可用 API 端点（共 {len(endpoints)} 个）", ""]
    
    for ep in endpoints:
        lines.append(f"🔹 {ep['path']}")
        lines.append(f"   方法: {ep['method']}")
        lines.append(f"   说明: {ep['desc']}")
        if ep.get('params'):
            lines.append(f"   参数: {', '.join(ep['params'])}")
        lines.append("")
    
    return "\n".join(lines)


# ==================== 部件管理 ====================

@mcp.tool()
def open_part(file_path: str) -> str:
    """
    打开 NX 部件文件
    
    参数:
        file_path: 部件文件的完整路径（.prt 文件）
    
    示例:
        open_part("C:/Projects/test_part.prt")
        open_part("/home/user/models/sample.prt")
    
    说明:
        - 会自动刷新当前工作部件
        - 支持连续打开不同文件
        - 返回部件的基本信息（名称、路径、单位等）
    """
    result = nx_client.call("/api/open_part", {"file_path": file_path})
    return format_result(result)


@mcp.tool()
def save_part() -> str:
    """
    保存当前部件
    自动保存初始文件路径下的output文件夹内，并以时间戳为后缀重命名
    """

    
    result = nx_client.call("/api/save_part")
    return format_result(result)


@mcp.tool()
def get_part_info() -> str:
    """
    获取当前部件的详细信息
    
    返回:
        - 部件名称
        - 文件路径
        - 单位
        - 是否已修改
        - Leaf 名称
    
    说明:
        - 需要先打开部件
        - 实时获取最新信息
    """
    result = nx_client.call("/api/get_part_info")
    return format_result(result)


# ==================== CAM 环境 ====================

@mcp.tool()
def switch_to_manufacturing() -> str:
    """
    切换到 CAM 加工环境
    
    说明:
        - 将 NX 切换到加工模块
        - 初始化 CAM 会话
        - 创建工序前必须执行此操作
        - 需要先打开部件
    
    典型流程:
        1. open_part("xxx.prt")
        2. switch_to_manufacturing()
        3. process_nx_crafts() 或其他 CAM 操作
    """
    result = nx_client.call("/api/switch_to_manufacturing")
    return format_result(result)


# ==================== 工序处理 ====================

@mcp.tool()
def process_nx_crafts(judgement_M: bool = False) -> str:
    """
    处理 NX 工艺（创建 CAM 工序）
    
    参数:
        judgement_M: 是否进行 M 代码判断（默认: False）
    
    说明:
        - 根据工艺定义自动创建 CAM 工序
        - 需要先切换到 CAM 环境（switch_to_manufacturing）
        - 会调用 modules/procsse_sort.py 中的工艺处理逻辑
    
    典型流程:
        1. open_part("xxx.prt")
        2. switch_to_manufacturing()
        3. process_nx_crafts()
    """
    result = nx_client.call(
        "/api/process_nx_crafts", 
        {"judgement_M": judgement_M},
        timeout=120  # 工艺处理可能耗时较长
    )
    return format_result(result)


@mcp.tool()
def drilling_automation() -> str:
    """
    自动打孔工作流程
    
    说明:
        - 执行完整的自动打孔流程
        - 需要先打开部件
        - 调用 modules/Drilling_Automation/main_workflow.py
        - 该操作可能耗时较长
    
    典型流程:
        1. open_part("xxx.prt")
        2. drilling_automation()
    """
    result = nx_client.call(
        "/api/Drilling_Automation",
        {},
        timeout=180  # 自动打孔可能耗时很长
    )
    return format_result(result)


# ==================== 工作流示例 ====================

@mcp.tool()
def complete_cam_workflow(file_path: str, judgement_M: bool = False) -> str:
    """
    完整的 CAM 工作流程（一键执行）
    
    参数:
        file_path: 部件文件路径
        judgement_M: 是否进行 M 代码判断（默认: False）
    
    流程:
        1. 打开部件
        2. 切换到 CAM 环境
        3. 处理工艺
        4. 保存部件
    
    示例:
        complete_cam_workflow("C:/Projects/test.prt")
    """
    steps = []
    
    # Step 1: 打开部件
    steps.append("📂 步骤 1/4: 打开部件...")
    result = nx_client.call("/api/open_part", {"file_path": file_path})
    if not result.get("success"):
        return "❌ 打开部件失败\n\n" + format_result(result)
    steps.append("  ✓ 部件已打开")
    
    # Step 2: 切换环境
    steps.append("\n🔧 步骤 2/4: 切换到 CAM 环境...")
    result = nx_client.call("/api/switch_to_manufacturing")
    if not result.get("success"):
        return "❌ 切换环境失败\n\n" + format_result(result)
    steps.append("  ✓ 已切换到 CAM 环境")
    
    # Step 3: 处理工艺
    steps.append("\n⚙️ 步骤 3/4: 处理工艺...")
    result = nx_client.call(
        "/api/process_nx_crafts", 
        {"judgement_M": judgement_M},
        timeout=120
    )
    if not result.get("success"):
        return "❌ 工艺处理失败\n\n" + format_result(result)
    steps.append("  ✓ 工艺处理完成")
    
    # Step 4: 保存部件
    steps.append("\n💾 步骤 4/4: 保存部件...")
    result = nx_client.call("/api/save_part")
    if not result.get("success"):
        return "❌ 保存失败\n\n" + format_result(result)
    steps.append("  ✓ 部件已保存")
    
    # 完成
    steps.append("\n" + "="*50)
    steps.append("✅ 完整工作流程执行成功！")
    
    return "\n".join(steps)


# ==================== 主程序 ====================

if __name__ == "__main__":
    print("="*70)
    print(f"MCP 服务: {MCP_SERVICE_NAME}")
    print(f"版本: {MCP_SERVICE_VERSION}")
    print(f"NX 服务地址: {NX_SERVICE_URL}")
    print("="*70)
    print()
    
    mcp.run()