"""
http_server.py - NX HTTP 服务（极简版）
作者: 鲁统林
启动时获取session，打开文件后刷新workPart
"""

import json
import os
import time
import traceback
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer

NX_SERVICE_HOST = "127.0.0.1"
NX_SERVICE_PORT = 8765


_session = None
_current_workpart = None
_request_count = 0
_start_time = datetime.now()
_last_error = None


def print_log(message, level="INFO"):
    """简化的日志输出"""
    timestamp = datetime.now().strftime("%H:%M:%S")


    emoji_map = {
        "INFO": "ℹ️",
        "WARN": "⚠️",
        "ERROR": "❌",
        "SUCCESS": "✅",
        "DEBUG": "🔍"
    }

    emoji = emoji_map.get(level.upper(), "")
    print(f"[{timestamp}] {emoji} {message}", flush=True)


def init_session():
    """
    初始化 Session（仅在服务启动时调用一次）
    """
    global _session
    try:
        import NXOpen
        _session = NXOpen.Session.GetSession()
        print_log("Session 初始化成功", "SUCCESS")
        return True
    except Exception as e:
        print_log(f"Session 初始化失败: {e}", "ERROR")
        traceback.print_exc()
    return False

def get_session():
    """
    获取 Session（单例模式）
    """
    global _session
    if _session is None:
        try:
            import NXOpen
            _session = NXOpen.Session.GetSession()
        except Exception as e:
            raise RuntimeError(f"无法获取 Session: {e}")
    return _session

def refresh_workpart():
    """
    刷新当前工作部件（每次打开文件后调用）
    """
    global _current_workpart

    try:
        session = get_session()
        _current_workpart = session.Parts.Work
        
        if _current_workpart is None:
            print_log("当前没有激活的部件", "WARN")
            return None
        
        print_log(f"工作部件已刷新: {_current_workpart.Name}", "DEBUG")
        return _current_workpart

    except Exception as e:
        print_log(f"刷新工作部件失败: {e}", "ERROR")
        _current_workpart = None
        return None
def get_workpart():
    """
    获取当前工作部件
    如果没有缓存或需要实时获取，则刷新
    """
    global _current_workpart

    # 实时获取最新的 WorkPart
    session = get_session()
    _current_workpart = session.Parts.Work

    if _current_workpart is None:
        raise ValueError("NX 中当前没有打开或激活的部件 (WorkPart is None)")

    return _current_workpart

def success_response(data=None, message=None):
    """成功响应（修正版）"""
    response = {"success": True}
    
    if data is not None:
        response["data"] = data
    
    if message:
        # 如果 data 不存在，先创建一个字典
        if "data" not in response:
            response["data"] = {}
            
        # 只有当 data 是字典时，才能往里面塞 message
        if isinstance(response["data"], dict):
            response["data"]["message"] = message
        else:
            # 如果 data 是字符串或列表，无法插入 message，建议放在外层
            response["_message_info"] = message
            
    return response

def error_response(error_message, details=None):
    """错误响应"""
    response = {"success": False, "error": error_message}
    if details:
        response["details"] = details
    return response


class NXRequestHandler(BaseHTTPRequestHandler):
    """HTTP 请求处理器（极简版）"""

    def log_message(self, format, *args):
        """简化 HTTP 日志"""
        if "/health" not in self.path and self.path != "/":
            print_log(format % args, "INFO")

    def _send_json(self, data, status=200):
        """发送 JSON 响应"""
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        
        response = json.dumps(data, ensure_ascii=False, indent=2)
        self.wfile.write(response.encode('utf-8'))

    def do_OPTIONS(self):
        """处理 CORS 预检请求"""
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        """处理 GET 请求"""
        global _request_count, _start_time, _last_error, _current_workpart
        
        if self.path == "/":
            uptime = (datetime.now() - _start_time).total_seconds()
            
            current_part_name = None
            try:
                if _current_workpart:
                    current_part_name = _current_workpart.Name
                else:
                    # 尝试实时获取
                    session = get_session()
                    work_part = session.Parts.Work
                    if work_part:
                        current_part_name = work_part.Name
            except Exception as e:
                print_log(f"获取部件名称失败: {e}", "DEBUG")
            
            response_data = {
                "service": "NX HTTP Service",
                "version": "11.0 - 极简版",
                "status": "running",
                "uptime_seconds": round(uptime, 2),
                "request_count": _request_count,
                "session_ready": _session is not None,
                "current_part": current_part_name
            }
            
            if _last_error:
                response_data["last_error"] = _last_error
            
            self._send_json(response_data)
        
        elif self.path == "/health":
            self._send_json({
                "status": "ok",
                "session_ready": _session is not None,
                "workpart_ready": _current_workpart is not None,
                "timestamp": datetime.now().isoformat()
            })
        
        elif self.path == "/api/endpoints":
            endpoints = self._get_all_endpoints()
            self._send_json({
                "success": True,
                "endpoints": endpoints,
                "count": len(endpoints)
            })
        
        else:
            self._send_json({"error": "Not Found"}, 404)

    def do_POST(self):
        """处理 POST 请求"""
        global _request_count
        _request_count += 1
        
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)
        
        try:
            params = json.loads(body) if body else {}
        except json.JSONDecodeError:
            self._send_json({"success": False, "error": "Invalid JSON"}, 400)
            return
        
        try:
            print_log(f"API: {self.path}", "INFO")
            
            result = self._handle_request(self.path, params)
            self._send_json(result)
        
        except Exception as e:
            global _last_error
            _last_error = str(e)
            
            print_log(f"请求处理错误: {e}", "ERROR")
            traceback.print_exc()
            
            self._send_json({
                "success": False,
                "error": str(e),
                "traceback": traceback.format_exc()
            })

    # ==================== 请求路由 ====================

    def _handle_request(self, path, params):
        """请求路由分发"""
        
        # 路由映射表
        routes = {
            # 部件管理
            "/api/open_part": lambda: self._open_part(params),
            "/api/save_part": lambda: self._save_part(),
            "/api/get_part_info": lambda: self._get_part_info(),
            # CAM 环境
            "/api/switch_to_manufacturing": lambda: self._switch_to_manufacturing(),
            # 工序处理
            "/api/process_nx_crafts": lambda: self._process_nx_crafts(params),
            "/api/Drilling_Automation": lambda: self._Drilling_Automation(params),
        }
        
        handler = routes.get(path)
        if handler:
            return handler()
        else:
            return error_response(f"未知端点: {path}")

    # ==================== 内部函数 ====================

    def _open_part(self, params):
        """打开部件文件"""
        global _current_workpart, _last_error

        try:
            file_path = params.get("file_path")
            if not file_path:
                return error_response("缺少必需参数：file_path")
            if not os.path.exists(file_path):
                return error_response(f"文件不存在: {file_path}")
            print_log(f"正在打开部件: {file_path}", "INFO")
            # 获取 Session
            session = get_session()
            
            # 打开部件文件
            try:
                base_part, load_status = session.Parts.OpenBaseDisplay(file_path)
                if load_status:
                    load_status.Dispose()
            except Exception as e:
                error_msg = f"打开部件失败: {e}"
                print_log(error_msg, "ERROR")
                traceback.print_exc()
                return error_response(error_msg)
            # 刷新工作部件
            workPart = refresh_workpart()
            if workPart is None:
                return error_response("打开部件后，无法获取工作部件")
            print_log(f"成功打开部件: {workPart.Name}", "SUCCESS")
            # 收集部件信息
            part_info = {
                "part_name": workPart.Name,
                "file_path": workPart.FullPath,
                "unit": str(workPart.PartUnits),
                "is_modified": workPart.IsModified
            }
            return success_response(part_info, message=f"成功打开部件: {workPart.Name}")
        
        except Exception as e:
            error_msg = f"打开部件时出错: {e}"
            print_log(error_msg, "ERROR")
            traceback.print_exc()
            _last_error = error_msg
            return error_response(error_msg)

    def _save_part(self):
        """保存部件到 output 子文件夹（带时间戳）"""
        try:
            workPart = get_workpart() 
            save_path = None
            part_path = workPart.FullPath
            # 防御性编程：如果是新建文件没保存过，FullPath可能为空
            if not part_path:
                return error_response("当前部件未保存过，无法获取路径")
            # --- 路径处理逻辑 ---
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            dir_name, file_name = os.path.split(part_path)
            name, ext = os.path.splitext(file_name)
            output_dir = os.path.join(dir_name, "output")
            os.makedirs(output_dir, exist_ok=True)
            save_path = os.path.join(output_dir, f"{name}_{timestamp}{ext}")
            workPart.SaveAs(save_path)
            return success_response({
                "file_path": save_path,
                "file_name": os.path.basename(save_path)
            }, message="文件保存成功")
            
        except Exception as e:
            # 加上错误捕获，防止保存失败导致服务崩溃
            return error_response(f"保存失败: {str(e)}")

    def _get_part_info(self):
        """获取当前部件信息"""
        try:
            workPart = get_workpart()
            
            part_info = {
                "part_name": workPart.Name,
                "file_path": workPart.FullPath,
                "unit": str(workPart.PartUnits),
                "is_modified": workPart.IsModified,
                "leaf_name": workPart.Leaf
            }
            
            return success_response(part_info)
        
        except Exception as e:
            return error_response(f"获取部件信息失败: {str(e)}")

    def _switch_to_manufacturing(self):
        """切换到加工环境（修正版）"""
        try:
            # 1. 获取当前的 session 和 workPart
            import NXOpen.UF
            session = get_session()
            workPart = get_workpart() # 记得用我们刚才讨论的实时获取函数
            uf = NXOpen.UF.UFSession.GetUFSession()

            if session.ApplicationName != "UG_APP_MANUFACTURING":
                session.ApplicationSwitchImmediate("UG_APP_MANUFACTURING")
            if workPart.CAMSetup is None:
                workPart.CAMSetup.New()
            uf.Cam.InitSession()
            print("切换到加工环境")
            return success_response({
                "environment": "manufacturing"
            }, message="已切换到加工环境")

        except Exception as e:
            print_log(f"切换环境失败: {e}", "ERROR")
            traceback.print_exc()
            return error_response(str(e))

    # ==================== 工序处理 ====================

    def _process_nx_crafts(self, params):
        """处理NX工艺（创建CAM工序）"""
        try:
            workPart = get_workpart()
            
            judgement_M = params.get("judgement_M", False)
            
            from modules.procsse_sort import Procsse_sort
            
            ps = Procsse_sort()
            craft_result = ps.process_nx_crafts(workPart, judgement_M=judgement_M)
            
            return success_response({
                "craft_result": craft_result
            }, message="工艺处理完成")
        
        except Exception as e:
            print_log(f"处理工艺失败: {e}", "ERROR")
            traceback.print_exc()
            return error_response(str(e))

    def _Drilling_Automation(self, params):
        """自动打孔工作流程"""
        try:
            session = get_session()
            workPart = get_workpart()
            
            from modules.Drilling_Automation.main_workflow import MainWorkflow
            
            mw = MainWorkflow(session, workPart)
            result = mw.run_workflow()
            
            return success_response({
                "workflow_result": result
            }, message="自动打孔完成")
        
        except Exception as e:
            print_log(f"自动打孔失败: {e}", "ERROR")
            traceback.print_exc()
            return error_response(str(e))

    # ==================== API 文档 ====================

    def _get_all_endpoints(self):
        """获取所有API端点"""
        return [
            {
                "path": "/api/open_part",
                "method": "POST",
                "desc": "打开部件文件",
                "params": ["file_path"]
            },
            {
                "path": "/api/save_part",
                "method": "POST",
                "desc": "保存部件",
                "params": ["save_path (可选)"]
            },
            {
                "path": "/api/get_part_info",
                "method": "POST",
                "desc": "获取当前部件信息",
                "params": []
            },
            {
                "path": "/api/switch_to_manufacturing",
                "method": "POST",
                "desc": "切换到加工环境",
                "params": []
            },
            {
                "path": "/api/process_nx_crafts",
                "method": "POST",
                "desc": "处理NX工艺",
                "params": ["judgement_M (可选)"]
            },
            {
                "path": "/api/Drilling_Automation",
                "method": "POST",
                "desc": "自动打孔工作流程",
                "params": []
            }
        ]

def main():
    """启动服务"""
    print("="*70, flush=True)
    print("NX HTTP Service - v11.0 极简版", flush=True)
    print("="*70, flush=True)
    print(f"监听地址: http://{NX_SERVICE_HOST}:{NX_SERVICE_PORT}", flush=True)
    print(f"启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", flush=True)
    print("="*70, flush=True)
    print("", flush=True)


    # 初始化 Session（启动时获取）
    print_log("正在初始化 Session...", "INFO")
    if init_session():
        print_log("Session 已就绪", "SUCCESS")
    else:
        print_log("Session 初始化失败，将在首次调用时重试", "WARN")

    print("", flush=True)

    server = HTTPServer((NX_SERVICE_HOST, NX_SERVICE_PORT), NXRequestHandler)

    print_log(f"服务已启动: {NX_SERVICE_HOST}:{NX_SERVICE_PORT}", "SUCCESS")
    print_log("按 Ctrl+C 停止服务", "INFO")
    print("", flush=True)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("", flush=True)
        print_log("正在关闭服务...", "WARN")
        server.shutdown()
        print_log("服务已停止", "INFO")

if __name__ == "__main__":  # ✅ 正确：前后各两个下划线
    main()