# -*- coding: utf-8 -*-
import ctypes
import json
import os
import sys
import time
import re
import gc

# NXOpen相关导入
try:
    import NXOpen
    import NXOpen.UF
    import NXOpen.CAM
except ImportError:
    print("❌ 错误: 必须在 NX 环境下运行 (File -> Execute -> NX Open...)")
    sys.exit(1)

# ============================================================================
# 🔧 通用DLL加载器 (新版本核心，自动适配C++注册宏)
# ============================================================================
class UniversalLoader:
    def __init__(self, dll_path):
        if not os.path.exists(dll_path):
            raise FileNotFoundError(f"DLL not found: {dll_path}")
        # 根据参考代码 截图cpp.py，直接使用 CDLL
        self.dll = ctypes.CDLL(dll_path)
        print("✅ 以 CDLL(cdecl) 方式加载DLL成功")
        
        self.functions = {}
        self._load_metadata()
    
    def _load_metadata(self):
        """尝试读取C++注册的元数据"""
        try:
            self.dll.get_func_count.restype = ctypes.c_int
            count = self.dll.get_func_count()
            
            self.dll.get_func_info.argtypes = [ctypes.c_int, ctypes.c_char_p, ctypes.c_int]
            buf = ctypes.create_string_buffer(4096)
            
            for i in range(count):
                if self.dll.get_func_info(i, buf, 4096) == 0:
                    try:
                        info = json.loads(buf.value.decode())
                        self._register_func(info)
                    except:
                        pass
            
            print(f"   已注册DLL接口: {list(self.functions.keys())}")
        except Exception:
            print("⚠️ 警告: 无法读取元数据，尝试直接调用函数...")
    
    def _register_func(self, info):
        """注册函数参数和返回值类型"""
        func_name = info['name']
        if not hasattr(self.dll, func_name):
            return
        c_func = getattr(self.dll, func_name)
        self.functions[func_name] = info
        
        # 映射参数类型
        argtypes = []
        for p in info.get('params', []):
            if p['type'] == 4:
                argtypes.append(ctypes.c_char_p)
            else:
                argtypes.append(ctypes.c_int)
        
        c_func.argtypes = argtypes
        c_func.restype = ctypes.c_int
    
    def __getattr__(self, name):
        """动态调用DLL函数（无元数据时默认配置）"""
        if not hasattr(self.dll, name):
            raise AttributeError(f"DLL中未找到函数: {name}")
        
        c_func = getattr(self.dll, name)
        
        # 无元数据时默认配置
        if name not in self.functions:
            c_func.argtypes = []
            c_func.restype = ctypes.c_int
            
        return c_func

# ============================================================================
# 📊 NX CAM工序参数导出器 (保留核心业务逻辑)
# ============================================================================
class NXOperationExporter:
    """NX工序参数导出器（仅提取转速和刀路时间）"""
    
    def __init__(self, base_dir, workpiece_name):
        """初始化NX会话和基础对象
        :param base_dir: 基础目录
        :param workpiece_name: 工件名称（用于文件名）
        """
        self.theSession = NXOpen.Session.GetSession()
        self.theUFSession = NXOpen.UF.UFSession.GetUFSession()
        self.ui = NXOpen.UI.GetUI()
        self.lw = self.theSession.ListingWindow
        
        # 清理非法字符
        self.workpiece_name = re.sub(r'[\\/:*?"<>|]', '_', workpiece_name)
        
        # 定义顶级文件夹
        self.txt_folder = os.path.join(base_dir, "工件信息TXT")
        self.dims_folder = os.path.join(base_dir, "尺寸信息TXT")
        self.json_folder = os.path.join(base_dir, "JSON数据")
        self.img_dir = os.path.join(base_dir, "screen-shot") # 基础截图目录
        
        # 自动创建文件夹
        self._create_folders()
        
        # 生成文件路径
        self.txt_output_path = os.path.join(self.txt_folder, f"{self.workpiece_name}.txt")
        self.dims_output_path = os.path.join(self.dims_folder, f"{self.workpiece_name}_尺寸.txt")
        self.json_output_path = os.path.join(self.json_folder, f"{self.workpiece_name}.json")

        # 需要的参数ID映射
        self.needed_params = {
            124: "Toolpath Time",          # 刀路时间
            142: "Toolpath Cutting Time",  # 切削时间
            4005: "Spindle RPM",           # 主轴转速
        }
    
    def _create_folders(self):
        """自动创建所需的文件夹（不存在则创建）"""
        try:
            for folder in [self.txt_folder, self.dims_folder, self.json_folder, self.img_dir]:
                os.makedirs(folder, exist_ok=True)
                self.lw.WriteLine(f"【文件夹准备】: {folder} (已存在/创建成功)")
        except Exception as e:
            self.lw.WriteLine(f"【创建文件夹失败】: {str(e)}")
    
    def init_log(self):
        """初始化日志窗口"""
        self.lw.Open()
        self.lw.WriteLine("="*50)
        self.lw.WriteLine(f"正在启动[{self.workpiece_name}]工序参数导出程序...")
        self.lw.WriteLine("="*50)
    
    def get_all_operations(self):
        """获取所有CAM工序"""
        workPart = self.theSession.Parts.Work
        camSetup = workPart.CAMSetup
        if not camSetup:
            self.lw.WriteLine("\n【错误】未找到CAM设置！")
            return []
        opCollection = camSetup.CAMOperationCollection
        operations = [op for op in opCollection]
        
        if not operations:
            print(f"   ⚠️ [Debug] {self.workpiece_name}: 未找到任何工序")
            return []
        
        print(f"   ✅ [Debug] {self.workpiece_name}: 检测到 {len(operations)} 个工序")
        return operations
    
    def get_param_value(self, obj_tag, param_id):
        """读取指定参数的值"""
        val = None
        val_type = "Unknown"
        
        # 尝试不同类型
        try:
            val = self.theUFSession.Param.AskDoubleValue(obj_tag, param_id)
            val_type = "Double"
        except:
            try:
                val = self.theUFSession.Param.AskIntValue(obj_tag, param_id)
                val_type = "Int"
            except:
                try:
                    val = self.theUFSession.Param.AskStringValue(obj_tag, param_id)
                    val_type = "String"
                except:
                    pass
        
        return val, val_type
    
    def debug_inspect_op(self, operation):
        """调试：检查工序属性"""
        try:
            print(f"   [Inspect] 工序类型: {type(operation)}")
            attrs = [x for x in dir(operation) if "Feed" in x or "Speed" in x or "ToolpathTime" in x]
            print(f"   [Inspect] 相关属性: {attrs}")
            
            # 检查刀路时间方法
            if hasattr(operation, "GetToolpathTime"):
                try:
                    time_val = operation.GetToolpathTime()
                    print(f"   [Inspect] 刀路时间: {time_val}")
                except Exception as ex:
                    print(f"   [Inspect] 读取刀路时间失败: {ex}")
        except:
            pass
    
    def get_cam_operation_native_attrs(self, operation):
        """通过NXOpen.CAM原生接口获取参数"""
        native_attrs = {}
        try:
            # 获取主轴转速
            feeds_builder = None
            if hasattr(operation, "GetFeeds"):
                feeds_builder = operation.GetFeeds()
            elif hasattr(operation, "Feeds"):
                feeds_builder = operation.Feeds
            
            if feeds_builder and hasattr(feeds_builder, "SpindleSpeedBuilder"):
                spindle_val = feeds_builder.SpindleSpeedBuilder.Value
                native_attrs["Spindle_RPM_Native"] = spindle_val.Value if hasattr(spindle_val, "Value") else spindle_val
            
            # 获取刀路时间
            if hasattr(operation, "GetToolpathTime"):
                native_attrs["Toolpath_Time_Native"] = operation.GetToolpathTime()
            
            # 尝试参数ID 73
            val_73, _ = self.get_param_value(operation.Tag, 73)
            if val_73 is not None:
                native_attrs["Spindle_RPM_ID73"] = val_73
                
        except Exception as e:
            print(f"       ⚠️ 读取原生属性失败: {e}")
        return native_attrs
    
    def collect_operation_params(self, operation):
        """收集单个工序的参数"""
        op_name = operation.Name
        obj_tag = operation.Tag
        print(f"   [Debug] 处理工序: {op_name}")
        self.debug_inspect_op(operation)
        
        collected_params = []
        # 扫描需要的参数
        for param_id, display_name in self.needed_params.items():
            val, val_type = self.get_param_value(obj_tag, param_id)
            if val is not None:
                collected_params.append({
                    "id": param_id,
                    "display_name": display_name,
                    "type": val_type,
                    "value": val
                })
        
        # 补充原生接口数据
        native_attrs = self.get_cam_operation_native_attrs(operation)
        
        # 补充刀路时间
        if not any(p['id'] == 124 for p in collected_params) and "Toolpath_Time_Native" in native_attrs:
            collected_params.append({
                "id": 124,
                "display_name": "Toolpath Time",
                "type": "Double",
                "value": native_attrs["Toolpath_Time_Native"]
            })
        
        # 补充切削时间
        if not any(p['id'] == 142 for p in collected_params) and "Toolpath_Time_Native" in native_attrs:
            collected_params.append({
                "id": 142,
                "display_name": "Toolpath Cutting Time",
                "type": "Double",
                "value": native_attrs["Toolpath_Time_Native"]
            })
        
        # 补充主轴转速
        if not any(p['id'] == 4005 for p in collected_params):
            rpm_val = native_attrs.get("Spindle_RPM_Native") or native_attrs.get("Spindle_RPM_ID73")
            if rpm_val is not None:
                collected_params.append({
                    "id": 4005,
                    "display_name": "Spindle RPM",
                    "type": "Double",
                    "value": rpm_val
                })
        
        self.lw.WriteLine(f"  工序[{op_name}]收集到 {len(collected_params)} 个有效参数")
        
        return {
            "operation_name": op_name,
            "total_params": len(collected_params),
            "parameters": collected_params
        }
    
    def collect_all_operations_data(self, operations):
        """收集所有工序数据"""
        all_data = []
        for op in operations:
            all_data.append(self.collect_operation_params(op))
        return all_data
    
    def build_result_data(self, all_data):
        """构建JSON数据结构"""
        return {
            "meta_data": {
                "export_time": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
                "workpiece_name": self.workpiece_name,
                "total_operations": len(all_data),
                "total_parameters": sum(op["total_params"] for op in all_data)
            },
            "operations": all_data
        }
    
    def save_to_json(self, result_data):
        """保存JSON文件"""
        try:
            with open(self.json_output_path, "w", encoding='utf-8') as f:
                json.dump(result_data, f, indent=4, ensure_ascii=False)
            
            print("="*30)
            print("       工序参数导出成功！")
            print(f"JSON路径: {self.json_output_path}")
            print("="*30)
        except Exception as e:
            print(f"\n【保存JSON失败】: {str(e)}")
    
    def run(self):
        """执行导出流程"""
        self.init_log()
        operations = self.get_all_operations()
        if not operations:
            return
        all_data = self.collect_all_operations_data(operations)
        result_data = self.build_result_data(all_data)
        self.save_to_json(result_data)

# ============================================================================
# 📏 辅助函数：尺寸提取、路径处理、权限检查
# ============================================================================
def get_workpiece_name(part_path=None):
    """获取工件名称（不含扩展名）"""
    if part_path:
        return os.path.splitext(os.path.basename(part_path))[0]
    # 从NX会话获取
    session = NXOpen.Session.GetSession()
    workPart = session.Parts.Work
    part_path = workPart.FullPath or workPart.Name
    return os.path.splitext(os.path.basename(part_path))[0]

def get_workpiece_path():
    """获取当前工作部件的完整路径"""
    session = NXOpen.Session.GetSession()
    workPart = session.Parts.Work
    return workPart.FullPath or workPart.Name

def process_cam_operations(txt_output_path):
    """处理CAM工序信息并写入TXT"""
    session = NXOpen.Session.GetSession()
    workPart = session.Parts.Work
    workpiece_path = get_workpiece_path()

    # 写入路径信息
    try:
        with open(txt_output_path, "w", encoding='gbk') as f:
            f.write(f"当前工作部件完整路径: {workpiece_path}\n")
            f.write("="*60 + "\n\n")
    except Exception as e:
        print(f"❌ 写入TXT失败: {e}")
        return

    list_window = session.ListingWindow
    list_window.SelectDevice(NXOpen.ListingWindow.DeviceType.File, txt_output_path)
    list_window.Open()
    infoTool = session.Information

    camSetup = workPart.CAMSetup
    if not camSetup:
        list_window.WriteLine("未找到CAM设置！")
        return
    opCollection = camSetup.CAMOperationCollection
    operations = [op for op in opCollection]

    if not operations:
        list_window.WriteLine("没有找到工序！")
        return

    # 输出工序详情
    for op in operations:
        infoTool.DisplayCamObjectsDetails([op])
        list_window.WriteLine("="*40)
    list_window.Close()

def extract_dimensions_from_text(text):
    """解析Note文本中的尺寸"""
    patterns = [
        r"(\d+\.?\d*)\s*[\*x×]\s*(\d+\.?\d*)\s*[\*x×]\s*(\d+\.?\d*)",
        r"(\d+\.?\d*)\s*[lL].*?(\d+\.?\d*)\s*[wW].*?(\d+\.?\d*)\s*[tThH]",
        r"长\s*(\d+\.?\d*).*?宽\s*(\d+\.?\d*).*?高\s*(\d+\.?\d*)"
    ]
    
    text = text.replace("×", "*").replace("x", "*")
    
    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            return f"{m.group(1)}*{m.group(2)}*{m.group(3)}"
    return None

def read_dwg_objects_and_annotations(dims_output_path):
    """提取尺寸信息并写入TXT"""
    session = NXOpen.Session.GetSession()
    workPart = session.Parts.Work
    try:
        notes = list(workPart.Notes)
    except:
        notes = []

    dimension_str = ""
    for n in notes:
        try:
            text = " ".join(n.GetText())
            dim = extract_dimensions_from_text(text)
            if dim:
                dimension_str = dim
                break
        except Exception as e:
            continue
    
    # 写入尺寸文件
    try:
        with open(dims_output_path, "w", encoding='gbk') as f:
            f.write(dimension_str if dimension_str else "")
    except Exception as e:
        print(f"❌ 写入尺寸TXT失败: {e}")
    
    # 日志提示
    print_to_info_window(f"尺寸提取结果: {dimension_str or '未找到有效尺寸'}")

def print_to_info_window(message):
    """打印到NX信息窗口"""
    session = NXOpen.Session.GetSession()
    session.ListingWindow.Open()
    session.ListingWindow.WriteLine(str(message))

def check_directory_permission(dir_path):
    """检查目录写入权限"""
    if not os.path.exists(dir_path):
        return False
    test_file = os.path.join(dir_path, "tmp_permission_test.txt")
    try:
        with open(test_file, "w", encoding="utf-8") as f:
            f.write("test")
        os.remove(test_file)
        return True
    except Exception as e:
        print(f"❌ 目录无写入权限: {e}")
        return False

# ============================================================================
# 🚀 主程序
# ============================================================================
def main():
    print(f"--- 🚀 整合任务启动 ---")
    
    # 1. 解析参数
    # 参数顺序: 脚本名, prt_path, output_base_dir, dll_path
    if len(sys.argv) < 4:
        # 如果参数不足，给出提示并退出（或者可以保留硬编码作为测试，但这里按要求改为纯参数驱动）
        print("用法错误: 请传入 <prt_path> <output_base_dir> <dll_path>")
        # 调试模式下即使没有参数也可能运行，这时候需要返回
        # 为了演示，我们打印当前参数
        print(f"当前参数: {sys.argv}")
        return

    prt_path = sys.argv[1]
    output_base_dir = sys.argv[2]
    dll_path = sys.argv[3]

    print(f"📄 PRT路径: {prt_path}")
    print(f"📂 输出根目录: {output_base_dir}")
    print(f"🔌 DLL路径: {dll_path}")

    # 2. 前置检查
    if not os.path.exists(dll_path):
        print(f"❌ DLL不存在: {dll_path}")
        return
    if not os.path.exists(prt_path):
        print(f"❌ PRT文件不存在: {prt_path}")
        return
    
    # 尝试创建输出根目录
    try:
        os.makedirs(output_base_dir, exist_ok=True)
    except Exception as e:
        print(f"❌ 无法创建输出根目录 {output_base_dir}: {e}")
        return
        
    if not check_directory_permission(output_base_dir):
        return

    session = NXOpen.Session.GetSession()
    base_part = None

    # 3. 打开部件
    try:
        print(f"📂 打开部件: {os.path.basename(prt_path)}")
        base_part, _ = session.Parts.OpenBaseDisplay(prt_path)
    except Exception as e:
        print(f"❌ 打开部件失败: {e}")
        return

    # 4. 初始化导出器 (会自动在 output_base_dir 下创建 4 个子文件夹)
    workpiece_name = get_workpiece_name(prt_path)
    exporter = NXOperationExporter(output_base_dir, workpiece_name)
    
    # 5. 处理CAM工序和尺寸
    print("📊 处理CAM工序信息...")
    process_cam_operations(exporter.txt_output_path)
    print("📏 提取尺寸信息...")
    read_dwg_objects_and_annotations(exporter.dims_output_path)

    # 6. 加载DLL并执行截图
    print("🔌 加载DLL...")
    try:
        plugin = UniversalLoader(dll_path)
        
        # 构造专用截图目录路径: output_base_dir/screen-shot/{workpiece_name}
        # NXOperationExporter 初始化时其实只创建了 screen-shot 文件夹，这里需要额外创建专用子文件夹
        img_root_dir = os.path.join(output_base_dir, "screen-shot")
        specific_output_dir = os.path.join(img_root_dir, workpiece_name)
        
        # 确保专用目录存在
        os.makedirs(specific_output_dir, exist_ok=True)
        
        # 兼容性处理：添加末尾分隔符并转义反斜杠
        specific_output_dir = os.path.abspath(specific_output_dir)
        if not specific_output_dir.endswith(os.sep):
            specific_output_dir += os.sep
        # 某些DLL实现可能需要双反斜杠路径
        specific_output_dir = specific_output_dir.replace("\\", "\\\\")
        
        print(f"📷 专用截图目录(DLL参数): {specific_output_dir}")

        # 直接传递专用目录路径给 DLL
        arg = specific_output_dir.encode('utf-8')

        rc = None
        func_name = None

        # 优先尝试新版本函数名 FlipAndShotForPy
        try:
            try:
                rc = plugin.FlipAndShotForPy(arg)
                func_name = "FlipAndShotForPy"
            except TypeError:
                rc = plugin.FlipAndShotForPy()
                func_name = "FlipAndShotForPy()"
        except AttributeError:
            # 兼容旧版本函数名 FlipAndShotforPython
            try:
                try:
                    rc = plugin.FlipAndShotforPython(arg)
                    func_name = "FlipAndShotforPython"
                except TypeError:
                    rc = plugin.FlipAndShotforPython()
                    func_name = "FlipAndShotforPython()"
            except AttributeError:
                print(f"❌ DLL中未找到合适的截图函数 (FlipAndShotForPy / FlipAndShotforPython)")
                return

        print(f"✅ 调用{dll_path}::{func_name} -> 返回值: {rc}")
        if rc == 0:
            print_to_info_window("【截图成功】已保存到指定目录")
        else:
            print_to_info_window(f"【截图失败】DLL返回错误码: {rc}")
    except Exception as e:
        print(f"❌ DLL操作失败: {e}")
        # 继续执行后续导出，不完全中断

    # 7. 执行CAM参数导出
    print("📊 导出CAM工序参数...")
    exporter.run()

    # 8. 清理工作
    if base_part:
        try:
            base_part.Close(NXOpen.BasePart.CloseWholeTree.TrueValue,
                            NXOpen.BasePart.CloseModified.CloseModified, None)
        except Exception:
            pass
    base_part = None
    gc.collect()
    print("--- 🎯 所有任务完成 ---")
    print_to_info_window("所有文件已按类型存储到对应文件夹！")

if __name__ == "__main__":
    main()