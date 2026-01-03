# -*- coding: utf-8 -*-
import ctypes
import json
import os
import sys
import time
import glob
import gc

# ============================================================================
# ⚙️ 配置区域 (请根据实际情况修改)
# ============================================================================
# 1. DLL 路径 (确保是最新编译的 AutoAlignDll)
DLL_PATH = r"C:\Projects\NCv4.7\core\DLL\FlipAndShot\jietujietu.dll"

# 2. 要处理的 PRT 文件路径
PART_PATH = r"C:\Projects\NCv4.7\output\06_CAM\Final_CAM_PRT\LB-06.prt"

# 3. 用于保存截图的根目录
OUT_PATH = r"C:\Projects\Fighting2_chaifen\Screenshots"
# ============================================================================

try:
    import NXOpen
    import NXOpen.UF
except ImportError:
    print("❌ 错误: 必须在 NX 环境下运行 (File -> Execute -> NX Open...)")
    sys.exit(1)


import os
# -->-->-->-->-->-->-->-->-->-->-->-->-->-->-->-->-->
# 路径生成函数
def generate_output_paths(prt_path, output_root_dir):
    # 1. 提取文件名 (例如: "DIE-xxx1.prt")
    filename_with_ext = os.path.basename(prt_path)
    
    # 2. 去除后缀 (例如: "DIE-xxx1")
    filename_no_ext = os.path.splitext(filename_with_ext)[0]
    
    # 3. 拼接子文件夹路径 (例如: "E:\work\screen-shot\DIE-xxx1")
    # 这就是你需要传给 DLL 的路径
    specific_output_dir = os.path.join(output_root_dir, filename_no_ext)
    
    # 4. 如果文件夹不存在，则创建它 (非常重要，否则DLL保存时会报错)
    if not os.path.exists(specific_output_dir):
        os.makedirs(specific_output_dir)
        print(f"已创建文件夹: {specific_output_dir}")

    # 1. 强制添加结尾斜杠
    if not specific_output_dir.endswith(os.sep):
        specific_output_dir += os.sep
    
    # 2. 【修改点】将所有单斜杠替换为双斜杠
    specific_output_dir = specific_output_dir.replace("\\", "\\\\")
        
    
    return specific_output_dir
    #预期生成 "E:\\work\\screen-shot\\DIE-xxx1\\"
# -->-->-->-->-->-->-->-->-->-->-->-->-->-->-->-->-->


# ============================================================================
# 🔧 通用加载器 (自动适配 C++ 注册宏)
# ============================================================================
class UniversalLoader:
    def __init__(self, dll_path):
        if not os.path.exists(dll_path): 
            raise FileNotFoundError(f"DLL not found: {dll_path}")
        self.dll = ctypes.CDLL(dll_path)
        self.functions = {}
        self._load_metadata()
    
    def _load_metadata(self):
        """尝试读取 C++ 注册的元数据"""
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
                    except: pass
            
            print(f"   已注册接口: {list(self.functions.keys())}")
        except Exception:
            print("⚠️ 警告: 无法读取元数据，尝试直接调用...")
    
    def _register_func(self, info):
        func_name = info['name']
        if not hasattr(self.dll, func_name): return
        c_func = getattr(self.dll, func_name)
        self.functions[func_name] = info
        
        # 映射参数类型
        argtypes = []
        for p in info.get('params', []):
            if p['type'] == 4: argtypes.append(ctypes.c_char_p)
            else: argtypes.append(ctypes.c_int)
        
        c_func.argtypes = argtypes
        c_func.restype = ctypes.c_int
    
    def __getattr__(self, name):
        """动态调用 DLL 函数"""
        if not hasattr(self.dll, name):
            raise AttributeError(f"DLL 中未找到函数: {name}")
        
        c_func = getattr(self.dll, name)
        
        # 如果没有元数据，默认无参数或根据调用推断(不推荐)，这里假设无参
        if name not in self.functions:
            c_func.argtypes = []
            c_func.restype = ctypes.c_int
            
        return c_func

# ============================================================================
# 🚀 主程序
# ============================================================================
def main():
    print(f"--- 🚀 翻转和截图  任务启动 ---")

    # 1. 检查文件是否存在
    if not os.path.exists(DLL_PATH):
        print(f"❌ DLL 不存在: {DLL_PATH}")
        return
    if not os.path.exists(PART_PATH):
        print(f"❌ PRT 不存在: {PART_PATH}")
        return

    session = NXOpen.Session.GetSession()
    base_part = None

    try:
        # 2. 打开部件
        print(f"📂 打开: {os.path.basename(PART_PATH)}")
        base_part, _ = session.Parts.OpenBaseDisplay(PART_PATH)
    except Exception as e:
        print(f"❌ 打开失败: {e}")
        return

    # 3. 加载 DLL
    print("🔌 加载 DLL...")
    try:
        plugin = UniversalLoader(DLL_PATH)
    except Exception as e:
        print(f"❌ DLL 加载失败: {e}")
        return

    # 4. 执行核心功能
    print("🚀 正在计算并移动实体...")
    try:
        specific_output_dir = generate_output_paths(PART_PATH, OUT_PATH)
        print(f"   输出路径: {specific_output_dir}")
        arg = specific_output_dir.encode('utf-8') if specific_output_dir else None
        rc = plugin.FlipAndShotForPy(arg)
        print(f"FlipAndShotForPy({specific_output_dir!r}) -> {rc}")
        
    except AttributeError:
        print("❌ 错误: DLL中没有找到 RunAutoAlign 函数。")
        print("   请确认您编译的是最新的 C++ 代码，且函数名拼写正确。")
        return
    except Exception as e:
        print(f"❌ 运行时错误: {e}")
        return   
    finally:
        # 6. 清理工作
        if base_part:
            try:
                # 关闭部件 (不做保存，因为已经另存为了)
                base_part.Close(NXOpen.BasePart.CloseWholeTree.TrueValue, 
                                NXOpen.BasePart.CloseModified.CloseModified, None)
            except: pass
        base_part = None
        gc.collect()
        print("--- 截图完成 ---")

if __name__ == "__main__":
    main()