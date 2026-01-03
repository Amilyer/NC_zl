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
DLL_PATH = r"C:\Projects\贴图\x64\Debug\NX_Open_Wizard1.dll"

# 2. 要处理的 PRT 文件路径
PART_PATH = r"C:\Projects\NC\output\02_Process\2_Merged_PRT\UP-12.prt"
# ============================================================================

try:
    import NXOpen
    import NXOpen.UF
except ImportError:
    print("❌ 错误: 必须在 NX 环境下运行 (File -> Execute -> NX Open...)")
    sys.exit(1)

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
    print(f"--- 🚀 自动对齐(贴图)任务启动 ---")

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
        # 调用 C++ 导出的无参函数 RunAutoAlign
        ret = plugin.RunAutoAlign()
        
        if ret == 0:
            print("✅ 成功: 实体已根据 2D/3D 信息自动对齐。")
        elif ret == -1:
            print("⚠️ 失败: 没有活动的显示部件。")
            return
        elif ret == 2:
            print("⚠️ 警告: 未找到匹配的移动向量 (可能2D图纸和3D模型不匹配)。")
            return
        else:
            print(f"⚠️ 警告: 未知返回码 {ret}")
            return
            
    except AttributeError:
        print("❌ 错误: DLL中没有找到 RunAutoAlign 函数。")
        print("   请确认您编译的是最新的 C++ 代码，且函数名拼写正确。")
        return
    except Exception as e:
        print(f"❌ 运行时错误: {e}")
        return

    # 5. 保存结果
    print("💾 保存结果...")
    try:
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        dir_name, file_name = os.path.split(PART_PATH)
        name, ext = os.path.splitext(file_name)
        
        # 创建 output 文件夹
        output_dir = os.path.join(dir_name, "output")
        os.makedirs(output_dir, exist_ok=True)
        
        save_path = os.path.join(output_dir, f"{name}_Aligned_{timestamp}{ext}")
        base_part.SaveAs(save_path)
        print(f"✅ 已保存至: {save_path}")
        
    except Exception as e:
        print(f"❌ 保存失败: {e}")
    
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
        print("--- 🏁 任务结束 ---")

if __name__ == "__main__":
    main()