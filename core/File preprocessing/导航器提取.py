# -*- coding: utf-8 -*-
import ctypes
import gc
import glob
import json
import os
import sys

# ============================================================================
# ⚙️ 配置区域
# ============================================================================
# 1. 编译好的 DLL 路径
DLL_PATH = r"C:\Projects\NC\modules\导航器提取.dll"

# 2. 输入文件夹 (PRT文件)
INPUT_FOLDER = r"C:\Projects\NC\file\cleaned_prt"

# 3. 输出文件夹名称
OUTPUT_DIR_NAME = "Navigator_Reports"

# 4. 目标图层 (0=所有图层, 20=指定图层)
TARGET_LAYER = 0  # 建议先设为 0 以确保能找到实体
# ============================================================================

try:
    import NXOpen
    import NXOpen.UF
except ImportError:
    print("❌ 错误: 必须在 NX 环境下运行 (File -> Execute -> NX Open...)")
    sys.exit(1)

# ============================================================================
# 🔧 辅助函数: 确保 CAM 环境就绪
# ============================================================================
def ensure_cam_setup_ready(the_session, work_part):
    """
    智能准备 CAM 环境 (修复 'Current part does not contain valid setup' 错误)
    """
    try:
        # 1. 检查 CAM 会话是否启动
        if not the_session.IsCamSessionInitialized():
            # print("   ⚡ 启动 CAM 会话...")
            the_session.CreateCamSession()

        # 2. 检查部件内是否存在 Setup
        # 尝试访问 CAMSetup，如果未初始化或不存在，通常需要在 try 块中处理
        try:
            if work_part.CamSetup.IsInitialized():
                return True
        except:
            pass # 继续向下尝试创建

        # 3. 创建 Setup (如果不存在)
        print("   ⚡ 当前部件没有有效的 Setup，正在自动创建 'hole_making' 环境...")
        
        # 获取默认的 Setup 模板 (通常是 mill_planar, hole_making 等)
        # 这里使用 hole_making 作为通用模板
        try:
            work_part.CreateCamSetup("hole_making")
            print("   ✅ CAM Setup (hole_making) 创建成功。")
            return True
        except Exception as e:
            # 如果 hole_making 失败，尝试 mill_planar
            print(f"   ⚠️ hole_making 创建失败，尝试 mill_planar... ({e})")
            work_part.CreateCamSetup("mill_planar")
            print("   ✅ CAM Setup (mill_planar) 创建成功。")
            return True

    except Exception as ex:
        print(f"   ❌ 自动创建 CAM Setup 失败: {ex}")
        return False

# ============================================================================
# 🔧 通用加载器 (无需修改)
# ============================================================================
class UniversalLoader:
    def __init__(self, dll_path):
        if not os.path.exists(dll_path): raise FileNotFoundError(f"DLL not found: {dll_path}")
        self.dll = ctypes.CDLL(dll_path)
        self.functions = {}
        self._load_metadata()

    def _load_metadata(self):
        try:
            self.dll.get_func_count.restype = ctypes.c_int
            count = self.dll.get_func_count()
        except AttributeError:
            raise Exception("DLL 不支持通用接口 (缺少 get_func_count)")

        self.dll.get_func_info.argtypes = [ctypes.c_int, ctypes.c_char_p, ctypes.c_int]
        buf = ctypes.create_string_buffer(4096)
        
        for i in range(count):
            if self.dll.get_func_info(i, buf, 4096) == 0:
                info = json.loads(buf.value.decode())
                self._register_func(info)

    def _register_func(self, info):
        func_name = info['name']
        if not hasattr(self.dll, func_name): return
        c_func = getattr(self.dll, func_name)
        self.functions[func_name] = info
        argtypes = []
        for p in info['params']:
            if p['type'] == 4: argtypes.append(ctypes.c_char_p)
            else: argtypes.append(ctypes.c_int)
        c_func.argtypes = argtypes
        c_func.restype = ctypes.c_int 

    def __getattr__(self, name):
        if name not in self.functions: raise AttributeError(f"DLL 中未找到函数: {name}")
        def wrapper(**kwargs):
            return self._invoke(name, kwargs)
        return wrapper

    def _invoke(self, name, kwargs):
        info = self.functions[name]
        c_func = getattr(self.dll, name)
        args = []
        for p in info['params']:
            pname = p['name']
            if pname in kwargs:
                val = kwargs[pname]
                if p['type'] == 4 and isinstance(val, str):
                    val = val.encode('gbk') # 路径编码
                args.append(val)
            else:
                args.append(0)
        return c_func(*args)

# ============================================================================
# 🚀 主程序
# ============================================================================
def main():
    print("🚀 启动批量特征识别...")
    
    if not os.path.exists(DLL_PATH):
        print(f"❌ DLL 未找到: {DLL_PATH}")
        return

    # 准备输出目录
    global_output_root = os.path.join(INPUT_FOLDER, f"{OUTPUT_DIR_NAME}")
    os.makedirs(global_output_root, exist_ok=True)
    print(f"📂 结果将保存至: {global_output_root}")

    # 获取 PRT 列表
    prt_files = glob.glob(os.path.join(INPUT_FOLDER, "*.prt"))
    print(f"📂 发现 {len(prt_files)} 个文件")

    session = NXOpen.Session.GetSession()
    
    # 加载 DLL
    try:
        plugin = UniversalLoader(DLL_PATH)
        print("✅ DLL 加载成功")
    except Exception as e:
        print(f"❌ DLL 加载失败: {e}")
        return

    for i, file_path in enumerate(prt_files):
        file_name = os.path.basename(file_path)
        print(f"\n[{i+1}/{len(prt_files)}] 处理: {file_name}")
        
        base_part = None
        try:
            # 1. 打开部件
            base_part, _ = session.Parts.OpenBaseDisplay(file_path)
            
            # 2. [关键步骤] 确保 CAM 环境就绪
            # 如果这一步失败，C++ 会报 CAM Setup is NULL
            if not ensure_cam_setup_ready(session, base_part):
                print("   ❌ 无法初始化 CAM 环境，跳过此文件。")
                continue

            # 3. 调用 DLL 函数
            # 参数名必须与 C++ JSON 中的 name 一致
            ret = plugin.RunFeatureRecognition(
                output_dir=global_output_root,
                target_layer=TARGET_LAYER
            )
            
            if ret == 0:
                print("   ✅ 识别成功")
            else:
                print(f"   ⚠️ 识别失败或无特征 (Code: {ret})")

        except Exception as e:
            print(f"   ❌ 异常: {e}")
        finally:
            # 关闭部件 (保存修改，因为我们可能创建了 CAM Setup)
            if base_part:
                try:
                    # 如果创建了 Setup，需要保存，否则下次打开还是没有
                    save_mode = NXOpen.BasePart.CloseModified.CloseModified
                    base_part.Close(NXOpen.BasePart.CloseWholeTree.TrueValue, save_mode, None)
                except: pass
            base_part = None
            gc.collect()

    print("\n🎉 全部完成")

if __name__ == "__main__":
    main()