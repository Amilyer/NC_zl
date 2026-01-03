# -*- coding: utf-8 -*-
import ctypes
import json
import os
import sys
import glob
import gc

# ============================================================================
# ⚙️ 配置区域 (请根据实际情况修改)
# ============================================================================
# 1. 您的 C++ DLL 路径 (确保指向最新编译的 strict priority 版本)
DLL_PATH = r"D:\cc++_pro\final_jiejue\final\最终_根据加工顺序生成加工方向\x64\Debug\NX_Open_Wizard1.dll"

# 2. PRT 输入文件夹 (包含 .prt 文件)
INPUT_FOLDER = r"C:\Users\Admin\Desktop\test"

# 3. 优先级/特征 CSV 所在文件夹 (通常与 PRT 在一起)
FEATURE_CSV_FOLDER = r"C:\Users\Admin\Desktop\test\Geometry_Analysis_Reports"

# 4. CSV 的文件名后缀匹配规则
#    例如: PRT名是 "DIE-03.prt", CSV名是 "DIE-03.csv"
FEATURE_CSV_SUFFIX = ".csv"

# 5. 结果输出文件夹名称 (将自动创建在 INPUT_FOLDER 下)
OUTPUT_DIR_NAME = "final_direction"

# 6. 目标图层 (0=分析所有图层, 1-256=指定图层)
TARGET_LAYER = 20
# ============================================================================

try:
    import NXOpen
    import NXOpen.UF
except ImportError:
    print("❌ 错误: 必须在 NX 环境下运行 (File -> Execute -> NX Open...)")
    sys.exit(1)


# ============================================================================
# 🔧 通用加载器 (无需修改，自动适配 C++ 注册宏)
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
            # 4 代表字符串(char*), 其他视为 int
            if p['type'] == 4:
                argtypes.append(ctypes.c_char_p)
            else:
                argtypes.append(ctypes.c_int)
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
                # 字符串转字节流 (GBK 兼容 Windows 中文路径)
                if p['type'] == 4 and isinstance(val, str):
                    val = val.encode('gbk')
                args.append(val)
            else:
                print(f"⚠️ 参数 '{pname}' 未提供，默认传 0")
                args.append(0)

        # 调用 C++ 函数并返回结果
        return c_func(*args)


# ============================================================================
# 🚀 主程序
# ============================================================================
def main():
    print("🚀 启动几何分析 (Strict Priority Mode)...")

    if not os.path.exists(DLL_PATH):
        print(f"❌ DLL 未找到: {DLL_PATH}")
        return

    # 准备输出目录
    global_output_dir = os.path.join(INPUT_FOLDER, OUTPUT_DIR_NAME)
    os.makedirs(global_output_dir, exist_ok=True)
    print(f"📂 结果保存目录: {global_output_dir}")

    # 扫描 PRT 文件
    prt_files = glob.glob(os.path.join(INPUT_FOLDER, "*.prt"))
    print(f"📂 发现 {len(prt_files)} 个 PRT 文件")

    # 获取 NX 会话
    session = NXOpen.Session.GetSession()

    # 加载 DLL
    try:
        plugin = UniversalLoader(DLL_PATH)
        print(f"✅ DLL 加载成功")
        # 打印一下注册的函数，方便调试
        print(f"   已注册接口: {json.dumps(plugin.functions, indent=2)}")
    except Exception as e:
        print(f"❌ DLL 加载失败: {e}")
        return

    # 循环处理文件
    success_count = 0
    for i, file_path in enumerate(prt_files):
        file_name = os.path.basename(file_path)
        print(f"\n[{i + 1}/{len(prt_files)}] 处理: {file_name}")

        base_part = None
        try:
            # 1. 打开部件
            base_part, _ = session.Parts.OpenBaseDisplay(file_path)

            # 2. 构造路径
            part_name_only = os.path.splitext(file_name)[0]

            # A. 结果输出 CSV 路径
            output_csv_path = os.path.join(global_output_dir, f"{part_name_only}.csv")

            # B. 优先级输入 CSV 路径 (根据配置拼接)
            # 逻辑：在 FEATURE_CSV_FOLDER 中寻找 "文件名 + 后缀"
            priority_csv_path = os.path.join(FEATURE_CSV_FOLDER, part_name_only + FEATURE_CSV_SUFFIX)

            # 检查特征文件是否存在
            if not os.path.exists(priority_csv_path):
                print(f"   ⚠️ 警告: 未找到优先级 CSV 文件: {priority_csv_path}")
                print(f"      C++ 将使用默认优先级 (+Z > -Z ...)")
                # 如果文件不存在，传空字符串或不存在的路径，C++ 端会 handle 成默认值
            else:
                print(f"   📄 加载优先级定义: {os.path.basename(priority_csv_path)}")

            print(f"   -> 正在调用 C++ 分析模块...")

            # 3. 调用 DLL 接口: RunGeometryAnalysis
            # 【重要】参数名必须与 C++ 代码中 PARAM() 定义的一致:
            # PARAM(input_csv_path, TYPE_STRING)
            # PARAM(output_csv_path, TYPE_STRING)
            # PARAM(target_layer, TYPE_INT)

            ret_code = plugin.RunGeometryAnalysis(
                input_csv_path=priority_csv_path,  # <--- 修改此处名称匹配 C++
                output_csv_path=output_csv_path,
                target_layer=TARGET_LAYER
            )

            # 4. 检查返回值
            if ret_code == 0:
                print("   ✅ 分析成功，CSV 已生成。")
                success_count += 1
            elif ret_code == 2:
                print("   ⚠️ 分析完成但无结果 (可能是空图层或无特征)。")
            else:
                print(f"   ❌ 分析失败 (Code: {ret_code})")

        except Exception as e:
            print(f"   ❌ Python 异常: {e}")

        finally:
            # 5. 关闭部件 (不保存，只读取)
            if base_part:
                try:
                    base_part.Close(NXOpen.BasePart.CloseWholeTree.TrueValue,
                                    NXOpen.BasePart.CloseModified.CloseModified,
                                    None)
                except:
                    pass

            base_part = None
            gc.collect()

    print(f"\n🎉 全部完成! 成功处理 {success_count}/{len(prt_files)} 个文件。")


if __name__ == "__main__":
    main()