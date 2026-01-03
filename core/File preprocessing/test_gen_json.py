# -*- coding: utf-8 -*-
"""
单独测试脚本: 调用 生成工程单json.py
"""
import os
import sys
import NXOpen
import importlib.util

# ==============================================================================
# 配置
# ==============================================================================
# 测试用的 PRT 文件路径 (请修改此处)
TEST_PRT_PATH = r"C:\Projects\NC\output\06_CAM\Final_CAM_PRT\DIE-03.prt"
# 输出目录
OUTPUT_DIR = r"C:\Users\admin\Desktop\新建文件夹"

# ==============================================================================

def load_module_from_file(module_name, file_path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

def main():
    if not os.path.exists(TEST_PRT_PATH):
        print(f"❌ 找不到测试文件: {TEST_PRT_PATH}")
        return

    # 1. 启动 NX Session
    try:
        session = NXOpen.Session.GetSession()
    except Exception as e:
        print(f"❌ 无法连接 NX Session: {e}")
        return

    # 2. 打开部件
    print(f"📂 正在打开: {TEST_PRT_PATH}")
    base_part, load_status = session.Parts.OpenBaseDisplay(TEST_PRT_PATH)
    
    if not session.Parts.Work:
        print("❌ 打开部件失败 (Work Part is None)")
        return

    # 3. 加载生成模块
    current_dir = os.path.dirname(os.path.abspath(__file__))
    json_generator_path = os.path.join(current_dir, "生成工程单json.py")
    
    try:
        json_gen_module = load_module_from_file("generate_eo_json", json_generator_path)
        print("✅ 模块加载成功")
    except Exception as e:
        print(f"❌ 模块加载失败: {e}")
        return

    # 4. 执行
    print("🚀 开始执行生成逻辑...")
    try:
        # 确保输出目录存在
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
            
        json_gen_module.main(OUTPUT_DIR)
        print("✅ 执行完成")
    except Exception as e:
        print(f"❌ 执行出错: {e}")
        import traceback
        traceback.print_exc()

    base_part.Close(NXOpen.BasePart.CloseWholeTree.TrueValue, NXOpen.BasePart.CloseModified.CloseModified, None)

if __name__ == "__main__":
    main()
