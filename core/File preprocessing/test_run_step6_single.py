# -*- coding: utf-8 -*-
"""
单独测试 Step 6 逻辑的脚本 (test_run_step6_single.py)
功能：
针对指定的一个 PRT 文件执行 Step 6 的完整逻辑：
1. 图层归一化 (Layer 1)
2. 特征清理 (Layer 20: 删除孔、删除指定颜色面、移除参数)
3. 另存为到测试目录
"""

import os
import shutil
import traceback
import gc

# 确保项目根目录在 sys.path
import sys
from pathlib import Path

current_dir = Path(__file__).parent.absolute()
if str(current_dir) not in sys.path:
    sys.path.insert(0, str(current_dir))

import config
from path_manager import PathManager

# 导入业务模块
try:
    import layer_manager
    import feature_cleaner
    import NXOpen
    _MODULES_LOADED = True
except ImportError:
    print("❌ 无法导入必要模块 (layer_manager/feature_cleaner/NXOpen)")
    _MODULES_LOADED = False

def run_step6_single_file(file_path):
    """
    对单个文件执行 Step 6 逻辑
    :param file_path: 输入文件的绝对路径
    """
    
    if not _MODULES_LOADED:
        print("❌ 模块未加载，跳过执行")
        return

    # 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"❌ 文件不存在: {file_path}")
        return

    print(f"🚀 开始测试 Step 6 (单文件模式)")
    print(f"📄 目标文件: {file_path}")

    # 定义输出目录 (测试用)
    output_dir = os.path.join(os.path.dirname(file_path), "test_step6_output")
    os.makedirs(output_dir, exist_ok=True)
    
    file_name = os.path.basename(file_path)
    output_path = os.path.join(output_dir, file_name)
    
    print(f"📂 输出目录: {output_dir}")

    # 初始化工具
    try:
        session = NXOpen.Session.GetSession()
        lm = layer_manager.LayerManager()
        fc = feature_cleaner.FeatureCleaner()
    except Exception as e:
        print(f"❌ 初始化 NX 工具失败: {e}")
        return

    base_part = None
    try:
        # 打开部件
        print("  > 打开部件...")
        base_part, _ = session.Parts.OpenBaseDisplay(file_path)
        
        # --- 核心业务逻辑 CALL START ---
        
        # A. 图层归一化: Move All -> Layer 1
        print(f"  > [1/3] 归一化图层 (Move All -> {config.LAYER_SOURCE})...")
        lm.process_part(base_part, config.LAYER_SOURCE)
        
        # B. 复制图层 (该步骤在 process_part 或 clean_part 中可能涉及，参照 run_step6.py 注释，显式复制已被注释掉)
        # print(f"  > [2/3] 复制图层 ({config.LAYER_SOURCE} -> {config.LAYER_TARGET})...")
        # lm.copy_layer_objects(base_part, config.LAYER_SOURCE, config.LAYER_TARGET)
        
        # C. 特征清理: Layer 20 (Holes, Color Faces, Params)
        print(f"  > [3/3] 特征清理 (Layer {config.LAYER_TARGET}, Color {config.COLOR_INDEX_TARGET})...")
        fc.clean_part(base_part, config.LAYER_TARGET, config.COLOR_INDEX_TARGET)
        
        # --- 核心业务逻辑 CALL END ---
        
        # 保存结果
        # 如果输出文件已存在，先尝试删除
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except Exception as e:
                print(f"  ⚠️ 删除旧文件失败: {e}")

        print(f"  > 另存为: {output_path}")
        base_part.SaveAs(output_path)
        print(f"✅ 测试成功! 文件已保存。")

    except Exception as e:
        print(f"❌ 处理失败: {e}")
        traceback.print_exc()
    finally:
        # 关闭部件
        if base_part:
            try:
                base_part.Close(NXOpen.BasePart.CloseWholeTree.TrueValue, 
                              NXOpen.BasePart.CloseModified.CloseModified, 
                              None)
            except:
                pass
        # 垃圾回收
        base_part = None
        gc.collect()

    print("=" * 60)

if __name__ == "__main__":
    # 在这里指定要测试的文件路径
    # 示例: C:\Projects\NC\output\02_Textured_PRT\xxx.prt
    # 请根据实际情况修改下面的路径
    target_file = r"C:\Projects\NC\output\04_PRT_with_Tool\GU-04.prt" 
    
    # 也可以使用命令行参数传入
    if len(sys.argv) > 1:
        target_file = sys.argv[1]

    run_step6_single_file(target_file)
