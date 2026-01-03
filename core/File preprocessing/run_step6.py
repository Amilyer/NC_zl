# -*- coding: utf-8 -*-
"""
步骤 6: 图层处理与特征清理 (run_step6.py)
功能：
1. 读取 Step 5 处理后的 PRT 文件
2. 批量执行图层标准化：将所有层移动到 Layer 1
3. 复制图层：Layer 1 -> Layer 20
4. 特征清理：在 Layer 20 上删除孔、删除指定颜色面、移除参数
注意：
   本脚本只负责文件遍历和调用，具体业务逻辑由 layer_manager 和 feature_cleaner 模块实现。
"""

import os
import shutil
import glob
import traceback
import config
from path_manager import PathManager
import gc

# 导入业务模块
try:
    import layer_manager
    import feature_cleaner
    import NXOpen
    _MODULES_LOADED = True
except ImportError:
    print("❌ 无法导入必要模块 (layer_manager/feature_cleaner/NXOpen)")
    _MODULES_LOADED = False

def run_step6_logic(pm: PathManager):
    """步骤 6 联合处理逻辑 (单线程顺序执行)"""
    
    if not _MODULES_LOADED:
        print("❌ 模块未加载，跳过执行")
        return

    # 1. 路径准备
    input_dir = pm.get_textured_prt_dir()
    output_dir = pm.get_cleaned_prt_dir()
    
    if not os.path.exists(input_dir):
        print(f"❌ 输入目录不存在: {input_dir}")
        return

    # 确保输出目录存在 (如果存在则先清理)
    if os.path.exists(output_dir):
        try:
            shutil.rmtree(output_dir)
            print(f"🗑️ 已清理输出目录: {output_dir}")
        except Exception as e:
            print(f"⚠️ 清理目录失败: {e}")
            
    os.makedirs(output_dir, exist_ok=True)

    # 获取文件列表
    prt_files = glob.glob(os.path.join(input_dir, "*.prt"))
    if not prt_files:
        print(f"⚠️ 输入目录为空: {input_dir}")
        return

    print(f"📂 找到 {len(prt_files)} 个 PRT 文件，开始处理...")

    # 2. 初始化工具
    try:
        session = NXOpen.Session.GetSession()
        lm = layer_manager.LayerManager()
        fc = feature_cleaner.FeatureCleaner()
    except Exception as e:
        print(f"❌ 初始化 NX 工具失败: {e}")
        return

    # 3. 循环处理
    success_count = 0
    
    for i, file_path in enumerate(prt_files):
        file_name = os.path.basename(file_path)
        print(f"\nProcessing [{i+1}/{len(prt_files)}]: {file_name}")
        
        base_part = None
        try:
            # 打开部件
            base_part, _ = session.Parts.OpenBaseDisplay(file_path)
            
            # --- 核心业务逻辑 CALL START ---
            
            # A. 图层归一化: Move All -> Layer 1
            print(f"  > [1/3] 归一化图层 (Move All -> {config.LAYER_SOURCE})...")
            lm.process_part(base_part, config.LAYER_SOURCE)
            
            # B. 复制图层: Layer 1 -> Layer 20
            print(f"  > [2/3] 复制图层 ({config.LAYER_SOURCE} -> {config.LAYER_TARGET})...")
            # 修复：删除孔.py 内部已经包含了复制逻辑，此处不需要再次复制，否则会产生重叠体
            # lm.copy_layer_objects(base_part, config.LAYER_SOURCE, config.LAYER_TARGET)
            
            # C. 特征清理: Layer 20 (Holes, Color Faces, Params)
            print(f"  > [3/3] 特征清理 (Layer {config.LAYER_TARGET}, Color {config.COLOR_INDEX_TARGET})...")
            fc.clean_part(base_part, config.LAYER_TARGET, config.COLOR_INDEX_TARGET)
            
            # --- 核心业务逻辑 CALL END ---
            
            # 保存结果
            output_path = os.path.join(output_dir, file_name)
            
            # 如果输出文件已存在，先尝试删除（防止 SaveAs 报错）
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                    print(f"  ℹ️ 已删除旧文件: {output_path}")
                except Exception as e:
                    print(f"  ⚠️ 删除旧文件失败 (可能被占用): {e}")

            base_part.SaveAs(output_path)
            print(f"  ✅ 保存成功: {output_path}")
            success_count += 1

        except Exception as e:
            print(f"  ❌ 处理失败: {e}")
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

    print(f"\n🎉 步骤 6 完成! 成功: {success_count}/{len(prt_files)}")
    print("=" * 60)

def main():
    print("=" * 60)
    print("🚀 步骤 6: 图层处理与特征清理 (单线程托管版)")
    print("=" * 60)

    pm = PathManager(config.FILE_INPUT_PRT, config.FILE_INPUT_DXF)
    run_step6_logic(pm)

if __name__ == "__main__":
    main()
