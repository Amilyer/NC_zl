# -*- coding: utf-8 -*-
import os
import sys
import importlib.util
import traceback
import NXOpen
from pathlib import Path

# 添加项目根目录到 sys.path
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(os.path.dirname(current_dir))
if project_root not in sys.path:
    sys.path.append(project_root)

from path_manager import PathManager, get_path_manager

def load_eo_module():
    """Dynamically load the 创建工程单 module"""
    module_path = os.path.join(current_dir, "创建工程单.py")
    spec = importlib.util.spec_from_file_location("eo_module", module_path)
    eo_module = importlib.util.module_from_spec(spec)
    sys.modules["eo_module"] = eo_module
    spec.loader.exec_module(eo_module)
    return eo_module

def run_step15_logic(pm: PathManager):
    print("🚀 开始执行 Step 15: 生成工程单Excel ...")
    
    # 1. Load module
    try:
        eo_module = load_eo_module()
        print("✅ 工程单生成模块加载成功")
    except Exception as e:
        print(f"❌ 加载模块失败: {e}")
        return

    # 2. Setup paths
    eo_root = pm.get_engineering_order_root()
    txt_dir = pm.get_eo_txt_dir()
    dims_dir = pm.get_eo_dims_dir()
    json_dir = pm.get_eo_json_dir()
    output_excel_dir = pm.get_eo_excel_dir()
    
    print(f"📂 工件信息目录: {txt_dir}")
    print(f"📂 尺寸信息目录: {dims_dir}")
    print(f"📂 JSON数据目录: {json_dir}")
    print(f"📂 Excel输出目录: {output_excel_dir}")
    
    # 清理输出目录
    import shutil
    if output_excel_dir.exists():
        try:
            shutil.rmtree(output_excel_dir)
            print(f"🗑️ 已清理输出目录: {output_excel_dir}")
        except Exception as e:
            print(f"⚠️ 清理目录失败: {e}")
    output_excel_dir.mkdir(parents=True, exist_ok=True)
    
    if not txt_dir.exists():
        print(f"⚠️ 工件信息目录不存在: {txt_dir}")
        return

    # 3. Find files to process
    # Iterate through JSON files as they are the main data source, or TXT
    # Let's use JSON files as the driver since step 14 generates them.
    json_files = list(json_dir.glob("*.json"))
    if not json_files:
        print("⚠️ 未找到JSON数据文件的工件")
        return

    print(f"📋 发现 {len(json_files)} 个工件数据，开始生成Excel...")

    for i, json_path in enumerate(json_files):
        part_name = json_path.stem # e.g. "Model1"
        print(f"\n[{i+1}/{len(json_files)}] 正在处理: {part_name}")
        
        # Construct corresponding paths
        workpiece_txt_path = txt_dir / f"{part_name}.txt"
        dims_txt_path = dims_dir / f"{part_name}_尺寸.txt"
        
        # Check existence
        if not workpiece_txt_path.exists():
            print(f"   ⚠️ 缺少工件信息TXT: {workpiece_txt_path}")
            # continue? Or let the module handle it (it warns).
            
        if not dims_txt_path.exists():
            print(f"   ⚠️ 缺少尺寸信息TXT: {dims_txt_path}")
            
        try:
            # Call main logic
            # main(workpiece_txt_path, dims_txt_path, json_path, excel_save_dir, tool_excel_path, image_folder)
            tool_excel_path = pm.get_tool_params_excel_path()
            screenshot_root = pm.get_eo_screenshot_dir()
            image_folder = screenshot_root / part_name
            
            if not image_folder.exists():
                print(f"   ⚠️ 缺少截图文件夹: {image_folder}")
            
            eo_module.main(
                str(workpiece_txt_path), 
                str(dims_txt_path), 
                str(json_path), 
                str(output_excel_dir),
                str(tool_excel_path),
                str(image_folder)
            )
            print(f"   ✅ 生成成功 -> {output_excel_dir / f'{part_name}.xlsx'}")
        except Exception as e:
            print(f"   ❌ 生成失败: {e}")
            traceback.print_exc()

    print("\n✅ Step 15 (Excel Generation) 完成")

def main():
    import config
    from path_manager import init_path_manager
    pm = init_path_manager()
    run_step15_logic(pm)

if __name__ == "__main__":
    main()

