"""
Step 11: 生成 CAM 刀路配置文件 (JSON)
主要功能：
1. 遍历 Step 8 输出目录下的所有 PRT 文件
2. 调用 '生成爬面文件.py' 生成爬面/往复等高 JSON
3. 调用 '生成面铣文件.py' 生成面铣 JSON（半精铣/全精铣）
"""

import importlib.util
import os
import sys

# 添加当前目录到 path 以便导入模块
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from path_manager import init_path_manager
except ImportError:
    # 尝试在上级目录查找
    sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from path_manager import init_path_manager

def import_crawling_module():
    """动态导入生成爬面文件模块"""
    module_name = "generate_crawling_files"
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, "生成爬面文件.py")
    
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到文件: {file_path}")

    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def import_face_milling_module():
    """动态导入生成面铣文件模块"""
    module_name = "generate_face_milling_files"
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, "生成面铣文件.py")
    
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到文件: {file_path}")

    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    spec.loader.exec_module(module)
    return module

def import_spiral_module():
    """动态导入生成螺旋文件模块"""
    module_name = "generate_spiral_files"
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, "生成螺旋文件.py")
    
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到文件: {file_path}")

    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

def import_corner_cleaning_module():
    """动态导入生成清角文件模块"""
    module_name = "generate_corner_cleaning_files"
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, "生成清角文件.py")
    
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到文件: {file_path}")

    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

def main():
    print("=" * 80)
    print("  Step 11: 生成 CAM 刀路配置文件 (爬面/往复等高/面铣/螺旋)")
    print("=" * 80)

    # 1. 初始化 PathManager
    # 1. 初始化 PathManager
    try:
        import config
        pm = init_path_manager(config.FILE_INPUT_PRT, config.FILE_INPUT_DXF)
    except Exception as e:
        # Fallback if config import fails (unlikely)
        print(f"[WARN] Config loading failed: {e}")
        pm = init_path_manager(r"C:\Projects\NC\input\3D_PRT", r"C:\Projects\NC\input\2D_DXF")
    
    # 2. 导入生成模块
    try:
        crawling_module = import_crawling_module()
        face_milling_module = import_face_milling_module()
        spiral_module = import_spiral_module()
        corner_cleaning_module = import_corner_cleaning_module()
    except Exception as e:
        print(f"[ERROR] 无法导入生成模块: {e}")
        return

    # 3. 获取输入/输出目录
    # 修改：直接使用 Step 8 的输出 (04_PRT_with_Tool)，跳过 Step 9
    prt_dir = pm.get_final_prt_dir()
    output_dir = pm.get_cam_json_dir()     # 05_CAM/Toolpath_JSON

    print(f"[INFO] 🔗 输入目录 (Step 8 Output, Skipping Step 9): {prt_dir}")
    print(f"[INFO] 📂 输出目录: {output_dir}")

    if not prt_dir.exists():
        print(f"[ERROR] 零件目录不存在: {prt_dir}")
        return

    # 3.1 清理输出目录
    if output_dir.exists():
        print(f"[INFO] 正在清理旧的 JSON 文件: {output_dir}")
        for old_json in output_dir.glob("*.json"):
            try:
                old_json.unlink()
            except Exception as e:
                print(f"  [WARN] 删除失败 {old_json.name}: {e}")
    else:
        output_dir.mkdir(parents=True, exist_ok=True)

    # 4. 遍历处理
    prt_files = list(prt_dir.glob("*.prt"))
    if not prt_files:
        print("[WARN] 目录中没有 PRT 文件")
        return

    success_count = 0
    fail_count = 0

    for prt_file in prt_files:
        part_name = prt_file.stem
        print(f"\n>> 正在处理: {part_name}")

        # 构造输入文件路径
        # .../03_Analysis/Navigator_Reports/csv/{part_name}_FeatureRecognition_Log.csv
        feature_log = pm.get_nav_csv_dir() / f"{part_name}_FeatureRecognition_Log.csv"
        
        # .../03_Analysis/Face_Info/face_csv/{part_name}_face_data.csv
        face_data = pm.get_face_csv_dir() / f"{part_name}_face_data.csv"
        
        # .../03_Analysis/Geometry_Analysis/{part_name}.csv
        direction_csv = pm.get_analysis_geo_dir() / f"{part_name}.csv"
        
        tool_json = pm.get_tool_params_json()

        # 检查必要文件是否存在
        missing_files = []
        if not feature_log.exists(): missing_files.append(f"特征日志 ({feature_log.name})")
        if not face_data.exists(): missing_files.append(f"面数据 ({face_data.name})")
        if not direction_csv.exists(): missing_files.append(f"方向分析 ({direction_csv.name})")

        if missing_files:
            print(f"[SKIP] 跳过 {part_name}，缺少文件: {', '.join(missing_files)}")
            fail_count += 1
            continue

        try:
            # 调用 main1
            # 签名: main1(prt_folder, feature_log_csv, face_data_csv, direction_csv, tool_json, output_dir)
            # 注意：prt_folder 参数现在传入的是完整的 PRT 文件路径
            crawling_module.main1(
                str(prt_file),
                str(feature_log),
                str(face_data),
                str(direction_csv),
                str(tool_json),
                str(output_dir)
            )
            
            # 调用面铣模块
            # 签名: main1(csv_face, csv_tag, out_dir, prt_folder, excel_params, tool_json)
            print("[INFO] 正在生成面铣文件...")
            face_milling_module.main1(
                csv_face=str(face_data),
                csv_tag=str(direction_csv),
                out_dir=str(output_dir),
                prt_folder=str(prt_file),
                excel_params=str(pm.get_part_params_excel()),
                tool_json=str(tool_json)
            )

            # 调用螺旋模块
            # 签名: main1(prt_folder, face_data_csv, csv_file, json_file, direction_file, output_dir, excel_params)
            print("[INFO] 正在生成螺旋文件...")
            spiral_module.main1(
                prt_folder=str(prt_file),
                face_data_csv=str(face_data),
                csv_file=str(feature_log),
                json_file=str(tool_json),
                direction_file=str(direction_csv),
                output_dir=str(output_dir),
                
            )
            
            # 调用清角模块
            # 签名: main1(raw_input_files, face_csv, tool_json, part_xlsx, direction_file, output_dir)
            print("[INFO] 正在生成清角文件...")
            # 构造清角模块需要的输入文件列表 (来自螺旋模块的输出)
            spiral_json_1 = output_dir / f"{part_name}_半精_螺旋.json"
            spiral_json_2 = output_dir / f"{part_name}_半精_螺旋_往复等高.json"
            
            # 确保文件存在才加入列表 (虽然刚生成应该存在，但为了健壮性)
            corner_input_files = []
            if spiral_json_1.exists(): corner_input_files.append(str(spiral_json_1))
            if spiral_json_2.exists(): corner_input_files.append(str(spiral_json_2))
            
            if corner_input_files:
                corner_cleaning_module.main1(
                    raw_input_files=corner_input_files,
                    face_csv=str(face_data),
                    tool_json=str(tool_json),
                    part_xlsx=str(pm.get_part_params_excel()),
                    direction_file=str(direction_csv),
                    output_dir=str(output_dir)
                )
            else:
                print(f"[WARN] 没有找到螺旋模块的输出文件，跳过生成清角文件: {part_name}")

            success_count += 1
        except Exception as e:
            print(f"[ERROR] 处理 {part_name} 时发生异常: {e}")
            import traceback
            traceback.print_exc()
            fail_count += 1

    print("\n" + "=" * 80)
    print("  Step 11 完成")
    print(f"  成功: {success_count}")
    print(f"  失败: {fail_count}")
    print("=" * 80)

if __name__ == "__main__":
    main()