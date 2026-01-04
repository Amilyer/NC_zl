"""
Step 10: 生成开粗行腔文件 (run_step10.py)
春江潮水连海平，海上明月共潮生
涟涟随波千万里，何处春江无月明
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

def import_cavity_roughing_module():
    """动态导入生成开粗行腔文件模块"""
    module_name = "generate_cavity_roughing_files"
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, "生成开粗行腔文件.py")
    
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到文件: {file_path}")

    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    spec.loader.exec_module(module)
    return module


def import_zlevel_roughing_module():
    """动态导入生成开粗往复等高模块"""
    module_name = "generate_zlevel_roughing_files"
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, "生成开粗往复等高.py")
    
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到文件: {file_path}")

    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module



def main():
    print("=" * 80)
    print("  Step 10: 生成开粗行腔文件")
    print("=" * 80)

    # 1. 初始化
    pm = init_path_manager(r"C:\Projects\NC\input\3D_PRT", r"C:\Projects\NC\input\2D_DXF")
    pm = init_path_manager(r"C:\Projects\NC\input\3D_PRT", r"C:\Projects\NC\input\2D_DXF")
    cavity_module = import_cavity_roughing_module()
    zlevel_module = import_zlevel_roughing_module()


    # 2. 目录: 使用 Step 9 输出的钻孔 PRT 文件
    prt_dir = pm.get_step9_drilled_dir()
    
    # 新增：分开存放开粗JSON和PRT
    out_json_dir = pm.get_cam_roughing_json_dir()
    print(f"[INFO] 零件目录: {prt_dir}")
    print(f"[INFO] JSON输出目录: {out_json_dir}")

    # 3. 遍历
    prt_files = list(prt_dir.glob("*.prt"))
    if not prt_files:
        print("[WARN] 目录中没有 PRT 文件")
        return

    import shutil

    success_count = 0
    fail_count = 0

    for prt_file in prt_files:
        part_name = prt_file.stem
        print(f"\n>> 正在处理: {part_name}")

        # 标准化路径拼接 (基于已验证的命名规则)
        feature_csv = pm.get_nav_csv_dir() / f"{part_name}_FeatureRecognition_Log.csv"
        face_csv = pm.get_face_csv_dir() / f"{part_name}_face_data.csv"
        mach_csv = pm.get_analysis_geo_dir() / f"{part_name}.csv"
        knife_json = pm.get_tool_params_json()
        
        out_json = out_json_dir / f"{part_name}_行腔.json"
        
        # 确保关键文件存在
        if not (feature_csv.exists() and face_csv.exists() and mach_csv.exists()):
             missing = []
             if not feature_csv.exists(): missing.append("特征日志")
             if not face_csv.exists(): missing.append("面数据")
             if not mach_csv.exists(): missing.append("方向分析")
             print(f"[SKIP] 跳过，缺少文件: {', '.join(missing)}")
             print(f"       期望特征路径: {feature_csv}") # 打印出来方便排查
             fail_count += 1
             continue

        try:
            # 2.1 生成行腔开粗 JSON
            print("   -> 生成行腔开粗...")
            cavity_module.generate_cavity_json_v2(
                str(prt_file), # 使用原始PRT
                str(feature_csv),
                str(face_csv),
                str(mach_csv),
                str(knife_json),
                str(out_json),
                str(pm.get_part_params_excel()),
                str(pm.get_counterbore_csv_dir() / f"{part_name}.csv") # path_main: 沉头孔数量文件
            )
            
            # 2.2 生成往复等高开粗 JSON
            print("   -> 生成往复等高开粗...")
            zlevel_module.main1(
                str(prt_file),
                str(feature_csv),
                str(face_csv),
                str(mach_csv),
                str(knife_json),
                str(out_json_dir)  # 注意：这里传入的是目录，不是具体文件路径
            )
            
            success_count += 1
        except Exception as e:
            print(f"[ERROR] 处理 {part_name} 失败: {e}")
            import traceback
            traceback.print_exc()
            fail_count += 1

    print("\n" + "=" * 80)
    print(f"Step 10 完成: 成功 {success_count}, 失败 {fail_count}")
    print("=" * 80)

if __name__ == "__main__":
    main()
