# -*- coding: utf-8 -*-
"""
步骤 7: 联合提取与分析 (run_step7.py) [单进程版]啦啦啦啦啦
功能：
1. 读取 Step 6 处理后的 PRT (Cleaned)
2. 4阶段流程:
   - Phase 1: 面信息提取
   - Phase 2: 导航器提取 (Layer 20) -> 生成 CSV 和 PRT
   - Phase 3: 沉头数量统计 (调用 获取沉头数量.py)
   - Phase 4: 几何分析 (使用 Phase 3 的 CSV 作为输入)
3. 另存为到 output/03_Analysis/Face_Info/prt (供 Step 8 使用)
"""

import os
import sys
import shutil
import traceback
import importlib.util
import config
from path_manager import PathManager

# 导入功能模块
try:
    import NXOpen
    from face_extractor import FaceExtractor
    from navigator_extractor import NavigatorExtractor
    from navigator_extractor import NavigatorExtractor
    from geometry_strict_runner import GeometryStrictRunner # New runner
    from nx_processor import NXProcessor
except ImportError:
    pass

def load_counterbore_module():
    """动态导入 '获取沉头数量.py'"""
    module_name = "get_counterbore_count"
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, "获取沉头数量.py")
    
    if os.path.exists(file_path):
        spec = importlib.util.spec_from_file_location(module_name, file_path)
        module = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
        return module
    else:
        raise FileNotFoundError(f"找不到文件: {file_path}")

def run_step7_logic(pm: PathManager):
    print("=" * 60)
    print("🚀 步骤 7: 联合提取与分析 (4阶段流程)")
    print("=" * 60)

    # 1. 路径准备
    input_dir = str(pm.get_cleaned_prt_dir()) 
    
    # 最终输出: Step 8 Input
    output_final_dir = str(pm.get_analysis_face_prt_dir())
    
    # 子目录准备
    dir_face = str(pm.get_analysis_face_dir())
    dir_nav_20_csv = str(pm.get_nav_csv_dir())
    dir_nav_20_prt = str(pm.get_nav_prt_dir())
    dir_counterbore = str(pm.get_counterbore_csv_dir()) # Phase 3 Output
    dir_geo = str(pm.get_analysis_geo_dir())     # Phase 4 Output (Flattened as per request)
    
    print(f"📂 输入目录: {input_dir}")
    print(f"📂 最终输出: {output_final_dir}")
    print("-" * 50)

    if not os.path.exists(input_dir):
        print(f"❌ 输入目录不存在: {input_dir}")
        return

    # 清理并重建输出目录
    dirs_to_clean = [
        output_final_dir, dir_face, 
        dir_nav_20_csv, dir_nav_20_prt,
        dir_counterbore, dir_geo
    ]
    for d in dirs_to_clean:
        if os.path.exists(d):
            try: shutil.rmtree(d)
            except: pass
        os.makedirs(d, exist_ok=True)

    # 2. 初始化提取器
    try:
        fe = FaceExtractor(str(pm.get_face_info_dll_path()))
        ne = NavigatorExtractor(str(pm.get_navigator_dll_path()))
        # Phase 4 (Strict Geometry) - 使用新的 Runner
        ga20_runner = GeometryStrictRunner(str(pm.get_geometry_analysis_dll_path_20()))
        
        # 动态加载沉头数量模块
        counterbore_mod = load_counterbore_module()
        ProcessInfoHandler = counterbore_mod.ProcessInfoHandler
        
        print("✅ 模块加载成功")
    except Exception as e:
        print(f"❌ 初始化提取器失败: {e}")
        traceback.print_exc()
        return

    # 3. 获取文件列表
    prt_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.prt')]
    total = len(prt_files)
    
    if not prt_files:
        print("❌ 未找到 PRT 文件")
        return

    # 4. 启动 NX 会话
    session = NXOpen.Session.GetSession()
    nx_proc = NXProcessor() 

    print(f"📂 开始处理 {total} 个文件...")
    
    success_count = 0
    
    for idx, filename in enumerate(prt_files):
        # 基础文件路径
        input_path = os.path.join(input_dir, filename)
        final_output_path = os.path.join(output_final_dir, filename)
        
        # 中间路径
        prt_20_path = os.path.join(dir_nav_20_prt, filename)
        csv_name = filename.replace('.prt', '.csv')
        
        prefix = f"[{idx+1}/{total}] {filename}"
        print(f"\n{prefix} 处理中...")
        
        try:
            # =================================================================
            # Phase 1: 面信息提取 (Input: Cleaned PRT)
            # =================================================================
            print("  > [1/4] 面信息提取...")
            if not nx_proc.open_part(input_path):
                print(f"  ❌ 无法打开部件: {input_path}")
                continue
            
            work_part = session.Parts.Work
            fe.process_part(dir_face, target_layer=config.LAYER_FACE_INFO_TARGET)
            
            # =================================================================
            # Phase 2: 导航器提取 (Layer 20)
            # =================================================================
            print("  > [2/4] 导航器提取 (Layer 20)...")
            
            # 提取导航器信息
            ne.process_part(work_part, dir_nav_20_csv, target_layer=config.LAYER_NAV_20)
            
            # 保存中间 PRT 到 Layer 20 目录
            if os.path.exists(prt_20_path):
                try: os.remove(prt_20_path)
                except: pass
            work_part.SaveAs(prt_20_path)
            print(f"    (Saved Navigator PRT: {os.path.basename(prt_20_path)})")
            
            # =================================================================
            # Phase 3: 沉头数量统计 (Counterbore Count)
            # =================================================================
            print("  > [3/4] 沉头数量统计...")
            csv_counterbore = os.path.join(dir_counterbore, csv_name)
            
            try:
                # 实例化处理类 (传入当前 Session 和 WorkPart)
                # 注意：work_part 现在是 Phase 2 处理后的状态 (已包含 Layer 20 特征?)
                # 获取沉头数量脚本是读取 Notes 和几何信息
                handler = ProcessInfoHandler(session, work_part)
                handler.get_hole_num(csv_counterbore)
                print(f"    (Generated Counterbore CSV: {os.path.basename(csv_counterbore)})")
            except Exception as e:
                print(f"    ❌ 沉头统计失败: {e}")
                traceback.print_exc()
                # 如果此步失败，Phase 4 也会受影响
                continue

            # =================================================================
            # Phase 4: 几何分析 (使用 Phase 3 的结果)
            # =================================================================
            print("  > [4/4] 几何分析 (Strict Priority)...")
            csv_geo_final = os.path.join(dir_geo, csv_name)
            
            if not os.path.exists(csv_counterbore):
                print("    ⚠️ 未找到沉头 CSV，无法进行几何分析")
            else:
                # 调用 DLL (使用 Runner)
                res = ga20_runner.run_analysis(priority_csv_path=str(csv_counterbore), output_csv_path=str(csv_geo_final), target_layer=config.LAYER_NAV_20)
                if res == 0:
                     print(f"    ✅ 几何分析完成 (Output: {os.path.basename(csv_geo_final)})")
                else:
                     print(f"    ⚠️ 几何分析返回非零代码: {res}")

            # =================================================================
            # Final: 另存为最终结果
            # =================================================================
            # Step 8 requires this PRT (Wait, Step 8 config calls get_nav_layer20_prt_dir()?)
            # 任务描述说: "另存为到 output/03_Analysis/Face_Info/prt (供 Step 8 使用)"
            # 但之前 Step 8 配置改成了 Layer 20 PRT。
            # 为了保险，我们还是按原始需求保存到 Face_Info/prt (run_step7.py原始注释这么写的)
            # 并且让 Step 8 能够找到它。
            # 不过 wait，implementation plan 说 Step 8 input configured to Layer 20 PRT.
            # 既然 Phase 2 已经保存了 Layer 20 PRT，这里是否还需要保存？
            # 这里的 SaveAs(output_final_dir) 是 Step 7 的最终产出。
            
            output_final_path = os.path.join(output_final_dir, filename)
            if os.path.exists(output_final_path):
                try: os.remove(output_final_path)
                except: pass
            work_part.SaveAs(output_final_path)
            print(f"  ✅ 最终保存: {os.path.basename(output_final_path)}")
            
            nx_proc.close_all()
            success_count += 1

        except Exception as e:
            print(f"  ❌ 处理异常: {e}")
            traceback.print_exc()
            try: nx_proc.close_all()
            except: pass
            
    print("-" * 50)
    print(f"🎉 步骤 7 完成! 成功: {success_count}/{total}")

def main():
    pm = PathManager(config.FILE_INPUT_PRT, config.FILE_INPUT_DXF)
    run_step7_logic(pm)

if __name__ == "__main__":
    main()
