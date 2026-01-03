# -*- coding: utf-8 -*-
"""
步骤 5: 贴图自动化 (run_step5.py) [单进程版]
功能：
1. 读取 Step 4 输出的 PRT 文件
2. 批量将零件移动到原点并贴图
注意：本脚本采用单进程顺序执行，确保稳定性。
"""
import sys
import os
import shutil
import time
import traceback
import gc
import gc

import config
from path_manager import PathManager

# -----------------------------------------------------------------------------
# 环境配置 (NX)
# -----------------------------------------------------------------------------
NX_BASE_DIR = r"C:\Program Files\Siemens\NX2312" 
NX_PYTHON_DIR = os.path.join(NX_BASE_DIR, "NXBIN", "python")
NX_MANAGED_DIR = os.path.join(NX_BASE_DIR, "NXBIN", "managed")

for p in [NX_PYTHON_DIR, NX_MANAGED_DIR]:
    if os.path.exists(p) and p not in sys.path:
        sys.path.append(p)

# -----------------------------------------------------------------------------
# 导入依赖
# -----------------------------------------------------------------------------
try:
    import NXOpen
    from nx_processor import NXProcessor
    from texture_mapper import TextureMapper
    
    # 钻孔模块路径 (core/NX_Drilling_Automation2)
    DRILL_MODULE_PATH = os.path.join(config.PROJECT_ROOT, "core", "NX_Drilling_Automation2")
    if os.path.exists(DRILL_MODULE_PATH) and DRILL_MODULE_PATH not in sys.path:
        sys.path.insert(0, DRILL_MODULE_PATH)
        
    import move_main
    import 放平3d体 # Flatten 3D Body module
    
except ImportError as e:
    print(f"❌ 依赖模块导入失败: {e}")
    print(f"   请检查路径: {sys.path[:5]}")
    sys.exit(1)


def process_single_file(file_path: str, pm: PathManager, index: int):
    """
    处理单个文件的核心逻辑
    """
    filename = os.path.basename(file_path)
    result = {
        "success": False,
        "message": "",
        "file": filename
    }
    
    nx = None
    try:
        # 准备参数
        drill_json_path = str(pm.get_drill_table_json())
        knife_json_path = str(pm.get_knife_table_json())
        texture_dll_path = str(pm.get_texture_dll_path())
        output_dir = str(pm.get_textured_prt_dir()) # Step 5 output
        
        # 1. 初始化 NX
        nx = NXProcessor()
        
        # 2. 打开部件
        if not nx.open_part(file_path):
            result["message"] = "无法打开部件"
            return result
            
        session = nx.get_session()
        work_part = nx.get_current_part()

        # 3. 移动原点 (调用 move_main)
        # 注意：move_main 可能需要自己的 config 上下文，这里直接调用
        try:
             print(f"   [Debug] Drill JSON: {drill_json_path}")
             print(f"   [Debug] Knife JSON: {knife_json_path}")
             move_main.move_to_origin(session, work_part, drill_json_path, knife_json_path)
        except Exception as e:
             # 有时 move_main 可能因为图层或其他原因失败但不致命
             print(f"   ⚠️ 移动原点失败: {e}")
             traceback.print_exc()
             # result["message"] = f"移动原点失败: {e}"
             # return result
        

        # 3.5 放平 3D 体 (新增)
        try:
             print("   > 执行放平逻辑...")
             放平3d体.execute_alignment(work_part)
        except Exception as e:
             print(f"   ⚠️ 放平逻辑警告: {e}")

        # 4. 贴图
        try:
            tm = TextureMapper(texture_dll_path)
            tm.apply_texture()
        except Exception as e:
            print(f"   ⚠️ 贴图警告: {e}")

        # 5. 保存
        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
            
        output_file_path = os.path.join(output_dir, filename)
        
        # 使用 SaveAs 保存到新位置
        try:
            work_part.SaveAs(output_file_path)
            result["success"] = True
            result["message"] = "成功"
        except Exception as e:
            result["message"] = f"保存失败: {e}"

        nx.close_all()
        return result

    except Exception as e:
        result["message"] = f"处理异常: {e}"
        # traceback.print_exc()
        if nx:
            try: nx.close_all() 
            except: pass
        return result
    finally:
        gc.collect()


def run_step5_logic():
    print("=" * 60)
    print("🚀 步骤 5: 移动原点与贴图 (单进程版)")
    print("=" * 60)

    # 初始化管理器
    pm = PathManager(config.FILE_INPUT_PRT, config.FILE_INPUT_DXF)
    
    # 1. 准备目录
    source_prt_dir = str(pm.get_merged_prt_dir()) # Step 4 output
    output_dir = str(pm.get_textured_prt_dir())    # Step 5 output
    
    if not os.path.exists(source_prt_dir):
        print(f"❌ 源目录不存在: {source_prt_dir}")
        print("请先运行步骤 4")
        return

    # 清理输出目录
    if os.path.exists(output_dir):
        try: shutil.rmtree(output_dir)
        except: pass
    os.makedirs(output_dir, exist_ok=True)

    print(f"源目录: {source_prt_dir}")
    print(f"输出目录: {output_dir}")
    print("-" * 50)

    # 2. 获取文件列表
    prt_files = [os.path.join(source_prt_dir, f) for f in os.listdir(source_prt_dir) if f.lower().endswith('.prt')]
    
    if not prt_files:
        print(f"❌ 未找到 PRT 文件")
        return

    # 3. 循环处理
    results = []
    completed = 0
    total = len(prt_files)
    
    start_time = time.perf_counter()

    for idx, f_path in enumerate(prt_files):
        try:
            res = process_single_file(f_path, pm, idx + 1)
            results.append(res)
            
            completed += 1
            status_icon = "✅" if res["success"] else "❌"
            print(f"[{completed}/{total}] {status_icon} {res['file']}")
            if not res["success"]:
                print(f"    原因: {res['message']}")
                
        except KeyboardInterrupt:
            print("\n⚠️ 用户中断")
            break
        except Exception as e:
            print(f"❌ 循环错误: {e}")
            
        sys.stdout.flush()

    # 4. 统计
    print("-" * 50)
    success_count = sum(1 for r in results if r["success"])
    print(f"📊 处理完成 | 成功: {success_count} | 失败: {len(results) - success_count}")
    print(f"⏱️ 总耗时: {(time.perf_counter() - start_time):.2f} 秒")
    
    gc.collect()

if __name__ == "__main__":
    run_step5_logic()
