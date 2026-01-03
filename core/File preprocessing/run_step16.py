# -*- coding: utf-8 -*-
"""
Step 16: NC程序导出 (多进程并发版)
功能：遍历 CAM PRT 文件，生成 NC 代码
"""
import os
import sys
import importlib.util
import traceback
from pathlib import Path
import time
from concurrent.futures import ProcessPoolExecutor, as_completed

# 添加项目根目录到 sys.path
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(os.path.dirname(current_dir))
if project_root not in sys.path:
    sys.path.append(project_root)
if current_dir not in sys.path:
    sys.path.append(current_dir)

from path_manager import PathManager, get_path_manager

# 并发配置
MAX_WORKERS = 20  # 并发进程数

def load_nc_module():
    """Dynamically load the nc_processor module"""
    module_path = os.path.join(current_dir, "nc_processor.py")
    spec = importlib.util.spec_from_file_location("nc_module", module_path)
    nc_module = importlib.util.module_from_spec(spec)
    sys.modules["nc_module"] = nc_module
    spec.loader.exec_module(nc_module)
    return nc_module

def process_single_file(args):
    """
    处理单个 PRT 文件的 NC 代码生成（子进程执行）
    """
    prt_file, output_root = args
    
    import sys
    import os
    import time
    import traceback
    from pathlib import Path
    
    # [DEBUG] 多进程调试信息
    pid = os.getpid()
    start_time = time.time()
    part_name = Path(prt_file).stem
    print(f"[DEBUG] PID={pid} | 开始处理: {part_name}")
    
    # 确保路径正确
    current_dir = os.path.dirname(os.path.abspath(__file__))
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)
    
    try:
        import NXOpen
        session = NXOpen.Session.GetSession()
        nc_module = load_nc_module()
    except Exception as e:
        return {"success": False, "file": str(prt_file), "error": f"初始化失败: {e}"}
    
    part_out_dir = Path(output_root) / part_name
    
    try:
        # 创建输出目录
        part_out_dir.mkdir(parents=True, exist_ok=True)
        
        # 打开文件
        base_part, _ = session.Parts.OpenBaseDisplay(str(prt_file))
        
        # 生成 NC 代码
        nc_module.main(str(part_out_dir))
        
        # 关闭文件
        base_part.Close(NXOpen.BasePart.CloseWholeTree.TrueValue, NXOpen.BasePart.CloseModified.CloseModified, None)
        
        # [DEBUG] 完成信息
        elapsed = time.time() - start_time
        print(f"[DEBUG] PID={pid} | 完成: {part_name} | 耗时: {elapsed:.1f}s")
        return {"success": True, "file": str(prt_file), "output_dir": str(part_out_dir), "pid": pid, "elapsed": elapsed}
    except Exception as e:
        elapsed = time.time() - start_time
        print(f"[DEBUG] PID={pid} | 失败: {part_name} | 耗时: {elapsed:.1f}s | 错误: {e}")
        return {"success": False, "file": str(prt_file), "error": str(e), "traceback": traceback.format_exc(), "pid": pid}

def run_step16_logic(pm: PathManager):
    print("=" * 80)
    print(f"  Step 16: NC程序导出 (多进程并发版, workers={MAX_WORKERS})")
    print("=" * 80)

    # 1. Setup paths
    input_dir = pm.get_final_cam_prt_dir()
    output_root = pm.get_nc_output_dir()
    
    print(f"📂 输入目录: {input_dir}")
    print(f"📂 输出目录: {output_root}")
    
    # 清理输出目录（主进程执行一次）
    import shutil
    if output_root.exists():
        try:
            shutil.rmtree(output_root)
            print(f"🗑️ 已清理输出目录: {output_root}")
        except Exception as e:
            print(f"⚠️ 清理目录失败: {e}")
    output_root.mkdir(parents=True, exist_ok=True)
    
    if not input_dir.exists():
        print(f"⚠️ 输入目录不存在: {input_dir}")
        return
        
    prt_files = list(input_dir.glob("*.prt"))
    if not prt_files:
        print("⚠️ 未找到PRT文件")
        return

    print(f"[INFO] 找到 {len(prt_files)} 个PRT文件")

    # 2. 准备并发任务参数
    task_args = [
        (str(prt_file), str(output_root))
        for prt_file in prt_files
    ]
    
    # 3. 并发处理文件
    success_count = 0
    fail_count = 0
    
    print(f"\n🚀 开始并发处理 {len(prt_files)} 个文件...")
    
    with ProcessPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_file = {executor.submit(process_single_file, args): args[0] for args in task_args}
        
        for i, future in enumerate(as_completed(future_to_file)):
            prt_path = future_to_file[future]
            part_name = Path(prt_path).stem
            
            try:
                result = future.result()
                if result["success"]:
                    success_count += 1
                    print(f"[{i+1}/{len(prt_files)}] ✅ {part_name}")
                else:
                    fail_count += 1
                    print(f"[{i+1}/{len(prt_files)}] ❌ {part_name}: {result.get('error', '未知错误')}")
            except Exception as e:
                fail_count += 1
                print(f"[{i+1}/{len(prt_files)}] ❌ {part_name}: {e}")

    print("\n" + "=" * 80)
    print(f"Step 16 完成: 成功 {success_count}, 失败 {fail_count}")
    print("=" * 80)

def main():
    """统一入口 - 根据配置选择执行模式"""
    import config
    
    from path_manager import init_path_manager
    pm = init_path_manager()
    run_step16_logic(pm)

if __name__ == "__main__":
    # 独立运行时检查 NX 环境
    try:
        import NXOpen
        s = NXOpen.Session.GetSession()
    except:
        print("⚠️ 需要在NX环境或通过run_journal运行")
        sys.exit(1)
    main()