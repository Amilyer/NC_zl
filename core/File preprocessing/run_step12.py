"""
Step 12: 自动生成开粗刀轨 (多进程并发版)
功能：遍历 PRT 文件，查找 Step 10 生成的开粗 JSON 配置文件，调用 '创建开粗刀轨.py' 生成加工程序
"""

import importlib.util
import os
import sys
import config
from path_manager import init_path_manager
from pathlib import Path
import shutil
from concurrent.futures import ProcessPoolExecutor, as_completed

def clean_dir(dir_path):
    for item in os.listdir(dir_path):
        item_path = os.path.join(dir_path, item)
        try:
            if os.path.isfile(item_path):
                os.unlink(item_path)
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
        except Exception as e:
            print(f"Error cleaning {item_path}: {e}")

# ----------------------------------------------------------------------------------------------------------------------
def import_module_from_path(module_name, file_path):
    """动态导入指定路径的模块 (支持中文路径/文件名)"""
    try:
        if not os.path.exists(file_path):
             return None
        spec = importlib.util.spec_from_file_location(module_name, file_path)
        if spec and spec.loader:
            module = importlib.util.module_from_spec(spec)
            sys.modules[module_name] = module
            spec.loader.exec_module(module)
            return module
    except Exception as e:
        print(f"[ERROR] 无法导入模块 {module_name}: {e}")
    return None

def import_roughing_toolpath_module():
    """动态导入创建开粗刀轨模块"""
    module_name = "create_roughing_toolpath"
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, "创建开粗刀轨.py")
    return import_module_from_path(module_name, file_path)

# ----------------------------------------------------------------------------------------------------------------------
# 并发配置
MAX_WORKERS = getattr(config, 'PROCESS_MAX_WORKERS', 8)

def process_single_file(args):
    """
    处理单个 PRT 文件的开粗刀轨生成（子进程执行）
    """
    prt_file, json_dir, output_dir, project_root = args
    
    import sys
    import os
    import time
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
    
    # 动态导入模块（子进程内）
    try:
        tp_module = import_roughing_toolpath_module()
        if not tp_module:
            return {"success": False, "file": str(prt_file), "error": "导入模块失败"}


        # [DEBUG] 设置测试模式
        tp_module.CONFIG["TEST_MODE"] = True




        
    except Exception as e:
        return {"success": False, "file": str(prt_file), "error": f"导入模块异常: {e}"}
    
    part_name = Path(prt_file).stem
    json_path = Path(json_dir)
    
    # 查找 JSON 文件
    cavity_json = json_path / f"{part_name}_行腔.json"
    reciprocating_json = json_path / f"{part_name}_开粗_往复等高.json"
    if not reciprocating_json.exists():
        reciprocating_json = json_path / f"{part_name}_往复等高.json"
    
    cavity_path = str(cavity_json) if cavity_json.exists() else None
    reciprocating_path = str(reciprocating_json) if reciprocating_json.exists() else None
    
    try:
        saved_path = tp_module.generate_toolpath_workflow(
            part_path=str(prt_file),
            cavity_json_path=cavity_path,
            reciprocating_json_path=reciprocating_path,
            save_dir=str(output_dir)
        )
        # [DEBUG] 完成信息
        elapsed = time.time() - start_time
        print(f"[DEBUG] PID={pid} | 完成: {part_name} | 耗时: {elapsed:.1f}s")
        return {"success": True, "file": str(prt_file), "saved_path": saved_path, "pid": pid, "elapsed": elapsed}
    except Exception as e:
        import traceback
        elapsed = time.time() - start_time
        print(f"[DEBUG] PID={pid} | 失败: {part_name} | 耗时: {elapsed:.1f}s | 错误: {e}")
        return {"success": False, "file": str(prt_file), "error": str(e), "traceback": traceback.format_exc(), "pid": pid}


def main():
    print("=" * 80)
    print(f"  Step 12: 自动生成开粗刀轨 (多进程并发版, workers={MAX_WORKERS})")
    print("=" * 80)

    # 1. 初始化PathManager
    pm = init_path_manager(config.FILE_INPUT_PRT, config.FILE_INPUT_DXF)
    
    # 2. 从 Step 9 输出目录读取 PRT 文件 (作为开粗起始)
    prt_dir = pm.get_step9_drilled_dir()
    if not prt_dir.exists():
        print(f"[ERROR] 输入目录不存在: {prt_dir}")
        print("请先运行 Step 11")
        return
        
    # 获取所有PRT文件
    prt_files = list(prt_dir.glob("*.prt"))
    if not prt_files:
        print(f"[ERROR] 在 {prt_dir} 中找不到PRT文件")
        return
        
    print(f"[INFO] 输入目录: {prt_dir}")
    print(f"[INFO] 找到 {len(prt_files)} 个PRT文件")
    
    # 3. 确保输出目录存在
    output_dir = pm.get_cam_roughing_prt_dir()
    if output_dir.exists():
        print(f"[INFO] 清理输出目录: {output_dir}")
        clean_dir(str(output_dir))
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"[INFO] 输出目录: {output_dir}")

    # 4. 准备并发任务参数
    json_dir = pm.get_cam_roughing_json_dir()
    project_root = str(Path(__file__).parent.parent.parent)
    
    task_args = [
        (str(prt_file), str(json_dir), str(output_dir), project_root)
        for prt_file in prt_files
    ]
    
    # 5. 并发处理文件
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
    print(f"Step 12 完成: 成功 {success_count}, 失败 {fail_count}")
    print("=" * 80)

if __name__ == "__main__":
    main()
