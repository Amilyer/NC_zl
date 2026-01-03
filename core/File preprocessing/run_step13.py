"""
Step 13: 自动刀轨生成 (精铣 - 多进程并发版)
功能：遍历 PRT 文件，查找 JSON 配置文件，调用 '创建刀轨.py' 生成精加工程序
"""

import importlib.util
import os
import sys
import config
from path_manager import init_path_manager
from pathlib import Path
import shutil
import re
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
    """动态导入指定路径的模块"""
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

def import_finishing_toolpath_module():
    """动态导入创建刀轨模块"""
    module_name = "create_toolpath"
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, "创建刀轨.py")
    return import_module_from_path(module_name, file_path)

# ----------------------------------------------------------------------------------------------------------------------
# 并发配置
MAX_WORKERS = getattr(config, 'PROCESS_MAX_WORKERS', 8)

def find_json_files(json_dir, part_name):
    """扫描目录下的json，进行配对"""
    # 解析除去时间戳的基础名称 (例: DIE-05_2024... -> DIE-05)
    base_name = re.sub(r'_\d{8}_\d{6}$', '', part_name)
    
    def get_path(suffix):
        # 1. 尝试完全匹配 (PartName_Suffix.json)
        p1 = os.path.join(json_dir, f"{part_name}_{suffix}.json")
        if os.path.exists(p1): return p1
        
        # 2. 尝试基础名称匹配 (BaseName_Suffix.json)
        p2 = os.path.join(json_dir, f"{base_name}_{suffix}.json")
        if os.path.exists(p2): return p2
        return None

    return {
        "half_spiral_json_path": get_path("半精_螺旋"),
        "half_spiral_reciprocating_json_path": get_path("半精_螺旋_往复等高"),
        "half_surface_json_path": get_path("半精_爬面"),
        "half_jiao_json_path": get_path("半精_清角"),
        "half_mian_json_path": get_path("半精_面铣"),
        "mian_json_path": get_path("全精_面铣"),
        "spiral_json_path": get_path("全精_螺旋"),
        "spiral_reciprocating_json_path": get_path("全精_螺旋_往复等高"),
        "reciprocating_json_path": get_path("全精_往复等高"),
        "surface_json_path": get_path("全精_爬面"),
        "gen_json_path": get_path("全精_清根")
    }

def process_single_file(args):
    """
    处理单个 PRT 文件的精加工刀轨生成（子进程执行）
    """
    prt_file, json_dir, output_dir = args
    
    import sys
    import os
    import re
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
        tp_module = import_finishing_toolpath_module()
        if not tp_module:
            return {"success": False, "file": str(prt_file), "error": "导入模块失败"}




        tp_module.CONFIG["TEST_MODE"] = True


        
    except Exception as e:
        return {"success": False, "file": str(prt_file), "error": f"导入模块异常: {e}"}
    
    base_name = re.sub(r'_\d{8}_\d{6}$', '', part_name)
    
    try:
        # 查找对应的 JSON 文件
        json_config = find_json_files(str(json_dir), base_name)
        
        # 准备参数
        workflow_args = json_config.copy()
        workflow_args["part_path"] = str(prt_file)
        workflow_args["save_dir"] = str(output_dir)
        
        # 调用工作流函数
        saved_path = tp_module.generate_toolpath_workflow(**workflow_args)
        
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
    print(f"  Step 13: 自动刀轨生成 (精铣 - 多进程并发版, workers={MAX_WORKERS})")
    print("=" * 80)

    # 1. 初始化PathManager
    pm = init_path_manager(config.FILE_INPUT_PRT, config.FILE_INPUT_DXF)
    
    # 2. 输入目录 (Step 12 输出的开粗后PRT)
    input_dir = pm.get_cam_roughing_prt_dir()
    if not input_dir.exists():
        print(f"[ERROR] 输入目录不存在: {input_dir}")
        print("请先运行 Step 12")
        return

    # 获取所有PRT文件
    prt_files = list(input_dir.glob("*.prt"))
    if not prt_files:
        print(f"[ERROR] 在 {input_dir} 中找不到PRT文件")
        return
        
    print(f"[INFO] 输入目录: {input_dir}")
    print(f"[INFO] 找到 {len(prt_files)} 个PRT文件")

    # 3. 确保输出目录存在
    output_dir = pm.get_final_cam_prt_dir()
    if output_dir.exists():
        print(f"[INFO] 清理输出目录: {output_dir}")
        clean_dir(str(output_dir))
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"[INFO] 输出目录: {output_dir}")
    
    # 4. 准备并发任务参数
    json_dir = pm.get_cam_json_dir()
    
    task_args = [
        (str(prt_file), str(json_dir), str(output_dir))
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
    print(f"Step 13 完成: 成功 {success_count}, 失败 {fail_count}")
    print("=" * 80)

if __name__ == "__main__":
    main()