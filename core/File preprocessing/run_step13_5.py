# -*- coding: utf-8 -*-
"""
Step 13.5: 过切检查 (Gouge Check)
功能：
1. 读取 Step 13 生成的最终 PRT 文件
2. 调用 guoqiejiancha.py 执行过切检查
3. 生成 Excel/TXT/JSON 报告
"""

import os
import sys
import glob
import traceback
import importlib.util

# Ensure we can import modules from project root logic if needed
# Add current directory to sys.path to ensure we can import path_manager
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

try:
    from path_manager import init_path_manager
except ImportError:
    # If standard import fails, try to look up parent (though implementation_plan says it's in core/File preprocessing)
    # This is a fallback
    print("⚠ 无法直接导入 path_manager，尝试调整 sys.path")
    sys.path.append(os.path.dirname(os.path.dirname(current_dir))) 
    from core.File_preprocessing.path_manager import init_path_manager

def load_module_from_file(module_name, file_path):
    """动态加载指定路径的 Python 模块"""
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None:
        raise ImportError(f"无法从 {file_path} 创建模块规范")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

def run_step13_5():
    print("🚀 开始执行 Step 13.5: 过切检查 (Gouge Check) ...")
    
    # 1. 初始化路径管理器
    # 这会读取 config.py 配置并确保基本目录存在
    try:
        pm = init_path_manager()
    except Exception as e:
        print(f"❌ 初始化 PathManager 失败: {e}")
        return

    # 2. 确定输入目录 (Step 13 输出的 PRT)
    # 使用 path_manager 中定义的 Final_CAM_PRT (06_CAM/Final_CAM_PRT)
    try:
        input_prt_dir = pm.get_final_cam_prt_dir()
    except AttributeError:
        # Fallback in case method name is slightly different in loaded version
        # Based on file read, method is get_final_cam_prt_dir
        input_prt_dir = os.path.join(pm.dir_cam, 'Final_CAM_PRT')

    if not os.path.exists(input_prt_dir):
        print(f"❌ 输入目录不存在: {input_prt_dir}")
        print("请检查 Step 13 是否已成功执行并生成文件。")
        return

    # 3. 确定输出目录 (08_Gouge_Check_Reports)
    # 在 06_CAM 下创建一个新目录用于存放过切检查报告
    output_root = os.path.join(pm.dir_cam, "08_Gouge_Check_Reports")
    
    if not os.path.exists(output_root):
        os.makedirs(output_root)
        print(f"已创建输出目录: {output_root}")

    # 4. 加载 guoqiejiancha.py 模块
    # 该脚本应位于当前脚本同级目录
    guoqie_script_path = os.path.join(current_dir, "guoqiejiancha.py")
    if not os.path.exists(guoqie_script_path):
        print(f"❌ 找不到过切检查脚本文件: {guoqie_script_path}")
        return
        
    try:
        guoqie_module = load_module_from_file("guoqiejiancha", guoqie_script_path)
    except Exception as e:
        print(f"❌ 加载过切检查模块失败: {e}")
        traceback.print_exc()
        return

    # 5. 遍历 PRT 文件并处理
    prt_search_pattern = os.path.join(input_prt_dir, "*.prt")
    prt_files = glob.glob(prt_search_pattern)
    
    if not prt_files:
        print(f"⚠ 在 {input_prt_dir} 未找到任何 PRT 文件")
        return

    print(f"找到 {len(prt_files)} 个 PRT 文件，准备开始检查...")

    success_count = 0
    fail_count = 0
    skipped_count = 0

    for i, prt_path in enumerate(prt_files, 1):
        prt_name = os.path.basename(prt_path)
        print(f"\n[{i}/{len(prt_files)}] 正在处理: {prt_name}")
        
        try:
            # 调用 guoqiejiancha.main(part_path, root_dir)
            # guoqiejiancha.main 会自动在 root_dir 下创建 excel, txt, json, prt 子目录
            result = guoqie_module.main(prt_path, output_root)
            
            if result:
                print(f"✅ {prt_name} 检查完成")
                success_count += 1
            else:
                print(f"❌ {prt_name} 检查失败 (返回 False)")
                fail_count += 1
        except Exception as e:
            print(f"❌ {prt_name} 处理异常: {e}")
            traceback.print_exc()
            fail_count += 1

    print("\n" + "="*50)
    print(f"Step 13.5 执行完毕")
    print(f"共扫描: {len(prt_files)}")
    print(f"成功: {success_count}, 失败: {fail_count}")
    print(f"报告根目录: {output_root}")
    print("="*50)

if __name__ == "__main__":
    try:
        run_step13_5()
    except Exception as e:
        print(f"❌ 程序运行出错: {e}")
        traceback.print_exc()
