# -*- coding: utf-8 -*-
"""
步骤 14: 生成工程单数据 (run_step14.py)
功能：
1. 遍历含有CAM工序的PRT文件 (通常是 06_CAM/Final_CAM_PRT)
2. 调用 'final_jietu.py' (子进程) 提取工序参数并截图
3. 生成 JSON、TXT、尺寸信息文件到 06_CAM/Engineering_Order_Data
"""

import os
import sys
import glob
import shutil
import subprocess
import traceback

import config
from path_manager import PathManager

def run_step14_logic(pm: PathManager):
    print("=" * 80)
    print("🚀 步骤 14: 生成工程单数据 (JSON/TXT/截图)")
    print("=" * 80)

    # 1. 确定路径
    # 输入: 最终CAM PRT文件
    input_dir = pm.get_final_cam_prt_dir()
    # 输出根目录
    output_root = pm.get_engineering_order_root()
    # 脚本路径 (final_jietu.py)
    current_dir = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(current_dir, "final_jietu.py")
    # DLL路径 (根据实际存在的文件)
    dll_path = os.path.join(config.DLL_DIR, "FlipAndShot", "jietujietu.dll")

    # 清理输出目录
    if os.path.exists(output_root):
        try:
            shutil.rmtree(output_root)
            print(f"🗑️ 已清理旧输出目录: {output_root}")
        except Exception as e:
            print(f"⚠️ 清理旧目录失败: {e}")

    print(f"📂 输入PRT目录: {input_dir}")
    print(f"📂 数据输出目录: {output_root}")
    print(f"📜 调用脚本: {script_path}")
    print(f"🔌 DLL路径: {dll_path}")

    if not os.path.exists(input_dir):
        print(f"❌ 输入目录不存在: {input_dir}")
        return

    if not os.path.exists(script_path):
        print(f"❌ 脚本不存在: {script_path}")
        return
        
    if not os.path.exists(dll_path):
        print(f"❌ DLL不存在: {dll_path}")
        return

    prt_files = glob.glob(os.path.join(input_dir, "*.prt"))
    if not prt_files:
        print("⚠️ 未找到PRT文件")
        return

    success_count = 0
    fail_count = 0

    # 2. 遍历处理
    for i, prt_path in enumerate(prt_files):
        file_name = os.path.basename(prt_path)
        print(f"\nProcessing [{i+1}/{len(prt_files)}]: {file_name}")

        # 构造命令行参数
        # python final_jietu.py <prt_path> <output_base_dir> <dll_path>
        cmd = [
            sys.executable,
            script_path,
            prt_path,
            str(output_root),
            str(dll_path)
        ]
        
        try:
            # 调用子进程
            result = subprocess.run(cmd, capture_output=False, text=True, check=False)
            
            if result.returncode == 0:
                print(f"✅ [{file_name}] 处理成功")
                success_count += 1
            else:
                print(f"❌ [{file_name}] 处理失败 (Exit Code {result.returncode})")
                fail_count += 1
                
        except Exception as e:
            print(f"❌ 调用异常: {e}")
            traceback.print_exc()
            fail_count += 1

    print("\n" + "=" * 80)
    print(f"Step 14 完成: 成功 {success_count}, 失败 {fail_count}")
    print("=" * 80)

def main():
    pm = PathManager(config.FILE_INPUT_PRT, config.FILE_INPUT_DXF)
    run_step14_logic(pm)

if __name__ == "__main__":
    main()
