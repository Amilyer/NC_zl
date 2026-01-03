# -*- coding: utf-8 -*-
"""
自动执行脚本 (run_all_steps.py)
功能：
1. 自动扫描当前目录下的 run_step*.py 脚本
2. 按步骤号顺序执行 (1, 2, 3...)
3. 上一个脚本执行成功后，自动执行下一个
4. 遇到错误立即停止
"""

import sys
import os
import re
import time
import subprocess
from pathlib import Path

def get_step_number(filename):
    """从文件名中提取步骤号，用于排序"""
    match = re.search(r'run_step(\d+)', filename, re.IGNORECASE)
    if match:
        return int(match.group(1))
    return float('inf')

def main():
    # 获取当前脚本所在目录
    current_dir = Path(__file__).parent
    
    # 1. 扫描所有 run_step*.py 文件
    print(f"📂 正在扫描脚本目录: {current_dir}")
    scripts = list(current_dir.glob("run_step*.py"))
    
    # 过滤掉非步骤脚本（如果需要）
    scripts = [s for s in scripts if get_step_number(s.name) != float('inf')]
    
    # 2. 按步骤号排序
    scripts.sort(key=lambda p: get_step_number(p.name))
    
    if not scripts:
        print("❌ 未找到任何 run_step*.py 脚本")
        return

    print(f"📋 找到 {len(scripts)} 个待执行脚本:")
    for s in scripts:
        print(f"   - {s.name}")
    print("=" * 60)
    
    total_start = time.perf_counter()

    # 3. 顺序执行
    for i, script_path in enumerate(scripts):
        script_name = script_path.name
        step_num = get_step_number(script_name)
        
        print(f"\n🚀 [{i+1}/{len(scripts)}] 正在执行: {script_name} ...")
        print("-" * 60)
        
        step_start = time.perf_counter()
        
        try:
            # 使用当前 Python 解释器启动子进程
            # cwd 设置为脚本所在目录，确保相对路径正确
            # check=True 会在返回码非 0 时抛出 CalledProcessError
            result = subprocess.run(
                [sys.executable, str(script_path)],
                cwd=str(current_dir),
                check=True
            )
            
            duration = time.perf_counter() - step_start
            print("-" * 60)
            print(f"✅ {script_name} 执行成功 (耗时: {duration:.2f}s)")
            
        except subprocess.CalledProcessError as e:
            print("-" * 60)
            print(f"❌ {script_name} 执行失败 (返回码: {e.returncode})")
            print("⛔ 自动化流程已终止")
            sys.exit(e.returncode)
            
        except Exception as e:
            print("-" * 60)
            print(f"❌ {script_name} 发生未知错误: {e}")
            print("⛔ 自动化流程已终止")
            sys.exit(1)
            
        # 可选：步骤间稍微暂停，便于观察或释放资源
        time.sleep(1)

    total_duration = time.perf_counter() - total_start
    print("\n" + "=" * 60)
    print(f"🎉 所有步骤执行完成！")
    print(f"⏱️ 总耗时: {total_duration:.2f}s")
    print("=" * 60)

if __name__ == "__main__":
    main()
