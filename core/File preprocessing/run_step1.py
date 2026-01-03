# -*- coding: utf-8 -*-
"""
步骤 1: 3D PRT 拆分 (run_step1.py)
功能：
1. 读取输入的 3D 组立 PRT 文件
2. 拆分为独立的零件文件
3. 生成 CSV 拆分报告
"""
print("hello,the pretty cat named Qin~")

import time
import config
from path_manager import init_path_manager
from prt_split import split_prt_file_with_output

def run_processing_loop(pm, input_prt_path):
    print("=" * 60)
    print("🚀 步骤 1: 3D PRT 拆分")
    print("=" * 60)
    
    start_time = time.perf_counter()

    # 1. 准备目录
    output_dir = pm.get_split_prt_dir()
    
    print(f"📦 输入文件: {input_prt_path.name}")
    print(f"📂 输出目录: {output_dir}")
    print("-" * 50)

    # 2. 执行核心逻辑
    try:
        csv_path, out_dir_str = split_prt_file_with_output(
            str(input_prt_path),
            str(output_dir)
        )
        
        if csv_path and out_dir_str:
            print(f"✅ 拆分成功")
            print(f"   CSV报告: {csv_path}")
            
            # 统计数量 (使用 pathlib)
            # 兼容大小写扩展名 (Windows文件名通常不敏感，但glob可能敏感，这里简单匹配 .prt)
            # 若需严格不区分大小写，可遍历检查 suffix
            count = len([p for p in output_dir.glob('*') if p.suffix.lower() == '.prt'])
            print(f"   生成数量: {count} 个文件")
            
        else:
            print("❌ 拆分失败: 返回路径为空")

    except Exception as e:
        print(f"❌ 发生异常: {e}")
        import traceback
        traceback.print_exc()

    print("-" * 50)
    print(f"⏱️ 总耗时: {(time.perf_counter() - start_time):.2f} 秒")


def main():
    # 1. 初始化路径
    pm = init_path_manager(config.FILE_INPUT_PRT_STR, config.FILE_INPUT_DXF_STR)
    
    # 2. 检查输入
    input_prt = pm.get_input_3d_prt()
    if not input_prt.exists():
        print(f"❌ 找不到输入文件: {input_prt}")
        return

    # 3. 运行
    run_processing_loop(pm, input_prt)

if __name__ == "__main__":
    main()
