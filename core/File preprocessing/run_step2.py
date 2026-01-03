# -*- coding: utf-8 -*-
"""
步骤 2: 2D DXF 拆分 (run_step2.py)
功能：
1. 读取输入的 2D 组立 DXF 文件
2. 识别图框并拆分为独立子图 DXF
3. 清理非几何图层
"""

import shutil
import time
import config
from path_manager import init_path_manager
from dxf_split import split_dxf_file_with_output

def run_processing_loop(pm, input_dxf_path):
    print("=" * 60)
    print("🚀 步骤 2: 2D DXF 拆分")
    print("=" * 60)
    
    start_time = time.perf_counter()

    # 1. 准备目录
    dxf_split_dir = pm.get_split_dxf_dir()
    
    # 清理旧数据
    if dxf_split_dir.exists():
        print(f"🧹 清理旧数据: {dxf_split_dir}")
        shutil.rmtree(dxf_split_dir)
    dxf_split_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"📦 输入文件: {input_dxf_path.name}")
    print(f"📂 输出目录: {dxf_split_dir}")
    print("-" * 50)

    # 2. 执行核心逻辑
    try:
        exported_dir = split_dxf_file_with_output(
            str(input_dxf_path),
            str(dxf_split_dir)
        )
        
        if exported_dir:
            # 统计数量 (使用 pathlib)
            count = len([p for p in dxf_split_dir.glob('*') if p.suffix.lower() == '.dxf'])
            
            print(f"✅ 拆分成功")
            print(f"   生成数量: {count} 个文件")
            print(f"   输出路径: {exported_dir}")
        else:
            print("❌ 拆分失败: 返回路径为空")

    except Exception as e:
        print(f"❌ 发生异常: {e}")
        import traceback
        traceback.print_exc()

    print("-" * 50)
    print(f"⏱️ 总耗时: {(time.perf_counter() - start_time):.2f} 秒")


def main():
    # 1. 初始化
    pm = init_path_manager(config.FILE_INPUT_PRT_STR, config.FILE_INPUT_DXF_STR)
    
    # 2. 检查输入
    input_dxf = pm.get_input_2d_dxf()
    if not input_dxf.exists():
        print(f"❌ 找不到输入文件: {input_dxf}")
        return

    # 3. 运行
    run_processing_loop(pm, input_dxf)

if __name__ == "__main__":
    main()
