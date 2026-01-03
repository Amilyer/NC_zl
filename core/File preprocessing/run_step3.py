# -*- coding: utf-8 -*-
"""
步骤 3: 数据配对与格式转换 (run_step3.py)
功能
"""

print("hello,the pretty cat named Qin~")

import shutil
import time
from pathlib import Path

import config
from path_manager import init_path_manager, PathManager

# 导入业务逻辑
from dxf_info_extractor import extract_dxf_info
from data_matcher import match_data
from dxf_to_prt import batch_convert_dxf_to_prt

def run_step3_1(pm: PathManager) -> str:
    """提取 DXF 信息"""
    print("\n📐 [Step 3.1] 提取DXF尺寸信息")
    
    input_dir = pm.get_split_dxf_dir()
    output_csv = pm.get_2d_report_csv()
    
    # 检查输入 (pathlib iterdir)
    if not input_dir.exists() or not any(input_dir.iterdir()):
        print(f"   ❌ 输入目录为空或不存在: {input_dir}")
        return None

    # 运行提取
    result = extract_dxf_info(str(input_dir), str(output_csv))
    
    if result:
        print(f"   ✅ 提取完成: {Path(result).name}")
    else:
        print("   ❌ 提取失败")
    return result

def run_step3_2(pm: PathManager) -> str:
    """数据配对"""
    print("\n🔗 [Step 3.2] 数据配对")
    
    dxf_csv = pm.get_2d_report_csv()
    prt_csv = pm.get_3d_report_csv()
    output_csv = pm.get_match_result_csv()

    if not dxf_csv.exists():
        print(f"   ❌ 缺少 2D CSV: {dxf_csv}")
        return None
    if not prt_csv.exists():
        print(f"   ❌ 缺少 3D CSV: {prt_csv}")
        return None

    result = match_data(str(dxf_csv), str(prt_csv), str(output_csv))
    
    if result:
        print(f"   ✅ 配对完成: {Path(result).name}")
    else:
        print("   ❌ 配对失败")
    return result

def run_step3_3(pm: PathManager):
    """DXF 转 PRT"""
    print("\n🔄 [Step 3.3] DXF 转 PRT")
    
    input_dir = pm.get_split_dxf_dir()
    output_dir = pm.get_dxf_prt_dir()
    
    # 清理输出
    if output_dir.exists():
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"   DIR: {output_dir}")
    batch_convert_dxf_to_prt(str(input_dir), str(output_dir))
    print(f"   ✅ 转换流程调用结束 (具体结果见上方日志)")

def run_processing_loop(pm):
    print("=" * 60)
    print("🚀 步骤 3: 数据配对与格式转换")
    print("=" * 60)
    start_time = time.perf_counter()
    
    # 1. 检查前置 Step 1 & 2
    if not pm.get_split_dxf_dir().exists():
        print("❌ 错误: 找不到 Step 2 输出的 DXF 目录，请先运行 Step 2")
        return
    if not pm.get_3d_report_csv().exists():
        print("❌ 错误: 找不到 Step 1 输出的 CSV 报告，请先运行 Step 1")
        return

    # 2. 依次运行子步骤
    try:
        # Step 3.1
        if not run_step3_1(pm):
            print("💥 Step 3.1 失败，流程终止")
            return
            
        # Step 3.2
        if not run_step3_2(pm):
            print("💥 Step 3.2 失败，流程终止")
            return
        
        # Step 3.3
        run_step3_3(pm)
        
        print("\n" + "-" * 50)
        print(f"🎉 步骤 3 全部完成 | 耗时: {(time.perf_counter() - start_time):.2f}s")

    except Exception as e:
        print(f"\n❌ 未知异常: {e}")
        import traceback
        traceback.print_exc()

def main():
    pm = init_path_manager(config.FILE_INPUT_PRT_STR, config.FILE_INPUT_DXF_STR)
    run_processing_loop(pm)

if __name__ == "__main__":
    main()
