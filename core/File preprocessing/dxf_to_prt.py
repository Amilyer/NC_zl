# -*- coding: utf-8 -*-
"""
批量DXF转PRT（NX外部Python运行，强制多段线转直线/圆弧）
核心：保留ImportPolylineTo=ArcLines，支持多进程并发
"""
import NXOpen
import NXOpen.Annotations
import NXOpen.Preferences
import os
import glob
import tempfile
import sys
from concurrent.futures import ProcessPoolExecutor, as_completed
from typing import Dict, Any

import config
# 尝试从配置获取进程数，默认8
MAX_WORKERS = getattr(config, 'PROCESS_MAX_WORKERS', 8)


def _import_single(dxf_file: str, prt_file: str) -> Dict[str, Any]:
    """单文件转换封装（用于多进程）"""
    result = {
        "success": False,
        "dxf_file": dxf_file,
        "prt_file": prt_file,
        "message": ""
    }

    if os.path.exists(prt_file):
        result["message"] = "PRT已存在"
        return result

    try:
        # 调用核心导入函数
        success = import_dxf_file(dxf_file, prt_file)
        result["success"] = success
        result["message"] = "转换成功" if success else "转换失败"
    except Exception as e:
        result["message"] = str(e)

    return result


def import_dxf_file(input_file, output_file):
    """
    导入单个DXF文件并保存为PRT文件（核心：ImportPolylineTo=ArcLines）
    返回: True=成功, False=失败
    """
    if not os.path.exists(input_file):
        return False

    # 确保输出目录存在
    output_dir = os.path.dirname(output_file)
    os.makedirs(output_dir, exist_ok=True)

    try:
        # 初始化NX会话
        theSession = NXOpen.Session.GetSession()
        workPart = theSession.Parts.Work
        displayPart = theSession.Parts.Display

        # 创建撤销标记
        markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "DXF导入")

        # 初始化DXF导入器
        dxfdwgImporter1 = theSession.DexManager.CreateDxfdwgImporter()
        dxfdwgImporter1.Units = NXOpen.DxfdwgImporter.UnitsEnum.Metric
        dxfdwgImporter1.ImportTo = NXOpen.DxfdwgImporter.ImportToEnum.New
        dxfdwgImporter1.ConvModelData = True
        dxfdwgImporter1.ConvLayoutData = True
        dxfdwgImporter1.ImportCurvesType = NXOpen.DxfdwgImporter.ImportCurvesAs.Curves

        # 核心：强制多段线导入为直线/圆弧
        dxfdwgImporter1.ImportPolylineTo = NXOpen.DxfdwgImporter.ImportPolylinesAs.ArcLines
        dxfdwgImporter1.ImportDimensionType = NXOpen.DxfdwgImporter.ImportDimensionsAs.Group
        # 输入输出文件
        dxfdwgImporter1.InputFile = input_file
        dxfdwgImporter1.OutputFile = output_file

        # 其他配置
        dxfdwgImporter1.HealBodies = True
        dxfdwgImporter1.FileOpenFlag = False

        # 临时映射文件
        temp_dir = tempfile.gettempdir()
        text_font_file = os.path.join(temp_dir, "text_font_mapping.txt")
        line_font_file = os.path.join(temp_dir, "line_font_mapping.txt")
        cross_hatch_file = os.path.join(temp_dir, "cross_hatch_mapping.txt")

        for file_path in [text_font_file, line_font_file, cross_hatch_file]:
            if not os.path.exists(file_path):
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write("# 临时字体/线型映射文件")

        dxfdwgImporter1.TextFontMappingFile = text_font_file
        dxfdwgImporter1.LineFontMappingFile = line_font_file
        dxfdwgImporter1.CrossHatchMappingFile = cross_hatch_file

        # 图层配置
        dxfdwgImporter1.AvoidUsedNXLayers = True
        dxfdwgImporter1.ReadLayerNumFromPrefix = False
        dxfdwgImporter1.TranslateUnselectedLayer = False
        dxfdwgImporter1.DestForUnselectedLayer = 256
        dxfdwgImporter1.ProcessingOrder = NXOpen.DxfdwgImporter.ProcessingOrderAs.Alphabetical
        dxfdwgImporter1.SkipEmptyLayer = True
        dxfdwgImporter1.UnSelectedLayers = None
        dxfdwgImporter1.AspectRatioOption = NXOpen.DxfdwgImporter.AspectRatioOptions.UseSameAsACADWidthFactor
        dxfdwgImporter1.ProcessHoldFlag = True

        # 提交导入
        theSession.SetUndoMarkName(markId1, "导入DXF文件")
        nXObject1 = dxfdwgImporter1.Commit()

        # 切换到建模应用
        theSession.ApplicationSwitchImmediate("UG_APP_MODELING")

        # 清理资源
        dxfdwgImporter1.Destroy()

        # 校验PRT文件
        return os.path.exists(output_file) and os.path.getsize(output_file) > 0

    except Exception as e:
        if 'dxfdwgImporter1' in locals():
            try:
                dxfdwgImporter1.Destroy()
            except:
                pass
        return False


def batch_convert_dxf_to_prt(input_dir: str, output_dir: str):
    """
    批量将DXF文件转换为PRT文件（多进程版本）
    """
    # 基础校验
    if not os.path.exists(input_dir):
        print(f"❌ 输入目录不存在: {input_dir}")
        return

    os.makedirs(output_dir, exist_ok=True)
    print(f"✅ 输出目录: {output_dir}")

    # 收集DXF文件
    dxf_files = glob.glob(os.path.join(input_dir, "*.dxf")) + glob.glob(os.path.join(input_dir, "*.DXF"))
    dxf_files = list(set(dxf_files))

    if not dxf_files:
        print(f"❌ 在目录 {input_dir} 中未找到DXF文件")
        return

    # 构建任务列表
    tasks = []
    for f in dxf_files:
        base = os.path.splitext(os.path.basename(f))[0]
        out = os.path.join(output_dir, f"{base}.prt")
        tasks.append((f, out))

    # 过滤已存在文件
    filtered_tasks = []
    for dxf_file, prt_file in tasks:
        if os.path.exists(prt_file):
            # print(f"⏭️ 跳过: {os.path.basename(prt_file)}（已存在）")
            pass
        else:
            filtered_tasks.append((dxf_file, prt_file))

    if not filtered_tasks:
        print("✅ 所有文件已转换完成")
        return

    # 多进程转换
    print(f"🚀 开始批量转换（多进程）")
    print(f"  输入目录: {input_dir}")
    print(f"  输出目录: {output_dir}")
    print(f"  任务数: {len(filtered_tasks)}, 进程数: {MAX_WORKERS}")
    print("-" * 50)

    results = []
    completed = 0
    total = len(filtered_tasks)

    with ProcessPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_task = {
            executor.submit(_import_single, dxf, prt): (dxf, prt)
            for dxf, prt in filtered_tasks
        }

        for future in as_completed(future_to_task):
            dxf_file, prt_file = future_to_task[future]
            try:
                result = future.result()
                results.append(result)
                completed += 1
                name = os.path.basename(dxf_file)
                status = "✅" if result["success"] else "❌"
                # print(f"[{completed}/{total}] {status} {name}: {result['message']}")
                sys.stdout.flush()
            except Exception as e:
                completed += 1
                print(f"[{completed}/{total}] ❌ 进程错误 {os.path.basename(dxf_file)}: {str(e)}")
                sys.stdout.flush()

    # 统计结果
    print("-" * 50)
    success = sum(1 for r in results if r["success"])
    failed = len(results) - success
    print(f"📊 转换完成!")
    print(f"✅ 成功: {success} 个文件")
    print(f"❌ 失败: {failed} 个文件")
    print(f"📈 总处理: {len(filtered_tasks)} 个文件")


if __name__ == '__main__':
    # 可根据需要修改路径或添加命令行参数解析
    INPUT_DIRECTORY = r"C:\Users\Admin\Desktop\223\file\CAD_pictures\Export\M250195-P1 2D图"
    OUTPUT_DIRECTORY = r"C:\Users\Admin\Desktop\223\file\CAD_pictures\Export\1"
    # 执行批量转换
    batch_convert_dxf_to_prt(INPUT_DIRECTORY, OUTPUT_DIRECTORY)
    print("\n✅ 批量转换任务全部完成!")
