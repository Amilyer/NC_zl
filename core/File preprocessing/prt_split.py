# -*- coding: utf-8 -*-
"""
PRT拆分模块 - 独立版本
从3D PRT文件中拆分出单个实体并导出
"""

import csv
import os

import config

try:
    from tqdm import tqdm
except ImportError:
    def tqdm(iterable, *args, **kwargs):
        return iterable

try:
    import NXOpen
    import NXOpen.UF as UF
except ImportError:
    print("警告: NXOpen 库未找到。此模块需要在 NX 环境中运行。")
    NXOpen = None
    UF = None


def get_uf_session():
    """获取UF会话"""
    return UF.UFSession.GetUFSession()


def sanitize_filename(name):
    """清理文件名中的非法字符"""
    invalid_chars = r'<>:"/\|?*'
    for char in invalid_chars:
        name = name.replace(char, "_")
    return name


def open_part(session, part_path):
    """保证返回 NXOpen.Part 类型"""
    try:
        result = session.Parts.Open(part_path)

        # tuple unpack
        if isinstance(result, tuple):
            part = result[0]
        else:
            part = result

        if isinstance(part, NXOpen.Part):
            return part

        raise TypeError("Open() 未返回 NXOpen.Part 对象")

    except Exception as e:
        raise Exception(f"打开部件失败: {part_path}\n错误信息: {str(e)}")


def export_body_prt(pyuf, output_path, body):
    """导出单个实体为独立 .prt"""
    try:
        pyuf.Part.Export(output_path, 1, [body.Tag])
        return True
    except Exception as e:
        print("导出失败: ", e)
        return False


def measure_body(session, body):
    """使用 NXOpen 方法测量实体包络盒"""
    try:
        sellist = [body]
        ptanchor = NXOpen.Point3d(0.0, 0.0, 0.0)

        # 使用 NXOpen 的测量方法
        boxpoints, dir, edgelengths, ptorigin, ptextreme, volume = session.Measurement.GetBoundingBoxProperties(
            sellist, 1, ptanchor, False)

        length = edgelengths[0]
        width = edgelengths[1]
        height = edgelengths[2]

        return round(length, 3), round(width, 3), round(height, 3)

    except Exception as e:
        print("测量失败: ", e)
        return None, None, None


def split_prt_file(target_part_path: str):
    """
    PRT拆分主函数（兼容旧版）
    
    Args:
        target_part_path: 目标PRT文件路径
    
    Returns:
        tuple: (CSV报告路径, 输出目录路径)
    """
    # 使用默认的输出目录
    work_dir = os.path.dirname(target_part_path)
    output_dir = os.path.join(work_dir, "prt")
    return split_prt_file_with_output(target_part_path, output_dir)


def split_prt_file_with_output(target_part_path: str, output_dir: str):
    """
    PRT文件拆分 - 支持自定义输出目录
    
    Args:
        target_part_path: 输入的PRT文件路径
        output_dir: 输出目录路径
    
    Returns:
        tuple: (CSV报告路径, 输出目录路径)
    """
    if not NXOpen or not UF:
        print("❌ NXOpen 库未加载，无法执行PRT拆分。")
        return None, None

    if not os.path.exists(target_part_path):
        print("❌ 文件不存在: ", target_part_path)
        return None, None

    print("=" * 60)
    print("PRT拆分模块 (优化版)")
    print("=" * 60)
    print(f"输入文件: {target_part_path}")

    session = NXOpen.Session.GetSession()
    pyuf = get_uf_session()

    # 打开部件
    part = open_part(session, target_part_path)
    session.Parts.SetDisplay(part, False, False)
    session.Parts.SetWork(part)

    workPart = session.Parts.Work
    
    # 清理并创建输出目录
    if os.path.exists(output_dir):
        print(f"正在清理输出目录: {output_dir}")
        import shutil
        try:
            # 删除目录下所有内容
            for filename in os.listdir(output_dir):
                file_path = os.path.join(output_dir, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                except Exception as e:
                    print(f"删除文件失败 {file_path}: {e}")
        except Exception as e:
            print(f"清理目录失败: {e}")
            
    os.makedirs(output_dir, exist_ok=True)

    print(f"输出目录: {output_dir}\n")

    # 创建CSV报告
    csv_path = os.path.join(output_dir, "_Export_Report.csv")
    f = open(csv_path, "w", newline="", encoding="utf-8-sig")
    writer = csv.writer(f)
    # 增加 Status 和 Note 列
    writer.writerow(['文件名', '长度_L (mm)', '宽度_W (mm)', '高度_T (mm)', '状态', '备注'])

    # 从配置读取排除图层，如果未配置则使用默认值
    excluded_layers = set(getattr(config, 'PRT_SPLIT_LAYERS_EXCLUDED', [60, 70, 100, 101, 255]))
    exported_names = set()

    # 添加计数器
    skipped_count = 0
    export_count = 0
    failed_count = 0

    bodies = list(workPart.Bodies)
    print(f"发现 {len(bodies)} 个实体\n")

    for body in tqdm(bodies, desc="导出实体", unit="个"):
        if body.Layer in excluded_layers:
            # print(f"跳过图层 {body.Layer} ：{body.Name}") # 减少刷屏
            skipped_count += 1
            continue

        try:
            name = body.Name or body.JournalIdentifier.replace("(", "").replace(")", "")
        except:
            name = f"BODY_{body.Tag}"

        # 1. 基础清理
        base_safe_name = sanitize_filename(name)
        
        # 2. 重名处理
        safe_name = base_safe_name
        counter = 1
        note = ""
        
        while safe_name in exported_names:
            safe_name = f"{base_safe_name}_{counter}"
            counter += 1
            note = f"重名重命名 (原名: {base_safe_name})"
            
        exported_names.add(safe_name)

        new_prt = os.path.join(output_dir, safe_name + ".prt")
        print(f"导出: {safe_name}.prt ... ", end="")

        # 3. 导出
        if not export_body_prt(pyuf, new_prt, body):
            print("❌ 导出失败")
            writer.writerow([safe_name + ".prt", 0, 0, 0, "Failed", "导出API调用失败"])
            failed_count += 1
            continue

        # 4. 测量
        L, W, T = measure_body(session, body)

        if L is None:
            print("⚠ 尺寸测量失败")
            writer.writerow([safe_name + ".prt", 0, 0, 0, "Warning", "导出成功但测量失败"])
            # 仍然算作导出成功，只是测量失败
            export_count += 1
            continue

        writer.writerow([safe_name + ".prt", L, W, T, "Success", note])
        print(f"✔ (L:{L}, W:{W}, T:{T}) {note}")
        export_count += 1

    f.close()
    print("\n完成，数据已写入：", csv_path)
    print("\n统计:")
    print(f"  - 成功导出: {export_count}")
    print(f"  - 失败: {failed_count}")
    print(f"  - 跳过 (排除图层): {skipped_count}")
    print(f"  - 总计实体: {len(bodies)}")
    print("\n--- 导出流程结束 ---")

    # 清理内存
    try:
        if part:
            part.Close(NXOpen.BasePart.CloseWholeTree.TrueValue, NXOpen.BasePart.CloseModified.UseResponses, None)
    except Exception as e:
        print(f"关闭部件时出错: {e}")

    import gc
    gc.collect()

    return csv_path, output_dir


def main():
    """独立运行入口"""
    # 使用配置文件中的路径
    target_path = config.FILE_INPUT_PRT
    
    if not os.path.exists(target_path):
        print(f"❌ 输入文件不存在: {target_path}")
        print("请在 config.py 中配置正确的 FILE_INPUT_PRT 路径")
        return

    split_prt_file(target_path)


if __name__ == "__main__":
    main()
