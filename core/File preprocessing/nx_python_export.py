"""
使用 NX Open Python API 导出 STL - 命令行独立运行版本

注意: 此脚本需要:
1. 安装 Siemens NX
2. 配置 NX Python 环境
3. ✅ 可以从命令行直接运行，无需打开 NX GUI！

运行方式:
    方法1: 直接运行 (如果环境配置正确)
        python nx_python_export.py
    
    方法2: 使用 NX 批处理模式
        run_journal nx_python_export.py
    
    方法3: 使用 ugii_env 启动
        ugii_env python nx_python_export.py

适用场景: 完全自动化的批处理，无需 GUI
"""

from pathlib import Path

import NXOpen
import NXOpen.UF


def get_nx_version():
    """获取 NX 版本信息"""
    try:
        theSession = NXOpen.Session.GetSession()
        version = theSession.GetEnvironmentVariableValue("UGII_VERSION")
        return version if version else "Unknown"
    except:
        return "Unknown"


def export_prt_to_stl_batch(input_folder, output_folder, 
                            triangle_tolerance=0.05, 
                            angle_tolerance=10.0):
    """
    批量导出 PRT 为 STL
    
    Args:
        input_folder: PRT 文件夹
        output_folder: STL 输出文件夹
        triangle_tolerance: 三角形容差 (mm)
        angle_tolerance: 角度容差 (度)
    """
    # 获取 NX Session
    theSession = NXOpen.Session.GetSession()
    theUfSession = NXOpen.UF.UFSession.GetUFSession()
    
    # 显示 NX 版本
    nx_version = get_nx_version()
    print(f"NX 版本: {nx_version}")
    
    # 创建输出文件夹
    output_path = Path(output_folder)
    output_path.mkdir(exist_ok=True)
    
    # 获取所有 PRT 文件
    input_path = Path(input_folder)
    prt_files = list(input_path.glob("*.prt"))
    
    print(f"找到 {len(prt_files)} 个 PRT 文件")
    
    success_count = 0
    
    for i, prt_file in enumerate(prt_files, 1):
        try:
            print(f"\n[{i}/{len(prt_files)}] 处理: {prt_file.name}")
            
            # 打开零件（兼容不同 NX 版本）
            part_load_status = None
            try:
                # 方法1: 新版本 NX (2020+)，返回 (part, status)
                part, part_load_status = theSession.Parts.OpenActiveDisplay(
                    str(prt_file),
                    NXOpen.DisplayPartOption.AllowAdditional
                )
            except (TypeError, ValueError, SyntaxError):
                # 方法2: 老版本 NX (12-19)，需要传入 status 对象
                try:
                    part_load_status = NXOpen.PartLoadStatus()
                    part = theSession.Parts.OpenActiveDisplay(
                        str(prt_file),
                        NXOpen.DisplayPartOption.AllowAdditional,
                        part_load_status
                    )
                except:
                    # 方法3: 最简单的方式
                    part = theSession.Parts.Open(str(prt_file))
                    theSession.Parts.SetDisplay(part, False, False)
            
            # 输出 STL 文件路径
            stl_file = output_path / prt_file.with_suffix('.stl').name
            
            print(f"  目标路径: {stl_file}")
            
            # 导出 STL
            export_single_stl(part, str(stl_file), 
                            triangle_tolerance, angle_tolerance)
            
            # 验证文件是否真的创建了
            if stl_file.exists():
                file_size = stl_file.stat().st_size / 1024  # KB
                print(f"  文件大小: {file_size:.2f} KB")
            else:
                print("  ⚠ 警告: 文件未创建！")
                # 尝试查找可能的位置
                possible_locations = [
                    Path(stl_file.name),  # 当前目录
                    Path.home() / stl_file.name,  # 用户主目录
                    Path(part.FullPath).parent / stl_file.name,  # PRT文件目录
                ]
                for loc in possible_locations:
                    if loc.exists():
                        print(f"  找到文件: {loc}")
                        break
            
            # 关闭零件（兼容不同版本）
            try:
                # 方法1: 尝试标准关闭
                theSession.Parts.CloseAll(1, part_load_status)  # 1 = CloseWholeTree.True
            except (TypeError, AttributeError):
                # 方法2: 简化调用
                try:
                    theSession.Parts.CloseAll(1)
                except:
                    # 方法3: 只关闭当前零件
                    try:
                        part.Close(1, 0, None)  # CloseWholeTree=1, CloseModified=0
                    except:
                        pass  # 忽略关闭错误
            
            print(f"  ✓ 成功: {stl_file.name}")
            success_count += 1
            
        except Exception as e:
            print(f"  ✗ 失败: {e}")
            import traceback
            print("  详细错误信息:")
            traceback.print_exc()
    
    print(f"\n完成！成功转换 {success_count}/{len(prt_files)} 个文件")
    return success_count


def export_single_stl(part, stl_path, triangle_tolerance, angle_tolerance):
    """导出单个零件为 STL（使用 UF API - 最稳定）"""
    theSession = NXOpen.Session.GetSession()
    theUfSession = NXOpen.UF.UFSession.GetUFSession()
    
    # 确保使用绝对路径
    stl_path_abs = str(Path(stl_path).resolve())
    
    try:
        # 收集所有实体
        solid_bodies = [body for body in part.Bodies if body.IsSolidBody]
        
        if len(solid_bodies) == 0:
            raise Exception("未找到实体")
        
        # 打开 STL 文件
        file_base_name = Path(stl_path_abs).name
        header = f"Header: {file_base_name}"
        
        # 使用 UF API 打开二进制 STL 文件
        file_handle = theUfSession.Std.OpenBinaryStlFile(
            stl_path_abs,
            False,  # append = False (新文件)
            header
        )
        
        # 导出每个实体到 STL
        for body in solid_bodies:
            # Tag.Null 在某些 NX 版本中直接用 0 表示
            num_errors, info_error = theUfSession.Std.PutSolidInStlFile(
                file_handle,
                0,  # Tag.Null
                body.Tag,
                0.0,  # max_facet_edges (0 = 自动)
                0.0,  # adjacency_tolerance (0 = 使用triangle_tolerance)
                triangle_tolerance
            )
        
        # 关闭文件
        theUfSession.Std.CloseStlFile(file_handle)
        
    except Exception as e:
        raise Exception(f"导出 STL 失败: {str(e)}")

    return stl_path_abs


# ========================================
# 命令行主程序
# ========================================
if __name__ == '__main__':
    import argparse
    import sys
    
    parser = argparse.ArgumentParser(
        description='NX PRT 批量导出为 STL - 命令行版本',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 基本用法
  python nx_python_export.py -i input_folder -o output_folder
  
  # 指定精度
  python nx_python_export.py -i input_folder -o output_folder -t 0.01 -a 5
  
  # 使用 NX 批处理模式
  run_journal nx_python_export.py
        """
    )

    parser.add_argument('-i', '--input',
                       help='输入 PRT 文件夹路径', default="D:\变形程度整理prt\容易变形prt")
    parser.add_argument('-o', '--output',
                       help='输出 STL 文件夹路径', default="D:\变形程度整理prt\stl")
    parser.add_argument('-t', '--triangle-tolerance', type=float, default=0.05,
                       help='三角形容差 (mm), 默认: 0.05')
    parser.add_argument('-a', '--angle-tolerance', type=float, default=5.0,
                       help='角度容差 (度), 默认: 5.0')

    args = parser.parse_args()

    # 交互式模式
    if not args.input or not args.output:
        print("\n" + "="*60)
        print("NX PRT 批量导出工具 - 命令行版本")
        print("="*60)
        print("\n请输入参数:\n")

        input_folder = input("输入 PRT 文件夹路径: ").strip()
        output_folder = input("输出 STL 文件夹路径: ").strip()

        triangle_tol = input("三角形容差 (mm) [默认 0.05]: ").strip()
        triangle_tolerance = float(triangle_tol) if triangle_tol else 0.05

        angle_tol = input("角度容差 (度) [默认 5.0]: ").strip()
        angle_tolerance = float(angle_tol) if angle_tol else 5.0
    else:
        input_folder = args.input
        output_folder = args.output
        triangle_tolerance = args.triangle_tolerance
        angle_tolerance = args.angle_tolerance

    # 验证路径
    if not Path(input_folder).exists():
        print(f"\n✗ 错误: 输入文件夹不存在: {input_folder}")
        sys.exit(1)

    # 显示配置
    print("\n" + "="*60)
    print("配置信息")
    print("="*60)
    print(f"输入文件夹: {input_folder}")
    print(f"输出文件夹: {output_folder}")
    print(f"三角形容差: {triangle_tolerance} mm")
    print(f"角度容差: {angle_tolerance} 度")
    print("="*60 + "\n")

    # 执行转换
    try:
        success_count = export_prt_to_stl_batch(
            input_folder,
            output_folder,
            triangle_tolerance=triangle_tolerance,
            angle_tolerance=angle_tolerance
        )

        print("\n" + "="*60)
        if success_count > 0:
            print(f"✓ 批量转换完成！成功 {success_count} 个文件")
        else:
            print("✗ 没有成功转换任何文件")
        print("="*60)

    except Exception as e:
        print(f"\n✗ 错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
