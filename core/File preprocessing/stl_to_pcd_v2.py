"""
STL 转 PCD 转换工具 V2 - 使用 numpy-stl

解决 Open3D 读取二进制 STL 时的编码问题
"""

from pathlib import Path

import numpy as np
import open3d as o3d
from stl import mesh as stl_mesh


def stl_to_pcd(stl_file: str, output_pcd: str = None, 
               point_count: int = 50000, 
               visualize: bool = False) -> list:
    """
    将 STL 文件转换为 PCD 点云
    
    使用 numpy-stl 读取 STL，避免编码问题
    """
    print(f"\n{'='*60}")
    print("STL 转 PCD (V2)")
    print(f"{'='*60}")
    
    stl_path = Path(stl_file)
    
    if not stl_path.exists():
        print(f"✗ 文件不存在: {stl_file}")
        return [False, None]
    
    if output_pcd is None:
        output_pcd = str(stl_path.with_suffix('.pcd'))
    
    try:
        # 1. 使用 numpy-stl 读取 STL（支持二进制格式）
        print(f"读取 STL: {stl_path.name}")
        stl_object = stl_mesh.Mesh.from_file(str(stl_path))
        
        print(f"  三角形数: {len(stl_object.vectors)}")
        
        # 2. 转换为 Open3D 的 TriangleMesh
        # 提取顶点和三角形
        vertices = stl_object.vectors.reshape(-1, 3)
        
        # 去重顶点并创建索引映射
        vertices_unique, indices = np.unique(vertices, axis=0, return_inverse=True)
        triangles = indices.reshape(-1, 3)
        
        print(f"  顶点数: {len(vertices_unique)}")
        print(f"  三角形数: {len(triangles)}")
        
        # 3. 创建 Open3D 三角网格
        mesh = o3d.geometry.TriangleMesh()
        mesh.vertices = o3d.utility.Vector3dVector(vertices_unique)
        mesh.triangles = o3d.utility.Vector3iVector(triangles)
        mesh.compute_vertex_normals()
        
        # 4. 在网格表面均匀采样点云
        print(f"\n在网格表面采样 {point_count} 个点...")
        pcd = mesh.sample_points_uniformly(number_of_points=point_count)
        
        print(f"  实际点数: {len(pcd.points)}")
        
        # 6. 估计法线
        print("估计法线...")
        pcd.estimate_normals(
            search_param=o3d.geometry.KDTreeSearchParamHybrid(
                radius=0.1, max_nn=30
            )
        )
        
        # 7. 可视化（在保存前，这样即使保存失败也能看到）
        if visualize:
            print("\n显示点云...")
            o3d.visualization.draw_geometries(
                [pcd],
                window_name=f"点云: {stl_path.name}",
                width=800,
                height=600
            )
        
        # 8. 保存 PCD
        print(f"\n保存 PCD: {Path(output_pcd).name}")
        
        # 确保使用绝对路径
        output_pcd_abs = str(Path(output_pcd).resolve())
        print(f"  完整路径: {output_pcd_abs}")
        
        # 确保输出目录存在
        Path(output_pcd_abs).parent.mkdir(parents=True, exist_ok=True)
        
        o3d.io.write_point_cloud(output_pcd_abs, pcd)
        
        if Path(output_pcd_abs).exists():
            file_size = Path(output_pcd_abs).stat().st_size / 1024
            print("✓ 转换成功")
            print(f"  文件大小: {file_size:.1f} KB")
        else:
            print("✗ 保存失败")
            if not visualize:  # 只有在不可视化时才返回False
                return [False, None]
        
        print(f"{'='*60}\n")
        return [True, output_pcd_abs]
        
    except Exception as e:
        import traceback
        print(f"✗ 转换失败: {e}")
        traceback.print_exc()
        return [False, None]


def batch_convert(input_folder: str, output_folder: str = None,
                 point_count: int = 50000) -> int:
    """批量转换"""
    input_path = Path(input_folder)
    
    if not input_path.exists():
        print(f"✗ 文件夹不存在: {input_folder}")
        return 0
    
    if output_folder:
        output_path = Path(output_folder)
        output_path.mkdir(exist_ok=True)
    else:
        output_path = input_path
    
    # 查找 STL 文件（不区分大小写）
    stl_files = []
    for ext in ['*.stl', '*.STL']:
        stl_files.extend(input_path.glob(ext))
    
    # 去重（避免重复文件）
    stl_files = list(set(stl_files))
    
    if len(stl_files) == 0:
        print(f"✗ 未找到 STL 文件: {input_folder}")
        return 0
    
    print(f"\n{'='*60}")
    print("批量转换 STL → PCD (V2)")
    print(f"{'='*60}")
    print(f"输入文件夹: {input_folder}")
    print(f"输出文件夹: {output_path}")
    print(f"找到 {len(stl_files)} 个 STL 文件")
    print(f"{'='*60}\n")
    
    success_count = 0
    
    for i, stl_file in enumerate(stl_files, 1):
        print(f"[{i}/{len(stl_files)}] 处理: {stl_file.name}")
        
        output_pcd = output_path / stl_file.with_suffix('.pcd').name
        
        if stl_to_pcd(str(stl_file), str(output_pcd), point_count):
            success_count += 1
    
    print(f"{'='*60}")
    print(f"完成！成功转换 {success_count}/{len(stl_files)} 个文件")
    print(f"{'='*60}\n")
    
    return success_count


def main():
    """命令行主程序"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='STL 转 PCD 转换工具 V2 - 使用 numpy-stl',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 转换单个文件
  python stl_to_pcd_v2.py model.stl
  
  # 批量转换
  python stl_to_pcd_v2.py "D:\\变形程度整理prt\\stl" --batch -n 10000
  
  # 可视化
  python stl_to_pcd_v2.py model.stl --visualize
        """
    )
    
    parser.add_argument('input', nargs='?', help='输入 STL 文件或文件夹')
    parser.add_argument('-o', '--output', help='输出 PCD 文件')
    parser.add_argument('-n', '--points', type=int, default=50000,
                       help='目标点数 (默认: 50000)')
    parser.add_argument('--batch', action='store_true',
                       help='批量转换文件夹')
    parser.add_argument('--visualize', action='store_true',
                       help='可视化显示点云')
    
    args = parser.parse_args()
    
    if not args.input:
        # 交互式模式
        print("\n╔══════════════════════════════════════╗")
        print("║  STL 转 PCD 转换工具 V2              ║")
        print("╚══════════════════════════════════════╝")
        
        choice = input("\n选择: 1)单个文件 2)批量转换: ").strip()
        
        if choice == "1":
            stl_file = input("STL 文件路径: ").strip().strip('"')
            point_count = input("采样点数 [50000]: ").strip()
            point_count = int(point_count) if point_count else 50000
            visualize = input("可视化? (y/n) [n]: ").strip().lower() == 'y'
            
            stl_to_pcd(stl_file, point_count=point_count, visualize=visualize)
        
        elif choice == "2":
            folder = input("文件夹路径: ").strip().strip('"')
            point_count = input("采样点数 [10000]: ").strip()
            point_count = int(point_count) if point_count else 50000
            
            batch_convert(folder, point_count=point_count)
    
    else:
        if args.batch:
            batch_convert(args.input, args.output, args.points)
        else:
            stl_to_pcd(args.input, args.output, args.points, args.visualize)


if __name__ == '__main__':
    main()
