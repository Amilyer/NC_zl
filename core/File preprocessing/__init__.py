"""
CVFH 特征提取和检索模块

使用示例:
    import cvfh_module as cvfh
    
    # 提取特征
    feature = cvfh.extract_feature("path/to/file.pcd")
    
    # 搜索最近邻
    indices, distances = cvfh.search_knn(feature, k=5)
"""

import os
import sys
from pathlib import Path

# 添加 DLL 搜索路径
_module_dir = Path(__file__).parent
_dll_dir = _module_dir / "dll"

if sys.platform == 'win32' and _dll_dir.exists():
    # Python 3.8+ 使用 add_dll_directory
    if hasattr(os, 'add_dll_directory'):
        os.add_dll_directory(str(_dll_dir))
    # 兼容旧版本
    else:
        dll_path = str(_dll_dir)
        if dll_path not in os.environ.get('PATH', ''):
            os.environ['PATH'] = dll_path + os.pathsep + os.environ.get('PATH', '')

# 切换到模块目录加载 DLL（DLL 内部使用相对路径 data/）
_original_cwd = os.getcwd()
os.chdir(str(_module_dir))

# 导入核心模块

# 恢复原始工作目录
os.chdir(_original_cwd)

