# -*- coding: utf-8 -*-
"""
NX钻孔自动化模块包

模块说明：
- drill_config.py: 配置模块，包含常量、路径配置、参数设置等
- utils.py: 工具函数模块，包含通用功能：日志输出、数学计算、坐标处理等
- geometry.py: 几何体处理模块，包含坐标系创建、圆识别、草图创建等功能
- path_optimization.py: 路径优化模块，包含TSP路径优化算法
- process_info.py: 加工信息处理模块，包含注释解析、加工参数提取等功能
- parameter_parser.py: 参数解析模块，包含加工参数解析、孔属性提取等功能
- drilling_operations.py: 钻孔操作模块，包含刀具创建、工序设置、钻孔操作等功能
- mirror_operations.py: 镜像操作模块，包含镜像曲线、边界检测等功能
- drill_library.py: 钻刀库模块，包含钻孔参数查询、材质匹配等功能
- main_workflow.py: 主流程模块，包含完整钻孔自动化的主流程控制
- main.py: 主程序入口
"""

__version__ = "2.0.0"
__author__ = "NX钻孔自动化团队"
__description__ = "NX钻孔自动化脚本 - 模块化重构版本"