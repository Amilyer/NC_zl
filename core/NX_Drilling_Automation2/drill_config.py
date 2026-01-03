# -*- coding: utf-8 -*-
"""
配置模块
包含常量、路径配置、参数设置等
"""

# 文件路径配置
DRILL_JSON_PATH = r"D:\项目\NC编程\代码实现\tool_file\drill_table.json"
MILL_JSON_PATH = r"D:\项目\NC编程\代码实现\tool_file\knife_table.json"

# MCS坐标系配置
DEFAULT_MCS_NAME = "MCS"
DEFAULT_B_MCS_NAME = "B_MCS"
DEFAULT_FRONT_MCS_NAME = "FRONT_MCS"
DEFAULT_SIDE_MCS_NAME = "RIGHT_MCS"


# 几何体配置
DEFAULT_ORIENT_GEOMETRY_NAME = "WORKPIECE_Z"
DEFAULT_SAFE_DISTANCE = 50.0


# 钻孔参数配置
DEFAULT_TIP_DIAMETER = 1.0  # 刀尖直径
DEFAULT_STEP_DISTANCE = 0.0  # 步距
DEFAULT_DRILL_DIAMETER = 1.0  # 钻孔直径
DEFAULT_DRILL_DEPTH = 1.0  # 钻孔深度


# 钻刀参数配置
TIP_ANGLE = 118.0  # 刀尖角度
TIP_LEN = 2.4034424761102000  # 刀尖长度
CORNER_RADIUS = 0.0  # 拐角半径
LENGTH = 80.0  # 长度
BLADE_LENGTH = 35.0  # 刀刃长度
BLADE_NUMBER = 2  # 刀刃数

# 铣孔参数配置
DEFAULT_X_TOP_OFFSET = 0.5        # 顶部偏移
DEFAULT_AXIAL_MAX_DISTANCE = 0.0  # 轴向步距最大距离
DEFAULT_RADIAL_MAX_DISTANCE = 0.0 # 径向步距最大距离
DEFAULT_STRAT_DIAMETER = 0.0      # 起始直径
DEFAULT_RADIAL_TOOL_NUMBER = 1.0    # 刀路数

# 通用参数配置
DEFAULT_SPINDLE_SPEED = 1000.0  # 主轴转速
DEFAULT_FEED_RATE = 100.0  # 给进率
DEFAULT_TOP_OFFSET = 1.0  # 顶部偏移
DEFAULT_ALL_BOTTOM_OFFSET = 0.0  # 底部偏移
DEFAULT_RAPTO_OFFSET = 0.0
DEFAULT_VECTOR = (0.0, 0.0, 1.0) # 按Z轴旋转
DEFAULT_ORIGIN_POINT = (0.0, 0.0, 0.0)

# 板料余量
DEFAULT_MATERIAL_AMOUNT = 0.7
# 程序组配置
DEFAULT_GEOMETRY_GROUP = "GEOMETRY"

# 加工说明配置
MATERIALS = ["45#", "A3", "CR12", "CR12MOV", "SKD11", "SKH-9", "DC53", "P20", "T00L0X33", "T00L0X44", "合金铜",
             "TOOLOX33", "TOOLOX44"]  # 由于从加工说明中提取出来时。得到的是"T00L0X33", "T00L0X44"，但应为"TOOLOX33", "TOOLOX44"，暂时不知情况，打个补丁

# 刀具驱动点配置
TOOL_DRIVE_POINT_TIP = "SYS_CL_TIP"
TOOL_DRIVE_POINT_SHOULDER = "SYS_CL_SHOULDER"

# 循环类型配置
CYCLE_DRILL_DEEP = "Drill,Deep"
CYCLE_DRILL_STANDARD = "Drill"

# 孔类型配置
HOLE_TYPE_MAP = {
    "螺纹孔": "threaded",
    "入子孔": "insert",
    "弹簧孔": "spring",
    "沉头孔": "countersink",
    "一般孔": "general"
}
