# -*- coding: UTF-8 -*-
"""
配置文件 (Configuration)
功能：定义项目全局常量、路径、参数
"""

from pathlib import Path

# ==============================================================================
# 1. 核心路径配置 (Core Paths)
# ==============================================================================
PROJECT_ROOT = Path(r"C:\Projects\NCv4.8.1")
INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"

# 默认输入文件
FILE_INPUT_PRT = INPUT_DIR / "M250209-P3-2025.11.21.prt"
FILE_INPUT_DXF = INPUT_DIR / "M250209-P3-2025.12.13.dxf"

# 文件夹
# FILE_INPUT_PRT = INPUT_DIR / "M250209-P4.2025.11.25.prt"
# FILE_INPUT_DXF = INPUT_DIR / "M250209-P4.2025.11.25.dxf"

# ==============================================================================
# 2. 资源与模型路径 (Resources & Models)
# ==============================================================================
# 机器学习模型
DIR_MODELS = INPUT_DIR / "point_cloud" / "models" / "saved"
FILE_MODEL_RF = DIR_MODELS / "rf_model_core.pkl"
FILE_MODEL_SCALER = DIR_MODELS / "scaler_core.pkl"
FILE_MODEL_CONFIG = DIR_MODELS / "config_core.pkl"

# DLL 基础目录
DLL_DIR = PROJECT_ROOT / "core" / "DLL"

# 具体 DLL 文件
FILE_DLL_TEXTURE = DLL_DIR / "贴图.dll"
FILE_DLL_FACE_INFO = DLL_DIR / "遍历读取写入面的信息.dll"
FILE_DLL_NAVIGATOR = DLL_DIR / "导航器提取.dll"
FILE_DLL_GEOMETRY_ANALYSIS_1 = DLL_DIR / "加工面方向判断1.dll"
FILE_DLL_GEOMETRY_ANALYSIS_20 = DLL_DIR / "加工面方向判断.dll"
FILE_DLL_SCREENSHOT = DLL_DIR / "FlipAndShot" / "lalalalala.dll"


# ==============================================================================
# 3. 全局处理参数 (Global Processing Parameters)
# ==============================================================================
# 多进程
PROCESS_MAX_WORKERS = 8

# STL/PCD 处理
PROCESS_STL_TOLERANCE = 0.05
PROCESS_STL_ANGLE_TOLERANCE = 5
DIMENSION_TOLERANCE = 1.0  # mm
PROCESS_PCD_POINT_COUNT = 50000

# ==============================================================================
# 4. 各步骤特定配置 (Step-Specific Configs)
# ==============================================================================

# --- Step 1: PRT 拆分 ---
PRT_SPLIT_LAYERS_EXCLUDED = [60, 70, 100, 101, 255]

# --- Step 2: DXF 拆分 ---
DXF_SPLIT_KEYWORDS_REQUIRED = ['加工说明']
DXF_SPLIT_KEYWORDS_EXCLUDED = ['厂内标准件', '订购', '装配', '组配']
DXF_SPLIT_MAP_CLASSIFY = {
    "Drawing1": ["PH", "PPS", "PS", "PU", "subdrawing_001", "subdrawing_003", "U1", "U2", "UP"]
}

# --- Step 7: 删除孔和面 ---
LAYER_WORK = 20
LAYER_SOURCE = 1
LAYER_TARGET = 20
COLOR_INDEX_TARGET = 186  # 红色

# --- Step 8: 面信息提取 ---
DIR_NAME_FACE_INFO = "Face_Info_Reports"
LAYER_FACE_INFO_TARGET = 20

# --- Step 9: 导航器提取 ---
DIR_NAME_NAVIGATOR = "导航器特征"
LAYER_NAVIGATOR_TARGET = 20
LAYER_NAV_1 = 1
LAYER_NAV_20 = 20
LAYER_NAVIGATOR_TARGET_MERGED = 0

# --- Step 10: 几何分析 ---
DIR_NAME_GEOMETRY_ANALYSIS = "Geometry_Analysis_Reports"
LAYER_GEOMETRY_ANALYSIS_TARGET = 20

# --- Step 12: 刀具创建 (含钻头) ---
FILE_TOOL_PARAMS_WITH_DRILL = INPUT_DIR / "铣刀参数_追加钻头数据.xlsx"

# --- Step 8: 简单刀具创建 (仅铣刀) ---
FILE_MILL_TOOLS_EXCEL = INPUT_DIR / "铣刀参数.json"

# --- Step 16: NC 代码生成 ---
NC_POST_NAME = "钢料通用"
NC_GROUP_NAMES = [
    "D1", "D2", "背面打点", "背面钻孔", "背面铣孔", "开粗", "全精", "半精",
    "正面", "正面打点工序", "正面钻孔", "正面铰孔", "正面铣孔",
    "背面", "背面打点", "背面钻孔", "背面铰孔", "背面铣孔",
    "侧面", "侧面打点", "侧面钻孔", "侧面铰孔", "侧面背铣孔", "侧面正铣孔"
]

# ==============================================================================
# 5. 路径字符串转换 (Path String Convertors)
# ==============================================================================
def get_str_path(path):
    """将 Path 对象转换为字符串路径"""
    return str(path) if isinstance(path, Path) else path

# 为所有配置项添加字符串版本的访问器 (Compatibility)
FILE_INPUT_PRT_STR = get_str_path(FILE_INPUT_PRT)
FILE_INPUT_DXF_STR = get_str_path(FILE_INPUT_DXF)
FILE_MODEL_RF_STR = get_str_path(FILE_MODEL_RF)
FILE_MODEL_SCALER_STR = get_str_path(FILE_MODEL_SCALER)
FILE_MODEL_CONFIG_STR = get_str_path(FILE_MODEL_CONFIG)
FILE_DLL_TEXTURE_STR = get_str_path(FILE_DLL_TEXTURE)
FILE_DLL_FACE_INFO_STR = get_str_path(FILE_DLL_FACE_INFO)
FILE_DLL_NAVIGATOR_STR = get_str_path(FILE_DLL_NAVIGATOR)
FILE_DLL_GEOMETRY_ANALYSIS_1_STR = get_str_path(FILE_DLL_GEOMETRY_ANALYSIS_1)
FILE_DLL_GEOMETRY_ANALYSIS_20_STR = get_str_path(FILE_DLL_GEOMETRY_ANALYSIS_20)
FILE_DLL_SCREENSHOT_STR = get_str_path(FILE_DLL_SCREENSHOT)
FILE_TOOL_PARAMS_WITH_DRILL_STR = get_str_path(FILE_TOOL_PARAMS_WITH_DRILL)
FILE_MILL_TOOLS_EXCEL_STR = get_str_path(FILE_MILL_TOOLS_EXCEL)
PROJECT_ROOT_STR = get_str_path(PROJECT_ROOT)
DIR_MODELS_STR = get_str_path(DIR_MODELS)