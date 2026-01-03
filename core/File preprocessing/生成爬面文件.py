#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
特征驱动生成脚本：基于 FeatureRecognition_Log 和 face_csv 生成半爬面(往复等高) 与 爬面 两个 JSON

流程：
1. 从 FeatureRecognition_Log.csv 中剔除含 "POCKET" 特征的所有面
2. 从 face_csv 中剔除"平面且法向与Z轴平行"的面（顶面/底面）
3. 剩余面按邻接关系分组 → 生成半爬面(往复等高) JSON（智能选刀）
4. 在上述基础上再剔除"法向与Z轴垂直"的面
5. 重新按邻接关系分组 → 生成爬面 JSON（固定刀具 10R5）
"""

import os
import json
import math
import re
from collections import defaultdict, deque
from pathlib import Path

import pandas as pd

# ==================================================================================
# 配置
# ==================================================================================

# 基础路径配置
BASE_PATH = r"C:\Projects\NC\file"
WORKPIECE_FOLDER = os.path.join(r"C:\Projects\NC_New\NC\output\04_PRT_with_Tool")

# 运行时路径（由 setup_paths_for_part 函数填充）
CONFIG = {
    # 输入文件路径（运行时填充）
    "FEATURE_LOG_CSV": "",      # FeatureRecognition_Log.csv
    "FACE_DATA_CSV": "",        # face_data.csv
    "PART_PATH": "",
    "TOOL_JSON_PATH": os.path.join(r"C:\Projects\NC_New\NC\input\铣刀参数.json"),  # 所有零件共用
    "DIRECTION_CSV_PATH": "",
    "EXCEL_PARAMS_PATH": r"C:\Projects\NC\file\output_file\零件参数.xlsx",  # 零件参数表

    # 输出 JSON（运行时填充）
    "BANPAMIAN_JSON_PATH": "",  # 半爬面(往复等高)
    "BANPAMIAN_VERTICAL_JSON_PATH": "",  # 垂直侧面往复等高(余量为0)
    "PAMIAN_JSON_PATH": "",     # 爬面

    # 垂直面往复等高固定刀具
    "VERTICAL_FIXED_TOOL": "D10",
    # 选刀角度阈值：小于此角度用球刀，大于等于用牛鼻刀
    "TOOL_CATEGORY_ANGLE_THRESHOLD": 45.0,
    # 刀具类别
    "TOOL_CATEGORY_BALL": "钨钢球刀",
    "TOOL_CATEGORY_BULL_NOSE": "钨钢牛鼻刀",
    "TOOL_CATEGORY_FLAT": "钨钢平刀",  # 钨钢平刀类别
    # 默认刀具（兖底）
    "DEFAULT_TOOL": "10R5",

    # 往复等高默认刀具（用于智能选刀时的兜底）
    "BANPAMIAN_DEFAULT_TOOL": "17R0.8",
    # 爬面固定刀具
    "PAMIAN_FIXED_TOOL": "10R5",

    # 爬面固定参数
    "PAMIAN_FIXED": {
        "STEEP_METHOD": "非陡峭",
        "STEEP_ANGLE": 90.0,
        "CUT_PATTERN": "往复",
        "STEPOVER_TYPE": "恒定",
        "STEPOVER_UNIT": "mm",
        "CUT_ANGLE": 45.0,
        "OVERLAP_TYPE": "距离",
        "OVERLAP_DISTANCE": 0.2,
        "INTOL": 0.02,
        "OUTTOL": 0.02,
    },

    # 往复等高固定参数
    "BANPAMIAN_FIXED": {
        "REFERENCE_TOOL": "NULL",
    },

    # 默认值（当无法从表中获取时使用）
    "DEFAULT_CUT_DEPTH": 0.5,
    "DEFAULT_FEED": 0.0,
    "DEFAULT_SPINDLE_RPM": 6000.0,
    "DEFAULT_TRAVERSE": 0.0,
    "DEFAULT_LAYER": 20,
}

# 角度阈值
PARALLEL_Z_THRESHOLD = 1.0      # 法向与Z轴平行的阈值（度），angle < 此值视为平行
VERTICAL_Z_THRESHOLD = 1.0      # 法向与Z轴垂直的阈值（度），|angle - 90| < 此值视为垂直

# 材质到切深列名的映射
MATERIAL_COLUMN_MAP = {
    '45#': '45#,A3,切深',
    'A3': '45#,A3,切深',
    'CR12': 'CR12热处理前切深',
    'CR12MOV': 'CR12mov,SKD11,SKH-9,DC53,热处理前切深',
    'SKD11': 'CR12mov,SKD11,SKH-9,DC53,热处理前切深',
    'SKH-9': 'CR12mov,SKD11,SKH-9,DC53,热处理前切深',
    'DC53': 'CR12mov,SKD11,SKH-9,DC53,热处理前切深',
    'P20': 'P20切深',
    'TOOLOX33': 'TOOLOX33 TOOLOX44切深',
    'TOOLOX44': 'TOOLOX33 TOOLOX44切深',
    'T00L0X33': 'TOOLOX33 TOOLOX44切深',
    'T00L0X44': 'TOOLOX33 TOOLOX44切深',
    '合金铜': '合金铜切深',
}

# 热处理后的材质列名映射
MATERIAL_COLUMN_MAP_HEAT_TREATED = {
    'CR12': 'CR12热处理后切深',
    'CR12MOV': 'CR12mov,SKD11,SKH-9,DC53,热处理后切深',
    'SKD11': 'CR12mov,SKD11,SKH-9,DC53,热处理后切深',
    'SKH-9': 'CR12mov,SKD11,SKH-9,DC53,热处理后切深',
    'DC53': 'CR12mov,SKD11,SKH-9,DC53,热处理后切深',
}


# ==================================================================================
# 路径组装辅助函数
# ==================================================================================
def extract_part_name_from_path(prt_path: str) -> str:
    """
    从完整的PRT文件路径中提取零件名（不含.prt后缀）
    例如：C:\\Projects\\NC\\output\\04_PRT_with_Tool\\UP-02.prt -> UP-02
          D:\\Data\\DIE-03_modified.prt -> DIE-03_modified
    
    Args:
        prt_path: PRT文件的完整路径
        
    Returns:
        零件名（不含.prt后缀）
    """
    # 获取文件名（包含扩展名）
    filename = os.path.basename(prt_path)
    # 去掉扩展名
    part_name, _ = os.path.splitext(filename)
    return part_name


def extract_part_code(part_name: str) -> str:
    """
    从 PART_NAME 提取 PART_CODE（零件代号）
    例如：DIE-03_modified -> DIE-03
          PU-25-JXJX_modified -> PU-25-JXJX
    """
    match = re.match(r"([A-Za-z]+-\d+(?:-[A-Za-z0-9]+)?)", part_name)
    if match:
        return match.group(1).upper()
    return part_name.split('_')[0].upper()


# ==================================================================================
# 图层映射函数
# ==================================================================================
def get_layer_by_direction(direction: str) -> int:
    """根据方向返回对应的图层"""
    layer_map = {
        '+Z': 20, '-Z': 70,
        '+X': 40, '-X': 30,
        '+Y': 60, '-Y': 50
    }
    return layer_map.get(direction.strip(), CONFIG["DEFAULT_LAYER"])


# ==================================================================================
# 材质提取函数
# ==================================================================================
# 延迟加载的全局变量
_MATERIAL_DF = None


def _load_material_excel(excel_path: str) -> pd.DataFrame:
    """延迟加载零件参数表，只加载一次"""
    global _MATERIAL_DF
    if _MATERIAL_DF is None:
        try:
            df = pd.read_excel(excel_path)
            print(f"[INFO] 成功加载零件参数表，共 {len(df)} 行")
            _MATERIAL_DF = df
        except Exception as e:
            print(f"[ERROR] 读取零件参数表失败: {e}")
            _MATERIAL_DF = pd.DataFrame()
    return _MATERIAL_DF


def get_material_from_filename(prt_folder: str, excel_path: str):
    """
    根据PRT文件名称从零件参数表中获取材质和热处理信息
    
    参数:
        prt_folder: PRT文件所在的文件夹或PRT文件路径
        excel_path: 零件参数表路径
    
    返回:
        Tuple[材质字符串, 是否热处理]
    """
    global _MATERIAL_DF

    prt_folder = Path(prt_folder)
    excel_path = Path(excel_path)

    # 如果传入的是文件路径，提取文件夹
    if prt_folder.is_file():
        prt_path = prt_folder
        prt_folder = prt_folder.parent
    else:
        # 查找文件夹下唯一的 .prt 文件
        if not prt_folder.exists():
            print(f"[ERROR] 零件文件夹不存在: {prt_folder}")
            return "45#", False
        
        prt_files = list(prt_folder.glob("*.prt"))
        if not prt_files:
            print("[WARN] 未找到任何 .prt 文件，使用默认材质 45#")
            return "45#", False
        if len(prt_files) > 1:
            print(f"[WARN] 发现多个prt文件，使用第一个: {prt_files[0].name}")
        prt_path = prt_files[0]

    filename = prt_path.stem  # 去掉 .prt 后缀

    # 提取前缀（如 DIE-03_modified → DIE-03）
    match = re.match(r"([A-Z]+-\d+)", filename.upper())
    if match:
        prefix = match.group(1)  # 如 DIE-03
    else:
        prefix = filename.upper().split('_')[0]

    # 加载Excel表
    df = _load_material_excel(excel_path)
    if df.empty:
        print("[WARN] 零件参数表为空，使用默认材质 45#")
        return "45#", False

    # 智能匹配：优先精确匹配"文件名称"或"零件名称"或"编号"列包含该前缀
    mask = (
        df.astype(str).apply(lambda col: col.str.contains(prefix, case=False, na=False)).any(axis=1)
    )
    matched_rows = df[mask]

    if matched_rows.empty:
        print(f"[WARN] 未在零件参数表中找到编号包含 '{prefix}' 的记录，使用默认材质 45#")
        return "45#", False

    row = matched_rows.iloc[0]  # 取第一条匹配的

    # 提取材质
    material_candidates = ['材质']
    material = None
    for col in material_candidates:
        if col in row and pd.notna(row[col]):
            material = str(row[col]).strip()
            break

    if not material:
        print("[WARN] 未找到材质信息，使用默认 45#")
        material = "45#"

    # 判断是否热处理：热处理列有内容即为已热处理
    heat_treatment_candidates = ['热处理']
    is_heat_treated = False
    for col in heat_treatment_candidates:
        if col in row and pd.notna(row[col]):
            ht_text = str(row[col]).strip()
            if ht_text and ht_text.lower() not in ['无', '否', '-', '']:
                is_heat_treated = True
                print(f"[INFO] 检测到热处理: {ht_text}")
                break

    print(f"[INFO] 材质: {material}, 热处理: {is_heat_treated}")
    return material, is_heat_treated


def get_material_and_heat_treatment(part_file: str):
    """
    从零件参数表中获取材质和热处理信息（兼容旧接口）
    
    参数:
        part_file: PRT文件路径
    
    返回:
        Tuple[材质字符串, 是否热处理]
    """
    excel_path = CONFIG.get("EXCEL_PARAMS_PATH", r"D:\Projects\NC\output\00_Resources\CSV_Reports\零件参数.xlsx")
    return get_material_from_filename(part_file, excel_path)


# ==================================================================================
# 辅助函数 - 解析点坐标
# ==================================================================================
def parse_point(point_str: str) -> tuple:
    """解析点坐标字符串，格式: 'x,y,z'"""
    try:
        coords = point_str.strip().split(',')
        if len(coords) == 3:
            return tuple(map(float, coords))
    except (ValueError, AttributeError):
        pass
    return None


def find_bottom_face_in_group(face_tags: list, face_df: pd.DataFrame) -> int:
    """在一组面中找到Z坐标最小的面（最底层面）"""
    point_col = "Face Data - Point"
    min_z = float('inf')
    bottom_face = None
    
    for tag in face_tags:
        # 从face_df中找到该面的行
        row = face_df[face_df["Face Tag"] == tag]
        if row.empty:
            continue
            
        # 获取面的点坐标
        point_str = row.iloc[0].get(point_col)
        if pd.notna(point_str):
            coords = parse_point(str(point_str))
            if coords and len(coords) == 3:
                z_value = coords[2]  # Z坐标
                if z_value < min_z:
                    min_z = z_value
                    bottom_face = tag
    
    return bottom_face


# ==================================================================================
# 刀具参数读取
# ==================================================================================
def read_tool_parameters(tool_file: str) -> pd.DataFrame:
    """读取铣刀参数表（JSON 版本）"""
    try:
        with open(tool_file, 'r', encoding='utf-8') as f:
            tools_list = json.load(f)
        
        df = pd.DataFrame(tools_list)
        
        # 确保数值列类型正确
        if '直径' in df.columns:
            df['直径'] = pd.to_numeric(df['直径'], errors='coerce')
        
        speed_feed_columns = ['转速(普)', '进给(普)', '横越(普)']
        for col in speed_feed_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        print(f"[INFO] 成功读取 {len(df)} 把刀具参数")
        return df

    except Exception as e:
        print(f"[ERROR] 读取刀具参数表时出错: {e}")
        return pd.DataFrame()


# ==================================================================================
# 刀具R角提取
# ==================================================================================
def extract_r_angle_from_tool_name(tool_name: str) -> float:
    """
    从刀具名称中提取R角
    例如：
    - "10R5" -> 5.0
    - "6R3.5" -> 3.5
    - "D10" -> 0.0 (平刀无R角)
    - "4R2" -> 2.0
    """
    if not tool_name:
        return 0.0
    
    # 使用正则匹配 R 后面的数字（支持小数）
    match = re.search(r'R(\d+\.?\d*)', tool_name, re.IGNORECASE)
    if match:
        return float(match.group(1))
    return 0.0


def select_clearing_tool_from_ball_mills(tools_df: pd.DataFrame, min_rad_data: float) -> str:
    """
    从钨钢球刀类别中选择清根刀具
    
    规则：选择R角 < min_rad_data 且R角最大的刀具（刚好小于最小Rad Data）
    
    Args:
        tools_df: 刀具参数表
        min_rad_data: 需清根面中的最小短半径
    
    Returns:
        选中的刀具名称，若无合适刀具返回None
    """
    if tools_df.empty or min_rad_data <= 0:
        return None
    
    # 筛选钨钢球刀类别
    ball_mills = tools_df[tools_df['类别'] == CONFIG["TOOL_CATEGORY_BALL"]].copy()
    
    if ball_mills.empty:
        print(f"[WARN] 未找到钨钢球刀类别的刀具")
        return None
    
    # 筛选R角 < min_rad_data 的刀具
    valid_tools = ball_mills[ball_mills['R角'] < min_rad_data].copy()
    
    if valid_tools.empty:
        print(f"[WARN] 钨钢球刀中无R角 < {min_rad_data} 的刀具")
        return None
    
    # 选择R角最大的刀具（刚好小于min_rad_data）
    max_r_row = valid_tools.loc[valid_tools['R角'].idxmax()]
    tool_name = max_r_row['刀具名称']
    
    print(f"[INFO] 清根选刀: 最小Rad Data={min_rad_data}, 选择刀具 '{tool_name}' (R角={max_r_row['R角']})")
    return tool_name


def get_max_diameter_tool(tool_names: list, tools_df: pd.DataFrame) -> str:
    """
    从刀具名称列表中找出直径最大的刀具
    
    Args:
        tool_names: 刀具名称列表
        tools_df: 刀具参数表
    
    Returns:
        直径最大的刀具名称
    """
    if not tool_names or tools_df.empty:
        return None
    
    max_diameter = -1
    max_tool = None
    
    for tool_name in tool_names:
        tool_row = tools_df[tools_df['刀具名称'] == tool_name]
        if not tool_row.empty:
            diameter = tool_row.iloc[0].get('直径', 0)
            if diameter > max_diameter:
                max_diameter = diameter
                max_tool = tool_name
    
    return max_tool


def get_min_rad_data_for_component(component: list, face_df: pd.DataFrame) -> float:
    """
    获取组内所有面的最小非零Rad Data值
    
    Args:
        component: 面标签列表
        face_df: 面数据DataFrame
    
    Returns:
        最小非零Rad Data值，如果都是零或空则返回None
    """
    rad_data_col = "Face Data - Rad Data"
    if rad_data_col not in face_df.columns:
        return None
    
    non_zero_values = []
    for face_tag in component:
        row = face_df[face_df["Face Tag"] == face_tag]
        if not row.empty:
            rad_data = row.iloc[0].get(rad_data_col)
            if pd.notna(rad_data):
                try:
                    val = float(rad_data)
                    if val > 0:
                        non_zero_values.append(val)
                except (ValueError, TypeError):
                    pass
    
    return min(non_zero_values) if non_zero_values else None


def get_min_radius_for_vertical_component(component: list, face_df: pd.DataFrame) -> float:
    """
    获取垂直面组内所有面的最小非零Radius值（专门用于垂直面组的R角判断）
    
    Args:
        component: 面标签列表
        face_df: 面数据DataFrame
    
    Returns:
        最小非零Radius值，如果都是零或空则返回None
    """
    radius_col = "Face Data - Radius"
    if radius_col not in face_df.columns:
        return None
    
    non_zero_values = []
    for face_tag in component:
        row = face_df[face_df["Face Tag"] == face_tag]
        if not row.empty:
            radius_data = row.iloc[0].get(radius_col)
            if pd.notna(radius_data):
                try:
                    val = float(radius_data)
                    if val > 0:
                        non_zero_values.append(val)
                except (ValueError, TypeError):
                    pass
    
    return min(non_zero_values) if non_zero_values else None


def select_tool_for_vertical_with_radius(tools_df: pd.DataFrame, min_rad_data: float) -> str:
    """
    根据最小Rad Data为有R角的垂直面组选刀
    

    
    规则：
    1. 计算最大允许直径 d = min_rad_data * 2
    2. 从钨钢平刀类别中选择直径 <= d 的最大刀具
    
    Args:
        tools_df: 刀具参数表
        min_rad_data: 组内最小非零Rad Data
    
    Returns:
        选中的刀具名称
    """
    if tools_df.empty or min_rad_data is None or min_rad_data <= 0:
        return CONFIG["VERTICAL_FIXED_TOOL"]  # 默认返回D10
    
    # 计算最大允许直径（刀具直径不能超过这个值）
    max_allowed_diameter = min_rad_data * 2
    
    # 筛选钨钢平刀类别
    flat_tools = tools_df[tools_df['类别'] == CONFIG["TOOL_CATEGORY_FLAT"]].copy()
    
    if flat_tools.empty:
        print(f"[WARN] 未找到钨钢平刀类别的刀具，使用默认D10")
        return CONFIG["VERTICAL_FIXED_TOOL"]
    
    # 筛选直径 <= max_allowed_diameter 的刀具
    valid_tools = flat_tools[flat_tools['直径'] < max_allowed_diameter].copy()
    
    if valid_tools.empty:
        # 没有满足条件的刀具，选择该类别中直径最小的
        min_diameter_row = flat_tools.loc[flat_tools['直径'].idxmin()]
        tool_name = min_diameter_row['刀具名称']
        print(f"[WARN] 钨钢平刀中无直径 <= {max_allowed_diameter:.1f}mm 的刀具，选择最小直径: {tool_name} (直径={min_diameter_row['直径']}mm)")
        return tool_name
    
    # 选择满足条件的最大刀具（直径最大但<=最大允许直径）
    valid_tools = valid_tools.sort_values('直径', ascending=False)
    selected_row = valid_tools.iloc[0]
    tool_name = selected_row['刀具名称']
    
    print(f"[INFO] 垂直面组有R角: 最小Rad Data={min_rad_data:.2f}, 最大允许直径={max_allowed_diameter:.1f}mm, 选择 '{tool_name}' (直径={selected_row['直径']}mm)")
    return tool_name


# ==================================================================================
# 智能选刀
# ==================================================================================
def get_max_angle_for_component(component: list, angle_map: dict) -> float:
    """获取一个分组中所有面的最大陡峭角度"""
    angles = [angle_map[tag] for tag in component if tag in angle_map and angle_map[tag] is not None]
    return max(angles) if angles else None


def get_min_angle_for_component(component: list, angle_map: dict) -> float:
    """获取一个分组中所有面的最小陡峭角度（最缓面）"""
    angles = [angle_map[tag] for tag in component if tag in angle_map and angle_map[tag] is not None]
    return min(angles) if angles else None


def select_tool_by_angle_and_category(tools_df: pd.DataFrame, max_angle: float, default_tool: str) -> str:
    """
    根据最大陡峭角度智能选刀（按类别限制）
    
    规则：
    1. max_angle < 60° → 在"钨钢球刀"类别中选
    2. max_angle >= 60° → 在"钨钢牛鼻刀"类别中选
    3. d = 0.8 / sin(max_angle)，如果 d < 10 则扩大到 10
    4. 在对应类别中选 直径 >= d 且最接近 d 的刀具
    5. 如果该类别中没有满足条件的刀具，选该类别中直径最大的刀具
    """
    if tools_df.empty or '直径' not in tools_df.columns or '类别' not in tools_df.columns:
        print(f"[WARN] 刀具表为空或缺少必要列，使用默认刀具: {default_tool}")
        return default_tool
    
    if max_angle is None or max_angle <= 0:
        print(f"[WARN] 无效的最大角度，使用默认刀具: {default_tool}")
        return default_tool
    
    # 根据角度选择刀具类别
    angle_threshold = CONFIG["TOOL_CATEGORY_ANGLE_THRESHOLD"]
    if max_angle < angle_threshold:
        category = CONFIG["TOOL_CATEGORY_BALL"]  # 钨钢球刀
        category_desc = "缓面"
    else:
        category = CONFIG["TOOL_CATEGORY_BULL_NOSE"]  # 钨钢牛鼻刀
        category_desc = "陡面"
    
    # 筛选该类别的刀具
    category_tools = tools_df[tools_df['类别'] == category].copy()
    
    if category_tools.empty:
        print(f"[WARN] 未找到类别 '{category}' 的刀具，使用默认刀具: {default_tool}")
        return default_tool
    
    # 计算所需直径
    angle_rad = math.radians(max_angle)
    sin_val = math.sin(angle_rad)
    
    if sin_val <= 0.001:
        print(f"[WARN] sin({max_angle}°) 接近0，使用默认刀具: {default_tool}")
        return default_tool
    
    d = 0.8 / sin_val
    if d < 10:
        d = 10
    
    print(f"[DEBUG] {category_desc}组: 最大角度={max_angle:.2f}°, 类别={category}, 需求直径d={d:.3f}mm")
    
    # 在该类别中选择刀具
    valid_tools = category_tools[category_tools['直径'] >= d].copy()
    
    if valid_tools.empty:
        # 该类别中没有满足条件的刀具，选该类别中直径最大的
        max_diameter_row = category_tools.loc[category_tools['直径'].idxmax()]
        tool_name = max_diameter_row['刀具名称']
        print(f"[WARN] 类别 '{category}' 中无直径 >= {d:.3f}mm 的刀具，选择最大直径: {tool_name} (直径={max_diameter_row['直径']}mm)")
        return tool_name
    
    # 选择直径 >= d 且最接近 d 的刀具
    valid_tools = valid_tools.sort_values('直径', ascending=True)
    selected_row = valid_tools.iloc[0]
    tool_name = selected_row['刀具名称']
    
    print(f"[INFO] 智能选刀: 类别={category}, 选择刀具 '{tool_name}' (直径={selected_row['直径']}mm)")
    return tool_name


def select_tool_by_angle_from_flat_category(tools_df: pd.DataFrame, max_angle: float, default_tool: str) -> str:
    """
    根据最大陡峭角度从钨钢平刀类别中智能选刀（用于半精_爬面的斜面组）
    
    规则：
    1. 固定从"钨钢平刀"类别中选刀
    2. d = 0.8 / sin(max_angle)，如果 d < 10 则扩大到 10
    3. 在钨钢平刀类别中选 直径 >= d 且最接近 d 的刀具
    4. 如果该类别中没有满足条件的刀具，选该类别中直径最大的刀具
    """
    if tools_df.empty or '直径' not in tools_df.columns or '类别' not in tools_df.columns:
        print(f"[WARN] 刀具表为空或缺少必要列，使用默认刀具: {default_tool}")
        return default_tool
    
    if max_angle is None or max_angle <= 0:
        print(f"[WARN] 无效的最大角度，使用默认刀具: {default_tool}")
        return default_tool
    
    # 固定使用钨钢平刀类别
    category = CONFIG["TOOL_CATEGORY_FLAT"]  # 钨钢平刀
    
    # 筛选该类别的刀具
    category_tools = tools_df[tools_df['类别'] == category].copy()
    
    if category_tools.empty:
        print(f"[WARN] 未找到类别 '{category}' 的刀具，使用默认刀具: {default_tool}")
        return default_tool
    
    # 计算所需直径
    angle_rad = math.radians(max_angle)
    sin_val = math.sin(angle_rad)
    
    if sin_val <= 0.001:
        print(f"[WARN] sin({max_angle}°) 接近0，使用默认刀具: {default_tool}")
        return default_tool
    
    d = 0.8 / sin_val
    if d < 10:
        d = 10
    
    print(f"[DEBUG] 半精_爬面斜面组: 最大角度={max_angle:.2f}°, 类别={category}, 需求直径d={d:.3f}mm")
    
    # 在该类别中选择刀具
    valid_tools = category_tools[category_tools['直径'] >= d].copy()
    
    if valid_tools.empty:
        # 该类别中没有满足条件的刀具，选该类别中直径最大的
        max_diameter_row = category_tools.loc[category_tools['直径'].idxmax()]
        tool_name = max_diameter_row['刀具名称']
        print(f"[WARN] 类别 '{category}' 中无直径 >= {d:.3f}mm 的刀具，选择最大直径: {tool_name} (直径={max_diameter_row['直径']}mm)")
        return tool_name
    
    # 选择直径 >= d 且最接近 d 的刀具
    valid_tools = valid_tools.sort_values('直径', ascending=True)
    selected_row = valid_tools.iloc[0]
    tool_name = selected_row['刀具名称']
    
    print(f"[INFO] 智能选刀（钨钢平刀）: 类别={category}, 选择刀具 '{tool_name}' (直径={selected_row['直径']}mm)")
    return tool_name


def select_tool_for_banpamian_slope(tools_df: pd.DataFrame, min_angle: float, max_angle: float, default_tool: str) -> str:
    """
    根据最缓面和最陡面角度智能选刀（用于半精_爬面的斜面组）
    
    规则：
    1. 根据最缓面角度选择刀具类别：
       - min_angle < 45° → 钨钢球刀
       - min_angle >= 45° → 钨钢牛鼻刀
    2. 根据最陡面角度计算所需直径：d = 0.8 / sin(max_angle)，最小10mm
    3. 在选定类别中选择直径 >= d 且最接近 d 的刀具
    4. 如果该类别中没有满足条件的刀具，选该类别中直径最大的刀具
    """
    if tools_df.empty or '直径' not in tools_df.columns or '类别' not in tools_df.columns:
        print(f"[WARN] 刀具表为空或缺少必要列，使用默认刀具: {default_tool}")
        return default_tool
    
    if max_angle is None or max_angle <= 0:
        print(f"[WARN] 无效的最大角度，使用默认刀具: {default_tool}")
        return default_tool
    
    if min_angle is None:
        min_angle = max_angle  # 如果没有最小角度，使用最大角度
    
    # 根据最缓面角度选择刀具类别
    angle_threshold = 45.0
    if min_angle < angle_threshold:
        category = CONFIG["TOOL_CATEGORY_BALL"]  # 钨钢球刀
        category_desc = f"最缓面{min_angle:.1f}°<45°，选球刀"
    else:
        category = CONFIG["TOOL_CATEGORY_BULL_NOSE"]  # 钨钢牛鼻刀
        category_desc = f"最缓面{min_angle:.1f}°>=45°，选牛鼻刀"
    
    # 筛选该类别的刀具
    category_tools = tools_df[tools_df['类别'] == category].copy()
    
    if category_tools.empty:
        print(f"[WARN] 未找到类别 '{category}' 的刀具，使用默认刀具: {default_tool}")
        return default_tool
    
    # 根据最陡面角度计算所需直径
    angle_rad = math.radians(max_angle)
    sin_val = math.sin(angle_rad)
    
    if sin_val <= 0.001:
        print(f"[WARN] sin({max_angle}°) 接近0，使用默认刀具: {default_tool}")
        return default_tool
    
    d = 0.8 / sin_val
    if d < 10:
        d = 10
    
    print(f"[DEBUG] 半精_爬面斜面组: {category_desc}, 最陡面={max_angle:.1f}°, 需求直径d={d:.3f}mm")
    
    # 在该类别中选择刀具
    valid_tools = category_tools[category_tools['直径'] >= d].copy()
    
    if valid_tools.empty:
        # 该类别中没有满足条件的刀具，选该类别中直径最大的
        max_diameter_row = category_tools.loc[category_tools['直径'].idxmax()]
        tool_name = max_diameter_row['刀具名称']
        print(f"[WARN] 类别 '{category}' 中无直径 >= {d:.3f}mm 的刀具，选择最大直径: {tool_name} (直径={max_diameter_row['直径']}mm)")
        return tool_name
    
    # 选择直径 >= d 且最接近 d 的刀具
    valid_tools = valid_tools.sort_values('直径', ascending=True)
    selected_row = valid_tools.iloc[0]
    tool_name = selected_row['刀具名称']
    
    print(f"[INFO] 智能选刀（{category}）: 选择刀具 '{tool_name}' (直径={selected_row['直径']}mm)")
    return tool_name


def is_component_vertical(component: list, angle_map: dict) -> bool:
    """
    判断一个分组是否为"垂直面组"
    规则：组内所有面都是垂直面（与Z轴夹角接近90°）
    """
    for tag in component:
        angle = angle_map.get(tag)
        if angle is not None and not is_vertical_to_z(angle):
            return False
    return True


def classify_component_by_slope_count(component: list, angle_map: dict, threshold: float) -> str:
    gentle_count = 0
    steep_count = 0
    for tag in component:
        angle = angle_map.get(tag)
        if angle is None:
            steep_count += 1
            continue
        if angle <= threshold:
            gentle_count += 1
        else:
            steep_count += 1
    return "steep" if steep_count >= gentle_count else "gentle"


def get_tool_parameters(tools_df: pd.DataFrame, tool_name: str, material: str, is_heat_treated: bool) -> dict:
    """根据刀具名称、材质和热处理状态获取切深、进给、转速、横越、刀具类别"""
    result = {
        'cut_depth': CONFIG["DEFAULT_CUT_DEPTH"],
        'feed': CONFIG["DEFAULT_FEED"],
        'spindle_rpm': CONFIG["DEFAULT_SPINDLE_RPM"],
        'traverse': CONFIG["DEFAULT_TRAVERSE"],
        'category': "",  # 刀具类别
    }
    
    if tools_df.empty:
        return result
    
    tool_row = tools_df[tools_df['刀具名称'] == tool_name]
    if tool_row.empty:
        print(f"[WARN] 未找到刀具 '{tool_name}'，使用默认参数")
        return result
    
    tool_row = tool_row.iloc[0]
    
    # 获取刀具类别
    if '类别' in tool_row.index and pd.notna(tool_row['类别']):
        result['category'] = str(tool_row['类别'])
    
    material_upper = material.upper() if material else "45#"
    
    if is_heat_treated and material_upper in MATERIAL_COLUMN_MAP_HEAT_TREATED:
        column_name = MATERIAL_COLUMN_MAP_HEAT_TREATED[material_upper]
    elif material_upper in MATERIAL_COLUMN_MAP:
        column_name = MATERIAL_COLUMN_MAP[material_upper]
    else:
        column_name = '45#,A3,切深'
    
    if column_name in tool_row.index:
        depth = tool_row[column_name]
        if pd.notna(depth):
            result['cut_depth'] = float(depth)
    
    if '转速(普)' in tool_row.index and pd.notna(tool_row['转速(普)']):
        result['spindle_rpm'] = float(tool_row['转速(普)'])
    
    if '进给(普)' in tool_row.index and pd.notna(tool_row['进给(普)']):
        result['feed'] = float(tool_row['进给(普)'])
    
    if '横越(普)' in tool_row.index and pd.notna(tool_row['横越(普)']):
        result['traverse'] = float(tool_row['横越(普)'])
    
    return result


# ==================================================================================
# 方向映射读取
# ==================================================================================
def read_direction_mapping(direction_file: str) -> dict:
    """读取方向映射文件"""
    try:
        df = pd.read_csv(direction_file)
        direction_map = {}
        for column in df.columns:
            col_name = column.strip()
            for face_tag in df[column].dropna():
                try:
                    direction_map[int(face_tag)] = col_name
                except (ValueError, TypeError):
                    continue
        print(f"[INFO] 成功读取方向映射，共 {len(direction_map)} 个面标签")
        return direction_map
    except Exception as e:
        print(f"[WARN] 读取方向映射文件时出错: {e}")
        return {}


def get_layer_for_component(face_tags: list, direction_map: dict, component_name: str = "") -> int:
    """根据面组件中的面标签获取指定图层"""
    if not direction_map:
        return CONFIG["DEFAULT_LAYER"]
    
    for tag in face_tags:
        if tag in direction_map:
            direction = direction_map[tag]
            return get_layer_by_direction(direction)
    
    return CONFIG["DEFAULT_LAYER"]


# ==================================================================================
# 向量与角度计算
# ==================================================================================
def parse_vector(vec_str):
    """解析向量字符串"""
    if pd.isna(vec_str) or not vec_str:
        return None
    vec_str = str(vec_str).strip().strip('"')
    numbers = re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", vec_str)
    if len(numbers) >= 3:
        return (float(numbers[0]), float(numbers[1]), float(numbers[2]))
    return None


def calculate_angle_with_z(normal):
    """计算法向量与Z轴的夹角（度）"""
    if normal is None:
        return None
    nx, ny, nz = normal
    length = math.sqrt(nx * nx + ny * ny + nz * nz)
    if length < 1e-10:
        return None
    cos_val = abs(nz) / length
    cos_val = min(1.0, max(0.0, cos_val))
    angle_rad = math.acos(cos_val)
    return math.degrees(angle_rad)


def is_parallel_to_z(angle: float) -> bool:
    """判断是否与Z轴平行（顶面/底面）"""
    if angle is None:
        return False
    return angle < PARALLEL_Z_THRESHOLD


def is_vertical_to_z(angle: float) -> bool:
    """判断是否与Z轴垂直"""
    return abs(angle - 90.0) < VERTICAL_Z_THRESHOLD


def is_top_plane(normal_z: float, tolerance: float = 0.999) -> bool:
    """
    判断是否为顶部平面（法向量朝上）
    参数:
        normal_z: 法向量的Z分量
        tolerance: 判断阈值，默认0.999（约1度）
    返回:
        bool: 如果法向量朝上（Z分量接近1）返回True
    """
    return normal_z > tolerance


def add_adjacent_planes_to_json(json_data: dict, face_df: pd.DataFrame, angle_map: dict) -> dict:
    """
    为往复等高JSON中的每个组添加邻接的平面（包括顶部和底部平面）
    
    参数:
        json_data: 往复等高JSON数据
        face_df: 包含面信息的DataFrame，需要有'Face Tag'和'Adjacent Face Tags'列
        angle_map: 面Tag到角度的映射
    返回:
        修改后的JSON数据
    """
    print("\n" + "=" * 50)
    print("  添加邻接平面到往复等高组")
    print("=" * 50)
    
    # 1. 构建邻接关系映射（双向）
    adjacency_map = {}
    for _, row in face_df.iterrows():
        face_tag = str(row['Face Tag'])
        adjacent_str = str(row.get('Adjacent Face Tags', ''))
        
        if adjacent_str and adjacent_str != 'nan':
            adjacent_faces = [s.strip() for s in adjacent_str.split(';') if s.strip()]
            
            # 建立双向邻接关系
            if face_tag not in adjacency_map:
                adjacency_map[face_tag] = set()
            adjacency_map[face_tag].update(adjacent_faces)
            
            # 反向关系
            for adj_face in adjacent_faces:
                if adj_face not in adjacency_map:
                    adjacency_map[adj_face] = set()
                adjacency_map[adj_face].add(face_tag)
    
    print(f"    构建邻接关系映射，共 {len(adjacency_map)} 个面有邻接关系")
    
    # 2. 识别所有平面（包括顶部和底部平面）
    all_planes = set()  # 所有平面
    top_planes = set()  # 顶部平面
    bottom_planes = set()  # 底部平面
    
    normal_col = "Face Normal"
    if normal_col not in face_df.columns:
        normal_col = "Face Data - Normal Direction"
    
    for _, row in face_df.iterrows():
        face_tag = row['Face Tag']
        angle = angle_map.get(face_tag)
        
        # 检查是否为平面（与Z轴夹角接近0度）
        if angle is not None and angle < 1.0:
            # 获取法向量
            normal = parse_vector(row.get(normal_col))
            if normal:
                face_tag_str = str(face_tag)
                all_planes.add(face_tag_str)
                
                # 判断是顶部还是底部平面
                if normal[2] > 0.999:  # 法向量Z分量>0.999，朝上
                    top_planes.add(face_tag_str)
                elif normal[2] < -0.999:  # 法向量Z分量<-0.999，朝下
                    bottom_planes.add(face_tag_str)
    
    print(f"    识别到 {len(all_planes)} 个平面：")
    print(f"      - 顶部平面: {len(top_planes)} 个")
    print(f"      - 底部平面: {len(bottom_planes)} 个")
    
    # 3. 为每个组添加邻接的平面
    total_added = 0
    for group_name, group_data in json_data.items():
        if '面ID列表' not in group_data:
            continue
            
        # 获取当前组的所有面（转为字符串）
        current_faces = set(str(face_id) for face_id in group_data['面ID列表'])
        
        # 找出所有邻接面
        all_adjacent = set()
        for face_id in current_faces:
            if face_id in adjacency_map:
                all_adjacent.update(adjacency_map[face_id])
        
        # 筛选出邻接的平面（包括顶部和底部，排除已在组中的）
        adjacent_planes = (all_adjacent & all_planes) - current_faces
        
        if adjacent_planes:
            # 统计顶部和底部平面数量
            adjacent_top_count = len(adjacent_planes & top_planes)
            adjacent_bottom_count = len(adjacent_planes & bottom_planes)
            
            # 将平面加入组
            added_faces = [int(face_id) for face_id in adjacent_planes]
            group_data['面ID列表'].extend(added_faces)
            total_added += len(added_faces)
            print(f"    {group_name}: 添加 {len(added_faces)} 个平面 (顶部: {adjacent_top_count}, 底部: {adjacent_bottom_count})")
    
    print(f"\n    总计添加 {total_added} 个平面到各组")
    print("=" * 50)
    
    return json_data


def build_adjacency_graph(df: pd.DataFrame, target_tags: set) -> dict:
    """构建邻接图"""
    graph = defaultdict(set)

    for _, row in df.iterrows():
        tag = row["Face Tag"]
        if tag not in target_tags:
            continue

        adj_str = row.get("Adjacent Face Tags")
        if pd.isna(adj_str):
            continue

        for s in str(adj_str).split(";"):
            s = s.strip()
            if not s:
                continue
            try:
                neighbor = int(s)
            except ValueError:
                continue

            if neighbor in target_tags:
                graph[tag].add(neighbor)
                graph[neighbor].add(tag)

    # 确保孤立点也在图中
    for t in target_tags:
        _ = graph[t]

    return graph


def find_connected_components(graph: dict, tags: set) -> list:
    """找出所有连通分量"""
    visited = set()
    components = []

    for start in tags:
        if start in visited:
            continue

        queue = deque([start])
        visited.add(start)
        comp = []

        while queue:
            v = queue.popleft()
            comp.append(v)
            for nb in graph[v]:
                if nb not in visited:
                    visited.add(nb)
                    queue.append(nb)

        components.append(sorted(comp))

    return components


def build_adjacency_graph_by_color(df: pd.DataFrame, target_tags: set, color_map: dict) -> dict:
    """
    构建按颜色限制的邻接图
    只有相邻且颜色相同的面才会连边
    """
    graph = defaultdict(set)

    for _, row in df.iterrows():
        tag = row["Face Tag"]
        if tag not in target_tags:
            continue

        tag_color = color_map.get(tag)
        
        adj_str = row.get("Adjacent Face Tags")
        if pd.isna(adj_str):
            continue

        for s in str(adj_str).split(";"):
            s = s.strip()
            if not s:
                continue
            try:
                neighbor = int(s)
            except ValueError:
                continue

            # 只有相邻且颜色相同才连边
            if neighbor in target_tags and color_map.get(neighbor) == tag_color:
                graph[tag].add(neighbor)
                graph[neighbor].add(tag)

    # 确保孤立点也在图中
    for t in target_tags:
        _ = graph[t]

    return graph


def find_components_by_color(df: pd.DataFrame, target_tags: set, color_map: dict) -> list:
    """
    按颜色+邻接关系分组
    每组内的面颜色一致且相邻
    """
    # 构建按颜色限制的邻接图
    graph = build_adjacency_graph_by_color(df, target_tags, color_map)
    
    # 找连通分量
    components = find_connected_components(graph, target_tags)
    
    return components


def load_pocket_face_tags(feature_log_csv: str) -> set:
    """
    从 FeatureRecognition_Log.csv 中加载需要剔除的 POCKET 特征的 FACE_TAG
    注意：STEP1POCKET 类型的面不剔除，保留下来
    """
    try:
        # 尝试不同编码读取
        for encoding in ['utf-8', 'gbk', 'gb2312', 'latin1']:
            try:
                df = pd.read_csv(feature_log_csv, encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            print(f"[ERROR] 无法读取特征日志文件（编码问题）")
            return set()

        print(f"[INFO] 读取特征日志: {len(df)} 行")

        # 筛选特征类型中含 "POCKET" 但不是 "STEP1POCKET" 的行
        pocket_rows = df[
            df['Type'].str.contains('POCKET', case=False, na=False) &
            ~df['Type'].str.upper().str.strip().eq('STEP1POCKET')
        ]

        # 收集 FACE_TAG
        pocket_tags = set()
        for tag in pocket_rows['Attribute'].dropna():
            try:
                pocket_tags.add(int(tag))
            except (ValueError, TypeError):
                continue

        print(f"[INFO] 找到 {len(pocket_tags)} 个需剔除的 POCKET 特征面（不含 STEP1POCKET）")
        return pocket_tags

    except Exception as e:
        print(f"[ERROR] 读取特征日志时出错: {e}")
        return set()


def load_step1pocket_face_tags(feature_log_csv: str) -> set:
    """从 FeatureRecognition_Log.csv 中加载 STEP1POCKET 的 FACE_TAG"""
    try:
        for encoding in ['utf-8', 'gbk', 'gb2312', 'latin1']:
            try:
                df = pd.read_csv(feature_log_csv, encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            print(f"[ERROR] 无法读取特征日志文件（编码问题）")
            return set()

        step1_rows = df[df['Type'].str.upper().str.strip().eq('STEP1POCKET')]
        step1_tags = set()
        for tag in step1_rows['Attribute'].dropna():
            try:
                step1_tags.add(int(tag))
            except (ValueError, TypeError):
                continue

        print(f"[INFO] 找到 {len(step1_tags)} 个 STEP1POCKET 面")
        return step1_tags
    except Exception as e:
        print(f"[ERROR] 读取 STEP1POCKET 时出错: {e}")
        return set()


def get_adjacent_face_count(face_df: pd.DataFrame, face_tag: int) -> int:
    """统计 face_df 中指定面标签的邻接面数量（去重后计数）"""
    row = face_df[face_df["Face Tag"] == face_tag]
    if row.empty:
        return 0

    adj_str = row.iloc[0].get("Adjacent Face Tags")
    if pd.isna(adj_str):
        return 0
    neighbors = set()
    for s in str(adj_str).split(";"):
        s = s.strip()
        if not s:
            continue
        try:
            neighbors.add(int(s))
        except ValueError:
            continue
    return len(neighbors)


def generate_banpamian_json(components: list, group_tools: list, group_tool_params: list,
                            direction_map: dict, fixed_params: dict) -> dict:
    """生成半精_爬面(往复等高) JSON 配置"""
    json_config = {}

    for idx, comp in enumerate(components, start=1):
        op_name = f"往复等高{idx}"
        layer = get_layer_for_component(comp, direction_map, op_name)
        
        tool_name = group_tools[idx - 1]
        tool_params = group_tool_params[idx - 1]
        category = str(tool_params.get('category', '') or '')

        # 根据刀具类别设置侧面/底面余量
        wall_allow = 0.0
        floor_allow = 0.0
        if '飞刀' in category:
            wall_allow = 0.2
            floor_allow = 0.5
        elif '钨钢' in category:
            wall_allow = 0.05
            floor_allow = 0.5

        json_config[op_name] = {
            "工序": "往复等高_SIMPLE",
            "面ID列表": comp,
            "刀具名称": tool_name,
            "刀具类别": category,
            "部件侧面余量": wall_allow,
            "部件底面余量": floor_allow,
            "切深": tool_params['cut_depth'],
            "指定图层": layer,
            "参考刀具": fixed_params.get("REFERENCE_TOOL", "NULL"),
            "进给": tool_params['feed'],
            "转速": tool_params['spindle_rpm'],
            "横越": tool_params['traverse'],
        }

    return json_config


def generate_banpamian_zero_stock_json(components: list, group_tools: list, group_tool_params: list,
                                       direction_map: dict, fixed_params: dict,
                                       angle_map: dict = None) -> dict:
    """
    生成 全精_往复等高(垂直侧面及陡面往复等高) JSON 配置（侧面/底面余量固定为0）
    用于颜色6、垂直Z轴的侧面
    
    切深规则（按单个面判断）：
    - 垂直面（与Z轴夹角≈90°）：切深固定为 10
    - 非垂直面：切深固定为 0.1
    
    注意：同一组内可能包含垂直面和非垂直面，会被拆分成多个工序
    
    Args:
        components: 面组列表
        group_tools: 每组对应的刀具名称列表
        group_tool_params: 每组对应的刀具参数列表
        direction_map: 方向映射
        fixed_params: 固定参数
        angle_map: 面ID→角度映射（用于判断垂直面）
    """
    json_config = {}
    op_idx = 1

    for comp_idx, comp in enumerate(components):
        tool_name = group_tools[comp_idx]
        tool_params = group_tool_params[comp_idx]
        category = str(tool_params.get('category', '') or '')

        # 侧面/底面余量固定为0
        wall_allow = 0.0
        floor_allow = 0.0

        # 按单个面判断垂直/非垂直，拆分成两个子组
        vertical_faces = []
        non_vertical_faces = []
        
        for face_tag in comp:
            if angle_map is not None:
                angle = angle_map.get(face_tag)
                if angle is not None and is_vertical_to_z(angle):
                    vertical_faces.append(face_tag)
                else:
                    non_vertical_faces.append(face_tag)
            else:
                # 无angle_map时默认为非垂直面
                non_vertical_faces.append(face_tag)
        
        # 生成垂直面工序（切深=10）
        if vertical_faces:
            op_name = f"往复等高{op_idx}"
            layer = get_layer_for_component(vertical_faces, direction_map, op_name)
            json_config[op_name] = {
                "工序": "往复等高_SIMPLE",
                "面ID列表": sorted(vertical_faces),
                "刀具名称": tool_name,
                "刀具类别": category,
                "部件侧面余量": wall_allow,
                "部件底面余量": floor_allow,
                "切深": 10.0,
                "指定图层": layer,
                "参考刀具": fixed_params.get("REFERENCE_TOOL", "NULL"),
                "进给": tool_params['feed'],
                "转速": tool_params['spindle_rpm'],
                "横越": tool_params['traverse'],
            }
            op_idx += 1
        
        # 生成非垂直面工序（切深=0.1）
        if non_vertical_faces:
            op_name = f"往复等高{op_idx}"
            layer = get_layer_for_component(non_vertical_faces, direction_map, op_name)
            json_config[op_name] = {
                "工序": "往复等高_SIMPLE",
                "面ID列表": sorted(non_vertical_faces),
                "刀具名称": tool_name,
                "刀具类别": category,
                "部件侧面余量": wall_allow,
                "部件底面余量": floor_allow,
                "切深": 0.1,
                "指定图层": layer,
                "参考刀具": fixed_params.get("REFERENCE_TOOL", "NULL"),
                "进给": tool_params['feed'],
                "转速": tool_params['spindle_rpm'],
                "横越": tool_params['traverse'],
            }
            op_idx += 1

    return json_config


def generate_pamian_json(components: list, group_tools: list, group_tool_params: list,
                         direction_map: dict, fixed_params: dict) -> dict:
    """生成全精爬面 JSON 配置（支持每组不同刀具）"""
    json_config = {}

    for idx, comp in enumerate(components, start=1):
        op_name = f"爬面{idx}"
        layer = get_layer_for_component(comp, direction_map, op_name)
        
        tool_name = group_tools[idx - 1]
        tool_params = group_tool_params[idx - 1]
        
        json_config[op_name] = {
            "工序": "爬面_SIMPLE",
            "面ID列表": comp,
            "刀具名称": tool_name,
            "刀具类别": tool_params['category'],
            "陡峭空间范围方法": fixed_params.get("STEEP_METHOD", "非陡峭"),
            "陡峭壁角度": fixed_params.get("STEEP_ANGLE", 90.0),
            "非陡峭切削模式": fixed_params.get("CUT_PATTERN", "往复"),
            "步距类型": fixed_params.get("STEPOVER_TYPE", "恒定"),
            "切深": tool_params['cut_depth'],
            "步距单位": fixed_params.get("STEPOVER_UNIT", "mm"),
            "剖切角类型": "指定",
            "剖切角_与XC夹角": fixed_params.get("CUT_ANGLE", 45.0),
            "重叠区域类型": fixed_params.get("OVERLAP_TYPE", "距离"),
            "重叠距离": fixed_params.get("OVERLAP_DISTANCE", 0.2),
            "指定图层": layer,
            "内公差": fixed_params.get("INTOL", 0.02),
            "外公差": fixed_params.get("OUTTOL", 0.02),
            "进给": tool_params['feed'],
            "转速": tool_params['spindle_rpm'],
            "横越": tool_params['traverse'],
        }

    return json_config


# ==================================================================================
# 工序合并逻辑
# ==================================================================================
def regroup_by_tool_and_direction(ops: dict, direction_map: dict, operation_prefix: str = "往复等高",
                                   include_cut_depth: bool = False) -> dict:
    """
    将所有工序的面打散，按 刀具+加工方向 (可选+切深) 重新分组
    
    Args:
        ops: 原始工序字典 {op_name: op_dict}
        direction_map: 面ID -> 方向 的映射 {face_tag: "+Z"/"-Z"/"+X"等}
        operation_prefix: 输出工序名前缀
        include_cut_depth: 是否将切深作为分组条件（仅全精_往复等高需要）
    
    Returns:
        重新分组后的工序字典
    """
    if not ops:
        return {}
    
    # 1. 打散所有面，记录每个面对应的刀具、切深和工序参数
    face_to_tool = {}  # face_tag -> 刀具名称
    face_to_cut_depth = {}  # face_tag -> 切深
    face_to_params = {}  # face_tag -> 工序参数（不含面ID列表）
    
    for op_name, op in ops.items():
        tool_name = op.get("刀具名称")
        cut_depth = op.get("切深", 0.1)  # 默认切深0.1
        # 复制参数（不含面ID列表）
        params = {k: v for k, v in op.items() if k != "面ID列表"}
        
        for face_tag in op.get("面ID列表", []):
            face_to_tool[face_tag] = tool_name
            face_to_cut_depth[face_tag] = cut_depth
            face_to_params[face_tag] = params
    
    # 2. 按 刀具+方向 (可选+切深) 重新分组
    grouped = {}
    
    for face_tag, tool_name in face_to_tool.items():
        direction = direction_map.get(face_tag, "UNKNOWN")
        cut_depth = face_to_cut_depth.get(face_tag, 0.1)
        
        if include_cut_depth:
            key = (tool_name, direction, cut_depth)
        else:
            key = (tool_name, direction)
        
        if key not in grouped:
            grouped[key] = {
                "params": face_to_params[face_tag],
                "faces": []
            }
        grouped[key]["faces"].append(face_tag)
    
    # 3. 生成新的工序字典
    result = {}
    for idx, (key, data) in enumerate(grouped.items(), 1):
        op_name = f"{operation_prefix}{idx}"
        op = data["params"].copy()
        op["面ID列表"] = sorted(data["faces"])
        # 根据方向重新计算图层（方向在key的第二个位置）
        direction = key[1]
        op["指定图层"] = get_layer_by_direction(direction)
        result[op_name] = op
    
    group_desc = "刀具+方向+切深" if include_cut_depth else "刀具+方向"
    print(f"    [重新分组] 原 {len(ops)} 个工序 -> 按{group_desc}分为 {len(result)} 个工序")
    return result


def _build_operation_group_key(op):
    """生成用于分组合并的键，忽略面ID列表（用于往复等高）"""
    # 使用与merge_operations_by_params相同的默认值逻辑，确保一致性
    return (
        op.get("工序"),
        op.get("刀具名称"),
        op.get("刀具类别", ""),
        op.get("切深"),
        op.get("指定图层"),
        op.get("参考刀具", "NULL"),
        op.get("转速", 0),
        op.get("进给", 0),
        op.get("横越", 0),
        op.get("部件侧面余量", 0),
        op.get("部件底面余量", 0),
    )


def merge_operations_by_params(ops: dict, operation_prefix: str = "往复等高") -> dict:
    """
    将除面ID列表外参数完全一致的工序合并，减少输出的组数量
    ops: dict mapping op_name -> op_dict
    """
    grouped = {}

    for _, op in ops.items():
        key = _build_operation_group_key(op)
        if key not in grouped:
            grouped[key] = {
                "工序": op.get("工序"),
                "刀具名称": op.get("刀具名称"),
                "刀具类别": op.get("刀具类别", ""),
                "切深": op.get("切深"),
                "指定图层": op.get("指定图层"),
                "参考刀具": op.get("参考刀具", "NULL"),
                "转速": op.get("转速", 0),
                "进给": op.get("进给", 0),
                "横越": op.get("横越", 0),
                "部件侧面余量": op.get("部件侧面余量", 0),
                "部件底面余量": op.get("部件底面余量", 0),
                "面ID列表": []
            }
        grouped[key]["面ID列表"].extend(op.get("面ID列表", []))

    merged_ops = {}
    for idx, merged in enumerate(grouped.values(), 1):
        # 去重并排序面ID，确保为 int
        merged["面ID列表"] = sorted({int(fid) for fid in merged["面ID列表"]})
        merged_ops[f"{operation_prefix}{idx}"] = merged

    return merged_ops


def _build_pamian_group_key(op):
    """爬面专用的合并键，不包含面ID列表"""
    keys = [
        "工序",
        "刀具名称",
        "刀具类别",
        "陡峭空间范围方法",
        "陡峭壁角度",
        "非陡峭切削模式",
        "步距类型",
        "切深",
        "步距单位",
        "剖切角类型",
        "剖切角_与XC夹角",
        "重叠区域类型",
        "重叠距离",
        "指定图层",
        "内公差",
        "外公差",
        "进给",
        "转速",
        "横越",
    ]
    return tuple(op.get(k) for k in keys)


def merge_pamian_operations(ops: dict, prefix="爬面") -> dict:
    """合并爬面 JSON 的工序"""
    if not ops:
        return {}
    
    grouped = {}
    original_count = len(ops)

    for op_name, op in ops.items():
        key = _build_pamian_group_key(op)
        if key not in grouped:
            grouped[key] = {
                k: op[k] for k in op if k != "面ID列表"
            }
            grouped[key]["面ID列表"] = []

        grouped[key]["面ID列表"].extend(op.get("面ID列表", []))

    # 去重并建立新编号
    merged = {}
    for idx, op in enumerate(grouped.values(), 1):
        op["面ID列表"] = sorted({int(i) for i in op["面ID列表"]})
        merged[f"{prefix}{idx}"] = op

    if len(merged) < original_count:
        print(f"    [DEBUG] 合并前: {original_count} 个操作, 合并后: {len(merged)} 个操作")
    
    return merged


def save_json(data, output_path):
    """保存 JSON 文件，自动创建目录"""
    # 如果数据为空，不保存文件
    if not data or (isinstance(data, dict) and len(data) == 0):
        print(f"[INFO] 数据为空，跳过保存: {output_path}")
        return
    
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[SUCCESS] 已保存: {output_path}")


# ==================================================================================
# 主流程
# ==================================================================================
def process_single_part(part_name: str) -> bool:
    """
    处理单个零件
    
    Args:
        part_name: 零件文件名（不含 .prt 后缀）
    
    Returns:
        bool: 处理成功返回 True
    """
    part_code = extract_part_code(part_name)
    
    print("=" * 70)
    print(f"  特征驱动生成脚本")
    print(f"  零件: {part_name} (代号: {part_code})")
    print("=" * 70)

    # -------------------------------------------------------------------------
    # Step 1: 获取材质和热处理信息
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [1] 获取材质和热处理信息")
    print("-" * 50)
    
    material, is_heat_treated = get_material_and_heat_treatment(CONFIG["PART_PATH"])
    if not material:
        material = "45#"
        print(f"[WARN] 未识别到材质，使用默认值: {material}")

    # -------------------------------------------------------------------------
    # Step 2: 读取刀具参数表
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [2] 读取刀具参数表")
    print("-" * 50)
    
    tools_df = read_tool_parameters(CONFIG["TOOL_JSON_PATH"])

    # -------------------------------------------------------------------------
    # Step 3: 读取方向映射
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [3] 读取方向映射")
    print("-" * 50)
    
    direction_map = read_direction_mapping(CONFIG["DIRECTION_CSV_PATH"])

    # -------------------------------------------------------------------------
    # Step 4: 读取面数据并计算角度
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [4] 读取面数据并计算角度")
    print("-" * 50)
    
    face_df = pd.read_csv(CONFIG["FACE_DATA_CSV"])
    print(f"[INFO] 读取面数据: {len(face_df)} 个面")
    
    # 确定法向量列名
    normal_col = "Face Normal"
    if normal_col not in face_df.columns:
        normal_col = "Face Data - Normal Direction"
    
    # 计算每个面与Z轴的夹角，同时构建颜色映射
    angle_map = {}
    color_map = {}
    for _, row in face_df.iterrows():
        tag = row["Face Tag"]
        normal = parse_vector(row.get(normal_col))
        angle = calculate_angle_with_z(normal)
        if angle is not None:
            angle_map[tag] = angle
        # 构建颜色映射
        if "Face Color" in row.index:
            try:
                color_map[tag] = int(row["Face Color"])
            except (ValueError, TypeError):
                pass

    # -------------------------------------------------------------------------
    # Step 5: 从特征日志中获取 POCKET 面
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [5] 从特征日志中剔除 POCKET 特征面")
    print("-" * 50)
    
    pocket_tags = load_pocket_face_tags(CONFIG["FEATURE_LOG_CSV"])

    # -------------------------------------------------------------------------
    # Step 6: 第一次过滤 - 只保留颜色6/66的面 + 剔除POCKET面 + 剔除平行Z轴面
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [6] 第一次过滤：只保留颜色6/66 + 剔除POCKET + 剔除平行Z轴面")
    print("-" * 50)
    
    # 获取所有面的 Face Tag
    all_tags = set(face_df["Face Tag"])
    print(f"    原始面数: {len(all_tags)}")
    
    # 只保留颜色为 6/66 的面（所有面都按颜色过滤，不区分垂直/非垂直）
    ALLOWED_COLORS = {6, 66}
    tags_with_valid_color = {tag for tag in all_tags if color_map.get(tag) in ALLOWED_COLORS}
    print(f"    只保留颜色(6/66)后: {len(tags_with_valid_color)} 个面")
    
    # 剔除 POCKET 面（不含 STEP1POCKET）
    tags_after_pocket = tags_with_valid_color - pocket_tags
    print(f"    剔除 POCKET 后: {len(tags_after_pocket)} 个面")

    # 对保留的 STEP1POCKET 面按邻接数量再过滤：邻接数>2才保留
    step1_tags = load_step1pocket_face_tags(CONFIG["FEATURE_LOG_CSV"])
    step1_in_candidates = step1_tags & tags_after_pocket
    step1_to_remove = set()
    for t in step1_in_candidates:
        adj_cnt = get_adjacent_face_count(face_df, t)
        if adj_cnt <= 2:
            step1_to_remove.add(t)
    if step1_to_remove:
        print(f"    剔除邻接数<=2的 STEP1POCKET 面: {len(step1_to_remove)} 个")
    tags_after_pocket -= step1_to_remove
    print(f"    STEP1POCKET 邻接过滤后: {len(tags_after_pocket)} 个面")

    # 剔除平行Z轴的面（顶面/底面）
    parallel_z_tags = {tag for tag in tags_after_pocket if is_parallel_to_z(angle_map.get(tag))}
    tags_for_banpamian = tags_after_pocket - parallel_z_tags
    print(f"    剔除平行Z轴面后: {len(tags_for_banpamian)} 个面")

    # -------------------------------------------------------------------------
    # Step 7: 构建邻接图并分组 → 生成半爬面(往复等高) JSON
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [7] 分组并生成半爬面(往复等高) JSON")
    print("-" * 50)
    
    if tags_for_banpamian:
        # 按颜色+邻接关系分组（每组颜色一致）
        initial_components = find_components_by_color(face_df, tags_for_banpamian, color_map)
        print(f"    按颜色+邻接分为 {len(initial_components)} 个初始组")
        
        # 处理混合组：分离垂直面
        banpamian_components = []
        extracted_vertical_faces = []
        
        for idx, comp in enumerate(initial_components, start=1):
            is_pure_vertical = is_component_vertical(comp, angle_map)
            
            if is_pure_vertical:
                # 纯垂直面组，直接加入半精_爬面
                banpamian_components.append(comp)
                print(f"    初始组 {idx}: {len(comp)} 个面 - 纯垂直面组")
            else:
                # 混合组或纯斜面组，分离垂直面
                vertical_faces = []
                non_vertical_faces = []
                
                for face_tag in comp:
                    angle = angle_map.get(face_tag)
                    if angle is not None and is_vertical_to_z(angle):
                        vertical_faces.append(face_tag)
                    else:
                        non_vertical_faces.append(face_tag)
                
                if vertical_faces and non_vertical_faces:
                    # 混合组：分离处理
                    print(f"    初始组 {idx}: {len(comp)} 个面 - 混合组（{len(vertical_faces)} 垂直面 + {len(non_vertical_faces)} 斜面）")
                    banpamian_components.append(non_vertical_faces)  # 斜面组
                    extracted_vertical_faces.extend(vertical_faces)  # 垂直面待重组
                elif non_vertical_faces:
                    # 纯斜面组
                    banpamian_components.append(comp)
                    print(f"    初始组 {idx}: {len(comp)} 个面 - 纯斜面组")
                else:
                    # 应该不会到这里（因为前面已判断不是纯垂直面组）
                    banpamian_components.append(comp)
        
        # 将提取的垂直面按邻接关系重新分组并加入半精_爬面
        if extracted_vertical_faces:
            print(f"\n    从混合组提取 {len(extracted_vertical_faces)} 个垂直面，重新分组...")
            extracted_vertical_tags = set(extracted_vertical_faces)
            vertical_components = find_components_by_color(face_df, extracted_vertical_tags, color_map)
            print(f"    垂直面重新分为 {len(vertical_components)} 个组")
            banpamian_components.extend(vertical_components)  # 将重新分组的垂直面加入半精_爬面
        
        print(f"\n    最终分组数: {len(banpamian_components)} 个组")
        
        # 对每组智能选刀
        group_tools = []
        group_tool_params = []
        
        for idx, comp in enumerate(banpamian_components, start=1):
            max_angle = get_max_angle_for_component(comp, angle_map)
            min_angle = get_min_angle_for_component(comp, angle_map)
            is_vertical = is_component_vertical(comp, angle_map)
            
            if is_vertical:
                # 检查垂直面组是否有R角（使用Face Data - Radius列）
                min_radius = get_min_radius_for_vertical_component(comp, face_df)
                if min_radius is not None and min_radius > 0:
                    # 有R角，根据Radius选刀
                    tool_name = select_tool_for_vertical_with_radius(tools_df, min_radius)
                    print(f"    最终组 {idx}: {len(comp)} 个面, 垂直面组(有R角) -> 选择刀具 {tool_name}")
                else:
                    # 无R角，使用默认D10
                    tool_name = CONFIG["VERTICAL_FIXED_TOOL"]
                    print(f"    最终组 {idx}: {len(comp)} 个面, 垂直面组(无R角) -> 固定刀具 {tool_name}")
            else:
                # 斜面组：根据最缓面选择刀具类别（球刀或牛鼻刀），根据最陡面计算直径
                print(f"    最终组 {idx}: {len(comp)} 个面, 最缓面={min_angle:.1f}°, 最陡面={max_angle:.1f}°" if min_angle and max_angle else f"    最终组 {idx}: {len(comp)} 个面, 无有效角度")
                tool_name = select_tool_for_banpamian_slope(tools_df, min_angle, max_angle, CONFIG["DEFAULT_TOOL"])
            
            group_tools.append(tool_name)
            
            tool_params = get_tool_parameters(tools_df, tool_name, material, is_heat_treated)
            group_tool_params.append(tool_params)
        
        # 生成 JSON（所有组都使用往复等高工序）
        banpamian_json = generate_banpamian_json(
            components=banpamian_components,
            group_tools=group_tools,
            group_tool_params=group_tool_params,
            direction_map=direction_map,
            fixed_params=CONFIG["BANPAMIAN_FIXED"]
        )
        # 按刀具+方向重新分组
        try:
            banpamian_json = regroup_by_tool_and_direction(banpamian_json, direction_map, operation_prefix="往复等高")
            print(f"\n    半爬面重新分组后工序数: {len(banpamian_json)}")
        except Exception as e:
            print(f"[WARN] 半爬面重新分组失败: {e}")
            import traceback
            traceback.print_exc()
        save_json(banpamian_json, CONFIG["BANPAMIAN_JSON_PATH"])
    else:
        print("[WARN] 无可用面，未生成半爬面 JSON")
        banpamian_components = []

    # -------------------------------------------------------------------------
    # Step 7b: 对第一次过滤后的组进行分流（陡面组 vs 缓面组）
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [7b] 对第一次过滤后的组进行分流（陡面组 vs 缓面组）")
    print("-" * 50)
    # 陡、缓面分界值
    slope_threshold = 35.0
    # slope_threshold = float(CONFIG["TOOL_CATEGORY_ANGLE_THRESHOLD"])
    steep_components = []  # 陡面组（用于全精_往复等高）
    gentle_components = []  # 缓面组（用于全精_爬面）
    
    if banpamian_components:
        for idx, comp in enumerate(banpamian_components, start=1):
            group_type = classify_component_by_slope_count(comp, angle_map, slope_threshold)
            if group_type == "steep":
                steep_components.append(comp)
                print(f"    组 {idx}: {len(comp)} 个面 -> 陡面组（全精_往复等高）")
            else:
                gentle_components.append(comp)
                print(f"    组 {idx}: {len(comp)} 个面 -> 缓面组（全精_爬面）")
        
        print(f"\n    陡面组数: {len(steep_components)}")
        print(f"    缓面组数: {len(gentle_components)}")
    else:
        print("[WARN] 无可用组进行分流")

    # -------------------------------------------------------------------------
    # Step 8: 处理缓面组 - 生成全精_爬面 JSON（删除垂直面）
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [8] 处理缓面组 - 生成全精_爬面 JSON（删除垂直面）")
    print("-" * 50)
    
    pamian_json = {}
    removed_vertical_faces = []  # 从缓面组中删除的垂直面，后续加入全精_往复等高
    
    if gentle_components:
        # 对每个缓面组，删除垂直面
        pamian_components_processed = []
        pamian_group_tools = []
        pamian_group_tool_params = []
        
        for idx, comp in enumerate(gentle_components, start=1):
            # 分离垂直面和非垂直面
            vertical_faces = [tag for tag in comp if is_vertical_to_z(angle_map.get(tag))]
            non_vertical_faces = [tag for tag in comp if not is_vertical_to_z(angle_map.get(tag))]
            
            if vertical_faces:
                print(f"\n    缓面组 {idx}: 删除 {len(vertical_faces)} 个垂直面")
                removed_vertical_faces.extend(vertical_faces)
            
            # 只保留非垂直面的组用于爬面
            if non_vertical_faces:
                pamian_components_processed.append(non_vertical_faces)
                
                # 对非垂直面组选刀
                max_angle = get_max_angle_for_component(non_vertical_faces, angle_map)
                print(f"    缓面组 {idx} (处理后): {len(non_vertical_faces)} 个面, 最大角度={max_angle:.2f}°" if max_angle else f"    缓面组 {idx} (处理后): {len(non_vertical_faces)} 个面, 无有效角度")
                tool_name = select_tool_by_angle_and_category(tools_df, max_angle, CONFIG["DEFAULT_TOOL"])
                pamian_group_tools.append(tool_name)
                tool_params = get_tool_parameters(tools_df, tool_name, material, is_heat_treated)
                pamian_group_tool_params.append(tool_params)
        
        if pamian_components_processed:
            # 高度过滤
            height_map = {}
            if 'Height' in face_df.columns:
                for _, row in face_df.iterrows():
                    tag = row["Face Tag"]
                    height = row.get("Height")
                    if pd.notna(height):
                        try:
                            height_map[tag] = float(height)
                        except (ValueError, TypeError):
                            pass
            
            PAMIAN_MIN_HEIGHT = 2.1
            components_after_height = []
            filtered_count = 0
            for comp in pamian_components_processed:
                if len(comp) == 1:
                    face_tag = comp[0]
                    face_height = height_map.get(face_tag, 999)
                    if face_height <= PAMIAN_MIN_HEIGHT:
                        filtered_count += 1
                        continue
                components_after_height.append(comp)
            
            if filtered_count > 0:
                print(f"    过滤掉 {filtered_count} 个单面且高度≤{PAMIAN_MIN_HEIGHT}mm 的组")
            
            # 更新工具列表（只保留通过高度过滤的组）
            if len(components_after_height) < len(pamian_components_processed):
                # 需要重新匹配工具
                pamian_group_tools_filtered = []
                pamian_group_tool_params_filtered = []
                for comp in components_after_height:
                    max_angle = get_max_angle_for_component(comp, angle_map)
                    tool_name = select_tool_by_angle_and_category(tools_df, max_angle, CONFIG["DEFAULT_TOOL"])
                    pamian_group_tools_filtered.append(tool_name)
                    tool_params = get_tool_parameters(tools_df, tool_name, material, is_heat_treated)
                    pamian_group_tool_params_filtered.append(tool_params)
                pamian_group_tools = pamian_group_tools_filtered
                pamian_group_tool_params = pamian_group_tool_params_filtered
            
            if components_after_height:
                pamian_json = generate_pamian_json(
                    components=components_after_height,
                    group_tools=pamian_group_tools,
                    group_tool_params=pamian_group_tool_params,
                    direction_map=direction_map,
                    fixed_params=CONFIG["PAMIAN_FIXED"]
                )
                print(f"    生成全精_爬面JSON，合并前工序数: {len(pamian_json)}")
                # 按刀具+方向重新分组
                try:
                    pamian_json = regroup_by_tool_and_direction(pamian_json, direction_map, operation_prefix="爬面")
                    print(f"    全精_爬面重新分组后工序数: {len(pamian_json)}")
                except Exception as e:
                    print(f"[WARN] 全精_爬面重新分组失败: {e}")
                    import traceback
                    traceback.print_exc()
                save_json(pamian_json, CONFIG["PAMIAN_JSON_PATH"])
            else:
                print("[WARN] 高度过滤后无可用组，未生成全精_爬面 JSON")
        else:
            print("[WARN] 缓面组处理后无可用面，未生成全精_爬面 JSON")
    else:
        print("[WARN] 无缓面组，未生成全精_爬面 JSON")
    
    print(f"\n    从缓面组删除的垂直面数: {len(removed_vertical_faces)}")

    # -------------------------------------------------------------------------
    # Step 9: 生成全精_往复等高 JSON（陡面组 + 从缓面组删除的垂直面）
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [9] 生成全精_往复等高 JSON（陡面组 + 从缓面组删除的垂直面）")
    print("-" * 50)
    
    zlevel_components_all = []
    zlevel_group_tools_all = []
    zlevel_group_tool_params_all = []
    
    # 1. 添加陡面组
    if steep_components:
        print(f"    添加 {len(steep_components)} 个陡面组")
        for idx, comp in enumerate(steep_components, start=1):
            max_angle = get_max_angle_for_component(comp, angle_map)
            is_vertical = is_component_vertical(comp, angle_map)
            
            if is_vertical:
                # 检查垂直面组是否有R角（使用Face Data - Radius列）
                min_radius = get_min_radius_for_vertical_component(comp, face_df)
                if min_radius is not None and min_radius > 0:
                    # 有R角，根据Radius选刀
                    tool_name = select_tool_for_vertical_with_radius(tools_df, min_radius)
                    print(f"    陡面组 {idx}: {len(comp)} 个面, 垂直面组(有R角) -> 选择刀具 {tool_name}")
                else:
                    # 无R角，使用默认D10
                    tool_name = CONFIG["VERTICAL_FIXED_TOOL"]
                    print(f"    陡面组 {idx}: {len(comp)} 个面, 垂直面组(无R角) -> 固定刀具 {tool_name}")
            else:
                print(f"    陡面组 {idx}: {len(comp)} 个面, 最大角度={max_angle:.2f}°" if max_angle else f"    陡面组 {idx}: {len(comp)} 个面, 无有效角度")
                tool_name = select_tool_by_angle_and_category(tools_df, max_angle, CONFIG["DEFAULT_TOOL"])
            
            zlevel_components_all.append(comp)
            zlevel_group_tools_all.append(tool_name)
            tool_params = get_tool_parameters(tools_df, tool_name, material, is_heat_treated)
            zlevel_group_tool_params_all.append(tool_params)
    
    # 2. 将从缓面组删除的垂直面按邻接关系分组并添加
    if removed_vertical_faces:
        print(f"    添加从缓面组删除的 {len(removed_vertical_faces)} 个垂直面")
        removed_vertical_tags = set(removed_vertical_faces)
        # 按颜色+邻接关系分组
        removed_vertical_components = find_components_by_color(face_df, removed_vertical_tags, color_map)
        print(f"    按颜色+邻接分为 {len(removed_vertical_components)} 个组")
        
        for idx, comp in enumerate(removed_vertical_components, start=len(zlevel_components_all) + 1):
            # 检查垂直面组是否有R角（使用Face Data - Radius列）
            min_radius = get_min_radius_for_vertical_component(comp, face_df)
            if min_radius is not None and min_radius > 0:
                # 有R角，根据Radius选刀
                tool_name = select_tool_for_vertical_with_radius(tools_df, min_radius)
                print(f"    删除的垂直面组 {idx}: {len(comp)} 个面(有R角) -> 选择刀具 {tool_name}")
            else:
                # 无R角，使用默认D10
                tool_name = CONFIG["VERTICAL_FIXED_TOOL"]
                print(f"    删除的垂直面组 {idx}: {len(comp)} 个面(无R角) -> 固定刀具 {tool_name}")
            
            zlevel_components_all.append(comp)
            zlevel_group_tools_all.append(tool_name)
            tool_params = get_tool_parameters(tools_df, tool_name, material, is_heat_treated)
            zlevel_group_tool_params_all.append(tool_params)
    
    # 生成全精_往复等高 JSON
    if zlevel_components_all:
        zlevel_banpamian_json = generate_banpamian_zero_stock_json(
            components=zlevel_components_all,
            group_tools=zlevel_group_tools_all,
            group_tool_params=zlevel_group_tool_params_all,
            direction_map=direction_map,
            fixed_params=CONFIG["BANPAMIAN_FIXED"],
            angle_map=angle_map
        )
        # 按刀具+方向+切深重新分组（全精_往复等高需要按切深分组）
        try:
            zlevel_banpamian_json = regroup_by_tool_and_direction(zlevel_banpamian_json, direction_map, operation_prefix="往复等高", include_cut_depth=True)
            print(f"\n    全精_往复等高重新分组后工序数: {len(zlevel_banpamian_json)}")
        except Exception as e:
            print(f"[WARN] 全精_往复等高重新分组失败: {e}")
            import traceback
            traceback.print_exc()
        
        # 添加邻接的平面（包括顶部和底部）
        try:
            zlevel_banpamian_json = add_adjacent_planes_to_json(zlevel_banpamian_json, face_df, angle_map)
        except Exception as e:
            print(f"[WARN] 添加邻接平面失败: {e}")
            import traceback
            traceback.print_exc()
        
        save_json(zlevel_banpamian_json, CONFIG["BANPAMIAN_VERTICAL_JSON_PATH"])
    else:
        print("[WARN] 无符合条件的往复等高面组，未生成全精_往复等高 JSON")

    # -------------------------------------------------------------------------
    # Step 10: 生成全精_清根 JSON（基于爬面组的Rad Data判断）
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [10] 生成全精_清根 JSON（基于爬面组的Rad Data判断）")
    print("-" * 50)
    
    # Step 10.1: 构建 face_tag -> rad_data 映射
    rad_data_map = {}
    rad_data_col = "Face Data - Rad Data"
    if rad_data_col in face_df.columns:
        for _, row in face_df.iterrows():
            tag = row["Face Tag"]
            rad_data_val = row.get(rad_data_col)
            if pd.notna(rad_data_val):
                try:
                    rad_data_map[tag] = float(rad_data_val)
                except (ValueError, TypeError):
                    pass
        print(f"[INFO] 读取Rad Data映射: {len(rad_data_map)} 个面")
    else:
        print(f"[WARN] 未找到列 '{rad_data_col}'，跳过清根JSON生成")
    
    # Step 10.2: 遍历全精_爬面的每个组，找出需要清根的面
    clearing_faces = []  # 需要清根的面列表
    source_tools = set()  # 来源爬面组的刀具集合
    
    # 检查爬面组变量是否存在
    has_pamian_components = 'components_after_height' in locals() and components_after_height
    has_pamian_tools = 'pamian_group_tools' in locals() and pamian_group_tools
    
    if rad_data_map and has_pamian_components and has_pamian_tools and len(components_after_height) == len(pamian_group_tools):
        print(f"\n    检查 {len(components_after_height)} 个爬面组...")
        
        for group_idx, (comp, group_tool) in enumerate(zip(components_after_height, pamian_group_tools), start=1):
            # 从刀具名称提取R角
            tool_r_angle = extract_r_angle_from_tool_name(group_tool)
            print(f"    爬面组 {group_idx}: 刀具={group_tool}, R角={tool_r_angle}")
            
            if tool_r_angle <= 0:
                print(f"      刀具无R角，跳过此组")
                continue
            
            # 检查组内每个面
            group_clearing_faces = []
            for face_tag in comp:
                rad_data = rad_data_map.get(face_tag, 0)
                
                # 排除 Rad Data = 0 的面
                if rad_data == 0:
                    continue
                
                # 如果 rad_data < tool_r_angle，说明刀进不去，需要清根
                if rad_data < tool_r_angle:
                    group_clearing_faces.append(face_tag)
            
            if group_clearing_faces:
                print(f"      发现 {len(group_clearing_faces)} 个需清根的面 (Rad Data < {tool_r_angle})")
                clearing_faces.extend(group_clearing_faces)
                source_tools.add(group_tool)
    
    # Step 10.3: 如果有需要清根的面，生成清根JSON
    if clearing_faces:
        print(f"\n    共 {len(clearing_faces)} 个面需要清根")
        
        # 将清根面按邻接关系分组
        clearing_tags = set(clearing_faces)
        clearing_components = find_components_by_color(face_df, clearing_tags, color_map)
        print(f"    按颜色+邻接分为 {len(clearing_components)} 个清根组")
        
        # 对每个组找到最底层面，并根据最底层面的Rad Data选刀
        groups_with_tools = []  # [(comp, bottom_face_tag, rad_data, tool_name)]
        
        for comp_idx, comp in enumerate(clearing_components, start=1):
            print(f"\n    清根组 {comp_idx}: {len(comp)} 个面")
            
            # 找到最底层的面
            bottom_face = find_bottom_face_in_group(comp, face_df)
            
            if bottom_face is None:
                print(f"      [WARN] 未找到最底层面，跳过此组")
                continue
            
            # 获取最底层面的Rad Data
            bottom_rad_data = rad_data_map.get(bottom_face, 0)
            
            if bottom_rad_data <= 0:
                print(f"      [WARN] 最底层面(Tag={bottom_face})的Rad Data={bottom_rad_data}，跳过此组")
                continue
            
            print(f"      最底层面: Tag={bottom_face}, Rad Data={bottom_rad_data}")
            
            # 根据最底层面的Rad Data选择清根刀具
            clearing_tool = select_clearing_tool_from_ball_mills(tools_df, bottom_rad_data)
            
            if clearing_tool:
                groups_with_tools.append((comp, bottom_face, bottom_rad_data, clearing_tool))
                print(f"      选择刀具: {clearing_tool}")
                print(f"      组内保留 {len(comp)} 个面用于生成刀轨")
            else:
                print(f"      [WARN] 未能选择合适的清根刀具")
        
        # 生成清根JSON
        if groups_with_tools:
            print(f"\n    生成清根JSON: {len(groups_with_tools)} 个清根工序")
            
            # 确定参考刀具（来源刀具中直径最大的）
            reference_tool = get_max_diameter_tool(list(source_tools), tools_df)
            print(f"    参考刀具: {reference_tool} (来源刀具中直径最大)")
            
            # 生成清根JSON（每个组一个工序，包含组内所有面）
            clearing_json = {}
            for idx, (comp, bottom_face, rad_data, tool_name) in enumerate(groups_with_tools, start=1):
                op_name = f"清根{idx}"
                
                # 获取该刀具的参数
                tool_params = get_tool_parameters(tools_df, tool_name, material, is_heat_treated)
                
                clearing_json[op_name] = {
                    "工序": "清根_SIMPLE",
                    "面ID列表": comp,  # 包含整个组的所有面
                    "刀具名称": tool_name,
                    "刀具类别": tool_params.get('category', ''),
                    "部件侧面余量": 0.0,
                    "部件底面余量": 0.0,
                    "切深": tool_params['cut_depth'],
                    "指定图层": get_layer_for_component(comp, direction_map, f"清根{idx}"),
                    "参考刀具": reference_tool if reference_tool else "NULL",
                    "进给": tool_params['feed'],
                    "转速": tool_params['spindle_rpm'],
                    "横越": tool_params['traverse'],
                }
                print(f"    {op_name}: 包含 {len(comp)} 个面, 基于最底层面Tag={bottom_face}, Rad Data={rad_data}, 刀具={tool_name}")
            
            # 添加邻接的平面（包括顶部和底部，与全精_往复等高一样）
            try:
                clearing_json = add_adjacent_planes_to_json(clearing_json, face_df, angle_map)
            except Exception as e:
                print(f"[WARN] 清根添加邻接平面失败: {e}")
                import traceback
                traceback.print_exc()
            
            # 保存清根JSON
            save_json(clearing_json, CONFIG["CLEARING_JSON_PATH"])
            print(f"    清根JSON工序数: {len(clearing_json)}, 共包含面数: {sum(len(op['面ID列表']) for op in clearing_json.values())}")
        else:
            print("[WARN] 未能选择合适的清根刀具，跳过清根JSON生成")
    else:
        print("[INFO] 无需要清根的面，不生成清根JSON")

    # -------------------------------------------------------------------------
    # 完成
    # -------------------------------------------------------------------------
    print("\n" + "=" * 70)
    print(f"  零件 {part_name} 处理完成！")
    print(f"  材质: {material}, 热处理: {is_heat_treated}")
    print("=" * 70)
    
    return True


def main1(prt_folder, feature_log_csv, face_data_csv, direction_csv, tool_json, output_dir):
    r"""
    主入口函数 - 处理单个零件
    
    Args:
        prt_folder: PRT文件路径
        feature_log_csv: 特征日志CSV路径
        face_data_csv: 面数据CSV路径
        direction_csv: 方向CSV路径
        tool_json: 刀具参数JSON路径
        output_dir: 输出目录
    """
    part_code = extract_part_name_from_path(prt_folder)

    output_banpamian_json = os.path.join(output_dir, f"{part_code}_半精_爬面.json")
    output_vertical_json = os.path.join(output_dir, f"{part_code}_全精_往复等高.json")
    output_pamian_json = os.path.join(output_dir, f"{part_code}_全精_爬面.json")
    output_clearing_json = os.path.join(output_dir, f"{part_code}_全精_清根.json")
    
    # 配置所有路径
    CONFIG["PART_PATH"] = prt_folder
    CONFIG["FEATURE_LOG_CSV"] = feature_log_csv
    CONFIG["FACE_DATA_CSV"] = face_data_csv
    CONFIG["DIRECTION_CSV_PATH"] = direction_csv
    CONFIG["TOOL_JSON_PATH"] = tool_json
    CONFIG["BANPAMIAN_JSON_PATH"] = output_banpamian_json
    CONFIG["BANPAMIAN_VERTICAL_JSON_PATH"] = output_vertical_json
    CONFIG["PAMIAN_JSON_PATH"] = output_pamian_json
    CONFIG["CLEARING_JSON_PATH"] = output_clearing_json
    
    # 处理零件
    process_single_part(part_code)


def main():
    
    prt_folder = r"D:\Projects\NC\output\03_Analysis\Face_Info\prt\UP-12.prt"
    feature_log_csv = rf'D:\Projects\NC\output\03_Analysis\Navigator_Reports\UP-12_FeatureRecognition_Log.csv'
    face_data_csv = rf"D:\Projects\NC\output\03_Analysis\Face_Info\face_csv\UP-12_face_data.csv"
    direction_csv = rf'D:\Projects\NC\output\03_Analysis\Geometry_Analysis\UP-12.csv'
    tool_json = r"D:\Projects\NC\input\铣刀参数.json"
    output_dir = r"C:\Projects\NC\file\toolpath_json"
    
    main1(prt_folder, feature_log_csv, face_data_csv, direction_csv, tool_json, output_dir)


if __name__ == "__main__":
    main()