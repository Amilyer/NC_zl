#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
生成开粗_往复等高 JSON文件
只处理颜色为150的垂直面，使用垂直面组的选刀规则
"""

import json
import os
import pandas as pd
from typing import Dict, List, Tuple, Set, Any

# =============================
# 配置参数
# =============================
CONFIG = {
    # 颜色过滤
    "TARGET_COLOR": 150,  # 只保留150号色的面（垂直面）
    
    # 固定参数
    "VERTICAL_FIXED_TOOL": "D10",  # 垂直面无R角时的固定刀具
    "DEFAULT_TOOL": "D12R0.5",  # 默认刀具
    "TOOL_CATEGORY_FLAT": "钨钢平刀",  # 平刀类别
    
    # 开粗往复等高固定参数
    "KAICU_FIXED": {
        "部件侧面余量": 0.3,
        "部件底面余量": 0.05
    },
    
    # 文件路径（将在main1中设置）
    "PART_PATH": "",
    "FACE_DATA_CSV": "",
    "DIRECTION_CSV_PATH": "",
    "TOOL_JSON_PATH": "",
    "KAICU_JSON_PATH": ""
}

# =============================
# 文件读取和保存
# =============================
def read_csv(filepath: str) -> pd.DataFrame:
    """读取CSV文件"""
    if not os.path.exists(filepath):
        print(f"[ERROR] 文件不存在: {filepath}")
        return pd.DataFrame()
    try:
        df = pd.read_csv(filepath, encoding='utf-8')
        return df
    except UnicodeDecodeError:
        try:
            df = pd.read_csv(filepath, encoding='gbk')
            return df
        except Exception as e:
            print(f"[ERROR] 读取CSV失败: {e}")
            return pd.DataFrame()

def read_json(filepath: str) -> dict:
    """读取JSON文件"""
    if not os.path.exists(filepath):
        print(f"[ERROR] 文件不存在: {filepath}")
        return {}
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"[ERROR] 读取JSON失败: {e}")
        return {}

def save_json(data: dict, filepath: str):
    """保存JSON文件"""
    try:
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"[INFO] 成功保存: {filepath}")
    except Exception as e:
        print(f"[ERROR] 保存JSON失败: {e}")

# =============================
# 基础数据处理
# =============================
def extract_part_name_from_path(path: str) -> str:
    """从路径中提取零件名"""
    basename = os.path.basename(path)
    name = os.path.splitext(basename)[0]
    return name

def parse_material_info(part_path: str) -> Tuple[str, bool]:
    """解析材质信息"""
    part_name = extract_part_name_from_path(part_path).upper()
    
    # 判断材质 - 简化版，实际应从参数表读取
    if part_name.startswith("GU") or part_name.startswith("UP") or part_name.startswith("DIE"):
        material = "45#"  # 使用具体材质名称
    else:
        material = "45#"  # 默认材质
    
    # 判断热处理（简化处理，实际可能需要从其他数据源获取）
    is_heat_treated = False  # 默认未热处理
    
    return material, is_heat_treated

def load_tools_data(tool_json_path: str) -> pd.DataFrame:
    """加载刀具参数"""
    try:
        with open(tool_json_path, 'r', encoding='utf-8') as f:
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

# =============================
# 面分组相关函数
# =============================
def build_adjacency_map(face_df: pd.DataFrame) -> Dict[str, Set[str]]:
    """构建面的邻接映射"""
    adjacency_map = {}
    
    # 解析邻接关系
    for _, row in face_df.iterrows():
        face_tag = str(row['Face Tag'])
        adjacent_str = row.get('Adjacent Face Tags', '')
        
        if pd.notna(adjacent_str) and adjacent_str and str(adjacent_str) != 'nan':
            # 初始化该面的邻接集合
            if face_tag not in adjacency_map:
                adjacency_map[face_tag] = set()
            
            # 解析邻接面列表（支持分号分隔）
            adjacent_faces = [s.strip() for s in str(adjacent_str).split(';') if s.strip()]
            adjacency_map[face_tag].update(adjacent_faces)
            
            # 建立双向邻接关系
            for adj_face in adjacent_faces:
                if adj_face not in adjacency_map:
                    adjacency_map[adj_face] = set()
                adjacency_map[adj_face].add(face_tag)
    
    return adjacency_map

def find_connected_components(face_tags: Set[str], adjacency_map: Dict[str, Set[str]]) -> List[List[str]]:
    """根据邻接关系将面分组（使用深度优先搜索）"""
    components = []
    visited = set()
    
    def dfs(face_tag: str, component: List[str]):
        """深度优先搜索找连通分量"""
        if face_tag in visited or face_tag not in face_tags:
            return
        
        visited.add(face_tag)
        component.append(face_tag)
        
        # 访问所有邻接的面
        if face_tag in adjacency_map:
            for adj_face in adjacency_map[face_tag]:
                if adj_face in face_tags and adj_face not in visited:
                    dfs(adj_face, component)
    
    # 对每个未访问的面开始DFS
    for face_tag in face_tags:
        if face_tag not in visited:
            component = []
            dfs(face_tag, component)
            if component:
                components.append(component)
    
    return components

# =============================
# R角相关函数
# =============================
def get_min_radius_for_vertical_component(component: list, face_df: pd.DataFrame) -> float:
    """
    获取垂直面组内所有面的最小非零Radius值（专门用于垂直面组的R角判断）
    使用 Face Data - Radius 列
    
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
        # 转换为字符串进行比较
        row = face_df[face_df["Face Tag"].astype(str) == str(face_tag)]
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

# =============================
# 刀具选择函数
# =============================
def select_tool_for_vertical_with_radius(tools_df: pd.DataFrame, min_radius: float) -> str:
    """
    根据垂直面组的最小R角选择刀具
    规则：选择直径 < min_radius * 2 的最大平刀
    
    Args:
        tools_df: 刀具参数表
        min_radius: 组内最小非零Radius
        
    Returns:
        选中的刀具名称
    """
    if tools_df.empty or min_radius is None or min_radius <= 0:
        return CONFIG["VERTICAL_FIXED_TOOL"]  # 默认返回D10
    
    # 计算最大允许直径
    max_allowed_diameter = min_radius * 2
    
    # 筛选钨钢平刀类别
    flat_tools = tools_df[tools_df['类别'] == CONFIG["TOOL_CATEGORY_FLAT"]].copy()
    
    if flat_tools.empty:
        print(f"[WARN] 未找到钨钢平刀类别的刀具，使用默认D10")
        return CONFIG["VERTICAL_FIXED_TOOL"]
    
    # 筛选直径 < max_allowed_diameter 的刀具
    valid_tools = flat_tools[flat_tools['直径'] < max_allowed_diameter].copy()
    
    if valid_tools.empty:
        # 没有满足条件的刀具，选择该类别中直径最小的
        min_diameter_row = flat_tools.loc[flat_tools['直径'].idxmin()]
        tool_name = min_diameter_row['刀具名称']
        print(f"[WARN] 钨钢平刀中无直径 < {max_allowed_diameter:.1f}mm 的刀具，选择最小直径: {tool_name} (直径={min_diameter_row['直径']}mm)")
        return tool_name
    
    # 选择满足条件的最大刀具
    valid_tools = valid_tools.sort_values('直径', ascending=False)
    selected_row = valid_tools.iloc[0]
    tool_name = selected_row['刀具名称']
    
    print(f"[INFO] 垂直面组有R角: 最小Radius={min_radius:.2f}, 最大允许直径={max_allowed_diameter:.1f}mm, 选择 '{tool_name}' (直径={selected_row['直径']}mm)")
    return tool_name

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

def get_tool_parameters(tools_df: pd.DataFrame, tool_name: str, material: str, is_heat_treated: bool) -> dict:
    """根据刀具名称、材质和热处理状态获取切深、进给、转速、横越、刀具类别"""
    # 默认参数
    result = {
        'cut_depth': 0.5,
        'feed': 0.0,
        'spindle_rpm': 6000.0,
        'traverse': 0.0,
        'category': "",
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
    
    # 根据材质和热处理状态选择切深列
    material_upper = material.upper() if material else "45#"
    
    if is_heat_treated and material_upper in MATERIAL_COLUMN_MAP_HEAT_TREATED:
        column_name = MATERIAL_COLUMN_MAP_HEAT_TREATED[material_upper]
    elif material_upper in MATERIAL_COLUMN_MAP:
        column_name = MATERIAL_COLUMN_MAP[material_upper]
    else:
        column_name = '45#,A3,切深'
    
    # 获取切深
    if column_name in tool_row.index:
        depth = tool_row[column_name]
        if pd.notna(depth):
            result['cut_depth'] = float(depth)
    
    # 获取转速、进给、横越（使用"普"列）
    if '转速(普)' in tool_row.index and pd.notna(tool_row['转速(普)']):
        result['spindle_rpm'] = float(tool_row['转速(普)'])
    
    if '进给(普)' in tool_row.index and pd.notna(tool_row['进给(普)']):
        result['feed'] = float(tool_row['进给(普)'])
    
    if '横越(普)' in tool_row.index and pd.notna(tool_row['横越(普)']):
        result['traverse'] = float(tool_row['横越(普)'])
    
    return result

# =============================
# 方向和图层处理
# =============================
def load_direction_data(csv_path: str) -> Dict[str, str]:
    """加载方向数据（从CSV文件的列名和Face Tag映射）"""
    try:
        df = read_csv(csv_path)
        if df.empty:
            return {}
        
        direction_map = {}
        # CSV文件的每一列代表一个方向，列名是方向（如+Z, -Z等）
        # 列中的值是该方向的Face Tag
        for column in df.columns:
            col_name = column.strip()
            for face_tag in df[column].dropna():
                try:
                    # 将Face Tag映射到方向
                    direction_map[int(face_tag)] = col_name
                except (ValueError, TypeError):
                    continue
        
        print(f"[INFO] 成功读取方向映射，共 {len(direction_map)} 个面标签")
        return direction_map
    except Exception as e:
        print(f"[WARN] 读取方向映射文件时出错: {e}")
        return {}

def get_layer_by_direction(direction: str) -> int:
    """根据方向返回对应的图层号"""
    layer_map = {
        '+Z': 20, '-Z': 70,
        '+X': 40, '-X': 30,
        '+Y': 60, '-Y': 50
    }
    # 默认图层70（-Z方向）
    return layer_map.get(direction.strip(), 70)

def get_layer_for_component(component: list, direction_map: dict, default_layer: str = "") -> int:
    """获取组的图层号（整数）"""
    if not direction_map:
        return 70  # 默认图层
    
    # 查找组内第一个有方向映射的面
    for face_tag in component:
        # 尝试将face_tag转换为整数进行查找
        try:
            tag_int = int(face_tag)
            if tag_int in direction_map:
                direction = direction_map[tag_int]
                return get_layer_by_direction(direction)
        except (ValueError, TypeError):
            pass
    
    return 70  # 默认图层

# =============================
# JSON生成函数
# =============================
def generate_kaicu_json(components: List[List[str]], 
                        group_tools: List[str],
                        group_tool_params: List[dict],
                        direction_map: dict,
                        fixed_params: dict) -> dict:
    """生成开粗往复等高JSON"""
    json_data = {}
    
    for idx, (comp, tool_name, tool_params) in enumerate(zip(components, group_tools, group_tool_params), start=1):
        op_name = f"开粗往复等高{idx}"
        
        json_data[op_name] = {
            "工序": "往复等高_SIMPLE",
            "面ID列表": comp,
            "刀具名称": tool_name,
            "刀具类别": tool_params.get('category', ''),
            "部件侧面余量": fixed_params["部件侧面余量"],
            "部件底面余量": fixed_params["部件底面余量"],
            "切深": tool_params['cut_depth'],
            "指定图层": get_layer_for_component(comp, direction_map),
            "进给": tool_params['feed'],
            "转速": tool_params['spindle_rpm'],
            "横越": tool_params['traverse']
        }
    
    return json_data

def regroup_by_tool_and_direction(json_data: dict, direction_map: dict, operation_prefix: str = "开粗往复等高") -> dict:
    """按刀具和方向重新分组工序"""
    if not json_data:
        return json_data
    
    # 按刀具+方向+切深分组
    groups = {}
    
    for op_name, op_data in json_data.items():
        tool_name = op_data["刀具名称"]
        layer = op_data["指定图层"]
        cut_depth = op_data.get("切深", 1.0)
        
        # 生成分组键
        group_key = (tool_name, layer, cut_depth)
        
        if group_key not in groups:
            groups[group_key] = {
                "faces": [],
                "data": op_data.copy()
            }
        
        # 合并面ID列表
        groups[group_key]["faces"].extend(op_data["面ID列表"])
    
    # 生成新的JSON
    new_json = {}
    for idx, ((tool_name, layer, cut_depth), group) in enumerate(groups.items(), start=1):
        op_name = f"往复等高{idx}"
        
        group["data"]["面ID列表"] = group["faces"]
        new_json[op_name] = group["data"]
    
    return new_json

# =============================
# 主处理函数
# =============================
def process_single_part(part_name: str) -> bool:
    """处理单个零件，生成开粗往复等高JSON"""
    
    print("=" * 70)
    print(f"  开始处理零件: {part_name}")
    print("=" * 70)
    
    # -------------------------------------------------------------------------
    # Step 1: 读取基础数据
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [1] 读取基础数据")
    print("-" * 50)
    
    # 解析材质信息
    material, is_heat_treated = parse_material_info(CONFIG["PART_PATH"])
    print(f"[INFO] 材质: {material}, 热处理: {is_heat_treated}")
    
    # 读取面数据
    face_df = read_csv(CONFIG["FACE_DATA_CSV"])
    if face_df.empty:
        print("[ERROR] 面数据为空")
        return False
    print(f"[INFO] 读取 {len(face_df)} 个面")
    
    # 读取刀具参数
    tools_df = load_tools_data(CONFIG["TOOL_JSON_PATH"])
    if tools_df.empty:
        print("[ERROR] 刀具参数为空")
        return False
    print(f"[INFO] 读取 {len(tools_df)} 个刀具参数")
    
    # 读取方向数据
    direction_map = load_direction_data(CONFIG["DIRECTION_CSV_PATH"])
    print(f"[INFO] 读取 {len(direction_map)} 个面的方向数据")
    
    # -------------------------------------------------------------------------
    # Step 2: 筛选150号色的面
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [2] 筛选150号色的面（垂直面）")
    print("-" * 50)
    
    # 构建颜色映射
    color_map = {}
    if 'Face Color' in face_df.columns:
        for _, row in face_df.iterrows():
            tag = str(row['Face Tag'])  # 确保是字符串
            color = row.get('Face Color')
            if pd.notna(color):
                try:
                    color_map[tag] = int(color)
                except (ValueError, TypeError):
                    pass
    
    print(f"[INFO] 读取颜色映射: {len(color_map)} 个面")
    
    # 筛选150号色的面
    target_faces = []
    for tag, color in color_map.items():
        if color == CONFIG["TARGET_COLOR"]:
            target_faces.append(tag)
    
    print(f"[INFO] 找到 {len(target_faces)} 个150号色的面")
    
    if not target_faces:
        print("[WARN] 未找到150号色的面，无法生成开粗往复等高JSON")
        return False
    
    # -------------------------------------------------------------------------
    # Step 3: 按邻接关系分组
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [3] 按邻接关系分组")
    print("-" * 50)
    
    # 构建邻接映射
    adjacency_map = build_adjacency_map(face_df)
    
    # 按邻接关系分组
    target_face_set = set(target_faces)
    components = find_connected_components(target_face_set, adjacency_map)
    
    print(f"[INFO] 分为 {len(components)} 个组")
    
    # -------------------------------------------------------------------------
    # Step 4: 对每组选择刀具
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [4] 对每组选择刀具")
    print("-" * 50)
    
    group_tools = []
    group_tool_params = []
    
    for idx, comp in enumerate(components, start=1):
        # 检查垂直面组是否有R角
        min_radius = get_min_radius_for_vertical_component(comp, face_df)
        
        if min_radius is not None and min_radius > 0:
            # 有R角，根据Radius选刀
            tool_name = select_tool_for_vertical_with_radius(tools_df, min_radius)
            print(f"    组 {idx}: {len(comp)} 个面, 垂直面组(有R角) -> 选择刀具 {tool_name}")
        else:
            # 无R角，使用默认D10
            tool_name = CONFIG["VERTICAL_FIXED_TOOL"]
            print(f"    组 {idx}: {len(comp)} 个面, 垂直面组(无R角) -> 固定刀具 {tool_name}")
        
        group_tools.append(tool_name)
        
        # 获取刀具参数
        tool_params = get_tool_parameters(tools_df, tool_name, material, is_heat_treated)
        group_tool_params.append(tool_params)
    
    # -------------------------------------------------------------------------
    # Step 5: 生成开粗往复等高JSON
    # -------------------------------------------------------------------------
    print("\n" + "-" * 50)
    print("  [5] 生成开粗往复等高JSON")
    print("-" * 50)
    
    kaicu_json = generate_kaicu_json(
        components=components,
        group_tools=group_tools,
        group_tool_params=group_tool_params,
        direction_map=direction_map,
        fixed_params=CONFIG["KAICU_FIXED"]
    )
    
    print(f"[INFO] 生成 {len(kaicu_json)} 个工序")
    
    # 按刀具+方向+切深重新分组
    try:
        kaicu_json = regroup_by_tool_and_direction(kaicu_json, direction_map, operation_prefix="开粗往复等高")
        print(f"[INFO] 重新分组后工序数: {len(kaicu_json)}")
    except Exception as e:
        print(f"[WARN] 重新分组失败: {e}")
        import traceback
        traceback.print_exc()
    
    # 保存JSON
    save_json(kaicu_json, CONFIG["KAICU_JSON_PATH"])
    
    # -------------------------------------------------------------------------
    # 完成
    # -------------------------------------------------------------------------
    print("\n" + "=" * 70)
    print(f"  零件 {part_name} 处理完成！")
    print(f"  材质: {material}, 热处理: {is_heat_treated}")
    print("=" * 70)
    
    return True

# =============================
# 主入口函数
# =============================
def main1(prt_folder, feature_log_csv, face_data_csv, direction_csv, tool_json, output_dir):
    """
    主入口函数 - 处理单个零件
    
    Args:
        prt_folder: PRT文件路径
        feature_log_csv: 特征日志CSV路径（未使用）
        face_data_csv: 面数据CSV路径
        direction_csv: 方向CSV路径
        tool_json: 刀具参数JSON路径
        output_dir: 输出目录
    """
    part_code = extract_part_name_from_path(prt_folder)
    
    output_kaicu_json = os.path.join(output_dir, f"{part_code}_开粗_往复等高.json")
    
    # 配置所有路径
    CONFIG["PART_PATH"] = prt_folder
    CONFIG["FACE_DATA_CSV"] = face_data_csv
    CONFIG["DIRECTION_CSV_PATH"] = direction_csv
    CONFIG["TOOL_JSON_PATH"] = tool_json
    CONFIG["KAICU_JSON_PATH"] = output_kaicu_json
    
    # 处理零件
    process_single_part(part_code)


def main():
    
    prt_folder = r"D:\Projects\NC\output\03_Analysis\Face_Info\prt\B2-01.prt"
    feature_log_csv = rf'D:\Projects\NC\output\03_Analysis\Navigator_Reports\B2-01_FeatureRecognition_Log.csv'
    face_data_csv = rf"D:\Projects\NC\output\03_Analysis\Face_Info\face_csv\B2-01_face_data.csv"
    direction_csv = rf'D:\Projects\NC\output\03_Analysis\Geometry_Analysis\B2-01.csv'
    tool_json = r"D:\Projects\NC\input\铣刀参数.json"
    output_dir = r"C:\Projects\NC\file\toolpath_json"
    
    main1(prt_folder, feature_log_csv, face_data_csv, direction_csv, tool_json, output_dir)


if __name__ == "__main__":
    main()
