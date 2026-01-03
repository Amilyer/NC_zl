#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
清角JSON生成脚本（最终完美版）
功能：
1. 支持多文件输入，自动过滤空路径和无效路径。
2. 动态生成键名（往复等高1, 往复等高2...），根据有效文件数量自动排序。
3. 零件参数Excel匹配列修正为【文件名称】。
4. 智能切深匹配（材质+热处理字符串包含匹配）。
"""

import os
import json
import pandas as pd
import numpy as np
from pathlib import Path
from typing import List, Dict, Tuple, Optional

# ==================================================================================
# 配置区域
# ==================================================================================

CONFIG = {
    "DEFAULT_TOOL_CATEGORY": "钨钢平刀",
    "DEFAULT_CUT_DEPTH": 0.25,
    "DEFAULT_FEED": 2000.0,
    "DEFAULT_SPINDLE_RPM": 3000.0,
    "DEFAULT_TRAVERSE": 8000.0,
    "DEFAULT_LAYER": 70,
    "DEFAULT_WALL_ALLOWANCE": 0.0,
    "DEFAULT_FLOOR_ALLOWANCE": 0.0,
}

# ==================================================================================
# 核心类
# ==================================================================================

class CornerCleaningProcessor:
    """清角处理器类，封装所有相关功能"""
    
    def __init__(self):
        """初始化处理器"""
        self._material_df = None  # 零件参数表缓存
        self.material = None
        self.is_heat_treated = False
        
    def filter_valid_files(self, file_list: list) -> list:
        """
        清洗文件列表：
        1. 过滤空字符串或None
        2. 过滤不存在的路径
        """
        valid_files = []
        print("-" * 30)
        print("正在检查输入文件有效性...")
        
        for f in file_list:
            # 1. 检查是否为空字符串
            if not f or not isinstance(f, str) or not f.strip():
                print(f"[SKIP] 跳过空路径或无效格式")
                continue
                
            # 2. 检查文件是否存在
            clean_path = f.strip()
            if os.path.exists(clean_path):
                valid_files.append(clean_path)
                print(f"[OK] 文件存在: {os.path.basename(clean_path)}")
            else:
                print(f"[SKIP] 文件不存在: {clean_path}")
                
        print(f"有效文件数量: {len(valid_files)}")
        print("-" * 30)
        return valid_files
    
    def _load_material_excel(self, excel_path: str | Path) -> pd.DataFrame:
        """延迟加载零件参数表，只加载一次"""
        if self._material_df is None:
            try:
                df = pd.read_excel(excel_path)
                print(f"成功加载零件参数表，共 {len(df)} 行")
                self._material_df = df
            except Exception as e:
                print(f"读取零件参数表失败: {e}")
                self._material_df = pd.DataFrame()
        return self._material_df
    
    def load_material_info(self, part_name_key: str, excel_path: str | Path):
        """
        从Excel读取材质 (匹配【文件名称】列)
        
        Args:
            part_name_key: 零件名称关键字
            excel_path: Excel文件路径
        """
        if not os.path.exists(excel_path):
            print(f"[WARN] 零件参数Excel不存在: {excel_path}")
            self.material = None
            self.is_heat_treated = False
            return
        
        try:
            df = self._load_material_excel(excel_path)
            
            if '文件名称' not in df.columns:
                print(f"[ERROR] Excel中缺少【文件名称】列")
                self.material = None
                self.is_heat_treated = False
                return
            
            # 匹配逻辑 (精确匹配 + 去后缀匹配)
            row = df[df['文件名称'] == part_name_key]
            if row.empty:
                clean_names = df['文件名称'].astype(str).apply(lambda x: os.path.splitext(x)[0])
                row = df[clean_names == part_name_key]
            
            if row.empty:
                print(f"[WARN] Excel中未找到零件: {part_name_key}")
                self.material = None
                self.is_heat_treated = False
                return
            
            material = str(row.iloc[0]['材质']).strip()
            ht_val = row.iloc[0]['热处理']
            
            is_heat_treated = False
            if pd.notna(ht_val) and str(ht_val).strip() != "":
                is_heat_treated = True
                
            self.material = material
            self.is_heat_treated = is_heat_treated
            
            print(f"[INFO] 零件材质信息 -> 材质: {material} | 热处理: {'YES' if is_heat_treated else 'NO'}")
            
        except Exception as e:
            print(f"[ERROR] 读取零件Excel出错: {e}")
            self.material = None
            self.is_heat_treated = False
    
    def load_direction_data(self, csv_path: str) -> Dict[int, str]:
        try:
            if not csv_path or not os.path.exists(csv_path):
                return {}

            try:
                df = pd.read_csv(csv_path, encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(csv_path, encoding='gbk')

            if df.empty:
                return {}

            direction_map: Dict[int, str] = {}
            for column in df.columns:
                col_name = str(column).strip()
                for face_tag in df[column].dropna():
                    try:
                        direction_map[int(face_tag)] = col_name
                    except (ValueError, TypeError):
                        continue

            return direction_map
        except Exception as e:
            print(f"[WARN] 读取方向映射文件时出错: {e}")
            return {}

    def get_layer_by_direction(self, direction: str) -> int:
        layer_map = {
            '+Z': 20, '-Z': 70,
            '+X': 40, '-X': 30,
            '+Y': 60, '-Y': 50
        }
        return layer_map.get(direction.strip(), CONFIG["DEFAULT_LAYER"])

    def get_layer_for_face_ids(self, face_ids: list, direction_map: dict) -> int:
        if not direction_map:
            return CONFIG["DEFAULT_LAYER"]

        for face_tag in face_ids:
            try:
                tag_int = int(face_tag)
                if tag_int in direction_map:
                    direction = direction_map[tag_int]
                    return self.get_layer_by_direction(direction)
            except (ValueError, TypeError):
                continue

        return CONFIG["DEFAULT_LAYER"]
    
    def load_face_data(self, csv_path: str) -> pd.DataFrame:
        """加载face_csv数据，筛选圆柱面(Face Type=16)"""
        try:
            if not os.path.exists(csv_path):
                print(f"[WARN] CSV文件不存在: {csv_path}")
                return pd.DataFrame()
            
            df = pd.read_csv(csv_path)
            cylinder_faces = df[df['Face Type'] == 16].copy()
            print(f"[INFO] 加载圆柱面数据: 共 {len(cylinder_faces)} 个圆柱面")
            return cylinder_faces
        except Exception as e:
            print(f"[WARN] 读取面数据CSV失败: {e}")
            return pd.DataFrame()
    
    def get_min_cylinder_radius_for_faces(self, face_ids: list, cylinder_df: pd.DataFrame) -> float:
        """根据面ID列表获取最小圆柱面半径"""
        if cylinder_df.empty or not face_ids:
            return 5.0
        
        matched = cylinder_df[cylinder_df['Face Tag'].isin(face_ids)]
        if matched.empty:
            return 5.0
        
        valid_radii = matched[matched['Face Data - Radius'] > 0]['Face Data - Radius']
        if valid_radii.empty:
            return 5.0
        
        return float(valid_radii.min())
    
    def merge_operations_by_layer(self, json_files: list, tool_map: dict) -> Dict[int, dict]:
        """
        读取多个JSON文件，按"指定图层"合并分组
        
        Returns:
            dict: {图层号: {"face_ids": [...], "max_tool_name": "xxx", "max_tool_diameter": float}}
        """
        layer_groups: Dict[int, dict] = {}
        
        for json_file in json_files:
            try:
                with open(json_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                
                for op_key, op_val in data.items():
                    layer = op_val.get("指定图层", CONFIG["DEFAULT_LAYER"])
                    face_ids = op_val.get("面ID列表", [])
                    tool_name = op_val.get("刀具名称", "")
                    
                    if layer not in layer_groups:
                        layer_groups[layer] = {
                            "face_ids": [],
                            "max_tool_name": None,
                            "max_tool_diameter": 0
                        }
                    
                    layer_groups[layer]["face_ids"].extend(face_ids)
                    
                    if tool_name and tool_name in tool_map:
                        tool_diameter = tool_map[tool_name].get("直径", 0)
                        if tool_diameter > layer_groups[layer]["max_tool_diameter"]:
                            layer_groups[layer]["max_tool_diameter"] = tool_diameter
                            layer_groups[layer]["max_tool_name"] = tool_name
                            
            except Exception as e:
                print(f"[WARN] 读取JSON文件失败 {json_file}: {e}")
                continue
        
        for layer in layer_groups:
            layer_groups[layer]["face_ids"] = sorted(list(set(layer_groups[layer]["face_ids"])))
        
        return layer_groups
    
    def select_tool_by_radius(self, min_radius: float, tool_list: list) -> dict:
        """
        根据最小圆柱面半径选刀
        规则：刀具直径 ≤ 2*半径，且取满足条件中直径最大的刀
        """
        max_diameter = 2 * min_radius
        flat_tools = [t for t in tool_list if t.get("类别") == "钨钢平刀"]
        
        if not flat_tools:
            print(f"[WARN] 未找到钨钢平刀类别的刀具")
            return None
        
        valid_tools = [t for t in flat_tools if t["直径"] <= max_diameter]
        
        if valid_tools:
            selected = max(valid_tools, key=lambda x: x["直径"])
            print(f"    -> 选刀: {selected['刀具名称']} (直径={selected['直径']}mm, 最大允许={max_diameter}mm)")
            return selected
        
        smallest = min(flat_tools, key=lambda x: x["直径"])
        print(f"    -> [WARN] 无满足条件的刀具，使用最小刀: {smallest['刀具名称']} (直径={smallest['直径']}mm)")
        return smallest
    
    def calculate_cut_parameters(self, tool_c: dict):
        """计算切深、转速、进给"""
        if not tool_c:
            return (CONFIG["DEFAULT_CUT_DEPTH"], CONFIG["DEFAULT_SPINDLE_RPM"], 
                    CONFIG["DEFAULT_FEED"], CONFIG["DEFAULT_TRAVERSE"])

        rpm = tool_c.get("转速(普)", CONFIG["DEFAULT_SPINDLE_RPM"])
        feed = tool_c.get("进给(普)", CONFIG["DEFAULT_FEED"])
        traverse = tool_c.get("横越(普)", CONFIG["DEFAULT_TRAVERSE"])
        cut_depth = CONFIG["DEFAULT_CUT_DEPTH"]
        
        if self.material:
            matched_keys = [k for k in tool_c.keys() if self.material in k]
            target_key = None
            
            if matched_keys:
                if self.is_heat_treated:
                    candidates = [k for k in matched_keys if "热处理后" in k]
                    if candidates:
                        target_key = candidates[0]
                else:
                    candidates = [k for k in matched_keys if "热处理后" not in k]
                    if candidates:
                        pure_keys = [k for k in candidates if "热处理" not in k]
                        target_key = pure_keys[0] if pure_keys else candidates[0]
            
            if target_key:
                cut_depth = tool_c[target_key]
                print(f"    -> [参数] 匹配键: '{target_key}' | 值: {cut_depth}")
                
        return cut_depth, rpm, feed, traverse
    
    def extract_part_name(self, filepath: str) -> str:
        """从文件路径提取零件名"""
        name = os.path.basename(filepath)
        if "_" in name:
            return name.split("_")[0]
        return os.path.splitext(name)[0]
    
    def process_corner_cleaning(self, raw_input_files: list, face_csv: str, 
                               tool_json: str, part_xlsx: str, direction_file: str, output_dir: str):
        """
        完整的清角处理流程
        
        Args:
            raw_input_files: 原始输入文件列表
            face_csv: 面数据CSV路径
            tool_json: 刀具参数JSON路径
            part_xlsx: 零件参数Excel路径
            direction_file: 方向文件路径
            output_dir: 输出目录
        """
        # 1. 过滤无效文件
        valid_inputs = self.filter_valid_files(raw_input_files)
        
        if not valid_inputs:
            print("[ERROR] 没有有效的输入文件，程序终止。")
            return

        # 2. 提取零件号 (使用第一个有效文件)
        part_name_key = self.extract_part_name(valid_inputs[0])
        output_path = os.path.join(output_dir, f"{part_name_key}_半精_清角.json")
        
        print("="*70)
        print(f"  开始生成清角JSON")
        print(f"  零件: {part_name_key}")
        print("="*70)

        # 3. 加载材质信息
        self.load_material_info(part_name_key, part_xlsx)
        
        # 4. 加载刀具库
        try:
            with open(tool_json, "r", encoding="utf-8") as f:
                all_tools_list = json.load(f)
                tool_map = {t["刀具名称"]: t for t in all_tools_list}
        except Exception as e:
            print(f"[FATAL] 无法读取铣刀参数JSON: {e}")
            return

        # 5. 加载圆柱面数据
        cylinder_df = self.load_face_data(face_csv)

        # 6. 按图层合并分组
        print("\n--- 按图层合并分组 ---")
        layer_groups = self.merge_operations_by_layer(valid_inputs, tool_map)
        
        for layer, group in layer_groups.items():
            print(f"  图层 {layer}: {len(group['face_ids'])} 个面ID, 最大刀具: {group['max_tool_name']} (D={group['max_tool_diameter']}mm)")

        # 7. 遍历每个图层组生成清角工序
        final_data = {}
        
        for idx, (layer, group) in enumerate(sorted(layer_groups.items()), 1):
            print(f"\n--- 处理图层 {layer} ---")
            
            face_ids = group["face_ids"]
            ref_tool_name = group["max_tool_name"] if group["max_tool_name"] else "NULL"
            
            # 获取该组最小圆柱面半径
            min_radius = self.get_min_cylinder_radius_for_faces(face_ids, cylinder_df)
            print(f"    -> 最小圆柱面半径: {min_radius}mm")
            
            # 根据半径选刀
            selected_tool = self.select_tool_by_radius(min_radius, all_tools_list)
            
            # 计算切削参数
            cut_depth, rpm, feed, traverse = self.calculate_cut_parameters(selected_tool)
            
            # 动态生成 Key: 清角1, 清角2...
            node_key = f"清角{idx}"
            final_data[node_key] = {
                "工序": "清角_SIMPLE",
                "面ID列表": face_ids,
                "刀具名称": selected_tool["刀具名称"] if selected_tool else "未命名",
                "刀具类别": selected_tool["类别"] if selected_tool else CONFIG["DEFAULT_TOOL_CATEGORY"],
                "部件侧面余量": 0.03,
                "部件底面余量": 0.03,
                "切深": cut_depth,
                "指定图层": layer,
                "参考刀具": ref_tool_name,
                "进给": float(feed),
                "转速": float(rpm),
                "横越": float(traverse)
            }
            
            print(f"    -> 参考刀具: {ref_tool_name}")

        # 8. 保存
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(final_data, f, ensure_ascii=False, indent=2)
        
        print("\n" + "="*70)
        print(f"已保存至: {output_path}")
        print("="*70)

# ==================================================================================
# 主函数
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

def main1(raw_input_files: list, face_csv: str, tool_json: str, 
         part_xlsx: str, direction_file:str,output_dir: str):
    """
    主入口函数 - 处理清角加工
    
    Args:
        raw_input_files: 原始输入文件列表
        face_csv: 面数据CSV路径
        tool_json: 刀具参数JSON路径
        part_xlsx: 零件参数Excel路径
        direction_file: 方向文件路径
        output_dir: 输出目录
    """
    # 创建处理器实例
    processor = CornerCleaningProcessor()
    
    # 执行处理流程
    processor.process_corner_cleaning(
        raw_input_files=raw_input_files,
        face_csv=face_csv,
        tool_json=tool_json,
        part_xlsx=part_xlsx,
        direction_file=direction_file,
        output_dir=output_dir
    )

def main():
    """示例主函数"""
    # 零件名称 - 可以根据需要修改
    part_name = "UP-01"
    
    # 构建输入文件路径
    input_json_files = [
        os.path.join(r'C:\Projects\NC\file\toolpath_json', f"{part_name}_半精_螺旋.json"),
        os.path.join(r'C:\Projects\NC\file\toolpath_json', f"{part_name}_半精_螺旋_往复等高.json")
    ]
    
    # 输入文件路径
    face_csv = os.path.join(r'C:\Projects\NC\file\Face_Info\face_csv', f"{part_name}_face_data.csv")
    tool_json = os.path.join(r'D:\Projects\NC\input\铣刀参数.json')
    part_xlsx = os.path.join(r'D:\Projects\NC\output\00_Resources\CSV_Reports\零件参数2.xlsx')
    direction_file = os.path.join(r'D:\Projects\NC\output\03_Analysis\Geometry_Analysis2', f"{part_name}.prt.csv")
    output_dir = r"C:\Projects\NC\file\toolpath_json"
    
    print(f"零件代码: {part_name}")
    print(f"输出目录: {output_dir}")
    print(f"面数据CSV: {face_csv}")
    print(f"输出文件: {part_name}_半精_清角.json")
    
    # 调用主处理函数
    main1(input_json_files, face_csv, tool_json, part_xlsx, direction_file,output_dir)

# ==================================================================================
# 执行入口
# ==================================================================================

if __name__ == "__main__":
    main()