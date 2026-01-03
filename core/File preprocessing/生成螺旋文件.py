
import pandas as pd
from pathlib import Path
import re
import NXOpen
import json
import os
import numpy as np
from typing import Dict, List, Tuple, Optional
import warnings

class SpiralProcessor:
    """螺旋加工处理器类，封装所有相关功能"""
    
    def __init__(self):
        """初始化处理器"""
        self._material_df = None  # 零件参数表缓存
        self._layer_map = {  # 方向到图层的映射
            '+Z': 20, '-Z': 70,
            '+X': 40, '-X': 30,
            '+Y': 60, '-Y': 50
        }
        self.material = None
        self.is_heat_treated = False
        self.spiral_data = {}
        self.half_spiral_data = {}
        self._face_data_dict = {}  # 新增：缓存面数据，用于判断开放/封闭
        
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
    
    def get_layer_by_direction(self, direction: str) -> int:
        """根据方向返回对应的图层"""
        return self._layer_map.get(direction, 20)
    

    def parse_tool_R_angle(self, tool_name: str) -> float:
        """
        从刀具名称中解析R角大小
        示例: "63R6" -> 6.0, "50R5" -> 5.0, "32R0.8" -> 0.8, "D10平刀" -> 0.0
        """
        if not isinstance(tool_name, str):
            return 0.0
        
        # 匹配R后面的数字（支持小数）
        import re
        pattern = r'R(\d+\.?\d*)'
        match = re.search(pattern, tool_name, re.IGNORECASE)
        
        if match:
            try:
                return float(match.group(1))
            except:
                return 0.0
        return 0.0
    
    def load_material_info(self, prt_folder: str | Path, excel_path: str | Path):
        """
        根据 final_workpiece_prt 文件夹下的 .prt 文件名，匹配零件参数表获取材质和热处理状态
        """
        prt_folder = Path(prt_folder)
        excel_path = Path(excel_path)
        
        if not prt_folder.exists():
            print(f"零件路径不存在: {prt_folder}")
            self.material = "45#"
            self.is_heat_treated = False
            return
        
        # 1. 判断是文件还是文件夹
        if prt_folder.is_file():
             prt_files = [prt_folder]
        else:
             prt_files = list(prt_folder.glob("*.prt"))

        if not prt_files:
            print("未找到任何 .prt 文件，使用默认材质 45#")
            self.material = "45#"
            self.is_heat_treated = False
            return
        
        if len(prt_files) > 1:
            print(f"警告：发现多个prt文件，使用第一个: {prt_files[0].name}")
        
        prt_path = prt_files[0]
        filename = prt_path.stem
        print(f"检测到零件文件: {prt_path.name}")
        
        # 2. 提取前缀
        match = re.match(r"([A-Z]+-\d+)", filename.upper())
        if match:
            prefix = match.group(1)
        else:
            if '-' in filename:
                parts = filename.upper().split('_')[0].split('-')
                prefix = parts[0] + '-' + parts[1]
            else:
                prefix = filename.upper()
        
        print(f"解析零件编号: {prefix}")
        
        # 3. 加载Excel表
        df = self._load_material_excel(excel_path)
        if df.empty:
            print("零件参数表为空，使用默认材质 45#")
            self.material = "45#"
            self.is_heat_treated = False
            return
        
        # 4. 智能匹配
        mask = (
            df.astype(str).apply(lambda col: col.str.contains(prefix, case=False, na=False)).any(axis=1)
        )
        matched_rows = df[mask]
        
        if matched_rows.empty:
            print(f"未在零件参数表中找到编号包含 '{prefix}' 的记录，使用默认材质 45#")
            self.material = "45#"
            self.is_heat_treated = False
            return
        
        row = matched_rows.iloc[0]
        
        # 提取材质
        material_candidates = ['材质']
        material = None
        for col in material_candidates:
            if col in row and pd.notna(row[col]):
                material = str(row[col]).strip()
                break
        
        if not material:
            print("未找到材质信息，使用默认 45#")
            material = "45#"
        
        # 判断是否热处理
        heat_treatment_candidates = ['热处理']
        is_heat_treated = False
        for col in heat_treatment_candidates:
            if col in row and pd.notna(row[col]):
                ht_text = str(row[col]).strip()
                if ht_text and ht_text.lower() not in ['无', '否', '-', '']:
                    is_heat_treated = True
                    print(f"检测到热处理: {ht_text}")
                    break
        
        self.material = material
        self.is_heat_treated = is_heat_treated
        print(f"→ 材质: {material} | 是否热处理: {is_heat_treated}")
    
    def parse_point(self, s: str) -> Tuple[float, float, float]:
        """解析点坐标字符串"""
        try:
            return tuple(map(float, s.strip().split(',')))
        except:
            return (0.0, 0.0, 0.0)
    
    def is_z_up(self, face_row: Dict) -> bool:
        """判断法向量是否朝上"""
        try:
            normal_str = face_row.get("Face Normal", "0,0,0")
            nx, ny, nz = self.parse_point(normal_str)
            return nz > 0.999
        except:
            return False
    
    def load_face_data(self, face_csv_path: str | Path):
        """
        加载面数据CSV，用于判断开放/封闭
        
        Args:
            face_csv_path: 面数据CSV文件路径
        """
        try:
            df = pd.read_csv(face_csv_path)
            print(f"成功加载面数据，共 {len(df)} 行")
            
            # 创建面数据字典，键为面标签
            for _, row in df.iterrows():
                face_tag = str(row.get("Face Tag", "")).strip()
                if face_tag:
                    # 转换为字典存储
                    self._face_data_dict[int(face_tag)] = {
                        "point": row.get("Face Data - Point", "0,0,0"),
                        "normal": row.get("Face Normal", "0,0,1"),
                        "adjacent_tags": row.get("Adjacent Face Tags", "")
                    }
            
            print(f"成功缓存 {len(self._face_data_dict)} 个面的数据")
            return True
        except Exception as e:
            print(f"加载面数据失败: {e}")
            return False
    
    def is_open_face(self, face_tag: int) -> bool:
        """
        判断一个面是否是开放面（参考上一对话中的判断逻辑）
        
        Args:
            face_tag: 面标签
            
        Returns:
            True: 开放面
            False: 封闭面
        """
        if not self._face_data_dict:
            print(f"警告：未加载面数据，无法判断面 {face_tag} 的开放/封闭状态")
            return False
        
        face_info = self._face_data_dict.get(face_tag)
        if not face_info:
            print(f"警告：未找到面 {face_tag} 的数据")
            return False
        
        # 检查是否是Z向上的面
        normal_str = face_info.get("normal", "0,0,1")
        if not self.is_z_up({"Face Normal": normal_str}):
            return False
        
        # 获取当前面的Z坐标
        try:
            point_str = face_info.get("point", "0,0,0")
            z_self = self.parse_point(point_str)[2]
        except:
            return False
        
        # 获取相邻面标签
        adj_tags_str = face_info.get("adjacent_tags", "")
        adj_tags = [t.strip() for t in adj_tags_str.split(';') if t.strip()]
        
        # 获取相邻面的Z坐标
        zs = []
        for t in adj_tags:
            try:
                t_int = int(t)
                adj_info = self._face_data_dict.get(t_int)
                if adj_info:
                    adj_point_str = adj_info.get("point", "0,0,0")
                    z_adj = self.parse_point(adj_point_str)[2]
                    zs.append(z_adj)
            except:
                continue
        
        if not zs:
            return False  # 没有相邻面，默认为封闭面
        
        # 判断是否是开放面：存在任意一个相邻面的Z坐标小于当前面的Z坐标
        is_open = any(z < z_self - 1e-6 for z in zs)
        
        return is_open
    
    def is_closed_pocket(self, bottom_tag: int) -> bool:
        """
        判断一个槽是否是封闭槽
        
        Args:
            bottom_tag: 底面标签
            
        Returns:
            True: 封闭槽
            False: 开放槽
        """
        if not self._face_data_dict:
            return True
        
        # 获取底面信息
        bottom_info = self._face_data_dict.get(bottom_tag)
        if not bottom_info:
            return True
        
        # 获取底面法向量
        bottom_normal_str = bottom_info.get("Face Normal", "0,0,1")
        bn_x, bn_y, bn_z = self.parse_point(bottom_normal_str)
        
        # 底面应该是朝上的（法向量朝-Z方向）
        if bn_z > -0.9:
            print(f"底面 {bottom_tag} 法向量 {bottom_normal_str} 不是朝上，可能不是有效的底面")
        
        # 获取相邻面信息
        adj_faces = self.get_adjacent_faces_info(bottom_tag)
        
        if len(adj_faces) == 0:
            print(f"底面 {bottom_tag} 没有相邻面，是开放面")
            return False
        
        # 分析相邻面
        vertical_faces = [f for f in adj_faces if f['is_vertical']]
        horizontal_faces = [f for f in adj_faces if f['is_horizontal']]
        
        print(f"底面 {bottom_tag} 有 {len(adj_faces)} 个相邻面，其中垂直面: {len(vertical_faces)}，水平面: {len(horizontal_faces)}")
        
        # 如果存在水平相邻面，说明这个底面是开放的（连接到其他水平面）
        if len(horizontal_faces) > 0:
            print(f"底面 {bottom_tag} 有水平相邻面，判断为开放")
            for face in horizontal_faces:
                print(f"  水平相邻面 {face['tag']}，法向量 {face['normal']}")
            return False
        
        # 检查相邻面是否都是侧壁（法向量应该大致水平且指向槽内）
        # 对于封闭槽，所有相邻面应该是侧壁
        if len(vertical_faces) < 3:  # 至少需要3个侧壁才能形成封闭
            print(f"底面 {bottom_tag} 只有 {len(vertical_faces)} 个垂直相邻面，不足以形成封闭")
            return False
        
        # 检查相邻面是否形成闭合环
        # 更简单的判断：如果所有相邻面都是垂直的，且数量足够，通常就是封闭的
        # 对于复杂情况，可能需要更复杂的几何判断
        
        print(f"底面 {bottom_tag} 判断为封闭")
        return True
    
    def read_tool_parameters_with_depth(self, json_file: str | Path) -> pd.DataFrame:
        """
        从JSON文件读取铣刀参数表，包含切削深度、转速、进给、横越信息
        """
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                tool_data = json.load(f)
            
            df = pd.DataFrame(tool_data)

            cutting_depth_columns = [
                '45#,A3,切深', 'CR12热处理前切深', 'CR12mov,SKD11,SKH-9,DC53,热处理前切深',
                'P20切深', 'TOOLOX33 TOOLOX44切深', '合金铜切深',
                'CR12热处理后切深', 'CR12mov,SKD11,SKH-9,DC53,热处理后切深'
            ]
            
            speed_feed_columns = ['转速(普)', '进给(普)', '横越(普)']
            
            selected_columns = ['类别', '刀具名称', '直径'] + cutting_depth_columns + speed_feed_columns
            available_columns = [col for col in selected_columns if col in df.columns]

            df = df[available_columns].copy()
            df = df.dropna(subset=['刀具名称', '直径'])
            df['直径'] = pd.to_numeric(df['直径'], errors='coerce')
            df = df.dropna(subset=['直径'])
            
            for col in speed_feed_columns:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
            
            # 如果已热处理，过滤掉所有飞刀类型的刀具
            if self.is_heat_treated:
                original_count = len(df)
                df = df[~df['类别'].str.contains('飞刀|飞', na=False, case=False)]
                filtered_count = len(df)
                print(f"已热处理零件，过滤掉所有飞刀刀具：过滤前 {original_count} 把，过滤后 {filtered_count} 把")

            print(f"成功读取 {len(df)} 把刀具，包含刀具类别、切削深度和转速/进给信息")
            
            # 打印所有刀具类别用于调试
            categories = df['类别'].unique().tolist()
            print(f"刀具类别列表: {categories}")
            
            return df

        except Exception as e:
            print(f"读取刀具参数JSON文件时出错: {e}")
            return pd.DataFrame()
    
    def get_cutting_depth(self, tool_row) -> float:
        """根据材质和热处理状态获取对应的切削深度"""
        material_mapping = {
            '45#': '45#,A3,切深',
            'A3': '45#,A3,切深',
            'CR12': 'CR12热处理后切深' if self.is_heat_treated else 'CR12热处理前切深',
            'CR12MOV': 'CR12mov,SKD11,SKH-9,DC53,热处理后切深' if self.is_heat_treated else 'CR12mov,SKD11,SKH-9,DC53,热处理前切深',
            'SKD11': 'CR12mov,SKD11,SKH-9,DC53,热处理后切深' if self.is_heat_treated else 'CR12mov,SKD11,SKH-9,DC53,热处理前切深',
            'SKH-9': 'CR12mov,SKD11,SKH-9,DC53,热处理后切深' if self.is_heat_treated else 'CR12mov,SKD11,SKH-9,DC53,热处理前切深',
            'DC53': 'CR12mov,SKD11,SKH-9,DC53,热处理后切深' if self.is_heat_treated else 'CR12mov,SKD11,SKH-9,DC53,热处理前切深',
            'P20': 'P20切深',
            'TOOLOX33': 'TOOLOX33 TOOLOX44切深',
            'TOOLOX44': 'TOOLOX33 TOOLOX44切深',
            'T00L0X33': 'TOOLOX33 TOOLOX44切深',
            'T00L0X44': 'TOOLOX33 TOOLOX44切深',
            '合金铜': '合金铜切深'
        }

        if self.material and self.material.upper() in material_mapping:
            column_name = material_mapping[self.material.upper()]
            if column_name in tool_row:
                return tool_row[column_name]

        return tool_row.get('45#,A3,切深', 0.1)
    
    def read_direction_mapping(self, direction_file: str | Path) -> dict:
        """读取方向映射文件，返回面标签到方向的映射字典"""
        try:
            df = pd.read_csv(direction_file)
            direction_map = {}
            
            for column in df.columns:
                for face_tag in df[column].dropna():
                    direction_map[int(face_tag)] = column.strip()

            print(f"成功读取方向映射，共 {len(direction_map)} 个面标签")
            return direction_map

        except Exception as e:
            print(f"读取方向映射文件时出错: {e}")
            return {}
    
    def validate_tool_for_finish_machining(self, tool_row, operation_type, is_semi_finish):
        """
        验证刀具是否适合精加工
        
        Args:
            tool_row: 刀具行数据
            operation_type: 工序类型（'螺旋'或'往复等高'）
            is_semi_finish: 是否为半精加工
            
        Returns:
            bool: 是否有效
            str: 错误信息（如果无效）
        """
        tool_category = tool_row.get('类别', '')
        tool_name = tool_row.get('刀具名称', '')

        # ============ 新增：半精加工+非热处理的R角检查 ============
        if is_semi_finish and not self.is_heat_treated:
            # 只检查飞刀的R角
            if '装刀粒样式' in str(tool_category) and '飞刀' in str(tool_category):
                tool_R_angle = self.parse_tool_R_angle(tool_name)
                if tool_R_angle > 1.0:
                    return False, f"半精加工+非热处理不能使用R角大于1mm的装刀粒样式飞刀: {tool_name} (R角={tool_R_angle}mm)"
        # =========================================================
        
        # 全精加工验证
        if not is_semi_finish:
            # 检查是否是飞刀
            if '飞' in str(tool_category).lower():
                return False, f"全精加工坚决不允许使用飞刀: {tool_category}"
            
            # ============ 新增：全精加工必须是钨钢平刀 ============
            if '钨钢平刀' not in str(tool_category):
                return False, f"全精加工必须使用钨钢平刀，当前刀具类别: {tool_category}"
            # ============ 新增结束 ============
        
        # 半精加工验证
        else:
            if self.is_heat_treated and '钨钢平刀' not in str(tool_category):
                return False, f"半精热处理零件必须使用钨钢平刀，当前刀具类别: {tool_category}"
        
        return True, "刀具验证通过"
    

    def get_adjacent_faces_info(self, face_tag: int) -> List[Dict]:
        """
        获取相邻面的详细信息
        """
        if not self._face_data_dict:
            return []
        
        face_info = self._face_data_dict.get(face_tag)
        if not face_info:
            return []
        
        adj_tags_str = face_info.get("Adjacent Face Tags", "")
        adj_tags = [t.strip() for t in adj_tags_str.split(';') if t.strip()]
        
        result = []
        for adj_tag in adj_tags:
            adj_info = self._face_data_dict.get(int(adj_tag))
            if adj_info:
                normal_str = adj_info.get("Face Normal", "0,0,1")
                nx, ny, nz = self.parse_point(normal_str)
                result.append({
                    'tag': int(adj_tag),
                    'normal': (nx, ny, nz),
                    'is_vertical': abs(nz) > 0.9,  # 法向量接近垂直
                    'is_horizontal': abs(nz) < 0.1,  # 法向量接近水平
                })
        
        return result
    

    def check_adjacent_z_relative_to_bottom(self, bottom_tag: int) -> bool:
        """
        检查底面的所有相邻面的Z坐标是否都高于底面
        
        Args:
            bottom_tag: 底面标签
            
        Returns:
            True: 所有相邻面Z > 底面Z
            False: 至少有一个相邻面Z <= 底面Z
        """
        if not self._face_data_dict:
            print(f"警告：未加载面数据，无法判断相邻面Z坐标关系")
            return False  # 无法判断时保守返回False（视为不封闭）
        
        bottom_info = self._face_data_dict.get(bottom_tag)
        if not bottom_info:
            print(f"警告：未找到底面 {bottom_tag} 的数据")
            return False
        
        # 获取底面Z坐标
        try:
            point_str = bottom_info.get("point", "0,0,0")
            z_self = self.parse_point(point_str)[2]
        except Exception as e:
            print(f"解析底面Z坐标失败: {e}")
            return False
        
        # 获取相邻面标签
        adj_tags_str = bottom_info.get("adjacent_tags", "")
        adj_tags = [t.strip() for t in adj_tags_str.split(';') if t.strip()]
        
        if not adj_tags:
            print(f"底面 {bottom_tag} 没有相邻面，无法形成封闭槽")
            return False
        
        # 检查每个相邻面的Z坐标
        lower_z_found = False
        for adj_tag in adj_tags:
            try:
                t_int = int(adj_tag)
                adj_info = self._face_data_dict.get(t_int)
                if adj_info:
                    adj_point_str = adj_info.get("point", "0,0,0")
                    z_adj = self.parse_point(adj_point_str)[2]
                    
                    # 如果有任何一个相邻面Z <= 底面Z，则不满足封闭条件
                    if z_adj <= z_self + 1e-6:  # 允许微小误差
                        lower_z_found = True
                        print(f"  相邻面 {t_int} Z={z_adj:.3f} <= 底面Z={z_self:.3f}，判定为开放")
                        break
                    else:
                        print(f"  相邻面 {t_int} Z={z_adj:.3f} > 底面Z={z_self:.3f}")
            except Exception as e:
                print(f"  处理相邻面 {adj_tag} 时出错: {e}")
                continue
        
        # 如果没有找到Z <= 底面的相邻面，则满足第一个封闭条件
        return not lower_z_found

    
    def group_features_with_tools(self, csv_path, json_file, direction_file, 
                                 face_csv_path: str | Path = None,  # 新增：面数据CSV路径
                                 is_semi_finish: bool = False):
        """
        分组特征并匹配刀具
        
        Args:
            is_semi_finish: 是否为半精加工模式，如果为True则只选择钨钢平刀
            face_csv_path: 面数据CSV路径，用于判断开放/封闭
        """
        # 如果提供了面数据CSV路径，则加载面数据用于判断开放/封闭
        if face_csv_path and os.path.exists(face_csv_path):
            self.load_face_data(face_csv_path)
            print("已加载面数据，将进行开放/封闭判断")
        else:
            print("警告：未提供面数据CSV路径，无法进行开放/封闭判断，将默认所有槽为封闭")
        
        # 读取刀具参数，传递热处理状态以过滤飞刀
        tools_df = self.read_tool_parameters_with_depth(json_file)
        if tools_df.empty:
            print("无法读取刀具参数，无法进行刀具匹配")
            return {}, {}
        
        # ============== 修改开始：根据新规则过滤刀具 ==============
        if is_semi_finish:
            # 半精加工模式
            if self.is_heat_treated:
                # 热处理为True：只选钨钢平刀
                original_count = len(tools_df)
                # 确保是钨钢平刀
                tools_df = tools_df[tools_df['类别'].str.contains('钨钢平刀', na=False, case=False)]
                filtered_count = len(tools_df)
                print(f"半精加工+热处理：只选择钨钢平刀，过滤前 {original_count} 把，过滤后 {filtered_count} 把")
            else:
                # 热处理为False：只从"钨钢平刀"和"装刀粒样式 飞刀"中选择
                original_count = len(tools_df)
                # 创建匹配模式：钨钢平刀 或 装刀粒样式 飞刀
                pattern = r'钨钢平刀|装刀粒样式\s*飞刀|装刀粒样式飞刀'
                tools_df = tools_df[tools_df['类别'].str.contains(pattern, na=False, case=False, regex=True)]
                filtered_count = len(tools_df)
                print(f"半精加工+非热处理：只选择钨钢平刀或装刀粒样式飞刀，过滤前 {original_count} 把，过滤后 {filtered_count} 把")
                
                # 打印选择的刀具类型用于调试
                if not tools_df.empty:
                    categories = tools_df['类别'].unique().tolist()
                    print(f"允许的刀具类别: {categories}")
                else:
                    print(f"警告：没有找到钨钢平刀或装刀粒样式飞刀")
        else:
            # 全精加工模式：无论是否热处理都坚决不选飞刀
            original_count = len(tools_df)
            
            # 坚决不选任何飞刀，包括各种可能的飞刀名称
            fly_pattern = re.compile(r'飞.*刀|飞', re.IGNORECASE)
            tools_df = tools_df[~tools_df['类别'].apply(lambda x: bool(fly_pattern.search(str(x))))]
            
            # ============ 新增：只选择钨钢平刀 ============
            tools_df = tools_df[tools_df['类别'].str.contains('钨钢平刀', na=False, case=False)]
            
            filtered_count = len(tools_df)
            print(f"全精加工：坚决不选飞刀，只选钨钢平刀，过滤前 {original_count} 把，过滤后 {filtered_count} 把")
            
            # 打印剩余的刀具类别，用于验证
            remaining_categories = tools_df['类别'].unique()
            print(f"全精加工允许的刀具类别: {remaining_categories}")
            
            # 如果有疑似飞刀残留，发出严重警告
            if any('飞' in str(cat).lower() for cat in remaining_categories):
                print(f"⚠️ 严重警告：全精加工刀具列表中仍有疑似飞刀！请检查刀具参数表！")
                print(f"  疑似飞刀类别: {[cat for cat in remaining_categories if '飞' in str(cat).lower()]}")
    # ============ 修改结束 ============
        # ============== 修改结束 ==============

        # 读取方向映射
        direction_map = self.read_direction_mapping(direction_file)
        if not direction_map:
            print("警告：无法读取方向映射文件，指定图层功能将不可用")

        # 读取特征数据
        df = pd.read_csv(csv_path)

        # 检查必要列
        required_columns = ['ID', 'Type', 'Role', 'Length', 'Width', 'Height', 'Attribute', 'Dia']
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            print(f"缺少列: {missing_columns}")
            return {}, {}

        df = df[required_columns]
        
        # 数据预处理：转为数值型
        df['Height'] = pd.to_numeric(df['Height'], errors='coerce')
        df['Length'] = pd.to_numeric(df['Length'], errors='coerce')
        df['Width'] = pd.to_numeric(df['Width'], errors='coerce')
        df['Dia'] = pd.to_numeric(df['Dia'], errors='coerce')

        # 筛选特征：包含 POCKET 但不是 STEP1POCKET
        pocket_mask = df['Type'].str.contains('POCKET', case=False, na=False)
        not_step1_mask = df['Type'].ne('STEP1POCKET')
        pocket_df = df[pocket_mask & not_step1_mask]
        
        print(f"原始特征数量: {len(df['ID'].unique())}")
        print(f"筛选后特征数量: {len(pocket_df['ID'].unique())}")

        # 初始化两个结果字典
        spiral_result = {}  # 螺旋 (侧壁高度一致且封闭)
        reciprocating_result = {}    # 往复等高 (侧壁高度不一致 或 开放)

        for feature_id, g in pocket_df.groupby('ID'):
            # 1. 寻找底面 (Bottom) 以确定基准
            bottom = g[g['Role'] == 'Bottom']
            if bottom.empty:
                continue

            bottom_row = bottom.iloc[0]
            bottom_tag = int(bottom_row['Attribute'])
            sr = min(float(bottom_row['Length']), float(bottom_row['Width']))  # 最短边
            tags = g['Attribute'].astype(int).tolist()
            
            # 获取特征的直径
            wall_faces = g[g['Role'].str.contains('Wall', case=False, na=False)]
            if not wall_faces.empty:
                positive_dia = wall_faces[wall_faces['Dia'] > 0]
                if not positive_dia.empty and positive_dia['Dia'].notna().any():
                    feature_dia = positive_dia['Dia'].min()
                else:
                    feature_dia = 0
            else:
                feature_dia = 0

            # 2. 判断侧壁高度一致性
            sides = g[g['Role'].str.contains('Wall', case=False, na=False)]
            is_height_consistent = True
            
            if not sides.empty:
                heights = sides['Height'].dropna()
                if not heights.empty:
                    min_h = heights.min()
                    max_h = heights.max()
                    if (max_h - min_h) > 0.001:
                        is_height_consistent = False
                        print(f"特征 {feature_id}: 侧壁高度不一致 (min={min_h:.2f}, max={max_h:.2f})")
                    else:
                        print(f"特征 {feature_id}: 侧壁高度一致 (min={min_h:.2f}, max={max_h:.2f})")
            
            # 3. 判断相邻面Z坐标关系（新逻辑）
            all_adj_higher = False
            if self._face_data_dict:
                all_adj_higher = self.check_adjacent_z_relative_to_bottom(bottom_tag)
                print(f"特征 {feature_id}: 相邻面Z坐标 {'全部高于' if all_adj_higher else '存在低于或等于'}底面")
            else:
                print(f"特征 {feature_id}: 未加载面数据，无法进行Z坐标判断")
            
            # 4. 根据新规则判断封闭/开放状态
            # 封闭槽：相邻面Z都高于底面 AND 侧壁高度一致
            is_closed = all_adj_higher and is_height_consistent
            
            # 打印最终判断结果
            if is_closed:
                print(f"→ 特征 {feature_id}: **封闭槽** (满足两个条件)")
            else:
                # 判断是哪一个条件不满足
                reasons = []
                if not all_adj_higher:
                    reasons.append("相邻面Z坐标不满足")
                if not is_height_consistent:
                    reasons.append("侧壁高度不一致")
                print(f"→ 特征 {feature_id}: **开放槽** (原因: {'; '.join(reasons)})")
            
            # 5. 根据新规则判断加工方式
            # 只有封闭槽才用螺旋，其他用往复等高
            is_spiral_candidate = is_closed
            print(f"特征 {feature_id}: 高度一致={is_height_consistent}, 封闭={is_closed}, 适用加工方式={'螺旋' if is_spiral_candidate else '往复等高'}")

            # 6. 刀具选择逻辑
            tool_selection_type = "最短边"
            best_tool = None  # 先初始化为None
            selected_from_range = False

            # ============== 修正：半精加工 + 非热处理的刀具选择逻辑 ==============
            if is_semi_finish:
                # 半精加工模式：基于最短边（SR）计算
                lower_bound = sr / 2
                upper_bound = sr * 3 / 4
                middle_value = (lower_bound + upper_bound) / 2
                
                if not self.is_heat_treated:
                    # 半精加工 + 非热处理：在范围内优先选择飞刀
                    
                    # 1. 在选择范围内找飞刀（装刀粒样式飞刀）
                    flying_tools_in_range = tools_df[
                        tools_df['类别'].str.contains('装刀粒样式\s*飞刀|装刀粒样式飞刀', na=False, case=False, regex=True) &
                        (tools_df['直径'] > lower_bound) &
                        (tools_df['直径'] < upper_bound)
                    ].copy()
                    
                    # 过滤掉R角>1的飞刀
                    valid_flying_tools = []
                    for idx, tool_row in flying_tools_in_range.iterrows():
                        tool_name = tool_row['刀具名称']
                        tool_R_angle = self.parse_tool_R_angle(tool_name)
                        
                        if tool_R_angle <= 1.0:  # 只保留R角≤1的飞刀
                            valid_flying_tools.append((idx, tool_row, tool_R_angle))
                    
                    # 2. 如果范围内有符合条件的飞刀（R角≤1），选择最接近中间值的
                    if valid_flying_tools:
                        # 选择最接近中间值的飞刀
                        valid_flying_df = pd.DataFrame([tool_row for _, tool_row, _ in valid_flying_tools])
                        valid_flying_df['差值'] = abs(valid_flying_df['直径'] - middle_value)
                        best_tool = valid_flying_df.loc[valid_flying_df['差值'].idxmin()]
                        print(f"特征 {feature_id}: 选择了R角≤1的飞刀 {best_tool['刀具名称']}")
                    
                    # 3. 如果范围内没有符合条件的飞刀（要么没有飞刀，要么飞刀R角>1）
                    else:
                        print(f"特征 {feature_id}: 范围内没有R角≤1的飞刀，尝试选择钨钢平刀")
                        
                        # 在范围内找钨钢平刀
                        tungsten_tools_in_range = tools_df[
                            tools_df['类别'].str.contains('钨钢平刀', na=False, case=False) &
                            (tools_df['直径'] > lower_bound) &
                            (tools_df['直径'] < upper_bound)
                        ].copy()
                        
                        if not tungsten_tools_in_range.empty:
                            # 选择最接近中间值的钨钢平刀
                            tungsten_tools_in_range['差值'] = abs(tungsten_tools_in_range['直径'] - middle_value)
                            best_tool = tungsten_tools_in_range.loc[tungsten_tools_in_range['差值'].idxmin()]
                            print(f"特征 {feature_id}: 选择了钨钢平刀 {best_tool['刀具名称']}")
                        
                        # 4. 如果范围内也没有钨钢平刀，则在全局范围内选择
                        else:
                            print(f"特征 {feature_id}: 范围内没有合适刀具，在全局选择")
                            
                            # 优先在全局找钨钢平刀
                            global_tungsten_tools = tools_df[
                                tools_df['类别'].str.contains('钨钢平刀', na=False, case=False)
                            ].copy()
                            
                            if not global_tungsten_tools.empty:
                                global_tungsten_tools['差值'] = abs(global_tungsten_tools['直径'] - middle_value)
                                best_tool = global_tungsten_tools.loc[global_tungsten_tools['差值'].idxmin()]
                                print(f"特征 {feature_id}: 全局选择了钨钢平刀 {best_tool['刀具名称']}")
                            
                            else:
                                # 最后尝试全局找R角≤1的飞刀
                                global_flying_tools = tools_df[
                                    tools_df['类别'].str.contains('装刀粒样式\s*飞刀|装刀粒样式飞刀', na=False, case=False, regex=True)
                                ].copy()
                                
                                valid_global_flying = []
                                for idx, tool_row in global_flying_tools.iterrows():
                                    tool_name = tool_row['刀具名称']
                                    tool_R_angle = self.parse_tool_R_angle(tool_name)
                                    if tool_R_angle <= 1.0:
                                        valid_global_flying.append((idx, tool_row, tool_R_angle))
                                
                                if valid_global_flying:
                                    valid_global_df = pd.DataFrame([tool_row for _, tool_row, _ in valid_global_flying])
                                    valid_global_df['差值'] = abs(valid_global_df['直径'] - middle_value)
                                    best_tool = valid_global_df.loc[valid_global_df['差值'].idxmin()]
                                    print(f"特征 {feature_id}: 全局选择了R角≤1的飞刀 {best_tool['刀具名称']}")
                                else:
                                    # 没有合适的刀具，跳过此特征
                                    print(f"特征 {feature_id}: 没有找到合适的刀具，跳过此特征")
                                    continue


                else:
                    # 半精加工 + 热处理：原来的选择逻辑保持不变，只选钨钢平刀
                    suitable_tools = tools_df[
                        (tools_df['直径'] > lower_bound) &
                        (tools_df['直径'] < upper_bound)
                    ].copy()
                    
                    if not suitable_tools.empty:
                        suitable_tools['差值'] = abs(suitable_tools['直径'] - middle_value)
                        best_tool = suitable_tools.loc[suitable_tools['差值'].idxmin()]
                        if not selected_from_range:
                            tool_selection_type = "范围选择(最佳)"
                        selected_from_range = True
                    else:
                        # 范围内无匹配，选择全局最接近
                        tools_df_copy = tools_df.copy()
                        tools_df_copy['差值'] = abs(tools_df_copy['直径'] - middle_value)
                        best_tool = tools_df_copy.loc[tools_df_copy['差值'].idxmin()]
                        if tool_selection_type == "最短边":
                            tool_selection_type = "全局选择(最接近)"
                        print(f"特征 {feature_id}: SR={sr:.2f}mm 无完美匹配刀具，使用最接近的刀具")
            
            else:
                # 全精加工模式：基于特征直径（圆柱面直径）选择刀具
                if feature_dia > 0:
                    # 基于特征直径选择刀具
                    target_diameter = feature_dia
                    
                    # 规则1: 查找直径等于特征直径的刀具
                    exact_match_tools = tools_df[tools_df['直径'] == target_diameter].copy()
                    
                    if not exact_match_tools.empty:
                        # 有直径完全匹配的刀具，选择第一个
                        best_tool = exact_match_tools.iloc[0]
                        tool_selection_type = f"精确匹配(直径={target_diameter:.2f}mm)"
                        print(f"特征 {feature_id}: 找到精确匹配刀具 {best_tool['刀具名称']} (直径={target_diameter:.2f}mm)")
                    else:
                        # 规则2: 选择直径小于特征直径且最接近的刀具
                        smaller_tools = tools_df[tools_df['直径'] < target_diameter].copy()
                        
                        if not smaller_tools.empty:
                            # 计算与目标直径的差值
                            smaller_tools.loc[:, '差值'] = target_diameter - smaller_tools['直径']
                            # 选择差值最小的（即最接近但小于目标直径）
                            best_tool = smaller_tools.loc[smaller_tools['差值'].idxmin()]
                            tool_selection_type = f"接近匹配(直径={best_tool['直径']:.2f}mm < {target_diameter:.2f}mm)"
                            print(f"特征 {feature_id}: 未找到精确匹配，选择最接近的刀具 {best_tool['刀具名称']} (直径={best_tool['直径']:.2f}mm < 目标直径{target_diameter:.2f}mm)")
                        else:
                            # 规则3: 没有更小的刀具，选择全局最接近的
                            tools_df_copy = tools_df.copy()
                            tools_df_copy.loc[:, '差值'] = abs(tools_df_copy['直径'] - target_diameter)
                            best_tool = tools_df_copy.loc[tools_df_copy['差值'].idxmin()]
                            tool_selection_type = f"全局选择(最近, 直径={best_tool['直径']:.2f}mm)"
                            print(f"特征 {feature_id}: 没有更小的刀具，使用全局最接近的刀具 {best_tool['刀具名称']} (直径={best_tool['直径']:.2f}mm)")
                else:
                    # 特征直径=0，说明不是圆柱槽，退回到基于最短边的选择
                    lower_bound = sr / 2
                    upper_bound = sr * 3 / 4
                    middle_value = (lower_bound + upper_bound) / 2
                    
                    suitable_tools = tools_df[
                        (tools_df['直径'] > lower_bound) &
                        (tools_df['直径'] < upper_bound)
                    ].copy()
                    
                    # 全精加工模式再次检查，确保没有飞刀
                    flying_tools_mask = suitable_tools['类别'].str.contains('飞', na=False, case=False)
                    if flying_tools_mask.any():
                        print(f"警告：在全精加工模式下发现飞刀，进行二次过滤")
                        suitable_tools = suitable_tools[~flying_tools_mask]
                    
                    if not suitable_tools.empty:
                        suitable_tools.loc[:, '差值'] = abs(suitable_tools['直径'] - middle_value)
                        best_tool = suitable_tools.loc[suitable_tools['差值'].idxmin()]
                        tool_selection_type = "范围选择(基于SR)"
                        selected_from_range = True
                        print(f"特征 {feature_id}: 非圆柱槽，基于最短边SR={sr:.2f}mm选择刀具 {best_tool['刀具名称']}")
                    else:
                        # 范围内无匹配，选择全局最接近
                        tools_df_copy = tools_df.copy()
                        tools_df_copy.loc[:, '差值'] = abs(tools_df_copy['直径'] - middle_value)
                        best_tool = tools_df_copy.loc[tools_df_copy['差值'].idxmin()]
                        tool_selection_type = "全局选择(基于SR)"
                        print(f"特征 {feature_id}: SR={sr:.2f}mm 无完美匹配刀具，使用全局最接近的刀具")
            # ============== 修改结束 ==============

            # 检查是否选择了刀具
            if best_tool is None:
                # 如果没有选择刀具，使用一个默认的刀具选择逻辑
                print(f"警告：特征 {feature_id} 未找到合适的刀具，使用默认选择")
                tools_df_copy = tools_df.copy()
                tools_df_copy.loc[:, '差值'] = abs(tools_df_copy['直径'] - middle_value)
                best_tool = tools_df_copy.loc[tools_df_copy['差值'].idxmin()]
                tool_selection_type = "默认选择"

            # ============ 新增：刀具验证 ============
            operation_type_for_validation = '螺旋' if is_spiral_candidate else '往复等高'
            is_valid, error_msg = self.validate_tool_for_finish_machining(
                best_tool, 
                operation_type=operation_type_for_validation,
                is_semi_finish=is_semi_finish
            )
            
            if not is_valid:
                print(f"刀具验证失败: {error_msg}")
                continue  # 跳过这个特征
            # ========================================

            tool_name = best_tool['刀具名称']
            tool_category = best_tool['类别'] if '类别' in best_tool else "未知"
            tool_diameter = best_tool['直径']
            cutting_depth = self.get_cutting_depth(best_tool)

            spindle_speed = float(best_tool.get('转速(普)')) if '转速(普)' in best_tool and pd.notna(best_tool.get('转速(普)')) else None
            feed_rate = float(best_tool.get('进给(普)')) if '进给(普)' in best_tool and pd.notna(best_tool.get('进给(普)')) else None
            traverse_rate = float(best_tool.get('横越(普)')) if '横越(普)' in best_tool and pd.notna(best_tool.get('横越(普)')) else None

            # 打印选择结果，如果是半精加工模式特别标注
            mode_note = "【半精模式】" if is_semi_finish else "【全精模式】"
            print(f"特征 {feature_id}{mode_note}: 最短边={sr:.2f}mm → {tool_selection_type}选择刀具 {tool_name} (类别={tool_category}, 直径={tool_diameter}mm, 转速={spindle_speed}, 进给={feed_rate}, 横越={traverse_rate})")


            # 6. 获取指定图层
            specified_layer = 0
            if tags and direction_map:
                first_tag = tags[0]
                direction = direction_map.get(first_tag)
                if direction:
                    specified_layer = self.get_layer_by_direction(direction)

            # 7. 根据加工方式存入不同的字典
            if is_spiral_candidate:
                target_dict = spiral_result
                operation_type = "螺旋"
            else:
                target_dict = reciprocating_result
                operation_type = "往复等高"
            
            # 使用(最短边, 刀具名称)作为键，避免相同最短边但不同刀具的冲突
            group_key = (sr, tool_name, operation_type)

            if group_key not in target_dict:
                target_dict[group_key] = {
                    'face_tags': [int(tag) for tag in tags],
                    'tool_name': str(tool_name),
                    'tool_category': str(tool_category),
                    'tool_diameter': float(tool_diameter),
                    'short_side': float(sr),
                    'feature_dia': float(feature_dia) if feature_dia > 0 else 0,
                    'is_heat_treated': bool(self.is_heat_treated),
                    'material': str(self.material),
                    'cutting_depth': float(cutting_depth),
                    'spindle_speed': spindle_speed,
                    'feed_rate': feed_rate,
                    'traverse_rate': traverse_rate,
                    'specified_layer': int(specified_layer),
                    'selection_type': str(tool_selection_type),
                    'operation_type': operation_type,  # 新增：记录加工方式
                    'is_height_consistent': bool(is_height_consistent),  # 新增：记录高度一致性
                    'is_closed': bool(is_closed)  # 新增：记录封闭状态
                }
            else:
                target_dict[group_key]['face_tags'].extend([int(tag) for tag in tags])
        
        # 在返回结果前进行刀具整合
        spiral_result = self.consolidate_tools(spiral_result)
        reciprocating_result = self.consolidate_tools(reciprocating_result)
        
        self.spiral_data = spiral_result
        self.half_spiral_data = reciprocating_result
        
        # 打印统计信息
        print(f"\n加工方式统计:")
        print(f"  螺旋加工: {len(spiral_result)} 组")
        print(f"  往复等高: {len(reciprocating_result)} 组")
        
        return spiral_result, reciprocating_result

    def consolidate_tools(self, result_dict: dict) -> dict:
        """
        整合刀具，减少刀具数量
        
        规则：
        1. 只有刀具类型相同的刀具可以整合
        2. 只有指定图层相同的刀具可以整合
        3. 钨钢平刀直径相差2之内可以整合，其他刀相差5之内可以整合
        4. 最多三把直径不同的刀当作一组，直径相近的刀可以整合，比如12，13，14，15，16，17；其中(12，13，14)一组，(15，16，17)一组
        5. 选择组内直径最小的刀
        
        Args:
            result_dict: 原始的分组结果字典
        
        Returns:
            整合后的结果字典
        """
        if not result_dict:
            return {}
        
        # 将字典转换为列表，便于处理
        items = list(result_dict.items())
        
        # 首先按刀具类别和指定图层分组
        category_layer_groups = {}
        for key, data in items:
            category = data['tool_category']
            specified_layer = data['specified_layer']
            group_key = (category, specified_layer)
            
            if group_key not in category_layer_groups:
                category_layer_groups[group_key] = []
            category_layer_groups[group_key].append((key, data))
        
        # 打印原始分组信息，用于调试
        print(f"\n原始刀具分组信息:")
        for (category, layer), group_items in category_layer_groups.items():
            print(f"  类别 '{category}'，图层 {layer}:")
            for key, data in group_items:
                print(f"    - 刀具: {data['tool_name']}, 直径: {data['tool_diameter']}mm, 面数: {len(data['face_tags'])}")
        
        consolidated_result = {}
        new_key_counter = 1
        
        # 对每个(刀具类别, 指定图层)组合进行整合
        for (category, layer), group_items in category_layer_groups.items():
            # 按刀具直径排序
            group_items.sort(key=lambda x: x[1]['tool_diameter'])
            
            # 确定阈值
            threshold = 2 if category == "钨钢平刀" else 5
            
            # 打印该组的阈值信息
            print(f"\n整合类别 '{category}'，图层 {layer}，阈值: {threshold}mm")
            
            # 使用贪心算法进行整合
            i = 0
            while i < len(group_items):
                # 创建新组
                group = []
                base_item = group_items[i]
                base_diameter = base_item[1]['tool_diameter']
                group.append(base_item)
                
                # 查找可以加入当前组的刀具
                j = i + 1
                while j < len(group_items) and len(group) < 3:
                    current_item = group_items[j]
                    current_diameter = current_item[1]['tool_diameter']
                    
                    # 检查是否满足阈值条件
                    if abs(current_diameter - base_diameter) <= threshold:
                        group.append(current_item)
                        j += 1
                    else:
                        # 如果直径差超过阈值，则不能加入
                        break
                
                # 打印组信息
                if len(group) > 1:
                    tool_names = [item[1]['tool_name'] for item in group]
                    diameters = [item[1]['tool_diameter'] for item in group]
                    print(f"  整合组 {new_key_counter}: {tool_names} (直径: {diameters}mm) → 选择 {group[0][1]['tool_name']}")
                
                # 选择组内直径最小的刀具作为代表
                # group已经按直径排序，所以第一个就是最小的
                best_item = group[0]
                best_data = best_item[1].copy()
                
                # 合并所有面标签
                all_face_tags = []
                for _, data in group:
                    all_face_tags.extend(data['face_tags'])
                
                # 更新数据
                best_data['face_tags'] = all_face_tags
                if len(group) > 1:
                    best_data['selection_type'] = f"整合({len(group)}把刀具)"
                
                # 添加到结果
                consolidated_result[f"group_{new_key_counter}"] = best_data
                new_key_counter += 1
                
                # 移动到下一组
                i = j
        
        print(f"\n整合完成: {len(items)}组 → {len(consolidated_result)}组")
        
        # 打印整合后的总结
        print(f"\n整合后分组信息:")
        for group_key, data in consolidated_result.items():
            print(f"  {group_key}: 刀具: {data['tool_name']}, 类别: {data['tool_category']}, 直径: {data['tool_diameter']}mm, 图层: {data['specified_layer']}, 面数: {len(data['face_tags'])}")
        
        return consolidated_result


    def print_enhanced_result(self, result: dict, title: str):
        """打印增强的分组结果"""
        if not result:
            print(f"\n[{title}] 没有生成任何分组结果")
            return

        print(f"\n=== {title} 分组结果 ===")
        for key, data in sorted(result.items()):
            operation_type = data.get('operation_type', '未知')
            height_consistent = "高度一致" if data.get('is_height_consistent', False) else "高度不一致"
            closed_status = "封闭" if data.get('is_closed', False) else "开放"
            
            if data.get('feature_dia', 0) > 0:
                selection_info = data.get('selection_type', '')
                speed_info = f"转速={data.get('spindle_speed')} | 进给={data.get('feed_rate')} | 横越={data.get('traverse_rate')}"
                print(f"  组{key}: {operation_type} | 最小Dia: {data['feature_dia']:.2f}mm | 刀具: {data['tool_name']} (类别={data.get('tool_category', '未知')}, 直径={data['tool_diameter']}mm) | 状态: {height_consistent}/{closed_status} | {speed_info} | 包含面: {len(data['face_tags'])}")
            else:
                selection_info = data.get('selection_type', '')
                speed_info = f"转速={data.get('spindle_speed')} | 进给={data.get('feed_rate')} | 横越={data.get('traverse_rate')}"
                print(f"  组{key}: {operation_type} | 最短边: {data['short_side']:.2f}mm | 刀具: {data['tool_name']} (类别={data.get('tool_category', '未知')}, 直径={data['tool_diameter']}mm) | 状态: {height_consistent}/{closed_status} | {speed_info} | 包含面: {len(data['face_tags'])}")


    def save_result_to_separate_json(self, result: dict, 
                               half_finish_output: str | Path,
                               full_finish_output: str | Path,
                               operation_name: str,
                               prefix: str = "",
                               is_semi_finish: bool = False):
        """
        将结果保存为两个独立的JSON文件：半精铣和全精铣
        
        Args:
            is_semi_finish: 是否为半精加工模式
        """
        if not result:
            print(f"结果为空，跳过保存 {operation_name} 的JSON文件")
            return
        
        # 打印结果摘要，用于调试
        mode_note = "【半精模式】" if is_semi_finish else "【全精模式】"
        print(f"\n保存 {operation_name} 结果{mode_note}:")
        
        # 创建输出目录
        for output_file in [half_finish_output, full_finish_output]:
            os.makedirs(os.path.dirname(output_file), exist_ok=True)
        
        # 按照特征直径或最短边排序
        sorted_items = sorted(result.items(), key=lambda x: x[1].get('tool_diameter', 0))
        
        # 1. 创建半精铣JSON数据
        half_json_data = {}
        if is_semi_finish:
            # 半精加工模式：已经在选择阶段过滤了，这里直接使用所有结果
            for i, (key, info) in enumerate(sorted_items, 1):
                semi_key = f"{prefix}{i}" if prefix else f"工序{i}"
                
                # 根据刀具类别设置侧面余量
                tool_category = info.get('tool_category', '')
                if tool_category == "钨钢平刀":
                    side_allowance = 0.05
                else:
                    side_allowance = 0.2
                bottom_allowance = 0.5
                
                # 验证刀具名称
                tool_name = info['tool_name']
                if tool_name != tool_name.strip():
                    print(f"警告: 刀具名称 '{tool_name}' 包含空格")
                
                half_json_data[semi_key] = {
                    "工序": operation_name,
                    "面ID列表": info['face_tags'],
                    "刀具名称": tool_name,
                    "刀具类别": tool_category,
                    "切深": info['cutting_depth'],
                    "指定图层": info['specified_layer'],
                    "参考刀具": None,
                    "转速": info.get('spindle_speed'),
                    "进给": info.get('feed_rate'),
                    "横越": info.get('traverse_rate'),
                    "部件侧面余量": side_allowance,
                    "部件底面余量": bottom_allowance
                }
        else:
            # 精加工模式：半精文件不需要生成
            half_json_data = {}
        
        # 2. 创建全精铣JSON数据
        full_json_data = {}
        if not is_semi_finish:
            # 精加工模式：生成全精文件
            for i, (key, info) in enumerate(sorted_items, 1):
                full_key = f"{prefix}{i}" if prefix else f"工序{i}"
                
                # ============ 最终验证：全精加工坚决不使用飞刀 ============
                tool_category = info.get('tool_category', '')
                if '飞' in str(tool_category).lower():
                    print(f"❌ 严重错误：全精加工发现飞刀 {info['tool_name']} (类别: {tool_category})，跳过此工序")
                    continue
                
                # 全精加工必须是钨钢平刀
                if '钨钢平刀' not in str(tool_category):
                    print(f"❌ 严重错误：全精加工必须用钨钢平刀，当前刀具: {info['tool_name']} (类别: {tool_category})，跳过此工序")
                    continue
                # ======================================================
                
                side_allowance = 0.0
                bottom_allowance = 0.0
                
                full_json_data[full_key] = {
                    "工序": operation_name,
                    "面ID列表": info['face_tags'],
                    "刀具名称": info['tool_name'],
                    "刀具类别": info.get('tool_category', '未知'),
                    "切深": 10.0,
                    "指定图层": info['specified_layer'],
                    "参考刀具": None,
                    "转速": info.get('spindle_speed'),
                    "进给": info.get('feed_rate'),
                    "横越": info.get('traverse_rate'),
                    "部件侧面余量": side_allowance,
                    "部件底面余量": bottom_allowance
                }
        else:
            # 半精加工模式：全精文件不需要生成
            full_json_data = {}

        # 保存半精铣JSON文件
        if half_json_data:
            try:
                with open(half_finish_output, 'w', encoding='utf-8') as f:
                    json.dump(half_json_data, f, ensure_ascii=False, indent=4)
                print(f"半精铣JSON已保存: {half_finish_output} ({len(half_json_data)}个工序)")
            except Exception as e:
                print(f"保存半精铣 {half_finish_output} 时出错: {e}")
        elif is_semi_finish:
            print(f"半精铣结果为空，未创建文件: {half_finish_output}")
        
        # 保存全精铣JSON文件
        if full_json_data:
            try:
                with open(full_finish_output, 'w', encoding='utf-8') as f:
                    json.dump(full_json_data, f, ensure_ascii=False, indent=4)
                print(f"全精铣JSON已保存: {full_finish_output} ({len(full_json_data)}个工序)")
                
                # 验证全精加工中是否有飞刀
                for key, data in full_json_data.items():
                    tool_category = data.get('刀具类别', '')
                    if '飞' in str(tool_category).lower():
                        print(f"⚠️ 最终验证警告：全精加工JSON中仍有飞刀: {data['刀具名称']} (类别: {tool_category})")
                
            except Exception as e:
                print(f"保存全精铣 {full_finish_output} 时出错: {e}")
        elif not is_semi_finish:
            print(f"全精铣结果为空，未创建文件: {full_finish_output}")

    def process_spiral_data(self, output_half_spiral: str | Path, output_full_spiral: str | Path):
        """处理螺旋数据并保存结果"""
        self.print_enhanced_result(self.spiral_data, "螺旋 (侧壁等高且封闭)")
        self.save_result_to_separate_json(
            result=self.spiral_data,
            half_finish_output=output_half_spiral,
            full_finish_output=output_full_spiral,
            operation_name="D4-螺旋_SIMPLE",
            prefix=""
        )
    
    def process_half_spiral_data(self, output_half_half_spiral: str | Path, output_full_half_spiral: str | Path):
        """处理半螺旋数据并保存结果"""
        self.print_enhanced_result(self.half_spiral_data, "往复等高 (侧壁不等高或开放)")
        self.save_result_to_separate_json(
            result=self.half_spiral_data,
            half_finish_output=output_half_half_spiral,
            full_finish_output=output_full_half_spiral,
            operation_name="往复等高_SIMPLE",
            prefix=""
        )
    
    def full_processing_pipeline(self, prt_folder: str | Path, excel_params: str | Path,
                            csv_file: str | Path, json_file: str | Path, direction_file: str | Path,
                            output_half_spiral: str | Path, output_full_spiral: str | Path,
                            output_half_half_spiral: str | Path, output_full_half_spiral: str | Path,
                            face_csv_path: str | Path = None):  # 新增：面数据CSV路径
        """
        完整的处理流程，从数据加载到结果保存
        """
        print("=" * 60)
        print("开始螺旋加工特征处理流程")
        print("=" * 60)
        
        # 1. 加载材质信息
        print("\n步骤1: 加载材质和热处理信息")
        print("-" * 40)
        self.load_material_info(prt_folder, excel_params)
        
        # 2. 处理半精加工数据（只选择钨钢平刀）
        print("\n步骤2: 分组特征并匹配刀具 - 半精加工模式（只选钨钢平刀）")
        print("-" * 40)
        # 先清除之前的数据
        self.spiral_data = {}
        self.half_spiral_data = {}
        
        # 半精加工模式
        semi_normal_result, semi_half_result = self.group_features_with_tools(
            csv_file, json_file, direction_file, 
            face_csv_path=face_csv_path,  # 传递面数据CSV路径
            is_semi_finish=True
        )
        
        # 处理螺旋数据（半精）
        if semi_normal_result:
            print("\n步骤3: 处理螺旋数据（侧壁等高且封闭）- 半精加工")
            print("-" * 40)
            self.print_enhanced_result(semi_normal_result, "螺旋 (侧壁等高且封闭) - 半精加工")
            self.save_result_to_separate_json(
                result=semi_normal_result,
                half_finish_output=output_half_spiral,
                full_finish_output=output_full_spiral,  # 半精加工时这个参数不会被使用
                operation_name="D4-螺旋_SIMPLE",
                prefix="",
                is_semi_finish=True  # 新增参数，表示是半精加工
            )
        
        # 处理往复等高数据（半精）
        if semi_half_result:
            print("\n步骤4: 处理往复等高数据（侧壁不等高或开放）- 半精加工")
            print("-" * 40)
            self.print_enhanced_result(semi_half_result, "往复等高 (侧壁不等高或开放) - 半精加工")
            self.save_result_to_separate_json(
                result=semi_half_result,
                half_finish_output=output_half_half_spiral,
                full_finish_output=output_full_half_spiral,  # 半精加工时这个参数不会被使用
                operation_name="往复等高_SIMPLE",
                prefix="",
                is_semi_finish=True  # 新增参数，表示是半精加工
            )
        
        # 3. 处理精加工数据（使用所有刀具）
        print("\n步骤5: 分组特征并匹配刀具 - 精加工模式（使用所有刀具）")
        print("-" * 40)
        # 先清除之前的数据
        self.spiral_data = {}
        self.half_spiral_data = {}
        
        # 精加工模式
        full_normal_result, full_half_result = self.group_features_with_tools(
            csv_file, json_file, direction_file, 
            face_csv_path=face_csv_path,  # 传递面数据CSV路径
            is_semi_finish=False
        )
        
        # 处理螺旋数据（精加工）
        if full_normal_result:
            print("\n步骤6: 处理螺旋数据（侧壁等高且封闭）- 精加工")
            print("-" * 40)
            self.print_enhanced_result(full_normal_result, "螺旋 (侧壁等高且封闭) - 精加工")
            self.save_result_to_separate_json(
                result=full_normal_result,
                half_finish_output=output_half_spiral,  # 精加工时这个参数不会被使用
                full_finish_output=output_full_spiral,
                operation_name="D4-螺旋_SIMPLE",
                prefix="",
                is_semi_finish=False  # 新增参数，表示是精加工
            )
        
        # 处理往复等高数据（精加工）
        if full_half_result:
            print("\n步骤7: 处理往复等高数据（侧壁不等高或开放）- 精加工")
            print("-" * 40)
            self.print_enhanced_result(full_half_result, "往复等高 (侧壁不等高或开放) - 精加工")
            self.save_result_to_separate_json(
                result=full_half_result,
                half_finish_output=output_half_half_spiral,  # 精加工时这个参数不会被使用
                full_finish_output=output_full_half_spiral,
                operation_name="往复等高_SIMPLE",
                prefix="",
                is_semi_finish=False  # 新增参数，表示是精加工
            )
        
        # 5. 输出总结
        print("\n" + "=" * 60)
        print("处理完成！已生成四个JSON文件：")
        print(f"1. 半精螺旋: {output_half_spiral}")
        print(f"2. 全精螺旋: {output_full_spiral}")
        print(f"3. 半精螺旋_往复等高: {output_half_half_spiral}")
        print(f"4. 全精螺旋_往复等高: {output_full_half_spiral}")
        print("=" * 60)


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


def main1(prt_folder, face_data_csv, csv_file, json_file, direction_file, output_dir):
    """
    主入口函数 - 处理单个零件的螺旋加工
    
    Args:
        prt_folder: PRT文件路径 或 文件夹路径
        face_data_csv: 面数据CSV路径（新增：用于判断开放/封闭）
        csv_file: 特征识别日志CSV路径
        json_file: 刀具参数JSON路径
        direction_file: 方向映射CSV路径
        output_dir: 输出目录
    """
    # 处理 prt_folder 参数，支持文件或文件夹
    prt_path_obj = Path(prt_folder)
    if prt_path_obj.is_file():
        prt_files = [prt_path_obj]
    else:
        prt_files = list(prt_path_obj.glob("*.prt"))

    if not prt_files:
        print(f"在路径 {prt_folder} 中未找到任何 .prt 文件")
        return
    
    
    excel_params = r'C:\Projects\NC\output\00_Resources\CSV_Reports\零件参数.xlsx'
    
    part_code = extract_part_name_from_path(str(prt_files[0]))
    
    # 生成输出文件路径
    output_half_spiral = os.path.join(output_dir, f"{part_code}_半精_螺旋.json")
    output_full_spiral = os.path.join(output_dir, f"{part_code}_全精_螺旋.json")
    output_half_half_spiral = os.path.join(output_dir, f"{part_code}_半精_螺旋_往复等高.json")
    output_full_half_spiral = os.path.join(output_dir, f"{part_code}_全精_螺旋_往复等高.json")
    
    print(f"零件代码: {part_code}")
    print(f"输出目录: {output_dir}")
    print(f"面数据CSV: {face_data_csv}")
    print(f"输出文件:")
    print(f"  1. 半精螺旋: {output_half_spiral}")
    print(f"  2. 全精螺旋: {output_full_spiral}")
    print(f"  3. 半精螺旋_往复等高: {output_half_half_spiral}")
    print(f"  4. 全精螺旋_往复等高: {output_full_half_spiral}")
    
    # 创建处理器实例
    processor = SpiralProcessor()
    
    # 执行完整处理流程
    processor.full_processing_pipeline(
        prt_folder=prt_folder,
        excel_params=excel_params,
        csv_file=csv_file,
        json_file=json_file,
        direction_file=direction_file,
        output_half_spiral=output_half_spiral,
        output_full_spiral=output_full_spiral,
        output_half_half_spiral=output_half_half_spiral,
        output_full_half_spiral=output_full_half_spiral,
        face_csv_path=face_data_csv  # 新增：传递面数据CSV路径
    )


def main():
    """示例主函数"""
    
    # 零件名称 - 可以根据需要修改
    part_name = "DIE-03"
    
    # 输入文件路径
    prt_folder = os.path.join(r'C:\Projects\NC\output\prt')
    
    # 使用part_name动态构建CSV和方向文件路径
    face_data_csv = os.path.join(r'C:\Projects\NC\output\03_Analysis\Face_Info\face_csv',  f"{part_name}_face_data.csv")
    csv_file = os.path.join(r'C:\Projects\NC\output\03_Analysis\Navigator_Reports',  f"{part_name}_FeatureRecognition_Log.csv")
    json_file = os.path.join(r'C:\Projects\NC\input\铣刀参数.json')
    direction_file = os.path.join(r'C:\Projects\NC\output\03_Analysis\Geometry_Analysis', f"{part_name}.prt.csv")
    output_dir = r"C:\Projects\NC\output\json"
    
    # 调用主处理函数
    main1(prt_folder, face_data_csv, csv_file, json_file, direction_file, output_dir)


if __name__ == '__main__':
    main()