# -*- coding: utf-8 -*-
"""
钻刀库模块
包含钻孔参数查询、材质匹配等功能
"""

import json
import os
from utils import print_to_info_window
import drill_config

class DrillLibrary:
    """钻刀库管理类"""

    def __init__(self, json_path=None):
        self.json_path = json_path or drill_config.DRILL_JSON_PATH
        self.data = self._load_json_data()

    def _load_json_data(self):
        """读取 JSON 文件内容"""

        if not os.path.exists(self.json_path):
            print_to_info_window(f"❌ 文件不存在：{self.json_path}")
            return []

        try:
            with open(self.json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            print_to_info_window(f"✅ 已成功读取 JSON 文件，共 {len(data)} 条记录")
            return data
        except Exception as e:
            print_to_info_window(f"❌ 读取 JSON 文件失败: {e}")
            return []

    def get_drill_parameters(self, material_type, diameter, dtype, have_mill=""):
        """根据材质和直径获取钻孔参数"""

        if not self.data:
            return None, None, None, None

        # 标记（用于判断数据库中是否存在相应材料的参数）
        found_material = 0
        # 标记2（用于判断数据库中是否有相应的刀具）
        found_tool = 0

        tool_name = None
        step_distance = 0.0
        feed_rate = 40.0
        cycle_type = "Drill,Deep"
        # 判断是钻孔还是铰孔
        if (dtype == "钻" or (dtype != "铰" and dtype != "铣")):
            expression = "1-2"
            if have_mill:
                operator = "<" if "精铣" in have_mill else "<="
                expression = f'round(drill_parse["直径"], 1) {operator} round(diameter, 1) and "U" in drill_parse["钻头名"]'
            for drill_parse in self.data:
                if material_type in drill_parse["材质"]:
                    found_material += 1
                    if have_mill:
                        if eval(expression):
                            tool_name = drill_parse["钻头名"]
                            step_distance = drill_parse["step"]
                            feed_rate = drill_parse["进给"]

                            if drill_parse["钻孔模式"] == "深孔":
                                cycle_type = drill_config.CYCLE_DRILL_DEEP
                            elif drill_parse["钻孔模式"] == "标准钻":
                                cycle_type = drill_config.CYCLE_DRILL_STANDARD

                                found_tool += 1
                    else:
                        if round(drill_parse["直径"], 1) == round(diameter, 1) and ("A" in drill_parse["钻头名"] or "U" in drill_parse["钻头名"]):
                            tool_name = drill_parse["钻头名"]
                            step_distance = drill_parse["step"]
                            feed_rate = drill_parse["进给"]

                            if drill_parse["钻孔模式"] == "深孔":
                                cycle_type = drill_config.CYCLE_DRILL_DEEP
                            elif drill_parse["钻孔模式"] == "标准钻":
                                cycle_type = drill_config.CYCLE_DRILL_STANDARD

                                found_tool += 1
        elif (dtype == "铰"):
            for drill_parse in self.data:
                if material_type in drill_parse["材质"]:
                    found_material += 1
                    if round(drill_parse["直径"], 1) == round(diameter, 1) and "J" in drill_parse["钻头名"]:
                        tool_name = drill_parse["钻头名"]
                        step_distance = drill_parse["step"]
                        feed_rate = drill_parse["进给"]

                        if drill_parse["钻孔模式"] == "深孔":
                            cycle_type = drill_config.CYCLE_DRILL_DEEP
                        elif drill_parse["钻孔模式"] == "标准钻":
                            cycle_type = drill_config.CYCLE_DRILL_STANDARD

                        found_tool += 1
        # 检查结果
        if found_material == 0:
            print_to_info_window(f"数据库中缺少{material_type}材质的相关参数")
            return None, None, None, None

        if found_tool == 0:
            tool_name = f"A_{round(diameter,1)}"
            print_to_info_window(f"刀具库中缺少相应的刀具。所需钻刀直径为：{diameter}")
        else:
            print_to_info_window(f"刀具库中存在相应的刀具。刀具名为：{tool_name}")

        return tool_name, step_distance, feed_rate, cycle_type

    def get_mrill_parameters(self, material_type, diameter,deviation, processing_method=None):
        data = read_drill_table_json(drill_config.MILL_JSON_PATH)
        if not data:
            return None, None, None, None
        # 标记（用于判断数据库中是否存在相应材料的参数）
        found_material = 0
        # 标记2（用于判断数据库中是否有相应的刀具）
        found_tool = 0
        tool_name = None
        R1 = 0.0
        mill_depth = 0.3

        max_diameter = 0.0
        for mill_parse in data:
            if material_type in mill_parse["材质"]:
                found_material += 1
                # 一般孔刀具参数获取
                if processing_method:
                    operation = f'round(diameter, 1) - round(mill_parse["直径"], 1) {">" if processing_method == "精铣" else ">="} round(mill_parse["直径"], 1) * 0.7'
                    if eval(operation):
                        if mill_parse["直径"] > max_diameter:
                            max_diameter = mill_parse["直径"]
                            tool_name = mill_parse["刀具名"]
                            R1 = mill_parse["R角"]
                            mill_depth = mill_parse["步距"]
                    continue
                # 沉头孔刀具参数获取
                if round(diameter, 1) - round(mill_parse["直径"], 1) > round(mill_parse["直径"], 1) * 0.7:
                    if mill_parse["直径"] > max_diameter:
                        max_diameter = mill_parse["直径"]
                        tool_name = mill_parse["刀具名"]
                        R1 = mill_parse["R角"]
                        mill_depth = mill_parse["步距"]
                # if round(mill_parse["直径"], 1) < round(diameter, 1) and round(mill_parse["直径"], 1) > round(deviation, 1) and "D" in mill_parse["刀具名"]:
                #     if mill_parse["直径"] > max_diameter:
                #         max_diameter = mill_parse["直径"]
                #         tool_name = mill_parse["刀具名"]
                #         R1 = mill_parse["R角"]
                #         mill_depth = mill_parse["步距"]
                        found_tool += 1
        diameter = max_diameter
        # 检查结果
        if found_material == 0:
            print_to_info_window(f"数据库中缺少{material_type}材质的相关参数")
            return None, None, None

        if found_tool == 0:
            tool_name = f"D_{diameter}"
            print_to_info_window(f"刀具库中缺少相应的刀具。所需铣刀直径为：{diameter}")
        else:
            print_to_info_window(f"刀具库中存在相应的刀具。刀具名为：{tool_name}")

        return diameter, tool_name, R1, mill_depth

    def validate_material(self, material_type):
        """验证材质是否在数据库中"""

        if not self.data:
            return False

        for drill_parse in self.data:
            if material_type in drill_parse["材质"]:
                return True

        return False

    def get_tool_name(self, diameter):
        """根据直径生成默认刀具名称"""
        return f"A_{diameter}"


def read_drill_table_json(json_path=None):
    """兼容函数：读取钻孔参数表"""

    json_path = json_path or drill_config.DRILL_JSON_PATH
    if not os.path.exists(json_path):
        print_to_info_window(f"❌ 文件不存在：{json_path}")
        return []

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        print_to_info_window(f"✅ 已成功读取 JSON 文件，共 {len(data)} 条记录")
        return data
    except Exception as e:
        print_to_info_window(f"❌ 读取 JSON 文件失败: {e}")
        return []


