# -*- coding: utf-8 -*-
"""
参数解析模块
包含加工参数解析、孔属性提取等功能
"""

import re
import copy
import traceback
import drill_config
from drill_library import read_drill_table_json
from utils import print_to_info_window


class ParameterParser:
    """参数解析器"""

    def __init__(self):
        pass

    def init_template(self):
        """初始化孔属性模板"""

        return {
            "num": 1,  # 数量（字符串开头-之前的整数-默认为1）
            "circle_type": None,  # 孔类型
            "is_through_hole": None,  # 是否通孔（布尔值）
            "real_diamater": None,  # 主孔直径（含偏置）
            "real_diamater_head": None,  # 沉头直径（含偏置）
            "real_diamater_threading_hole": None,  # 穿线孔直径（含偏置）
            "luowen": {  # 螺纹信息字典
                "diamater": None,  # 螺径（mm）
                "pitch": None,  # 螺距（mm）
                "direction": None,  # 加工方向（正/背）
                "luowen_depth": None  # 螺纹深度（mm）
            },
            "depth": {  # 深度信息字典
                "head_depth": None,  # 沉头深度（mm）
                "hole_depth": None  # 主孔深度（mm）
            },
            "is_bz": None,  # 能否背面打孔（布尔值）
            "is_zzANDbz": False,  # 是否需要正、背两面钻孔
            "main_hole_processing": "钻",  # 主孔加工方式（默认"钻"）
            "bias": {
                "value": None,  # 公差偏置数值
                "operation": None  # + 或 - 或 None
            },
            "is_side": False,
            "is_divided": False
        }

    def get_depth(self, text, hole_type=None):
        """从加工说明文本中提取孔的深度值"""

        try:
            if hole_type == "沉头":
                pattern = r'沉头<O>\d*\.?\d+深(\d*\.?\d+)'  # 匹配沉头孔深度
            elif hole_type == "螺纹":
                pattern = r"M\d+[xX]?P?\d*\.?\d+(?:[^，。！？；：\"'()（）【】{}、.!?;:(){}[\]]*)深(\d*\.?\d+)"  # 匹配螺纹孔深度
            elif hole_type == "钻深":
                pattern = r"钻深(\d*\.?\d+)"  # 匹配主孔钻孔深度
            else:
                pattern = r"深(\d*\.?\d+)"  # 通用匹配

            match = re.search(pattern, text)
            return float(match.group(1)) if match else None
        except Exception as e:
            err_info = f"❌ 提取深度失败（文本：{text[:20]}...）\n异常类型：{type(e).__name__}\n异常位置：{traceback.extract_tb(e.__traceback__)}\n异常描述：{str(e)}"
            print_to_info_window(err_info)
            return None

    def get_special_bias(self, text, bias):
        """
        根据特殊关键词提取公差偏置
        """
        # "攻" 或者 *.1 或者 *.2
        if "攻" in text or re.fullmatch(r"^.*(\.1|\.2)$", text):
            bias['value'] = 0.0
            return bias
        keywords = ["顶针沉头", "引导针沉头", "弹簧孔"]
        for keyword in keywords:
            if keyword in text:
                bias['value'], bias['operation'] = 1.0, "+"
                return bias
        if "铰" in text:
            bias['value'], bias['operation'] = 0.2, "-"
            return bias
        if "导柱沉头" in text:
            bias['value'], bias['operation'] = 2.0, "+"
        if "沉头" in text:
            if "外导柱" in text:
                bias['value'], bias['operation'] = 4.0, "+"
                return bias
            if "导套" in text or "螺丝" in text or "等高套筒" in text:
                bias['value'], bias['operation'] = 2.0, "+"
                return bias
            if "合销" in text:
                bias['value'], bias['operation'] = 1.0, "+"
                return bias
        if "钻穿" in text and "等高套筒" in text:
            bias['value'], bias['operation'] = 1.0, "+"
            return bias
        if "过孔" in text:
            if "螺丝过孔" in text:
                bias['value'], bias['operation'] = 1.0, "+"
                return bias
            if "导柱过孔" in text:
                bias['value'], bias['operation'] = 2.0, "+"
                return bias
            bias['value'], bias['operation'] = 1.0, "+"
            return bias

        if "顶针孔" in text or "引导孔" in text:
            bias['value'], bias['operation'] = 0.5, "+"
            return bias
        if "穿线孔" in text:
            bias['value'], bias['operation'] = 1.0, "+"
            return bias
        return bias

    def get_bias(self, text):
        """提取公差偏置"""

        bias = {
            "value": None,
            "operation": None
        }  # 公差偏置
        try:
            pattern = r'单([+-]\d+\.?\d*|\.?\d+)'  # 匹配正负整数/小数偏置值
            matches = re.finditer(pattern, text)
            for match in matches:
                start_index = match.start()  # 匹配字符串的首字符索引
                match_str = match.group(1)  # 匹配到的字符串
                if "单+" == text[start_index: start_index + 2]:
                    bias["value"] = 2 * float(match_str) if match_str else 0.0
                    bias["operation"] = "+"
                    return bias

                if "单-" in text[start_index: start_index + 2]:
                    bias["value"] = 2 * float(match_str) if match_str else 0.0
                    bias["operation"] = "-"
                    return bias
            # 当前加工说明不存在像（单+0.01）这样的公差信息
            # 执行特殊关键字匹配获取公差
            bias = self.get_special_bias(text, bias)

            return bias
        except Exception as e:
            err_info = f"❌ 提取偏置失败（文本：{text[:20]}...）\n异常类型：{type(e).__name__}\n异常位置：{traceback.extract_tb(e.__traceback__)}\n异常描述：{str(e)}"
            print_to_info_window(err_info)
            return None

    def get_d(self, text):
        """提取孔的直径"""

        try:
            if len(text) <= 1:  # 文本过短，无有效直径信息
                return None
            index_1 = text.find("<O>")  # 定位直径标识
            if index_1 == -1:  # 无<O>标识，不提取直径
                return None
            # 提取<O>后第一个整数/小数作为基准直径
            matches = re.findall(r'\d*\.?\d+', text[index_1 + 2:])
            d = float(matches[0]) if matches else None
            return d
        except Exception as e:
            err_info = f"❌ 提取直径失败（文本：{text[:20]}...）\n异常类型：{type(e).__name__}\n异常位置：{traceback.extract_tb(e.__traceback__)}\n异常描述：{str(e)}"
            print_to_info_window(err_info)
            return None

    def get_circle_attribute(self, text, template, count, _size):
        """解析加工说明文本，提取孔的完整属性"""

        try:
            _template = copy.deepcopy(template)  # 深拷贝，避免修改原始模板
            # 提取公差偏置
            _template['bias'] = self.get_bias(text)
            if "侧" in text:
                _template["is_side"] = True
            # 1. 解析螺纹规格（如M12xP1.75、M10P1.5）
            r = r"M\d+[xX]?P\d*\.?\d+"  # 匹配螺纹正则（兼容x/X分隔螺径和螺距）
            match = re.search(r, text)
            if match:
                luowen_depth = self.get_depth(text, "螺纹")  # 提取螺纹深度
                _template["luowen"]["luowen_depth"] = luowen_depth
                # 提取到螺纹深度
                if luowen_depth:
                    # 若是螺纹孔：主孔深度 = 螺纹深度 + 5mm
                    _template["depth"]["hole_depth"] = luowen_depth + 5.0
                else:
                    pattern = r"攻深(\d*\.?\d+)"  # 匹配主孔钻孔深度
                    match_1 = re.search(pattern, text)
                    if match_1:
                        luowen_depth = float(match_1.group(1))
                        _template["luowen"]["luowen_depth"] = luowen_depth
                        _template["depth"]["hole_depth"] = luowen_depth + 5.0
                if not _template["luowen"]["luowen_depth"] and "攻穿" in text and _size:
                    _template["luowen"]["luowen_depth"] = _size[0] if _template['is_side'] else _size[2]

                idx_r = text.find(match[0])  # 螺纹规格在文本中的起始位置
                idx_l = text.find(",") if "," in text else 0  # 螺纹规格前的分隔符位置
                # 判断螺纹加工方向（正/背）
                if "正" in text[idx_l: idx_r]:
                    _template['luowen']['direction'] = "正"
                else:
                    _template['luowen']['direction'] = "背"
                # 提取螺径（第一个数值）和螺距（第二个数值）
                matches = re.findall(r'\d+\.?\d*', match.group(0))
                if len(matches) >= 1:
                    _template['luowen']['diamater'] = int(matches[0])
                if len(matches) >= 2:
                    _template['luowen']['pitch'] = float(matches[1])

            # 2. 解析主孔加工方式
            if "割" in text:
                _template["main_hole_processing"] = "割"  # 线切割加工
            if "铰" in text:
                _template["main_hole_processing"] = "铰"  # 铰孔加工

            # 3. 沉头孔特殊处理（区分正/背沉头，提取沉头直径、深度）
            if "沉头" in text:
                if count["keys"] and _template["is_through_hole"]:
                    count["keys"].pop(-1)
                # 判断沉头方向（正/背）
                if "背沉头" in text or "背面沉头" in text or "背面" in text:
                    circle_type = "背沉头"
                    _template["is_bz"] = True
                    count["back"] += 1
                elif "正沉头" in text or "正面沉头" in text or "正面" in text:
                    circle_type = "正沉头"
                    _template["is_bz"] = False
                    count["correct"] += 1

                elif _template["is_through_hole"] and "背面" in text:
                    circle_type = "背沉头"
                    _template["is_bz"] = False
                    count["correct"] += 1
                else:
                    circle_type = "正沉头"
                    _template["is_bz"] = False
                    count["correct"] += 1

                # 提取沉头属性
                match = re.search(r'<O>\d+(\.\d+)?x\d+(\.\d+)?', text)

                if match:
                    _template['depth']['head_depth'] = match.group(0).split('x')[-1]
                else:
                    match = re.search(r'沉头<O>\d+(\.\d+)?深\d+(\.\d+)?', text)
                    _template['depth']['head_depth'] = self.get_depth(text, "沉头")

                # 正则匹配<O>后的浮点数，捕获组1为目标浮点数
                pattern = r'<O>(\d+(\.\d+)?)'
                # findall返回元组列表，每个元组的第一个元素是浮点数字符串
                matches = re.findall(pattern, text)
                # 提取并转换为float列表
                float_list = [float(match[0]) for match in matches]
                l = len(float_list)
                if l > 1:
                    _template['real_diamater_head'] = max(float_list)  # 沉头直径
                    _template['real_diamater'] = min(float_list)  # 主孔直径
                elif l == 1:
                    if match:
                        _template['real_diamater_head'] = float_list[0]  # 沉头直径
                    else:
                        _template['real_diamater'] = float_list[0]  # 主孔直径

                _template['circle_type'] = circle_type
                _template["depth"]["hole_depth"] = self.get_depth(text, "钻深")  # 主孔深度
                return _template, count
            else:
                # 4. 非沉头孔处理（螺纹孔/入子孔/弹簧孔/一般孔）
                if match:
                    _template["circle_type"] = "螺纹孔"
                elif "入子孔" in text:
                    _template["circle_type"] = "入子孔"
                elif "弹簧孔" in text:
                    _template["circle_type"] = "弹簧孔"
                else:
                    _template["circle_type"] = "一般孔"

                if not match:
                    # 提取一般孔深度（区分"钻深"和通用"深"）
                    _template["depth"]["hole_depth"] = self.get_depth(text,
                                                                      "钻深") if "钻深" in text else self.get_depth(
                        text)
                # 5. 穿线孔处理（分离穿线孔直径和主孔直径）
                idx1 = text.find("穿线孔")
                if idx1 != -1:
                    _template['real_diamater_threading_hole'] = self.get_d(text[idx1:])  # 穿线孔直径
                    _template['real_diamater'] = self.get_d(text[:idx1])  # 主孔直径
                else:
                    _template['real_diamater'] = self.get_d(text)  # 普通孔直径
            return _template, count
        except Exception as e:
            err_info = f"❌ 提取孔属性失败（文本：{text[:20]}...）\n异常类型：{type(e).__name__}\n异常位置：{traceback.extract_tb(e.__traceback__)}\n异常描述：{str(e)}"
            print_to_info_window(err_info)
            return copy.deepcopy(template), count  # 异常时返回空模板

    def get_is_bz_ThroughHole(self, text):
        """判断是否支持背面打孔和是否为通孔"""

        try:
            # 判断是否为通孔
            if any(keyword in text for keyword in
                   ["钻穿", "攻穿", "割", "铰通", "通", "thru", "through", "贯穿", "透", "镗", "塘", "搪"]):
                return False, True
            # 正钻背攻/背钻正攻/正钻背沉头/背钻正沉头
            elif "正" in text and "背" in text:
                return False, True
            # 判断是否支持背面打孔
            if "深" in text and "背" in text:
                return True, False  # (背面打孔, 非通孔)
            return False, False  # (非背面打孔, 非通孔)
        except Exception as e:
            err_info = f"❌ 判断通孔/背面打孔失败（文本：{text[:20]}...）\n异常类型：{type(e).__name__}\n异常位置：{traceback.extract_tb(e.__traceback__)}\n异常描述：{str(e)}"
            print_to_info_window(err_info)
            return False, False

    def is_both_sides_drill(self, datas, _template, _size, datas_dict):
        depth = _size[0] if _template['is_side'] else _size[2]
        # 获取材质
        _material = datas_dict["材质"]
        material = ""
        if len(_material) > 0:
            material = _material[0]
        diameter = _template["real_diamater"]
        if not diameter:
            diameter = _template["real_diamater_threading_hole"]
        if diameter and _template["is_through_hole"] and "沉头孔" not in f'{_template["circle_type"]}':
            max_diameter = 0
            length = 0
            for data in datas:
                if material == data["材质"] and round(data["直径"], 1) == round(diameter, 1):
                    if data["直径"] <= max_diameter:
                        continue
                    max_diameter = data["直径"]
                    length = data["有效长度"]
            if depth and length and length < depth:
                _template["is_zzANDbz"] = True
                _template["depth"]["hole_depth"] = depth / 2
            else:
                if not _template["depth"]["hole_depth"]:
                    _template["depth"]["hole_depth"] = depth
        return _template

    def process_hole_data(self, datas_dicts):
        """主处理函数：解析孔加工信息"""

        print_to_info_window("=" * 50)
        print_to_info_window("开始处理孔加工数据...")
        print_to_info_window("=" * 50)

        # 初始化孔属性模板
        template = self.init_template()
        result = {}  # 存储最终格式化结果
        datas = read_drill_table_json(drill_config.DRILL_JSON_PATH)  # 初始化钻刀库
        is_divided = False

        try:
            # 统计总注释数量
            total = len(datas_dicts["主体加工说明"])
            processed = 0  # 成功处理数量
            skipped = 0  # 跳过数量
            print_to_info_window(f"\n开始解析数据（共{total}个key）...")

            # 统计正背沉头数量的字典,"keys": 通孔key列表L、W、M...
            count = {"correct": 0, "back": 0, "keys": []}
            # 遍历每条注释进行解析
            datas_dict = datas_dicts["主体加工说明"]
            diameter_map = {"M3": 2.5, "M4": 3.3, "M5": 4.2, "M6": 5.0, "M8": 6.8, "M10": 8.5, "M12": 10.3,
                            "M16": 14.0, "M20": 17.5, "M24": 21.0, "M30": 26.0,
                            }
            str_dict = str(datas_dict)
            if "分中加工" in str_dict:
                is_divided = True
            elif "分中" in str_dict:
                is_divided = True
            elif "中加工" in str_dict:
                is_divided = True
            for key in datas_dict:
                try:
                    # 1. 跳过不需要处理的key
                    if key in ["其它注释", "加工对象", "尺寸", "重量", "数量", "热处理", "材质"]:
                        skipped += 1
                        continue
                    text = str(datas_dict[key])
                    # 跳过含"槽"、"让位"、"导轨面"的加工项
                    if any(keyword in text for keyword in ["槽", "让位", "导轨面"]):
                        skipped += 1
                        continue
                    # 排除标注了暂不加工的孔
                    index = text.find("暂不加工")
                    if index != -1:
                        left_text = text[:index]
                        right_text = text[index:]
                        left_idx = left_text.rfind(',')
                        right_idx = right_text.find(',')
                        len_left_text = len(left_text)
                        if left_idx != -1:
                            text = text[: left_idx + 1] + (
                                text[right_idx + len_left_text + 1:] if right_idx != -1 else "")

                        else:
                            idx = left_text.find('-')
                            if idx != -1:
                                text = text[: idx + 1] + (
                                    text[right_idx + len_left_text + 1:] if right_idx != -1 else "")

                    # 2. 解析单条孔加工注释
                    _template = copy.deepcopy(template)

                    # 提取数量
                    idx = text.find("-")
                    if idx != -1 and text[:idx].strip().isdigit():
                        _template['num'] = int(text[:idx].strip())
                        text = text[idx + 1:].strip()

                    pattern = r'(\d+(?:\.\d+)?)L[x*](\d+(?:\.\d+)?)W[x*](\d+(?:\.\d+)?)T'
                    match = re.search(pattern, datas_dicts["主体加工说明"]["尺寸"] if datas_dicts["主体加工说明"][
                        "尺寸"] else f"{datas_dicts['主体加工说明']}")
                    # 长宽厚
                    _size = []
                    if match:
                        _size = [float(match.group(1)), float(match.group(2)), float(match.group(3))]


                    # 判断是否背面打孔、是否为通孔
                    _template['is_bz'], _template["is_through_hole"] = self.get_is_bz_ThroughHole(text)
                    # 通孔，存入通孔keys列表
                    if _template["is_through_hole"]:
                        count["keys"].append(key)

                    # 提取铣孔属性
                    mill_match = re.search(r'<O>(\d+(\.\d+))', text)
                    if mill_match and "铣" in text and not _template["is_through_hole"]:
                        if "精铣" in text:
                            _template["main_hole_processing"] = "一般孔精铣"
                        else:
                            _template["main_hole_processing"] = "一般孔铣"
                        depth_mill_match = re.search(r'铣<O>\d*\.?\d+深(\d*\.?\d+)', text)
                        mill_depth = float(depth_mill_match.group(1)) if depth_mill_match else self.get_depth(text,
                                                                                                              "铣深")
                        if not mill_depth:
                            mill_depth = re.search(r'深\d+(\.\d+)', text)
                            if mill_depth:
                                mill_depth = mill_depth.group(0)
                        _template["depth"]["hole_depth"] = mill_depth
                        _template["real_diamater"] = float(mill_match.group(1))
                        # 正反面判断
                        if "背" in text:
                            _template["mill_direction"] = "背"
                        if "深度准" in text and _template["depth"]["hole_depth"]:
                            _template["depth"]["hole_depth"] -= 0.2
                        if "侧" in text:
                            _template["is_side"] = True

                        # 是否双面钻孔
                        _template = self.is_both_sides_drill(datas, _template, _size, datas_dict)
                        result[key] = _template
                        processed += 1
                        continue
                    # 提取孔属性并更新模板
                    _template, count = self.get_circle_attribute(text, _template, count, _size)

                    # 存在"深度准", 则钻孔深度 - 0.5mm
                    if "深度准" in text and _template["depth"]["hole_depth"]:
                        _template["depth"]["hole_depth"] -= 0.5

                    if "背攻深" in text and _template["luowen"]["luowen_depth"]:
                        _template["luowen"]["luowen_depth"] += 0.5

                    # 根据螺纹直径获取孔直径
                    m = re.findall(r'M\d+', text)
                    if _template["real_diamater"] is None and m:
                        if m[0] in diameter_map:
                            _template["real_diamater"] = diameter_map[m[0]]
                    # 是否双面钻孔
                    _template = self.is_both_sides_drill(datas, _template, _size, datas_dict)
                    # 存入结果字典
                    result[key] = _template
                    processed += 1

                except Exception as e:
                    # 单条key解析失败，记录异常并继续处理下一条
                    err_info = f"❌ 处理key[{key}]时失败\n异常类型：{type(e).__name__}\n异常位置：{traceback.extract_tb(e.__traceback__)}\n异常描述：{str(e)}"
                    print_to_info_window(err_info)
                    continue
            # 判断-背沉头>正沉头且二者不同为0，所有非沉头通孔设置为背面打孔
            if count["correct"] or count["back"]:
                for key in count["keys"]:
                    if (count["back"] - count["correct"]) >= 0:
                        result[key]["is_bz"] = True
                    else:
                        result[key]["is_bz"] = False

            return result, is_divided

        except Exception as e:
            # 主流程异常
            err_info = f"❌ 主流程处理失败\n异常类型：{type(e).__name__}\n异常位置：{traceback.extract_tb(e.__traceback__)}\n异常描述：{str(e)}"
            print_to_info_window(err_info)
            return None, None
