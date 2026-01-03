#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kiabi9.0 类封装版本 - 模拟生成爬面文件的结构
"""

import os
import csv
import json
import sys
import re
from collections import defaultdict, deque
from typing import Dict, List, Tuple
import pandas as pd
from pathlib import Path
import math


class KiabiProcessor:
    """
    Kiabi 9.0 类封装版 - 生成半精和全精加工JSON文件
    """

    def __init__(self):
        """初始化配置"""
        self.CONFIG = {
            # 输入文件路径
            "CSV_FACE": "",
            "CSV_TAG": "",
            "OUT_DIR": "",

            # 输出文件名（运行时根据PRT文件名生成）
            "OUT_JSON_HALF": "",
            "OUT_JSON_FULL": "",

            # 材质和刀具参数
            "PRT_FOLDER": "",
            "EXCEL_PARAMS": "",
            "TOOL_JSON": "",

            # 零件信息
            "PART_NAME": "",  # 零件名（不含.prt后缀）

            # 图层映射
            "LAYER_MAP": {"+Z": 20, "+Y": 60, "-Y": 50, "-X": 30, "+X": 40, "-Z": 70},
            "COL_NAMES": ['+Z', '-Z', '+X', '-X', '+Y', '-Y'],

            # 全局变量
            "MIN_LENGTH_THRESHOLD": 4.0,  # 最小长度阈值(mm)
            "CLUSTER_TOLERANCE": 5.0,  # 开放面聚类容忍度(mm)
        }

        # 内部缓存
        self._MATERIAL_DF = None
        self._TOOL_DF = None
        self.tag2row = {}
        self.col_tags = defaultdict(list)
        self.result = {}

    def setup_paths(self, csv_face: str, csv_tag: str, out_dir: str,
                    prt_folder: str, excel_params: str, tool_json: str):
        """
        设置所有路径参数

        Args:
            csv_face: 面数据CSV路径
            csv_tag: 标签数据CSV路径
            out_dir: 输出目录
            prt_folder: PRT文件所在文件夹
            excel_params: 零件参数Excel文件路径
            tool_json: 刀具参数JSON文件路径
        """
        self.CONFIG["CSV_FACE"] = csv_face
        self.CONFIG["CSV_TAG"] = csv_tag
        self.CONFIG["OUT_DIR"] = out_dir
        self.CONFIG["PRT_FOLDER"] = prt_folder
        self.CONFIG["EXCEL_PARAMS"] = excel_params
        self.CONFIG["TOOL_JSON"] = tool_json

        # 获取零件名（不含.prt后缀）
        self._extract_part_name()

        # 根据零件名生成输出文件名
        if self.CONFIG["PART_NAME"]:
            self.CONFIG["OUT_JSON_HALF"] = f"{self.CONFIG['PART_NAME']}_半精_面铣.json"
            self.CONFIG["OUT_JSON_FULL"] = f"{self.CONFIG['PART_NAME']}_全精_面铣.json"
        else:
            self.CONFIG["OUT_JSON_HALF"] = "半精_面铣.json"
            self.CONFIG["OUT_JSON_FULL"] = "全精_面铣.json"

        # 确保输出目录存在
        os.makedirs(out_dir, exist_ok=True)

    def _extract_part_name(self):
        """从PRT文件夹或文件路径中提取零件名（不含.prt后缀）"""
        path_input = Path(self.CONFIG["PRT_FOLDER"])

        if not path_input.exists():
            print(f"[WARN] 零件路径不存在: {path_input}")
            return

        if path_input.is_file():
            # 如果是文件，直接使用
            if path_input.suffix.lower() == '.prt':
                self.CONFIG["PART_NAME"] = path_input.stem
                print(f"[INFO] 零件名: {self.CONFIG['PART_NAME']}")
            else:
                print(f"[WARN] 输入文件不是 .prt 文件: {path_input}")
        else:
            # 如果是目录，则查找PRT文件
            prt_files = list(path_input.glob("*.prt"))
            if not prt_files:
                print("[WARN] 未找到任何 .prt 文件")
                return

            if len(prt_files) > 1:
                print(f"[WARN] 发现多个 prt 文件，使用第一个: {prt_files[0].name}")

            # 提取零件名（不含.prt后缀）
            part_name = prt_files[0].stem
            self.CONFIG["PART_NAME"] = part_name
            print(f"[INFO] 零件名: {part_name}")

    def _load_material_excel(self, excel_path: str | Path) -> pd.DataFrame:
        """加载零件参数表"""
        if self._MATERIAL_DF is None:
            try:
                df = pd.read_excel(excel_path)
                print(f"[INFO] 成功加载零件参数表，共 {len(df)} 行")
                self._MATERIAL_DF = df
            except Exception as e:
                print(f"[ERROR] 读取零件参数表失败: {e}")
                self._MATERIAL_DF = pd.DataFrame()
        return self._MATERIAL_DF

    def get_material_from_filename(self) -> Tuple[str, bool]:
        """
        从PRT文件名获取材质和热处理信息

        Returns:
            tuple: (材质, 是否热处理)
        """
        path_input = Path(self.CONFIG["PRT_FOLDER"])
        excel_path = Path(self.CONFIG["EXCEL_PARAMS"])

        if not path_input.exists():
            print(f"[WARN] 零件路径不存在: {path_input}")
            return "45#", False

        target_file = None
        if path_input.is_file():
            target_file = path_input
        else:
            prt_files = list(path_input.glob("*.prt"))
            if not prt_files:
                print("[WARN] 未找到任何 .prt 文件，使用默认材质 45#")
                return "45#", False
            if len(prt_files) > 1:
                print(f"[WARN] 发现多个 prt 文件，使用第一个: {prt_files[0].name}")
            target_file = prt_files[0]

        filename = target_file.stem
        print(f"[INFO] 检测到零件文件: {target_file.name}")

        match = re.match(r"([A-Z]+-\d+)", filename.upper())
        prefix = match.group(1) if match else filename.upper().split('_')[0]
        print(f"[INFO] 解析零件编号: {prefix}")

        df = self._load_material_excel(excel_path)
        if df.empty:
            print("[WARN] 零件参数表为空，使用默认材质 45#")
            return "45#", False

        mask = df.astype(str).apply(
            lambda col: col.str.contains(prefix, case=False, na=False)
        ).any(axis=1)
        matched = df[mask]

        if matched.empty:
            print(f"[WARN] 未在零件参数表中找到编号包含 '{prefix}' 的记录，使用默认材质 45#")
            return "45#", False

        row = matched.iloc[0]
        material = None

        # 提取材质
        for col in ["材质"]:
            if col in row and pd.notna(row[col]):
                material = str(row[col]).strip()
                break

        if not material:
            material = "45#"

        # 检查热处理
        is_heat = False
        for col in ["热处理"]:
            if col in row and pd.notna(row[col]):
                ht_text = str(row[col]).strip()
                if ht_text and ht_text.lower() not in ["无", "否", "-", ""]:
                    is_heat = True
                    print(f"[INFO] 检测到热处理: {ht_text}")
                    break

        print(f"[INFO] 材质: {material} | 是否热处理: {is_heat}")
        return material, is_heat

    def read_tool_parameters_from_json(self) -> pd.DataFrame:
        """
        从JSON文件读取刀具参数

        Returns:
            pd.DataFrame: 刀具参数表
        """
        if self._TOOL_DF is None:
            try:
                # 读取JSON文件
                with open(self.CONFIG["TOOL_JSON"], 'r', encoding='utf-8') as f:
                    tool_data = json.load(f)

                # 转换为DataFrame
                df = pd.DataFrame(tool_data)

                # 需要的列
                need_columns = [
                    '刀具名称', '直径', '类别', 'R角',
                    '45#,A3,切深', 'CR12热处理前切深', 'CR12热处理后切深',
                    'CR12mov,SKD11,SKH-9,DC53,热处理前切深',
                    'CR12mov,SKD11,SKH-9,DC53,热处理后切深',
                    'P20切深', 'TOOLOX33 TOOLOX44切深', '合金铜切深',
                    '横越(普)', '进给(普)', '转速(普)'
                ]

                # 只保留存在的列
                available_columns = [c for c in need_columns if c in df.columns]
                if available_columns:
                    df = df[available_columns]
                else:
                    print(f"[WARN] 刀具JSON文件中未找到任何需要的列")
                    df = pd.DataFrame()

                # 确保关键列存在
                if '刀具名称' not in df.columns or '直径' not in df.columns:
                    print(f"[ERROR] 刀具JSON文件缺少必要列（刀具名称或直径）")
                    df = pd.DataFrame()
                else:
                    # 清理数据
                    df = df.dropna(subset=['刀具名称', '直径'])
                    df['直径'] = pd.to_numeric(df['直径'], errors='coerce')
                    df = df.dropna(subset=['直径'])

                    # 确保类别和R角列有默认值
                    if '类别' not in df.columns:
                        df['类别'] = '未知'
                    if 'R角' not in df.columns:
                        df['R角'] = 0.0
                    else:
                        df['R角'] = pd.to_numeric(df['R角'], errors='coerce').fillna(0.0)

                self._TOOL_DF = df
                print(f"[INFO] 成功从JSON读取 {len(self._TOOL_DF)} 把刀具")
                print(f"[INFO] 刀具类别分布: {df['类别'].value_counts().to_dict()}")

            except FileNotFoundError:
                print(f"[ERROR] 刀具JSON文件不存在: {self.CONFIG['TOOL_JSON']}")
                self._TOOL_DF = pd.DataFrame()
            except json.JSONDecodeError as e:
                print(f"[ERROR] 刀具JSON文件格式错误: {e}")
                self._TOOL_DF = pd.DataFrame()
            except Exception as e:
                print(f"[ERROR] 读取刀具JSON文件失败: {e}")
                self._TOOL_DF = pd.DataFrame()

        return self._TOOL_DF

    def get_cutting_depth(self, tool_row, material: str, is_heat: bool) -> float:
        """根据材质和热处理状态获取切深"""
        mapping = {
            '45#': '45#,A3,切深', 'A3': '45#,A3,切深',
            'CR12': 'CR12热处理后切深' if is_heat else 'CR12热处理前切深',
            'CR12MOV': 'CR12mov,SKD11,SKH-9,DC53,热处理后切深' if is_heat else 'CR12mov,SKD11,SKH-9,DC53,热处理前切深',
            'SKD11': 'CR12mov,SKD11,SKH-9,DC53,热处理后切深' if is_heat else 'CR12mov,SKD11,SKH-9,DC53,热处理前切深',
            'SKH-9': 'CR12mov,SKD11,SKH-9,DC53,热处理后切深' if is_heat else 'CR12mov,SKD11,SKH-9,DC53,热处理前切深',
            'DC53': 'CR12mov,SKD11,SKH-9,DC53,热处理后切深' if is_heat else 'CR12mov,SKD11,SKH-9,DC53,热处理前切深',
            'P20': 'P20切深',
            'TOOLOX33': 'TOOLOX33 TOOLOX44切深', 'TOOLOX44': 'TOOLOX33 TOOLOX44切深',
            'T00L0X33': 'TOOLOX33 TOOLOX44切深', 'T00L0X44': 'TOOLOX33 TOOLOX44切深',
            '合金铜': '合金铜切深'
        }

        if material and material.upper() in mapping:
            col = mapping[material.upper()]
            if col in tool_row and pd.notna(tool_row[col]):
                try:
                    value = tool_row[col]
                    if isinstance(value, str) and value.upper() == "NONE":
                        return float(tool_row.get('45#,A3,切深', 0.1))
                    return float(value)
                except (ValueError, TypeError):
                    pass

        # 默认值
        return float(tool_row.get('45#,A3,切深', 0.1))

    def select_tool_and_depth(self, material: str, is_heat: bool, sr: float,
                              is_open: bool, tools: pd.DataFrame) -> Tuple[str, float, float, float, float, float, str]:
        """
        智能选刀和切深（修改后的规则）

        Args:
            material: 材质
            is_heat: 是否热处理
            sr: 面最短边长度
            is_open: 是否是开放面
            tools: 刀具数据表

        Returns:
            tuple: (刀具名称, 直径, 切深, 转速, 进给, 横越, 刀具类别)
        """
        # 复制一份，避免修改原数据
        tools = tools.copy()



        # 规则1：过滤掉钨钢球刀
        if '类别' in tools.columns:
            original_count = len(tools)
            tools = tools[tools['类别'] != '钨钢球刀'].copy()
            print(f"[INFO] 过滤掉钨钢球刀后剩余 {len(tools)} 把刀具 (过滤前: {original_count})")

        # 规则2：如果是否热处理True，过滤掉飞刀
        if is_heat and '类别' in tools.columns:
            original_count = len(tools)
            tools = tools[~tools['类别'].str.contains('飞刀')].copy()
            print(f"[INFO] 热处理材料，过滤掉飞刀后剩余 {len(tools)} 把刀具 (过滤前: {original_count})")
        # 新增规则：过滤掉钨钢牛鼻刀
        if '类别' in tools.columns:
            original_count = len(tools)
            tools = tools[tools['类别'] != '钨钢牛鼻刀'].copy()
            print(f"[INFO] 过滤掉钨钢牛鼻刀后剩余 {len(tools)} 把刀具 (过滤前: {original_count})")

        # 按直径排序
        tools["直径"] = pd.to_numeric(tools["直径"], errors="coerce")
        tools = tools.dropna(subset=["直径"]).sort_values("直径")

        if len(tools) == 0:
            print(f"[WARN] 没有可用的刀具")
            # 返回默认值
            return "D10", 10.0, 0.1, 3000.0, 2800.0, 8000.0, "钨钢平刀"



        # ---------- 1. 全局预过滤：硬料直接踢飞刀 ----------
        if is_heat:  # 热后硬料
            tools = tools[~tools['类别'].str.contains('飞刀', na=False)]

        if is_open:
            if is_heat:
                # 硬料固定 ϕ10
                dia = 10.0
                cand = tools[tools["直径"] == dia].copy()
                if cand.empty:
                    tools["_diff"] = (tools["直径"] - dia).abs()
                    cand = tools.nsmallest(1, "_diff").copy()

                # 如果直径相同，优先选择类别为"钨钢平刀"且R角为0的刀具
                if len(cand) > 1:
                    cand.loc[:, "_priority"] = cand.apply(
                        lambda row: 0 if (row.get('类别') == '钨钢平刀' and row.get('R角', 0) == 0) else 1,
                        axis=1
                    )
                    cand = cand.sort_values(["_priority", "直径"])

                row = cand.iloc[0]
            else:
                # 软料 - 修改后的规则
                if sr >= 40:
                    # SR>=40mm，目标刀具17R0.8
                    target_dia = 17.0
                    print(f"[INFO] 开放面软料SR≥40mm，面最短边: {sr}mm, 目标刀具: 17R0.8")

                    # 首先尝试找17R0.8的刀具
                    cand = tools[
                        (tools["直径"] == target_dia) &
                        (tools["R角"] == 0.8)
                        ].copy()

                    if cand.empty:
                        # 如果没有17R0.8，找直径17的其他刀具
                        print(f"[INFO] 未找到17R0.8刀具，寻找直径17的其他刀具")
                        cand = tools[tools["直径"] == target_dia].copy()

                        if cand.empty:
                            # 如果没有直径17的刀具，找最接近17的刀具
                            print(f"[INFO] 未找到直径17的刀具，寻找最接近17mm的刀具")
                            tools["_diff"] = (tools["直径"] - target_dia).abs()
                            cand = tools.nsmallest(3, "_diff").copy()  # 取最接近的3把
                        else:
                            print(f"[INFO] 找到 {len(cand)} 把直径17mm的刀具")
                    else:
                        print(f"[INFO] 找到 {len(cand)} 把17R0.8刀具")
                else:
                    # SR<40mm，直径=面最短边×0.5（向上取整，上限17mm）
                    target_dia = min(17.0, math.ceil(sr * 0.5))
                    print(
                        f"[INFO] 开放面软料SR<40mm，面最短边: {sr}mm, 计算直径: {sr * 0.5:.1f}mm, 目标直径: {target_dia}mm")

                    # 寻找直径>=目标直径且<=17的刀具
                    cand = tools[
                        (tools["直径"] >= target_dia) &
                        (tools["直径"] <= 17.0)
                        ].copy()

                    if cand.empty:
                        # 如果没有符合条件的刀具，找最接近目标直径的刀具
                        print(f"[INFO] 直径在[{target_dia}, 17]范围内无刀具，选择最接近目标直径的刀具")
                        tools["_diff"] = (tools["直径"] - target_dia).abs()
                        cand = tools.nsmallest(3, "_diff").copy()
                    else:
                        print(f"[INFO] 找到 {len(cand)} 把直径在[{target_dia}, 17]范围内的刀具")

                # 在候选刀具中排序：优先选择R角为0.8的刀具
                if not cand.empty:
                    if len(cand) > 1:
                        # 计算优先级：R角0.8优先，然后直径最小的优先
                        cand.loc[:, "_priority_r"] = cand.apply(
                            lambda row: 0 if (row.get('R角', 0) == 0.8) else 1,
                            axis=1
                        )
                        cand.loc[:, "_priority_dia"] = cand["直径"]

                        cand = cand.sort_values(
                            ["_priority_r", "_priority_dia"]
                        )

                    # 输出候选刀具信息
                    print(f"[INFO] 候选刀具排序结果:")
                    for idx, (_, tool_row) in enumerate(cand.iterrows()):
                        print(f"  {idx + 1}. {tool_row['刀具名称']}, 直径: {tool_row['直径']}mm, "
                              f"类别: {tool_row.get('类别', '未知')}, R角: {tool_row.get('R角', 0)}")

                    row = cand.iloc[0]
                else:
                    print(f"[WARN] 没有候选刀具，使用备用选择")
                    row = tools.iloc[0]  # 备用选择
        else:
            # 封闭面 - 修改后的规则：最大刀也是17R0.8
            need_dia = math.ceil(sr / 2)  # 需求直径向上取整
            need_dia = min(need_dia, 17.0)  # 限制最大为17mm
            print(f"[INFO] 封闭面，面最短边: {sr}mm, 需求直径: {need_dia}mm (上限17mm)")

            # 寻找直径>=需求直径且<=17的刀具
            cand = tools[
                (tools["直径"] >= need_dia) &
                (tools["直径"] <= 17.0)
                ].copy()

            if cand.empty:
                # 如果没有符合条件的刀具，找最接近需求直径的刀具
                print(f"[INFO] 直径在[{need_dia}, 17]范围内无刀具，选择最接近需求直径的刀具")
                tools["_diff"] = (tools["直径"] - need_dia).abs()
                cand = tools.nsmallest(3, "_diff").copy()
            else:
                print(f"[INFO] 找到 {len(cand)} 把直径在[{need_dia}, 17]范围内的刀具")

            # 在候选刀具中排序：优先选择17R0.8的刀具
            if not cand.empty:
                if len(cand) > 1:
                    # 计算优先级：先按直径排序（从小到大），然后优先R角0.8
                    cand = cand.sort_values("直径")  # 先按直径从小到大排序

                    # 标记是否为17R0.8
                    cand.loc[:, "_is_17r08"] = cand.apply(
                        lambda row: 1 if (row['直径'] == 17.0 and row.get('R角', 0) == 0.8) else 0,
                        axis=1
                    )

                    # 优先选择17R0.8，然后按直径从小到大
                    cand = cand.sort_values(
                        ["_is_17r08", "直径"],
                        ascending=[False, True]
                    )

                # 输出候选刀具信息
                print(f"[INFO] 封闭面候选刀具排序结果:")
                for idx, (_, tool_row) in enumerate(cand.iterrows()):
                    print(f"  {idx + 1}. {tool_row['刀具名称']}, 直径: {tool_row['直径']}mm, "
                          f"类别: {tool_row.get('类别', '未知')}, R角: {tool_row.get('R角', 0)}")

                row = cand.iloc[0]
            else:
                # 如果没有符合条件的刀具，选择直径最大的刀具（但不超过17mm）
                print(f"[INFO] 没有直径在[{need_dia}, 17]范围内的刀具，选择最大直径刀具（上限17mm）")
                tools_within_17 = tools[tools["直径"] <= 17.0].copy()
                if not tools_within_17.empty:
                    tools_within_17 = tools_within_17.sort_values("直径", ascending=False)
                    row = tools_within_17.iloc[0]
                else:
                    print(f"[WARN] 没有直径≤17mm的刀具，使用默认刀具")
                    row = tools.iloc[0]

        # 获取切深
        depth = self.get_cutting_depth(row, material, is_heat)

        # 获取切削参数
        def get_float_value(col_name, default=0):
            if col_name in row and pd.notna(row[col_name]):
                try:
                    value = row[col_name]
                    if isinstance(value, str) and value.upper() == "NONE":
                        return default
                    return float(value)
                except (ValueError, TypeError):
                    return default
            return default

        speed = get_float_value("转速(普)", 0)
        feed = get_float_value("进给(普)", 0)
        rapid = get_float_value("横越(普)", 0)

        # 输出选刀信息
        tool_name = row["刀具名称"]
        tool_dia = float(row["直径"])
        tool_category = row.get('类别', '未知')
        tool_r = row.get('R角', 0)
        print(
            f"[INFO] 选择刀具: {tool_name}, 直径: {tool_dia}mm, 类别: {tool_category}, R角: {tool_r}, 切深: {depth}mm")

        return tool_name, tool_dia, depth, speed, feed, rapid, tool_category

    # -------------------- 原 kaibi 工具函数 --------------------
    @staticmethod
    def parse_point(s: str) -> Tuple[float, float, float]:
        """解析点坐标字符串"""
        return tuple(map(float, s.strip().split(',')))

    @staticmethod
    def shortest_edge(row: Dict[str, str]) -> float:
        """计算面的最短边"""
        L = float(row["Long"])
        W = float(row["Width"])
        H = float(row["Height"])
        return min(L, W) if L > 0 and W > 0 else H

    @staticmethod
    def is_z_direction(row: Dict[str, str]) -> bool:
        """判断法向量是否在Z方向（朝上或朝下）"""
        try:
            nx, ny, nz = KiabiProcessor.parse_point(row["Face Normal"])
            return abs(nz) > 0.9
        except:
            return False

    def load_data(self):
        """加载CSV数据"""
        print("\n" + "-" * 50)
        print("  [1] 加载CSV数据")
        print("-" * 50)

        # 加载面数据
        try:
            with open(self.CONFIG["CSV_FACE"], newline='', encoding='utf-8') as f:
                for r in csv.DictReader(f):
                    self.tag2row[r["Face Tag"]] = r
            print(f"[INFO] 成功加载面数据，共 {len(self.tag2row)} 个面")
        except Exception as e:
            print(f"[ERROR] 加载面数据失败: {e}")
            sys.exit(1)

        # 加载标签数据
        try:
            with open(self.CONFIG["CSV_TAG"], newline='', encoding='utf-8') as f:
                for r in csv.DictReader(f, fieldnames=self.CONFIG["COL_NAMES"]):
                    for col in self.CONFIG["LAYER_MAP"]:
                        v = r[col].strip()
                        if v:
                            self.col_tags[col].append(v)
            print(f"[INFO] 成功加载标签数据")
        except Exception as e:
            print(f"[ERROR] 加载标签数据失败: {e}")
            sys.exit(1)

    def process(self):
        """主处理流程"""
        print("\n" + "-" * 50)
        print("  [2] 获取材质和刀具参数")
        print("-" * 50)

        # 获取材质和热处理信息
        material, is_heat = self.get_material_from_filename()

        # 读取刀具参数
        tools_df = self.read_tool_parameters_from_json()
        if tools_df.empty:
            sys.exit("[ERROR] 刀具表读取失败，程序终止")

        print("\n" + "-" * 50)
        print("  [3] 处理各个方向的面")
        print("-" * 50)

        # 主处理逻辑
        result = {}
        index = 1
        X = self.CONFIG["MIN_LENGTH_THRESHOLD"]

        # 第一步：收集所有朝上面的Z坐标，找出最大Z值
        all_up_faces_z = []
        for col in self.CONFIG["COL_NAMES"]:
            tags = self.col_tags[col]
            for tag in tags:
                row = self.tag2row.get(tag)
                if not row:
                    continue

                # 检查面类型、颜色和方向
                if row["Face Type"] != "22":
                    continue
                if row.get("Face Color", "").strip() != "6":
                    continue
                if not self.is_z_direction(row):
                    continue

                # 获取面的Z坐标
                z_self = self.parse_point(row["Face Data - Point"])[2]
                all_up_faces_z.append(z_self)

        # 找出最大Z值（模型最上面的面）
        max_z = max(all_up_faces_z) if all_up_faces_z else 0
        print(f"[INFO] 检测到所有朝上面的最大Z值: {max_z:.2f}mm，将过滤掉Z值等于此值的面")

        # 第二步：处理各个方向的面，过滤掉Z值最大的面
        for col in self.CONFIG["COL_NAMES"]:
            tags = self.col_tags[col]
            close_list = []  # 封闭面列表
            open_list = []  # 开放面列表

            print(f"[DEBUG] 处理方向 {col}: 共有 {len(tags)} 个标签")

            # 分类处理
            for tag in tags:
                row = self.tag2row.get(tag)
                # 检查面类型、颜色和方向
                if not row:
                    print(f"[DEBUG] 标签 {tag} 未在面数据中找到")
                    continue

                # 检查基本条件
                if row["Face Type"] != "22":
                    continue
                if row.get("Face Color", "").strip() != "6":
                    continue
                if not self.is_z_direction(row):
                    continue

                # 获取面的Z坐标
                z_self = self.parse_point(row["Face Data - Point"])[2]

                # 新增规则：过滤掉Z值最大的面（模型最上面的面）
                if abs(z_self - max_z) < 1e-6:
                    # 获取法向量
                    try:
                        nx, ny, nz = self.parse_point(row["Face Normal"])
                        # 只过滤法向量Z分量>0.999的面（接近完全水平）
                        if nz > 0.999:
                            print(f"[INFO] 过滤掉模型最上面的水平面 {tag} (Z={z_self:.2f}mm, nz={nz:.6f})")
                            continue
                        else:
                            print(f"[INFO] 保留倾斜顶面 {tag} (Z={z_self:.2f}mm, nz={nz:.6f})")
                    except:
                        print(f"[INFO] 过滤掉模型最上面的面 {tag} (Z={z_self:.2f}mm)")
                        continue

                adj_tags = [t.strip() for t in row["Adjacent Face Tags"].split(';') if t.strip()]

                # 获取相邻面的Z坐标
                zs = []
                for t in adj_tags:
                    if t in self.tag2row:
                        try:
                            zs.append(self.parse_point(self.tag2row[t]["Face Data - Point"])[2])
                        except:
                            continue

                if not zs:
                    continue

                # 判断是否是开放面
                is_open = any(z < z_self - 1e-6 for z in zs)
                length = self.shortest_edge(row)

                if length < X:
                    continue

                if is_open:
                    open_list.append((tag, length))
                else:
                    close_list.append((tag, length))

            # 处理封闭面
            close_grp = defaultdict(list)
            for t, l in close_list:
                close_grp[int(l)].append(t)

            # 处理开放面 - 聚类
            open_list.sort(key=lambda x: x[1])
            open_clusters = []
            if open_list:
                min_len, tags = open_list[0][1], [open_list[0][0]]
                for t, l in open_list[1:]:
                    if l - min_len <= self.CONFIG["CLUSTER_TOLERANCE"]:
                        tags.append(t)
                    else:
                        if min_len >= X:
                            open_clusters.append((min_len, tags))
                        min_len, tags = l, [t]
                if min_len >= X:
                    open_clusters.append((min_len, tags))

            # 获取图层
            layer = self.CONFIG["LAYER_MAP"][col]

            # 处理开放面组
            for min_len, tags in open_clusters:
                tool_name, tool_diam, depth, speed, feed, rapid, tool_category = self.select_tool_and_depth(
                    material, is_heat, min_len, True, tools_df)
                # 计算运动类型 - 取面组中所有面的运动类型，按多数决定
                motion_types = []
                for tag in tags:
                    motion_type = self.calculate_motion_type(tag)
                    motion_types.append(motion_type)

                # 统计运动类型
                cut_count = motion_types.count("切削")
                follow_count = motion_types.count("跟随")

                # 确定最终运动类型
                final_motion_type = "切削" if cut_count >= follow_count else "跟随"

                key = f"MIAN1_{index}"
                index += 1
                result[key] = {
                    "工序": "MIAN1_SIMPLE",
                    "面ID列表": [int(t) for t in tags],
                    "刀具名称": tool_name,
                    "切深": depth,
                    "指定图层": layer,
                    "参考刀具": "NULL",
                    "转速": speed,
                    "进给": feed,
                    "横越": rapid,
                    "刀具类别": tool_category,
                    "切深值": depth, # 临时存储，现在将用于半精加工JSON
                    "运动类型": final_motion_type  # 添加运动类型
                }

            # 处理封闭面组
            for _, tags in sorted(close_grp.items()):
                # 检查封闭面第一个面的颜色
                first_tag_row = self.tag2row[tags[0]]
                if first_tag_row.get("Face Color", "").strip() != "6":
                    continue

                min_len = int(self.shortest_edge(first_tag_row))
                tool_name, tool_diam, depth, speed, feed, rapid, tool_category = self.select_tool_and_depth(
                    material, is_heat, min_len, False, tools_df)
                # 计算运动类型 - 取面组中所有面的运动类型，按多数决定
                motion_types = []
                for tag in tags:
                    motion_type = self.calculate_motion_type(tag)
                    motion_types.append(motion_type)

                # 统计运动类型
                cut_count = motion_types.count("切削")
                follow_count = motion_types.count("跟随")

                # 确定最终运动类型
                final_motion_type = "切削" if cut_count >= follow_count else "跟随"

                key = f"MIAN1_{index}"
                index += 1
                result[key] = {
                    "工序": "MIAN1_SIMPLE",
                    "面ID列表": [int(t) for t in tags],
                    "刀具名称": tool_name,
                    "切深": depth,
                    "指定图层": layer,
                    "参考刀具": "NULL",
                    "转速": speed,
                    "进给": feed,
                    "横越": rapid,
                    "刀具类别": tool_category,
                    "切深值": depth,  # 临时存储，现在将用于半精加工JSON
                    "运动类型": final_motion_type  # 添加运动类型
                }

        self.result = result
        print(f"[INFO] 处理完成，共生成 {len(result)} 个加工工序")

    def _merge_small_face_groups(self, result_dict):
        """
        合并面ID列表≤2的相同刀具工序

        Args:
            result_dict: 原始结果字典

        Returns:
            合并后的结果字典
        """
        # 按刀具名称分组
        tool_groups = {}
        for key, data in result_dict.items():
            tool_name = data["刀具名称"]
            if tool_name not in tool_groups:
                tool_groups[tool_name] = []
            tool_groups[tool_name].append((key, data))

        merged_result = {}
        new_index = 1

        # 处理每个刀具组
        for tool_name, group_items in tool_groups.items():
            # 如果该刀具只有一个工序，直接保留
            if len(group_items) == 1:
                key, data = group_items[0]
                merged_result[f"MIAN1_{new_index}"] = data
                new_index += 1
                continue

            # 如果该刀具有多个工序，检查是否需要合并
            large_groups = []  # 面ID列表>2的工序
            small_groups = []  # 面ID列表≤2的工序
            all_face_ids = []  # 所有小面的面ID（用于合并）

            # 分类处理
            for key, data in group_items:
                face_ids = data["面ID列表"]
                if len(face_ids) <= 2:
                    small_groups.append((key, data))
                    all_face_ids.extend(face_ids)
                else:
                    large_groups.append((key, data))

            # 如果有小面组，合并它们
            if small_groups and len(all_face_ids) > 0:
                # 使用第一个小面组的参数作为合并后的参数
                first_key, first_data = small_groups[0]
                merged_data = first_data.copy()
                merged_data["面ID列表"] = all_face_ids

                merged_result[f"MIAN1_{new_index}"] = merged_data
                new_index += 1
                print(f"[INFO] 合并了{len(small_groups)}个小面组，刀具: {tool_name}, 合并后共{len(all_face_ids)}个面")

            # 保留大面组（面ID列表>2）
            for key, data in large_groups:
                merged_result[f"MIAN1_{new_index}"] = data
                new_index += 1

        print(f"[INFO] 合并完成: 原始{len(result_dict)}个工序 → 合并后{len(merged_result)}个工序")
        return merged_result

    def calculate_motion_type(self, face_tag: str) -> str:
        """
        根据面的Area、Long、Width计算运动类型

        Args:
            face_tag: 面标签

        Returns:
            str: "切削" 或 "跟随"
        """
        if face_tag not in self.tag2row:
            return "跟随"  # 默认值

        row = self.tag2row[face_tag]

        try:
            # 获取Area、Long、Width值
            area = float(row.get("Area", 0))
            long = float(row.get("Long", 0))
            width = float(row.get("Width", 0))

            # 计算NEW_Area
            new_area = long * width

            # 防止除零错误
            if new_area <= 0:
                return "跟随"

            # 计算比例
            ratio = area / new_area

            # 根据比例判断运动类型
            if ratio >= 0.76:
                return "跟随"
            elif ratio<=0.5:
                return "跟随"
            else:
                return "切削"

        except (ValueError, TypeError, KeyError):
            return "跟随"  # 如果计算出错，返回默认值

    def generate_json_files(self):
        """生成JSON文件"""
        print("\n" + "-" * 50)
        print("  [4] 生成JSON文件")
        print("-" * 50)

        # 检查结果是否为空
        if not self.result:
            print("[WARN] 处理结果为空，没有需要加工的面，跳过JSON文件生成")
            print("[INFO] 可能的原因：")
            print("  1. 所有面都被过滤掉了（如模型最上面的面）")
            print("  2. 面尺寸小于最小阈值")
            print("  3. 面类型或颜色不符合条件")
            print("  4. CSV文件中没有有效的面数据")
            return

        # 确保输出目录存在
        os.makedirs(self.CONFIG["OUT_DIR"], exist_ok=True)

        # 输出文件名
        half_path = os.path.join(self.CONFIG["OUT_DIR"], self.CONFIG["OUT_JSON_HALF"])
        full_path = os.path.join(self.CONFIG["OUT_DIR"], self.CONFIG["OUT_JSON_FULL"])

        print(f"[INFO] 半精加工文件: {self.CONFIG['OUT_JSON_HALF']}")
        print(f"[INFO] 全精加工文件: {self.CONFIG['OUT_JSON_FULL']}")

        # ---------- 1. 生成半精加工JSON文件 ----------
        half_result = {}
        for key, v in self.result.items():
            # 从结果中获取切深值（根据材质和热处理状态从铣刀参数.json中获取的值）
            cut_depth_per_pass = v["切深值"]
            print(f"[INFO] 工序 {key}: 刀具 {v['刀具名称']} 的每刀切削深度设置为 {cut_depth_per_pass}mm")

            half_result[key] = {
                "工序": v["工序"],
                "面ID列表": v["面ID列表"],
                "刀具名称": v["刀具名称"],
                "刀具类别": v["刀具类别"],
                "切深": v["切深"],
                "指定图层": v["指定图层"],
                "参考刀具": "NULL",
                "转速": v["转速"],
                "进给": v["进给"],
                "横越": v["横越"],
                "最终底面余量": 0.05 if v["刀具类别"] in ["钨钢平刀", "钨钢牛鼻刀"] else 0.2,
                "壁余量": 3.0,
                "底面毛坯厚度": 0.5,
                "每刀切削深度": cut_depth_per_pass,  # 使用根据材质和热处理状态计算的实际切深值
                "部件余量": 0.1,
                "运动类型": v["运动类型"]  # 添加运动类型
            }
        # 在这里调用合并函数，合并面ID列表≤2的相同刀具工序
        print(f"[INFO] 开始合并半精加工中的面ID列表≤2的相同刀具工序...")
        half_result = self._merge_small_face_groups(half_result)

        # 检查半精加工结果是否为空
        if half_result:
            with open(half_path, 'w', encoding='utf-8') as f:
                json.dump(half_result, f, ensure_ascii=False, indent=4)
            print(f"✅ 半精加工文件已生成: {half_path}")
            print(f"[INFO] 半精加工JSON中的每刀切削深度已根据材质和热处理状态设置")
        else:
            print(f"[WARN] 半精加工结果为空，不生成 {self.CONFIG['OUT_JSON_HALF']} 文件")

        # ---------- 2. 生成全精加工JSON文件 ----------
        full_result = {}

        # 检查是否所有刀具直径都大于10mm
        all_diameters_gt_10 = True
        diameters_info = []

        for key, v in self.result.items():
            # 提取刀具直径信息
            tool_name = v["刀具名称"]
            # 从刀具名称中提取直径（假设格式为"D10"、"D32"等）
            dia_match = re.search(r'D(\d+(\.\d+)?)', tool_name.upper())
            if dia_match:
                dia = float(dia_match.group(1))
            else:
                # 如果无法从名称提取，使用刀具参数表中的直径
                dia = v.get("刀具直径", 0)  # 注意：这里需要确保刀具直径已保存在结果中
                if dia == 0:
                    # 如果结果中没有保存直径，需要从刀具参数表中查找
                    tools_df = self.read_tool_parameters_from_json()
                    tool_row = tools_df[tools_df["刀具名称"] == tool_name]
                    if not tool_row.empty:
                        dia = float(tool_row.iloc[0]["直径"])
                    else:
                        dia = 0

            diameters_info.append((key, tool_name, dia))

            # 检查是否所有直径都大于10
            if dia <= 10:
                all_diameters_gt_10 = False

        print(f"[INFO] 刀具直径检查:")
        for key, tool_name, dia in diameters_info:
            print(f"  {key}: {tool_name}, 直径: {dia}mm")

        # 如果所有刀具直径都大于10mm，则在全精加工中使用D10或直径10mm的刀具
        d10_name = None
        d10_category = None
        d10_speed = 0
        d10_feed = 0
        d10_rapid = 0

        if all_diameters_gt_10 and diameters_info:
            print(f"[INFO] 所有刀具直径都大于10mm，全精加工将统一使用D10或直径10mm的刀具")

            # 查找D10或直径10mm的刀具
            tools_df = self.read_tool_parameters_from_json()
            d10_candidates = []

            # 首先查找名称为"D10"的刀具
            d10_tool = tools_df[tools_df["刀具名称"].str.upper() == "D10"]
            if not d10_tool.empty:
                d10_candidates.append(d10_tool.iloc[0])
                print(f"[INFO] 找到名称为D10的刀具: {d10_tool.iloc[0]['刀具名称']}")

            # 如果没有找到"D10"，查找直径为10mm的刀具
            if not d10_candidates:
                dia10_tools = tools_df[tools_df["直径"] == 10.0]
                if not dia10_tools.empty:
                    # 优先选择钨钢平刀且R角为0的刀具
                    priority_tools = dia10_tools[
                        (dia10_tools["类别"] == "钨钢平刀") &
                        (dia10_tools["R角"] == 0)
                        ]

                    if not priority_tools.empty:
                        d10_candidates.append(priority_tools.iloc[0])
                        print(f"[INFO] 找到直径10mm的钨钢平刀R0刀具: {priority_tools.iloc[0]['刀具名称']}")
                    else:
                        # 如果没有优先刀具，选择第一个直径10mm的刀具
                        d10_candidates.append(dia10_tools.iloc[0])
                        print(f"[INFO] 找到直径10mm的刀具: {dia10_tools.iloc[0]['刀具名称']}")

            if d10_candidates:
                d10_tool_row = d10_candidates[0]
                d10_name = d10_tool_row["刀具名称"]
                d10_category = d10_tool_row.get("类别", "未知")

                # 获取D10刀具的切削参数
                def get_d10_float_value(col_name, default=0):
                    if col_name in d10_tool_row and pd.notna(d10_tool_row[col_name]):
                        try:
                            value = d10_tool_row[col_name]
                            if isinstance(value, str) and value.upper() == "NONE":
                                return default
                            return float(value)
                        except (ValueError, TypeError):
                            return default
                    return default

                d10_speed = get_d10_float_value("转速(普)", 0)
                d10_feed = get_d10_float_value("进给(普)", 0)
                d10_rapid = get_d10_float_value("横越(普)", 0)

                print(f"[INFO] 全精加工统一使用刀具: {d10_name}, 直径: {d10_tool_row['直径']}mm, 类别: {d10_category}")
                print(f"[INFO] 切削参数 - 转速: {d10_speed}, 进给: {d10_feed}, 横越: {d10_rapid}")
            else:
                print(f"[WARN] 未找到D10或直径10mm的刀具，全精加工将使用原刀具")

        # 生成全精加工结果
        for key, v in self.result.items():
            # 如果所有刀具直径都大于10mm且找到了D10刀具，则使用D10刀具
            if all_diameters_gt_10 and d10_name:
                # 使用D10刀具的参数
                full_result[key] = {
                    "工序": v["工序"],
                    "面ID列表": v["面ID列表"],
                    "刀具名称": d10_name,
                    "刀具类别": d10_category,
                    "切深": 0.0,  # 全精加工切深为0
                    "指定图层": v["指定图层"],
                    "参考刀具": "NULL",
                    "转速": d10_speed,
                    "进给": d10_feed,
                    "横越": d10_rapid,
                    "最终底面余量": 0.0,
                    "壁余量": 0.0,
                    "底面毛坯厚度": 0.1,
                    "每刀切削深度": 0.0,  # 全精加工每刀切削深度为0
                    "部件余量": 0.12,
                    "运动类型": v["运动类型"]  # 添加运动类型
                }
            else:
                # 使用原刀具
                full_result[key] = {
                    "工序": v["工序"],
                    "面ID列表": v["面ID列表"],
                    "刀具名称": v["刀具名称"],
                    "刀具类别": v["刀具类别"],
                    "切深": v["切深"],
                    "指定图层": v["指定图层"],
                    "参考刀具": "NULL",
                    "转速": v["转速"],
                    "进给": v["进给"],
                    "横越": v["横越"],
                    "最终底面余量": 0.0,
                    "壁余量": 0.0,
                    "底面毛坯厚度": 0.1,
                    "每刀切削深度": 0.0,  # 全精加工每刀切削深度为0
                    "部件余量": 0.12,
                    "运动类型": v["运动类型"]  # 添加运动类型
                }
        # 在这里调用合并函数，合并面ID列表≤2的相同刀具工序
        print(f"[INFO] 开始合并全精加工中的面ID列表≤2的相同刀具工序...")
        full_result = self._merge_small_face_groups(full_result)

        # 检查全精加工结果是否为空
        if full_result:
            with open(full_path, 'w', encoding='utf-8') as f:
                json.dump(full_result, f, ensure_ascii=False, indent=4)
            print(f"✅ 全精加工文件已生成: {full_path}")
        else:
            print(f"[WARN] 全精加工结果为空，不生成 {self.CONFIG['OUT_JSON_FULL']} 文件")

        # 输出总结信息
        if half_result or full_result:
            print(f"\n✅ 总共生成 {len(half_result)} 个半精加工工序和 {len(full_result)} 个全精加工工序")
            if half_result:
                print(f"✅ 半精加工JSON中的每刀切削深度已根据材质和热处理状态从铣刀参数.json中获取")
            print(f"✅ 两个文件中的键名都统一为 MIAN1_x 格式")
        else:
            print(f"\n[WARN] 未生成任何JSON文件，没有有效的加工工序")

    def run(self):
        """运行完整流程"""
        print("=" * 70)
        print("  Kiabi 9.0 加工路径生成脚本")

        # 显示零件信息
        if self.CONFIG["PART_NAME"]:
            print(f"  零件: {self.CONFIG['PART_NAME']}")
        else:
            print("  零件: 未知")

        print("=" * 70)

        # 加载数据
        self.load_data()

        # 处理数据
        self.process()

        # 生成JSON文件
        self.generate_json_files()

        print("\n" + "=" * 70)
        print("  处理完成！")
        print("=" * 70)


def main1(csv_face, csv_tag, out_dir, prt_folder, excel_params, tool_json):
    """
    供 run_step8.py 调用的接口函数
    """
    processor = KiabiProcessor()
    processor.setup_paths(
        csv_face=csv_face,
        csv_tag=csv_tag,
        out_dir=out_dir,
        prt_folder=prt_folder,
        excel_params=excel_params,
        tool_json=tool_json
    )
    processor.run()


def main():
    """主函数 - 示例使用"""
    # 创建处理器实例
    processor = KiabiProcessor()

    # 配置路径
    processor.setup_paths(
        csv_face=r"C:\Projects\03_Analysis\Face_Info\face_csv\DIE-03_face_data.csv",
        csv_tag=r"C:\Projects\03_Analysis\Geometry_Analysis\DIE-03.prt.csv",
        out_dir=r"C:\Projects\NC\file\json",
        prt_folder=r"C:\Projects\04_PRT_with_Tool\DIE-03.prt",
        excel_params=r"C:\Projects\NC\file\output_file\零件参数.xlsx",
        tool_json=r"C:\Projects\NC\file\铣刀参数.json"
    )

    # 运行处理流程
    processor.run()


if __name__ == "__main__":
    main()