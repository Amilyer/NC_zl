# -*- coding: utf-8 -*-
"""
NX 2312 自动钻孔 + 钻刀创建 + 路径优化脚本 - 重构版本
模块化重构，提升代码可读性和可维护性

功能：
1. 从部件曲线中识别圆孔
2. 通过最近邻 + 2-opt 优化钻孔路径
3. 自动创建钻刀（STD_DRILL）
4. 自动创建钻孔工序并生成刀轨
5. 自动创建程序组，并将工序放入其中
"""

import os
import NXOpen
import NXOpen.CAM
import math
import json
import NXOpen.Features
import NXOpen.GeometricUtilities
import re
import copy
import traceback
import NXOpen.UF

# 导入模块
from main_workflow import MainWorkflow
from utils import print_to_info_window
from process_info import ProcessInfoHandler
from parameter_parser import ParameterParser

def lw_write(message):
    """兼容函数：日志输出"""
    print_to_info_window(message)


if __name__ == "__main__":
    session = NXOpen.Session.GetSession()
    work_part = session.Parts.Work
    pih = ProcessInfoHandler()
    pp = ParameterParser()
    annotated_data = pih.extract_and_process_notes(work_part)
    lw_write(annotated_data)
    processed_result = pp.process_hole_data(annotated_data)
    for info in processed_result:
        lw_write(f"key:{info},processed_result:{processed_result[info]}")
    