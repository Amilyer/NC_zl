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
import config
from process_info import ProcessInfoHandler
from parameter_parser import ParameterParser
from geometry import GeometryHandler

from utils import print_to_info_window


def cam(theSession, workPart):
    theSession.Parts.SetDisplay(workPart, False, False)
    theSession.Parts.SetWork(workPart)
    markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Enter 加工")

    theSession.ApplicationSwitchImmediate("UG_APP_MANUFACTURING")

    result1 = theSession.IsCamSessionInitialized()

    theSession.CreateCamSession()

    theSession.CAMSession.PathDisplay.SetAnimationSpeed(5)

    theSession.CAMSession.PathDisplay.SetIpwResolution(NXOpen.CAM.PathDisplay.IpwResolutionType.Medium)

    markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "初始化组装")

    theSession.CAMSession.SpecifyConfiguration("${UGII_CAM_CONFIG_DIR}cam_express")

    theSession.CAMSession.SpecifyConfiguration("${UGII_CAM_CONFIG_DIR}cam_general")

    cAMSetup1 = workPart.CreateCamSetup("hole_making")

    kinematicConfigurator1 = workPart.CreateKinematicConfigurator()

    kinematicConfigurator2 = workPart.KinematicConfigurator


def open_part(session, part_path):
    """保证返回 NXOpen.Part 类型"""
    try:
        result = session.Parts.Open(part_path)
        # tuple unpack
        if isinstance(result, tuple):
            part = result[0]
        else:
            part = result

        if isinstance(part, NXOpen.Part):
            return part

        raise TypeError("Open() 未返回 NXOpen.Part 对象")

    except Exception as e:
        # 统一异常格式：异常信息 + 位置（文件名、行号、函数名）
        exc_info = traceback.extract_tb(e.__traceback__)[-1]
        err_msg = (f"[open_part] 异常：{str(e)} | "
                   f"异常位置：文件{exc_info.filename}，行{exc_info.lineno}，函数{exc_info.name}")
        print(err_msg)
        raise  # 重新抛出异常，让调用者处理


def lw_write(message):
    """兼容函数：日志输出"""
    print_to_info_window(message)


def drill_start(file_directory, drill_path, knfie_path):
    """主函数"""
    try:
        # session = NXOpen.Session.GetSession()
        # work_part = session.Parts.Work

        config.DRILL_JSON_PATH = drill_path
        config.MILL_JSON_PATH = knfie_path
        session = NXOpen.Session.GetSession()
        files_path = os.scandir(file_directory)
        for file_path in files_path:
            if file_path.path.endswith(".prt"):

                work_part = open_part(session, file_path.path)
                displayPart = session.Parts.Display
                # 创建加工环境
                cam(session, work_part)
                geometry_handler = GeometryHandler(session, work_part)
                process_info_handler = ProcessInfoHandler()
                parameter_parser = ParameterParser()
                # 第一步：预处理，提取加工信息
                print_to_info_window("第一步：预处理，提取加工信息")
                annotated_data = process_info_handler.extract_and_process_notes(work_part)
                processed_result = parameter_parser.process_hole_data(annotated_data)
                annotated_data = annotated_data["主体加工说明"]
                if not processed_result:
                    print(("加工信息处理失败，流程终止"))
                    return False

                # 获取材料类型
                material_type = process_info_handler.get_material(annotated_data["材质"][0])
                if material_type is None:
                    print("获取材质类型失败")
                    return False
                # 获取毛坯尺寸
                lwh_point = process_info_handler.get_workpiece_dimensions(annotated_data["尺寸"])
                # 获取加工坐标原点
                original_point = geometry_handler.get_start_point(lwh_point)
                if original_point[0] - 1e-7 != 0.0 or original_point[1] - 1e-7 != 0.0:
                    # 移动2D图到绝对坐标系
                    geometry_handler.move_objects_point_to_point(original_point, config.DEFAULT_ORIGIN_POINT)
                # 文件保存
                partSaveStatus1 = work_part.Save(NXOpen.BasePart.SaveComponents.TrueValue,
                                                 NXOpen.BasePart.CloseAfterSave.FalseValue)
                partSaveStatus1.Dispose()
                # 关闭当前部件
                work_part.Close(NXOpen.BasePart.CloseWholeTree.FalseValue, NXOpen.BasePart.CloseModified.UseResponses,
                                None)
                work_part = None



    except Exception as e:
        print_to_info_window(f"❌ 主程序执行异常: {str(e)}")
        traceback.print_exc()
        return False


if __name__ == "__main__":
    # 运行主程序
    drill_start(r"D:\项目\NC编程\test_file", r"D:\项目\NC编程\代码实现\tool_file\drill_table.json", r"D:\项目\NC编程\代码实现\tool_file\knife_table.json")
