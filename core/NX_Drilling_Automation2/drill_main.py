# # -*- coding: utf-8 -*-
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

import NXOpen
import NXOpen.CAM
import NXOpen.Features
import NXOpen.GeometricUtilities
import traceback
import NXOpen.UF
import drill_config
# 导入模块
from main_workflow import MainWorkflow
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


def drill_start(session, work_part, file_name, drill_path, knfie_path, is_save=False):
    """主函数"""
    try:
        drill_config.DRILL_JSON_PATH = drill_path
        drill_config.MILL_JSON_PATH = knfie_path
        displayPart = session.Parts.Display
        # 创建加工环境
        cam(session, work_part)
        # 创建工作流控制器
        workflow = MainWorkflow(session, work_part, file_name)

        # 执行完整工作流
        result = workflow.run_workflow()
        if result:
            print_to_info_window(f"✅ NX钻孔自动化流程执行成功---成功文件：{work_part.Name}.prt")
        else:
            print_to_info_window(f"❌ NX钻孔自动化流程执行失败---失败文件：{work_part.Name}.prt")
        if is_save:
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
    # for name in os.listdir(r"C:\Users\Admin\Desktop\3_Cleaned_PRT"):
    #     file_path = os.path.join(r'C:\Users\Admin\Desktop\3_Cleaned_PRT', name)

    file_path = r"C:\Projects\NCv4.7\output\04_PRT_with_Tool\B1-01.prt"
    session = NXOpen.Session.GetSession()
    work_part = open_part(session, file_path)
    name = file_path.split("\\")[-1]
    # 运行主程序
    drill_start(session, work_part, name, r"C:\Projects\NCv4.7\input\drill_table.json", r"C:\Projects\NCv4.7\input\knife_table.json", True)
