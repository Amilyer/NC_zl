#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""NX自动化工具 - 简化版（修复版）"""
import os
import time

import NXOpen
import NXOpen.UF

# ==================================================================================
# 配置
# ==================================================================================
CONFIG = {
    "PART_PATH": r"C:\Projects\NC\file\final_workpiece_with_tools\DIE-03_modified.prt",
}
# ==================================================================================
# NX工具类
# ==================================================================================
class NXTools:
    def __init__(self, workPart):
        self.workPart = workPart
        self.session = NXOpen.Session.GetSession()
        self.uf = NXOpen.UF.UFSession.GetUFSession()

    # 工作流辅助函数

    def switch_to_manufacturing(self):
        try:
        # 1. 检查 CAM 会话
            if not self.session.IsCamSessionInitialized():
                print("CAM 会话未初始化，正在启动...")
                self.session.CreateCamSession()
        # 2. 检查 Setup 是否存在
            try:
                if self.workPart.CAMSetup.IsInitialized():
                    return True
            except:
                pass # 继续向下尝试创建
        # 3. 创建 Setup
            print("当前部件没有有效的 Setup，正在自动创建 'hole_making' 环境...")
            self.workPart.CreateCamSetup("hole_making")
            print("✅ CAM Setup (hole_making) 创建成功。")
            return True
        except Exception as ex:
            print(f"❌ 自动创建 CAM Setup 失败: {ex}")
            return False

    def save_part(self,workPart):
        """保存部件到 output 子文件夹（带时间戳）"""
        try:

            save_path = None
            part_path = workPart.FullPath
            
            # 防御性编程：如果是新建文件没保存过，FullPath可能为空
            if not part_path:
                return "当前部件未保存过，无法获取路径"

            # --- 路径处理逻辑 ---
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            dir_name, file_name = os.path.split(part_path)
            name, ext = os.path.splitext(file_name)
            
            output_dir = os.path.join(dir_name, "output")
            os.makedirs(output_dir, exist_ok=True)
            
            save_path = os.path.join(output_dir, f"{name}_{timestamp}{ext}")

            workPart.SaveAs(save_path)

            return {
                "file_path": save_path,
                "file_name": os.path.basename(save_path)
            }
            
        except Exception as e:
            # 加上错误捕获，防止保存失败导致服务崩溃
            return f"保存失败: {str(e)}"
        

    def call_nx_dll(self, session):
        try:
            # 1. 获取 Session
            
            
            # 2. 定义 DLL 路径 (建议使用绝对路径)
            # 注意：路径中不要有中文，且使用 raw string (r"...") 避免转义问题
            dll_path = r"C:\Projects\NC\modules\NX_Opentest.dll"
            
            if not os.path.exists(dll_path):
                print(f"错误: 找不到文件 {dll_path}")
                return

            # 3. 准备参数 (如果 C++ 的 ufusr 接受参数)
            # 如果 C++ 端是 void ufusr(char *param, int *retcod, int param_len), 
            # 这里的 args 将传给 param
            args = []  # 或者 args = ["参数1", "参数2"]
            
            print(f"正在执行: {dll_path}")
            
            # 4. 调用 ExecuteUserFunction
            # 参数说明:
            # - libraryName: DLL 的完整路径
            # - entryName: C++ 入口函数名，通常是 "ufusr"
            # - inputArgs: 传递给 DLL 的参数列表 (通常是字符串或对象数组)
            return_value = session.ExecuteUserFunction(dll_path, "ufusr", args)
            
            print(f"DLL 执行完毕，返回值: {return_value}")

        except Exception as ex:
            print(f"执行 DLL 失败: {str(ex)}")

# ==================================================================================
# 主程序（仅用于测试）
# ==================================================================================
def main():
    # 获取配置
    part_path = CONFIG["PART_PATH"]
    print(f"部件: {part_path}")
    # 执行流程
    session = NXOpen.Session.GetSession()
    # 打开部件
    base_part, load_status = session.Parts.OpenBaseDisplay(part_path)
    # 创建工具实例
    workPart = session.Parts.Work
    nx = NXTools(workPart)
    # 设置CAM环境 
    nx.switch_to_manufacturing()



    # 生成NX工艺排序
    # from modules.procsse_sort import Procsse_sort
    # ps = Procsse_sort()
    # craft_result=ps.process_nx_crafts(workPart,judgement_M=False)
    # print(craft_result)
    
    # 自动打孔
    from modules.Drilling_Automation.main_workflow import MainWorkflow
    mw = MainWorkflow(session, workPart)
    mw.run_workflow()



    print(nx.save_part(workPart))
    # 清理资源
    if load_status:
        load_status.Dispose()
    return
if __name__ == "__main__":
    main()