#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
NX CAM自动化工具 - 刀具创建模块
独立功能：从JSON创建所有铣刀并直接保存PRT文件
新增功能：读取R角、长度、刃长参数
"""

from datetime import datetime
import traceback
import NXOpen
import NXOpen.CAM
import NXOpen.UF
import os
import json
from contextlib import contextmanager


# ==================================================================================
# 配置
# ==================================================================================
CONFIG = {
    "PART_PATH": r'C:\Projects\NC\output\04_PRT_with_Tool\DIE-14.prt',
    "AUTO_SAVE": True,
    "JSON_TOOLS_PATH": r'C:\Projects\NC\input\铣刀参数.json',
}


# ==================================================================================
# ToolCreator 刀具创建类
# ==================================================================================
class ToolCreator:
    def __init__(self, work_part):
        self.work_part = work_part
        self.session = NXOpen.Session.GetSession()
        self.uf = NXOpen.UF.UFSession.GetUFSession()
        self.created_count = 0
        self.skipped_count = 0

    def print_log(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        emoji_map = {
            "INFO": "ℹ️", "WARN": "⚠️", "ERROR": "❌",
            "SUCCESS": "✅", "DEBUG": "🔍", "START": "🚀", "END": "🏁"
        }
        emoji = emoji_map.get(level.upper(), "")
        print(f"[{timestamp}] {emoji} {message}", flush=True)

    def print_separator(self, char="=", length=80):
        print(char * length, flush=True)

    def print_header(self, title):
        self.print_separator()
        print(f"  {title}".center(80), flush=True)
        self.print_separator()

    def switch_to_manufacturing(self):
        """切换到加工环境"""
        if self.session.ApplicationName != "UG_APP_MANUFACTURING":
            self.session.ApplicationSwitchImmediate("UG_APP_MANUFACTURING")
        if self.work_part.CAMSetup is None:
            self.work_part.CAMSetup.New()
        self.uf.Cam.InitSession()
        self.print_log("切换到加工环境", "SUCCESS")
        return True

    @contextmanager
    def undo_mark_context(self, name):
        """创建撤销标记上下文"""
        mark_id = self.session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, name)
        try:
            yield mark_id
        except Exception as e:
            self.print_log(f"执行 '{name}' 时发生错误: {e}", "ERROR")
            raise e
        finally:
            self.session.DeleteUndoMark(mark_id, None)

    def save_part_directly(self):
        """直接保存当前工作部件，不另存为新文件"""
        if CONFIG["AUTO_SAVE"]:
            try:
                # 直接保存当前工作部件
                self.work_part.Save(
                    NXOpen.BasePart.SaveComponents.TrueValue, 
                    NXOpen.BasePart.CloseAfterSave.FalseValue
                )
                self.print_log(f"刀具创建完成，已直接保存到: {self.work_part.FullPath}", "SUCCESS")
                return True
            except Exception as e:
                self.print_log(f"保存文件失败: {str(e)}", "ERROR")
                return False
        return True

    def load_mill_tools_from_json(self, json_path):
        """从JSON文件加载铣刀参数并创建所有刀具，按直径从大到小排序"""
        self.print_log(f"开始从JSON加载铣刀参数: {json_path}", "START")
        
        try:
            # 读取JSON文件
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 假设JSON结构为列表：[{"ToolName": "D10R0.5", "Diameter": 10.0, "Cor1Rad": 0.5, "Length": 50, "FluteLn": 30}, ...]
            # 或者字典：{"D10R0.5": {...}, ...}
            # 这里需要根据实际JSON结构调整，假设是之前项目常用的字典格式，key是刀具名，value是参数
            
            tool_list = []
            if isinstance(data, dict):
                 for name, params in data.items():
                     params['刀具名称'] = name
                     tool_list.append(params)
            elif isinstance(data, list):
                tool_list = data
            
            tool_data = []
            # 统一参数名并验证
            for item in tool_list:
                # 兼容中英文键名
                name = item.get('刀具名称') or item.get('ToolName')
                dia = item.get('直径') or item.get('Diameter')
                rad = item.get('R角') or item.get('Cor1Rad') or item.get('R1') or 0.0
                length = item.get('长度') or item.get('Length') or item.get('Height') or 50.0
                flute = item.get('刃长') or item.get('FluteLn') or item.get('FluteLength') or 30.0
                
                if name and dia is not None:
                     tool_data.append({
                         '刀具名称': str(name).strip(),
                         '直径': float(dia),
                         'R角': float(rad),
                         '长度': float(length),
                         '刃长': float(flute)
                     })

            # 记录刀具总数
            total_tools = len(tool_data)
            self.print_log(f"从JSON读取到 {total_tools} 个刀具参数", "INFO")
            
            # === 按直径从大到小排序 ===
            tool_data.sort(key=lambda x: x['直径'], reverse=True)
            
            self.created_count = 0
            self.skipped_count = 0
            
            # 遍历，创建刀具
            for index, row in enumerate(tool_data):
                tool_name = row['刀具名称']
                
                try:
                    diameter = row['直径']
                    R1 = row['R角']
                    length = row['长度']
                    flute_length = row['刃长']
                    
                    # 计算当前刀具的排序位置
                    position = index + 1
                    
                    tool = self.get_or_create_mill_tool(
                        tool_type="MILL",
                        diameter=diameter,
                        R1=R1,
                        length=length,
                        flute_length=flute_length,
                        parent_group_name="GENERIC_MACHINE", 
                        tool_name=tool_name
                    )
                    
                    if tool:
                        self.created_count += 1
                    else:
                        self.skipped_count += 1
                        
                except Exception as e:
                    self.print_log(f"❌ 创建刀具 {tool_name} 失败: {str(e)}", "ERROR")
                    self.skipped_count += 1
            
            self.print_log(f"刀具创建完成: 成功 {self.created_count} 个, 跳过 {self.skipped_count} 个", "SUCCESS")
            return True
            
        except Exception as e:
            self.print_log(f"读取JSON文件失败: {str(e)}", "ERROR")
            traceback.print_exc()
            return False

    def get_or_create_mill_tool(self, tool_type="MILL", diameter=1.0, R1=0.0,
                                length=50.0, flute_length=30.0,
                                parent_group_name="GENERIC_MACHINE", tool_name="milling_tool"):
        """获取或创建铣刀工具，如果已存在则更新参数"""
        
        try:
            # 获取父刀具组
            parent_group = self.work_part.CAMSetup.CAMGroupCollection.FindObject(parent_group_name)
            if parent_group is None:
                raise ValueError(f"未找到刀具组 {parent_group_name}")

            # 查找已有的铣刀
            tool_obj = None
            try:
                tool_obj = self.work_part.CAMSetup.CAMGroupCollection.FindObject(tool_name)
                # self.print_log(f"✔ 已找到铣刀工具: {tool_name}，将更新参数", "DEBUG")
            except Exception:
                # self.print_log(f"未找到铣刀工具: {tool_name}，将创建新刀具", "DEBUG")
                tool_obj = None

            # 如果刀具不存在，创建新刀具
            if tool_obj is None:
                tool_obj = self.work_part.CAMSetup.CAMGroupCollection.CreateTool(
                    parent_group,
                    "hole_making",  # 使用hole_making类别
                    tool_type,
                    NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue,
                    tool_name
                )

            # 创建铣刀的 Builder
            mill_builder = self.work_part.CAMSetup.CAMGroupCollection.CreateMillToolBuilder(tool_obj)

            # 设置参数 - 无论刀具是否已存在，都会设置这些参数
            mill_builder.TlDiameterBuilder.Value = diameter
            
            # 根据你的参考函数，R角应该使用TlCor1RadBuilder
            if hasattr(mill_builder, "TlCor1RadBuilder"):
                mill_builder.TlCor1RadBuilder.Value = R1
            elif hasattr(mill_builder, "TlR1Builder"):  # 备用属性名
                mill_builder.TlR1Builder.Value = R1
            
            # 根据你的参考函数，长度应该使用TlHeightBuilder
            if hasattr(mill_builder, "TlHeightBuilder"):
                mill_builder.TlHeightBuilder.Value = length
            elif hasattr(mill_builder, "TlLengthBuilder"):  # 备用属性名
                mill_builder.TlLengthBuilder.Value = length
            
            # 根据你的参考函数，刃长应该使用TlFluteLnBuilder
            if hasattr(mill_builder, "TlFluteLnBuilder"):
                mill_builder.TlFluteLnBuilder.Value = flute_length
            elif hasattr(mill_builder, "TlFluteLengthBuilder"):  # 备用属性名
                mill_builder.TlFluteLengthBuilder.Value = flute_length

            # 提交并销毁 Builder
            mill_builder.Commit()
            mill_builder.Destroy()

            return tool_obj

        except Exception as e:
            self.print_log(f"创建铣刀工具失败: {str(e)}", "ERROR")
            return None
        

    def print_summary(self):
        """打印刀具创建摘要"""
        # 简化摘要
        print(f"   [Summary] 刀具创建: 成功 {self.created_count}, 跳过/失败 {self.skipped_count}", flush=True)


# ==================================================================================
# 主流程
# ==================================================================================
def create_tools_workflow(part_path, json_path):
    """刀具创建主工作流"""
    session = NXOpen.Session.GetSession()
    base_part, load_status = session.Parts.OpenBaseDisplay(part_path)
    work_part = session.Parts.Work

    creator = ToolCreator(work_part)
    creator.print_header("NX CAM 刀具创建工具")
    creator.print_log(f"零件: {work_part.Name}", "INFO")
    
    # 切换到加工环境
    creator.switch_to_manufacturing()
    
    # 从JSON创建所有刀具
    success = creator.load_mill_tools_from_json(json_path)
    
    # 打印摘要
    creator.print_summary()
    
    # 直接保存当前工作部件，不另存为
    if success and CONFIG["AUTO_SAVE"]:
        save_success = creator.save_part_directly()
        if save_success:
            creator.print_log(f"刀具创建完成，文件已直接保存", "END")
        else:
            creator.print_log(f"文件保存失败", "ERROR")
    else:
        creator.print_log(f"刀具创建完成（未保存）", "INFO")
    
    # 清理资源
    if load_status:
        load_status.Dispose()
    
    return success


def process_part(work_part, json_path):
    """
    供 run_step8.py 调用的接口
    """
    try:
        creator = ToolCreator(work_part)
        
        # 确保在加工环境下
        # creator.switch_to_manufacturing() # Already done in main loop usually, but harmless to verify
        if creator.session.ApplicationName != "UG_APP_MANUFACTURING":
             creator.session.ApplicationSwitchImmediate("UG_APP_MANUFACTURING")
        
        # 创建刀具
        success = creator.load_mill_tools_from_json(json_path)
        
        # 打印摘要
        creator.print_summary()
        
        return success
    except Exception as e:
        print(f"❌ 刀具创建过程出错: {e}")
        return False

def main():
    """主函数"""
    try:
        success = create_tools_workflow(
            CONFIG["PART_PATH"],
            CONFIG["JSON_TOOLS_PATH"]
        )
        
        if not success:
            print("刀具创建失败，请检查错误信息。")
            return 1
            
        return 0
        
    except Exception as e:
        print(f"❌ 主程序异常: {e}", flush=True)
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit_code = main()
    exit(exit_code)