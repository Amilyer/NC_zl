# -*- coding: utf-8 -*-
"""
刀具创建模块 (create_tools.py)
功能：从Excel读取铣刀参数并创建所有刀具
"""

import time
import traceback

import NXOpen
import NXOpen.CAM
import NXOpen.UF
import pandas as pd


class ToolCreator:
    def __init__(self, work_part):
        self.work_part = work_part
        self.session = NXOpen.Session.GetSession()
        self.uf = NXOpen.UF.UFSession.GetUFSession()
        self.created_count = 0
        self.skipped_count = 0

    def print_log(self, message, level="INFO"):
        timestamp = time.strftime("%H:%M:%S")
        emoji_map = {
            "INFO": "ℹ️", "WARN": "⚠️", "ERROR": "❌",
            "SUCCESS": "✅", "DEBUG": "🔍", "START": "🚀", "END": "🏁"
        }
        emoji = emoji_map.get(level.upper(), "")
        print(f"[{timestamp}] {emoji} {message}")

    def switch_to_manufacturing(self):
        """切换到加工环境"""
        try:
            # 检查核心对象有效性
            if not self.session:
                self.print_log("会话对象无效", "ERROR")
                return False
                
            if not self.work_part or self.work_part.IsDisposed:
                self.print_log("工作部件无效或已释放", "ERROR")
                return False

            # 检查是否已经在制造模块
            module_name = self.session.ApplicationName
            if module_name != "UG_APP_MANUFACTURING":
                self.print_log(f"正在从 {module_name} 切换到 UG_APP_MANUFACTURING...", "INFO")
                self.session.ApplicationSwitchImmediate("UG_APP_MANUFACTURING")
                time.sleep(0.1)  # 短暂等待模块切换完成
            
            # 初始化 CAM 会话
            if not self.session.IsCamSessionInitialized():
                self.print_log("CAM 会话未初始化，正在启动...", "INFO")
                self.session.CreateCamSession()
                time.sleep(0.1)  # 等待初始化完成
                
            # 确保 Setup 存在
            cam_setup_ready = False
            try:
                if self.work_part.CAMSetup is not None:
                    cam_setup_ready = True
                    self.print_log("CAM Setup 已存在", "SUCCESS")
            except Exception as e:
                self.print_log(f"检查 CAMSetup 时出错: {e}", "WARN")

            if not cam_setup_ready:
                # 尝试创建默认 Setup，优先使用mill_contour更适合铣削操作
                self.print_log("正在创建 CAM Setup...", "INFO")
                setup_created = False
                for setup_type in ["mill_contour", "mill_planar", "hole_making"]:
                    try:
                        self.work_part.CreateCamSetup(setup_type)
                        self.print_log(f"✅ CAM Setup ({setup_type}) 创建成功。", "SUCCESS")
                        setup_created = True
                        break
                    except Exception as e:
                        self.print_log(f"⚠ 创建 {setup_type} Setup 失败: {e}", "WARN")
                
                if not setup_created:
                    self.print_log("❌ 所有类型的 Setup 创建均失败", "ERROR")
                    return False
            
            self.print_log("已切换到加工环境", "SUCCESS")
            return True
        except Exception as e:
            self.print_log(f"切换加工环境失败: {e}", "ERROR")
            traceback.print_exc()
            return False

    def load_mill_tools_from_excel(self, excel_path):
        """从Excel文件加载铣刀参数并创建所有刀具"""
        self.print_log(f"开始从Excel加载铣刀参数: {excel_path}", "START")
        
        try:
            # 读取Excel文件，跳过第一行，第二行作为列名
            # 使用 sheet_name=0 读取第一个工作表
            df = pd.read_excel(excel_path, sheet_name=0, header=1)
            
            # 修改：提取需要的列：刀具名称、直径、R角、长度、刃长
            required_columns = ['刀具名称', '直径', 'R角', '长度', '刃长']
            
            # 检查列是否存在
            if not all(col in df.columns for col in required_columns):
                missing_cols = [col for col in required_columns if col not in df.columns]
                self.print_log(f"Excel文件中缺少必要的列: {missing_cols}", "ERROR")
                return False
            
            # 过滤有效数据（去除空值）
            tool_data = df[required_columns].dropna()

            # === 按直径从大到小排序 ===
            tool_data = tool_data.sort_values(by='直径', ascending=False)
            
            # 显示排序信息
            diameters = tool_data['直径'].tolist()
            if diameters:
                self.print_log(f"刀具直径范围: {min(diameters):.2f}mm ~ {max(diameters):.2f}mm", "INFO")
                self.print_log(f"排序方式: 按直径从大到小 (降序)", "SUCCESS")
                
                # 显示排序后的前几个刀具
                sample_tools = tool_data.head(min(5, len(tool_data)))
                sample_info = ", ".join([f"{row['刀具名称']}({row['直径']}mm)" 
                                        for _, row in sample_tools.iterrows()])
                self.print_log(f"排序后前{len(sample_tools)}个刀具: {sample_info}", "DEBUG")
            # === 排序结束 ===
            
            self.created_count = 0
            self.skipped_count = 0
            
            # 遍历每一行，创建刀具
            for index, row in tool_data.iterrows():
                tool_name = str(row['刀具名称']).strip()
                
                # 跳过表头或无效行
                if tool_name == '刀具名称' or not tool_name:
                    continue
                
                try:
                    diameter = float(row['直径'])
                    R1 = float(row['R角'])
                    length = float(row['长度'])
                    flute_length = float(row['刃长'])
                    
                    # 修改：调用更新后的刀具创建函数，传入新参数
                    tool = self.get_or_create_mill_tool(
                        tool_type="MILL",
                        diameter=diameter,
                        R1=R1,
                        length=length,  # 新增参数
                        flute_length=flute_length,  # 新增参数
                        parent_group_name="GENERIC_MACHINE", 
                        tool_name=tool_name
                    )
                    
                    if tool:
                        self.created_count += 1
                        # 简化输出：不再逐个打印成功信息
                        # self.print_log(f"✅ 创建刀具: {tool_name} ...", "SUCCESS")
                    else:
                        self.skipped_count += 1
                        
                except Exception as e:
                    self.print_log(f"❌ 创建刀具 {tool_name} 失败: {str(e)}", "ERROR")
                    self.skipped_count += 1
            
            self.print_log(f"刀具创建完成: 成功 {self.created_count} 个, 跳过 {self.skipped_count} 个", "SUCCESS")
            return True
            
        except Exception as e:
            self.print_log(f"读取Excel文件失败: {str(e)}", "ERROR")
            return False

    def get_or_create_mill_tool(self, tool_type="MILL", diameter=1.0, R1=0.0,
                                length=50.0, flute_length=30.0,
                                parent_group_name="GENERIC_MACHINE", tool_name="milling_tool"):
        """获取或创建铣刀工具，如果已存在则更新参数"""
        
        try:
            # 获取父刀具组
            parent_group = None
            try:
                parent_group = self.work_part.CAMSetup.CAMGroupCollection.FindObject(parent_group_name)
            except:
                pass
                
            if parent_group is None:
                # 尝试查找任何可用的 MACHINE_TOOL 组
                for group in self.work_part.CAMSetup.CAMGroupCollection:
                    if isinstance(group, NXOpen.CAM.MachineTool):
                        parent_group = group
                        break
            
            if parent_group is None:
                try:
                    for group in self.work_part.CAMSetup.CAMGroupCollection:
                        if group.Type == NXOpen.CAM.CAMGroupType.MachineTool:
                            parent_group = group
                            break
                    if parent_group is None:
                        for group in self.work_part.CAMSetup.CAMGroupCollection:
                            if group.IsToolGroup():
                                parent_group = group
                                break
                    if parent_group is None:
                        raise ValueError(f"未找到刀具组 {parent_group_name} 且无法自动定位替代组")
                    print(f"⚠ 未找到指定刀具组 {parent_group_name}，使用替代组: {parent_group.Name}")
                except Exception as e:
                    raise ValueError(f"未找到刀具组 {parent_group_name}，错误: {str(e)}")

            # 查找已有的铣刀
            tool_obj = None
            try:
                tool_obj = self.work_part.CAMSetup.CAMGroupCollection.FindObject(tool_name)
                self.print_log(f"✔ 已找到铣刀工具: {tool_name}，将更新参数", "DEBUG")
            except Exception:
                self.print_log(f"未找到铣刀工具: {tool_name}，将创建新刀具", "DEBUG")
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
            
            # R角处理
            if hasattr(mill_builder, "TlCor1RadBuilder"):
                mill_builder.TlCor1RadBuilder.Value = R1
            elif hasattr(mill_builder, "TlR1Builder"):
                mill_builder.TlR1Builder.Value = R1
            
            # 长度处理
            if hasattr(mill_builder, "TlHeightBuilder"):
                mill_builder.TlHeightBuilder.Value = length
            elif hasattr(mill_builder, "TlLengthBuilder"):
                mill_builder.TlLengthBuilder.Value = length
            
            # 刃长处理
            if hasattr(mill_builder, "TlFluteLnBuilder"):
                mill_builder.TlFluteLnBuilder.Value = flute_length
            elif hasattr(mill_builder, "TlFluteLengthBuilder"):
                mill_builder.TlFluteLengthBuilder.Value = flute_length

            # 提交并销毁 Builder
            mill_builder.Commit()
            mill_builder.Destroy()

            return tool_obj

        except Exception as e:
            self.print_log(f"创建铣刀工具失败: {str(e)}", "ERROR")
            return None
