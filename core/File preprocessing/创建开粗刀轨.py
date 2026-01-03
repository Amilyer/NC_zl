#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
NX CAM自动化工具 - 刀轨生成模块
精简版：只保留行腔和往复等高工序
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
    "PART_PATH": r'C:\Projects\NC\output\05_Drilled_PRT\B1-01.prt',
    "AUTO_SAVE": True,
    "TEST_MODE": True,  # 设置为False时生成实际刀轨
    
    # JSON文件路径
    "JSON_CAVITY_PATH": r'C:\Projects\NC\output\06_CAM\Roughing_JSON\B1-01_行腔.json',
    "JSON_RECIPROCATING_PATH": r'C:\Projects\NC\output\B1-01_开粗_往复等高.json',
}


# ==================================================================================
# 操作模板配置
# ==================================================================================
OPERATION_CONFIGS = {
    "往复等高_SIMPLE": {
        "operation_name": "45#备料",
        "operation_type": "往复等高-D4",
        "operation_subtype": "往复等高-D4",
        "builder_type": "zlevel",
        "description": "往复等高精加工",
        "special_config": {
            "depth_per_cut": 0.0335,
            "part_stock": 0.102,
            "engage_closed_type": NXOpen.CAM.NcmPlanarEngRetBuilder.EngRetTypes.RampOnShape,
            # ============ 新增默认值 ============
            "spindle_rpm": 1700.0,
            "feed_per_tooth": 2000.0,
            "feed_rapid": 8000.0,
            # ===================================
             # 新增：步距类型和每刀深度参数
            "stepover_type": NXOpen.CAM.StepoverBuilder.StepoverTypes.Constant,  # 恒定步距
            "global_depth_per_cut": 0.1  # 默认每刀深度0.1mm
        }
    },
    "行腔_SIMPLE": {
        "operation_name": "45#备料",
        "operation_type": "行腔_D4",
        "operation_subtype": "行腔_D4_1",
        "builder_type": "cavity",
        "description": "型腔铣精加工",
        "special_config": {
            "cut_pattern": NXOpen.CAM.CutPatternBuilder.Types.FollowPeriphery,
            "stepover_percent": 70.0,
            "depth_per_cut": 0.5,
            "cut_direction": NXOpen.CAM.CutDirection.Types.Climb,
            # ============ 新增默认值 ============
            "spindle_rpm": 1700.0,
            "feed_per_tooth": 2000.0,
            "feed_rapid": 8000.0,
            # ===================================
            # 新增参数
            "reference_tool": None,
            "stepover_type": NXOpen.CAM.StepoverBuilder.StepoverTypes.Constant
        }
    }
}

# 测试用例
TEST_CASES = []


# ==================================================================================
# ToolpathGenerator 刀轨生成类
# ==================================================================================
class ToolpathGenerator:
    BUILDER_MAP = {
        'cavity': 'CreateCavityMillingBuilder',
        'zlevel': 'CreateZlevelMillingBuilder',
    }

    LAYER_TO_GEOMETRY = {
        20: "WORKPIECE_0",
        30: "WORKPIECE_1", 
        40: "WORKPIECE_2",
        50: "WORKPIECE_3",
        60: "WORKPIECE_4",
        70: "WORKPIECE_5"
    }

    LAYER_TO_PROGRAM_GROUP = {
        20: "正",
        30: "左",
        40: "右",
        50: "前", 
        60: "后",
        70: "反"
    }
    # ==================================================

    def __init__(self, work_part, save_dir=None):
        self.work_part = work_part
        self.session = NXOpen.Session.GetSession()
        self.uf = NXOpen.UF.UFSession.GetUFSession()
        self.operation_count = 0
        self.success_count = 0
        self.failed_count = 0
        self.test_results = []
        self.save_dir = save_dir  # 保存目录参数

    def create_rough_program_group(self):
        """创建开粗程序组及其子程序组（正、反、左、右、前、后）"""
        self.print_log("开始创建开粗程序组结构...", "START")
        template_name = "45#备料"
        try:
            with self.undo_mark_context("创建开粗程序组结构"):
                # 获取CAM设置
                cam_setup = self.work_part.CAMSetup
                cam_groups = cam_setup.CAMGroupCollection
                
                # 查找NC_PROGRAM根组（如果没有则使用PROGRAM）
                try:
                    nc_program_group = cam_groups.FindObject("NC_PROGRAM")
                except:
                    self.print_log("未找到NC_PROGRAM组，使用默认PROGRAM组", "WARN")
                    nc_program_group = cam_groups.FindObject("PROGRAM")
                
                # ============ 1. 创建或获取开粗程序组 ============
                rough_program_name = "开粗"
                try:
                    rough_program_group = cam_groups.FindObject(rough_program_name)
                    self.print_log(f"程序组 '{rough_program_name}' 已存在", "DEBUG")
                except:
                    rough_program_group = cam_groups.CreateProgram(
                        nc_program_group, 
                        template_name, 
                        "PROGRAM", 
                        NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue, 
                        rough_program_name
                    )
                    self.print_log(f"创建程序组: {rough_program_name}", "SUCCESS")
                
                # ============ 2. 创建子程序组（正、反、左、右、前、后） ============
                sub_groups = {}
            
                # 要创建的子程序组列表
                sub_group_names = ["正", "左", "右", "前", "后", "反"]
                
                for direction in sub_group_names:
                    # 根据方向获取对应的图层
                    layer_for_direction = None
                    for layer_num, dir_name in self.LAYER_TO_PROGRAM_GROUP.items():
                        if dir_name == direction:
                            layer_for_direction = layer_num
                            break
                    
                    if layer_for_direction is None:
                        # 如果没有找到对应的图层，使用默认图层20
                        layer_for_direction = 20
                        self.print_log(f"方向 '{direction}' 没有对应的图层映射，使用默认图层20", "WARN")
                    
                    # 构建子程序组名称：方向_开粗_图层
                    sub_name = f"{direction}_开粗_{layer_for_direction}"
                    
                    try:
                        # 尝试查找已存在的子程序组
                        sub_group = cam_groups.FindObject(sub_name)
                        self.print_log(f"子程序组 '{sub_name}' 已存在", "DEBUG")
                    except:
                        # 如果不存在则创建
                        sub_group = cam_groups.CreateProgram(
                            rough_program_group,  # 父组：开粗程序组
                            template_name, 
                            "PROGRAM", 
                            NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue, 
                            sub_name
                        )
                        self.print_log(f"创建子程序组: {sub_name}", "SUCCESS")
                    
                    sub_groups[direction] = sub_group
                
                self.print_log("开粗程序组结构创建/获取完成", "SUCCESS")
                
                # 返回主程序和子程序组的字典
                return {
                    "main_group": rough_program_group,
                    "sub_groups": sub_groups
                }
                    
        except Exception as e:
            self.print_log(f"创建开粗程序组结构失败: {e}", "ERROR")
            traceback.print_exc()
            return None

    def get_rough_program_group(self, layer=20):
        """获取开粗程序组，根据图层返回对应的子程序组
        参数:
            layer: 图层编号，默认为20
        返回:
            对应图层的子程序组对象
        """
        try:
            # 根据图层获取对应的方向
            direction = self.LAYER_TO_PROGRAM_GROUP.get(layer, "正")  # 默认使用"正"
            
            # 构建子程序组名称：方向_开粗_图层
            sub_group_name = f"{direction}_开粗_{layer}"
            
            cam_groups = self.work_part.CAMSetup.CAMGroupCollection
            
            # 首先尝试直接查找子程序组
            try:
                sub_group = cam_groups.FindObject(sub_group_name)
                self.print_log(f"使用程序组: 开粗/{sub_group_name} (图层{layer})", "DEBUG")
                return sub_group
            except:
                self.print_log(f"未找到子程序组 {sub_group_name}，尝试查找主开粗程序组", "WARN")
            
            # 如果找不到子程序组，尝试查找主开粗程序组
            try:
                main_group = cam_groups.FindObject("开粗")
                self.print_log(f"使用主开粗程序组: 开粗", "WARN")
                return main_group
            except:
                self.print_log(f"未找到主开粗程序组，使用默认PROGRAM", "ERROR")
                return cam_groups.FindObject("PROGRAM")
                
        except Exception as e:
            self.print_log(f"获取程序组失败: {e}，使用默认PROGRAM", "ERROR")
            try:
                return self.work_part.CAMSetup.CAMGroupCollection.FindObject("PROGRAM")
            except:
                return None

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

    def save_part(self, part_path):
        """保存零件文件"""
        if CONFIG["AUTO_SAVE"]:
            # 指定保存文件夹
            if self.save_dir:
                save_dir = self.save_dir
            else:
                save_dir = r'C:\Projects\NC\output\05_CAM\Daogui_prt'
            
            # 确保保存文件夹存在
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)
                self.print_log(f"创建保存文件夹: {save_dir}", "INFO")
            
            # 获取原始文件名
            original_dir, file_name = os.path.split(part_path)
            
            # 构建新的保存路径（不含时间戳）
            save_path = os.path.join(save_dir, file_name)
            
            # 保存文件
            try:
                # 检查保存路径是否与当前部件路径相同
                current_path = os.path.normpath(self.work_part.FullPath)
                target_path = os.path.normpath(save_path)
                
                if current_path.lower() == target_path.lower():
                    # 如果路径相同，直接保存
                    self.work_part.Save(NXOpen.BasePart.SaveComponents.TrueValue, NXOpen.BasePart.CloseModified.UseResponses)
                    self.print_log(f"部件已保存 (覆盖原文件): {save_path}", "SUCCESS")
                else:
                    # 如果路径不同，另存为
                    self.work_part.SaveAs(save_path)
                    self.print_log(f"刀轨生成完成，另存至: {save_path}", "SUCCESS")
            except Exception as e:
                self.print_log(f"保存失败: {e}", "ERROR")
                return part_path
            return save_path
        return part_path

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

    # ==================== 面查找方法 ====================
    def _find_faces_by_attr_id(self, target_ids, layer=None):
        """根据属性ID查找面，可选按图层过滤
        
        参数:
            target_ids: 要查找的属性ID列表（字符串列表）
            layer: 指定的图层编号（整数），None表示不进行图层过滤
        """
        found = []
        try:
            layer_filter_msg = f"图层{layer}" if layer is not None else "所有图层"
            self.print_log(f"开始查找ID为 {target_ids} 的面 ({layer_filter_msg})...", "DEBUG")
            
            # 遍历所有体，并按图层过滤
            bodies_to_search = []
            
            for body in self.work_part.Bodies:
                # 如果指定了图层，检查体是否在指定图层上
                if layer is not None:
                    # 获取体所在的图层
                    body_layer = body.Layer
                    if body_layer != layer:
                        continue  # 跳过不在指定图层上的体
                bodies_to_search.append(body)
            
            self.print_log(f"搜索范围内的体数量: {len(bodies_to_search)}", "DEBUG")
            
            # 遍历筛选后的体
            for body in bodies_to_search:
                # 遍历体的所有面
                for face in body.GetFaces():
                    try:
                        # 检查面是否有"FACE_TAG"属性
                        if face.HasUserAttribute("FACE_TAG", NXOpen.NXObject.AttributeType.String, -1):
                            face_id = face.GetStringAttribute("FACE_TAG")
                            if face_id in target_ids:
                                found.append(face)
                                # self.print_log(f"找到面: Tag={face.Tag}, FACE_TAG={face_id}, 图层={body.Layer}", "DEBUG")
                    except Exception as e:
                        continue  # 跳过无法读取属性的面
            
            self.print_log(f"通过属性ID找到 {len(found)} 个面 ({layer_filter_msg})", "DEBUG")
        except Exception as e:
            self.print_log(f"查找面时出错: {e}", "ERROR")
        
        return found

    def find_face_by_tag(self, face_tag):
        """根据Tag查找面"""
        try:
            return NXOpen.Utilities.NXObjectManager.Get(face_tag)
        except:
            return None

    def _get_valid_faces(self, inputs, layer=None):
        """获取有效面（支持Tag列表、ID列表、面对象列表），可选按图层过滤
        
        参数:
            inputs: 面输入列表
            layer: 指定的图层编号（整数），None表示不进行图层过滤
        """
        if not inputs:
            return []
        
        # 如果已经是面对象，检查是否符合图层要求
        if isinstance(inputs[0], NXOpen.Face):
            if layer is not None:
                # 过滤出指定图层的面
                filtered_faces = []
                for face in inputs:
                    # 获取面所在的体，从而获取图层信息
                    body = face.GetBody()
                    if body and body.Layer == layer:
                        filtered_faces.append(face)
                
                if filtered_faces:
                    self.print_log(f"过滤得到 {len(filtered_faces)} 个指定图层({layer})的面", "SUCCESS")
                    return filtered_faces
                else:
                    self.print_log(f"在指定图层({layer})上未找到有效面", "WARN")
                    return []
            else:
                return inputs
        
        # 如果是字符串列表，尝试作为属性ID查找
        if isinstance(inputs[0], str):
            self.print_log(f"按属性ID查找面: {inputs} (图层={layer if layer is not None else '所有'})", "DEBUG")
            faces_by_id = self._find_faces_by_attr_id(inputs, layer)
            if faces_by_id:
                layer_msg = f"图层{layer}" if layer is not None else "所有图层"
                self.print_log(f"通过属性ID在{layer_msg}上找到 {len(faces_by_id)} 个面", "SUCCESS")
                return faces_by_id
            else:
                # 如果没有找到，尝试将字符串作为Tag处理
                self.print_log("尝试将输入作为Tag处理...", "DEBUG")
                try:
                    tag_inputs = [int(tag) for tag in inputs]
                    faces_by_tag = []
                    for tag in tag_inputs:
                        face = self.find_face_by_tag(tag)
                        if face:
                            # 检查图层过滤
                            if layer is not None:
                                body = face.GetBody()
                                if body and body.Layer == layer:
                                    faces_by_tag.append(face)
                            else:
                                faces_by_tag.append(face)
                    
                    if faces_by_tag:
                        layer_msg = f"图层{layer}" if layer is not None else "所有图层"
                        self.print_log(f"通过Tag在{layer_msg}上找到 {len(faces_by_tag)} 个面", "SUCCESS")
                        return faces_by_tag
                except ValueError:
                    pass
        
        # 如果是整数列表，作为Tag处理
        elif isinstance(inputs[0], int):
            faces_by_tag = []
            for tag in inputs:
                face = self.find_face_by_tag(tag)
                if face:
                    # 检查图层过滤
                    if layer is not None:
                        body = face.GetBody()
                        if body and body.Layer == layer:
                            faces_by_tag.append(face)
                    else:
                        faces_by_tag.append(face)
            
            if faces_by_tag:
                layer_msg = f"图层{layer}" if layer is not None else "所有图层"
                self.print_log(f"通过Tag在{layer_msg}上找到 {len(faces_by_tag)} 个面", "SUCCESS")
                return faces_by_tag
        
        self.print_log(f"未找到任何有效面: {inputs} (图层={layer if layer is not None else '所有'})", "WARN")
        return []
    

    def _set_geometry_with_one_set(self, builder, faces):
        """设置一个几何集（用于往复等高工序）"""
        try:
            # 初始化几何数据
            builder.CutAreaGeometry.InitializeData(False)
            
            # 获取几何列表
            geometry_list = builder.CutAreaGeometry.GeometryList
            
            # 获取第一个几何集（默认的）
            geometry_set = geometry_list.FindItem(0)
            
            # 设置面到几何集
            if faces:
                rule_opts = self.work_part.ScRuleFactory.CreateRuleOptions()
                rule_opts.SetSelectedFromInactive(False)
                
                # 创建面选择规则
                rule = self.work_part.ScRuleFactory.CreateRuleFaceDumb(faces, rule_opts)
                rule_opts.Dispose()
                
                # 替换规则
                geometry_set.ScCollector.ReplaceRules([rule], False)
                
                self.print_log(f"设置 {len(faces)} 个面到几何集", "DEBUG")
            
            return True
            
        except Exception as e:
            self.print_log(f"设置几何集失败: {e}", "ERROR")
            return False

    # ==================== JSON测试用例加载 ====================
    def load_cavity_assignments_from_json(self, json_path):
        """从JSON文件加载行腔工序分配结果并生成测试用例"""
        self.print_log(f"读取行腔工序分配JSON文件: {json_path}", "START")
        
        try:
            # 读取JSON文件
            with open(json_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
            
            self.print_log(f"成功读取JSON，共 {len(data)} 个行腔组", "SUCCESS")
            
            test_cases = []

            # 用于跟踪每个图层出现的次数
            layer_counter = {}
            
            for group_name, group_data in data.items():
                try:
                    # 提取关键数据
                    operation_type = group_data.get('工序', '行腔_SIMPLE')
                    normal_face_ids = group_data.get('普通面ID列表', [])
                    yellow_face_ids = group_data.get('黄色面ID列表', [])
                    tool_name = group_data['刀具名称']
                    
                    # 提取行腔特定参数
                    depth_per_cut = group_data.get('切深', 0.5)
                    depth_per_cut = float(depth_per_cut)
                    reference_tool = group_data.get('参考刀具', None)
                    layer = group_data.get('指定图层', 20)
                    
                    # 其他行腔相关参数
                    stepover_percent = group_data.get('步距百分比', 70.0)

                    # ============ 新增：读取转速、进给、横越参数 ============
                    spindle_rpm = group_data.get('转速', 1700.0)  # 默认值1700
                    feed_per_tooth = group_data.get('进给', 2000.0)  # 默认值2000
                    feed_rapid = group_data.get('横越', 8000.0)  # 默认值8000
                    # ====================================================

                    # ============ 新增：读取最终余量参数 ============
                    final_stock = group_data.get('最终余量', 0.8)  # 默认值0.8mm
                    final_stock = float(final_stock)  # 确保是浮点数
                    # ==============================================

                    # ============ 修改：根据图层判断是否使用圆弧进刀 ============
                    # 统计当前图层出现的次数
                    if layer not in layer_counter:
                        layer_counter[layer] = 1
                    else:
                        layer_counter[layer] += 1
                    
                    # 如果是图层中第一个出现的行腔组，不设置圆弧进刀，后续的设置为圆弧进刀
                    use_arc_engagement = (layer_counter[layer] > 1)
                    # ==========================================================
                    
                    
                    # 将面ID列表中的整数转换为字符串
                    normal_face_ids_str = [str(face_id) for face_id in normal_face_ids]
                    yellow_face_ids_str = [str(face_id) for face_id in yellow_face_ids]
                    
                    # 创建测试用例
                    test_case = (
                        operation_type, 
                        {
                        "normal_faces": normal_face_ids_str,  # 普通面
                        "yellow_faces": yellow_face_ids_str   # 黄色面
                        }, 
                        tool_name, 
                        {
                            "max_depth": depth_per_cut,
                            "reference_tool": reference_tool,
                            "layer": layer,
                            "stepover_percent": stepover_percent,
                            # ============ 新增：传递转速、进给、横越参数 ============
                            "spindle_rpm": spindle_rpm,
                            "feed_per_tooth": feed_per_tooth,
                            "feed_rapid": feed_rapid,
                            # ====================================================
                            # ============ 新增：传递最终余量参数 ============
                            "final_stock": final_stock,  # 传递最终余量
                            # ==============================================
                            # ============ 修改：传递圆弧进刀参数 ============
                            "use_arc_engagement": use_arc_engagement,  # 传递圆弧进刀参数
                            # ==============================================
                        }
                    )
                    
                    test_cases.append(test_case)
                    
                    # 详细日志
                    log_msg = f"行腔组 '{group_name}': 刀具={tool_name}, "
                    log_msg += f"普通面数量={len(normal_face_ids_str)}, "
                    log_msg += f"黄色面数量={len(yellow_face_ids_str)}, "
                    log_msg += f"最大深度={depth_per_cut}mm"
                    if reference_tool:
                        log_msg += f", 参考刀具={reference_tool}"
                    log_msg += f", 步距={stepover_percent}%"
                    log_msg += f", 转速={spindle_rpm}RPM, 进给={feed_per_tooth}mm/齿, 横越={feed_rapid}mm/min"
                    log_msg += f", 圆弧进刀={use_arc_engagement}"

                    self.print_log(log_msg, "DEBUG")
                    
                except Exception as e:
                    self.print_log(f"解析行腔组 '{group_name}' 数据失败: {e}", "ERROR")
                    continue
            
            self.print_log(f"成功生成 {len(test_cases)} 个行腔测试用例", "SUCCESS")
            return test_cases
            
        except Exception as e:
            self.print_log(f"读取行腔JSON文件失败: {e}", "ERROR")
            return []

    def load_reciprocating_zlevel_assignments_from_json(self, json_path):
        """
        从JSON文件加载往复等高刀具分配结果并生成测试用例
        参数:
            json_path: JSON文件路径
        返回:
            test_cases: 生成的测试用例列表
        """
        self.print_log(f"读取往复等高分配JSON文件: {json_path}", "START")
        
        try:
            # 读取JSON文件
            with open(json_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
            
            self.print_log(f"成功读取JSON，共 {len(data)} 个往复等高组", "SUCCESS")
            
            
            test_cases = []
            
            for group_name, group_data in data.items():
                try:
                    # 提取关键数据
                    operation_type = group_data['工序']
                    face_ids = group_data['面ID列表']
                    tool_name = group_data['刀具名称']
                    global_depth_per_cut = group_data['切深']  # 每刀深度
                    layer = group_data.get('指定图层', 20)  # 默认图层20
                    reference_tool = group_data.get('参考刀具', None)
                    # ============ 新增：读取转速、进给、横越参数 ============
                    spindle_rpm = group_data.get('转速', 1700.0)  # 默认值1700
                    feed_per_tooth = group_data.get('进给', 2000.0)  # 默认值2000
                    feed_rapid = group_data.get('横越', 8000.0)  # 默认值8000
                    # ====================================================

                    part_stock = group_data.get('部件侧面余量',0.0)  # 默认侧面余量
                    floor_stock = group_data.get('部件底面余量',0.0)   # 默认底面余量
                    
                    # 处理参考刀具（如果为"NULL"则设为None）
                    if reference_tool == "NULL":
                        reference_tool = None
                    
                    # 将面ID列表中的整数转换为字符串
                    face_ids_str = [str(face_id) for face_id in face_ids]
                    
                    # 创建测试用例
                    # 格式: (operation_type, face_ids, tool_name, extra_params)
                    test_case = (
                        operation_type,  # 使用JSON中的工序类型
                        face_ids_str, 
                        tool_name, 
                        {
                            "global_depth_per_cut": global_depth_per_cut,  # 每刀深度
                            "stepover_type": NXOpen.CAM.StepoverBuilder.StepoverTypes.Constant,  # 恒定步距
                            "layer": layer,
                            "reference_tool": reference_tool,
                            # ============ 新增：传递转速、进给、横越参数 ============
                            "spindle_rpm": spindle_rpm,
                            "feed_per_tooth": feed_per_tooth,
                            "feed_rapid": feed_rapid,
                            # ====================================================
                            # ============ 新增：传递余量参数 ============
                            "part_stock": part_stock,
                            "floor_stock": floor_stock,
                            # ===========================================
                        }
                    )
                    
                    test_cases.append(test_case)
                    
                    self.print_log(
                        f"往复等高组 '{group_name}': 刀具={tool_name}, "
                        f"面数量={len(face_ids_str)}, 每刀深度={global_depth_per_cut}mm, 图层={layer}, "
                        f"参考刀具={reference_tool if reference_tool else '无'}, "
                        f"转速={spindle_rpm}RPM, 进给={feed_per_tooth}mm/齿, "
                        f"横越={feed_rapid}mm/min, "
                        f"侧面余量={part_stock}mm, 底面余量={floor_stock}mm, ", 
                        "DEBUG"
                    )
                    
                except Exception as e:
                    self.print_log(f"解析往复等高组 '{group_name}' 数据失败: {e}", "ERROR")
                    continue
            
            self.print_log(f"成功生成 {len(test_cases)} 个往复等高测试用例", "SUCCESS")
            return test_cases
            
        except Exception as e:
            self.print_log(f"读取往复等高JSON文件失败: {e}", "ERROR")
            return []

    # ==================== 工序创建核心方法 ====================
    def _set_geometry_with_two_sets(self, builder, normal_faces, yellow_faces,final_stock=0.8):
        """设置两个几何集，分别对应普通面和黄色面，并设置不同余量"""
        try:
            # 初始化几何数据
            builder.CutAreaGeometry.InitializeData(False)
            
            # 获取几何列表
            geometry_list = builder.CutAreaGeometry.GeometryList
            
            # 获取第一个几何集（默认的）
            geometry_set1 = geometry_list.FindItem(0)
            
            # 设置普通面到第一个几何集
            if normal_faces:
                rule_opts1 = self.work_part.ScRuleFactory.CreateRuleOptions()
                rule_opts1.SetSelectedFromInactive(False)
                
                # 创建面选择规则
                rule1 = self.work_part.ScRuleFactory.CreateRuleFaceDumb(normal_faces, rule_opts1)
                rule_opts1.Dispose()
                
                # 替换规则
                geometry_set1.ScCollector.ReplaceRules([rule1], False)
                
                # 设置自定义余量：普通面余量为0mm
                geometry_set1.CustomStock = True
                geometry_set1.FinalStock = 0.0
                
                self.print_log(f"设置 {len(normal_faces)} 个普通面到几何集1，余量=0mm", "SUCCESS")
            
            # 创建第二个几何集用于黄色面
            geometry_set2 = builder.CutAreaGeometry.CreateGeometrySet()
            geometry_list.Append(geometry_set2)
            
            # 设置黄色面到第二个几何集
            if yellow_faces:
                rule_opts2 = self.work_part.ScRuleFactory.CreateRuleOptions()
                rule_opts2.SetSelectedFromInactive(False)
                
                # 创建面选择规则
                rule2 = self.work_part.ScRuleFactory.CreateRuleFaceDumb(yellow_faces, rule_opts2)
                rule_opts2.Dispose()
                
                # 替换规则
                sc_collector2 = geometry_set2.ScCollector
                sc_collector2.ReplaceRules([rule2], False)
                
                # 设置自定义余量：黄色面余量为0.8mm
                geometry_set2.CustomStock = True
                geometry_set2.FinalStock = final_stock
                
                self.print_log(f"设置 {len(yellow_faces)} 个黄色面到几何集2，余量=0.8mm", "SUCCESS")
            
            return True
            
        except Exception as e:
            self.print_log(f"设置双几何集失败: {e}", "ERROR")
            return False

    def generate_toolpath(self, operation):
        """生成刀轨"""
        if not CONFIG["TEST_MODE"]:
            try:
                self.work_part.CAMSetup.GenerateToolPath([operation])
                self.print_log(f"刀轨生成完成: {operation.Name}", "SUCCESS")
                return operation
            except Exception as e:
                self.print_log(f"刀轨生成警告: {e}", "WARN")
                return operation
        else:
            self.print_log(f"跳过刀路生成（测试模式）: {operation.Name}", "DEBUG")
            return operation

    def _finalize_operation(self, operation, tool_name):
        """完成操作设置"""
        try:
            tool = self.work_part.CAMSetup.CAMGroupCollection.FindObject(tool_name)
            self.work_part.CAMSetup.MoveObjects(
                NXOpen.CAM.CAMSetup.View.MachineTool, [operation], tool, NXOpen.CAM.CAMSetup.Paste.Inside
            )
            self.print_log(f"移动到刀具组: {tool_name}", "DEBUG")
        except:
            self.print_log(f"未找到刀具 {tool_name}，跳过移动", "WARN")

        return self.generate_toolpath(operation)

    def _configure_auto_clearance(self, builder, safe_distance=50.0):
        """配置自动安全平面"""
        try:
            if hasattr(builder, 'NonCuttingBuilder'):
                if hasattr(builder.NonCuttingBuilder, 'ClearanceBuilder'):
                    builder.NonCuttingBuilder.ClearanceBuilder.ClearanceType = NXOpen.CAM.NcmClearanceBuilder.ClearanceTypes.Automatic
                    builder.NonCuttingBuilder.ClearanceBuilder.SafeDistance = safe_distance
                elif hasattr(builder.NonCuttingBuilder, 'TransferCommonClearanceBuilder'):
                    builder.NonCuttingBuilder.TransferCommonClearanceBuilder.ClearanceType = NXOpen.CAM.NcmClearanceBuilder.ClearanceTypes.Automatic
                    builder.NonCuttingBuilder.TransferCommonClearanceBuilder.SafeDistance = safe_distance
                elif hasattr(builder.NonCuttingBuilder, 'CommonClearanceBuilder'):
                    builder.NonCuttingBuilder.CommonClearanceBuilder.ClearanceType = NXOpen.CAM.NcmClearanceBuilder.ClearanceTypes.Automatic
                    builder.NonCuttingBuilder.CommonClearanceBuilder.SafeDistance = safe_distance
            self.print_log(f"设置自动安全平面，安全距离: {safe_distance}mm", "SUCCESS")
        except Exception as e:
            self.print_log(f"安全平面设置警告: {e}", "DEBUG")

    # ==================== 特定工序参数配置 ====================
    def _configure_zlevel_params(self, builder, config):
        """配置深度轮廓铣参数"""
        try:
            special = config.get('special_config', {})

            if config['operation_type'] == "往复等高-D4":
                # 往复等高设置
                builder.CutLevel.RangeType = special.get('cut_level_range', NXOpen.CAM.CutLevel.RangeTypes.Automatic)
                
                # ============ 新增：设置步距类型和每刀深度 ============
                stepover_type = special.get('stepover_type', NXOpen.CAM.StepoverBuilder.StepoverTypes.Constant)
                builder.CutLevel.GlobalDepthPerCut.StepoverType = stepover_type
                
                global_depth = special.get('global_depth_per_cut', 0.1)
                builder.CutLevel.GlobalDepthPerCut.DistanceBuilder.Value = global_depth
                self.print_log(f"设置往复等高每刀深度: {global_depth}mm", "SUCCESS")
                # ====================================================

                # ============ 新增：设置转速、进给、横越参数 ============
                # 设置主轴转速
                if 'spindle_rpm' in special:
                    builder.FeedsBuilder.SpindleRpmBuilder.Value = special['spindle_rpm']
                    self.print_log(f"设置主轴转速: {special['spindle_rpm']}RPM", "SUCCESS")
                
                # 设置每齿进给
                if 'feed_per_tooth' in special:
                    builder.FeedsBuilder.FeedPerToothBuilder.Value = special['feed_per_tooth']
                    self.print_log(f"设置每齿进给: {special['feed_per_tooth']}mm/齿", "SUCCESS")
                
                # 设置横越速度
                if 'feed_rapid' in special:
                    builder.FeedsBuilder.FeedRapidOutput.Value = NXOpen.CAM.FeedRapidOutputMode.G1
                    builder.FeedsBuilder.FeedRapidOutput.InheritanceStatus = False
                    builder.FeedsBuilder.FeedRapidBuilder.Value = special['feed_rapid']
                    self.print_log(f"设置横越速度: {special['feed_rapid']}mm/min", "SUCCESS")
                # ====================================================
                # ============ 新增：硬编码设置进刀参数 ============
                try:
                    # 1. 设置封闭区域进刀方式为"沿形状斜进刀"
                    if hasattr(builder.NonCuttingBuilder, 'EngageClosedAreaBuilder'):
                        builder.NonCuttingBuilder.EngageClosedAreaBuilder.EngRetType = NXOpen.CAM.NcmPlanarEngRetBuilder.EngRetTypes.RampOnShape
                        # 设置螺旋斜坡角度为1.0度
                        if hasattr(builder.NonCuttingBuilder.EngageClosedAreaBuilder, 'HelicalRampAngleBuilder'):
                            builder.NonCuttingBuilder.EngageClosedAreaBuilder.HelicalRampAngleBuilder.Value = 1.0
                            self.print_log("设置封闭区域进刀方式: 沿形状斜进刀，斜坡角度1.0°", "SUCCESS")
                    
                    # 2. 设置开放区域进刀方式为"圆弧进刀"，半径5.0mm
                    if hasattr(builder.NonCuttingBuilder, 'EngageOpenAreaBuilder'):
                        # 首先确保进刀类型是正确的
                        builder.NonCuttingBuilder.EngageOpenAreaBuilder.EngRetType = NXOpen.CAM.NcmPlanarEngRetBuilder.EngRetTypes.Arc
                        # 设置圆弧半径
                        if hasattr(builder.NonCuttingBuilder.EngageOpenAreaBuilder, 'RadiusBuilder'):
                            builder.NonCuttingBuilder.EngageOpenAreaBuilder.RadiusBuilder.Value = 5.0
                            self.print_log("设置开放区域进刀方式: 圆弧进刀，半径5.0mm", "SUCCESS")
                    
                    # 3. 设置退刀方式与进刀一致
                    if hasattr(builder.NonCuttingBuilder, 'RetractClosedAreaBuilder'):
                        builder.NonCuttingBuilder.RetractClosedAreaBuilder.RetractType = NXOpen.CAM.NcmPlanarEngRetBuilder.EngRetTypes.Linear
                    
                    if hasattr(builder.NonCuttingBuilder, 'RetractOpenAreaBuilder'):
                        builder.NonCuttingBuilder.RetractOpenAreaBuilder.RetractType = NXOpen.CAM.NcmPlanarEngRetBuilder.EngRetTypes.Arc
                        if hasattr(builder.NonCuttingBuilder.RetractOpenAreaBuilder, 'RadiusBuilder'):
                            builder.NonCuttingBuilder.RetractOpenAreaBuilder.RadiusBuilder.Value = 5.0
                except Exception as e:
                    self.print_log(f"设置进刀参数时出错: {e}", "WARN")
                # ====================================================

                # ============ 新增：分别设置余量参数 ============
                try:
                    # 获取余量参数
                    part_stock = special.get('part_stock', 0.102)  # 侧面余量，使用原来的默认值
                    floor_stock = special.get('floor_stock', 0.0)  # 底面余量，使用原来的默认值
                    
                    # 取消底面与侧面余量一致（基础关闭）
                    if hasattr(builder.CutParameters, "FloorSameAsPartStock"):
                        builder.CutParameters.FloorSameAsPartStock = False
                    
                    # 设置侧面余量
                    builder.CutParameters.PartStock.Value = part_stock
                    
                    # 显式写入底面余量，防止被继承/默认覆盖
                    if hasattr(builder.CutParameters.FloorStock, "InheritanceStatus"):
                        builder.CutParameters.FloorStock.InheritanceStatus = False
                    builder.CutParameters.FloorStock.Value = floor_stock
                    
                    self.print_log(f"设置部件侧面余量: {part_stock}mm, 底面余量: {floor_stock}mm", "SUCCESS")
                except Exception as e:
                    self.print_log(f"设置余量参数失败: {e}", "WARN")
                # ===========================================

                # ======================================================================================================
                # 1. 启用边缘延伸
                try:
                    builder.CutParameters.ExtendAtEdges.Status = True
                    builder.CutParameters.ExtendAtEdges.Distance.Intent = NXOpen.CAM.ParamValueIntent.PartUnits
                    builder.CutParameters.ExtendAtEdges.Distance.Value = 2.0
                    self.print_log("启用边缘延伸功能并设为 2.0mm", "SUCCESS")
                except Exception as e:
                    self.print_log(f"设置边缘延伸参数失败: {e}", "WARN")
                
                # 2. 立即 Commit 一次（模仿宏：先让当前段生效）
                try:
                    builder.Commit()          # 第一次提交
                    self.print_log("特殊文件：第一次 Commit 完成", "DEBUG")
                except Exception as e:
                    self.print_log(f"第一次 Commit 失败: {e}", "WARN")

                # 3. 先把范围改成“用户定义”，宏里随后能写 TopHeight
                try:
                    builder.CutLevel.RangeType = NXOpen.CAM.CutLevel.RangeTypes.UserDefined
                    self.print_log("已切换切削层范围为“用户定义”", "DEBUG")
                except Exception as e:
                    self.print_log(f"切换范围类型失败: {e}", "WARN")

                # 4. 重新拿句柄，写顶部高度（宏里 TopZc 的等价物）
                try:
                    builder.CutLevel.TopZc = 0.0
                    # builder.CutLevel.InitializeData()   # 宏里紧接着的动作
                    self.print_log("设置顶部高度 = 0.0 并 InitializeData", "SUCCESS")
                except Exception as e:
                    self.print_log(f"设置顶部高度失败: {e}", "WARN")
            
                # =============================================================================================  

        except Exception as e:
            self.print_log(f"深度轮廓铣参数配置警告: {e}", "DEBUG")

    def _configure_cavity_params(self, builder, config, use_arc_engagement=False):
        """配置行腔铣参数"""
        try:
            special = config.get('special_config', {})

            # 设置切削模式
            builder.CutPattern.CutPattern = special.get('cut_pattern')
            # 设置步距
            builder.BndStepover.StepoverType = NXOpen.CAM.StepoverBuilder.StepoverTypes.PercentToolFlat
            builder.BndStepover.PercentToolFlatBuilder.Value = special.get('stepover_percent', 70.0)

            # 设置最大加工深度和步距类型
            builder.CutLevel.GlobalDepthPerCut.StepoverType = special.get(
                'stepover_type', 
                NXOpen.CAM.StepoverBuilder.StepoverTypes.Constant
            )
            max_depth = special.get('global_depth_per_cut', 10.1)
            builder.CutLevel.GlobalDepthPerCut.DistanceBuilder.Value = max_depth
            
            self.print_log(f"设置型腔铣最大加工深度: {max_depth}mm", "SUCCESS")

            builder.CutParameters.CutDirection.Type = special.get('cut_direction')
            builder.CutParameters.CutOrder = NXOpen.CAM.CutParametersCutOrderTypes.DepthFirst

            # ============ 新增：设置转速、进给、横越参数 ============
            # 设置主轴转速
            if 'spindle_rpm' in special:
                builder.FeedsBuilder.SpindleRpmBuilder.Value = special['spindle_rpm']
                self.print_log(f"设置主轴转速: {special['spindle_rpm']}RPM", "SUCCESS")
            
            # 设置每齿进给
            if 'feed_per_tooth' in special:
                builder.FeedsBuilder.FeedPerToothBuilder.Value = special['feed_per_tooth']
                self.print_log(f"设置每齿进给: {special['feed_per_tooth']}mm/齿", "SUCCESS")
            
            # 设置横越速度
            if 'feed_rapid' in special:
                builder.FeedsBuilder.FeedRapidOutput.Value = NXOpen.CAM.FeedRapidOutputMode.G1
                builder.FeedsBuilder.FeedRapidOutput.InheritanceStatus = False
                builder.FeedsBuilder.FeedRapidBuilder.Value = special['feed_rapid']
                self.print_log(f"设置横越速度: {special['feed_rapid']}mm/min", "SUCCESS")
            # ====================================================

            # 设置参考刀具
            reference_tool_name = special.get('reference_tool')
            if reference_tool_name and reference_tool_name != "无":
                try:
                    reference_tool = self.work_part.CAMSetup.CAMGroupCollection.FindObject(reference_tool_name)
                    if reference_tool:
                        builder.ReferenceTool = reference_tool
                        self.print_log(f"设置参考刀具: {reference_tool_name}", "SUCCESS")
                    else:
                        self.print_log(f"未找到参考刀具: {reference_tool_name}", "WARN")
                except Exception as e:
                    self.print_log(f"设置参考刀具失败: {e}", "WARN")

            # ============ 新增：根据是否为槽设置开放区域进刀参数 ============
            if use_arc_engagement:
                try:
                    # 检查是否有EngageOpenAreaBuilder属性
                    if hasattr(builder.NonCuttingBuilder, 'EngageOpenAreaBuilder'):
                        # 设置进刀类型为圆弧
                        builder.NonCuttingBuilder.EngageOpenAreaBuilder.EngRetType = NXOpen.CAM.NcmPlanarEngRetBuilder.EngRetTypes.Arc
                        self.print_log("设置开放区域进刀类型: 圆弧", "SUCCESS")
                        
                        # 设置半径类型为"刀具直径百分比"
                        builder.NonCuttingBuilder.EngageOpenAreaBuilder.RadiusBuilder.Intent = NXOpen.CAM.ParamValueIntent.ToolDep
                        # 设置半径值为50（即50%）
                        builder.NonCuttingBuilder.EngageOpenAreaBuilder.RadiusBuilder.Value = 50.0
                        self.print_log("设置圆弧半径: 刀具直径的50%", "SUCCESS")
                    else:
                        self.print_log("当前builder没有EngageOpenAreaBuilder属性", "WARN")
                except Exception as e:
                    self.print_log(f"设置开放区域进刀参数失败: {e}", "WARN")
            else:
                self.print_log("非槽加工，使用默认进刀参数", "DEBUG")
            # ==============================================================

        except Exception as e:
            self.print_log(f"型腔铣参数配置警告: {e}", "DEBUG")

    def create_operation(self, operation_key, face_inputs, tool_name, **params):
        """
        通用工序创建方法
        """
        if operation_key not in OPERATION_CONFIGS:
            raise ValueError(f"未知的工序类型: {operation_key}")

        config = OPERATION_CONFIGS[operation_key].copy()

        if params:
            config['special_config'] = config.get('special_config', {}).copy()
            
            # 更新通用参数
            depth_val = None
            if 'global_depth_per_cut' in params:
                depth_val = params['global_depth_per_cut']
            elif 'depth_per_cut' in params:
                depth_val = params['depth_per_cut']
            elif 'max_depth' in params: # 兼容行腔的 max_depth
                depth_val = params['max_depth']

            if depth_val is not None:
                # 统一写入 'global_depth_per_cut'，不再写 'depth_per_cut'
                config['special_config']['global_depth_per_cut'] = depth_val

            if 'reference_tool' in params:
                config['special_config']['reference_tool'] = params['reference_tool']
            if 'stepover_type' in params:
                config['special_config']['stepover_type'] = params['stepover_type']

            # ============ 新增：处理转速、进给、横越参数 ============
            if 'spindle_rpm' in params:
                config['special_config']['spindle_rpm'] = params['spindle_rpm']
            if 'feed_per_tooth' in params:
                config['special_config']['feed_per_tooth'] = params['feed_per_tooth']
            if 'feed_rapid' in params:
                config['special_config']['feed_rapid'] = params['feed_rapid']
            # ====================================================


            # ============ 新增：处理图层参数 ============
            if 'layer' in params:
                config['special_config']['layer'] = params['layer']
            # ===========================================

            # ============ 新增：处理最终余量参数 ============
            if 'final_stock' in params:
                # 注意：这里只是保存到config，但不用于全局余量设置
                config['special_config']['final_stock'] = params['final_stock']
            # ===============================================


            # ============ 新增：处理圆弧进刀参数 ============
            if 'use_arc_engagement' in params:
                config['special_config']['use_arc_engagement'] = params['use_arc_engagement']
            # ==============================================


            # ============ 新增：处理余量参数 ============
            if 'part_stock' in params:
                config['special_config']['part_stock'] = params['part_stock']
            if 'floor_stock' in params:
                config['special_config']['floor_stock'] = params['floor_stock']
            
            # ====================================================

            # 处理行腔特有参数
            if operation_key == "行腔_SIMPLE":
                if 'stepover_percent' in params:
                    config['special_config']['stepover_percent'] = params['stepover_percent']

        self.print_log(f"创建 {operation_key} 工序", "START")
        self.operation_count += 1

        try:
            # ==================== 根据工序类型处理面输入 ====================
            if operation_key == "行腔_SIMPLE":
                # 行腔工序：face_inputs是字典，包含normal_faces和yellow_faces
                normal_faces_input = face_inputs.get("normal_faces", [])
                yellow_faces_input = face_inputs.get("yellow_faces", [])
                final_stock = params.get('final_stock', 0.8)  # 从参数获取最终余量
            else:
                # 往复等高工序：face_inputs是面ID列表
                normal_faces_input = face_inputs  # 直接使用列表
                yellow_faces_input = []
                final_stock = 0.0  # 往复等高不使用最终余量
            # 1. 获取有效面对象

            layer = params.get('layer', 20)  # 从参数中获取图层
            normal_faces = self._get_valid_faces(normal_faces_input,layer) if normal_faces_input else []
            yellow_faces = self._get_valid_faces(yellow_faces_input,layer) if yellow_faces_input else []
            
            self.print_log(f"找到 {len(normal_faces)} 个普通面，{len(yellow_faces)} 个黄色面", "DEBUG")

            with self.undo_mark_context(f"创建{operation_key}"):
                groups = self.work_part.CAMSetup.CAMGroupCollection

                # 2. 创建操作
                try:
                    tool_group = groups.FindObject(tool_name)
                except:
                    tool_group = groups.FindObject("NONE")
                    self.print_log(f"刀具 {tool_name} 不存在，使用NONE", "WARN")


                # ============ 新增：根据图层选择几何体 ============
                # 获取图层参数，默认为20
                layer = params.get('layer', 20)
                
                # 根据图层映射获取几何体名称
                workpiece_geometry = self.LAYER_TO_GEOMETRY.get(layer, "WORKPIECE_1")
                
                # 查找几何体组，如果找不到则使用默认的WORKPIECE_1
                try:
                    workpiece_group = groups.FindObject(workpiece_geometry)
                    self.print_log(f"使用图层{layer}对应的几何体: {workpiece_geometry}", "SUCCESS")
                except:
                    workpiece_group = groups.FindObject("WORKPIECE_1")
                    self.print_log(f"图层{layer}对应的几何体{workpiece_geometry}不存在，使用默认的WORKPIECE_1", "WARN")
                # ==================================================


                # ============ 修改：使用开粗程序组 ============
                # 获取开粗程序组，如果不存在则创建
                # 获取图层参数，默认为20
                layer = params.get('layer', 20)
                # 根据图层获取对应的程序组
                program_group = self.get_rough_program_group(layer)
                if not program_group:
                    self.print_log("无法获取程序组，使用默认PROGRAM", "ERROR")
                    program_group = groups.FindObject("PROGRAM")
                # =============================================


                # ============ 新增：生成自定义工序名称 ============
                # 规则：工序类型_刀具名称_图层_序号
                base_name = f"{operation_key}_{tool_name}"
                
                # 尝试生成唯一名称
                operation_name = base_name
                suffix = 1
                
                # 检查名称是否已存在
                while True:
                    try:
                        # 尝试查找是否已存在同名操作
                        existing_op = self.work_part.CAMSetup.CAMOperationCollection.FindObject(operation_name)
                        # 如果存在，添加后缀
                        operation_name = f"{base_name}_{suffix}"
                        suffix += 1
                    except:
                        # 名称不存在，跳出循环
                        break
                
                custom_operation_name = operation_name
                self.print_log(f"自定义工序名称: {custom_operation_name}", "DEBUG")
                # ===================================================
                op = self.work_part.CAMSetup.CAMOperationCollection.Create(
                    program_group,
                    groups.FindObject("METHOD"),
                    tool_group,
                    workpiece_group,
                    config['operation_name'],
                    config['operation_type'],
                    NXOpen.CAM.OperationCollection.UseDefaultName.FalseValue, # 关键：使用自定义名称
                    custom_operation_name  # 自定义工序名称
                )

                # 3. 创建Builder
                builder_method_name = self.BUILDER_MAP[config['builder_type']]
                builder_method = getattr(self.work_part.CAMSetup.CAMOperationCollection, builder_method_name)
                builder = builder_method(op)

                try:
                    # ============ 4. 根据工序类型设置几何集 ============
                    if operation_key == "行腔_SIMPLE":
                        # 行腔工序：设置两个几何集，分别设置不同余量
                        if normal_faces or yellow_faces:
                            self._set_geometry_with_two_sets(builder, normal_faces, yellow_faces, final_stock)
                            self.print_log(f"行腔工序：普通面余量=0mm，黄色面余量={final_stock}mm", "SUCCESS")
                    else:
                        # 往复等高工序：设置一个几何集，使用统一的余量（在参数配置中设置）
                        if normal_faces:
                            self._set_geometry_with_one_set(builder, normal_faces)
                            self.print_log(f"往复等高工序：设置 {len(normal_faces)} 个面到几何集", "SUCCESS")
                    # ==============================================

                    # 5. 配置特定参数
                    if config['builder_type'] == 'zlevel':
                        self._configure_zlevel_params(builder, config)
                    elif config['builder_type'] == 'cavity':
                        use_arc_engagement = config['special_config'].get('use_arc_engagement', False)
                        self._configure_cavity_params(builder, config, use_arc_engagement)

                    # 6. 配置安全距离
                    self._configure_auto_clearance(builder, params.get('safe_distance', 50.0))

                    # 7. 提交
                    committed_op = builder.Commit()

                finally:
                    builder.Destroy()

                # 8. 完成操作
                final_op = self._finalize_operation(committed_op, tool_name)

            self.success_count += 1
            self.print_log(f"{operation_key} 创建成功: {final_op.Name}", "SUCCESS")

            result = {
                "status": "Success",
                "name": final_op.Name,
                "type": operation_key,
                "tag": final_op.Tag,
                "normal_faces_count": len(normal_faces),
                "yellow_faces_count": len(yellow_faces),
                "message": f"{config['description']}创建完成，普通面{len(normal_faces)}个，黄色面{len(yellow_faces)}个"
            }
            self.test_results.append(result)
            return result

        except Exception as e:
            self.failed_count += 1
            self.print_log(f"{operation_key} 创建失败: {e}", "ERROR")
            traceback.print_exc()

            result = {
                "status": "Failed",
                "error": str(e),
                "type": operation_key,
                "message": "工序创建失败"
            }
            self.test_results.append(result)
            return result

    def print_summary(self):
        """打印执行摘要"""
        self.print_separator("=")
        success_rate = (self.success_count / self.operation_count * 100) if self.operation_count > 0 else 0

        print(f"""
  刀轨生成摘要
  ----------------------------------------
  总工序数:   {self.operation_count}
  成功:       {self.success_count} ✅
  失败:       {self.failed_count} ❌
  成功率:     {success_rate:.1f}%
  程序组:     开粗 (所有工序)
        """.strip(), flush=True)

        if self.test_results:
            self.print_separator("-")
            print("  详细结果:")
            for i, result in enumerate(self.test_results, 1):
                status_emoji = "✅" if result['status'] == "Success" else "❌"
                name = result.get('name', result['type'])
                layer = result.get('layer', '未知')
                workpiece = result.get('workpiece', '未知')
                program_group = result.get('program_group', '开粗')
                print(f"  {i}. {name} ({result['type']}) {status_emoji}")
                print(f"      程序组: {program_group}, 图层: {layer}, 几何体: {workpiece}")
                if result.get('message'):
                    print(f"     信息: {result['message']}")
                if result.get('error'):
                    print(f"     错误: {result['error']}")

        self.print_separator("=")


# ==================================================================================
# 主流程
# ==================================================================================
def generate_toolpath_workflow(part_path, cavity_json_path=None, reciprocating_json_path=None,save_dir=None):
    """刀轨生成主工作流"""
    session = NXOpen.Session.GetSession()
    base_part, load_status = session.Parts.OpenBaseDisplay(part_path)
    work_part = session.Parts.Work

    generator = ToolpathGenerator(work_part,save_dir=save_dir)
    generator.print_header("NX CAM 刀轨生成工具 - 开粗版")
    generator.print_log(f"零件: {work_part.Name}", "INFO")
    generator.print_log(f"测试模式: {'开启' if CONFIG['TEST_MODE'] else '关闭'}", "INFO")
    
    # 切换到加工环境
    generator.switch_to_manufacturing()


    # ==================== 创建开粗程序组 ====================
    rough_program_group = generator.create_rough_program_group()
    if rough_program_group:
        generator.print_log(f"开粗程序组已准备就绪 (Tag: {rough_program_group})", "SUCCESS")
    
    # ==================== 加载JSON测试用例 ====================
    cavity_test_cases = generator.load_cavity_assignments_from_json(cavity_json_path)
    reciprocating_test_cases = generator.load_reciprocating_zlevel_assignments_from_json(reciprocating_json_path)

    # 合并所有测试用例
    all_test_cases = []
    all_test_cases.extend(cavity_test_cases)
    all_test_cases.extend(reciprocating_test_cases)
    all_test_cases.extend(TEST_CASES)
    
    generator.print_log(f"总测试用例数: {len(all_test_cases)}", "INFO")
    generator.print_log(
        f"其中行腔: {len(cavity_test_cases)} 个, "
        f"往复等高: {len(reciprocating_test_cases)} 个", 
        "INFO"
    )

    
    # 执行所有测试用例
    for test_case in all_test_cases:
        try:
            if len(test_case) == 4:
                op_key, face_inputs, tool_name, extra_params = test_case
                # 检查是否是行腔工序，需要特殊处理面输入
                if op_key == "行腔_SIMPLE" and isinstance(face_inputs, dict):
                    # 行腔工序使用新的双几何集逻辑
                    generator.create_operation(op_key, face_inputs, tool_name, **extra_params)
                else:
                    # 其他工序保持原逻辑（向后兼容）
                    generator.create_operation(op_key, face_inputs, tool_name, **extra_params)
            else:
                op_key, face_ids, tool_name = test_case
                extra_params = {}
            
                generator.create_operation(op_key, face_ids, tool_name, **extra_params)
        except Exception as e:
            generator.print_log(f"测试异常: {e}", "ERROR")
    
    # 打印摘要
    generator.print_summary()
    
    # 保存零件
    saved_path = generator.save_part(part_path)
    
    # 清理资源
    if load_status:
        load_status.Dispose()
    
    generator.print_log("所有工序创建完成", "END")
    return saved_path


def main():
    """主函数"""
    try:
        saved_path = generate_toolpath_workflow(
            part_path=CONFIG["PART_PATH"],
            cavity_json_path=CONFIG["JSON_CAVITY_PATH"],  # 对应的是 xx_行腔.json
            reciprocating_json_path=CONFIG["JSON_RECIPROCATING_PATH"],
            # 新增：保存目录参数
            save_dir=r'C:\Users\admin\Desktop\新建文件夹\Daogui_prt'
        )
        
        print(f"✅ 刀轨生成完成，文件已保存至: {saved_path}")
        return 0
        
    except Exception as e:
        print(f"❌ 主程序异常: {e}", flush=True)
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit_code = main()
    exit(exit_code)