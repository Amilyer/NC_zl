#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
NX CAM自动化工具 - 刀轨生成模块
精简版：只保留螺旋铣、半螺旋、半爬面、爬面往复等高、爬面和面铣工序
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
    "PART_PATH": r'C:\Projects\NC\output\06_CAM\Roughing_PRT\UP-01.prt',
    "AUTO_SAVE": True,
    "TEST_MODE": True,  # 设置为False时生成实际刀轨
    
    # JSON文件路径

    # 半精
    "JSON_HALF_SPIRAL_PATH": r'C:\Users\Admin\Desktop\12.14修改版(1)\Toolpath_JSON\DIE_03_半精_螺旋.json',
    "JSON_HALF_SPIRAL_RECIPROCATING_PATH": r'C:\Projects\NC\output\json\GU-01_半精_螺旋_往复等高.json',
    "JSON_HALF_SURFACE_PATH": r'C:\Projects\NC\output\json\DIE-05_半精_爬面.json',
    "JSON_HALF_JIAO_PATH": r'C:\Projects\NC\output\06_CAM\Toolpath_JSON\UP-01_半精_清角.json',
    "JSON_HALF_MIAN_PATH": r'C:\Projects\NC\output\json\DIE-05_半精_面铣.json',


    # 全精
    "JSON_MIAN_PATH": r'C:\Projects\NC\output\json\DIE-05_全精_面铣.json',
    "JSON_SPIRAL_PATH": r'C:\Users\Admin\Desktop\12.14修改版(1)\Toolpath_JSON\DIE-03_全精_螺旋.json',
    "JSON_SPIRAL_RECIPROCATING_PATH": r'C:\Projects\NC\output\json\GU-01_全精_螺旋_往复等高.json',
    "JSON_RECIPROCATING_PATH": r'C:\Projects\NC\output\json\DIE-05_全精_往复等高.json',
    "JSON_SURFACE_PATH": r'C:\Projects\NC\output\json\DIE-05_全精_爬面.json',
    "JSON_GEN_PATH": r'C:\Projects\NC\output\json\GU-01_全精_清根.json',



}


# ==================================================================================
# 操作模板配置
# ==================================================================================
OPERATION_CONFIGS = {
    "MIAN1_SIMPLE": {
        "operation_name": "45#备料",
        "operation_type": "MIAN1",
        "operation_subtype": "MIAN1",
        "builder_type": "volume_25d",
        "description": "平面铣精加工",
        "special_config": {
            "cut_pattern": NXOpen.CAM.CutPatternBuilder.Types.FollowPeriphery,
            "stepover_type": NXOpen.CAM.StepoverBuilder.StepoverTypes.PercentToolFlat,
            "stepover_distance": 13.7,  # 默认步距值
            # ============ 新增默认值 ============
            "spindle_rpm": 1700.0,
            "feed_per_tooth": 2000.0,
            "feed_rapid": 8000.0,
            # ===================================
            "cut_direction": NXOpen.CAM.CutDirection.Types.Climb,
            "pattern_direction": NXOpen.CAM.CutParametersPatternDirectionTypes.Inward,
            "floor_stock": 1e-17,
            "wall_stock": 1e-17,
            "engage_closed_type": NXOpen.CAM.NcmPlanarEngRetBuilder.EngRetTypes.RampOnShape,
            "engage_open_type": NXOpen.CAM.NcmPlanarEngRetBuilder.EngRetTypes.Linear
        }
    },
    "D4-螺旋_SIMPLE": {
        "operation_name": "45#备料",
        "operation_type": "D4-螺旋", 
        "operation_subtype": "D4-螺旋",
        "builder_type": "zlevel",
        "description": "螺旋铣削精加工",
        "special_config": {
            "cut_level_range": NXOpen.CAM.CutLevel.RangeTypes.Automatic,
            "global_depth_per_cut": 10.1,
            # ============ 新增默认值 ============
            "spindle_rpm": 1700.0,
            "feed_per_tooth": 2000.0,
            "feed_rapid": 8000.0,
            # ===================================
            # ============ 新增余量参数 ============
            "part_stock": 0.2,      # 部件侧面余量
            "floor_stock": 0.8,     # 部件底面余量
            # ===================================
            "stepover_type": NXOpen.CAM.StepoverBuilder.StepoverTypes.Constant
        }
    },
    "爬面_SIMPLE": {
        "operation_name": "45#备料",
        "operation_type": "爬面",
        "operation_subtype": "爬面",
        "builder_type": "surface",
        "description": "曲面轮廓精加工",
        "special_config": {
            "cut_direction": NXOpen.CAM.SurfaceContourBuilder.CutDirectionTypes.Climb,
            "cut_angle": 45.0,
            "part_stock": 0.0,
            "engage_type": NXOpen.CAM.NcmScEngRetBuilder.EngRetTypes.PlungeLift,
            # ============ 新增默认值 ============
            "spindle_rpm": 1700.0,
            "feed_per_tooth": 2000.0,
            "feed_rapid": 8000.0,
            # ===================================
            # ============ 新增：步距参数 ============
            "stepover_distance": 0.3  # 默认步距值
        }
    },
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
            "global_depth_per_cut": 0.1,  # 默认每刀深度0.1mm
            # 新增：切削层范围类型 - 使用 UserDefined 以便手动设置范围深度生效
            "cut_level_range": NXOpen.CAM.CutLevel.RangeTypes.UserDefined
        }
    },
    "清根_SIMPLE": {
        "operation_name": "45#备料",
        "operation_type": "清根",
        "operation_subtype": "清根",
        "builder_type": "surface",  # 新增构建器类型
        "description": "清根加工",
        "special_config": {
            # 从journal中提取的关键参数
            "flow_overlap_distance": 0.5,  # 重叠距离
            # ============ 新增默认值 ============
            "spindle_rpm": 1700.0,
            "feed_per_tooth": 2000.0,
            "feed_rapid": 8000.0,
            # ===================================
        }
    },
    "清角_SIMPLE": {
    "operation_name": "45#备料",
    "operation_type": "D4-清角",
    "operation_subtype": "D4-清角",
    "builder_type": "zlevel",  # 使用zlevel构建器
    "description": "清角加工",
    "special_config": {
        "reference_tool": None,
        "cut_direction": NXOpen.CAM.CutDirection.Types.Mixed,
        "cut_order": NXOpen.CAM.CutParametersCutOrderTypes.DepthFirstAlways,
        "part_stock": 0.0,
        # ============ 新增默认值 ============
        "spindle_rpm": 1700.0,
        "feed_per_tooth": 2000.0,
        "feed_rapid": 8000.0,
        # ===================================
        "merge_distance": 3.0,
        # 新增：步距类型和每刀深度参数
        "stepover_type": NXOpen.CAM.StepoverBuilder.StepoverTypes.Constant,
        "global_depth_per_cut": 0.1  # 默认每刀深度0.1mm
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
        'volume_25d': 'CreateVolumeBased25dMillingOperationBuilder',
        'zlevel': 'CreateZlevelMillingBuilder',
        'surface': 'CreateSurfaceContourBuilder',
        'flowcut': 'CreateSurfaceContourBuilder',  # 清根使用相同的构建器
    }

    LAYER_TO_GEOMETRY = {
        20: "WORKPIECE_0",
        30: "WORKPIECE_1", 
        40: "WORKPIECE_2",
        50: "WORKPIECE_3",
        60: "WORKPIECE_4",
        70: "WORKPIECE_5"
    }

    # 新增：图层到加工方向的映射
    LAYER_TO_DIRECTION = {
        20: "正",
        30: "左",
        40: "右",
        50: "前",
        60: "后",
        70: "反"
    }

    # 新增：所有加工方向列表
    DIRECTIONS = ["正", "左", "右", "前", "后", "反"]

    def __init__(self, work_part, save_dir=None):
        self.work_part = work_part
        self.session = NXOpen.Session.GetSession()
        self.uf = NXOpen.UF.UFSession.GetUFSession()
        self.operation_count = 0
        self.success_count = 0
        self.failed_count = 0
        self.test_results = []
        self.save_dir = save_dir  # 保存目录参数

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
    

    #-----------------创建程序组------------------
    def create_program_groups(self):
        """创建半精和全精程序组及其子程序组"""
        self.print_log("开始创建程序组...", "START")
        template_name = "45#备料"
        try:
            with self.undo_mark_context("创建程序组"):
                # 获取CAM设置
                cam_setup = self.work_part.CAMSetup
                cam_groups = cam_setup.CAMGroupCollection
                
                # 查找NC_PROGRAM根组（如果没有则创建）
                try:
                    nc_program_group = cam_groups.FindObject("NC_PROGRAM")
                except:
                    self.print_log("未找到NC_PROGRAM组，使用默认PROGRAM组", "WARN")
                    nc_program_group = cam_groups.FindObject("PROGRAM")
                
                # 创建主程序组（半精和全精）
                program_groups = {}
                for stage in ["半精", "全精"]:
                    try:
                        program_group = cam_groups.FindObject(stage)
                        self.print_log(f"程序组 '{stage}' 已存在", "DEBUG")
                    except:
                        program_group = cam_groups.CreateProgram(
                            nc_program_group, 
                            template_name, 
                            "PROGRAM", 
                            NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue, 
                            stage
                        )
                        self.print_log(f"创建程序组: {stage}", "SUCCESS")
                    
                    # 在程序组下创建子程序组（正、反、左、右、前、后）
                    direction_groups = {}
                    for direction in self.DIRECTIONS:
                        # 修改这里：根据方向获取对应的图层编号
                        # 查找方向对应的图层编号
                        layer_for_direction = None
                        for layer_num, dir_name in self.LAYER_TO_DIRECTION.items():
                            if dir_name == direction:
                                layer_for_direction = layer_num
                                break
                        
                        if layer_for_direction is None:
                            # 如果没有找到对应的图层，使用默认图层20
                            layer_for_direction = 20
                            self.print_log(f"方向 '{direction}' 没有对应的图层映射，使用默认图层20", "WARN")
                        
                        # 构建子程序组名称：方向_阶段_图层
                        sub_group_name = f"{direction}_{stage}_{layer_for_direction}"
                        
                        try:
                            sub_group = cam_groups.FindObject(sub_group_name)
                            self.print_log(f"子程序组 '{sub_group_name}' 已存在", "DEBUG")
                        except:
                            sub_group = cam_groups.CreateProgram(
                                program_group,  # 父组是主程序组
                                template_name, 
                                "PROGRAM", 
                                NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue, 
                                sub_group_name
                            )
                            self.print_log(f"创建子程序组: {sub_group_name}", "SUCCESS")
                        
                        direction_groups[direction] = sub_group
                    
                    program_groups[stage] = {
                        "main": program_group,
                        "directions": direction_groups
                    }
                
                self.print_log("程序组结构创建完成", "SUCCESS")
                return program_groups
                
        except Exception as e:
            self.print_log(f"创建程序组失败: {e}", "ERROR")
            traceback.print_exc()
            return {}
    
    def get_program_group_by_stage_and_layer(self, stage, layer):
        """根据加工阶段和图层获取对应的子程序组"""
        try:
            # 根据图层获取加工方向
            direction = self.LAYER_TO_DIRECTION.get(layer)
            if not direction:
                self.print_log(f"未知图层: {layer}，使用默认方向'正'", "WARN")
                direction = "正"
                layer = 20  # 使用默认图层20
            
            # 构建子程序组名称
            sub_group_name = f"{direction}_{stage}_{layer}"
            cam_groups = self.work_part.CAMSetup.CAMGroupCollection
            
            try:
                return cam_groups.FindObject(sub_group_name)
            except:
                # 如果找不到子程序组，尝试获取主程序组
                self.print_log(f"未找到子程序组 '{sub_group_name}'，使用主程序组", "WARN")
                return cam_groups.FindObject(stage)
                
        except Exception as e:
            self.print_log(f"获取程序组失败: {e}，使用默认PROGRAM", "WARN")
            try:
                return cam_groups.FindObject("PROGRAM")
            except:
                return None

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
                    self.work_part.Save(NXOpen.BasePart.SaveComponents.TrueValue,
                                        NXOpen.BasePart.CloseModified.UseResponses)
                    self.print_log(f"部件已保存 (覆盖原文件): {save_path}", "SUCCESS")
                else:
                    # 如果路径不同，另存为
                    self.work_part.SaveAs(save_path)
                    self.print_log(f"刀轨生成完成，另存至: {save_path}", "SUCCESS")
            except Exception as e:
                self.print_log(f"保存失败: {e}", "ERROR")
                # 尝试强制关闭以释放句柄 (可选)
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

    # ==================== JSON测试用例加载 ====================
    """从 xx_螺旋.JSON 文件加载刀具分配结果并生成测试用例"""
    def load_spiral_from_json(self, json_path, stage="半精"):
        
        self.print_log(f"读取刀具分配JSON文件: {json_path}", "START")
        
        try:
            # 读取JSON文件
            with open(json_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
            
            self.print_log(f"成功读取JSON，共 {len(data)} 个螺旋组", "SUCCESS")
            
            test_cases = []
            
            for group_name, group_data in data.items():
                try:
                    # 提取关键数据
                    operation_type = group_data['工序']
                    face_ids = group_data['面ID列表']
                    tool_name = group_data['刀具名称']
                    max_depth = group_data['切深']
                    layer = group_data.get('指定图层', 20)
                    reference_tool = group_data.get('参考刀具', None)

                    # ============ 新增：读取转速、进给、横越参数 ============
                    spindle_rpm = group_data.get('转速', 1700.0)  # 默认值1700
                    feed_per_tooth = group_data.get('进给', 2000.0)  # 默认值2000
                    feed_rapid = group_data.get('横越', 8000.0)  # 默认值8000
                    # ====================================================
                    
                    # ============ 新增：余量参数 ============
                    part_stock = group_data.get('部件侧面余量',0.2)  # 默认侧面余量
                    floor_stock = group_data.get('部件底面余量',0.5)  # 默认底面余量
                    

                    
                    # 将面ID列表中的整数转换为字符串
                    face_ids_str = [str(face_id) for face_id in face_ids]
                    
                    # 创建测试用例
                    test_case = (
                        operation_type, 
                        face_ids_str, 
                        tool_name, 
                        {"max_depth": max_depth, "layer": layer, "reference_tool": reference_tool,
                            # ============ 新增：传递余量参数 ============
                            "part_stock": part_stock,
                            "floor_stock": floor_stock,
                            # ===========================================
                            # ============ 传递转速、进给、横越参数 ============
                            "spindle_rpm": spindle_rpm,
                            "feed_per_tooth": feed_per_tooth,
                            "feed_rapid": feed_rapid,
                            # ================================================
                            "stage": stage  # 添加阶段信息
                            
                        }
                    )
                    
                    test_cases.append(test_case)
                    
                    # 修正：将多个字符串合并为一个字符串
                    log_message = (
                        f"螺旋组 '{group_name}': 刀具={tool_name}, "
                        f"面数量={len(face_ids_str)}, 深度={max_depth}mm, 图层={layer}, "
                        f"侧面余量={part_stock}mm, 底面余量={floor_stock}mm, "
                        f"转速={spindle_rpm}RPM, 进给={feed_per_tooth}mm/齿, 横越={feed_rapid}mm/min"
                    )
                    self.print_log(log_message, "DEBUG")
                    
                except Exception as e:
                    self.print_log(f"解析螺旋组 '{group_name}' 数据失败: {e}", "ERROR")
                    continue
            
            self.print_log(f"成功生成 {len(test_cases)} 个测试用例", "SUCCESS")
            return test_cases
            
        except Exception as e:
            self.print_log(f"读取JSON文件失败: {e}", "ERROR")
            return []
        
    """从【xx_螺旋_往复等高.json】加载往复等高工序分配结果并生成测试用例"""
    def load_spiral_reciprocating_from_json(self, json_path,stage="半精"):
        
        self.print_log(f"读取螺旋（往复等高）分配JSON文件: {json_path}", "START")
        
        try:
            with open(json_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
            
            self.print_log(f"成功读取螺旋（往复等高）JSON，共 {len(data)} 个半螺旋组", "SUCCESS")
            
            test_cases = []
            
            for group_name, group_data in data.items():
                try:
                    operation_type = group_data['工序']        # 应该是 "往复等高_SIMPLE"
                    face_ids = group_data['面ID列表']
                    tool_name = group_data['刀具名称']
                    cutting_depth = float(group_data['切深'])  # 每刀切深
                    layer = group_data.get('指定图层', 20)
                    reference_tool = group_data.get('参考刀具', None)

                    # ============ 新增：读取转速、进给、横越参数 ============
                    spindle_rpm = group_data.get('转速', 1700.0)  # 默认值1700
                    feed_per_tooth = group_data.get('进给', 2000.0)  # 默认值2000
                    feed_rapid = group_data.get('横越', 8000.0)  # 默认值8000
                    # ====================================================

                    # ============ 新增：余量参数 ============
                    part_stock = group_data.get('部件侧面余量',0.2)  # 默认侧面余量
                    floor_stock = group_data.get('部件底面余量',0.5)  # 默认底面余量
                    
                    # 转成字符串列表（兼容你的_face查找逻辑）
                    face_ids_str = [str(fid) for fid in face_ids]
                    
                    # 关键：往复等高使用的是 ZLEVEL 工序，切深参数要传 global_depth_per_cut
                    test_case = (
                        "往复等高_SIMPLE",           # 固定使用这个key，对应OPERATION_CONFIGS里的配置
                        face_ids_str,
                        tool_name,
                        {
                            "global_depth_per_cut": cutting_depth,   # ← 重点！这里传每刀切深
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
                            "stage": stage  # 添加阶段信息
                        }
                    )
                    
                    test_cases.append(test_case)
                    
                    self.print_log(
                        f"半螺旋组 '{group_name}': 刀具={tool_name}, "
                        f"面数量={len(face_ids_str)}, 每刀切深={cutting_depth}mm, 图层={layer}, "
                        f"转速={spindle_rpm}RPM, 进给={feed_per_tooth}mm/齿, "
                        f"横越={feed_rapid}mm/min, 图层={layer}", 
                        "DEBUG"
                    )
                    
                except Exception as e:
                    self.print_log(f"解析半螺旋组 '{group_name}' 失败: {e}", "ERROR")
                    continue
            
            self.print_log(f"成功生成 {len(test_cases)} 个往复等高测试用例", "SUCCESS")
            return test_cases
            
        except Exception as e:
            self.print_log(f"读取半螺旋JSON失败: {e}", "ERROR")
            return []


    """从【半精_爬面.json、全精_往复等高.json】加载往复等高工序分配结果并生成测试用例"""
    def load_half_surface_from_json(self, json_path,stage="半精"):

        self.print_log(f"读取JSON文件: {json_path}", "START")
        
        try:
            with open(json_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
            
            self.print_log(f"成功读取JSON，共 {len(data)} 个组", "SUCCESS")
            
            test_cases = []
            
            for group_name, group_data in data.items():
                try:
                    
                    operation_type = group_data['工序']
                    face_ids = group_data['面ID列表']
                    tool_name = group_data['刀具名称']
                    cutting_depth = float(group_data['切深'])  # 每刀切深
                    layer = group_data.get('指定图层', 20)
                    reference_tool = group_data.get('参考刀具', None)


                    # ============ 新增：读取转速、进给、横越参数 ============
                    spindle_rpm = group_data.get('转速', 1700.0)  # 默认值1700
                    feed_per_tooth = group_data.get('进给', 2000.0)  # 默认值2000
                    feed_rapid = group_data.get('横越', 8000.0)  # 默认值8000
                    # ====================================================

                    # ============ 新增：读取侧面/底面余量参数 ============
                    part_stock = group_data.get('部件侧面余量',0.2)  # 默认侧面余量
                    floor_stock = group_data.get('部件底面余量',0.5)  # 默认底面余量
                    # ==================================================
                    
                    # 转成字符串列表（兼容_face查找逻辑）
                    face_ids_str = [str(fid) for fid in face_ids]
                    
                    # 关键：半爬面使用往复等高工序配置，固定使用"往复等高_SIMPLE"作为操作键
                    test_case = (
                        "往复等高_SIMPLE",           # 固定使用这个key，对应OPERATION_CONFIGS里的配置
                        face_ids_str,
                        tool_name,
                        {
                            "global_depth_per_cut": cutting_depth,   # 传每刀切深
                            "layer": layer,
                            "reference_tool": reference_tool,
                            # ============ 新增：传递转速、进给、横越参数 ============
                            "spindle_rpm": spindle_rpm,
                            "feed_per_tooth": feed_per_tooth,
                            "feed_rapid": feed_rapid,
                            # ====================================================
                            # ============ 新增：传递侧面/底面余量参数 ============
                            "part_stock": part_stock,
                            "floor_stock": floor_stock,
                            # ==================================================
                            "stage": stage  # 添加阶段信息
                        }
                    )
                    
                    test_cases.append(test_case)
                    
                    self.print_log(
                        f"半爬面组 '{group_name}': 刀具={tool_name}, "
                        f"面数量={len(face_ids_str)}, 每刀切深={cutting_depth}mm, 图层={layer}, "
                        f"转速={spindle_rpm}RPM, 进给={feed_per_tooth}mm/齿, "
                        f"横越={feed_rapid}mm/min", 
                        "DEBUG"
                    )
                    
                except Exception as e:
                    self.print_log(f"解析半爬面组 '{group_name}' 失败: {e}", "ERROR")
                    continue
            
            self.print_log(f"成功生成 {len(test_cases)} 个半爬面测试用例", "SUCCESS")
            return test_cases
            
        except Exception as e:
            self.print_log(f"读取半爬面JSON失败: {e}", "ERROR")
            return []


    """从全精_爬面.JSON文件加载爬面刀具分配结果并生成测试用例"""
    def load_surface_from_json(self, json_path,stage="全精"):
        """

        
        参数:
            json_path: JSON文件路径
        返回:
            test_cases: 生成的测试用例列表
        """
        self.print_log(f"读取爬面分配JSON文件: {json_path}", "START")
        
        try:
            # 读取JSON文件
            with open(json_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
            
            self.print_log(f"成功读取JSON，共 {len(data)} 个爬面组", "SUCCESS")
            
            test_cases = []
            
            for group_name, group_data in data.items():
                try:
                    # 提取关键数据
                    # 注意：这里将"侧壁爬面_SIMPLE"转换为"爬面_SIMPLE"
                    operation_type = group_data['工序']
                    face_ids = group_data['面ID列表']
                    tool_name = group_data['刀具名称']
                    stepover_distance = group_data['切深']  # 使用切深作为步距
                    cut_angle = group_data.get('切削角度', 45.0)  # 获取切削角度，默认为45度
                    layer = group_data.get('指定图层', 20)  # 默认图层20

                    # ============ 新增：读取转速、进给、横越参数 ============
                    spindle_rpm = group_data.get('转速', 1700.0)  # 默认值1700
                    feed_per_tooth = group_data.get('进给', 2000.0)  # 默认值2000
                    feed_rapid = group_data.get('横越', 8000.0)  # 默认值8000
                    # ====================================================
                    
                    # 将面ID列表中的整数转换为字符串
                    face_ids_str = [str(face_id) for face_id in face_ids]
                    
                    # 创建测试用例
                    # 格式: (operation_type, face_ids, tool_name, extra_params)
                    test_case = (
                        "爬面_SIMPLE",  # 使用固定的操作类型
                        face_ids_str, 
                        tool_name, 
                        {
                            "stepover_distance": stepover_distance,
                            "cut_angle": cut_angle,
                            "layer": layer,
                            # ============ 新增：传递转速、进给、横越参数 ============
                            "spindle_rpm": spindle_rpm,
                            "feed_per_tooth": feed_per_tooth,
                            "feed_rapid": feed_rapid,
                            # ====================================================
                            "stage": stage  # 添加阶段信息
                        }
                    )
                    
                    test_cases.append(test_case)
                    
                    self.print_log(
                        f"爬面组 '{group_name}': 刀具={tool_name}, "
                        f"面数量={len(face_ids_str)}, 步距={stepover_distance}mm, "
                        f"转速={spindle_rpm}RPM, 进给={feed_per_tooth}mm/齿, "
                        f"横越={feed_rapid}mm/min, 图层={layer}", 
                        "DEBUG"
                    )
                    
                except Exception as e:
                    self.print_log(f"解析爬面组 '{group_name}' 数据失败: {e}", "ERROR")
                    continue
            
            self.print_log(f"成功生成 {len(test_cases)} 个爬面测试用例", "SUCCESS")
            return test_cases
            
        except Exception as e:
            self.print_log(f"读取爬面JSON文件失败: {e}", "ERROR")
            return []
        

    def load_jiao_from_json(self, json_path, stage="半精"):

        self.print_log(f"读取JSON文件: {json_path}", "START")

        try:
            with open(json_path, 'r', encoding='utf-8') as file:
                data = json.load(file)

            self.print_log(f"成功读取JSON，共 {len(data)} 个组", "SUCCESS")

            test_cases = []

            for group_name, group_data in data.items():
                try:

                    operation_type = group_data['工序']
                    face_ids = group_data['面ID列表']
                    tool_name = group_data['刀具名称']
                    cutting_depth = float(group_data['切深'])  # 每刀切深
                    layer = group_data.get('指定图层', 20)
                    reference_tool = group_data.get('参考刀具', None)

                    # ============ 新增：读取转速、进给、横越参数 ============
                    spindle_rpm = group_data.get('转速', 1700.0)  # 默认值1700
                    feed_per_tooth = group_data.get('进给', 2000.0)  # 默认值2000
                    feed_rapid = group_data.get('横越', 8000.0)  # 默认值8000
                    # ====================================================

                    # ============ 新增：读取侧面/底面余量参数 ============
                    part_stock = group_data.get('部件侧面余量', 0.03)  # 默认侧面余量
                    floor_stock = group_data.get('部件底面余量', 0.03)  # 默认底面余量
                    # ==================================================

                    # 转成字符串列表（兼容_face查找逻辑）
                    face_ids_str = [str(fid) for fid in face_ids]

                    # 关键：清角使用D4-清角配置，固定使用"清角_SIMPLE"作为操作键
                    test_case = (
                        "清角_SIMPLE",  # 固定使用这个key，对应OPERATION_CONFIGS里的配置
                        face_ids_str,
                        tool_name,
                        {
                            "global_depth_per_cut": cutting_depth,  # 传每刀切深
                            "layer": layer,
                            "reference_tool": reference_tool,
                            # ============ 新增：传递转速、进给、横越参数 ============
                            "spindle_rpm": spindle_rpm,
                            "feed_per_tooth": feed_per_tooth,
                            "feed_rapid": feed_rapid,
                            # ====================================================
                            # ============ 新增：传递侧面/底面余量参数 ============
                            "part_stock": part_stock,
                            "floor_stock": floor_stock,
                            # ==================================================
                            "stage": stage  # 添加阶段信息
                        }
                    )

                    test_cases.append(test_case)

                    self.print_log(
                        f"清角组 '{group_name}': 刀具={tool_name}, 参考道具={reference_tool if reference_tool else '无'}"
                        f"面数量={len(face_ids_str)}, 每刀切深={cutting_depth}mm, 图层={layer}, "
                        f""
                        f"转速={spindle_rpm}RPM, 进给={feed_per_tooth}mm/齿, "
                        f"横越={feed_rapid}mm/min",
                        "DEBUG"
                    )

                except Exception as e:
                    self.print_log(f"解析半爬面组 '{group_name}' 失败: {e}", "ERROR")
                    continue

            self.print_log(f"成功生成 {len(test_cases)} 个半爬面测试用例", "SUCCESS")
            return test_cases

        except Exception as e:
            self.print_log(f"读取半爬面JSON失败: {e}", "ERROR")
            return []
    
        
    def load_gen_from_json(self, json_path, stage="全精"):
        """从xx_全精_清根.json  文件加载清根刀具分配结果并生成测试用例"""
        self.print_log(f"读取清根分配JSON文件: {json_path}", "START")
        
        try:
            with open(json_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
            
            self.print_log(f"成功读取JSON，共 {len(data)} 个清根组", "SUCCESS")
            
            test_cases = []
            
            for group_name, group_data in data.items():
                try:
                    operation_type = group_data['工序']
                    face_ids = group_data['面ID列表']
                    tool_name = group_data['刀具名称']
                    flow_overlap_distance = group_data.get('重叠距离', 0.5)  # 默认0.5mm
                    reference_tool = group_data.get('参考刀具', None)  # 参考刀具
                    layer = group_data.get('指定图层', 20)

                    # ============ 新增：读取转速、进给、横越参数 ============
                    spindle_rpm = group_data.get('转速', 1700.0)
                    feed_per_tooth = group_data.get('进给', 2000.0)
                    feed_rapid = group_data.get('横越', 8000.0)
                    # ====================================================

                    # 将面ID列表转换为字符串
                    face_ids_str = [str(face_id) for face_id in face_ids]
                    
                    # 创建测试用例
                    test_case = (
                        "清根_SIMPLE",           # 固定使用这个key，对应OPERATION_CONFIGS里的配置
                        face_ids_str,
                        tool_name,
                        {
                            "flow_overlap_distance": 0.5,
                            "reference_tool": reference_tool,
                            "layer": layer,
                            "spindle_rpm": spindle_rpm,
                            "feed_per_tooth": feed_per_tooth,
                            "feed_rapid": feed_rapid,
                            "stage": stage
                        }
                    )
                    
                    test_cases.append(test_case)
                    
                    self.print_log(
                        f"清根组 '{group_name}': 刀具={tool_name}, "
                        f"面数量={len(face_ids_str)}, 重叠距离={flow_overlap_distance}mm, "
                        f"参考刀具={reference_tool if reference_tool else '无'}, "
                        f"转速={spindle_rpm}RPM, 进给={feed_per_tooth}mm/齿, "
                        f"横越={feed_rapid}mm/min, 图层={layer}", 
                        "DEBUG"
                    )
                    
                except Exception as e:
                    self.print_log(f"解析清根组 '{group_name}' 数据失败: {e}", "ERROR")
                    continue
            
            self.print_log(f"成功生成 {len(test_cases)} 个清根测试用例", "SUCCESS")
            return test_cases
            
        except Exception as e:
            self.print_log(f"读取清根JSON文件失败: {e}", "ERROR")
            return []


    """从xx_面铣.JSON文件加载面铣刀具分配结果并生成测试用例"""
    def load_mian_from_json(self, json_path,stage="全精"):
        """
        
        参数:
            json_path: JSON文件路径
        返回:
            test_cases: 生成的测试用例列表
        """
        self.print_log(f"读取面铣分配JSON文件: {json_path}", "START")
        
        try:
            # 读取JSON文件
            with open(json_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
            
            self.print_log(f"成功读取JSON，共 {len(data)} 个面铣组", "SUCCESS")
            
            test_cases = []
            
            for group_name, group_data in data.items():
                try:
                    # 提取关键数据
                    operation_type = group_data['工序']
                    face_ids = group_data['面ID列表']
                    tool_name = group_data['刀具名称']
                    stepover_distance = group_data['切深']  # 使用切深作为步距
                    layer = group_data.get('指定图层', 20)  # 默认图层20
                    reference_tool = group_data.get('参考刀具', None)

                    # ============ 新增：读取转速、进给、横越参数 ============
                    spindle_rpm = group_data.get('转速', 1700.0)  # 默认值1700
                    feed_per_tooth = group_data.get('进给', 2000.0)  # 默认值2000
                    feed_rapid = group_data.get('横越', 8000.0)  # 默认值8000
                    # ====================================================

                    # ============ 新增：根据余量参数 ============                    
                     # 从JSON中直接读取所有余量参数
                    floor_stock = group_data.get('最终底面余量', None)  # 最终底面余量
                    wall_stock = group_data.get('壁余量', None)  # 壁余量
                    blank_distance = group_data.get('底面毛坯厚度', None)  # 底面毛坯厚度
                    depth_per_cut = group_data.get('每刀切削深度', None)  # 每刀切削深度
                    part_stock = group_data.get('部件余量', None)  # 部件余量
                    # ======================================================

                    # ============ 新增：读取运动类型参数 ============
                    motion_type_str = group_data.get('运动类型', '切削')  # 默认值为"切削"
                    self.print_log(f"读取运动类型: {motion_type_str}", "DEBUG")
                    # ===============================================
                    

                    # 将面ID列表中的整数转换为字符串
                    face_ids_str = [str(face_id) for face_id in face_ids]
                    
                    # 创建测试用例
                    # 格式: (operation_type, face_ids, tool_name, extra_params)
                    test_case = (
                        operation_type,  # 使用JSON中的工序类型
                        face_ids_str, 
                        tool_name, 
                        {
                            "stepover_distance": stepover_distance,
                            "stepover_type": NXOpen.CAM.StepoverBuilder.StepoverTypes.Constant,  # 设置为恒定步距
                            "layer": layer,
                            "reference_tool": reference_tool,
                            # ============ 新增：传递转速、进给、横越参数 ============
                            "spindle_rpm": spindle_rpm,
                            "feed_per_tooth": feed_per_tooth,
                            "feed_rapid": feed_rapid,
                            # ====================================================

                            # ============ 余量参数（从JSON读取） ============
                            "floor_stock": floor_stock,
                            "wall_stock": wall_stock,
                            "blank_distance": blank_distance,  # 新增：底面毛坯厚度
                            "depth_per_cut": depth_per_cut,
                            "part_stock": part_stock,
                            # ===========================================

                            # ============ 新增：传递运动类型参数 ============
                            "motion_type": motion_type_str,
                            # ================================================

                            "stage": stage  # 添加阶段信息
                        }
                    )
                    
                    test_cases.append(test_case)
                    
                    self.print_log(
                        f"面铣组 '{group_name}': 刀具={tool_name}, "
                        f"面数量={len(face_ids_str)}, 步距={stepover_distance}mm, 图层={layer}, "
                        f"转速={spindle_rpm}RPM, 进给={feed_per_tooth}mm/齿, 横越={feed_rapid}mm/min, "
                        f"参考刀具={reference_tool if reference_tool else '无'}"
                        f"最终底面余量={floor_stock}mm, 壁余量={wall_stock}mm,底面毛坯厚度={blank_distance}mm, "
                        f"每刀切削深度={depth_per_cut}mm, 部件余量={part_stock}mm"
                        f"运动类型={motion_type_str}",  # 新增：日志中显示运动类型
                        "DEBUG"
                    )
                    
                except Exception as e:
                    self.print_log(f"解析面铣组 '{group_name}' 数据失败: {e}", "ERROR")
                    continue
            
            self.print_log(f"成功生成 {len(test_cases)} 个面铣测试用例", "SUCCESS")
            return test_cases
            
        except Exception as e:
            self.print_log(f"读取面铣JSON文件失败: {e}", "ERROR")
            return []

    # ==================== 工序创建核心方法 ====================
    def _set_geometry(self, builder, valid_faces):
        """设置几何体"""
        if not valid_faces:
            return

        try:
            builder.CutAreaGeometry.InitializeData(False)
            item = builder.CutAreaGeometry.GeometryList.FindItem(0)
            rule_opts = self.work_part.ScRuleFactory.CreateRuleOptions()
            rule = self.work_part.ScRuleFactory.CreateRuleFaceDumb(valid_faces, rule_opts)
            rule_opts.Dispose()
            item.ScCollector.ReplaceRules([rule], False)
            self.print_log(f"设置 {len(valid_faces)} 个面作为切削区域", "SUCCESS")
        except Exception as e:
            self.print_log(f"设置几何体失败: {e}", "ERROR")

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
    def _configure_face_milling_params(self, builder, config):
        """配置面铣参数"""
        try:
            special = config.get('special_config', {})
            # 设置步距类型和距离
            stepover_type = special.get('stepover_type')
            builder.BndStepover.StepoverType = stepover_type           

            # ============ 固定设置：最大距离属性 ============
            # 1. 设置意图为刀具直径百分比
            builder.BndStepover.DistanceBuilder.Intent = NXOpen.CAM.ParamValueIntent.ToolDep
            # 2. 固定值为0.5（0.5%刀具直径）
            builder.BndStepover.DistanceBuilder.Value = 50.0
            
            self.print_log("固定设置面铣最大距离：0.5%刀具直径", "SUCCESS")
            # =============================================

            builder.CutPattern.CutPattern = special.get('cut_pattern')         
            builder.CutParameters.CutDirection.Type = special.get('cut_direction')
            builder.CutParameters.PatternDirection = special.get('pattern_direction')

            # ============ 设置余量参数（从JSON读取） ============
            # 1. 设置最终底面余量
            if 'floor_stock' in special and special['floor_stock'] is not None:
                builder.CutParameters.FloorStock.Value = special['floor_stock']
                self.print_log(f"设置最终底面余量: {special['floor_stock']}mm", "SUCCESS")

            # 2. 设置壁余量
            if 'wall_stock' in special and special['wall_stock'] is not None:
                builder.CutParameters.WallStock.Value = special['wall_stock']
                self.print_log(f"设置壁余量: {special['wall_stock']}mm", "SUCCESS")

            # 3. 设置部件余量
            if 'part_stock' in special and special['part_stock'] is not None:
                builder.CutParameters.PartStock.Value = special['part_stock']
                self.print_log(f"设置部件余量: {special['part_stock']}mm", "SUCCESS")

            # 4. 设置底面毛坯厚度（重要！）
            if 'blank_distance' in special and special['blank_distance'] is not None:
                builder.CutParameters.BlankDistance.Value = special['blank_distance']
                self.print_log(f"设置底面毛坯厚度: {special['blank_distance']}mm", "SUCCESS")

            # 5. 设置每刀切削深度
            if 'depth_per_cut' in special and special['depth_per_cut'] is not None:
                try:
                    # 尝试设置每刀切削深度
                    builder.DepthPerCut.Value = special['depth_per_cut']
                    self.print_log(f"设置每刀切削深度: {special['depth_per_cut']}mm", "SUCCESS")
                except AttributeError as e:
                    self.print_log(f"当前builder没有DepthPerCut属性: {e}", "WARN")
                    # 尝试其他可能的属性名
                    try:
                        if hasattr(builder, 'DepthPerCutBuilder'):
                            builder.DepthPerCutBuilder.Value = special['depth_per_cut']
                            self.print_log(f"通过DepthPerCutBuilder设置每刀切削深度: {special['depth_per_cut']}mm",
                                           "SUCCESS")
                    except:
                        self.print_log("无法设置每刀切削深度参数", "WARN")
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


            # ============ 新增：设置跨区域运动类型 ============
            if 'motion_type' in special:
                motion_type_str = special['motion_type']
                try:
                    # 检查是否有AcrossVoids属性
                    if hasattr(builder.CutParameters, 'AcrossVoids'):
                        if motion_type_str == "切削":
                            builder.CutParameters.AcrossVoids.MotionType = NXOpen.CAM.AcrossVoids.MotionTypes.Cut
                            self.print_log("设置跨区域运动类型: 切削", "SUCCESS")
                        elif motion_type_str == "跟随":
                            builder.CutParameters.AcrossVoids.MotionType = NXOpen.CAM.AcrossVoids.MotionTypes.Follow
                            self.print_log("设置跨区域运动类型: 跟随", "SUCCESS")
                        else:
                            self.print_log(f"未知的运动类型: {motion_type_str}，使用默认值", "WARN")
                    else:
                        self.print_log("当前builder没有AcrossVoids属性", "DEBUG")
                except Exception as e:
                    self.print_log(f"设置运动类型失败: {e}", "WARN")
            # ====================================================

        except Exception as e:
            self.print_log(f"面铣参数配置警告: {e}", "DEBUG")

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
                
                # ============ 新增：设置余量参数 ============
                cp = builder.CutParameters
                if hasattr(cp, "FloorSameAsPartStock"):
                    cp.FloorSameAsPartStock = False
                    
                # 设置部件侧面余量
                part_stock = special.get('part_stock', 0.2)  # 默认0.2mm
                builder.CutParameters.PartStock.Value = part_stock
                self.print_log(f"设置部件侧面余量: {part_stock}mm", "SUCCESS")
                
                # 设置部件底面余量
                floor_stock = special.get('floor_stock', 0.8)  # 默认0.8mm
                builder.CutParameters.FloorStock.Value = floor_stock
                self.print_log(f"设置部件底面余量: {floor_stock}mm", "SUCCESS")
                # ===========================================

                # ============ 新增：设置切削顺序为始终深度优先 ============
                builder.CutParameters.CutOrder = NXOpen.CAM.CutParametersCutOrderTypes.DepthFirstAlways
                self.print_log("设置切削顺序: 始终深度优先", "SUCCESS")
                # =====================================================

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

            elif config['operation_type'] == "D4-螺旋":
                # 螺旋铣设置
                builder.CutLevel.RangeType = special.get('cut_level_range', NXOpen.CAM.CutLevel.RangeTypes.Automatic)
                builder.CutLevel.GlobalDepthPerCut.StepoverType = special.get(
                    'stepover_type', 
                    NXOpen.CAM.StepoverBuilder.StepoverTypes.Constant
                )
                max_depth = special.get('global_depth_per_cut', 10.1)
                builder.CutLevel.GlobalDepthPerCut.DistanceBuilder.Value = max_depth
                self.print_log(f"设置螺旋铣最大加工深度: {max_depth}mm", "SUCCESS")

                # ============ 新增：设置余量参数 ============
                # 设置部件侧面余量
                part_stock = special.get('part_stock', 0.2)  # 默认0.2mm
                builder.CutParameters.PartStock.Value = part_stock
                self.print_log(f"设置部件侧面余量: {part_stock}mm", "SUCCESS")
                
                # 设置部件底面余量
                floor_stock = special.get('floor_stock', 0.8)  # 默认0.8mm
                builder.CutParameters.FloorStock.Value = floor_stock
                self.print_log(f"设置部件底面余量: {floor_stock}mm", "SUCCESS")
                # ===========================================

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

                # ============ 仅对全精螺旋铣设置"非切削移动"高度起点为"当前层" ============
                # 获取加工阶段（从special_config中读取stage参数）
                stage = special.get('stage', '半精')  # 默认值为半精
                
                if stage == "全精":
                    try:
                        # 设置封闭区域进刀高度起点为当前层
                        if hasattr(builder.NonCuttingBuilder, 'EngageClosedAreaBuilder'):
                            if hasattr(builder.NonCuttingBuilder.EngageClosedAreaBuilder, 'HeightFrom'):
                                builder.NonCuttingBuilder.EngageClosedAreaBuilder.HeightFrom = NXOpen.CAM.NcmPlanarEngRetBuilder.MeasureHeightFrom.CurrentLevel
                                self.print_log(f"全精螺旋铣：设置封闭区域进刀高度起点为当前层", "SUCCESS")
                            
                    except Exception as e:
                        self.print_log(f"设置全精螺旋铣进刀高度起点时出错: {e}", "WARN")

                    # ============ 新增：层到层转移方法设置 ============
                    # 关键代码：设置层到层转移为使用转移方法
                    try:
                        if hasattr(builder.CutParameters.LevelToLevel, 'Type'):
                            builder.CutParameters.LevelToLevel.Type = NXOpen.CAM.LevelToLevel.Types.UseTransferMethod
                            self.print_log("设置层到层转移方法: 使用转移方法", "SUCCESS")
                        
                            
                    except Exception as e:
                        self.print_log(f"设置层到层转移方法时出错: {e}", "WARN")
                    # ====================================================
                else:
                    self.print_log(f"半精螺旋铣：保持默认进刀高度起点设置", "DEBUG")
                # =====================================================================


            elif config['operation_type'] == "D4-清角":
                # 清角特殊设置
                builder.ReferenceTool = NXOpen.CAM.Tool.Null
                builder.CutParameters.MergeDistance.Value = special.get('merge_distance', 3.0)
                builder.MinCutLength.Value = 0.5
                builder.CutParameters.CutDirection.Type = special.get('cut_direction')
                builder.CutParameters.CutOrder = special.get('cut_order')
                builder.CutParameters.PartStock.Value = special.get('part_stock', 0.0)
                
                # ============ 设置步距类型和每刀深度 ============
                stepover_type = special.get('stepover_type', NXOpen.CAM.StepoverBuilder.StepoverTypes.Constant)
                builder.CutLevel.GlobalDepthPerCut.StepoverType = stepover_type
                
                global_depth = special.get('global_depth_per_cut', 0.1)
                builder.CutLevel.GlobalDepthPerCut.DistanceBuilder.Value = global_depth
                self.print_log(f"设置清角每刀深度: {global_depth}mm", "SUCCESS")
                # ====================================================

                # ============ 新增：设置余量参数 ============
                # 设置部件侧面余量
                part_stock = special.get('part_stock', 0.03)  # 默认0.2mm
                builder.CutParameters.PartStock.Value = part_stock
                self.print_log(f"设置部件侧面余量: {part_stock}mm", "SUCCESS")

                # 设置部件底面余量
                floor_stock = special.get('floor_stock', 0.03)  # 默认0.8mm
                builder.CutParameters.FloorStock.Value = floor_stock
                self.print_log(f"设置部件底面余量: {floor_stock}mm", "SUCCESS")
                # ===========================================
                
                # ============ 设置转速、进给、横越参数 ============
                if 'spindle_rpm' in special:
                    builder.FeedsBuilder.SpindleRpmBuilder.Value = special['spindle_rpm']
                    self.print_log(f"设置主轴转速: {special['spindle_rpm']}RPM", "SUCCESS")
                
                if 'feed_per_tooth' in special:
                    builder.FeedsBuilder.FeedPerToothBuilder.Value = special['feed_per_tooth']
                    self.print_log(f"设置每齿进给: {special['feed_per_tooth']}mm/齿", "SUCCESS")
                
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

        except Exception as e:
            self.print_log(f"深度轮廓铣参数配置警告: {e}", "DEBUG")

    def _configure_surface_params(self, builder, config):
        """配置爬面参数"""
        try:
            special = config.get('special_config', {})

            builder.CutDirection = special.get('cut_direction')

            # 设置切削角度
            cutAngle = builder.DmareaMillingBuilder.NonSteepCutting.CutAngleBuilder
            cutAngle.Type = NXOpen.CAM.CutAngle.Types.Specify
            cutAngle.Value = special.get('cut_angle', 45.0)

            builder.CutParameters.PartStock.Value = special.get('part_stock', 0.0)

            # ============ 新增：设置转速、进给、横越参数 ============
            # 设置主轴转速
            if 'spindle_rpm' in special:
                builder.FeedsBuilder.SpindleRpmBuilder.Value = special['spindle_rpm']
                self.print_log(f"设置主轴转速: {special['spindle_rpm']}RPM", "SUCCESS")
            else:
                # 默认值
                builder.FeedsBuilder.SpindleRpmBuilder.Value = 1700.0
            
            # 设置每齿进给
            if 'feed_per_tooth' in special:
                builder.FeedsBuilder.FeedPerToothBuilder.Value = special['feed_per_tooth']
                self.print_log(f"设置每齿进给: {special['feed_per_tooth']}mm/齿", "SUCCESS")
            else:
                # 默认值
                builder.FeedsBuilder.FeedPerToothBuilder.Value = 2000.0
            
            # 设置横越速度
            if 'feed_rapid' in special:
                # 设置横越输出模式为G1
                builder.FeedsBuilder.FeedRapidOutput.Value = NXOpen.CAM.FeedRapidOutputMode.G1
                builder.FeedsBuilder.FeedRapidOutput.InheritanceStatus = False
                # 设置横越速度值
                builder.FeedsBuilder.FeedRapidBuilder.Value = special['feed_rapid']
                self.print_log(f"设置横越速度: {special['feed_rapid']}mm/min", "SUCCESS")
            else:
                # 默认值
                builder.FeedsBuilder.FeedRapidOutput.Value = NXOpen.CAM.FeedRapidOutputMode.G1
                builder.FeedsBuilder.FeedRapidOutput.InheritanceStatus = False
                builder.FeedsBuilder.FeedRapidBuilder.Value = 8000.0
            # ====================================================

            # ============ 新增：设置步距 ============
            stepover_distance = special.get('stepover_distance', 0.3)

            # 1）步距类型设为恒定（对应 journal-9 的两行）
            builder.DmareaMillingBuilder.SteepCutting.DepthPerCut.StepoverType = NXOpen.CAM.StepoverBuilder.StepoverTypes.Constant
            builder.DmareaMillingBuilder.NonSteepCutting.Stepover.StepoverType = NXOpen.CAM.StepoverBuilder.StepoverTypes.Constant

            # 2）步距数值设为 0.3（或配置里的值）
            if hasattr(builder.DmareaMillingBuilder, 'StepoverBuilder'):
                if hasattr(builder.DmareaMillingBuilder.StepoverBuilder, 'DistanceBuilder'):
                    builder.DmareaMillingBuilder.StepoverBuilder.DistanceBuilder.Intent = NXOpen.CAM.ParamValueIntent.PartUnits
                    builder.DmareaMillingBuilder.StepoverBuilder.DistanceBuilder.Value = stepover_distance
                    self.print_log(f"设置爬面步距: {stepover_distance}mm", "SUCCESS")
            # =======================================

        except Exception as e:
            self.print_log(f"曲面轮廓铣参数配置警告: {e}", "DEBUG")


    def _configure_flowcut_params(self, builder, config):
        """配置清根参数"""
        try:
            special = config.get('special_config', {})
            
            # ============ 设置重叠距离（关键参数） ============
            if 'flow_overlap_distance' in special:
                builder.FlowBuilder.FlowOverlapDistBuilder.Value = special['flow_overlap_distance']
                self.print_log(f"设置重叠距离: {special['flow_overlap_distance']}mm", "SUCCESS")
            
            # ============ 设置转速、进给、横越参数 ============
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
            
        except Exception as e:
            self.print_log(f"清根参数配置警告: {e}", "DEBUG")

    def create_operation(self, operation_key, face_inputs, tool_name, **params):
        """
        通用工序创建方法
        """
        if operation_key not in OPERATION_CONFIGS:
            raise ValueError(f"未知的工序类型: {operation_key}")

        config = OPERATION_CONFIGS[operation_key].copy()

        # 获取加工阶段，默认为"半精"
        stage = params.get('stage', '半精')

        # 获取图层参数，默认为20
        layer = params.get('layer', 20)

        if params:
            config['special_config'] = config.get('special_config', {}).copy()
            
            # 更新通用参数
            depth_val = None
            if 'global_depth_per_cut' in params:
                depth_val = params['global_depth_per_cut']
            elif 'depth_per_cut' in params:
                depth_val = params['depth_per_cut']
            elif 'max_depth' in params: # 兼容螺旋铣的 max_depth
                depth_val = params['max_depth']

            if depth_val is not None:
                # 统一写入 'global_depth_per_cut'，不再写 'depth_per_cut'
                config['special_config']['global_depth_per_cut'] = depth_val

            if 'reference_tool' in params:  # 参考刀具
                config['special_config']['reference_tool'] = params['reference_tool']
            if 'flow_overlap_distance' in params:  # 清根 重叠距离
                config['special_config']['flow_overlap_distance'] = params['flow_overlap_distance']
            if 'stepover_distance' in params:
                config['special_config']['stepover_distance'] = params['stepover_distance']
            if 'stepover_type' in params:
                config['special_config']['stepover_type'] = params['stepover_type']

            # ============ 新增：处理运动类型参数 ============
            if 'motion_type' in params:
                config['special_config']['motion_type'] = params['motion_type']
            # ===============================================


            # ============ 新增：处理图层参数 ============
            if 'layer' in params:
                config['special_config']['layer'] = params['layer']
            # ===========================================


            # ============ 新增：处理余量参数 ============
            if 'part_stock' in params:
                config['special_config']['part_stock'] = params['part_stock']
            if 'floor_stock' in params:
                config['special_config']['floor_stock'] = params['floor_stock']
            if 'wall_stock' in params:
                config['special_config']['wall_stock'] = params['wall_stock']
            # ============ 新增：底面毛坯厚度参数 ============
            if 'blank_distance' in params:
                config['special_config']['blank_distance'] = params['blank_distance']
            # ===========================================

            # ============ 新增：处理每刀切削深度参数 ============
            if 'depth_per_cut' in params:
                config['special_config']['depth_per_cut'] = params['depth_per_cut']

            # ============ 新增：处理转速、进给、横越参数 ============
            if 'spindle_rpm' in params:
                config['special_config']['spindle_rpm'] = params['spindle_rpm']
            if 'feed_per_tooth' in params:
                config['special_config']['feed_per_tooth'] = params['feed_per_tooth']
            if 'feed_rapid' in params:
                config['special_config']['feed_rapid'] = params['feed_rapid']
            # ====================================================


            # ============ 新增：处理爬面参数 ============
            if operation_key == "爬面_SIMPLE":
                if 'cut_angle' in params:
                    config['special_config']['cut_angle'] = params['cut_angle']

            # ============ 新增：将阶段信息放入special_config ============
            config['special_config']['stage'] = stage
            # ============================================================

        self.print_log(f"创建 {operation_key} 工序", "START")
        self.operation_count += 1

        try:
            # 1. 获取有效面（按图层过滤）
            valid_faces = self._get_valid_faces(face_inputs, layer)

            with self.undo_mark_context(f"创建{operation_key}"):
                groups = self.work_part.CAMSetup.CAMGroupCollection

                # 2. 根据阶段和图层获取对应的子程序组
                program_group = self.get_program_group_by_stage_and_layer(stage, layer)
                if not program_group:
                    program_group = groups.FindObject("PROGRAM")
                    self.print_log("使用默认PROGRAM组", "WARN")

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
                    NXOpen.CAM.OperationCollection.UseDefaultName.FalseValue,
                    custom_operation_name
                )

                # 3. 创建Builder
                builder_method_name = self.BUILDER_MAP[config['builder_type']]
                builder_method = getattr(self.work_part.CAMSetup.CAMOperationCollection, builder_method_name)
                builder = builder_method(op)

                try:
                    # 4. 设置几何（如果有有效面）
                    if valid_faces:
                        self._set_geometry(builder, valid_faces)

                    # 5. 配置特定参数
                    if config['builder_type'] == 'volume_25d':
                        self._configure_face_milling_params(builder, config)
                    elif config['builder_type'] == 'zlevel':
                        self._configure_zlevel_params(builder, config)
                    elif config['builder_type'] == 'surface':
                        if config['operation_type'] == "清根":  # 清根工序
                            self._configure_flowcut_params(builder, config)
                        else:  # 普通爬面
                            self._configure_surface_params(builder, config)

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
                "stage": stage,  # 添加阶段信息
                "layer": layer,
                "workpiece": workpiece_geometry,
                "message": f"{config['description']}创建完成，使用了 {len(valid_faces)} 个切削面"
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

        # 按阶段统计
        semi_results = [r for r in self.test_results if r.get('stage') == '半精']
        finish_results = [r for r in self.test_results if r.get('stage') == '全精']

        print(f"""
  刀轨生成摘要
  ----------------------------------------
  总工序数:   {self.operation_count}
  成功:       {self.success_count} ✅
  失败:       {self.failed_count} ❌
  成功率:     {success_rate:.1f}%
    按阶段统计:
  半精工序:   {len(semi_results)} 个
  全精工序:   {len(finish_results)} 个
        """.strip(), flush=True)

        if self.test_results:
            self.print_separator("-")
            print("  详细结果:")
            for i, result in enumerate(self.test_results, 1):
                status_emoji = "✅" if result['status'] == "Success" else "❌"
                name = result.get('name', result['type'])
                layer = result.get('layer', '未知')
                workpiece = result.get('workpiece', '未知')
                print(f"  {i}. {name} ({result['type']}) {status_emoji}")
                print(f"     图层: {layer}, 几何体: {workpiece}")
                if result.get('message'):
                    print(f"     信息: {result['message']}")
                if result.get('error'):
                    print(f"     错误: {result['error']}")

        self.print_separator("=")


# ==================================================================================
# 主流程  
# ==================================================================================



def generate_toolpath_workflow(part_path, half_spiral_json_path=None, half_spiral_reciprocating_json_path=None, 
                               half_surface_json_path=None,half_jiao_json_path=None,half_mian_json_path=None,mian_json_path=None,
                               spiral_json_path=None, spiral_reciprocating_json_path=None,reciprocating_json_path=None, surface_json_path=None,
                               gen_json_path=None,save_dir=None):
    """刀轨生成主工作流"""
    session = NXOpen.Session.GetSession()
    base_part, load_status = session.Parts.OpenBaseDisplay(part_path)
    work_part = session.Parts.Work

    generator = ToolpathGenerator(work_part, save_dir=save_dir)
    generator.print_header("NX CAM 刀轨生成工具 - 精简版")
    generator.print_log(f"零件: {work_part.Name}", "INFO")
    generator.print_log(f"测试模式: {'开启' if CONFIG['TEST_MODE'] else '关闭'}", "INFO")
    
    # 切换到加工环境
    generator.switch_to_manufacturing()

    # ==================== 创建程序组 ====================
    program_groups = generator.create_program_groups()
    if program_groups:
        generator.print_log("程序组创建/获取成功", "SUCCESS")
        for stage, group in program_groups.items():
            generator.print_log(f"程序组: {stage} (Tag: {group})", "DEBUG")
    
    # ==================== 加载JSON测试用例 ====================
    # 半精
    half_spiral_test_cases = generator.load_spiral_from_json(half_spiral_json_path, stage="半精")
    half_spiral_reciprocating_test_cases = generator.load_spiral_reciprocating_from_json(half_spiral_reciprocating_json_path, stage="半精")
    half_surface_test_cases = generator.load_half_surface_from_json(half_surface_json_path, stage="半精")
    half_jiao_test_cases = generator.load_jiao_from_json(half_jiao_json_path, stage="半精")
    half_mian_test_cases = generator.load_mian_from_json(half_mian_json_path, stage="半精")

    # 全精
    mian_test_cases = generator.load_mian_from_json(mian_json_path, stage="全精")
    spiral_test_cases = generator.load_spiral_from_json(spiral_json_path, stage="全精")
    spiral_reciprocating_test_cases = generator.load_spiral_reciprocating_from_json(spiral_reciprocating_json_path, stage="全精")
    reciprocating_test_cases = generator.load_half_surface_from_json(reciprocating_json_path, stage="全精")
    surface_test_cases = generator.load_surface_from_json(surface_json_path, stage="全精")
    gen_test_cases = generator.load_gen_from_json(gen_json_path, stage="全精")




    # 合并所有测试用例
    all_test_cases = []
    #半精
    all_test_cases.extend(half_spiral_test_cases)
    all_test_cases.extend(half_spiral_reciprocating_test_cases)
    all_test_cases.extend(half_surface_test_cases)
    all_test_cases.extend(half_jiao_test_cases)
    all_test_cases.extend(half_mian_test_cases)

    #全精
    all_test_cases.extend(mian_test_cases)
    all_test_cases.extend(spiral_test_cases)
    all_test_cases.extend(spiral_reciprocating_test_cases)
    all_test_cases.extend(reciprocating_test_cases)
    all_test_cases.extend(surface_test_cases)
    all_test_cases.extend(gen_test_cases)

    all_test_cases.extend(TEST_CASES)
    
    generator.print_log(f"总测试用例数: {len(all_test_cases)}", "INFO")
    generator.print_log(
        f"其中螺旋铣: {len(half_spiral_test_cases)} + {len(spiral_test_cases)}个, "
        f"往复等高: {len(half_spiral_reciprocating_test_cases)} + {len(spiral_reciprocating_test_cases)} 个, "
         f"清根: {len(gen_test_cases)} 个",
        "INFO"
    )

    # 执行所有测试用例
    for test_case in all_test_cases:
        try:
            if len(test_case) == 4:
                op_key, face_ids, tool_name, extra_params = test_case
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

            #半精
            half_spiral_json_path=CONFIG["JSON_HALF_SPIRAL_PATH"],                     # 对应的是 xx_半精_螺旋.json
            half_spiral_reciprocating_json_path=CONFIG["JSON_HALF_SPIRAL_RECIPROCATING_PATH"], # 对应的是 xx_半精_螺旋_往复等高.json
            half_surface_json_path=CONFIG["JSON_HALF_SURFACE_PATH"],  # 对应的是 xx_半精_爬面.json
            half_jiao_json_path=CONFIG["JSON_HALF_JIAO_PATH"],  # 对应的是 xx_半精_清角.json
            half_mian_json_path=CONFIG["JSON_HALF_MIAN_PATH"], # 对应的是 xx_半精_面铣.json

            #全精
            mian_json_path=CONFIG["JSON_MIAN_PATH"],  # 对应的是 xx_全精_面铣.json
            spiral_json_path=CONFIG["JSON_SPIRAL_PATH"], # 对应的是 xx_全精_螺旋.json
            spiral_reciprocating_json_path=CONFIG["JSON_SPIRAL_RECIPROCATING_PATH"], # 对应的是 xx_全精_螺旋_往复等高.json
            reciprocating_json_path=CONFIG["JSON_RECIPROCATING_PATH"], # 对应的是 xx_全精_往复等高.json
            surface_json_path=CONFIG["JSON_SURFACE_PATH"],  # 对应的是 xx_全精_爬面.json
            gen_json_path=CONFIG["JSON_GEN_PATH"],  # 对应的是 xx_全精_清根.json

            # 新增：保存目录参数
            save_dir=r'C:\Projects\NC\output\05_CAM\Daogui_prt'

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