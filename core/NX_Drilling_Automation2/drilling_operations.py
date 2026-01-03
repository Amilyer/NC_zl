# -*- coding: utf-8 -*-
"""
钻孔操作模块
包含刀具创建、工序设置、钻孔操作等功能
"""

import NXOpen
import NXOpen.CAM
from utils import print_to_info_window, handle_exception
from geometry import GeometryHandler
import drill_config

class DrillingOperationHandler:
    """钻孔操作处理器"""

    def __init__(self, session, work_part):
        self.session = session
        self.work_part = work_part
        self.cam_setup = work_part.CAMSetup

    def get_or_create_program_group(self, parent_group_name="NC_PROGRAM", category="hole_making", group_name="A"):
        """获取或创建程序组"""

        try:
            # 先检测是否存在
            existing = self.cam_setup.CAMGroupCollection.FindObject(group_name)
            if existing:
                print_to_info_window(f"✔ 已存在程序组: {group_name}")
                return existing
        except:
            pass

        # 创建程序组
        try:
            parent_group = self.cam_setup.CAMGroupCollection.FindObject(parent_group_name)
            program = self.cam_setup.CAMGroupCollection.CreateProgram(
                parent_group,
                category,
                "PROGRAM",
                NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue,
                group_name
            )
            print_to_info_window(f"🆕 已创建程序组: {group_name}")
            return program
        except Exception as ex:
            return handle_exception("创建程序失败", str(ex))

    def set_tool_drive_point(self, hole_drilling_builder, point_type="SYS_CL_TIP"):
        """设置钻孔刀具驱动点"""

        # 参数检查
        if point_type not in ["SYS_CL_TIP", "SYS_CL_SHOULDER"]:
            raise ValueError("point_type 必须为 SYS_CL_TIP 或 SYS_CL_SHOULDER")

        try:
            # 正确方式：直接调用 SetToolDrivePoint()
            hole_drilling_builder.SetToolDrivePoint(point_type)
            print_to_info_window(f"✔ 设置刀具驱动点: {point_type}")
        except Exception as e:
            print_to_info_window(f"⚠ 无法设置刀具驱动点: {str(e)}")

    # def set_bottom_stock(self, hole_drilling_builder, value=0.0):
    #     """设置钻孔底部余量"""
    #     cut_params = hole_drilling_builder.CuttingParameters
    #     cut_params.BottomStock.Value = value
    #     print_to_info_window(f"✔ 设置底部余量: {value} mm")
    def set_bottom_stock(self, hole_drilling_builder, value=0.0):
        """
        设置钻孔底偏置（Bottom Offset）
        等同于录制宏里的：
            holeDrillingBuilder.CuttingParameters.BottomOffset.Distance = value
        """
        cut_params = hole_drilling_builder.CuttingParameters
        cut_params.BottomOffset.Distance = float(value)

        print_to_info_window(f"✔ 设置钻孔底偏置 BottomOffset = {value} mm")

    def set_cycle_deep_drill(self, hole_drilling_builder, step_distance=3.0, cycle_type="Drill,Deep"):
        """设置循环类型"""

        hole_drilling_builder.CycleTable.CycleType = cycle_type
        hole_drilling_builder.CycleTable.AxialStepover.StepoverType = NXOpen.CAM.StepoverBuilder.StepoverTypes.Constant
        hole_drilling_builder.CycleTable.AxialStepover.DistanceBuilder.Value = step_distance

        if cycle_type == "Drill,Deep":
            print_to_info_window(f"✔ 设置循环类型: 深孔钻，步进距离 {step_distance} mm")
        elif cycle_type == "Drill":
            print_to_info_window(f"✔ 设置循环类型: 标准钻，步进距离 {step_distance} mm")

    def set_extend_path_offsets(self, hole_drilling_builder, top_offset=0.0, all_bottom_offset=0.0, rapto_offset=0.0):
        """
        设置延伸路径偏置：顶偏置、底偏置、Rapto 偏置
        NX 2312 录制宏中真实可用的接口：
            TopOffset.Distance
            BottomOffset.Distance
            RaptoOffset.Distance
        """

        cut_params = hole_drilling_builder.CuttingParameters

        # 顶偏置
        try:
            cut_params.TopOffset.Distance = top_offset
        except:
            pass

        # 底偏置
        try:
            cut_params.BottomOffset.Distance = all_bottom_offset
        except:
            pass

        # Rapto 偏置
        try:
            cut_params.RaptoOffset.Distance = rapto_offset
        except:
            pass

        print_to_info_window(
            f"✔ 设置延伸路径偏置：顶 {top_offset} mm, 底 {all_bottom_offset} mm, Rapto {rapto_offset} mm"
        )

    # 创建钻刀
    def get_or_create_drill_tool(self, drill_name="STD_DRILL", diameter=1.0, tip_diameter=1.0,
                                 parent_group_name="GENERIC_MACHINE", tool_name="z-zxz"):
        """获取或创建钻刀工具"""

        try:
            parent_group = self.cam_setup.CAMGroupCollection.FindObject(parent_group_name)
            if parent_group is None:
                raise ValueError(f"未找到刀具组 {parent_group_name}")

            # 查找已有刀具
            try:
                tool = self.cam_setup.CAMGroupCollection.FindObject(tool_name)
                print_to_info_window(f"✔ 已找到钻刀工具: {tool_name}")
                return tool
            except Exception:
                pass

            # 创建钻刀
            tool_obj = self.cam_setup.CAMGroupCollection.CreateTool(
                parent_group,
                "hole_making",
                drill_name,
                NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue,
                tool_name
            )

            # 创建对应的 Builder
            if drill_name.upper() == "CENTERDRILL":
                drill_builder = self.cam_setup.CAMGroupCollection.CreateDrillCenterBellToolBuilder(tool_obj)
            else:
                drill_builder = self.cam_setup.CAMGroupCollection.CreateDrillStdToolBuilder(tool_obj)

            # 设置参数
            drill_builder.TlDiameterBuilder.Value = diameter
            if hasattr(drill_builder, "TlTipDiameterBuilder"):
                drill_builder.TlTipDiameterBuilder.Value = tip_diameter
            # 设置钻刀参数
            if drill_name.upper() == "STD_DRILL":
                drill_builder.TlPointAngBuilder.Value = drill_config.TIP_ANGLE  # 设置刀尖角度
                drill_builder.TlPointLengthBuilder.Value = drill_config.TIP_LEN  # 刀尖长度
                drill_builder.TlCor1RadBuilder.Value = drill_config.CORNER_RADIUS  # 拐角半径
                drill_builder.TlHeightBuilder.Value = drill_config.LENGTH  # 长度
                drill_builder.TlFluteLnBuilder.Value = drill_config.BLADE_LENGTH  # 刀刃长度
                drill_builder.TlNumFlutesBuilder.Value = drill_config.BLADE_NUMBER  # 刀刃数

            drill_builder.Commit()
            drill_builder.Destroy()

            print_to_info_window(f"🆕 已创建钻刀工具: {tool_name}（直径 {diameter}mm，刀尖 {tip_diameter}mm）")
            return tool_obj

        except Exception as ex:
            return handle_exception("创建/获取钻刀工具失败", str(ex))

    # 创建铣刀
    def get_or_create_mill_tool(self, tool_type="MILL", diameter=1.0, R1=0.0,
                                parent_group_name="GENERIC_MACHINE", tool_name="milling_tool"):
        """获取或创建铣刀工具"""

        try:
            # 获取父刀具组
            parent_group = self.cam_setup.CAMGroupCollection.FindObject(parent_group_name)
            if parent_group is None:
                raise ValueError(f"未找到刀具组 {parent_group_name}")

            # 查找已有的铣刀
            try:
                tool = self.cam_setup.CAMGroupCollection.FindObject(tool_name)
                print_to_info_window(f"✔ 已找到铣刀工具: {tool_name}")
                return tool
            except Exception:
                pass

            # 创建铣刀
            tool_obj = self.cam_setup.CAMGroupCollection.CreateTool(
                parent_group,
                "hole_making",
                tool_type,
                NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue,
                tool_name
            )

            # 创建铣刀的 Builder
            mill_builder = self.cam_setup.CAMGroupCollection.CreateMillToolBuilder(tool_obj)

            # 设置参数
            mill_builder.TlDiameterBuilder.Value = diameter
            if hasattr(mill_builder, "TlR1Builder"):
                mill_builder.TlR1Builder.Value = R1

            # 提交并销毁 Builder
            mill_builder.Commit()
            mill_builder.Destroy()

            print_to_info_window(f"🆕 已创建铣刀工具: {tool_name}（直径 {diameter}mm，R1 {R1}mm）")
            return tool_obj

        except Exception as e:
            print_to_info_window(f"❌ 错误: {str(e)}")
            raise

    # 钻孔工序
    def create_drill_operation(
            self,
            operation_type="DRILL",
            tool_type="STD_DRILL",
            tool_name="CENTERDRILL_D3.0",
            geometry_name=drill_config.DEFAULT_MCS_NAME,  # ⚠️ 保留参数，但不再使用
            orient_geometry_name="WORKPIECE",  # ✅ MCS 名称
            parent_group_name="NC_PROGRAM",
            group_name="A",
            method_group_name="METHOD",
            hole_features=None,
            predefined_depth=drill_config.DEFAULT_DRILL_DEPTH,
            diameter=3.0,
            tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
            operation_name=None,
            drive_point="SYS_CL_TIP",
            is_through=False,
            step_distance=drill_config.DEFAULT_STEP_DISTANCE,
            feed_rate=drill_config.DEFAULT_FEED_RATE,
            cycle_type="Drill"
    ):
        """创建钻孔工序（创建时直接绑定 MCS 作为 Geometry）"""

        mark_id = None
        try:
            # -------------------------------------------------
            # Undo Mark
            # -------------------------------------------------
            mark_id = self.session.SetUndoMark(
                NXOpen.Session.MarkVisibility.Visible,
                "创建钻孔工序"
            )

            cam_groups = self.cam_setup.CAMGroupCollection

            # -------------------------------------------------
            # Program Group
            # -------------------------------------------------
            try:
                program_group = cam_groups.FindObject(group_name)
            except:
                program_group = self.get_or_create_program_group(
                    parent_group_name=parent_group_name,
                    group_name=group_name
                )

            # -------------------------------------------------
            # Method Group
            # -------------------------------------------------
            try:
                method_group = cam_groups.FindObject(method_group_name)
            except:
                raise ValueError(f"找不到方法组: {method_group_name}")

            # -------------------------------------------------
            # MCS（作为 Geometry 使用）
            # -------------------------------------------------
            try:
                mcs_group = cam_groups.FindObject(orient_geometry_name)
            except:
                raise ValueError(f"找不到 MCS 几何组: {orient_geometry_name}")

            # -------------------------------------------------
            # Operation Name
            # -------------------------------------------------
            if operation_name is None:
                operation_name = f"{operation_type}_AUTO"
                
            # -------------------------------------------------
            # Tool 获取或创建钻刀工具
            # -------------------------------------------------
            tool_obj = self.get_or_create_drill_tool(
                tool_type,
                round(diameter, 1),
                tip_diameter,
                "GENERIC_MACHINE",
                tool_name
            )
            tool_group = tool_obj if tool_obj else self.cam_setup.CAMGroupCollection.FindObject("NONE")
            # -------------------------------------------------
            # 创建工序（⭐ Geometry = MCS）
            # -------------------------------------------------                                
            operation = self.cam_setup.CAMOperationCollection.Create(
                program_group,  # 工序属于哪个程序
                method_group,   # 用什么加工方法（钻孔 / 铣削 / 参数模板）
                tool_group,  # 使用哪把刀
                mcs_group,  # 加工哪些几何体
                "hole_making",
                operation_type,
                NXOpen.CAM.OperationCollection.UseDefaultName.FalseValue,
                operation_name
            )                                                

            # -------------------------------------------------
            # Builder
            # -------------------------------------------------
            builder = self.cam_setup.CAMOperationCollection.CreateHoleDrillingBuilder(operation)

            # -------------------------------------------------
            # Feeds & Speeds
            # -------------------------------------------------
            builder.FeedsBuilder.SurfaceSpeedBuilder.Value = drill_config.DEFAULT_SPINDLE_SPEED
            builder.FeedsBuilder.FeedCutBuilder.Value = feed_rate

            # -------------------------------------------------
            # 深度逻辑（通孔 / 盲孔分离）
            # -------------------------------------------------
            if is_through:
                # 通孔：不用 PredefinedDepth
                builder.PredefinedDepth.Status = False
            else:
                builder.PredefinedDepth.Status = True
                builder.PredefinedDepth.Value = predefined_depth

            # -------------------------------------------------
            # 自定义加工参数
            # -------------------------------------------------
            # Drive Point（防御性设置）
            try:
                self.set_tool_drive_point(builder, drive_point)
            except:
                pass

            self.set_cycle_deep_drill(builder, step_distance, cycle_type)

            self.set_extend_path_offsets(
                builder,
                top_offset=drill_config.DEFAULT_TOP_OFFSET,
                all_bottom_offset=drill_config.DEFAULT_ALL_BOTTOM_OFFSET,
                rapto_offset=drill_config.DEFAULT_RAPTO_OFFSET
            )

            # -------------------------------------------------
            # 底部余量（仅通孔有效）
            # -------------------------------------------------
            bottom_offset = diameter * 0.7 if is_through else 0.0
            self.set_bottom_stock(builder, bottom_offset)

            # -------------------------------------------------
            # 绑定孔特征
            # -------------------------------------------------
            if hole_features:
                feature_geometry = builder.GetFeatureGeometry()
                geometry_list = feature_geometry.GeometryList

                for i, feature in enumerate(hole_features):
                    feature_set = feature_geometry.AddFeatureSet(
                        NXOpen.CAM.CAMFeature.Null,
                        f"NXHOLE_{i + 1}"
                    )
                    feature_set.CreateFeature([feature])

                    cam_feature = feature_set.GetFeature()

                    # 通孔强制标记（NX 2312 可用）
                    if is_through:
                        cam_feature.OverrideAttributeValue("IS_THROUGH", True)

                    geometry_list.Append(feature_set)

            # -------------------------------------------------
            # Commit 提交工序
            # -------------------------------------------------
            operation_obj = builder.Commit()
            builder.Destroy()

            # -------------------------------------------------
            # Clean Undo
            # -------------------------------------------------
            self.session.DeleteUndoMark(mark_id, None)

            # -----------------------------
            # 生成刀轨
            # -----------------------------
            try:
                self.cam_setup.GenerateToolPath([operation_obj])
                print_to_info_window(f"✅ 已创建钻孔工序: {operation_name}")
            except Exception as ex:
                print_to_info_window(f"⚠️ 生成刀轨失败: {ex}")

            self.session.DeleteUndoMark(mark_id, None)
            return operation_obj

        except Exception as ex:
            if mark_id:
                self.session.DeleteUndoMark(mark_id, None)
            return handle_exception("创建钻孔工序失败", str(ex))

    # 铣孔工序
    def create_hole_milling_operation(
            self,
            operation_type="HOLE_MILLING",
            tool_type="MILL",
            tool_name="MILL_D20_R1",
            geometry_name="MCS",
            orient_geometry_name="WORKPIECE",
            parent_group_name="NC_PROGRAM",
            group_name="X",
            method_group_name="METHOD",
            hole_features=None,
            predefined_depth=10.0,
            diameter=20.0,
            operation_name=None,
            corner_radius=1.0,  # ⭐ R角半径（替代钻孔刀尖直径）
            axial_distance=0.3,

    ):
        """
        创建孔铣（HOLE_MILLING）工序。
        """

        try:
            mark_id = self.session.SetUndoMark(
                NXOpen.Session.MarkVisibility.Visible,
                "创建孔铣工序"
            )

            cam_groups = self.cam_setup.CAMGroupCollection

            # 获取组
            try:
                program_group = cam_groups.FindObject(group_name)
            except:
                program_group = self.get_or_create_program_group(parent_group_name=parent_group_name,
                                                                 group_name=group_name)
            try:
                method_group = cam_groups.FindObject(method_group_name)
            except:
                raise ValueError(f"找不到方法组: {method_group_name}")

            try:
                geom_group = cam_groups.FindObject(geometry_name)
            except:
                raise ValueError(f"找不到几何组: {geometry_name}")

            try:
                orient_geometry = cam_groups.FindObject(orient_geometry_name)
            except:
                raise ValueError(f"找不到定向几何体: {orient_geometry_name}")

            if operation_name is None:
                operation_name = "HOLE_MILLING_AUTO"
                
            
            # -----------------------------
            # 创建铣刀（刀具）
            # -----------------------------
            tool_obj = self.get_or_create_mill_tool(
                tool_type=tool_type,
                diameter=diameter,
                R1=corner_radius,
                parent_group_name="GENERIC_MACHINE",
                tool_name=tool_name
            )
            
            tool_group = tool_obj if tool_obj else cam_groups.FindObject("NONE")

            # -----------------------------                            
            # 创建孔铣工序
            # -----------------------------
            operation = self.cam_setup.CAMOperationCollection.Create(
                program_group,
                method_group,
                tool_group,
                orient_geometry,
                "hole_making",
                operation_type,
                NXOpen.CAM.OperationCollection.UseDefaultName.FalseValue,
                operation_name
            )
            builder = self.cam_setup.CAMOperationCollection.CreateCylinderMillingBuilder(operation)

            # -----------------------------
            # 基本切削参数
            # -----------------------------
            builder.FeedsBuilder.SpindleRpmBuilder.Value = drill_config.DEFAULT_SPINDLE_SPEED
            builder.FeedsBuilder.FeedCutBuilder.Value = drill_config.DEFAULT_FEED_RATE

            builder.PredefinedDepth.Value = predefined_depth
            builder.PredefinedDepth.Status = True

            self.set_extend_path_offsets(builder, top_offset=drill_config.DEFAULT_X_TOP_OFFSET,
                                         all_bottom_offset=drill_config.DEFAULT_ALL_BOTTOM_OFFSET,
                                         rapto_offset=drill_config.DEFAULT_RAPTO_OFFSET)

            # 每次下深（严格对应录制宏）
            builder.AxialDistance.Value = axial_distance
            builder.AxialDistance.Intent = NXOpen.CAM.ParamValueIntent.PartUnits

            # 设置轴向步距类型为“刀路数 Number”
            builder.AxialStepover.StepoverType = NXOpen.CAM.StepoverBuilder.StepoverTypes.Number
            # 设置刀路数量，例如 3 刀
            builder.AxialStepover.DistanceBuilder.Value = drill_config.DEFAULT_RADIAL_TOOL_NUMBER

            # 设置径向步距最大距离
            builder.RadialStepover.DistanceBuilder.Value = drill_config.DEFAULT_RADIAL_MAX_DISTANCE
            builder.RadialStepover.DistanceBuilder.Intent = NXOpen.CAM.ParamValueIntent.PartUnits

            # 设置最小螺旋直径
            builder.MinimumHelixDiameter.Intent = NXOpen.CAM.ParamValueIntent.ToolDep
            builder.MinimumHelixDiameter.Value = 70.0

            # -----------------------------
            # 绑定几何体（孔）
            # -----------------------------
            if hole_features:
                feature_geo = builder.GetFeatureGeometry()
                geo_list = feature_geo.GeometryList
                feature_geo.SetDefaultAttribute("AXIAL_STEPOVER", drill_config.DEFAULT_AXIAL_MAX_DISTANCE)
                for i, feature in enumerate(hole_features):
                    feature_set = feature_geo.AddFeatureSet(
                        NXOpen.CAM.CAMFeature.Null,
                        f"MILLHOLE_{i + 1}"
                    )
                    feature_set.CreateFeature([feature])
                    geo_list.Append(feature_set)
                    try:
                        created_feature = feature_set.GetFeature()
                        # 方法1：OverrideAttributeValue（如果可用）
                        created_feature.OverrideAttributeValue("START_DIAMETER", drill_config.DEFAULT_STRAT_DIAMETER)
                    except Exception as ex:
                        print_to_info_window(f"⚠️ 设置 START_DIAMETER 失败: {ex}")
            # -----------------------------
            # 提交工序
            # -----------------------------
            operation_obj = builder.Commit()
            builder.Destroy()

            # -----------------------------
            # 生成刀轨
            # -----------------------------
            try:
                self.cam_setup.GenerateToolPath([operation_obj])
                print_to_info_window(f"✅ 已创建孔铣工序: {operation_name}")
            except Exception as ex:
                print_to_info_window(f"⚠️ 生成刀轨失败: {ex}")

            self.session.DeleteUndoMark(mark_id, None)
            return operation_obj

        except Exception as ex:
            self.session.DeleteUndoMark(mark_id, None)
            return handle_exception("创建孔铣工序失败", str(ex))
