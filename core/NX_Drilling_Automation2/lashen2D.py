# -*- coding: utf-8 -*-
"""
主流程模块
包含完整钻孔自动化的主流程控制
"""

from utils import print_to_info_window, handle_exception, point_from_angle, analyze_arc, safe_origin, euclidean_distance
from geometry import GeometryHandler
import drill_config
import NXOpen
import NXOpen.CAM
import math
import NXOpen.Features
import NXOpen.GeometricUtilities

import traceback
import NXOpen.UF
from process_info import ProcessInfoHandler
from parameter_parser import ParameterParser
from mirror_operations import MirrorHandler


class MainWorkflow2:
    """主流程控制器"""

    def __init__(self, session, work_part, uf_session):
        self.session = session
        self.work_part = work_part
        self.uf_session = uf_session
        self.mirror_handler = MirrorHandler(session, work_part)
        # 初始化各个处理器
        self.geometry_handler = GeometryHandler(session, work_part)
        self.process_info_handler = ProcessInfoHandler()
        self.parameter_parser = ParameterParser()

    def run_workflow(self):
        """执行完整的工作流程"""

        try:
            print_to_info_window("=" * 60)
            print_to_info_window("开始执行NX钻孔自动化流程")
            print_to_info_window("=" * 60)

            # 第一步：预处理，提取加工信息
            print_to_info_window("第一步：预处理，提取加工信息")
            annotated_data = self.process_info_handler.extract_and_process_notes(self.work_part)
            print_to_info_window(annotated_data)
            processed_result = self.parameter_parser.process_hole_data(annotated_data)
            print_to_info_window(annotated_data)
            annotated_data = annotated_data["主体加工说明"]
            if not processed_result:
                return handle_exception("加工信息处理失败，流程终止")
            # 获取毛坯尺寸
            print_to_info_window(annotated_data)
            lwh_point = self.process_info_handler.get_workpiece_dimensions(annotated_data["尺寸"])
            # 获取加工坐标原点
            original_point = self.geometry_handler.get_start_point(lwh_point)
            if original_point[0] != 0.0 and original_point[1] != 0.0:
                # 移动2D图到绝对坐标系
                self.geometry_handler.move_objects_point_to_point(original_point, drill_config.DEFAULT_ORIGIN_POINT)
            # 适合窗口
            self.work_part.ModelingViews.WorkView.Fit()
            point2 = self.diagonal_point(lwh_point)
            # 拉伸毛坯
            self.create_block_by_diagonal(drill_config.DEFAULT_ORIGIN_POINT, point2)
            # 获取所有实线封闭圆
            z_circle_list = self._extract_and_classify_circles(lwh_point)["front_circles"]

            c_circle_list = self._extract_and_classify_circles(lwh_point)["side_circles"]

            x_circle_obj_list = self._extract_and_classify_circles(lwh_point)["x_circles"]

            x_side_circle_list = self._extract_and_classify_circles(lwh_point)["x_side_circles"]
            center_point_list = self._extract_and_classify_circles(lwh_point)["center_point_list"]

            print_to_info_window("---------------------------------------")


            #
            body = self.find_body_by_point((0.0,0.0,0.0))
            # 主视图实心圆匹配
            tag_circle_map = self.process_info_handler.fenlei(self.work_part, z_circle_list ,lwh_point, None, None)

            self.solid_circle(processed_result, tag_circle_map, lwh_point, body )
            print_to_info_window("------------------------实现封闭圆处理完成-----------------------")
            # 侧视图实心圆匹配
            c_tag_circle_map = self.process_info_handler.fenlei(self.work_part, c_circle_list ,lwh_point, None, None)
            self.solid_circle(processed_result, c_tag_circle_map, lwh_point, body)
            print_to_info_window("------------------------侧面实现封闭圆处理完成-----------------------")
            # 主视图虚心圆匹配
            tag_circle_map = self.process_info_handler.fenlei(self.work_part, x_circle_obj_list ,lwh_point, None, None)
            origin_bs = "Z"
            self.dashed_circle(processed_result, tag_circle_map, lwh_point, body, origin_bs)
            print_to_info_window("------------------------虚线圆处理完成-----------------------")
            # 侧视图虚心圆匹配
            tag_circle_map = self.process_info_handler.fenlei(self.work_part, x_side_circle_list ,lwh_point, None, None)
            origin_bs = "X"
            self.dashed_circle(processed_result, tag_circle_map, lwh_point, body, origin_bs)
            print_to_info_window("------------------------侧面虚线圆处理完成-----------------------")
            # 槽拉伸
            is_closed_list = self.find_groove_obj(center_point_list)
            unique_list = self.unique_closed_paths(is_closed_list)
            for curve_list in unique_list:
                tool_body_list = self.perform_extrude(curve_list,(0.0,0.0,0.0),(0.0,0.0,-1.0),15)
                final_body_list = self.subtract([body], tool_body_list)
                self.remove_parameters(final_body_list)
            print_to_info_window("------------------------封闭槽处理完成-----------------------")

        except Exception as e:
            return handle_exception("主流程执行失败", str(e))
    # ---------------------------
    # 获取体的对角点
    # ---------------------------
    def diagonal_point(self, lwh_point):
        # 获取所有曲线
        curves_list = self.geometry_handler.get_all_curves()
        # 对角点
        x_point = None
        y_point = None
        for curve in curves_list:
            if isinstance(curve, NXOpen.Line):
                p1 = (curve.StartPoint.X, curve.StartPoint.Y, curve.StartPoint.Z)
                p2 = (curve.EndPoint.X, curve.EndPoint.Y, curve.EndPoint.Z)
                if curve.StartPoint.X == 0.0 and curve.StartPoint.Y == 0.0:
                    if curve.EndPoint.X == 0.0:
                        y_point = curve.EndPoint.Y
                    else:
                        x_point = curve.EndPoint.X
                elif curve.EndPoint.X == 0.0 and curve.EndPoint.Y == 0.0:
                    if curve.StartPoint.X == 0.0:
                        y_point = curve.StartPoint.Y
                    else:
                        x_point = curve.StartPoint.X
        return (x_point, y_point, -lwh_point[2])

    def create_block_by_diagonal(self, point1, point2):
        """
        通过对角点创建长方体

        创建一个长方体，由两个对角点定义。这是创建规则几何体的最简单方法。

        Args:
            point1 (NXOpen.Point3d): 第一个对角点坐标
            point2 (NXOpen.Point3d): 第二个对角点（对角位置）

        Returns:
            NXOpen.Body: 创建的长方体Body对象
            None: 创建失败时返回

        Note:
            - 点的顺序不影响结果
            - 创建的Body会自动添加到workPart中
            - Body默认在当前图层

        Example:
            >>> # 创建一个100x50x30的长方体
            >>> p1 = NXOpen.Point3d(0, 0, 0)
            Point( 1900.00000000, 1200.00000000, -58.00000000 )
            >>> p2 = NXOpen.Point3d(1900.0, 1200.0, -58.0)
            >>> body = toolbox.create_block_by_diagonal(p1, p2)
            >>> if body:
            ...     print(f"长方体创建成功，Tag={body.Tag}")
        """
        try:

            # 创建长方体特征构建器
            blockBuilder = self.work_part.Features.CreateBlockFeatureBuilder(None)
            blockBuilder.Type = NXOpen.Features.BlockFeatureBuilder.Types.DiagonalPoints

            p1 = NXOpen.Point3d(point1[0], point1[1], point1[2])
            p2 = NXOpen.Point3d(point2[0], point2[1], point2[2])
            # 设置对角点
            blockBuilder.SetTwoDiagonalPoints(p1, p2)

            # 提交特征并获取Body
            feature = blockBuilder.CommitFeature()
            body = feature.GetBodies()[0]

            # 清理构建器
            blockBuilder.Destroy()

            print_to_info_window(f"  ✓ 毛坯创建成功 (Tag: {body.Tag})")
            return body

        except Exception as e:
            print_to_info_window(f"  ✗ 创建毛坯出错: {e}")
            traceback.print_exc()
            return None
    # 实心圆拉伸
    def solid_circle(self, processed_result, tag_circle_map, lwh_point, body):
        for tag in processed_result:
            for key in tag_circle_map:
                if tag != key:
                    continue
                for arc_obj in tag_circle_map[key]:
                    if (processed_result[tag]["circle_type"] != "正沉头" and processed_result[tag]["circle_type"] != "背沉头"):

                        predefined_depth = (lwh_point[2] if processed_result[tag]["is_through_hole"]
                                            else processed_result[tag]["depth"]["hole_depth"])
                        if predefined_depth is None :
                            predefined_depth = lwh_point[2]
                        tool_body_list = self.perform_extrude(arc_obj[0][0], arc_obj[0][1], arc_obj[0][3],
                                                            predefined_depth)
                        final_body_list = self.subtract([body], tool_body_list)
                        self.remove_parameters(final_body_list)
                    elif processed_result[tag]["circle_type"] == "正沉头" or processed_result[tag]["circle_type"] == "背沉头":
                        real_diamater = round(float(processed_result[tag]["real_diamater"]), 1)
                        hole_depth = processed_result[tag]["depth"]["hole_depth"]
                        real_diamater_head = round(float(processed_result[tag]["real_diamater_head"]), 1)
                        head_depth = round(float(processed_result[tag]["depth"]["head_depth"]), 1)
                        if real_diamater == round(arc_obj[0][2], 1):
                            if hole_depth is None:
                                hole_depth = lwh_point[2]
                            tool_body_list = self.perform_extrude(arc_obj[0][0], arc_obj[0][1], arc_obj[0][3],
                                                                hole_depth)
                            final_body_list = self.subtract([body], tool_body_list)
                            self.remove_parameters(final_body_list)
                        elif real_diamater_head == round(arc_obj[0][2], 1):
                            if head_depth is None:
                                head_depth = lwh_point[2]
                            tool_body_list = self.perform_extrude(arc_obj[0][0], arc_obj[0][1], arc_obj[0][3],
                                                                head_depth)
                            final_body_list = self.subtract([body], tool_body_list)
                            self.remove_parameters(final_body_list)

    def dashed_circle(self, processed_result, tag_circle_map, lwh_point, body, origin_bs):
        for tag in processed_result:
            for key in tag_circle_map:
                if tag != key:
                    continue
                for arc_obj in tag_circle_map[key]:
                    if processed_result[tag]["circle_type"] != "背沉头":
                        real_diamater = round(float(processed_result[tag]["real_diamater"]), 1)
                        hole_depth = processed_result[tag]["depth"]["hole_depth"]
                        if real_diamater == round(arc_obj[0][2], 1):
                            if origin_bs == "Z":
                                origin = (arc_obj[0][1][0], arc_obj[0][1][1], lwh_point[2])
                            elif origin_bs == "X":
                                origin = (lwh_point[0], arc_obj[0][1][1], arc_obj[0][1][2])
                            if hole_depth is None:
                                hole_depth = lwh_point[2]
                            tool_body_list = self.perform_extrude(arc_obj[0][0], origin, arc_obj[0][3],
                                                                hole_depth)
                            final_body_list = self.subtract([body], tool_body_list)
                            self.remove_parameters(final_body_list)
                        # 背沉头
                    elif processed_result[tag]["circle_type"] == "背沉头":
                        real_diamater = round(float(processed_result[tag]["real_diamater"]), 1)
                        hole_depth = processed_result[tag]["depth"]["hole_depth"]
                        real_diamater_head = round(float(processed_result[tag]["real_diamater_head"]), 1)
                        head_depth = round(float(processed_result[tag]["depth"]["head_depth"]), 1)
                        if real_diamater == round(arc_obj[0][2], 1):
                            if origin_bs == "Z":
                                origin = (arc_obj[0][1][0], arc_obj[0][1][1], lwh_point[2])
                            elif origin_bs == "X":
                                origin = (lwh_point[0], arc_obj[0][1][1], arc_obj[0][1][2])
                            if hole_depth is None:
                                hole_depth = lwh_point[2]
                            tool_body_list = self.perform_extrude(arc_obj[0][0], origin, arc_obj[0][3],
                                                                hole_depth)
                            final_body_list = self.subtract([body], tool_body_list)
                            self.remove_parameters(final_body_list)
                        elif real_diamater_head == round(arc_obj[0][2], 1):
                            if origin_bs == "Z":
                                origin = (arc_obj[0][1][0], arc_obj[0][1][1], lwh_point[2])
                            elif origin_bs == "X":
                                origin = (lwh_point[0], arc_obj[0][1][1], arc_obj[0][1][2])
                            if head_depth is None:
                                head_depth = lwh_point[2]
                            tool_body_list = self.perform_extrude(arc_obj[0][0], origin, arc_obj[0][3],
                                                                head_depth)
                            final_body_list = self.subtract([body], tool_body_list)
                            self.remove_parameters(final_body_list)

    def find_groove_obj(self, center_point_list):
        """获取所有槽"""
        curves = self.geometry_handler.get_all_curves()
        # 获取所有非闭合圆弧
        arc_list = []
        # 获取所有直线和非闭合圆
        arc_line_list = []
        # 存放封闭对象
        is_closed_list = []
        for curve in curves:
            if isinstance(curve, NXOpen.Line):
                arc_line_list.append(curve)
            elif isinstance(curve, NXOpen.Arc):
                arc = analyze_arc(curve)
                if not arc:
                    key = (curve.CenterPoint.X,curve.CenterPoint.Y,curve.CenterPoint.Z)
                    key2 = (round(curve.CenterPoint.X,3), round(curve.CenterPoint.Y,3), round(curve.CenterPoint.Z,3))
                    if key2 in center_point_list:
                        continue
                    arc_list.append(curve)
                    arc_line_list.append(curve)
        # 判断是否有和圆弧相连且组成封闭的区域
        for arc in arc_list:
                # 圆弧起点
                start_point = self.get_arc_point(arc)[0]
                p_list = self.find_connected_path(arc, start_point, arc_line_list)
                is_closed_list.append(p_list)
        return is_closed_list

    def find_connected_path(self, start_line, start_point, lines):
        """
        连通图搜索：从 start_line 出发，查找是否能通过连接线段回到 start_point。
        支持 NXOpen.Line 和 NXOpen.Arc。

        返回：按顺序的线段（或圆弧）列表，如果没有形成闭合，返回 None。
        """

        def to_tuple(pt):
            """统一坐标格式"""
            return (round(pt[0], 2), round(pt[1], 2), round(pt[2], 2))

        def endpoints(obj):
            """返回线段或圆弧的两个端点"""
            if isinstance(obj, NXOpen.Line):
                return to_tuple((obj.StartPoint.X,obj.StartPoint.Y,obj.StartPoint.Z)), to_tuple((obj.EndPoint.X,obj.EndPoint.Y,obj.EndPoint.Z))
            elif isinstance(obj, NXOpen.Arc):
                points = self.get_arc_point(obj)
                return to_tuple(points[0]), to_tuple(points[1])

        start_pt = to_tuple(start_point)

        visited = set()
        path = []

        def dfs(current_point):
            """DFS：从当前点出发，寻找闭合路径"""
            for geom in lines:
                if geom in visited:
                    continue

                p0, p1 = endpoints(geom)

                # 判断是否连接
                if current_point == p0:
                    next_point = p1
                elif current_point == p1:
                    next_point = p0
                else:
                    continue  # 不连通

                # 记录路径
                visited.add(geom)
                path.append(geom)

                # ★ 找到闭合点：next_point 回到起始点
                if next_point == start_pt:
                    return True

                # 继续搜索
                if dfs(next_point):
                    return True

                # 没走通 → 回溯
                path.pop()
                visited.remove(geom)

            return False

        # 起始线加入路径
        visited.add(start_line)
        path.append(start_line)

        # 起始线的两个端点
        p0, p1 = endpoints(start_line)

        # 从起始线的终点（你需求中的方向）开始搜索
        if dfs(p1):
            return path
        else:
            return None

    def safe_tag(self,obj):
        try:
            # 访问一个属性来确认对象有效
            _ = obj.JournalIdentifier
            return obj.Tag
        except:
            return None

    def unique_closed_paths(self,is_closed_list):

        unique = set()
        result = []

        for path in is_closed_list:
            if not path:
                continue

            # 过滤掉无效对象
            safe_tags = []
            for obj in path:
                tag = self.safe_tag(obj)
                if tag is not None:
                    safe_tags.append(tag)

            if not safe_tags:
                continue

            key = frozenset(safe_tags)

            if key not in unique:
                unique.add(key)
                result.append(path)

        return result

    def get_arc_point(self, curve):
        """
        获取圆弧的起点和终点
        """
        arc_data = self.uf_session.Curve.AskArcData(curve.Tag)

        # 获取 StartAngle 和 EndAngle
        start_angle = arc_data.StartAngle
        end_angle = arc_data.EndAngle

        # 起点
        sp = NXOpen.Point3d(
            arc_data.ArcCenter[0] + arc_data.Radius * math.cos(start_angle),
            arc_data.ArcCenter[1] + arc_data.Radius * math.sin(start_angle),
            arc_data.ArcCenter[2]
        )

        # 终点
        ep = NXOpen.Point3d(
            arc_data.ArcCenter[0] + arc_data.Radius * math.cos(end_angle),
            arc_data.ArcCenter[1] + arc_data.Radius * math.sin(end_angle),
            arc_data.ArcCenter[2]
        )

        return ((sp.X, sp.Y, sp.Z), (ep.X, ep.Y, ep.Z))


    def get_curve_display_info(self, curve):
        """
        获取曲线显示属性（兼容不同 NX 版本）
        可返回: color, font, layer, width（若存在）
        """
        try:
            tag = curve.Tag
            props = self.uf_session.Obj.AskDisplayProperties(tag)

            info = {}
            # 不同版本结构体字段名可能不同，因此逐一判断是否存在
            if hasattr(props, "Color"):
                info["color"] = props.Color
            if hasattr(props, "Font"):
                info["font"] = props.Font
            return info

        except Exception as e:
            print_to_info_window(f"⚠️ 无法获取曲线显示属性: {e}")
            return None

    def describe_font(self, font_id):
        """将 Font 编号转为线型文字描述"""
        font_map = {
            1: "实线(Solid)",
            2: "虚线(Dashed)",
            3: "点线(Dotted)",
            4: "点划线(Dash-dot)",
            5: "中心线(Centerline)",
        }
        return font_map.get(font_id, f"未知({font_id})")

    def _extract_and_classify_circles(self,lwh_point):
        """提取和分类图中的圆"""

        circle_obj_list = []  # 用于存放正面实线圆和背面实线圆
        side_circle_list = []  # 用于存放侧面实线圆
        x_circle_obj_list = [] # 用于存放正面虚线圆和背面虚线圆
        x_side_circle_list = [] # 用于存放侧面虚线圆
        all_circle_list = []  # 存放所有圆
        curves = self.geometry_handler.get_all_curves()
        seen = []
        for curve in curves:
            if not (hasattr(curve, "CenterPoint") and hasattr(curve, "Radius")):
                continue
            info = self.get_curve_display_info(curve)
            if not info:
                continue
            arc = analyze_arc(curve)
            if arc is not None:
                c = arc.CenterPoint
                key = (round(c.X, 3), round(c.Y, 3), round(c.Z, 3))
                # 获取实线圆
                if info.get("font", 0) == 1:
                    # 侧面圆
                    if round(c.X,1) == lwh_point[0] :
                        side_circle_list.append((arc, key, arc.Radius * 2, (-1.0, 0.0, 0.0)))
                        all_circle_list.append(key)
                    elif key[2] == 0.0:
                        circle_obj_list.append((arc, key, arc.Radius * 2, (0.0, 0.0, -1.0)))
                        all_circle_list.append(key)
                # 获取虚线圆
                elif info.get("font", 0) == 2:
                    # 侧面圆
                    if round(c.X,1) == lwh_point[0] :
                        # 复制虚线圆
                        to_point = (0.0, c.Y, c.Z)
                        new_obj = self.geometry_handler.move_object_point_to_point(object=arc, from_point=c, to_point = to_point, layer=10, is_copy=True)[0]
                        x_side_circle_list.append((new_obj, key, arc.Radius * 2, (1.0, 0.0, 0.0)))
                        all_circle_list.append(key)
                    elif key[2] == 0.0:
                        # 复制虚线圆
                        to_point = (c.X, c.Y, -lwh_point[2])
                        new_obj = self.geometry_handler.move_object_point_to_point(object=arc, from_point=c, to_point = to_point, layer=10, is_copy=True)[0]
                        x_circle_obj_list.append((new_obj, key, arc.Radius * 2, (0.0, 0.0, 1.0)))
                        all_circle_list.append(key)


        return {
            "front_circles": circle_obj_list,
            "side_circles": side_circle_list,
            "x_side_circles": x_side_circle_list,
            "x_circles": x_circle_obj_list,
            "center_point_list": all_circle_list
        }


    def perform_extrude(self, curve_list, origin, vector, distance):
        """
        对指定的曲线组执行一次拉伸，曲线组需要组成闭合轮廓。

        Args:
            curve_list (list[NXOpen.IBaseCurve]): 闭合轮廓的所有曲线
        """

        try:
            extrudeBuilder = self.work_part.Features.CreateExtrudeBuilder(NXOpen.Features.Feature.Null)
            if isinstance(curve_list, NXOpen.Arc):
                curve_list = [curve_list]
            section = self.work_part.Sections.CreateSection(0.0095, 0.01, 0.5)
            extrudeBuilder.Section = section
            extrudeBuilder.AllowSelfIntersectingSection(True)
            extrudeBuilder.DistanceTolerance = 0.01

            # 创建新体
            extrudeBuilder.BooleanOperation.Type = NXOpen.GeometricUtilities.BooleanOperation.BooleanType.Create

            # 拉伸距离
            extrudeBuilder.Limits.StartExtend.Value.SetFormula("-1")
            extrudeBuilder.Limits.EndExtend.Value.SetFormula(f"{distance+5}")

            # 方向
            origin = NXOpen.Point3d(*origin)
            vector = NXOpen.Vector3d(*vector)
            direction = self.work_part.Directions.CreateDirection(origin, vector,
                                                                  NXOpen.SmartObject.UpdateOption.WithinModeling)
            extrudeBuilder.Direction = direction

            # Section设置
            section.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.OnlyCurves)

            # 创建选择规则
            scRuleFactory = self.work_part.ScRuleFactory
            ruleOptions = scRuleFactory.CreateRuleOptions()

            # 添加所有曲线
            rules = []
            for curve in curve_list:
                rule = scRuleFactory.CreateRuleBaseCurveDumb([curve], ruleOptions)
                rules.append(rule)

            helpPoint = NXOpen.Point3d(0.0, 0.0, 0.0)

            # 将所有曲线加入截面
            section.AddToSection(
                rules,
                NXOpen.NXObject.Null,
                NXOpen.NXObject.Null,
                NXOpen.NXObject.Null,
                helpPoint,
                NXOpen.Section.Mode.Create,
                False
            )

            # 提交拉伸
            feature = extrudeBuilder.Commit()
            extrudeBuilder.Destroy()

            return list(feature.GetBodies())

        except NXOpen.NXException:
            return []


    def subtract(self, target_bodies, tool_bodies):
        """
        从目标体上减去工具体。
        (修正) 修正了 "移除参数" 后 feature 对象失效的问题。
        """
        print_to_info_window(target_bodies[0].Tag)
        if not tool_bodies or not target_bodies:
            return []

        # ... (创建和设置 builder 的代码保持不变) ...
        builder = self.work_part.Features.CreateBooleanBuilderUsingCollector(None)
        builder.Operation = NXOpen.Features.Feature.BooleanType.Subtract
        builder.Tolerance = 0.01
        sc_factory = self.work_part.ScRuleFactory
        rule_options = sc_factory.CreateRuleOptions()
        rule_options.SetSelectedFromInactive(False)
        target_collector = self.work_part.ScCollectors.CreateCollector()
        target_rule = sc_factory.CreateRuleBodyDumb(target_bodies, True, rule_options)
        target_collector.ReplaceRules([target_rule], False)
        builder.TargetBodyCollector = target_collector
        tool_collector = self.work_part.ScCollectors.CreateCollector()
        tool_rule = sc_factory.CreateRuleBodyDumb(tool_bodies, True, rule_options)
        tool_collector.ReplaceRules([tool_rule], False)
        builder.ToolBodyCollector = tool_collector

        feature = builder.Commit()
        builder.Destroy()
        rule_options.Dispose()

        # (核心修正)
        # 1. 在 feature 被销毁前，只获取一次实体列表
        result_bodies = list(feature.GetBodies())

        # 2. 对这些实体执行“移除参数”操作。此操作会销毁 feature 对象。
        if result_bodies:
            self.remove_parameters(result_bodies)

            # 3. 直接返回我们第一次获取的实体列表。
            #    这个列表中的实体引用在移除参数后仍然有效，它们现在指向的是无参数的实体。
            #    绝对不能再次调用 feature.GetBodies()。
            return result_bodies

        return []

    def find_body_by_point(self, point):
        tolerance = 0.01
        for body in self.get_all_bodies():
            bbox = self.uf_session.ModlGeneral.AskBoundingBox(body.Tag)
            if (bbox[0] - tolerance <= point[0] <= bbox[3] + tolerance and
                bbox[1] - tolerance <= point[1] <= bbox[4] + tolerance and
                bbox[2] - tolerance <= point[2] <= bbox[5] + tolerance):
                return body
        return None


    def get_all_bodies(self):
        """返回当前工作部件里所有实体（Solid Body）的列表，并打印到信息窗口"""

        if self.work_part is None:
            return []


        bodies = list(self.work_part.Bodies)  # BodyCollection -> List[Body]
        if not bodies:
            return []

        for i, body in enumerate(bodies, 1):
            try:
                jid = body.JournalIdentifier
                typ = body.GetSolidBodyType()  # 这是方法，不是属性
            except Exception as e:
                pass
        return bodies

    # 给对象列表去参
    def remove_parameters(self, body_list):
        """对给定的体对象列表执行“移除参数”操作。"""
        if not body_list:
            return
        try:
            builder = self.work_part.Features.CreateRemoveParametersBuilder()
            # 在某些版本中，Add可能只接受单个对象，使用循环更安全
            for body in body_list:
                builder.Objects.Add(body)
            builder.Commit()
            builder.Destroy()
        except NXOpen.NXException as e:
            pass

if __name__ == "__main__":
    session = NXOpen.Session.GetSession()
    work_part = session.Parts.Work
    uf_session = NXOpen.UF.UFSession.GetUFSession()
    # 创建工作流控制器
    workflow = MainWorkflow2(session, work_part, uf_session)
    workflow.run_workflow()