# -*- coding: utf-8 -*-
"""
主流程模块
包含完整钻孔自动化的主流程控制
"""

import math
from utils import print_to_info_window, handle_exception, point_from_angle, analyze_arc
from geometry import GeometryHandler
from path_optimization import PathOptimizer
from process_info import ProcessInfoHandler
from parameter_parser import ParameterParser
from drilling_operations import DrillingOperationHandler
from mirror_operations import MirrorHandler
from drill_library import DrillLibrary
import config
import NXOpen


class MainWorkflow:
    """主流程控制器"""

    def __init__(self, session, work_part):
        self.session = session
        self.work_part = work_part

        # 初始化各个处理器
        self.geometry_handler = GeometryHandler(session, work_part)
        self.path_optimizer = PathOptimizer()
        self.process_info_handler = ProcessInfoHandler()
        self.parameter_parser = ParameterParser()
        self.drilling_handler = DrillingOperationHandler(session, work_part)
        self.mirror_handler = MirrorHandler(session, work_part)
        self.drill_library = DrillLibrary()

    def run_workflow(self):
        """执行完整的工作流程"""

        try:
            print_to_info_window("=" * 60)
            print_to_info_window("开始执行NX钻孔自动化流程")
            print_to_info_window("=" * 60)

            # 第一步：预处理，提取加工信息
            print_to_info_window("第一步：预处理，提取加工信息")
            annotated_data = self.process_info_handler.extract_and_process_notes(self.work_part)
            processed_result = self.parameter_parser.process_hole_data(annotated_data)
            annotated_data = annotated_data["主体加工说明"]
            if not processed_result:
                return handle_exception("加工信息处理失败，流程终止")

            # 获取材料类型
            material_type = self.process_info_handler.get_material(annotated_data["材质"][0])
            if material_type is None:
                print_to_info_window("获取材质类型失败")
                return False
            # 获取毛坯尺寸
            lwh_point = self.process_info_handler.get_workpiece_dimensions(annotated_data["尺寸"])
            # 获取加工坐标原点
            original_point = self.geometry_handler.get_start_point(lwh_point)
            if original_point[0] != 0.0 and original_point[1] != 0.0:
                # 移动2D图到绝对坐标系
                self.geometry_handler.move_objects_point_to_point(original_point,config.DEFAULT_ORIGIN_POINT)
            # 适合窗口
            self.work_part.ModelingViews.WorkView.Fit()
            print_to_info_window("第二步：创建MCS坐标系及几何体")
            mcs = self.geometry_handler.create_mcs_with_safe_plane(origin_point=config.DEFAULT_ORIGIN_POINT)
            workpiece = self.geometry_handler.create_workpiece_geometry()
            # 旋转曲线和注释
            #self.geometry_handler.rotate_objects(config.DEFAULT_VECTOR,config.DEFAULT_ORIGIN_POINT,90)
            # 第三步：获取图中所有的完整圆
            print_to_info_window("第三步：获取图中所有的完整圆")
            # 判断侧面图在轴正方向还是负方向
            has_y_negative = self.mirror_handler.judge_side_negative(lwh_point)
            circle_groups = self._extract_and_classify_circles(lwh_point,has_y_negative)
            if not circle_groups or (not circle_groups["front_circles"] and not circle_groups["side_circles"]):
                return handle_exception("未找到任何有效圆孔，流程终止")
            # 输出找到的圆数量信息
            front_count = len(circle_groups["front_circles"])
            side_count = len(circle_groups["side_circles"])
            print_to_info_window(f"✅ 找到正面或背面加工圆: {front_count} 个，侧面加工圆: {side_count} 个")

            # 第四步：对圆进行分类并设置基准圆
            print_to_info_window("第四步：对圆进行分类并设置基准圆")
            base_circles = self._setup_base_circles(circle_groups["front_circles"],config.DEFAULT_ORIGIN_POINT)
            if base_circles:
                print_to_info_window(f"✅ 设置基准圆: {len(base_circles)} 个")
            else:
                print_to_info_window("⚠️ 未设置基准圆")

            # 第五步：根据加工说明分类正反面圆
            print_to_info_window("第五步：根据加工说明分类正反面圆")
            classified_circles = self._classify_circles_by_processing(
                circle_groups["front_circles"],
                processed_result,
                lwh_point,
                circle_groups["arc_list"]
            )
            # 根据加工说明分类侧面圆
            classified_side_circles = self._classify__side_circles_by_processing(
                circle_groups["side_circles"],
                processed_result,
                lwh_point
            )
            # for circle_group in classified_circles:
            #     for info in classified_circles[circle_group]:
            #         print_to_info_window(info)
            # for circle_group in classified_side_circles:
            #     for info in classified_circles[circle_group]:
            #         print_to_info_window(info)
            return True
            # 输出分类结果
            front_count = len(classified_circles["front_circles"])
            z_sink_count = len(classified_circles["z_sink_circles"])
            back_count = len(classified_circles["back_circles"])
            b_sink_count = len(classified_circles["b_sink_circles"])
            side_count = len(classified_side_circles["side_circles"])
            side_z_sink_count = len(classified_side_circles["side_z_sink_circles"])

            print_to_info_window(f"✅ 分类结果: 正面圆 {front_count} 个，正沉头 {z_sink_count} 个，背面圆 {back_count} 个，背沉头 {b_sink_count} 个，侧面圆 {side_count} 个，侧面正沉头 {side_z_sink_count} 个")

            print_to_info_window("第六步：镜像处理")
            # 第六步：镜像处理
            mirrored_curves = self.mirror_handler.mirror_curves()
            # 移动镜像对象至256图层
            markId1 = self.session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Move Layer")
            self.work_part.Layers.MoveDisplayableObjects(256, mirrored_curves)
            # 第七步：打点和钻孔
            print_to_info_window("第七步：打点和钻孔操作")
            self._perform_drilling(classified_circles, classified_side_circles, lwh_point, base_circles, material_type,
                                   mirrored_curves)

            print_to_info_window("=" * 60)
            print_to_info_window("NX钻孔自动化流程执行完成")
            print_to_info_window("=" * 60)

            return True

        except Exception as e:
            return handle_exception("主流程执行失败", str(e))

    def _extract_and_classify_circles(self, lwh_point,has_y_negative):
        """提取和分类图中的圆"""

        seen = set()  # 用于过滤圆心相同的圆
        circle_obj_list = []  # 用于存放正面加工圆和背面加工圆
        side_circle_list = []  # 用于存放侧面加工圆
        arc_list = []  # 用于存放非闭合圆弧
        curves = self.geometry_handler.get_all_curves()

        for curve in curves:
            if not (hasattr(curve, "CenterPoint") and hasattr(curve, "Radius")):
                continue

            arc_list.append((curve,(curve.CenterPoint.X,curve.CenterPoint.Y,curve.CenterPoint.Z), curve.Radius * 2))

            arc = analyze_arc(curve)
            if not arc:
                continue

            c = arc.CenterPoint
            key = (round(c.X, 4), round(c.Y, 4), round(c.Z, 4))
            if key in seen:
                continue

            seen.add(key)

            # 侧面图在y轴方向处理规则
            if has_y_negative[1] is not None:
                # 判断圆是否在边界内
                if abs(key[1]) < lwh_point[1]:
                    # 在边界内说明是正面加工圆或背面加工圆
                    circle_obj_list.append((arc, key, arc.Radius * 2))
                else:
                    # 侧面加工圆
                    side_circle_list.append((arc, key, arc.Radius * 2))
            # 侧面图在X轴方向处理规则
            elif has_y_negative[0] is not None:
                # 判断圆是否在边界内
                if abs(key[0]) < lwh_point[0]:
                    # 在边界内说明是正面加工圆或背面加工圆
                    circle_obj_list.append((arc, key, arc.Radius * 2))
                else:
                    # 侧面加工圆
                    side_circle_list.append((arc, key, arc.Radius * 2))
            else:
                circle_obj_list.append((arc, key, arc.Radius * 2))

        return {
            "front_circles": circle_obj_list,
            "side_circles": side_circle_list,
            "arc_list": arc_list
        }

    def _setup_base_circles(self, front_circles,original_point):
        """设置基准圆"""

        if not front_circles:
            return None

        nearest = self.geometry_handler.find_nearest_circle(front_circles, original_point)

        if not nearest:
            print_to_info_window("⚠️ 未找到最近圆")
            return None

        # 判断是否存在与之呈L型的圆
        base_circles = self._find_l_shape_circles(front_circles, nearest)

        # 如果不存在L型结构，创建新圆
        if len(base_circles) < 3:
            base_circles = self._create_base_circles(front_circles[0][1][0])

        return base_circles

    def _find_l_shape_circles(self, circles, nearest_circle):
        """查找L型结构的基准圆"""

        base_circles = []

        for circle in circles:
            if (round(circle[1][0], 1) == round(nearest_circle[1][0], 1) and
                    round(circle[1][1], 1) != round(nearest_circle[1][1], 1) and
                    round(nearest_circle[2], 1) <= 3.0):
                base_circles.append((nearest_circle[0], nearest_circle[1], nearest_circle[2], None, None, None))
                base_circles.append((circle[0], circle[1], circle[2], None, None, None))
            elif (round(circle[1][0], 1) != round(nearest_circle[1][0], 1) and
                  round(circle[1][1], 1) == round(nearest_circle[1][1], 1) and
                  round(nearest_circle[2], 1) <= 3.0):
                base_circles.append((circle[0], circle[1], circle[2], None, None, None))

        return base_circles

    def _create_base_circles(self, x_coord):
        """创建基准圆"""

        point = (0.0, 0.0, 0.0)
        points_list = []
        base_circles = []

        # 判断图在y轴左侧还是右侧
        if x_coord < 0.0:
            # 第一个圆
            new_center = point_from_angle(point, 135, 4)
            points_list.append(new_center)
            # 第二个圆
            new_center2 = (new_center[0] - 5.0, new_center[1], 0.0)
            points_list.append(new_center2)
            # 第三个圆
            new_center3 = (new_center[0], new_center[1] + 5.0, 0.0)
            points_list.append(new_center3)
            # 画圆
            arc_list = self.geometry_handler.create_circle_sketch(points_list, 3)[1]

            if arc_list:
                for arc in arc_list:
                    base_circles.append((arc, (arc.CenterPoint.X, arc.CenterPoint.Y, arc.CenterPoint.Z),
                                        arc.Radius * 2, None, None, None))
        else:
            # 第一个圆
            new_center = point_from_angle(point, 45, 4)
            points_list.append(new_center)
            # 第二个圆
            new_center2 = (new_center[0] + 5.0, new_center[1], 0.0)
            points_list.append(new_center2)
            # 第三个圆
            new_center3 = (new_center[0], new_center[1] + 5.0, 0.0)
            points_list.append(new_center3)
            # 画圆
            arc_list = self.geometry_handler.create_circle_sketch(points_list, 3)[1]

            if arc_list:
                for arc in arc_list:
                    base_circles.append((arc, (arc.CenterPoint.X, arc.CenterPoint.Y, arc.CenterPoint.Z),
                                        arc.Radius * 2, None, None, None))
        return base_circles

    def _classify_circles_by_processing(self, front_circles, processed_result, lwh_point, arc_list):
        """根据加工说明分类正反面圆"""

        print_to_info_window(f"开始分类圆: 正面或背面加工圆 {len(front_circles)} 个")

        # 根据注释对圆分类
        tag_circle_map = self.process_info_handler.fenlei(self.work_part, front_circles, lwh_point, arc_list)
        print_to_info_window(f"标签-圆映射: {len(tag_circle_map)} 个标签")
        print_to_info_window(f"加工结果: {len(processed_result)} 个加工说明")
        # 存放背面打点的圆对象
        back_circle_list = []
        # 存放正面打点的圆对象
        front_circle_list = []
        # 存放正沉头孔对象
        z_sink_circle_list = []
        # 存放背沉头对象
        b_sink_circle_list = []
        # 获取加工说明里的标签
        matched_tags = 0
        for tag in processed_result:
            tag_matched = False
            for tag_key in tag_circle_map.keys():
                if tag == tag_key:
                    tag_matched = True
                    matched_tags += 1
                    # 判断是正面打点还是背面打点
                    if processed_result[tag]["is_bz"] or processed_result[tag]["is_zzANDbz"]:
                        print_to_info_window(f"标签 {tag}: 背面钻孔")
                        for circle_info in tag_circle_map[tag_key]:
                            circle_obj = circle_info[0][0]
                            circle_center_point = circle_info[0][1]
                            real_diameter = (processed_result[tag]["real_diamater"]
                                             if processed_result[tag]["real_diamater"] is not None
                                             else circle_info[0][2])
                            # 获取关键字，用于后续设置参数的依据
                            main_key = processed_result[tag]["main_hole_processing"]
                            # 判断是否通孔
                            predefined_depth = (lwh_point[2] + 5.0
                                                if processed_result[tag]["is_through_hole"] and not processed_result[tag]["is_zzANDbz"]
                                                else processed_result[tag]["depth"]["hole_depth"]
                                                )
                            # 判断是否是穿线孔 穿线孔钻完后距离割边要有2mm余量
                            if (processed_result[tag]["main_hole_processing"] == "割" and processed_result[tag][
                                "real_diamater"] is not None):
                                real_diameter = real_diameter - 2
                            back_circle_list.append(
                                (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key))
                            # 处理背沉头
                            if (processed_result[tag]["circle_type"] == "背沉头"):
                                circle_obj = circle_info[0][0]
                                circle_center_point = circle_info[0][1]
                                real_diameter = processed_result[tag]["real_diamater_head"]
                                predefined_depth = processed_result[tag]["depth"]["head_depth"]
                                main_key = "铣"
                                b_sink_circle_list.append(
                                    (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth,
                                     main_key))
                    if (not processed_result[tag]["is_bz"]) or processed_result[tag]["is_zzANDbz"]:
                        print_to_info_window(f"标签 {tag}: 正面钻孔")
                        for circle_info in tag_circle_map[tag_key]:
                            circle_obj = circle_info[0][0]
                            circle_center_point = circle_info[0][1]
                            real_diameter = (processed_result[tag]["real_diamater"]
                                             if processed_result[tag]["real_diamater"] is not None
                                             else circle_info[0][2])
                            # 获取关键字，用于后续设置参数的依据
                            main_key = processed_result[tag]["main_hole_processing"]
                            # 判断是否通孔
                            predefined_depth = (lwh_point[2] + 5.0
                                                if processed_result[tag]["is_through_hole"] and not processed_result[tag]["is_zzANDbz"]
                                                else processed_result[tag]["depth"]["hole_depth"]
                                                )
                            # 判断是否是穿线孔 穿线孔钻完后距离割边要有2mm余量
                            if (processed_result[tag]["main_hole_processing"] == "割" and processed_result[tag][
                                "real_diamater"] is not None):
                                real_diameter = real_diameter - 2
                            front_circle_list.append(
                                (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key))
                            # 处理正沉头
                            if (processed_result[tag]["circle_type"] == "正沉头"):
                                circle_obj = circle_info[0][0]
                                circle_center_point = circle_info[0][1]
                                real_diameter = processed_result[tag]["real_diamater_head"]
                                predefined_depth = processed_result[tag]["depth"]["head_depth"]
                                main_key = "铣"
                                z_sink_circle_list.append(
                                    (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth,
                                     main_key))

            if not tag_matched:
                print_to_info_window(f"⚠️ 标签 {tag} 未匹配到任何圆")

        print_to_info_window(
            f"分类结果: 正面圆 {len(front_circle_list)} 个，背面圆 {len(back_circle_list)} 个，正沉头 {len(z_sink_circle_list)} 个，背沉头 {len(b_sink_circle_list)} 个")

        return {
            "front_circles": front_circle_list,
            "back_circles": back_circle_list,
            "z_sink_circles": z_sink_circle_list,
            "b_sink_circles": b_sink_circle_list
        }

    def _classify__side_circles_by_processing(self, side_circles, processed_result, lwh_point):
        """根据加工说明分类侧面圆"""

        print_to_info_window(f"开始分类圆: 侧面加工圆 {len(side_circles)} 个")

        # 根据注释对圆分类
        tag_circle_map = self.process_info_handler.fenlei(self.work_part, side_circles,None ,None)

        print_to_info_window(f"标签-圆映射: {len(tag_circle_map)} 个标签")
        print_to_info_window(f"加工结果: {len(processed_result)} 个加工说明")

        # 存放侧面打点的圆对象
        side_circle_list = []
        # 存放侧面正沉头孔对象
        side_z_sink_circle_list = []
        # 存放侧面背沉头孔对象
        side_b_sink_circle_list = []
        # 获取加工说明里的标签
        matched_tags = 0
        for tag in processed_result:
            tag_matched = False
            for tag_key in tag_circle_map.keys():
                if tag == tag_key:
                    tag_matched = True
                    matched_tags += 1
                    print_to_info_window(f"标签 {tag}: 侧面钻孔")
                    for circle_info in tag_circle_map[tag_key]:
                        # 圆对象
                        circle_obj = circle_info[0][0]
                        # 圆心坐标
                        circle_center_point = circle_info[0][1]
                        # 直径
                        real_diameter = (processed_result[tag]["real_diamater"]
                                         if processed_result[tag]["real_diamater"] is not None
                                         else circle_info[0][2])

                        # 获取关键字，用于后续设置参数的依据
                        main_key = processed_result[tag]["main_hole_processing"]
                        # 判断是否通孔
                        predefined_depth = (lwh_point[2] + 5.0
                                            if processed_result[tag]["is_through_hole"]
                                            else processed_result[tag]["depth"]["hole_depth"]
                                            )
                        # 判断是否是穿线孔 穿线孔钻完后距离割边要有2mm余量
                        if (processed_result[tag]["main_hole_processing"] == "割" and processed_result[tag][
                            "real_diamater"] is not None):
                            real_diameter = real_diameter - 2
                        # 添加到列表中
                        side_circle_list.append(
                            (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key))
                        # 处理正沉头
                        if (processed_result[tag]["circle_type"] == "正沉头"):
                            circle_obj = circle_info[0][0]
                            circle_center_point = circle_info[0][1]
                            real_diameter = processed_result[tag]["real_diamater_head"]
                            predefined_depth = processed_result[tag]["depth"]["head_depth"]
                            main_key = "铣"
                            side_z_sink_circle_list.append(
                                (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key))
                        elif (processed_result[tag]["circle_type"] == "背沉头"):
                            circle_obj = circle_info[0][0]
                            circle_center_point = circle_info[0][1]
                            real_diameter = processed_result[tag]["real_diamater_head"]
                            predefined_depth = processed_result[tag]["depth"]["head_depth"]
                            main_key = "铣"
                            side_b_sink_circle_list.append(
                                (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key))

            if not tag_matched:
                print_to_info_window(f"⚠️ 标签 {tag} 未匹配到任何圆")

        print_to_info_window(f"分类结果: 侧面圆 {len(side_circle_list)} 个，匹配标签 {matched_tags} 个")

        return {
            "side_circles": side_circle_list,
            "side_z_sink_circles": side_z_sink_circle_list,
            "side_b_sink_circles": side_b_sink_circle_list
        }

    def _perform_drilling(self, classified_circles, classified_side_circles, lwh_point, base_circles, material_type,
                          mirrored_curves):
        """执行镜像处理和钻孔操作"""

        # 正面圆处理
        if classified_circles["front_circles"]:
            print_to_info_window("开始正面钻孔操作...")
            self._perform_front_drilling(classified_circles["front_circles"], base_circles, material_type)
            print_to_info_window("✅ 正面钻孔操作完成")
            # 正沉头处理
            if classified_circles["z_sink_circles"]:
                print_to_info_window("开始正沉头铣孔操作...")
                self._perform_front_milling(classified_circles["z_sink_circles"], material_type, geometry_name=config.DEFAULT_MCS_NAME, parent_group_name="正面", group_name="正面铣孔")
                print_to_info_window("✅ 正沉头铣孔操作完成")
        else:
            print_to_info_window("⚠️ 没有正面圆需要钻孔")

        # 背面圆处理
        if classified_circles["back_circles"]:
            print_to_info_window("开始背面钻孔操作...")
            self._perform_back_drilling(classified_circles["back_circles"], base_circles, mirrored_curves,
                                        material_type)
            print_to_info_window("✅ 背面钻孔操作完成")
            # 背沉头处理
            if classified_circles["b_sink_circles"]:
                print_to_info_window("开始背沉头铣孔操作...")
                # 背面铣孔
                self._perform_back_milling(classified_circles["b_sink_circles"], material_type, mirrored_curves)
                print_to_info_window("✅ 背沉头铣孔操作完成")
        else:
            print_to_info_window("⚠️ 没有背面圆需要钻孔")

        # 侧面圆处理
        if classified_side_circles["side_circles"]:
            print_to_info_window("开始侧面钻孔操作...")
            self._perform_side_drilling(classified_side_circles["side_circles"], lwh_point, material_type)
            print_to_info_window("✅ 侧面钻孔操作完成")
            print_to_info_window("开始侧面正面铣孔操作")
            self._perform_side_milling(classified_side_circles["side_z_sink_circles"], material_type, mirrored_curves,False)
            print_to_info_window("侧面正面铣孔操作结束")
            print_to_info_window("开始侧面背面铣孔操作")
            self._perform_side_milling(classified_side_circles["side_b_sink_circles"], material_type, mirrored_curves,True)
            print_to_info_window("侧面背面铣孔操作结束")
        else:
            print_to_info_window("⚠️ 没有侧面圆需要钻孔")

    def _perform_front_drilling(self, front_circles, base_circles, material_type):
        """执行正面钻孔操作"""

        print_to_info_window(f"处理正面圆: {len(front_circles)} 个")
        print_to_info_window(f"基准圆: {len(base_circles) if base_circles else 0} 个")

        # 添加基准圆到正面打点列表
        complete_front_circles = front_circles + base_circles

        print_to_info_window(f"正面圆总数: {len(complete_front_circles)} 个")

        # 正面圆打点路径优化
        front_optimization = self.path_optimizer.optimize_drilling_path(complete_front_circles)

        # 创建正面打点工序
        if front_optimization:
            print_to_info_window(f"✅ 正面圆路径优化完成，共 {len(front_optimization)} 个孔")

            # 检查正面圆是否有效
            valid_holes = [hole for hole in front_optimization if hole is not None]
            if not valid_holes:
                print_to_info_window("⚠️ 正面圆优化结果为空，跳过正面钻孔")
                return
            # 创建程序组
            self.drilling_handler.get_or_create_program_group(group_name="正面")
            self.drilling_handler.create_drill_operation(
                operation_type="DRILLING",
                tool_type="CENTERDRILL",
                tool_name="zxz",
                geometry_name=config.DEFAULT_MCS_NAME,
                orient_geometry_name=config.DEFAULT_ORIENT_GEOMETRY_NAME,
                parent_group_name="正面",
                group_name="正面打点工序",
                method_group_name="METHOD",
                hole_features=front_optimization,
                predefined_depth=config.DEFAULT_DRILL_DEPTH,
                diameter=config.DEFAULT_DRILL_DIAMETER,
                tip_diameter=config.DEFAULT_TIP_DIAMETER,
                operation_name="z_zxz",
                drive_point=config.TOOL_DRIVE_POINT_TIP,
                bottom_offset=0.0,
                step_distance=config.DEFAULT_STEP_DISTANCE,
                cycle_type=config.CYCLE_DRILL_STANDARD
            )
            print_to_info_window("✅ 正面打点工序创建成功")
        else:
            print_to_info_window("⚠️ 正面圆路径优化失败，跳过正面打点")
        # 正面钻孔
        tag_list = [front_circle[3] for front_circle in front_circles]
        # 去重
        tag_list = list(set(tag_list))
        for tag in tag_list:
            if (tag is None):
                continue
            temp_list = []
            # 加工类型
            dtype = None
            # 孔直径
            diameter = None
            # 孔深度
            predefined_depth = None
            for front_circle in front_circles:
                if (tag is not None and tag == front_circle[3]):
                    dtype = front_circle[5]
                    diameter = front_circle[2]
                    predefined_depth = front_circle[4]
                    temp_list.append(front_circle)
            tool_name, step_distance, feed_rate, cycle_type = self.drill_library.get_drill_parameters(material_type,
                                                                                                      diameter, dtype)
            if (tool_name is None and step_distance is None and feed_rate is None and cycle_type is None):
                raise SystemExit("结束程序")
            if (len(temp_list) == 0):
                continue
            # 路径优化
            z_front_optimization = self.path_optimizer.optimize_drilling_path(temp_list)
            if (dtype != "铰"):
                self.drilling_handler.create_drill_operation(
                    operation_type="DRILLING",
                    tool_type="STD_DRILL",
                    tool_name=tool_name,
                    geometry_name=config.DEFAULT_MCS_NAME,
                    orient_geometry_name=config.DEFAULT_ORIENT_GEOMETRY_NAME,
                    parent_group_name="正面",
                    group_name="正面钻孔",
                    method_group_name="METHOD",
                    hole_features=z_front_optimization,
                    predefined_depth=float(predefined_depth),
                    diameter=float(diameter),
                    tip_diameter=config.DEFAULT_TIP_DIAMETER,
                    operation_name=f"Z_{tag}_{tool_name}",
                    drive_point=config.TOOL_DRIVE_POINT_TIP,
                    bottom_offset=0.0,
                    step_distance=float(step_distance),
                    feed_rate=float(feed_rate),
                    cycle_type=cycle_type
                )
            else:
                self.drilling_handler.create_drill_operation(
                    operation_type="DRILLING",
                    tool_type="REAMER",
                    tool_name=tool_name,
                    geometry_name=config.DEFAULT_MCS_NAME,
                    orient_geometry_name=config.DEFAULT_ORIENT_GEOMETRY_NAME,
                    parent_group_name="正面",
                    group_name="正面铰孔",
                    method_group_name="METHOD",
                    hole_features=z_front_optimization,
                    predefined_depth=float(predefined_depth)-5.0,
                    diameter=float(diameter)-0.2,
                    tip_diameter=config.DEFAULT_TIP_DIAMETER,
                    operation_name=f"Z_J_{tag}_{tool_name}",
                    drive_point=config.TOOL_DRIVE_POINT_TIP,
                    bottom_offset=0.0,
                    step_distance=float(step_distance),
                    feed_rate=float(feed_rate),
                    cycle_type=cycle_type
                )

    # 铣孔工序
    def _perform_front_milling(self, sink_circles, material_type, geometry_name,parent_group_name, group_name):
        """执行铣孔操作"""
        # 铣孔
        tag_list = [sink_circle[3] for sink_circle in sink_circles]
        # 去重
        tag_list = list(set(tag_list))
        for tag in tag_list:
            if (tag is None):
                continue
            temp_list = []
            # 孔直径
            diameter = None
            # 孔深度
            predefined_depth = None
            for sink_circle in sink_circles:
                if (tag is not None and tag == sink_circle[3]):
                    diameter,start_diameter,deviation = self._computer_mill_diameter(sink_circle[1], sink_circle[2])
                    predefined_depth = sink_circle[4]
                    temp_list.append(sink_circle)
            final_diameter, tool_name, R1, mill_depth = self.drill_library.get_mrill_parameters(material_type, diameter,deviation)
            if (tool_name is None and R1 is None and mill_depth is None):
                raise SystemExit("结束程序")
            if (len(temp_list) == 0):
                continue
            # 路径优化
            sink_optimization = self.path_optimizer.optimize_drilling_path(temp_list)
            self.drilling_handler.create_hole_milling_operation(
                tool_name=tool_name,
                geometry_name=geometry_name,
                orient_geometry_name=config.DEFAULT_ORIENT_GEOMETRY_NAME,
                parent_group_name=parent_group_name,
                group_name=group_name,
                method_group_name="METHOD",
                hole_features=sink_optimization,
                predefined_depth=float(predefined_depth),
                diameter=float(final_diameter),
                corner_radius=float(R1),
                axial_distance=float(mill_depth),
                operation_name=f"X_{tag}_{tool_name}"
            )

    def _perform_back_drilling(self, back_circles, base_circles, mirrored_curves, material_type):
        """执行背面钻孔操作"""

        if mirrored_curves:
            # 处理镜像后的圆
            final_back_circles = self._process_mirrored_circles(mirrored_curves, back_circles, base_circles)
            if final_back_circles:
                # 背面圆打点路径优化
                back_optimization = self.path_optimizer.optimize_drilling_path(final_back_circles)
                # 创建程序组
                self.drilling_handler.get_or_create_program_group(group_name="背面")
                # 创建背面打点工序
                if back_optimization:
                    self.drilling_handler.create_drill_operation(
                        operation_type="DRILLING",
                        tool_type="CENTERDRILL",
                        tool_name="zxz",
                        geometry_name=config.DEFAULT_MCS_NAME,
                        orient_geometry_name=config.DEFAULT_ORIENT_GEOMETRY_NAME,
                        parent_group_name="背面",
                        group_name="背面打点",
                        method_group_name="METHOD",
                        hole_features=back_optimization,
                        predefined_depth=config.DEFAULT_DRILL_DEPTH,
                        diameter=config.DEFAULT_DRILL_DIAMETER,
                        tip_diameter=config.DEFAULT_TIP_DIAMETER,
                        operation_name="b_zxz",
                        drive_point=config.TOOL_DRIVE_POINT_TIP,
                        bottom_offset=0.0,
                        step_distance=config.DEFAULT_STEP_DISTANCE,
                        cycle_type=config.CYCLE_DRILL_STANDARD
                    )
            # 背面钻孔
            tag_list = [final_back_circle[3] for final_back_circle in final_back_circles]
            # 去重
            tag_list = list(set(tag_list))
            for tag in tag_list:
                if (tag is None):
                    continue
                temp_list = []
                # 加工类型
                dtype = None
                # 孔直径
                diameter = None
                # 孔深度
                predefined_depth = None
                for final_back_circle in final_back_circles:
                    if (tag == final_back_circle[3]):
                        dtype = final_back_circle[5]
                        diameter = final_back_circle[2]
                        predefined_depth = final_back_circle[4]
                        temp_list.append(final_back_circle)

                tool_name, step_distance, feed_rate, cycle_type = self.drill_library.get_drill_parameters(material_type,
                                                                                                          diameter,
                                                                                                          dtype)
                if (tool_name is None and step_distance is None and feed_rate is None and cycle_type is None):
                    raise SystemExit("结束程序")
                if (len(temp_list) == 0):
                    continue
                # 路径优化
                b_front_optimization = self.path_optimizer.optimize_drilling_path(temp_list)
                # 钻孔
                if (dtype != "铰"):
                    self.drilling_handler.create_drill_operation(
                        operation_type="DRILLING",
                        tool_type="STD_DRILL",
                        tool_name=tool_name,
                        geometry_name=config.DEFAULT_MCS_NAME,
                        orient_geometry_name=config.DEFAULT_ORIENT_GEOMETRY_NAME,
                        parent_group_name="背面",
                        group_name="背面钻孔",
                        method_group_name="METHOD",
                        hole_features=b_front_optimization,
                        predefined_depth=float(predefined_depth),
                        diameter=float(diameter),
                        tip_diameter=config.DEFAULT_TIP_DIAMETER,
                        operation_name=f"B_{tag}_{tool_name}",
                        drive_point=config.TOOL_DRIVE_POINT_TIP,
                        bottom_offset=0.0,
                        step_distance=float(step_distance),
                        feed_rate=float(feed_rate),
                        cycle_type=cycle_type
                    )
                else:
                    self.drilling_handler.create_drill_operation(
                        operation_type="DRILLING",
                        tool_type="REAMER",
                        tool_name=tool_name,
                        geometry_name=config.DEFAULT_MCS_NAME,
                        orient_geometry_name=config.DEFAULT_ORIENT_GEOMETRY_NAME,
                        parent_group_name="背面",
                        group_name="背面铰孔",
                        method_group_name="METHOD",
                        hole_features=b_front_optimization,
                        predefined_depth=float(predefined_depth)-5.0,
                        diameter=float(diameter)-0.2,
                        tip_diameter=config.DEFAULT_TIP_DIAMETER,
                        operation_name=f"B_J_{tag}_{tool_name}",
                        drive_point=config.TOOL_DRIVE_POINT_TIP,
                        bottom_offset=0.0,
                        step_distance=float(step_distance),
                        feed_rate=float(feed_rate),
                        cycle_type=cycle_type
                    )

    # 背面铣孔
    def _perform_back_milling(self, b_sink_circles, material_type, mirrored_curves):
        """执行背面铣孔操作"""
        if mirrored_curves:
            # 处理镜像后的圆
            final_b_sink_circles = self._process_mirrored_circles(mirrored_curves, b_sink_circles, None)
            # 创建铣孔工序
            self._perform_front_milling(final_b_sink_circles, material_type, geometry_name=config.DEFAULT_MCS_NAME, parent_group_name="背面", group_name="背面铣孔")

    def _process_mirrored_circles(self, mirrored_curves, back_circles, base_circles):
        """处理镜像后的圆"""

        final_back_circles = []
        seen = set()
        for curve in mirrored_curves:
            if not (hasattr(curve, "CenterPoint") and hasattr(curve, "Radius")):
                continue

            # arc = analyze_arc(curve)
            # if not arc:
            #     continue
            key = (round(curve.CenterPoint.X, 2), round(curve.CenterPoint.Y, 2), round(curve.CenterPoint.Z, 2))
            if key in seen:
                continue
            seen.add(key)
            # 查找匹配的原始圆
            if (base_circles is not None):
                for back_circle in back_circles:
                    # 获取第一个镜像后的基准圆
                    if (key[0] == round(-base_circles[0][1][0], 2) and
                            key[1] == round(base_circles[0][1][1], 2) and
                            key[2] == round(base_circles[0][1][2], 2)):
                        final_back_circles.append((curve, (curve.CenterPoint.X, curve.CenterPoint.Y, curve.CenterPoint.Z),
                                                   base_circles[0][2], None, None, None))

                    elif (key[0] == round(-back_circle[1][0], 2) and
                          key[1] == round(back_circle[1][1], 2) and
                          key[2] == round(back_circle[1][2], 2)):
                        final_back_circles.append((curve, (curve.CenterPoint.X, curve.CenterPoint.Y, curve.CenterPoint.Z),
                                                   back_circle[2], back_circle[3], back_circle[4], back_circle[5]))
            else:
                for back_circle in back_circles:
                    if (key[0] == round(-back_circle[1][0], 2) and
                            key[1] == round(back_circle[1][1], 2) and
                            key[2] == round(back_circle[1][2], 2)):
                        final_back_circles.append((curve, (curve.CenterPoint.X, curve.CenterPoint.Y, curve.CenterPoint.Z),
                                                   back_circle[2], back_circle[3], back_circle[4], back_circle[5]))
        return final_back_circles

    def _perform_side_drilling(self, side_circles, lwh_point, material_type):
        # 获取侧面最近的边界线
        min_y = self.mirror_handler.select_boundary_curve(lwh_point)
        if min_y is not None:
            # 判断侧面图在y轴正方向还是负方向
            has_y_negative = self.mirror_handler.judge_y_negative()
            if has_y_negative:
                # 创建侧面MCS坐标系
                side_mcs = self.geometry_handler.create_mcs_with_safe_plane(
                    origin_point=(0.0, min_y - lwh_point[2], 0.0),
                    mcs_name=config.DEFAULT_SIDE_MCS_NAME,
                    safe_distance=config.DEFAULT_SAFE_DISTANCE
                )
            else:
                # 创建侧面MCS坐标系
                side_mcs = self.geometry_handler.create_mcs_with_safe_plane(
                    origin_point=(0.0, min_y, 0.0),
                    mcs_name=config.DEFAULT_SIDE_MCS_NAME,
                    safe_distance=config.DEFAULT_SAFE_DISTANCE
                )
            # 侧面圆打点路径优化
            side_optimization = self.path_optimizer.optimize_drilling_path(side_circles)
            # 创建侧面打点工序
            if side_optimization:
                # 创建程序组
                self.drilling_handler.get_or_create_program_group(group_name="侧面")
                self.drilling_handler.create_drill_operation(
                    operation_type="DRILLING",
                    tool_type="CENTERDRILL",
                    tool_name="zxz",
                    geometry_name=config.DEFAULT_SIDE_MCS_NAME,
                    orient_geometry_name=config.DEFAULT_ORIENT_GEOMETRY_NAME,
                    parent_group_name="侧面",
                    group_name="侧面打点",
                    method_group_name="METHOD",
                    hole_features=side_optimization,
                    predefined_depth=config.DEFAULT_DRILL_DEPTH,
                    diameter=config.DEFAULT_DRILL_DIAMETER,
                    tip_diameter=config.DEFAULT_TIP_DIAMETER,
                    operation_name="c_zxz",
                    drive_point=config.TOOL_DRIVE_POINT_TIP,
                    bottom_offset=0.0,
                    step_distance=config.DEFAULT_STEP_DISTANCE,
                    cycle_type=config.CYCLE_DRILL_STANDARD
                )
            # 侧面钻孔
            tag_list = [side_circle[3] for side_circle in side_circles]
            # 去重
            tag_list = list(set(tag_list))
            for tag in tag_list:
                if (tag is None):
                    continue
                temp_list = []
                # 加工类型
                dtype = None
                # 孔直径
                diameter = None
                # 孔深度
                predefined_depth = None
                for side_circle in side_circles:
                    if (tag is not None and tag == side_circle[3]):
                        dtype = side_circle[5]
                        diameter = side_circle[2]
                        predefined_depth = side_circle[4]
                        temp_list.append(side_circle)
                tool_name, step_distance, feed_rate, cycle_type = self.drill_library.get_drill_parameters(material_type,
                                                                                                          diameter,
                                                                                                          dtype)
                if (tool_name is None and step_distance is None and feed_rate is None and cycle_type is None):
                    raise SystemExit("结束程序")
                if (len(temp_list) == 0):
                    continue
                # 路径优化
                c_front_optimization = self.path_optimizer.optimize_drilling_path(temp_list)
                if (dtype != "铰"):
                    self.drilling_handler.create_drill_operation(
                        operation_type="DRILLING",
                        tool_type="STD_DRILL",
                        tool_name=tool_name,
                        geometry_name=config.DEFAULT_SIDE_MCS_NAME,
                        orient_geometry_name=config.DEFAULT_ORIENT_GEOMETRY_NAME,
                        parent_group_name="侧面",
                        group_name="侧面钻孔",
                        method_group_name="METHOD",
                        hole_features=c_front_optimization,
                        predefined_depth=float(predefined_depth),
                        diameter=float(diameter),
                        tip_diameter=config.DEFAULT_TIP_DIAMETER,
                        operation_name=f"C_{tag}_{tool_name}",
                        drive_point=config.TOOL_DRIVE_POINT_TIP,
                        bottom_offset=0.0,
                        step_distance=float(step_distance),
                        feed_rate=float(feed_rate),
                        cycle_type=cycle_type
                    )
                else:
                    self.drilling_handler.create_drill_operation(
                        operation_type="DRILLING",
                        tool_type="REAMER",
                        tool_name=tool_name,
                        geometry_name=config.DEFAULT_SIDE_MCS_NAME,
                        orient_geometry_name=config.DEFAULT_ORIENT_GEOMETRY_NAME,
                        parent_group_name="侧面",
                        group_name="侧面铰孔",
                        method_group_name="METHOD",
                        hole_features=c_front_optimization,
                        predefined_depth=float(predefined_depth)-5.0,
                        diameter=float(diameter)-0.2,
                        tip_diameter=config.DEFAULT_TIP_DIAMETER,
                        operation_name=f"J_{tag}_{tool_name}",
                        drive_point=config.TOOL_DRIVE_POINT_TIP,
                        bottom_offset=0.0,
                        step_distance=float(step_distance),
                        feed_rate=float(feed_rate),
                        cycle_type=cycle_type
                    )

    # 侧面铣孔
    def _perform_side_milling(self, side_circles, material_type, mirrored_curves, bs):
        if bs:
            # 处理镜像后的圆
            final_c_sink_circles = self._process_mirrored_circles(mirrored_curves, side_circles, None)
            # 创建铣孔工序
            self._perform_front_milling(final_c_sink_circles, material_type, geometry_name=config.DEFAULT_SIDE_MCS_NAME, parent_group_name="侧面", group_name="侧面背铣孔")
        else:
            self._perform_front_milling(side_circles, material_type, geometry_name=config.DEFAULT_SIDE_MCS_NAME, parent_group_name="侧面", group_name="侧面正铣孔")

    def _computer_mill_diameter(self, center_point, diameter):
        """沉头铣刀直径计算逻辑：取内圆半径的2/1+内外圆半径的差，得出的值可+-0.5mm"""
        curves = self.geometry_handler.get_all_curves()
        for curve in curves:
            if not (hasattr(curve, "CenterPoint") and hasattr(curve, "Radius")):
                continue

            arc = analyze_arc(curve)
            if not arc:
                continue

            c = arc.CenterPoint
            key = (round(c.X, 1), round(c.Y, 1), round(c.Z, 1))
            if (key[0] == round(center_point[0], 1) and key[1] == round(center_point[1], 1) and key[2] == round(
                    center_point[2], 1) and round(arc.Radius * 2, 1) != round(diameter, 1)):
                if (round(arc.Radius * 2, 1) > round(diameter, 1)):
                    start_diameter = round(diameter, 1)
                    deviation = (round(arc.Radius * 2, 1) - round(diameter, 1)) / 2
                    tool_diameter = round(abs(diameter / 4 + arc.Radius - diameter / 2), 1)
                else:
                    start_diameter = round(arc.Radius * 2, 1)
                    deviation =  (round(diameter, 1) - round(arc.Radius * 2, 1)) / 2
                    tool_diameter = round(abs(arc.Radius / 2 + diameter / 2 - arc.Radius), 1)
                return tool_diameter ,start_diameter ,deviation
