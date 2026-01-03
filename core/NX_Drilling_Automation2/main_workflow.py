# -*- coding: utf-8 -*-
"""
主流程模块
包含完整钻孔自动化的主流程控制
"""
import re
import ERROR
from log_record import ExceptionLogger
from utils import print_to_info_window, handle_exception, point_from_angle, analyze_arc, get_circle_params, safe_origin, \
    endpoints, get_arc_point, find_connected_path, create_point, vec_norm, vec_sub, EPS, point_in_polygon_robust, \
    find_one_valid_point, move_layer, remove_parameters, rotate_body, switch_to_manufacturing, delete_body, \
    is_mcs_exists
from geometry import GeometryHandler
from path_optimization import PathOptimizer
from process_info import ProcessInfoHandler
from parameter_parser import ParameterParser
from drilling_operations import DrillingOperationHandler
from mirror_operations import MirrorHandler
from drill_library import DrillLibrary
import drill_config
from reconstruct_threading_hole import RedCircleExtractor
import NXOpen


class MainWorkflow:
    """主流程控制器"""

    def __init__(self, session, work_part, file_name):
        self.session = session
        self.work_part = work_part
        file_name = str(file_name).split(".")[0]
        # 初始化日志管理器
        self.logger = ExceptionLogger(fr"core\NX_Drilling_Automation2\error_record_xlsx\{file_name}_error_record.xlsx")
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
            print_to_info_window("正在进行初始化，检查必要信息是否存在")
            # 切换到加工环境
            switch_to_manufacturing(self.session, self.work_part)
            # 第一步：预处理，提取加工信息
            print_to_info_window("第一步：预处理，提取加工信息")
            annotated_data = self.process_info_handler.extract_and_process_notes(self.work_part)
            processed_result, is_divided = self.parameter_parser.process_hole_data(annotated_data)
            annotated_data = annotated_data["主体加工说明"]
            if not processed_result:
                raise ERROR.InfoInvalidError()
            # 获取材料类型
            if not annotated_data["材质"]:
                raise ERROR.MaterialInvalidError()
            material_type = self.process_info_handler.get_material(annotated_data["材质"][0])

            # 获取毛坯尺寸
            if not annotated_data["尺寸"]:
                raise ERROR.SizeInvalidError()
            lwh_point = self.process_info_handler.get_workpiece_dimensions(annotated_data["尺寸"])

            # 获取加工坐标原点
            original_point = self.geometry_handler.get_start_point(lwh_point, "主视图")
            print_to_info_window(original_point)

            if original_point is None:
                raise ERROR.BoundLineInvalidError()

            print_to_info_window("初始化完成，进入钻孔流程")

        except ERROR.NXError as e:
            self.logger.log_exception(e)
            raise

        if abs(original_point[0]) > 1e-7 or abs(original_point[1]) > 1e-7:
            # 移动2D图到绝对坐标系
            try:
                self.geometry_handler.move_objects_point_to_point(original_point, drill_config.DEFAULT_ORIGIN_POINT)
            except:
                print_to_info_window("2D图已在加工坐标原点，无需移动")
        # 获取主视图四个边界点
        minx, miny, maxx, maxy = self.geometry_handler.main_view_bound_points(lwh_point)
        # 适合窗口
        self.work_part.ModelingViews.WorkView.Fit()
        # 获取主体
        main_body = list(self.work_part.Bodies)[0]
        print_to_info_window("第二步：创建MCS坐标系及几何体")
        # 中分
        if is_divided:
            # 主视图MCS
            mcs = self.geometry_handler.create_mcs_with_safe_plane(
                origin_point=(lwh_point[0] / 2, lwh_point[1] / 2, 0.0))
        else:
            mcs = self.geometry_handler.create_mcs_with_safe_plane(origin_point=drill_config.DEFAULT_ORIGIN_POINT)
            workpiece = self.geometry_handler.create_workpiece_geometry()
        # 旋转曲线和注释
        # self.geometry_handler.rotate_objects(drill_config.DEFAULT_VECTOR,drill_config.DEFAULT_ORIGIN_POINT,90)
        # 第三步：获取图中所有的完整圆
        print_to_info_window("第三步：获取图中所有的完整圆")
        circle_groups = self._extract_and_classify_circles(lwh_point, minx, miny, maxx, maxy)
        if not circle_groups or (not circle_groups["front_circles"] and not circle_groups["side_circles"]):
            return handle_exception("未找到任何有效圆孔，流程终止")
        # 输出找到的圆数量信息
        front_count = len(circle_groups["front_circles"])
        side_count = len(circle_groups["side_circles"])
        closed_bound_count = len(circle_groups["z_closed_bound_circle"])
        z_view_circle_count = len(circle_groups["z_view_circle"])
        print_to_info_window(
            f"✅ 找到正面或背面加工圆: {front_count} 个，正视图加工圆：{z_view_circle_count} 个，侧面加工圆: {side_count} 个,线割槽: {closed_bound_count} 个")

        # 判断是否存在正视图加工圆，有正视图加工圆镜像并旋转体
        if z_view_circle_count > 0:
            if main_body is not None:
                # 复制体
                body2 = \
                self.geometry_handler.move_bodies_point_to_point([main_body], (0.0, 0.0, 0.0), (0.0, 0.0, 0.0), 50,
                                                                 True)[0]
                remove_parameters(self.work_part, list(self.work_part.Bodies))
                # 旋转体
                body2 = rotate_body(self.work_part, body2, (0.0, 0.0, 0.0), (1.0, 0.0, 0.0), -90)  # 第三视角法，正视图在主视图下方
                remove_parameters(self.work_part, list(self.work_part.Bodies))
                # 移动距离
                p1, p2 = self.geometry_handler.compute_bound_line_distince(lwh_point)
                # 移动体
                body2 = self.geometry_handler.move_bodies_point_to_point([body2], (0.0, 0.0, 0.0), p2, 50, False)[0]
                # 将移动后的实体移动至50层
                move_layer(self.session, self.work_part, [body2], 50)
                print_to_info_window("已将正面体移至50层")
                # 创建正视图MCS坐标系
                front_original_point = self.geometry_handler.get_start_point(lwh_point, "正视图")
                mcs = self.geometry_handler.create_mcs_with_safe_plane(origin_point=front_original_point,
                                                                       mcs_name="FRONT_MCS")
        side_body = None
        right_original_point = None
        if side_count > 0:
            if main_body is not None:
                # 复制体
                body2 = \
                self.geometry_handler.move_bodies_point_to_point([main_body], (0.0, 0.0, 0.0), (0.0, 0.0, 0.0), 70,
                                                                 True)[0]
                remove_parameters(self.work_part, list(self.work_part.Bodies))
                # 旋转体
                body2 = rotate_body(self.work_part, body2, (lwh_point[0], 0.0, 0.0), (0.0, 1.0, 0.0),
                                    -90)  # 第三视角法，右视图在主视图右侧
                remove_parameters(self.work_part, list(self.work_part.Bodies))
                # 移动距离
                p1, p2 = self.geometry_handler.compute_bound_line_distince(lwh_point)
                print_to_info_window(p1)
                # 移动体
                body2 = \
                self.geometry_handler.move_bodies_point_to_point([body2], (lwh_point[0], 0.0, 0.0), p1, 50, False)[0]
                # 将移动后的实体移动至40层
                move_layer(self.session, self.work_part, [body2], 40)
                print_to_info_window("已将右面体移至40层")

                if body2:
                    side_body = body2
                # 创建右视图MCS坐标系
                right_original_point = self.geometry_handler.get_start_point(lwh_point, "右视图")
                mcs = self.geometry_handler.create_mcs_with_safe_plane(origin_point=right_original_point,
                                                                       mcs_name="RIGHT_MCS")
        # 存放线割槽圆对象
        real_cut_circle = []
        if closed_bound_count > 0:
            # 在线割槽内画圆
            print_to_info_window("开始在线割槽内画圆")
            for circle_info in circle_groups["z_closed_bound_circle"]:
                # 画圆
                arc = self.geometry_handler.create_circle_robust(circle_info[0], circle_info[1])
                real_cut_circle.append((arc, circle_info[0], circle_info[1]))

        else:
            print_to_info_window("无线割槽或线割槽内已有圆")

        # 第四步：对圆进行分类并设置基准圆
        print_to_info_window("第四步：对圆进行分类并设置基准圆")
        base_circles = self._setup_base_circles(circle_groups["front_circles"],
                                                drill_config.DEFAULT_ORIGIN_POINT)
        if base_circles:
            print_to_info_window(f"✅ 设置基准圆: {len(base_circles)} 个")
        else:
            print_to_info_window("⚠️ 未设置基准圆")

        # # 第4.5步：创建穿线孔
        # # 创建提取器并获取红色闭合圆
        extractor = RedCircleExtractor(self.session, self.work_part, target_color=186)
        red_circles = extractor.get_red_closed_circles()
        #
        # # 计算内圆参数（传入毛坯厚度）
        inner_circle_params = extractor.calculate_inner_circle_params_precise(lwh_point[2])
        threading_circles = []  # 画好的穿线圆
        threading_circles_groups = {
            "front_circles": [],
            "side_circles": [],
            "arc_list": [],
            "z_closed_bound_circle": [],
            "z_view_circle": [],
            "all_circles": [],
        }  # 预制穿线圆分组
        red_center_list = []  # 红色穿线圆圆心坐标
        if inner_circle_params:
            # 穿线圆
            threading_circles = []
            circle_params, red_center_list = get_circle_params(inner_circle_params)
            for circle_param in circle_params:
                threading_circles.append(
                    self.geometry_handler.create_circle_robust(circle_param[0], circle_param[1]))
            if threading_circles:
                threading_circles_groups = self._extract_and_classify_circles(lwh_point, minx, miny, maxx, maxy,
                                                                              threading_circles)

        # 第五步：根据加工说明分类正反面圆
        print_to_info_window("第五步：根据加工说明分类正反面圆")
        classified_circles = self._classify_circles_by_processing(
            circle_groups["front_circles"],
            processed_result,
            lwh_point,
            arc_list=circle_groups["arc_list"],
            all_circles=circle_groups["all_circles"]
        )
        if threading_circles_groups["front_circles"]:
            threading_classified_circles = self._classify_circles_by_processing(
                threading_circles_groups["front_circles"],
                processed_result,
                lwh_point,
                all_circles=circle_groups["all_circles"],
                threading=True,  # 穿线孔
                red_center_list=red_center_list

            )
            #  删除红色封闭圆，替换为上一步画好的穿线孔圆
            classified_circles = self.update_circles(threading_classified_circles, classified_circles,
                                                     red_circles)

        # 根据加工说明分类侧面圆
        classified_side_circles = self._classify__side_circles_by_processing(
            circle_groups["side_circles"],
            processed_result,
            lwh_point,
            all_circles=circle_groups["all_circles"]
        )
        if threading_circles_groups["side_circles"]:
            threading_classified_circles = self._classify__side_circles_by_processing(
                threading_circles_groups["side_circles"],
                processed_result,
                lwh_point,
                all_circles=circle_groups["all_circles"],
                threading=True,  # 穿线孔
                red_center_list=red_center_list

            )
            #  删除红色封闭圆，替换为上一步画好的穿线孔圆
            classified_side_circles = self.update_circles(threading_classified_circles, classified_side_circles,
                                                          red_circles)
        # 根据加工说明分类正视图圆
        classified_front_circles = self._classify__front_circles_by_processing(
            circle_groups["z_view_circle"],
            processed_result,
            lwh_point,
            all_circles=circle_groups["all_circles"]
        )
        if threading_circles_groups["z_view_circle"]:
            threading_classified_circles = self._classify__front_circles_by_processing(
                threading_circles_groups["z_view_circle"],
                processed_result,
                lwh_point,
                all_circles=circle_groups["all_circles"],
                threading=True,  # 穿线孔
                red_center_list=red_center_list

            )
            #  删除红色封闭圆，替换为上一步画好的穿线孔圆
            classified_side_circles = self.update_circles(threading_classified_circles,
                                                          classified_front_circles,
                                                          red_circles)
        # 输出分类结果
        front_count = len(classified_circles["front_circles"])
        z_sink_count = len(classified_circles["z_sink_circles"])
        back_count = len(classified_circles["back_circles"])
        b_sink_count = len(classified_circles["b_sink_circles"])
        side_z_count = len(classified_side_circles["side_z_circles"])
        side_b_count = len(classified_side_circles["side_b_circles"])
        side_z_sink_count = len(classified_side_circles["side_z_sink_circles"])
        side_b_sink_count = len(classified_side_circles["side_b_sink_circles"])
        front_view_circle_count = len(classified_front_circles["front_view_circle"])
        front_view_back_circle_count = len(classified_front_circles["front_view_back_circle"])
        z_view_sink_circle_count = len(classified_front_circles["z_view_sink_circle"])
        b_view_sink_circle_count = len(classified_front_circles["b_view_sink_circle"])
        print_to_info_window(
            f"✅ 分类结果: 正面圆 {front_count} 个，正沉头 {z_sink_count} 个，背面圆 {back_count} 个，背沉头 {b_sink_count} 个")
        print_to_info_window(
            f"✅ 分类结果: 正视图圆：{front_view_circle_count}个，正视图正沉头：{z_view_sink_circle_count}个，正视图背面圆：{front_view_back_circle_count}个，正视图背沉头：{b_view_sink_circle_count}个")

        print_to_info_window(
            f"✅ 分类结果: 侧面正面圆 {side_z_count} 个，侧面正沉头 {side_z_sink_count} 个，侧面背面圆 {side_b_count} 个，侧面背沉头 {side_b_sink_count} 个")
        if back_count > 0 or b_sink_count > 0:
            if is_divided:
                # 反面MCS
                b_mcs = self.geometry_handler.create_mcs_with_safe_plane(
                    origin_point=(-lwh_point[0] / 2, lwh_point[1] / 2, 0.0), mcs_name="B_MCS")
                workpiece = self.geometry_handler.create_workpiece_geometry()
            if main_body is not None:
                # 复制体
                body2 = \
                    self.geometry_handler.move_bodies_point_to_point([main_body], (0.0, 0.0, 0.0), (0.0, 0.0, 0.0), 70,
                                                                     True)[0]
                remove_parameters(self.work_part, list(self.work_part.Bodies))
                # 获取体的中心点
                body_center_point = (lwh_point[0] / 2, lwh_point[1] / 2, -lwh_point[2] / 2)
                # 旋转体
                body2 = rotate_body(self.work_part, body2, body_center_point, (0.0, 1.0, 0.0), 180)
                remove_parameters(self.work_part, list(self.work_part.Bodies))
                body3 = self.geometry_handler.move_bodies_point_to_point([body2], (0.0, 0.0, 0.0),
                                                                         (-lwh_point[0], 0.0, 0.0), 70, False)[0]
                # 将旋转后实体移动至70层
                move_layer(self.session, self.work_part, [body3], 70)
                print_to_info_window("已将反面体移至70层")

        # 判断是否有正面处理的孔，存在正面孔，线割槽在正面处理，否则背面处理
        if front_count > 0 and len(real_cut_circle) > 0:
            for cut_circle in real_cut_circle:
                if cut_circle[2] == 1.0:
                    classified_circles["front_circles"].append(
                        (cut_circle[0], cut_circle[1], cut_circle[2], "线割槽", 1.0, "槽", False))
                else:
                    classified_circles["front_circles"].append(
                        (cut_circle[0], cut_circle[1], cut_circle[2], "线割槽", lwh_point[2], "槽", True))
        elif back_count > 0 and closed_bound_count > 0:
            for cut_circle in real_cut_circle:
                if cut_circle[2] == 1.0:
                    classified_circles["back_circles"].append(
                        (cut_circle[0], cut_circle[1], cut_circle[2], "线割槽", 1.0, "槽", False))
                else:
                    classified_circles["back_circles"].append(
                        (cut_circle[0], cut_circle[1], cut_circle[2], "线割槽", lwh_point[2], "槽", True))

        print_to_info_window("第六步：镜像处理")
        # 第六步：镜像处理
        b_mirrored_curves = []
        side_mirrored_curves = []
        mirror_curves = self.mirror_handler.select_boundary_curve(lwh_point)
        if len(mirror_curves["main_mirror_curves"]) > 0 and back_count > 0:
            b_mirrored_curves = self.mirror_handler.mirror_curves("主视图", (0.0, 0.0, 0.0),
                                                                  mirror_curves["main_mirror_curves"])  # 主视图镜像
            # 主视图背面
            markId1 = self.session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Move Layer")
            self.work_part.Layers.MoveDisplayableObjects(70, b_mirrored_curves)
        if front_view_back_circle_count > 0 or b_view_sink_circle_count > 0:  # 正视图镜像
            front_mirrored_curves = self.mirror_handler.mirror_curves("正视图", front_original_point,
                                                                      mirror_curves[
                                                                          "front_mirror_curves"])  # 正视图镜像
            # 正视图视图背面
            markId1 = self.session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Move Layer")
            self.work_part.Layers.MoveDisplayableObjects(60, front_mirrored_curves)
        if side_b_sink_count + side_b_count > 0 and right_original_point:  # 侧视图镜像
            if side_body is not None:
                side_body_center_point = (
                    right_original_point[0] + lwh_point[2] / 2, right_original_point[1] + lwh_point[1] / 2,
                    -(right_original_point[2] + lwh_point[0] / 2))
                # 复制体
                body2 = \
                self.geometry_handler.move_bodies_point_to_point([side_body], (0.0, 0.0, 0.0), (0.0, 0.0, 0.0), 40,
                                                                 True)[0]
                remove_parameters(self.work_part, list(self.work_part.Bodies))
                # 旋转体
                body2 = rotate_body(self.work_part, body2, side_body_center_point, (0.0, 1.0, 0.0),
                                    180)  # 第三视角法，右视图在主视图右侧
                remove_parameters(self.work_part, list(self.work_part.Bodies))

                # 旋转体
                body2 = rotate_body(self.work_part, body2, side_body_center_point, (0.0, 0.0, 1.0),
                                    180)  # 第三视角法，右视图在主视图右侧
                remove_parameters(self.work_part, list(self.work_part.Bodies))

                # 移动距离
                tag_point = (side_body_center_point[0], -side_body_center_point[1], side_body_center_point[2])
                # 移动体
                body2 = \
                self.geometry_handler.move_bodies_point_to_point([body2], side_body_center_point, tag_point, 40, False)[
                    0]
                move_layer(self.session, self.work_part, [body2], 30)
                print_to_info_window("已将右面背面体移至30层")
            side_mirrored_curves = self.mirror_handler.mirror_curves("侧视图", right_original_point,
                                                                     mirror_curves[
                                                                         "right_mirror_curves"])  # 侧视图镜像
            # 侧视图背面
            markId1 = self.session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Move Layer")
            self.work_part.Layers.MoveDisplayableObjects(30, side_mirrored_curves)
        # 第七步：打点和钻孔
        print_to_info_window("第七步：打点和钻孔操作")
        self._perform_drilling(classified_circles, classified_side_circles, classified_front_circles, lwh_point,
                               base_circles, material_type,
                               b_mirrored_curves, side_mirrored_curves)

        # # 第八步：复制标签注释至镜像位置
        # # 遍历注释对象
        # """face_vector = 1  # 1=调整x坐标，0=调整y坐标"""
        # notes = list(self.work_part.Notes)
        # mirrored_notes = []
        #
        # # 预编译正则表达式，提升循环内执行效率
        # note_text_pattern = re.compile(r'^(?:[a-zA-Z]+|[a-zA-Z]+\d+[a-zA-Z]+|(?=.*[a-zA-Z])(?=.*\d)[a-zA-Z\d]+)$')
        #
        # for note in notes:
        #     # 获取注释文本并格式化
        #     texts = note.GetText()
        #     text_str = " ".join(texts).strip()
        #
        #     # 过滤符合正则规则的注释
        #     if note_text_pattern.fullmatch(text_str):
        #         # 获取注释原始坐标并四舍五入（避免重复计算）
        #         x, y, z = safe_origin(note)
        #         x_round = round(x, 4)
        #         y_round = round(y, 4)
        #         z_round = round(z, 4)
        #
        #         # 定义原始点
        #         from_point = NXOpen.Point3d(x_round, y_round, z_round)
        #
        #         # 根据face_vector决定坐标变化逻辑
        #         if face_vector == 1:
        #             # face_vector=1：x坐标取反（x - 2x = -x）
        #             to_x = -x_round
        #             to_y = y_round
        #         elif face_vector == 0:
        #             # face_vector=0：y坐标取反（y - 2y = -y）
        #             to_x = x_round
        #             to_y = -y_round
        #         else:
        #             # 异常值处理：默认不修改坐标
        #             to_x = x_round
        #             to_y = y_round
        #
        #         # 定义目标点
        #         to_point = NXOpen.Point3d(to_x, to_y, z_round)
        #
        #         # 移动注释并收集结果
        #         mirrored_note = self.geometry_handler.move_objects_point_to_point(
        #             from_point, to_point, 256, True, note
        #         )
        #         mirrored_notes += mirrored_note
        #
        # # 移动标签注释至256图层
        # markId1 = self.session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Move Layer")
        # self.work_part.Layers.MoveDisplayableObjects(256, mirrored_notes)
        # print_to_info_window("=" * 60)
        # print_to_info_window("NX钻孔自动化流程执行完成")
        # print_to_info_window("=" * 60)

        return True

    def update_circles(self, threading_classified_circles, classified_circles, red_circles):
        """
        过滤所有红色圆邻近区域的圆
        过滤规则：
            过滤掉：圆心x/y/z绝对值与任意红色圆圆心x/y/z差≤红色圆半径的圆（覆盖目标范围）；

        """
        # 预处理：提取红色圆的圆心
        red_centers = []
        if red_circles:
            for red_circle in red_circles:
                red_centers.append([red_circle[1], red_circle[2]])

        # 初始化结果字典
        processed_circles = {}
        for circle_type, threading_circles in threading_classified_circles.items():
            if threading_circles is None:
                continue

            original_circles = classified_circles.get(circle_type, [])
            filtered_circles = []

            # 遍历原始圆，执行简化过滤
            for circle in original_circles:
                cx, cy, cz = circle[1][0], circle[1][1], circle[1][2]  # 当前圆圆心
                need_filter = False

                # 核心规则：x/y/z绝对值差≤红色圆半径则过滤
                for (rx, ry, rz), r in red_centers:
                    # 计算绝对值差（匹配对称负坐标）
                    x_diff = abs(abs(cx) - abs(rx))
                    y_diff = abs(abs(cy) - abs(ry))
                    z_diff = abs(abs(cz) - abs(rz))
                    if x_diff <= r and y_diff <= r and z_diff <= r:  # 阈值r(红色圆半径)
                        need_filter = True
                        break

                # 不满足过滤条件则保留
                if not need_filter:
                    filtered_circles.append(circle)

            # 追加穿线圆
            filtered_circles += threading_circles
            processed_circles[circle_type] = filtered_circles

        return processed_circles

    def _extract_and_classify_circles(self, lwh_point, minx, miny, maxx, maxy, curves=None):
        """提取和分类图中的圆"""
        uf_session = NXOpen.UF.UFSession.GetUFSession()
        seen = set()  # 用于过滤圆心相同的圆
        circle_obj_list = []  # 用于存放正面加工圆和背面加工圆
        all_circles = []  # 存放所有封闭圆
        circle_obj_list_z = []  # 用于存放正视图
        side_circle_list = []  # 用于存放侧面加工圆
        arc_list = []  # 用于存放非闭合圆弧
        red_line_list = []  # 红色线
        red_arc_list = []  # 红色非闭合圆弧
        closed_bound_list = []  # 线割槽
        z_closed_bound_circle_list = []  # 线割槽中的封闭圆
        if not curves:
            curves = self.geometry_handler.get_all_curves()

        for curve in curves:
            # 获取线tag
            tag = curve.Tag
            # 获取线的属性
            props = uf_session.Obj.AskDisplayProperties(tag)
            # 提取红色的实线
            if props.Color == 186 and isinstance(curve, NXOpen.Line) and props.Font == 1:
                points = endpoints(curve)
                if abs(points[0][0]) < lwh_point[0] and abs(points[0][1]) < lwh_point[1]:
                    red_line_list.append(curve)
            if not (hasattr(curve, "CenterPoint") and hasattr(curve, "Radius")):
                continue
            arc_list.append((curve, (curve.CenterPoint.X, curve.CenterPoint.Y, curve.CenterPoint.Z), curve.Radius * 2))

            arc = analyze_arc(curve)
            if not arc:
                # 提取非闭合红色圆弧
                if props.Color == 186 and props.Font == 1:
                    points = endpoints(curve)
                    if abs(points[0][0]) < lwh_point[0] and abs(points[0][1]) < lwh_point[1]:
                        red_arc_list.append(curve)
                continue

            c = arc.CenterPoint
            key = (round(c.X, 4), round(c.Y, 4), round(c.Z, 4))
            all_circles.append((arc, key, 2 * arc.Radius))
            if key in seen:
                continue

            seen.add(key)
            # 判断圆心属于哪个视图
            three_view_bs = self.mirror_handler.three_view_area(arc.CenterPoint, minx, miny, maxx, maxy)
            if three_view_bs.get("主视图") is not None:
                # 在边界内说明是正面加工圆或背面加工圆
                circle_obj_list.append((arc, key, arc.Radius * 2))
            elif three_view_bs.get("正视图") is not None:
                circle_obj_list_z.append((arc, key, arc.Radius * 2))
            elif three_view_bs.get("侧视图") is not None:
                side_circle_list.append((arc, key, arc.Radius * 2))

        print_to_info_window("------------------- 开始处理线割槽和线割圆 -------------------")
        # 计算封闭区域
        for red_arc in red_arc_list:
            if red_arc in [item for sublist in closed_bound_list for item in sublist]:
                continue
            start_point = get_arc_point(red_arc)[0]
            path_list = find_connected_path(red_arc, start_point, red_line_list + red_arc_list)
            if path_list is not None:
                closed_bound_list.append(path_list)
        # 判断每个封闭区域内是否存在一个圆
        for closed_bound in closed_bound_list:
            segments = []
            for item in closed_bound:
                st = endpoints(item)[0]
                et = endpoints(item)[1]
                segments.append(((st[0], st[1]), (et[0], et[1])))
            # 预处理数据
            polygon_vertices = []
            seg_list = []
            for seg in segments:
                p1 = create_point(seg[0][0], seg[0][1])
                p2 = create_point(seg[1][0], seg[1][1])
                seg_list.append((p1, p2))
                if not polygon_vertices or vec_norm(vec_sub(polygon_vertices[-1], p1)) > EPS:
                    polygon_vertices.append(p1)
            if seg_list:
                last_p2 = seg_list[-1][1]
                if vec_norm(vec_sub(polygon_vertices[-1], last_p2)) > EPS:
                    polygon_vertices.append(last_p2)
            # 遍历圆弧列表
            exist_bs = None  # 存在标识
            for arc in arc_list:
                # 过滤掉组成封闭区域的圆弧
                if arc[0] in closed_bound:
                    continue
                arc_points = (arc[0].CenterPoint.X, arc[0].CenterPoint.Y)
                if point_in_polygon_robust(arc_points, polygon_vertices):
                    exist_bs = True
                    break
            if exist_bs == True and exist_bs is not None:
                continue
            # 计算封闭区域内的线割圆
            if lwh_point[2] >= 50:
                arc_center = [10.0, 7.0, 5.0, 3.0, 1.0]
            elif lwh_point[2] < 50 and lwh_point[2] >= 30:
                arc_center = [7.0, 5.0, 3.0, 1.0]
            elif lwh_point[2] < 30:
                arc_center = [5.0, 3.0, 1.0]
            for r in arc_center:
                r = r / 2
                circle_center_point = find_one_valid_point(closed_bound, r + 2)
                if circle_center_point is not None:
                    z_closed_bound_circle_list.append(((circle_center_point[0], circle_center_point[1], 0.0), r))
                    break

        return {
            "front_circles": circle_obj_list,
            "side_circles": side_circle_list,
            "arc_list": arc_list,
            "z_closed_bound_circle": z_closed_bound_circle_list,
            "z_view_circle": circle_obj_list_z,
            "all_circles": all_circles,
        }

    def _setup_base_circles(self, front_circles, original_point):
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
                base_circles.append((nearest_circle[0], nearest_circle[1], nearest_circle[2], None, None, None, False))
                base_circles.append((circle[0], circle[1], circle[2], None, None, None, False))
            elif (round(circle[1][0], 1) != round(nearest_circle[1][0], 1) and
                  round(circle[1][1], 1) == round(nearest_circle[1][1], 1) and
                  round(nearest_circle[2], 1) <= 3.0):
                base_circles.append((circle[0], circle[1], circle[2], None, None, None, False))

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
                                         arc.Radius * 2, None, None, None, False))
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
                                         arc.Radius * 2, None, None, None, False))
        return base_circles

    def _classify_circles_by_processing(self, front_circles, processed_result, lwh_point, arc_list=None,
                                        all_circles=None, threading=False, red_center_list=None):
        """根据加工说明分类正反面圆"""

        print_to_info_window(f"开始分类圆: 正面或背面加工圆 {len(front_circles)} 个")

        # 根据注释对圆分类
        tag_circle_map = self.process_info_handler.fenlei(self.work_part, front_circles, lwh_point, arc_list,
                                                          processed_result, all_circles, red_center_list)
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
                            is_through = processed_result[tag]["is_through_hole"]
                            predefined_depth = (lwh_point[2]
                                                if is_through and not processed_result[tag]["is_zzANDbz"]
                                                else processed_result[tag]["depth"]["hole_depth"]
                                                )
                            # 判断是否是穿线孔 穿线孔钻完后距离割边要有2mm余量
                            if (processed_result[tag]["main_hole_processing"] == "割" and processed_result[tag][
                                "real_diamater"] is not None):
                                real_diameter = real_diameter - 2
                            # 画的穿线圆处理
                            if threading:
                                real_diameter = circle_info[0][2]
                                predefined_depth = lwh_point[2]
                            back_circle_list.append(
                                (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key,
                                 is_through))

                            # 处理背沉头
                            if processed_result[tag]["circle_type"] == "背沉头":
                                circle_obj = circle_info[0][0]
                                circle_center_point = circle_info[0][1]
                                real_diameter = processed_result[tag]["real_diamater_head"]
                                predefined_depth = processed_result[tag]["depth"]["head_depth"]
                                main_key = "铣"
                                b_sink_circle_list.append(
                                    (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth,
                                     main_key, False))

                    if not processed_result[tag]["is_bz"] or processed_result[tag]["is_zzANDbz"]:
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
                            is_through = processed_result[tag]["is_through_hole"]
                            predefined_depth = (lwh_point[2]
                                                if processed_result[tag]["is_through_hole"] and not
                            processed_result[tag]["is_zzANDbz"]
                                                else processed_result[tag]["depth"]["hole_depth"]
                                                )
                            # 判断是否是穿线孔 穿线孔钻完后距离割边要有2mm余量
                            if (processed_result[tag]["main_hole_processing"] == "割" and processed_result[tag][
                                "real_diamater"] is not None):
                                real_diameter = real_diameter - 2

                            # 画的穿线圆处理
                            if threading:
                                real_diameter = circle_info[0][2]
                                predefined_depth = lwh_point[2]
                            front_circle_list.append(
                                (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key,
                                 is_through))

                            # 处理正沉头
                            if (processed_result[tag]["circle_type"] == "正沉头"):
                                circle_obj = circle_info[0][0]
                                circle_center_point = circle_info[0][1]
                                real_diameter = processed_result[tag]["real_diamater_head"]
                                predefined_depth = processed_result[tag]["depth"]["head_depth"]
                                main_key = "铣"
                                z_sink_circle_list.append(
                                    (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth,
                                     main_key, False))

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

    def _classify__side_circles_by_processing(self, side_circles, processed_result, lwh_point, arc_list=None,
                                              all_circles=None, threading=False, red_center_list=None):
        """根据加工说明分类侧面圆"""

        print_to_info_window(f"开始分类圆: 侧面加工圆 {len(side_circles)} 个")

        # 根据注释对圆分类
        tag_circle_map = self.process_info_handler.fenlei(self.work_part, side_circles, lwh_point, arc_list,
                                                          processed_result, all_circles, red_center_list)

        print_to_info_window(f"标签-圆映射: {len(tag_circle_map)} 个标签")
        print_to_info_window(f"加工结果: {len(processed_result)} 个加工说明")

        # 存放侧面正面打点的圆对象
        side_z_circle_list = []
        # 存放侧面背面打点的圆对象
        side_b_circle_list = []
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
                        is_through = processed_result[tag]["is_through_hole"]
                        hole_depth = processed_result[tag]["depth"]["hole_depth"]
                        predefined_depth = lwh_point[0] if is_through else hole_depth
                        # 判断是否是穿线孔 穿线孔钻完后距离割边要有2mm余量
                        if (processed_result[tag]["main_hole_processing"] == "割" and processed_result[tag][
                            "real_diamater"] is not None):
                            real_diameter = real_diameter - 2

                        # 画的穿线圆处理
                        if threading:
                            real_diameter = circle_info[0][2]
                            predefined_depth = lwh_point[0]
                        if processed_result[tag]["is_zzANDbz"] and not threading:
                            predefined_depth = hole_depth if hole_depth else lwh_point[0] / 2
                            side_z_circle_list.append(
                                (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key,
                                 is_through))
                            side_b_circle_list.append(
                                (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key,
                                 is_through))
                        elif processed_result[tag]["is_bz"]:
                            side_b_circle_list.append(
                                (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key,
                                 is_through))
                        else:
                            side_z_circle_list.append(
                                (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key,
                                 is_through))

                        # 处理正沉头
                        if (processed_result[tag]["circle_type"] == "正沉头"):
                            circle_obj = circle_info[0][0]
                            circle_center_point = circle_info[0][1]
                            real_diameter = processed_result[tag]["real_diamater_head"]
                            predefined_depth = processed_result[tag]["depth"]["head_depth"]
                            main_key = "铣"
                            side_z_sink_circle_list.append(
                                (
                                    circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key,
                                    False))
                            continue
                        if processed_result[tag]["circle_type"] == "背沉头":
                            circle_obj = circle_info[0][0]
                            circle_center_point = circle_info[0][1]
                            real_diameter = processed_result[tag]["real_diamater_head"]
                            predefined_depth = processed_result[tag]["depth"]["head_depth"]
                            main_key = "铣"
                            side_b_sink_circle_list.append(
                                (
                                    circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key,
                                    False))
                            continue

            if not tag_matched:
                print_to_info_window(f"⚠️ 标签 {tag} 未匹配到任何圆")

        print_to_info_window(
            f"分类结果: 侧面正面圆 {len(side_z_circle_list)} 个，侧面背面圆 {len(side_b_circle_list)} 个，匹配标签 {matched_tags} 个")

        return {
            "side_z_circles": side_z_circle_list,
            "side_b_circles": side_b_circle_list,
            "side_z_sink_circles": side_z_sink_circle_list,
            "side_b_sink_circles": side_b_sink_circle_list
        }

    def _perform_drilling(self, classified_circles, classified_side_circles, classified_front_circles, lwh_point,
                          base_circles, material_type,
                          mirrored_curves, side_mirrored_curves):
        """执行镜像处理和钻孔操作"""

        # 正面圆处理
        if classified_circles["front_circles"]:
            print_to_info_window("开始正面钻孔操作...")
            temp_list_mill = self._perform_front_drilling(classified_circles["front_circles"], base_circles,
                                                          material_type)
            print_to_info_window("✅ 正面钻孔操作完成")
            # 一般孔铣孔
            if temp_list_mill:
                print_to_info_window("开始一般孔铣孔操作...")
                self._perform_front_milling(temp_list_mill, material_type, geometry_name=drill_config.DEFAULT_MCS_NAME,
                                            parent_group_name="正面", group_name="正面铣孔")
                print_to_info_window("✅ 一般孔铣孔操作完成")

            # 正沉头处理
            if classified_circles["z_sink_circles"]:
                print_to_info_window("开始正沉头铣孔操作...")
                self._perform_front_milling(classified_circles["z_sink_circles"], material_type,
                                            geometry_name=drill_config.DEFAULT_MCS_NAME, parent_group_name="正面",
                                            group_name="正面铣孔")
                print_to_info_window("✅ 正沉头铣孔操作完成")
        else:
            print_to_info_window("⚠️ 没有正面圆需要钻孔")

        # 背面圆处理
        if classified_circles["back_circles"]:
            print_to_info_window("开始背面钻孔操作...")

            temp_list_mill = self._perform_back_drilling(classified_circles["back_circles"], base_circles,
                                                         mirrored_curves,
                                                         material_type)

            print_to_info_window("✅ 背面钻孔操作完成")
            # 一般孔铣孔
            if temp_list_mill:
                print_to_info_window("开始一般孔背面铣孔操作...")
                self._perform_front_milling(temp_list_mill, material_type, geometry_name=drill_config.DEFAULT_MCS_NAME,
                                            parent_group_name="背面", group_name="背面铣孔")
                print_to_info_window("✅ 一般孔背面铣孔操作完成")

            # 背沉头处理
            if classified_circles["b_sink_circles"]:
                print_to_info_window("开始背沉头铣孔操作...")
                # 背面铣孔
                self._perform_back_milling(classified_circles["b_sink_circles"], material_type, mirrored_curves)
                print_to_info_window("✅ 背沉头铣孔操作完成")
        else:
            print_to_info_window("⚠️ 没有背面圆需要钻孔")

        # 侧面正面圆处理
        if classified_side_circles["side_z_circles"]:
            print_to_info_window("开始侧面正面钻孔操作...")
            temp_list_mill = self._perform_side_drilling(classified_side_circles["side_z_circles"],
                                                         material_type, base_circles)
            print_to_info_window("✅ 侧面正面钻孔操作完成")

            # 一般孔铣孔
            if temp_list_mill:
                print_to_info_window("开始侧面正面一般孔铣孔操作...")
                self._perform_side_milling(temp_list_mill, material_type, mirrored_curves, False)
                print_to_info_window("✅ 侧面正面一般孔铣孔操作完成")

            print_to_info_window("开始侧面正面沉头孔铣孔操作")
            self._perform_side_milling(classified_side_circles["side_z_sink_circles"], material_type, mirrored_curves,
                                       False)
            print_to_info_window("侧面正面沉头孔铣孔操作结束")

        # 侧面背面圆处理
        if classified_side_circles["side_b_circles"]:
            print_to_info_window("开始侧面背面钻孔操作...")
            temp_list_mill = self._perform_side_drilling(classified_side_circles["side_b_circles"],
                                                         material_type, base_circles, side_mirrored_curves, "c_b")
            print_to_info_window("✅ 侧面背面钻孔操作完成")

            # 一般孔铣孔
            if temp_list_mill:
                print_to_info_window("开始侧面背面一般孔铣孔操作...")
                self._perform_side_milling(temp_list_mill, material_type, mirrored_curves, False)
                print_to_info_window("✅ 侧面背面一般孔铣孔操作完成")
            print_to_info_window("开始侧面背面沉头孔铣孔操作")
            self._perform_side_milling(classified_side_circles["side_b_sink_circles"], material_type, mirrored_curves,
                                       True)
            print_to_info_window("侧面背面沉头孔铣孔操作结束")
        else:
            print_to_info_window("⚠️ 没有侧面圆需要钻孔")

        # 正视图圆处理
        if classified_front_circles["front_view_circle"]:
            print_to_info_window("开始正视图钻孔操作...")
            temp_list_mill = self._perform_front_view_drilling(classified_front_circles["front_view_circle"],
                                                               material_type)
            print_to_info_window("✅ 正视图钻孔操作完成")

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
                geometry_name=drill_config.DEFAULT_MCS_NAME,
                orient_geometry_name=drill_config.DEFAULT_MCS_NAME,
                parent_group_name="正面",
                group_name="正面打点工序",
                method_group_name="METHOD",
                hole_features=front_optimization,
                predefined_depth=drill_config.DEFAULT_DRILL_DEPTH,
                diameter=drill_config.DEFAULT_DRILL_DIAMETER,
                tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                operation_name="z_zxz",
                drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                is_through=False,
                step_distance=drill_config.DEFAULT_STEP_DISTANCE,
                cycle_type=drill_config.CYCLE_DRILL_STANDARD
            )
            print_to_info_window("✅ 正面打点工序创建成功")
        else:
            print_to_info_window("⚠️ 正面圆路径优化失败，跳过正面打点")
        # 正面钻孔
        tag_list = [front_circle[3] for front_circle in front_circles]
        # 去重
        tag_list = list(set(tag_list))
        temp_list_mill = []  # 一般孔铣孔列表
        for tag in tag_list:
            if (tag is None):
                continue
            temp_list = []
            temp_list_hinge = []  # 铰之前先钻，单独存放铰孔的钻孔参数
            # 加工类型
            dtype = None
            # 孔直径
            diameter = None
            # 孔深度
            predefined_depth = None
            have_mill = ""
            # 是否通孔
            is_through = False
            for front_circle in front_circles:
                if (tag is not None and tag == front_circle[3]):
                    if "一般孔" in front_circle[5]:  # 一般孔铣或一般孔精铣
                        dtype = "钻"
                        diameter = front_circle[2] - 0.5 if "精铣" in front_circle[5] else front_circle[2]
                        predefined_depth = front_circle[4]
                        if predefined_depth > 2 * diameter:
                            temp_list.append(front_circle)
                        temp_list_mill.append(front_circle)
                        have_mill = "精铣" if "精铣" in front_circle[5] else "铣"
                        continue
                    # 铰孔拆分为钻与铰，通过铰的参数获取钻的参数，这个钻单独处理
                    if front_circle[5] == "铰":
                        dtype = "铰"

                        if not front_circle[4] or not front_circle[2]:
                            continue
                        diameter = front_circle[2] - 0.2  # 钻直径
                        predefined_depth = front_circle[4] - (diameter * 0.7)  # 铰深
                        # 钻参数
                        temp_list_hinge.append((front_circle[0], front_circle[1], diameter, front_circle[3],
                                                front_circle[4] + diameter * 0.7, "钻", front_circle[6]))
                        # 铰参数
                        temp_list.append((front_circle[0], front_circle[1], front_circle[2], front_circle[3],
                                          predefined_depth, front_circle[5], False))
                        diameter = front_circle[2]  # 铰直径
                        continue
                    is_through = front_circle[6]
                    dtype = front_circle[5]
                    diameter = front_circle[2]
                    predefined_depth = front_circle[4]
                    temp_list.append(front_circle)

            # 创建铰孔工序之前的钻孔工序
            if temp_list_hinge:
                if not predefined_depth:
                    continue
                tool_name, step_distance, feed_rate, cycle_type = self.drill_library.get_drill_parameters(material_type,
                                                                                                          temp_list_hinge[
                                                                                                              0][2],
                                                                                                          temp_list_hinge[
                                                                                                              0][-1])
                if (tool_name is None and step_distance is None and feed_rate is None and cycle_type is None):
                    continue
                if (len(temp_list) == 0):
                    continue

                # 路径优化
                z_front_optimization = self.path_optimizer.optimize_drilling_path(temp_list)
                self.drilling_handler.create_drill_operation(
                    operation_type="DRILLING",
                    tool_type="STD_DRILL",
                    tool_name=tool_name,
                    geometry_name=drill_config.DEFAULT_MCS_NAME,
                    orient_geometry_name=drill_config.DEFAULT_MCS_NAME,
                    parent_group_name="正面",
                    group_name="正面钻孔",
                    method_group_name="METHOD",
                    hole_features=z_front_optimization,
                    predefined_depth=float(temp_list_hinge[0][4]),
                    diameter=float(temp_list_hinge[0][2]),
                    tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                    operation_name=f"Z_{tag}_{tool_name}",
                    drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                    is_through=is_through,
                    step_distance=float(step_distance),
                    feed_rate=float(feed_rate),
                    cycle_type=cycle_type
                )
            if not predefined_depth:
                continue
            if diameter >= 42:
                self.special_hole_handle(tag, temp_list, predefined_depth, dtype, material_type,
                                         drill_config.DEFAULT_MCS_NAME, "正面", "Z", is_through)
                continue

            tool_name, step_distance, feed_rate, cycle_type = self.drill_library.get_drill_parameters(material_type,
                                                                                                      diameter,
                                                                                                      dtype,
                                                                                                      have_mill)
            if (tool_name is None and step_distance is None and feed_rate is None and cycle_type is None):
                continue
            if (len(temp_list) == 0):
                continue

            # 路径优化
            z_front_optimization = self.path_optimizer.optimize_drilling_path(temp_list)
            if (dtype != "铰"):
                self.drilling_handler.create_drill_operation(
                    operation_type="DRILLING",
                    tool_type="STD_DRILL",
                    tool_name=tool_name,
                    geometry_name=drill_config.DEFAULT_MCS_NAME,
                    orient_geometry_name=drill_config.DEFAULT_MCS_NAME,
                    parent_group_name="正面",
                    group_name="正面钻孔",
                    method_group_name="METHOD",
                    hole_features=z_front_optimization,
                    predefined_depth=float(predefined_depth),
                    diameter=float(diameter),
                    tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                    operation_name=f"Z_{tag}_{tool_name}",
                    drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                    is_through=is_through,
                    step_distance=float(step_distance),
                    feed_rate=float(feed_rate),
                    cycle_type=cycle_type
                )
            else:
                self.drilling_handler.create_drill_operation(
                    operation_type="DRILLING",
                    tool_type="REAMER",
                    tool_name=tool_name,
                    geometry_name=drill_config.DEFAULT_MCS_NAME,
                    orient_geometry_name=drill_config.DEFAULT_MCS_NAME,
                    parent_group_name="正面",
                    group_name="正面铰孔",
                    method_group_name="METHOD",
                    hole_features=z_front_optimization,
                    predefined_depth=float(predefined_depth),
                    diameter=float(diameter),
                    tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                    operation_name=f"Z_J_{tag}_{tool_name}",
                    drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                    is_through=False,
                    step_distance=float(step_distance),
                    feed_rate=float(feed_rate),
                    cycle_type=cycle_type
                )
        return temp_list_mill

    # 处理特殊孔
    def special_hole_handle(self, tag, hole_list, predefined_depth, dtype, material_type, geometry_name,
                            parent_group_name, bs, is_through):
        diameter = 42.0
        tool_name, step_distance, feed_rate, cycle_type = self.drill_library.get_drill_parameters(material_type,
                                                                                                  diameter, dtype)
        if (tool_name is None and step_distance is None and feed_rate is None and cycle_type is None):
            return
        if (len(hole_list) == 0):
            return
        # 路径优化
        z_front_optimization = self.path_optimizer.optimize_drilling_path(hole_list)
        if (dtype != "铰"):
            self.drilling_handler.create_drill_operation(
                operation_type="DRILLING",
                tool_type="STD_DRILL",
                tool_name=tool_name,
                geometry_name=geometry_name,
                orient_geometry_name=geometry_name,
                parent_group_name=parent_group_name,
                group_name=f"{parent_group_name}钻孔",
                method_group_name="METHOD",
                hole_features=z_front_optimization,
                predefined_depth=float(predefined_depth),
                diameter=float(diameter),
                tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                operation_name=f"{bs}_{tag}_{tool_name}",
                drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                is_through=is_through,
                step_distance=float(step_distance),
                feed_rate=float(feed_rate),
                cycle_type=cycle_type
            )
        # 铣孔
        final_diameter, tool_name, R1, mill_depth = self.drill_library.get_mrill_parameters(material_type, diameter,
                                                                                            None)
        if (tool_name is None and R1 is None and mill_depth is None):
            return
        # 路径优化
        sink_optimization = self.path_optimizer.optimize_drilling_path(hole_list)
        self.drilling_handler.create_hole_milling_operation(
            tool_name=tool_name,
            geometry_name=geometry_name,
            orient_geometry_name=geometry_name,
            parent_group_name=parent_group_name,
            group_name=f"{parent_group_name}铣孔",
            method_group_name="METHOD",
            hole_features=sink_optimization,
            predefined_depth=float(predefined_depth),
            diameter=float(final_diameter),
            corner_radius=float(R1),
            axial_distance=float(mill_depth),
            operation_name=f"X_{tag}_{tool_name}"
        )

    # 铣孔工序
    def _perform_front_milling(self, sink_circles, material_type, geometry_name, parent_group_name, group_name,
                               direction=""):
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
            processing_method = None  #
            for sink_circle in sink_circles:
                if (tag is not None and tag == sink_circle[3]):
                    # diameter,start_diameter,deviation = self._computer_mill_diameter(sink_circle[1], sink_circle[2])
                    if "一般孔" in sink_circle[5]:
                        processing_method = "精铣" if "精铣" in sink_circle[5] else "铣"
                        diameter = sink_circle[2]
                        predefined_depth = sink_circle[4]
                        temp_list.append(sink_circle)
                        continue
                    diameter = sink_circle[2]
                    predefined_depth = sink_circle[4]
                    temp_list.append(sink_circle)

            if not predefined_depth:
                continue
            final_diameter, tool_name, R1, mill_depth = self.drill_library.get_mrill_parameters(material_type, diameter,
                                                                                                None, processing_method)
            if (tool_name is None and R1 is None and mill_depth is None):
                continue
            if (len(temp_list) == 0):
                continue

            # 路径优化
            sink_optimization = self.path_optimizer.optimize_drilling_path(temp_list)
            self.drilling_handler.create_hole_milling_operation(
                tool_name=tool_name,
                geometry_name=geometry_name,
                orient_geometry_name=geometry_name,
                parent_group_name=parent_group_name,
                group_name=group_name,
                method_group_name="METHOD",
                hole_features=sink_optimization,
                predefined_depth=float(predefined_depth),
                diameter=float(final_diameter),
                corner_radius=float(R1),
                axial_distance=float(mill_depth),
                operation_name=f"{direction}X_{tag}_{tool_name}"
            )

    def _perform_back_drilling(self, back_circles, base_circles, mirrored_curves, material_type):
        """执行背面钻孔操作"""
        bool = is_mcs_exists(self.work_part, "B_MCS")
        if bool:
            b_mcs = drill_config.DEFAULT_B_MCS_NAME
        else:
            b_mcs = drill_config.DEFAULT_MCS_NAME
        if len(mirrored_curves) > 0:
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
                        geometry_name=b_mcs,
                        orient_geometry_name=b_mcs,
                        parent_group_name="背面",
                        group_name="背面打点",
                        method_group_name="METHOD",
                        hole_features=back_optimization,
                        predefined_depth=drill_config.DEFAULT_DRILL_DEPTH,
                        diameter=drill_config.DEFAULT_DRILL_DIAMETER,
                        tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                        operation_name=f"b_zxz",
                        drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                        is_through=False,
                        step_distance=drill_config.DEFAULT_STEP_DISTANCE,
                        cycle_type=drill_config.CYCLE_DRILL_STANDARD
                    )
            # 背面钻孔
            tag_list = [final_back_circle[3] for final_back_circle in final_back_circles]
            # 去重
            tag_list = list(set(tag_list))
            temp_list_mill = []  # 一般孔铣孔列表
            for tag in tag_list:
                if (tag is None):
                    continue
                temp_list = []
                temp_list_hinge = []  # 铰孔前的钻孔列表
                # 加工类型
                dtype = None
                # 孔直径
                diameter = None
                # 孔深度
                predefined_depth = None
                have_mill = ""
                # 是否通孔
                is_through = False
                for final_back_circle in final_back_circles:
                    if (tag == final_back_circle[3]):
                        if "一般孔" in final_back_circle[5]:  # 一般孔铣或一般精铣
                            dtype = "钻"
                            diameter = final_back_circle[2] - 0.5 if "精铣" in final_back_circle else final_back_circle[
                                2]
                            predefined_depth = final_back_circle[4]
                            if predefined_depth > 2 * diameter:
                                temp_list.append(final_back_circle)
                            temp_list_mill.append(final_back_circle)
                            have_mill = "精铣" if "精铣" in final_back_circle[5] else "铣"
                            continue
                        # 铰孔拆分为钻与铰，通过铰的参数获取钻的参数，这个钻单独处理
                        if final_back_circle[5] == "铰":
                            dtype = "铰"

                            if not final_back_circle[4] or not final_back_circle[2]:
                                continue
                            diameter = final_back_circle[2] - 0.2  # 钻直径
                            predefined_depth = final_back_circle[4] - (diameter * 0.7)  # 铰深
                            # 钻参数
                            temp_list_hinge.append(
                                (final_back_circle[0], final_back_circle[1], diameter, final_back_circle[3],
                                 final_back_circle[4] + diameter * 0.7, "钻", False))
                            # 铰参数
                            temp_list.append(
                                (final_back_circle[0], final_back_circle[1], final_back_circle[2], final_back_circle[3],
                                 predefined_depth, final_back_circle[5], False))
                            diameter = final_back_circle[2]  # 铰直径
                            continue
                        is_through = final_back_circle[6]
                        dtype = final_back_circle[5]
                        diameter = final_back_circle[2]
                        predefined_depth = final_back_circle[4]
                        temp_list.append(final_back_circle)

                # 创建铰孔工序之前的钻孔工序
                if temp_list_hinge:
                    if not predefined_depth:
                        continue
                    tool_name, step_distance, feed_rate, cycle_type = self.drill_library.get_drill_parameters(
                        material_type,
                        temp_list_hinge[0][2],
                        temp_list_hinge[0][5], )
                    if (tool_name is None and step_distance is None and feed_rate is None and cycle_type is None):
                        continue
                    if (len(temp_list) == 0):
                        continue

                    # 路径优化
                    b_front_optimization = self.path_optimizer.optimize_drilling_path(temp_list)
                    self.drilling_handler.create_drill_operation(
                        operation_type="DRILLING",
                        tool_type="STD_DRILL",
                        tool_name=tool_name,
                        geometry_name=b_mcs,
                        orient_geometry_name=b_mcs,
                        parent_group_name="背面",
                        group_name="背面钻孔",
                        method_group_name="METHOD",
                        hole_features=b_front_optimization,
                        predefined_depth=float(temp_list_hinge[0][4]),
                        diameter=float(temp_list_hinge[0][2]),
                        tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                        operation_name=f"B_{tag}_{tool_name}",
                        drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                        is_through=is_through,
                        step_distance=float(step_distance),
                        feed_rate=float(feed_rate),
                        cycle_type=cycle_type
                    )

                if not predefined_depth:
                    continue
                if diameter and diameter >= 42:
                    self.special_hole_handle(tag, temp_list, predefined_depth, dtype, material_type,
                                             drill_config.DEFAULT_MCS_NAME, "背面", "B", is_through)
                    continue

                tool_name, step_distance, feed_rate, cycle_type = self.drill_library.get_drill_parameters(material_type,
                                                                                                          diameter,
                                                                                                          dtype,
                                                                                                          have_mill)
                if (tool_name is None and step_distance is None and feed_rate is None and cycle_type is None):
                    continue
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
                        geometry_name=b_mcs,
                        orient_geometry_name=b_mcs,
                        parent_group_name="背面",
                        group_name="背面钻孔",
                        method_group_name="METHOD",
                        hole_features=b_front_optimization,
                        predefined_depth=float(predefined_depth),
                        diameter=float(diameter),
                        tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                        operation_name=f"B_{tag}_{tool_name}",
                        drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                        is_through=is_through,
                        step_distance=float(step_distance),
                        feed_rate=float(feed_rate),
                        cycle_type=cycle_type
                    )
                else:
                    self.drilling_handler.create_drill_operation(
                        operation_type="DRILLING",
                        tool_type="REAMER",
                        tool_name=tool_name,
                        geometry_name=b_mcs,
                        orient_geometry_name=b_mcs,
                        parent_group_name="背面",
                        group_name="背面铰孔",
                        method_group_name="METHOD",
                        hole_features=b_front_optimization,
                        predefined_depth=float(predefined_depth),
                        diameter=float(diameter),
                        tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                        operation_name=f"B_J_{tag}_{tool_name}",
                        drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                        is_through=False,
                        step_distance=float(step_distance),
                        feed_rate=float(feed_rate),
                        cycle_type=cycle_type
                    )
            return temp_list_mill

    # 背面铣孔
    def _perform_back_milling(self, b_sink_circles, material_type, mirrored_curves):
        """执行背面铣孔操作"""
        if mirrored_curves:
            # 处理镜像后的圆
            final_b_sink_circles = self._process_mirrored_circles(mirrored_curves, b_sink_circles, None)
            bool = is_mcs_exists(self.work_part, "B_MCS")
            if bool:
                b_mcs = drill_config.DEFAULT_B_MCS_NAME
            else:
                b_mcs = drill_config.DEFAULT_MCS_NAME
            # 创建铣孔工序
            self._perform_front_milling(final_b_sink_circles, material_type,
                                        geometry_name=b_mcs,
                                        parent_group_name="背面", group_name="背面铣孔")

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
                        final_back_circles.append(
                            (curve, (curve.CenterPoint.X, curve.CenterPoint.Y, curve.CenterPoint.Z),
                             base_circles[0][2], None, None, None))

                    elif (key[0] == round(-back_circle[1][0], 2) and
                          key[1] == round(back_circle[1][1], 2) and
                          key[2] == round(back_circle[1][2], 2)):
                        final_back_circles.append(
                            (curve, (curve.CenterPoint.X, curve.CenterPoint.Y, curve.CenterPoint.Z),
                             back_circle[2], back_circle[3], back_circle[4], back_circle[5], back_circle[6]))
            else:
                for back_circle in back_circles:
                    if (key[0] == round(-back_circle[1][0], 2) and
                            key[1] == round(back_circle[1][1], 2) and
                            key[2] == round(back_circle[1][2], 2)):
                        final_back_circles.append(
                            (curve, (curve.CenterPoint.X, curve.CenterPoint.Y, curve.CenterPoint.Z),
                             back_circle[2], back_circle[3], back_circle[4], back_circle[5], back_circle[6]))
        return final_back_circles

    def _perform_side_drilling(self, side_circles, material_type, base_circles, mirrored_curves=None, name="c"):
        if mirrored_curves:
            # 处理镜像后的圆
            side_circles = self._process_mirrored_circles(mirrored_curves, side_circles, base_circles)
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
                geometry_name=drill_config.DEFAULT_SIDE_MCS_NAME,
                orient_geometry_name=drill_config.DEFAULT_SIDE_MCS_NAME,
                parent_group_name="侧面",
                group_name="侧面正面打点" if name == "c" else "侧面背面打点",
                method_group_name="METHOD",
                hole_features=side_optimization,
                predefined_depth=drill_config.DEFAULT_DRILL_DEPTH,
                diameter=drill_config.DEFAULT_DRILL_DIAMETER,
                tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                operation_name=f"{name}_zxz",
                drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                is_through=False,
                step_distance=drill_config.DEFAULT_STEP_DISTANCE,
                cycle_type=drill_config.CYCLE_DRILL_STANDARD
            )
        # 侧面钻孔
        tag_list = [side_circle[3] for side_circle in side_circles]
        # 去重
        tag_list = list(set(tag_list))
        temp_list_mill = []  # 一般孔铣孔列表
        for tag in tag_list:
            if (tag is None):
                continue
            temp_list = []
            temp_list_hinge = []
            # 加工类型
            dtype = None
            # 孔直径
            diameter = None
            # 孔深度
            predefined_depth = None
            have_mill = ""
            # 是否通孔
            is_through = False
            for side_circle in side_circles:
                if (tag is not None and tag == side_circle[3]):
                    if "一般孔" in side_circle[5]:  # 一般孔铣或一般孔精铣
                        dtype = "钻"
                        diameter = side_circle[2] - 0.5 if "精铣" in side_circle else side_circle[2]
                        predefined_depth = side_circle[4]
                        if predefined_depth > 2 * diameter:
                            temp_list.append(side_circle)
                        temp_list_mill.append(side_circle)
                        have_mill = "精铣" if "精铣" in side_circle[5] else "铣"
                        continue
                    # 铰孔拆分为钻与铰，通过铰的参数获取钻的参数，这个钻单独处理
                    if side_circle[5] == "铰":
                        dtype = "铰"
                        if not side_circle[2] and not side_circles[4]:
                            continue
                        diameter = side_circle[2] - 0.2  # 钻直径
                        predefined_depth = side_circle[4] - (diameter * 0.7)  # 铰深
                        # 钻参数
                        temp_list_hinge.append((side_circle[0], side_circle[1], diameter, side_circle[3],
                                                side_circle[4] + diameter * 0.7, "钻", side_circle[6]))
                        # 铰参数
                        temp_list.append((side_circle[0], side_circle[1], side_circle[2], side_circle[3],
                                          predefined_depth, side_circle[5], False))
                        diameter = side_circle[2]  # 铰直径
                    else:
                        is_through = side_circle[6]
                        dtype = side_circle[5]
                        diameter = side_circle[2]
                        predefined_depth = side_circle[4]
                        temp_list.append(side_circle)

            if temp_list_hinge:
                if not predefined_depth:
                    continue
                tool_name, step_distance, feed_rate, cycle_type = self.drill_library.get_drill_parameters(
                    material_type,
                    temp_list_hinge[0][2],
                    temp_list_hinge[0][-1])
                if (tool_name is None and step_distance is None and feed_rate is None and cycle_type is None):
                    continue
                if (len(temp_list) == 0):
                    continue

                # 路径优化
                c_front_optimization = self.path_optimizer.optimize_drilling_path(temp_list)
                self.drilling_handler.create_drill_operation(
                    operation_type="DRILLING",
                    tool_type="STD_DRILL",
                    tool_name=tool_name,
                    geometry_name=drill_config.DEFAULT_SIDE_MCS_NAME,
                    orient_geometry_name=drill_config.DEFAULT_SIDE_MCS_NAME,
                    parent_group_name="侧面",
                    group_name="侧面正面钻孔" if name == "c" else "侧面背面钻孔",
                    method_group_name="METHOD",
                    hole_features=c_front_optimization,
                    predefined_depth=float(temp_list_hinge[0][4]),
                    diameter=float(temp_list_hinge[0][2]),
                    tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                    operation_name=f"{name.upper()}_{tag}_{tool_name}",
                    drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                    is_through=is_through,
                    step_distance=float(step_distance),
                    feed_rate=float(feed_rate),
                    cycle_type=cycle_type
                )

            if not predefined_depth:
                continue
            tool_name, step_distance, feed_rate, cycle_type = self.drill_library.get_drill_parameters(
                material_type,
                diameter,
                dtype,
                have_mill)
            if (tool_name is None and step_distance is None and feed_rate is None and cycle_type is None):
                continue
            if (len(temp_list) == 0):
                continue

            # 路径优化
            c_front_optimization = self.path_optimizer.optimize_drilling_path(temp_list)
            if (dtype != "铰"):
                self.drilling_handler.create_drill_operation(
                    operation_type="DRILLING",
                    tool_type="STD_DRILL",
                    tool_name=tool_name,
                    geometry_name=drill_config.DEFAULT_SIDE_MCS_NAME,
                    orient_geometry_name=drill_config.DEFAULT_SIDE_MCS_NAME,
                    parent_group_name="侧面",
                    group_name="侧面正面钻孔" if name == "c" else "侧面背面钻孔",
                    method_group_name="METHOD",
                    hole_features=c_front_optimization,
                    predefined_depth=float(predefined_depth),
                    diameter=float(diameter),
                    tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                    operation_name=f"{name.upper()}_{tag}_{tool_name}",
                    drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                    is_through=is_through,
                    step_distance=float(step_distance),
                    feed_rate=float(feed_rate),
                    cycle_type=cycle_type
                )
            else:
                self.drilling_handler.create_drill_operation(
                    operation_type="DRILLING",
                    tool_type="REAMER",
                    tool_name=tool_name,
                    geometry_name=drill_config.DEFAULT_SIDE_MCS_NAME,
                    orient_geometry_name=drill_config.DEFAULT_SIDE_MCS_NAME,
                    parent_group_name="侧面",
                    group_name="侧面正铰孔" if name == "c" else "侧面背铰孔",
                    method_group_name="METHOD",
                    hole_features=c_front_optimization,
                    predefined_depth=float(predefined_depth),
                    diameter=float(diameter),
                    tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                    operation_name=f"{'C_B_' if name == 'c' else 'C_'}J_{'B_' if name == 'c' else ''}{tag}_{tool_name}",
                    drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                    is_through=False,
                    step_distance=float(step_distance),
                    feed_rate=float(feed_rate),
                    cycle_type=cycle_type
                )
        return temp_list_mill

    # 侧面铣孔
    def _perform_side_milling(self, side_circles, material_type, mirrored_curves, bs):
        if bs:
            # 处理镜像后的圆
            final_c_sink_circles = self._process_mirrored_circles(mirrored_curves, side_circles, None)
            # 创建铣孔工序
            self._perform_front_milling(final_c_sink_circles, material_type,
                                        geometry_name=drill_config.DEFAULT_SIDE_MCS_NAME,
                                        parent_group_name="侧面", group_name="侧面背铣孔", direction="B_")
        else:
            self._perform_front_milling(side_circles, material_type, geometry_name=drill_config.DEFAULT_SIDE_MCS_NAME,
                                        parent_group_name="侧面", group_name="侧面正铣孔", direction="Z_")

    def _perform_front_view_drilling(self, front_circles, material_type):
        """执行正视图钻孔操作"""

        print_to_info_window(f"处理正视图圆: {len(front_circles)} 个")

        # 添加基准圆到正面打点列表
        complete_front_circles = front_circles

        print_to_info_window(f"正视图圆总数: {len(complete_front_circles)} 个")

        # 正面圆打点路径优化
        front_optimization = self.path_optimizer.optimize_drilling_path(complete_front_circles)

        # 创建正面打点工序
        if front_optimization:
            print_to_info_window(f"✅ 正视图圆路径优化完成，共 {len(front_optimization)} 个孔")

            # 检查正面圆是否有效
            valid_holes = [hole for hole in front_optimization if hole is not None]
            if not valid_holes:
                print_to_info_window("⚠️ 正视图圆优化结果为空，跳过正面钻孔")
                return
            # 创建程序组
            self.drilling_handler.get_or_create_program_group(group_name="前面")
            self.drilling_handler.create_drill_operation(
                operation_type="DRILLING",
                tool_type="CENTERDRILL",
                tool_name="zxz",
                geometry_name=drill_config.DEFAULT_FRONT_MCS_NAME,
                orient_geometry_name=drill_config.DEFAULT_FRONT_MCS_NAME,
                parent_group_name="前面",
                group_name="正视图打点工序",
                method_group_name="METHOD",
                hole_features=front_optimization,
                predefined_depth=drill_config.DEFAULT_DRILL_DEPTH,
                diameter=drill_config.DEFAULT_DRILL_DIAMETER,
                tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                operation_name="z_view_zxz",
                drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                is_through=False,
                step_distance=drill_config.DEFAULT_STEP_DISTANCE,
                cycle_type=drill_config.CYCLE_DRILL_STANDARD
            )
            print_to_info_window("✅ 正视图打点工序创建成功")
        else:
            print_to_info_window("⚠️ 正视图路径优化失败，跳过正面打点")
        # 正面钻孔
        tag_list = [front_circle[3] for front_circle in front_circles]
        # 去重
        tag_list = list(set(tag_list))
        temp_list_mill = []  # 一般孔铣孔列表
        for tag in tag_list:
            if (tag is None):
                continue
            temp_list = []
            temp_list_hinge = []  # 铰之前先钻，单独存放铰孔的钻孔参数
            # 加工类型
            dtype = None
            # 孔直径
            diameter = None
            # 孔深度
            predefined_depth = None
            have_mill = ""
            # 是否通孔
            is_through = False
            for front_circle in front_circles:
                if (tag is not None and tag == front_circle[3]):
                    if "一般孔" in front_circle[5]:  # 一般孔铣或一般孔精铣
                        dtype = "钻"
                        diameter = front_circle[2] - 0.5 if "精铣" in front_circle[5] else front_circle[2]
                        predefined_depth = front_circle[4]
                        temp_list.append(front_circle)
                        temp_list_mill.append(front_circle)
                        have_mill = "精铣" if "精铣" in front_circle[5] else "铣"
                        continue
                    # 铰孔拆分为钻与铰，通过铰的参数获取钻的参数，这个钻单独处理
                    if front_circle[5] == "铰":
                        dtype = "铰"

                        if not front_circle[4] or not front_circle[2]:
                            continue
                        diameter = front_circle[2] - 0.2  # 钻直径
                        predefined_depth = front_circle[4] - (diameter * 0.7)  # 铰深
                        # 钻参数
                        temp_list_hinge.append((front_circle[0], front_circle[1], diameter, front_circle[3],
                                                front_circle[4] + diameter * 0.7, "钻", front_circle[6]))
                        # 铰参数
                        temp_list.append((front_circle[0], front_circle[1], front_circle[2], front_circle[3],
                                          predefined_depth, front_circle[5], False))
                        diameter = front_circle[2]  # 铰直径
                        continue
                    is_through = front_circle[6]
                    dtype = front_circle[5]
                    diameter = front_circle[2]
                    predefined_depth = front_circle[4]
                    temp_list.append(front_circle)

            # 创建铰孔工序之前的钻孔工序
            if temp_list_hinge:
                if not predefined_depth:
                    continue
                tool_name, step_distance, feed_rate, cycle_type = self.drill_library.get_drill_parameters(material_type,
                                                                                                          temp_list_hinge[
                                                                                                              0][2],
                                                                                                          temp_list_hinge[
                                                                                                              0][-1])
                if (tool_name is None and step_distance is None and feed_rate is None and cycle_type is None):
                    continue
                if (len(temp_list) == 0):
                    continue

                # 路径优化
                z_front_optimization = self.path_optimizer.optimize_drilling_path(temp_list)
                self.drilling_handler.create_drill_operation(
                    operation_type="DRILLING",
                    tool_type="STD_DRILL",
                    tool_name=tool_name,
                    geometry_name=drill_config.DEFAULT_FRONT_MCS_NAME,
                    orient_geometry_name=drill_config.DEFAULT_FRONT_MCS_NAME,
                    parent_group_name="前面",
                    group_name="正视图钻孔",
                    method_group_name="METHOD",
                    hole_features=z_front_optimization,
                    predefined_depth=float(temp_list_hinge[0][4]),
                    diameter=float(temp_list_hinge[0][2]),
                    tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                    operation_name=f"Z_view_{tag}_{tool_name}",
                    drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                    is_through=is_through,
                    step_distance=float(step_distance),
                    feed_rate=float(feed_rate),
                    cycle_type=cycle_type
                )
            if not predefined_depth:
                continue
            tool_name, step_distance, feed_rate, cycle_type = self.drill_library.get_drill_parameters(material_type,
                                                                                                      diameter,
                                                                                                      dtype,
                                                                                                      have_mill)
            if (tool_name is None and step_distance is None and feed_rate is None and cycle_type is None):
                continue
            if (len(temp_list) == 0):
                continue

            # 路径优化
            z_front_optimization = self.path_optimizer.optimize_drilling_path(temp_list)
            if (dtype != "铰"):
                self.drilling_handler.create_drill_operation(
                    operation_type="DRILLING",
                    tool_type="STD_DRILL",
                    tool_name=tool_name,
                    geometry_name=drill_config.DEFAULT_FRONT_MCS_NAME,
                    orient_geometry_name=drill_config.DEFAULT_FRONT_MCS_NAME,
                    parent_group_name="前面",
                    group_name="正视图钻孔",
                    method_group_name="METHOD",
                    hole_features=z_front_optimization,
                    predefined_depth=float(predefined_depth),
                    diameter=float(diameter),
                    tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                    operation_name=f"Z_view_{tag}_{tool_name}",
                    drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                    is_through=is_through,
                    step_distance=float(step_distance),
                    feed_rate=float(feed_rate),
                    cycle_type=cycle_type
                )
            else:
                self.drilling_handler.create_drill_operation(
                    operation_type="DRILLING",
                    tool_type="REAMER",
                    tool_name=tool_name,
                    geometry_name=drill_config.DEFAULT_FRONT_MCS_NAME,
                    orient_geometry_name=drill_config.DEFAULT_FRONT_MCS_NAME,
                    parent_group_name="前面",
                    group_name="正视图铰孔",
                    method_group_name="METHOD",
                    hole_features=z_front_optimization,
                    predefined_depth=float(predefined_depth),
                    diameter=float(diameter),
                    tip_diameter=drill_config.DEFAULT_TIP_DIAMETER,
                    operation_name=f"Z_view_J_{tag}_{tool_name}",
                    drive_point=drill_config.TOOL_DRIVE_POINT_TIP,
                    is_through=False,
                    step_distance=float(step_distance),
                    feed_rate=float(feed_rate),
                    cycle_type=cycle_type
                )
        return temp_list_mill

    def _classify__front_circles_by_processing(self, z_view_circle, processed_result, lwh_point, all_circles=None,
                                               threading=False, red_center_list=None):
        """根据加工说明分类正反面圆"""

        print_to_info_window(f"开始分类圆: 正视图圆 {len(z_view_circle)} 个")

        # 根据注释对圆分类
        tag_circle_map = self.process_info_handler.fenlei(self.work_part, z_view_circle, lwh_point, None,
                                                          all_circles, threading, red_center_list)
        print_to_info_window(f"标签-圆映射: {len(tag_circle_map)} 个标签")
        print_to_info_window(f"加工结果: {len(processed_result)} 个加工说明")
        # 存放正面打点的圆对象
        front_view_circle_list = []
        # 存放正沉头孔对象
        z_view_sink_circle_list = []
        # 存放背面打点的圆对象
        front_view_back_circle_list = []
        # 存放背沉头孔对象
        b_view_sink_circle_list = []
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
                            is_through = processed_result[tag]["is_through_hole"]
                            predefined_depth = (lwh_point[2]
                                                if processed_result[tag]["is_through_hole"] and not
                            processed_result[tag]["is_zzANDbz"]
                                                else processed_result[tag]["depth"]["hole_depth"]
                                                )
                            # 判断是否是穿线孔 穿线孔钻完后距离割边要有2mm余量
                            if (processed_result[tag]["main_hole_processing"] == "割" and processed_result[tag][
                                "real_diamater"] is not None):
                                real_diameter = real_diameter - 2
                            # 画的穿线圆处理
                            if threading:
                                real_diameter = circle_info[0][2]
                                predefined_depth = lwh_point[2]
                            front_view_back_circle_list.append(
                                (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key,
                                 is_through))
                            # 处理背沉头
                            if (processed_result[tag]["circle_type"] == "背沉头"):
                                circle_obj = circle_info[0][0]
                                circle_center_point = circle_info[0][1]
                                real_diameter = processed_result[tag]["real_diamater_head"]
                                predefined_depth = processed_result[tag]["depth"]["head_depth"]
                                main_key = "铣"
                                b_view_sink_circle_list.append(
                                    (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth,
                                     main_key, False))
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
                            is_through = processed_result[tag]["is_through_hole"]
                            predefined_depth = (lwh_point[2]
                                                if processed_result[tag]["is_through_hole"] and not
                            processed_result[tag]["is_zzANDbz"]
                                                else processed_result[tag]["depth"]["hole_depth"]
                                                )
                            # 判断是否是穿线孔 穿线孔钻完后距离割边要有2mm余量
                            if (processed_result[tag]["main_hole_processing"] == "割" and processed_result[tag][
                                "real_diamater"] is not None):
                                real_diameter = real_diameter - 2

                            # 画的穿线圆处理
                            if threading:
                                real_diameter = circle_info[0][2]
                                predefined_depth = lwh_point[2]

                            front_view_circle_list.append(
                                (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth, main_key,
                                 is_through))
                            # 处理正沉头
                            if (processed_result[tag]["circle_type"] == "正沉头"):
                                circle_obj = circle_info[0][0]
                                circle_center_point = circle_info[0][1]
                                real_diameter = processed_result[tag]["real_diamater_head"]
                                predefined_depth = processed_result[tag]["depth"]["head_depth"]
                                main_key = "铣"
                                z_view_sink_circle_list.append(
                                    (circle_obj, circle_center_point, real_diameter, tag_key, predefined_depth,
                                     main_key, False))

            if not tag_matched:
                print_to_info_window(f"⚠️ 标签 {tag} 未匹配到任何圆")

        print_to_info_window(
            f"分类结果: 正面圆 {len(front_view_circle_list)} 个，背面圆 {len(front_view_back_circle_list)} 个，正沉头 {len(z_view_sink_circle_list)} 个，背沉头 {len(b_view_sink_circle_list)} 个")

        return {
            "front_view_circle": front_view_circle_list,
            "front_view_back_circle": front_view_back_circle_list,
            "z_view_sink_circle": z_view_sink_circle_list,
            "b_view_sink_circle": b_view_sink_circle_list
        }
