# -*- coding: utf-8 -*-
"""
几何体处理模块
包含坐标系创建、圆识别、草图创建等功能
"""

import math
import NXOpen
import NXOpen.CAM
from utils import print_to_info_window, handle_exception, euclidean_distance, safe_origin, endpoints
import drill_config

class GeometryHandler:
    """几何体处理类"""

    def __init__(self, session, work_part):
        self.session = session
        self.work_part = work_part

    def create_mcs_with_safe_plane(self, origin_point=None, mcs_name=None, safe_distance=None):
        """创建 MCS 坐标系，并在包容体上方设置安全平面"""

        origin_point = origin_point or drill_config.DEFAULT_ORIGIN_POINT
        mcs_name = mcs_name or drill_config.DEFAULT_MCS_NAME
        safe_distance = safe_distance or drill_config.DEFAULT_SAFE_DISTANCE

        try:
            # 获取 CAM 根几何组
            geom_group = self.work_part.CAMSetup.CAMGroupCollection.FindObject(drill_config.DEFAULT_GEOMETRY_GROUP)
            if geom_group is None:
                return handle_exception("未找到 GEOMETRY 组，无法创建 MCS")

            # 删除同名 MCS
            try:
                existing = self.work_part.CAMSetup.CAMGroupCollection.FindObject(f"GEOMETRY/{mcs_name}")
                if existing:
                    existing.Delete()
                    print_to_info_window(f"已删除同名 MCS: {mcs_name}")
            except Exception:
                pass

            # 创建几何组 MCS
            mcs_group = self.work_part.CAMSetup.CAMGroupCollection.CreateGeometry(
                geom_group,
                "hole_making",
                "MCS",
                NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue,
                mcs_name
            )

            # 创建 MillOrientGeomBuilder
            builder = self.work_part.CAMSetup.CAMGroupCollection.CreateMillOrientGeomBuilder(mcs_group)

            # 创建坐标系
            origin3 = NXOpen.Point3d(*origin_point)
            x_dir = NXOpen.Vector3d(1.0, 0.0, 0.0)
            y_dir = NXOpen.Vector3d(0.0, 1.0, 0.0)
            xform = self.work_part.Xforms.CreateXform(origin3, x_dir, y_dir,
                                                      NXOpen.SmartObject.UpdateOption.AfterModeling, 1.0)
            csys = self.work_part.CoordinateSystems.CreateCoordinateSystem(xform,
                                                                           NXOpen.SmartObject.UpdateOption.AfterModeling)
            builder.Mcs = csys

            # 设置安全平面（固定平面）
            builder.TransferClearanceBuilder.ClearanceType = NXOpen.CAM.NcmClearanceBuilder.ClearanceTypes.Plane

            # 创建安全平面
            plane_safe = self.work_part.Planes.CreatePlane(
                NXOpen.Point3d(origin_point[0], origin_point[1], origin_point[2] + safe_distance),
                NXOpen.Vector3d(0.0, 0.0, 1.0),
                NXOpen.SmartObject.UpdateOption.AfterModeling
            )
            builder.TransferClearanceBuilder.PlaneXform = plane_safe

            # 提交 MCS
            nx_obj = builder.Commit()
            builder.Destroy()
            print_to_info_window(f"✅ MCS '{mcs_name}' 创建完成，安全高度 {safe_distance} mm 已设置")
            return nx_obj

        except Exception as ex:
            return handle_exception("创建 MCS 坐标系失败", str(ex))

    def create_workpiece_geometry(self, parent_mcs_name=None, dtype="hole_making", new_group_name="WORKPIECE_Z"):
        """在指定 MCS 下创建 WORKPIECE 几何组"""

        parent_mcs_name = parent_mcs_name or drill_config.DEFAULT_MCS_NAME

        try:
            # 获取父 MCS
            parent_mcs = self.work_part.CAMSetup.CAMGroupCollection.FindObject(parent_mcs_name)
            if parent_mcs is None:
                return handle_exception(f"未找到 MCS: {parent_mcs_name}")

            # 创建几何组
            nCGroup = self.work_part.CAMSetup.CAMGroupCollection.CreateGeometry(
                parent_mcs,
                dtype,
                "WORKPIECE",
                NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue,
                new_group_name
            )

            # 创建 MillGeomBuilder
            mill_geom_builder = self.work_part.CAMSetup.CAMGroupCollection.CreateMillGeomBuilder(nCGroup)

            # 提交
            nXObject = mill_geom_builder.Commit()
            mill_geom_builder.Destroy()

            print_to_info_window(f"✅ 已创建 WORKPIECE 几何组: {new_group_name}")
            return nCGroup

        except Exception as ex:
            return handle_exception("创建 WORKPIECE 几何组失败", str(ex))

    def get_all_curves(self):
        """获取所有曲线"""
        return list(self.work_part.Curves)

    # 获取主视图四个边界点
    def main_view_bound_points(self, lwh_point):
        curves = self.get_all_curves()
        l_bound_list = []
        k_bound_list = []
        minx, miny, maxx, maxy = float("inf"), float("inf"), float("-inf"), float("-inf")
        for curve in curves:
            if isinstance(curve, NXOpen.Line):
                p1 = (curve.StartPoint.X, curve.StartPoint.Y, curve.StartPoint.Z)
                p2 = (curve.EndPoint.X, curve.EndPoint.Y, curve.EndPoint.Z)
                distance = euclidean_distance(p2, p1)
                if round(distance, 1) == round(lwh_point[0], 1):
                    l_bound_list.append(curve)
                if round(distance, 1) == round(lwh_point[1], 1):
                    k_bound_list.append(curve)
        for l_curve in l_bound_list:
            lp1 = (round(l_curve.StartPoint.X, 3), round(l_curve.StartPoint.Y, 3), round(l_curve.StartPoint.Z, 3))
            lp2 = (round(l_curve.EndPoint.X, 3), round(l_curve.EndPoint.Y, 3), round(l_curve.EndPoint.Z, 3))
            for k_curve in k_bound_list:
                kp1 = (round(k_curve.StartPoint.X, 3), round(k_curve.StartPoint.Y, 3), round(k_curve.StartPoint.Z, 3))
                kp2 = (round(k_curve.EndPoint.X, 3), round(k_curve.EndPoint.Y, 3), round(k_curve.EndPoint.Z, 3))
                # 判断交点
                if lp1 == kp1:
                    minx, miny, maxx, maxy = self.compare_value(lp1, minx, miny, maxx, maxy)
                if lp1 == kp2:
                    minx, miny, maxx, maxy = self.compare_value(lp1, minx, miny, maxx, maxy)
                if lp2 == kp1:
                    minx, miny, maxx, maxy = self.compare_value(lp2, minx, miny, maxx, maxy)
                if lp2 == kp2:
                    minx, miny, maxx, maxy = self.compare_value(lp2, minx, miny, maxx, maxy)
        return minx, miny, maxx, maxy

    def compare_value(self, line, minx, miny, maxx, maxy):
        if minx > line[0]:
            minx = line[0]
        if maxx < line[0]:
            maxx = line[0]
        if miny > line[1]:
            miny = line[1]
        if maxy < line[1]:
            maxy = line[1]
        return minx, miny, maxx, maxy


    def compute_bound_line_distince(self, lwh_point):
        """计算边界距离"""
        curves = self.get_all_curves()
        right_min_point = float("inf")
        front_max_point = float("-inf")
        p1 = None
        p2 = None
        for curve_item in curves:
            if isinstance(curve_item, NXOpen.Line):
                points = endpoints(curve_item)
                length = euclidean_distance(points[0], points[1])
                # 找出侧视图板料线最小的X轴的点即为移动距离
                if abs(length - lwh_point[2]) < 0.1 and points[0][0] > lwh_point[0]:
                    if points[0][0] < points[1][0]:
                        if points[0][0] < right_min_point:
                            right_min_point = points[0][0]
                            p1 = (points[0][0],0.0,0.0)
                    else:
                        if points[1][0] < right_min_point:
                            right_min_point = points[1][0]
                            p1 = (points[1][0],0.0,0.0)
                # 正视图
                if abs(length - lwh_point[2]) < 0.1 and points[0][1] < -0.1:
                    if points[0][1] > points[1][1]:
                        if points[0][1] > front_max_point:
                            front_max_point = points[0][1]
                            p2 = (0.0,points[0][1],0.0)
                    else:
                        if points[1][1] > front_max_point:
                            front_max_point = points[1][1]
                            p2 = (0.0,points[1][1],0.0)

        return (p1, p2)


    def find_nearest_circle(self, circle_list, point):
        """查找离坐标最近的点"""

        if not circle_list:
            return None

        min_dist = float("inf")
        nearest_circle = None

        for circle in circle_list:
            cx, cy, cz = float(circle[1][0]), float(circle[1][1]), float(circle[1][2])
            dx = point[0] - cx
            dy = point[1] - cy
            dz = point[2] - cz
            dist = math.sqrt(dx * dx + dy * dy + dz * dz)
            if dist < min_dist:
                min_dist = dist
                nearest_circle = (circle[0], (cx, cy, cz), circle[2], dist)

        return nearest_circle

    def create_circle_robust(self, center_xyz, radius):
        """
        直接创建 NXOpen.Arc 几何对象（严格匹配NX2312官方文档重载）
        入参：center_xyz (x,y,z) 圆心坐标，radius 半径（数字）
        返回：NXOpen.Arc 对象 / None（失败时写入 ListingWindow）
        """
        workPart = self.work_part
        try:
            # 1. 解析输入参数为NXOpen基础类型
            px, py, pz = (float(v) for v in center_xyz)
            center = NXOpen.Point3d(px, py, pz)  # 圆心（Point3d）
            radius_val = float(radius)  # 半径（double）
            x_direction = NXOpen.Vector3d(1.0, 0.0, 0.0)  # X方向（Vector3d）
            y_direction = NXOpen.Vector3d(0.0, 1.0, 0.0)  # Y方向（Vector3d）

            # 2. 角度转换：弧度，完整圆=0 ~ 2π
            start_angle = 0.0  # 起始角度（弧度）
            end_angle = 2 * math.pi  # 终止角度（弧度，完整圆）

            # 3. 调用CreateArc创建完整圆
            # 参数：center + x_direction + y_direction + radius + start_angle + end_angle
            arc_obj = workPart.Curves.CreateArc(
                center,
                x_direction,
                y_direction,
                radius_val,
                start_angle,
                end_angle
            )

            return arc_obj

        except Exception as e:
            print_to_info_window(f"直接创建Arc失败: {str(e)}")
            return None

    def create_circle_sketch(self, center_point_list, diameter=3.0, sketch_name="SKETCH_000"):
        """在当前工作部件中创建草图，并绘制一个圆"""

        try:
            theSession = NXOpen.Session.GetSession()
            workPart = self.work_part

            # 开始任务环境
            theSession.BeginTaskEnvironment()

            # 创建草图坐标系
            origin = NXOpen.Point3d(0.0, 0.0, 0.0)
            zDir = NXOpen.Vector3d(0.0, 0.0, 1.0)
            plane = workPart.Planes.CreatePlane(origin, zDir, NXOpen.SmartObject.UpdateOption.WithinModeling)

            direction = workPart.Directions.CreateDirection(origin, NXOpen.Vector3d(1.0, 0.0, 0.0),
                                                            NXOpen.SmartObject.UpdateOption.WithinModeling)
            point = workPart.Points.CreatePoint(origin)

            xform = workPart.Xforms.CreateXformByPlaneXDirPoint(plane, direction, point,
                                                                NXOpen.SmartObject.UpdateOption.WithinModeling,
                                                                1.0, False, False)
            csy = workPart.CoordinateSystems.CreateCoordinateSystem(xform,
                                                                    NXOpen.SmartObject.UpdateOption.WithinModeling)

            # 创建草图
            sketchBuilder = workPart.Sketches.CreateSimpleSketchInPlaceBuilder()
            sketchBuilder.CoordinateSystem = csy

            nXObject = sketchBuilder.Commit()
            sketch = nXObject

            sketchBuilder.Destroy()

            arc_list = []
            # 激活草图
            sketch.Activate(NXOpen.Sketch.ViewReorient.TrueValue)
            # 适合窗口
            self.work_part.ModelingViews.WorkView.Fit()
            for center_point in center_point_list:
                if sketch_name == "SKETCH_threading_circle":
                    diameter = center_point[2]
                # 绘制圆
                nXMatrix = sketch.Orientation
                radius = float(diameter) / 2.0
                center = NXOpen.Point3d(float(center_point[0]), float(center_point[1]), 0.0)
                arc = workPart.Curves.CreateArc(center, nXMatrix, radius, 0.0, 2.0 * math.pi)
                # sketch.AddGeometry(arc, NXOpen.Sketch.InferConstraintsOption.InferNoConstraints)
                sketch.Update()
                arc_list.append(arc)
            # 完成草图
            workRegionBuilder = workPart.Sketches.CreateWorkRegionBuilder()
            workRegionBuilder.Scope = NXOpen.SketchWorkRegionBuilder.ScopeType.EntireSketch
            workRegionBuilder.Commit()
            workRegionBuilder.Destroy()

            # 计算状态并退出草图
            sketch.CalculateStatus()
            sketch.Deactivate(NXOpen.Sketch.ViewReorient.TrueValue, NXOpen.Sketch.UpdateLevel.Model)

            # 设置草图名
            try:
                sketch.SetName(sketch_name)
            except Exception:
                pass

            # 结束任务环境
            theSession.EndTaskEnvironment()

            return (sketch, arc_list)

        except Exception as e:
            handle_exception("创建草图失败", str(e))
            return None
    # 计算两点之间的距离
    def euclidean_point(self, point1, point2):
        return math.sqrt((point1[0] - point2[0])**2 + (point1[1] - point2[1])**2 + (point1[2] - point2[2])**2)

    # -------------------------------
    # 数值工具
    # -------------------------------
    def fix(self, v, tol=1e-6):
        return 0.0 if abs(v) < tol else v

    def point_tuple(self, p):
        return (self.fix(p.X), self.fix(p.Y), self.fix(p.Z))

    def euclidean_point(self, p1, p2):
        return math.sqrt(
            (p1[0] - p2[0]) ** 2 +
            (p1[1] - p2[1]) ** 2 +
            (p1[2] - p2[2]) ** 2
        )

    # -------------------------------
    # 线段格式化
    # -------------------------------
    def point_format(self, line):
        p1 = line.StartPoint
        p2 = line.EndPoint
        return self.point_tuple(p1), self.point_tuple(p2)

    # -------------------------------
    # 线段相交（2D）
    # -------------------------------
    def line_intersection_2d(self, p1, p2, p3, p4, tol=1e-6):
        def cross(a, b):
            return a[0] * b[1] - a[1] * b[0]

        r = (p2[0] - p1[0], p2[1] - p1[1])
        s = (p4[0] - p3[0], p4[1] - p3[1])

        denom = cross(r, s)
        if abs(denom) < tol:
            return None

        t = cross((p3[0] - p1[0], p3[1] - p1[1]), s) / denom
        u = cross((p3[0] - p1[0], p3[1] - p1[1]), r) / denom

        if -tol <= t <= 1 + tol and -tol <= u <= 1 + tol:
            x = p1[0] + t * r[0]
            y = p1[1] + t * r[1]
            return (self.fix(x), self.fix(y), 0.0)

        return None

    # -------------------------------
    # 获取所有直线
    # -------------------------------
    def get_all_lines(self):
        lines = []
        for curve in self.work_part.Curves:
            if isinstance(curve, NXOpen.Line):
                lines.append(curve)
        return lines

    # -------------------------------
    # 主函数：获取加工原点
    # -------------------------------
    def get_start_point(self, lwh_point, view_name, length_tol=1.0, dist_tol=1.0):

        # ========= 1. 获取“0”注释 =========
        zero_notes = []
        for note in self.work_part.Notes:
            try:
                text = note.GetText()[0]
                if text.strip() in ["0", "0.0", "0.00", "-0"]:
                    zero_notes.append(note)
            except:
                pass

        if len(zero_notes) < 2:
            return None
        l_target = None
        w_target = None
        note_points = [safe_origin(n) for n in zero_notes]
        if view_name == "主视图":
            l_target, w_target = lwh_point[0], lwh_point[1]
        elif view_name == "正视图":
            l_target, w_target = lwh_point[0], lwh_point[2]
        elif view_name == "侧视图":
            l_target, w_target = lwh_point[1], lwh_point[2]
        l_lines = []
        w_lines = []
        if l_target is None or w_target is None:
            return None
        for line in self.get_all_lines():
            p1, p2 = self.point_format(line)
            length = self.euclidean_point(p1, p2)

            if abs(length - l_target) < length_tol:
                l_lines.append(line)
            elif abs(length - w_target) < length_tol:
                w_lines.append(line)

        if not l_lines or not w_lines:
            return None
        l_target = 0
        w_target = 0
        # ========= 3. 查找所有交点 =========
        best_point = None
        min_dist = float("inf")

        for l_line in l_lines:
            lp1, lp2 = self.point_format(l_line)
            for w_line in w_lines:
                wp1, wp2 = self.point_format(w_line)

                ip = self.line_intersection_2d(lp1, lp2, wp1, wp2)
                if not ip:
                    continue

                # ========= 4. 距离注释最近原则 =========
                for np in note_points:
                    d = self.euclidean_point(ip, np)
                    if d < min_dist:
                        min_dist = d
                        best_point = ip

        if not best_point:
            return None

        x = self.fix(best_point[0])
        y = self.fix(best_point[1])
        return (x, y, 0.0)


    def calculate_extended_line_intersection(self, line1, line2, tolerance=0.001):
        """
        计算两条直线延长后的交点（适用于不相交但延长后相交的线段）

        参数:
            line1, line2: NXOpen.Line 对象
            tolerance: 容差（默认0.001mm）

        返回:
            (success, point, status, parameters)
            - success: bool, 是否成功计算
            - point: (x, y, z) 交点坐标
            - status: str, 状态描述
            - parameters: (t, s) 交点在两条直线上的参数值
                         t<0或t>1表示交点在line1线段外
                         s<0或s>1表示交点在line2线段外
        """
        try:
            # 获取直线端点
            p1 = (line1.StartPoint.X, line1.StartPoint.Y, line1.StartPoint.Z)
            p2 = (line1.EndPoint.X, line1.EndPoint.Y, line1.EndPoint.Z)
            p3 = (line2.StartPoint.X, line2.StartPoint.Y, line2.StartPoint.Z)
            p4 = (line2.EndPoint.X, line2.EndPoint.Y, line2.EndPoint.Z)

            # 计算方向向量
            v1 = (p2[0] - p1[0], p2[1] - p1[1], p2[2] - p1[2])
            v2 = (p4[0] - p3[0], p4[1] - p3[1], p4[2] - p3[2])

            # 检查是否平行（叉积接近零向量）
            cross = (
                v1[1] * v2[2] - v1[2] * v2[1],
                v1[2] * v2[0] - v1[0] * v2[2],
                v1[0] * v2[1] - v1[1] * v2[0]
            )
            cross_norm = math.sqrt(cross[0] ** 2 + cross[1] ** 2 + cross[2] ** 2)

            if cross_norm < tolerance:
                return False, None, "直线平行，无交点", None

            # 求解参数 t 和 s: p1 + t*v1 = p3 + s*v2
            # 构建线性方程组求解
            # 使用最小二乘法求解超定方程组
            v1v1 = v1[0] ** 2 + v1[1] ** 2 + v1[2] ** 2
            v2v2 = v2[0] ** 2 + v2[1] ** 2 + v2[2] ** 2
            v1v2 = v1[0] * v2[0] + v1[1] * v2[1] + v1[2] * v2[2]

            rhs_x = p3[0] - p1[0]
            rhs_y = p3[1] - p1[1]
            rhs_z = p3[2] - p1[2]

            v1_rhs = v1[0] * rhs_x + v1[1] * rhs_y + v1[2] * rhs_z
            v2_rhs = v2[0] * rhs_x + v2[1] * rhs_y + v2[2] * rhs_z

            denom = v1v1 * v2v2 - v1v2 * v1v2

            if abs(denom) < tolerance:
                return False, None, "计算失败", None

            t = (v1_rhs * v2v2 - v2_rhs * v1v2) / denom
            s = (v1v1 * v2_rhs - v1v2 * v1_rhs) / denom

            # 计算交点（在line1上）
            intersection = (
                p1[0] + t * v1[0],
                p1[1] + t * v1[1],
                p1[2] + t * v1[2]
            )

            # 判断交点相对于线段的位置
            status = "延长线交点"
            if 0 <= t <= 1 and 0 <= s <= 1:
                status = "线段内交点"
            elif 0 <= t <= 1:
                status += "（line1在线段内）"
            elif 0 <= s <= 1:
                status += "（line2在线段内）"

            return True, intersection, status, (t, s)

        except Exception as e:
            return False, None, f"计算错误: {str(e)}", None

    def rotate_objects(self, rotation_vector, rotation_point, angle_deg, layer=9):
        """
        旋转所有曲线对象
        :param rotation_vector: tuple 或 NXOpen.Vector3d
        :param rotation_point: tuple 或 NXOpen.Point3d
        :param angle_deg: 旋转角度（度）
        :param layer: 图层
        :return: 已提交对象列表
        """
        curve_list = self.get_all_curves()  # 获取所有曲线

        # 类型转换（循环外完成）
        if isinstance(rotation_point, tuple):
            rotation_point = NXOpen.Point3d(*rotation_point)
        if isinstance(rotation_vector, tuple):
            rotation_vector = NXOpen.Vector3d(*rotation_vector)

        # 创建 MoveObjectBuilder
        moveBuilder = self.work_part.BaseFeatures.CreateMoveObjectBuilder(NXOpen.Features.MoveObject.Null)
        moveBuilder.Layer = layer

        # 设置旋转轴
        direction = self.work_part.Directions.CreateDirection(rotation_point, rotation_vector,
                                                              NXOpen.SmartObject.UpdateOption.WithinModeling)
        direction.ProtectFromDelete()
        axis = self.work_part.Axes.CreateAxis(NXOpen.Point.Null, direction,
                                              NXOpen.SmartObject.UpdateOption.WithinModeling)
        moveBuilder.TransformMotion.AngularAxis = axis

        # 设置旋转类型和角度
        moveBuilder.TransformMotion.DeltaEnum = NXOpen.GeometricUtilities.ModlMotion.Delta.ReferenceAcsWorkPart
        moveBuilder.TransformMotion.Option = NXOpen.GeometricUtilities.ModlMotion.Options.Angle
        moveBuilder.TransformMotion.Angle.SetFormula(str(angle_deg))

        # 添加所有曲线
        moveBuilder.ObjectToMoveObject.Add(curve_list)

        # 提交
        curve_objs = moveBuilder.Commit()

        # 清理
        moveBuilder.Destroy()
        direction.ReleaseDeleteProtection()

        note_list = list(self.work_part.Notes)
        # 类型转换（循环外完成）
        if isinstance(rotation_point, tuple):
            rotation_point = NXOpen.Point3d(*rotation_point)
        if isinstance(rotation_vector, tuple):
            rotation_vector = NXOpen.Vector3d(*rotation_vector)

        # 创建 MoveObjectBuilder
        moveBuilder = self.work_part.BaseFeatures.CreateMoveObjectBuilder(NXOpen.Features.MoveObject.Null)
        moveBuilder.Layer = layer

        # 设置旋转轴
        direction = self.work_part.Directions.CreateDirection(rotation_point, rotation_vector,
                                                              NXOpen.SmartObject.UpdateOption.WithinModeling)
        direction.ProtectFromDelete()
        axis = self.work_part.Axes.CreateAxis(NXOpen.Point.Null, direction,
                                              NXOpen.SmartObject.UpdateOption.WithinModeling)
        moveBuilder.TransformMotion.AngularAxis = axis

        # 设置旋转类型和角度
        moveBuilder.TransformMotion.DeltaEnum = NXOpen.GeometricUtilities.ModlMotion.Delta.ReferenceAcsWorkPart
        moveBuilder.TransformMotion.Option = NXOpen.GeometricUtilities.ModlMotion.Options.Angle
        moveBuilder.TransformMotion.Angle.SetFormula(str(angle_deg))

        # 添加所有曲线
        moveBuilder.ObjectToMoveObject.Add(note_list)

        # 提交
        note_obj = moveBuilder.Commit()

        # 清理
        moveBuilder.Destroy()
        direction.ReleaseDeleteProtection()

    def move_objects_point_to_point(self, from_point, to_point, layer=9, is_copy=False, mill_note=None):
        """
        点对点移动，可选择移动或复制，from_point/to_point 可以是 NXOpen.Point3d 或 (x,y,z)
        :param from_point:  移动起点 NXOpen.Point3d 或 (x, y, z)
        :param to_point:    移动终点 NXOpen.Point3d 或 (x, y, z)
        :param layer:       图层
        :param is_copy:     False=移动（默认），True=复制
        :param mill_note:  单独移动M、M2、M2M格式的注释(复制为新对象，移动至镜像上)
        :return: list of committed objects (复制时为新对象)
        """
        work_part = self.work_part

        # helper: 把输入规范成 NXOpen.Point3d
        def _to_point3d(p):
            if isinstance(p, NXOpen.Point3d):
                return p
            # 支持 list/tuple/其他可迭代
            try:
                x, y, z = p
                return NXOpen.Point3d(float(x), float(y), float(z))
            except Exception:
                raise ValueError("from_point/to_point 必须是 NXOpen.Point3d 或 可迭代三元组 (x,y,z)")

        # 收集对象（保持你原来的选择逻辑）
        un_move_obj_list = []
        if not mill_note:
            note_list = list(work_part.Notes)
            line_list = list(work_part.Lines)
            new_line_list = []
            new_spline_list = []
            # 过滤掉长度为0的直线
            for line in line_list:
                p1 = (round(line.StartPoint.X, 1), round(line.StartPoint.Y, 1))
                p2 = (round(line.EndPoint.X, 1), round(line.EndPoint.Y, 1))
                if p1 != p2:
                    new_line_list.append(line)
                else:
                    un_move_obj_list.append(line)
            arc_list = list(work_part.Arcs)
            spline_list = list(work_part.Splines)
            # 过滤掉长度为0的样条
            for spline in spline_list:
                if spline.GetLength() > 0.01:
                    new_spline_list.append(spline)
                else:
                    un_move_obj_list.append(spline)
            objects = new_line_list + arc_list + new_spline_list + note_list
        else:
            objects = [mill_note]
        if len(un_move_obj_list) > 0:
            self.delete_nx_objects(un_move_obj_list)
        # 创建 MoveObjectBuilder
        builder = work_part.BaseFeatures.CreateMoveObjectBuilder(
            NXOpen.Features.MoveObject.Null
        )

        # 根据录制宏使用 MoveObjectResult 控制 复制/移动
        if is_copy:
            builder.MoveObjectResult = NXOpen.Features.MoveObjectBuilder.MoveObjectResultOptions.CopyOriginal
        else:
            builder.MoveObjectResult = NXOpen.Features.MoveObjectBuilder.MoveObjectResultOptions.MoveOriginal

        # 保留录制宏的 TransformMotion 配置
        tm = builder.TransformMotion
        tm.DistanceAngle.OrientXpress.AxisOption = NXOpen.GeometricUtilities.OrientXpressBuilder.Axis.Passive
        tm.DistanceAngle.OrientXpress.PlaneOption = NXOpen.GeometricUtilities.OrientXpressBuilder.Plane.Passive
        tm.OrientXpress.AxisOption = NXOpen.GeometricUtilities.OrientXpressBuilder.Axis.Passive
        tm.OrientXpress.PlaneOption = NXOpen.GeometricUtilities.OrientXpressBuilder.Plane.Passive

        tm.DeltaEnum = NXOpen.GeometricUtilities.ModlMotion.Delta.ReferenceAcsWorkPart
        tm.Option = NXOpen.GeometricUtilities.ModlMotion.Options.PointToPoint

        builder.Layer = layer

        # 处理输入点（支持 NXOpen.Point3d 或 三元组）
        from_p3d = _to_point3d(from_point)
        to_p3d = _to_point3d(to_point)

        # 创建 Point 对象（Journal 中是 CreatePoint），并移除参数
        fromPt = work_part.Points.CreatePoint(from_p3d)
        work_part.Points.RemoveParameters(fromPt)

        toPt = work_part.Points.CreatePoint(to_p3d)
        work_part.Points.RemoveParameters(toPt)

        tm.FromPoint = fromPt
        tm.ToPoint = toPt

        # 添加要移动/复制的对象
        for obj in objects:
            builder.ObjectToMoveObject.Add(obj)

        # 提交并获取受影响对象（复制时为新对象）
        feature = builder.Commit()
        try:
            result_objects = builder.GetCommittedObjects()
        except Exception:
            # 某些 NX 版本在 Commit 后必须通过返回的 feature 获取对象：
            try:
                result_objects = feature.GetObjects()  # 视版本 API 不同而定
            except Exception:
                result_objects = []

        builder.Destroy()

        return list(result_objects)

    # 删除对象
    def delete_nx_objects(self, objects):
        """
        使用 NX UpdateManager 删除对象（曲线、体等）。

        Args:
            objects: 单个 NX 对象或对象列表
        """
        theSession = NXOpen.Session.GetSession()

        try:

            # 如果是单个对象，转成列表
            if not isinstance(objects, (list, tuple)):
                objects = [objects]

            # 过滤空对象
            valid_objs = [obj for obj in objects if obj]
            if not valid_objs:
                return False

            # 创建撤销标记
            markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible,"Delete Immovable")  # Delete Immovable 删除不可移动对象

            # 清空可能存在的错误列表
            theSession.UpdateManager.ClearErrorList()

            # 添加对象到删除列表
            nErrs1 = theSession.UpdateManager.AddObjectsToDeleteList(valid_objs)

            # 执行删除
            undo_mark_id = theSession.NewestVisibleUndoMark
            nErrs2 = theSession.UpdateManager.DoUpdate(undo_mark_id)

            # 删除撤销标记
            theSession.DeleteUndoMark(markId1, None)

            return True

        except NXOpen.NXException as ex:
            return False

        except Exception as ex:
            return False

    def move_bodies_point_to_point(self, body_list, from_point, to_point, layer=9, is_copy=True):
        """
        只移动或复制体对象，from_point/to_point 可以是 NXOpen.Point3d 或 (x,y,z)
        :param from_point:  移动起点 NXOpen.Point3d 或 (x, y, z)
        :param to_point:    移动终点 NXOpen.Point3d 或 (x, y, z)
        :param layer:       图层
        :param is_copy:     False=移动（默认），True=复制
        :return: list of committed body objects (复制时为新对象)
        """
        work_part = self.work_part

        # helper: 把输入规范成 NXOpen.Point3d
        def _to_point3d(p):
            if isinstance(p, NXOpen.Point3d):
                return p
            # 支持 list/tuple/其他可迭代
            try:
                x, y, z = p
                return NXOpen.Point3d(float(x), float(y), float(z))
            except Exception:
                raise ValueError("from_point/to_point 必须是 NXOpen.Point3d 或 可迭代三元组 (x,y,z)")

        # 创建 MoveObjectBuilder
        builder = work_part.BaseFeatures.CreateMoveObjectBuilder(
            NXOpen.Features.MoveObject.Null
        )

        # 根据录制宏使用 MoveObjectResult 控制 复制/移动
        if is_copy:
            builder.MoveObjectResult = NXOpen.Features.MoveObjectBuilder.MoveObjectResultOptions.CopyOriginal
        else:
            builder.MoveObjectResult = NXOpen.Features.MoveObjectBuilder.MoveObjectResultOptions.MoveOriginal

        # 保留录制宏的 TransformMotion 配置
        tm = builder.TransformMotion
        tm.DistanceAngle.OrientXpress.AxisOption = NXOpen.GeometricUtilities.OrientXpressBuilder.Axis.Passive
        tm.DistanceAngle.OrientXpress.PlaneOption = NXOpen.GeometricUtilities.OrientXpressBuilder.Plane.Passive
        tm.OrientXpress.AxisOption = NXOpen.GeometricUtilities.OrientXpressBuilder.Axis.Passive
        tm.OrientXpress.PlaneOption = NXOpen.GeometricUtilities.OrientXpressBuilder.Plane.Passive

        tm.DeltaEnum = NXOpen.GeometricUtilities.ModlMotion.Delta.ReferenceAcsWorkPart
        tm.Option = NXOpen.GeometricUtilities.ModlMotion.Options.PointToPoint

        builder.Layer = layer

        # 处理输入点（支持 NXOpen.Point3d 或 三元组）
        from_p3d = _to_point3d(from_point)
        to_p3d = _to_point3d(to_point)

        # 创建 Point 对象（Journal 中是 CreatePoint），并移除参数
        fromPt = work_part.Points.CreatePoint(from_p3d)
        work_part.Points.RemoveParameters(fromPt)

        toPt = work_part.Points.CreatePoint(to_p3d)
        work_part.Points.RemoveParameters(toPt)

        tm.FromPoint = fromPt
        tm.ToPoint = toPt

        # 添加要移动/复制的体对象
        for body in body_list:
            builder.ObjectToMoveObject.Add(body)

        # 提交并获取受影响对象（复制时为新对象）
        feature = builder.Commit()
        try:
            result_objects = builder.GetCommittedObjects()
        except Exception:
            # 某些 NX 版本在 Commit 后必须通过返回的 feature 获取对象：
            try:
                result_objects = feature.GetObjects()  # 视版本 API 不同而定
            except Exception:
                result_objects = []

        builder.Destroy()

        return list(result_objects)
