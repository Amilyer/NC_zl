# -*- coding: utf-8 -*-
"""
镜像操作模块
包含镜像曲线、边界检测等功能
"""

import math
import NXOpen
from utils import print_to_info_window, handle_exception, analyze_arc, get_arc_point, safe_origin
from geometry import GeometryHandler
from path_optimization import PathOptimizer
from log_record import ExceptionLogger

class MirrorHandler:
    """镜像操作处理器"""

    def __init__(self, session, work_part):
        self.session = session
        self.work_part = work_part
        self.geometry_handler = GeometryHandler(session, work_part)
        # 初始化日志管理器
        self.logger = ExceptionLogger("error_record.xlsx")
    # def mirror_curves(self,has_y_negative):
    #     """镜像操作，face_Vector=1表示根据yz面镜像"""
    #     try:
    #         face_vector = 1
    #         # 创建镜像曲线构建器
    #         mirrorCurveBuilder1 = self.work_part.Features.CreateMirrorCurveBuilder(NXOpen.Features.Feature.Null)
    #         if has_y_negative[0] is not None:
    #             face_vector = 0
    #         elif has_y_negative[1] is not None:
    #             face_vector = 1
    #         origin1 = NXOpen.Point3d(0.0, 0.0, 0.0)
    #         if face_vector == 0:
    #             normal1 = NXOpen.Vector3d(0.0, 1.0, 0.0)
    #         elif face_vector == 1:
    #             normal1 = NXOpen.Vector3d(1.0, 0.0, 0.0)
    #         plane1 = self.work_part.Planes.CreatePlane(origin1, normal1, NXOpen.SmartObject.UpdateOption.WithinModeling)
    #         mirrorCurveBuilder1.NewPlane = plane1
    #
    #         unit1 = self.work_part.UnitCollection.FindObject("MilliMeter")
    #         expression1 = self.work_part.Expressions.CreateSystemExpressionWithUnits("0", unit1)
    #         expression2 = self.work_part.Expressions.CreateSystemExpressionWithUnits("0", unit1)
    #
    #         mirrorCurveBuilder1.Curve.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.OnlyCurves)
    #
    #         # 选择所有曲线进行镜像
    #         curves1 = []
    #         try:
    #             for curve in self.work_part.Curves:
    #                 curves1.append(curve)
    #         except Exception as e:
    #             self.session.ListingWindow.Open()
    #             self.session.ListingWindow.WriteLine(f"⚠️ 警告：部分曲线未找到 {e}")
    #
    #         # 创建选择规则
    #         selectionIntentRuleOptions1 = self.work_part.ScRuleFactory.CreateRuleOptions()
    #         selectionIntentRuleOptions1.SetSelectedFromInactive(False)
    #         curveDumbRule1 = self.work_part.ScRuleFactory.CreateRuleBaseCurveDumb(curves1, selectionIntentRuleOptions1)
    #         selectionIntentRuleOptions1.Dispose()
    #
    #         mirrorCurveBuilder1.Curve.AllowSelfIntersection(True)
    #         mirrorCurveBuilder1.Curve.AllowDegenerateCurves(False)
    #
    #         rules1 = [curveDumbRule1]
    #         helpPoint1 = NXOpen.Point3d(0.0, 0.0, 0.0)
    #         mirrorCurveBuilder1.Curve.AddToSection(rules1, NXOpen.NXObject.Null, NXOpen.NXObject.Null,
    #                                                NXOpen.NXObject.Null, helpPoint1, NXOpen.Section.Mode.Create, False)
    #
    #         mirrorCurveBuilder1.PlaneOption = NXOpen.Features.MirrorCurveBuilder.PlaneOptions.NewPlane
    #
    #         # 设置镜像平面
    #         geom1 = []
    #         plane1.SetGeometry(geom1)
    #         plane1.SetMethod(NXOpen.PlaneTypes.MethodType.FixedX)
    #         plane1.Evaluate()
    #
    #         # 执行镜像操作
    #         nXObject1 = mirrorCurveBuilder1.Commit()
    #
    #         mirrorCurveBuilder1.Curve.CleanMappingData()
    #         mirrorCurveBuilder1.Destroy()
    #
    #         # 提取镜像后的曲线对象集合
    #         print_to_info_window("=== 镜像曲线结果 ===")
    #
    #         mirrored_curves = []
    #         if isinstance(nXObject1, NXOpen.Features.MirrorCurve):
    #             mirrored_feature = nXObject1
    #             try:
    #                 entities = mirrored_feature.GetEntities()
    #                 for ent in entities:
    #                     mirrored_curves.append(ent)
    #             except Exception as e:
    #                 print_to_info_window(f"⚠️ 无法获取镜像实体: {e}")
    #         else:
    #             print_to_info_window("❌ 未检测到有效的 MirrorCurve 特征。")
    #
    #         # 清理表达式
    #         try:
    #             self.work_part.Expressions.Delete(expression2)
    #         except NXOpen.NXException as ex:
    #             ex.AssertErrorCode(1050029)
    #         try:
    #             self.work_part.Expressions.Delete(expression1)
    #         except NXOpen.NXException as ex:
    #             ex.AssertErrorCode(1050029)
    #
    #         return mirrored_curves, face_vector
    #
    #     except Exception as e:
    #         return handle_exception("镜像操作失败", str(e))

    def mirror_curves(self, view_name, origin1, circle_list):
        """镜像操作，face_Vector=1表示根据yz面镜像"""
        try:
            # 创建镜像曲线构建器
            mirrorCurveBuilder1 = self.work_part.Features.CreateMirrorCurveBuilder(NXOpen.Features.Feature.Null)
            if view_name == "主视图":
                normal1 = NXOpen.Vector3d(1.0, 0.0, 0.0)
                origin1 = NXOpen.Point3d(origin1[0], origin1[1], origin1[2])
            elif view_name == "正视图":
                normal1 = NXOpen.Vector3d(1.0, 0.0, 0.0)
                origin1 = NXOpen.Point3d(origin1[0], origin1[1], origin1[2])
            elif view_name == "侧视图":
                normal1 = NXOpen.Vector3d(0.0, 1.0, 0.0)
                origin1 = NXOpen.Point3d(origin1[0], origin1[1], origin1[2])
            plane1 = self.work_part.Planes.CreatePlane(origin1, normal1, NXOpen.SmartObject.UpdateOption.WithinModeling)
            mirrorCurveBuilder1.NewPlane = plane1
            unit1 = self.work_part.UnitCollection.FindObject("MilliMeter")
            expression1 = self.work_part.Expressions.CreateSystemExpressionWithUnits("0", unit1)
            expression2 = self.work_part.Expressions.CreateSystemExpressionWithUnits("0", unit1)

            mirrorCurveBuilder1.Curve.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.OnlyCurves)

            # 创建选择规则
            selectionIntentRuleOptions1 = self.work_part.ScRuleFactory.CreateRuleOptions()
            selectionIntentRuleOptions1.SetSelectedFromInactive(False)
            curveDumbRule1 = self.work_part.ScRuleFactory.CreateRuleBaseCurveDumb(circle_list, selectionIntentRuleOptions1)
            selectionIntentRuleOptions1.Dispose()

            mirrorCurveBuilder1.Curve.AllowSelfIntersection(True)
            mirrorCurveBuilder1.Curve.AllowDegenerateCurves(False)

            rules1 = [curveDumbRule1]
            helpPoint1 = NXOpen.Point3d(0.0, 0.0, 0.0)
            mirrorCurveBuilder1.Curve.AddToSection(rules1, NXOpen.NXObject.Null, NXOpen.NXObject.Null,
                                                   NXOpen.NXObject.Null, helpPoint1, NXOpen.Section.Mode.Create, False)

            mirrorCurveBuilder1.PlaneOption = NXOpen.Features.MirrorCurveBuilder.PlaneOptions.NewPlane

            # 设置镜像平面
            geom1 = []
            plane1.SetGeometry(geom1)
            plane1.SetMethod(NXOpen.PlaneTypes.MethodType.FixedX)
            plane1.Evaluate()

            # 执行镜像操作
            nXObject1 = mirrorCurveBuilder1.Commit()

            mirrorCurveBuilder1.Curve.CleanMappingData()
            mirrorCurveBuilder1.Destroy()

            # 提取镜像后的曲线对象集合
            print_to_info_window("=== 镜像曲线结果 ===")

            mirrored_curves = []
            if isinstance(nXObject1, NXOpen.Features.MirrorCurve):
                mirrored_feature = nXObject1
                try:
                    entities = mirrored_feature.GetEntities()
                    for ent in entities:
                        mirrored_curves.append(ent)
                except Exception as e:
                    print_to_info_window(f"⚠️ 无法获取镜像实体: {e}")
            else:
                print_to_info_window("❌ 未检测到有效的 MirrorCurve 特征。")

            # 清理表达式
            try:
                self.work_part.Expressions.Delete(expression2)
            except NXOpen.NXException as ex:
                ex.AssertErrorCode(1050029)
            try:
                self.work_part.Expressions.Delete(expression1)
            except NXOpen.NXException as ex:
                ex.AssertErrorCode(1050029)

            return mirrored_curves

        except Exception as e:
            return handle_exception("镜像操作失败", str(e))




    def judge_side_negative(self,lwh_point):
        """判断是否有侧面加工圆 x或y为0时说明在x轴或y轴负方向，1在正方向"""

        curves = self.geometry_handler.get_all_curves()
        side_x = None
        side_y = None
        for curve in curves:
            if not (hasattr(curve, "CenterPoint") and hasattr(curve, "Radius")):
                continue

            arc = analyze_arc(curve)
            if not arc:
                continue

            c = arc.CenterPoint
            if round(c.Y, 4) < 0.0:
                side_y = 0
            elif round(c.Y, 4) > lwh_point[2] and (abs(round(c.X, 4)) > lwh_point[0] or abs(round(c.Y, 4)) > lwh_point[1]):
                if round(c.X, 4) > lwh_point[2]:
                    side_y = None
                    if round(c.X, 4) < 0.0:
                        side_x = 0
                    elif round(c.X, 4) > lwh_point[2]:
                        side_x = 1
                    return (side_x,side_y)
                side_y = 1
        return (side_x, side_y)

    """三视图加工区域所处象限"""
    def three_view_area(self, center_point, minx, miny, maxx, maxy):
        """第一象限 00 第二象限：01 第三象限 10 第四象限 11"""
        def get_quadrant_code(center_point, view, three_view_bs):
            if center_point.X > 0 and center_point.Y > 0:
                three_view_bs[view] = "00"
            elif center_point.X < 0 and center_point.Y > 0:
                three_view_bs[view] = "01"
            elif center_point.X < 0 and center_point.Y < 0:
                three_view_bs[view] = "10"
            elif center_point.X > 0 and center_point.Y < 0:
                three_view_bs[view] = "11"

        three_view_bs = {
            "主视图" : None,
            "正视图" : None,
            "侧视图" : None
        }
        if minx >= 0:
            # 俯/仰视图区域：长×宽
            if minx <= round(center_point.X, 4) and round(center_point.X, 4) <= maxx and round(center_point.Y, 4) >= miny and round(center_point.Y, 4) <= maxy:
                get_quadrant_code(center_point,"主视图",three_view_bs)
            # 正视图区域：长×高
            elif round(center_point.Y, 4)  < miny:
                get_quadrant_code(center_point, "正视图", three_view_bs)
            # 侧视图区域：宽×高
            elif round(center_point.X, 4) > maxx:
                get_quadrant_code(center_point, "侧视图", three_view_bs)



        return three_view_bs

    def select_boundary_curve(self, lwh_point):
        # 主视图曲线和注释
        main_curves_list = []
        # 正视图曲线
        font_curves_list = []
        # 右视图曲线
        right_curves_list = []
        curves = self.geometry_handler.get_all_curves()
        for curve in curves:
            if isinstance(curve, NXOpen.Line):
                center_point = (curve.StartPoint.X, curve.StartPoint.Y)
            elif isinstance(curve, NXOpen.Arc):
                center_point = (curve.CenterPoint.X, curve.CenterPoint.Y)

            # 俯/仰视图区域：长×宽
            if abs(round(center_point[0], 4)) <= lwh_point[0] and abs(round(center_point[1], 4)) <= lwh_point[1] and center_point[1] >=0:
                main_curves_list.append(curve)
            # 正视图区域：长×高
            elif (round(center_point[1], 4) > lwh_point[1] or round(center_point[1], 4) < 0.0) and abs(round(center_point[0], 4)) <= lwh_point[0]:
                font_curves_list.append(curve)
            # 侧视图区域：宽×高
            elif abs(round(center_point[0], 4)) > lwh_point[0] and abs(round(center_point[1], 4)) <= lwh_point[1]:
                right_curves_list.append(curve)

        return {
            "main_mirror_curves": main_curves_list,
            "front_mirror_curves": font_curves_list,
            "right_mirror_curves": right_curves_list
        }


    # def select_boundary_curve(self, lwh_point, has_y_negative):
    #     """找到侧面离坐标系最近的边的y轴坐标"""
    #     curves = self.geometry_handler.get_all_curves()
    #     po = PathOptimizer()
    #     # y轴负方向的侧视图
    #     if has_y_negative[1] == 0:
    #         for curve in curves:
    #             if (type(curve).__name__ == "Line"):
    #                 if curve.StartPoint.Y < 0.0 and curve.EndPoint.Y < 0.0 and abs(po._euclidean_length(curve) - lwh_point[2]) < 0.1:
    #                     return (None,po._line_min_y_point(curve))
    #     # # y轴正方向的侧视图
    #     # elif has_y_negative[1] == 1:
    #     #     for curve in curves:
    #     #         if (type(curve).__name__ == "Line"):
    #     #             if curve.StartPoint.Y > lwh_point[1] and curve.EndPoint.Y < 0.0 and abs(po._euclidean_length(curve) - lwh_point[2]) < 0.1:
    #     #                 return po._line_min_y_point(curve)
    #     # X轴正方向侧视图
    #     elif has_y_negative[0] == 1:
    #         for curve in curves:
    #             if (type(curve).__name__ == "Line"):
    #                 if curve.StartPoint.X > 0.0 and curve.EndPoint.Y > 0.0 and abs(po._euclidean_length(curve) - lwh_point[2]) < 0.1:
    #                     return (po._line_min_x_point(curve),None)
    #     return None, None

    def mirror_all_body_features_except_mirrors(self, features_to_mirror, mirror_plane="YZ"):
        """
        自动镜像当前部件中所有非镜像的体特征（即能生成几何体的特征）。

        :param mirror_plane: 镜像平面，可选 "XY", "YZ", "ZX"
        :return: 成功镜像的特征对象
        """
        # 遍历获取的特征，逐个进行镜像
        successful_mirrors = []
        for feat in features_to_mirror:
            try:
                mirrored_feat = self.create_mirror_feature(feat, mirror_plane=mirror_plane)  # 逐个传递特征
                if mirrored_feat is not None:
                    successful_mirrors.append(mirrored_feat)
                    print_to_info_window(f"已镜像特征: {mirrored_feat.Tag}")
            except Exception as ex:
                print_to_info_window(f"镜像失败: {feat.Tag} | 错误: {str(ex)}")

        print_to_info_window(f"成功镜像了 {len(successful_mirrors)} 个特征。")
        return successful_mirrors

    def create_mirror_feature(self, feature_to_mirror, mirror_plane="YZ"):
        """
        创建指定特征的镜像副本。

        :param feature_to_mirror: 要镜像的特征对象 (NXOpen.Features.Feature)
        :param mirror_plane: 镜像平面，可选 "XY", "YZ", "ZX"
        :return: 镜像特征对象 (NXOpen.Features.Feature)，失败时返回 None
        :raises ValueError: 如果 mirror_plane 不合法
        """
        workPart = self.work_part

        # 创建镜像构建器
        mirrorBuilder1 = workPart.Features.CreateGeomcopyBuilder(NXOpen.Features.Feature.Null)
        mirrorBuilder1.PatternService.PatternType = NXOpen.GeometricUtilities.PatternDefinition.PatternEnum.Mirror

        # 定义矩阵和法向量
        matrix = NXOpen.Matrix3x3()
        normal = None

        if mirror_plane == "XY":
            matrix.Xx = 1.0
            matrix.Xy = 0.0
            matrix.Xz = 0.0
            matrix.Yx = 0.0
            matrix.Yy = 1.0
            matrix.Yz = 0.0
            matrix.Zx = 0.0
            matrix.Zy = 0.0
            matrix.Zz = 1.0
            normal = NXOpen.Vector3d(0.0, 1.0, 0.0)
        elif mirror_plane == "YZ":
            matrix.Xx = 0.0
            matrix.Xy = 1.0
            matrix.Xz = 0.0
            matrix.Yx = 0.0
            matrix.Yy = 0.0
            matrix.Yz = 1.0
            matrix.Zx = 1.0
            matrix.Zy = 0.0
            matrix.Zz = 0.0
            normal = NXOpen.Vector3d(0.0, 0.0, 1.0)
        elif mirror_plane == "ZX":
            matrix.Xx = 1.0
            matrix.Xy = 0.0
            matrix.Xz = 0.0
            matrix.Yx = 0.0
            matrix.Yy = 0.0
            matrix.Yz = 1.0
            matrix.Zx = 0.0
            matrix.Zy = 1.0
            matrix.Zz = 0.0
            normal = NXOpen.Vector3d(0.0, 1.0, 0.0)
        else:
            raise ValueError("mirror_plane must be one of 'XY', 'YZ', 'ZX'")

        # 创建镜像平面
        origin = NXOpen.Point3d(0.0, 0.0, 0.0)
        plane1 = workPart.Planes.CreatePlane(origin, normal, NXOpen.SmartObject.UpdateOption.WithinModeling)
        plane1.SetMethod(NXOpen.PlaneTypes.MethodType.FixedX)
        plane1.Matrix = matrix
        plane1.SetAlternate(NXOpen.PlaneTypes.AlternateType.One)
        plane1.Evaluate()

        # 设置镜像定义
        mirrorBuilder1.PatternService.MirrorDefinition.NewPlane = plane1
        mirrorBuilder1.PatternService.MirrorDefinition.PlaneOption = NXOpen.GeometricUtilities.MirrorPattern.PlaneOptions.New
        mirrorBuilder1.CsysMirrorOption = NXOpen.Features.MirrorBuilder.CsysMirrorOptions.MirrorZAndX

        try:
            mirrorBuilder1.FeatureList.Add(feature_to_mirror)  # 直接传递特征对象，而不是列表
            # 提交操作
            markId = self.session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "镜像特征")
            nXObject = mirrorBuilder1.Commit()
            mirrorBuilder1.Destroy()
            return nXObject
        except Exception as ex:
            mirrorBuilder1.Destroy()
            print_to_info_window(f"镜像失败: {str(ex)}")
            return None

    def mirror_body(self, body, origin, mirror_plane="YZ", rotate_angle=45):
        """
        创建指定体的镜像副本并返回镜像后的几何体。

        :param body: 要镜像的体对象
        :param mirror_plane: 镜像平面，可选 "XY", "YZ", "ZX"
        :param rotate_angle: 镜像角度（默认为 45 度）
        :return: 镜像后的几何体 (NXOpen.Body)，失败时返回 None
        """

        geomcopyBuilder1 = self.work_part.Features.CreateGeomcopyBuilder(NXOpen.Features.Feature.Null)

        # 设置镜像平面的法向量和矩阵
        matrix = NXOpen.Matrix3x3()
        normal = None

        if mirror_plane == "XY":
            matrix.Xx = 1.0
            matrix.Xy = 0.0
            matrix.Xz = 0.0
            matrix.Yx = 0.0
            matrix.Yy = 1.0
            matrix.Yz = 0.0
            matrix.Zx = 0.0
            matrix.Zy = 0.0
            matrix.Zz = 1.0
            normal = NXOpen.Vector3d(0.0, 1.0, 0.0)
        elif mirror_plane == "YZ":
            matrix.Xx = 0.0
            matrix.Xy = 1.0
            matrix.Xz = 0.0
            matrix.Yx = 0.0
            matrix.Yy = 0.0
            matrix.Yz = 1.0
            matrix.Zx = 1.0
            matrix.Zy = 0.0
            matrix.Zz = 0.0
            normal = NXOpen.Vector3d(0.0, 0.0, 1.0)
        elif mirror_plane == "ZX":
            matrix.Xx = 1.0
            matrix.Xy = 0.0
            matrix.Xz = 0.0
            matrix.Yx = 0.0
            matrix.Yy = 0.0
            matrix.Yz = 1.0
            matrix.Zx = 0.0
            matrix.Zy = 1.0
            matrix.Zz = 0.0
            normal = NXOpen.Vector3d(0.0, 1.0, 0.0)
        else:
            raise ValueError("mirror_plane must be one of 'XY', 'YZ', 'ZX'")

        # 创建镜像平面
        origin = NXOpen.Point3d(origin[0], origin[1], origin[2])
        plane1 = self.work_part.Planes.CreatePlane(origin, normal, NXOpen.SmartObject.UpdateOption.WithinModeling)
        plane1.SetMethod(NXOpen.PlaneTypes.MethodType.FixedX)
        plane1.Matrix = matrix
        plane1.SetAlternate(NXOpen.PlaneTypes.AlternateType.One)
        plane1.Evaluate()

        # 设置镜像定义
        geomcopyBuilder1.MirrorPlane = plane1
        geomcopyBuilder1.Type = NXOpen.Features.GeomcopyBuilder.TransformTypes.Mirror
        geomcopyBuilder1.RotateAngle.SetFormula(str(rotate_angle))
        geomcopyBuilder1.NumberOfCopies.SetFormula("1")

        try:
            geomcopyBuilder1.GeometryToInstance.Add(body)  # 直接传递几何体对象
            # 提交镜像操作
            markId = self.session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "镜像几何体")
            # 获取镜像后的几何体
            mirrored_body = geomcopyBuilder1.CommitFeature().GetBodies()[0]  # 直接获取第一个镜像体
            geomcopyBuilder1.Destroy()
            return mirrored_body  # 返回镜像后的几何体
        except Exception as ex:
            geomcopyBuilder1.Destroy()
            print(f"镜像失败: {str(ex)}")
            return None
