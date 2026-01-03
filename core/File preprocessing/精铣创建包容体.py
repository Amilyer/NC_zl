import time
import traceback
import NXOpen
import NXOpen.Features
import NXOpen.GeometricUtilities
import NXOpen.CAM
import NXOpen.UF
import NXOpen.Layer
import os

import math


# 如果你的环境中没有 scipy，请注释掉下面这行
# try:
#     from scipy.__config__ import CONFIG
# except ImportError:
#     pass

def open_prt_file_simple(prt_path):
    """改进的文件打开函数，确保CAM环境正确初始化"""
    try:
        if not os.path.exists(prt_path):
            print(f"PRT文件不存在: {prt_path}")
            return None

        print(f"部件: {prt_path}")
        session = NXOpen.Session.GetSession()
        base_part, load_status = session.Parts.OpenBaseDisplay(prt_path)
        workPart = session.Parts.Work

        # 切换到制造模块
        try:
            session.ApplicationSwitchImmediate("UG_APP_MANUFACTURING")
            print("✅ 已切换到制造模块")
        except Exception as e:
            print(f"⚠ 制造模块切换警告: {e}")
            try:
                session.ApplicationSwitchImmediate("Manufacturing")
                print("✅ 已通过备用名称切换到制造模块")
            except Exception as e2:
                print(f"❌ 制造模块切换失败: {e2}")

        # 初始化CAM会话
        try:
            uf = NXOpen.UF.UFSession.GetUFSession()
            uf.Cam.InitSession()
            print("✅ CAM会话初始化完成")
        except Exception as e:
            print(f"❌ CAM会话初始化失败: {e}")

        session.Parts.SetDisplay(workPart, False, False)
        session.Parts.SetWork(workPart)
        print(f"成功打开PRT文件: {prt_path}")
        return workPart

    except Exception as e:
        print(f"打开PRT文件时出错: {str(e)}")
        return None

def save_part(part_path, work_part):
    # Remove timestamp logic as requested
    save_path = part_path
    
    # Intelligent Save logic
    try:
        current_path = work_part.FullPath
        if os.path.normpath(save_path).lower() == os.path.normpath(current_path).lower():
             work_part.Save(NXOpen.BasePart.SaveComponents.TrueValue, NXOpen.BasePart.CloseAfterSave.FalseValue)
             print(f"已保存(更新): {save_path}", "SUCCESS")
        else:
             work_part.SaveAs(save_path)
             print(f"已另存为: {save_path}", "SUCCESS")
        return save_path
    except Exception as e:
        print(f"保存失败: {e}")
        return None

def close_part(part=None):
    theSession = NXOpen.Session.GetSession()
    try:
        if part:
            part.Close(NXOpen.BasePart.CloseWholeTree.TrueValue, NXOpen.BasePart.CloseModified.CloseModified, None)
            print(f"已关闭部件: {part.Name}")
        return True
    except Exception as e:
        print(f"关闭部件时出错: {e}")
        return False

def find_body_by_features(work_part):
    """通过遍历特征找到图层为20的体"""
    try:
        features_list = []
        for f in work_part.Features:
            try:
                if hasattr(f, 'GetBodies') and len(f.GetBodies()) > 0 and f.FeatureType != "MIRROR":
                    features_list.append(f)
            except Exception as e:
                print(f"⚠ 遍历特征时出错: {e}")
                continue

        if len(features_list) == 0:
            print("❌ 未找到体特征")
            return None

        for feature in features_list:
            try:
                bodies = feature.GetBodies()
                for body in bodies:
                    if body.Layer == 20:
                        print(f"✓ 找到图层20的体: {body.Name} (来自特征: {feature.Name})")
                        return body
            except Exception as e:
                print(f"⚠ 处理特征 {feature.Name} 时出错: {e}")
                continue

        print("❌ 未找到图层为20的体")
        return None
    except Exception as e:
        print(f"❌ 查找图层20的体时出错: {e}")
        traceback.print_exc()
        return None

def create_tooling_box_from_body(work_part: NXOpen.Part, target_body: NXOpen.Body):
    """根据目标实体自动创建包容体 (仅用于定位)"""
    the_session = NXOpen.Session.GetSession()
    mark_id = the_session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "创建包容体")
    tooling_box_builder = None

    try:
        tooling_box_builder = work_part.Features.ToolingFeatureCollection.CreateToolingBoxBuilder(NXOpen.Features.ToolingBox.Null)
        tooling_box_builder.Type = NXOpen.Features.ToolingBoxBuilder.Types.BoundedBlock

        for offset in [tooling_box_builder.OffsetPositiveX, tooling_box_builder.OffsetNegativeX, tooling_box_builder.OffsetPositiveY, tooling_box_builder.OffsetNegativeY, tooling_box_builder.OffsetPositiveZ, tooling_box_builder.OffsetNegativeZ]:
            offset.SetFormula("0")

        # 设置包容体方向与WCS一致
        matrix = NXOpen.Matrix3x3()
        matrix.Xx, matrix.Xy, matrix.Xz = 1.0, 0.0, 0.0
        matrix.Yx, matrix.Yy, matrix.Yz = 0.0, 1.0, 0.0
        matrix.Zx, matrix.Zy, matrix.Zz = 0.0, 0.0, 1.0
        tooling_box_builder.SetBoxMatrixAndPosition(matrix, NXOpen.Point3d(0.0, 0.0, 0.0))

        rule_options = work_part.ScRuleFactory.CreateRuleOptions()
        rule_options.SetSelectedFromInactive(False)
        body_rule = work_part.ScRuleFactory.CreateRuleBodyDumb([target_body], True, rule_options)
        rule_options.Dispose()

        sc_collector = tooling_box_builder.BoundedObject
        sc_collector.ReplaceRules([body_rule], False)
        tooling_box_builder.CalculateBoxSize()

        tooling_box_feature = tooling_box_builder.Commit()
        the_session.SetUndoMarkName(mark_id, "包容体创建完成")

        bodies = tooling_box_feature.GetBodies()
        if bodies and len(bodies) > 0:
            print(f"✅ 成功创建包容体 (用于定位)")
            return bodies[0] 
        else:
            print("❌ 包容体创建失败")
            return None
    
    except Exception as e:
        print(f"❌ 创建包容体失败: {e}")
        traceback.print_exc()
        # 回滚操作
        the_session.UndoToMark(mark_id, False)
        return None
    
    finally:
        # 确保Builder被销毁
        if tooling_box_builder:
            tooling_box_builder.Destroy()

def find_face_parallel_to_xy(body, extreme_type='max'):
    """寻找Z方向最极端的水平面（用于安全平面）"""
    session = NXOpen.UF.UFSession.GetUFSession()
    found_face = None
    extreme_value = float('-inf') if extreme_type == 'max' else float('inf')

    try:
        faces = body.GetFaces()
        for face in faces:
            try:
                if face.SolidFaceType == NXOpen.Face.FaceType.Planar:
                    try:
                        bbox = session.ModlGeneral.AskBoundingBox(face.Tag)
                        z_min, z_max = bbox[2], bbox[5]
                        if abs(z_max - z_min) < 0.001: 
                            current_z = z_max if extreme_type == 'max' else z_min
                            if ((extreme_type == 'max' and current_z > extreme_value) or 
                               (extreme_type == 'min' and current_z < extreme_value)):
                                extreme_value = current_z
                                found_face = face
                    except Exception as e:
                        print(f"  ⚠ 获取面边界框时出错: {e}")
            except Exception as e:
                print(f"  ⚠ 检查面类型时出错: {e}")
                continue
    except Exception as e:
        print(f"❌ 获取面列表时出错: {e}")
        traceback.print_exc()
    
    return found_face

def print_to_info_window(message):
    """将消息输出到NX的信息窗口和日志文件"""
    theSession = NXOpen.Session.GetSession()
    theSession.ListingWindow.Open()
    theSession.ListingWindow.WriteLine(str(message))
    theSession.LogFile.WriteLine(str(message))

def left_down_point(body):
    theUfSession = NXOpen.UF.UFSession.GetUFSession()
    bbox = theUfSession.ModlGeneral.AskBoundingBox(body.Tag)
    # 返回最小X, 最小Y, 最大Z (作为起点和安全平面 Z 方向的参考)
    return (bbox[0], bbox[1], bbox[5]) 

def ask_arc_center_abs(edge_tag: int):
    """
    返回绝对坐标系圆心 [x,y,z] 和半径 r
    与 C++ 完全一致，2312 亲测
    """
    uf = NXOpen.UF.UFSession.GetUFSession()   # 必须加这句！！！
    evaluator = uf.Eval.Initialize(edge_tag)
    try:
        arc_obj = uf.Eval.AskArc(evaluator)   # 无参版本，返回结构体
        # print_to_info_window(f"arc_obj:{arc_obj}")
        # 用 dir() 当场看字段名
        if hasattr(arc_obj, 'Center'):
            center = arc_obj.Center          # 有的版本叫 Center
        elif hasattr(arc_obj, 'center'):
            center = arc_obj.center          # 有的版本叫 center
        else:                                # 再不行就按索引 0 取
            center = arc_obj[0]

        if hasattr(arc_obj, 'Radius'):
            radius = arc_obj.Radius
        elif hasattr(arc_obj, 'radius'):
            radius = arc_obj.radius
        else:
            radius = arc_obj[6]

        # 确保 center 是 tuple/list
        if isinstance(center, (tuple, list)) and len(center) >= 3:
            return tuple(center[:3]), radius
        else:
            raise RuntimeError('无法解析圆心字段')
    finally:
        # 2312 没有 Free，也没有 Close，用 Dispose 模式
        if hasattr(evaluator, 'Dispose'):
            evaluator.Dispose()
        # 保险起见再置空
        evaluator = None
 
def find_red_cyl_face_center(body, color_index=186, prefer_lower_z=False):
    """
    寻找红色圆柱面端面圆心坐标：
    1、先找到红色孔面（非盲孔）
    2、通过圆弧边获取圆心和最大/小z
    
    :param prefer_lower_z=False 选择上端面
    """
    if body is None:
        return None

    # 备选：Body 颜色
    body_color = -1
    try:
        body_color = body.Color
    except:
        pass

    for face in body.GetFaces():
        # 判断面的颜色是否正确
        # 原有颜色逻辑（一模一样）
        face_color = -1
        try:
            face_color = face.Color
        except:
            pass
        if face_color <= 0 or face_color is None:
            face_color = body_color
        if face_color != color_index:
            continue
        
        if face.SolidFaceType != NXOpen.Face.FaceType.Cylindrical:
            continue
        
        # 判断是否为孔面
        hole_data_tuple = face.GetHoleData()    
        if not hole_data_tuple:
            continue
        hole_data, is_hole = hole_data_tuple
        # 通过边获取孔面的一个圆心和最大z
        point = None
        # 记录是否已获取一个圆心、是否已存在闭合圆弧
        is_closed = False
        final_z = None
        for edge in face.GetEdges():
            if edge.SolidEdgeType != NXOpen.Edge.EdgeType.Circular:
                continue
            # 通过圆弧端点获取最大、小z，并判断是否闭合
            vertices = edge.GetVertices()
            if (not is_closed and abs(vertices[0].X - vertices[1].X) < 0.001 
                and abs(vertices[0].Y - vertices[1].Y) < 0.001 and abs(vertices[0].Z - vertices[1].Z) < 0.001):
                    try:
                        # 解析闭合圆弧边获取圆心数据
                        center, radius = ask_arc_center_abs(edge.Tag)
                        point = [center[0], center[1], center[2]]
                        is_closed = True
                    except Exception as e: 
                        print(f"{e}:无法解析圆心字段，请寻找下一个圆弧")
                        continue
            # 更新最大/小z
            if not final_z:
                final_z = point[2]
            else:
                final_z = (min(final_z, vertices[0].Z, vertices[1].Z) 
                            if prefer_lower_z else max(final_z, vertices[0].Z, vertices[1].Z))
        # 未获取圆心或未找到闭合圆弧线，舍弃
        if not point:
            print_to_info_window(f"警告：未找到圆柱面{'下' if prefer_lower_z else '上'}端面圆心，寻找下一个红色圆柱面")
            continue
        # 判断孔类型，只保留贯穿孔（排除盲孔）
        if hole_data.GetDepthLimit() != NXOpen.ResizeHoleData.Depthlimit.ThroughNext:
            continue
        print(f"获取孔的{'上' if prefer_lower_z == False else '下'}端面圆心：{point[0]:.6f}, {point[1]:.6f}, {point[2]:.6f}")  
        return point
    print_to_info_window("仍未找到红色圆柱面")
    return None

def initialize_cam_environment(work_part):
    try:
        if work_part.CAMSetup is None:
            print("初始化CAM环境...")
            cam_setup = work_part.CAMSetups.CreateSetup(
                NXOpen.CAM.CAMSetup.CAMSetupType.Mill,
                NXOpen.CAM.CAMSetup.GeneralToolpathOutputType.ProgramAndToolLocation,
                "mill_contour"
            )
            print("✅ CAM环境初始化完成")
            return cam_setup
        else:
            print("✅ CAM环境已存在")
            return work_part.CAMSetup
    except Exception as e:
        print(f"❌ CAM环境初始化失败: {e}")
        return None

def create_mcs_with_safe_plane(work_part, tooling_box, points, mcs_name="MCS_20", safe_distance=1.0):
    """创建MCS坐标系并设置安全平面"""
    session = NXOpen.Session.GetSession()
    if not ensure_cam_setup_ready(session, work_part): return None

    # 用包容体的顶面来计算安全平面
    top_face = find_face_parallel_to_xy(tooling_box, "max")
    if not top_face:
        print("⚠ 未找到包容体顶面，无法创建安全平面")
        return None

    try:
        existing = work_part.CAMSetup.CAMGroupCollection.FindObject(f"GEOMETRY/{mcs_name}")
        if existing: existing.Delete()
    except: pass

    try:
        geom_group = work_part.CAMSetup.CAMGroupCollection.FindObject("GEOMETRY")
        if geom_group is None: return None
            
        mcs_group = work_part.CAMSetup.CAMGroupCollection.CreateGeometry(
            geom_group, "mill_contour", "MCS",
            NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue, mcs_name
        )
        builder = work_part.CAMSetup.CAMGroupCollection.CreateMillOrientGeomBuilder(mcs_group)
        
        # 使用传入的点作为坐标系原点
        origin3 = NXOpen.Point3d(points[0], points[1], points[2]) 
        x_dir = NXOpen.Vector3d(1.0, 0.0, 0.0)
        y_dir = NXOpen.Vector3d(0.0, 1.0, 0.0)
        xform = work_part.Xforms.CreateXform(origin3, x_dir, y_dir, NXOpen.SmartObject.UpdateOption.AfterModeling, 1.0)
        csys = work_part.CoordinateSystems.CreateCoordinateSystem(xform, NXOpen.SmartObject.UpdateOption.AfterModeling)
        builder.Mcs = csys
        
        # 设置安全平面
        builder.TransferClearanceBuilder.ClearanceType = NXOpen.CAM.NcmClearanceBuilder.ClearanceTypes.Plane
        
        # 创建临时平面用于安全平面设置
        plane_safe = work_part.Planes.CreatePlane(NXOpen.Point3d(0.0, 0.0, 0.0), NXOpen.Vector3d(0.0, 0.0, 1.0), NXOpen.SmartObject.UpdateOption.AfterModeling)
        plane_safe.SetMethod(NXOpen.PlaneTypes.MethodType.Distance)
        plane_safe.SetGeometry([top_face])
        expr = plane_safe.Expression
        expr.RightHandSide = str(safe_distance)
        plane_safe.SetAlternate(NXOpen.PlaneTypes.AlternateType.One)
        plane_safe.Evaluate()
        builder.TransferClearanceBuilder.PlaneXform = plane_safe

        nx_obj = builder.Commit()
        builder.Destroy()
        print(f"✅ MCS 创建完成: {mcs_name}")
        return nx_obj

    except Exception as e:
        print(f"❌ 创建MCS时出错: {e}")
        return None

def create_cam_workpiece(work_part, parent_group, part_body, blank_body=None, workpiece_name="WORKPIECE_20"):
    """
    在指定 MCS 下创建 CAM 几何体 (WORKPIECE)。
    parent_group: 父级 MCS 组对象
    part_body: 加工体
    blank_body: 毛坯体。传 None 则不设置毛坯。
    """
    print(f"开始创建工件: {workpiece_name} (父级: {parent_group.Name if parent_group else '未知'})")
    uf_session = NXOpen.UF.UFSession.GetUFSession()
    geom_builder = None

    try:
        if parent_group is None:
            print(f"❌ 父级组对象为空")
            return None

        # 检查重名工件并删除
        try:
            existing = parent_group.FindObject(workpiece_name)
            if existing:
                print(f"  发现重名工件 {workpiece_name}，正在删除...")
                uf_session.Obj.DeleteObject(existing.Tag)
                time.sleep(0.1)
        except Exception as e:
            print(f"  ⚠ 检查重名工件时出错: {e}")

        # 创建 WORKPIECE 几何体组
        try:
            nc_group = work_part.CAMSetup.CAMGroupCollection.CreateGeometry(
                parent_group, "mill_contour", "WORKPIECE",
                NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue, workpiece_name
            )
        except Exception as e:
            print(f"  ❌ 创建几何体组失败: {e}")
            return None

        # 创建几何体构建器
        try:
            geom_builder = work_part.CAMSetup.CAMGroupCollection.CreateMillGeomBuilder(nc_group)
            sc_rule_factory = work_part.ScRuleFactory
        except Exception as e:
            print(f"  ❌ 创建几何体构建器失败: {e}")
            return None

        # ---------------- 设置加工体 (Part) ----------------
        if part_body:
            print("  正在设置加工几何体...")
            try:
                geom_builder.PartGeometry.InitializeData(False)
                geometry_set = geom_builder.PartGeometry.GeometryList.FindItem(0)
                rule_opt = sc_rule_factory.CreateRuleOptions()
                rule_opt.SetSelectedFromInactive(False)
                body_dumb_rule = sc_rule_factory.CreateRuleBodyDumb([part_body], True, rule_opt)
                rule_opt.Dispose()
                sc_collector = geometry_set.ScCollector
                sc_collector.ReplaceRules([body_dumb_rule], False)
                print("  ✅ 加工几何体设置完成")
            except Exception as e:
                print(f"  ❌ 设置加工几何体失败: {e}")
        
        # ---------------- 设置毛坯体 (Blank) - 关键：不设置 ----------------
        if blank_body:
            print("  正在设置毛坯几何体...")
            # ... (如果您将来需要设置毛坯，在这里添加逻辑)
        else:
            print("  ℹ️ 跳过毛坯设置 (blank_body=None)")

        # 提交并销毁构建器
        try:
            nx_obj = geom_builder.Commit()
            print(f"✅ CAM工件几何体创建完成: {workpiece_name}")
            return nx_obj
        except Exception as e:
            print(f"❌ 提交几何体构建器失败: {e}")
            return None

    except Exception as e:
        print(f"❌ 创建工件时出错: {e}")
        traceback.print_exc()
        return None
    finally:
        # 确保构建器被销毁
        if geom_builder:
            try:
                geom_builder.Destroy()
            except:
                pass

def bbox_center_of_body(body):
    """返回体的包围盒中心，类型为 NXOpen.Point3d。
    作为默认的“轴心”计算方法，稳健且总是可用。
    """
    try:
        uf = NXOpen.UF.UFSession.GetUFSession()
        bbox = uf.ModlGeneral.AskBoundingBox(body.Tag)
        cx = (bbox[0] + bbox[3]) / 2.0
        cy = (bbox[1] + bbox[4]) / 2.0
        cz = (bbox[2] + bbox[5]) / 2.0
        return NXOpen.Point3d(cx, cy, cz)
    except Exception:
        # 兜底：返回原点，尽量不抛异常破坏流程
        return NXOpen.Point3d(0.0, 0.0, 0.0)

def rotate_bodies_by_object(
    bodies,
    angle_degrees=-90,
    axis_direction=NXOpen.Vector3d(0.0, 1.0, 0.0),
    axis_origin=None,
        move_result=NXOpen.Features.MoveObjectBuilder.MoveObjectResultOptions.CopyOriginal,
        layer=1,
        undo_mark_name=None
):
    print(f"开始旋转 {len(bodies)} 个体，目标图层: {layer}")
    theSession = NXOpen.Session.GetSession()
    workPart = theSession.Parts.Work

    if undo_mark_name is None:
        undo_mark_name = f"旋转几何体 {angle_degrees}度"
    markId = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, undo_mark_name)

    moveBuilder = workPart.BaseFeatures.CreateMoveObjectBuilder(NXOpen.Features.MoveObject.Null)
    moveBuilder.TransformMotion.Option = NXOpen.GeometricUtilities.ModlMotion.Options.Angle
    moveBuilder.TransformMotion.DeltaEnum = NXOpen.GeometricUtilities.ModlMotion.Delta.ReferenceAcsWorkPart
    moveBuilder.TransformMotion.Angle.SetFormula(str(angle_degrees))

    for param in [moveBuilder.TransformMotion.DistanceValue, moveBuilder.TransformMotion.DeltaXc,
                  moveBuilder.TransformMotion.DeltaYc, moveBuilder.TransformMotion.DeltaZc]:
        param.SetFormula("0")
    
    # 如果未提供轴心，优先尝试使用体的质量质心作为轴心
    if axis_origin is None:
        try:
            if bodies and len(bodies) > 0:
                # 默认使用包围盒中心作为轴心（稳健且总可用）
                bbox_ct = bbox_center_of_body(bodies[0])
                axis_origin = bbox_ct
                print(f"  使用包围盒中心作为轴心: ({axis_origin.X:.6f}, {axis_origin.Y:.6f}, {axis_origin.Z:.6f})")
            else:
                axis_origin = NXOpen.Point3d(0.0, 0.0, 0.0)
        except Exception:
            axis_origin = NXOpen.Point3d(0.0, 0.0, 0.0)

    # 规范 axis_origin 的类型：支持 tuple/list、NXOpen.Body 或 NXOpen.Point3d
    if isinstance(axis_origin, (tuple, list)):
        axis_origin = NXOpen.Point3d(axis_origin[0], axis_origin[1], axis_origin[2])
    elif isinstance(axis_origin, NXOpen.Body):
        axis_origin = bbox_center_of_body(axis_origin)
    elif axis_origin is None:
        # 若未传入轴心，使用第一个体的包围盒中心作为轴心
        axis_origin = bbox_center_of_body(bodies[0]) if bodies and len(bodies) > 0 else NXOpen.Point3d(0.0, 0.0, 0.0)

    # 打印用于调试的轴点和方向，方便在 NX 中验证
    try:
        print(f"  旋转轴点: ({axis_origin.X:.6f}, {axis_origin.Y:.6f}, {axis_origin.Z:.6f}), 方向: ({axis_direction.X:.6f}, {axis_direction.Y:.6f}, {axis_direction.Z:.6f})")
    except Exception:
        pass

    # 创建 Direction 对象，将轴点和轴方向信息编码进去
    # 参考 NX_Drilling_Automation2/geometry.py 的方式：CreateDirection(Point3d, Vector3d, UpdateOption)
    direction = workPart.Directions.CreateDirection(axis_origin, axis_direction, NXOpen.SmartObject.UpdateOption.WithinModeling)
    
    # 创建 Axis 对象，第一个参数传 Point.Null（轴点已在 Direction 中），第二个参数是 Direction
    # 参考同样的文件：CreateAxis(Point.Null, direction, UpdateOption)
    axis = workPart.Axes.CreateAxis(NXOpen.Point.Null, direction, NXOpen.SmartObject.UpdateOption.WithinModeling)
    
    moveBuilder.TransformMotion.AngularAxis = axis

    for body in bodies:
        moveBuilder.ObjectToMoveObject.Add(body)

    moveBuilder.MoveObjectResult = move_result
    moveBuilder.LayerOption = NXOpen.Features.MoveObjectBuilder.LayerOptionType.AsSpecified
    moveBuilder.Layer = layer
    moveBuilder.NumberOfCopies = 1

    try:
        moveBuilder.Commit()
        committed_objects = moveBuilder.GetCommittedObjects()
        moveBuilder.Destroy()
        theSession.SetUndoMarkName(markId, f"旋转完成: {angle_degrees}度")
        return committed_objects
    except Exception as e:
        print(f"❌ 旋转操作失败: {e}")
        moveBuilder.Destroy()
        return None

def set_work_layer(layer_number):
    """设置工作图层并隐藏其他图层"""
    try:
        theSession = NXOpen.Session.GetSession()
        workPart = theSession.Parts.Work
        stateArray = [NXOpen.Layer.StateInfo(layer_number, NXOpen.Layer.State.WorkLayer)]
        workPart.Layers.ChangeStates(stateArray, True)
        print(f"已将工作图层设置为: {layer_number}")
        return True
    except Exception as ex:
        print(f"设置工作图层时出错: {ex}")
        return False

def ensure_cam_setup_ready(the_session, work_part):
    """
    智能准备 CAM 环境 (修复 'Current part does not contain valid setup' 错误)
    """
    try:
        # 检查输入参数有效性
        if not the_session:
            print("❌ 会话对象无效")
            return False
            
        if not work_part:
            print("❌ 工作部件无效")
            return False

        # 1. 检查 CAM 会话
        if not the_session.IsCamSessionInitialized():
            print("CAM 会话未初始化，正在启动...")
            the_session.CreateCamSession()
            time.sleep(0.1)  # 短暂等待初始化完成

        # 2. 检查 Setup 是否存在
        # 尝试访问 CAMSetup，如果未初始化或不存在，通常需要在 try 块中处理
        cam_setup_ready = False
        
        try:
            if work_part.CAMSetup is not None:
                print("✅ CAM Setup 已初始化")
                cam_setup_ready = True
        except Exception as e:
            print(f"⚠ 检查 CAMSetup 时出错: {e}")
            # 继续向下尝试创建

        # 3. 创建 Setup，优先使用mill_contour更适合铣削操作
        if not cam_setup_ready:
            print("当前部件没有有效的 Setup，正在自动创建 CAM 环境...")
            setup_created = False
            for setup_type in ["mill_contour", "mill_planar", "hole_making"]:
                try:
                    work_part.CreateCamSetup(setup_type)
                    print(f"✅ CAM Setup ({setup_type}) 创建成功。")
                    setup_created = True
                    break
                except Exception as e:
                    print(f"⚠ 创建 {setup_type} Setup 失败: {e}")
            
            if not setup_created:
                print("❌ 所有类型的 Setup 创建均失败")
                return False
                
        return True

    except Exception as ex:
        print(f"❌ 自动创建 CAM Setup 失败: {ex}")
        traceback.print_exc()
        return False

# ----------------- 旋转操作封装 -----------------
# 注意：这里我们使用 CopyOriginal 创建副本
def rotate_x_minus_90(bodies):
    return rotate_bodies_by_object(bodies, -90, NXOpen.Vector3d(1.0, 0.0, 0.0), layer=30, undo_mark_name="X轴旋转-90度-30")
def rotate_x_plus_90(bodies):
    return rotate_bodies_by_object(bodies, 90, NXOpen.Vector3d(1.0, 0.0, 0.0), layer=40, undo_mark_name="X轴旋转+90度-40")
def rotate_y_minus_90(bodies):
    return rotate_bodies_by_object(bodies, -90, NXOpen.Vector3d(0.0, 1.0, 0.0), layer=50, undo_mark_name="Y轴旋转-90度-50")
def rotate_y_plus_90(bodies):
    return rotate_bodies_by_object(bodies, 90, NXOpen.Vector3d(0.0, 1.0, 0.0), layer=60, undo_mark_name="Y轴旋转+90度-60")
def rotate_y_minus_180(bodies):
    return rotate_bodies_by_object(bodies, 180, NXOpen.Vector3d(0.0, 1.0, 0.0), layer=70, undo_mark_name="y轴旋转180度-70")
def rotate_x_minus_180(bodies):
    return rotate_bodies_by_object(bodies, 180, NXOpen.Vector3d(1.0, 0.0, 0.0), layer=80, undo_mark_name="X轴旋转180度-80")

# ----------------- 核心 CAM 几何体和MCS坐标系创建函数 -----------------
def create_mcs_and_workpiece_for_body(work_part, target_body, operation_name, index):
    """
    为目标实体创建 MCS 和不带毛坯的 WORKPIECE
    该函数是外部调用的主要入口
    """
    
    body_layer = target_body.Layer
    
    try:
        print(f"  正在处理图层: {body_layer}")
        set_work_layer(body_layer)

        # 1. 创建包容体 (仅用于计算安全平面)
        print(f"  为 {operation_name} 计算 MCS 边界...")
        tooling_box = create_tooling_box_from_body(work_part, target_body)

        # 2. 创建 MCS
        if tooling_box:
            # 正Z/负Z方向（图层20、70）：MCS坐标原点设为割孔上表面圆心；若无割孔，则坐标系原点设为包容体左上角
            # 其他方向：MCS坐标原点设为包容体左上角
            if body_layer == 20 or body_layer == 70:
                points = find_red_cyl_face_center(target_body)
                if not points:
                    points = left_down_point(tooling_box)
            else:
                points = left_down_point(tooling_box)

            mcs_name = f"MCS_{operation_name}_{index}"
            mcs_obj = create_mcs_with_safe_plane(
                work_part, 
                tooling_box, 
                points, 
                mcs_name=mcs_name, 
                safe_distance=1.0
            )
            
            if mcs_obj:
                # 3. 创建 WORKPIECE (不设置毛坯)
                workpiece_name = f"WORKPIECE_{index}"
                wp_obj = create_cam_workpiece(
                    work_part, 
                    mcs_obj,  # 直接传递对象
                    part_body=target_body, 
                    blank_body=None,        # <-- 关键：不设置毛坯
                    workpiece_name=workpiece_name
                )
                
                # 4. 删除临时包容体
                try:
                    theSession = NXOpen.Session.GetSession()
                    delete_mark_id = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "删除临时包容体")
                    theSession.DeleteObject(tooling_box)
                    theSession.UpdateManager.DoUpdate(delete_mark_id)
                    print(f"✅ 临时包容体已删除")
                except Exception as e:
                    print(f"⚠ 删除临时包容体失败: {e}")
                
                return wp_obj is not None
            else:
                return False
        else:
            return False

    except Exception as e:
        print(f"❌ {operation_name} 处理出错: {e}")
        traceback.print_exc()
        return False

# ----------------- 主流程 -----------------
# ----------------- 方向分析逻辑 (CSV版) -----------------
import csv

def read_machining_directions_from_csv(part_name):
    """
    从几何分析CSV报告中读取需要加工的方向
    
    Args:
        part_name: 部件名称 (不含路径和扩展名)
        
    Returns:
        set: 包含需要加工的方向集合 {'+Z', '-Z', '+X', '-X', '+Y', '-Y'}
        如果未找到文件或读取失败，返回 None (默认全做)
    """
    needed_directions = set()
    
    # 构造 CSV 路径
    # 尝试从 config 获取路径，或者使用默认路径结构
    csv_path = None
    try:
        import config
        # 假设 config 中有相关配置，或者根据已知结构拼接
        # 结构: output/03_Analysis/Geometry_Analysis
        project_root = getattr(config, 'PROJECT_ROOT', None)
        if project_root:
            # 优先尝试 Geometry_Analysis
            temp_path = os.path.join(str(project_root), "output", "03_Analysis", "Geometry_Analysis", f"{part_name}.csv")
            # 兼容 prt.csv 命名 (例如 DIE-05.prt.csv)
            temp_path_prt = os.path.join(str(project_root), "output", "03_Analysis", "Geometry_Analysis", f"{part_name}.prt.csv")
            
            if os.path.exists(temp_path):
                csv_path = temp_path
            elif os.path.exists(temp_path_prt):
                csv_path = temp_path_prt
    except ImportError:
        pass
        
    if not csv_path or not os.path.exists(csv_path):
        print(f"  ⚠ 未找到几何分析报告 (CSV): {csv_path or part_name}")
        return None
        
    print(f"  正在读取加工方向报告: {csv_path}")
    
    try:
        with open(csv_path, 'r', encoding='utf-8-sig', errors='ignore') as f:
            reader = csv.reader(f)
            headers = next(reader, None)
            
            if not headers:
                print("  ❌ CSV 文件为空")
                return None
                
            # 清理表头 (去除空格)
            headers = [h.strip() for h in headers if h.strip()]
            
            # 建立列索引映射: Direction -> Column Index
            # 期望表头包含: +Z, -Z, +X, -X, +Y, -Y
            col_map = {}
            for i, h in enumerate(headers):
                if h in ['+Z', '-Z', '+X', '-X', '+Y', '-Y']:
                    col_map[h] = i
            
            # 检查每一列是否有数据
            has_data = {d: False for d in col_map.keys()}
            
            for row in reader:
                for direction, col_idx in col_map.items():
                    if col_idx < len(row):
                        val = row[col_idx].strip()
                        if val: # 如果有值 (Face ID)
                            has_data[direction] = True
            
            # 统计结果
            for direction, has in has_data.items():
                if has:
                    needed_directions.add(direction)
                    
            print(f"  检测到的加工方向 (来自CSV): {list(needed_directions)}")
            return needed_directions

    except Exception as e:
        print(f"❌ 读取 CSV 失败: {e}")
        return None

# ----------------- 主流程 -----------------
def process_file_auto(source_path, target_path):
    print("=" * 50)
    print("开始自动处理文件 (逻辑：为每个方向创建一个不带毛坯的 WORKPIECE)")
    
    part = open_prt_file_simple(source_path)
    if not part: return False

    original_body = find_body_by_features(part)
    if not original_body:
        close_part(part)
        return False

    # 获取部件名称
    part_name = os.path.splitext(os.path.basename(source_path))[0]

    # 分析加工方向 (读取 CSV)
    needed_dirs = read_machining_directions_from_csv(part_name)
    
    # 定义操作映射
    # Map: (Operation Name, Function, Required Direction on Original Body)
    # CSV 中的方向: +Z, -Z, +X, -X, +Y, -Y
    # 注意: CSV 方向基于原始坐标系
    # 旋转后朝上的面 (Z+) 对应 原始体的哪个方向?
    # verify mapping:
    # rotate_x_minus_90 (-90 X): Y- -> Z+  => Needs -Y
    # rotate_x_plus_90  (+90 X): Y+ -> Z+  => Needs +Y
    # rotate_y_minus_90 (-90 Y): X+ -> Z+  => Needs +X
    # rotate_y_plus_90  (+90 Y): X- -> Z+  => Needs -X
    # rotate_y_minus_180(180 Y): Z- -> Z+  => Needs -Z
    # rotate_x_minus_180(180 X): Z- -> Z+  => Needs -Z
    
    all_operations = [
        ("X轴负90度", rotate_x_minus_90, "-X"),
        ("X轴正90度", rotate_x_plus_90, "+X"),
        ("Y轴负90度", rotate_y_minus_90, "-Y"),
        ("Y轴正90度", rotate_y_plus_90, "+Y"),
        ("Z轴负180度", rotate_y_minus_180, "-Z")
    ]

    bodies_to_rotate = [original_body]
    success_count = 0
    
    # 1. 处理原始实体 (图层 20) -> 对应 +Z
    # 如果 needed_dirs 不为 None 且不包含 +Z，是否跳过? 
    # 原则上原始方向通常是主要方向，但如果 CSV 明确说没面，可以跳过?
    # 按照用户要求 "看看对应的方向下是否有面id即可了"，如果 +Z 没面，也应该跳过。
    
    run_original = True
    if needed_dirs is not None and "+Z" not in needed_dirs:
        print("  ℹ️ 跳过 原始方向 (方向 +Z 无加工面)")
        run_original = False
    
    if run_original:
        print("\n[第一步] 处理原始实体 (图层 20)")
        if create_mcs_and_workpiece_for_body(part, original_body, "ORIGINAL_DIRECTION", 0):
            success_count += 1

    # 2. 旋转并处理所有副本 (根据分析结果筛选)
    for i, (op_name, op_function, req_dir) in enumerate(all_operations):
        # 筛选逻辑
        if needed_dirs is not None:
             # CSV 里的方向标记如果不在 needed_dirs 里，说明没面，跳过
            if req_dir not in needed_dirs:
                print(f"  ℹ️ 跳过 {op_name} (方向 {req_dir} 无加工面)")
                continue
        
        print(f"\n执行操作: {op_name}")
        rotated_bodies = op_function(bodies_to_rotate)
        
        if rotated_bodies and len(rotated_bodies) > 0:
            # 传入旋转后的副本
            if create_mcs_and_workpiece_for_body(part, rotated_bodies[0], op_name, i + 1): 
                success_count += 1
        else:
            print(f"❌ {op_name} 旋转操作失败")

    final_path = source_path
    if success_count > 0:
        saved = save_part(target_path, part)
        if saved:
            final_path = saved
    
    close_part(part)
    print(f"🎉 文件处理完成! 成功创建 {success_count} 组 CAM 几何体 (无毛坯)。")
    return success_count > 0, final_path

def main():
    # --- 请在这里修改你的文件路径 ---
    source_path = r"C:\Users\admin\Desktop\test-mcs-hole.prt"
    save_path = r"C:\Users\admin\Desktop\result\model2.prt"
    # --------------------------------

    if not os.path.exists(os.path.dirname(source_path)):
        print(f"路径不存在，请检查: {os.path.dirname(source_path)}")
        return

    process_file_auto(source_path, save_path)

if __name__ == '__main__':
    main()