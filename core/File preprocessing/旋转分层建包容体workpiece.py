import time
import NXOpen
import NXOpen.Features
import NXOpen.GeometricUtilities
import NXOpen.CAM
import NXOpen.UF
import NXOpen.Layer
import os
import traceback

# 如果你的环境中没有 scipy，请注释掉下面这行
try:
    from scipy.__config__ import CONFIG
except ImportError:
    pass

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
    time.sleep(1)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    dir_name, file_name = os.path.split(part_path)
    name, ext = os.path.splitext(file_name)
    save_path = os.path.join(dir_name, f"{name}_{timestamp}{ext}")
    work_part.SaveAs(save_path)
    print(f"保存至: {save_path}", "SUCCESS")
    return save_path

def close_part(part=None):
    theSession = NXOpen.Session.GetSession()
    try:
        if part and hasattr(part, 'Close'):
            try:
                part.Close(NXOpen.BasePart.CloseWholeTree.TrueValue, NXOpen.BasePart.CloseModified.UseResponses, None)
                print(f"已关闭部件: {getattr(part, 'Name', str(part))}")
            except Exception as e:
                print(f"关闭部件 (Close API) 时出错: {e}")
        return True
    except Exception as e:
        print(f"关闭部件时出错: {e}")
        return False

def find_body_by_features(work_part):
    """通过遍历特征找到体，只返回图层为20的体"""
    features_to_mirror = []
    for feat in work_part.Features:
        if hasattr(feat, 'GetBodies', ) and len(feat.GetBodies()) > 0:
            if feat.FeatureType != "MIRROR":
                features_to_mirror.append(feat)

    print(f"找到 {len(features_to_mirror)} 个符合条件的特征")

    if len(features_to_mirror) == 0:
        print("❌ 未找到符合条件的体特征")
        return None

    for feature in features_to_mirror:
        try:
            bodies = feature.GetBodies()
            for body in bodies:
                if body.Layer == 20:
                    print(f"✓ 找到图层20的体: {body.Name} (来自特征: {feature.Name})")
                    return body
        except Exception as e:
            print(f"❌ 获取特征 {feature.Name} 的体时出错: {e}")
            continue

    print("❌ 未找到图层为20的体")
    return None

def create_tooling_box_from_body(work_part: NXOpen.Part, target_body: NXOpen.Body):
    """根据目标实体自动创建包容体"""
    the_session = NXOpen.Session.GetSession()
    mark_id = the_session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "创建包容体")

    tooling_box_builder = work_part.Features.ToolingFeatureCollection.CreateToolingBoxBuilder(
        NXOpen.Features.ToolingBox.Null
    )
    tooling_box_builder.Type = NXOpen.Features.ToolingBoxBuilder.Types.BoundedBlock

    for offset in [
        tooling_box_builder.OffsetPositiveX, tooling_box_builder.OffsetNegativeX,
        tooling_box_builder.OffsetPositiveY, tooling_box_builder.OffsetNegativeY,
        tooling_box_builder.OffsetPositiveZ, tooling_box_builder.OffsetNegativeZ,
        tooling_box_builder.RadialOffset, tooling_box_builder.Clearance,
    ]:
        offset.SetFormula("0")

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
    tooling_box_builder.SetSelectedOccurrences([target_body], [])
    tooling_box_builder.CalculateBoxSize()

    tooling_box_feature = tooling_box_builder.Commit()
    tooling_box_builder.Destroy()
    the_session.SetUndoMarkName(mark_id, "包容体创建完成")

    bodies = tooling_box_feature.GetBodies()
    if bodies and len(bodies) > 0:
        print(f"✅ 成功创建包容体 (临时用于定位)")
        # 注意：这里的包容体是特征体，需要返回其主体 (body)
        return bodies[0] 
    else:
        print("❌ 包容体创建失败")
        return None

def left_down_point(body):
    theUfSession = NXOpen.UF.UFSession.GetUFSession()
    bbox = theUfSession.ModlGeneral.AskBoundingBox(body.Tag)
    # 返回最小X, 最小Y, 最大Z (作为起点和安全平面 Z 方向的参考)
    return (bbox[0], bbox[1], bbox[5]) 

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


def body_centroid(body, prefer_mass=False):
    """返回体的质心点（若可用使用质量质心，否则退回到包围盒中心）。
    prefer_mass: 如果为 True，会尝试通过 UF 的质量属性接口获取质量中心；若失败回退到 bbox 中心。
    注意：不同 NX 版本下 UF 的质量属性接口名可能不同，这里尝试安全调用并降级处理。
    """
    if body is None:
        return NXOpen.Point3d(0.0, 0.0, 0.0)

    if prefer_mass:
        try:
            uf = NXOpen.UF.UFSession.GetUFSession()
            # 尝试常见的 UF 接口名，若不存在会抛异常并走到下面的返回
            # UF 的返回格式随版本不同，这里作宽松处理：如果返回包含质心坐标就使用
            if hasattr(uf.Modl, 'AskMassProps'):
                props = uf.Modl.AskMassProps(body.Tag)
            elif hasattr(uf.Modl, 'AskMassProperties'):
                props = uf.Modl.AskMassProperties(body.Tag)
            else:
                props = None

            if props:
                # 常见返回结构： [mass, cgx, cgy, cgz, ...] 或者类似，做最小长度检查
                if hasattr(props, '__len__') and len(props) >= 4:
                    return NXOpen.Point3d(float(props[1]), float(props[2]), float(props[3]))
        except Exception:
            pass

    # 兜底使用包围盒中心
    return bbox_center_of_body(body)

def find_face_parallel_to_xy(body, extreme_type='min'):
    session = NXOpen.UF.UFSession.GetUFSession()
    found_face = None
    extreme_value = float('inf') if extreme_type == 'min' else float('-inf')

    # 遍历所有面，寻找Z方向最极端的平面
    for face in body.GetFaces():
        if face.SolidFaceType == NXOpen.Face.FaceType.Planar:
            try:
                # 获取面的边界框，用于估算 Z 坐标
                bbox = session.ModlGeneral.AskBoundingBox(face.Tag)
                z_min, z_max = bbox[2], bbox[5]
                
                # 检查是否为水平面 (Z范围很小)
                if abs(z_max - z_min) < 0.001: 
                    current_z = z_min
                    if extreme_type == 'min' and current_z < extreme_value:
                        extreme_value = current_z
                        found_face = face
                    elif extreme_type == 'max' and current_z > extreme_value:
                        extreme_value = current_z
                        found_face = face
            except NXOpen.NXException:
                continue
    return found_face


def read_machining_directions_from_csv(part_name):
    """
    从几何分析CSV报告中读取需要加工的方向，返回集合如 {'+Z','-X',...}
    """
    try:
        import config
        from path_manager import init_path_manager
        pm = init_path_manager(config.FILE_INPUT_PRT, config.FILE_INPUT_DXF)
        geo_dir = pm.get_analysis_geo_dir()
        path1 = os.path.join(str(geo_dir), f"{part_name}.prt.csv")
        path2 = os.path.join(str(geo_dir), f"{part_name}.csv")
        csv_path = None
        if os.path.exists(path1):
            csv_path = path1
        elif os.path.exists(path2):
            csv_path = path2
    except Exception:
        csv_path = None

    if not csv_path:
        return None

    needed = set()
    try:
        import csv
        with open(csv_path, 'r', encoding='utf-8-sig', errors='ignore') as f:
            reader = csv.reader(f)
            headers = next(reader, None)
            if not headers:
                return None
            headers = [h.strip() for h in headers if h.strip()]
            col_map = {}
            for i, h in enumerate(headers):
                if h in ['+Z', '-Z', '+X', '-X', '+Y', '-Y']:
                    col_map[h] = i

            has_data = {d: False for d in col_map.keys()}
            for row in reader:
                for direction, col_idx in col_map.items():
                    if col_idx < len(row) and row[col_idx].strip():
                        has_data[direction] = True

            for direction, has in has_data.items():
                if has:
                    needed.add(direction)
        return needed
    except Exception:
        return None

def switch_to_manufacturing():
    """切换到加工环境"""
    try:
        session = NXOpen.Session.GetSession()
        work_part = session.Parts.Work
        # 检查核心对象有效性
        if not session:
            print("会话对象无效", "ERROR")
            return False

        # 检查是否已经在制造模块
        module_name = session.ApplicationName
        if module_name != "UG_APP_MANUFACTURING":
            print(f"正在从 {module_name} 切换到 UG_APP_MANUFACTURING...", "INFO")
            session.ApplicationSwitchImmediate("UG_APP_MANUFACTURING")
            time.sleep(0.1)  # 短暂等待模块切换完成
        
        # 初始化 CAM 会话
        if not session.IsCamSessionInitialized():
            print("CAM 会话未初始化，正在启动...", "INFO")
            session.CreateCamSession()
            time.sleep(0.1)  # 等待初始化完成
            
        # 确保 Setup 存在
        cam_setup_ready = False
        try:
            if work_part.CAMSetup is not None and work_part.CAMSetup.IsInitialized():
                cam_setup_ready = True
                print("CAM Setup 已存在", "SUCCESS")
        except Exception as e:
            print(f"检查 CAMSetup 时出错: {e}", "WARN")

        if not cam_setup_ready:
            # 尝试创建默认 Setup，优先使用mill_contour更适合铣削操作
            print("正在创建 CAM Setup...", "INFO")
            setup_created = False
            for setup_type in ["mill_contour", "mill_planar", "hole_making"]:
                try:
                    work_part.CreateCamSetup(setup_type)
                    print(f"✅ CAM Setup ({setup_type}) 创建成功。", "SUCCESS")
                    setup_created = True
                    break
                except Exception as e:
                    print(f"⚠ 创建 {setup_type} Setup 失败: {e}", "WARN")
            
            if not setup_created:
                print("❌ 所有类型的 Setup 创建均失败", "ERROR")
                return False
        
        print("已切换到加工环境", "SUCCESS")
        return True
    except Exception as e:
        print(f"切换加工环境失败: {e}", "ERROR")
        traceback.print_exc()
        return False

def create_mcs_with_safe_plane(work_part, tooling_box, points, mcs_name="MCS_1", safe_distance=1.0):
    """创建MCS坐标系并设置安全平面"""
    switch_to_manufacturing()
    # 用包容体的顶面来计算安全平面
    top_face = find_face_parallel_to_xy(tooling_box, "max")
    if not top_face:
        print("⚠ 未找到包容体顶面，无法创建安全平面")
        return None

    try:
        existing = work_part.CAMSetup.CAMGroupCollection.FindObject(f"GEOMETRY/{mcs_name}")
        if existing:
            existing.Delete()
            print(f"已删除同名 MCS: {mcs_name}")
    except:
        pass

    try:
        geom_group = work_part.CAMSetup.CAMGroupCollection.FindObject("GEOMETRY")
        if geom_group is None:
            return None
            
        mcs_group = work_part.CAMSetup.CAMGroupCollection.CreateGeometry(
            geom_group, "mill_contour", "MCS",
            NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue, mcs_name
        )
        builder = work_part.CAMSetup.CAMGroupCollection.CreateMillOrientGeomBuilder(mcs_group)
        
        # 使用包容体的左下角作为坐标系原点 (Xmin, Ymin, Zmax)
        origin3 = NXOpen.Point3d(points[0], points[1], points[2]) 
        x_dir = NXOpen.Vector3d(1.0, 0.0, 0.0)
        y_dir = NXOpen.Vector3d(0.0, 1.0, 0.0)
        xform = work_part.Xforms.CreateXform(origin3, x_dir, y_dir, NXOpen.SmartObject.UpdateOption.AfterModeling, 1.0)
        csys = work_part.CoordinateSystems.CreateCoordinateSystem(xform, NXOpen.SmartObject.UpdateOption.AfterModeling)
        builder.Mcs = csys
        
        # 设置安全平面
        builder.TransferClearanceBuilder.ClearanceType = NXOpen.CAM.NcmClearanceBuilder.ClearanceTypes.Plane
        plane_safe = work_part.Planes.CreatePlane(NXOpen.Point3d(0.0, 0.0, 0.0), NXOpen.Vector3d(0.0, 0.0, 1.0),
                                                  NXOpen.SmartObject.UpdateOption.AfterModeling)
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

def create_cam_workpiece(work_part, mcs_name, part_body, blank_body=None, workpiece_name="WORKPIECE"):
    """
    在指定 MCS 下创建 CAM 几何体 (WORKPIECE)。
    part_body: 加工体
    blank_body: 毛坯体。传 None 则不设置毛坯。
    """
    print(f"开始创建工件: {workpiece_name} (父级 MCS: {mcs_name})")
    uf_session = NXOpen.UF.UFSession.GetUFSession()

    try:
        # 查找父级MCS
        orient_geometry = None
        try:
            orient_geometry = work_part.CAMSetup.CAMGroupCollection.FindObject(f"GEOMETRY/{mcs_name}")
        except:
            orient_geometry = work_part.CAMSetup.CAMGroupCollection.FindObject(mcs_name)

        if orient_geometry is None:
            print(f"❌ 找不到指定的MCS父组: {mcs_name}")
            return None

        # 检查重名工件
        try:
            existing = orient_geometry.FindObject(workpiece_name)
            if existing:
                uf_session.Obj.DeleteObject(existing.Tag)
                time.sleep(0.1)
        except:
            pass

        # 创建 WORKPIECE 几何体组
        nc_group = work_part.CAMSetup.CAMGroupCollection.CreateGeometry(
            orient_geometry, 
            "mill_contour", 
            "WORKPIECE",
            NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue, 
            workpiece_name
        )
        geom_builder = work_part.CAMSetup.CAMGroupCollection.CreateMillGeomBuilder(nc_group)
        sc_rule_factory = work_part.ScRuleFactory

        # ---------------- 设置加工体 (Part) ----------------
        if part_body:
            print("  正在设置加工几何体...")
            geom_builder.PartGeometry.InitializeData(False)
            geometry_set = geom_builder.PartGeometry.GeometryList.FindItem(0)
            rule_opt = sc_rule_factory.CreateRuleOptions()
            rule_opt.SetSelectedFromInactive(False)
            
            body_dumb_rule = sc_rule_factory.CreateRuleBodyDumb([part_body], True, rule_opt)
            rule_opt.Dispose()
            
            sc_collector = geometry_set.ScCollector
            sc_collector.ReplaceRules([body_dumb_rule], False)
        
        # ---------------- 设置毛坯体 (Blank) - 关键：检查 blank_body ----------------
        if blank_body:
            print("  正在设置毛坯几何体...")
            geom_builder.BlankGeometry.InitializeData(False)
            geometry_set_blank = geom_builder.BlankGeometry.GeometryList.FindItem(0)
            
            rule_opt2 = sc_rule_factory.CreateRuleOptions()
            rule_opt2.SetSelectedFromInactive(False)
            
            body_dumb_rule2 = sc_rule_factory.CreateRuleBodyDumb([blank_body], True, rule_opt2)
            rule_opt2.Dispose()
            
            sc_collector2 = geometry_set_blank.ScCollector
            sc_collector2.ReplaceRules([body_dumb_rule2], False)
        else:
            print("  ℹ️ 跳过毛坯设置 (用户未指定毛坯)")

        nx_obj = geom_builder.Commit()
        geom_builder.Destroy()
        print(f"✅ CAM工件几何体创建完成: {workpiece_name}")
        return nx_obj

    except Exception as e:
        print(f"❌ 创建工件时出错: {e}")
        return None

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
        # 尝试从 committed_objects 中提取 Body 实例，保证上层代码得到实际的 body
        bodies_out = []
        try:
            for obj in committed_objects:
                try:
                    if isinstance(obj, NXOpen.Body):
                        bodies_out.append(obj)
                    elif hasattr(obj, 'GetBodies'):
                        try:
                            bs = obj.GetBodies()
                            if bs:
                                for b in bs:
                                    if isinstance(b, NXOpen.Body):
                                        bodies_out.append(b)
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            bodies_out = []

        moveBuilder.Destroy()
        theSession.SetUndoMarkName(markId, f"旋转完成: {angle_degrees}度")
        if bodies_out:
            print(f"旋转产生 {len(bodies_out)} 个 body 返回")
            return bodies_out
        else:
            # 回退：返回原始 committed_objects，便于调用方进一步诊断
            print(f"⚠ 未从 committed_objects 中提取到 body，返回原始 committed_objects (count={len(committed_objects)})")
            return committed_objects
    except Exception as e:
        print(f"❌ 旋转操作失败: {e}")
        moveBuilder.Destroy()
        return None

def set_work_layer(layer_number):
    try:
        theSession = NXOpen.Session.GetSession()
        workPart = theSession.Parts.Work
        # 隐藏除当前层外的所有层
        stateArray = [NXOpen.Layer.StateInfo(layer_number, NXOpen.Layer.State.WorkLayer)]
        workPart.Layers.ChangeStates(stateArray, True) # True: 不可见状态设置为隐藏
        print(f"已将工作图层设置为: {layer_number}")
        return True
    except Exception as ex:
        print(f"设置工作图层时出错: {ex}")
        return False

# ----------------- 旋转操作封装 -----------------
# 注意：这里我们使用 CopyOriginal 创建副本
def rotate_x_minus_90(bodies):
    if not bodies:
        return None
    axis_origin = bbox_center_of_body(bodies[0])
    print(f"显式传入轴心 (bbox center) 给 X轴负90: ({axis_origin.X:.6f}, {axis_origin.Y:.6f}, {axis_origin.Z:.6f})")
    return rotate_bodies_by_object(bodies, -90, NXOpen.Vector3d(1.0, 0.0, 0.0), axis_origin=axis_origin, layer=30, undo_mark_name="X_L30")


def rotate_x_plus_90(bodies):
    if not bodies:
        return None
    axis_origin = bbox_center_of_body(bodies[0])
    print(f"显式传入轴心 (bbox center) 给 X轴正90: ({axis_origin.X:.6f}, {axis_origin.Y:.6f}, {axis_origin.Z:.6f})")
    return rotate_bodies_by_object(bodies, 90, NXOpen.Vector3d(1.0, 0.0, 0.0), axis_origin=axis_origin, layer=40, undo_mark_name="X_L40")


def rotate_y_minus_90(bodies):
    if not bodies:
        return None
    axis_origin = bbox_center_of_body(bodies[0])
    print(f"显式传入轴心 (bbox center) 给 Y轴负90: ({axis_origin.X:.6f}, {axis_origin.Y:.6f}, {axis_origin.Z:.6f})")
    return rotate_bodies_by_object(bodies, -90, NXOpen.Vector3d(0.0, 1.0, 0.0), axis_origin=axis_origin, layer=50, undo_mark_name="Y_L50")


def rotate_y_plus_90(bodies):
    if not bodies:
        return None
    axis_origin = bbox_center_of_body(bodies[0])
    print(f"显式传入轴心 (bbox center) 给 Y轴正90: ({axis_origin.X:.6f}, {axis_origin.Y:.6f}, {axis_origin.Z:.6f})")
    return rotate_bodies_by_object(bodies, 90, NXOpen.Vector3d(0.0, 1.0, 0.0), axis_origin=axis_origin, layer=60, undo_mark_name="Y_L60")


def rotate_y_minus_180(bodies):
    if not bodies:
        return None
    axis_origin = bbox_center_of_body(bodies[0])
    print(f"显式传入轴心 (bbox center) 给 Y轴180: ({axis_origin.X:.6f}, {axis_origin.Y:.6f}, {axis_origin.Z:.6f})")
    return rotate_bodies_by_object(bodies, 180, NXOpen.Vector3d(0.0, 1.0, 0.0), axis_origin=axis_origin, layer=70, undo_mark_name="BACK_L70")


def rotate_x_minus_180(bodies):
    if not bodies:
        return None
    axis_origin = bbox_center_of_body(bodies[0])
    print(f"显式传入轴心 (bbox center) 给 X轴180: ({axis_origin.X:.6f}, {axis_origin.Y:.6f}, {axis_origin.Z:.6f})")
    return rotate_bodies_by_object(bodies, 180, NXOpen.Vector3d(1.0, 0.0, 0.0), axis_origin=axis_origin, layer=80, undo_mark_name="X轴旋转180度")

# ----------------- 核心 CAM 几何体创建函数 -----------------
def create_mcs_for_body(work_part, target_body, operation_name, index):
    """为任何实体（原始或旋转）创建包容体、MCS，以及不带毛坯的 WORKPIECE"""
    try:
        body_layer = target_body.Layer
        print(f"  正在处理图层: {body_layer}")
        if not set_work_layer(body_layer):
            return False

        # 1. 创建包容体 (仅用于计算MCS定位和安全平面)
        print(f"  为 {operation_name} 计算 MCS 边界...")
        # 注意：这里创建的包容体是特征体，在建模历史中
        tooling_box = create_tooling_box_from_body(work_part, target_body)

        if tooling_box:
            # 使用包容体来确定 MCS 的原点和安全平面
            points = left_down_point(tooling_box)
            mcs_name = f"{operation_name}_{index}"
            
            # 2. 创建 MCS
            mcs_obj = create_mcs_with_safe_plane(
                work_part, 
                tooling_box, # 用来计算安全平面
                points, 
                mcs_name=mcs_name, 
                safe_distance=1.0
            )
            
            if mcs_obj:
                # 3. 创建 WORKPIECE (关键：blank_body=None)
                workpiece_name = f"WORKPIECE_{index}"
                create_cam_workpiece(
                    work_part, 
                    mcs_name, 
                    part_body=target_body, 
                    blank_body=None,        # <-- 关键：不设置毛坯
                    workpiece_name=workpiece_name
                )
                return True
            else:
                return False
        else:
            return False

    except Exception as e:
        print(f"❌ {operation_name} 处理出错: {e}")
        import traceback
        traceback.print_exc()
        return False

# ----------------- 主流程 -----------------
def process_file_auto(source_path, target_path):
    print("=" * 50)
    print("开始自动处理文件 (逻辑：为每个方向创建一个不带毛坯的 WORKPIECE)")

    part = open_prt_file_simple(source_path)
    if not part:
        return False

    original_body = find_body_by_features(part)
    if not original_body:
        close_part(part)
        return False

    # 尝试读取需要的加工方向（若 CSV 可用）
    part_name = os.path.splitext(os.path.basename(source_path))[0]
    needed_dirs = read_machining_directions_from_csv(part_name)
    if needed_dirs is None:
        print("⚠ 未找到或无法解析几何分析 CSV，默认生成所有方向的包容体")
    else:
        print(f"检测到的需要加工方向: {list(needed_dirs)}")

    bodies_to_rotate = [original_body]
    # 每个操作同时带上方向代码，便于与 CSV 结果匹配
    operations = [
        ("X_L30", rotate_x_minus_90, "-X"),
        ("X_L40", rotate_x_plus_90, "+X"),
        ("Y_L50", rotate_y_minus_90, "-Y"),
        ("Y_L60", rotate_y_plus_90, "+Y"),
        ("BACK_L70", rotate_y_minus_180, "-Z"),
        # ("X轴正180度", rotate_x_minus_180, None),
    ]

    success_count = 0

    # 1. 处理原始实体 (对应 +Z)
    print("\n[第一步] 处理原始实体 (图层 20)")
    if needed_dirs is None or "+Z" in needed_dirs:
        if create_mcs_for_body(part, original_body, "ORIGINAL_DIRECTION", 0):
            success_count += 1
    else:
        print("跳过原始方向 (+Z) 的包容体创建（CSV 指示不需要）")

    # 2. 旋转并处理选定的副本（仅当 CSV 指示需要时）
    for i, (op_name, op_function, dir_code) in enumerate(operations):
        # 当 dir_code 为 None 时，保守处理为总是执行（或按需修改）
        if needed_dirs is not None and dir_code is not None and dir_code not in needed_dirs:
            print(f"跳过 {op_name}（方向 {dir_code} 未在 CSV 中标记）")
            continue

        print(f"\n执行操作: {op_name}")
        rotated_bodies = op_function(bodies_to_rotate)

        if rotated_bodies and len(rotated_bodies) > 0:
            if create_mcs_for_body(part, rotated_bodies[0], op_name, i + 1):
                success_count += 1
        else:
            print(f"❌ {op_name} 旋转操作失败")

    if success_count > 0:
        # 如果 source_path 与 target_path 不同，保存为目标路径；
        # 如果相同，则避免生成一份未命名的时间戳备份，仅直接保存覆盖（留给上层流程统一管理）
        try:
            if os.path.abspath(target_path) != os.path.abspath(source_path):
                save_part(target_path, part)
            else:
                try:
                    part.Save(NXOpen.BasePart.SaveComponents.TrueValue, NXOpen.BasePart.CloseAfterSave.FalseValue)
                    print(f"已保存（覆盖）: {target_path}")
                except Exception as e:
                    print(f"覆盖保存失败，尝试 SaveAs: {e}")
                    save_part(target_path, part)
        except Exception as e:
            print(f"保存步骤发生异常: {e}")

    # 关闭部件以释放文件锁
    try:
        close_part(part)
    except Exception:
        pass
    print(f"🎉 文件处理完成! 成功创建 {success_count} 组 CAM 几何体 (无毛坯)。")
    return success_count > 0

# def process_part(work_part):
#     """
#     供 run_step8.py 调用的接口，直接处理已打开的 work_part
#     """
#     print("=" * 50)
#     print("开始处理部件 (逻辑：为每个方向创建一个不带毛坯的 WORKPIECE)")
    
#     original_body = find_body_by_features(work_part)
#     if not original_body:
#         return False

#     bodies_to_rotate = [original_body]
#     operations = [
#         # (操作名称, 旋转函数)
#         ("X轴负90度-30", rotate_x_minus_90),
#         ("X轴正90度-40", rotate_x_plus_90),
#         ("Y轴负90度-50", rotate_y_minus_90),
#         ("Y轴正90度-60", rotate_y_plus_90),
#         ("Y轴正180度-70", rotate_y_minus_180),
#         # ("X轴正180度-80", rotate_x_minus_180)
#     ]

#     success_count = 0
    
#     # 1. 处理原始实体 (图层 20)
#     print("\n[第一步] 处理原始实体 (图层 20)")
#     if create_mcs_for_body(work_part, original_body, "ORIGINAL_DIRECTION", 0):
#         success_count += 1

#     # 2. 旋转并处理所有副本
#     for i, (op_name, op_function) in enumerate(operations):
#         print(f"\n执行操作: {op_name}")
#         rotated_bodies = op_function(bodies_to_rotate)
        
#         if rotated_bodies and len(rotated_bodies) > 0:
#             # 传入旋转后的副本
#             if create_mcs_for_body(work_part, rotated_bodies[0], op_name, i + 1): 
#                 success_count += 1
#         else:
#             print(f"❌ {op_name} 旋转操作失败")

#     return success_count > 0

def main():
    # --- 请在这里修改你的文件路径 ---
    source_path = r"E:\Desktop\3mian_modified.prt"
    save_path = r"E:\Desktop\3mian_modified2.prt"
    # --------------------------------

    if not os.path.exists(os.path.dirname(source_path)):
        print(f"路径不存在，请检查: {os.path.dirname(source_path)}")
        return

    process_file_auto(source_path, save_path)

if __name__ == '__main__':
    main()