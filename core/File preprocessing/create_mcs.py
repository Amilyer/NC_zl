# -*- coding: utf-8 -*-
"""
MCS与工件创建模块 (create_mcs.py)
功能：
1. 查找图层20的实体
2. 创建包容体和MCS坐标系
3. 设置安全平面
4. 创建CAM工件几何体(WORKPIECE)
"""

import time
import traceback
import NXOpen
import NXOpen.Features
import NXOpen.GeometricUtilities
import NXOpen.CAM
import NXOpen.UF
import NXOpen.Layer

# ============================================================================
# 🔧 实用函数：几何体和图层操作
# ============================================================================

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

def left_down_point(body):
    """获取包容体的最小XYZ点"""
    try:
        theUfSession = NXOpen.UF.UFSession.GetUFSession()
        bbox = theUfSession.ModlGeneral.AskBoundingBox(body.Tag)
        # 返回 Xmin, Ymin, Zmax (作为 MCS 原点和安全平面的 Z 参考)
        return (bbox[0], bbox[1], bbox[5])
    except Exception as e:
        print(f"❌ 获取包容体边界框时出错: {e}")
        traceback.print_exc()
        return None 

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

# ============================================================================
# 🔧 CAM 操作函数
# ============================================================================

def ensure_cam_setup_ready(the_session, work_part):
    """
    智能准备 CAM 环境 (简化版，采用用户提供逻辑)

    行为:
    1. 如果 CAM 会话未初始化，调用 Session.CreateCamSession()
    2. 如果当前部件已有已初始化的 CAM Setup，直接返回 True
    3. 否则自动创建一个 'hole_making' 类型的 CAM Setup 并返回 True
    """
    def print_to_info_window(msg: str):
        try:
            ss = NXOpen.Session.GetSession()
            lw = ss.ListingWindow
            lw.Open()
            lw.WriteLine(msg)
        except Exception:
            try:
                print(msg)
            except Exception:
                pass

    try:
        if not the_session:
            print_to_info_window("❌ 会话对象无效")
            return False

        if not work_part:
            print_to_info_window("❌ 工作部件无效")
            return False

        # 1. 检查 CAM 会话
        try:
            if not the_session.IsCamSessionInitialized():
                print_to_info_window("CAM 会话未初始化，正在启动...")
                the_session.CreateCamSession()
        except Exception:
            # 有些 NX 版本下 IsCamSessionInitialized 可能不存在，尝试直接创建
            try:
                the_session.CreateCamSession()
            except Exception as e:
                print_to_info_window(f"⚠ 无法通过 Session.CreateCamSession() 初始化: {e}")

        # 2. 检查 Setup 是否存在
        try:
            if work_part.CAMSetup is not None and work_part.CAMSetup.IsInitialized():
                return True
        except Exception:
            # 继续尝试创建
            pass

        # 3. 创建 Setup（使用 hole_making，针对钻孔场景更稳妥）
        print_to_info_window("当前部件没有有效的 Setup，正在自动创建 'hole_making' 环境...")
        try:
            work_part.CreateCamSetup("hole_making")
            print_to_info_window("✅ CAM Setup (hole_making) 创建成功。")
            return True
        except Exception as ex:
            print_to_info_window(f"❌ 自动创建 CAM Setup 失败: {ex}")
            traceback.print_exc()
            return False

    except Exception as ex:
        print_to_info_window(f"❌ 自动创建 CAM Setup 失败: {ex}")
        traceback.print_exc()
        return False

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
        
        # 使用包容体的左下角作为坐标系原点 (Xmin, Ymin, Zmax)
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

# ============================================================================
# 🚀 暴露的公共接口
# ============================================================================

def create_mcs_and_workpiece_for_body(work_part, target_body):
    """
    为图层 20 实体创建 MCS 和不带毛坯的 WORKPIECE
    该函数是外部调用的主要入口
    """
    
    body_layer = target_body.Layer
    if body_layer != 20:
         print(f"❌ 实体不在图层 20，操作终止。")
         return False

    operation_name = "ORIGINAL"
    
    try:
        print(f"  正在处理图层: {body_layer}")
        set_work_layer(body_layer)

        # 1. 创建包容体 (仅用于计算MCS定位和安全平面)
        print(f"  为 {operation_name} 计算 MCS 边界...")
        tooling_box = create_tooling_box_from_body(work_part, target_body)

        if tooling_box:
            # 2. 创建 MCS
            points = left_down_point(tooling_box)
            mcs_name = "MCS_1"
            mcs_obj = create_mcs_with_safe_plane(
                work_part, 
                tooling_box, 
                points, 
                mcs_name=mcs_name, 
                safe_distance=1.0
            )
            
            if mcs_obj:
                # 3. 创建 WORKPIECE (不设置毛坯)
                workpiece_name = "WORKPIECE_1"
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
