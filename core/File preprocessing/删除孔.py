import NXOpen
import NXOpen.Features
import NXOpen.GeometricUtilities
import NXOpen.UF
import NXOpen.Layer
import os
import traceback
# ==============================================================================
# 全局设置图层
# ==============================================================================

SOURCE_LAYER = 1  # 源数据图层
TARGET_LAYER = 20  # 处理图层


# ==============================================================================
# 工具函数
# ==============================================================================

def print_to_info_window(message):
    print(str(message))
    try:
        session = NXOpen.Session.GetSession()
        try:
            session.LogFile.WriteLine(str(message))
        except:
            pass
    except:
        pass


def save_part(part):
    """保存部件"""
    try:
        part.Save(NXOpen.BasePart.SaveComponents.TrueValue, NXOpen.BasePart.CloseAfterSave.FalseValue)
        print_to_info_window("部件保存成功")
        return True
    except Exception as ex:
        print_to_info_window(f"保存失败: {ex}")
        return False


def copy_bodies_from_source_to_target_layer(source_layer, target_layer):
    """复制源图层实体到目标图层"""
    try:
        theSession = NXOpen.Session.GetSession()
        workPart = theSession.Parts.Work
        all_bodies = workPart.Bodies
        bodies_to_copy = [body for body in all_bodies if body.Layer == source_layer]

        if not bodies_to_copy:
            print_to_info_window(f"图层 {source_layer} 为空，无法复制。")
            return 0

        workPart.Layers.CopyObjects(target_layer, bodies_to_copy)
        print_to_info_window(f"已将 {len(bodies_to_copy)} 个实体从图层 {source_layer} 复制到 {target_layer}")
        return len(bodies_to_copy)
    except Exception as ex:
        print_to_info_window(f"图层复制失败: {ex}")
        return 0


def set_work_layer(layer_number):
    """设置工作图层"""
    try:
        theSession = NXOpen.Session.GetSession()
        workPart = theSession.Parts.Work
        uf_session = NXOpen.UF.UFSession.GetUFSession()
        original_work_layer = uf_session.Layer.AskWorkLayer()
        stateArray = [NXOpen.Layer.StateInfo(layer_number, NXOpen.Layer.State.WorkLayer)]
        workPart.Layers.ChangeStates(stateArray, False)
        return original_work_layer
    except:
        return 1


def hide_all_layers_except_work_layer(work_layer):
    """隐藏非工作图层"""
    try:
        uf_session = NXOpen.UF.UFSession.GetUFSession()
        for layer in range(1, 257):
            if layer != work_layer:
                uf_session.Layer.SetStatus(layer, NXOpen.UF.UFConstants.UF_LAYER_INACTIVE_LAYER)
    except:
        pass


def restore_all_layers():
    """恢复显示所有图层"""
    try:
        uf_session = NXOpen.UF.UFSession.GetUFSession()
        for layer in range(1, 257):
            uf_session.Layer.SetStatus(layer, NXOpen.UF.UFConstants.UF_LAYER_ACTIVE_LAYER)
    except:
        pass


def get_bodies_on_target_layer(workPart, layer_index):
    """获取指定图层的实体"""
    target_bodies = []
    for body in workPart.Bodies:
        if body.Layer == layer_index:
            target_bodies.append(body)
    return target_bodies


# ==============================================================================
# 核心几何算法
# ==============================================================================

def merge_sort_fast(arr):
    temp = [None] * len(arr)
    _merge_sort(arr, temp, 0, len(arr) - 1)
    return arr


def _merge_sort(arr, temp, left, right):
    if left >= right: return
    mid = (left + right) // 2
    _merge_sort(arr, temp, left, mid)
    _merge_sort(arr, temp, mid + 1, right)
    _merge(arr, temp, left, mid, right)


def _merge(arr, temp, left, mid, right):
    i, j = left, mid + 1
    k = left
    while i <= mid and j <= right:
        if arr[i][1] <= arr[j][1]:
            temp[k] = arr[i];
            i += 1
        else:
            temp[k] = arr[j];
            j += 1
        k += 1
    while i <= mid: temp[k] = arr[i]; i += 1; k += 1
    while j <= right: temp[k] = arr[j]; j += 1; k += 1
    for x in range(left, right + 1): arr[x] = temp[x]


def _get_tagged_object(tag_or_obj):
    try:
        if isinstance(tag_or_obj, (int,)):
            return NXOpen.TaggedObjectManager.GetTaggedObject(tag_or_obj)
    except:
        pass
    return tag_or_obj


def _ensure_edge(obj, workPart):
    if isinstance(obj, NXOpen.Edge): return obj
    obj = _get_tagged_object(obj)
    if isinstance(obj, NXOpen.Face):
        try:
            edges = obj.GetEdges()
            if edges: return edges[0]
        except:
            pass
    if isinstance(obj, NXOpen.Body):
        try:
            faces = obj.GetFaces()
            if faces:
                for f in faces:
                    try:
                        es = f.GetEdges()
                        if es: return es[0]
                    except:
                        continue
        except:
            pass
    return None


def _ensure_face(obj):
    o = _get_tagged_object(obj)
    if isinstance(o, NXOpen.Face): return o
    return None


def to_tuple(pt):
    return (round(pt[0], 1), round(pt[1], 1), round(pt[2], 1))


def create_extrude_from_edge(workPart, edge, start_target_face, target_body=None):
    theSession = NXOpen.Session.GetSession()
    edge_obj = _ensure_edge(edge, workPart)
    face_obj = _ensure_face(start_target_face)
    if target_body is None:
        tb = [NXOpen.Body.Null]
    else:
        tb = [target_body]

    extrudeBuilder = None
    try:
        extrudeBuilder = workPart.Features.CreateExtrudeBuilder(NXOpen.Features.Feature.Null)
        section = workPart.Sections.CreateSection(0.00095, 0.001, 0.050000000000000003)
        extrudeBuilder.Section = section
        extrudeBuilder.AllowSelfIntersectingSection(True)
        section.AllowSelfIntersection(True)
        section.AllowDegenerateCurves(False)
        section.SetAllowedEntityTypes(NXOpen.Section.AllowTypes.OnlyCurves)
        section.DistanceTolerance = 0.001
        section.ChainingTolerance = 0.00095

        sel_opts = workPart.ScRuleFactory.CreateRuleOptions()
        sel_opts.SetSelectedFromInactive(False)
        edge_rule = workPart.ScRuleFactory.CreateRuleEdgeDumb([edge_obj], sel_opts)
        sel_opts.Dispose()

        help_point = NXOpen.Point3d(0.0, 0.0, 0.0)
        try:
            verts = edge_obj.GetVertices()
            if verts: help_point = verts[0].Point
        except:
            pass

        section.AddToSection([edge_rule], edge_obj, NXOpen.NXObject.Null, NXOpen.NXObject.Null,
                             help_point, NXOpen.Section.Mode.Create, False)

        origin = NXOpen.Point3d(0.0, 0.0, 0.0)
        vec = NXOpen.Vector3d(0.0, 0.0, 1.0)
        direction = workPart.Directions.CreateDirection(origin, vec, NXOpen.SmartObject.UpdateOption.WithinModeling)
        extrudeBuilder.Direction = direction
        extrudeBuilder.DistanceTolerance = 0.001
        extrudeBuilder.Offset.StartOffset.SetFormula("0")
        extrudeBuilder.Offset.EndOffset.SetFormula("0.1")
        extrudeBuilder.Draft.FrontDraftAngle.SetFormula("2")
        extrudeBuilder.Draft.BackDraftAngle.SetFormula("2")
        extrudeBuilder.Limits.StartExtend.TrimType = NXOpen.GeometricUtilities.Extend.ExtendType.UntilExtended
        extrudeBuilder.Limits.StartExtend.Value.SetFormula("42")
        extrudeBuilder.Limits.EndExtend.Value.SetFormula("0")
        extrudeBuilder.Limits.StartExtend.Target = face_obj
        extrudeBuilder.BooleanOperation.SetTargetBodies(tb)
        extrudeBuilder.BooleanOperation.Type = NXOpen.GeometricUtilities.BooleanOperation.BooleanType.Unite

        sv = extrudeBuilder.SmartVolumeProfile
        sv.OpenProfileSmartVolumeOption = False
        sv.CloseProfileRule = NXOpen.GeometricUtilities.SmartVolumeProfileBuilder.CloseProfileRuleType.Fci
        extrudeBuilder.ParentFeatureInternal = False

        feature = extrudeBuilder.CommitFeature()
        try:
            extrudeBuilder.Destroy()
        except:
            pass
        return feature

    except Exception:
        if extrudeBuilder:
            try:
                extrudeBuilder.Destroy()
            except:
                pass
        raise


def delete_faces(faces_to_delete):
    theSession = NXOpen.Session.GetSession()
    workPart = theSession.Parts.Work
    deleteFaceBuilder = workPart.Features.CreateDeleteFaceBuilder(NXOpen.Features.Feature.Null)
    fr = deleteFaceBuilder.FaceRecognized
    fr.CoplanarEnabled = False
    fr.CoplanarAxesEnabled = False
    fr.CoaxialEnabled = False
    fr.SameOrbitEnabled = False
    fr.EqualDiameterEnabled = False
    fr.TangentEnabled = False
    fr.SymmetricEnabled = False
    fr.OffsetEnabled = False
    fr.RigidBodyFaceEnabled = False
    deleteFaceBuilder.Type = NXOpen.Features.DeleteFaceBuilder.SelectTypes.Hole
    deleteFaceBuilder.FaceEdgeBlendPreference = NXOpen.Features.DeleteFaceBuilder.FaceEdgeBlendPreferenceOptions.Cliff

    sc_rule_options = workPart.ScRuleFactory.CreateRuleOptions()
    sc_rule_options.SetSelectedFromInactive(False)
    face_dumb_rule = workPart.ScRuleFactory.CreateRuleFaceDumb(faces_to_delete, sc_rule_options)
    deleteFaceBuilder.FaceCollector.ReplaceRules([face_dumb_rule], False)
    sc_rule_options.Dispose()

    try:
        feature = deleteFaceBuilder.Commit()
    except NXOpen.NXException as ex:
        raise
    finally:
        deleteFaceBuilder.Destroy()
    return feature


def remove_parameters(theSession, workPart, body_list):
    if not body_list: return True
    valid_bodies = []
    for b in body_list:
        try:
            _ = b.Tag
            valid_bodies.append(b)
        except:
            pass
    if not valid_bodies: return True

    mark_id = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Remove Parameters")
    builder = None
    try:
        builder = workPart.Features.CreateRemoveParametersBuilder()
        builder.Objects.Add(valid_bodies)
        builder.Commit()
        return True
    except:
        theSession.UndoToMark(mark_id, "Remove Parameters")
        return False
    finally:
        if builder:
            try:
                builder.Destroy()
            except:
                pass


# ==============================================================================
# 删孔逻辑
# ==============================================================================

def delete_conical_hole_main(theSession, workPart):
    delete_hole_list = set()
    # 使用全局 TARGET_LAYER
    bodies = get_bodies_on_target_layer(workPart, TARGET_LAYER)
    print_to_info_window(f"扫描锥面孔... (实体数: {len(bodies)})")

    for body in bodies:
        conical_face_list_for_body = []
        try:
            faces = body.GetFaces()
            for face in faces:
                if face in delete_hole_list: continue
                try:
                    if face.Color == 6 or face.Color == 181: continue
                    if int(str(face.SolidFaceType)) != 2: continue
                except:
                    continue

                edges = face.GetEdges()
                if len(edges) == 2:
                    pair_faces = []
                    for edge in edges:
                        try:
                            connected_faces = edge.GetFaces()
                            if len(connected_faces) != 2: continue
                            has_cone = False
                            for cf in connected_faces:
                                if str(cf.SolidFaceType) == "3": has_cone = True; break
                            if has_cone:
                                for cf in connected_faces: pair_faces.append(cf)
                        except:
                            pass
                    if pair_faces:
                        for pf in pair_faces:
                            if pf not in delete_hole_list:
                                delete_hole_list.add(pf)
                                conical_face_list_for_body.append(pf)
            if len(conical_face_list_for_body) > 0:
                unique_faces = list(set(conical_face_list_for_body))
                try:
                    delete_faces(unique_faces)
                except Exception as e:
                    for f in unique_faces:
                        try:
                            f.Color = 181; f.RedisplayObject()
                        except:
                            pass
        except Exception as e:
            pass


def delete_hole_main(theSession, workPart):
    hole_list = []
    delete_hole_list = set()
    bodies = get_bodies_on_target_layer(workPart, TARGET_LAYER)

    for body in bodies:
        faces = body.GetFaces()
        for face in faces:
            if face in delete_hole_list: continue
            if face.Color == 6: continue
            try:
                if int(str(face.SolidFaceType)) != 2: continue
            except:
                continue
            edges = face.GetEdges()
            if len(edges) == 2:
                for edge in edges:
                    fs = edge.GetFaces()
                    if len(fs) != 2: continue
                    delete_hole_list.add(face)
                    hole_list.append(face)
    if len(hole_list) > 0:
        for hole in hole_list:
            try:
                delete_faces([hole])
            except:
                hole.Color = 181
                hole.RedisplayObject()


def delete_holes_main(theSession, workPart, bodies):
    """ 执行逻辑封装 """
    for body in bodies:
        delete_face_list = set()
        faces = body.GetFaces()
        for face in faces:
            if face in delete_face_list: continue
            try:
                dtype = str(face.SolidFaceType)
            except:
                continue
            if int(dtype) == 2:
                edges = face.GetEdges()
                if len(edges) > 2:
                    f_edge_obj = None
                    ref_face = None
                    arr = []
                    for edge in edges:
                        sub_faces = edge.GetFaces()
                        for face2 in sub_faces:
                            if face != face2: ref_face = face2
                        vertices = edge.GetVertices()
                        if not vertices: continue
                        st = (vertices[0].X, vertices[0].Y, vertices[0].Z)
                        et = (vertices[1].X, vertices[1].Y, vertices[1].Z)
                        arr.append((ref_face, st[2]))
                        if to_tuple(st) == to_tuple(et):
                            f_edge_obj = edge

                    sorted_arr = merge_sort_fast(arr)
                    if f_edge_obj is not None and sorted_arr and len(sorted_arr) >= 2:
                        try:
                            target_ref = sorted_arr[-2][0]
                            create_extrude_from_edge(workPart, f_edge_obj, target_ref, body)
                            print_to_info_window(f"  成功填堵(>2边) Face: {face.Tag}")
                        except Exception as e:
                            try:
                                face.Color = 186; face.RedisplayObject()
                            except:
                                pass

    # Step 2
    print_to_info_window("Step 2: 删除锥面孔...")
    delete_conical_hole_main(theSession, workPart)

    # Step 3
    print_to_info_window("Step 3: 处理 = 2 边孔...")
    bodies = get_bodies_on_target_layer(workPart, TARGET_LAYER)
    for body in bodies:
        delete_face_list = set()
        faces = body.GetFaces()
        for face in faces:
            if face in delete_face_list: continue
            try:
                dtype = str(face.SolidFaceType)
            except:
                continue
            if int(dtype) == 2:
                edges = face.GetEdges()
                if len(edges) == 2:
                    f_edge = None
                    min_num = float("inf")
                    try:
                        v0 = edges[0].GetVertices()
                        v1 = edges[1].GetVertices()
                        if not v0 or not v1: continue
                        if round(v0[0].Z, 1) == round(v1[0].Z, 1): continue
                    except:
                        continue

                    for edge_idx in range(len(edges)):
                        vertices = edges[edge_idx].GetVertices()
                        if not vertices: continue
                        st = (vertices[0].X, vertices[0].Y, vertices[0].Z)
                        et = (vertices[1].X, vertices[1].Y, vertices[1].Z)
                        if to_tuple(st) == to_tuple(et):
                            if st[2] < min_num:
                                min_num = st[2]
                                f_edge = edges[edge_idx]
                    if f_edge is not None:
                        ref_face_edge = edges[0] if f_edge != edges[0] else edges[1]
                        sub_faces = ref_face_edge.GetFaces()
                        for face2 in sub_faces:
                            if face != face2:
                                try:
                                    create_extrude_from_edge(workPart, f_edge, face2, body)
                                    print_to_info_window(f"  成功填堵(=2边) Face: {face.Tag}")
                                except Exception as e:
                                    pass

    # Step 4
    print_to_info_window("Step 4: 删除剩余简单孔...")
    delete_hole_main(theSession, workPart)

    # Step 5
    print_to_info_window("Step 5: 移除参数...")
    final_bodies = get_bodies_on_target_layer(workPart, TARGET_LAYER)
    remove_parameters(theSession, workPart, final_bodies)

    print_to_info_window("方案二步骤结束。")


def delete_faces(faces_to_delete):
    """
    严格按录制宏逻辑执行删除面（同步建模 - 删除面）
    :param faces_to_delete: list of NXOpen.Face 需要删除的面列表
    :return: NXOpen.Features.DeleteFace 返回的特征对象
    """

    theSession = NXOpen.Session.GetSession()
    workPart = theSession.Parts.Work

    # -------------------------------------------------------------------------
    # 录制宏：创建 DeleteFaceBuilder
    # -------------------------------------------------------------------------
    markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "删除面-开始")

    deleteFaceBuilder = workPart.Features.CreateDeleteFaceBuilder(NXOpen.Features.Feature.Null)

    deleteFaceBuilder.FaceRecognized.RelationScope = 1023

    # 创建平面（录制代码需要）
    origin = NXOpen.Point3d(0.0, 0.0, 0.0)
    normal = NXOpen.Vector3d(0.0, 0.0, 1.0)
    plane = workPart.Planes.CreatePlane(origin, normal, NXOpen.SmartObject.UpdateOption.WithinModeling)

    deleteFaceBuilder.CapPlane = plane

    # 创建表达式（录制代码需要）
    unit = deleteFaceBuilder.MaxHoleDiameter.Units
    expr1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit)
    expr2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit)

    # -------------------------------------------------------------------------
    # 录制宏：识别选项全部关闭
    # -------------------------------------------------------------------------
    fr = deleteFaceBuilder.FaceRecognized
    fr.CoplanarEnabled = False
    fr.CoplanarAxesEnabled = False
    fr.CoaxialEnabled = False
    fr.SameOrbitEnabled = False
    fr.EqualDiameterEnabled = False
    fr.TangentEnabled = False
    fr.SymmetricEnabled = False
    fr.OffsetEnabled = False
    fr.RigidBodyFaceEnabled = False

    # 设置删除类型
    deleteFaceBuilder.Type = NXOpen.Features.DeleteFaceBuilder.SelectTypes.Hole
    deleteFaceBuilder.UseHoleDiameter = False

    deleteFaceBuilder.MaxHoleDiameter.SetFormula("5")
    deleteFaceBuilder.MaxBlendRadius.SetFormula("5")

    deleteFaceBuilder.FaceEdgeBlendPreference = NXOpen.Features.DeleteFaceBuilder.FaceEdgeBlendPreferenceOptions.Cliff

    deleteFaceBuilder.Type = NXOpen.Features.DeleteFaceBuilder.SelectTypes.Blend

    deleteFaceBuilder.CapPlane = NXOpen.Plane.Null

    fr.CloneScope = 511
    fr.UseFindClone = True
    fr.UseFindRelated = False
    fr.UseFaceBrowse = True
    fr.RelationScope = 0

    # -------------------------------------------------------------------------
    # 替换录制中的选择面逻辑 → 使用 faces_to_delete
    # -------------------------------------------------------------------------
    sc_rule_options = workPart.ScRuleFactory.CreateRuleOptions()
    sc_rule_options.SetSelectedFromInactive(False)

    # face dumb rule
    face_dumb_rule = workPart.ScRuleFactory.CreateRuleFaceDumb(faces_to_delete, sc_rule_options)

    deleteFaceBuilder.FaceCollector.ReplaceRules([face_dumb_rule], False)

    sc_rule_options.Dispose()

    # -------------------------------------------------------------------------
    # 执行删除操作
    # -------------------------------------------------------------------------
    markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "删除面执行")

    deleteFaceBuilder.Type = NXOpen.Features.DeleteFaceBuilder.SelectTypes.Hole
    feature = deleteFaceBuilder.Commit()

    theSession.DeleteUndoMark(markId2, None)
    theSession.SetUndoMarkName(markId1, "删除面完成")


    deleteFaceBuilder.Destroy()

    for expr in [expr1, expr2]:
        try:
            workPart.Expressions.Delete(expr)
        except NXOpen.NXException:
            pass

    plane.DestroyPlane()

    return feature

def delete_hole_v1_main(theSession, workPart, need_copy_layer=True):
    """ 执行逻辑封装 """
    if need_copy_layer:
        print_to_info_window("=== 初始化图层操作 ===")
        copy_count = copy_bodies_from_source_to_target_layer(SOURCE_LAYER, TARGET_LAYER)
        if copy_count == 0:
            print_to_info_window("没有实体可处理，程序停止。")
            return None

        set_work_layer(TARGET_LAYER)
        hide_all_layers_except_work_layer(TARGET_LAYER)
    else:
        # 第二轮会进入这里，跳过复制，直接处理现有实体
        print_to_info_window("=== 跳过图层复制（继续处理当前实体） ===")

    bodies = get_bodies_on_target_layer(workPart, TARGET_LAYER)
    hole_list = []
    for body in bodies:
        faces = body.GetFaces()
        for face in faces:
            if face.Color == 6:
                continue

            edges = face.GetEdges()
            if len(edges) == 1:
                hole_list.append(face)

            if len(edges) == 2:
                hole_list.append(face)
    try:
        if hole_list:
            delete_faces(hole_list)
    except :
        return bodies
    # 移除参数
    remove_parameters(theSession,workPart,list(workPart.Bodies))
    # Step 5
    print_to_info_window("Step 5: 移除参数...")
    final_bodies = get_bodies_on_target_layer(workPart, TARGET_LAYER)
    remove_parameters(theSession, workPart, final_bodies)
    print_to_info_window("方案一流程部分完成！")
    return final_bodies


# ==============================================================================
# 封装的统一入口函数 (供 feature_cleaner.py 调用)
# ==============================================================================

def run_delete_logic(theSession, workPart):
    """
    执行完整的删孔流程 (两轮循环: 简单孔 -> 复制 -> 复杂孔 -> 简单孔 -> 复杂孔)
    """
    try:
        for i in range(2):
            print_to_info_window(f"\n======== [删除孔] 开始第 {i + 1} 轮循环 ========")

            # 方案一：
            # 第一轮 (i==0) 传 True (复制图层)，第二轮 (i==1) 传 False (跳过复制)
            print_to_info_window(f"方案一开始执行 (轮次: {i + 1})")
            bodies = delete_hole_v1_main(theSession, workPart, need_copy_layer=(i == 0))

            # 如果方案一没有返回有效实体则重新获取一下
            if bodies is None:
                bodies = get_bodies_on_target_layer(workPart, TARGET_LAYER)

            # 方案二：继续对同一组实体进行处理
            print_to_info_window(f"方案二开始执行 (轮次: {i + 1})")
            delete_holes_main(theSession, workPart, bodies)

            print_to_info_window(f"======== [删除孔] 结束第 {i + 1} 轮循环 ========\n")
        
        # 恢复显示
        restore_all_layers()
        print_to_info_window("✅ 外部孔清理流程执行完毕")
        return True
    
    except Exception as ex:
        print_to_info_window(f"❌ 删孔流程错误: {ex}")
        print_to_info_window(traceback.format_exc())
        return False


# ==============================================================================
# 主函数 (入口)
# ==============================================================================

def main():
    file_path = r"C:\Users\Admin\Desktop\ppp\13.prt"

    theSession = NXOpen.Session.GetSession()

    if not os.path.exists(file_path):
        print_to_info_window(f"找不到文件: {file_path}")
        return

    print_to_info_window(f"正在打开: {file_path}")

    try:
        # 打开部件
        base_part, load_status = theSession.Parts.OpenBaseDisplay(file_path)
        workPart = theSession.Parts.Work

        if workPart is None:
            print_to_info_window("打开部件失败")
            return

        # 调用封装的逻辑
        run_delete_logic(theSession, workPart)
        
        # 保存
        save_part(workPart)

    except Exception as ex:
        print_to_info_window(f" 错误: {ex}")
        print_to_info_window(traceback.format_exc())


if __name__ == '__main__':
    main()