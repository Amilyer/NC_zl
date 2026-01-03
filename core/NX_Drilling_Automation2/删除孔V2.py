import NXOpen
import NXOpen.Features
import NXOpen.GeometricUtilities

def print_to_info_window(message):
    """输出信息到 NX 信息窗口"""
    session = NXOpen.Session.GetSession()
    session.ListingWindow.Open()
    session.ListingWindow.WriteLine(str(message))
    try:
        session.LogFile.WriteLine(str(message))
    except Exception:
        pass

def merge_sort_fast(arr):
    # 初始化临时数组，长度与原数组一致，用None填充（元组元素更语义化）
    temp = [None] * len(arr)
    _merge_sort(arr, temp, 0, len(arr) - 1)
    return arr

def _merge_sort(arr, temp, left, right):
    if left >= right:
        return
    mid = (left + right) // 2
    _merge_sort(arr, temp, left, mid)
    _merge_sort(arr, temp, mid + 1, right)
    _merge(arr, temp, left, mid, right)

def _merge(arr, temp, left, mid, right):
    i, j = left, mid + 1
    k = left

    # 核心修改：比较元组的第二个元素 arr[i][1] 和 arr[j][1]
    while i <= mid and j <= right:
        if arr[i][1] <= arr[j][1]:
            temp[k] = arr[i]
            i += 1
        else:
            temp[k] = arr[j]
            j += 1
        k += 1

    # 处理剩余元素（直接复制元组，无需修改）
    while i <= mid:
        temp[k] = arr[i]
        i += 1
        k += 1

    while j <= right:
        temp[k] = arr[j]
        j += 1
        k += 1

    # 将 temp 中的元组复制回 arr
    for x in range(left, right + 1):
        arr[x] = temp[x]


def _get_tagged_object(tag_or_obj):
    """如果传入是 int/long tag，返回对应对象；否则返回原对象。"""
    try:
        # 如果是数字 tag，尝试获取 TaggedObject
        if isinstance(tag_or_obj, (int,)):
            obj = NXOpen.TaggedObjectManager.GetTaggedObject(tag_or_obj)
            return obj
    except Exception:
        pass
    return tag_or_obj

def _ensure_edge(obj, workPart):
    """
    确保返回 NXOpen.Edge。
    如果 obj 是 Edge 直接返回；
    如果 obj 是 Body/Feature 尝试从其 Faces -> Edges 里取第一个 Edge（自动尝试，谨慎）。
    否则返回 None。
    """
    # 1) 如果已经是 Edge
    if isinstance(obj, NXOpen.Edge):
        return obj

    # 2) 如果是 tag，先获取真实对象
    obj = _get_tagged_object(obj)

    # 3) 如果是 Face，尝试取一个边
    if isinstance(obj, NXOpen.Face):
        try:
            edges = obj.GetEdges()
            if edges and len(edges) > 0:
                return edges[0]
        except Exception:
            pass

    # 4) 如果是 Body，尝试通过面取边（自动尝试第一个 face 的第一个 edge）
    if isinstance(obj, NXOpen.Body):
        try:
            faces = obj.GetFaces()
            if faces and len(faces) > 0:
                for f in faces:
                    try:
                        es = f.GetEdges()
                        if es and len(es) > 0:
                            return es[0]
                    except Exception:
                        continue
        except Exception:
            pass

    # 5) 如果是 Feature（如 BRep feature），尝试 FindObject 查找 EDGE（不保证总成功）
    # 这里不暴力尝试更多 API 以免出错
    return None

def _ensure_face(obj):
    """确保返回 NXOpen.Face 或 None。obj 可以是 face 对象或 tag。"""
    o = _get_tagged_object(obj)
    if isinstance(o, NXOpen.Face):
        return o
    return None

def create_extrude_from_edge(workPart, edge, start_target_face, target_body=None):
    theSession = NXOpen.Session.GetSession()

    # 允许用户传入 tag（int）或对象，我们先尝试把它们解析回来
    edge_candidate = _get_tagged_object(edge)
    face_candidate = _get_tagged_object(start_target_face)
    body_candidate = _get_tagged_object(target_body) if target_body is not None else None

    # 尝试确保 edge 是 NXOpen.Edge
    edge_obj = _ensure_edge(edge_candidate, workPart)
    if edge_obj is None:
        raise TypeError("参数 'edge' 必须是 NXOpen.Edge（或可解析到 Edge 的 tag/对象）。"
                        "\n你传入的对象类型是: {} "
                        "\n建议：在 Journal 中用如下方式获取 edge："
                        "\n  brep = workPart.Features.FindObject(\"UNPARAMETERIZED_FEATURE(1)\")"
                        "\n  edge = brep.FindObject(\"EDGE * 1 * 11 {...}\")"
                        "\n然后把该 edge 作为参数传入。".format(type(edge_candidate)))

    # 确保 face 是 NXOpen.Face
    face_obj = _ensure_face(face_candidate)
    if face_obj is None:
        raise TypeError("参数 'start_target_face' 必须是 NXOpen.Face（或能解析到 Face 的 tag）。"
                        "\n你传入的对象类型是: {} "
                        "\n建议：用 brep.FindObject(...) 或从 body 的 Faces 中选择一个 Face，并把它传入。".format(type(face_candidate)))

    # 若 target_body 提供，确保是 Body（否则置为 Body.Null）
    if body_candidate is None:
        tb = [NXOpen.Body.Null]
    else:
        if not isinstance(body_candidate, NXOpen.Body):
            raise TypeError("参数 'target_body' 必须是 NXOpen.Body 或 None。")
        tb = [body_candidate]

    # ---------- 与之前相同的拉伸构建逻辑（只使用已校验的 edge_obj / face_obj / tb） ----------
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
        edges_arr = [edge_obj]   # 保证这里是 Edge
        edge_rule = workPart.ScRuleFactory.CreateRuleEdgeDumb(edges_arr, sel_opts)
        sel_opts.Dispose()

        help_point = NXOpen.Point3d(0.0, 0.0, 0.0)
        try:
            verts = edge_obj.GetVertices()
            if verts and len(verts) > 0:
                vp = verts[0].Point
                help_point = NXOpen.Point3d(vp.X, vp.Y, vp.Z)
        except Exception:
            pass

        rules = [edge_rule]
        section.AddToSection(rules, edge_obj, NXOpen.NXObject.Null, NXOpen.NXObject.Null,
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
        extrudeBuilder.Limits.StartExtend.Target = NXOpen.DisplayableObject.Null
        extrudeBuilder.Limits.StartExtend.Value.SetFormula("42")
        extrudeBuilder.Limits.EndExtend.Value.SetFormula("0")

        # set provided start face
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
        except Exception:
            pass

        return feature

    except Exception:
        if extrudeBuilder is not None:
            try:
                extrudeBuilder.Destroy()
            except Exception:
                pass
        raise


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
    # deleteFaceBuilder.Heal = True
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

    # -------------------------------------------------------------------------
    # 清理
    # -------------------------------------------------------------------------
    deleteFaceBuilder.Destroy()

    # 尝试删除表达式（若被依赖会捕获异常）
    for expr in [expr1, expr2]:
        try:
            workPart.Expressions.Delete(expr)
        except NXOpen.NXException:
            pass

    plane.DestroyPlane()

    return feature


def to_tuple(pt):
    """统一坐标格式"""
    return (round(pt[0], 1), round(pt[1], 1), round(pt[2], 1))

def remove_parameters(theSession, workPart, body_list):
    """移除参数"""
    if not body_list:
        return True
    valid_bodies = []
    for b in body_list:
        try:
            _ = b.Tag
            valid_bodies.append(b)
        except:
            pass
    if not valid_bodies:
        return True

    print_to_info_window(f"正在对 {len(valid_bodies)} 个实体执行去参...")
    mark_id = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Remove Parameters")
    builder = None
    try:
        builder = workPart.Features.CreateRemoveParametersBuilder()
        builder.Objects.Add(valid_bodies)
        builder.Commit()
        theSession.SetUndoMarkName(mark_id, "Remove Parameters Success")
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


def delete_conical_hole_main(theSession,workPart):
    # 记录所有被删除的面
    delete_hole_list = set()
    # 删除带有锥面的孔
    conical_face_list = []
    bodies = list(workPart.Bodies)
    for body in bodies:
        faces = body.GetFaces()
        for face in faces:
            if face in delete_hole_list:
                continue
            if face.Color == 6:
                continue
            dtype = str(face.SolidFaceType)
            # 获取圆柱面
            if int(dtype) != 2:
                continue
            edges = face.GetEdges()
            edge_num = len(edges)
            if edge_num == 2:
                for edge in edges:
                    try:
                        faces = edge.GetFaces()
                        if len(faces) != 2:
                            continue
                        if "3" in [str(item.SolidFaceType) for item in faces]:
                            for face2 in faces:
                                delete_hole_list.add(face2)
                                conical_face_list.append(face2)
                    except:
                        print_to_info_window(f"face:{face.Tag},删除失败,面标记为粉色")
                        face.color = 181
                        face.RedisplayObject()
    # 删除锥面孔
    if len(conical_face_list) > 0:
        delete_faces(conical_face_list)
    # 移除参数
    remove_parameters(theSession, workPart, list(workPart.Bodies))

def delete_hole_main(theSession, workPart):
    # 简单孔
    hole_list = []
    # 记录所有被删除的面
    delete_hole_list = set()
    bodies = list(workPart.Bodies)
    for body in bodies:
        faces = body.GetFaces()
        for face in faces:
            if face in delete_hole_list:
                continue
            if face.Color == 6:
                continue
            dtype = str(face.SolidFaceType)
            # 获取圆柱面
            if int(dtype) != 2:
                continue
            edges = face.GetEdges()
            edge_num = len(edges)
            if edge_num == 2:
                for edge in edges:
                    faces = edge.GetFaces()
                    if len(faces) != 2:
                        continue
                    if "3" not in [str(item.SolidFaceType) for item in faces]:
                        delete_hole_list.add(face)
                        hole_list.append(face)
    # 删除普通孔
    if len(hole_list) > 0:
        for hole in hole_list:
            try:
                delete_faces([hole])
                # 移除参数
                remove_parameters(theSession, workPart, list(workPart.Bodies))
            except:
                print_to_info_window(f"face:{hole.Tag},删除失败,面标记为粉色(注：如果有粉色面，说明是真的删除失败！！！！)")
                hole.Color = 181
                hole.RedisplayObject()


def final_delete_holes_main(theSession, workPart):
    """ 执行顺序：1.拉伸边数大于2的孔 2.删除带有锥面的孔 3.拉伸边数等于2的孔 4.删除剩余孔   严格按照顺序执行，否则会报错"""
    # 获取体
    body = list(workPart.Bodies)[0]
    # 存放已经被消除的面
    delete_face_list = set()
    # 获取所有面
    faces = body.GetFaces()
    # 先拉伸边数大于2的面
    for face in faces:
        if face in delete_face_list:
            continue
        dtype = str(face.SolidFaceType)
        # 获取圆柱面
        if int(dtype) == 2:
            # 获取圆柱面的边
            edges = face.GetEdges()
            if len(edges) < 2:
                continue
            elif len(edges) > 2:
                f_edge_obj = None
                ref_face = None
                arr = []
                min_z = float("inf")
                for edge in edges:
                    faces = edge.GetFaces()
                    for face2 in faces:
                        if face != face2:
                            ref_face = face2
                    vertices = edge.GetVertices()  # 返回 list[NXOpen.Point3d]
                    st = (vertices[0].X, vertices[0].Y, vertices[0].Z)
                    et = (vertices[1].X, vertices[1].Y, vertices[1].Z)
                    arr.append((ref_face,st[2]))
                    if to_tuple(st) == to_tuple(et):
                            f_edge_obj = edge
                sorted_arr = merge_sort_fast(arr)
                if f_edge_obj is not None:
                    create_extrude_from_edge(workPart, f_edge_obj, sorted_arr[-2][0], body)
    # 调用删除锥面孔函数
    delete_conical_hole_main(theSession, workPart)
    # 再执行边数等于2的面
    # 获取体
    body = list(workPart.Bodies)[0]
    # 存放已经被消除的面
    delete_face_list = set()
    # 获取所有面
    faces = body.GetFaces()
    for face in faces:
        if face in delete_face_list:
            continue
        dtype = str(face.SolidFaceType)
        # 获取圆柱面
        if int(dtype) == 2:
            # 获取圆柱面的边
            edges = face.GetEdges()
            if len(edges) == 2:
                f_edge = None
                min_num = float("inf")
                if round(edges[0].GetVertices()[0].Z,1) == round(edges[1].GetVertices()[0].Z,1):
                    continue
                for edge_idx in range(len(edges)):
                    vertices = edges[edge_idx].GetVertices()
                    st = (vertices[0].X, vertices[0].Y, vertices[0].Z)
                    et = (vertices[1].X, vertices[1].Y, vertices[1].Z)
                    if to_tuple(st) == to_tuple(et):
                        if st[2] < min_num:
                            min_num = st[2]
                            f_edge = edges[edge_idx]
                if f_edge is not None:
                    for delete_face in f_edge.GetFaces():
                        delete_face_list.add(delete_face)
                    ref_face_edge = edges[0] if f_edge != edges[0] else edges[1]
                    faces = ref_face_edge.GetFaces()
                    for face2 in faces:
                        delete_face_list.add(face2)
                        if face != face2:
                            create_extrude_from_edge(workPart, f_edge, face2, body)
    # 删除剩余孔
    delete_hole_main(theSession,workPart)
    # 整体去参
    remove_parameters(theSession, workPart, list(workPart.Bodies))

if __name__ == '__main__':
    theSession = NXOpen.Session.GetSession()
    workPart = theSession.Parts.Work
    final_delete_holes_main(theSession,workPart)
