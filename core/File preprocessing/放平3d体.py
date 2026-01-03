# -*- coding: utf-8 -*-
import NXOpen
import NXOpen.UF as UF
import NXOpen.Features
import NXOpen.GeometricUtilities
import math
import re
import traceback
import os


# ==========================================
#  日志与工具
# ==========================================
def print_log(message):
    try:
        s = NXOpen.Session.GetSession()
        if not s.ListingWindow.IsOpen: s.ListingWindow.Open()
        s.ListingWindow.WriteLine(str(message))
    except:
        print(str(message))


# ==========================================
#  数学工具 (向量/几何)
# ==========================================
class MathUtils:
    @staticmethod
    def to_point3d(obj):
        """安全获取坐标点"""
        if hasattr(obj, "Coordinates"):
            return obj.Coordinates
        return obj

    @staticmethod
    def get_len(p1, p2):
        pp1 = MathUtils.to_point3d(p1)
        pp2 = MathUtils.to_point3d(p2)
        return math.sqrt((pp1.X - pp2.X) ** 2 + (pp1.Y - pp2.Y) ** 2 + (pp1.Z - pp2.Z) ** 2)

    @staticmethod
    def is_point_same(p1, p2, tol=0.01):
        pp1 = MathUtils.to_point3d(p1)
        pp2 = MathUtils.to_point3d(p2)
        return ((pp1.X - pp2.X) ** 2 + (pp1.Y - pp2.Y) ** 2 + (pp1.Z - pp2.Z) ** 2) < tol ** 2

    @staticmethod
    def normalize(v):
        l = math.sqrt(v.X ** 2 + v.Y ** 2 + v.Z ** 2)
        if l < 1e-9: return NXOpen.Vector3d(0.0, 0.0, 1.0)
        return NXOpen.Vector3d(v.X / l, v.Y / l, v.Z / l)

    @staticmethod
    def vector_add(p, v, scale=1.0):
        pp = MathUtils.to_point3d(p)
        return NXOpen.Point3d(pp.X + v.X * scale, pp.Y + v.Y * scale, pp.Z + v.Z * scale)

    @staticmethod
    def dot_product(v1, v2):
        return v1.X * v2.X + v1.Y * v2.Y + v1.Z * v2.Z

    @staticmethod
    def point_dist_to_plane(pt, origin, normal):
        ppt = MathUtils.to_point3d(pt)
        porigin = MathUtils.to_point3d(origin)
        vec = NXOpen.Vector3d(ppt.X - porigin.X, ppt.Y - porigin.Y, ppt.Z - porigin.Z)
        return MathUtils.dot_product(vec, normal)

    @staticmethod
    def get_axis_enum(vec):
        n = MathUtils.normalize(vec)
        if abs(n.X) > 0.95: return 0  # X轴
        if abs(n.Y) > 0.95: return 1  # Y轴
        if abs(n.Z) > 0.95: return 2  # Z轴
        return -1


# ==========================================
#  测量
# ==========================================
class MeasurementHelper:
    @staticmethod
    def check_alignment_and_get_dims(bodies):
        """
        对比 AABB (Type 1) 和 BestFit (Type 0) 的边长。
        """
        session = NXOpen.Session.GetSession()
        pt_anchor = NXOpen.Point3d(0.0, 0.0, 0.0)
        TOLERANCE = 0.1  # 判定公差稍微放宽，避免浮点误差

        result = {
            'is_aligned': False,
            'dims_best': [0.0, 0.0, 0.0],
            'max_z': 0.0
        }

        if not bodies: return result

        try:
            # 1. Type 1: 轴对齐包容盒 (AABB)
            res_aabb = session.Measurement.GetBoundingBoxProperties(bodies, 1, pt_anchor, False)
            edges_aabb = res_aabb[2]
            dims_aabb_sorted = sorted([edges_aabb[0], edges_aabb[1], edges_aabb[2]])

            # 计算当前最高点 Z
            corners_aabb = res_aabb[0]
            max_z = -1e9
            for p in corners_aabb:
                if p.Z > max_z: max_z = p.Z
            result['max_z'] = max_z

            # 2. 获取 Type 0: 最适合包容盒 (OBB/BestFit)
            res_best = session.Measurement.GetBoundingBoxProperties(bodies, 0, pt_anchor, False)
            edges_best = res_best[2]
            dims_best_sorted = sorted([edges_best[0], edges_best[1], edges_best[2]])
            result['dims_best'] = dims_best_sorted

            # 3. 对比判断逻辑
            diff_0 = abs(dims_aabb_sorted[0] - dims_best_sorted[0])
            diff_1 = abs(dims_aabb_sorted[1] - dims_best_sorted[1])
            diff_2 = abs(dims_aabb_sorted[2] - dims_best_sorted[2])

            print_log(f"  > [尺寸检测] 固有(BestFit): {[round(x, 2) for x in dims_best_sorted]}")
            print_log(f"  > [尺寸检测] 投影(AABB)   : {[round(x, 2) for x in dims_aabb_sorted]}")

            if diff_0 < TOLERANCE and diff_1 < TOLERANCE and diff_2 < TOLERANCE:
                result['is_aligned'] = True
            else:
                result['is_aligned'] = False

            return result

        except Exception as e:
            print_log(f"  !! 测量失败: {e}")
            return result


# ==========================================
# Note 解析
# ==========================================
def get_req_dims_from_note(workPart):
    for n in workPart.Notes:
        try:
            txt = " ".join(n.GetText()).replace("×", "x").replace("*", "x")
            m = re.search(r"(\d+\.?\d*)\s*[lL].*?(\d+\.?\d*)\s*[wW].*?(\d+\.?\d*)\s*[tThH]", txt, re.I)
            if m:
                dims = [float(m.group(1)), float(m.group(2)), float(m.group(3))]
                print_log(f"  > [Note] 找到尺寸: L={dims[0]}, W={dims[1]}, T={dims[2]}")
                return dims
        except:
            continue
    return None


# ==========================================
#  矩形拓扑识别
# ==========================================
def find_closed_rectangle(workPart, L, W):
    """
    在 2D 线框中寻找长宽分别为 L 和 W 的闭合矩形。
    返回: (L边的方向向量, W边的方向向量)
    """
    # 1. 收集所有 Line 对象
    curves = [c for c in workPart.Curves if isinstance(c, NXOpen.Line)]
    line_data = []
    for c in curves:
        line_data.append((MathUtils.get_len(c.StartPoint, c.EndPoint), c.StartPoint, c.EndPoint))

    tol = 0.1  # 长度匹配公差

    # 2. 筛选候选线段
    l_cands = [x for x in line_data if abs(x[0] - L) <= tol]
    w_cands = [x for x in line_data if abs(x[0] - W) <= tol]

    if len(l_cands) < 2 or len(w_cands) < 2:
        print_log(" [拓扑] 候选线数量不足，无法构成矩形。")
        return None, None

    # 3. 匹配闭合环路 (L1-L2-W1-W2)
    for i in range(len(l_cands)):
        for j in range(i + 1, len(l_cands)):
            l1, l2 = l_cands[i], l_cands[j]
            for m in range(len(w_cands)):
                for n in range(m + 1, len(w_cands)):
                    w1, w2 = w_cands[m], w_cands[n]

                    # 收集所有端点
                    pts = [l1[1], l1[2], l2[1], l2[2], w1[1], w1[2], w2[1], w2[2]]

                    # 找出唯一不重复的点 (如果有闭合矩形，应该只有4个唯一角点)
                    uniq = []
                    for p in pts:
                        is_new = True
                        for u in uniq:
                            if MathUtils.is_point_same(p, u):
                                is_new = False
                                break
                        if is_new:
                            uniq.append(p)

                    if len(uniq) == 4:
                        # 验证：每个角点是否连接了一长一短，只要能围成4个点，且两长两短，基本就是目标矩形
                        print_log("[拓扑] 成功定位闭合矩形 2D 线框。")

                        # 计算 L 边向量 (取 l1)
                        vec_l = NXOpen.Vector3d(l1[2].X - l1[1].X, l1[2].Y - l1[1].Y, l1[2].Z - l1[1].Z)
                        # 计算 W 边向量 (取 w1)
                        vec_w = NXOpen.Vector3d(w1[2].X - w1[1].X, w1[2].Y - w1[1].Y, w1[2].Z - w1[1].Z)

                        return vec_l, vec_w

    print_log("[拓扑] 未找到闭合矩形。")
    return None, None


# ==========================================
#  孔数量
# ==========================================
def analyze_top_bottom_smart(workPart, bodies, root_point, z_vec, z_length):
    session = NXOpen.Session.GetSession()
    try:
        unit = workPart.UnitCollection.FindObject("MilliMeter")
    except:
        unit = workPart.PartUnits

    count_near, count_far = 0, 0
    norm_z = MathUtils.normalize(z_vec)
    plane_tolerance = 3.0

    print_log(f"[孔检测](仅检查平行于Z轴的平面)...")

    for body in bodies:
        for face in body.GetFaces():
            if face.SolidFaceType == NXOpen.Face.FaceType.Planar:
                for edge in face.GetEdges():
                    try:
                        props = session.Measurement.GetArcAndCurveProperties([edge], unit, True)
                        if props[7] and props[3] > 0:
                            if MathUtils.is_point_same(props[5], props[6], 0.001):
                                dist = MathUtils.point_dist_to_plane(props[7], root_point, norm_z)
                                if abs(dist) < plane_tolerance:
                                    count_near += 1
                                elif abs(dist - z_length) < plane_tolerance:
                                    count_far += 1
                    except:
                        continue

    print_log(f"Near端孔数: {count_near}, Far端孔数: {count_far}")

    if count_far > count_near:
        print_log("结果: Far端孔更多，需反转Z轴。")
        return True
    return False


# ==========================================
#  执行Z 轴平移
# ==========================================
def perform_safe_z_move(workPart, bodies, dist_z):
    delete_list = []
    try:
        vec_x = NXOpen.Vector3d(1.0, 0.0, 0.0)
        vec_y = NXOpen.Vector3d(0.0, 1.0, 0.0)
        vec_z = NXOpen.Vector3d(0.0, 0.0, 1.0)
        origin_start = NXOpen.Point3d(0.0, 0.0, 0.0)

        p1_s = workPart.Planes.CreatePlane(origin_start, vec_x, NXOpen.SmartObject.UpdateOption.WithinModeling)
        p2_s = workPart.Planes.CreatePlane(origin_start, vec_y, NXOpen.SmartObject.UpdateOption.WithinModeling)
        p3_s = workPart.Planes.CreatePlane(origin_start, vec_z, NXOpen.SmartObject.UpdateOption.WithinModeling)
        delete_list.extend([p1_s, p2_s, p3_s])

        xform_s = workPart.Xforms.CreateXform(p1_s, p2_s, p3_s, NXOpen.SmartObject.UpdateOption.WithinModeling, 1.0)
        csys_s = workPart.CoordinateSystems.CreateCoordinateSystem(xform_s,
                                                                   NXOpen.SmartObject.UpdateOption.WithinModeling)
        delete_list.extend([xform_s, csys_s])

        origin_end = NXOpen.Point3d(0.0, 0.0, float(dist_z))
        p1_e = workPart.Planes.CreatePlane(origin_end, vec_x, NXOpen.SmartObject.UpdateOption.WithinModeling)
        p2_e = workPart.Planes.CreatePlane(origin_end, vec_y, NXOpen.SmartObject.UpdateOption.WithinModeling)
        p3_e = workPart.Planes.CreatePlane(origin_end, vec_z, NXOpen.SmartObject.UpdateOption.WithinModeling)
        delete_list.extend([p1_e, p2_e, p3_e])

        xform_e = workPart.Xforms.CreateXform(p1_e, p2_e, p3_e, NXOpen.SmartObject.UpdateOption.WithinModeling, 1.0)
        csys_e = workPart.CoordinateSystems.CreateCoordinateSystem(xform_e,
                                                                   NXOpen.SmartObject.UpdateOption.WithinModeling)
        delete_list.extend([xform_e, csys_e])

        mb = workPart.BaseFeatures.CreateMoveObjectBuilder(NXOpen.Features.MoveObject.Null)
        mb.TransformMotion.Option = NXOpen.GeometricUtilities.ModlMotion.Options.CsysToCsys
        mb.TransformMotion.FromCsys = csys_s
        mb.TransformMotion.ToCsys = csys_e
        mb.ObjectToMoveObject.Add(bodies)
        mb.Commit()
        mb.Destroy()

    except Exception as e:
        print_log(f"  !! 安全平移失败: {e}")
    finally:
        try:
            session = NXOpen.Session.GetSession()
            um = session.UpdateManager
            delete_list.reverse()
            for obj in delete_list:
                if obj: um.AddToDeleteList(obj)
            clean_mark = session.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "CleanupZ")
            um.DoUpdate(clean_mark)
            session.DeleteUndoMark(clean_mark, None)
        except:
            pass


def perform_point_to_point_move(workPart, bodies, start_pt, end_pt):
    """
    执行点对点移动将其拉到绝对坐标原点。
    """
    s = NXOpen.Session.GetSession()
    delete_list = []

    try:
        # 1. 创建构建器
        mb = workPart.BaseFeatures.CreateMoveObjectBuilder(NXOpen.Features.MoveObject.Null)
        mb.TransformMotion.Option = NXOpen.GeometricUtilities.ModlMotion.Options.PointToPoint

        # 2. 创建临时 Smart Point 对象
        # 起点
        pt_from = workPart.Points.CreatePoint(start_pt)
        # 终点
        pt_to = workPart.Points.CreatePoint(end_pt)

        delete_list.extend([pt_from, pt_to])

        mb.TransformMotion.FromPoint = pt_from
        mb.TransformMotion.ToPoint = pt_to

        # 3. 添加要移动的对象
        mb.ObjectToMoveObject.Add(bodies)

        # 4. 提交
        mb.Commit()
        mb.Destroy()

    except Exception as e:
        print_log(f"  !! 点对点平移失败: {e}")

    finally:
        # 清理创建的临时点，避免污染模型
        try:
            um = s.UpdateManager
            for obj in delete_list:
                if obj: um.AddToDeleteList(obj)

            clean_mark = s.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "CleanupP2P")
            um.DoUpdate(clean_mark)
            s.DeleteUndoMark(clean_mark, None)
        except:
            pass

# ==========================================
#  主逻辑
# ==========================================
def execute_alignment(workPart):
    s = NXOpen.Session.GetSession()
    delete_list = []

    try:
        all_bodies = list(workPart.Bodies)
        if not all_bodies:
            print_log("!! 错误: 未找到任何 Body。")
            return False

        # 1. Note 解析
        req_dims = get_req_dims_from_note(workPart)

        print_log("[预检查] 验证当前姿态 (Type 1 vs Type 0)...")
        align_status = MeasurementHelper.check_alignment_and_get_dims(all_bodies)

        is_aligned = align_status['is_aligned']
        measured_dims = align_status['dims_best']  # [S, M, L]

        if not req_dims:
            # 如果没找到 Note，回退到使用测量尺寸，但此时无法进行2D图纸约束
            req_dims = [measured_dims[2], measured_dims[1], measured_dims[0]]
            print_log(f"[提示] 未找Note，使用测量尺寸: L={req_dims[0]:.2f}, W={req_dims[1]:.2f}")

        # -----------------------------------------------------------------
        # 分支 A: 已经摆正
        # -----------------------------------------------------------------
        # if is_aligned:
        #     print_log("  => 判定结果: [已摆正]")
        #     return True
        if is_aligned:
            print_log("  => 判定结果: [已摆正]")
            # 1. 重新获取 AABB 以确定最小角点
            pt_anchor = NXOpen.Point3d(0.0, 0.0, 0.0)
            res_aabb = s.Measurement.GetBoundingBoxProperties(all_bodies, 1, pt_anchor, False)
            corners = res_aabb[0]

            # 寻找最小坐标 (MinX, MinY, MinZ)
            min_x, min_y, min_z = 1e9, 1e9, 1e9
            for p in corners:
                if p.X < min_x: min_x = p.X
                if p.Y < min_y: min_y = p.Y
                if p.Z < min_z: min_z = p.Z

            current_min_pt = NXOpen.Point3d(min_x, min_y, min_z)
            target_origin = NXOpen.Point3d(0.0, 0.0, 0.0)

            # 2. 计算距离，判断是否已在原点
            dist = math.sqrt(min_x ** 2 + min_y ** 2 + min_z ** 2)

            if dist < 0.01:
                print_log("     (位置已在绝对原点，无操作)")
            else:
                print_log(f"     (位置偏离原点 {dist:.2f}，执行点到点归零...)")
                print_log(f"     [从: {min_x:.1f},{min_y:.1f},{min_z:.1f} -> 到: 0,0,0]")
                perform_point_to_point_move(workPart, all_bodies, current_min_pt, target_origin)
                print_log("     (归零完成)")

                # 移动后保存
                NXOpen.UF.UFSession.GetUFSession().Part.Save()

            return True

        # -----------------------------------------------------------------
        # 分支 B: 倾斜 -> 执行放平逻辑
        # -----------------------------------------------------------------
        print_log("判定结果: [倾斜] -> 开始执行归正。")

        # 2. 寻找参考 2D 矩形方向 (基于 Note 尺寸)
        # vec_l_ref 来自 2D 图纸，代表 L 边应该在的方向 (绝对坐标系)
        vec_l_ref, vec_w_ref = find_closed_rectangle(workPart, req_dims[0], req_dims[1])

        # 3. 创建 ToolingBox 分析 3D 实体
        print_log("ToolingBox 分析...")
        tc = workPart.Features.ToolingFeatureCollection
        bb = tc.CreateToolingBoxBuilder(NXOpen.Features.ToolingBox.Null)
        bb.Type = NXOpen.Features.ToolingBoxBuilder.Types.BoundedBlock
        bb.NonAlignedMinimumBox = True

        mat = NXOpen.Matrix3x3();
        mat.Xx = 1.0;
        mat.Yy = 1.0;
        mat.Zz = 1.0
        bb.SetBoxMatrixAndPosition(mat, NXOpen.Point3d(0.0, 0.0, 0.0))
        zero = "0"
        for o in [bb.OffsetPositiveX, bb.OffsetNegativeX, bb.OffsetPositiveY, bb.OffsetNegativeY, bb.OffsetPositiveZ,
                  bb.OffsetNegativeZ]:
            o.SetFormula(zero)

        r = workPart.ScRuleFactory.CreateRuleBodyDumb(all_bodies, True, workPart.ScRuleFactory.CreateRuleOptions())
        bb.BoundedObject.ReplaceRules([r], False)
        bb.CalculateBoxSize()
        bbox_feat = bb.Commit()
        bb.Destroy()
        delete_list.append(bbox_feat)

        # 4. 提取 ToolingBox 的方向向量
        bbox_body = bbox_feat.GetBodies()[0]
        edges = bbox_body.GetEdges()
        root_v_obj = edges[0].GetVertices()[0]
        root = MathUtils.to_point3d(root_v_obj)

        neighbors = []
        for e in edges:
            p1 = MathUtils.to_point3d(e.GetVertices()[0])
            p2 = MathUtils.to_point3d(e.GetVertices()[1])
            tgt = None
            if MathUtils.is_point_same(p1, root):
                tgt = p2
            elif MathUtils.is_point_same(p2, root):
                tgt = p1

            if tgt:
                v_raw = NXOpen.Vector3d(tgt.X - root.X, tgt.Y - root.Y, tgt.Z - root.Z)
                l = math.sqrt(v_raw.X ** 2 + v_raw.Y ** 2 + v_raw.Z ** 2)
                l_v = l if l > 1e-9 else 1.0
                neighbors.append({'len': l, 'vec': NXOpen.Vector3d(v_raw.X / l_v, v_raw.Y / l_v, v_raw.Z / l_v)})

        req_l, req_w = req_dims[0], req_dims[1]

        # -----------------------------------------------------------------
        # 5. 确定轴向映射规则 (根据板料线)
        # -----------------------------------------------------------------
        t_idx_l = 1
        t_idx_w = 0
        # 如果找到了 2D 图纸中的闭合矩形，按图纸方向指定
        if vec_l_ref:
            axis_enum = MathUtils.get_axis_enum(vec_l_ref)
            if axis_enum != -1:
                t_idx_l = axis_enum
                print_log(f"[规则] 2D图纸指示: L边 ({req_l}) 应对应轴: {t_idx_l} (0=X, 1=Y)")

                # 互斥设定 W 轴 (假设平面情况，非X即Y)
                if t_idx_l == 0:
                    t_idx_w = 1
                else:
                    t_idx_w = 0
            else:
                print_log("  > [警告] 2D图纸 L 边方向不正，使用默认逻辑。")
        else:
            print_log("  > [警告] 未找到匹配尺寸的 2D 闭合矩形，使用默认逻辑。")

        # 6. 填充向量
        vecs = [None, None, None]
        assigned_indices = set()
        z_len_val = 0.0
        # --- 子函数：在 neighbors 中找最接近 target_len 的边 ---
        def get_best_match_index(target_len, exclude_idxs):
            best_idx = -1
            min_diff = 1e9
            for i, n in enumerate(neighbors):
                if i in exclude_idxs: continue
                diff = abs(n['len'] - target_len)
                if diff < min_diff:
                    min_diff = diff
                    best_idx = i
            return best_idx, min_diff

        # 1. 优先匹配 L
        idx_l, diff_l = get_best_match_index(req_l, assigned_indices)
        if idx_l != -1:
            vecs[t_idx_l] = neighbors[idx_l]['vec']
            assigned_indices.add(idx_l)
            print_log(
                f"  > [匹配] L边 ({req_l}) 匹配到测量边: {neighbors[idx_l]['len']:.2f} (误差: {diff_l:.2f}) -> 轴 {t_idx_l}")
        else:
            print_log("  !! [错误] 无法找到 L 边的匹配项")

        # 2. 其次匹配 W (在剩下的里面找最接近 W 的)
        idx_w, diff_w = get_best_match_index(req_w, assigned_indices)
        if idx_w != -1:
            vecs[t_idx_w] = neighbors[idx_w]['vec']
            assigned_indices.add(idx_w)
            print_log(
                f"  > [匹配] W边 ({req_w}) 匹配到测量边: {neighbors[idx_w]['len']:.2f} (误差: {diff_w:.2f}) -> 轴 {t_idx_w}")
        else:
            print_log("  !! [错误] 无法找到 W 边的匹配项")

        # 3. 剩下的给 Z
        # 找到那个还没被分配的 neighbors index
        idx_z = -1
        for i in range(3):
            if i not in assigned_indices:
                idx_z = i
                break

        # 找到 vecs 里面那个还是 None 的槽位 (通常是 t_idx_l 和 t_idx_w 之外的那个)
        slot_z = -1
        for i in range(3):
            if vecs[i] is None:
                slot_z = i
                break

        if idx_z != -1 and slot_z != -1:
            vecs[slot_z] = neighbors[idx_z]['vec']
            z_len_val = neighbors[idx_z]['len']
            print_log(f"  > [匹配] 剩余边 (厚度) 匹配到测量边: {neighbors[idx_z]['len']:.2f} -> 轴 {slot_z}")

        # 最终检查
        if not (vecs[0] and vecs[1] and vecs[2]):
            print_log("!! 无法解析包围盒向量，跳过。")
            return False

        # 7. 孔方向检测 (确定Z轴正反)
        invert = analyze_top_bottom_smart(workPart, all_bodies, root, vecs[2], z_len_val)
        final_o = root if invert else MathUtils.vector_add(root, vecs[2], z_len_val)
        final_z = NXOpen.Vector3d(-vecs[2].X, -vecs[2].Y, -vecs[2].Z) if invert else vecs[2]

        # 8. 构建坐标系并移动
        planes = []
        for v in [vecs[0], vecs[1], final_z]:
            planes.append(workPart.Planes.CreatePlane(final_o, v, NXOpen.SmartObject.UpdateOption.WithinModeling))
        delete_list.extend(planes)

        fx = workPart.Xforms.CreateXform(planes[0], planes[1], planes[2],
                                         NXOpen.SmartObject.UpdateOption.WithinModeling, 1.0)
        fcs = workPart.CoordinateSystems.CreateCoordinateSystem(fx, NXOpen.SmartObject.UpdateOption.WithinModeling)
        delete_list.extend([fx, fcs])

        tx = workPart.Xforms.CreateXform(NXOpen.SmartObject.UpdateOption.WithinModeling, 1.0)
        tcs = workPart.CoordinateSystems.CreateCoordinateSystem(tx, NXOpen.SmartObject.UpdateOption.WithinModeling)
        delete_list.extend([tx, tcs])

        print_log("  > 执行移动 (CsysToCsys)...")
        mb = workPart.BaseFeatures.CreateMoveObjectBuilder(NXOpen.Features.MoveObject.Null)
        mb.TransformMotion.Option = NXOpen.GeometricUtilities.ModlMotion.Options.CsysToCsys
        mb.TransformMotion.FromCsys = fcs
        mb.TransformMotion.ToCsys = tcs
        mb.ObjectToMoveObject.Add(all_bodies)
        try:
            mb.Commit()
        except Exception as e:
            if "No motion" not in str(e): raise e
        mb.Destroy()

        # 9. 移动后 Z 轴高度归零复查
        print_log("  > Z轴高度复查...")
        mark = s.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Update")
        s.UpdateManager.DoUpdate(mark)
        s.DeleteUndoMark(mark, None)

        new_bodies = list(workPart.Bodies)
        res_check = MeasurementHelper.check_alignment_and_get_dims(new_bodies)
        check_z = res_check['max_z']

        if check_z > 0.1:
            print_log(f"  > 修正下移 {check_z:.2f}...")
            perform_safe_z_move(workPart, new_bodies, -check_z)

        # 10. 清理与保存
        print_log("  > 清理临时对象...")
        delete_list.reverse()
        um = s.UpdateManager
        for o in delete_list: um.AddToDeleteList(o)

        clean_mark = s.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Cleanup")
        um.DoUpdate(clean_mark)
        s.DeleteUndoMark(clean_mark, None)

        UF.UFSession.GetUFSession().Part.Save()
        print_log("✅ 完成。")
        return True

    except Exception as e:
        print_log(f"❌ 错误: {e}")
        print_log(traceback.format_exc())
        return False


def main():
    part_path = r"C:\Projects\NC\output\02_Process\3_Cleaned_PRT\DIE-03.prt"
    if os.path.exists(part_path):
        s = NXOpen.Session.GetSession()
        s.Parts.OpenBaseDisplay(part_path)
        execute_alignment(s.Parts.Work)



if __name__ == "__main__":
    main()