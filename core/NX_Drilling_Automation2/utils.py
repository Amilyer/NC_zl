# -*- coding: utf-8 -*-
"""
工具函数模块
包含通用功能：日志输出、数学计算、坐标处理等
"""

import math
import traceback
import NXOpen
import sys
import NXOpen.Annotations

def print_to_info_window(message):
    """输出信息到 NX 信息窗口"""
    print(message)
    # session = NXOpen.Session.GetSession()
    # session.ListingWindow.Open()
    # session.ListingWindow.WriteLine(str(message))
    # try:
    #     session.LogFile.WriteLine(str(message))
    # except Exception:
    #     pass


def euclidean_distance(point1, point2):
    """计算两个点之间的欧氏距离"""
    return math.sqrt((point1[0] - point2[0]) ** 2 +
                     (point1[1] - point2[1]) ** 2 +
                     (point1[2] - point2[2]) ** 2)


def point_from_angle(point, angle_deg, distance):
    """从起点沿角度方向（相对X轴）移动距离，计算目标点坐标（二维）"""
    rad = math.radians(angle_deg)
    x2 = point[0] + distance * math.cos(rad)
    y2 = point[1] + distance * math.sin(rad)
    return (x2, y2, 0.0)


def safe_origin(note):
    """尝试多种方式读取坐标"""
    try:
        if hasattr(note, "AnnotationOrigin"):
            o = note.AnnotationOrigin
            return (o.X, o.Y, o.Z)
        elif hasattr(note, "Origin"):
            o = note.Origin
            return (o.X, o.Y, o.Z)
        elif hasattr(note, "GetOrigin"):
            o = note.GetOrigin()
            return (o.X, o.Y, o.Z)
    except Exception as e:
        return (None, None, None)
    return (None, None, None)


def analyze_arc(arc):
    """分析单个 Arc 对象，判断是否完整圆"""
    try:
        delta = abs(arc.EndAngle - arc.StartAngle)
        is_full_circle = math.isclose(delta % (2 * math.pi), 2 * math.pi, rel_tol=1e-6) or math.isclose(delta,
                                                                                                        2 * math.pi,
                                                                                                        rel_tol=1e-6)
        return arc if is_full_circle else None
    except Exception:
        return None


def handle_exception(error_msg, details=None):
    """统一异常处理"""
    full_msg = f"❌ {error_msg}"
    if details:
        full_msg += f"\n详情: {details}"
    print_to_info_window(full_msg)
    return None


def get_circle_params(inner_circle_params):
    """
    :param inner_circle_params: 红色实线圆与其内圆参数
    :return: 处理后的参数列表，整体出错返回空列表
    """
    # 初始化最终结果列表
    result_list = []
    red_center_list = []

    try:
        # 遍历每个字典处理
        for item in inner_circle_params:
            # 提取两个圆心坐标（保留原始浮点精度）
            red_center = item['red_circle_center']  # 元组：(x, y, z)
            inner_center = item['inner_circle_center']  # 元组：(x, y, z)
            inner_dia = float(item['inner_circle_diameter'])  # 转为浮点保证精度

            # 高精度计算三维欧几里得距离（z轴为0不影响，保留通用逻辑）
            dx = red_center[0] - inner_center[0]
            dy = red_center[1] - inner_center[1]
            dz = red_center[2] - inner_center[2]
            distance = math.sqrt(dx ** 2 + dy ** 2 + dz ** 2)  # 距离公式

            # 严格判断距离是否 < 3.000000mm（6位小数精度）
            if distance < 3.000000:
                # 条件1：距离<3mm → 取red_circle_center的0、1索引值 + inner_dia
                x = red_center[0]
                y = red_center[1]
            else:
                # 条件2：距离≥3mm → 取inner_circle_center的0、1索引值 + inner_dia
                x = inner_center[0]
                y = inner_center[1]

            # 按格式存入列表（子列表包裹，符合你示例的[[x,y,dia]]格式）
            result_list.append([[x, y, 0.0], inner_dia / 2.0])
            red_center_list.append(red_center)

    except Exception as e:
        # 捕获函数执行过程中的所有异常，打印统一提示
        print(f"处理圆参数时发生错误：{str(e)}")
        return [], []

    return result_list, red_center_list

def to_tuple(pt):
    """统一坐标格式"""
    return (round(pt[0], 1), round(pt[1], 1), round(pt[2], 1))

# 适当放宽递归深度，因为大图路径确实可能很长
sys.setrecursionlimit(1000)

def build_adjacency_map(lines):
    """
    核心优化：构建邻接表（空间索引）
    Key: 点坐标 (x, y, z)
    Value: 连接该点的线段列表 [(line_obj, 另一端的点坐标), ...]
    """
    adj_map = {}

    for geom in lines:
        try:
            p0, p1 = endpoints(geom)  # 假设 endpoints 返回的是 NXOpen.Point3d 或类似结构
            # 统一转为 tuple 方便做字典 key
            p0_t = to_tuple(p0)
            p1_t = to_tuple(p1)

            # 过滤掉零长度的线段（脏数据）
            if p0_t == p1_t:
                continue

            # 记录 p0 端点的连接
            if p0_t not in adj_map:
                adj_map[p0_t] = []
            adj_map[p0_t].append((geom, p1_t))

            # 记录 p1 端点的连接
            if p1_t not in adj_map:
                adj_map[p1_t] = []
            adj_map[p1_t].append((geom, p0_t))

        except Exception:
            continue

    return adj_map


def find_connected_path(start_line, start_point, lines):
    """
    优化后的连通图搜索：使用邻接表代替列表遍历
    """
    # 1. 预处理：构建图 (耗时极短，仅需遍历一次)
    adj_map = build_adjacency_map(lines)

    # 2. 准备起点和终点
    start_p0, start_p1 = endpoints(start_line)
    start_p0_t = to_tuple(start_p0)
    start_p1_t = to_tuple(start_p1)

    # 确定搜索的起始方向：
    # 如果传入的 start_point 是 p0，那我们要往 p1 走，反之亦然
    target_start_pt = to_tuple(start_point)

    # 我们的目标是最终回到这个 target_start_pt
    goal_pt = target_start_pt

    # 当前位置设为线的另一头
    current_pos = start_p1_t if target_start_pt == start_p0_t else start_p0_t

    # 3. 初始化 DFS 状态
    visited_lines = set()
    visited_lines.add(start_line)

    path = [start_line]

    # 统计数据（可选，用于调试）
    stats = {"steps": 0}

    def dfs_fast(curr_pt):
        stats["steps"] += 1

        # 熔断机制：虽然有了空间索引，但如果几万步还没闭合，说明图本身有问题
        if stats["steps"] > 1000:
            return False

        # ★ 闭合检查：如果当前点回到了目标点，成功！
        if curr_pt == goal_pt:
            return True

        # ★ 极速查找：直接从字典取相邻线，不再遍历整个列表
        # 如果当前点是死路（没线连着），直接返回 False
        if curr_pt not in adj_map:
            return False

        # 获取连接该点的所有候选线
        # candidates 是 [(line_obj, next_point_tuple), ...]
        candidates = adj_map[curr_pt]

        # 排序优化（贪心策略）：可选
        # 如果 candidates 很多，可以优先选择与上一条线夹角较小的线（平滑过渡）
        # 这里先保持简单 DFS

        for next_line, next_pt in candidates:
            # 如果这条线已经走过，跳过
            if next_line in visited_lines:
                continue

            # 记录状态
            visited_lines.add(next_line)
            path.append(next_line)

            # 递归下一步
            if dfs_fast(next_pt):
                return True

            # 回溯 (Backtrack)
            path.pop()
            visited_lines.remove(next_line)

        return False

    # 4. 执行搜索
    # print(f"DEBUG: 图构建完成，节点数: {len(adj_map)}。开始快速搜索...")

    if dfs_fast(current_pos):
        # print(f"DEBUG: ✅ 成功闭合！路径长度: {len(path)}, 计算步数: {stats['steps']}")
        return path
    else:
        # print(f"DEBUG: ❌ 未能闭合。")
        return None



# -------------------------- 全局配置（浮点精度阈值） --------------------------
EPS = 1e-6  # 适配输入数据的浮点精度（原始数据有1e-7级误差）
TARGET_DISTANCE = 7.0  # 目标距离
DISTANCE_TOLERANCE = 0.1  # 距离判断的容差（处理浮点误差）
GRID_STEP = 0.5  # 网格步长（越小越密，精度越高，效率越低）
REFINE_AREA = (490, 105, 500, 115)  # 细化区域：(xmin, ymin, xmax, ymax)
REFINE_STEP = 0.1  # 细化区域的网格步长
MANUAL_POINT = (496.50000010, 110.88087815)  # 手动指定的点（优先验证）

# -------------------------- 基础几何工具函数 --------------------------
def get_arc_point(curve):
    """
    获取圆弧的起点和终点
    """
    uf_session = NXOpen.UF.UFSession.GetUFSession()
    arc_data = uf_session.Curve.AskArcData(curve.Tag)

    # 获取 StartAngle 和 EndAngle
    start_angle = arc_data.StartAngle
    end_angle = arc_data.EndAngle

    # 起点
    sp = NXOpen.Point3d(
        arc_data.ArcCenter[0] + arc_data.Radius * math.cos(start_angle),
        arc_data.ArcCenter[1] + arc_data.Radius * math.sin(start_angle),
        arc_data.ArcCenter[2]
    )

    # 终点
    ep = NXOpen.Point3d(
        arc_data.ArcCenter[0] + arc_data.Radius * math.cos(end_angle),
        arc_data.ArcCenter[1] + arc_data.Radius * math.sin(end_angle),
        arc_data.ArcCenter[2]
    )

    return ((sp.X, sp.Y, sp.Z), (ep.X, ep.Y, ep.Z))


def endpoints(obj):
    """返回线段或圆弧的两个端点"""
    if isinstance(obj, NXOpen.Line):
        return (obj.StartPoint.X, obj.StartPoint.Y, obj.StartPoint.Z), (obj.EndPoint.X, obj.EndPoint.Y, obj.EndPoint.Z)
    elif isinstance(obj, NXOpen.Arc):
        points = get_arc_point(obj)
        return (points[0], points[1])

def create_point(x, y):
    """创建二维点（舍弃z轴）"""
    return (float(x), float(y))

def vec_sub(p1, p2):
    """向量减法：p1 - p2"""
    return (p1[0] - p2[0], p1[1] - p2[1])

def vec_add(p1, p2):
    """向量加法：p1 + p2"""
    return (p1[0] + p2[0], p1[1] + p2[1])

def vec_mul(p, k):
    """向量数乘：p * k"""
    return (p[0] * k, p[1] * k)

def vec_div(p, k):
    """向量数除：p / k（避免除零）"""
    return (p[0] / k, p[1] / k) if k > EPS else (0.0, 0.0)

def vec_dot(p1, p2):
    """向量点积"""
    return p1[0] * p2[0] + p1[1] * p2[1]

def vec_cross(p1, p2):
    """二维向量叉积（返回标量，p1×p2）"""
    return p1[0] * p2[1] - p1[1] * p2[0]

def vec_norm(p):
    """向量模长"""
    return math.hypot(p[0], p[1])

def vec_normalize(p):
    """向量单位化（处理零向量）"""
    n = vec_norm(p)
    return vec_div(p, n) if n > EPS else (0.0, 0.0)

def line_from_two_points(p1, p2):
    """由两点构造直线：ax + by + c = 0（归一化）"""
    a = p2[1] - p1[1]
    b = p1[0] - p2[0]
    c = p2[0] * p1[1] - p1[0] * p2[1]
    norm = math.hypot(a, b)
    if norm > EPS:
        a /= norm
        b /= norm
        c /= norm
    return (a, b, c)

def line_signed_distance(line, p):
    """点到直线的有符号距离"""
    a, b, c = line
    return a * p[0] + b * p[1] + c

def line_intersection(line1, line2):
    """求两条直线的交点（返回点或None）"""
    a1, b1, c1 = line1
    a2, b2, c2 = line2
    det = a1 * b2 - a2 * b1
    if abs(det) < EPS:
        return None
    x = (b1 * c2 - b2 * c1) / det
    y = (a2 * c1 - a1 * c2) / det
    return create_point(x, y)

def segment_contains_point(seg, p):
    """判断点p是否在线段seg上"""
    p1, p2 = seg
    if abs(line_signed_distance(line_from_two_points(p1, p2), p)) > EPS:
        return False
    min_x = min(p1[0], p2[0]) - EPS
    max_x = max(p1[0], p2[0]) + EPS
    min_y = min(p1[1], p2[1]) - EPS
    max_y = max(p1[1], p2[1]) + EPS
    return min_x <= p[0] <= max_x and min_y <= p[1] <= max_y

def segment_clip_by_half_plane(seg, half_plane_line):
    """用半平面裁剪线段（返回裁剪后的线段或None）"""
    p1, p2 = seg
    d1 = line_signed_distance(half_plane_line, p1)
    d2 = line_signed_distance(half_plane_line, p2)

    if d1 >= -EPS and d2 >= -EPS:
        return seg
    if d1 < -EPS and d2 < -EPS:
        return None
    seg_line = line_from_two_points(p1, p2)
    intersect_p = line_intersection(seg_line, half_plane_line)
    if intersect_p is None:
        return None
    return (p1, intersect_p) if d1 >= -EPS else (intersect_p, p2)

def point_in_polygon_robust(p, polygon_vertices):
    """鲁棒的射线法判断点是否在多边形内部（含边界）"""
    px, py = p
    n = len(polygon_vertices)
    inside = False
    j = n - 1

    for i in range(n):
        vi = polygon_vertices[i]
        vj = polygon_vertices[j]
        xi, yi = vi
        xj, yj = vj

        if segment_contains_point((vi, vj), p):
            return True
        if abs(yi - yj) < EPS:
            j = i
            continue
        if ((yi > py + EPS) != (yj > py + EPS)):
            x_intersect = (xj - xi) * (py - yi) / (yj - yi) + xi
            if px < x_intersect + EPS:
                inside = not inside
        j = i

    return inside

def point_to_segment_distance_robust(p, seg):
    """鲁棒计算点到线段的最短距离"""
    p1, p2 = seg
    if vec_norm(vec_sub(p1, p2)) < EPS:
        return vec_norm(vec_sub(p, p1))

    vec_seg = vec_sub(p2, p1)
    vec_p = vec_sub(p, p1)
    t = vec_dot(vec_p, vec_seg) / vec_dot(vec_seg, vec_seg)
    t = max(0.0, min(1.0, t))

    foot_x = p1[0] + t * vec_seg[0]
    foot_y = p1[1] + t * vec_seg[1]
    return vec_norm(vec_sub(p, (foot_x, foot_y)))

def get_polygon_clockwise(polygon_vertices):
    """判断多边形顶点的顺时针/逆时针方向"""
    n = len(polygon_vertices)
    area = 0.0
    for i in range(n):
        p1 = polygon_vertices[i]
        p2 = polygon_vertices[(i + 1) % n]
        area += (p2[0] - p1[0]) * (p2[1] + p1[1])
    return area > EPS

def get_inner_normal_robust(seg, polygon_vertices, clockwise):
    """鲁棒计算线段的内法向量（单位向量）"""
    p1, p2 = seg
    vec_seg = vec_sub(p2, p1)
    vec_seg_unit = vec_normalize(vec_seg)
    left_normal = (-vec_seg_unit[1], vec_seg_unit[0])
    right_normal = (vec_seg_unit[1], -vec_seg_unit[0])
    initial_normal = right_normal if clockwise else left_normal

    mid_p = ((p1[0]+p2[0])/2, (p1[1]+p2[1])/2)
    test_p = vec_add(mid_p, vec_mul(initial_normal, EPS))
    if point_in_polygon_robust(test_p, polygon_vertices):
        return initial_normal
    else:
        return (-initial_normal[0], -initial_normal[1])

def get_half_plane_for_edge_robust(seg, d_target, polygon_vertices, clockwise):
    """鲁棒生成边的“距离≥d_target”的半平面直线"""
    p1, p2 = seg
    normal = get_inner_normal_robust(seg, polygon_vertices, clockwise)
    offset_p1 = vec_add(p1, vec_mul(normal, d_target))
    offset_p2 = vec_add(p2, vec_mul(normal, d_target))
    half_plane_line = line_from_two_points(offset_p1, offset_p2)

    test_p = vec_add(p1, vec_mul(normal, d_target/2))
    if line_signed_distance(half_plane_line, test_p) < -EPS:
        a, b, c = half_plane_line
        half_plane_line = (-a, -b, -c)
    return half_plane_line

# -------------------------- 核心验证函数（仅返回第一个符合条件的点） --------------------------
def is_point_satisfy_condition(p, seg_list, d_target):
    """验证点是否满足距离条件（仅返回布尔值）"""
    has_target = False
    all_ge = True
    for idx, seg in enumerate(seg_list):
        dist = point_to_segment_distance_robust(p, seg)
        # 检查是否有边的距离≈d_target（容差内）且垂足在边上
        if abs(dist - d_target) <= DISTANCE_TOLERANCE:
            p1, p2 = seg
            vec_seg = vec_sub(p2, p1)
            vec_p = vec_sub(p, p1)
            t = vec_dot(vec_p, vec_seg) / vec_dot(vec_seg, vec_seg) if vec_dot(vec_seg, vec_seg) > EPS else 0.0
            if 0.0 - EPS <= t <= 1.0 + EPS:
                has_target = True
        # 检查其他边的距离≥d_target - 容差
        if dist < d_target - DISTANCE_TOLERANCE:
            all_ge = False
            break
    return has_target and all_ge

# -------------------------- 几何构造法（快速查找第一个点） --------------------------
def find_first_point_by_construction(closed_bound, d_target):
    """几何构造法：找到第一个符合条件的点并返回，否则返回None"""
    segments = []
    for item in closed_bound:
        st = endpoints(item)[0]
        et = endpoints(item)[1]
        segments.append(((st[0],st[1]), (et[0],et[1])))
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
    if len(polygon_vertices) < 3 or len(seg_list) < 3:
        return None

    clockwise = get_polygon_clockwise(polygon_vertices)

    # 遍历每条边，找到第一个符合条件的点
    for edge_idx in range(len(seg_list)):
        current_seg = seg_list[edge_idx]
        p1, p2 = current_seg
        normal = get_inner_normal_robust(current_seg, polygon_vertices, clockwise)
        if vec_norm(normal) < EPS:
            continue
        # 生成偏移线段
        offset_p1 = vec_add(p1, vec_mul(normal, d_target))
        offset_p2 = vec_add(p2, vec_mul(normal, d_target))
        offset_seg = (offset_p1, offset_p2)

        # 裁剪偏移线段
        clipped_seg = offset_seg
        for j in range(len(seg_list)):
            if j == edge_idx:
                continue
            other_seg = seg_list[j]
            half_plane_line = get_half_plane_for_edge_robust(other_seg, d_target, polygon_vertices, clockwise)
            clipped_seg = segment_clip_by_half_plane(clipped_seg, half_plane_line)
            if clipped_seg is None:
                break

        # 验证特征点，找到第一个符合条件的点
        if clipped_seg is not None:
            clip_p1, clip_p2 = clipped_seg
            # 按顺序验证：中点（优先）、端点1、端点2
            test_points = [
                ((clip_p1[0]+clip_p2[0])/2, (clip_p1[1]+clip_p2[1])/2),
                clip_p1,
                clip_p2,
            ]
            for test_p in test_points:
                if point_in_polygon_robust(test_p, polygon_vertices) and is_point_satisfy_condition(test_p, seg_list, d_target):
                    return (round(test_p[0], 6), round(test_p[1], 6))
    return None

# -------------------------- 网格遍历法（查找第一个点） --------------------------
def find_first_point_by_grid(closed_bound, d_target):
    """网格遍历法：找到第一个符合条件的点并返回，否则返回None"""
    segments = []
    for item in closed_bound:
        st = endpoints(item)[0]
        et = endpoints(item)[1]
        segments.append(((st[0],st[1]), (et[0],et[1])))
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
    if len(polygon_vertices) < 3 or len(seg_list) < 3:
        return None

    # 第二步：遍历细化区域（细步长，次优先）
    rxmin, rymin, rxmax, rymax = REFINE_AREA
    x = rxmin
    while x <= rxmax + EPS:
        y = rymin
        while y <= rymax + EPS:
            p = (x, y)
            if point_in_polygon_robust(p, polygon_vertices) and is_point_satisfy_condition(p, seg_list, d_target):
                return (round(p[0], 6), round(p[1], 6))
            y += REFINE_STEP
        x += REFINE_STEP

    # 第三步：遍历全局网格（粗步长，最后优先级）
    xs = [p[0] for p in polygon_vertices]
    ys = [p[1] for p in polygon_vertices]
    xmin, xmax = min(xs), max(xs)
    ymin, ymax = min(ys), max(ys)
    x = xmin
    while x <= xmax + EPS:
        y = ymin
        while y <= ymax + EPS:
            p = (x, y)
            if point_in_polygon_robust(p, polygon_vertices) and is_point_satisfy_condition(p, seg_list, d_target):
                return (round(p[0], 6), round(p[1], 6))
            y += GRID_STEP
        x += GRID_STEP

    return None

# -------------------------- 主函数：仅返回一个符合条件的点 --------------------------
def find_one_valid_point(segments, d_target=TARGET_DISTANCE):
    """
    仅返回一个符合条件的点，优先级：
    1. 几何构造法找到的第一个点
    3. 细化区域网格中的第一个点
    4. 全局网格中的第一个点
    5. None（未找到）
    """
    # 1. 几何构造法（最高优先级）
    result = find_first_point_by_construction(segments, d_target)
    if result is not None:
        return result
    # 2. 网格遍历法（次优先级）
    result = find_first_point_by_grid(segments, d_target)
    return result

def create_rotate_feature(
        session,
        workPart,
        object_to_rotate,
        axis_origin,
        axis_direction,
        angle,
):
    """
    创建指定对象的旋转特征

    :param object_to_rotate: 要旋转的对象 (NXOpen.Features.Feature 或 NXOpen.Body)
    :param axis_origin: 旋转轴原点 (NXOpen.Point3d)
    :param axis_direction: 旋转轴方向矢量 (NXOpen.Vector3d)
    :param angle: 旋转角度(度)
    :return: 旋转后的特征对象 (NXOpen.Features.Feature)，失败时返回None
    """

    # 创建移动对象构建器
    move_builder = workPart.BaseFeatures.CreateMoveObjectBuilder(NXOpen.Features.MoveObject.Null)
    try:
        # 确保 axis_origin 和 axis_direction 是正确的类型
        axis_origin = NXOpen.Point3d(axis_origin[0], axis_origin[1], axis_origin[2])
        axis_direction = NXOpen.Vector3d(axis_direction[0], axis_direction[1], axis_direction[2])

        # 配置旋转参数
        move_builder.TransformMotion.Option = NXOpen.GeometricUtilities.ModlMotion.Options.Angle
        move_builder.TransformMotion.DeltaEnum = NXOpen.GeometricUtilities.ModlMotion.Delta.ReferenceAcsWorkPart

        # 创建旋转轴
        direction = workPart.Directions.CreateDirection(
            axis_origin,
            axis_direction,
            NXOpen.SmartObject.UpdateOption.WithinModeling
        )
        axis = workPart.Axes.CreateAxis(
            NXOpen.Point.Null,
            direction,
            NXOpen.SmartObject.UpdateOption.WithinModeling
        )
        move_builder.TransformMotion.AngularAxis = axis

        # 设置旋转角度
        move_builder.TransformMotion.Angle.SetFormula(str(angle))

        # 旋转对象
        move_builder.MoveObjectResult = NXOpen.Features.MoveObjectBuilder.MoveObjectResultOptions.MoveOriginal

        # 确保正确添加对象
        if isinstance(object_to_rotate, NXOpen.Features.Feature):
            # # 如果是特征，尝试获取其包含的体
            bodies = [object_to_rotate]
            for body in bodies:
                move_builder.ObjectToMoveObject.Add(body)
        elif isinstance(object_to_rotate, NXOpen.Features.Brep):
            # 如果是单独的体对象
            move_builder.ObjectToMoveObject.Add(object_to_rotate)
        else:
            raise ValueError("传入的对象类型无效，必须是 Body 或 Feature 对象。")

        # 执行旋转操作
        mark_id = session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "旋转对象")
        result = move_builder.Commit()
        print_to_info_window("旋转成功")
    except Exception as ex:
        print_to_info_window(f"旋转失败: {str(ex)}")
        return None
    finally:
        move_builder.Destroy()

def move_layer(session, work_part, objs, layer_num):
    markId1 = session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Move Layer")
    work_part.Layers.MoveDisplayableObjects(layer_num, objs)

    # 给对象列表去参
    def remove_parameters(self, body_list):
        """对给定的体对象列表执行“移除参数”操作。"""
        if not body_list:
            return
        try:
            builder = self.work_part.Features.CreateRemoveParametersBuilder()
            # 在某些版本中，Add可能只接受单个对象，使用循环更安全
            for body in body_list:
                builder.Objects.Add(body)
            builder.Commit()
            builder.Destroy()
        except NXOpen.NXException as e:
            pass

# 给对象列表去参
def remove_parameters(work_part, body_list):
    """对给定的体对象列表执行“移除参数”操作。"""
    if not body_list:
        return
    try:
        builder = work_part.Features.CreateRemoveParametersBuilder()
        # 在某些版本中，Add可能只接受单个对象，使用循环更安全
        # for body in body_list:
        #     builder.Objects.Add(body)
        #     builder.Commit()
        #     builder.Destroy()
        builder.Objects.Add(body_list)
        builder.Commit()
        builder.Destroy()
    except NXOpen.NXException as e:
        pass

def rotate_body(work_part, body, axis_point, direction_vector, angle):
    """
    旋转指定的体对象。

    :param work_part: 当前的工作零件 (NXOpen.Part)
    :param body: 要旋转的体对象 (NXOpen.Body)
    :param axis_point: 旋转轴的原点 (NXOpen.Point3d)
    :param direction_vector: 旋转轴的方向矢量 (NXOpen.Vector3d)
    :param angle: 旋转角度 (度)
    :return: 旋转后的体对象或None（如果失败）
    """
    try:
        # 获取当前会话
        the_session = NXOpen.Session.GetSession()

        # 创建旋转特征构建器
        move_object_builder = work_part.BaseFeatures.CreateMoveObjectBuilder(NXOpen.Features.MoveObject.Null)

        # 配置旋转轴和方向
        origin = NXOpen.Point3d(axis_point[0], axis_point[1], axis_point[2])
        direction = NXOpen.Vector3d(direction_vector[0], direction_vector[1], direction_vector[2])

        # 创建旋转轴方向
        axis_direction = work_part.Directions.CreateDirection(origin, direction, NXOpen.SmartObject.UpdateOption.WithinModeling)
        axis = work_part.Axes.CreateAxis(NXOpen.Point.Null, axis_direction, NXOpen.SmartObject.UpdateOption.WithinModeling)

        # 设置旋转角度
        move_object_builder.TransformMotion.AngularAxis = axis
        move_object_builder.TransformMotion.Option = NXOpen.GeometricUtilities.ModlMotion.Options.Angle
        move_object_builder.TransformMotion.Angle.SetFormula(str(angle))

        # 添加要旋转的体
        move_object_builder.ObjectToMoveObject.Add(body)

        # 设置撤销标记
        mark_id = the_session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "旋转体")

        # 执行旋转操作
        move_object_builder.Commit()

        # 删除撤销标记
        the_session.DeleteUndoMark(mark_id, None)

        # 销毁构建器
        move_object_builder.Destroy()

        return body  # 返回旋转后的体对象

    except NXOpen.NXException as e:
        print(f"错误: {e}")
        return None


def switch_to_manufacturing(session, work_part):
    """检查并切换到NX加工（Manufacturing）环境。"""
    try:
        if session.ApplicationName == "UG_APP_MANUFACTURING":
            print_to_info_window(f"已切换到加工环境")
            return True

        session.ApplicationSwitchImmediate("UG_APP_MANUFACTURING")
        if work_part.CAMSetup is None:
            work_part.CAMSetup.New()
        return True
    except Exception as e:
        print_to_info_window(f"❌ 切换到加工环境失败: {e}")
        return False


def switch_to_modeling(session):
    """检查并切换到NX建模（Modeling）环境。"""
    try:
        if session.ApplicationName == "UG_APP_MODELING":
            print_to_info_window("已处于建模环境。")
            return True

        session.ApplicationSwitchImmediate("UG_APP_MODELING")
        print_to_info_window("✅ 成功切换到建模环境。")
        return True
    except Exception as e:
        print_to_info_window(f"❌ 切换到建模环境失败: {e}")
        return False

def delete_nx_objects(session, objects):
    """
    使用 NX UpdateManager 删除对象（曲线、体等）。

    Args:
        objects: 单个 NX 对象或对象列表
        step_name: 日志步骤名
    """

    try:

        # 如果是单个对象，转成列表
        if not isinstance(objects, (list, tuple)):
            objects = [objects]

        # 过滤空对象
        valid_objs = [obj for obj in objects if obj]
        if not valid_objs:
            return False

        # 创建撤销标记
        markId1 = session.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, step_name)

        # 清空可能存在的错误列表
        session.UpdateManager.ClearErrorList()

        # 添加对象到删除列表
        nErrs1 = session.UpdateManager.AddObjectsToDeleteList(valid_objs)

        # 执行删除
        undo_mark_id = session.NewestVisibleUndoMark
        nErrs2 = session.UpdateManager.DoUpdate(undo_mark_id)

        # 删除撤销标记
        session.DeleteUndoMark(markId1, None)

        return True

    except NXOpen.NXException as ex:
        return False

    except Exception as ex:
        return False


def delete_body(session, body_to_delete):
    """
    删除指定的体对象。

    :param session: 当前会话对象 (NXOpen.Session)
    :param body_to_delete: 要删除的体对象 (NXOpen.Body)
    :return: 返回布尔值，指示删除是否成功
    """
    try:
        workPart = session.Parts.Work  # 获取当前工作部件

        # 设置 Undo 标记
        markId1 = session.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Delete")

        # 清除更新管理器的错误列表
        session.UpdateManager.ClearErrorList()

        # 设置可见的 Undo 标记
        markId2 = session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Delete")

        # 将对象添加到删除列表
        objects_to_delete = [body_to_delete]  # 将体对象添加到删除列表
        nErrs1 = session.UpdateManager.AddObjectsToDeleteList(objects_to_delete)

        # 执行删除更新
        id1 = session.NewestVisibleUndoMark
        nErrs2 = session.UpdateManager.DoUpdate(id1)

        # 删除 Undo 标记
        session.DeleteUndoMark(markId1, None)

        # 打印成功信息
        print_to_info_window(f"成功删除体对象")

        return True

    except Exception as ex:
        # 捕获异常并打印错误信息
        print_to_info_window(f"删除体对象时发生错误: {str(ex)}")
        return False

def is_mcs_exists(work_part, mcs_name):
    """🔧 判断是否存在名为 `mcs_name` 的 MCS（CAM GEOMETRY 组）。

    返回: True 如果存在，否则 False。
    兼容性说明：该函数只检查 `work_part.CAMSetup.CAMGroupCollection` 下的 GEOMETRY/{mcs_name} 对象
    """
    try:
        if work_part is None:
            return False
        # 检查是否有 CAMSetup 并可访问 CAMGroupCollection
        camsetup = getattr(work_part, 'CAMSetup', None)
        if camsetup is None:
            return False
        try:
            geo_objects = camsetup.CAMGroupCollection.FindObject(f"GEOMETRY")
            for obj in geo_objects.GetMembers():
                if obj.Name == mcs_name:
                    return True
        except Exception:
            return False
    except Exception:
        return False