import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Polygon

# =============================
# 用户输入参数
# =============================
vertices = np.array([
    [11, 89],
    [22, 73],
    [25, 80],
    [11, 65]
])

r = 1           # 圆半径 (mm)
min_arc_gap = 2 # 要求：到最近边的弧距离 = 2mm，其余 > 2mm

# 采样设置
num_samples_per_side = 6    # 每条边上尝试生成多少个候选圆心
safe_margin_ratio = 0.1     # 靠近顶点部分裁剪掉的比例

# 图形设置
figsize = (10, 9)
save_fig = False
output_filename = "valid_circles_in_quadrilateral.png"
# =============================
# 工具函数
# =============================

def is_point_in_polygon(x, y, poly):
    """
    使用射线法判断点是否在多边形内部（包括边界）
    """
    n = len(poly)
    inside = False
    p1x, p1y = poly[0]
    for i in range(1, n + 1):
        p2x, p2y = poly[i % n]
        if min(p1y, p2y) < y <= max(p1y, p2y):
            if p1x + (y - p1y) * (p2x - p1x) / (p2y - p1y) > x:
                inside = not inside
        p1x, p1y = p2x, p2y
    return inside

def point_to_line_signed_distance(x0, y0, x1, y1, x2, y2):
    """
    计算点 (x0,y0) 到直线 (x1,y1)-(x2,y2) 的有符号距离
    正值表示在左侧 → 假设多边形是逆时针方向
    """
    dx = x2 - x1
    dy = y2 - y1
    norm = np.hypot(dx, dy)
    if norm == 0:
        return np.inf
    # 单位法向量指向左侧（左手法则）
    nx = -dy / norm
    ny = dx / norm
    # 向量从起点到目标点
    px = x0 - x1
    py = y0 - y1
    dist = nx * px + ny * py
    return dist

def ensure_counterclockwise(vertices):
    """
    确保顶点为逆时针顺序（便于统一内法向为左侧）
    """
    x = vertices[:, 0]
    y = vertices[:, 1]
    area = 0.5 * np.sum(x[:-1] * y[1:] - x[1:] * y[:-1])
    if area < 0:
        return vertices[::-1]  # 反转为逆时针
    return vertices.copy()

def get_offset_segment_points(x1, y1, x2, y2, offset, n_points=5, margin_ratio=0.1):
    """
    在距离边内侧 offset 处生成 n_points 个点（避开两端）
    """
    dx = x2 - x1
    dy = y2 - y1
    length = np.hypot(dx, dy)
    if length < 1e-6:
        return []

    # 单位法向量（向左即内侧）
    nx = -dy / length
    ny = dx / length

    # 偏移后线段的两个端点
    ox1 = x1 + nx * offset
    oy1 = y1 + ny * offset
    ox2 = x2 + nx * offset
    oy2 = y2 + ny * offset

    # 参数化取点，去掉靠近端点的部分
    start_t = margin_ratio
    end_t = 1 - margin_ratio
    points = []
    for i in range(n_points):
        t = start_t + (end_t - start_t) * i / (n_points - 1) if n_points > 1 else 0.5
        x = ox1 + t * (ox2 - ox1)
        y = oy1 + t * (oy2 - oy1)
        points.append((x, y))
    return points
# =============================
# 主逻辑：查找所有符合条件的圆心
# =============================

# 确保顶点为逆时针
verts_ccw = ensure_counterclockwise(np.vstack([vertices, vertices[0]]))
vertices = verts_ccw[:-1]

valid_centers = []  # 存储有效圆心
offset = r + min_arc_gap  # 圆心到最近边的距离应为 r + 2 = 5 mm

print("🔍 开始搜索符合条件的圆心...")
print(f"圆半径 r = {r} mm，要求最近边弧距离 = {min_arc_gap} mm，其余 > {min_arc_gap} mm\n")

for i in range(4):
    # 当前边
    p1 = vertices[i]
    p2 = vertices[(i+1) % 4]
    x1, y1 = p1
    x2, y2 = p2

    # 获取在该边内侧 offset 距离上的采样点
    candidates = get_offset_segment_points(
        x1, y1, x2, y2,
        offset=offset,
        n_points=num_samples_per_side,
        margin_ratio=safe_margin_ratio
    )

    for cx, cy in candidates:
        # 先快速检查是否在多边形内（含边界）
        if not is_point_in_polygon(cx, cy, vertices):
            continue

        # 计算圆心到四条边的有符号距离（正值表示在内侧）
        distances = []
        for j in range(4):
            a1 = vertices[j]
            a2 = vertices[(j+1) % 4]
            d = point_to_line_signed_distance(cx, cy, a1[0], a1[1], a2[0], a2[1])
            distances.append(d)

        # 所有距离必须 ≥ r，否则圆会越界
        if any(d < r - 1e-6 for d in distances):
            continue

        # 弧到各边的距离 = 圆心到边距离 - r
        edge_gaps = [d - r for d in distances]

        min_gap_val = min(edge_gaps)
        min_idx = np.argmin(edge_gaps)

        # 必须是当前这条边（i）为最小，且等于 min_arc_gap
        if abs(min_gap_val - min_arc_gap) > 0.1:  # 容差 0.1mm
            continue

        # 其他边的距离必须 > min_arc_gap
        other_gaps = [gap for idx, gap in enumerate(edge_gaps) if idx != i]
        if all(gap > min_arc_gap + 0.1 for gap in other_gaps):
            valid_centers.append((cx, cy))
            side_name = ['AB', 'BC', 'CD', 'DA'][i]
            print(f"✅ 找到有效圆心: ({cx:.2f}, {cy:.2f}) → 靠近边 {side_name}")
# =============================
# 绘图：仅展示符合条件的圆
# =============================

if len(valid_centers) == 0:
    print("❌ 未找到任何满足条件的圆。")
else:
    fig, ax = plt.subplots(figsize=figsize)

    # 绘制原始四边形
    poly_patch = Polygon(vertices, closed=True, edgecolor='black', facecolor='none', linewidth=2, label='封闭区域')
    ax.add_patch(poly_patch)

    # 标注顶点
    for idx, (x, y) in enumerate(vertices):
        ax.annotate(f'P{idx}\n({int(x)},{int(y)})', (x, y), xytext=(5, 5),
                    textcoords='offset points', fontsize=10, color='darkblue')

    # 绘制每个有效圆
    colors_cycle = ['red', 'green', 'blue', 'orange', 'purple']
    for idx, (cx, cy) in enumerate(valid_centers):
        circle = Circle((cx, cy), r, color=colors_cycle[idx % len(colors_cycle)],
                       fill=False, linewidth=2, linestyle='-')
        ax.add_patch(circle)
        ax.plot(cx, cy, 'o', color='black', markersize=5)
        ax.annotate(f'C{idx}', (cx, cy), xytext=(0, -10), textcoords='offset points',
                    fontsize=9, ha='center')

    # 设置图像
    all_x = vertices[:, 0]
    all_y = vertices[:, 1]
    padding = 10
    ax.set_xlim(all_x.min() - padding, all_x.max() + padding)
    ax.set_ylim(all_y.min() - padding, all_y.max() + padding)
    ax.set_aspect('equal')
    ax.grid(True, linestyle='--', alpha=0.5)
    ax.set_title(f'✅ 符合条件的圆（共 {len(valid_centers)} 个）\n'
                 f'半径={r}mm，最近边弧距=2mm，其余>2mm', fontsize=14)
    ax.set_xlabel("X (mm)")
    ax.set_ylabel("Y (mm)")

    if save_fig:
        plt.savefig(output_filename, dpi=150, bbox_inches='tight')
        print(f"\n📊 图像已保存为: {output_filename}")

    plt.tight_layout()
    plt.show()

    # 输出最终结果
    print("\n🎉 所有符合条件的圆心坐标列表（保留两位小数）：")
    for idx, (x, y) in enumerate(valid_centers):
        print(f"C{idx}: ({x:.2f}, {y:.2f})")

