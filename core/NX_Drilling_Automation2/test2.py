import math
from matplotlib import pyplot as plt
from matplotlib.patches import Circle


def compute_offset_inner_center(center, R, r2, offset=2.0, angle_deg=0):
    """
    Compute the center of a small circle so the distance between arcs = offset.

    Means:
        distance(big_center, small_center) = R - r2 - offset

    Args:
        center (tuple): (x,y) center of big circle
        R (float): radius of big circle
        r2 (float): radius of small circle
        offset (float): desired spacing between arcs, e.g., 2 mm
        angle_deg (float): direction of small circle

    Returns:
        (x2, y2)
    """
    if r2 <= 0 or R <= 0:
        raise ValueError("R 和 r2 必须为正数")
    if r2 + offset > R:
        raise ValueError("内圆 + 偏置 大于外圆，无法放置")

    theta = math.radians(angle_deg)
    d = R - r2 - offset  # 关键：圆心距离 = R - r2 - 2mm

    x2 = center[0] + d * math.cos(theta)
    y2 = center[1] + d * math.sin(theta)
    return (x2, y2)


def plot_circles_offset(center, R, r2, offset=2.0, angle_deg=0, show=True):
    """
    可视化展示：外圆 + 内圆（保持 offset mm 间距）
    """
    center2 = compute_offset_inner_center(center, R, r2, offset, angle_deg)

    fig, ax = plt.subplots(figsize=(6, 6))

    big = Circle(center, R, fill=False, linewidth=1.5)
    ax.add_patch(big)

    small = Circle(center2, r2, fill=False, linewidth=1.5)
    ax.add_patch(small)

    # 中心点
    ax.plot(center[0], center[1], 'o', markersize=6)
    ax.plot(center2[0], center2[1], 'x', markersize=6)

    # 连线
    ax.plot([center[0], center2[0]], [center[1], center2[1]], '--')

    ax.set_aspect('equal', adjustable='box')

    margin = max(R, r2) * 0.3 + 1
    ax.set_xlim(center[0] - R - margin, center[0] + R + margin)
    ax.set_ylim(center[1] - R - margin, center[1] + R + margin)

    ax.set_title(
        f"Arc spacing = {offset}mm, angle={angle_deg}°, inner center={center2}"
    )
    ax.set_xlabel("X")
    ax.set_ylabel("Y")

    if show:
        plt.show()

    return center2


# -------------- 示例运行 ----------------
if __name__ == "__main__":
    big_center = (0, 0)
    R = 8  # 外圆半径
    r2 = 3.5  # 内圆半径
    offset = 2  # 圆弧距离 = 2mm
    angle = 45   # 内圆方向角度

    c2 = plot_circles_offset(big_center, R, r2, offset, angle_deg=angle)
    print("内圆圆心：", c2)
