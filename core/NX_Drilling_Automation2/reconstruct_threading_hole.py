import NXOpen
from typing import List, Tuple, Dict, Any
import NXOpen.Features
import NXOpen.GeometricUtilities
import math
from utils import analyze_arc


class RedCircleExtractor:
    def __init__(self, session, work_part: NXOpen.Part, target_color: int = 186):
        """
        红色闭合圆提取器 + 内圆圆心/半径精准计算器
        核心规则：
        1. 红色圆直径<4.000000mm：跳过
        2. 红色圆直径=4.000000mm：内圆直径3.000000mm，圆心与红色圆相同
        3. 红色圆直径>4.000000mm：内圆与外圆内缩2mm的辅助圆内切（仅1个点距离=2mm，其余>2mm）
        :param work_part: NX当前工作部件
        :param target_color: 目标颜色（默认红色186）
        """
        self.session = session
        self.work_part = work_part
        self.target_color = target_color
        self.coord_precision = 6  # 高精度坐标（6位小数，杜绝计算误差）
        self.red_closed_circles: List[Tuple[NXOpen.Arc, Tuple[float, float, float], float]] = []  # 红色闭合圆列表
        self.TARGET_DIST = 2.0  # 唯一匹配点的距离（固定2.0mm）
        self.SPECIAL_DIAMETER = 4.0  # 特殊判断直径阈值
        self.SPECIAL_INNER_DIAMETER = 3.0  # 红色圆=4mm时的内圆直径

    def _filter_red_closed_circles(self):
        """筛选并去重红色闭合圆"""
        seen_centers = set()
        self.red_closed_circles.clear()
        print("\n===== 开始筛选红色闭合圆（去重） =====")

        for curve in self.work_part.Curves:
            # 步骤1：仅处理Arc类型
            if not isinstance(curve, NXOpen.Arc):
                continue

            # 步骤2：筛选红色几何
            display_obj = curve if not hasattr(curve, 'GetDisplayableObject') else curve.GetDisplayableObject()
            if display_obj.Color != self.target_color:
                continue

            # 步骤3：筛选闭合圆
            valid_circle = analyze_arc(curve)
            if not valid_circle:
                continue

            # 步骤4：圆心去重（高精度）
            center = valid_circle.CenterPoint
            center_key = (
                round(center.X, self.coord_precision),
                round(center.Y, self.coord_precision),
                round(center.Z, self.coord_precision)
            )
            if center_key in seen_centers:
                print(f"  跳过重复圆心圆：{center_key}")
                continue

            # 筛选实线样式
            if str(curve.LineFont) != '1':
                continue

            # 步骤5：收集结果（圆对象、高精度圆心坐标、直径）
            diameter = round(valid_circle.Radius * 2, self.coord_precision)
            self.red_closed_circles.append((valid_circle, center_key, diameter))
            seen_centers.add(center_key)
            print(f"  新增红色闭合圆：圆心={center_key}，直径={diameter:.6f}mm")

        print(f"\n筛选完成：共提取{len(self.red_closed_circles)}个去重后的红色闭合圆")
        print("===== 红色闭合圆筛选完成 =====\n")

    def get_red_closed_circles(self) -> List[Tuple[NXOpen.Arc, Tuple[float, float, float], float]]:
        """
        对外提供的核心方法：获取所有去重后的红色闭合圆
        :return: 红色闭合圆列表，格式[(圆NX对象, 高精度圆心坐标, 高精度直径), ...]
        """
        self._filter_red_closed_circles()
        return self.red_closed_circles

    def _get_diameter_list(self, T: float) -> List[int]:
        """根据零件厚度T获取直径列表"""
        if T > 50:
            return [10, 7, 5, 3, 1]
        elif 30 < T <= 50:
            return [7, 5, 3, 1]
        else:  # T <= 30
            return [5, 3, 1]

    def _verify_tangency(self, outer_center: Tuple[float, float, float], inner_center: Tuple[float, float, float],
                         outer_radius: float, inner_radius: float) -> Tuple[bool, float]:
        """
        验证内圆与外圆内缩2mm的辅助圆是否内切（仅一个交点）
        :param outer_center: 外圆圆心
        :param inner_center: 内圆圆心
        :param outer_radius: 外圆半径
        :param inner_radius: 内圆半径
        :return: (是否内切, 切点到外圆的距离)
        """
        # 辅助圆半径：外圆内缩2mm
        assist_radius = round(outer_radius - self.TARGET_DIST, self.coord_precision)
        # 内外圆心距
        center_dist = round(
            math.hypot(inner_center[0] - outer_center[0], inner_center[1] - outer_center[1]),
            self.coord_precision
        )
        # 内切条件：圆心距 + 内圆半径 = 辅助圆半径（允许1e-6浮点误差）
        is_tangent = abs(center_dist + inner_radius - assist_radius) < 1e-6

        # 计算切点到外圆的距离（验证是否=2mm）
        if is_tangent:
            # 切点在两圆心连线上，距离外圆的距离=外圆半径 - 辅助圆半径=2mm
            tangent_dist = round(outer_radius - assist_radius, self.coord_precision)
        else:
            tangent_dist = 0.0

        return is_tangent, tangent_dist

    def calculate_inner_circle_params_precise(self, T: float) -> List[Dict[str, Any]]:
        """
        核心计算函数：
        - 红色圆直径<4.000000mm：跳过
        - 红色圆直径=4.000000mm：内圆直径3.000000mm，圆心与红色圆相同
        - 红色圆直径>4.000000mm：内圆与外圆内缩2mm的辅助圆内切（仅1个点距离=2mm，其余>2mm）
        仅计算坐标+半径，不创建实际圆
        :param T: 零件厚度
        :return: 内圆精准参数列表
        """
        if not self.red_closed_circles:
            print("\n❌ 无红色闭合圆，跳过内圆参数计算")
            return []

        diameter_list = self._get_diameter_list(T)

        inner_circle_params = []
        # 内圆圆心可选方向（保留原四个方向，可扩展）
        direction_angles = {
            "上左": math.pi * 3 / 4,  # 135°
            "下左": math.pi * 5 / 4,  # 225°
            "上右": math.pi * 1 / 4,  # 45°
            "下右": math.pi * 7 / 4  # 315°
        }

        # 遍历每个红色闭合圆
        for idx, (red_circle, red_center, red_diameter) in enumerate(self.red_closed_circles, 1):
            red_radius = round(red_diameter / 2, self.coord_precision)

            # ========== 新增：特殊判断逻辑 ==========
            # 1. 红色圆直径 < 4.000000mm，跳过
            if red_diameter < self.SPECIAL_DIAMETER - 1e-6:  # 浮点误差兼容
                print(f"  ❌ 红色圆直径({red_diameter:.6f}mm) < 4.000000mm，跳过该红色圆")
                continue
            # 2. 红色圆直径 = 4.000000mm，强制内圆参数
            elif abs(red_diameter - self.SPECIAL_DIAMETER) < 1e-6:
                inner_diameter = self.SPECIAL_INNER_DIAMETER
                inner_radius = round(inner_diameter / 2, self.coord_precision)
                inner_center = red_center  # 圆心与红色圆相同
                print(f"  ✅ 红色圆直径=4.000000mm，执行特殊规则：")
                print(f"    内圆直径固定为{inner_diameter:.6f}mm，圆心与红色圆相同：{inner_center}")

                # 封装特殊参数
                params = {
                    "red_circle_center": red_center,
                    "red_circle_diameter": red_diameter,
                    "red_circle_radius": red_radius,
                    "inner_circle_center": inner_center,
                    "inner_circle_radius": inner_radius,
                    "inner_circle_diameter": inner_diameter,
                    "matched_direction": "中心重合",
                    "tangent_point_distance": self.TARGET_DIST,
                    "verification": f"红色圆直径=4.000000mm，内圆直径固定为3.000000mm，圆心与红色圆重合"
                }
                inner_circle_params.append(params)
                continue
            # ========== 特殊判断结束 ==========

            # 辅助圆半径（外圆内缩2mm）
            assist_radius = round(red_radius - self.TARGET_DIST, self.coord_precision)
            if assist_radius <= 0:
                print(f"  ❌ 辅助圆半径({assist_radius:.6f}mm)≤0，无法找到内圆，跳过")
                continue

            matched_d = None
            matched_direction = ""
            inner_center = None
            inner_radius = 0.0
            tangent_point_dist = 0.0

            # 遍历候选直径列表，找第一个符合条件的内圆直径
            for d in diameter_list:
                current_inner_radius = round(d / 2, self.coord_precision)

                # 验证必要条件：辅助圆半径 ≥ 内圆半径（否则无法内切）
                if assist_radius < current_inner_radius:
                    continue

                # 计算内圆圆心到外圆圆心的距离（内切条件）
                center_dist = round(assist_radius - current_inner_radius, self.coord_precision)

                # 遍历方向找符合条件的内圆圆心
                for dir_name, angle in direction_angles.items():
                    # 计算该方向的内圆圆心坐标
                    dx = round(center_dist * math.cos(angle), self.coord_precision)
                    dy = round(center_dist * math.sin(angle), self.coord_precision)
                    current_inner_center = (
                        round(red_center[0] + dx, self.coord_precision),
                        round(red_center[1] + dy, self.coord_precision),
                        red_center[2]
                    )

                    # 验证内切条件（仅一个交点）
                    is_tangent, dist = self._verify_tangency(
                        red_center, current_inner_center, red_radius, current_inner_radius
                    )
                    if not is_tangent:
                        continue

                    # 验证其余点距离>2mm（内切几何特性已保证，此处二次验证）
                    # 取内圆上任意非切点（如圆心垂直方向）验证距离
                    test_angle = angle + math.pi / 2  # 垂直方向
                    test_dx = round(current_inner_radius * math.cos(test_angle), self.coord_precision)
                    test_dy = round(current_inner_radius * math.sin(test_angle), self.coord_precision)
                    test_point = (
                        round(current_inner_center[0] + test_dx, self.coord_precision),
                        round(current_inner_center[1] + test_dy, self.coord_precision),
                        current_inner_center[2]
                    )
                    # 测试点到外圆的距离 = 外圆半径 - 测试点到外圆圆心的距离
                    test_dist = round(
                        red_radius - math.hypot(test_point[0] - red_center[0], test_point[1] - red_center[1]),
                        self.coord_precision
                    )

                    if test_dist <= self.TARGET_DIST:
                        continue

                    # 所有条件满足
                    matched_d = d
                    matched_direction = dir_name
                    inner_center = current_inner_center
                    inner_radius = current_inner_radius
                    tangent_point_dist = dist
                    break

                if matched_d is not None:
                    break

            if matched_d is None:
                print(f"  ❌ 无符合条件的内圆直径，跳过该红色圆")
                continue

            # 封装最终参数
            params = {
                "red_circle_center": red_center,
                "red_circle_diameter": red_diameter,
                "red_circle_radius": red_radius,
                "inner_circle_center": inner_center,
                "inner_circle_radius": inner_radius,
                "inner_circle_diameter": matched_d,
                "matched_direction": matched_direction,  # 内圆圆心方向
                "tangent_point_distance": tangent_point_dist,  # 唯一匹配点距离（=2mm）
                "verification": f"内圆与外圆内缩2mm的辅助圆内切，仅1个点距离=2mm，其余点距离>2mm"
            }
            inner_circle_params.append(params)

        print(f"\n===== 内圆参数精准计算完成：共找到{len(inner_circle_params)}组有效参数 =====\n")
        return inner_circle_params
