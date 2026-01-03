# -*- coding: utf-8 -*-
"""
路径优化模块
包含TSP路径优化算法
"""

import math
from utils import print_to_info_window
from geometry import GeometryHandler

class PathOptimizer:
    """路径优化器"""
    
    def __init__(self):
        pass
    
    def _euclidean_center(self, a, b):
        """计算两个圆心之间的欧氏距离"""
        c1, c2 = a.CenterPoint, b.CenterPoint
        return math.sqrt((c1.X - c2.X)**2 + (c1.Y - c2.Y)**2 + (c1.Z - c2.Z)**2)

    def _line_min_x_point(self, line):
        """计算直线最小的y值"""
        l1, l2 = line.StartPoint, line.EndPoint
        if l1.X < l2.X:
            return l1.X
        else:
            return l2.X

    def _line_min_y_point(self, line):
        """计算直线最小的y值"""
        l1, l2 = line.StartPoint, line.EndPoint
        if l1.Y < l2.Y:
            return l1.Y
        else:
            return l2.Y

    def _line_max_y_point(self, line):
        """计算直线最小的y值"""
        l1, l2 = line.StartPoint, line.EndPoint
        if l1.Y < l2.Y:
            return l1.Y
        else:
            return l2.Y

    def _euclidean_length(self, line):
        """计算一条直线的长度"""
        l1, l2 = line.StartPoint, line.EndPoint
        return math.sqrt((l1.X - l2.X)**2 + (l1.Y - l2.Y)**2 + (l1.Z - l2.Z)**2)

    def nearest_neighbor_tsp(self, circle_obj_list, start_idx=0):
        """最近邻贪心算法"""
        
        n = len(circle_obj_list)
        if n == 0:
            return []
        
        visited = [False] * n
        path = [start_idx]
        visited[start_idx] = True
        current = start_idx
        
        for _ in range(n - 1):
            nearest = None
            min_dist = float('inf')
            
            for j in range(n):
                if visited[j]:
                    continue
                d = self._euclidean_center(circle_obj_list[current][0], circle_obj_list[j][0])
                if d < min_dist:
                    min_dist = d
                    nearest = j
            
            if nearest is not None:
                path.append(nearest)
                visited[nearest] = True
                current = nearest
        
        return path

    def two_opt(self, path, circle_obj_list):
        """2-opt 优化路径"""
        
        improved = True
        while improved:
            improved = False
            for i in range(1, len(path) - 2):
                for j in range(i + 1, len(path)):
                    if j - i == 1:
                        continue
                    
                    d1 = (self._euclidean_center(circle_obj_list[path[i - 1]][0], circle_obj_list[path[i]][0]) + 
                          self._euclidean_center(circle_obj_list[path[j - 1]][0], circle_obj_list[path[j]][0]))
                    
                    d2 = (self._euclidean_center(circle_obj_list[path[i - 1]][0], circle_obj_list[path[j - 1]][0]) + 
                          self._euclidean_center(circle_obj_list[path[i]][0], circle_obj_list[path[j]][0]))
                    
                    if d2 < d1:
                        path[i:j] = reversed(path[i:j])
                        improved = True
        
        return path

    def optimize_drilling_path(self, circle_list, start_point=None):
        """钻孔路径优化主函数"""
        
        if not circle_list:
            print_to_info_window("❌ 未找到任何圆孔")
            return []

        # 坐标原点
        point = start_point or (0.0, 0.0, 0.0)
        # 获取起始点
        geometry_handler = GeometryHandler(None,None)  # 临时创建，仅用于查找
        nearest_circle = geometry_handler.find_nearest_circle(circle_list, point)
        
        if not nearest_circle:
            print_to_info_window("⚠️ 未找到起始圆")
            start_idx = 0
        else:
            # 将起始点替换到列表表头
            for idx in range(len(circle_list)):
                if ((round(circle_list[idx][1][0], 2) == round(nearest_circle[1][0], 2)) and 
                    (round(circle_list[idx][1][1], 2) == round(nearest_circle[1][1], 2)) and 
                    (round(circle_list[idx][1][2], 2) == round(nearest_circle[1][2], 2))):
                    temp = circle_list[0]
                    circle_list[0] = circle_list[idx]
                    circle_list[idx] = temp
                    break
            start_idx = 0

        # 路径优化
        idx_list = self.nearest_neighbor_tsp(circle_list, start_idx)
        idx_list = self.two_opt(idx_list, circle_list)
        
        optimization_list = []
        for idx in idx_list:
            optimization_list.append(circle_list[idx][0])
        
        print_to_info_window(f"✅ 共找到 {len(optimization_list)} 个圆孔，路径优化完成。")
        return optimization_list


def drilling_path_optimization(circle_list):
    """路径优化兼容函数（保持原有接口）"""
    optimizer = PathOptimizer()
    return optimizer.optimize_drilling_path(circle_list)