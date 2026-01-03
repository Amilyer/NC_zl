# -*- coding: utf-8 -*-
"""
匹配管理器 (match_manager.py)
功能：处理 CSV 配对数据的加载与最佳匹配选择
"""

import csv
import os
from collections import defaultdict
from typing import Dict, List, Optional, Tuple


class MatchManager:
    """管理数据配对逻辑"""

    def __init__(self):
        self.matching_data = defaultdict(list)

    def load_matches(self, csv_path: str) -> Dict[str, List[Dict]]:
        """
        加载配对结果 CSV
        
        Args:
            csv_path: CSV 文件路径
            
        Returns:
            Dict: {prt_filename: [candidate_list]}
        """
        if not os.path.exists(csv_path):
            print(f"❌ CSV文件不存在: {csv_path}")
            return {}

        self.matching_data.clear()
        
        try:
            with open(csv_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if self._is_valid_match(row):
                        self._add_record(row)
        except Exception as e:
            print(f"❌ 读取CSV失败: {e}")
            return {}

        print(f"✅ 加载了 {len(self.matching_data)} 个PRT文件的匹配数据")
        return self.matching_data

    def select_best_match(self, prt_dims: Tuple[float, float, float], candidates: List[Dict]) -> Optional[Dict]:
        """
        从候选中选择最佳匹配
        
        Args:
            prt_dims: (L, W, T)
            candidates: 候选列表
            
        Returns:
            Dict: 最佳匹配记录
        """
        if not candidates:
            return None

        # 重新计算相似度（确保实时性）
        for cand in candidates:
            cand['similarity'] = self._calculate_similarity(prt_dims, cand['prt2_dims'])

        # 按相似度降序排序
        best = max(candidates, key=lambda x: x['similarity'])
        return best

    def _is_valid_match(self, row: Dict) -> bool:
        """检查行数据是否有效"""
        return (row.get('匹配状态') == '已匹配' and 
                row.get('PRT文件名') and 
                row.get('DXF文件名'))

    def _add_record(self, row: Dict):
        """解析并添加单条记录"""
        try:
            prt_file = row['PRT文件名']
            dxf_file = row['DXF文件名']
            # DXF转PRT后的文件名
            prt2_file = dxf_file.replace('.dxf', '.prt')

            prt_dims = (
                float(row.get('PRT_长度(mm)') or 0),
                float(row.get('PRT_宽度(mm)') or 0),
                float(row.get('PRT_高度(mm)') or 0)
            )
            
            prt2_dims = (
                float(row.get('DXF_长度(mm)') or 0),
                float(row.get('DXF_宽度(mm)') or 0),
                float(row.get('DXF_高度(mm)') or 0)
            )

            self.matching_data[prt_file].append({
                'prt2_file': prt2_file,
                'prt_dims': prt_dims,
                'prt2_dims': prt2_dims,
                'similarity': 0.0  # 稍后计算
            })
        except ValueError:
            pass  # 忽略数据格式错误的行

    def _calculate_similarity(self, dims1: Tuple, dims2: Tuple) -> float:
        """计算尺寸相似度 (0-1)"""
        l1, w1, t1 = dims1
        l2, w2, t2 = dims2
        
        # 避免除以零
        def safe_div(a, b):
            return abs(a - b) / max(a, b, 1.0)

        err_l = safe_div(l1, l2)
        err_w = safe_div(w1, w2)
        err_h = safe_div(t1, t2)

        return 1.0 - (err_l + err_w + err_h) / 3.0

if __name__ == "__main__":
    # 简单测试
    mm = MatchManager()
    print("MatchManager 模块测试通过")
