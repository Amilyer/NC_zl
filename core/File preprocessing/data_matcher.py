# -*- coding: utf-8 -*-
"""
数据配对模块 - 用于将DXF(2D图纸)和PRT(3D模型)文件进行智能匹配

主要功能：
- 三层筛选匹配机制（文件名匹配、尺寸匹配、相似度匹配）
- 支持渐进式容差匹配
- 基于分类和相似度的智能匹配算法

主函数: match_data(dxf_csv, prt_csv, output_csv, tolerance=1.0) -> Optional[str]
"""

import os
import re
from typing import Optional, Tuple, List, Dict

# 检查依赖
try:
    import pandas as pd
    _PANDAS_AVAILABLE = True
except ImportError:
    pd = None
    _PANDAS_AVAILABLE = False


# ==============================================================================
# 内部辅助函数（不对外暴露）
# ==============================================================================

def _reorder_dimensions(length: float, width: float, height: float) -> Tuple[float, float, float]:
    """
    对长宽高进行重排，确保满足 长 ≥ 宽 ≥ 高
    
    Args:
        length: 原始长度
        width: 原始宽度  
        height: 原始高度
        
    Returns:
        Tuple[float, float, float]: 重排后的 (长度, 宽度, 高度)，满足 长 ≥ 宽 ≥ 高
    """
    # 将三个维度放入列表并排序（降序）
    dimensions = [length, width, height]
    dimensions.sort(reverse=True)
    
    # 返回重排后的结果：长 ≥ 宽 ≥ 高
    return (dimensions[0], dimensions[1], dimensions[2])


def _extract_core_pattern(filename) -> str:
    """提取文件名核心模式用于匹配"""
    if pd.isna(filename):
        return ""
    filename = re.sub(r'^\d+_', '', str(filename))  # 移除前缀数字
    filename = re.sub(r'\.(dxf|prt)$', '', filename, flags=re.IGNORECASE)  # 移除扩展名
    return filename.upper()


def _three_dimensions_match(
    prt_dims: Tuple[float, float, float], 
    dxf_dims: Tuple[float, float, float], 
    tolerance: float = 1.0
) -> bool:
    """
    三维精准匹配 - 检查PRT和DXF的三个维度是否都在容差范围内匹配
    
    匹配规则：
    1. 首先检查2D图(DXF)的每个维度+0.1mm是否都大于等于3D图(PRT)的对应维度
    2. 然后检查三个维度是否都在容差范围内匹配（DXF尺寸不超过PRT尺寸+容差）
    3. 三个维度都必须同时满足上述条件
    
    Args:
        prt_dims: PRT的长宽高 (L, W, T)
        dxf_dims: DXF的长宽高 (L, W, T)
        tolerance: 容差值，默认为1.0mm
        
    Returns:
        bool: 如果三个维度都匹配则返回True，否则返回False
    """
    # 检查是否有无效数据
    if not all(dim > 0 for dim in prt_dims) or not all(dim > 0 for dim in dxf_dims):
        return False

    # 获取PRT和DXF的尺寸
    prt_L, prt_W, prt_T = prt_dims
    dxf_L, dxf_W, dxf_T = dxf_dims

    # 首先检查2D图(DXF)的每个维度+0.1mm是否都大于等于3D图(PRT)的对应维度
    if not (dxf_L + 0.1 >= prt_L and dxf_W + 0.1 >= prt_W and dxf_T + 0.1 >= prt_T):
        return False

    # 检查三个维度是否都在容差范围内匹配（DXF尺寸不超过PRT尺寸+容差）
    match_L = (dxf_L - prt_L <= tolerance)  # 长度对比
    match_W = (dxf_W - prt_W <= tolerance)  # 宽度对比
    match_T = (dxf_T - prt_T <= tolerance)  # 高度对比

    # 三个维度都必须匹配
    return match_L and match_W and match_T





def _dimensions_similar(
    prt_dims: Tuple[float, float, float], 
    dxf_dims: Tuple[float, float, float], 
    tolerance: float = 1.0
) -> bool:
    """
    比较PRT和DXF的尺寸是否在容差范围内相似。
    规则：2D图(DXF)对应的长度+0.1mm要大于等于3D图(PRT)对应的长度，
         2D图(DXF)对应的宽度+0.1mm要大于等于3D图(PRT)对应的宽度，
         2D图(DXF)对应的高度+0.1mm要大于等于3D图(PRT)对应的高度，
          在此基础之上，如果至少有两个维度在容差范围内匹配则返回True，否则返回False。
    
    Args:
        prt_dims: PRT的长宽高 (L, W, T)
        dxf_dims: DXF的长宽高 (L, W, T)
        tolerance: 容差值，默认为1.0mm
        
    Returns:
        bool: 如果满足条件（DXF各维度+0.1mm>=PRT对应维度且至少两个维度在容差内匹配）则返回True，否则返回False
    """
    # 检查是否有无效数据
    if not all(dim > 0 for dim in prt_dims) or not all(dim > 0 for dim in dxf_dims):
        return False

    # 获取PRT和DXF的尺寸
    prt_L, prt_W, prt_T = prt_dims
    dxf_L, dxf_W, dxf_T = dxf_dims

    # 首先检查2D图(DXF)的每个维度+0.1mm是否都大于等于3D图(PRT)的对应维度
    if not (dxf_L + 0.1 >= prt_L and dxf_W + 0.1 >= prt_W and dxf_T + 0.1 >= prt_T):
        return False

    # 按照长度、宽度、高度一一对应比较，检查是否在容差范围内（DXF尺寸不超过PRT尺寸+容差）
    match_L = (dxf_L - prt_L <= tolerance)  # 长度对比
    match_W = (dxf_W - prt_W <= tolerance)  # 宽度对比
    match_T = (dxf_T - prt_T <= tolerance)  # 高度对比

    # 计算匹配的维度数量
    match_count = sum([match_L, match_W, match_T])
    
    # 如果至少有两个维度匹配，则返回 True
    return match_count >= 2


def _classify_files_by_category(filename: str) -> str:
    """
    根据文件名中的关键词对文件进行分类，从文件名中动态提取分类数据
    
    规则：
    1. 基于_extract_core_pattern操作，筛选第一个连接线"-"之前的数据
    2. 区分"-"和下划线，没有"-"即使有下划线也不提取
    3. 过滤掉像'UNPARAMETERIZED_FEATURE9.prt'或'JOIN.16.prt'这样的无效数据
    4. 过滤纯数字或长度小于2的无效分类
    
    Args:
        filename: 文件名
        
    Returns:
        str: 分类标签 (从文件名中提取的有效分类，或 OTHER)
    """
    if pd.isna(filename):
        return "OTHER"
        
    filename = str(filename)
    
    # 首先提取核心模式（移除前缀数字和扩展名）
    core_pattern = _extract_core_pattern(filename)
    
    # 检查是否包含连接线"-"，没有则返回OTHER
    if '-' not in core_pattern:
        return "OTHER"
    
    # 提取第一个"-"之前的数据
    category_part = core_pattern.split('-')[0].strip()
    
    # 过滤无效数据：如果提取的部分包含FEATURE、JOIN、UNPARAMETERIZED等关键词，视为无效
    invalid_keywords = ['FEATURE', 'JOIN', 'UNPARAMETERIZED', 'PARAMETERIZED']
    for invalid_word in invalid_keywords:
        if invalid_word in category_part:
            return "OTHER"
    
    # 过滤掉纯数字或太短的数据
    if len(category_part) < 2 or category_part.isdigit():
        return "OTHER"
    
    # 返回有效的分类数据
    return category_part


def _progressive_matching(
    prt_dims: Tuple[float, float, float], 
    dxf_dims: Tuple[float, float, float]
) -> Optional[str]:
    """
    基于维度组合的渐进式匹配函数
    
    算法说明：
    1. 首先检查基本尺寸条件：2D各维度+0.1mm ≥ 3D对应维度
    2. 计算体积相似度：(1 - |体积差|/PRT体积) × 100%
    3. 计算形状比例相似度：基于长宽比、宽高比、长高比的差异
    4. 综合相似度 = 体积相似度×0.6 + 形状相似度×0.4
    5. 渐进式阈值检查：95%、90%、85%、80%、75%
    
    Args:
        prt_dims: PRT的长宽高 (L, W, T) - 已排序，长≥宽≥高
        dxf_dims: DXF的长宽高 (L, W, T) - 已排序，长≥宽≥高
        
    Returns:
        str: 匹配规则名称，如"第二层-维度组合90%"，如果未匹配则返回None
    """
    # 检查是否有无效数据
    if not all(dim > 0 for dim in prt_dims) or not all(dim > 0 for dim in dxf_dims):
        return None

    prt_L, prt_W, prt_T = prt_dims
    dxf_L, dxf_W, dxf_T = dxf_dims
    
    # 检查基本尺寸条件：2D各维度+0.1mm >= 3D对应维度
    if not (dxf_L + 0.1 >= prt_L and dxf_W + 0.1 >= prt_W and dxf_T + 0.1 >= prt_T):
        return None
    
    # 计算体积相似度
    prt_volume = prt_L * prt_W * prt_T
    dxf_volume = dxf_L * dxf_W * dxf_T
    
    # 体积相似度 (0-100%，越高表示越相似)
    volume_diff = abs(prt_volume - dxf_volume)
    volume_similarity = max(0, 100 - (volume_diff / prt_volume * 100)) if prt_volume > 0 else 0
    
    # 计算形状比例相似度
    # PRT形状比例
    prt_ratio_LW = prt_L / prt_W if prt_W > 0 else 1.0
    prt_ratio_WT = prt_W / prt_T if prt_T > 0 else 1.0
    prt_ratio_LT = prt_L / prt_T if prt_T > 0 else 1.0
    
    # DXF形状比例
    dxf_ratio_LW = dxf_L / dxf_W if dxf_W > 0 else 1.0
    dxf_ratio_WT = dxf_W / dxf_T if dxf_T > 0 else 1.0
    dxf_ratio_LT = dxf_L / dxf_T if dxf_T > 0 else 1.0
    
    # 计算各比例差异
    ratio_LW_diff = abs(prt_ratio_LW - dxf_ratio_LW)
    ratio_WT_diff = abs(prt_ratio_WT - dxf_ratio_WT)
    ratio_LT_diff = abs(prt_ratio_LT - dxf_ratio_LT)
    
    # 形状相似度 (0-100%，越高表示越相似)
    avg_ratio_diff = (ratio_LW_diff + ratio_WT_diff + ratio_LT_diff) / 3
    shape_similarity = max(0, 100 - avg_ratio_diff * 10)  # 调整系数，使合理的比例差异得到适当的相似度
    
    # 综合相似度评分 (体积相似度权重60%，形状相似度权重40%)
    overall_similarity = volume_similarity * 0.6 + shape_similarity * 0.4
    
    # 渐进式相似度阈值检查（从95%到75%，步长5%）
    for similarity_threshold in range(95, 74, -5):
        if overall_similarity >= similarity_threshold:
            return f"第二层-维度组合{similarity_threshold}%"
    
    return None


def _create_match_record(prt_row, dxf_row=None, status='已匹配', match_priority='') -> dict:
    """创建匹配记录"""
    record = {
        'PRT文件名': prt_row['文件名'] if prt_row is not None else '',
        'PRT_长度(mm)': prt_row['长度_L (mm)'] if prt_row is not None else '',
        'PRT_宽度(mm)': prt_row['宽度_W (mm)'] if prt_row is not None else '',
        'PRT_高度(mm)': prt_row['高度_T (mm)'] if prt_row is not None else '',
        '': '',  # 分隔列
        'DXF文件名': dxf_row['文件名'] if dxf_row is not None else '',
        'DXF_长度(mm)': dxf_row['长度_L (mm)'] if dxf_row is not None else '',
        'DXF_宽度(mm)': dxf_row['宽度_W (mm)'] if dxf_row is not None else '',
        'DXF_高度(mm)': dxf_row['高度_T (mm)'] if dxf_row is not None else '',
        '匹配优先级': match_priority if match_priority else ('文件名+尺寸匹配' if status == '已匹配' and not match_priority else ''),
        '匹配状态': status
    }
    return record


def _filename_match_first_layer_rule(prt_pattern: str, dxf_pattern: str) -> bool:
    """
    第一层文件名匹配规则：双向核心模式包含关系
    
    Args:
        prt_pattern: PRT文件核心模式
        dxf_pattern: DXF文件核心模式
        
    Returns:
        bool: 是否匹配成功
    """
    return prt_pattern and dxf_pattern and (prt_pattern in dxf_pattern or dxf_pattern in prt_pattern)


def _first_layer_matching(df_dxf, df_prt, tolerance: float = 1.0) -> tuple:
    """
    第一层筛选：文件名核心模式匹配 + 三维精准匹配
    
    处理流程：
    1. 遍历所有未匹配的PRT文件
    2. 对每个PRT文件：
       - 检查分类，如果为OTHER则标记为"未匹配-PRT"并跳过
       - 在所有未匹配的DXF文件中查找文件名匹配的候选
       - 对文件名匹配的DXF进行三维尺寸精准匹配检查
       - 找到第一个三维匹配后立即停止搜索（优先匹配策略）
    
    匹配规则（在满足文件名匹配的前提下）：
    1. 文件名核心模式双向包含检查（PRT模式在DXF模式中或反之）
    2. 过滤分类为OTHER的文件，不参与匹配
    3. 只进行三维精确匹配：DXF各维度+0.1mm≥PRT对应维度且三个维度都在容差范围内
    4. 找到第一个三维匹配后立即停止搜索（优先匹配策略）
    5. 匹配成功的优先级为："文件名+三维"
    
    Args:
        df_dxf: DXF数据DataFrame（包含已匹配状态）
        df_prt: PRT数据DataFrame（包含已匹配状态）
        tolerance: 尺寸匹配容差值
        
    Returns:
        tuple: (matched_records, df_dxf, df_prt) - 匹配记录列表和更新后的DataFrame
    """
    matched_records = []
    
    # 第一层筛选：文件名核心模式匹配 + 三维尺寸匹配
    for prt_idx, prt_row in df_prt.iterrows():
        prt_pattern = prt_row['核心模式']

        # 检查PRT是否已经匹配
        if df_prt.loc[prt_idx, '已匹配']:
            continue

        # 获取PRT的尺寸数据
        try:
            prt_length = float(prt_row['长度_L (mm)']) if prt_row['长度_L (mm)'] else 0.0
            prt_width = float(prt_row['宽度_W (mm)']) if prt_row['宽度_W (mm)'] else 0.0
            prt_height = float(prt_row['高度_T (mm)']) if prt_row['高度_T (mm)'] else 0.0
            prt_dims = (prt_length, prt_width, prt_height)
        except (ValueError, TypeError):
            prt_dims = (0.0, 0.0, 0.0)

        # 对PRT进行分类检查，如果为OTHER则不参与匹配
        prt_category = _classify_files_by_category(prt_row['文件名'])
        if prt_category == "OTHER":
            # 直接标记为未匹配-PRT，匹配优先级为"文件名模糊"
            matched_records.append(_create_match_record(prt_row, None, '未匹配-PRT', '文件名模糊'))
            continue

        # 收集所有可能的匹配候选（只考虑三维匹配）
        best_match = None
        best_match_dxf_idx = -1

        for dxf_idx, dxf_row in df_dxf.iterrows():
            if dxf_row['已匹配']:
                continue

            # 对DXF进行分类检查，如果为OTHER则不参与匹配
            dxf_category = _classify_files_by_category(dxf_row['文件名'])
            if dxf_category == "OTHER":
                continue

            # 转换文件名为大写用于后续比较（保持一致性）
            dxf_filename = str(dxf_row['文件名']).upper()
            # 提取DXF文件核心模式用于文件名匹配
            dxf_pattern = _extract_core_pattern(dxf_row['文件名'])
            # 文件名匹配检查 - 使用第一层规则
            if _filename_match_first_layer_rule(prt_pattern, dxf_pattern):
                # 获取DXF的尺寸数据
                try:
                    dxf_length = float(dxf_row['长度_L (mm)']) if dxf_row['长度_L (mm)'] else 0.0
                    dxf_width = float(dxf_row['宽度_W (mm)']) if dxf_row['宽度_W (mm)'] else 0.0
                    dxf_height = float(dxf_row['高度_T (mm)']) if dxf_row['高度_T (mm)'] else 0.0
                    dxf_dims = (dxf_length, dxf_width, dxf_height)
                except (ValueError, TypeError):
                    dxf_dims = (0.0, 0.0, 0.0)
                
                # 三维尺寸匹配检查 - 只进行三维精确匹配
                if (_three_dimensions_match(prt_dims, dxf_dims, tolerance)):
                    # 找到三维匹配
                    best_match = dxf_row
                    best_match_dxf_idx = dxf_idx
                    break  # 找到三维匹配后立即停止搜索
        
        # 应用最佳匹配结果（只处理三维匹配）
        if best_match is not None:
            matched_records.append(_create_match_record(prt_row, best_match, '已匹配', '文件名+三维'))
            df_dxf.loc[best_match_dxf_idx, '已匹配'] = True
            df_prt.loc[prt_idx, '已匹配'] = True  # 标记PRT为已匹配

    return matched_records, df_dxf, df_prt


def _middle_layer_matching(df_dxf, df_prt, first_layer_matched) -> tuple:
    """
    中间层筛选：文件名匹配 + 三维渐进式最大差值匹配
    
    核心特征：
    - 三维尺寸匹配：长度、宽度、高度全部参与匹配
    - 最大差值原则：以三个维度差值中的最大值作为判断标准
    - 渐进式容差：从2mm到10mm逐步放宽（共9级）
    
    处理流程：
    1. 获取第一层筛选后仍未匹配的DXF和PRT数据
    2. 过滤掉分类为OTHER的文件
    3. 对每个未匹配的PRT文件：
       - 在所有未匹配的DXF文件中查找文件名匹配的候选
       - 对通过文件名匹配的DXF进行三维渐进式尺寸差检查
       - 从容差2mm到10mm逐级检查，找到匹配立即停止
    
    匹配规则：
    1. 使用与第一层相同的文件名核心模式匹配规则（双向包含）
    2. 对文件名匹配的文件进行三维渐进式尺寸差检查（2-10mm）
    3. 要求2D图各维度+0.1mm ≥ 3D图对应维度
    4. 取三个维度差值的最大值进行容差判断（最大差值原则）
    5. 匹配优先级为："中间层{X}mm"（X为2-10）
    
    Args:
        df_dxf: DXF数据DataFrame（包含已匹配状态）
        df_prt: PRT数据DataFrame（包含已匹配状态）
        first_layer_matched: 第一层筛选的匹配记录
        
    Returns:
        tuple: (matched_records, df_dxf, df_prt) - 匹配记录列表和更新后的DataFrame
    """
    matched_records = first_layer_matched.copy()
    
    # 获取未匹配的DXF和PRT数据
    unmatched_dxf_df = df_dxf[~df_dxf['已匹配']].copy()
    unmatched_prt_df = df_prt[~df_prt['已匹配']].copy()
    
    # 过滤掉分类为OTHER的文件，它们不参与中间层匹配
    unmatched_dxf_df = unmatched_dxf_df[unmatched_dxf_df['文件名'].apply(_classify_files_by_category) != 'OTHER']
    unmatched_prt_df = unmatched_prt_df[unmatched_prt_df['文件名'].apply(_classify_files_by_category) != 'OTHER']
    
    # 处理未匹配的PRT数据
    for prt_idx, prt_row in unmatched_prt_df.iterrows():
        # 检查这个PRT是否已经匹配
        if prt_row['文件名'] in [m['PRT文件名'] for m in matched_records]:
            continue
            
        # 获取PRT的尺寸数据
        try:
            prt_length = float(prt_row['长度_L (mm)']) if prt_row['长度_L (mm)'] else 0.0
            prt_width = float(prt_row['宽度_W (mm)']) if prt_row['宽度_W (mm)'] else 0.0
            prt_height = float(prt_row['高度_T (mm)']) if prt_row['高度_T (mm)'] else 0.0
            prt_dims = (prt_length, prt_width, prt_height)
        except (ValueError, TypeError):
            prt_dims = (0.0, 0.0, 0.0)
            
        prt_pattern = prt_row['核心模式']
        
        # 查找匹配的DXF文件
        for dxf_idx, dxf_row in unmatched_dxf_df.iterrows():
            # 检查这个DXF是否已经匹配
            if dxf_row['文件名'] in [m['DXF文件名'] for m in matched_records if m['匹配状态'] == '已匹配' and m['DXF文件名']]:
                continue
                
            dxf_filename = str(dxf_row['文件名']).upper()
            dxf_pattern = _extract_core_pattern(dxf_row['文件名'])
            
            # 文件名匹配检查 - 使用与第一层完全相同的规则
            if _filename_match_first_layer_rule(prt_pattern, dxf_pattern):
                # 获取DXF的尺寸数据
                try:
                    dxf_length = float(dxf_row['长度_L (mm)']) if dxf_row['长度_L (mm)'] else 0.0
                    dxf_width = float(dxf_row['宽度_W (mm)']) if dxf_row['宽度_W (mm)'] else 0.0
                    dxf_height = float(dxf_row['高度_T (mm)']) if dxf_row['高度_T (mm)'] else 0.0
                    dxf_dims = (dxf_length, dxf_width, dxf_height)
                except (ValueError, TypeError):
                    dxf_dims = (0.0, 0.0, 0.0)
                
                # 检查基本尺寸条件：2D各维度+0.1mm >= 3D对应维度
                if not (dxf_length + 0.1 >= prt_length and dxf_width + 0.1 >= prt_width and dxf_height + 0.1 >= prt_height):
                    continue
                
                # 计算三个维度的差值（取绝对值，表示尺寸差异程度）
                length_diff = abs(prt_length - dxf_length)
                width_diff = abs(prt_width - dxf_width)
                height_diff = abs(prt_height - dxf_height)
                
                # 找出三个维度差值中的最大值作为匹配判断依据
                max_diff = max(length_diff, width_diff, height_diff)
                
                # 渐进式匹配检查（从容差2mm到10mm，步长1mm）
                for tolerance in range(2, 11):
                    if max_diff <= tolerance:
                        # 中间层匹配成功
                        matched_records.append(_create_match_record(prt_row, dxf_row, '已匹配', f'中间层{tolerance}mm'))
                        # 标记DXF和PRT为已匹配
                        original_dxf_idx = df_dxf[df_dxf['文件名'] == dxf_row['文件名']].index[0]
                        df_dxf.loc[original_dxf_idx, '已匹配'] = True
                        df_prt.loc[prt_idx, '已匹配'] = True
                        break
                
                # 如果已经找到匹配，跳出循环
                if prt_row['文件名'] in [m['PRT文件名'] for m in matched_records]:
                    break
    
    return matched_records, df_dxf, df_prt


def _second_layer_matching(df_dxf, df_prt, middle_layer_matched) -> tuple:
    """
    第二层筛选：基于分类的渐进式维度组合匹配
    
    处理流程：
    1. 获取中间层筛选后仍未匹配的DXF和PRT数据
    2. 为数据添加分类标签（基于文件名动态提取）
    3. 过滤掉分类为OTHER的文件
    4. 对每个未匹配的PRT文件：
       - 查找同一类别的未匹配DXF文件
       - 使用_progressive_matching函数进行相似度计算
       - 找到匹配后立即停止搜索
    5. 处理未匹配的DXF文件，分类为OTHER的标记特殊优先级
    
    匹配规则：
    1. 分别对未匹配的DXF和PRT文件按类别分类（基于文件名关键词）
    2. 对同一类别的文件进行体积相似度和形状比例相似度计算
    3. 综合相似度 = 体积相似度×0.6 + 形状相似度×0.4
    4. 渐进式相似度阈值检查（95%、90%、85%、80%、75%）
    5. 过滤分类为OTHER的文件，不参与第二层匹配
    6. 匹配优先级为："第二层-维度组合{X}%"（X为75-95）
    
    Args:
        df_dxf: DXF数据DataFrame（包含已匹配状态）
        df_prt: PRT数据DataFrame（包含已匹配状态）
        middle_layer_matched: 中间层筛选的匹配记录
        
    Returns:
        tuple: (all_matched_records, unmatched_dxf_records) - 所有匹配记录和未匹配DXF记录
    """
    matched_records = middle_layer_matched.copy()
    
    # 中间层筛选：对未匹配的数据进行文件名匹配和三维渐进式尺寸匹配（最大差值原则）
    
    # 获取未匹配的DXF
    unmatched_dxf_df = df_dxf[~df_dxf['已匹配']].copy()
    
    # 为未匹配的数据添加分类标签
    unmatched_dxf_df['分类'] = unmatched_dxf_df['文件名'].apply(_classify_files_by_category)
    
    # 获取未匹配的PRT数据（注意：这里只处理真正未匹配的PRT，即不在matched列表中的PRT）
    unmatched_prt_df = df_prt[~df_prt['已匹配']].copy()
    unmatched_prt_df['分类'] = unmatched_prt_df['文件名'].apply(_classify_files_by_category)
    
    # 过滤掉分类为OTHER的文件，它们不参与第二层匹配
    unmatched_dxf_df = unmatched_dxf_df[unmatched_dxf_df['分类'] != 'OTHER']
    unmatched_prt_df = unmatched_prt_df[unmatched_prt_df['分类'] != 'OTHER']
    
    # 处理未匹配的PRT数据
    for prt_idx, prt_row in unmatched_prt_df.iterrows():
        # 检查这个PRT是否已经匹配或者已经在matched列表中（双重保险，避免重复处理已在matched列表中的PRT）
        if prt_row['文件名'] in [m['PRT文件名'] for m in matched_records]:
            continue
            
        # 获取PRT的尺寸数据
        try:
            prt_length = float(prt_row['长度_L (mm)']) if prt_row['长度_L (mm)'] else 0.0
            prt_width = float(prt_row['宽度_W (mm)']) if prt_row['宽度_W (mm)'] else 0.0
            prt_height = float(prt_row['高度_T (mm)']) if prt_row['高度_T (mm)'] else 0.0
            prt_dims = (prt_length, prt_width, prt_height)
        except (ValueError, TypeError):
            prt_dims = (0.0, 0.0, 0.0)
            
        # 对PRT进行分类
        prt_category = prt_row['分类']
        
        # 查找同一分类的未匹配DXF文件
        same_category_dxf = unmatched_dxf_df[unmatched_dxf_df['分类'] == prt_category]
        
        found_second_layer = False
        for dxf_idx, dxf_row in same_category_dxf.iterrows():
            # 检查这个DXF是否已经被匹配（防止重复匹配）
            if dxf_row['文件名'] in [m['DXF文件名'] for m in matched_records if m['匹配状态'] == '已匹配' and m['DXF文件名']]:
                continue
                
            # 获取DXF的尺寸数据
            try:
                dxf_length = float(dxf_row['长度_L (mm)']) if dxf_row['长度_L (mm)'] else 0.0
                dxf_width = float(dxf_row['宽度_W (mm)']) if dxf_row['宽度_W (mm)'] else 0.0
                dxf_height = float(dxf_row['高度_T (mm)']) if dxf_row['高度_T (mm)'] else 0.0
                dxf_dims = (dxf_length, dxf_width, dxf_height)
            except (ValueError, TypeError):
                dxf_dims = (0.0, 0.0, 0.0)
            
            # 进行渐进式匹配
            match_rule = _progressive_matching(prt_dims, dxf_dims)
            if match_rule:
                # 第二层匹配成功
                matched_records.append(_create_match_record(prt_row, dxf_row, '已匹配', match_rule))
                # 标记DXF和PRT为已匹配
                # 使用索引标记原始DataFrame中的条目
                original_dxf_idx = df_dxf[df_dxf['文件名'] == dxf_row['文件名']].index[0]
                df_dxf.loc[original_dxf_idx, '已匹配'] = True
                df_prt.loc[prt_idx, '已匹配'] = True
                found_second_layer = True
                break
                
        if not found_second_layer:
            # 仍然未匹配，但要确保不会重复添加
            if prt_row['文件名'] not in [m['PRT文件名'] for m in matched_records]:
                matched_records.append(_create_match_record(prt_row, None, '未匹配-PRT'))

    # 未匹配的DXF（包括分类为OTHER的DXF文件）
    unmatched_dxf_records = []
    for _, row in df_dxf[~df_dxf['已匹配']].iterrows():
        # 对于分类为OTHER的DXF文件，标记为未匹配-DXF，匹配优先级为"文件名模糊"
        dxf_category = _classify_files_by_category(row['文件名'])
        if dxf_category == "OTHER":
            unmatched_dxf_records.append(_create_match_record(None, row, '未匹配-DXF', '文件名模糊'))
        else:
            unmatched_dxf_records.append(_create_match_record(None, row, '未匹配-DXF'))

    return matched_records, unmatched_dxf_records


def _do_matching(df_dxf, df_prt, tolerance: float = 1.0) -> tuple:
    """
    执行完整的三层筛选匹配逻辑
    
    预处理阶段：
    1. 对DXF数据的长宽高进行重排（长≥宽≥高）
    2. 添加核心模式（用于文件名匹配）
    3. 添加已匹配标记（防止重复匹配）
    
    三层筛选流程：
    1. 第一层：文件名匹配 + 三维精准匹配（容差默认为1mm）
       - 最严格的匹配，要求文件名和三维尺寸都匹配
       - 匹配成功立即标记，不参与后续筛选
    2. 中间层：文件名匹配 + 渐进式最大差值匹配（2-10mm）
       - 对第一层未匹配的文件进行更宽松的尺寸匹配
       - 取三个维度差值的最大值进行容差判断
    3. 第二层：分类匹配 + 渐进式维度组合匹配（75-95%相似度）
       - 对仍未匹配的文件按分类进行相似度计算
       - 基于体积相似度和形状比例相似度的综合评分
    
    Args:
        df_dxf: DXF数据DataFrame
        df_prt: PRT数据DataFrame
        tolerance: 第一层筛选的尺寸匹配容差值，默认为1.0mm
        
    Returns:
        tuple: (all_matched_records, unmatched_dxf_records) - 所有匹配记录和未匹配DXF记录
    """
    # 预处理：对DXF数据的长宽高进行重排，确保满足 长 ≥ 宽 ≥ 高
    for idx, row in df_dxf.iterrows():
        try:
            length = float(row['长度_L (mm)']) if row['长度_L (mm)'] else 0.0
            width = float(row['宽度_W (mm)']) if row['宽度_W (mm)'] else 0.0
            height = float(row['高度_T (mm)']) if row['高度_T (mm)'] else 0.0
            
            # 对长宽高进行重排
            reordered_dims = _reorder_dimensions(length, width, height)
            
            # 更新DataFrame中的数据
            df_dxf.loc[idx, '长度_L (mm)'] = reordered_dims[0]
            df_dxf.loc[idx, '宽度_W (mm)'] = reordered_dims[1]
            df_dxf.loc[idx, '高度_T (mm)'] = reordered_dims[2]
        except (ValueError, TypeError):
            # 如果转换失败，保持原始数据不变
            pass
    
    # 预处理：添加核心模式和已匹配标记
    df_dxf['核心模式'] = df_dxf['文件名'].apply(_extract_core_pattern)
    df_prt['核心模式'] = df_prt['文件名'].apply(_extract_core_pattern)
    df_dxf['已匹配'] = False
    df_prt['已匹配'] = False  # 添加PRT已匹配标记

    # 第一层筛选：文件名核心模式匹配 + 三维尺寸匹配
    first_layer_matched, df_dxf, df_prt = _first_layer_matching(df_dxf, df_prt, tolerance)

    # 中间层筛选：对第一层未匹配的文件进行文件名匹配 + 渐进式尺寸匹配（2-10mm）
    middle_layer_matched, df_dxf, df_prt = _middle_layer_matching(df_dxf, df_prt, first_layer_matched)

    # 第二层筛选：对未匹配的数据进行分类和渐进式匹配
    all_matched, unmatched = _second_layer_matching(df_dxf, df_prt, middle_layer_matched)

    return all_matched, unmatched


def _print_stats(matched: list, unmatched: list):
    """
    打印匹配结果统计信息
    
    统计内容：
    - 匹配成功的DXF文件数量
    - 未匹配的PRT文件数量  
    - 未匹配的DXF文件数量
    
    Args:
        matched: 匹配记录列表
        unmatched: 未匹配DXF记录列表
    """
    # 统计匹配成功的DXF文件数量（已匹配状态且DXF文件名不为空）
    matched_dxf_count = len([m for m in matched if m['匹配状态'] == '已匹配' and m['DXF文件名']])
    prt_miss = len([m for m in matched if m['匹配状态'] == '未匹配-PRT'])
    dxf_miss = len(unmatched)

    print(f"\n=== 匹配统计 ===")
    print(f"  匹配成功: {matched_dxf_count}")
    print(f"  未匹配PRT: {prt_miss}")
    print(f"  未匹配DXF: {dxf_miss}")


# ==============================================================================
# 主函数（唯一对外暴露的接口）
# ==============================================================================

def match_data(dxf_csv: str, prt_csv: str, output_csv: str, tolerance: float = 1.0) -> Optional[str]:
    """
    将DXF和PRT的尺寸数据进行配对，采用三层筛选机制：
    第一层：文件名核心模式匹配 + 三维精准匹配（去掉二维精准匹配）
    中间层：对第一层未匹配的文件进行文件名匹配 + 三维渐进式尺寸匹配（2-10mm，最大差值原则）
    第二层：对未匹配的数据按类别分类后进行渐进式容差匹配
    
    中间层筛选详细规则：
    1. 采用与第一层相同的文件名匹配规则（PRT核心模式包含在DXF文件名中）
    2. 对通过文件名匹配的文件进行9级尺寸差检查（2mm到10mm）：
       ① 如果2D各维度+0.1mm≥3D对应维度，且三维尺寸差值的最大值≤2mm，匹配成功，优先级为"中间层2mm"
       ② 如果2D各维度+0.1mm≥3D对应维度，且三维尺寸差值的最大值≤3mm，匹配成功，优先级为"中间层3mm"
       ...
       ⑨ 如果2D各维度+0.1mm≥3D对应维度，且三维尺寸差值的最大值≤10mm，匹配成功，优先级为"中间层10mm"
    
    第二层筛选详细规则：
    1. 分别对未匹配的DXF和PRT文件按类别分类（基于文件名关键词动态提取分类）
    2. 对同一类别的文件进行体积相似度和形状比例相似度计算：
       - 体积相似度 = (1 - |体积差|/PRT体积) × 100%
       - 形状相似度基于长宽比、宽高比、长高比的差异计算
    3. 综合相似度 = 体积相似度×0.6 + 形状相似度×0.4
    4. 渐进式相似度阈值检查（95%、90%、85%、80%、75%）
    5. 过滤分类为OTHER的文件，不参与第二层匹配
    
    Args:
        dxf_csv: DXF信息CSV文件路径（包含2D图纸数据）
        prt_csv: PRT信息CSV文件路径（包含3D模型数据）
        output_csv: 匹配结果输出CSV文件路径
        tolerance: 第一层筛选的尺寸匹配容差值（毫米），默认为1.0mm
    
    Returns:
        Optional[str]: 成功返回输出文件路径，失败返回None
    """
    if not _PANDAS_AVAILABLE:
        print("错误: pandas库未安装")
        return None

    if not os.path.exists(dxf_csv):
        print(f"错误: DXF文件不存在 - {dxf_csv}")
        return None

    if not os.path.exists(prt_csv):
        print(f"错误: PRT文件不存在 - {prt_csv}")
        return None

    try:
        df_dxf = pd.read_csv(dxf_csv)
        df_prt = pd.read_csv(prt_csv)
    except Exception as e:
        print(f"错误: 读取CSV失败 - {e}")
        return None

    df_dxf.columns = df_dxf.columns.str.strip()
    df_prt.columns = df_prt.columns.str.strip()

    # 输出初始文件数量
    initial_dxf_count = len(df_dxf)
    initial_prt_count = len(df_prt)
    print(f"开始匹配：共有 {initial_prt_count} 个3D文件和 {initial_dxf_count} 个2D文件需要参与匹配")

    matched, unmatched = _do_matching(df_dxf, df_prt, tolerance)

    # 计算匹配结果统计
    # 只统计实际匹配成功的DXF文件数量（每个DXF只能匹配一次）
    matched_count = len([m for m in matched if m['匹配状态'] == '已匹配' and m['DXF文件名']])
    unmatched_prt_count = len([m for m in matched if m['匹配状态'] == '未匹配-PRT'])
    unmatched_dxf_count = len(unmatched)
    
    # 输出匹配结果统计
    print(f"匹配完成：")
    print(f"  成功匹配 {matched_count} 个文件")
    print(f"  剩余未匹配 3D文件 {unmatched_prt_count} 个")
    print(f"  剩余未匹配 2D文件 {unmatched_dxf_count} 个")

    # 保存结果
    final_df = pd.DataFrame(matched + unmatched)
    os.makedirs(os.path.dirname(output_csv), exist_ok=True)
    final_df.to_csv(output_csv, index=False, encoding='utf-8-sig')

    _print_stats(matched, unmatched)
    print(f"结果已保存: {output_csv}")

    return output_csv