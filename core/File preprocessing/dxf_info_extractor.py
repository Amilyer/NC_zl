# -*- coding: utf-8 -*-
"""
DXF信息提取模块
主函数: extract_dxf_info(input_folder, output_csv) -> str
"""

import csv
import glob
import os
import re
import warnings
from collections import OrderedDict
from math import sqrt
from typing import Dict, List, Optional, Tuple

warnings.filterwarnings("ignore")

# 检查依赖
try:
    import ezdxf
    _EZDXF_AVAILABLE = True
except ImportError:
    ezdxf = None
    _EZDXF_AVAILABLE = False


# ==============================================================================
# 内部辅助类和函数（不对外暴露）
# ==============================================================================

class _DrawingInfoExtractor:
    """图纸信息提取器"""

    def extract_drawing_info(self, texts: List) -> Dict:
        """从文字中提取尺寸信息"""
        dimensions = None

        for text in texts:
            content = text['content'].strip()

            # 多种尺寸格式匹配
            patterns = [
                (r'(\d+\.?\d*)\s*L\s*\*\s*(\d+\.?\d*)\s*W\s*\*\s*(\d+\.?\d*)\s*T', re.IGNORECASE),
                (r'L\s*(\d+\.?\d*)\s*\*\s*W\s*(\d+\.?\d*)\s*\*\s*T\s*(\d+\.?\d*)', re.IGNORECASE),
                (r'L\s*[=:]\s*(\d+\.?\d*)\s*W\s*[=:]\s*(\d+\.?\d*)\s*T\s*[=:]\s*(\d+\.?\d*)', re.IGNORECASE),
                (r'(\d+\.?\d*)\s*[×xX]\s*(\d+\.?\d*)\s*[×xX]\s*(\d+\.?\d*)', 0),
                (r'(\d+\.?\d*)\s*L\s+(\d+\.?\d*)\s*W\s+(\d+\.?\d*)\s*T', re.IGNORECASE),
                (r'(\d+\.?\d*)\s*L\s*[xX*]\s*(\d+\.?\d*)\s*W\s*[xX*]\s*(\d+\.?\d*)\s*T', re.IGNORECASE),
                (r'[（(]\s*(\d+\.?\d*)\s*L\s*[xX*]\s*(\d+\.?\d*)\s*W\s*[xX*]\s*(\d+\.?\d*)\s*T\s*[,:]?.*?[）)]', re.IGNORECASE),
            ]

            for pattern, flags in patterns:
                match = re.search(pattern, content, flags) if flags else re.search(pattern, content)
                if match:
                    l, w, t = match.groups()
                    dimensions = f"{l}L*{w}W*{t}T"
                    break
            if dimensions:
                break

        return {'dimensions': dimensions}

    def parse_dimensions(self, dimensions_str: str) -> Dict:
        """解析尺寸字符串"""
        if not dimensions_str:
            return {}
        try:
            result = {}
            l_match = re.search(r'(\d+\.?\d*)\s*L', dimensions_str, re.IGNORECASE)
            w_match = re.search(r'(\d+\.?\d*)\s*W', dimensions_str, re.IGNORECASE)
            t_match = re.search(r'(\d+\.?\d*)\s*T', dimensions_str, re.IGNORECASE)
            if l_match:
                result['length'] = float(l_match.group(1))
            if w_match:
                result['width'] = float(w_match.group(1))
            if t_match:
                result['thickness'] = float(t_match.group(1))
            return result
        except:
            return {}


class _TextProcessor:
    """文字处理器"""

    def process_text_list(self, texts: List) -> List:
        if not texts:
            return []
        seen = set()
        result = []
        for text in texts:
            content = text['content'].strip()
            if content and content not in seen:
                seen.add(content)
                result.append(text)
        return result


class _CADAnalyzer:
    """CAD文件分析器"""

    def __init__(self):
        self.all_texts = []
        self.frame_blocks = []
        self.sub_drawings = {}
        self.info_extractor = _DrawingInfoExtractor()
        self.text_processor = _TextProcessor()
        self.processed_files = set()

    def analyze_file(self, file_path: str) -> Dict:
        """分析单个CAD文件"""
        file_basename = os.path.basename(file_path)
        if file_basename in self.processed_files:
            return {}

        self.processed_files.add(file_basename)
        self.all_texts = []
        self.frame_blocks = []
        self.sub_drawings = {}

        try:
            doc = ezdxf.readfile(file_path)
            msp = doc.modelspace()
            self._extract_texts(msp)
            self._identify_regions(msp)
            self._create_regions()
            self._assign_texts()
            return self.sub_drawings
        except Exception as e:
            print(f"  分析失败: {e}")
            return {}

    def _extract_texts(self, msp):
        """提取文字"""
        for entity_type in ['TEXT', 'MTEXT', 'ATTRIB', 'ATTDEF']:
            try:
                for entity in msp.query(entity_type):
                    info = self._process_text_entity(entity)
                    if info:
                        self.all_texts.append(info)
            except:
                continue

    def _process_text_entity(self, entity) -> Optional[Dict]:
        try:
            content = self._get_text_content(entity)
            position = self._get_text_position(entity)
            if content and position:
                return {
                    'content': self._clean_content(content),
                    'position': position,
                    'entity_type': entity.dxftype()
                }
        except:
            pass
        return None

    def _get_text_content(self, entity) -> Optional[str]:
        t = entity.dxftype()
        try:
            if t == 'TEXT':
                return entity.dxf.text
            elif t == 'MTEXT':
                return entity.get_text() if hasattr(entity, 'get_text') else getattr(entity.dxf, 'text', None)
            elif t in ['ATTRIB', 'ATTDEF']:
                return entity.dxf.text
        except:
            pass
        return None

    def _get_text_position(self, entity) -> Optional[Tuple[float, float]]:
        try:
            if hasattr(entity.dxf, 'insert'):
                p = entity.dxf.insert
                return (float(p.x), float(p.y))
            elif hasattr(entity.dxf, 'position'):
                p = entity.dxf.position
                return (float(p.x), float(p.y))
        except:
            pass
        return None

    def _clean_content(self, content: str) -> str:
        if not content:
            return ""
        content = re.sub(r'\{\\[^}]*\}', '', content)
        content = re.sub(r'\\[A-Za-z][^;]*;', '', content)
        
        def unicode_replace(m):
            try:
                return chr(int(m.group(1), 16))
            except:
                return m.group(0)
        
        content = re.sub(r'\\U\+([0-9A-Fa-f]{4})', unicode_replace, content)
        replacements = {'%%c': 'Φ', '%%C': 'Φ', '%%d': '°', '%%D': '°', '%%p': '±', '%%P': '±'}
        for old, new in replacements.items():
            content = content.replace(old, new)
        return re.sub(r'\s+', ' ', content).strip()

    def _identify_regions(self, msp):
        """识别子图区域"""
        entities = []
        for entity in msp.query('LINE LWPOLYLINE CIRCLE ARC POLYLINE'):
            bounds = self._get_entity_bounds(entity)
            if bounds:
                entities.append({
                    'bounds': bounds,
                    'center': ((bounds['min_x'] + bounds['max_x']) / 2,
                               (bounds['min_y'] + bounds['max_y']) / 2)
                })

        if not entities:
            return

        # 简单聚类
        clusters = []
        visited = set()
        for i, ent in enumerate(entities):
            if i in visited:
                continue
            cluster = [ent]
            visited.add(i)
            cx, cy = ent['center']
            for j, other in enumerate(entities):
                if j not in visited:
                    ox, oy = other['center']
                    if sqrt((cx - ox) ** 2 + (cy - oy) ** 2) < 300:
                        cluster.append(other)
                        visited.add(j)
            clusters.append(cluster)

        for i, cluster in enumerate(clusters):
            min_x = min(e['bounds']['min_x'] for e in cluster)
            max_x = max(e['bounds']['max_x'] for e in cluster)
            min_y = min(e['bounds']['min_y'] for e in cluster)
            max_y = max(e['bounds']['max_y'] for e in cluster)
            margin = 50
            self.frame_blocks.append({
                'block_name': f'region_{i + 1}',
                'bounds': {
                    'min_x': min_x - margin, 'max_x': max_x + margin,
                    'min_y': min_y - margin, 'max_y': max_y + margin,
                    'width': max_x - min_x + 2 * margin,
                    'height': max_y - min_y + 2 * margin
                }
            })

    def _get_entity_bounds(self, entity) -> Optional[Dict]:
        try:
            t = entity.dxftype()
            if t == 'LINE':
                s, e = entity.dxf.start, entity.dxf.end
                return {'min_x': min(s.x, e.x), 'max_x': max(s.x, e.x),
                        'min_y': min(s.y, e.y), 'max_y': max(s.y, e.y)}
            elif t in ['CIRCLE', 'ARC']:
                c, r = entity.dxf.center, entity.dxf.radius
                return {'min_x': c.x - r, 'max_x': c.x + r,
                        'min_y': c.y - r, 'max_y': c.y + r}
            elif t in ['LWPOLYLINE', 'POLYLINE']:
                pts = entity.get_points(format='xy')
                if pts:
                    xs, ys = zip(*pts)
                    return {'min_x': min(xs), 'max_x': max(xs),
                            'min_y': min(ys), 'max_y': max(ys)}
        except:
            pass
        return None

    def _create_regions(self):
        """创建子图区域"""
        self.frame_blocks.sort(key=lambda fb: (-round(fb['bounds']['min_y'] / 100), round(fb['bounds']['min_x'] / 100)))
        for i, fb in enumerate(self.frame_blocks):
            self.sub_drawings[f"region_{i + 1:03d}"] = {
                'bounds': fb['bounds'],
                'texts': []
            }

    def _assign_texts(self):
        """分配文字到区域"""
        for text in self.all_texts:
            x, y = text['position']
            for region_data in self.sub_drawings.values():
                b = region_data['bounds']
                if b['min_x'] <= x <= b['max_x'] and b['min_y'] <= y <= b['max_y']:
                    region_data['texts'].append(text)
                    break

        for region_data in self.sub_drawings.values():
            region_data['texts'] = self.text_processor.process_text_list(region_data['texts'])


def _write_csv(data: List[Dict], output_path: str):
    """写入CSV"""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        fieldnames = ['文件名', '长度_L (mm)', '宽度_W (mm)', '高度_T (mm)', '备注']
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in data:
            writer.writerow({
                '文件名': row.get('文件名', ''),
                '长度_L (mm)': row.get('长度_L (mm)', ''),
                '宽度_W (mm)': row.get('宽度_W (mm)', ''),
                '高度_T (mm)': row.get('高度_T (mm)', ''),
                '备注': row.get('备注', '')
            })


# ==============================================================================
# 主函数（唯一对外暴露的接口）
# ==============================================================================

def extract_dxf_info(input_folder: str, output_csv: str) -> Optional[str]:
    """
    从DXF文件夹中提取尺寸信息，生成CSV报告
    
    Args:
        input_folder: 包含DXF文件的文件夹路径
        output_csv: 输出CSV文件路径
    
    Returns:
        str: 成功返回输出文件路径，失败返回None
    """
    if not _EZDXF_AVAILABLE:
        print("错误: ezdxf库未安装")
        return None

    if not os.path.isdir(input_folder):
        print(f"错误: 文件夹不存在 - {input_folder}")
        return None

    dxf_files = glob.glob(os.path.join(input_folder, "*.dxf"))
    if not dxf_files:
        print(f"错误: 未找到DXF文件 - {input_folder}")
        return None

    print(f"找到 {len(dxf_files)} 个DXF文件")

    analyzer = _CADAnalyzer()
    all_data = []

    for dxf_file in dxf_files:
        file_name = os.path.basename(dxf_file)
        # print(f"  处理: {file_name}")

        sub_drawings = analyzer.analyze_file(dxf_file)

        if sub_drawings:
            for region_data in sub_drawings.values():
                info = analyzer.info_extractor.extract_drawing_info(region_data['texts'])
                dims = analyzer.info_extractor.parse_dimensions(info['dimensions']) if info['dimensions'] else {}

                all_data.append({
                    '文件名': file_name,
                    '长度_L (mm)': dims.get('length', ''),
                    '宽度_W (mm)': dims.get('width', ''),
                    '高度_T (mm)': dims.get('thickness', ''),
                    '备注': '' if dims else '未提取到尺寸'
                })
        else:
            all_data.append({
                '文件名': file_name,
                '长度_L (mm)': '',
                '宽度_W (mm)': '',
                '高度_T (mm)': '',
                '备注': '未识别到区域'
            })

    # 去重：每个文件保留第一条有效记录
    merged = OrderedDict()
    for rec in all_data:
        fn = rec['文件名']
        if fn in merged and (merged[fn].get('长度_L (mm)') or merged[fn].get('宽度_W (mm)')):
            continue
        merged[fn] = rec

    _write_csv(list(merged.values()), output_csv)
    
    valid_count = len([r for r in merged.values() if r.get('长度_L (mm)')])
    print(f"完成: {len(merged)} 个文件, {valid_count} 个有尺寸信息")
    
    return output_csv