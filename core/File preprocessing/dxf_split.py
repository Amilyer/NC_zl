import os
import re
import warnings
from collections import Counter, defaultdict
from math import cos, pi, radians, sin, sqrt
from typing import Dict, List, Optional, Tuple

warnings.filterwarnings("ignore")

try:
    import ezdxf
    from ezdxf.addons import Importer
    EZDXF_AVAILABLE = True
except ImportError:
    print("警告: ezdxf 库未找到。CAD 2D处理功能将无法运行。")
    ezdxf = None
    Importer = None
    EZDXF_AVAILABLE = False

if EZDXF_AVAILABLE:
    
    # 子图编号与文件名提取工具类
    class ProfessionalDrawingNumberExtractor:
        """专业图纸编号提取器 + 子图文件名提取"""

        def __init__(self):
            # 用于命名子图文件的正则（按优先级）
            self.filename_patterns = [
                re.compile(r'编号\s*[：:]\s*(\S+)'),
                re.compile(r'加工说明:\([^)]+\)_(\S+)'),
                re.compile(r'加工说明\([^)]+\):(\S+)'),
            ]

            self.explicit_label_re = re.compile(r'编号\s*[：:]\s*(\S+)')
            self.primary_patterns = [
                r'PH-[A-Z0-9]+', r'DIE-[A-Z0-9]+', r'[A-Z]{1,2}[0-9]{1,3}-[A-Z]{1,2}',
                r'[A-Z]{1,2}[0-9]{2,3}', r'[A-Z]{2,4}-[0-9]{1,3}',
            ]
            self.secondary_patterns = [
                r'[A-Z]{2,4}[0-9]{1,2}', r'[A-Z]{2,4}', r'MA-?[A-Z0-9]*', r'[A-Z][0-9]',
            ]
            self.excluded_terms = {
                '图纸', '设计', '审核', '标准', '规格', '材料', '备注', '品名', '编号',
                '数量', '热处理', '加工说明', '修改', '尺寸', '所有', '全周', '已订购',
                'TITLE', 'DRAWING', 'DESIGN', 'SCALE', 'DATE', '制图', '日期',
                '单位', '比例', '共页', '第页', '版本', 'PCS', '深', '攻', '钻',
                '割', '铰', '倒角', '沉头', '背', '穿', '让位', '合销', '导套',
                '螺丝', '基准', '弹簧', '定位', '精铣', '慢丝', '线割', '垂直度',
                '位置度', '加工', '夹板', '入子', '连接块', '外形', '绿色', '虚线',
                '直身', '拼装', '零件', '模板', '精磨'
            }
            self.cad_annotations = {
                'M', 'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10',
                'G', 'G1', 'G2', 'G3', 'G4', 'G5', 'G6', 'G7', 'G8', 'G9',
                'L', 'L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'L9',
                'U', 'U1', 'U2', 'U3', 'U4', 'U5', 'X', 'X1', 'X2', 'X3', 'X4', 'X5', 'X6', 'X7', 'X8', 'X9',
                'K', 'K1', 'K2', 'K3', 'K4', 'K5', 'A', 'A1', 'A2', 'A3', 'A4', 'A5',
                'Q', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'f1', 'f2', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9'
            }

        def extract_region_filename_by_patterns(self, subdrawing_data: Dict) -> Optional[str]:
            # 从区域文字中提取子图文件名
            texts = subdrawing_data.get('texts', []) or []
            for t in texts:
                c = (t.get('content') or '').strip()
                if not c:
                    continue
                for rx in self.filename_patterns:
                    m = rx.search(c)
                    if m and m.group(1):
                        return self.generate_safe_filename(self._clean_candidate_after_label(m.group(1)))
            for i, t in enumerate(texts):
                c = (t.get('content') or '').strip()
                if c in ('编号', '编号：', '编号:'):
                    for j in range(i + 1, min(i + 6, len(texts))):
                        nxt = (texts[j].get('content') or '').strip()
                        if nxt:
                            return self.generate_safe_filename(self._clean_candidate_after_label(nxt))
            return None

        def extract_drawing_number_from_region(self, subdrawing_data: Dict) -> Optional[str]:
            # 从区域文字中提取图纸编号
            bounds = subdrawing_data['bounds']
            texts = subdrawing_data['texts']
            filtered_texts = self._preprocess_texts(texts)
            if not filtered_texts:
                return None
            extraction_methods = [
                self._extract_from_explicit_labels,
                self._extract_from_key_positions,
                self._extract_from_pattern_matching
            ]
            for method in extraction_methods:
                result = method(bounds, filtered_texts)
                if result and self._validate_drawing_number(result):
                    return result
            return None

        def _preprocess_texts(self, texts: List) -> List:
            # 预处理区域文字，过滤无效内容
            content_frequency = Counter([text['content'].strip() for text in texts])
            processed = []
            for text in texts:
                content = text['content'].strip()
                if not content or len(content) > 30:
                    continue
                if any(term in content for term in self.excluded_terms):
                    continue
                if content in self.cad_annotations:
                    continue
                if len(content) <= 2 and content_frequency[content] > 5:
                    continue
                if self._is_dimension_or_value(content):
                    continue
                processed.append(text)
            return processed

        def _is_dimension_or_value(self, content: str) -> bool:
            # 判断文字是否为尺寸或数值
            dimension_patterns = [
                r'^\d+\.?\d*$', r'^\d+\.?\d*[LWTHDRC]$', r'^Φ\d+\.?\d*$', r'^R\d+\.?\d*$',
                r'^\d+\.?\d*°$', r'^\d+\.?\d*mm$', r'^M\d+x\d+\.?\d*$', r'^\d+\.?\d*深$',
                r'^C\d+\.?\d*$', r'^HRC\d+-\d+$', r'^\d+\.?\d*[×xX]\d+\.?\d*',
            ]
            return any(re.match(pattern, content) for pattern in dimension_patterns)

        def _extract_from_explicit_labels(self, bounds: Dict, texts: List) -> Optional[str]:
            # 从显式标签提取编号
            for t in texts:
                m = self.explicit_label_re.search(t['content'].strip())
                if m:
                    cand = self._clean_candidate_after_label(m.group(1))
                    if self._validate_drawing_number(cand):
                        return cand
            for i, t in enumerate(texts):
                c = t['content'].strip()
                if c in ('编号', '编号：', '编号:'):
                    for j in range(i + 1, min(i + 6, len(texts))):
                        cand = self._clean_candidate_after_label(texts[j]['content'].strip())
                        if self._validate_drawing_number(cand):
                            return cand
            return None

        def _clean_candidate_after_label(self, s: str) -> str:
            # 清理标签后的编号候选
            cleaned = (s or '').strip()
            if not cleaned:
                return cleaned
            cleaned = cleaned.split()[0]
            cleaned = cleaned.strip('，,。.;；:：)]】）\'"').strip('([【（\'"')
            cleaned = re.sub(r'^[\s\-_]+|[\s\-_]+$', '', cleaned)
            return cleaned[:64] if len(cleaned) > 64 else cleaned

        def _extract_from_key_positions(self, bounds: Dict, texts: List) -> Optional[str]:
            # 从关键位置提取编号
            position_zones = [
                {'name': 'top_left', 'bounds': self._define_zone_bounds(bounds, 0, 0.35, 0.7, 1.0), 'weight': 2.5},
                {'name': 'title_block', 'bounds': self._define_zone_bounds(bounds, 0.7, 1.0, 0, 0.25), 'weight': 2.0},
                {'name': 'top_right','bounds': self._define_zone_bounds(bounds,  0.65, 1.0, 0.75,1.0),'weight': 1.5 },
                {'name': 'bottom_left', 'bounds': self._define_zone_bounds(bounds, 0, 0.35, 0, 0.25), 'weight': 1.8},
            ]
            best_candidate, best_score = None, 0.0
            for zone in position_zones:
                zone_texts = self._get_texts_in_bounds(texts, zone['bounds'])
                for text in zone_texts:
                    content = text['content'].strip()
                    quality = self._calculate_quality_score(content)
                    if quality > 0:
                        score = quality * zone['weight']
                        if score > best_score:
                            best_score, best_candidate = score, content
            return best_candidate if best_score > 2.0 else None

        def _extract_from_pattern_matching(self, bounds: Dict, texts: List) -> Optional[str]:
            # 从模式匹配提取编号
            candidates = []
            pattern_groups = [
                (self.primary_patterns, 3.0),
                (self.secondary_patterns, 2.0),
            ]
            for text in texts:
                content = text['content'].strip()
                for patterns, base_w in pattern_groups:
                    for pat in patterns:
                        if re.match(pat + '$', content):
                            pos_w = self._calculate_position_weight(text['position'], bounds)
                            candidates.append((content, base_w * pos_w))
                            break
            if candidates:
                return max(candidates, key=lambda x: x[1])[0]
            return None

        def _validate_drawing_number(self, content: str) -> bool:
            # 校验提取到的编号是否合法
            if not content or len(content) > 16:
                return False
            invalid_patterns = [
                r'^[:：].*', r'.*[:：]\s*$', r'^\d+\.\d+$', r'^[0-9]{4,}$',
                r'.*说明.*', r'.*加工.*', r'.*深$', r'.*磨$', r'^[\d\.\-\+\s]+$',
                r'.*PCS.*',
            ]
            if any(re.match(p, content) for p in invalid_patterns):
                return False
            valid_patterns = [
                r'^[A-Z]{1,4}[0-9]*$',
                r'^[A-Z]+-[A-Z0-9]+$',
                r'^[A-Z]{2,4}$',
            ]
            return any(re.match(p, content) for p in valid_patterns)

        def _calculate_quality_score(self, content: str) -> float:
            # 计算编号内容的质量分数
            if not content:
                return 0.0
            score = 0.0
            n = len(content)
            if 2 <= n <= 6:
                score += 3.0
            elif n == 1:
                score += 0.5
            else:
                score += 1.0
            if re.match(r'^[A-Z]+-[A-Z0-9]+$', content):
                score += 4.0
            elif re.match(r'^[A-Z]{1,3}[0-9]{1,3}$', content):
                score += 3.5
            elif re.match(r'^[A-Z]{2,4}$', content):
                score += 2.5
            elif re.match(r'^[A-Z][0-9]$', content):
                score += 2.0
            if re.search(r'[0-9]', content):
                score += 1.0
            if content in self.cad_annotations:
                score = 0.0
            return score

        def _calculate_position_weight(self, position: Tuple, bounds: Dict) -> float:
            # 计算编号在区域中的位置权重
            x, y = position
            width, height = bounds['width'], bounds['height']
            xn = (x - bounds['min_x']) / width
            yn = (y - bounds['min_y']) / height
            x_w = 1.2 - xn * 0.4
            y_w = 1.0 if yn > 0.75 else (0.9 if yn < 0.25 else 0.4)
            return x_w * y_w

        def _define_zone_bounds(self, bounds: Dict, x_start: float, x_end: float,
                                y_start: float, y_end: float) -> Dict:
            # 定义区域边界
            w, h = bounds['width'], bounds['height']
            return {
                'min_x': bounds['min_x'] + w * x_start,
                'max_x': bounds['min_x'] + w * x_end,
                'min_y': bounds['min_y'] + h * y_start,
                'max_y': bounds['min_y'] + h * y_end,
            }

        def _get_texts_in_bounds(self, texts: List, zone_bounds: Dict) -> List:
            # 获取区域内的文字
            res = []
            for t in texts:
                x, y = t['position']
                if (zone_bounds['min_x'] <= x <= zone_bounds['max_x'] and
                        zone_bounds['min_y'] <= y <= zone_bounds['max_y']):
                    res.append(t)
            return res

        def generate_safe_filename(self, name: str) -> str:
            # 生成安全的文件名
            if not name:
                return "未知编号"
            s = re.sub(r'[<>:"/\\|?*]', '_', name.strip()).replace(' ', '_')
            s = s.rstrip(' .')
            return s if len(s) <= 80 else s[:80]

    # 文字实体过滤与处理工具
    class IntelligentTextProcessor:
        """智能文字处理器"""

        def __init__(self):
            self.noise_patterns = [
                r'^\d+\.?\d*$', r'^[\d\.\-\+\s]+$', r'^\d+\.?\d*[LWTHDRC]$',
                r'^Φ\d+\.?\d*', r'^R\d+\.?\d*', r'^M\d+x',
                r'^\d+\.?\d*°$', r'^\d+\.?\d*mm$', r'^\d+\.?\d*[×xX]\d+\.?\d*',
                r'.*深$', r'.*攻$', r'.*钻$',
            ]
            self.meaningful_keywords = [
                '品名', '编号', '材料', '热处理', '数量',
                '加工说明', '尺寸', '修改', '备注', '规格', '型号'
            ]

        def process_text_list(self, texts: List[Dict]) -> List[Dict]:
            # 过滤并处理文字实体，去除无用信息
            if not texts:
                return []
            counter = Counter([t['content'].strip() for t in texts])
            processed = []
            for t in texts:
                c = t['content'].strip()
                if self._should_keep_text(c, counter):
                    processed.append(t)
            return processed

        def _should_keep_text(self, content: str, counter: Counter) -> bool:
            # 判断文字是否应保留
            if not content:
                return False
            if len(content) > 50:
                return False
            if any(k in content for k in self.meaningful_keywords):
                return True
            if any(re.match(p, content) for p in self.noise_patterns):
                return False
            if len(content) <= 3 and counter[content] > 8:
                return False
            if len(content) <= 1 and counter[content] > 3:
                return False
            return True

    # CAD块分析与拆分工具
    class OptimizedCADBlockAnalyzer:
        """优化的CAD块分析器 + 修复BlockLayout错误 + 添加拆分条件筛选"""

        def __init__(self):
            self.all_texts = []
            self.all_entities = []
            self.frame_blocks = []
            self.sub_drawings = {}
            self.layer_colors = {}
            self.text_processor = IntelligentTextProcessor()
            self.number_extractor = ProfessionalDrawingNumberExtractor()
            self.cutting_detector = StrictCuttingDetector()
            self.doc = None
            self.msp = None
            self.classify_map = None

            # 拆分条件配置
            self.required_keywords = ['加工说明']  # 必须包含的关键词
            self.excluded_keywords = ['厂内标准件', '订购', '装配', '组配']  # 排除的关键词

        def analyze_cad_file(self, file_path: str) -> Dict:
            # CAD文件详细分析，提取图层、文字、实体、图框块，识别子图区域
            print(f"开始放宽分析CAD文件: {file_path}")
            try:
                doc = ezdxf.readfile(file_path)
                msp = doc.modelspace()
                self.doc = doc
                self.msp = msp

                self._extract_layer_colors(doc)
                self._extract_all_texts(msp)
                self._extract_all_entities(msp)
                self._identify_frame_blocks(msp)
                self._create_subdrawing_regions()
                self._assign_texts_to_regions()
                self._analyze_cutting_contours_for_regions()

                # if self.sub_drawings:
                #     self._analyze_cutting_contours_for_regions()

                print(f"放宽分析完成，识别出 {len(self.sub_drawings)} 个满足条件的子图区域")
                return self.sub_drawings
            except Exception as e:
                print(f"文件分析失败: {str(e)}")
                import traceback
                traceback.print_exc()
                return {}

        def _should_process_region(self, region_data: Dict) -> bool:
            # 判断区域是否满足拆分条件（包含/排除关键词）
            """检查区域是否满足拆分条件"""
            texts = region_data.get('texts', [])
            text_contents = [t.get('content', '') for t in texts]

            has_required = any(any(keyword in content for keyword in self.required_keywords)
                               for content in text_contents)
            if not has_required:
                # print(f"   [条件检查] 不包含'{self.required_keywords[0]}'，跳过")
                return False

            has_excluded = any(any(keyword in content for keyword in self.excluded_keywords)
                               for content in text_contents)
            if has_excluded:
                # print("   [条件检查] 包含排除词汇，跳过")
                return False

            return True

        def _safe_spline_points(self, entity):
            # 安全提取样条点
            pts = []
            try:
                if hasattr(entity, 'control_points') and entity.control_points:
                    for p in entity.control_points:
                        try:
                            pts.append((float(p[0]), float(p[1]), float(p[2]) if len(p) > 2 else 0.0))
                        except Exception:
                            pts.append((float(p.x), float(p.y), float(getattr(p, 'z', 0.0))))
                if not pts and hasattr(entity, 'fit_points') and entity.fit_points:
                    for p in entity.fit_points:
                        pts.append((float(p.x), float(p.y), float(getattr(p, 'z', 0.0))))
                if not pts and hasattr(entity, 'vertices'):
                    for v in entity.vertices:
                        if hasattr(v, 'dxf') and hasattr(v.dxf, 'location'):
                            pts.append((float(v.dxf.location.x), float(v.dxf.location.y),
                                        float(getattr(v.dxf.location, 'z', 0.0))))
                if not pts and hasattr(entity.dxf, 'start_point') and hasattr(entity.dxf, 'end_point'):
                    s = entity.dxf.start_point
                    e = entity.dxf.end_point
                    pts = [(float(s.x), float(s.y), float(getattr(s, 'z', 0.0))),
                           (float(e.x), float(e.y), float(getattr(e, 'z', 0.0)))]
            except Exception:
                pass
            return pts

        def _point_in_bounds(self, pt, bounds: Dict) -> bool:
            # 判断点是否在区域内
            if pt is None:
                return False
            x, y = pt
            return (bounds['min_x'] <= x <= bounds['max_x']) and (bounds['min_y'] <= y <= bounds['max_y'])

        def _ellipse_start_end_points(self, entity) -> Tuple[
            Optional[Tuple[float, float]], Optional[Tuple[float, float]]]:
            # 获取椭圆起止点
            try:
                tool = entity.construction_tool()
                sp = (float(tool.start_point.x), float(tool.start_point.y))
                ep = (float(tool.end_point.x), float(tool.end_point.y))
                return sp, ep
            except Exception:
                pass

            try:
                c = entity.dxf.center
                maj = entity.dxf.major_axis
                ratio = float(getattr(entity.dxf, 'ratio', 0.5) or 0.5)
                a_x, a_y = float(maj.x), float(maj.y)
                b_x, b_y = -a_y * ratio, a_x * ratio

                t0 = float(getattr(entity.dxf, 'start_param', 0.0) or 0.0)
                t1 = float(getattr(entity.dxf, 'end_param', 2 * pi) or 2 * pi)

                cx, cy = float(c.x), float(c.y)

                def eval_point(t):
                    return (cx + a_x * cos(t) + b_x * sin(t),
                            cy + a_y * cos(t) + b_y * sin(t))

                sp = eval_point(t0)
                ep = eval_point(t1)
                return sp, ep
            except Exception:
                return None, None

        def _ellipse_hits_region_by_endpoints(self, entity, bounds: Dict) -> bool:
            # 判断椭圆端点是否在区域内
            sp, ep = self._ellipse_start_end_points(entity)
            return self._point_in_bounds(sp, bounds) or self._point_in_bounds(ep, bounds)

        def export_regions_to_dxf(self, output_dir: str):
            # 导出所有满足条件的子图为独立DXF文件
            if not self.doc or not self.msp:
                print("导出失败：未加载文档。")
                return
            os.makedirs(output_dir, exist_ok=True)
            all_msp_entities = list(self.msp)

            doc_units = self.doc.units
            print(f"源文件单位类型: {doc_units} (0=无单位, 1=英寸, 3=毫米, 4=厘米)")
            unit_to_mm = {
                0: 1.0, 1: 25.4, 2: 304.8, 3: 1.0, 4: 10.0, 5: 1000.0,
            }
            scale = unit_to_mm.get(doc_units, 1.0)
            print(f"单位转换因子（转为毫米）: {scale}")

            def get_entity_bounds_generic(e) -> Optional[Dict]:
                try:
                    entity_type = e.dxftype()
                    bounds = None

                    if entity_type == 'LINE':
                        start = e.dxf.start
                        end = e.dxf.end
                        min_x = min(start.x, end.x) * scale
                        max_x = max(start.x, end.x) * scale
                        min_y = min(start.y, end.y) * scale
                        max_y = max(start.y, end.y) * scale
                        bounds = {'min_x': min_x, 'max_x': max_x, 'min_y': min_y, 'max_y': max_y,
                                  'start': (start.x * scale, start.y * scale),
                                  'end': (end.x * scale, end.y * scale)}

                    elif entity_type == 'CIRCLE':
                        center = e.dxf.center
                        radius = e.dxf.radius
                        min_x = (center.x - radius) * scale
                        max_x = (center.x + radius) * scale
                        min_y = (center.y - radius) * scale
                        max_y = (center.y + radius) * scale
                        bounds = {'min_x': min_x, 'max_x': max_x, 'min_y': min_y, 'max_y': max_y}
                    elif entity_type == 'ARC':
                        center = e.dxf.center
                        radius = e.dxf.radius
                        start_angle = e.dxf.start_angle
                        end_angle = e.dxf.end_angle
                        from math import cos, radians, sin
                        angles = [start_angle, end_angle]
                        # 补充极值点（0, 90, 180, 270度在圆弧范围内时也要算进去）
                        for a in [0, 90, 180, 270]:
                            if start_angle < end_angle:
                                if start_angle <= a <= end_angle:
                                    angles.append(a)
                            else:
                                if a >= start_angle or a <= end_angle:
                                    angles.append(a)
                        pts = []
                        for ang in angles:
                            rad = radians(ang)
                            x = center.x + radius * cos(rad)
                            y = center.y + radius * sin(rad)
                            pts.append((x * scale, y * scale))
                        xs, ys = zip(*pts)
                        min_x = min(xs)
                        max_x = max(xs)
                        min_y = min(ys)
                        max_y = max(ys)
                        bounds = {'min_x': min_x, 'max_x': max_x, 'min_y': min_y, 'max_y': max_y}

                    elif entity_type in ('LWPOLYLINE', 'POLYLINE'):
                        pts = []
                        if entity_type == 'LWPOLYLINE':
                            pts = [(p[0] * scale, p[1] * scale) for p in e.get_points(format='xy')]
                        else:
                            for v in e.vertices:
                                pts.append((v.dxf.x * scale, v.dxf.y * scale))
                        if pts:
                            xs, ys = zip(*pts)
                            bounds = {'min_x': min(xs), 'max_x': max(xs), 'min_y': min(ys), 'max_y': max(ys)}

                    elif entity_type == 'ELLIPSE':
                        center = e.dxf.center
                        major_axis = e.dxf.major_axis
                        ratio = float(getattr(e.dxf, 'ratio', 0.5) or 0.5)
                        min_x = (center.x - major_axis.x) * scale
                        max_x = (center.x + major_axis.x) * scale
                        min_y = (center.y - (major_axis.y * ratio)) * scale
                        max_y = (center.y + (major_axis.y * ratio)) * scale
                        bounds = {'min_x': min_x, 'max_x': max_x, 'min_y': min_y, 'max_y': max_y}

                    elif entity_type in ('TEXT', 'MTEXT', 'ATTRIB', 'ATTDEF', 'LEADER', 'TABLE'):
                        pos = getattr(e.dxf, 'insert', None) or getattr(e.dxf, 'pos', None)
                        if pos:
                            height = float(getattr(e.dxf, 'height', 2.5) or 2.5) * scale
                            width = height * 5
                            min_x = pos.x * scale - width / 2
                            max_x = pos.x * scale + width / 2
                            min_y = pos.y * scale - height / 2
                            max_y = pos.y * scale + height / 2
                            bounds = {'min_x': min_x, 'max_x': max_x, 'min_y': min_y, 'max_y': max_y}
                    elif entity_type == 'DIMENSION':
                        # 获取标注原点与文本位置
                        defpoint = e.dxf.defpoint
                        textpoint = e.dxf.text_midpoint or e.dxf.text_location
                        
                        # 获取文本内容和高度
                        text_content = e.dxf.text
                        text_height = float(getattr(e.dxf, 'text_height', 2.5)) * scale
                        
                        # 计算边界框
                        # 假设文本框的宽度为高度的2倍（可根据实际字体调整）
                        text_width = text_height * 2
                        
                        # 计算边界框的四个角点
                        min_x = (textpoint.x - text_width / 2) * scale
                        max_x = (textpoint.x + text_width / 2) * scale
                        min_y = (textpoint.y - text_height / 2) * scale
                        max_y = (textpoint.y + text_height / 2) * scale
                        
                        # 同时考虑标注线的端点，确保整个标注线也被包含
                        min_x = min(min_x, defpoint.x * scale)
                        max_x = max(max_x, defpoint.x * scale)
                        min_y = min(min_y, defpoint.y * scale)
                        max_y = max(max_y, defpoint.y * scale)
                        
                        bounds = {'min_x': min_x, 'max_x': max_x, 'min_y': min_y, 'max_y': max_y}

                    elif entity_type == 'INSERT':
                        block_def = e.doc.blocks.get(e.dxf.name)
                        if block_def:
                            bounds = self._calculate_block_bounds(block_def, e)
                            if bounds:
                                bounds['min_x'] *= scale
                                bounds['max_x'] *= scale
                                bounds['min_y'] *= scale
                                bounds['max_y'] *= scale

                    elif entity_type == 'SPLINE':
                        pts = self._safe_spline_points(e)
                        if pts:
                            pts_mm = [(p[0] * scale, p[1] * scale) for p in pts]
                            xs, ys = zip(*pts_mm)
                            buffer = 1.0
                            bounds = {
                                'min_x': min(xs) - buffer,
                                'max_x': max(xs) + buffer,
                                'min_y': min(ys) - buffer,
                                'max_y': max(ys) + buffer
                            }

                    return bounds

                except Exception as ex:
                    print(f"计算实体 {e.dxftype()} 边界时出错: {str(ex)}")
                    return None

            def intersect(a: Dict, b: Dict, buffer: float = 0.1) -> bool:
                return not (
                        a['max_x'] + buffer < b['min_x'] or
                        a['min_x'] - buffer > b['max_x'] or
                        a['max_y'] + buffer < b['min_y'] or
                        a['min_y'] - buffer > b['max_y']
                )

            export_count = 0
            for idx, (region_id, region) in enumerate(self.sub_drawings.items(), start=1):
                bounds = region['bounds']
                region_bounds = {
                    'min_x': bounds['min_x'] * scale,
                    'max_x': bounds['max_x'] * scale,
                    'min_y': bounds['min_y'] * scale,
                    'max_y': bounds['max_y'] * scale
                }
                # print(f"\n处理子图 {region_id}，图框边界（毫米）: {region_bounds}")

                frame_block_name = region['frame_block']['block_name']

                selected_entities = []
                for e in all_msp_entities:
                    try:
                        if e.dxftype() == 'INSERT' and e.dxf.name == frame_block_name:
                            continue

                        ent_bounds = get_entity_bounds_generic(e)
                        if not ent_bounds:
                            continue
                        # 直线特殊处理：只有两端都在区域内才选入
                        if e.dxftype() == 'LINE' and ent_bounds:
                            sx, sy = ent_bounds['start']
                            ex, ey = ent_bounds['end']
                            if (region_bounds['min_x'] <= sx <= region_bounds['max_x'] and
                                region_bounds['min_y'] <= sy <= region_bounds['max_y'] and
                                region_bounds['min_x'] <= ex <= region_bounds['max_x'] and
                                region_bounds['min_y'] <= ey <= region_bounds['max_y']):
                                selected_entities.append(e)
                        elif ent_bounds and intersect(ent_bounds, region_bounds):
                            selected_entities.append(e)
                            
                    except Exception as ex:
                        print(f"筛选实体 {e.dxftype()} 时出错: {str(ex)}")
                        continue

                # print(f"子图 {region_id} 初始筛选出 {len(selected_entities)} 个实体")

                if not selected_entities:
                    # print(f"[导出提示] {region_id} 内未找到可导出的实体，跳过。")
                    continue
                
                # 复制文字样式和标注样式
                new_doc = ezdxf.new(dxfversion=self.doc.dxfversion)
                try:
                    new_doc.units = self.doc.units
                    
                    # 1. 先复制文字样式（标注依赖文字样式）
                    for text_style in self.doc.styles:
                        style_name = text_style.dxf.name
                        if style_name not in new_doc.styles:
                            new_text_style = new_doc.styles.new(style_name)
                        else:
                            new_text_style = new_doc.styles.get(style_name)
                        
                        # 复制所有文字样式属性
                        attrs_to_copy = ['font', 'bigfont', 'height', 'width', 'oblique', 'flags', 'generation_flags']
                        for attr in attrs_to_copy:
                            try:
                                if hasattr(text_style.dxf, attr):
                                    setattr(new_text_style.dxf, attr, getattr(text_style.dxf, attr))
                            except Exception:
                                pass
                    
                    # 2. 再复制标注样式（完整复制所有属性）
                    for dim_style in self.doc.dimstyles:
                        style_name = dim_style.dxf.name
                        if style_name not in new_doc.dimstyles:
                            new_doc.dimstyles.new(style_name)
                        new_dim = new_doc.dimstyles.get(style_name)
                        
                        # 复制所有标注样式属性（关键：包含小数位数和零抑制）
                        dim_attrs_to_copy = [
                            # 文字相关
                            'dimtxsty', 'dimtxt', 'dimtad', 'dimgap', 'dimjust', 'dimtih', 'dimtoh',
                            # 数值格式（关键属性）
                            'dimdec', 'dimzin', 'dimlunit', 'dimdsep', 'dimrnd', 'dimtfac',
                            # 线条和箭头
                            'dimscale', 'dimasz', 'dimblk', 'dimblk1', 'dimblk2', 'dimdle', 'dimdli',
                            'dimexe', 'dimexo', 'dimclrd', 'dimclre', 'dimclrt',
                            # 单位和测量
                            'dimlfac', 'dimpost', 'dimapost', 'dimalt', 'dimaltd', 'dimaltf',
                            # 公差
                            'dimtol', 'dimlim', 'dimtp', 'dimtm', 'dimtolj',
                            # 其他
                            'dimse1', 'dimse2', 'dimtad', 'dimfrac', 'dimlwd', 'dimlwe'
                        ]
                        
                        for attr in dim_attrs_to_copy:
                            try:
                                if hasattr(dim_style.dxf, attr):
                                    setattr(new_dim.dxf, attr, getattr(dim_style.dxf, attr))
                            except Exception:
                                pass
                        
                        # 特别处理测量值（保留原始设置）
                        try:
                            if hasattr(dim_style, 'dxfattribs'):
                                for key, value in dim_style.dxfattribs().items():
                                    if key not in ['handle', 'owner']:
                                        try:
                                            setattr(new_dim.dxf, key, value)
                                        except Exception:
                                            pass
                        except Exception:
                            pass
                            
                            # 显式保证小数点分隔符和小数位数与原图一致
                            try:
                                if hasattr(dim_style.dxf, 'dimdsep'):
                                    new_dim.dxf.dimdsep = dim_style.dxf.dimdsep
                                else:
                                    new_dim.dxf.dimdsep = '.'
                            except Exception:
                                new_dim.dxf.dimdsep = '.'
                            try:
                                if hasattr(dim_style.dxf, 'dimdec'):
                                    new_dim.dxf.dimdec = dim_style.dxf.dimdec
                            except Exception:
                                pass
                except Exception as e:
                    print(f"复制文字样式和标注样式时出错: {str(e)}")

                # 3. 复制所有块定义（包括箭头块，标注样式需要）（有大用，不能省）
                try:
                    for src_block in self.doc.blocks:
                        block_name = src_block.dxf.name
                        # 跳过模型空间和图纸空间
                        if block_name in ('*Model_Space', '*Paper_Space', '*Paper_Space0'):
                            continue
                        # 如果新文档中没有这个块，则创建
                        if block_name not in new_doc.blocks:
                            try:
                                new_block = new_doc.blocks.new(name=block_name)
                                # 复制块的属性
                                try:
                                    new_block.dxf.description = src_block.dxf.description
                                except Exception:
                                    pass
                                try:
                                    new_block.dxf.base_point = src_block.dxf.base_point
                                except Exception:
                                    pass
                            except Exception as e:
                                print(f"创建块 {block_name} 时出错: {e}")
                except Exception as e:
                    print(f"复制块定义时出错: {str(e)}")

                target_msp = new_doc.modelspace()
                importer = Importer(self.doc, new_doc)
                importer.import_entities(selected_entities, target_msp)
                importer.finalize()
                for block in new_doc.blocks:
                    block_name = block.dxf.name
                    if not isinstance(block_name, str):
                        print(f"跳过无效块名（非字符串类型）: {block_name}")
                        continue
                    src_block = self.doc.blocks.get(block_name)
                    if src_block and block:
                        try:
                            block.dxf.units = src_block.dxf.units
                        except Exception:
                            pass
                        try:
                            block.dxf.xscale = src_block.dxf.xscale
                            block.dxf.yscale = src_block.dxf.yscale
                            block.dxf.zscale = src_block.dxf.zscale
                        except Exception:
                            pass

                fname = self.number_extractor.extract_region_filename_by_patterns(region)
                if not fname:
                    drawing_number = self.number_extractor.extract_drawing_number_from_region(region)
                    if drawing_number:
                        fname = self.number_extractor.generate_safe_filename(drawing_number)
                if not fname:
                    fname = region_id

                def unique_path(path: str) -> str:
                    if not os.path.exists(path):
                        return path
                    root, ext = os.path.splitext(path)
                    k = 2
                    while True:
                        cand = f"{root}({k}){ext}"
                        if not os.path.exists(cand):
                            return cand
                        k += 1

                filename = f"{idx:03d}_{fname}.dxf"
                save_path = unique_path(os.path.join(output_dir, filename))
                new_doc.saveas(save_path)

                region['exported_dxf'] = save_path
                region['export_basename'] = fname
                print(f"[OK] 子图 {region_id} 已导出：{save_path}")
                print(f"    导出实体数量: {len(selected_entities)}")

                export_count += 1

            print(f"\n导出完成：共导出 {export_count} 个满足条件的子图")
            return output_dir

        def _extract_layer_colors(self, doc):
            # 提取所有图层的颜色信息
            print("正在提取图层颜色信息...")
            try:
                for layer in doc.layers:
                    self.layer_colors[layer.dxf.name] = getattr(layer.dxf, 'color', 7)
                print(f"图层颜色信息提取完成: {len(self.layer_colors)} 个图层")
            except Exception as e:
                print(f"图层颜色提取失败: {e}")

        def _extract_all_entities(self, msp):
            # 提取模型空间中的所有几何实体
            print("正在提取几何实体...")
            geometric_types = ['LINE', 'CIRCLE', 'ARC', 'LWPOLYLINE', 'POLYLINE', 'ELLIPSE', 'SPLINE', 'INSERT','DIMENSION']
            for entity_type in geometric_types:
                try:
                    for entity in msp.query(entity_type):
                        info = self._process_geometric_entity(entity)
                        if info:
                            self.all_entities.append(info)
                except Exception:
                    continue
            print(f"几何实体提取完成: {len(self.all_entities)} 个实体")

        def _process_geometric_entity(self, entity) -> Optional[Dict]:
            # 处理单个几何实体，返回字典信息
            try:
                t = entity.dxftype()
                layer = getattr(entity.dxf, 'layer', '0')
                color = getattr(entity.dxf, 'color', 256)
                handle = getattr(entity.dxf, 'handle', 'N/A')
                linetype = getattr(entity.dxf, 'linetype', 'ByLayer')
                center = self._calculate_entity_center(entity)
                perimeter = self._calculate_entity_perimeter(entity)

                entity_info = {
                    'type': t, 'layer': layer, 'entity_color': color, 'handle': handle,
                    'linetype': linetype, 'center': center, 'perimeter': perimeter
                }

                if t == 'INSERT':
                    entity_info['insert_entity'] = entity
                    entity_info['block_name'] = entity.dxf.name

                return entity_info
            except Exception as e:
                print(f"处理实体时出错: {e}")
                return None

        def _calculate_entity_center(self, entity):
            # 计算实体中心点
            try:
                t = entity.dxftype()
                if t in ['CIRCLE', 'ARC']:
                    c = entity.dxf.center
                    return (round(c.x, 2), round(c.y, 2))
                elif t == 'LINE':
                    s, e = entity.dxf.start, entity.dxf.end
                    return (round((s.x + e.x) / 2, 2), round((s.y + e.y) / 2, 2))
                elif t in ['LWPOLYLINE', 'POLYLINE']:
                    pts = entity.get_points(format='xy')
                    if pts:
                        xs = [p[0] for p in pts]
                        ys = [p[1] for p in pts]
                        return (round(sum(xs) / len(xs), 2), round(sum(ys) / len(ys), 2))
                elif t == 'ELLIPSE':
                    c = entity.dxf.center
                    return (round(c.x, 2), round(c.y, 2))
                elif t == 'SPLINE':
                    pts = self._safe_spline_points(entity)
                    if len(pts) >= 2:
                        xs = [p[0] for p in pts]
                        ys = [p[1] for p in pts]
                        return (round(sum(xs) / len(xs), 2), round(sum(ys) / len(ys), 2))
                elif t == 'INSERT':
                    insert_pt = entity.dxf.insert
                    return (round(insert_pt.x, 2), round(insert_pt.y, 2))
            except Exception:
                pass
            return (0.0, 0.0)

        def _calculate_entity_perimeter(self, entity):
            # 计算实体周长
            try:
                t = entity.dxftype()
                if t == 'CIRCLE':
                    r = entity.dxf.radius
                    return round(2 * pi * r, 2)
                elif t == 'ARC':
                    r = entity.dxf.radius
                    sa = radians(entity.dxf.start_angle)
                    ea = radians(entity.dxf.end_angle)
                    if ea < sa:
                        ea += 2 * pi
                    return round(r * (ea - sa), 2)
                elif t == 'LINE':
                    s, e = entity.dxf.start, entity.dxf.end
                    return round(sqrt((e.x - s.x) ** 2 + (e.y - s.y) ** 2), 2)
                elif t in ['LWPOLYLINE', 'POLYLINE']:
                    return round(self._calculate_polyline_length(entity), 2)
                elif t == 'ELLIPSE':
                    mx = entity.dxf.major_axis
                    a = (mx.x ** 2 + mx.y ** 2 + mx.z ** 2) ** 0.5
                    ratio = float(getattr(entity.dxf, 'ratio', 0.5) or 0.5)
                    b = a * ratio
                    h = ((a - b) ** 2) / ((a + b) ** 2) if (a + b) != 0 else 0.0
                    per = pi * (a + b) * (1 + (3 * h) / (10 + (4 - 3 * h) ** 0.5))
                    return round(per, 2)
                elif t == 'SPLINE':
                    pts = self._safe_spline_points(entity)
                    if len(pts) >= 2:
                        total_length = 0.0
                        for i in range(len(pts) - 1):
                            x1, y1 = pts[i][0], pts[i][1]
                            x2, y2 = pts[i + 1][0], pts[i + 1][1]
                            total_length += sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
                        if getattr(entity.dxf, 'flags', 0) & 1 and len(pts) > 2:
                            x1, y1 = pts[-1][0], pts[-1][1]
                            x2, y2 = pts[0][0], pts[0][1]
                            total_length += sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
                        return round(total_length, 2)
                elif t == 'INSERT':
                    block_def = entity.doc.blocks.get(entity.dxf.name)
                    if block_def:
                        bounds = self._calculate_block_bounds(block_def, entity)
                        if bounds:
                            return 2 * (bounds['width'] + bounds['height'])
            except Exception:
                pass
            return 0.0

        def _calculate_polyline_length(self, polyline):
            # 计算多段线长度
            try:
                pts = polyline.get_points(format='xy')
                if len(pts) < 2:
                    return 0.0
                total = 0.0
                for i in range(len(pts) - 1):
                    x1, y1 = pts[i]
                    x2, y2 = pts[i + 1]
                    total += sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
                if getattr(polyline, 'closed', False) and len(pts) > 2:
                    x1, y1 = pts[-1]
                    x2, y2 = pts[0]
                    total += sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
                return total
            except Exception:
                return 0.0

        def _extract_all_texts(self, msp):
            # 提取模型空间中的所有文字实体
            print("正在提取文字实体...")
            for t in ['TEXT', 'MTEXT', 'ATTRIB', 'ATTDEF']:
                try:
                    for e in msp.query(t):
                        info = self._process_text_entity(e)
                        if info:
                            self.all_texts.append(info)
                except Exception:
                    continue
            print(f"文字提取完成: {len(self.all_texts)} 个实体")

        def _process_text_entity(self, entity) -> Optional[Dict]:
            # 处理单个文字实体，返回字典信息
            try:
                content = self._extract_text_content(entity)
                if not content:
                    return None
                position = self._get_text_position(entity)
                if not position:
                    return None
                return {
                    'content': self._clean_text_content(content),
                    'position': position,
                    'entity_type': entity.dxftype()
                }
            except Exception:
                return None

        def _extract_text_content(self, entity) -> Optional[str]:
            # 提取文字内容
            t = entity.dxftype()
            try:
                if t == 'TEXT':
                    return entity.dxf.text
                elif t == 'MTEXT':
                    if hasattr(entity, 'get_text'):
                        return entity.get_text()
                    elif hasattr(entity, 'plain_text'):
                        return entity.plain_text()
                    return getattr(entity.dxf, 'text', None)
                elif t in ['ATTRIB', 'ATTDEF']:
                    return entity.dxf.text
                # elif t == 'DIMENSION':
                #     if hasattr(entity, 'get_measurement'):
                #         return str(entity.get_measurement())
                #     return getattr(entity.dxf, 'text', None)
            except Exception:
                pass
            return None

        def _get_text_position(self, entity) -> Optional[Tuple[float, float]]:
            # 获取文字位置
            try:
                if hasattr(entity.dxf, 'insert'):
                    p = entity.dxf.insert
                    return (float(p.x), float(p.y))
                elif hasattr(entity.dxf, 'position'):
                    p = entity.dxf.position
                    return (float(p.x), float(p.y))
                elif hasattr(entity.dxf, 'start'):
                    p = entity.dxf.start
                    return (float(p.x), float(p.y))
            except Exception:
                pass
            return None

        def _clean_text_content(self, content: str) -> str:
            # 清理文字内容
            if not content:
                return ""
            content = re.sub(r'\{\\[^}]*\}', '', content)
            content = re.sub(r'\\[A-Za-z][^;]*;', '', content)
            repl = {'%%c': 'Φ', '%%C': 'Φ', '%%d': '°', '%%D': '°', '%%p': '±', '%%P': '±', '%%u': '', '%%U': '',
                    '%%o': '', '%%O': ''}
            for k, v in repl.items():
                content = content.replace(k, v)
            return re.sub(r'\s+', ' ', content).strip()

        def set_classify_map(self, classify: Dict[str, List[str]]):
            # 设置分类映射
            self.classify_map = classify or {}

        def _identify_frame_blocks(self, msp):
            # 识别所有块引用作为图框块
            print("正在识别块图框（修复 BlockLayout 处理）...")
            self.frame_blocks = []
            all_inserts = []

            def collect_inserts(entity):
                if not hasattr(entity, 'dxftype'):
                    for child in entity:
                        collect_inserts(child)
                    return
                if entity.dxftype() == 'INSERT':
                    all_inserts.append(entity)
                if hasattr(entity, '__iter__'):
                    try:
                        for child in entity:
                            collect_inserts(child)
                    except TypeError:
                        pass

            collect_inserts(msp)
            print(f"共收集到 {len(all_inserts)} 个块引用（含匿名块）")

            block_candidates = []
            for idx, insert in enumerate(all_inserts):
                try:
                    block_name = insert.dxf.name
                    block_def = insert.doc.blocks.get(block_name)
                    if not block_def:
                        print(f"块 {block_name} 无定义，跳过")
                        continue

                    bounds = self._calculate_block_bounds(block_def, insert)
                    if not bounds:
                        print(f"块 {block_name} 无法计算边界，跳过")
                        continue

                    # print(f"块 {block_name} (序号{idx})：宽={bounds['width']:.1f}, 高={bounds['height']:.1f}")

                    if self._is_valid_frame_block(bounds):
                        block_candidates.append({
                            'type': 'block',
                            'block_name': block_name,
                            'insert_point': (insert.dxf.insert.x, insert.dxf.insert.y),
                            'bounds': bounds,
                            'insert_entity': insert
                        })
                except Exception as e:
                    print(f"处理块 {insert.dxf.name if hasattr(insert.dxf, 'name') else '未知'} 时出错: {e}")
                    continue

            self.frame_blocks = self._deduplicate_frames(block_candidates)
            print(f"有效图框块数量: {len(self.frame_blocks)}")

        def _calculate_block_bounds(self, block_def, insert) -> Optional[Dict]:
            # 修复块边界计算中的 BlockLayout 处理逻辑
            """修复块边界计算中的 BlockLayout 处理逻辑"""
            try:
                min_x = min_y = float('inf')
                max_x = max_y = float('-inf')
                has_entities = False

                sx = getattr(insert.dxf, 'xscale', 1.0)
                sy = getattr(insert.dxf, 'yscale', 1.0)
                rotation = getattr(insert.dxf, 'rotation', 0.0)
                rot_rad = radians(rotation)
                insert_pt = insert.dxf.insert

                def traverse_entity(entity):
                    nonlocal min_x, max_x, min_y, max_y, has_entities

                    if entity.dxftype() == 'INSERT':
                        nested_block = entity.doc.blocks.get(entity.dxf.name)
                        if nested_block:
                            for nested_ent in nested_block:
                                traverse_entity(nested_ent)
                        return

                    entity_bounds = self._get_entity_local_bounds(entity)
                    if not entity_bounds:
                        return
                    has_entities = True

                    local_corners = [
                        (entity_bounds['min_x'], entity_bounds['min_y']),
                        (entity_bounds['max_x'], entity_bounds['min_y']),
                        (entity_bounds['min_x'], entity_bounds['max_y']),
                        (entity_bounds['max_x'], entity_bounds['max_y']),
                    ]

                    for (x, y) in local_corners:
                        x_scaled = x * sx
                        y_scaled = y * sy
                        x_rot = x_scaled * cos(rot_rad) - y_scaled * sin(rot_rad)
                        y_rot = x_scaled * sin(rot_rad) + y_scaled * cos(rot_rad)
                        x_world = insert_pt.x + x_rot
                        y_world = insert_pt.y + y_rot
                        min_x = min(min_x, x_world)
                        max_x = max(max_x, x_world)
                        min_y = min(min_y, y_world)
                        max_y = max(max_y, y_world)

                for ent in block_def:
                    traverse_entity(ent)

                if not has_entities:
                    return None

                return {
                    'min_x': min_x, 'max_x': max_x,
                    'min_y': min_y, 'max_y': max_y,
                    'width': max_x - min_x, 'height': max_y - min_y
                }
            except Exception as e:
                print(f"计算块边界出错: {e}")
                return None

        def _get_entity_local_bounds(self, entity) -> Optional[Dict]:
            # 获取实体本地边界
            try:
                t = entity.dxftype()
                if t == 'LINE':
                    s, e = entity.dxf.start, entity.dxf.end
                    return {
                        'min_x': min(s.x, e.x), 'max_x': max(s.x, e.x),
                        'min_y': min(s.y, e.y), 'max_y': max(s.y, e.y)
                    }
                elif t in ['CIRCLE', 'ARC']:
                    c = entity.dxf.center
                    r = entity.dxf.radius
                    return {
                        'min_x': c.x - r, 'max_x': c.x + r,
                        'min_y': c.y - r, 'max_y': c.y + r
                    }
                elif t in ['LWPOLYLINE', 'POLYLINE']:
                    pts = entity.get_points(format='xy')
                    if pts:
                        xs, ys = zip(*pts)
                        return {'min_x': min(xs), 'max_x': max(xs), 'min_y': min(ys), 'max_y': max(ys)}
                elif t == 'ELLIPSE':
                    c = entity.dxf.center
                    major = entity.dxf.major_axis
                    ratio = getattr(entity.dxf, 'ratio', 0.5)
                    a = (major.x ** 2 + major.y ** 2) ** 0.5
                    b = a * ratio
                    return {'min_x': c.x - a, 'max_x': c.x + a, 'min_y': c.y - b, 'max_y': c.y + b}
                elif t == 'SPLINE':
                    pts = self._safe_spline_points(entity)
                    if pts:
                        xs = [p[0] for p in pts]
                        ys = [p[1] for p in pts]
                        return {'min_x': min(xs), 'max_x': max(xs), 'min_y': min(ys), 'max_y': max(ys)}
            except Exception as e:
                print(f"获取实体本地边界出错 ({entity.dxftype()}): {e}")
            return None

        def _is_valid_frame_block(self, bounds: Dict) -> bool:
            # 判断块是否为有效图框
            min_size = 30
            aspect_ratio = bounds['width'] / bounds['height'] if bounds['height'] != 0 else 0
            return (bounds['width'] > min_size and
                    bounds['height'] > min_size and
                    0.2 < aspect_ratio < 5)

        def _deduplicate_frames(self, candidates):
            # 图框去重
            if len(candidates) <= 1:
                return candidates

            candidates.sort(key=lambda x: (x['bounds']['width'] * x['bounds']['height']), reverse=True)
            unique = [candidates[0]]

            for c in candidates[1:]:
                c_bounds = c['bounds']
                overlap = False
                for u in unique:
                    u_bounds = u['bounds']
                    if (c_bounds['max_x'] > u_bounds['min_x'] and
                            c_bounds['min_x'] < u_bounds['max_x'] and
                            c_bounds['max_y'] > u_bounds['min_y'] and
                            c_bounds['min_y'] < u_bounds['max_y']):
                        overlap = True
                        break
                if not overlap:
                    unique.append(c)
            return unique

        def _create_subdrawing_regions(self):
            # 根据图框块创建子图区域并筛选
            """创建子图区域时直接应用筛选条件"""
            self.frame_blocks.sort(key=self._get_spatial_sort_key)
            valid_count = 0

            for i, fb in enumerate(self.frame_blocks):
                rid = f"subdrawing_{i + 1:03d}"

                region_data = {
                    'frame_block': fb,
                    'bounds': fb['bounds'],
                    'texts': [],
                    'cutting_analysis': {}
                }

                self._assign_texts_to_single_region(region_data)

                if self._should_process_region(region_data):
                    self.sub_drawings[rid] = region_data
                    valid_count += 1
                    # print(f"区域 {rid} 满足条件，将被处理")
                else:
                    pass
                    # print(f"区域 {rid} 不满足条件，跳过")

            print(f"区域创建完成：{len(self.frame_blocks)} 个图框，{valid_count} 个满足条件")

        def _assign_texts_to_single_region(self, region_data: Dict):
            # 为单个区域分配文字
            """为单个区域分配文字"""
            bounds = region_data['bounds']
            region_texts = []

            for text in self.all_texts:
                x, y = text['position']
                if (bounds['min_x'] <= x <= bounds['max_x'] and
                        bounds['min_y'] <= y <= bounds['max_y']):
                    region_texts.append(text)

            region_data['texts'] = self.text_processor.process_text_list(region_texts)

        def _get_spatial_sort_key(self, frame_block):
            # 区域空间排序键
            b = frame_block['bounds']
            tol = 100
            return (-round(b['min_y'] / tol), round(b['min_x'] / tol))

        def _assign_texts_to_regions(self):
            """分配文字到区域"""
            print("正在分配文字到区域...")

            for text in self.all_texts:
                text_x, text_y = text['position']
                assigned = False

                # 查找包含文字的区域
                for region_id, region_data in self.sub_drawings.items():
                    bounds = region_data['bounds']

                    if (bounds['min_x'] <= text_x <= bounds['max_x'] and
                            bounds['min_y'] <= text_y <= bounds['max_y']):
                        region_data['texts'].append(text)
                        assigned = True
                        break

                # 分配到最近区域
                if not assigned:
                    closest_region = self._find_closest_region((text_x, text_y))
                    if closest_region:
                        self.sub_drawings[closest_region]['texts'].append(text)

            # 处理各区域文字
            for region_id, region_data in self.sub_drawings.items():
                original_count = len(region_data['texts'])
                region_data['texts'] = self.text_processor.process_text_list(
                    region_data['texts'])
                processed_count = len(region_data['texts'])

                # print(f"区域 {region_id}: 处理前 {original_count} 个文字，"
                #    f"处理后 {processed_count} 个文字")
                
        def _analyze_cutting_contours_for_regions(self):
            """为各区域分析切割轮廓"""
            print("正在分析各区域的切割轮廓（采用严格策略）...")

            for region_id, region_data in self.sub_drawings.items():
                bounds = region_data['bounds']

                # 检测该区域的切割轮廓
                cutting_result = self.cutting_detector.detect_cutting_contours_in_region(
                    bounds, self.all_entities, self.layer_colors)

                region_data['cutting_analysis'] = cutting_result

                # print(f"区域 {region_id}: 检测到 {cutting_result['contour_count']} 个切割轮廓，"
                #    f"总长度 {cutting_result['total_cutting_length']:.2f}mm")


        def _find_closest_region(self, pos: Tuple[float, float]) -> Optional[str]:
            # 查找最近区域
            x, y = pos
            md = float('inf')
            cid = None
            for rid, r in self.sub_drawings.items():
                b = r['bounds']
                cx = (b['min_x'] + b['max_x']) / 2
                cy = (b['min_y'] + b['max_y']) / 2
                d = sqrt((x - cx) ** 2 + (y - cy) ** 2)
                if d < md:
                    md = d
                    cid = rid
            return cid


    class StrictCuttingDetector:
        """严格的切割轮廓检测器 - 仅识别直接红色实体"""

        def __init__(self):
            # 切割相关颜色代码（红色系）
            self.cutting_colors = [1]

            # 明确排除的图层模式
            self.exclude_layer_patterns = [
                r'.*text.*', r'.*dim.*', r'.*annotation.*', r'.*title.*',
                r'.*border.*', r'.*frame.*', r'.*标注.*', r'.*文字.*', r'.*尺寸.*',
                r'.*defpoints.*'  # AutoCAD默认点图层
            ]

            # 几何实体类型
            self.geometric_entities = ['LINE', 'CIRCLE', 'ARC', 'LWPOLYLINE', 'POLYLINE', 'ELLIPSE']
            self.BYLAYER_COLOR = 256

            # 图层颜色映射
            self.layer_colors = {}

        def detect_cutting_contours_in_region(self, region_bounds: Dict, all_entities: List,
                                            layer_colors: Dict) -> Dict:
            """检测指定区域内的切割轮廓 - 采用严格策略"""
            self.layer_colors = layer_colors

            # 获取区域内的所有几何实体
            region_entities = self._get_entities_in_bounds(all_entities, region_bounds)

            # 采用严格策略识别红色实体
            red_entities = []
            for entity in region_entities:
                if self._is_red_geometric_entity_strict(entity):
                    red_entities.append(entity)

            # 分类实体
            cutting_contours = []
            auxiliary_marks = []

            for entity_info in red_entities:
                if self._should_exclude_entity(entity_info):
                    auxiliary_marks.append(entity_info)
                else:
                    cutting_contours.append(entity_info)

            # 计算统计数据
            total_cutting_length = sum(e.get('perimeter', 0) for e in cutting_contours)
            contour_count = len(cutting_contours)

            return {
                'cutting_contours': cutting_contours,
                'auxiliary_marks': auxiliary_marks,
                'total_cutting_length': round(total_cutting_length, 2),
                'contour_count': contour_count,
                'contour_types': self._get_contour_types(cutting_contours),
                'cutting_analysis': self._generate_cutting_analysis(cutting_contours)
            }

        def _get_entities_in_bounds(self, entities: List, bounds: Dict) -> List:
            """获取边界内的实体"""
            region_entities = []

            for entity_info in entities:
                center = entity_info.get('center', (0, 0))
                if (bounds['min_x'] <= center[0] <= bounds['max_x'] and
                        bounds['min_y'] <= center[1] <= bounds['max_y']):
                    region_entities.append(entity_info)

            return region_entities

        def _is_red_geometric_entity_strict(self, entity_info: Dict) -> bool:
            entity_type = entity_info.get('type', '')
            if entity_type not in self.geometric_entities:
                return False
            entity_color = entity_info.get('entity_color', self.BYLAYER_COLOR)
            if entity_color not in self.cutting_colors:
                return False
            linetype = entity_info.get('linetype', 'ByLayer')
            if linetype.lower() not in ['continuous', 'bylayer']:
                return False
            return True

        def _should_exclude_entity(self, entity_info: Dict) -> bool:
            """判断是否应该排除实体"""
            layer_name = entity_info.get('layer', '')
            layer_name_lower = layer_name.lower()

            # 排除明确的文字、标注、边框图层
            for pattern in self.exclude_layer_patterns:
                if re.match(pattern, layer_name_lower, re.IGNORECASE):
                    return True

            return False

        def _get_contour_types(self, contours: List) -> Dict:
            """获取轮廓类型分布"""
            type_count = defaultdict(int)
            for contour in contours:
                type_count[contour.get('type', 'UNKNOWN')] += 1
            return dict(type_count)

        def _generate_cutting_analysis(self, contours: List) -> Dict:
            """生成切割分析数据"""
            if not contours:
                return {'summary': '未检测到切割轮廓'}

            perimeters = [c.get('perimeter', 0) for c in contours if c.get('perimeter', 0) > 0]

            analysis = {
                'total_length': sum(perimeters),
                'contour_count': len(contours),
                'avg_length': sum(perimeters) / len(perimeters) if perimeters else 0,
                'min_length': min(perimeters) if perimeters else 0,
                'max_length': max(perimeters) if perimeters else 0
            }

            # 生成描述性摘要
            if analysis['contour_count'] > 0:
                analysis['summary'] = (f"检测到{analysis['contour_count']}个切割轮廓，"
                                    f"总长度{analysis['total_length']:.2f}mm")
            else:
                analysis['summary'] = "未检测到有效切割轮廓"

            return analysis

    # 2D拆图主逻辑
    def IntegralMoudle_Splitting(input_file: str, output_root: str):
        print("集成CAD子图分析与切割轮廓检测系统")
        print("=" * 60)
        print(f"输入整图: {input_file}")
        print(f"子图DXF输出目录: {output_root}")

        if not os.path.exists(input_file):
            print(f"错误：文件不存在 - {input_file}")
            return None
        if not input_file.lower().endswith('.dxf'):
            print("错误：仅支持DXF格式文件，请先转换DWG为DXF")
            return None

        try:
            doc = ezdxf.readfile(input_file)
            print(f"调试：DXF版本 {doc.dxfversion}，单位 {doc.units}")
            print(f"调试：模型空间实体总数 {len(list(doc.modelspace()))}")
        except Exception as e:
            print(f"读取文件信息失败：{e}")

        analyzer = OptimizedCADBlockAnalyzer()
        results = analyzer.analyze_cad_file(input_file)
        if not results:
            print("未提取到任何满足条件的子图信息，可能原因：")
            print("1. 图框不包含'加工说明'文字")
            print("2. 图框包含排除词汇（厂内标准件、订购、装配、组配）")
            print("3. 图框不是块引用，且线条组合未被识别")
            print("4. 图框尺寸小于最小阈值（当前30单位）")
            return None

        if not output_root:
            output_root = os.path.join(os.path.dirname(input_file), "subdrawings_dxf")
        os.makedirs(output_root, exist_ok=True)
        print(f"\n开始导出满足条件的子图DXF到目录：{output_root}")

        exported_dir = analyzer.export_regions_to_dxf(output_root)

        print(f"\n分析与导出完成！子图DXF目录: {exported_dir}")

        return exported_dir

else:
    # ezdxf不可用时的占位函数
    def IntegralMoudle_Splitting(input_file: str, output_root: str):
        print("ezdxf 库未加载，跳过 2D 图纸拆分。")
        return None
    

# ==============================================================================
# 包装函数
# ==============================================================================

def split_dxf_file_with_output(input_dxf: str, output_dir: str):
    # DXF文件拆分 - 支持自定义输出目录，校验环境和输入，调用主拆分逻辑
    """
    DXF文件拆分 - 支持自定义输出目录
    
    Args:
        input_dxf: 输入的DXF文件路径
        output_dir: 输出目录路径
    
    Returns:
        str: 导出的子图目录路径，失败返回None
    """
    if not EZDXF_AVAILABLE:
        print("错误: ezdxf库未安装，无法进行DXF拆分")
        return None
    
    if not os.path.exists(input_dxf):
        print(f"错误: 输入文件不存在 - {input_dxf}")
        return None
    
    result = IntegralMoudle_Splitting(input_dxf, output_dir)
    
    import gc
    gc.collect()
    
    return result


if __name__ == "__main__":
    # 测试代码
    # input_dxf = "M250209-P4.2025.11.25.dxf"
    # output_dir = os.path.join(os.path.dirname(__file__), "split")
    # result = split_dxf_file_with_output(input_dxf, output_dir)
    import config
    from path_manager import init_path_manager
    
    pm = init_path_manager(config.FILE_INPUT_PRT, config.FILE_INPUT_DXF)
    result = split_dxf_file_with_output(pm.get_input_2d_dxf(), pm.get_split_dxf_dir())
    if result:
        print(f"\n✅ DXF拆分完成！\n   输出目录: {result}")
    else:
        print("\n❌ DXF拆分失败")