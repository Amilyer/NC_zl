import csv
import re
import NXOpen


class ProcessInfoHandler:
    """加工信息处理类"""

    def __init__(self, session, work_part):
        self.session = session
        self.work_part = work_part

    def safe_origin(self, note):
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

    def extract_number(self, text):
        """提取字符串开头的数字（支持整数/小数）"""
        match = re.match(r'(\d+(?:\.\d+)?)', text)
        return float(match.group(1)) if match else 0

    def sum_pcs(self, pcs_list):
        """求和PCS数量"""
        total = 0
        for pcs in pcs_list:
            match = re.search(r'(\d+)PCS', pcs, re.IGNORECASE)
            if match:
                total += int(match.group(1))
        return [f"{total - 1}PCS"] if total > 0 else []

    def merge_process_notes(self, note_list):
        """
        动态合并加工说明：
        - 保留字段：加工对象/尺寸/材质/热处理/重量/数量（按规则处理）
        - 其他字段：自动识别，列表类型合并去重，字符串类型数字相加
        """
        if not note_list:
            return {}

        merged = {}
        # 定义固定保留字段（仅这部分特殊处理）
        keep_fields = ["加工对象", "尺寸", "材质", "热处理", "重量", "数量"]

        # ------------------- 处理保留字段 -------------------
        # 1. 加工对象/尺寸：取第一个字典的原值
        for field in ["加工对象", "尺寸"]:
            merged[field] = note_list[0].get(field, "")

        # 2. 材质/热处理/重量：取第一个字典的列表（保持原样）
        for field in ["材质", "热处理", "重量"]:
            merged[field] = note_list[0].get(field, []).copy()

        # 3. 数量：求和处理
        all_pcs = []
        for note in note_list:
            all_pcs.extend(note.get("数量", []))
        merged["数量"] = self.sum_pcs(all_pcs)

        # ------------------- 动态处理非保留字段 -------------------
        # 收集所有非保留字段（去重）
        dynamic_keys = set()
        for note in note_list:
            for key in note.keys():
                if key not in keep_fields:
                    dynamic_keys.add(key)

        for key in dynamic_keys:
            # 收集所有字典中该键的有效值
            values = []
            for note in note_list:
                val = note.get(key)
                if val:  # 过滤空值
                    values.append(val)

            if not values:
                merged[key] = [] if isinstance(values, list) else ""
                continue

            # 处理列表类型（如"其它注释"）
            if isinstance(values[0], list):
                merged_list = []
                for val in values:
                    for item in val:
                        if item not in merged_list:
                            merged_list.append(item)
                merged[key] = merged_list

            # 处理字符串类型（如G/K/L/其他任意字符串键）
            elif isinstance(values[0], str):
                # 提取所有值的开头数字求和
                total_num = sum(self.extract_number(v) for v in values)
                # 保留第一个值的格式，替换开头数字
                first_val = values[0]
                if '-' in first_val:
                    prefix, suffix = first_val.split('-', 1)
                    merged[key] = f"{int(total_num)}-{suffix}"
                else:
                    # 无分隔符时，直接替换开头数字
                    num_match = re.match(r'\d+(?:\.\d+)?', first_val)
                    if num_match:
                        merged[key] = first_val.replace(num_match.group(), str(int(total_num)), 1)
                    else:
                        merged[key] = first_val  # 无数字则保留原值
        return merged

    def extract_and_process_notes(self, work_part):
        """从NX当前工作部件中提取所有注释（含多列排序逻辑）"""
        print("=== 开始提取加工说明注释 ===")
        note_list = list(work_part.Notes)
        print(f"=== 共找到 {len(note_list)} 条注释 ===")

        # 过滤非加工信息注释
        note_str = []
        notes = []
        for note in note_list:
            text_str = "".join(note.GetText()).strip()
            if re.fullmatch(r'[A-Za-z]\d*', text_str):
                continue
            try:
                float(text_str)
                continue
            except:
                note_str.append(text_str)
                notes.append(note)
        # 确保加工说明在开头（若不在则反转列表）
        if note_str and "加工说明" not in note_str[0]:
            notes = notes[::-1]

        parent = {}  # 存储父加工说明：{f"{x_str}_{y_str}": 加工说明文本}
        son = []  # 存储子条目：[(coords, text_str, 材质, 重量, 热处理, PCS)]
        # 材质匹配列表
        materials_master = ["45#", "A3", "CR12", "CR12MOV", "SKD11", "SKH-9",
                            "DC53", "P20", "TOOLOX33", "TOOLOX44", "合金铜",
                            "T00L0X33", "T00L0X44"]

        for idx, note in enumerate(notes, start=1):
            try:
                texts = note.GetText()
                text_str = "".join(texts).strip()
                if "HRC" not in text_str:
                    text_str = text_str.replace(" ", "")
                x, y, z = self.safe_origin(note)
                if x is None or y is None:
                    continue
                x_str, y_str = f"{x:.3f}", f"{y:.3f}"
                coords = (x_str, y_str)

                # 识别父加工说明
                if "加工说明" in text_str:
                    parent[f"{x_str}_{y_str}"] = text_str
                    continue

                # 提取材质
                material_list = []
                _material = None
                length = 0
                for mat in materials_master:
                    if mat.lower() in text_str.lower():
                        if len(mat) > length:
                            _material = mat
                            length = len(mat)
                corrected_mat = "TOOLOX33" if _material == "T00L0X33" else (
                    "TOOLOX44" if _material == "T00L0X44" else _material)
                if corrected_mat:
                    material_list.append(corrected_mat)

                # 提取重量
                weight_match = re.search(r'(\d+(?:\.\d+)?)KG', text_str)
                weight = f"{weight_match.group(1)}KG" if weight_match else None

                # 提取热处理
                heat_treatment = []
                if "HRC" in text_str:
                    hrc_idx = text_str.find("HRC")
                    heat_treatment.append(text_str[hrc_idx:].split()[0].replace(")", ""))

                # 提取PCS
                pcs_match = re.search(r'(\d+)PCS', text_str, re.IGNORECASE)
                pcs = f"{pcs_match.group(1)}PCS" if pcs_match else None

                # 加入子条目列表
                son.append((coords, text_str, material_list, weight, heat_treatment, pcs))

            except Exception as e:
                print(f"处理注释{idx}时出错：{str(e)}")
                continue

        # 处理数据排序与合并
        raw_data_dict = self.process_data(parent, son)
        # 整理最终结果格式
        final_dicts = {"主体加工说明": [], "共出板加工说明": []}
        for key in raw_data_dict:
            final_dict = {"其它注释": []}
            final_dict["尺寸"] = None
            for data in raw_data_dict[key]["文本列表"]:
                # 解析加工对象
                if "加工说明" in data:
                    text_parts = data.split(":")
                    final_dict["加工对象"] = text_parts[-1].strip() if len(text_parts) > 1 else None
                    continue
                # 解析带冒号的条目
                colon_idx = data.find(":")
                if colon_idx != -1 and "尺寸" not in data:
                    k = data[:colon_idx].strip()
                    value = data[colon_idx + 1:].strip()
                    if "GW" not in k and k not in final_dict:
                        final_dict[k] = value
                    continue
                # 解析尺寸（L*W*T）
                l_match = re.search(r'(\d+(?:\.\d+)?)L', data)
                w_match = re.search(r'(\d+(?:\.\d+)?)W', data)
                t_match = re.search(r'(\d+(?:\.\d+)?)T', data)
                if l_match and w_match and t_match:
                    final_dict['尺寸'] = f"{l_match.group(1)}L*{w_match.group(1)}W*{t_match.group(1)}T"
                    continue
                # 未分类的注释去重
                if data and data not in final_dict["其它注释"]:
                    final_dict["其它注释"].append(data)

            # 整理属性信息（去重+默认值）
            attrs = raw_data_dict[key]["属性信息"]
            final_dict["材质"] = list(set(attrs["材质"])) if attrs["材质"] else []
            final_dict["热处理"] = list(set(attrs["热处理"])) if attrs["热处理"] else []
            final_dict["重量"] = list(set(attrs["重量"])) if attrs["重量"] else []
            final_dict["数量"] = list(set(attrs["PCS"])) if attrs["PCS"] else ["1PCS"]
            if "共出板注解" in f"{final_dict}":
                final_dicts["共出板加工说明"].append(final_dict)
                continue
            final_dicts["主体加工说明"].append(final_dict)

        subject_len = len(final_dicts["主体加工说明"])

        if subject_len < 1:
            return final_dicts
        if subject_len > 1:
            final_dicts["主体加工说明"] = self.merge_process_notes(final_dicts["主体加工说明"])
        else:
            final_dicts["主体加工说明"] = final_dicts["主体加工说明"][0]

        return final_dicts

    # 辅助函数：判断文本是否满足"包含冒号且冒号前匹配[A-Za-z]\d*"
    def is_valid_prev(self, text):
        if ':' not in text:
            return False
        prefix = text.split(':', 1)[0]
        return re.fullmatch(r'[A-Za-z](?:\d+[A-Za-z]?)?', prefix) is not None

    # 独立函数：处理文本拼接和属性收集（含向上追溯+逗号去重逻辑）
    def merge_texts_and_collect_attrs(self, tuples_list):
        merged = []
        attrs = {"材质": [], "重量": [], "热处理": [], "PCS": []}  # 包含PCS字段

        for i in range(len(tuples_list)):
            _, text, materials, weight, heat_treatment, pcs = tuples_list[i]

            # 收集属性
            if materials:
                attrs["材质"].extend(materials)
            if weight:
                attrs["重量"].append(weight)
            if heat_treatment:
                attrs["热处理"].extend(heat_treatment)
            if pcs:
                attrs["PCS"].append(pcs)

            # 无已拼接元素时直接添加
            if not merged:
                merged.append(text)
                continue

            # 获取上下文信息（原列表中的下一个元素）
            has_next = i + 1 < len(tuples_list)
            next_text = tuples_list[i + 1][1] if has_next else ""

            # 判定是否需要拼接（原有条件）
            need_merge = False
            is_new_condition = False

            # 原有条件：当前元素开头是逗号
            original_condition = text.startswith(',')
            # 新增条件：当前元素无冒号，上下元素有冒号且上一个符合规则
            new_condition = (':' not in text and
                             has_next and
                             ':' in next_text
                             and "GW" not in next_text)

            if original_condition or new_condition:
                need_merge = True
                is_new_condition = new_condition

            if need_merge:
                # 从后往前追溯，找到第一个符合条件的元素
                target_idx = None
                for idx in reversed(range(len(merged))):
                    if self.is_valid_prev(merged[idx]):
                        target_idx = idx
                        break

                # 找到符合条件的元素则拼接，否则新增
                if target_idx is not None:
                    target_text = merged[target_idx]
                    if is_new_condition:
                        # 新增条件拼接规则：判断目标元素是否以逗号结尾
                        if target_text.endswith(','):
                            merged[target_idx] += text.lstrip(',')
                        else:
                            merged[target_idx] += f',{text.lstrip(",")}'
                    else:
                        # 原有条件拼接规则：处理逗号重复
                        if target_text.endswith(',') and text.startswith(','):
                            merged[target_idx] += text[1:]
                        else:
                            merged[target_idx] += text
                else:
                    merged.append(text)
            else:
                merged.append(text)

        return merged, attrs

    def process_data(self, original_dict, data_list):
        # 1. 解析父加工说明的信息：(x_parent, y_parent, parent_key, parent_text)
        parent_info = []
        for parent_key, parent_text in original_dict.items():
            try:
                x_str, y_str = parent_key.split('_')
                x_parent = float(x_str.strip())
                y_parent = float(y_str.strip())
                parent_info.append((x_parent, y_parent, parent_key, parent_text))
            except (ValueError, IndexError):
                continue  # 键格式错误则跳过

        # 2. 给每个子条目分配对应的父加工说明
        groups = {}  # key: parent_key, value: list[(sub_x, sub_y, text_str, materials, weight, heat_treatment, pcs)]
        for item in data_list:
            coords, text_str, materials, weight, heat_treatment, pcs = item
            try:
                sub_x = float(coords[0].strip())
                sub_y = float(coords[1].strip())
            except (ValueError, IndexError):
                continue

            # 找到【在子条目上方】的父（y_parent > sub_y）
            candidate_parents = [p for p in parent_info if p[1] > sub_y]
            if not candidate_parents:
                # 无上方父则取所有父中x最近的
                if not parent_info:
                    continue
                closest_parent = min(parent_info, key=lambda p: abs(p[0] - sub_x))
            else:
                # 取上方父中x最近的
                closest_parent = min(candidate_parents, key=lambda p: abs(p[0] - sub_x))

            parent_key = closest_parent[2]
            if parent_key not in groups:
                groups[parent_key] = []
            groups[parent_key].append((sub_x, sub_y, text_str, materials, weight, heat_treatment, pcs))

        # 3. 对每个父的子条目做【列排序】（x分组→列内y降序→列按父x距离排序）
        processed_groups = {}
        column_x_threshold = 200.0  # x细微差异的阈值（可根据图纸调整）
        for parent_key, subs in groups.items():
            if not subs:
                processed_groups[parent_key] = []
                continue

            # 步骤1：按x排序子条目，再按x阈值分组为“列”
            subs_sorted_x = sorted(subs, key=lambda s: s[0])
            columns = []
            current_col = [subs_sorted_x[0]]
            for s in subs_sorted_x[1:]:
                x_diff = abs(s[0] - current_col[0][0])
                if x_diff <= column_x_threshold:
                    current_col.append(s)
                else:
                    columns.append(current_col)
                    current_col = [s]
            columns.append(current_col)  # 加入最后一列

            # 步骤2：列内按y降序排序（从上到下）
            for col in columns:
                col.sort(key=lambda s: s[1], reverse=True)

            # 步骤3：列按“与父x的距离”从小到大排序
            parent_x = next(p[0] for p in parent_info if p[2] == parent_key)
            columns_sorted = sorted(columns, key=lambda col: abs(col[0][0] - parent_x))

            # 步骤4：合并列，转为目标格式
            sorted_subs = []
            for col in columns_sorted:
                sorted_subs.extend([(s[1], s[2], s[3], s[4], s[5], s[6]) for s in col])
            processed_groups[parent_key] = sorted_subs

        # 4. 按父的y坐标降序（从上到下）分配数字键，构建最终结果
        final_result = {}
        parent_info_sorted = sorted(parent_info, key=lambda p: p[1], reverse=True)
        for num_key, p in enumerate(parent_info_sorted, 1):
            parent_key = p[2]
            parent_text = p[3]
            subs = processed_groups.get(parent_key, [])
            merged_texts, attrs = self.merge_texts_and_collect_attrs(subs)
            final_result[str(num_key)] = {
                "文本列表": [parent_text] + merged_texts,
                "属性信息": attrs
            }

        return final_result

    def get_is_bz_ThroughHole(self, text):
        """判断是否支持背面打孔和是否为通孔"""
        # 判断是否为通孔
        if any(keyword in text for keyword in
               ["钻穿", "攻穿", "割", "铰通", "通", "thru", "through", "贯穿", "透", "镗", "塘", "搪"]):
            return True
        # 正钻背攻/背钻正攻/正钻背沉头/背钻正沉头
        if "正" in text and "背" in text:
            return True
        # 判断是否支持背面打孔
        if "深" not in text:
            return True
        return False

    def write_directional_data_to_csv(self, data, output_path):
        try:
            through = data['through_hole_num']
            back = through + data['sink_num']['back_sink']
            correct = through + data['sink_num']['correct_sink']

            headers = ['Z+', 'Z-', 'X+', 'X-', 'Y+', 'Y-']
            if back > correct:
                headers[0], headers[1] = headers[1], headers[0]

            values = {
                'Z+': str(correct),
                'Z-': str(back),
                'X+': '',
                'X-': '',
                'Y+': '',
                'Y-': ''
            }
            row = [values[col] for col in headers]

            # 使用 QUOTE_ALL 强制所有字段加引号
            with open(output_path, mode='w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, quoting=csv.QUOTE_ALL)
                writer.writerow(headers)
                writer.writerow(row)
            print("成功写入CSV")
        except Exception as e:
            print(f"写入CSV时出错：{str(e)}")

    def get_hole_num(self, csv_file_path):
        data_dicts = self.extract_and_process_notes(self.work_part)
        data_dicts["主体加工说明"] = [data_dicts["主体加工说明"]]
        count = {"through_hole_num": 0,
                 "sink_num": {"back_sink": 0, "correct_sink": 0}}
        for key1 in data_dicts:
            for data_dict in data_dicts[key1]:
                for key2 in data_dict:
                    # 1. 跳过不需要处理的key
                    if key2 in ["其它注释", "加工对象", "尺寸", "重量", "数量", "热处理", "材质"]:
                        continue
                    text = str(data_dict[key2])
                    # 跳过含"槽"、"让位"、"导轨面"的加工项
                    if any(keyword in text for keyword in ["槽", "让位", "导轨面"]):
                        continue
                    # 提取数量
                    num = 1
                    idx = text.find("-")
                    if idx != -1 and text[:idx].strip().isdigit():
                        num = int(text[:idx].strip())
                        text = text[idx + 1:].strip()
                    # 判断是否通孔
                    is_through_hole = self.get_is_bz_ThroughHole(text)

                    if "沉头" in text:
                        if "侧" in text:
                            continue
                        # 判断沉头方向（正/背）
                        if "背沉头" in text or "背面沉头" in text or "背面" in text:
                            count["sink_num"]["back_sink"] += num
                            continue

                        if is_through_hole and "背面" in text:
                            count["sink_num"]["back_sink"] += num
                            continue

                        count["sink_num"]["correct_sink"] += num
                    else:
                        if is_through_hole:
                            count["through_hole_num"] += 1

        self.write_directional_data_to_csv(count, csv_file_path)

def open_part(session, part_path):
    """保证返回 NXOpen.Part 类型"""
    try:
        result = session.Parts.Open(part_path)
        # tuple unpack
        if isinstance(result, tuple):
            part = result[0]
        else:
            part = result

        if isinstance(part, NXOpen.Part):
            return part

        raise TypeError("Open() 未返回 NXOpen.Part 对象")
    except Exception as e:
        # 统一异常格式：异常信息 + 位置（文件名、行号、函数名）
        exc_info = traceback.extract_tb(e.__traceback__)[-1]
        err_msg = (f"[open_part] 异常：{str(e)} | "
                   f"异常位置：文件{exc_info.filename}，行{exc_info.lineno}，函数{exc_info.name}")
        print(err_msg)

if __name__ == "__main__":
    session = NXOpen.Session.GetSession()
    work_part = open_part(session, r"D:\weichat_doc\xwechat_files\wxid_wfzh2viv4tkg12_3d41\msg\file\2025-12\LP-04.prt")
    p = ProcessInfoHandler(session, work_part)
    p.get_hole_num("./1.csv")





