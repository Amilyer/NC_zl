'''
2025.12.22, 10:15 第一版
黄雪
'''
import NXOpen
import csv
import os
import re
import json
import openpyxl
from pathlib import Path
import pandas as pd


def print_to_info_window(message):
    theSession = NXOpen.Session.GetSession()
    lw = theSession.ListingWindow
    lw.Open()
    lw.WriteLine(str(message))


def load_knife_table(json_path):
    if not os.path.exists(json_path):
        return []
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def read_csv_as_pandas(file_path):
    if not os.path.exists(file_path):
        return pd.DataFrame()
    try:
        return pd.read_csv(file_path, encoding='utf-8-sig')
    except:
        try:
            return pd.read_csv(file_path, encoding='gbk')
        except:
            return pd.DataFrame()


def get_tool(knife_data, tool_name):
    if not knife_data or not tool_name:
        return None
    for item in knife_data:
        if item.get("刀具名称") == tool_name:
            return item
    return None


def get_face_normals_dict(file_path):
    normals_map = {}
    df = read_csv_as_pandas(file_path)
    if df.empty:
        return normals_map
    key_tag = 'Face Tag'
    key_norm = 'Face Normal'
    df = df.dropna(subset=[key_tag, key_norm])
    df['tag_int'] = df[key_tag].astype(int)
    for _, row in df.iterrows():
        tag_id = row['tag_int']
        n_str = str(row[key_norm])
        parts = n_str.replace('"', '').split(',')
        if len(parts) >= 3:
            nx = float(parts[0].strip())
            ny = float(parts[1].strip())
            nz = float(parts[2].strip())
            normals_map[tag_id] = [nx, ny, nz]
    return normals_map


def get_grouped_face_tags(machining_path, path_main):
    df = read_csv_as_pandas(machining_path)
    if df.empty:
        return {}
    target_directions = ["+Z", "-Z", "+X", "-X", "+Y", "-Y"]
    final_groups = {}
    main_direction = get_main_direction(path_main)
    for col in df.columns:
        clean_col = col.strip()
        if clean_col in target_directions:
            tags = df[col].dropna().astype(float).astype(int).tolist()
            if tags:
                final_groups[clean_col] = tags
    return final_groups, main_direction


def get_main_direction(path_main):
    main_dir = '+Z'
    df = read_csv_as_pandas(path_main)
    if df.empty:
        return main_dir
    data_dict0 = df.to_dict('list')
    data_dict = dict()
    for k, v in data_dict0.items():
        if not pd.isna(v[0]):
            data_dict[k[1] + k[0]] = int(v[0])
    k_list = []
    for k, v in data_dict.items():
        if v:
            k_list.append(k)
    if not k_list:
        return main_dir
    k_list = sorted(k_list, key=lambda k: data_dict[k], reverse=True)
    k = k_list[0]
    if '+Z' in data_dict and data_dict['+Z'] == data_dict[k]: return '+Z'
    return k


def get_direction_map(faces_path):
    faces_data = read_csv_to_list(faces_path)
    normal_map = get_face_normals_dict(faces_path)
    direction_map = {'+X': [], '-X': [], '+Y': [], '-Y': [], '+Z': [], '-Z': []}
    for face in faces_data:
        if normal_map[int(face['Face Tag'])][0] == 1.0:
            direction_map['+X'].append(int(face['Face Tag']))
        if normal_map[int(face['Face Tag'])][0] == -1.0:
            direction_map['-X'].append(int(face['Face Tag']))
        if normal_map[int(face['Face Tag'])][1] == 1.0:
            direction_map['+Y'].append(int(face['Face Tag']))
        if normal_map[int(face['Face Tag'])][1] == -1.0:
            direction_map['-Y'].append(int(face['Face Tag']))
        if normal_map[int(face['Face Tag'])][2] == 1.0:
            direction_map['+Z'].append(int(face['Face Tag']))
        if normal_map[int(face['Face Tag'])][2] == -1.0:
            direction_map['-Z'].append(int(face['Face Tag']))
    return direction_map


def filter_sidewalls_by_normal(grouped_faces, normals_map):
    filtered_groups = {}
    for direction, tags in grouped_faces.items():
        valid_tags = []
        excluded_count = 0
        check_idx = -1
        d_upper = direction.upper()
        if "Z" in d_upper:
            check_idx = 2
        elif "X" in d_upper:
            check_idx = 0
        elif "Y" in d_upper:
            check_idx = 1
        for t in tags:
            if t not in normals_map or check_idx == -1:
                valid_tags.append(t)
                continue
            n_vec = normals_map[t]
            val = n_vec[check_idx]
            if abs(val) > 0.001:
                valid_tags.append(t)
            else:
                excluded_count += 1
        if valid_tags:
            filtered_groups[direction] = valid_tags
    return filtered_groups


def get_material_only(work_part):
    if work_part is None:
        return "45#"
    size_pattern = r'(\d+(?:\.\d+)?)L\D*?(\d+(?:\.\d+)?)W\D*?(\d+(?:\.\d+)?)T'
    material_db = ["45#", "A3", "CR12", "CR12MOV", "SKD11", "SKH-9", "DC53", "P20",
                   "T00L0X33", "T00L0X44", "合金铜", "TOOLOX33", "TOOLOX44"]
    for note in work_part.Notes:
        text = " ".join(note.GetText()).strip()
        if re.search(size_pattern, text, re.IGNORECASE):
            found = None
            max_l = 0
            for m in material_db:
                if m.lower() in text.lower() and len(m) > max_l:
                    found = m
                    max_l = len(m)
            if found:
                return found
    return "45#"


def get_material_by_prt_name(prt_name, excel_path):
    df = pd.read_excel(excel_path, engine='openpyxl')
    df = df[df['文件名称'] == prt_name]
    return df['材质'].tolist()[0]


def get_is_hot(prt_name, excel_path):
    df = pd.read_excel(excel_path, engine='openpyxl')
    df = df[df['文件名称'] == prt_name]
    if df['热处理'].tolist()[0]:
        return True
    else:
        return False


def read_csv_to_list(file_path):
    data = []
    if not os.path.exists(file_path):
        return []
    try:
        with open(file_path, mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                data.append(row)
    except UnicodeDecodeError:
        with open(file_path, mode='r', encoding='gbk') as f:
            reader = csv.DictReader(f)
            for row in reader:
                data.append(row)
    return data if data else []


def get_face_color_map(faces_path):
    color_map = {}
    raw_faces = read_csv_to_list(faces_path)
    if not raw_faces:
        return color_map
    k_tag = 'Face Tag'
    k_color = 'Face Color'
    for raw in raw_faces:
        tag_str = raw.get(k_tag, '0')
        color_str = raw.get(k_color, '0')
        if tag_str and color_str:
            tid = int(tag_str)
            color_val = int(color_str)
            color_map[tid] = color_val
    return color_map


def get_pos(feature_faces):
    map = dict()
    for f in feature_faces:
        map[int(f['Face Tag'])] = [float(p.strip()) for p in f['Face Point'].split(',')]
    return map


def get_sides(feature_faces):
    map = dict()
    for f in feature_faces:
        map[int(f['Face Tag'])] = [int(p.strip()) for p in f['Adjacent Face Tags'].split(';')]
    return map


def get_is_top(faces_path):
    R = dict()
    faces_data = read_csv_to_list(faces_path)
    normal_map = get_face_normals_dict(faces_path)
    direction_map = {'+X': [], '-X': [], '+Y': [], '-Y': [], '+Z': [], '-Z': []}
    pos_map = get_pos(faces_data)
    for face in faces_data:
        if normal_map[int(face['Face Tag'])][0] == 1.0:
            direction_map['+X'].append(face)
        if normal_map[int(face['Face Tag'])][0] == -1.0:
            direction_map['-X'].append(face)
        if normal_map[int(face['Face Tag'])][1] == 1.0:
            direction_map['+Y'].append(face)
        if normal_map[int(face['Face Tag'])][1] == -1.0:
            direction_map['-Y'].append(face)
        if normal_map[int(face['Face Tag'])][2] == 1.0:
            direction_map['+Z'].append(face)
        if normal_map[int(face['Face Tag'])][2] == -1.0:
            direction_map['-Z'].append(face)
        else:
            R[int(face['Face Tag'])] = False
    for direction in ['+X', '-X', '+Y', '-Y', '+Z', '-Z']:
        index = 0
        match direction[1]:
            case 'X':
                index = 0
            case 'Y':
                index = 1
            case 'Z':
                index = 2
        p_n = direction[0]
        group = direction_map[direction]
        all_pos = [pos_map[int(f['Face Tag'])][index] for f in group if int(f['Face Type']) == 22]
        if not all_pos:
            continue
        if p_n == '+':
            limit = max(all_pos)
        else:
            limit = min(all_pos)
        for face in group:
            if abs(pos_map[int(face['Face Tag'])][index] - limit) <= 0.01 and int(face['Face Type']) == 22:
                R[int(face['Face Tag'])] = True
            else:
                R[int(face['Face Tag'])] = False
    return R


def get_layer20_volume_diff(workPart):
    theSession = NXOpen.Session.GetSession()
    layer_20_bodies = []
    for body in workPart.Bodies:
        if body.Layer == 20:
            layer_20_bodies.append(body)
    if len(layer_20_bodies) != 2:
        return -1.0
    results_a = theSession.Measurement.GetBodyProperties([layer_20_bodies[0]], 0.999, False)
    vol_a = results_a[0][1]
    results_b = theSession.Measurement.GetBodyProperties([layer_20_bodies[1]], 0.999, False)
    vol_b = results_b[0][1]
    return abs(vol_a - vol_b)


def get_pos_diff_direction(workPart, direction):
    theSession = NXOpen.Session.GetSession()
    layer_20_bodies = []
    for body in workPart.Bodies:
        if body.Layer == 20:
            layer_20_bodies.append(body)
    if len(layer_20_bodies) != 2:
        return -1.0
    body = layer_20_bodies[1]
    faces = list(body.GetFaces())
    poss = []
    for face in faces:
        point = theSession.Measurement.GetFaceProperties([face], 0.99, NXOpen.Measurement.AlternateFace.Radius, True)[3]
        if 'X' in direction:
            poss.append(point.X)
        if 'Y' in direction:
            poss.append(point.Y)
        if 'Z' in direction:
            poss.append(point.Z)
    return max(poss) if '+' in direction else min(poss)


def get_volume_diff_direction(workPart, direction, faces_data):
    pos = get_pos_diff_direction(workPart, direction)
    sum = 0
    for face in faces_data:
        sum += abs(pos - face['pos']) * face['area_direction']
    return sum


def calculate_tool_sequence(slot_path, faces_path, face_ids, direction, work_part):
    current_ids = set([int(id) for id in face_ids])
    slot_map = get_slot_map(slot_path)
    r_map = get_R_map(slot_path)
    raw_slots = read_csv_to_list(slot_path)
    raw_faces = read_csv_to_list(faces_path)
    color_map = get_face_color_map(faces_path)
    if not (raw_faces and current_ids):
        return []
    top_map = get_is_top(faces_path)
    print(top_map)
    direction_map = get_direction_map(faces_path)
    normal_map = get_face_normals_dict(faces_path)
    full_tools = [63, 50, 32, 17, 10, 8, 6]
    tool_map = {63: "63R6", 50: "50R0.8", 32: "32R0.8", 17: "17R0.8", 10: "D10",
                8: "D8", 6: "D6", 5: "D5", 4: "D4"}
    tool_next_map = {63: 32, 50: 32, 32: 17, 17: 10, 10: 4,
                     8: 4, 6: 0}
    faces_data = []
    target_axis = 'Z'
    axis_index = 2
    if 'X' in direction:
        target_axis = 'X'
        axis_index = 0
    elif 'Y' in direction:
        target_axis = 'Y'
        axis_index = 1
    k_tag = 'Face Tag'
    k_len = 'Long'
    k_wid = 'Width'
    k_hgt = 'Height'
    k_area = 'Area'
    k_min_edge = 'Min Edge Length'
    k_pos = 'Face Point'
    k_type = 'Face Type'
    k_min_rad = 'Min Radius Of Curvature'
    for raw in raw_faces:
        tid = int(raw[k_tag])
        if tid in current_ids:
            l = raw[k_len]
            w = raw[k_wid]
            h = raw[k_hgt]
            area = float(raw[k_area])
            min_edge_val = float(raw[k_min_edge])
            face_type = int(raw[k_type])
            min_rad_val = float(raw[k_min_rad])
            val_raw = raw[k_pos]
            pos_parts = [p.strip() for p in str(val_raw).replace('"', '').split(',')]
            pos_val = float(pos_parts[axis_index]) if len(pos_parts) >= 3 else 0.0
            if target_axis == 'X':
                dim1, dim2 = float(w), float(h)
            elif target_axis == 'Y':
                dim1, dim2 = float(l), float(h)
            else:
                dim1, dim2 = float(l), float(w)
            if face_type == 22:
                solidity = 0.0
                if dim1 > 0 and dim2 > 0:
                    solidity = area / (dim1 * dim2)
                if solidity > 0.8:
                    final_d = min(dim1, dim2)
                else:
                    if min_edge_val > 0.001:
                        final_d = min_edge_val
                    else:
                        final_d = min(dim1, dim2)
            else:
                if min_rad_val > 0.001:
                    final_d = min_rad_val * 2.0
                else:
                    final_d = min(dim1, dim2)
            if not color_map[tid] == 186 or color_map[tid] == 150:
                faces_data.append({
                    'tag': tid,
                    'd': final_d,
                    'area': area,
                    'pos': pos_val,
                    'type': face_type,
                    'angle': abs(normal_map[tid][axis_index]),
                    'area_direction': area * abs(normal_map[tid][axis_index])
                })
    e = 0
    for f in faces_data:
        if (not top_map[f['tag']]) and color_map[f['tag']] != 186 and color_map[f['tag']] != 150:
            e = 1
    if e == 0:
        return [], faces_data
    faces_bottom = [f for f in faces_data if f['type'] == 22 and f['tag'] in direction_map[direction]]

    if len(faces_bottom) == len(faces_data):
        pos = faces_bottom[0]['pos']
        e = 0
        for f in faces_bottom:
            if abs(f['pos'] - pos) >= 0.01:
                e = 1
        if e == 0:
            return [], faces_data
    s1, s2, s3 = 4000000, 903000, 400000
    volume = get_volume_diff_direction(work_part, direction, faces_data)

    if volume <= s1:
        full_tools.pop(0)
    if volume <= s2:
        full_tools.pop(0)
    if volume <= s3 and len([f for f in faces_data if f['type'] == 22 and f['tag'] in direction_map[direction]]) >= 2:
        full_tools.pop(0)

    slots = [f for f in faces_data if f['tag'] in slot_map and slot_map[f['tag']]]
    v1 = get_volume_diff_direction(work_part, direction, slots)
    v2 = get_volume_diff_direction(work_part, direction, faces_data)

    if v1 > v2 * 0.4:
        while sum([f['area'] for f in slots if f['d'] >= full_tools[0]]) < sum([f['area'] for f in slots]) * 0.2:
            if full_tools[0] == 10: break
            full_tools.pop(0)

    seq = []
    seq.append(full_tools[0])
    if full_tools[0] == 63:
        need_32=False
        for f in slots:
            if 32 <= f['d'] < 63:
                need_32 = True
                break
        if need_32:
            seq.append(32)
        need_17 = False

        for f in slots:
            if 17 <= f['d'] < 32:
                need_17 = True
                break
        if need_17:
            seq.append(17)


    if full_tools[0] == 50:
        need_17 = False
        for f in slots:
            if 17 <= f['d'] < 50:
                need_17 = True
                break
        if need_17:
            seq.append(17)

    really_slot_map = get_really_slot_map(slot_path, raw_faces, direction)
    really_pockets = [f for f in faces_data if f['tag'] in really_slot_map and really_slot_map[f['tag']]]
    if len(really_pockets) == 0:
        return [tool_map.get(t, str(t)) for t in seq], faces_data
    pocket_sizes = [f['d'] for f in really_pockets]
    min_pocket = min(pocket_sizes)
    cands = [t for t in full_tools if t <= min_pocket]
    seq.append(max(cands) if cands else 6)
    seq = sorted(list(set(seq)), reverse=True)
    rs = [f for f in faces_data if f['tag'] in r_map and r_map[f['tag']]]
    if len(rs) == 0:
        return [tool_map.get(t, str(t)) for t in seq], faces_data
    r_sizes = [f['d'] for f in rs]
    min_r = min(r_sizes)
    cands = [t for t in full_tools if t <= min_r]
    seq.append(max(cands) if cands else 6)
    seq = set(seq)
    if 6 in seq:
        if 8 in seq:
            seq.remove(8)
        if 10 in seq:
            seq.remove(10)
    if 8 in seq:
        if 10 in seq:
            seq.remove(10)
    seq = sorted(list(seq), reverse=True)
    return [tool_map.get(t, str(t)) for t in seq], faces_data


def get_R_map(slot_path):
    slot_data = read_csv_to_list(slot_path)
    R = dict()
    for f in slot_data:
        if (('POCKET' in f['Type'] and f['Type'] != 'STEP1POCKET') or f['Type'] == 'SLOT_PARTIAL_U_SHAPED' or f[
            'Type'] == 'SLOT_PARTIAL_RECTANGULAR' or f['Type'] == 'SLOT_RECTANGULAR') and f[
            'FaceType'] == 'Cylindrical':
            R[int(f['Attribute'])] = True
        else:
            R[int(f['Attribute'])] = False
    return R


def get_slot_map(slot_path):
    slot_data = read_csv_to_list(slot_path)
    R = dict()
    for f in slot_data:
        if (('POCKET' in f['Type'] and f['Type'] != 'STEP1POCKET') or f['Type'] == 'SLOT_PARTIAL_U_SHAPED' or f[
            'Type'] == 'SLOT_PARTIAL_RECTANGULAR' or f['Type'] == 'SLOT_RECTANGULAR') and f[
            'IsBottom'] == 'Yes':
            R[int(f['Attribute'])] = True
        else:
            R[int(f['Attribute'])] = False
    return R


def get_pocket_map(slot_path):
    slot_data = read_csv_to_list(slot_path)
    R = dict()
    for f in slot_data:
        if (('POCKET' in f['Type'] and f['Type'] != 'STEP1POCKET') or f['Type'] == 'SLOT_PARTIAL_U_SHAPED' or f[
            'Type'] == 'SLOT_PARTIAL_RECTANGULAR' or f['Type'] == 'SLOT_RECTANGULAR'):
            R[int(f['Attribute'])] = True
        else:
            R[int(f['Attribute'])] = False
    return R


def get_really_slot_map(slot_path, feature_faces, direction):
    slot_data = read_csv_to_list(slot_path)
    pos_map = get_pos(feature_faces)
    sides_map = get_sides(feature_faces)
    axis_index = 2
    if 'X' in direction:
        axis_index = 0
    elif 'Y' in direction:
        axis_index = 1
    R = dict()
    for f in slot_data:
        if (('POCKET' in f['Type'] and f['Type'] != 'STEP1POCKET') or f['Type'] == 'SLOT_PARTIAL_U_SHAPED' or f[
            'Type'] == 'SLOT_PARTIAL_RECTANGULAR' or f['Type'] == 'SLOT_RECTANGULAR') and f[
            'IsBottom'] == 'Yes':
            fid = int(f['Attribute'])
            pos_it = pos_map[fid][axis_index]
            if '-' in direction:
                pos_it = -pos_it
            sum = 0
            for side in sides_map[fid]:
                pos_side = pos_map[side][axis_index]
                if '-' in direction:
                    pos_side = -pos_side
                if pos_it >= pos_side:
                    sum += 1
            if sum >= 3:
                R[int(f['Attribute'])] = True
            else:
                R[int(f['Attribute'])] = False
    return R


def get_step_by_material(tool_rec, material, is_hot):
    material = str(material)
    is_hot = False
    if not tool_rec or not material:
        return "0.5"
    m = material.upper().replace(" ", "")
    if not is_hot:
        if any(k in m for k in ["45#", "A3"]):
            key = "45#,A3,切深"
        elif "CR12" in m:
            key = "CR12热处理前切深"
        elif any(k in m for k in ["CR12MOV", "SKD11", "SKH-9", "DC53"]):
            key = "CR12mov,SKD11,SKH-9,DC53,热处理前切深"
        elif "P20" in m:
            key = "P20切深"
        elif any(k in m for k in ["TOOLOX33", "TOOLOX44"]):
            key = "TOOLOX33 TOOLOX44切深"
        elif "合金铜" in m:
            key = "合金铜切深"
        else:
            key = "45#,A3,切深"
        val = tool_rec[key]
        return str(val) if str(val).replace(".", "").isdigit() else "0.5"
    else:
        if any(k in m for k in ["CR12"]):
            key = "CR12热处理后切深"
        elif any(k in m for k in ["CR12mov", "SKD11", "SKH-9", "DC53"]):
            key = "CR12mov,SKD11,SKH-9,DC53,热处理后切深"
        val = tool_rec[key]
        return str(val) if str(val).replace(".", "").isdigit() else "0.5"


def generate_cavity_json_v2(part_file, slot_csv, face_csv, mach_csv, knife_json, out_json, path_prts, path_main):
    pocket_map = get_pocket_map(slot_csv)
    theSession = NXOpen.Session.GetSession()
    lw = theSession.ListingWindow
    lw.SelectDevice(NXOpen.ListingWindow.DeviceType.Window, "")
    lw.Open()
    base_part, load_status = theSession.Parts.OpenBaseDisplay(part_file)
    load_status.Dispose()
    direction_map = get_direction_map(face_csv)
    knife_data = load_knife_table(knife_json)
    face_color_map = get_face_color_map(face_csv)
    groups, top_but_not_top = get_grouped_face_tags(mach_csv, path_main)
    print('top_but_not_top', top_but_not_top, len(top_but_not_top))
    top_map = get_is_top(face_csv)
    face_normal_map = get_face_normals_dict(face_csv)
    output_data_dict = {}
    process_count = 1
    tool_map = {"63R6": 63, "50R0.8": 50, "32R0.8": 32, "17R0.8": 17, "D10": 10,
                "D8": 8, "D6": 6, "D5": 5, "D4": 4}

    filename_with_ext = os.path.basename(part_file)  # 获取 "DIE-20_modified.prt"
    pattern = r"([a-zA-Z]+-\d+)"
    match = re.search(pattern, filename_with_ext)
    if match:
        prt_name = match.group(1)  # 提取匹配到的部分，如 "DIE-20"
    else:
        # 如果文件名不符合规则（比如没有连字符），则退回到仅去后缀
        prt_name = os.path.splitext(filename_with_ext)[0]
    print('处理文件', prt_name)
    material_name = get_material_by_prt_name(prt_name, path_prts)
    print("材质：", material_name)
    for direction, fids in groups.items():
        print('正在处理', direction, '方向...')
        if not fids:
            continue
        seq, faces_data = calculate_tool_sequence(slot_csv, face_csv, fids, direction, base_part)
        if seq:
            ids_normal = []
            ids_yellow = []
            for fid in fids:
                fid = int(fid)

                if top_map[fid] and not fid in direction_map[top_but_not_top]: continue
                color = face_color_map[fid]
                if color == 150 or color == 186: continue
                if color == 6:
                    ids_yellow.append(fid)
                else:
                    ids_normal.append(fid)
            ref_tool = "NONE"
            layer_id = 0
            if '+Z' in direction:
                layer_id = 20
            elif '-Z' in direction:
                layer_id = 70
            elif '+X' in direction:
                layer_id = 40
            elif '-X' in direction:
                layer_id = 30
            elif '+Y' in direction:
                layer_id = 60
            elif '-Y' in direction:
                layer_id = 50

            for tool_name in seq:
                tool = get_tool(knife_data, tool_name)
                step_val = get_step_by_material(tool, material_name, False)
                step_str = str(step_val) if step_val else "0"
                if tool:
                    tool = {
                        "转速": float(tool.get("转速(普)")),
                        "进给": float(tool.get("进给(普)")),
                        "横越": float(tool.get("横越(普)")),
                    }
                if not tool:
                    tool = {
                        "转速": 2000.0,
                        "进给": 800.0,
                        "横越": 8000.0,
                    }
                inner_data_b = {
                    "工序": "行腔_SIMPLE",
                    "刀具名称": tool_name,
                    "加工方向": direction,
                    "指定图层": layer_id,
                    "切深": step_str,
                    "参考刀具": (
                        "D"+str(tool_map[ref_tool]) if 'R0.8' in ref_tool else ref_tool) if ref_tool != "NONE" else "NONE",
                    "最终余量": 0.8,
                    "是否为槽": False,
                    "普通面ID列表": ids_normal,
                    "黄色面ID列表": ids_yellow,
                    **tool
                }
                wrapper_key = "行腔{}".format(process_count)
                if inner_data_b['普通面ID列表'] or inner_data_b['黄色面ID列表']:
                    output_data_dict[wrapper_key] = inner_data_b
                    ref_tool = tool_name
                    process_count += 1
    try:
        with open(out_json, 'w', encoding='utf-8') as f:
            json.dump(output_data_dict, f, ensure_ascii=False, indent=4)
        print('生成成功')
    except:
        print('生成失败')


def main():
    specific_prt_path = r"C:\Projects\NC\output\04_PRT_with_Tool\GU-04.prt"
    path_slots = r"C:\Projects\NC\output\03_Analysis\Navigator_Reports\csv\GU-04_FeatureRecognition_Log.csv"
    path_all_faces = r"C:\Projects\NC\output\03_Analysis\Face_Info\face_csv\GU-04_face_data.csv"
    path_machining = r"C:\Projects\NC\output\03_Analysis\Geometry_Analysis\GU-04.csv"
    path_knife_json = r"C:\Projects\NC\铣刀参数.json"
    path_prts = r"C:\Projects\NC\output\00_Resources\CSV_Reports\零件参数.xlsx"
    path_output_json = r"C:\Projects\NC\行腔4.json"
    path_main = r"C:\Projects\NC\output\03_Analysis\Counterbore_Info\GU-04.csv"
    generate_cavity_json_v2(specific_prt_path, path_slots, path_all_faces,
                            path_machining, path_knife_json,
                            path_output_json, path_prts, path_main)


if __name__ == '__main__':
    main()
