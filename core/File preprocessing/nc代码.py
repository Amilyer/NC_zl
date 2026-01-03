# NX2312 - Postprocess NC Operations Directly (Skip Group Post)
# Run:
#   D:\NX23\NXBIN\run_journal.exe  E:\Desktop\nc代码(1).py

import os
import re
import time
import traceback

import NXOpen
import NXOpen.CAM

COORD_TOLERANCE = 1e-6


# ----------------------------
# Config
# ----------------------------
class Config:
    PART_PATH = r"E:\Desktop\Final_CAM_PRT\DIE-03.prt"
    OUT_ROOT = r"E:\Desktop"
    POST_NAME = "钢料通用"
    ENABLE_LOG = True

    # 组名（精确匹配）：用于限定要处理的工序所在的组范围
    GROUP_NAMES = [
        "D1", "D2", "背面打点", "背面钻孔", "背面铣孔", "开粗", "全精", "半精",
        "正面", "正面打点工序", "正面钻孔", "正面铰孔", "正面铣孔",
        "背面", "背面打点", "背面钻孔", "背面铰孔", "背面铣孔",
        "侧面", "侧面打点", "侧面钻孔", "侧面铰孔", "侧面背铣孔", "侧面正铣孔"
    ]

    # 移除了无用的FALLBACK_POST_EACH_OPERATION配置


# ----------------------------
# Logging
# ----------------------------
def log(msg: str):
    if not Config.ENABLE_LOG:
        return
    try:
        s = NXOpen.Session.GetSession()
        lw = s.ListingWindow
        lw.Open()
        lw.WriteLine(str(msg))
        try:
            s.LogFile.WriteLine(str(msg))
        except Exception:
            pass
    except Exception:
        print(msg, flush=True)


def try_remove_file(path, retries=3, delay=0.2):
    for _ in range(retries):
        try:
            if os.path.exists(path):
                os.remove(path)
            return True
        except Exception:
            time.sleep(delay)
    return False


def dedup_keep_order(items):
    seen = set()
    out = []
    for x in items:
        if x not in seen:
            out.append(x)
            seen.add(x)
    return out


# ----------------------------
# NC validity check
# ----------------------------
def _is_tool_name_nonempty(nc_text):
    m = re.search(r'\(Tool name\s*:\s*(.*?)\)', nc_text, re.IGNORECASE | re.DOTALL)
    return bool(m and m.group(1).strip())


def nc_text_has_valid_motion(nc_text, tolerance=COORD_TOLERANCE):
    if not nc_text or not nc_text.strip():
        return False

    meaningful = []
    for L in nc_text.splitlines():
        s = L.strip()
        if not s:
            continue
        if s.startswith('%') or s.startswith(';'):
            continue
        if s.startswith('(') and 'tool name' not in s.lower():
            continue
        meaningful.append(s)

    text = "\n".join(meaningful)

    if re.search(r'\bG0?1\b|\bG0?2\b|\bG0?3\b', text, re.IGNORECASE):
        return True
    if re.search(r'\bF\s*[-+]?\d*\.?\d+', text, re.IGNORECASE):
        return True

    for m in re.finditer(r'[XYZ]\s*([-+]?\d*\.?\d+)', text, re.IGNORECASE):
        try:
            if abs(float(m.group(1))) > tolerance:
                return True
        except Exception:
            return True

    if re.search(r'\bM06\b|\bT\d+\b', text, re.IGNORECASE):
        return _is_tool_name_nonempty(nc_text)

    return False


def replace_name_comment(nc_text, new_name):
    pattern = re.compile(r'\(\s*NAME\s*[:=]?\s*[^)]*\)', re.IGNORECASE)
    repl = f"(NAME: {new_name})"

    if pattern.search(nc_text):
        return pattern.sub(repl, nc_text, count=1)

    lines = nc_text.splitlines(True)
    if lines and lines[0].strip() == "%":
        lines.insert(1, repl + "\n")
    else:
        lines.insert(0, repl + "\n")
    return "".join(lines)


# ----------------------------
# CAM helpers
# ----------------------------
def group_has_operation(obj, depth=0, max_depth=30):
    if depth > max_depth:
        return False
    try:
        members = obj.GetMembers()
    except Exception:
        return False

    for m in members:
        if isinstance(m, NXOpen.CAM.Operation):
            return True
        if hasattr(m, "GetMembers"):
            if group_has_operation(m, depth + 1, max_depth):
                return True
    return False


def list_operations_in_group(obj, depth=0, max_depth=30):
    ops = []
    if depth > max_depth:
        return ops
    try:
        members = obj.GetMembers()
    except Exception:
        return ops

    for m in members:
        if isinstance(m, NXOpen.CAM.Operation):
            ops.append(m)
        elif hasattr(m, "GetMembers"):
            ops.extend(list_operations_in_group(m, depth + 1, max_depth))
    return ops


def find_group_exact(cam_setup, group_name):
    try:
        g = cam_setup.CAMGroupCollection.FindObject(group_name)
        if g:
            try:
                if getattr(g, "Name", "").strip() == group_name:
                    return g
            except Exception:
                return g
    except Exception:
        pass

    try:
        groups = list(cam_setup.CAMGroupCollection)
    except Exception:
        groups = []

    for g in groups:
        try:
            name = getattr(g, "Name", None)
            if name and str(name).strip() == group_name:
                return g
        except Exception:
            pass
    return None


# ----------------------------
# NX open / init
# ----------------------------
def open_part_as_work(session, part_path):
    if not os.path.exists(part_path):
        raise FileNotFoundError(part_path)

    log(f"正在打开部件: {part_path}")

    part = None
    load_status = None
    try:
        part, load_status = session.Parts.OpenBaseDisplay(part_path)
    except Exception:
        part, load_status = session.Parts.OpenDisplay(part_path)

    try:
        session.Parts.SetWork(part)
    except Exception:
        pass

    try:
        if load_status:
            load_status.Dispose()
    except Exception:
        pass

    if session.Parts.Work is None:
        raise RuntimeError("部件打开后 Work Part 仍为空")

    log("✅ 部件打开成功")
    return session.Parts.Work


def init_cam_environment(session):
    try:
        session.ApplicationSwitchImmediate("UG_APP_MANUFACTURING")
        log("✅ 已切换到 Manufacturing")
    except Exception as e:
        log(f"ℹ️ 切换 Manufacturing 失败/不可用：{e}")

    try:
        import NXOpen.UF
        uf = NXOpen.UF.UFSession.GetUFSession()
        uf.Cam.InitSession()
        log("✅ UF CAM InitSession 成功")
    except Exception as e:
        log(f"ℹ️ UF 不可用或 InitSession 失败，跳过：{e}")


def save_work_part(workPart):
    try:
        from NXOpen import BasePart
        workPart.Save(BasePart.SaveComponents.TrueValue, BasePart.CloseAfterSave.FalseValue)
        log("✅ Work Part 保存完成")
        return True
    except Exception as e:
        log(f"ℹ️ 保存失败/跳过：{e}")
        return False


# ----------------------------
# Postprocess wrapper (NO .All in NX2312)
# ----------------------------
def postprocess_objects(cam_setup, objects, post_name, out_nc_path):
    """
    objects: [CAMObject]  (group or [operation])
    返回 (ok:bool, err:str)
    """
    units = NXOpen.CAM.CAMSetup.OutputUnits.PostDefined

    warn_enum = NXOpen.CAM.CAMSetup.PostprocessSettingsOutputWarning
    review_enum = NXOpen.CAM.CAMSetup.PostprocessSettingsReviewTool
    mode_enum = NXOpen.CAM.CAMSetup.PostprocessSettingsPostMode

    # NX2312 常见：只有 PostDefined（没有 All）
    warn_val = warn_enum.PostDefined
    review_val = review_enum.PostDefined
    mode_val = mode_enum.Normal

    try:
        cam_setup.PostprocessWithPostModeSetting(
            objects,
            post_name,
            out_nc_path,
            units,
            warn_val,
            review_val,
            mode_val
        )
        return True, ""
    except Exception as e:
        return False, str(e)


# ----------------------------
# NC write helper
# ----------------------------
def finalize_nc(tmp_path, final_nc, label):
    try:
        with open(tmp_path, "r", encoding="utf-8", errors="ignore") as f:
            nc_text = f.read()
    except Exception as e:
        return False, f"读取临时 NC 失败: {e}（保留临时文件: {tmp_path}）"

    if not nc_text_has_valid_motion(nc_text):
        try_remove_file(tmp_path)
        return False, "无有效刀路，已删除"

    try:
        fixed = replace_name_comment(nc_text, label)
        tmp_final = final_nc + ".tmp"
        with open(tmp_final, "w", encoding="utf-8", errors="ignore") as f:
            f.write(fixed)

        if os.path.exists(final_nc):
            os.remove(final_nc)
        os.replace(tmp_final, final_nc)

        try_remove_file(tmp_path)
        return True, ""
    except Exception as e:
        return False, f"写入最终 NC 失败: {e}（保留临时文件: {tmp_path}）"


# ----------------------------
# Main
# ----------------------------
def main():
    session = NXOpen.Session.GetSession()
    log("=== 开始 NC 工序直接后处理 ===")

    workPart = open_part_as_work(session, Config.PART_PATH)
    init_cam_environment(session)

    try:
        cam_setup = workPart.CAMSetup
    except Exception as e:
        log(f"❌ 该部件没有 CAMSetup / 未启用 CAM：{e}")
        return

    out_dir = os.path.join(Config.OUT_ROOT, workPart.Name)
    os.makedirs(out_dir, exist_ok=True)

    log(f"输出目录: {out_dir}")
    log(f"Post: {Config.POST_NAME}")

    try:
        cam_setup.OutputBallCenter = False
    except Exception:
        pass

    group_names = dedup_keep_order(Config.GROUP_NAMES)
    # 用于收集所有找到的工序，避免重复处理（如果多个组包含相同工序）
    all_ops = []
    processed_op_names = set()

    # 第一步：收集所有指定组内的工序
    for group_name in group_names:
        log(f"\n—— 扫描组: {group_name} ——")

        nc_group = find_group_exact(cam_setup, group_name)
        if not nc_group:
            log(f"⚠ 未找到组: {group_name}")
            continue

        if not group_has_operation(nc_group):
            log(f"⚠ 组 {group_name} 下无 Operation（含子组），跳过")
            continue

        # 获取组内所有工序
        ops = list_operations_in_group(nc_group)
        if not ops:
            log(f"⚠ 组 {group_name} 下未找到任何工序")
            continue

        # 去重添加工序（按名称去重）
        for op in ops:
            try:
                op_name = op.Name
                if op_name not in processed_op_names:
                    processed_op_names.add(op_name)
                    all_ops.append((group_name, op))
                    log(f"  ✅ 找到工序: {op_name}")
            except Exception:
                log(f"  ⚠ 无法获取工序名称，跳过该工序")

    if not all_ops:
        log("❌ 未找到任何需要处理的工序")
        save_work_part(workPart)
        return

    log(f"\n=== 共找到 {len(all_ops)} 个待处理工序，开始逐个后处理 ===")

    # 第二步：逐个处理工序
    bad_ops = []
    for idx, (group_name, op) in enumerate(all_ops, start=1):
        try:
            op_name = op.Name
        except Exception:
            op_name = f"OP_{idx:03d}"

        log(f"\n—— 处理工序 {idx}/{len(all_ops)}: {op_name}（所属组：{group_name}）——")

        # 临时文件路径
        tmp_op = os.path.join(out_dir, f".tmp_op_{idx:03d}.nc")
        # 最终文件路径：组名_工序名.nc
        final_op = os.path.join(out_dir, f"{group_name}__{op_name}.nc")
        try_remove_file(tmp_op)

        # 后处理工序
        ok_op, err_op = postprocess_objects(cam_setup, [op], Config.POST_NAME, tmp_op)
        if not ok_op:
            bad_ops.append((op_name, err_op))
            try_remove_file(tmp_op)
            log(f"❌ 工序 {op_name} 后处理失败: {err_op}")
            continue

        # 处理NC文件并保存
        ok3, err3 = finalize_nc(tmp_op, final_op, f"{group_name}__{op_name}")
        if ok3:
            log(f"✔ 工序 {op_name} 输出完成: {final_op}")
        else:
            log(f"⚠ 工序 {op_name} 处理失败: {err3}")
            bad_ops.append((op_name, err3))

    # 输出失败统计
    if bad_ops:
        log(f"\n❌ 共有 {len(bad_ops)} 个工序处理失败：")
        for name, e in bad_ops:
            log(f"  - {name} :: {e}")
    else:
        log(f"\n✅ 所有工序处理成功！")

    # 保存部件
    save_work_part(workPart)
    log("\n=== 工序后处理全部完成 ===")


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        log(f"❌ 主程序出错: {ex}")
        log(traceback.format_exc())