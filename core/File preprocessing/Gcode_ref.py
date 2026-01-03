# nx_nc_post_lib.py  (NX2312)
# Library-style wrapper: run_nc_post(...) callable from other scripts
# Run with run_journal.exe, or import in other NXOpen scripts.

import os
import re
import time
import traceback

import NXOpen
import NXOpen.CAM

COORD_TOLERANCE = 1e-6


# ----------------------------
# Logging
# ----------------------------
def _log(enable, msg: str):
    if not enable:
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


def _try_remove_file(path, retries=3, delay=0.2):
    for _ in range(retries):
        try:
            if os.path.exists(path):
                os.remove(path)
            return True
        except Exception:
            time.sleep(delay)
    return False


def _dedup_keep_order(items):
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


def _nc_text_has_valid_motion(nc_text, tolerance=COORD_TOLERANCE):
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


def _replace_name_comment(nc_text, new_name):
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
def _group_has_operation(obj, depth=0, max_depth=30):
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
            if _group_has_operation(m, depth + 1, max_depth):
                return True
    return False


def _list_operations_in_group(obj, depth=0, max_depth=30):
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
            ops.extend(_list_operations_in_group(m, depth + 1, max_depth))
    return ops


def _find_group_exact(cam_setup, group_name):
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
def _open_part_as_work(session, part_path, enable_log):
    if not os.path.exists(part_path):
        raise FileNotFoundError(part_path)

    _log(enable_log, f"正在打开部件: {part_path}")

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

    _log(enable_log, "✅ 部件打开成功")
    return session.Parts.Work


def _init_cam_environment(session, enable_log):
    try:
        session.ApplicationSwitchImmediate("UG_APP_MANUFACTURING")
        _log(enable_log, "✅ 已切换到 Manufacturing")
    except Exception as e:
        _log(enable_log, f"ℹ️ 切换 Manufacturing 失败/不可用：{e}")

    try:
        import NXOpen.UF
        uf = NXOpen.UF.UFSession.GetUFSession()
        uf.Cam.InitSession()
        _log(enable_log, "✅ UF CAM InitSession 成功")
    except Exception as e:
        _log(enable_log, f"ℹ️ UF 不可用或 InitSession 失败，跳过：{e}")


def _save_work_part(workPart, enable_log):
    try:
        from NXOpen import BasePart
        workPart.Save(BasePart.SaveComponents.TrueValue, BasePart.CloseAfterSave.FalseValue)
        _log(enable_log, "✅ Work Part 保存完成")
        return True
    except Exception as e:
        _log(enable_log, f"ℹ️ 保存失败/跳过：{e}")
        return False


# ----------------------------
# Postprocess wrapper (NX2312: no .All)
# ----------------------------
def _postprocess_objects(cam_setup, objects, post_name, out_nc_path):
    units = NXOpen.CAM.CAMSetup.OutputUnits.PostDefined
    warn_val = NXOpen.CAM.CAMSetup.PostprocessSettingsOutputWarning.PostDefined
    review_val = NXOpen.CAM.CAMSetup.PostprocessSettingsReviewTool.PostDefined
    mode_val = NXOpen.CAM.CAMSetup.PostprocessSettingsPostMode.Normal

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


def _finalize_nc(tmp_path, final_nc, label):
    try:
        with open(tmp_path, "r", encoding="utf-8", errors="ignore") as f:
            nc_text = f.read()
    except Exception as e:
        return False, f"读取临时 NC 失败: {e}", ""

    if not _nc_text_has_valid_motion(nc_text):
        _try_remove_file(tmp_path)
        return False, "无有效刀路，已删除", ""

    try:
        fixed = _replace_name_comment(nc_text, label)
        tmp_final = final_nc + ".tmp"
        with open(tmp_final, "w", encoding="utf-8", errors="ignore") as f:
            f.write(fixed)

        if os.path.exists(final_nc):
            os.remove(final_nc)
        os.replace(tmp_final, final_nc)

        _try_remove_file(tmp_path)
        return True, "", final_nc
    except Exception as e:
        return False, f"写入最终 NC 失败: {e}", ""


# ======================================================================
# Public API
# ======================================================================
def run_nc_post(
    part_path: str,
    out_root: str = r"E:\Desktop",
    post_name: str = "钢料通用",
    group_names=None,
    fallback_each_op: bool = True,
    enable_log: bool = True,
    save_part: bool = True,
):
    """
    主入口：便于外部调用（import 本文件后直接调用）

    Returns dict:
      {
        "out_dir": str,
        "group_ok": [group_label...],
        "group_fail": [{"group":..., "error":..., "ops":[...]}...],
        "op_ok": [final_nc_path...],
        "op_fail": [{"group":..., "op":..., "error":...}...],
        "files": [all_output_files...]
      }
    """
    if group_names is None:
        group_names = []
    group_names = _dedup_keep_order(list(group_names))

    session = NXOpen.Session.GetSession()

    _log(enable_log, "=== 开始 NC 分组后处理 ===")
    workPart = _open_part_as_work(session, part_path, enable_log)
    _init_cam_environment(session, enable_log)

    cam_setup = workPart.CAMSetup

    out_dir = os.path.join(out_root, workPart.Name)
    os.makedirs(out_dir, exist_ok=True)

    _log(enable_log, f"输出目录: {out_dir}")
    _log(enable_log, f"Post: {post_name}")

    try:
        cam_setup.OutputBallCenter = False
    except Exception:
        pass

    result = {
        "out_dir": out_dir,
        "group_ok": [],
        "group_fail": [],
        "op_ok": [],
        "op_fail": [],
        "files": [],
    }

    for gidx, group_name in enumerate(group_names, start=1):
        _log(enable_log, f"\n—— 处理组: {group_name} ——")

        nc_group = _find_group_exact(cam_setup, group_name)
        if not nc_group:
            _log(enable_log, f"⚠ 未找到组: {group_name}")
            continue

        if not _group_has_operation(nc_group):
            _log(enable_log, f"⚠ 组 {group_name} 下无 Operation（含子组），跳过")
            continue

        label = nc_group.Name
        tmp_path = os.path.join(out_dir, f".tmp_group_{gidx:02d}.nc")  # ASCII tmp
        final_nc = os.path.join(out_dir, f"{label}.nc")

        _try_remove_file(tmp_path)
        _log(enable_log, f"→ 临时后处理(ASCII): {tmp_path}")

        ok, err = _postprocess_objects(cam_setup, [nc_group], post_name, tmp_path)
        if ok:
            ok2, err2, out_file = _finalize_nc(tmp_path, final_nc, label)
            if ok2:
                _log(enable_log, f"✔ 输出完成: {out_file}")
                result["group_ok"].append(label)
                result["files"].append(out_file)
            else:
                _log(enable_log, f"⚠ {err2}")
            continue

        # group failed
        _log(enable_log, f"❌ 组后处理失败: {err}")
        _try_remove_file(tmp_path)

        ops = _list_operations_in_group(nc_group)
        op_names = []
        for op in ops:
            try:
                op_names.append(op.Name)
            except Exception:
                op_names.append(str(op))

        result["group_fail"].append({"group": label, "error": err, "ops": op_names})

        if ops:
            _log(enable_log, "该组包含 Operations：")
            for n in op_names:
                _log(enable_log, f"  - {n}")

        if not fallback_each_op or not ops:
            continue

        _log(enable_log, "↳ 回退：逐个 Operation 单独后处理（能出多少出多少）…")

        for oidx, op in enumerate(ops, start=1):
            try:
                op_name = op.Name
            except Exception:
                op_name = f"OP_{oidx:02d}"

            tmp_op = os.path.join(out_dir, f".tmp_{gidx:02d}_{oidx:02d}.nc")  # ASCII tmp
            final_op = os.path.join(out_dir, f"{label}__{op_name}.nc")

            _try_remove_file(tmp_op)

            ok_op, err_op = _postprocess_objects(cam_setup, [op], post_name, tmp_op)
            if not ok_op:
                result["op_fail"].append({"group": label, "op": op_name, "error": err_op})
                _try_remove_file(tmp_op)
                continue

            ok3, err3, out_file = _finalize_nc(tmp_op, final_op, f"{label}__{op_name}")
            if ok3:
                _log(enable_log, f"  ✔ OP 输出: {out_file}")
                result["op_ok"].append(out_file)
                result["files"].append(out_file)
            else:
                _log(enable_log, f"  ⚠ OP {op_name}: {err3}")

    if save_part:
        _save_work_part(workPart, enable_log)

    _log(enable_log, "=== 全部完成 ===")
    return result


# ======================================================================
# Optional: standalone run
# ======================================================================
if __name__ == "__main__":
    try:
        # 这里给一个默认调用示例：你可以改参数或在别的脚本里 import 调用 run_nc_post
        default_groups = [
            "D1", "D2", "背面打点", "背面钻孔", "背面铣孔", "开粗", "全精", "半精",
            "正面", "正面打点工序", "正面钻孔", "正面铰孔", "正面铣孔",
            "背面", "背面铰孔", "侧面", "侧面打点", "侧面钻孔", "侧面铰孔"
        ]

        run_nc_post(
            part_path=r"E:\Desktop\Final_CAM_PRT\DIE-03.prt",
            out_root=r"E:\Desktop",
            post_name="钢料通用",
            group_names=default_groups,
            fallback_each_op=True,
            enable_log=True,
            save_part=True,
        )
    except Exception as ex:
        _log(True, f"❌ 主程序出错: {ex}")
        _log(True, traceback.format_exc())
