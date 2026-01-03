# NX2312 - NC Code Generation (Refactored)
# Implements main logic for orchestrating NX CAM post-processing
# Uses nc_file_manager for file handling

import os
import traceback
import NXOpen
import NXOpen.CAM
import nc_file_manager as ncfm
import config

# ----------------------------
# Config Definitions
# ----------------------------
DEFAULT_POST_NAME = config.NC_POST_NAME
FALLBACK_POST_EACH_OPERATION = True

# Groups to process
GROUP_NAMES = config.NC_GROUP_NAMES

# ----------------------------
# CAM Helpers
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
# NX Environment
# ----------------------------
def init_cam_environment(session):
    try:
        session.ApplicationSwitchImmediate("UG_APP_MANUFACTURING")
    except Exception as e:
        ncfm.log(f"ℹ️ 切换 Manufacturing 失败/不可用：{e}")

    try:
        import NXOpen.UF
        uf = NXOpen.UF.UFSession.GetUFSession()
        uf.Cam.InitSession()
    except Exception as e:
        ncfm.log(f"ℹ️ UF 不可用或 InitSession 失败，跳过：{e}")

# ----------------------------
# Post Processing Logic
# ----------------------------
def postprocess_objects(cam_setup, objects, post_name, out_nc_path):
    """
    Executes the NX Postprocess command.
    Returns (success: bool, error_msg: str)
    """
    units = NXOpen.CAM.CAMSetup.OutputUnits.PostDefined
    warn_enum = NXOpen.CAM.CAMSetup.PostprocessSettingsOutputWarning
    review_enum = NXOpen.CAM.CAMSetup.PostprocessSettingsReviewTool
    mode_enum = NXOpen.CAM.CAMSetup.PostprocessSettingsPostMode

    # NX2312: Use PostDefined
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


def process_work_part(work_part, out_dir, post_name=DEFAULT_POST_NAME, group_names=None):
    if group_names is None:
        group_names = GROUP_NAMES
    
    group_names = ncfm.dedup_keep_order(group_names)
    session = NXOpen.Session.GetSession()
    
    init_cam_environment(session)

    try:
        cam_setup = work_part.CAMSetup
    except Exception as e:
        ncfm.log(f"❌ 该部件没有 CAMSetup / 未启用 CAM：{e}")
        return

    # Ensure output dir exists
    if not os.path.exists(out_dir):
        os.makedirs(out_dir)

    ncfm.log(f"输出目录: {out_dir}")
    ncfm.log(f"Post: {post_name}")

    try:
        cam_setup.OutputBallCenter = False
    except Exception:
        pass

    for gidx, group_name in enumerate(group_names, start=1):
        ncfm.log(f"\n—— 处理组: {group_name} ——")

        nc_group = find_group_exact(cam_setup, group_name)
        if not nc_group:
            ncfm.log(f"⚠ 未找到组: {group_name}")
            continue

        if not group_has_operation(nc_group):
            ncfm.log(f"⚠ 组 {group_name} 下无 Operation（含子组），跳过")
            continue

        label = nc_group.Name
        tmp_path = os.path.join(out_dir, f".tmp_group_{gidx:02d}.nc")
        final_nc = os.path.join(out_dir, f"{label}.nc")
        
        ncfm.try_remove_file(tmp_path)
        ncfm.log(f"→ 临时后处理(ASCII): {tmp_path}")

        ok, err = postprocess_objects(cam_setup, [nc_group], post_name, tmp_path)
        
        if ok:
            ok2, err2, out_file = ncfm.finalize_nc(tmp_path, final_nc, label)
            if ok2:
                ncfm.log(f"✔ 输出完成: {out_file}")
            else:
                ncfm.log(f"⚠ {err2}")
            continue

        # If group post failed
        ncfm.log(f"❌ 组后处理失败: {err}")
        ncfm.try_remove_file(tmp_path)

        ops = list_operations_in_group(nc_group)
        if ops:
            ncfm.log("该组包含 Operations：")
            for op in ops:
                try:
                    ncfm.log(f"  - {op.Name}")
                except Exception:
                    ncfm.log(f"  - {op}")

        if not FALLBACK_POST_EACH_OPERATION or not ops:
            continue

        ncfm.log("↳ 回退：逐个 Operation 单独后处理（能出多少出多少）…")

        bad_ops = []
        for oidx, op in enumerate(ops, start=1):
            try:
                op_name = op.Name
            except Exception:
                op_name = f"OP_{oidx:02d}"

            tmp_op = os.path.join(out_dir, f".tmp_{gidx:02d}_{oidx:02d}.nc")
            final_op = os.path.join(out_dir, f"{label}__{op_name}.nc")
            
            ncfm.try_remove_file(tmp_op)

            ok_op, err_op = postprocess_objects(cam_setup, [op], post_name, tmp_op)
            if not ok_op:
                bad_ops.append((op_name, err_op))
                ncfm.try_remove_file(tmp_op)
                continue

            ok3, err3, out_file = ncfm.finalize_nc(tmp_op, final_op, f"{label}__{op_name}")
            if ok3:
                ncfm.log(f"  ✔ OP 输出: {out_file}")
            else:
                ncfm.log(f"  ⚠ OP {op_name}: {err3}")

        if bad_ops:
            ncfm.log("⚠ 以下 Operations 仍失败：")
            for name, e in bad_ops:
                ncfm.log(f"  - {name} :: {e}")

# ----------------------------
# Main Entry Point (Called by run_step16.py)
# ----------------------------
def main(output_dir):
    """
    Main entry point invoked by run_step16.py.
    Operates on the currently open Work Part.
    """
    ncfm.log("=== 开始 NC 代码生成 (nc_processor.py) ===")
    
    session = NXOpen.Session.GetSession()
    work_part = session.Parts.Work
    
    if work_part is None:
        ncfm.log("❌ 错误：当前没有活动的 Work Part")
        return

    process_work_part(work_part, output_dir)
    ncfm.log("=== NC 代码生成结束 ===")

if __name__ == "__main__":
    # For standalone testing within NX Journal
    try:
        # Check if we have arguments or just use a default test path
        # In a real journal run, we might rely on the currently open part
        s = NXOpen.Session.GetSession() 
        if s.Parts.Work:
            # Default output to desktop or temp if no arg provided
            out = r"E:\Desktop\NC_Output_Test"  # Example default
            main(out)
        else:
            ncfm.log("请先打开一个部件文件 (.prt)")
    except Exception as ex:
        ncfm.log(f"❌ 主程序出错: {ex}")
        ncfm.log(traceback.format_exc())
