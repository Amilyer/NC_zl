# nc_file_manager.py
# File management and utility functions for NC code generation
# Extracted from Gcode.py reference

import os
import re
import time

try:
    import NXOpen
except ImportError:
    NXOpen = None

COORD_TOLERANCE = 1e-6

# ----------------------------
# Logging
# ----------------------------
def log(msg: str, enable: bool = True):
    if not enable:
        return
    
    # Try logging to NX listing window if available
    try:
        if NXOpen:
            s = NXOpen.Session.GetSession()
            lw = s.ListingWindow
            lw.Open()
            lw.WriteLine(str(msg))
            try:
                s.LogFile.WriteLine(str(msg))
            except Exception:
                pass
            return
    except Exception:
        pass
    
    # Fallback to standard print
    print(msg, flush=True)


def try_remove_file(path, retries=3, delay=0.2):
    for _ in range(retries):
        try:
            if os.path.exists(path):
                os.remove(path)
            # Check if gone
            if not os.path.exists(path):
                return True
        except Exception:
            time.sleep(delay)
    return not os.path.exists(path)


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
    # Look for (Tool name : ...)
    m = re.search(r'\(Tool name\s*:\s*(.*?)\)', nc_text, re.IGNORECASE | re.DOTALL)
    # If match found, check if group 1 is not empty/space
    return bool(m and m.group(1).strip())


def nc_text_has_valid_motion(nc_text, tolerance=COORD_TOLERANCE):
    if not nc_text or not nc_text.strip():
        return False

    meaningful = []
    for L in nc_text.splitlines():
        s = L.strip()
        if not s:
            continue
        # Skip standard comments or program headers that don't imply motion
        if s.startswith('%') or s.startswith(';'):
            continue
        # Ignore comments except potentially tool name ones (handled elsewhere mostly, but here we just filter generic lines)
        if s.startswith('(') and 'tool name' not in s.lower():
            continue
        meaningful.append(s)

    text = "\n".join(meaningful)

    # 1. G01/G02/G03
    if re.search(r'\bG0?1\b|\bG0?2\b|\bG0?3\b', text, re.IGNORECASE):
        return True
    
    # 2. Feed rate
    if re.search(r'\bF\s*[-+]?\d*\.?\d+', text, re.IGNORECASE):
        return True

    # 3. Coordinate movement > tolerance
    for m in re.finditer(r'[XYZ]\s*([-+]?\d*\.?\d+)', text, re.IGNORECASE):
        try:
            if abs(float(m.group(1))) > tolerance:
                return True
        except Exception:
            return True

    # 4. Tool change (valid only if tool name present)
    if re.search(r'\bM06\b|\bT\d+\b', text, re.IGNORECASE):
        return _is_tool_name_nonempty(nc_text)

    return False


def replace_name_comment(nc_text, new_name):
    # Replace (NAME : ...) or insert it
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


def finalize_nc(tmp_path, final_nc, label):
    """
    Reads tmp_path, validates it, updates header, moves to final_nc.
    Returns (success: bool, error_msg: str, output_path: str)
    """
    try:
        with open(tmp_path, "r", encoding="utf-8", errors="ignore") as f:
            nc_text = f.read()
    except Exception as e:
        return False, f"读取临时 NC 失败: {e}", ""

    if not nc_text_has_valid_motion(nc_text):
        try_remove_file(tmp_path)
        return False, "无有效刀路，已删除", ""

    try:
        fixed = replace_name_comment(nc_text, label)
        # Write to a .tmp file alongside final destination to avoid partial writes
        tmp_final = final_nc + ".tmp"
        with open(tmp_final, "w", encoding="utf-8", errors="ignore") as f:
            f.write(fixed)

        if os.path.exists(final_nc):
            os.remove(final_nc)
        
        os.replace(tmp_final, final_nc)

        # Cleanup original temp
        try_remove_file(tmp_path)
        
        return True, "", final_nc
    except Exception as e:
        return False, f"写入最终 NC 失败: {e}", ""
