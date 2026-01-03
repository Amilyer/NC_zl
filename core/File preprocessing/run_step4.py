# -*- coding: utf-8 -*-
"""
步骤 4: PRT 合并与智能处理 (run_step4.py) [单进程重构版]
功能：
1. 读取配对结果
2. 批量合并 3D PRT 和 2D 转换后的 PRT 并进行 AI 预测
3. 导出 Excel 报表
注意：本脚本已完全移除多进程代码，采用最简单的顺序执行方式。
"""

import os
import sys
import shutil
import time
import traceback
import gc

import config
from path_manager import PathManager, init_path_manager

# -----------------------------------------------------------------------------
# 环境配置 (NX)
# -----------------------------------------------------------------------------
# 确保 NXBIN 路径在 sys.path 中，否则找不到 NXOpen
NX_BASE_DIR = r"C:\Program Files\Siemens\NX2312" # 默认或从 config 读取
NX_PYTHON_DIR = os.path.join(NX_BASE_DIR, "NXBIN", "python")
NX_MANAGED_DIR = os.path.join(NX_BASE_DIR, "NXBIN", "managed")

for p in [NX_PYTHON_DIR, NX_MANAGED_DIR]:
    if os.path.exists(p) and p not in sys.path:
        sys.path.append(p)

# -----------------------------------------------------------------------------
# 导入依赖 (确保环境路径正确)
# -----------------------------------------------------------------------------
try:
    from match_manager import MatchManager
    from parts_parameters2excel import prt_to_dict, dict_to_excel
    from nx_processor import NXProcessor
    import NXOpen
    
    # 尝试导入 AI 模块
    try:
        import joblib
        from ai_classifier import AIClassifier
        _AI_AVAILABLE = True
    except ImportError as e:
        print(f"⚠️ AI 模块导入失败: {e}")
        _AI_AVAILABLE = False
        
except ImportError as e:
    print(f"❌ 核心模块导入失败: {e}")
    print("   请检查 PYTHONPATH 或运行环境。")
    sys.exit(1)


def process_single_match(prt_file: str, candidates: list, pm: PathManager, index: int):
    """
    处理单个文件的核心函数
    """
    result = {
        "success": False,
        "message": "",
        "file": prt_file,
        "params": None,
        "label": None
    }
    
    nx = None
    try:
        # 获取路径参数
        split_prt_dir = str(pm.get_split_prt_dir())
        dxf_to_prt_dir = str(pm.get_dxf_prt_dir())
        output_dir = str(pm.get_merged_prt_dir())
        
        mm = MatchManager()
        
        # 1. 选择最佳匹配
        best_match = mm.select_best_match(candidates[0]['prt_dims'], candidates)
        if not best_match:
            result["message"] = "无有效匹配"
            return result

        # 路径构建
        prt_path = os.path.join(split_prt_dir, prt_file)
        prt2_path = os.path.join(dxf_to_prt_dir, best_match['prt2_file'])
        
        # 2. 初始化 NX 和 AI
        nx = NXProcessor()
        
        ai = None
        if _AI_AVAILABLE:
            ai = AIClassifier(pm)
            ai.load_models()

        # 3. NX 操作：打开 3D
        if not nx.open_part(prt_path):
            result["message"] = "无法打开3D文件"
            return result
            
        # --- 2D 文件路径修正逻辑 ---
        prt2_path = os.path.normpath(os.path.abspath(prt2_path))
        if not os.path.exists(prt2_path):
            target_name = os.path.basename(prt2_path)
            found_candidate = None
            try:
                for f in os.listdir(dxf_to_prt_dir):
                    if f.endswith(target_name) or (target_name in f):
                        found_candidate = os.path.join(dxf_to_prt_dir, f)
                        break
            except Exception:
                pass
            
            if found_candidate and os.path.exists(found_candidate):
                prt2_path = found_candidate
            else:
                result["message"] = f"找不到对应的2D文件: {target_name}"
                nx.close_all()
                return result
        # -------------------------

        # 4. NX 操作：导入 2D
        if not nx.import_part(prt2_path):
            nx.close_all()
            result["message"] = "导入2D文件失败"
            return result

        # 5. 准备保存路径
        save_name = f"{os.path.splitext(prt_file)[0]}.prt"
        save_path = os.path.join(output_dir, save_name)
        save_path = os.path.abspath(save_path)
        
        # 6. AI 预测
        label = None
        if ai and ai.is_loaded:
            base_name = os.path.splitext(prt_file)[0]
            label = ai.predict(nx.get_current_part(), base_name)
            result["label"] = label
        
        # 7. 提取参数
        params = {}
        try:
            params = prt_to_dict(
                index,
                nx.get_session(),
                nx.get_current_part(),
                {}, 
                label if label else "未知"
            )
            result["params"] = params
        except Exception as e:
            # print(f"⚠️ 参数提取警告: {e}")
            pass

        # 8. 保存结果
        if nx.save_as(save_path):
            result["success"] = True
            result["message"] = "成功"
        else:
            result["message"] = f"保存失败: {save_path}"
            
        nx.close_all()
        return result

    except Exception as e:
        result["message"] = f"处理异常: {e}"
        traceback.print_exc()
        if nx:
            try: nx.close_all() 
            except: pass
        return result
    finally:
        gc.collect()


def run_processing_loop(pm: PathManager):
    """
    执行主循环
    """
    print("=" * 60)
    print("🚀 步骤 4: PRT 合并与智能处理 (简化单进程版)")
    print("=" * 60)
    
    start_time = time.perf_counter()
    
    # 1. 准备目录
    output_dir = str(pm.get_merged_prt_dir())
    excel_path = str(pm.get_parts_excel())
    
    if os.path.exists(output_dir):
        try: shutil.rmtree(output_dir)
        except: pass
    os.makedirs(output_dir, exist_ok=True)
    
    # 2. 加载匹配数据
    print("🚀 加载匹配数据...")
    mm = MatchManager()
    csv_path = str(pm.get_match_result_csv())
    
    if not os.path.exists(csv_path):
        print(f"❌ 找不到配对结果 CSV: {csv_path}")
        return

    matches = mm.load_matches(csv_path)
    if not matches:
        print("❌ 无匹配数据，流程终止")
        return

    print(f"  待处理数量: {len(matches)}")
    print(f"  输出目录: {output_dir}")
    print("-" * 50)

    # 3. 开始循环
    results = []
    aggregated_params = {}
    completed = 0
    total = len(matches)
    
    match_items = list(matches.items())
    
    for idx, (prt_file, candidates) in enumerate(match_items):
        try:
            res = process_single_match(prt_file, candidates, pm, idx + 1)
            results.append(res)
            
            # 显示进度
            completed += 1
            status_icon = "✅" if res["success"] else "❌"
            label_info = f"| AI: {res['label']}" if res['label'] else ""
            print(f"[{completed}/{total}] {status_icon} {res['file']} {label_info}")
            
            if not res["success"]:
                print(f"    原因: {res['message']}")
                
            if res['params']:
                aggregated_params.update(res['params'])

        except KeyboardInterrupt:
            print("\n⚠️ 用户中断执行")
            break
        except Exception as e:
            print(f"❌ 未知错误: {e}")
        
        # 强制刷新输出
        sys.stdout.flush()

    # 4. 统计与报告
    print("-" * 50)
    success_count = sum(1 for r in results if r["success"])
    print(f"📊 处理完成 | 成功: {success_count} | 失败: {len(results) - success_count}")
    print(f"⏱️ 总耗时: {(time.perf_counter() - start_time):.2f} 秒")
    
    # 5. 生成 Excel
    if aggregated_params:
        try:
            dict_to_excel(aggregated_params, excel_path)
            print(f"✅ Excel 报表已生成: {excel_path}")
        except Exception as e:
            print(f"❌ Excel 生成失败: {e}")


def main():
    # 初始化
    pm = init_path_manager(config.FILE_INPUT_PRT_STR, config.FILE_INPUT_DXF_STR)
    
    # 运行
    run_processing_loop(pm)

if __name__ == "__main__":
    main()
