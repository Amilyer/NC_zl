# -*- coding: utf-8 -*-
"""
Step 9: Automatic Drilling (run_step9.py)
功能：
1. 读取 Step 8 (或 Step 7) 生成的 PRT (位于 output/04_PRT_with_Tool)
2. 调用 core/NX_Drilling_Automation2/drill_main.py 进行自动化钻孔
3. 另存为到 output/05_Drilled_PRT
"""
import os
import sys
import shutil
import traceback
import importlib.util
from pathlib import Path

# 添加项目根目录到 sys.path
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

import config
from path_manager import init_path_manager, PathManager

try:
    import NXOpen
    from nx_processor import NXProcessor
except ImportError:
    pass

def load_drill_module(pm: PathManager):
    """动态加载 drill_main.py 模块"""
    # 路径: core/NX_Drilling_Automation2/drill_main.py
    drill_script_path = pm.project_root / "core" / "NX_Drilling_Automation2" / "drill_main.py"
    
    if not drill_script_path.exists():
        raise FileNotFoundError(f"未找到钻孔脚本: {drill_script_path}")
        
    print(f"🔧 加载钻孔模块: {drill_script_path}")
    
    # 将包含 drill_main.py 的目录添加到 sys.path，以便它能导入同目录下的其他模块 (utils, main_workflow 等)
    drill_dir = str(drill_script_path.parent)
    if drill_dir not in sys.path:
        sys.path.insert(0, drill_dir)
        
    spec = importlib.util.spec_from_file_location("drill_main", str(drill_script_path))
    drill_module = importlib.util.module_from_spec(spec)
    sys.modules["drill_main"] = drill_module
    spec.loader.exec_module(drill_module)
    return drill_module

def run_step9_logic(pm: PathManager):
    print("=" * 60)
    print("🚀 Step 9: 自动化钻孔流程")
    print("=" * 60)

    # 1. 路径配置
    input_dir = pm.get_step8_prt_dir() # Input: output/04_PRT_with_Tool
    output_dir = pm.get_step9_drilled_dir() # Output: output/05_Drilled_PRT
    
    # 钻孔配置文件
    drill_json = str(pm.get_drill_table_json())
    knife_json = str(pm.get_knife_table_json())
    
    print(f"📂 输入目录: {input_dir}")
    print(f"📂 输出目录: {output_dir}")
    print(f"📄 Drill JSON: {drill_json}")
    print(f"📄 Knife JSON: {knife_json}")
    
    if not input_dir.exists():
        print(f"❌ 输入目录不存在: {input_dir}")
        return

    # 清理输出目录
    if output_dir.exists():
        try: shutil.rmtree(output_dir)
        except: pass
    output_dir.mkdir(parents=True, exist_ok=True)

    # 2. 加载钻孔模块
    try:
        drill_main = load_drill_module(pm)
        print("✅ 钻孔模块加载成功")
    except Exception as e:
        print(f"❌ 加载钻孔模块失败: {e}")
        traceback.print_exc()
        return

    # 3. 启动 NX 会话
    session = NXOpen.Session.GetSession()
    nx_proc = NXProcessor() # 用于管理打开/关闭

    prt_files = list(input_dir.glob("*.prt"))
    total = len(prt_files)
    success_count = 0
    
    print(f"📂 发现 {total} 个 PRT 文件")

    for idx, prt_file in enumerate(prt_files):
        filename = prt_file.name
        output_path = output_dir / filename
        prefix = f"[{idx+1}/{total}] {filename}"
        
        print(f"\n{prefix} 处理中...")
        
        try:
            # 打开文件
            if not nx_proc.open_part(str(prt_file)):
                print(f"  ❌ 无法打开文件: {prt_file}")
                continue
                
            work_part = session.Parts.Work
            
            # 调用钻孔逻辑
            print("   > 执行钻孔自动化 (drill_start)...")
            # 签名: drill_start(session, work_part, drill_path, knfie_path, is_save=False)
            drill_main.drill_start(
                session, 
                work_part, 
                filename,
                drill_json, 
                knife_json, 
                is_save=False # 不原地保存，手动另存为
            )
            
            # 另存为
            print(f"   > 另存为: {output_path}")
            # 确保目标文件不存在
            if output_path.exists():
                try: output_path.unlink()
                except: pass
                
            work_part.SaveAs(str(output_path))
            
            success_count += 1
            nx_proc.close_all()
            
        except Exception as e:
            print(f"  ❌ 处理异常: {e}")
            traceback.print_exc()
            try: nx_proc.close_all()
            except: pass
            
        # 强制垃圾回收
        import gc
        gc.collect()

    print("-" * 50)
    print(f"🎉 步骤 9 完成. 成功: {success_count}/{total}")

def main():
    pm = init_path_manager(config.FILE_INPUT_PRT, config.FILE_INPUT_DXF)
    run_step9_logic(pm)

if __name__ == "__main__":
    main()
