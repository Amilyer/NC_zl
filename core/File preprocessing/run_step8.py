# -*- coding: utf-8 -*-
"""
步骤 8: 综合处理 (run_step8.py)
功能：
1. 复制 PRT 文件到输出目录
2. 调用 workpiece_module 创建包容体和 MCS (当前已屏蔽)
3. 调用 tool_module 创建刀具
** 多进程版 **
"""
import os
import sys
import shutil
import glob
import time
import traceback
import importlib.util
from concurrent.futures import ProcessPoolExecutor, as_completed

# 1. 确保路径设置
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

try:
    import config
    from path_manager import init_path_manager, PathManager
except ImportError:
    pass

# 定义项目根目录
try:
    PROJECT_ROOT = config.PROJECT_ROOT_STR
except:
    PROJECT_ROOT = "C:/Projects/NC"

def import_module_from_path(module_name, file_path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

def ensure_cam_setup_worker(session, work_part):
    """Worker 内部使用的 CAM 环境设置函数"""
    try:
        module_name = session.ApplicationName
        if module_name != "UG_APP_MANUFACTURING":
            session.ApplicationSwitchImmediate("UG_APP_MANUFACTURING")
        
        if not session.IsCamSessionInitialized():
            session.CreateCamSession()
            
        cam_setup_ready = False
        try:
            if work_part.CAMSetup is not None:
                cam_setup_ready = True
        except:
            pass

        if not cam_setup_ready:
            try:
                work_part.CreateCamSetup("hole_making")
            except:
                return False
        return True
    except:
        return False

def process_single_prt_worker(input_path, output_path, excel_path, project_root):
    """
    Step 8 worker function
    """
    import sys
    import os
    import importlib.util
    import traceback
    
    # 重新添加路径
    if project_root not in sys.path:
        sys.path.insert(0, project_root)
    core_dir = os.path.join(project_root, "core", "File preprocessing")
    if core_dir not in sys.path:
        sys.path.insert(0, core_dir)

    import NXOpen

    # 动态导入模块 (在进程内)
    workpiece_module = None
    tool_module = None

    # 导入 tool_module
    tool_script = os.path.join(core_dir, "创建刀具.py")
    try:
        spec = importlib.util.spec_from_file_location("tool_module", tool_script)
        tool_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(tool_module)
    except Exception as e:
        return False, f"无法导入 tool_module: {e}"

    # 原先在此导入包容体创建脚本，但用户已要求移除原始开粗脚本的自动调用
    workpiece_module = None
    rotated_workpiece_module = None

    # 仅引入用于旋转分层构建包容体的脚本（如果可用）
    try:
        rotated_script = os.path.join(core_dir, "旋转分层建包容体workpiece.py")
        spec3 = importlib.util.spec_from_file_location("rotated_workpiece_module", rotated_script)
        rotated_workpiece_module = importlib.util.module_from_spec(spec3)
        spec3.loader.exec_module(rotated_workpiece_module)
    except Exception:
        rotated_workpiece_module = None

    file_name = os.path.basename(input_path)
    prefix = f"[{os.getpid()}] {file_name}"
    
    # print(f"{prefix}: 开始处理...")

    # 1. 复制文件
    try:
        shutil.copy2(input_path, output_path)
    except Exception as e:
        return False, f"{prefix}: 复制文件失败: {e}"

    session = None
    work_part = None

    try:
        # 针对旋转包容体，先读取几何分析 CSV 决定是否需要执行旋转构建
        if rotated_workpiece_module:
            try:
                part_name = os.path.splitext(os.path.basename(output_path))[0]
                try:
                    needed_dirs = rotated_workpiece_module.read_machining_directions_from_csv(part_name)
                except Exception:
                    needed_dirs = None

                # 如果 CSV 不存在或无法解析，保守调用模块以保证行为一致；
                # 否则仅在 CSV 标记了至少一个方向（包括 +Z 原始方向 或 其他旋转方向）时调用。
                call_module = False
                if needed_dirs is None:
                    call_module = True
                else:
                    if len(needed_dirs) > 0:
                        call_module = True

                if call_module:
                    try:
                        print(f"{prefix}: 正在调用 rotated_workpiece_module (旋转分层建包容体)... 需要的方向: {needed_dirs}")
                        success_r = rotated_workpiece_module.process_file_auto(output_path, output_path)
                        if success_r:
                            print(f"{prefix}: rotated_workpiece_module 已完成")
                        else:
                            print(f"{prefix}: rotated_workpiece_module 返回 False（继续执行刀具创建）")
                    except Exception as e:
                        print(f"{prefix}: rotated_workpiece_module 调用异常: {e}")
                else:
                    print(f"{prefix}: CSV 指示不需要任何方向的旋转，跳过 rotated_workpiece_module")
            except Exception as e:
                print(f"{prefix}: rotated_workpiece_module 处理出错: {e}")
        else:
            print(f"{prefix}: 未加载 rotated_workpiece_module，跳过旋转包容体创建")

        # 现在打开刚复制到输出目录的部件以进行刀具创建
        session = NXOpen.Session.GetSession()
        base_part, _ = session.Parts.OpenBaseDisplay(output_path)
        work_part = session.Parts.Work

        # 3. 切换到加工环境
        if not ensure_cam_setup_worker(session, work_part):
            # 尝试关闭
            try:
                work_part.Close(NXOpen.BasePart.CloseWholeTree.TrueValue, NXOpen.BasePart.CloseModified.UseResponses, None)
            except: pass
            return False, f"{prefix}: 无法初始化 CAM 环境"

        # # 包容体创建已被用户禁用，跳过此步
        # print(f"{prefix}: 跳过包容体创建（按用户要求）")

        # 5. 刀具创建
        # print(f"{prefix}: 执行刀具创建...")
        if tool_module:
            try:
                # 调用 process_part 接口
                success = tool_module.process_part(work_part, excel_path)
                if not success:
                    # 刀具创建失败通常不应阻断所有流程，但按要求返回 False
                     return False, f"{prefix}: 刀具创建部分失败"
            except Exception as e:
                return False, f"{prefix}: 刀具创建抛出异常: {e}"
        else:
            return False, f"{prefix}: tool_module 未加载"

        # 6. 保存并关闭
        work_part.Save(NXOpen.BasePart.SaveComponents.TrueValue, NXOpen.BasePart.CloseAfterSave.FalseValue)
        work_part.Close(NXOpen.BasePart.CloseWholeTree.TrueValue, NXOpen.BasePart.CloseModified.CloseModified, None)
        
        return True, f"{prefix}: 处理成功"

    except Exception as e:
        err_msg = f"{prefix}: 异常: {e}"
        try:
            if work_part:
                work_part.Close(NXOpen.BasePart.CloseWholeTree.TrueValue, NXOpen.BasePart.CloseModified.CloseModified, None)
        except:
            pass
        return False, err_msg
    finally:
        import gc
        gc.collect()

def main():
    print("=" * 60)
    print("🚀 步骤 8: 综合处理 (包容体[屏蔽] + 刀具) - 多进程版")
    print("=" * 60)

    pm = init_path_manager(config.FILE_INPUT_PRT, config.FILE_INPUT_DXF)
    
    input_dir = pm.get_analysis_face_prt_dir()
    output_dir = pm.get_final_prt_dir() 
    excel_path = pm.get_mill_tools_excel()

    print(f"📂 输入目录: {input_dir}")
    print(f"📂 输出目录: {output_dir}")
    print(f"📄 Excel路径: {excel_path}")
    print(f"⚙️  并行进程数: {config.PROCESS_MAX_WORKERS}")
    print("=" * 60)

    if not os.path.exists(input_dir):
        print(f"❌ 输入目录不存在: {input_dir}")
        return

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    else:
        # 清理输出目录
        try:
            for f in os.listdir(output_dir):
                fp = os.path.join(output_dir, f)
                if os.path.isfile(fp): os.unlink(fp)
        except: pass

    prt_files = glob.glob(os.path.join(input_dir, "*.prt"))
    if not prt_files:
        print("❌ 没有找到 PRT 文件")
        return

    total_files = len(prt_files)
    print(f"📂 发现 {total_files} 个 PRT 文件，准备开始...")

    success_count = 0
    failed_count = 0
    
    with ProcessPoolExecutor(max_workers=config.PROCESS_MAX_WORKERS) as executor:
        future_to_file = {
            executor.submit(process_single_prt_worker, f, os.path.join(output_dir, os.path.basename(f)), excel_path, PROJECT_ROOT): f
            for f in prt_files
        }
        
        print("\n🚀 正在并行处理任务...")
        for i, future in enumerate(as_completed(future_to_file)):
            original_file = future_to_file[future]
            fname = os.path.basename(original_file)
            
            try:
                success, msg = future.result()
                if success:
                    success_count += 1
                    print(f"[{i+1}/{total_files}] ✅ {fname}")
                else:
                    failed_count += 1
                    print(f"[{i+1}/{total_files}] ❌ {msg}")
            except Exception as e:
                failed_count += 1
                print(f"[{i+1}/{total_files}] ❌ {fname}: {e}")

    print("\n" + "=" * 60)
    print(f"🎉 全部完成!")
    print(f"✅ 成功: {success_count}")
    print(f"❌ 失败: {failed_count}")
    print("=" * 60)

if __name__ == '__main__':
    main()
