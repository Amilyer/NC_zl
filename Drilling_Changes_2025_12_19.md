# 钻孔功能代码变更汇总 (2025-12-19)

本文档详细记录了今日关于 NX 自动钻孔功能的代码重构与变更，重点包含 **Step 9 (`run_step9.py`)** 的核心代码修改。

## 1. 核心修复：`run_step9.py`

此文件是钻孔流程的入口。为了解决 `ModuleNotFoundError` 和 `FileNotFoundError`，我们对模块加载逻辑进行了彻底重写。

### 1.1 动态加载与路径修复 (Fix Import Errors)

旧版本直接尝试 import，容易因路径问题失败。新版本使用 `importlib` 并强制注入 `sys.path`。

**文件**: `c:\Projects\NC\core\File preprocessing\run_step9.py`

```python
def load_drill_module(pm: PathManager):
    """动态加载 drill_main.py 模块"""
    # 1. 精确定位脚本路径
    # drill_script_path = pm.project_root / "core" / "NX_Drilling_Automation2" / "drill_main.py"
    drill_script_path = pm.project_root / "core" / "NX_Drilling_Automation2" / "drill_main.py"
    
    # 2. 检查文件是否存在
    if not drill_script_path.exists():
        raise FileNotFoundError(f"未找到钻孔脚本: {drill_script_path}")
        
    print(f"🔧 加载钻孔模块: {drill_script_path}")
    
    # 3. [关键] 将父目录添加到 sys.path
    # 这样 drill_main.py 内部 import 同级模块 (如 drilling_operations) 才能成功
    drill_dir = str(drill_script_path.parent)
    if drill_dir not in sys.path:
        sys.path.insert(0, drill_dir)
        
    # 4. 动态加载模块
    spec = importlib.util.spec_from_file_location("drill_main", str(drill_script_path))
    drill_module = importlib.util.module_from_spec(spec)
    sys.modules["drill_main"] = drill_module
    spec.loader.exec_module(drill_module)
    return drill_module
```

### 1.2 标准化调用 (Main Logic)

在主循环中，我们明确了调用参数，并处理了异常。

```python
    # ... (在 run_step9_logic 函数中)

    # 加载模块
    try:
        drill_main = load_drill_module(pm)
        print("✅ 钻孔模块加载成功")
    except Exception as e:
        print(f"❌ 加载钻孔模块失败: {e}")
        return

    # ... (循环文件)
            
            # 调用钻孔逻辑
            print("   > 执行钻孔自动化 (drill_start)...")
            
            drill_main.drill_start(
                session, 
                work_part, 
                drill_json,   # 钻孔参数表路径
                knife_json,   # 刀具参数表路径
                is_save=False # 不原地保存，手动另存为
            )
            
            # 另存为到 output/05_Drilled_PRT
            print(f"   > 另存为: {output_path}")
            if output_path.exists():
                try: output_path.unlink()
                except: pass
                
            work_part.SaveAs(str(output_path))
```

## 2. 流程自动化：`run_all_steps.py`

为了解决内存泄漏和 NX 进程锁定问题，我们创建了全新的启动脚本。

**文件**: `c:\Projects\NC\core\File preprocessing\run_all_steps.py`

此脚本使用 `subprocess` 启动每一步，确保每跑完一步，内存和 NX 对象都会被 OS 强制回收。

```python
def main():
    # ...
    # 顺序执行脚本
    for i, script_path in enumerate(scripts):
        # ...
        try:
            # 使用 subprocess 启动新进程
            result = subprocess.run(
                [sys.executable, str(script_path)],
                cwd=str(current_dir),
                check=True
            )
            print(f"✅ {script_name} 执行成功")
            
        except subprocess.CalledProcessError as e:
            print(f"❌ {script_name} 执行失败 (返回码: {e.returncode})")
            sys.exit(e.returncode)
```

## 3. 路径配置 (Config & PathManager)

确认 `NX_Drilling_Automation2` 文件夹位于 `core` 目录下。

- **期望结构**:
  ```text
  c:\Projects\NC\core\
      ├── File preprocessing\
      │   ├── run_step9.py
      │   └── run_all_steps.py
      └── NX_Drilling_Automation2\  <-- 钻孔模块必须在此
          ├── drill_main.py
          └── drilling_operations.py
  ```

如果再次遇到 `FileNotFoundError`，请首先确认上述目录结构是否完整。

---
**版本**: 2025-12-19
**修改人**: Antigravity Agent
