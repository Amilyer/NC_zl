---
description: 创建新开发步骤的标准工作流规则
---

# 创建新工序步骤 (Create New Step) 标准流程

当用户请求创建一个新的处理步骤（如 `run_stepX.py`）时，**必须**严格遵循以下规范。

## 1. 注册路径与配置 (Infrastructure)

在编写脚本之前，首先确保基础设施就绪：

1.  **修改 `core/File preprocessing/path_manager.py`**:
    *   **注册输出目录**: 添加 `get_step{N}_{name}_dir(self)` 方法。
    *   **规范**: 使用 `self._get_dir(self.work_dir / 'XX_FolderName')` 确保目录自动创建。
    *   **禁止**: 绝对不要在脚本中硬编码路径 `C:\...`。

2.  **修改 `core/File preprocessing/config.py`**:
    *   如果有特定的常量（图层、颜色、公差），在此文件中定义，例如 `LAYER_STEP{N}_TARGET = 20`。

## 2. 脚本结构规范 (Script Structure)

新脚本 `run_step{N}.py` 必须遵循 **"初始化 -> 循环 -> 处理 -> 清理"** 的模式。

### A. 头部声明
```python
# -*- coding: utf-8 -*-
"""
步骤 {N}: {功能名称} (run_step{N}.py)
功能：
1. {功能点1}
2. {功能点2}
"""
import os, sys, time
from path_manager import init_path_manager
# 导入功能模块 (Function Module)
```

### B. 单文件处理函数 (`process_single_file`)
**逻辑要求**:
1.  **打开文件**: 使用 NX Open 或其他库打开目标 PRT。
2.  **匹配配置**: 根据文件名找到对应的配置参数（如需）。
3.  **调用功能函数**: 调用核心逻辑模块，传入对应参数（如需）。
4.  **关闭文件**: 无论成功失败，必须在 `finally` 块或通过逻辑保证关闭文件，释放内存。
5.  **返回结果**: 返回字典 `{'success': Bool, 'msg': Str}`。

```python
def process_single_file(file_path, pm, config_data):
    try:
        # 1. Open
        # 2. Match Config
        # 3. Call Function
        # 4. Save/Export
        return {"success": True, "message": "成功"}
    except Exception as e:
        return {"success": False, "message": str(e)}
    finally:
        # 5. Close File (Critical!)
        pass
```

### C. 主循环函数 (`run_processing_loop`)
**逻辑要求**:
1.  **清理输出目录**: 脚本启动时，先清空对应的 Output 目录，确保无脏数据。
2.  **简化的日志**: 不要在循环内打印大量调试信息。使用标准格式：
    *   `[M/N] ✅ Filename.prt | 附加信息`
    *   `[M/N] ❌ Filename.prt | 错误信息`
3.  **异常捕获**: 确保单个文件崩溃不影响整体循环。

## 3. 示例模板 (Template)

请参考以下结构编写代码：

```python
def run_processing_loop(pm):
    print("=" * 60)
    print(f"🚀 步骤 {N}: {Title}")
    print("=" * 60)
    
    # 1. 准备目录
    output_dir = pm.get_step{N}_output_dir()
    if os.path.exists(output_dir): shutil.rmtree(output_dir)
    os.makedirs(output_dir)
    
    # 2. 获取输入文件列表
    # ...
    
    # 3. 循环处理
    total = len(files)
    for i, file in enumerate(files):
        # 调用处理函数
        res = process_single_file(file, pm)
        
        # 简化的日志提醒
        status = "✅" if res['success'] else "❌"
        print(f"[{i+1}/{total}] {status} {os.path.basename(file)} {res['message']}")
```

## 4. 检查清单

- [ ] `path_manager.py` 已更新？
- [ ] `process_single_file` 包含打开/关闭/异常处理闭环？
- [ ] 日志输出是否简洁清晰（✅/❌）？
- [ ] 是否清理了旧的输出数据？