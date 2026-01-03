# NX钻孔自动化脚本 - 模块化重构版本

## 项目概述

本项目是对原有NX钻孔自动化脚本进行模块化重构的版本，将原有单一的Python脚本分解为多个逻辑清晰的模块，提升了代码的可读性、可维护性和可扩展性。

## 模块结构

```
NX_Drilling_Automation2/
├── config.py              # 配置模块
├── utils.py               # 工具函数模块
├── geometry.py            # 几何体处理模块
├── path_optimization.py  # 路径优化模块
├── process_info.py        # 加工信息处理模块
├── parameter_parser.py    # 参数解析模块
├── drilling_operations.py # 钻孔操作模块
├── mirror_operations.py   # 镜像操作模块
├── drill_library.py       # 钻刀库模块
├── main_workflow.py      # 主流程模块
├── main.py               # 主程序入口
├── __init__.py           # 包初始化文件
└── README.md            # 说明文档
```

## 功能模块说明

### 1. config.py - 配置模块
- 包含所有常量、路径配置、参数设置
- 统一的配置管理，便于修改和维护

### 2. utils.py - 工具函数模块
- 通用功能：日志输出、数学计算、坐标处理等
- 异常处理统一管理

### 3. geometry.py - 几何体处理模块
- MCS坐标系创建
- 圆识别和分析
- 草图创建和几何体操作

### 4. path_optimization.py - 路径优化模块
- 最近邻贪心算法
- 2-opt优化算法
- 钻孔路径优化

### 5. process_info.py - 加工信息处理模块
- NX注释提取和解析
- 圆分类和标签匹配
- 材料类型识别

### 6. parameter_parser.py - 参数解析模块
- 加工参数解析
- 孔属性提取
- 深度、直径、偏置等参数处理

### 7. drilling_operations.py - 钻孔操作模块
- 刀具创建和管理
- 钻孔工序设置
- 程序组创建

### 8. mirror_operations.py - 镜像操作模块
- 曲线镜像处理
- 边界检测
- 侧面加工支持

### 9. drill_library.py - 钻刀库模块
- 钻孔参数查询
- 材质匹配
- 刀具参数管理

### 10. main_workflow.py - 主流程模块
- 完整工作流程控制
- 各模块协调调用
- 错误处理和日志记录

## 使用说明

### 运行方式

1. **直接运行主程序**：
   ```python
   # 在NX环境中运行
   exec(open("E:/dataset/NC编程/代码结构优化结果/NX_Drilling_Automation2/main.py").read())
   ```

2. **模块化调用**：
   ```python
   from main_workflow import MainWorkflow
   
   session = NXOpen.Session.GetSession()
   work_part = session.Parts.Work
   workflow = MainWorkflow(session, work_part)
   result = workflow.run_workflow()
   ```

### 配置修改

所有配置参数都集中在 `config.py` 文件中，可以根据需要修改：

- `DRILL_JSON_PATH`: 钻孔参数JSON文件路径
- `DEFAULT_MCS_NAME`: 默认MCS坐标系名称
- `DEFAULT_SAFE_DISTANCE`: 安全距离
- 其他钻孔参数和默认值

## 功能特性

### 保持原有功能
- ✅ 完整的圆孔识别和分类
- ✅ 优化的钻孔路径规划
- ✅ 自动创建钻刀和钻孔工序
- ✅ 正面/背面/侧面加工支持
- ✅ 加工参数自动匹配

### 重构改进
- ✅ 模块化设计，逻辑清晰
- ✅ 统一的异常处理机制
- ✅ 配置集中管理
- ✅ 代码可读性大幅提升
- ✅ 易于维护和扩展

## 注意事项

1. **兼容性**：重构版本完全兼容原有功能，接口保持不变
2. **路径配置**：确保 `config.py` 中的文件路径配置正确
3. **NX版本**：支持NX 2312及以上版本
4. **依赖**：需要NX Open API支持

## 开发建议

1. **添加新功能**：在相应的模块中添加，保持模块职责单一
2. **修改配置**：统一在 `config.py` 中修改
3. **异常处理**：使用 `utils.handle_exception()` 统一处理
4. **日志输出**：使用 `utils.print_to_info_window()` 统一输出

## 版本历史

- v2.0.0: 模块化重构版本
- v1.0.0: 原始版本

## 技术支持

如有问题或建议，请联系开发团队。