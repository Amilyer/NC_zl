# -*- coding: utf-8 -*-
"""
几何分析器 (geometry_analyzer.py)
功能：调用 DLL 对 PRT 文件进行几何分析 (加工面方向判断)
"""


from dll_loader import UniversalLoader


class GeometryAnalyzer:
    """几何分析工具类"""

    def __init__(self, dll_path):
        self.dll_path = dll_path
        self.plugin = None
        self._load_plugin()

    def _load_plugin(self):
        try:
            self.plugin = UniversalLoader(self.dll_path)
            # print(f"✅ 几何分析 DLL 加载成功")
        except Exception as e:
            print(f"❌ 几何分析 DLL 加载失败: {e}")
            raise e

    def process_part(self, input_csv_path, output_csv_path, target_layer=0):
        """
        执行几何分析
        
        Args:
            input_csv_path: 输入 CSV 路径
            output_csv_path: 结果 CSV 路径
            target_layer: 目标图层
        """
        if not self.plugin:
            return -999

        try:
            # 动态获取 DLL 定义的参数名
            func_info = self.plugin.functions.get('RunGeometryAnalysis')
            if not func_info:
                print("❌ DLL 中未找到 RunGeometryAnalysis 函数")
                return -997
                
            params = func_info.get('params', [])
            if len(params) < 3:
                print(f"❌ DLL 参数数量不足: {len(params)} (预期至少 3 个)")
                return -996

            # 获取参数名 (不同 DLL 可能不同: input_feature_csv vs input_csv_path)
            param_input_name = params[0]['name']
            param_output_name = params[1]['name']
            param_layer_name = params[2]['name']
            
            # 构造调用参数
            kwargs = {
                param_input_name: input_csv_path,
                param_output_name: output_csv_path,
                param_layer_name: target_layer
            }
            
            # 调用 DLL 接口
            ret = self.plugin.RunGeometryAnalysis(**kwargs)
            return ret
        except Exception as e:
            print(f"   ❌ 几何分析调用异常: {e}")
            return -998
