# -*- coding: utf-8 -*-
import ctypes
import json
import os
import sys

# ============================================================================
# 🔧 通用加载器 (复用自 加工面方向判断.py)
# ============================================================================
class UniversalLoader:
    def __init__(self, dll_path):
        if not os.path.exists(dll_path):
            raise FileNotFoundError(f"DLL not found: {dll_path}")
        try:
            # 尝试加载 DLL
            self.dll = ctypes.CDLL(dll_path)
        except OSError as e:
            raise OSError(f"加载 DLL 失败 ({dll_path}): {e}")
            
        self.functions = {}
        self._load_metadata()

    def _load_metadata(self):
        try:
            self.dll.get_func_count.restype = ctypes.c_int
            count = self.dll.get_func_count()
        except AttributeError:
            raise Exception("DLL 不支持通用接口 (缺少 get_func_count)，请检查 DLL 版本")

        self.dll.get_func_info.argtypes = [ctypes.c_int, ctypes.c_char_p, ctypes.c_int]
        buf = ctypes.create_string_buffer(4096)

        for i in range(count):
            if self.dll.get_func_info(i, buf, 4096) == 0:
                info = json.loads(buf.value.decode())
                self._register_func(info)

    def _register_func(self, info):
        func_name = info['name']
        if not hasattr(self.dll, func_name): return
        c_func = getattr(self.dll, func_name)
        self.functions[func_name] = info

        argtypes = []
        for p in info['params']:
            # 4 代表字符串(char*), 其他视为 int
            if p['type'] == 4:
                argtypes.append(ctypes.c_char_p)
            else:
                argtypes.append(ctypes.c_int)
        c_func.argtypes = argtypes
        c_func.restype = ctypes.c_int

    def __getattr__(self, name):
        if name not in self.functions:
            raise AttributeError(f"DLL 中未找到函数: {name} (可用: {list(self.functions.keys())})")

        def wrapper(**kwargs):
            return self._invoke(name, kwargs)

        return wrapper

    def _invoke(self, name, kwargs):
        info = self.functions[name]
        c_func = getattr(self.dll, name)
        args = []
        for p in info['params']:
            pname = p['name']
            if pname in kwargs:
                val = kwargs[pname]
                # 字符串转字节流 (GBK 兼容 Windows 中文路径)
                if p['type'] == 4 and isinstance(val, str):
                    val = val.encode('gbk')
                args.append(val)
            else:
                # print(f"⚠️ 参数 '{pname}' 未提供，默认传 0")
                args.append(0)

        return c_func(*args)

class GeometryStrictRunner:
    def __init__(self, dll_path):
        self.dll_path = dll_path
        self.loader = None
        
    def load(self):
        if not self.loader:
            self.loader = UniversalLoader(self.dll_path)
    
    def run_analysis(self, priority_csv_path, output_csv_path, target_layer=20):
        """
        执行 RunGeometryAnalysis
        :param priority_csv_path: 优先级定义 CSV (通常是 Counterbore Info)
        :param output_csv_path: 结果输出路径
        :param target_layer: 目标图层
        :return: 0=成功, 其他=失败代码
        """
        self.load()
        
        # 必须匹配 RunGeometryAnalysis 的参数
        # 假设 DLL 定义为: RunGeometryAnalysis(input_csv_path, output_csv_path, target_layer)
        # 注意: 加工面方向判断.py 中使用了 input_csv_path=priority_csv_path
        
        # 先检查函数是否存在
        if 'RunGeometryAnalysis' not in self.loader.functions:
            raise AttributeError("DLL 中未找到 'RunGeometryAnalysis' 函数")

        # 动态传参
        # 获取形参名以确保正确传递 keys
        # params = self.loader.functions['RunGeometryAnalysis']['params']
        # param_names = [p['name'] for p in params]
        # 但我们这里直接按观察到的名字传，UniversalLoader 会匹配
        
        ret = self.loader.RunGeometryAnalysis(
            input_csv_path=priority_csv_path,
            output_csv_path=output_csv_path,
            target_layer=target_layer
        )
        return ret
