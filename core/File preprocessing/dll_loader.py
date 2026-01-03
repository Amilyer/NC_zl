# -*- coding: utf-8 -*-
"""
DLL 加载器 (dll_loader.py)
功能：通用 DLL 加载与接口调用封装
"""

import ctypes
import json
import os


class UniversalLoader:
    """通用 DLL 加载器"""
    
    def __init__(self, dll_path):
        if not os.path.exists(dll_path): 
            raise FileNotFoundError(f"DLL not found: {dll_path}")
        self.dll = ctypes.CDLL(dll_path)
        self.functions = {}
        self._load_metadata()

    def _load_metadata(self):
        # 获取函数数量
        try:
            self.dll.get_func_count.restype = ctypes.c_int
            count = self.dll.get_func_count()
        except AttributeError:
            raise Exception("DLL 不支持通用接口 (缺少 get_func_count)")

        # 定义 get_func_info 参数类型
        self.dll.get_func_info.argtypes = [ctypes.c_int, ctypes.c_char_p, ctypes.c_int]
        buf = ctypes.create_string_buffer(4096)
        
        for i in range(count):
            if self.dll.get_func_info(i, buf, 4096) == 0:
                json_str = buf.value.decode()
                try:
                    info = json.loads(json_str)
                    self._register_func(info)
                except json.JSONDecodeError:
                    print(f"⚠️ 警告: 第 {i} 个函数的元数据 JSON 解析失败")

    def _register_func(self, info):
        func_name = info['name']
        if not hasattr(self.dll, func_name): 
            return
        
        c_func = getattr(self.dll, func_name)
        self.functions[func_name] = info
        
        # 映射参数类型 (4=String, 1=Int, etc.)
        argtypes = []
        for p in info['params']:
            if p['type'] == 4: 
                argtypes.append(ctypes.c_char_p)
            else: 
                argtypes.append(ctypes.c_int)
            
        c_func.argtypes = argtypes
        c_func.restype = ctypes.c_int 

    def __getattr__(self, name):
        if name not in self.functions: 
            raise AttributeError(f"DLL 中未找到函数: {name}")
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
                # 字符串转字节流 (GBK 用于兼容 Windows 路径)
                if p['type'] == 4 and isinstance(val, str):
                    val = val.encode('gbk') 
                args.append(val)
            else:
                # print(f"⚠️ 参数 '{pname}' 未提供，默认传 0")
                args.append(0)
        
        # 调用 C++ 函数
        ret_val = c_func(*args)
        return {'return': ret_val}
