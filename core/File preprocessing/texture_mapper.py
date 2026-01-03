# -*- coding: utf-8 -*-
"""
贴图模块 (texture_mapper.py)
功能：封装贴图 DLL 调用逻辑，支持复用
"""
import os

from dll_loader import UniversalLoader


class TextureMapper:
    def __init__(self, dll_path):
        self.dll_path = dll_path
        self.loader = None
        self._init_loader()

    def _init_loader(self):
        """初始化 DLL 加载器"""
        if os.path.exists(self.dll_path):
            try:
                self.loader = UniversalLoader(self.dll_path)
            except Exception as e:
                print(f"❌ 贴图 DLL 加载失败: {e}")
        else:
            print(f"⚠️ 找不到贴图 DLL: {self.dll_path}")

    def apply_texture(self):
        """
        执行自动对齐贴图逻辑
        
        Returns:
            int: API 返回码
                 0: 成功
                 -1: 无显示部件
                 2: 无匹配向量
                 -999: DLL 未加载
                 -998: 异常
        """
        print("  → 执行贴图处理...")
        
        if not self.loader:
             print(f"    ⚠️ 贴图 DLL 未就绪 (路径: {self.dll_path})")
             return -999

        try:
            # 调用新的自动对齐接口 RunAutoAlign
            # DLL 直接操作当前打开的 NX 部件
            res = self.loader.RunAutoAlign()
            # 兼容返回值可能是字典的情况
            if isinstance(res, dict):
                ret_code = res.get('return', -999)
            else:
                ret_code = res
            
            if ret_code == 0:
                print("    ✓ 贴图成功 (自动对齐完成)")
            elif ret_code == -1:
                print("    ⚠️ 贴图失败: 无显示部件")
            elif ret_code == 2:
                print("    ⚠️ 贴图警告: 未找到匹配的移动向量")
            else:
                print(f"    ⚠️ 贴图返回码: {ret_code}")
            
            return ret_code

        except Exception as e:
            if "RunAutoAlign" in str(e):
                    print(f"    ❌ 贴图接口未找到 (RunAutoAlign): {e}")
            else:
                    print(f"    ❌ 贴图异常: {e}")
            # traceback.print_exc()
            return -998
