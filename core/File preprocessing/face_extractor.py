# -*- coding: utf-8 -*-
"""
面信息提取器 (face_extractor.py)
功能：调用 DLL 提取面信息和腔体分析
"""

import os

from dll_loader import UniversalLoader


class FaceExtractor:
    """面信息提取工具类"""

    def __init__(self, dll_path):
        self.dll_path = dll_path
        self.plugin = None
        self._load_plugin()

    def _load_plugin(self):
        try:
            self.plugin = UniversalLoader(self.dll_path)
            print(f"✅ DLL 加载成功: {os.path.basename(self.dll_path)}")
        except Exception as e:
            print(f"❌ DLL 加载失败: {e}")
            raise e

    def process_part(self, output_root, target_layer=0):
        """
        处理当前打开的部件
        
        Args:
            output_root: 输出根目录
            target_layer: 目标图层
            
        Returns:
            int: 返回码 (0=成功)
        """
        if not self.plugin:
            return -999

        try:
            # 调用 DLL 接口
            # 注意：这里假设 DLL 接口名为 run_extraction_and_save，参数匹配
            res = self.plugin.run_extraction_and_save(
                base_save_dir=output_root, 
                target_layer=target_layer
            )
            return res['return']
        except Exception as e:
            print(f"❌ 提取过程异常: {e}")
            return -998
