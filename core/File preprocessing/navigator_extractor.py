# -*- coding: utf-8 -*-
"""
导航器提取器 (navigator_extractor.py)
功能：准备 CAM 环境并调用导航器 DLL 进行特征识别
"""


from dll_loader import UniversalLoader

try:
    import NXOpen
    import NXOpen.CAM
    _NX_AVAILABLE = True
except ImportError:
    _NX_AVAILABLE = False

class NavigatorExtractor:
    """导航器提取工具类"""

    def __init__(self, dll_path):
        self.dll_path = dll_path
        self.plugin = None
        self._load_plugin()
        
        if _NX_AVAILABLE:
            self.session = NXOpen.Session.GetSession()

    def _load_plugin(self):
        try:
            self.plugin = UniversalLoader(self.dll_path)
            # print(f"✅ 导航器 DLL 加载成功")
        except Exception as e:
            print(f"❌ 导航器 DLL 加载失败: {e}")
            raise e

    def ensure_cam_setup_ready(self, work_part):
        """
        智能准备 CAM 环境
        """
        if not _NX_AVAILABLE: return False
        
        try:
            # 1. 检查 CAM 会话
            if not self.session.IsCamSessionInitialized():
                self.session.CreateCamSession()

            # 2. 检查 Setup
            try:
                if work_part.CamSetup.IsInitialized():
                    return True
            except:
                pass 

            # 3. 创建 Setup
            print("   ⚡ 正在创建 CAM Setup (hole_making)...")
            try:
                work_part.CreateCamSetup("hole_making")
                return True
            except Exception:
                print("   ⚠️ hole_making 失败，尝试 mill_planar...")
                work_part.CreateCamSetup("mill_planar")
                return True

        except Exception as ex:
            print(f"   ❌ CAM Setup 准备失败: {ex}")
            return False

    def process_part(self, work_part, output_dir, target_layer=0):
        """
        执行特征识别
        """
        if not self.plugin:
            return -999

        # 1. 准备环境
        if not self.ensure_cam_setup_ready(work_part):
            return -997

        # 2. 调用 DLL
        try:
            ret = self.plugin.run_feature_recognition(
                output_dir=str(output_dir),
                target_layer=target_layer
            )
            return ret
        except Exception as e:
            print(f"   ❌ 导航器调用异常: {e}")
            return -998
