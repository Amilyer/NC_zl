# -*- coding: utf-8 -*-
"""
AI 分类器 (ai_classifier.py)
功能：加载模型并预测零件类型
"""

import os
import sys
import traceback

import joblib

# -----------------------------------------------------------------------------
# 路径配置与依赖导入 (Refactored)
# -----------------------------------------------------------------------------

# 1. 动态添加依赖路径
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(CURRENT_DIR, "..", "..")) # Assuming c:/Projects/NC/
DEMO_LAST_DIR = os.path.abspath(os.path.join(CURRENT_DIR, "..", "demo_last"))

# 修复：强制添加当前环境的 site-packages 到 sys.path
# NX 环境可能会重置 sys.path，导致 pip 安装的包丢失
try:
    import site
    # 1. 尝试使用 site 模块获取
    site_packages = site.getsitepackages()
    for p in site_packages:
        if os.path.exists(p) and p not in sys.path:
            sys.path.insert(0, p) # 插入到最前，优先使用环境包
            
    # 2. 备用：根据 sys.executable 推断 (针对 Conda/Virtualenv)
    if sys.executable:
        env_root = os.path.dirname(sys.executable) # e.g., .../envs/NC_env
        libs_path = os.path.join(env_root, "Lib", "site-packages")
        if os.path.exists(libs_path) and libs_path not in sys.path:
            sys.path.insert(0, libs_path)
            print(f"DEBUG: Added site-packages: {libs_path}")

    # 3. 针对混合环境的特殊修复 (User Site Packages)
    # pip show numpy-stl 显示安装在了这里:
    user_site = r"C:\Users\lutonglin\AppData\Roaming\Python\Python313\site-packages"
    if os.path.exists(user_site) and user_site not in sys.path:
        sys.path.insert(0, user_site)
        print(f"DEBUG: Added user site-packages: {user_site}")
            
except Exception as e:
    print(f"DEBUG: Path setup warning: {e}")


POSSIBLE_LIB_PATHS = [
    DEMO_LAST_DIR,
    os.path.join(PROJECT_ROOT, "core", "demo_last"),
    os.path.join(PROJECT_ROOT, "input"), 
]



for p in POSSIBLE_LIB_PATHS:
    if p not in sys.path and os.path.exists(p):
        sys.path.insert(0, p)

print(f"DEBUG: sys.path for AI imports (Final): {sys.path[:5]} ...")

# 2. 顶层导入
try:
    from nx_python_export import export_single_stl
    from point_cloud.utils.features.core_features import (
        extract_core_features_from_file,
    )
    from stl_to_pcd_v2 import stl_to_pcd
    print("✅ AI 依赖库导入成功 (Top-level)")
except ImportError as e:
    print(f"⚠️ AI 依赖库导入失败 (Top-level): {e}")
    print("   这可能导致 AI 预测功能不可用。请检查 'numpy-stl' 是否安装，以及 helper 模块路径是否正确。")
    # 设置为 None，后续做 check
    export_single_stl = None
    extract_core_features_from_file = None
    stl_to_pcd = None


class AIClassifier:
    """AI 分类与预测逻辑"""

    def __init__(self, pm):
        """
        Args:
            pm: PathManager 实例
        """
        self.pm = pm
        self.model = None
        self.scaler = None
        self.config = None
        self.is_loaded = False

    def load_models(self) -> bool:
        """加载预训练模型"""
        # 如果顶层导入失败，直接返回
        if extract_core_features_from_file is None:
            print("❌ AI 依赖库未正确加载，无法启用预测功能。")
            return False

        try:
            # 加载模型文件
            model_path = self.pm.get_rf_model_path()
            scaler_path = self.pm.get_scaler_path()
            config_path = self.pm.get_model_config_path()

            if not (os.path.exists(model_path) and os.path.exists(scaler_path)):
                print(f"⚠️ 模型文件缺失: {model_path} 或 {scaler_path}")
                return False

            self.model = joblib.load(model_path)
            self.scaler = joblib.load(scaler_path)
            self.config = joblib.load(config_path)
            self.is_loaded = True
            print("✅ AI 模型文件加载成功")
            return True
        except Exception as e:
            print(f"❌ 模型文件加载失败: {e}")
            traceback.print_exc()
            return False

    def predict(self, work_part, part_name: str) -> str:
        """
        对当前部件进行预测
        """
        if not self.is_loaded:
            return None

        # 二次检查依赖
        if export_single_stl is None or stl_to_pcd is None:
            return None

        try:
            # 1. 导出 STL
            stl_path = self.pm.get_stl_path(part_name)
            os.makedirs(os.path.dirname(stl_path), exist_ok=True)
            
            exported_stl = export_single_stl(work_part, stl_path, 0.05, 5)
            if not exported_stl:
                print(f"❌ [AI Debug] STL 导出失败: {stl_path}")
                return None

            # 2. STL -> PCD
            pcd_path = self.pm.get_pcd_path(part_name)
            os.makedirs(os.path.dirname(pcd_path), exist_ok=True)
            
            _, final_pcd = stl_to_pcd(exported_stl, pcd_path, point_count=50000, visualize=False)
            if not final_pcd:
                print(f"❌ [AI Debug] PCD 转换失败: {pcd_path}")
                return None

            # 3. 特征提取与预测
            feature = extract_core_features_from_file(final_pcd)
            feature_scaled = self.scaler.transform(feature.reshape(1, -1))
            prediction_idx = self.model.predict(feature_scaled)[0]
            
            label = self.config['class_names'][prediction_idx]
            return label

        except Exception as e:
            print(f"⚠️ 预测过程出错: {e}")
            traceback.print_exc()
            return None

if __name__ == "__main__":
    print("AIClassifier 模块测试")
