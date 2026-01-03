# -*- coding: utf-8 -*-
"""
图层管理器 (layer_manager.py)
功能：管理 NX 图层操作，如查找特征并移动图层
"""

import traceback

# 尝试导入 NXOpen
try:
    import NXOpen
    import NXOpen.UF
    _NX_AVAILABLE = True
except ImportError:
    _NX_AVAILABLE = False

class LayerManager:
    """图层操作管理类"""

    def __init__(self):
        if not _NX_AVAILABLE:
            raise RuntimeError("NXOpen 模块未找到，请在 NX 环境中运行")
        
        self.session = NXOpen.Session.GetSession()
        self.uf_session = NXOpen.UF.UFSession.GetUFSession()

    def switch_to_manufacturing(self):
        """切换到制造应用 (UG_APP_MANUFACTURING)"""
        try:
            self.session.ApplicationSwitchImmediate("UG_APP_MANUFACTURING")
            self.uf_session.Cam.InitSession()
            return True
        except Exception as e:
            print(f"⚠️ 切换应用失败: {e}")
            return False

    def process_part(self, work_part, target_layer: int) -> bool:
        """
        处理单个部件：查找特征并移动到目标图层
        
        Args:
            work_part: NXOpen.Part 对象
            target_layer: 目标图层 ID
            
        Returns:
            bool: 是否成功移动
        """
        if not work_part:
            return False

        try:
            # 1. 查找特征
            features = self._find_features(work_part)
            if not features:
                print("⚠️ 未找到符合条件的特征 (Bodies/Curves/Notes)")
                return False

            # 2. 移动图层
            return self._move_objects(work_part, features, target_layer)

        except Exception as e:
            print(f"❌ 图层处理出错: {e}")
            traceback.print_exc()
            return False

    def _find_features(self, work_part) -> list:
        """查找所有需要移动的特征"""
        features = []
        
        # 收集曲线
        features.extend([c for c in work_part.Curves])
        # 收集注释
        features.extend([n for n in work_part.Notes])
        # 收集实体
        features.extend([b for b in work_part.Bodies])
        
        print(f"  - 找到 {len(features)} 个特征")
        return features

    def _move_objects(self, work_part, objects: list, layer_id: int) -> bool:
        """移动对象到指定图层"""
        if not objects:
            return False

        try:
            mark_id = self.session.SetUndoMark(
                NXOpen.Session.MarkVisibility.Visible, "Move Layer"
            )
            
            work_part.Layers.MoveDisplayableObjects(layer_id, objects)
            
            print(f"  ✓ 已将 {len(objects)} 个对象移动到图层 {layer_id}")
            return True
        except Exception as e:
            print(f"❌ 移动对象失败: {e}")
    def copy_layer_objects(self, work_part, source_layer: int, target_layer: int) -> bool:
        """复制源图层的所有实体到目标图层"""
        try:
            bodies = [b for b in work_part.Bodies if b.Layer == source_layer]
            if not bodies:
                print(f"  ⚠️ 源图层 {source_layer} 没有实体")
                return False
                
            self.session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Copy Layer")
            work_part.Layers.CopyObjects(target_layer, bodies)
            print(f"  ✓ 已复制 {len(bodies)} 个实体: {source_layer} -> {target_layer}")
            return True
        except Exception as e:
            print(f"❌ 复制图层失败: {e}")
            return False

if __name__ == "__main__":
    print("LayerManager 模块")
