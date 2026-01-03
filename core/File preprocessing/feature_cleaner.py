# -*- coding: utf-8 -*-
"""
特征清理器 (feature_cleaner.py)
功能：删除孔、指定颜色的面，以及移除参数
"""

import traceback

try:
    import NXOpen
    import NXOpen.Features
    import NXOpen.GeometricUtilities
    _NX_AVAILABLE = True
except ImportError:
    _NX_AVAILABLE = False

class FeatureCleaner:
    """特征清理工具类"""

    def __init__(self):
        if not _NX_AVAILABLE:
            raise RuntimeError("NXOpen 模块未找到")
        self.session = NXOpen.Session.GetSession()

    def clean_part(self, work_part, target_layer: int, target_color_index: int):
        """
        清理部件：删除孔 -> 删除颜色面 -> 移除参数
        """
        try:
            # 1. 删除孔
            self._delete_holes(work_part, target_layer)
            
            # 2. 删除指定颜色的面
            self._delete_color_faces(work_part, target_layer, target_color_index)
            
            # 3. 更新模型
            try:
                self.session.UpdateManager.DoUpdate(self.session.NewestVisibleUndoMark)
            except:
                pass
            
            # 4. 移除参数
            self._remove_parameters(work_part, target_layer)
            
            return True
        except Exception as e:
            print(f"❌ 特征清理失败: {e}")
            traceback.print_exc()
            return False

    def _delete_holes(self, work_part, layer):
        """
        调用外部 '删除孔.py' 进行清理
        """
        print("  - [External] 调用 '删除孔.py' 进行孔清理...")
        
        # 动态导入
        try:
            import os
            import importlib.util
            import sys
            
            module_name = "hole_deleter_module"
            current_dir = os.path.dirname(os.path.abspath(__file__))
            file_path = os.path.join(current_dir, "删除孔.py")
            
            if not os.path.exists(file_path):
                print(f"    ❌ 找不到文件: {file_path}")
                return

            spec = importlib.util.spec_from_file_location(module_name, file_path)
            if spec and spec.loader:
                module = importlib.util.module_from_spec(spec)
                sys.modules[module_name] = module
                spec.loader.exec_module(module)
                
                # 直接调用封装好的统一入口函数
                if hasattr(module, 'run_delete_logic'):
                    print("    -> 调用 run_delete_logic (包含方案一/二及双轮循环)...")
                    module.run_delete_logic(self.session, work_part)
                elif hasattr(module, 'delete_hole_v1_main'):
                    # 兼容旧版本：如果没有 run_delete_logic，至少跑一次 v1
                    print("    ⚠️ 未找到 run_delete_logic，尝试运行 delete_hole_v1_main...")
                    module.delete_hole_v1_main(self.session, work_part)
                else:
                    print("    ❌ 模块中未找到 'run_delete_logic' 或 'delete_hole_v1_main' 函数")
            else:
                print("    ❌ 模块加载失败")

        except Exception as e:
            print(f"    ❌ 调用外部脚本失败: {e}")
            import traceback
            traceback.print_exc()

    def _delete_color_faces(self, work_part, layer, color_index):
        """删除指定颜色的面"""
        print(f"  - 正在扫描颜色 ID {color_index} 的面...")
        
        # 收集所有符合颜色的面
        target_faces = []
        for body in work_part.Bodies:
            if body.Layer == layer:
                for face in body.GetFaces():
                    if face.Color == color_index:
                        target_faces.append(face)
        
        if not target_faces:
            print("    未找到目标颜色的面")
            return

        # 分组删除 (按连通性)
        groups = self._group_connected_faces(target_faces)
        print(f"    找到 {len(target_faces)} 个面，分为 {len(groups)} 组进行删除")
        
        success_count = 0
        for group in groups:
            if self._execute_delete_face(work_part, group):
                success_count += 1
                
        print(f"    颜色面删除完成: {success_count}/{len(groups)} 组")

    def _execute_delete_face(self, work_part, faces):
        """执行删除面操作"""
        try:
            builder = work_part.Features.CreateDeleteFaceBuilder(NXOpen.Features.Feature.Null)
            
            # 关闭智能识别，只删选中的
            fr = builder.FaceRecognized
            fr.CoplanarEnabled = False
            fr.CoaxialEnabled = False
            fr.EqualDiameterEnabled = False
            fr.TangentEnabled = False
            fr.SymmetricEnabled = False
            fr.OffsetEnabled = False
            
            builder.Type = NXOpen.Features.DeleteFaceBuilder.SelectTypes.Face
            builder.FaceEdgeBlendPreference = NXOpen.Features.DeleteFaceBuilder.FaceEdgeBlendPreferenceOptions.Cliff
            
            # 添加面到收集器
            scRuleFactory = work_part.ScRuleFactory
            ruleOptions = scRuleFactory.CreateRuleOptions()
            ruleOptions.SetSelectedFromInactive(False)
            faceRule = scRuleFactory.CreateRuleFaceDumb(faces, ruleOptions)
            ruleOptions.Dispose()
            builder.FaceCollector.ReplaceRules([faceRule], False)
            
            builder.Commit()
            builder.Destroy()
            return True
        except Exception:
            return False

    def _remove_parameters(self, work_part, layer):
        """移除参数"""
        bodies = [b for b in work_part.Bodies if b.Layer == layer]
        if not bodies: return
        
        try:
            builder = work_part.Features.CreateRemoveParametersBuilder()
            builder.Objects.Add(bodies)
            builder.Commit()
            builder.Destroy()
            print("  - 已移除实体参数")
        except Exception as e:
            print(f"  ⚠️ 移除参数失败: {e}")

    def _group_connected_faces(self, faces):
        """将面按连通性分组"""
        face_pool = set(faces)
        groups = []
        while face_pool:
            seed = face_pool.pop()
            group = {seed}
            stack = [seed]
            while stack:
                f = stack.pop()
                try:
                    for edge in f.GetEdges():
                        for neighbor in edge.GetFaces():
                            if neighbor in face_pool:
                                face_pool.remove(neighbor)
                                group.add(neighbor)
                                stack.append(neighbor)
                except: pass
            groups.append(list(group))
        return groups
