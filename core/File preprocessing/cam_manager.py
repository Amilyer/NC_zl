# -*- coding: utf-8 -*-
"""
CAM 管理模块 (cam_manager.py)
功能：封装 CAM 相关的核心操作，包括 MCS 创建、刀具创建、JSON 生成和刀轨生成。
"""

import os
import traceback

import pandas as pd

try:
    import NXOpen
    import NXOpen.CAM
    import NXOpen.Layer
    import NXOpen.UF
except ImportError:
    pass

# 导入 Step 13 的生成器模块
# 假设这些模块在 modules/综合json输出 下，或者已经在 sys.path 中
try:
    import 生成往复等高 as zlevel_gen
    import 生成爬面文件 as cam_gen
    import 生成螺旋文件 as spiral_gen
    import 生成行腔文件 as cavity_gen
    import 生成面铣文件 as face_gen
except ImportError:
    pass

# 导入 Step 14 的刀轨生成核心
try:
    import 创建刀轨 as toolpath_module
except ImportError:
    pass

class CAMManager:
    def __init__(self, work_part):
        self.work_part = work_part
        self.session = NXOpen.Session.GetSession()
        self.uf = NXOpen.UF.UFSession.GetUFSession()

    def log(self, message, level="INFO"):
        print(f"[{level}] {message}")

    # ==========================================================================
    # 1. MCS & Workpiece (原 Step 11)
    # ==========================================================================
    
    def create_mcs_and_workpiece(self, target_layer=20):
        """创建 MCS 和 Workpiece"""
        try:
            # 1. 找到目标体
            target_body = self._find_body_on_layer(target_layer)
            if not target_body:
                self.log("未找到目标图层的实体", "WARN")
                return False

            # 2. 确保 CAM Setup
            if not self._ensure_cam_setup():
                return False

            # 3. 创建包容体 (用于计算 MCS)
            tooling_box = self._create_tooling_box(target_body)
            if not tooling_box:
                return False

            # 4. 创建 MCS
            mcs_obj = self._create_mcs(tooling_box)
            if not mcs_obj:
                return False

            # 5. 创建 Workpiece
            wp_obj = self._create_workpiece(mcs_obj, target_body)
            
            # 清理包容体
            # self._delete_object(tooling_box) # 可选

            return wp_obj is not None

        except Exception as e:
            self.log(f"MCS/Workpiece 创建失败: {e}", "ERROR")
            traceback.print_exc()
            return False

    def _find_body_on_layer(self, layer):
        for body in self.work_part.Bodies:
            if body.Layer == layer:
                return body
        return None

    def _ensure_cam_setup(self):
        if not self.session.IsCamSessionInitialized():
            self.session.CreateCamSession()
        
        try:
            if self.work_part.CAMSetup.IsInitialized():
                return True
        except:
            pass
            
        try:
            self.work_part.CreateCamSetup("hole_making")
            return True
        except Exception as e:
            self.log(f"创建 Setup 失败: {e}", "ERROR")
            return False

    def _create_tooling_box(self, target_body):
        """根据目标实体自动创建包容体 (仅用于定位)"""
        try:
            self.session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "创建包容体")

            tooling_box_builder = self.work_part.Features.ToolingFeatureCollection.CreateToolingBoxBuilder(NXOpen.Features.ToolingBox.Null)
            tooling_box_builder.Type = NXOpen.Features.ToolingBoxBuilder.Types.BoundedBlock

            for offset in [tooling_box_builder.OffsetPositiveX, tooling_box_builder.OffsetNegativeX, tooling_box_builder.OffsetPositiveY, tooling_box_builder.OffsetNegativeY, tooling_box_builder.OffsetPositiveZ, tooling_box_builder.OffsetNegativeZ]:
                offset.SetFormula("0")

            # 设置包容体方向与WCS一致
            matrix = NXOpen.Matrix3x3()
            matrix.Xx, matrix.Xy, matrix.Xz = 1.0, 0.0, 0.0
            matrix.Yx, matrix.Yy, matrix.Yz = 0.0, 1.0, 0.0
            matrix.Zx, matrix.Zy, matrix.Zz = 0.0, 0.0, 1.0
            tooling_box_builder.SetBoxMatrixAndPosition(matrix, NXOpen.Point3d(0.0, 0.0, 0.0))

            rule_options = self.work_part.ScRuleFactory.CreateRuleOptions()
            rule_options.SetSelectedFromInactive(False)
            body_rule = self.work_part.ScRuleFactory.CreateRuleBodyDumb([target_body], True, rule_options)
            rule_options.Dispose()

            sc_collector = tooling_box_builder.BoundedObject
            sc_collector.ReplaceRules([body_rule], False)
            tooling_box_builder.CalculateBoxSize()

            tooling_box_feature = tooling_box_builder.Commit()
            tooling_box_builder.Destroy()

            bodies = tooling_box_feature.GetBodies()
            if bodies and len(bodies) > 0:
                return bodies[0]
            else:
                return None
        except Exception as e:
            self.log(f"包容体创建失败: {e}", "ERROR")
            return None

    def _left_down_point(self, body):
        """获取包容体的最小XYZ点"""
        bbox = self.uf.ModlGeneral.AskBoundingBox(body.Tag)
        # 返回 Xmin, Ymin, Zmax (作为 MCS 原点和安全平面的 Z 参考)
        return (bbox[0], bbox[1], bbox[5])

    def _find_face_parallel_to_xy(self, body, extreme_type='max'):
        """寻找Z方向最极端的水平面（用于安全平面）"""
        found_face = None
        extreme_value = float('-inf')

        for face in body.GetFaces():
            if face.SolidFaceType == NXOpen.Face.FaceType.Planar:
                try:
                    bbox = self.uf.ModlGeneral.AskBoundingBox(face.Tag)
                    z_min, z_max = bbox[2], bbox[5]
                    if abs(z_max - z_min) < 0.001:
                        current_z = z_max
                        if extreme_type == 'max' and current_z > extreme_value:
                            extreme_value = current_z
                            found_face = face
                except:
                    continue
        return found_face

    def _create_mcs(self, tooling_box, mcs_name="MCS_1", safe_distance=1.0):
        """创建MCS坐标系并设置安全平面"""
        # 用包容体的顶面来计算安全平面
        top_face = self._find_face_parallel_to_xy(tooling_box, "max")
        if not top_face:
            self.log("未找到包容体顶面，无法创建安全平面", "WARN")
            return None
            
        points = self._left_down_point(tooling_box)

        try:
            # 删除旧的
            try:
                existing = self.work_part.CAMSetup.CAMGroupCollection.FindObject(f"GEOMETRY/{mcs_name}")
                if existing: existing.Delete()
            except: pass

            geom_group = self.work_part.CAMSetup.CAMGroupCollection.FindObject("GEOMETRY")
            mcs_group = self.work_part.CAMSetup.CAMGroupCollection.CreateGeometry(
                geom_group, "mill_contour", "MCS",
                NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue, mcs_name
            )
            builder = self.work_part.CAMSetup.CAMGroupCollection.CreateMillOrientGeomBuilder(mcs_group)
            
            # 使用包容体的左下角作为坐标系原点 (Xmin, Ymin, Zmax)
            origin3 = NXOpen.Point3d(points[0], points[1], points[2]) 
            x_dir = NXOpen.Vector3d(1.0, 0.0, 0.0)
            y_dir = NXOpen.Vector3d(0.0, 1.0, 0.0)
            xform = self.work_part.Xforms.CreateXform(origin3, x_dir, y_dir, NXOpen.SmartObject.UpdateOption.AfterModeling, 1.0)
            csys = self.work_part.CoordinateSystems.CreateCoordinateSystem(xform, NXOpen.SmartObject.UpdateOption.AfterModeling)
            builder.Mcs = csys
            
            # 设置安全平面
            builder.TransferClearanceBuilder.ClearanceType = NXOpen.CAM.NcmClearanceBuilder.ClearanceTypes.Plane
            
            # 创建临时平面用于安全平面设置
            plane_safe = self.work_part.Planes.CreatePlane(NXOpen.Point3d(0.0, 0.0, 0.0), NXOpen.Vector3d(0.0, 0.0, 1.0), NXOpen.SmartObject.UpdateOption.AfterModeling)
            plane_safe.SetMethod(NXOpen.PlaneTypes.MethodType.Distance)
            plane_safe.SetGeometry([top_face])
            expr = plane_safe.Expression
            expr.RightHandSide = str(safe_distance)
            plane_safe.SetAlternate(NXOpen.PlaneTypes.AlternateType.One)
            plane_safe.Evaluate()
            builder.TransferClearanceBuilder.PlaneXform = plane_safe

            nx_obj = builder.Commit()
            builder.Destroy()
            return nx_obj

        except Exception as e:
            self.log(f"MCS 创建失败: {e}", "ERROR")
            return None

    def _create_workpiece(self, parent_mcs, part_body, workpiece_name="WORKPIECE_1"):
        try:
            # 检查重名工件并删除
            try:
                existing = parent_mcs.FindObject(workpiece_name)
                if existing: self.uf.Obj.DeleteObject(existing.Tag)
            except: pass

            wp = self.work_part.CAMSetup.CAMGroupCollection.CreateGeometry(
                parent_mcs, "mill_contour", "WORKPIECE",
                NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue, workpiece_name
            )
            builder = self.work_part.CAMSetup.CAMGroupCollection.CreateMillGeomBuilder(wp)
            
            # 设置 Part
            builder.PartGeometry.InitializeData(False)
            geom_set = builder.PartGeometry.GeometryList.FindItem(0)
            rule_opt = self.work_part.ScRuleFactory.CreateRuleOptions()
            rule_opt.SetSelectedFromInactive(False)
            rule = self.work_part.ScRuleFactory.CreateRuleBodyDumb([part_body], True, rule_opt)
            geom_set.ScCollector.ReplaceRules([rule], False)
            rule_opt.Dispose()
            
            # 暂不设置毛坯 (Blank)
            
            builder.Commit()
            builder.Destroy()
            return wp
        except Exception as e:
            self.log(f"Workpiece 创建失败: {e}", "ERROR")
            return None

    # ==========================================================================
    # 2. Tool Creation (原 Step 12)
    # ==========================================================================

    def create_tools_from_excel(self, excel_path):
        """从 Excel 创建刀具"""
        if not os.path.exists(excel_path):
            self.log("刀具 Excel 不存在", "ERROR")
            return False

        try:
            # 使用 sheet_name=0 读取第一个工作表
            df = pd.read_excel(excel_path, sheet_name=0, header=1)
            count = 0
            for _, row in df.iterrows():
                name = str(row['刀具名称']).strip()
                if not name or name == '刀具名称': continue
                
                try:
                    dia = float(row['直径'])
                    r = float(row['R角'])
                    self._create_mill_tool(name, dia, r)
                    count += 1
                except:
                    continue
            self.log(f"创建了 {count} 把刀具", "SUCCESS")
            return True
        except Exception as e:
            self.log(f"刀具创建失败: {e}", "ERROR")
            return False

    def _create_mill_tool(self, name, diameter, r_radius):
        try:
            # 检查是否存在
            try:
                self.work_part.CAMSetup.CAMGroupCollection.FindObject(name)
                return # 已存在
            except:
                pass

            parent = self._get_tool_parent_group()
            tool = self.work_part.CAMSetup.CAMGroupCollection.CreateTool(
                parent, "hole_making", "MILL",
                NXOpen.CAM.NCGroupCollection.UseDefaultName.FalseValue, name
            )
            builder = self.work_part.CAMSetup.CAMGroupCollection.CreateMillToolBuilder(tool)
            builder.TlDiameterBuilder.Value = diameter
            if hasattr(builder, "TlR1Builder"):
                builder.TlR1Builder.Value = r_radius
            builder.Commit()
            builder.Destroy()
        except Exception as e:
            self.log(f"创建刀具 {name} 失败: {e}", "WARN")

    def _get_tool_parent_group(self):
        # 尝试找 GENERIC_MACHINE 或创建
        try:
            return self.work_part.CAMSetup.CAMGroupCollection.FindObject("GENERIC_MACHINE")
        except:
            # 简单起见，挂在 NC_PROGRAM 下或者创建一个
            return self.work_part.CAMSetup.CAMGroupCollection.FindObject("NC_PROGRAM")

    # ==========================================================================
    # 3. JSON Generation (原 Step 13)
    # ==========================================================================

    def generate_jsons(self, prt_path, aux_files, pm):
        """生成所有刀轨配置 JSON"""
        prt_name = os.path.splitext(os.path.basename(prt_path))[0]
        
        face_csv = aux_files.get('face_csv')
        geo_csv = aux_files.get('geo_csv')
        feature_csv = aux_files.get('feature_csv')
        cavity_csv = aux_files.get('cavity_csv')
        
        if not (face_csv and geo_csv and feature_csv):
            self.log("缺少必要 CSV 文件，跳过 JSON 生成", "WARN")
            return False

        knife_json = pm.get_knife_table_json()
        tool_excel = pm.get_tool_params_excel_path()

        try:
            # 1. Cavity
            cavity_gen.generate_cavity_json_v2(prt_path, feature_csv, face_csv, geo_csv, knife_json, pm.get_json_output_path(prt_name, 'cavity'))
            # 2. ZLevel
            zlevel_gen.generate_zlevel_json(prt_path, face_csv, geo_csv, feature_csv, pm.get_json_output_path(prt_name, 'zlevel'))
            # 3. CAM (爬面)
            cam_gen.generate_cam_json(prt_path, cavity_csv or face_csv, tool_excel, pm.get_json_output_path(prt_name, 'cam'))
            # 4. Face
            face_gen.generate_face_milling_json(prt_path, tool_excel, face_csv, geo_csv, pm.get_json_output_path(prt_name, 'face'))
            # 5. Spiral
            res = spiral_gen.group_with_tools_and_depth(feature_csv, tool_excel, prt_path, geo_csv)
            spiral_gen.save_result_to_json(res, pm.get_json_output_path(prt_name, 'spiral'))
            
            return True
        except Exception as e:
            self.log(f"JSON 生成失败: {e}", "ERROR")
            traceback.print_exc()
            return False

    # ==========================================================================
    # 4. Toolpath Generation (原 Step 14)
    # ==========================================================================

    def generate_toolpaths(self, prt_path, pm):
        """生成最终刀轨"""
        prt_name = os.path.splitext(os.path.basename(prt_path))[0]
        
        # 获取 JSON 路径
        spiral = pm.get_json_output_path(prt_name, 'spiral')
        cavity = pm.get_json_output_path(prt_name, 'cavity')
        cam = pm.get_json_output_path(prt_name, 'cam')
        face = pm.get_json_output_path(prt_name, 'face')
        zlevel = pm.get_json_output_path(prt_name, 'zlevel')

        # 配置
        toolpath_module.CONFIG['TEST_MODE'] = False
        toolpath_module.CONFIG['AUTO_SAVE'] = True

        try:
            # 调用核心生成逻辑
            # 注意：这里我们已经在 NX 会话中打开了部件，但 toolpath_module.generate_toolpath_workflow 
            # 设计为接受文件路径并自己打开。
            # 为了避免冲突，我们可能需要先保存当前会话，或者修改 toolpath_module 以接受当前 part。
            # 鉴于 toolpath_module 逻辑复杂，最稳妥的方式是：
            # 在调用此方法前，外部先保存并关闭当前部件。然后让 toolpath_module 自己去打开处理。
            
            saved_path = toolpath_module.generate_toolpath_workflow(
                prt_path, spiral, cavity, cam, face, zlevel
            )
            return saved_path
        except Exception as e:
            self.log(f"刀轨生成失败: {e}", "ERROR")
            traceback.print_exc()
            return None
