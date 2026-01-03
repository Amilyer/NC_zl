# -*- coding: utf-8 -*-
"""
NX 处理器 (nx_processor.py)
功能：封装 NX Open API 操作
注意：必须在 NX Python 环境下运行
"""

import os

# 尝试导入 NXOpen
try:
    import NXOpen
    import NXOpen.Preferences
    _NX_AVAILABLE = True
except ImportError:
    _NX_AVAILABLE = False

class NXProcessor:
    """NX 操作封装类"""

    def __init__(self):
        if not _NX_AVAILABLE:
            raise RuntimeError("NXOpen 模块未找到，请在 NX 环境中运行")
        
        self.session = NXOpen.Session.GetSession()
        self.work_part = None
        self.display_part = None

    def open_part(self, prt_path: str) -> bool:
        """打开 PRT 文件"""
        if not os.path.exists(prt_path):
            print(f"❌ 文件不存在: {prt_path}")
            return False

        try:
            status = self.session.Parts.OpenBaseDisplay(prt_path)
            # 处理返回可能是元组的情况
            part = status[0] if isinstance(status, tuple) else status
            
            if part:
                self.session.Parts.SetWork(part)
                self.work_part = self.session.Parts.Work
                return True
            return False
        except Exception as e:
            print(f"❌ 打开文件失败: {e}")
            return False

    def import_part(self, file_path: str) -> bool:
        """导入外部部件 (PRT/STEP)"""
        if not self.work_part:
            print("❌ 未打开工作部件")
            return False

        try:
            importer = self.work_part.ImportManager.CreatePartImporter()
            importer.FileName = file_path
            importer.Scale = 1.0
            importer.CreateNamedGroup = True
            importer.ImportViews = False
            importer.ImportCamObjects = False
            importer.LayerOption = NXOpen.PartImporter.LayerOptionType.Original
            importer.DestinationCoordinateSystemSpecification = \
                NXOpen.PartImporter.DestinationCoordinateSystemSpecificationType.Work
            
            # 恢复原有的矩阵创建逻辑，防止内存错误
            element = NXOpen.Matrix3x3()
            element.Xx = 1.0; element.Xy = 0.0; element.Xz = 0.0
            element.Yx = 0.0; element.Yy = 1.0; element.Yz = 0.0
            element.Zx = 0.0; element.Zy = 0.0; element.Zz = 1.0
            importer.DestinationCoordinateSystem = self.work_part.NXMatrices.Create(element)
            
            # 设置原点
            importer.DestinationPoint = NXOpen.Point3d(0.0, 0.0, 0.0)
            
            importer.Commit()
            importer.Destroy()
            return True
        except Exception as e:
            print(f"❌ 导入失败: {e}")
            return False

    def save(self) -> bool:
        """保存当前部件"""
        if not self.work_part:
            return False
            
        try:
            self.work_part.Save(NXOpen.BasePart.SaveComponents.TrueValue, NXOpen.BasePart.CloseAfterSave.FalseValue)
            return True
        except Exception as e:
            print(f"❌ 保存失败: {e}")
            return False

    def save_as(self, output_path: str) -> bool:
        """另存为"""
        if not self.work_part:
            return False
            
        try:
            dir_name = os.path.dirname(output_path)
            if not os.path.exists(dir_name):
                os.makedirs(dir_name)
                
            self.work_part.SaveAs(output_path)
            return True
        except Exception as e:
            print(f"❌ 保存失败: {e}")
            return False

    def close_all(self):
        """关闭所有部件"""
        try:
            self.session.Parts.CloseAll(
                NXOpen.BasePart.CloseModified.UseResponses,
                None
            )
            self.work_part = None
        except:
            pass

    def get_current_part(self):
        return self.work_part

    def get_session(self):
        return self.session

if __name__ == "__main__":
    if _NX_AVAILABLE:
        print("✅ NX 环境检测正常")
    else:
        print("⚠️ 非 NX 环境，仅供代码检查")
