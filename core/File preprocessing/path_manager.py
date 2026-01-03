# -*- coding: utf-8 -*-
"""
路径管理器 (PathManager)
功能：统一管理项目中的所有文件路径
"""

from pathlib import Path
from typing import Optional
import config

class PathManager:
    def __init__(self, input_3d_prt: Optional[str] = None, input_2d_dxf: Optional[str] = None):
        # 1. 基础输入
        if input_3d_prt:
            self.input_3d_prt = Path(input_3d_prt).resolve()
        else:
            self.input_3d_prt = config.FILE_INPUT_PRT

        if input_2d_dxf:
            self.input_2d_dxf = Path(input_2d_dxf).resolve()
        else:
            self.input_2d_dxf = config.FILE_INPUT_DXF
        
        # 2. 核心目录
        self.project_root = Path(config.PROJECT_ROOT)
        self.input_dir = Path(config.INPUT_DIR)
        self.output_dir = Path(config.OUTPUT_DIR)
        self.work_dir = self.output_dir
        self.temp_dir = self.work_dir / '.temp'
        
        # 3. 一级子目录结构
        self.dir_resources = self.work_dir / '00_Resources'
        self.dir_input = self.work_dir / '01_Input'
        self.dir_process = self.work_dir / '02_Process'
        self.dir_analysis = self.work_dir / '03_Analysis'
        self.dir_step7 = self.work_dir / '04_PRT_with_Tool'
        self.dir_drilled = self.work_dir / '05_Drilled_PRT'
        self.dir_cam = self.work_dir / '06_CAM'
        self.dir_output = self.work_dir / '07_Output'

    def _get_dir(self, path: Path) -> Path:
        """获取目录并确保其存在"""
        path.mkdir(parents=True, exist_ok=True)
        return path

    def ensure_dirs(self):
        """一次性初始化核心目录"""
        self.get_split_prt_dir()
        self.get_split_dxf_dir()
        self.get_dxf_prt_dir()
        self.get_merged_prt_dir()
        self.ensure_temp_dirs()

    def ensure_temp_dirs(self):
        self.get_stl_dir()
        self.get_pcd_dir()

    # ==========================================================================
    # 01_Input
    # ==========================================================================
    def get_split_prt_dir(self) -> Path: return self._get_dir(self.dir_input / '3D_PRT_Split')
    def get_split_dxf_dir(self) -> Path: return self._get_dir(self.dir_input / '2D_DXF_Split')

    # ==========================================================================
    # 02_Process
    # ==========================================================================
    def get_dxf_prt_dir(self) -> Path:      return self._get_dir(self.dir_process / '1_DXF_to_PRT')
    def get_merged_prt_dir(self) -> Path:   return self._get_dir(self.dir_process / '2_Merged_PRT')
    def get_textured_prt_dir(self) -> Path: return self._get_dir(self.dir_process / '3_Textured_PRT')
    def get_cleaned_prt_dir(self) -> Path:  return self._get_dir(self.dir_process / '4_Cleaned_PRT')

    # ==========================================================================
    # 03_Analysis (Face Info, Navigator, Geometry)
    # ==========================================================================
    # Root Dirs
    def get_analysis_face_dir(self) -> Path: return self._get_dir(self.dir_analysis / 'Face_Info')
    def get_analysis_nav_dir(self) -> Path:  return self._get_dir(self.dir_analysis / 'Navigator_Reports')
    def get_analysis_geo_dir(self) -> Path:  return self._get_dir(self.dir_analysis / 'Geometry_Analysis')

    # Sub Dirs
    def get_face_csv_dir(self) -> Path:          return self._get_dir(self.get_analysis_face_dir() / 'face_csv')
    def get_analysis_face_prt_dir(self) -> Path: return self._get_dir(self.get_analysis_face_dir() / 'prt')
    
    def get_nav_csv_dir(self) -> Path:           return self._get_dir(self.get_analysis_nav_dir() / 'csv')
    def get_nav_prt_dir(self) -> Path:           return self._get_dir(self.get_analysis_nav_dir() / 'prt')
    def get_cavity_csv_dir(self) -> Path:        return self._get_dir(self.get_analysis_nav_dir() / 'cavity_csv')
    
    def get_feature_csv_dir(self) -> Path:       return self._get_dir(self.dir_analysis / 'Feature_Data')
    def get_counterbore_csv_dir(self) -> Path:   return self._get_dir(self.dir_analysis / 'Counterbore_Info')

    # ==========================================================================
    # 04_PRT_with_Tool (Step 7/8)
    # ==========================================================================
    def get_final_prt_dir(self) -> Path: return self._get_dir(self.dir_step7)
    def get_step8_prt_dir(self) -> Path: return self.get_final_prt_dir()

    # ==========================================================================
    # 05_Drilled_PRT (Step 9)
    # ==========================================================================
    def get_step9_drilled_dir(self) -> Path: return self._get_dir(self.dir_drilled)

    # ==========================================================================
    # 06_CAM (Step 10-15)
    # ==========================================================================
    # General
    def get_cam_json_dir(self) -> Path:      return self._get_dir(self.dir_cam / 'Toolpath_JSON')
    def get_final_cam_prt_dir(self) -> Path: return self._get_dir(self.dir_cam / 'Final_CAM_PRT')
    
    # Step 11 (Roughing)
    def get_cam_roughing_json_dir(self) -> Path: return self._get_dir(self.dir_cam / 'Roughing_JSON')
    def get_cam_roughing_prt_dir(self) -> Path:  return self._get_dir(self.dir_cam / 'Roughing_PRT')

    # Engineering Order (Step 14/15)
    def get_engineering_order_root(self) -> Path: return self._get_dir(self.dir_cam / 'Engineering_Order_Data')
    def get_eo_txt_dir(self) -> Path:   return self._get_dir(self.get_engineering_order_root() / '工件信息TXT')
    def get_eo_dims_dir(self) -> Path:  return self._get_dir(self.get_engineering_order_root() / '尺寸信息TXT')
    def get_eo_json_dir(self) -> Path:  return self._get_dir(self.get_engineering_order_root() / 'JSON数据')
    
    def get_eo_excel_dir(self) -> Path: return self._get_dir(self.dir_output / 'Engineering_Sheets')
    def get_eo_screenshot_dir(self) -> Path: return self._get_dir(self.get_engineering_order_root() / 'screen-shot')
    def get_nc_output_dir(self) -> Path: return self._get_dir(self.dir_output / 'NC_Code')

    # ==========================================================================
    # 00_Resources (Reports & Excels)
    # ==========================================================================
    def get_csv_data_dir(self) -> Path:    return self._get_dir(self.dir_resources / 'CSV_Reports')
    
    def get_3d_report_csv(self) -> Path:   return self.get_split_prt_dir() / '_Export_Report.csv'
    def get_2d_report_csv(self) -> Path:   return self.get_csv_data_dir() / 'dxf图纸信息汇总表.csv'
    def get_match_result_csv(self) -> Path: return self.get_csv_data_dir() / '数据配对结果.csv'
    
    def get_parts_excel(self) -> Path:       return self.get_csv_data_dir() / '零件参数.xlsx'
    def get_part_params_excel(self) -> Path:  return self.get_csv_data_dir() / '零件参数.xlsx'

    # ==========================================================================
    # External Resources (DLLs, Models, Configs)
    # ==========================================================================
    # DLLs
    def get_dll_path(self, dll_name: str) -> Path: return self.output_dir / dll_name
    def get_face_info_dll_path(self) -> Path:      return Path(config.FILE_DLL_FACE_INFO)
    def get_navigator_dll_path(self) -> Path:      return Path(config.FILE_DLL_NAVIGATOR)
    def get_geometry_analysis_dll_path_20(self) -> Path: return Path(config.FILE_DLL_GEOMETRY_ANALYSIS_20)
    def get_screenshot_dll_path(self) -> Path:     return Path(config.FILE_DLL_SCREENSHOT)
    def get_texture_dll_path(self) -> Path:        return Path(config.FILE_DLL_TEXTURE)

    # Models & Configs
    def get_tool_params_json(self) -> Path:       return self.input_dir / '铣刀参数.json'
    def get_knife_table_json(self) -> Path:       return self.input_dir / 'knife_table.json'
    def get_drill_table_json(self) -> Path:       return self.input_dir / 'drill_table.json'
    def get_tool_params_excel(self) -> Path:      return config.FILE_TOOL_PARAMS_WITH_DRILL
    def get_tool_params_excel_path(self) -> Path: return Path(config.FILE_TOOL_PARAMS_WITH_DRILL)
    
    def get_mill_tools_excel(self) -> Path:       return config.FILE_MILL_TOOLS_EXCEL
    
    def get_rf_model_path(self) -> Path:     return Path(config.FILE_MODEL_RF)
    def get_scaler_path(self) -> Path:       return Path(config.FILE_MODEL_SCALER)
    def get_model_config_path(self) -> Path: return Path(config.FILE_MODEL_CONFIG)
    def get_point_cloud_lib_dir(self) -> Path: return self.input_dir

    # ==========================================================================
    # Helper Methods
    # ==========================================================================
    def get_json_output_path(self, prt_name: str, json_type: str) -> Path:
        type_map = {
            'cavity': '行腔', 'zlevel': '往复等高', 
            'cam': '爬面', 'face': '面铣', 'spiral': '螺旋'
        }
        filename = f"{prt_name}_{type_map.get(json_type, json_type)}.json"
        return self.get_cam_json_dir() / filename

    # ==========================================================================
    # Temporary Files (.temp)
    # ==========================================================================
    def get_temp_dir(self) -> Path: return self._get_dir(self.temp_dir)
    def get_stl_dir(self) -> Path:  return self._get_dir(self.temp_dir / 'stl')
    def get_pcd_dir(self) -> Path:  return self._get_dir(self.temp_dir / 'pcd')
    def get_geometry_analysis_temp_dir(self) -> Path: return self._get_dir(self.temp_dir / 'Geometry_Analysis_Temp')
    
    def get_stl_path(self, name: str) -> Path: return self.get_stl_dir() / f"{name}.stl"
    def get_pcd_path(self, name: str) -> Path: return self.get_pcd_dir() / f"{name}.pcd"

    # ==========================================================================
    # Accessors (Getters)
    # ==========================================================================
    def get_input_3d_prt(self) -> Path: return self.input_3d_prt
    def get_input_2d_dxf(self) -> Path: return self.input_2d_dxf

# Global Instance
_path_manager: Optional[PathManager] = None

def init_path_manager(input_3d_prt: Optional[str] = None, input_2d_dxf: Optional[str] = None) -> PathManager:
    global _path_manager
    _path_manager = PathManager(input_3d_prt, input_2d_dxf)
    _path_manager.ensure_dirs()
    return _path_manager

def get_path_manager() -> PathManager:
    global _path_manager
    if _path_manager is None:
        raise RuntimeError("PathManager not initialized")
    return _path_manager