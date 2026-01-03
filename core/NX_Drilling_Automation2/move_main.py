
from utils import print_to_info_window, handle_exception, point_from_angle, analyze_arc
from geometry import GeometryHandler
from process_info import ProcessInfoHandler
from parameter_parser import ParameterParser
import drill_config

 
def move_to_origin(session, work_part, drill_path, knife_path):
    drill_config.DRILL_JSON_PATH = drill_path
    drill_config.MILL_JSON_PATH = knife_path
    geometry_handler = GeometryHandler(session, work_part)
    process_info_handler = ProcessInfoHandler()
    parameter_parser = ParameterParser()
 # 第一步：预处理，提取加工信息
    print_to_info_window("第一步：预处理，提取加工信息")
    annotated_data = process_info_handler.extract_and_process_notes(work_part)
    processed_result = parameter_parser.process_hole_data(annotated_data)
    annotated_data = annotated_data["主体加工说明"]
    if not processed_result:
        return handle_exception("加工信息处理失败，流程终止")

    # 获取材料类型
    if annotated_data["材质"]:
        material_type = process_info_handler.get_material(annotated_data["材质"][0])
    else:
        print_to_info_window("零件材质提取失败")
        return False

    # 获取毛坯尺寸
    if annotated_data["尺寸"]:
        lwh_point = process_info_handler.get_workpiece_dimensions(annotated_data["尺寸"])
    else:
        print_to_info_window("零件尺寸提取失败")
        return False
    # 获取加工坐标原点
    original_point = geometry_handler.get_start_point(lwh_point)
    try:
        if original_point[0] - 1e-7 != 0.0 or original_point[1] - 1e-7 != 0.0:
            # 移动2D图到绝对坐标系
            try:
                geometry_handler.move_objects_point_to_point(original_point,drill_config.DEFAULT_ORIGIN_POINT)
            except:
                print_to_info_window("2D图已在加工坐标原点，无需移动")
    except:
        print_to_info_window("获取加工坐标原点失败--2D图无(0, 0)点")
        return None