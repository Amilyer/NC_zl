import NXOpen
import math
import NXOpen.UF


def print_to_info_window(message):
    """输出信息到 NX 信息窗口"""
    session = NXOpen.Session.GetSession()
    session.ListingWindow.Open()
    session.ListingWindow.WriteLine(str(message))
    try:
        session.LogFile.WriteLine(str(message))
    except Exception:
        pass

if __name__ == '__main__':
    tag = 62528
    uf_session = NXOpen.UF.UFSession.GetUFSession()
    arc_data = uf_session.Curve.AskArcData(tag)

    # 获取 StartAngle 和 EndAngle
    start_angle = arc_data.StartAngle
    end_angle = arc_data.EndAngle

    # 起点
    sp = NXOpen.Point3d(
        arc_data.ArcCenter[0] + arc_data.Radius * math.cos(start_angle),
        arc_data.ArcCenter[1] + arc_data.Radius * math.sin(start_angle),
        arc_data.ArcCenter[2]
    )

    # 终点
    ep = NXOpen.Point3d(
        arc_data.ArcCenter[0] + arc_data.Radius * math.cos(end_angle),
        arc_data.ArcCenter[1] + arc_data.Radius * math.sin(end_angle),
        arc_data.ArcCenter[2]
    )

    print_to_info_window(f"sp:{sp},ep:{ep}")