import NXOpen
import NXOpen.Features
import NXOpen.GeometricUtilities


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
    theSession = NXOpen.Session.GetSession()
    workPart = theSession.Parts.Work
    import os

    print_to_info_window(os.environ.get("UGII_CAM_BASE_DIR"))
    print_to_info_window(os.environ.get("UGII_TOOL_LIBRARY_DIR"))
    print_to_info_window(os.environ.get("UGII_CAM_USER_DIR"))