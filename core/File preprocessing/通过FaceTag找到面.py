# -*- coding: utf-8 -*-
import NXOpen


class FaceHighlighter:
    def __init__(self):
        self.theSession = NXOpen.Session.GetSession()
        self.workPart = self.theSession.Parts.Work
        self.lw = self.theSession.ListingWindow
        self.lw.Open()

    def print_log(self, msg, level="INFO"):
        line = f"[{level}] {msg}"
        self.lw.WriteLine(line)
        self.theSession.LogFile.WriteLine(line)

    def get_face_user_attr(self, face, attr_title):
        """读取面上的用户属性，找不到返回 None"""
        try:
            for a in face.GetUserAttributes():
                if a.Title == attr_title:
                    if a.Type == NXOpen.NXObject.AttributeType.String:
                        return a.StringValue
                    elif a.Type == NXOpen.NXObject.AttributeType.Integer:
                        return a.IntegerValue
                    elif a.Type == NXOpen.NXObject.AttributeType.Real:
                        return a.RealValue
                    elif a.Type == NXOpen.NXObject.AttributeType.Boolean:
                        return a.BooleanValue
                    else:
                        return a.Value
            return None
        except Exception as e:
            self.print_log(f"读取属性失败 FaceTag={face.Tag}, attr={attr_title}: {e}", "ERROR")
            return None

    def find_faces_by_face_tag(self, face_tag_values, attr_title="FACE_TAG"):
        """
        你的输入就是 FACE_TAG 属性值列表：
          face_tag_values: [62883, 62837, ...] 或 ["62883", "62837", ...]
        返回匹配到的 Face 列表
        """
        targets = set(str(x).strip() for x in face_tag_values)

        found = []
        for body in self.workPart.Bodies:
            for face in body.GetFaces():
                v = self.get_face_user_attr(face, attr_title)
                if v is None:
                    continue
                if str(v).strip() in targets:
                    found.append(face)

        return found

    def highlight_by_face_tag(self, face_tag_values, attr_title="FACE_TAG"):
        """按 FACE_TAG 属性值高亮对应面（不改颜色）"""
        if not face_tag_values:
            self.print_log("FACE_TAG 输入为空", "WARN")
            return []

        self.print_log(f"按 {attr_title} 查找面: {face_tag_values}", "DEBUG")
        faces = self.find_faces_by_face_tag(face_tag_values, attr_title)

        if not faces:
            self.print_log(f"未找到任何面, {attr_title}={face_tag_values}", "WARN")
            return []

        for f in faces:
            try:
                f.Highlight()
                self.print_log(f"高亮面: FaceTag={f.Tag}, {attr_title}={self.get_face_user_attr(f, attr_title)}", "DEBUG")
            except Exception as e:
                self.print_log(f"高亮失败 FaceTag={f.Tag}: {e}", "ERROR")

        self.print_log(f"按 {attr_title} 匹配并高亮 {len(faces)} 个面", "SUCCESS")
        return faces


if __name__ == "__main__":
    tool = FaceHighlighter()

    # ✅ 你的输入就是 FACE_TAG
    face_tag_values = [50841,49179,50774,50816,49160,50847,50817,49164,49181,49010,49015,49020,48001,48012,48016,48027,48033,48037,48043,48045,48050,48054,48059,49032,49033,49034,49037,49038,49040,49041,49046,49049,49052,49079,49080]   
    # 这里写你的 FACE_TAG 列表
    tool.highlight_by_face_tag(face_tag_values, attr_title="FACE_TAG")
