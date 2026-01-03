class NXError(Exception):
    """NX 自动化基础异常（带错误码）"""
    def __init__(self, code, message):
        self.code = code
        super().__init__(f"{code} {message}")


class InfoInvalidError(NXError):
    def __init__(self, message="提取加工说明信息失败"):
        super().__init__("NX01", message)


class MaterialInvalidError(NXError):
    def __init__(self, message="提取材质信息失败"):
        super().__init__("NX02", message)


class SizeInvalidError(NXError):
    def __init__(self, message="零件尺寸提取失败"):
        super().__init__("NX03", message)


class BoundLineInvalidError(NXError):
    def __init__(self, message="缺少板料线或加工00标记"):
        super().__init__("NX04", message)
