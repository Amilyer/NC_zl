import importlib.util
import os
import sys


def test_import():
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(current_dir, "生成爬面文件.py")
        print(f"Testing import of: {file_path}")
        
        module_name = "test_module"
        spec = importlib.util.spec_from_file_location(module_name, file_path)
        module = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
        print("Import successful!")
    except Exception:
        print("Import failed!")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_import()
