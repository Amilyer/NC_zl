
import os
import sys

# Setup paths
current_dir = os.path.dirname(os.path.abspath(__file__))
core_dir = os.path.join(current_dir, "core", "File preprocessing")
if core_dir not in sys.path:
    sys.path.insert(0, core_dir)

try:
    import config
    from dll_loader import UniversalLoader
except ImportError as e:
    print(f"Error importing modules: {e}")
    sys.exit(1)

def inspect_dll(dll_path):
    print(f"\n{'='*50}")
    print(f"Inspecting DLL: {dll_path}")
    print(f"{'='*50}")
    
    if not os.path.exists(dll_path):
        print("❌ DLL file not found!")
        return

    try:
        loader = UniversalLoader(dll_path)
        print("\n✅ DLL Loaded Successfully")
        print(f"Functions found: {len(loader.functions)}")
        
        for func_name, info in loader.functions.items():
            print(f"\nFunction: {func_name}")
            print("-" * 30)
            print("  Parameters:")
            params = info.get('params', [])
            for i, p in enumerate(params):
                # type 4 usually means string/char*, 1 means int
                p_type = p.get('type')
                type_str = "String (char*)" if p_type == 4 else "Int" if p_type == 1 else f"Unknown({p_type})"
                print(f"    {i+1}. Name: '{p['name']}' | Type: {type_str} (Code: {p_type})")
            
            print("\n  Raw Meta:")
            print(f"    {info}")
            
    except Exception as e:
        print(f"❌ Error loading/inspecting DLL: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    # Check both DLLs
    
    # 1. Geometry Analysis 1
    dll1 = config.FILE_DLL_GEOMETRY_ANALYSIS_1
    inspect_dll(str(dll1))

    # 2. Geometry Analysis 20
    dll20 = config.FILE_DLL_GEOMETRY_ANALYSIS_20
    inspect_dll(str(dll20))
