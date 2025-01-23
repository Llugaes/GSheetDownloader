
import sys
import os

def _pyi_rth_encodings():
    import encodings
    import encodings.aliases
    import encodings.utf_8
    import encodings.ascii
    import encodings.latin_1
    import encodings.cp1252
    import encodings.utf_16
    import encodings.utf_32

# 初始化编码
_pyi_rth_encodings()

# 确保Python能找到基础库
if hasattr(sys, '_MEIPASS'):
    if sys._MEIPASS not in sys.path:
        sys.path.insert(0, sys._MEIPASS)
    
    # 添加 lib-dynload 路径
    lib_dynload = os.path.join(sys._MEIPASS, 'lib-dynload')
    if os.path.exists(lib_dynload) and lib_dynload not in sys.path:
        sys.path.insert(1, lib_dynload)
        
    # 添加 site-packages 路径
    site_packages = os.path.join(sys._MEIPASS, 'site-packages')
    if os.path.exists(site_packages) and site_packages not in sys.path:
        sys.path.append(site_packages)
