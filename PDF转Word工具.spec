# -*- mode: python ; coding: utf-8 -*-
"""
PDF转Word工具 - 瘦身版打包配置
优化措施：
1. 排除 OpenCV 人脸检测模型 (cv2.data)
2. 排除 OpenCV 视频IO模块 (opencv_videoio_ffmpeg)
3. 排除不需要的 Python 标准库模块
4. 启用字节码优化
"""
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = []

# 只收集必要的模块
tmp_ret = collect_all('pdf2docx')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pypdfium2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# 排除不需要的模块
excludes = [
    # OpenCV 不需要的子模块 (方案1)
    'cv2.data',
    # 标准库测试和开发模块 (方案2)
    'tkinter.test',
    'unittest',
    'pydoc',
    'doctest',
    'test',
    'tests',
    'distutils',
    'setuptools',
    'pip',
    # 其他不需要的模块
    'IPython',
    'jupyter',
    'notebook',
    'sphinx',
    'pydoc_data',
]

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=2,  # Python字节码优化级别
)

# 过滤掉不需要的二进制文件 (OpenCV视频模块)
a.binaries = [b for b in a.binaries if 'opencv_videoio_ffmpeg' not in b[0]]
# 过滤掉OpenCV人脸检测模型数据 (但保留config.py)
a.datas = [d for d in a.datas if 'cv2/data' not in d[0] and 'haarcascade' not in d[0]]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='PDF转Word工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)