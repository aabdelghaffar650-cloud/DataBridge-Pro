# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_submodules, collect_data_files, copy_metadata

# DataBridge backend bundle for Tauri/PyInstaller.
# Keep templates beside app.py in development; PyInstaller will bundle them here.
datas = [
    ('app.py', '.'),
    ('modules', 'modules'),
    ('templates', 'templates'),
]
binaries = []
hiddenimports = ['streamlit']

# Packages that need their full submodule tree + data files + binary
# extensions collected. This covers python-docx, PDF/report generation,
# Excel, and charting/image-export libraries used by the export tabs
# (Word / PDF / Excel / visuals / donor / M&E / UNDP reports).
COLLECT_ALL_PACKAGES = [
    'streamlit',
    'docx',        # python-docx
    'reportlab',
    'openpyxl',
    'plotly',
    'kaleido',
    'matplotlib',
    'PIL',
    'lxml',
]

for pkg in COLLECT_ALL_PACKAGES:
    tmp = collect_all(pkg)
    datas += tmp[0]
    binaries += tmp[1]
    hiddenimports += tmp[2]

# Explicit hidden imports for lazy/dynamic imports inside app.py and
# modules/reports.py that PyInstaller's static analysis can miss because
# they are imported inside functions/try-except blocks rather than at
# module top-level.
hiddenimports += [
    # python-docx
    'docx',
    'docx.document',
    'docx.shared',
    'docx.enum.text',
    'docx.enum.table',
    'docx.oxml',
    'docx.oxml.ns',
    'docx.oxml.shared',
    'docx.table',
    'docx.text.paragraph',
    # reportlab (PDF export)
    'reportlab',
    'reportlab.platypus',
    'reportlab.lib',
    'reportlab.lib.pagesizes',
    'reportlab.lib.styles',
    'reportlab.lib.units',
    'reportlab.pdfgen',
    'reportlab.pdfgen.canvas',
    # openpyxl (Excel export)
    'openpyxl',
    'openpyxl.workbook',
    'openpyxl.styles',
    # matplotlib (PDF/chart rasterization)
    'matplotlib',
    'matplotlib.pyplot',
    'matplotlib.backends.backend_agg',
    # plotly / kaleido (visuals export)
    'plotly',
    'plotly.graph_objects',
    'plotly.express',
    'kaleido',
    # Pillow (image handling used by docx/reportlab)
    'PIL',
    'PIL.Image',
    # lxml (python-docx XML backend)
    'lxml',
    'lxml.etree',
]

# Explicit submodule collection (belt-and-suspenders on top of collect_all)
for pkg in ['docx', 'reportlab', 'openpyxl', 'plotly', 'kaleido', 'matplotlib', 'PIL', 'lxml']:
    hiddenimports += collect_submodules(pkg)

# Package metadata that some libraries (esp. streamlit, altair, pandas,
# numpy) look up at runtime via importlib.metadata. Missing metadata
# causes "No module named X" / version-lookup errors even when the code
# itself is bundled.
for pkg in [
    'streamlit',
    'python-docx',
    'reportlab',
    'openpyxl',
    'plotly',
    'kaleido',
    'matplotlib',
    'Pillow',
    'altair',
    'pandas',
    'numpy',
]:
    try:
        datas += copy_metadata(pkg)
    except Exception:
        pass

# Extra data files (non-Python resources) for packages that ship them.
for pkg in ['altair', 'pandas']:
    try:
        datas += collect_data_files(pkg)
    except Exception:
        pass

# Deduplicate hiddenimports while preserving determinism.
hiddenimports = sorted(set(hiddenimports))


a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='databridge-backend',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
