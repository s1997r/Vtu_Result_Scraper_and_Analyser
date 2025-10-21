block_cipher = None

a = Analysis(
    ['vtu_gui.py'],  # Main entry point
    pathex=[],
    binaries=[],
    datas=[
        ('Tesseract-OCR', 'Tesseract-OCR'),  # For captcha handling
        ('icon.ico', '.'),  # Application icon
        ('analyzer.py', '.'),  # Include analyzer module
        ('student_data.py', '.'),  # Include student data parser
        ('vtu_marks_scraper.py', '.'),  # Include scraper module
        ('captcha_handler.py', '.')  # Include captcha handler
    ],
    hiddenimports=[
        'pytesseract',
        'webdriver_manager',
        'bs4',
        'selenium',
        'PIL',
        'pandas',
        'openpyxl',  # Needed for pandas Excel support
        'docx',  # For report generation
        'lxml',  # For docx table handling
        'numpy',  # pandas dependency
        'python-docx',  # Explicit docx import
        'dateutil',  # pandas dependency
        'pytz',  # pandas dependency
        'six',  # compatibility package
        'cryptography',  # selenium dependency
        'urllib3',  # selenium dependency
        'selenium.webdriver.common',
        'selenium.webdriver.chrome',
        'selenium.webdriver.support'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='VTU_Scraper',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to True for debugging, False for production
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # Application icon
)

# Additional build options for one-file executable
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='VTU_Scraper'
)

