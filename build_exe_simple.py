"""
Simplified build script for Anthrilo Label Generator - excludes unnecessary heavy dependencies
"""
import PyInstaller.__main__
import os
import shutil

# Clean previous builds
if os.path.exists('build'):
    shutil.rmtree('build')
if os.path.exists('dist'):
    shutil.rmtree('dist')

# Build executable with minimal dependencies
args = [
    'label_generator_gui.py',
    '--name=AnthriloLabelGenerator',
    '--onefile',
    '--windowed',
    '--icon=NONE',
    '--noconsole',
    
    # Essential imports only
    '--hidden-import=PIL.Image',
    '--hidden-import=PIL.ImageDraw',
    '--hidden-import=PIL.ImageFont',
    '--hidden-import=PIL._tkinter_finder',
    '--hidden-import=tkinter',
    '--hidden-import=tkinter.filedialog',
    '--hidden-import=tkinter.messagebox',
    '--hidden-import=tkinter.ttk',
    '--hidden-import=pandas',
    '--hidden-import=openpyxl',
    '--hidden-import=tempfile',
    '--hidden-import=pathlib',
    
    # Barcode module
    '--collect-all=barcode',
    
    # Exclude heavy unused packages
    '--exclude-module=scipy',
    '--exclude-module=torch',
    '--exclude-module=tensorflow',
    '--exclude-module=matplotlib',
    '--exclude-module=IPython',
    '--exclude-module=jupyter',
    '--exclude-module=pytest',
    '--exclude-module=notebook',
    '--exclude-module=zmq',
    '--exclude-module=PySide6',
    '--exclude-module=PyQt5',
    '--exclude-module=PyQt6',
]

# Add fonts if they exist
font_paths = [
    'C:\\Windows\\Fonts\\arial.ttf',
    'C:\\Windows\\Fonts\\arialbd.ttf'
]
for font_path in font_paths:
    if os.path.exists(font_path):
        font_name = os.path.basename(font_path)
        args.append(f'--add-data={font_path};.')
        print(f"[+] Adding font: {font_name}")

print("\n" + "="*60)
print("Building Anthrilo Label Generator (Simplified Build)")
print("="*60 + "\n")

PyInstaller.__main__.run(args)

# Move executable to APP folder
os.makedirs('APP', exist_ok=True)
if os.path.exists('dist/AnthriloLabelGenerator.exe'):
    if os.path.exists('APP/AnthriloLabelGenerator.exe'):
        os.remove('APP/AnthriloLabelGenerator.exe')
    shutil.move('dist/AnthriloLabelGenerator.exe', 'APP/AnthriloLabelGenerator.exe')
    
    file_size = os.path.getsize('APP/AnthriloLabelGenerator.exe') / (1024 * 1024)
    
    print("\n" + "="*60)
    print("[OK] BUILD SUCCESSFUL!")
    print("="*60)
    print(f"[OK] File: APP\\AnthriloLabelGenerator.exe")
    print(f"[OK] Size: {file_size:.2f} MB")
    print("\nTo run: .\\APP\\AnthriloLabelGenerator.exe")
    print("="*60)
else:
    print("\n" + "="*60)
    print("[ERROR] BUILD FAILED!")
    print("="*60)

# Clean up build files
if os.path.exists('build'):
    shutil.rmtree('build')
if os.path.exists('dist'):
    shutil.rmtree('dist')
if os.path.exists('AnthriloLabelGenerator.spec'):
    os.remove('AnthriloLabelGenerator.spec')

print("\n[OK] Cleanup complete!")

