@echo off
cd /d "%~dp0"
for %%f in (claudius_extra.py claudius_i18n.py claudius_theme.py claudius_lang.py claudius_print.py claudius_features.py claudius_view.py claudius_pro.py claudius_v3.py claudius_v3b.py claudius_text.py claudius_frame.py claudius_quick.py claudius_update.py splash.png FREEWARE.txt LICENSE) do if not exist "%%f" (echo MISSING FILE: %%f - copy ALL files of the program folder next to claudius.py & pause & exit /b 1)
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconsole --onefile --name "Claudius PDF Editor" --icon claudius.ico --splash splash.png --version-file version_info.txt --exclude-module pymupdf.layout --exclude-module pymupdf_layout --exclude-module pymupdf4llm --exclude-module onnxruntime --collect-data docx --collect-data pptx --collect-submodules docx --collect-submodules pptx --collect-submodules openpyxl --collect-all tkinterdnd2 --hidden-import claudius_extra --hidden-import claudius_i18n --hidden-import claudius_theme --hidden-import claudius_lang --hidden-import claudius_print --hidden-import claudius_features --hidden-import claudius_view --hidden-import claudius_pro --hidden-import claudius_v3 --hidden-import claudius_v3b --hidden-import claudius_text --hidden-import claudius_frame --hidden-import claudius_quick --hidden-import claudius_update claudius.py
echo.
echo Done. Your program is in the "dist" folder: dist\Claudius PDF Editor.exe
pause
