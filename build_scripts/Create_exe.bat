REM Change to the development directory
cd c:\Users\shalsh1\Downloads\My_Tools\Excel comparator\Development\

REM Build the Python script into a standalone executable using PyInstaller
c:\Data\MKS\BB400_POWERPACK\40_SW\90_Tools\Python311\python.exe -m PyInstaller --onefile .\src\excel_comparator.py

REM Copy the generated executable to the target directory
copy .\dist\excel_comparator.exe .\..\excel_comparator.exe /Y

REM Copy the configuration file to the target directory to be checked into PTC
@REM copy .\config.ini c:\MKS\DAG_FP48V\40_SW\55_Tool_Saved_Configurations\25_ET_Tools\DTC_ET_Decoder\config.ini /Y

REM Delete the PyInstaller spec file
del .\*.spec

REM Delete the build directory and all its contents
rmdir .\build /S /Q
rmdir .\dist /S /Q