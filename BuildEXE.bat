CLS
@echo off

set app_name=ECRev
set main_dir=%~dp0
set pyinstaller_dir=S:\Everyone\Management Software\pyinstaller-2.0\utils
set build_options=
set python=C:\Python27\python.exe



REM set makespec_options=--onefile --console --out="%main_dir%"
REM %python% "%pyinstaller_dir%\Makespec.py" %makespec_options% "%main_dir%\%app_name%.py"
REM @echo -----------------------------------
@echo on


%python% -O "%pyinstaller_dir%\Build.py" %build_options% "%main_dir%\%app_name%.spec"
@echo -----------------------------------


REM MOVE /Y "%main_dir%\dist\%app_name%.exe" "%main_dir%"
rmdir /s /q "%main_dir%\build"


@PAUSE