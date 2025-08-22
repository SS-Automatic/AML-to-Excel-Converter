@echo off
echo Собираю EXE файл...
pyinstaller --onefile --name "AML_Converter" aml_converter.py
echo Готово! EXE файл в папке 'dist'
pause