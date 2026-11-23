@echo off
setlocal
cd /d "%~dp0" || exit /b 1
set "BUILD_PYTHON=.venv-windows\Scripts\python.exe"
if not exist "%BUILD_PYTHON%" (
    py -3.12 -m venv .venv-windows || exit /b 1
)
"%BUILD_PYTHON%" -m pip install -r requirements.txt PyInstaller==6.16.0 || exit /b 1
"%BUILD_PYTHON%" -m unittest discover -s tests -v || exit /b 1
"%BUILD_PYTHON%" -m PyInstaller --noconfirm --clean --onefile --windowed --name ContractManager-v7 app.py || exit /b 1
echo dist\ContractManager-v7.exe
