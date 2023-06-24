@echo off

setlocal

REM Loading .env variables
for /F "tokens=1* delims==" %%i in (.env) do (
    if "%%i"=="PYTHON" (
        set "%%i=%%j"
    )
)

REM Creating virtual environment
if not exist "venv" (
    echo Creating virtual environment using Python at %PYTHON%...
    %PYTHON% -m venv venv
) else (
    echo Virtual environment already exists.
)

REM Activating the virtual environment
call venv\Scripts\activate

REM Installing the requirements
echo Installing the Python requirements...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo Installation completed!

endlocal
