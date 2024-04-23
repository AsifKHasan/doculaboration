:: gsheet->json->odt pipeline

@echo off
REM run libre office as a service
start "" "C:\Program Files\LibreOffice\program\soffice.exe" --headless --invisible --accept="socket,host=localhost,port=8100;urp;"
timeout /t 2 /nobreak > nul

:: parameters
set DOCUMENT=%1

:: json-from-gsheet
pushd .\gsheet-to-json\src
python json-from-gsheet.py --config "../conf/config.yml" --gsheet "%DOCUMENT%"

if errorlevel 1 (
  popd
  exit /b %errorlevel%
)

popd

:: odt-from-json
pushd .\json-to-odt\src
python odt-from-json.py --config "../conf/config.yml" --json "%DOCUMENT%"

if errorlevel 1 (
  popd
  exit /b %errorlevel%
)

popd
