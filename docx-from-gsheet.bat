:: gsheet->json->docx pipeline

@echo off

:: parameters
set DOCUMENT="MoL-LSG__eoi-and-tor"
set DOCULABORATION_BASE="d:/projects/asif@github/doculaboration"

cd %DOCULABORATION_BASE%

:: json-from-gsheet
pushd .\gsheet-to-json\src
.\json-from-gsheet.py --config "../conf/config.yml" --gsheet "%DOCUMENT%"

if errorlevel 1 (
  popd
  exit /b %errorlevel%
)

popd

:: docx-from-json
pushd .\json-to-docx\src
.\docx-from-json.py --config "../conf/config.yml" --json "%DOCUMENT%"

if errorlevel 1 (
  popd
  exit /b %errorlevel%
)

popd
