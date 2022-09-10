:: gsheet->json->context->pdf pipeline

@echo off

:: parameters
set DOCUMENT=%1

:: json-from-gsheet
pushd .\gsheet-to-json\src
@REM .\json-from-gsheet.py --config "../conf/config.yml" --gsheet "%DOCUMENT%"

if errorlevel 1 (
  popd
  exit /b %errorlevel%
)

popd

:: context-from-json
pushd .\json-to-context\src
.\context-from-json.py --config "../conf/config.yml" --json "%DOCUMENT%"

if errorlevel 1 (
  popd
  exit /b %errorlevel%
)

popd

:: context -> pdf
pushd .\out
ptime context --run %DOCUMENT%.tex
@REM move %DOCUMENT%.pdf %DOCUMENT%.tex.pdf

if errorlevel 1 (
  popd
  exit /b %errorlevel%
)

popd
