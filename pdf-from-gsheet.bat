:: gsheet->json->pandoc->pdf pipeline

@echo off

:: parameters
set DOCUMENT="MoL-DLRMS__eoi-and-tor"
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

:: pandoc-from-json
pushd .\json-to-pandoc\src
.\pandoc-from-json.py --config "../conf/config.yml" --json "%DOCUMENT%"

if errorlevel 1 (
  popd
  exit /b %errorlevel%
)

popd

:: pandoc-from-json
pushd .\out
ptime pandoc %DOCUMENT%.mkd ..\json-to-pandoc\conf\preamble.yml -s --pdf-engine=lualatex -f markdown -t latex -o %DOCUMENT%.pdf

if errorlevel 1 (
  popd
  exit /b %errorlevel%
)

popd
