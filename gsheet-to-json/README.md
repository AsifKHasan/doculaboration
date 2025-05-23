# gsheet-to-json

Generates json output from gsheet.

You will need a ```credential.json``` in ```conf``` which is not in the repo and should never be. Get your local copy and **never** commit it to repo

copy ```conf/config.yml.dist``` as ```conf/config.yml``` and do not commit the copied file

## software and tools required

You will need Poppler
- for Windows - ~https://blog.alivate.com.au/poppler-windows/~. https://github.com/oschwartz10612/poppler-windows/releases/tag/v24.08.0-0
- for Linux - https://poppler.freedesktop.org/

## Linux usage:
```
DOCULABORATION_BASE="/home/asif/projects/asif@github/doculaboration"
cd ${DOCULABORATION_BASE}
cd ./gsheet-to-json/src
./json-from-gsheet.py --config '../conf/config.yml'
```

## Windows usage:
```
set DOCULABORATION_BASE="D:\projects\asif@github\doculaboration"
cd %DOCULABORATION_BASE%
cd ./gsheet-to-json/src
python json-from-gsheet.py --config "../conf/config.yml"
```
