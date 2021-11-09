# gsheet-to-json

Generates json output from gsheet.

You will need a ```credential.json``` in ```conf``` which is not in the repo and should never be. Get your local copy and **never** commit it to repo

copy ```conf/config.yml.dist``` as ```conf/config.yml``` and do not commit the copied file

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
