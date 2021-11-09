# json-to-docx

Generates docx (WordML) from json generated through **gsheet-to-json**

copy ```conf/config.yml.dist``` as ```conf/config.yml``` and do not commit the copied file

## Linux usage:
```
DOCULABORATION_BASE="/home/asif/projects/asif@github/doculaboration"
cd ${DOCULABORATION_BASE}
cd ./json-to-docx/src
./docx-from-json.py --config '../conf/config.yml'
```

## Windows usage:
```
set DOCULABORATION_BASE="D:\projects\asif@github\doculaboration"
cd %DOCULABORATION_BASE%
cd ./json-to-docx/src
python docx-from-json.py --config "../conf/config.yml"
```
