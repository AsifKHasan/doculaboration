# json-to-docx

Generates docx (WordML) from json generated through **gsheet-to-json**

copy ```conf/config.yml.dist``` as ```conf/config.yml``` and do not commit the copied file

## software and tools required

1.  Python 3.8 or higher - <https://www.python.org/downloads/>
2.  Git -  <https://git-scm.com/downloads>
3.  MS Office. You can work without MS Office installed on Linux and Windows both, but you will not be able to generate indexes and pdf from the generated docx

## Windows usage:
you can run the windows command script *docx-from-gsheet.bat* in the root folder. It takes only one parameter - the name of the gsheet

```docx-from-gsheet.bat name-of-the-gsheet```

or you can run the python script this way (NOT PREFERRED)
```
set DOCULABORATION_BASE="D:\projects\asif@github\doculaboration"
cd %DOCULABORATION_BASE%
cd ./json-to-docx/src
python docx-from-json.py --config "../conf/config.yml"
```

## Linux usage:
you can run the bash script *docx-from-gsheet.sh* in the root folder. It takes only one parameter - the name of the gsheet

```./docx-from-gsheet.sh name-of-the-gsheet```

or you can run the python script this way (NOT PREFERRED)
```
DOCULABORATION_BASE="/home/asif/projects/asif@github/doculaboration"
cd ${DOCULABORATION_BASE}
cd ./json-to-docx/src
./docx-from-json.py --config '../conf/config.yml'
```
