# json-to-odt

Generates Openoffice odt from json generated through **gsheet-to-json**

copy ```conf/config.yml.dist``` as ```conf/config.yml``` and do not commit the copied file

## Linux usage:
```
DOCULABORATION_BASE="/home/asif/projects/asif@github/doculaboration"
cd ${DOCULABORATION_BASE}
cd ./json-to-odt/src
./odt-from-json.py --config '../conf/config.yml'
```

## Windows usage:
```
set DOCULABORATION_BASE="D:\projects\asif@github\doculaboration"
cd %DOCULABORATION_BASE%
cd ./json-to-odt/src
python odt-from-json.py --config "../conf/config.yml"
```


links for json-to-odt
https://mashupguide.net/1.0/html/ch17s04.xhtml
https://github.com/eea/odfpy/tree/master/examples
https://glot.io/snippets/f0nuzv7b8k
