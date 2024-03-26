# json-to-odt

Generates Openoffice odt from json generated through **gsheet-to-json**

copy ```conf/config.yml.dist``` as ```conf/config.yml``` and do not commit the copied file

## software and tools required

1.  Python 3.8 or higher - <https://www.python.org/downloads/>
2.  Git -  <https://git-scm.com/downloads>
3.  LibreOffice - <https://www.libreoffice.org/download/download/>. You can work without LibreOffice installed on Linux and Windows both, but you will not be able to generate indexes and pdf from the generated odt
    - On Windows machines, make sure to check ptyhon in custom installation setup
    - On Linux machines, run the following command to install LibreOffice python executable
        ```bash
        sudo apt install libreoffice-script-provider-python
        ```
        If you find a tough time where the embeded LibreOffice Python is installed, read [this issue](https://github.com/unoconv/unoconv/issues/49#issuecomment-897715262)

## Useful links
some useful links for json-to-odt
* https://mashupguide.net/1.0/html/ch17s04.xhtml
* https://github.com/eea/odfpy/tree/master/examples
* https://glot.io/snippets/f0nuzv7b8k

## Linux usage:
you can run the bash script *odt-from-gsheet.sh* in the root folder. It takes only one parameter - the name of the gsheet

```./odt-from-gsheet.sh name-of-the-gsheet```

or you can run the python script this way (NOT PREFERRED)
```
DOCULABORATION_BASE="/home/asif/projects/asif@github/doculaboration"
cd ${DOCULABORATION_BASE}
cd ./json-to-odt/src
./odt-from-json.py --config '../conf/config.yml'
```

## Windows usage:
you can run the windows command script *odt-from-gsheet.bat* in the root folder. It takes only one parameter - the name of the gsheet

```odt-from-gsheet.bat name-of-the-gsheet```

or you can run the python script this way (NOT PREFERRED)
```
set DOCULABORATION_BASE="D:\projects\asif@github\doculaboration"
cd %DOCULABORATION_BASE%
cd ./json-to-odt/src
python odt-from-json.py --config "../conf/config.yml"
```
