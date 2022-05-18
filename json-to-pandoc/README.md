# json-to-pandoc

Generates Pandoc markdown from json generated through **gsheet-to-json**

copy ```conf/config.yml.dist``` as ```conf/config.yml``` and do not commit the copied file

## Pandoc/LaTex generation
### Toolchain
1. Pandoc from https://pandoc.org/
2. TexLive from https://tug.org/texlive/
3. LuaLatex
4. We need fonts
 * Google Noto fonts from https://github.com/google/fonts
 * GNU free fonts from https://www.gnu.org/software/freefont/
 * Go Smallcaps font from https://www.fontmirror.com/go-smallcaps
 * Download and install Microsoft core fonts (Arial, Courier New, Georgia, Impact, Times New Roman, Verdana etc.)
 ```
 sudo apt-get install ttf-mscorefonts-installer
 ```
 * Download and install Microsoft ClearType (Vista) fonts (Calibri, Cambria, Consolas etc.)
 ```
 wget -qO- http://plasmasturm.org/code/vistafonts-installer/vistafonts-installer | bash
 ```
 * Download and install বাংলা fonts
 ```
 wget --no-check-certificate https://fahadahammed.com/extras/fonts/font.sh -O font.sh;chmod +x font.sh;bash font.sh;rm font.sh
 ```

### Linux usage:
cd to doculaboration root directory and run
```./pdf-from-gsheet.sh name-of-the-gsheet```

or (NOT PREFERRED)

```
DOCUMENT="name-of-the-gsheet"
DOCULABORATION_BASE="/home/asif/projects/asif@github/doculaboration"
cd ${DOCULABORATION_BASE}
cd ./out
time pandoc ${DOCUMENT}.mkd ../json-to-pandoc/conf/preamble.yml -s --pdf-engine=lualatex -f markdown -t latex -o ${DOCUMENT}.pdf
```

### Windows usage:
cd to doculaboration root directory and run
```pdf-from-gsheet.bat name-of-the-gsheet```

or (NOT PREFERRED)

```
set DOCUMENT="name-of-the-gsheet"
set DOCULABORATION_BASE="D:\projects\asif@github\doculaboration"
cd %DOCULABORATION_BASE%
cd ./out
ptime pandoc %DOCUMENT%.mkd ../json-to-pandoc/conf/preamble.yml -s --pdf-engine=lualatex -f markdown -t latex -o %DOCUMENT%.pdf
```
