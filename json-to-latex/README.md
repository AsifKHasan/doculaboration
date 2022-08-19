# json-to-latex

Generates LaTex from json generated through **gsheet-to-json**

copy ```conf/config.yml.dist``` as ```conf/config.yml``` and do not commit the copied file

## LaTex generation
### Toolchain
1. TexLive from https://tug.org/texlive/
2. LuaLatex
3. We need fonts
 * Google Noto fonts from https://github.com/google/fonts
 ```
 sudo apt install fonts-noto
 ```
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

### Windows usage:
cd to doculaboration root directory and run
```pdf-from-gsheet.bat name-of-the-gsheet```
