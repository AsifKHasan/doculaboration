# json-to-context

Generates ConTeXt from json generated through **gsheet-to-json**

copy ```conf/config.yml.dist``` as ```conf/config.yml``` and do not commit the copied file

## ConTeXt generation
### Toolchain
1. Ruby from https://www.ruby-lang.org/en/downloads/
2. ConTeXt from http://minimals.contextgarden.net/setup/context-setup-win64.zip [optional]
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
```./context-from-gsheet.sh name-of-the-gsheet```

### Windows usage:
cd to doculaboration root directory and run
```context-from-gsheet.bat name-of-the-gsheet```


# Building the font database
export OSFONTDIR=/home/asif/.fonts/lsaBanglaFonts
export OSFONTDIR=/usr/share/fonts/truetype/noto/
mtxrun --script fonts --reload


# if mtxrun has problem finding the mtx-fonts.lua file, it may be necessary to regenerate ConTeXt's file database:
context --generate


# Querying the font database
mtxrun --script fonts --list --all --pattern=*
mtxrun --script fonts --list --info --pattern=*