# doculaboration - Documentation Collaboration
CRDT (gsheet, etc.) based documentation collaboration pipeline to generate editable (docx, odt, etc.) and printable (pdf, etc.) output from input data

* *gsheet-to-json* is for generating json output from gsheet data. The output json is meant to be fed as input for document generation components of the pipeline
* *json-to-docx* is for generating docx (WordML) documents for editing
* *json-to-pandoc* is for generating pandoc markdown for generating editable/printable outputs

## (optional) update all python packages
```
pip list --outdated --format=freeze | grep -v '^\-e' | cut -d = -f 1  | xargs -n1 pip install -U
```

## Toolchain
1. Pandoc from https://pandoc.org/
2. TexLive from https://tug.org/texlive/
3. LuaLatex
4. We need fonts
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

## Linux usage:
```
DOCUMENT="name-of-the-gsheet"
DOCULABORATION_BASE="/home/asif/projects/asif@github/doculaboration"
cd ${DOCULABORATION_BASE}
cd ./out
time pandoc ${DOCUMENT}.mkd ../json-to-pandoc/conf/preamble.yml -s --pdf-engine=lualatex -f markdown -t latex -o ${DOCUMENT}.pdf
```

## Windows usage:
```
set DOCUMENT="name-of-the-gsheet"
set DOCULABORATION_BASE="D:\projects\asif@github\doculaboration"
cd %DOCULABORATION_BASE%
cd ./out
pandoc %DOCUMENT%.mkd ../json-to-pandoc/conf/preamble.yml -s --pdf-engine=lualatex -f markdown -t latex -o %DOCUMENT%.pdf
```
