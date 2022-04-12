# json-to-pandoc

Generates Pandoc markdown from json generated through **gsheet-to-json**

copy ```conf/config.yml.dist``` as ```conf/config.yml``` and do not commit the copied file

## Linux usage:
```
DOCULABORATION_BASE="/home/asif/projects/asif@github/doculaboration"
cd ${DOCULABORATION_BASE}
cd ./json-to-pandoc/src
./pandoc-from-json.py --config '../conf/config.yml'
```

## Windows usage:
```
set DOCULABORATION_BASE="D:\projects\asif@github\doculaboration"
cd %DOCULABORATION_BASE%
cd ./json-to-pandoc/src
python pandoc-from-json.py --config "../conf/config.yml"
```

TODO
1. latex block
2. verbatim block
3. superscript
4. subscript
5. footnote

TEXLive contains the otfinfo command line program, which can query this information; for example
otfinfo -i `kpsewhich lmroman10-regular.otf`


LuaTEX users only In order to load fonts by their name rather than by their filename (e.g., ‘Latin Modern Roman’ instead of ‘ec-lmr10’), you may need to run the script luaotfload-tool, which is distributed with the luaotfload package.
luaotfload-tool


some fonts to use (Preferred Family Name)
  Latin Modern Roman
  Latin Modern Sans
  Latin Modern Mono

  FreeSerif

  Open Sans

  Go Smallcaps
