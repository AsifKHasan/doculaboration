# doculaboration - Documentation Collaboration
CRDT (gsheet, etc.) based documentation collaboration pipeline to generate editable (docx, odt, etc.) and printable (pdf, etc.) output from input data

* *gsheet-to-json* is for generating json output from gsheet data. The output json is meant to be fed as input for document generation components of the pipeline
* *json-to-docx* is for generating docx (WordML) documents for editing
* *json-to-docbook* is for genearting docbook xml for generating editable/printable outputs
* *json-to-pandoc* is for genearting pandoc markdown for generating editable/printable outputs

to update all python packages
```
pip list --outdated --format=freeze | grep -v '^\-e' | cut -d = -f 1  | xargs -n1 pip install -U
```

to generate pdf from a docbook xml
```
pandoc -s -f docbook -t latex -o spectrum__written-exam-questions__rmstu__2021-08-07.pdf spectrum__written-exam-questions__rmstu__2021-08-07.xml
```

to generate pdf from a pandoc markdown
```
pandoc Tasnim.Kabir.Ratul__profile.mkd ../json-to-pandoc/conf/preamble.yml -s --pdf-engine=lualatex -f markdown -t latex -o Tasnim.Kabir.Ratul__profile.pdf

metadata.yaml
```

Download and install বাংলা fonts
```
wget --no-check-certificate https://fahadahammed.com/extras/fonts/font.sh -O font.sh;chmod +x font.sh;bash font.sh;rm font.sh
```

!{\color[rgb]{1,0,0}\setlength\arrayrulewidth{1.0pt}\vline}

\hhline{>{\arrayrulecolor[RGB]{255, 0, 0}}|-|>{\arrayrulecolor[RGB]{255, 0, 0}}-->{\arrayrulecolor[RGB]{0, 255, 0}}|}

\arrayrulecolor{red}\cline{1-1}
\arrayrulecolor{green}\cline{2-2}
\arrayrulecolor{blue}\cline{3-3}
\arrayrulecolor{red}\cline{4-4}
