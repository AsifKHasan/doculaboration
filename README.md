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
pandoc -s -f markdown -t latex -o spectrum__written-exam-questions__rmstu__2021-08-07.pdf spectrum__written-exam-questions__rmstu__2021-08-07.md
```
