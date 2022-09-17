# doculaboration - Documentation Collaboration
CRDT (gsheet, etc.) based documentation collaboration pipeline to generate editable (docx, odt, etc.) and printable (pdf, etc.) output from input data

* *gsheet-to-json* is for generating json output from gsheet data. The output json is meant to be fed as input for document generation components of the pipeline
* *json-to-docx* is for generating docx (WordML) documents
* *json-to-odt* is for generating odt (OpenOffice Text) documents
* *json-to-latex* is for generating LaTex for generating printable outputs
* *json-to-context* is for generating ConTeXt for generating printable outputs

## Get the scripts/programs
1. cd to ```d:\projects``` (for Windows) or ```~/projects``` (for Linux)
2. run ```git clone https://github.com/AsifKHasan/doculaboration.git```
3. cd to ```D:\projects\doculaboration``` (for Windows) or ```~/projects/doculaboration``` (for Linux)
4. run command ```pip install -r requirements.txt --upgrade```. See if there is any error or not.

## (optional) update all python packages
```
pip list --outdated --format=freeze | grep -v '^\-e' | cut -d = -f 1  | xargs -n1 pip install -U
```
