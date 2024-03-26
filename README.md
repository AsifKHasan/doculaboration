# doculaboration - Documentation Collaboration
CRDT (gsheet, etc.) based documentation collaboration pipeline to generate editable (docx, odt, etc.) and printable (pdf, etc.) output from input data

## Setup project credential and config
- Create a `Google Service Account`. Download the `Google Service Account Credential` file and rename it to `credential.json`
```bash
# copy the credential.json file and paste it to ./ghseet-to-json/conf/
cp </path/to/credential.json> <path/to/doculaboration>/ghseet-to-json/conf/

# in the project root folder

# rename the gsheet-to-json/conf/config.yml.dist to gsheet-to-json/conf/config.yml
mv gsheet-to-json/conf/config.yml.dist gsheet-to-json/conf/config.yml

```

## Run the application with a single command
Prerequisite:
- Setup project credential and config
- Install docker
```bash
# in the project root
docker compose up --build --detach
# you will see a service running on port http://localhost:8200
# you will see a website running on port http://localhost:8300
```

## Deploy the application
There is `.env` file in the project root. Modify the variables to customize the application according to your deployment preferences

Prerequisite:
- Setup project credential and config
- Install docker
```bash
# in the project root
docker compose up --build --detach
# you will see a service running on port http://localhost:8200
# you will see a website running on port http://localhost:8300
```

## Developers guide

### Understand the project

#### Core utils
* *gsheet-to-json* is for generating json output from gsheet data. The output json is meant to be fed as input for document generation components of the pipeline. See `gsheet-to-json/README.md` to learn more
* *json-to-docx* is for generating docx (WordML) documents. See `json-to-docx/README.md` to learn more
* *json-to-odt* is for generating odt (OpenOffice Text) documents. See `json-to-odt/README.md` to learn more
* *json-to-latex* is for generating LaTex for generating printable outputs. See `json-to-latex/README.md` to learn more
* *json-to-context* is for generating ConTeXt for generating printable outputs. See `json-to-context/README.md` to learn more

#### Application framework
* *api* is for api service(FastAPI) application. See `api/README.md` to learn more
* *ui* is for frontend(flutter) application. See `ui/README.md` to learn more
