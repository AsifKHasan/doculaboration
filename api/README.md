# api
This is the api service for Doculaboration.

## Environment Setup
- See the [Setup project credential and config](../README.md)
- Make sure `libreoffice` is installed with `python` and `uno` dependencies

## Run the application
```bash
# in the project root
uvicorn api.main:app --reload --port 8200 # development server
# you will find swagger service running on http://localhost:8200/docs
```
