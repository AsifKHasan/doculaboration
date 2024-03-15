# api

This is the api service for Doculaboration.

## Environment Setup

1. Setup venv
```bash
python3.8 -m venv venv
source venv/bin/activate
pip install -U pip
pip install -r requirements.txt
```
## Run the application

```bash
uvicorn api.main:app --reload # development server
```
