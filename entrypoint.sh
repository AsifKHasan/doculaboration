#!/bin/bash
soffice --headless --invisible --accept="socket,host=0,port=8100;urp;" &
uvicorn api.main:app --host 0.0.0.0 --port 8200
