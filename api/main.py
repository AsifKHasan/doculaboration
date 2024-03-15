from http.client import HTTPException
import subprocess
import os
from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

origins = [
    "http://localhost.tiangolo.com",
    "https://localhost.tiangolo.com",
    "http://localhost",
    "http://localhost:8080",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/pdfs/{gsheet_name}")
def read_item(gsheet_name: str) -> FileResponse:
    proc_status = subprocess.call(
        f"./odt-from-gsheet.sh {gsheet_name}",
        shell=True,
    )
    if proc_status != 0:
        raise HTTPException(
            status_code=500,
            detail=f"Maybe you forgot to share {gsheet_name} with Spectrum",
        )
    path = os.path.join("out", f"{gsheet_name}.odt.pdf")
    return FileResponse(
        path,
        media_type="application/pdf",
        filename=f"{gsheet_name}.pdf",
    )
