import asyncio
import os
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware


app = FastAPI()

origins = [
    "*"
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# Function to execute subprocess and stream output
async def execute_subprocess(gsheet_name: str):
    command = f"./api/odt-from-gsheet.sh {gsheet_name}"
    process = await asyncio.create_subprocess_shell(
        command,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE
    )

    while True:
        output = await process.stdout.readline()
        if not output:
            break
        yield f"{output.decode()}"

    await process.communicate()
    if process.returncode != 0:
        raise HTTPException(
            status_code=500,
            detail=f"Maybe you forgot to share {gsheet_name} with Spectrum"
        )

# Route for streaming subprocess output
@app.get("/process/{gsheet_name}")
async def stream_subprocess_output(gsheet_name: str) -> StreamingResponse:
    headers = {}
    headers["Content-Type"] = "text/event-stream"
    return StreamingResponse(content=execute_subprocess(gsheet_name), headers=headers,)

# Route for serving PDF file
@app.get("/pdf/{gsheet_name}/download")
def download_pdf(gsheet_name: str) -> FileResponse:
    path = os.path.join("out", f"{gsheet_name}.odt.pdf")
    return FileResponse(
        path,
        media_type="application/pdf",
        filename=f"{gsheet_name}.pdf",
    )

# Route for serving PDF file
@app.get("/json/{gsheet_name}/download")
def download_json(gsheet_name: str) -> FileResponse:
    path = os.path.join("out", f"{gsheet_name}.json")
    return FileResponse(
        path,
        media_type="text/json",
        filename=f"{gsheet_name}.json",
    )

# Route for serving PDF file
@app.get("/odt/{gsheet_name}/download")
def download_odt(gsheet_name: str) -> FileResponse:
    path = os.path.join("out", f"{gsheet_name}.odt")
    return FileResponse(
        path,
        media_type="application/odt",
        filename=f"{gsheet_name}.odt",
    )
