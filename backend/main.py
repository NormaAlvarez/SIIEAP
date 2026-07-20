from pathlib import Path
import shutil

from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pypdf import PdfReader

app = FastAPI(title="SIIEAP")

BASE = Path(__file__).parent.parent
FRONTEND = BASE / "frontend"
UPLOADS = BASE / "uploads"

UPLOADS.mkdir(exist_ok=True)

app.mount("/frontend", StaticFiles(directory=FRONTEND), name="frontend")


@app.get("/")
async def inicio():
    return FileResponse(FRONTEND / "index.html")


@app.post("/analizar")
async def analizar(file: UploadFile = File(...)):

    destino = UPLOADS / file.filename

    with open(destino, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    texto = ""

    if destino.suffix.lower() == ".pdf":

        lector = PdfReader(str(destino))

        for pagina in lector.pages:

            contenido = pagina.extract_text()

            if contenido:
                texto += contenido + "\n"

    return HTMLResponse(f"""
    <h2>Documento leído correctamente</h2>
    <hr>
    <pre style="white-space:pre-wrap;font-size:15px">
{texto}
    </pre>
    """)