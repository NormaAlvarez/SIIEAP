from pathlib import Path
import shutil

from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pypdf import PdfReader

app = FastAPI(title="SIIEAP")

BASE = Path(__file__).parent.parent

UPLOADS = BASE / "uploads"

UPLOADS.mkdir(exist_ok=True)

app.mount("/frontend", StaticFiles(directory=BASE / "frontend"), name="frontend")


@app.get("/")
def inicio():
    return FileResponse(BASE / "frontend" / "index.html")


@app.post("/analizar")
async def analizar(
    pregunta: str = Form(...),
    file: UploadFile = File(...)
):

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

    return HTMLResponse(pregunta = pregunta.lower()

parrafos = texto.split("\n")

resultado = []

for p in parrafos:

    if any(palabra in p.lower() for palabra in pregunta.split()):

        resultado.append(p)

if len(resultado)==0:

    resultado=["No encontré una respuesta directa en el documento."]

respuesta="<br><br>".join(resultado[:20])

return HTMLResponse(f"""

<h2>Respuesta encontrada</h2>

<hr>

<div style="font-size:16px">

{respuesta}

</div>

""")