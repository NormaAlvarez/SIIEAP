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

from pathlib import Path
import shutil

from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pypdf import PdfReader

# =====================================================
# CONFIGURACIÓN GENERAL
# =====================================================

app = FastAPI(
    title="SIIEAP",
    description="Sistema Inteligente de Investigación y Arquitectura Empresarial Pública",
    version="0.2"
)

BASE_DIR = Path(__file__).resolve().parent.parent

FRONTEND = BASE_DIR / "frontend"
UPLOADS = BASE_DIR / "uploads"

UPLOADS.mkdir(exist_ok=True)

app.mount(
    "/frontend",
    StaticFiles(directory=FRONTEND),
    name="frontend"
)


# =====================================================
# PÁGINA PRINCIPAL
# =====================================================

@app.get("/")
async def inicio():

    return FileResponse(FRONTEND / "index.html")


# =====================================================
# ANALIZAR DOCUMENTO
# =====================================================

@app.post("/analizar")
async def analizar(file: UploadFile = File(...)):

    destino = UPLOADS / file.filename

    with open(destino, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    texto = ""

    extension = destino.suffix.lower()

    if extension == ".pdf":

        lector = PdfReader(str(destino))

        for pagina in lector.pages:

            contenido = pagina.extract_text()

            if contenido:

                texto += contenido + "\n"

    else:

        texto = (
            "Esta versión solamente procesa documentos PDF.\n\n"
            "La versión 0.3 incorporará Word, Excel y PowerPoint."
        )

</> Python

# ===============================
# ESTADÍSTICAS DEL DOCUMENTO
# ===============================

palabras = len(texto.split())
caracteres = len(texto)
lineas = len(texto.splitlines())

paginas = len(lector.pages) if extension == ".pdf" else 0

vista = f"""
<div class="resultado">

<div class="cabecera">

<h1>📄 Documento Analizado</h1>

<h2>{file.filename}</h2>

</div>

<div class="estadisticas">

<div class="dato">
<strong>Páginas</strong><br>{paginas}
</div>

<div class="dato">
<strong>Palabras</strong><br>{palabras}
</div>

<div class="dato">
<strong>Caracteres</strong><br>{caracteres}
</div>

<div class="dato">
<strong>Líneas</strong><br>{lineas}
</div>

</div>

<h2>Contenido del documento</h2>

<div class="texto">

<pre>{texto}</pre>

</div>

</div>
"""
return HTMLResponse(vista)