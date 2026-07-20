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
    palabras = len(texto.split())
    caracteres = len(texto)

    if extension == ".pdf":
        paginas = len(lector.pages)
    else:
        paginas = 0

    return HTMLResponse(f"""
<div class="documento">

<div class="documento-header">

<h2>{file.filename}</h2>

<p>Documento analizado correctamente</p>

</div>

<div class="stats">

<div class="card-stat">
<div class="valor">{paginas}</div>
Páginas
</div>

<div class="card-stat">
<div class="valor">{palabras:,}</div>
Palabras
</div>

<div class="card-stat">
<div class="valor">{caracteres:,}</div>
Caracteres
</div>

</div>

<div class="documento-body">

<pre>{texto}</pre>

</div>

</div>
""")