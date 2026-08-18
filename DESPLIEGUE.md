# Cómo poner SIIEAP en internet, con clave de acceso y análisis con IA

Esto responde exactamente lo que preguntó: dónde reside el sistema, dónde
carga los archivos, y cómo queda protegido con usuario/clave.

## 1. Cree una cuenta gratuita en GitHub (si no tiene)
https://github.com/signup — es el lugar donde vive el código (no los datos
completos de 5.392 entidades, esos van en un archivo aparte, ver paso 3).

## 2. Suba la carpeta SIIEAP a un repositorio GitHub **privado**
- Cree un repositorio nuevo, marcado como **Private** (no público).
- Suba todo el contenido de esta carpeta `SIIEAP/` tal cual está.

## 3. Coloque sus archivos de datos en `SIIEAP/data/` ANTES de subir
Con estos 3 nombres exactos:
```
data/resultados_nacion.xlsx
data/resultados_territorio.xlsx
data/recomendaciones_consolidado.json.gz   ← use la versión comprimida aquí
                                               (Python sí la descomprime bien,
                                               a diferencia de la versión HTML)
```
El archivo de recomendaciones comprimido pesa ~52 MB — cabe sin problema en
GitHub (el límite es 100 MB por archivo).

## 4. Cree su clave de la API de Claude
En https://console.anthropic.com — necesita una cuenta con método de pago
cargado (el análisis con IA tiene costo por uso, ver sección de costos abajo).

## 5. Despliegue en Streamlit Community Cloud
1. Vaya a https://share.streamlit.io e inicie sesión con su cuenta de GitHub.
2. "New app" → seleccione su repositorio privado → archivo principal `app.py`.
3. Antes de desplegar, abra "Advanced settings" → "Secrets" y pegue:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-...su-llave-real..."
   APP_PASSWORD = "la-clave-que-usted-elija-para-sus-estudiantes"
   ```
4. Clic en "Deploy". En unos minutos tendrá una URL pública
   (`https://algo.streamlit.app`) protegida por la clave que puso.

## 6. Comparta con sus estudiantes
Les da la URL + la clave (`APP_PASSWORD`). Al entrar, verán la pantalla de
acceso antes que nada. Una vez dentro, cada uno puede:
- Elegir cualquier entidad (nación o territorio) — ya cargadas automáticamente.
- Ver su diagnóstico y sus recomendaciones oficiales (también automáticas).
- Pedir, con un clic, el análisis integral con IA — **una entidad a la vez**,
  nunca en bloque, tal como usted pidió.

## Sobre el costo del análisis con IA
Cada clic en "Generar análisis integral con IA" hace **una** llamada a la API
de Claude (no a las 5.392 entidades, solo a la que el usuario eligió). El
costo es pequeño por consulta individual, pero si tiene muchos estudiantes
consultando muchas entidades, sí se acumula — revise el panel de facturación
de Anthropic Console periódicamente las primeras semanas para calibrar el
uso real antes de abrirlo a todo el grupo.

## Qué pasa si NO configura ANTHROPIC_API_KEY o APP_PASSWORD
- Sin `APP_PASSWORD`: la app funciona igual pero sin puerta de acceso (cualquiera
  con el enlace entra directo) — útil para pruebas suyas, no para producción.
- Sin `ANTHROPIC_API_KEY`: todo el sistema funciona normal (diagnóstico,
  recomendaciones oficiales) excepto el botón de "Análisis integral con IA",
  que mostrará un mensaje de error claro explicando qué falta.
