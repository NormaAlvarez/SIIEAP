# Correcciones aplicadas al SIIEAP — agosto 2026

## 1. Bug de las dimensiones D1-D7 (causa raíz y corrección)
El sistema nunca leía las columnas oficiales D1-D7 del Excel de resultados;
las recalculaba internamente promediando los índices I01-I67, lo que casi
nunca coincidía con el valor oficial de Función Pública.

Corregido en:
- `backend/base_conocimiento/cargar_resultados_oficiales.py`: nueva función
  `_extraer_dimensiones_oficiales()`.
- `backend/motores/motor_diagnostico.py`: nuevo campo `promedio_oficial` en
  `ResultadoDimension` y función `valor_protagonista_dimension()`.
- Los 3 generadores de informe (docx y PDF) ahora usan siempre
  `valor_protagonista_dimension()`, nunca `r.promedio` directamente.

Verificado con 2 entidades reales distintas (Alcaldía de Barbosa y Alcaldía
de Caldas): las 7 dimensiones coinciden exactamente con Territorio/Nación.

## 2. Archivo de recomendaciones dividido (52 MB → 4 partes de máx. 18 MB)
`data/recomendaciones_consolidado.json.gz` se subía a GitHub sin problema
por línea de comandos, pero el navegador (arrastrar y soltar) tiene un
límite práctico de ~25 MB. Se dividió en `data/partes_recomendaciones/`
(parte_00 a parte_03) y `app.py` las reensambla automáticamente al
arrancar (`_reconstruir_gz_desde_partes()`). Probado: reconstruye
exactamente los mismos bytes que el archivo original (5.392 entidades).

## 3. Sin resaltado de color marcando brechas en las tablas de dimensiones
Se quitó el sombreado condicional rojo/verde en las tablas de dimensiones
de los 3 informes (docx y PDF) — ahora solo muestran el valor oficial, sin
marca visual de cuáles están por debajo de 60. El "Semáforo" con emoji del
informe ejecutivo se conservó porque es una función declarada y con
nombre propio, no una marca silenciosa.

## 4. Tres secciones nuevas en los 3 informes (técnico, ejecutivo, estudio
   de caso), basadas en contenido real de las guías anexadas — nada
   inventado:
   - **Análisis de riesgos institucionales**: escalas oficiales de la Guía
     de Gestión Integral del Riesgo V7, con nota metodológica explícita
     sobre cómo se adaptó al no tener datos de frecuencia de actividad.
   - **Capítulo de auditoría: hallazgos y plan de mejoramiento**: los 5
     roles de la Oficina de Control Interno (Decreto 648/2017), hallazgos
     con estructura Criterio-Condición-Criticidad (Norma 14.2 NOGAI™), y
     un plan de mejoramiento con la TOTALIDAD de las recomendaciones
     oficiales de Función Pública (no solo una por hallazgo), agrupadas
     por política con plazo sugerido según criticidad.
   - **Articulación con el Programa de Transparencia y Ética Pública**:
     conecta el Decreto 1122/2024 con el índice POL15 real de la entidad.

   Nuevo módulo: `backend/base_conocimiento/contenido_riesgo_auditoria.py`.
   `app.py` se actualizó (modo individual Y modo por lotes) para pasar la
   lista completa de recomendaciones a los 3 generadores docx.

## Recomendación antes de generar el lote completo
Correr `streamlit run app.py`, generar 2-3 entidades conocidas y comparar
visualmente contra el Excel oficial antes de reprocesar todo el país.
