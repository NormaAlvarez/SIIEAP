# SIIEAP — V2 (validado contra datos reales)

## Qué cambió respecto a V1

1. **Catálogo corregido a las 19 políticas oficiales** (antes: 18)
   - Se detectó y corrigió: la Tabla de Índices fuente no incluye POL14
     ("Índice de Seguimiento y Evaluación del Desempeño Institucional")
     porque no se descompone en índices — se agregó manualmente en
     `construir_catalogo.py`, función `_agregar_politicas_no_decompuestas`.
   - Se corrigió un error de tipeo en la fuente oficial: la Dimensión 4
     aparecía como "D04" en la Tabla de Índices pero como "D4" en los
     archivos reales de resultados FURAG. Se normalizó a "D4".

2. **Nuevo cargador de resultados oficiales reales**
   (`backend/base_conocimiento/cargar_resultados_oficiales.py`)
   - Lee directamente los archivos que publica Función Pública
     (`Resultados_vigXXXX_nacion.xlsx` / `_territorio.xlsx`).
   - Extrae automáticamente los 66 índices (I01-I68) y ahora también el
     puntaje directo de políticas sin índices propios (POL14).
   - Los índices/políticas que no aplican a una entidad (celda vacía en
     el Excel oficial) simplemente no entran al cálculo — así es como
     funciona el "criterio diferencial" en la práctica.

3. **Modelo de datos extendido** (`backend/modelos/entidades.py`)
   - Nueva clase `ResultadoPolitica` para políticas sin índices propios.

4. **Motor de diagnóstico actualizado** para usar `ResultadoPolitica`
   cuando una política del catálogo no tiene índices.

## Validación contra un caso real: ALCALDÍA DE SAN RAFAEL (Antioquia)

Usando `Resultados_vig2025_territorio.xlsx` sin modificar nada a mano:

- Índices reportados: 61 de 66 (los 5 faltantes son exactamente los que
  no le aplican: i20, i21 de Gobierno Digital e i30, i31, i32 de POL10
  Mejora Normativa — confirma la exención de categoría 6 que usted señaló).
- **IDI oficial publicado por Función Pública: 30.2**
- **IDI estimado por nuestro Motor de Diagnóstico: 29.78**

La diferencia (0.42 puntos) es esperable: nuestro motor promedia las 7
dimensiones con peso igual, mientras que la metodología oficial completa
del IDI puede aplicar ponderaciones internas no públicas por política.
Para uso diagnóstico (identificar brechas y priorizar) esta aproximación
es válida; para reportar un IDI "oficial" a un tercero, se debe citar
siempre el valor publicado por Función Pública, no el estimado.

## Pendiente — enfoques internacionales (CEPAL / ODS / OCDE)

Se buscó si Función Pública publica una tabla oficial que vincule las 19
políticas del IDI con los Objetivos de Desarrollo Sostenible (ODS) o con
marcos de la OCDE. **No se encontró un documento oficial de ese
cruce.** Lo que sí existe es una referencia conceptual (MIPG se describe
como alineado al "Whole-of-Government Approach" de la OCDE), pero no una
tabla política-por-política.

Antes de construir este módulo, se necesita una de estas dos cosas:
1. Un documento oficial de Función Pública, DNP o CEPAL con el cruce
   explícito (si usted lo tiene, compártalo), o
2. Aceptar que el cruce sería una **interpretación razonada del equipo**,
   no un dato oficial — y etiquetarlo así en el sistema para no
   presentarlo como si Función Pública lo hubiera certificado.

## Cómo correrlo

Igual que en V1 (ver requirements.txt y app.py). Para probar la carga de
datos reales:

```python
from backend.base_conocimiento.cargar_resultados_oficiales import buscar_entidad
from backend.motores.motor_diagnostico import diagnosticar

entidad, idi_oficial, grupo_par = buscar_entidad(
    "ruta/a/Resultados_vig2025_territorio.xlsx", "Territorio", "ALCALDIA DE SAN RAFAEL"
)
diagnostico = diagnosticar(entidad)
```

## Próximo paso sugerido

Conectar `cargar_resultados_oficiales.py` a la interfaz Streamlit (`app.py`)
para poder elegir una entidad real de una lista desplegable en vez de
digitar puntajes a mano. Este es el paso que convierte a SIIEAP en una
herramienta usable por cualquier entidad, no solo un caso de prueba.
