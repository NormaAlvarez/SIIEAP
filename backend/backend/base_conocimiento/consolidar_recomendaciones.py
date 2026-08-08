"""
Consolidador de Recomendaciones — TODAS las entidades.

Toma la carpeta descomprimida que Función Pública entrega (un .xlsx por
entidad, agrupados en carpetas por departamento + una carpeta "NACIÓN")
y construye UN SOLO archivo JSON indexado por entidad, para que la app no
tenga que pedir "suba el archivo de recomendaciones" entidad por entidad.

Uso:
    python consolidar_recomendaciones.py /ruta/a/carpeta_descomprimida salida.json

Cada entrada del JSON de salida:
{
  "ALCALDIA DE SAN RAFAEL": {
      "departamento_o_nacion": "ANTIOQUIA",
      "archivo_origen": "ANTIOQUIA/ALCALDIA DE SAN RAFAEL.xlsx",
      "recomendaciones": [
          {"politica_nombre": "...", "codigo_politica": "POL07", "texto": "..."},
          ...
      ]
  },
  ...
}

Esto también sirve como control de calidad: al final imprime cuántos
archivos se leyeron con éxito y cuántos fallaron (con el motivo), para no
descubrir errores silenciosamente en producción.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from cargar_recomendaciones import cargar_recomendaciones


def consolidar(carpeta_raiz: str, ruta_salida: str, offset: int = 0, limite: int | None = None) -> None:
    carpeta_raiz = Path(carpeta_raiz)
    archivos = sorted(carpeta_raiz.rglob("*.xlsx"))
    total_universo = len(archivos)

    if limite is not None:
        archivos = archivos[offset : offset + limite]
    elif offset:
        archivos = archivos[offset:]

    ruta_salida_p = Path(ruta_salida)
    consolidado = {}
    if ruta_salida_p.exists():
        consolidado = json.loads(ruta_salida_p.read_text(encoding="utf-8"))

    fallos = []

    for i, archivo in enumerate(archivos, 1):
        nombre_entidad = archivo.stem.strip().upper()
        departamento_o_nacion = archivo.parent.name

        try:
            recomendaciones = cargar_recomendaciones(archivo)
        except Exception as e:
            fallos.append((str(archivo), str(e)))
            continue

        consolidado[nombre_entidad] = {
            "departamento_o_nacion": departamento_o_nacion,
            "archivo_origen": str(archivo.relative_to(carpeta_raiz)),
            "recomendaciones": [
                {
                    "politica_nombre": r.politica_nombre,
                    "codigo_politica": r.codigo_politica,
                    "texto": r.texto,
                }
                for r in recomendaciones
            ],
        }

    ruta_salida_p.write_text(
        json.dumps(consolidado, ensure_ascii=False, indent=None), encoding="utf-8"
    )

    print(f"Lote procesado: offset={offset}, {len(archivos)} archivos de este lote.")
    print(f"Total acumulado en el archivo de salida: {len(consolidado)} entidades.")
    print(f"Universo total de archivos en la carpeta: {total_universo}")
    print(f"Fallidos en este lote: {len(fallos)}")
    for ruta, error in fallos:
        print(f"  - {ruta}: {error}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("carpeta_raiz")
    parser.add_argument("ruta_salida")
    parser.add_argument("--offset", type=int, default=0)
    parser.add_argument("--limite", type=int, default=None)
    args = parser.parse_args()
    consolidar(args.carpeta_raiz, args.ruta_salida, args.offset, args.limite)
