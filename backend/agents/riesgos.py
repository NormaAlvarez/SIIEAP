from analysis.document_parser import DocumentParser


class AgenteRiesgos:

    def __init__(self):
        self.parser = DocumentParser()

    def analizar(self, texto):

        claves = [
            "riesgo",
            "control",
            "probabilidad",
            "impacto",
            "amenaza",
            "mitigación"
        ]

        encontrados = self.parser.buscar(texto, claves)

        return {
            "agente": "Riesgos",
            "categoria": "Gestión del Riesgo",
            "hallazgos": encontrados,
            "cantidad": len(encontrados),
            "recomendaciones": [
                "Actualizar la matriz de riesgos",
                "Fortalecer controles"
            ],
            "nivel_riesgo": "ALTO"
        }