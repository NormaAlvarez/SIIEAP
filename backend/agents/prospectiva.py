from analysis.document_parser import DocumentParser


class AgenteProspectiva:

    def __init__(self):
        self.parser = DocumentParser()

    def analizar(self, texto):

        claves = [
            "futuro",
            "escenario",
            "tendencia",
            "innovación",
            "transformación",
            "2030",
            "2040"
        ]

        encontrados = self.parser.buscar(texto, claves)

        return {
            "agente": "Prospectiva",
            "categoria": "Prospectiva",
            "hallazgos": encontrados,
            "cantidad": len(encontrados),
            "recomendaciones": [
                "Construir escenarios",
                "Analizar tendencias"
            ],
            "nivel_riesgo": "MEDIO"
        }