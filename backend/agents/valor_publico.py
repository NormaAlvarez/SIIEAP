from analysis.document_parser import DocumentParser


class AgenteValorPublico:

    def __init__(self):
        self.parser = DocumentParser()

    def analizar(self, texto):

        claves = [
            "valor público",
            "ciudadano",
            "bienestar",
            "impacto",
            "resultado",
            "participación",
            "confianza"
        ]

        encontrados = self.parser.buscar(texto, claves)

        return {
            "agente": "Valor Público",
            "categoria": "Valor Público",
            "hallazgos": encontrados,
            "cantidad": len(encontrados),
            "recomendaciones": [
                "Incrementar el impacto institucional",
                "Fortalecer la generación de valor público"
            ],
            "nivel_riesgo": "MEDIO"
        }