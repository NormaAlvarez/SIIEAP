from analysis.document_parser import DocumentParser


class AgenteJuridico:

    def __init__(self):
        self.parser = DocumentParser()

    def analizar(self, texto):

        claves = [
            "ley",
            "decreto",
            "constitución",
            "resolución",
            "acuerdo",
            "conpes",
            "mipg",
            "función pública"
        ]

        encontrados = self.parser.buscar(texto, claves)

        return {
            "agente": "Jurídico",
            "categoria": "Normatividad",
            "hallazgos": encontrados,
            "cantidad": len(encontrados),
            "recomendaciones": [
                "Verificar cumplimiento normativo",
                "Actualizar el marco jurídico aplicable"
            ],
            "nivel_riesgo": "ALTO" if len(encontrados) < 2 else "BAJO"
        }