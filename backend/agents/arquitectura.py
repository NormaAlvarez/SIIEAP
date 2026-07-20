from analysis.document_parser import DocumentParser


class AgenteArquitectura:

    def __init__(self):
        self.parser = DocumentParser()

    def analizar(self, texto):

        claves = [
            "proceso",
            "arquitectura",
            "capacidad",
            "servicio",
            "cadena de valor",
            "gobierno",
            "tecnología"
        ]

        encontrados = self.parser.buscar(texto, claves)

        return {
            "agente": "Arquitectura Empresarial",
            "categoria": "Arquitectura",
            "hallazgos": encontrados,
            "cantidad": len(encontrados),
            "recomendaciones": [
                "Actualizar el mapa de procesos",
                "Fortalecer capacidades institucionales"
            ],
            "nivel_riesgo": "MEDIO"
        }