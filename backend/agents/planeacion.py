from analysis.document_parser import DocumentParser


class AgentePlaneacion:

    def __init__(self):
        self.parser = DocumentParser()

    def analizar(self, texto):

        claves = [
            "plan",
            "objetivo",
            "meta",
            "indicador",
            "programa",
            "proyecto",
            "estrategia"
        ]

        encontrados = self.parser.buscar(texto, claves)

        return {
            "agente": "Planeación",
            "categoria": "Planeación Estratégica",
            "hallazgos": encontrados,
            "cantidad": len(encontrados),
            "recomendaciones": [
                "Fortalecer la planeación",
                "Alinear objetivos con indicadores"
            ],
            "nivel_riesgo": "MEDIO"
        }