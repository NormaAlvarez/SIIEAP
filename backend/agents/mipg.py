from analysis.document_parser import DocumentParser


class AgenteMIPG:

    def __init__(self):
        self.parser = DocumentParser()

    def analizar(self, texto):

        claves = [
            "integridad",
            "planeación",
            "riesgos",
            "servicio al ciudadano",
            "gestión documental",
            "talento humano",
            "control interno",
            "evaluación"
        ]

        encontrados = self.parser.buscar(texto, claves)

        return {
            "agente": "MIPG",
            "categoria": "Modelo Integrado de Planeación y Gestión",
            "hallazgos": encontrados,
            "cantidad": len(encontrados),
            "recomendaciones": [
                "Fortalecer las políticas de gestión",
                "Revisar dimensiones de MIPG"
            ],
            "nivel_riesgo": "ALTO" if len(encontrados) < 2 else "BAJO"
        }