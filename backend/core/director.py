from agents.juridico import AgenteJuridico
from agents.mipg import AgenteMIPG
from agents.planeacion import AgentePlaneacion
from agents.prospectiva import AgenteProspectiva
from agents.arquitectura import AgenteArquitectura
from agents.riesgos import AgenteRiesgos
from agents.valor_publico import AgenteValorPublico


class DirectorIA:

    def __init__(self):

        self.agentes = [

            AgenteJuridico(),
            AgenteMIPG(),
            AgentePlaneacion(),
            AgenteProspectiva(),
            AgenteArquitectura(),
            AgenteRiesgos(),
            AgenteValorPublico()

        ]

    def analizar_documento(self, texto):

        informe = {

            "total_agentes": len(self.agentes),

            "resultados": []

        }

        hallazgos_totales = 0

        for agente in self.agentes:

            try:

                resultado = agente.analizar(texto)

                informe["resultados"].append(resultado)

                if "cantidad" in resultado:

                    hallazgos_totales += resultado["cantidad"]

            except Exception as e:

                informe["resultados"].append({

                    "agente": agente.__class__.__name__,

                    "error": str(e)

                })

        informe["hallazgos_totales"] = hallazgos_totales

        return informe