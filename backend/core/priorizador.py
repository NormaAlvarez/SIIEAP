class Priorizador:

    def clasificar(self,hallazgos):

        if hallazgos>30:

            return "CRÍTICO"

        if hallazgos>20:

            return "ALTO"

        if hallazgos>10:

            return "MEDIO"

        return "BAJO"