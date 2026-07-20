from collections import Counter
import re


class InstitutionalAnalyzer:

    def analizar(self, texto):

        texto = texto.lower()

        palabras = re.findall(r"\b[a-záéíóúñü]{4,}\b", texto)

        frecuencia = Counter(palabras)

        return frecuencia

    def top(self, texto, cantidad=20):

        frecuencia = self.analizar(texto)

        return frecuencia.most_common(cantidad)