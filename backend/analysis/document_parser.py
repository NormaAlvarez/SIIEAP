import re


class DocumentParser:

    def limpiar(self, texto):

        texto = texto.replace("\n", " ")

        texto = re.sub(r"\s+", " ", texto)

        return texto.strip()

    def buscar(self, texto, palabras):

        encontrados = []

        texto = texto.lower()

        for palabra in palabras:

            if palabra.lower() in texto:

                encontrados.append(palabra)

        return encontrados