async function enviarDocumento() {

    const archivo = document.getElementById("archivo").files[0];

    if (!archivo) {
        alert("Seleccione un documento.");
        return;
    }

    const datos = new FormData();

    datos.append("file", archivo);

    document.getElementById("respuesta").innerHTML =
        "<h3>Procesando documento...</h3>";

    try {

        const respuesta = await fetch("/analizar", {
            method: "POST",
            body: datos
        });

        const html = await respuesta.text();

        document.getElementById("respuesta").innerHTML = html;

    } catch (error) {

        document.getElementById("respuesta").innerHTML =
            "<h3>Error al procesar el documento.</h3>";

    }

}