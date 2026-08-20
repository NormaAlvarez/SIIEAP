"""Motor de Análisis Integral con IA — a demanda, una entidad a la vez.

Genera, SOLO cuando el usuario lo solicita para una entidad puntual (nunca
por lotes automáticos), un análisis que integra:
  - Las tres teorías del curso (NGP, post-NGP, Nuevo Institucionalismo)
  - La cadena de gobernanza (NGP → Gobernanza → Valor Público → Gobierno
    Abierto → Gestión por Resultados → Gestión del Conocimiento →
    Organizaciones que Aprenden)
  - Recomendaciones por ámbito técnico, jurídico y financiero
  - Valoración integral de riesgo (identificación, análisis de probabilidad
    e impacto, diseño de controles y riesgo residual), siguiendo la
    metodología de la Guía para la Gestión Integral del Riesgo en
    Entidades Públicas de Función Pública (v7, 2025)
  - Una prospectiva orientada a valor público

Este módulo NUNCA se ejecuta automáticamente para muchas entidades: se
invoca una vez, para la entidad que el usuario seleccionó, y siempre deja
explícito que es una interpretación de IA a partir de datos secundarios
(el IDI-MIPG), que debe ser validada y ajustada por el líder del proceso
con la información real y el contexto interno/externo de su entidad.

Requiere una variable de entorno o st.secrets con ANTHROPIC_API_KEY.

Nota de trazabilidad (corrección): el prompt exige desarrollar TODAS las
brechas detectadas, sin recortar, en varias secciones (técnica, jurídica,
financiera, riesgo). Para entidades con muchas brechas (20-30+), ese
desarrollo completo no cabe en una sola respuesta del modelo, sin importar
qué tan alto se ponga max_tokens — el límite siempre se puede volver a
alcanzar con una entidad todavía más grande. Por eso, en vez de solo subir
el límite, `generar_analisis_integral` ahora detecta cuándo la respuesta
se cortó por límite de tokens (`stop_reason == "max_tokens"`) y le pide al
modelo que continúe exactamente donde quedó, encadenando fragmentos hasta
completar el análisis (o hasta un tope de seguridad de continuaciones).
"""
from __future__ import annotations

import os

from anthropic import Anthropic

from backend.motores.motor_diagnostico import UMBRAL_BRECHA

MODELO = "claude-sonnet-4-6"

# Tokens de salida por cada llamada individual a la API. Sonnet 4.6 soporta
# hasta 64.000 tokens de salida en una sola llamada — se usa un margen de
# seguridad por debajo de ese tope. Usar un límite alto aquí (en vez de uno
# bajo con más continuaciones) importa por dos razones: (1) cada llamada de
# continuación reenvía TODA la conversación anterior como entrada, lo que
# suma latencia y costo de tokens de entrada en cada ronda; y (2) el tiempo
# real de generación está determinado por la velocidad de salida del modelo
# (aprox. 56 tokens/segundo) multiplicada por los tokens que efectivamente
# hacían falta — dividir esa misma cantidad en más llamadas no la reduce,
# solo agrega rondas de ida y vuelta innecesarias.
MAX_TOKENS_POR_LLAMADA = 60000

# Tope de seguridad de continuaciones encadenadas, para evitar un bucle
# indefinido (y un costo/tiempo de espera indefinidos) si algo saliera mal.
# Con MAX_TOKENS_POR_LLAMADA ya alto, la gran mayoría de entidades —incluso
# con 30+ brechas— deberían terminar en una sola llamada; este tope cubre
# el caso extremo de una entidad con un número de brechas fuera de lo común.
MAX_CONTINUACIONES = 2

PLANTILLA_SISTEMA = """Eres un analista experto en Administración Pública colombiana y en \
gestión integral del riesgo bajo el marco de Función Pública (MIPG, Guía para la Gestión \
Integral del Riesgo en Entidades Públicas, versión 7 de 2025, alineada con COSO-ERM 2017 \
e ISO 31000), apoyando a estudiantes y docentes de la Maestría en Administración Pública \
de la ESAP, y a las propias entidades públicas, en la lectura de su diagnóstico \
institucional IDI-MIPG.

Reglas estrictas que debes seguir siempre:

0. CONTROL DE EXTENSIÓN (aplica a TODO el documento, todas las secciones): sé denso y \
directo, no narrativo ni repetitivo. Nunca sacrifiques contenido exigido por brevedad — cada \
brecha, política, recomendación y elemento pedido en las reglas siguientes debe seguir \
apareciendo completo — pero redáctalo sin relleno: sin transiciones largas entre puntos, sin \
reformular la misma idea dos veces, sin repetir aclaraciones o advertencias que ya diste antes \
en el mismo documento (di cada advertencia metodológica UNA sola vez, donde corresponda, no en \
cada brecha o sección). Prefiere viñetas y listas densas sobre párrafos narrativos largos \
cuando el contenido lo permita (ver formato específico en los literales 3 y 3bis). Esto reduce \
la extensión del documento sin reducir su cobertura ni su rigor.

1. Basas tu análisis EXCLUSIVAMENTE en los datos reales que se te entregan (dimensiones, \
puntajes, brechas, recomendaciones oficiales). No inventas cifras que no se te dieron.

2. Para la lectura teórica usas SOLO estas tres corrientes (son las del curso, no agregues \
otras sin decirlo explícitamente): Nueva Gestión Pública (NGP), post-Nueva Gestión Pública \
(post-NGP), y Nuevo Institucionalismo en la Administración Pública. Dentro de estas tres \
corrientes puedes apoyarte en autores de referencia reconocidos cuando enriquezcan el \
análisis sin salirte de ellas: Oszlak (capacidad institucional) encaja dentro del Nuevo \
Institucionalismo; Mark Moore (1995, valor público y su "triángulo estratégico": \
legitimidad, valor sustantivo, capacidad operativa) encaja dentro de la post-NGP/gobernanza \
y es OBLIGATORIO usarlo explícitamente en la sección 7 (Plan de Mejoramiento Prospectivo).

2bis. ENFOQUES CONTEMPORÁNEOS (15 en total) — el informe ya trae, en su bloque fijo de \
contextualización, el desarrollo completo de estos 15 enfoques con su respaldo normativo \
interno (leyes, decretos y CONPES colombianos) y externo (OCDE, CEPAL, ONU, Agenda 2030/ODS): \
Estado Digital, Transformación Digital del Estado, Gobernanza Inteligente, Inteligencia \
Artificial en el sector público, Gobierno Abierto, Gobernanza de Datos, Administración \
Pública Basada en Evidencia, Capacidades Estatales, Resiliencia Institucional, Agenda 2030 y \
marcos multilaterales (OCDE/CEPAL/ONU), Gobierno como Plataforma, Interoperabilidad de datos \
entre entidades, Co-creación y Design Thinking en política pública, Administración Pública \
Conductual (Nudge) y Gemelos Digitales de Territorio. TU ANÁLISIS DEBE CONECTAR cada brecha \
relevante con el/los enfoque(s) contemporáneo(s) que mejor la explique(n) — por ejemplo, una \
brecha en Gobierno Digital conecta con "Estado Digital" y "Transformación Digital del Estado" \
(Decreto 767 de 2022); una brecha en gestión de riesgos de corrupción o transparencia conecta \
con "Gobierno Abierto" (Ley 1712 de 2014) y con el ODS 16; una brecha en gestión documental o \
de datos conecta con "Gobernanza de Datos" (Decreto 1389 de 2022, Ley 1581 de 2012); una \
brecha en control interno o evaluación conecta con "Capacidades Estatales" y "Resiliencia \
Institucional" (Ley 1523 de 2012). Cuando hagas esta conexión, cita el respaldo normativo \
correspondiente (interno Y, cuando aplique, externo/internacional) en lugar de dejarlo solo \
en el terreno de las tres teorías clásicas — esto es adicional a la lectura teórica del \
literal 2, no la reemplaza.

3. VALORACIÓN INTEGRAL DE RIESGO — NO selecciones solo "las brechas más críticas": desarrolla \
el ciclo completo de la Guía de Función Pública v7 para TODAS Y CADA UNA de las brechas \
detectadas que se te entregan, sin excepción y sin recortar la lista. El usuario necesita el \
panorama completo, no una muestra seleccionada.

FORMATO OBLIGATORIO PARA ESTA SECCIÓN (control de extensión — MISMO CONTENIDO, formato \
compacto): al inicio de la sección de Valoración Integral de Riesgo, escribe UNA sola vez el \
párrafo de precaución metodológica (ver literal f más abajo) — no lo repitas en cada brecha, \
ya que aplica a todas por igual. Luego, para cada brecha, en vez de redactar párrafos \
narrativos, usa un bloque de viñetas compacto con este formato fijo (sin perder ningún dato \
exigido, solo sin prosa de relleno):

   **[Código índice] Nombre de la brecha (puntaje) — Política**
   - Evento / causa inmediata / causa raíz / tipología / factor de riesgo: (una línea densa)
   - Probabilidad: Muy baja/Baja/Media/Alta/Muy alta — razón breve (frecuencia típica del proceso)
   - Impacto: Leve/Moderado/Mayor/Catastrófico — razón breve (afectación económica/reputacional)
   - Control(es): tipo (preventivo/detectivo/correctivo), naturaleza (manual/automático), \
responsable (cargo), acción (verbo fuerte)
   - Riesgo residual: nivel resultante — razón en una frase

Contenido exigido por brecha (NINGUNO de estos elementos se omite, solo se redactan compactos):

   a) Identificación y descripción del riesgo: Evento no deseado, Causa inmediata, Causa raíz \
   (humana, tecnológica, normativa, ambiental u organizacional), Tipología (Estratégico, \
   Operativo/de gestión, de Cumplimiento/legal, Tecnológico o de Seguridad de la Información, \
   Fiscal, de Corrupción o Integridad Pública, o Reputacional), y Factor de riesgo.

   b) Probabilidad: estimación CUALITATIVA razonando a partir de la frecuencia típica de \
   exposición (planeación estratégica ≈ 1 vez/año → muy baja; talento humano/jurídica/ \
   administrativa ≈ mensual → media; contabilidad/cartera/tecnología/tesorería ≈ semanal o \
   diaria → alta/muy alta), aclarando que es orientativa, no un dato medido.

   c) Impacto: estimación CUALITATIVA considerando afectación económica/presupuestal y \
   reputacional, según la naturaleza de la brecha.

   d) Controles: 1-2 controles con tipología (preventivo/detectivo/correctivo), naturaleza \
   (manual/automático), Responsable (cargo, no persona), Acción (verbo fuerte: verificar, \
   validar, conciliar, comparar, revisar, cotejar, detectar).

   e) Riesgo residual: si los controles se implementan bien, ¿el riesgo bajaría de nivel? \
   Explica el porqué en una frase.

   f) PRECAUCIÓN METODOLÓGICA (una sola vez, al INICIO de la sección — no por brecha): las \
   estimaciones de probabilidad, impacto y nivel residual son un PUNTO DE PARTIDA \
   metodológico calculado por IA a partir del tipo de proceso y la brecha detectada — NO \
   sustituyen el análisis del contexto interno/externo real, el mapa de riesgos por proceso, \
   ni la validación del líder del proceso y el equipo directivo, quienes deben confirmar o \
   ajustar estos valores con información real de la entidad, tal como exige la Guía de \
   Función Pública.

3bis. PROHIBIDO PRIORIZAR O RECORTAR: no organices las recomendaciones técnicas, jurídicas ni \
financieras en "horizontes de tiempo" (inmediato/mediano/largo plazo), ni en niveles de \
prioridad, ni selecciones solo un subconjunto "más importante". El usuario necesita el \
desarrollo COMPLETO de todas las brechas y políticas detectadas. Organiza en cambio por \
POLÍTICA o DIMENSIÓN MIPG (agrupando las brechas que correspondan a la misma política), de \
modo que cada política con brechas tenga su propio desarrollo técnico, jurídico y financiero \
completo — sin omitir ninguna.

FORMATO: igual que en el literal 3, usa viñetas compactas y densas por política en vez de \
párrafos narrativos extensos — el contenido exigido (todas las políticas, todas las \
recomendaciones técnicas/jurídicas/financieras) es el mismo, solo cambia que cada punto va \
directo al grano, sin frases de relleno ni transiciones largas entre puntos.

4. Todo lo que generes debe llevar, al final, una nota clara: "Este es un análisis generado \
por IA como punto de partida metodológico para la discusión del equipo directivo y el líder \
de proceso, siguiendo la estructura de la Guía para la Gestión Integral del Riesgo de \
Función Pública (v7, 2025); no es un dictamen oficial ni sustituye la construcción \
participativa del mapa de riesgos de la entidad ni la validación técnica y jurídica \
correspondiente."

5. Sé específico a los datos de ESTA entidad, no genérico. Cita los códigos de \
dimensión/índice reales que se te dan.

6. SEÑALES DE RIESGO MULTINIVEL (nueva sección obligatoria): organiza alertas —derivadas \
EXCLUSIVAMENTE de los índices MIPG entregados, nunca como dictamen ni acusación— agrupadas \
por el tipo de instancia de control que debería revisarlas: (a) Administrativo/Control \
Interno, (b) Disciplinario, (c) Fiscal, (d) Penal (siempre como alerta preventiva, jamás \
acusatoria). Aclara siempre que son lecturas indiciarias a partir de datos secundarios, que \
ameritan revisión por los organismos competentes, no una imputación de responsabilidad.

7. PLAN DE MEJORAMIENTO PROSPECTIVO (nueva sección obligatoria, y el ÚNICO lugar del informe \
donde SÍ corresponde una estructura por fases de tiempo): construye una hoja de ruta a 2-4 \
años organizada en fases (ej. 0-6 meses: contención de riesgo; 6-18 meses: fortalecimiento de \
bases; 18-36 meses: consolidación y comparabilidad con el grupo par; 36-48 meses: \
benchmarking territorial), donde cada fase agrupe acciones concretas que respondan a TODAS \
las brechas ya identificadas (ninguna brecha debe quedar fuera de las cuatro fases). Este \
plan debe estar explícitamente orientado a la toma de decisiones del equipo directivo y a \
asegurar valor público en el sentido de Mark Moore (1995): legitimidad y respaldo político- \
social, valor sustantivo para la ciudadanía, y capacidad operativa real — los tres vértices \
de su "triángulo estratégico" — de modo que el plan no sea solo una lista de tareas técnicas \
sino una propuesta de generación de valor público verificable. IMPORTANTE: esta es la ÚNICA \
sección del informe con estructura temporal por fases; las secciones 2, 3, 4 y 8 (recomendaciones \
técnicas, jurídicas, financieras y riesgo) deben seguir cubriendo TODAS las políticas/brechas \
sin agrupar por tiempo ni recortar nada, tal como exige la regla 3bis.

8. Estructura la respuesta en: (1) Lectura desde las tres teorías, conectando cada brecha \
relevante con el/los enfoque(s) contemporáneo(s) correspondiente(s) y su respaldo normativo \
interno/externo (literal 2bis), (2) Recomendaciones \
técnicas para TODAS las políticas con brechas (organizadas por política/dimensión, no por \
horizonte de tiempo ni prioridad), (3) Recomendaciones jurídicas para TODAS las políticas \
con brechas (basadas en las normas ya citadas en las recomendaciones oficiales entregadas, \
no normas inventadas), (4) Recomendaciones financieras, (5) Valoración integral de riesgo \
para TODAS las brechas detectadas (siguiendo el literal 3 completo: identificación, \
probabilidad, impacto, controles y riesgo residual, para cada una), (6) Señales de riesgo \
multinivel (literal 6: administrativo, disciplinario, fiscal, penal-preventivo), (7) Plan de \
Mejoramiento Prospectivo por fases orientado a valor público (literal 7), (8) Prospectiva \
orientada a valor público (cierre breve, distinto del plan de mejoramiento, con visión de \
futuro deseable para la entidad)."""

_INSTRUCCION_CONTINUAR = (
    "Tu respuesta anterior se cortó por límite de longitud, a mitad del texto. "
    "Continúa EXACTAMENTE donde quedaste, palabra por palabra, sin repetir nada de lo "
    "ya escrito, sin reiniciar ninguna sección desde el principio, sin agregar saludos, "
    "introducciones ni resúmenes de lo anterior. Retoma la frase u oración exactamente en "
    "el punto donde se cortó y sigue desarrollando el resto de las brechas, políticas y "
    "secciones que aún faltan, con el mismo nivel de detalle y siguiendo todas las reglas "
    "del mensaje de sistema original."
)


def construir_prompt_usuario(nombre_entidad, diag, recomendaciones_texto, idi_oficial=None):
    from backend.motores.motor_diagnostico import valor_protagonista_dimension

    if not diag.aplica_mipg_integral:
        return _construir_prompt_usuario_regimen_especial(nombre_entidad, diag, recomendaciones_texto)

    lineas = [f"# Diagnóstico real de: {nombre_entidad}", ""]
    idi_a_usar = idi_oficial if idi_oficial is not None else diag.idi_estimado
    etiqueta_idi = "IDI oficial (Función Pública)" if idi_oficial is not None else "IDI estimado (cálculo interno, sin dato oficial disponible)"
    lineas.append(f"{etiqueta_idi}: {idi_a_usar}")
    lineas.append(
        "IMPORTANTE: usa SIEMPRE la cifra anterior como el IDI de la entidad en todo el "
        "texto que redactes. No la recalcules ni la sustituyas por ningún otro número."
    )
    lineas.append("")
    lineas.append("## Resultado por dimensión (usa EXACTAMENTE estos valores, son el dato oficial u protagonista)")
    for r in diag.resultados_por_dimension:
        valor_protagonista = valor_protagonista_dimension(r)
        lineas.append(f"- {r.codigo} {r.nombre}: {valor_protagonista}, riesgo {r.nivel_riesgo} ({r.n_indices_evaluados}/{r.n_indices_esperados})")
    lineas.append("")
    lineas.append("## Brechas detectadas (< 60 puntos), ordenadas de más crítica a menos")
    lineas.append(
        f"Total de brechas detectadas: {len(diag.brechas)}. Debes cubrir TODAS, sin excepción, "
        "en las secciones de recomendaciones y valoración de riesgo."
    )
    for b in diag.brechas:
        lineas.append(f"- {b.codigo_indice} ({b.puntaje}): {b.nombre_indice} — {b.politica}")
    lineas.append("")
    if recomendaciones_texto:
        lineas.append("## Recomendaciones oficiales de Función Pública (texto real, para fundamentar tu análisis)")
        lineas.append(recomendaciones_texto[:12000])  # límite razonable de contexto
    else:
        lineas.append("## No se cargaron recomendaciones oficiales para esta entidad.")
    return "\n".join(lineas)


def _construir_prompt_usuario_regimen_especial(nombre_entidad, diag, recomendaciones_texto):
    """Prompt para entidades de régimen especial (MECI-only: Contralorías,
    Personerías, Concejos/Asambleas, universidades autónomas, CAR, Banco de
    la República — art. 40 Ley 489/1998, art. 2.2.22.3.4 Decreto 1499/2017).

    CORRECCIÓN (agosto 2026, a solicitud de Norma Álvarez, tras auditoría
    de la Contraloría Distrital de Medellín): la versión anterior de este
    prompt (compartida con las entidades de MIPG íntegro) le entregaba al
    modelo las 7 dimensiones D1-D7 —agrupando ahí puntajes de políticas que
    la entidad reportó de forma voluntaria— diciéndole que USARA esos
    valores como "el dato oficial o protagonista" y que calculara un
    "IDI-MIPG"/"IDI estimado" a partir de ellos. El resultado eran informes
    que se contradecían a sí mismos: por un lado advertían que para este
    régimen "el resultado oficial corresponde únicamente a la Dimensión 7",
    y dos párrafos más abajo presentaban un IDI y cuatro dimensiones
    inventadas como si fueran datos reales de la entidad.

    Esta versión, en cambio:
      1. Presenta el índice de Control Interno (MECI) —lo único que estas
         entidades realmente rinden en el FURAG de Función Pública— como
         la única cifra protagonista/oficial, con el puntaje oficial de la
         política (no un recálculo interno).
      2. NO arma un IDI-MIPG ni agrupa nada en dimensiones D1-D7: cada
         política que la entidad haya reportado adicionalmente (de forma
         voluntaria, no exigida por la norma para este régimen) se lista
         POR SEPARADO con su propio puntaje oficial.
      3. Instruye explícitamente analizar esas políticas voluntarias en
         riesgos, plan de mejoramiento y capítulo de auditoría (a petición
         expresa: no deben desaparecer del informe), pero marcadas como
         hallazgo informativo — no como incumplimiento normativo — ya que
         la norma no se las exige a este tipo de entidad.
    """
    lineas = [f"# Diagnóstico real de: {nombre_entidad}", ""]
    lineas.append(
        "## RÉGIMEN ESPECIAL (MECI-only) — instrucciones OBLIGATORIAS y distintas a una entidad de MIPG íntegro"
    )
    lineas.append(
        "Esta entidad NO está sujeta al Modelo Integrado de Planeación y Gestión (MIPG) en su "
        "integralidad (art. 40 Ley 489/1998, art. 2.2.22.3.4 Decreto 1499/2017): solo está obligada "
        "a la política de Control Interno (MECI). Es lo único que esta entidad realmente rinde en "
        "el FURAG de Función Pública."
    )
    lineas.append(
        "REGLAS ESTRICTAS que debes seguir en TODO el texto que redactes:\n"
        "  a) NO existe un 'IDI-MIPG' ni un 'IDI estimado' para esta entidad. NUNCA calcules, "
        "menciones ni insinúes una cifra de ese tipo, ni la sustituyas o compares con el índice de "
        "Control Interno.\n"
        "  b) NO agrupes puntajes de políticas en dimensiones 'D1', 'D2', 'D3', 'D5', etc. Esas "
        "dimensiones NO están definidas oficialmente para este régimen — preséntalas SIEMPRE por "
        "política individual (nombre de la política, no código de dimensión).\n"
        "  c) La ÚNICA cifra que debes tratar como protagonista/oficial de desempeño de la entidad "
        "es el índice de Control Interno (MECI) que se te da abajo. Dale énfasis explícito en la "
        "introducción y en la conclusión: es lo que la entidad realmente rinde ante Función Pública.\n"
        "  d) Las demás políticas reportadas son voluntarias (la norma no se las exige a este "
        "régimen). SÍ debes analizarlas — inclúyelas en las secciones de riesgos, plan de "
        "mejoramiento y capítulo de auditoría — pero acláralo explícitamente cada vez que las "
        "menciones: son hallazgos informativos de buenas prácticas, no incumplimientos normativos."
    )
    lineas.append("")

    meci_oficial = None
    if diag.codigo_politica_control_interno:
        meci_oficial = (diag.politicas_oficiales or {}).get(diag.codigo_politica_control_interno)
    if meci_oficial is None:
        meci_oficial = diag.idi_estimado  # respaldo si no hubiera dato oficial cargado

    lineas.append(f"## Índice de Control Interno (MECI) — dato OFICIAL de Función Pública: {meci_oficial}")
    lineas.append(
        "Esta es la cifra protagonista de todo el informe para esta entidad. Úsala siempre igual, "
        "sin recalcularla."
    )
    lineas.append("")

    brechas_obligatorias = [b for b in diag.brechas if b.obligatoria]
    brechas_voluntarias = [b for b in diag.brechas if not b.obligatoria]

    lineas.append(f"## Índices de Control Interno por debajo de 60 puntos (brecha real y exigible): {len(brechas_obligatorias)}")
    for b in brechas_obligatorias:
        lineas.append(f"- {b.codigo_indice} ({b.puntaje}): {b.nombre_indice}")
    lineas.append("")

    lineas.append(
        f"## Otras políticas reportadas voluntariamente por la entidad (NO exigidas a este régimen, "
        f"analizar cada una por separado, sin agrupar): {len(brechas_voluntarias)} con puntaje bajo 60"
    )
    if brechas_voluntarias:
        for b in brechas_voluntarias:
            lineas.append(f"- {b.codigo_indice} ({b.puntaje}): {b.nombre_indice} — política: {b.politica}")
    else:
        lineas.append("(Ninguna política voluntaria reportada por esta entidad cae bajo 60 puntos.)")
    lineas.append("")
    if diag.politicas_oficiales:
        otras_politicas_ok = {
            cod: p for cod, p in diag.politicas_oficiales.items()
            if cod != diag.codigo_politica_control_interno and p is not None and p >= UMBRAL_BRECHA
        }
        if otras_politicas_ok:
            lineas.append("## Otras políticas reportadas con puntaje oficial saludable (>= 60, para contexto, sin agrupar):")
            for cod, p in sorted(otras_politicas_ok.items()):
                lineas.append(f"- {cod}: {p}")
            lineas.append("")

    if recomendaciones_texto:
        lineas.append("## Recomendaciones oficiales de Función Pública (texto real, para fundamentar tu análisis)")
        lineas.append(recomendaciones_texto[:12000])
    else:
        lineas.append("## No se cargaron recomendaciones oficiales para esta entidad.")
    return "\n".join(lineas)


def generar_analisis_integral(nombre_entidad, diag, recomendaciones_texto, api_key=None, on_texto_parcial=None, idi_oficial=None):
    """
    Llama a la API de Claude para ESTA entidad. No se debe invocar en bucle
    sobre muchas entidades sin que el usuario lo pida explícitamente para
    cada una (evita costos inesperados y mantiene el principio de 'análisis
    a demanda, no por bloques').

    Si la respuesta se corta por límite de tokens antes de cubrir todas las
    brechas exigidas por el prompt (algo esperable en entidades con muchas
    brechas: 20, 30 o más), se piden automáticamente continuaciones
    encadenadas — hasta MAX_CONTINUACIONES veces — para completar el
    análisis en vez de entregarlo cortado a mitad de una sección.

    on_texto_parcial (opcional): función que se llama repetidamente con el
    texto acumulado hasta el momento (str), a medida que va llegando por
    streaming. Sin esto, una generación larga (varios minutos) no envía
    NINGÚN dato al navegador hasta que termina por completo — lo cual puede
    hacer que la conexión entre el navegador y Streamlit se corte por
    inactividad antes de que la respuesta esté lista. Con esto, la interfaz
    puede mostrar avance en vivo, lo que además mantiene la conexión activa.
    """
    api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError(
            "No hay ANTHROPIC_API_KEY configurada. Agréguela en st.secrets "
            "(Streamlit Community Cloud) o como variable de entorno."
        )

    cliente = Anthropic(api_key=api_key)
    mensaje_usuario = construir_prompt_usuario(nombre_entidad, diag, recomendaciones_texto, idi_oficial=idi_oficial)

    mensajes = [{"role": "user", "content": mensaje_usuario}]
    fragmentos = []
    texto_acumulado_total = ""

    for _intento in range(MAX_CONTINUACIONES + 1):
        # Se usa client.messages.stream(...) en vez de .create(...) porque,
        # con MAX_TOKENS_POR_LLAMADA tan alto, una generación que realmente
        # use varios miles de tokens puede tardar varios minutos — y la API
        # de Anthropic EXIGE streaming para operaciones que puedan superar
        # los 10 minutos (si no, rechaza la solicitud con el error
        # "Streaming is required for operations that may take longer than
        # 10 minutes"). Además, leer stream.text_stream (en vez de saltar
        # directo a get_final_message()) permite reportar avance en vivo
        # mediante on_texto_parcial mientras el texto va llegando.
        with cliente.messages.stream(
            model=MODELO,
            max_tokens=MAX_TOKENS_POR_LLAMADA,
            system=PLANTILLA_SISTEMA,
            messages=mensajes,
        ) as stream:
            if on_texto_parcial is not None:
                for delta in stream.text_stream:
                    texto_acumulado_total += delta
                    on_texto_parcial(texto_acumulado_total)
            respuesta = stream.get_final_message()

        fragmento = "".join(bloque.text for bloque in respuesta.content if hasattr(bloque, "text"))
        fragmentos.append(fragmento)

        if respuesta.stop_reason != "max_tokens":
            # Terminó de forma natural (o por otra razón que no es "se acabó
            # el espacio"): el análisis está completo.
            break

        # Se cortó por límite de tokens: se encadena la respuesta parcial
        # como turno del asistente y se pide que continúe exactamente
        # donde quedó, en la siguiente vuelta del bucle.
        mensajes.append({"role": "assistant", "content": fragmento})
        mensajes.append({"role": "user", "content": _INSTRUCCION_CONTINUAR})
    else:
        # Se agotaron las continuaciones permitidas sin que el modelo
        # terminara por sí solo. Es un caso extremo (entidad con un número
        # muy grande de brechas); se entrega igualmente todo lo generado
        # hasta el momento, con un aviso al final para que quede explícito
        # que puede faltar la cola de la última sección.
        fragmentos.append(
            "\n\n> ⚠️ Aviso automático: el análisis alcanzó el límite de "
            f"{MAX_CONTINUACIONES} continuaciones automáticas antes de "
            "terminar por sí solo (entidad con un número muy alto de "
            "brechas). Es posible que la última sección quede incompleta. "
            "Si es el caso, avise para ampliar el límite de continuaciones."
        )

    return "".join(fragmentos)
