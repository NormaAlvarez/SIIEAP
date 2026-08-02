"""Generador de Informe Técnico — Word (.docx) y PDF descargables.

Empaqueta, para UNA entidad, un informe técnico completo que el docente
puede entregar a sus estudiantes como evidencia de ejecución del sistema
y como fuente de trabajo para la asignatura. El informe integra:

  1. Portada institucional (SIIEAP, entidad, fecha)
  2. Contextualización de la Administración Pública Contemporánea:
     un bloque FIJO (igual en todos los informes) que recorre el arco
     teórico completo del microcurrículo — desde los antecedentes griegos
     de lo público/privado, pasando por el Nuevo Institucionalismo, la
     Nueva Gestión Pública (NGP) y la post-NGP, hasta el esquema
     integrador de la Administración Pública contemporánea (Estado
     Digital, Transformación Digital, Gobernanza Inteligente, IA,
     Gobierno Abierto, Gobernanza de Datos, Valor Público, Administración
     Pública Basada en Evidencia, Capacidades Estatales, Resiliencia
     Institucional, Agenda 2030/ODS, OCDE, CEPAL y Naciones Unidas).
  3. Resultado real del diagnóstico IDI-MIPG de la entidad (dimensiones,
     brechas).
  4. El análisis integral generado por IA (motor_analisis_ia.py): lectura
     desde las tres teorías del curso, recomendaciones técnicas/jurídicas/
     financieras, valoración de riesgo y prospectiva.
  5. Nota de trazabilidad y disclaimer académico.

Este módulo NO llama a la API de Claude — solo da formato a un texto de
análisis que ya fue generado antes por motor_analisis_ia.py.

Dependencias nuevas a agregar en requirements.txt:
    python-docx
    reportlab
"""
from __future__ import annotations

import io
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image,
)

from backend.motores.graficas_informe import generar_grafica_dimensiones, generar_grafica_brechas


# ---------------------------------------------------------------------------
# Bloque fijo: Contextualización de la Administración Pública Contemporánea
# ---------------------------------------------------------------------------

CONTEXTUALIZACION_TITULO = "Contextualización de la Administración Pública Contemporánea"

CONTEXTUALIZACION_INTRO = (
    "Este informe se enmarca en el recorrido teórico de la asignatura Enfoques y "
    "Teorías de la Administración Pública, que inicia en los antecedentes de la "
    "cultura griega para la distinción entre lo público y lo privado, y llega hasta "
    "los debates más recientes sobre la transformación digital del Estado. A "
    "continuación se sintetiza ese arco completo, como marco de lectura para el "
    "diagnóstico institucional que se presenta más adelante."
)

CONTEXTUALIZACION_ARCO_HISTORICO = [
    (
        "Antecedentes griegos de lo público y lo privado",
        "La distinción entre la esfera pública (la polis, los asuntos comunes, el "
        "koinon) y la esfera privada (el oikos, el hogar, el idion) es el punto de "
        "partida clásico para pensar qué es \"lo público\" y cómo se organiza su "
        "gobierno. De esa distinción original —lo que compete a todos frente a lo "
        "que compete a cada uno— se deriva la pregunta que atraviesa toda la "
        "disciplina: ¿quién decide sobre los asuntos comunes y con qué reglas?",
    ),
    (
        "Administración Pública clásica y el modelo burocrático weberiano",
        "La Administración Pública clásica (Wilson, Weber) responde a esa pregunta "
        "con un modelo de burocracia racional-legal: jerarquía, especialización, "
        "reglas escritas, impersonalidad y separación entre política y "
        "administración. Este modelo dota al Estado de previsibilidad y control, "
        "pero es también el que después será criticado por su rigidez, su lentitud "
        "y su desconexión de los resultados que efectivamente produce para el "
        "ciudadano.",
    ),
    (
        "Nuevo Institucionalismo en la Administración Pública",
        "March y Olsen (1989) proponen \"redescubrir las instituciones\": las reglas, "
        "normas y rutinas no son un telón de fondo neutro, sino que moldean "
        "directamente el comportamiento de las organizaciones públicas y explican "
        "por qué entidades con recursos similares logran resultados distintos. El "
        "institucionalismo distingue al menos tres variantes que se complementan: "
        "el institucionalismo racional (las reglas como incentivos que estructuran "
        "decisiones), el histórico (la dependencia de la trayectoria: lo que una "
        "entidad hizo antes condiciona lo que puede hacer ahora) y el sociológico "
        "(DiMaggio y Powell, 1983), que explica el isomorfismo institucional: las "
        "organizaciones se parecen entre sí no solo por eficiencia, sino por "
        "presión coercitiva (la norma obliga), mimética (se copia al que parece "
        "legítimo) y normativa (las profesiones y gremios difunden un estándar). "
        "Esta última variante es clave para leer el MIPG: explica por qué una "
        "entidad puede tener todos los formatos y comités que la norma exige "
        "(cumplimiento formal, isomorfismo) sin que eso se traduzca aún en una "
        "práctica real institucionalizada — la brecha entre \"reglas en el papel\" "
        "y \"reglas en uso\".",
    ),
    (
        "Nueva Gestión Pública (NGP)",
        "Surge en los años 80 y 90 (Hood, 1991; Osborne y Gaebler, 1992) como "
        "respuesta a la crisis fiscal del Estado de bienestar y a la percepción de "
        "ineficiencia burocrática. Retoma técnicas de gestión privada: "
        "descentralización de la autoridad gerencial, orientación a resultados e "
        "indicadores de desempeño, introducción de mecanismos de mercado y "
        "competencia, orientación al \"cliente-ciudadano\", y medición explícita del "
        "desempeño (los \"siete doctrinas\" de Hood). Su aporte fue instalar la "
        "cultura de la medición y la rendición de cuentas por resultados —la base "
        "misma de instrumentos como el FURAG—, pero ha sido criticada por "
        "fragmentar el aparato estatal en unidades autónomas difíciles de "
        "coordinar, por debilitar el ethos de lo público al importar lógicas de "
        "mercado, y por reducir el éxito de la gestión a indicadores que no "
        "siempre capturan el bienestar colectivo generado.",
    ),
    (
        "post-Nueva Gestión Pública (post-NGP) y Gobernanza Digital",
        "Ante la fragmentación que dejó la NGP, autores como Dunleavy et al. (2006, "
        "\"Digital Era Governance\") y Osborne (2006, \"Nueva Gobernanza Pública\") "
        "proponen reintegrar lo que la NGP separó: menos unidades autónomas y más "
        "coordinación interinstitucional, menos competencia interna y más redes de "
        "colaboración, menos cliente y más ciudadano co-productor del servicio. La "
        "post-NGP no abandona la orientación a resultados de la NGP, pero la "
        "combina con una lógica de red (gobernanza) y con las posibilidades que "
        "abre la digitalización del Estado para integrar de nuevo lo que estaba "
        "fragmentado.",
    ),
    (
        "Gobernanza Pública y Valor Público",
        "La gobernanza (Kooiman, Rhodes) reconoce que el Estado ya no gobierna en "
        "solitario: coordina una red de actores —mercado, sociedad civil, "
        "cooperación internacional— para producir resultados que ninguno lograría "
        "por separado. En paralelo, Mark Moore (1995, \"Creating Public Value\") "
        "propone el valor público como el criterio último de éxito de la gestión: "
        "no basta con ser eficiente (NGP) ni con coordinar bien una red "
        "(gobernanza) si el resultado no se traduce en bienestar colectivo "
        "legítimo y sostenible. El valor público exige, además, capacidad "
        "operativa real y legitimidad/respaldo político-social — el \"triángulo "
        "estratégico\" de Moore.",
    ),
    (
        "MIPG — Modelo Integrado de Planeación y Gestión",
        "El MIPG (Decreto 1499 de 2017) es la traducción institucional colombiana "
        "de estas tres corrientes: conserva del institucionalismo la atención a "
        "reglas y rutinas (dimensiones y políticas normadas), toma de la NGP la "
        "medición sistemática del desempeño (el FURAG y el IDI), y adopta de la "
        "gobernanza y el valor público la idea de que la meta final no es cumplir "
        "un formato sino generar resultados y valor para el ciudadano. Articula "
        "planeación, gestión del riesgo, talento humano, control interno y "
        "evaluación de resultados en un solo modelo de gestión pública.",
    ),
]

CONTEXTUALIZACION_TEORIAS_CONTEMPORANEAS_TITULO = (
    "Ampliación: las teorías y enfoques de la Administración Pública Contemporánea "
    "(no son solo tres, y se incluyen aquí las corrientes de vanguardia más recientes)"
)

CONTEXTUALIZACION_TEORIAS_CONTEMPORANEAS_INTRO = (
    "Las tres corrientes anteriores (Nuevo Institucionalismo, NGP, post-NGP) explican "
    "la base teórica del MIPG, pero la Administración Pública contemporánea ha seguido "
    "evolucionando en la última década con nuevos enfoques, cada uno con su propio "
    "cuerpo de literatura. Estos no son solo marcos conceptuales importados: cada uno "
    "se desarrolla aquí junto con la norma colombiana vigente que lo valida, citando "
    "el artículo específico cuando la fuente es una ley o decreto. Se incluyen, "
    "además de los diez enfoques del esquema integrador original, cinco corrientes "
    "adicionales de vanguardia administrativa (Gobierno como Plataforma, "
    "interoperabilidad de datos, co-creación y Design Thinking, Administración "
    "Pública Conductual y Gemelos Digitales de Territorio). Aclaración metodológica "
    "importante: los documentos CONPES (Consejo Nacional de Política Económica y "
    "Social) son documentos de POLÍTICA PÚBLICA, no leyes, y por tanto no tienen "
    "\"artículos\" — se citan aquí por su número y objetivo, no deben confundirse con "
    "normas de rango legal. Y dos salvedades honestas, señaladas explícitamente donde "
    "corresponde: la Inteligencia Artificial cuenta en Colombia con política pública "
    "(CONPES), pero todavía no con una ley integral que la regule; y la "
    "Administración Pública Conductual y los Gemelos Digitales de Territorio son, a "
    "la fecha de este informe, corrientes de vanguardia SIN desarrollo normativo "
    "propio verificado en Colombia — se incluyen como oportunidad de política "
    "pública, no como hecho normativo consumado."
)

CONTEXTUALIZACION_TEORIAS_CONTEMPORANEAS = [
    (
        "Estado Digital (e-Gobierno)",
        "El Estado Digital, o e-Gobierno, es la etapa en la que las tecnologías de la "
        "información dejan de ser un simple apoyo administrativo y se integran "
        "directamente en la prestación de servicios y en la toma de decisiones "
        "(Fountain, 2001, \"Building the Virtual State\"; Dunleavy et al., 2006). Su "
        "aporte clave es mostrar que la tecnología no se limita a automatizar un "
        "trámite existente: lo redefine, porque las reglas y la estructura de la "
        "organización terminan moldeadas por lo que el sistema de información permite "
        "o restringe (la \"tecnología en uso\" de Fountain). Colombia lo valida "
        "normativamente en el artículo 2, numeral 8, de la Ley 1341 de 2009 (que ordena "
        "al Gobierno fijar los mecanismos para la masificación del Gobierno en Línea, "
        "hoy Gobierno Digital), desarrollado por el Decreto 1008 de 2018 y actualizado "
        "por el Decreto 767 de 2022, cuyo artículo 2.2.9.1.1.3 (del Decreto 1078 de "
        "2015) fija los principios de la Política de Gobierno Digital vigente, una de "
        "las políticas de gestión y desempeño del MIPG.",
    ),
    (
        "Transformación Digital del Estado",
        "La Transformación Digital va un paso más allá del Estado Digital: no basta "
        "con digitalizar un procedimiento, sino que exige repensar de fondo la "
        "arquitectura institucional, los flujos de datos entre entidades y la cultura "
        "organizacional para operar bajo una lógica \"digital por defecto\" (OCDE, 2014, "
        "\"Recommendation on Digital Government Strategies\"). Colombia la desarrolla "
        "mediante el Decreto 1263 de 2022, que adiciona el Título 23 a la Parte 2 del "
        "Libro 2 del Decreto 1078 de 2015 con los lineamientos de transformación "
        "digital pública, y mediante la Estrategia Nacional Digital 2023-2026 del DNP "
        "(instrumento de política, no de rango legal), que se traducen en la Política "
        "de Gobierno Digital del MIPG.",
    ),
    (
        "Gobernanza Inteligente (Smart Governance)",
        "La Gobernanza Inteligente (Meijer y Bolívar, 2016) describe cómo los "
        "gobiernos usan datos, sensores y tecnologías de ciudad inteligente para "
        "coordinar la acción colectiva de forma más ágil y basada en evidencia casi "
        "en tiempo real. Extiende la lógica de red de la gobernanza pública (post-NGP) "
        "hacia un entorno digital instrumentado. En Colombia se apoya en el modelo de "
        "territorios y ciudades inteligentes que exige el Decreto 767 de 2022 y en el "
        "documento CONPES 3995 de 2020 (Política Nacional de Confianza y Seguridad "
        "Digital, sin artículos por ser documento de política, no ley), que da el "
        "marco de confianza sobre el que puede operar esa coordinación.",
    ),
    (
        "Inteligencia Artificial en el sector público",
        "La incorporación de Inteligencia Artificial en la Administración Pública "
        "(OCDE, 2019, \"Recomendación sobre Inteligencia Artificial\") ofrece "
        "oportunidades reales —automatización de análisis, diagnóstico predictivo, "
        "como el que realiza este mismo sistema— pero exige gestionar riesgos "
        "conocidos: sesgos algorítmicos, opacidad (\"caja negra\") y la necesidad de "
        "supervisión humana significativa sobre cualquier resultado generado por IA. "
        "En Colombia, esta corriente se ha traducido en política pública —los "
        "documentos CONPES 3975 de 2019 y CONPES 4144 de 2025 (Política Nacional de "
        "Inteligencia Artificial, con hoja de ruta a 2030), ninguno de los cuales tiene "
        "artículos por no ser leyes— pero, a la fecha de este informe, el país AÚN NO "
        "cuenta con una ley integral que regule la IA: existen varios proyectos de ley "
        "en trámite en el Congreso sobre la materia, aún sin sanción. Por eso este "
        "informe insiste en que el análisis generado con IA es siempre un punto de "
        "partida metodológico, nunca un dictamen definitivo.",
    ),
    (
        "Gobierno Abierto",
        "El Gobierno Abierto (Open Government Partnership, 2011) descansa en tres "
        "pilares que se refuerzan entre sí: transparencia (acceso a la información "
        "pública), participación ciudadana (involucrar a la ciudadanía en el diseño "
        "de la política) y colaboración (co-producción de soluciones entre Estado, "
        "sociedad civil y sector privado). Colombia es miembro de la Open Government "
        "Partnership desde 2011, y su pilar de transparencia está anclado en la Ley "
        "1712 de 2014 (Ley de Transparencia y del Derecho de Acceso a la Información "
        "Pública): el artículo 1 fija su objeto, el artículo 2 establece el principio "
        "de máxima publicidad (toda información en poder de una entidad pública es "
        "pública salvo excepción constitucional o legal expresa) y el artículo 3 fija "
        "los demás principios de la transparencia y el acceso a la información.",
    ),
    (
        "Gobernanza de Datos",
        "La Gobernanza de Datos trata el dato público como un activo estratégico que "
        "debe gestionarse con reglas claras de calidad, interoperabilidad, seguridad "
        "y uso responsable. En Colombia esta corriente tiene tres soportes legales "
        "concretos que operan en tensión productiva: el Decreto 1389 de 2022 "
        "(compilado como artículo 2.2.24.3.4 del Decreto 1078 de 2015), que crea el "
        "Comité Nacional de Datos para impulsar su gobernanza, uso y reutilización; el "
        "artículo 2 de la Ley 1712 de 2014, que exige apertura y máxima publicidad; y "
        "los artículos 1 y 2 de la Ley 1581 de 2012 (Habeas Data), que protegen los "
        "datos personales y ponen el límite a esa apertura cuando el dato es de una "
        "persona natural.",
    ),
    (
        "Administración Pública Basada en Evidencia",
        "La Administración Pública Basada en Evidencia (Evidence-Based Policy Making; "
        "Head, 2008; Nutley et al., 2007) exige que las decisiones de política "
        "pública se apoyen en datos y evaluaciones rigurosas, no solo en la intuición "
        "o en la presión política coyuntural. En Colombia, el propio Decreto 1499 de "
        "2017 que crea el MIPG es la traducción normativa de esta corriente: el "
        "artículo 2.2.22.3.2 lo define como marco de referencia para dirigir, "
        "planear, ejecutar, hacer seguimiento, evaluar y controlar la gestión "
        "pública, y el artículo 2.2.22.3.3 fija sus objetivos, obligando a medir el "
        "desempeño institucional (FURAG, IDI) como base de la toma de decisiones.",
    ),
    (
        "Capacidades Estatales",
        "Las Capacidades Estatales (Fukuyama, 2004; Grindle, 1997) son la aptitud "
        "real de un Estado o entidad para formular e implementar política pública, "
        "y suelen distinguirse al menos tres tipos: capacidad administrativa (talento "
        "humano, sistemas de información, procesos), capacidad fiscal (recursos "
        "disponibles y su ejecución) y capacidad política (legitimidad y respaldo "
        "para decidir y sostener una decisión en el tiempo). Colombia reconoce "
        "explícitamente estas asimetrías de capacidad entre entidades mediante el "
        "artículo 1 de la Ley 617 de 2000 (categorización presupuestal de "
        "departamentos, en desarrollo del artículo 302 de la Constitución) y su "
        "artículo 2 (categorización de distritos y municipios), junto con el "
        "artículo 3 de la Ley 1454 de 2011 (LOOT), que fija la autonomía, la "
        "descentralización y la asociatividad como principios rectores del "
        "ordenamiento territorial — la base normativa de la tipología y el 'Grupo "
        "par' que este sistema usa para no comparar entidades de capacidad muy "
        "distinta entre sí.",
    ),
    (
        "Resiliencia Institucional",
        "La Resiliencia Institucional (OCDE, 2020, en el marco de la respuesta a la "
        "pandemia) es la capacidad de una entidad para anticipar, absorber y "
        "recuperarse de choques —crisis sanitarias, desastres naturales, choques "
        "fiscales— sin perder su capacidad de generar valor público. Es, de todos los "
        "enfoques contemporáneos, el que tiene el respaldo legal colombiano más "
        "directo y antiguo: el artículo 8 de la Ley 1523 de 2012 crea el Sistema "
        "Nacional de Gestión del Riesgo de Desastres, y su artículo 42 obliga "
        "expresamente a toda entidad pública o privada que preste servicios públicos "
        "o desarrolle actividades de riesgo a realizar un análisis específico de "
        "riesgo, mientras el artículo 43 exige el respectivo plan de gestión del "
        "riesgo — desarrollado en detalle por el Decreto 2157 de 2017, que reglamenta "
        "justamente ese artículo 42. Esto incorpora la pregunta de resiliencia "
        "directamente al análisis institucional que hace este informe.",
    ),
    (
        "Agenda 2030, ODS y marcos multilaterales (OCDE, CEPAL, Naciones Unidas)",
        "La Agenda 2030 y los Objetivos de Desarrollo Sostenible (Naciones Unidas, "
        "2015) sitúan la gestión pública territorial dentro de compromisos globales, "
        "en particular el ODS 16 (paz, justicia e instituciones sólidas), que exige "
        "instituciones eficaces, responsables, transparentes y con participación "
        "ciudadana real. Colombia adoptó formalmente esta agenda mediante el "
        "documento CONPES 3918 de 2018 (estrategia de implementación de los ODS en el "
        "país, sin artículos por ser documento de política, no ley), y es miembro "
        "pleno de la OCDE desde 2020, lo que la obliga a alinear su gestión pública "
        "con los estándares técnicos de esa organización. La CEPAL adapta esos "
        "estándares al contexto latinoamericano y sus asimetrías estructurales, y "
        "Naciones Unidas los articula en metas verificables — el marco de referencia "
        "final frente al cual se mide, en último término, el desempeño institucional "
        "que reporta este sistema.",
    ),
    (
        "Gobierno como Plataforma",
        "El Gobierno como Plataforma (O'Reilly, 2010) propone que el Estado deje de "
        "construir un sistema aislado por cada trámite y en su lugar ofrezca "
        "infraestructura digital compartida y reutilizable (identidad, autenticación, "
        "interoperabilidad) sobre la cual cualquier entidad pueda montar sus "
        "servicios, evitando duplicar esfuerzos. Colombia lo implementa de forma "
        "concreta mediante el Decreto 620 de 2020: el artículo 2.2.17.2.2.1 garantiza "
        "el acceso a los servicios ciudadanos digitales base a través del "
        "'Articulador' (la Agencia Nacional Digital), y el artículo 2.2.17.2.2.2 "
        "establece que el servicio de interoperabilidad del Estado se presta de forma "
        "exclusiva a través de ese Articulador — es decir, la plataforma común que "
        "describe la teoría ya es, en Colombia, una obligación normativa concreta, "
        "fundamentada además en el artículo 147 de la Ley 1955 de 2019 (Plan Nacional "
        "de Desarrollo 2018-2022).",
    ),
    (
        "Interoperabilidad de datos entre entidades",
        "La interoperabilidad —la capacidad de que distintos sistemas de información "
        "intercambien datos sin fricción, siguiendo el modelo de referencia de países "
        "como Estonia (X-Road)— es la condición técnica que hace posible el Gobierno "
        "como Plataforma y la Gobernanza de Datos. En Colombia queda regulada por el "
        "mismo Decreto 620 de 2020 (artículo 2.2.17.2.2.2, servicio de "
        "interoperabilidad exclusivo del Articulador) y por el Decreto 1078 de 2015, "
        "que define el 'marco de interoperabilidad' como la estructura común de "
        "principios, recomendaciones y directrices —políticas, legales, "
        "organizacionales, semánticas y técnicas— que orientan el intercambio de "
        "información entre entidades del Estado.",
    ),
    (
        "Co-creación y Design Thinking en política pública",
        "La co-creación y el Design Thinking proponen que la ciudadanía deje de ser "
        "receptora pasiva de un servicio público y pase a codiseñarlo junto con la "
        "entidad, mediante prototipado rápido y validación directa con el usuario "
        "final. En Colombia esta corriente está parcialmente respaldada: el Decreto "
        "767 de 2022 incorpora la 'Innovación Pública Digital' como uno de los "
        "elementos estructurales de la Política de Gobierno Digital, y bajo ese "
        "marco el Centro de Innovación Pública Digital de MinTIC opera la "
        "metodología CoCrearE. Es importante precisar, con honestidad metodológica, "
        "que CoCrearE es una metodología operativa del Centro, no una norma con "
        "artículos propios: su respaldo legal es el Decreto 767 de 2022 que crea el "
        "marco dentro del cual esa metodología funciona.",
    ),
    (
        "Administración Pública Conductual (Nudge)",
        "La Administración Pública Conductual, o enfoque Nudge (Thaler y Sunstein, "
        "2008, \"Nudge: Improving Decisions About Health, Wealth, and Happiness\"), "
        "propone rediseñar la forma en que se presentan las opciones al ciudadano "
        "(el 'arquitecto de decisiones') para facilitar mejores decisiones sin "
        "restringir la libertad de elegir, apoyándose en evidencia de las ciencias "
        "del comportamiento. Es el enfoque contemporáneo que ha tenido más desarrollo "
        "internacional (la Behavioural Insights Team del Reino Unido, la oficina de "
        "Cass Sunstein en la administración Obama en Estados Unidos) pero, siendo "
        "estrictos con la evidencia disponible, no se identificó una ley o CONPES "
        "colombiano que institucionalice de forma específica una unidad o política "
        "de ciencias del comportamiento a nivel nacional. Se incluye aquí como "
        "corriente teórica vigente y de vanguardia, no como corriente ya "
        "normativizada en Colombia — una oportunidad de política pública más que un "
        "hecho normativo consumado.",
    ),
    (
        "Gemelos Digitales de Territorio (Digital Twins)",
        "Un gemelo digital de territorio es una réplica virtual, alimentada con "
        "datos reales y actualizados, de una ciudad o región, que permite simular el "
        "efecto de una decisión de política pública (una obra, una reubicación de "
        "población, un cambio de uso del suelo) antes de ejecutarla. Es, de los cinco "
        "enfoques añadidos en esta ampliación, el más incipiente en Colombia: no se "
        "identificó un decreto, ley o CONPES que regule específicamente los gemelos "
        "digitales de territorio a la fecha de este informe; el CONPES 4144 de "
        "2025 menciona tecnologías emergentes de forma general dentro de la Política "
        "Nacional de Inteligencia Artificial, pero sin un desarrollo propio para "
        "esta técnica en particular. Se incluye por su relevancia para el debate "
        "actual de vanguardia en gestión territorial, dejando explícito que aún es "
        "una oportunidad de política pública pendiente de desarrollo normativo en el "
        "país.",
    ),
]

CONTEXTUALIZACION_ESQUEMA_INTEGRADOR = [
    "Estado Digital",
    "Transformación Digital",
    "Gobernanza Inteligente",
    "Inteligencia Artificial",
    "Gobierno Abierto",
    "Gobernanza de Datos",
    "Valor Público",
    "Administración Pública Basada en Evidencia",
    "Capacidades Estatales",
    "Resiliencia Institucional",
    "Agenda 2030 y Objetivos de Desarrollo Sostenible (ODS)",
]

CONTEXTUALIZACION_CIERRE = (
    "Este esquema integrador evidencia que la Administración Pública contemporánea "
    "no es una ruptura frente a las teorías clásicas, sino su evolución: el Estado "
    "Digital y la Transformación Digital habilitan una Gobernanza Inteligente "
    "apoyada en Inteligencia Artificial; el Gobierno Abierto y la Gobernanza de "
    "Datos fortalecen el Valor Público y una Administración Pública Basada en "
    "Evidencia; y todo ello construye Capacidades Estatales y Resiliencia "
    "Institucional, en línea con la Agenda 2030 y los ODS y los marcos de "
    "referencia de organismos internacionales como la OCDE, la CEPAL y las "
    "Naciones Unidas, que orientan a los países en la modernización de su gestión "
    "pública hacia la innovación, la transparencia y la generación de valor "
    "público."
)

CADENA_INTERPRETACION_TITULO = "Cadena de interpretación institucional del SIIEAP"

CADENA_INTERPRETACION_INTRO = (
    "El SIIEAP no se limita a contar recomendaciones de Función Pública: cada "
    "brecha detectada se interpreta siguiendo una cadena de análisis que va desde "
    "el dato oficial hasta el plan de tratamiento y su seguimiento posterior. Esta "
    "es la ruta que sigue el sistema para cada brecha priorizada de la entidad:"
)

CADENA_INTERPRETACION_PASOS = [
    "Pregunta FURAG",
    "Dimensión – Política – Índice MIPG",
    "Estándar metodológico esperado",
    "Resultado oficial MIPG",
    "Recomendación oficial de Función Pública",
    "Brecha de implementación identificada",
    "Interpretación desde las teorías y enfoques de la Administración Pública",
    "Paradigma administrativo predominante",
    "Modelo de gestión pública relacionado",
    "Capacidad estatal comprometida",
    "Valor público afectado",
    "Gobernanza comprometida",
    "Transformación digital involucrada",
    "Normatividad aplicable",
    "Riesgo institucional",
    "Impacto en: índice, política, dimensión, IDI, capacidades institucionales y valor público",
    "Plan de tratamiento",
    "Indicadores KPI e indicadores KRI",
    "Seguimiento",
    "Nueva medición FURAG",
]

CADENA_INTERPRETACION_CIERRE = (
    "Esta cadena evita el error de tratar todas las recomendaciones como "
    "equivalentes: lo que importa no es cuántas brechas tiene una entidad, sino "
    "qué tan lejos está cada una del estándar exigido, qué capacidad estatal y "
    "qué valor público compromete, y qué tan bien se cierra el ciclo hasta la "
    "siguiente medición FURAG. El análisis de riesgo que se presenta más "
    "adelante en este informe sigue esta misma lógica de distancia al estándar, "
    "no de simple conteo."
)

DISCLAIMER_INFORME = (
    "Este informe fue generado por el Sistema Integral de Diagnóstico del "
    "Desempeño Institucional (SIIEAP), a partir de datos reales del Índice de "
    "Desempeño Institucional (IDI-MIPG) de Función Pública, complementado con un "
    "análisis generado por inteligencia artificial (Claude, Anthropic) como punto "
    "de partida académico y metodológico. No constituye un dictamen oficial de "
    "Función Pública ni sustituye la validación técnica, jurídica y del líder de "
    "proceso de la entidad analizada. Se entrega con fines académicos, como "
    "evidencia de ejecución del sistema y como insumo de trabajo para la "
    "asignatura."
)


def _fecha_hoy_es():
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
        "agosto", "septiembre", "octubre", "noviembre", "diciembre",
    ]
    hoy = datetime.now()
    return f"{hoy.day} de {meses[hoy.month - 1]} de {hoy.year}"


_COLOR_HEX_POR_RIESGO = {
    "alta": "F5B7B1",
    "media": "FAD7A0",
    "baja": "A9DFBF",
}


# Mapeo determinístico política MIPG -> enfoque(s) contemporáneo(s) + norma.
# Esto NO depende de que la IA lo mencione: se calcula siempre por código a
# partir de las brechas reales de la entidad, garantizando que la conexión
# con los 15 enfoques contemporáneos sea visible en TODOS los informes.
POLITICA_A_ENFOQUE_Y_NORMA = {
    "gestión estratégica del talento humano": ("Capacidades Estatales", "Ley 617 de 2000, art. 1-2; Ley 1454 de 2011, art. 3"),
    "integridad": ("Gobierno Abierto", "Ley 1712 de 2014; Ley 2013 de 2019 y Decreto 830 de 2021 (conflictos de interés)"),
    "planeación institucional": ("Administración Pública Basada en Evidencia", "Decreto 1499 de 2017, art. 2.2.22.3.2-3.3"),
    "gestión presupuestal y eficiencia del gasto": ("Capacidades Estatales", "Ley 610 de 2000, art. 1 (responsabilidad fiscal)"),
    "compras y contratación pública": ("Gobierno como Plataforma", "Decreto 620 de 2020, art. 2.2.17.2.2.1-2.2.2 (TVEC/Articulador)"),
    "fortalecimiento organizacional y simplificación de procesos": ("Capacidades Estatales", "Ley 617 de 2000; Ley 1454 de 2011"),
    "gobierno digital": ("Estado Digital / Transformación Digital del Estado", "Decreto 767 de 2022; Decreto 1263 de 2022"),
    "seguridad digital": ("Gobernanza de Datos", "CONPES 3995 de 2020 (Confianza y Seguridad Digital)"),
    "defensa jurídica": ("Capacidades Estatales", "Ley 87 de 1993 (control jurídico interno)"),
    "mejora normativa": ("Administración Pública Basada en Evidencia", "Decreto 1499 de 2017"),
    "servicio a las ciudadanías": ("Gobierno Abierto / Co-creación y Design Thinking", "Ley 1712 de 2014; Decreto 767 de 2022 (Innovación Pública Digital)"),
    "racionalización de trámites": ("Estado Digital", "Decreto 019 de 2012 (antitrámites); Decreto 767 de 2022"),
    "participación ciudadana en la gestión pública": ("Gobierno Abierto", "Ley 1712 de 2014; Colombia miembro OGP desde 2011"),
    "seguimiento y evaluación del desempeño institucional": ("Administración Pública Basada en Evidencia / Resiliencia Institucional", "Decreto 1499 de 2017; Ley 1523 de 2012"),
    "transparencia, acceso a la información y lucha contra la corrupción": ("Gobierno Abierto", "Ley 1712 de 2014, art. 1-3; ODS 16 (marco externo, ONU 2015)"),
    "gestión documental": ("Gobernanza de Datos", "Ley 594 de 2000 (Ley General de Archivos); Decreto 1389 de 2022"),
    "gestión de la información estadística": ("Gobernanza de Datos / Administración Pública Basada en Evidencia", "Decreto 1389 de 2022"),
    "gestión del conocimiento": ("Administración Pública Basada en Evidencia / Capacidades Estatales", "Decreto 1499 de 2017"),
    "control interno": ("Resiliencia Institucional / Capacidades Estatales", "Ley 87 de 1993, art. 1-2-9; Ley 1523 de 2012, art. 8"),
}


def _enfoque_y_norma_de_politica(nombre_politica: str):
    """Busca el enfoque contemporáneo y la norma asociada a una política MIPG,
    normalizando el nombre para tolerar variantes de mayúsculas/redacción."""
    clave = str(nombre_politica).strip().lower()
    clave = clave.replace("política de ", "").replace("política ", "")
    for politica_conocida, valor in POLITICA_A_ENFOQUE_Y_NORMA.items():
        if politica_conocida in clave or clave in politica_conocida:
            return valor
    return ("Capacidades Estatales", "Decreto 1499 de 2017 (MIPG)")


def _color_hex_riesgo(nivel_riesgo) -> str | None:
    if not nivel_riesgo:
        return None
    return _COLOR_HEX_POR_RIESGO.get(str(nivel_riesgo).strip().lower())


def _sombrear_celda(celda, color_hex: str) -> None:
    """Colorea el fondo de una celda de tabla de python-docx (no hay API
    pública directa, se manipula el XML de la celda)."""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    propiedades_celda = celda._tc.get_or_add_tcPr()
    sombreado = OxmlElement("w:shd")
    sombreado.set(qn("w:val"), "clear")
    sombreado.set(qn("w:color"), "auto")
    sombreado.set(qn("w:fill"), color_hex)
    propiedades_celda.append(sombreado)


# ---------------------------------------------------------------------------
# Generación del .docx
# ---------------------------------------------------------------------------

def generar_reporte_docx(nombre_entidad, diag, analisis_ia_texto, resultado_isvpt=None, resultado_360=None, idi_oficial=None):
    """Devuelve un BytesIO con el informe técnico en formato Word (.docx).

    resultado_isvpt (opcional): un motor_isvpt.ResultadoISVPT ya calculado
    para el grupo de comparación de la entidad. Si se pasa, se agrega una
    sección con el Índice Sintético de Valor Público Territorial (novedad
    metodológica inspirada en el ISDEL de Vélez Tamayo et al., 2026).
    resultado_360 (opcional): un motor_analisis_360.ResultadoAnalisis360 ya
    calculado, usado para el Resumen Ejecutivo (comparación real contra el
    grupo par, igual que en un informe profesional de auditoría).
    idi_oficial (opcional pero MUY recomendado): el IDI oficial publicado
    por Función Pública para esta entidad, leído directamente del archivo
    oficial. Si se proporciona, este informe SIEMPRE lo usa como la cifra
    protagonista (portada, resumen ejecutivo); diag.idi_estimado (el
    cálculo interno de SIIEAP) se muestra únicamente como nota metodológica
    de verificación, nunca como reemplazo del dato oficial.
    """
    doc = Document()

    # Tipografía base más grande y legible (el informe se lee en pantalla y se imprime)
    estilo_normal_doc = doc.styles["Normal"]
    estilo_normal_doc.font.size = Pt(11.5)
    estilo_normal_doc.font.name = "Calibri"
    estilo_normal_doc.paragraph_format.space_after = Pt(8)
    estilo_normal_doc.paragraph_format.line_spacing = 1.15

    # El IDI OFICIAL de Función Pública es siempre la cifra protagonista.
    # Si por alguna razón no se cargó (ej. captura manual), se usa el cálculo
    # interno como única cifra disponible, dejándolo claro en el texto.
    idi_protagonista = idi_oficial if idi_oficial is not None else diag.idi_estimado
    hay_diferencia_idi = (
        idi_oficial is not None and diag.idi_estimado is not None
        and round(idi_oficial, 2) != round(diag.idi_estimado, 2)
    )
    nivel_riesgo_global = "ALTO" if (idi_protagonista or 0) < 40 else ("MEDIO" if (idi_protagonista or 0) < 70 else "BAJO")

    # Portada — con contenido real, no un título vacío
    titulo = doc.add_heading("Modelo de Conocimiento Institucional del Sistema de Inteligencia Artificial para la Evaluación Integral del Desempeño Institucional en Entidades Públicas (SIIEAP)", level=0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run("Informe de Diagnóstico Institucional y Plan de Mejoramiento Prospectivo")
    run_sub.italic = True
    run_sub.font.size = Pt(13)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(nombre_entidad)
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run(f"Generado el {_fecha_hoy_es()} · Índice de Desempeño Institucional (IDI-MIPG), Decreto 1499 de 2017")

    p_autoria = doc.add_paragraph()
    p_autoria.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_autoria = p_autoria.add_run(
        "Docente: Norma Elizabeth Álvarez Grajales · Área de conocimiento: Entidades "
        "Públicas y del Desarrollo · Escuela Superior de Administración Pública (ESAP)"
    )
    run_autoria.italic = True
    run_autoria.font.size = Pt(9)

    doc.add_paragraph()
    tabla_portada = doc.add_table(rows=1, cols=3)
    tabla_portada.style = "Light Grid Accent 1"
    celdas_portada = tabla_portada.rows[0].cells
    celdas_portada[0].text = f"IDI oficial (Función Pública)\n{idi_protagonista}"
    celdas_portada[1].text = f"Nivel de riesgo global\n{nivel_riesgo_global}"
    celdas_portada[2].text = f"Brechas detectadas\n{len(diag.brechas)}"
    for celda in celdas_portada:
        for parrafo in celda.paragraphs:
            parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run_c in parrafo.runs:
                run_c.bold = True
                run_c.font.size = Pt(13)
    if resultado_360 is not None and resultado_360.promedio_idi is not None:
        p_grupo = doc.add_paragraph()
        p_grupo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_grupo.add_run(
            f"Grupo de comparación: {resultado_360.filtro_descripcion} "
            f"({resultado_360.n_entidades} entidades) · IDI promedio del grupo: {resultado_360.promedio_idi}"
        )
    if hay_diferencia_idi:
        p_nota_idi = doc.add_paragraph()
        p_nota_idi.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_nota_idi = p_nota_idi.add_run(
            f"(Nota: el cálculo interno de verificación de SIIEAP, revisando la suma/ponderación de los "
            f"índices reportados, arroja {diag.idi_estimado}. Esta es una cifra de validación metodológica "
            f"interna que NO reemplaza ni desvirtúa el IDI oficial de Función Pública, que es siempre el "
            f"que prevalece.)"
        )
        run_nota_idi.italic = True
        run_nota_idi.font.size = Pt(9)
        run_nota_idi.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_page_break()

    # 1. Resumen ejecutivo — cifras reales de ESTA entidad, arriba de todo
    doc.add_heading("1. Resumen ejecutivo", level=1)
    parrafos_resumen = [
        f"{nombre_entidad} obtuvo un Índice de Desempeño Institucional (IDI) oficial de {idi_protagonista} "
        f"sobre 100 en la vigencia analizada, con un nivel de riesgo global {nivel_riesgo_global}."
    ]
    if resultado_360 is not None and resultado_360.promedio_idi is not None:
        brecha_grupo = round((idi_protagonista or 0) - resultado_360.promedio_idi, 2)
        comparativo = "por debajo" if brecha_grupo < 0 else "por encima"
        parrafos_resumen.append(
            f"Frente a su grupo de comparación ({resultado_360.filtro_descripcion}, "
            f"{resultado_360.n_entidades} entidades), la entidad se ubica {abs(brecha_grupo)} puntos "
            f"{comparativo} del promedio del grupo ({resultado_360.promedio_idi})."
        )
        if resultado_360.percentil_entidad_referencia is not None:
            parrafos_resumen.append(
                f"Esto la ubica en el percentil {resultado_360.percentil_entidad_referencia}% de su grupo."
            )
    dims_criticas = sorted(diag.resultados_por_dimension, key=lambda r: (r.promedio if r.promedio is not None else 0))[:3]
    if dims_criticas:
        nombres_criticas = ", ".join(f"{r.codigo} {r.nombre} ({r.promedio})" for r in dims_criticas)
        parrafos_resumen.append(f"Las dimensiones más críticas son: {nombres_criticas}.")
    dims_fuertes = [r for r in diag.resultados_por_dimension if (r.promedio or 0) >= 60]
    if dims_fuertes:
        nombres_fuertes = ", ".join(f"{r.codigo} {r.nombre} ({r.promedio})" for r in dims_fuertes)
        parrafos_resumen.append(f"Como fortaleza relativa, se destaca(n): {nombres_fuertes}.")
    else:
        parrafos_resumen.append(
            "Ninguna dimensión alcanza el umbral de 60 puntos: el patrón de brechas simultáneas en "
            "múltiples dimensiones sugiere una debilidad estructural de capacidad institucional "
            "(Oszlak), más que fallas puntuales y aisladas de gestión."
        )
    if hay_diferencia_idi:
        parrafos_resumen.append(
            f"Nota metodológica: Función Pública reporta un IDI oficial de {idi_oficial} para esta "
            f"entidad, cifra que este informe usa siempre como referencia principal. El sistema SIIEAP, "
            f"al revisar internamente la suma y ponderación de los índices del archivo oficial, calcula "
            f"además un valor de verificación de {diag.idi_estimado}; esta diferencia es un insumo para "
            f"depurar la metodología interna del sistema, y en ningún caso reemplaza, desvirtúa ni "
            f"contradice la cifra oficial de Función Pública."
        )
    parrafos_resumen.append(
        f"Se detectaron {len(diag.brechas)} brechas de implementación por debajo del umbral esperado "
        "(< 60 puntos), todas desarrolladas en detalle en este informe — sin selección ni recorte."
    )
    for parrafo_r in parrafos_resumen:
        doc.add_paragraph(parrafo_r)

    # Contextualización académica (bloque fijo)
    doc.add_heading(CONTEXTUALIZACION_TITULO, level=1)
    doc.add_paragraph(CONTEXTUALIZACION_INTRO)

    doc.add_heading("Arco teórico: de los griegos a la Administración Pública contemporánea", level=2)
    for titulo_hito, texto_hito in CONTEXTUALIZACION_ARCO_HISTORICO:
        p = doc.add_paragraph()
        run_hito = p.add_run(f"{titulo_hito}. ")
        run_hito.bold = True
        p.add_run(texto_hito)

    doc.add_heading(CONTEXTUALIZACION_TEORIAS_CONTEMPORANEAS_TITULO, level=2)
    doc.add_paragraph(CONTEXTUALIZACION_TEORIAS_CONTEMPORANEAS_INTRO)
    for titulo_teoria, texto_teoria in CONTEXTUALIZACION_TEORIAS_CONTEMPORANEAS:
        p_teoria = doc.add_paragraph()
        run_titulo_teoria = p_teoria.add_run(f"{titulo_teoria}. ")
        run_titulo_teoria.bold = True
        p_teoria.add_run(texto_teoria)

    doc.add_heading("Esquema integrador de la Administración Pública contemporánea", level=2)
    p_esquema = doc.add_paragraph()
    run_esquema = p_esquema.add_run(" → ".join(CONTEXTUALIZACION_ESQUEMA_INTEGRADOR))
    run_esquema.italic = True
    doc.add_paragraph(CONTEXTUALIZACION_CIERRE)

    doc.add_page_break()

    # Cadena de interpretación institucional (bloque fijo)
    doc.add_heading(CADENA_INTERPRETACION_TITULO, level=1)
    doc.add_paragraph(CADENA_INTERPRETACION_INTRO)
    for paso in CADENA_INTERPRETACION_PASOS:
        doc.add_paragraph(paso, style="List Bullet")
    doc.add_paragraph()
    p_cierre_cadena = doc.add_paragraph()
    run_cierre_cadena = p_cierre_cadena.add_run(CADENA_INTERPRETACION_CIERRE)
    run_cierre_cadena.italic = True

    doc.add_page_break()

    # Resultado real del diagnóstico
    doc.add_heading("Resultado del diagnóstico institucional (datos reales)", level=1)
    p_idi = doc.add_paragraph()
    run_idi = p_idi.add_run(f"IDI estimado: {diag.idi_estimado}")
    run_idi.bold = True
    run_idi.font.size = Pt(14)

    tabla = doc.add_table(rows=1, cols=4)
    tabla.style = "Light Grid Accent 1"
    encabezados = tabla.rows[0].cells
    for celda, texto in zip(encabezados, ["Dimensión", "Promedio", "Riesgo", "Índices"]):
        celda.text = texto
        for parrafo in celda.paragraphs:
            for run_enc in parrafo.runs:
                run_enc.bold = True

    for r in diag.resultados_por_dimension:
        fila = tabla.add_row().cells
        fila[0].text = f"{r.codigo} {r.nombre}"
        fila[1].text = str(r.promedio)
        fila[2].text = str(r.nivel_riesgo)
        fila[3].text = f"{r.n_indices_evaluados}/{r.n_indices_esperados}"
        color_fondo = _color_hex_riesgo(r.nivel_riesgo)
        if color_fondo:
            _sombrear_celda(fila[2], color_fondo)

    doc.add_paragraph()
    try:
        buffer_grafica_dim = generar_grafica_dimensiones(diag)
        doc.add_picture(buffer_grafica_dim, width=Inches(6.2))
        parrafo_imagen = doc.paragraphs[-1]
        parrafo_imagen.alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception:
        pass  # si falla la gráfica, el informe sigue sin ella

    doc.add_heading("Brechas priorizadas (todas las detectadas)", level=2)
    if diag.brechas:
        try:
            buffer_grafica_brechas = generar_grafica_brechas(diag)
            if buffer_grafica_brechas:
                doc.add_picture(buffer_grafica_brechas, width=Inches(6.2))
                doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph()
        except Exception:
            pass
        for b in diag.brechas:
            doc.add_paragraph(
                f"{b.codigo_indice} ({b.puntaje}) — {b.nombre_indice} · {b.politica}",
                style="List Bullet",
            )
    else:
        doc.add_paragraph("No se detectaron brechas por debajo del umbral con los datos disponibles.")

    # Tabla de conexión política -> enfoque contemporáneo -> norma (garantizada por código)
    if diag.brechas:
        doc.add_page_break()
        doc.add_heading("Conexión de las brechas con los enfoques contemporáneos de la Administración Pública", level=2)
        doc.add_paragraph(
            "Cada política con brechas se conecta aquí, de forma sistemática, con el enfoque "
            "contemporáneo que mejor la explica (de los 15 desarrollados en la contextualización "
            "de este informe) y con la norma colombiana que lo respalda."
        )
        politicas_con_brecha = []
        vistas = set()
        for b in diag.brechas:
            if b.politica not in vistas:
                vistas.add(b.politica)
                politicas_con_brecha.append(b.politica)
        tabla_enfoques = doc.add_table(rows=1, cols=3)
        tabla_enfoques.style = "Light Grid Accent 1"
        enc_ef = tabla_enfoques.rows[0].cells
        for celda, texto in zip(enc_ef, ["Política con brecha", "Enfoque contemporáneo", "Norma"]):
            celda.text = texto
            for parrafo in celda.paragraphs:
                for run_enc in parrafo.runs:
                    run_enc.bold = True
        for politica in politicas_con_brecha:
            enfoque, norma = _enfoque_y_norma_de_politica(politica)
            fila = tabla_enfoques.add_row().cells
            fila[0].text = politica
            fila[1].text = enfoque
            fila[2].text = norma

    if resultado_isvpt is not None and resultado_isvpt.isvpt_entidad_referencia is not None:
        doc.add_page_break()
        doc.add_heading("🎯 El Termómetro del Valor Público: Índice Sintético de Valor Público Territorial (ISVPT)", level=1)
        doc.add_paragraph(
            "Como complemento al IDI oficial, este informe incorpora un ejercicio de índice "
            "sintético construido con la metodología académica validada por Vélez Tamayo, "
            "Ortiz-Muñoz y Cardona Montoya (2026) para el Índice Sintético de Desarrollo "
            "Económico Local (ISDEL) — normalización min-max y agregación aritmética simple, "
            "siguiendo las directrices de la OCDE (2008) para indicadores compuestos — "
            "aplicada aquí a las 7 dimensiones reales del IDI-MIPG dentro del grupo de "
            "comparación de esta entidad."
        )
        p_isvpt = doc.add_paragraph()
        run_isvpt = p_isvpt.add_run(
            f"ISVPT de {nombre_entidad}: {resultado_isvpt.isvpt_entidad_referencia} "
            f"(posición {resultado_isvpt.posicion_entidad_referencia} de "
            f"{resultado_isvpt.n_entidades} en su grupo de comparación)."
        )
        run_isvpt.bold = True

        if resultado_isvpt.subindices_por_dimension_entidad:
            doc.add_heading("Subíndices normalizados por dimensión (0 = el más rezagado del grupo, 1 = el mejor)", level=2)
            tabla_isvpt = doc.add_table(rows=1, cols=2)
            tabla_isvpt.style = "Light Grid Accent 1"
            enc = tabla_isvpt.rows[0].cells
            enc[0].text, enc[1].text = "Dimensión", "Subíndice normalizado"
            for celda in enc:
                for parrafo in celda.paragraphs:
                    for run_enc in parrafo.runs:
                        run_enc.bold = True
            for dim, valor in resultado_isvpt.subindices_por_dimension_entidad.items():
                fila = tabla_isvpt.add_row().cells
                fila[0].text = dim
                fila[1].text = str(valor)

        doc.add_paragraph()
        p_nota_isvpt = doc.add_paragraph()
        run_nota_isvpt = p_nota_isvpt.add_run(resultado_isvpt.nota_metodologica)
        run_nota_isvpt.italic = True
        run_nota_isvpt.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_page_break()

    # Análisis integral IA
    doc.add_heading("Análisis integral, señales de riesgo multinivel y Plan de Mejoramiento Prospectivo (generado por IA a partir de los datos reales)", level=1)
    for parrafo in analisis_ia_texto.split("\n"):
        if parrafo.strip():
            doc.add_paragraph(parrafo)

    doc.add_page_break()

    # Disclaimer
    doc.add_heading("Nota metodológica", level=2)
    p_disc = doc.add_paragraph()
    run_disc = p_disc.add_run(DISCLAIMER_INFORME)
    run_disc.italic = True
    run_disc.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------------------------
# Generación del PDF (reportlab)
# ---------------------------------------------------------------------------

def generar_reporte_pdf(nombre_entidad, diag, analisis_ia_texto, resultado_isvpt=None, resultado_360=None, idi_oficial=None):
    """Devuelve un BytesIO con el informe técnico en formato PDF.

    resultado_isvpt (opcional): ver docstring de generar_reporte_docx.
    resultado_360 (opcional): ver docstring de generar_reporte_docx.
    idi_oficial (opcional pero MUY recomendado): ver docstring de generar_reporte_docx.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=LETTER,
        topMargin=2 * cm, bottomMargin=2 * cm,
        leftMargin=2 * cm, rightMargin=2 * cm,
    )
    estilos = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle("TituloSIIEAP", parent=estilos["Title"], fontSize=20)
    estilo_h2 = ParagraphStyle("H2SIIEAP", parent=estilos["Heading2"], spaceBefore=14, fontSize=14)
    estilo_normal = ParagraphStyle("NormalSIIEAP", parent=estilos["Normal"], fontSize=11, leading=15)
    estilo_cursiva = ParagraphStyle("Cursiva", parent=estilo_normal, fontName="Helvetica-Oblique", textColor=colors.grey)

    idi_protagonista = idi_oficial if idi_oficial is not None else diag.idi_estimado
    hay_diferencia_idi = (
        idi_oficial is not None and diag.idi_estimado is not None
        and round(idi_oficial, 2) != round(diag.idi_estimado, 2)
    )
    nivel_riesgo_global = "ALTO" if (idi_protagonista or 0) < 40 else ("MEDIO" if (idi_protagonista or 0) < 70 else "BAJO")

    elementos = []

    # Portada — con contenido real, no un título vacío
    elementos.append(Paragraph("Modelo de Conocimiento Institucional del Sistema de Inteligencia Artificial para la Evaluación Integral del Desempeño Institucional en Entidades Públicas (SIIEAP)", estilo_titulo))
    elementos.append(Paragraph("Informe de Diagnóstico Institucional y Plan de Mejoramiento Prospectivo", ParagraphStyle("SubtituloSIIEAP", parent=estilo_normal, fontSize=13, fontName="Helvetica-Oblique")))
    elementos.append(Spacer(1, 14))
    elementos.append(Paragraph(f"<b>{nombre_entidad}</b>", ParagraphStyle("EntidadSIIEAP", parent=estilos["Heading2"], textColor=colors.HexColor("#1F3864"))))
    elementos.append(Paragraph(f"Generado el {_fecha_hoy_es()} · Índice de Desempeño Institucional (IDI-MIPG), Decreto 1499 de 2017", estilo_normal))
    elementos.append(Paragraph(
        "Docente: Norma Elizabeth Álvarez Grajales · Área de conocimiento: Entidades "
        "Públicas y del Desarrollo · Escuela Superior de Administración Pública (ESAP)",
        ParagraphStyle("AutoriaSIIEAP", parent=estilo_normal, fontSize=9, fontName="Helvetica-Oblique"),
    ))
    elementos.append(Spacer(1, 14))

    datos_tabla_portada = [
        ["IDI oficial (Función Pública)", "Nivel de riesgo global", "Brechas detectadas"],
        [str(idi_protagonista), nivel_riesgo_global, str(len(diag.brechas))],
    ]
    tabla_portada = Table(datos_tabla_portada, hAlign="CENTER", colWidths=[5 * cm, 5 * cm, 5 * cm])
    tabla_portada.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 11),
        ("FONTSIZE", (0, 1), (-1, 1), 15),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elementos.append(tabla_portada)
    if resultado_360 is not None and resultado_360.promedio_idi is not None:
        elementos.append(Spacer(1, 10))
        elementos.append(Paragraph(
            f"Grupo de comparación: {resultado_360.filtro_descripcion} "
            f"({resultado_360.n_entidades} entidades) · IDI promedio del grupo: {resultado_360.promedio_idi}",
            estilo_cursiva,
        ))
    if hay_diferencia_idi:
        elementos.append(Spacer(1, 8))
        elementos.append(Paragraph(
            f"(Nota: el cálculo interno de verificación de SIIEAP, revisando la suma/ponderación de los "
            f"índices reportados, arroja {diag.idi_estimado}. Esta es una cifra de validación metodológica "
            f"interna que NO reemplaza ni desvirtúa el IDI oficial de Función Pública, que es siempre el "
            f"que prevalece.)",
            ParagraphStyle("NotaIDI", parent=estilo_cursiva, fontSize=8),
        ))
    elementos.append(PageBreak())

    # 1. Resumen ejecutivo — cifras reales de ESTA entidad, arriba de todo
    elementos.append(Paragraph("1. Resumen ejecutivo", estilos["Heading1"]))
    parrafos_resumen = [
        f"{nombre_entidad} obtuvo un Índice de Desempeño Institucional (IDI) oficial de {idi_protagonista} "
        f"sobre 100 en la vigencia analizada, con un nivel de riesgo global {nivel_riesgo_global}."
    ]
    if resultado_360 is not None and resultado_360.promedio_idi is not None:
        brecha_grupo = round((idi_protagonista or 0) - resultado_360.promedio_idi, 2)
        comparativo = "por debajo" if brecha_grupo < 0 else "por encima"
        parrafos_resumen.append(
            f"Frente a su grupo de comparación ({resultado_360.filtro_descripcion}, "
            f"{resultado_360.n_entidades} entidades), la entidad se ubica {abs(brecha_grupo)} puntos "
            f"{comparativo} del promedio del grupo ({resultado_360.promedio_idi})."
        )
        if resultado_360.percentil_entidad_referencia is not None:
            parrafos_resumen.append(f"Esto la ubica en el percentil {resultado_360.percentil_entidad_referencia}% de su grupo.")
    dims_criticas = sorted(diag.resultados_por_dimension, key=lambda r: (r.promedio if r.promedio is not None else 0))[:3]
    if dims_criticas:
        nombres_criticas = ", ".join(f"{r.codigo} {r.nombre} ({r.promedio})" for r in dims_criticas)
        parrafos_resumen.append(f"Las dimensiones más críticas son: {nombres_criticas}.")
    dims_fuertes = [r for r in diag.resultados_por_dimension if (r.promedio or 0) >= 60]
    if dims_fuertes:
        nombres_fuertes = ", ".join(f"{r.codigo} {r.nombre} ({r.promedio})" for r in dims_fuertes)
        parrafos_resumen.append(f"Como fortaleza relativa, se destaca(n): {nombres_fuertes}.")
    else:
        parrafos_resumen.append(
            "Ninguna dimensión alcanza el umbral de 60 puntos: el patrón de brechas simultáneas en "
            "múltiples dimensiones sugiere una debilidad estructural de capacidad institucional "
            "(Oszlak), más que fallas puntuales y aisladas de gestión."
        )
    if hay_diferencia_idi:
        parrafos_resumen.append(
            f"Nota metodológica: Función Pública reporta un IDI oficial de {idi_oficial} para esta "
            f"entidad, cifra que este informe usa siempre como referencia principal. El sistema SIIEAP, "
            f"al revisar internamente la suma y ponderación de los índices del archivo oficial, calcula "
            f"además un valor de verificación de {diag.idi_estimado}; esta diferencia es un insumo para "
            f"depurar la metodología interna del sistema, y en ningún caso reemplaza, desvirtúa ni "
            f"contradice la cifra oficial de Función Pública."
        )
    parrafos_resumen.append(
        f"Se detectaron {len(diag.brechas)} brechas de implementación por debajo del umbral esperado "
        "(< 60 puntos), todas desarrolladas en detalle en este informe — sin selección ni recorte."
    )
    for parrafo_r in parrafos_resumen:
        elementos.append(Paragraph(parrafo_r, estilo_normal))
        elementos.append(Spacer(1, 4))
    elementos.append(PageBreak())

    # Contextualización
    elementos.append(Paragraph(CONTEXTUALIZACION_TITULO, estilos["Heading1"]))
    elementos.append(Paragraph(CONTEXTUALIZACION_INTRO, estilo_normal))
    elementos.append(Paragraph("Arco teórico: de los griegos a la Administración Pública contemporánea", estilo_h2))
    for titulo_hito, texto_hito in CONTEXTUALIZACION_ARCO_HISTORICO:
        elementos.append(Paragraph(f"<b>{titulo_hito}.</b> {texto_hito}", estilo_normal))
        elementos.append(Spacer(1, 6))

    elementos.append(Paragraph(CONTEXTUALIZACION_TEORIAS_CONTEMPORANEAS_TITULO, estilo_h2))
    elementos.append(Paragraph(CONTEXTUALIZACION_TEORIAS_CONTEMPORANEAS_INTRO, estilo_normal))
    elementos.append(Spacer(1, 4))
    for titulo_teoria, texto_teoria in CONTEXTUALIZACION_TEORIAS_CONTEMPORANEAS:
        elementos.append(Paragraph(f"<b>{titulo_teoria}.</b> {texto_teoria}", estilo_normal))
        elementos.append(Spacer(1, 6))

    elementos.append(Paragraph("Esquema integrador de la Administración Pública contemporánea", estilo_h2))
    elementos.append(Paragraph("<i>" + " → ".join(CONTEXTUALIZACION_ESQUEMA_INTEGRADOR) + "</i>", estilo_normal))
    elementos.append(Spacer(1, 8))
    elementos.append(Paragraph(CONTEXTUALIZACION_CIERRE, estilo_normal))
    elementos.append(PageBreak())

    # Cadena de interpretación institucional (bloque fijo)
    elementos.append(Paragraph(CADENA_INTERPRETACION_TITULO, estilos["Heading1"]))
    elementos.append(Paragraph(CADENA_INTERPRETACION_INTRO, estilo_normal))
    elementos.append(Spacer(1, 6))
    for paso in CADENA_INTERPRETACION_PASOS:
        elementos.append(Paragraph(f"• {paso}", estilo_normal))
    elementos.append(Spacer(1, 8))
    elementos.append(Paragraph(CADENA_INTERPRETACION_CIERRE, estilo_cursiva))
    elementos.append(PageBreak())

    # Resultado real del diagnóstico
    elementos.append(Paragraph("Resultado del diagnóstico institucional (datos reales)", estilos["Heading1"]))
    elementos.append(Paragraph(f"<b>IDI estimado: {diag.idi_estimado}</b>", ParagraphStyle("IDIGrande", parent=estilo_normal, fontSize=15)))
    elementos.append(Spacer(1, 8))

    datos_tabla = [["Dimensión", "Promedio", "Riesgo", "Índices"]]
    for r in diag.resultados_por_dimension:
        datos_tabla.append([
            f"{r.codigo} {r.nombre}", str(r.promedio), str(r.nivel_riesgo),
            f"{r.n_indices_evaluados}/{r.n_indices_esperados}",
        ])
    tabla = Table(datos_tabla, hAlign="LEFT", colWidths=[7 * cm, 2.5 * cm, 2.5 * cm, 2.5 * cm])
    estilo_tabla_dim = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    _COLOR_PDF_POR_RIESGO = {
        "alta": colors.HexColor("#F5B7B1"),
        "media": colors.HexColor("#FAD7A0"),
        "baja": colors.HexColor("#A9DFBF"),
    }
    for indice_fila, r in enumerate(diag.resultados_por_dimension, start=1):
        color_riesgo = _COLOR_PDF_POR_RIESGO.get(str(r.nivel_riesgo).strip().lower())
        if color_riesgo:
            estilo_tabla_dim.append(("BACKGROUND", (2, indice_fila), (2, indice_fila), color_riesgo))
    tabla.setStyle(TableStyle(estilo_tabla_dim))
    elementos.append(tabla)
    elementos.append(Spacer(1, 10))

    try:
        buffer_grafica_dim = generar_grafica_dimensiones(diag)
        elementos.append(Image(buffer_grafica_dim, width=16 * cm, height=16 * cm * 0.5))
        elementos.append(Spacer(1, 10))
    except Exception:
        pass

    elementos.append(Paragraph("Brechas priorizadas (todas las detectadas)", estilo_h2))
    if diag.brechas:
        try:
            buffer_grafica_brechas = generar_grafica_brechas(diag)
            if buffer_grafica_brechas:
                elementos.append(Image(buffer_grafica_brechas, width=16 * cm, height=16 * cm * 0.55))
                elementos.append(Spacer(1, 10))
        except Exception:
            pass
        for b in diag.brechas:
            elementos.append(Paragraph(
                f"• {b.codigo_indice} ({b.puntaje}) — {b.nombre_indice} · {b.politica}", estilo_normal,
            ))
    else:
        elementos.append(Paragraph("No se detectaron brechas por debajo del umbral con los datos disponibles.", estilo_normal))
    elementos.append(PageBreak())

    # Tabla de conexión política -> enfoque contemporáneo -> norma (garantizada por código)
    if diag.brechas:
        elementos.append(Paragraph("Conexión de las brechas con los enfoques contemporáneos de la Administración Pública", estilo_h2))
        elementos.append(Paragraph(
            "Cada política con brechas se conecta aquí, de forma sistemática, con el enfoque "
            "contemporáneo que mejor la explica y con la norma colombiana que lo respalda.",
            estilo_normal,
        ))
        elementos.append(Spacer(1, 6))
        politicas_con_brecha = []
        vistas = set()
        for b in diag.brechas:
            if b.politica not in vistas:
                vistas.add(b.politica)
                politicas_con_brecha.append(b.politica)
        datos_tabla_enfoques = [["Política con brecha", "Enfoque contemporáneo", "Norma"]]
        for politica in politicas_con_brecha:
            enfoque, norma = _enfoque_y_norma_de_politica(politica)
            datos_tabla_enfoques.append([politica, enfoque, norma])
        tabla_enfoques = Table(datos_tabla_enfoques, hAlign="LEFT", colWidths=[5 * cm, 5 * cm, 6 * cm])
        tabla_enfoques.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        elementos.append(tabla_enfoques)
        elementos.append(PageBreak())

    if resultado_isvpt is not None and resultado_isvpt.isvpt_entidad_referencia is not None:
        elementos.append(Paragraph("🎯 El Termómetro del Valor Público: Índice Sintético de Valor Público Territorial (ISVPT)", estilos["Heading1"]))
        elementos.append(Paragraph(
            "Como complemento al IDI oficial, este informe incorpora un ejercicio de índice "
            "sintético construido con la metodología académica validada por Vélez Tamayo, "
            "Ortiz-Muñoz y Cardona Montoya (2026) para el ISDEL — normalización min-max y "
            "agregación aritmética simple, siguiendo las directrices de la OCDE (2008) para "
            "indicadores compuestos — aplicada aquí a las 7 dimensiones reales del IDI-MIPG "
            "dentro del grupo de comparación de esta entidad.",
            estilo_normal,
        ))
        elementos.append(Spacer(1, 6))
        elementos.append(Paragraph(
            f"<b>ISVPT de {nombre_entidad}: {resultado_isvpt.isvpt_entidad_referencia} "
            f"(posición {resultado_isvpt.posicion_entidad_referencia} de "
            f"{resultado_isvpt.n_entidades} en su grupo de comparación).</b>",
            estilo_normal,
        ))
        elementos.append(Spacer(1, 8))

        if resultado_isvpt.subindices_por_dimension_entidad:
            elementos.append(Paragraph("Subíndices normalizados por dimensión (0 = más rezagado del grupo, 1 = mejor del grupo)", estilo_h2))
            datos_tabla_isvpt = [["Dimensión", "Subíndice normalizado"]]
            for dim, valor in resultado_isvpt.subindices_por_dimension_entidad.items():
                datos_tabla_isvpt.append([dim, str(valor)])
            tabla_isvpt = Table(datos_tabla_isvpt, hAlign="LEFT", colWidths=[10 * cm, 4 * cm])
            tabla_isvpt.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]))
            elementos.append(tabla_isvpt)
            elementos.append(Spacer(1, 8))

        elementos.append(Paragraph(resultado_isvpt.nota_metodologica, estilo_cursiva))
        elementos.append(PageBreak())

    # Análisis integral IA
    elementos.append(Paragraph("Análisis integral, señales de riesgo multinivel y Plan de Mejoramiento Prospectivo (generado por IA a partir de los datos reales)", estilos["Heading1"]))
    for parrafo in analisis_ia_texto.split("\n"):
        if parrafo.strip():
            texto_escapado = parrafo.replace("&amp;", "&amp;amp;").replace("&lt;", "&amp;lt;").replace("&gt;", "&amp;gt;")
            elementos.append(Paragraph(texto_escapado, estilo_normal))
            elementos.append(Spacer(1, 4))
    elementos.append(PageBreak())

    # Disclaimer
    elementos.append(Paragraph("Nota metodológica", estilos["Heading2"]))
    elementos.append(Paragraph(DISCLAIMER_INFORME, estilo_cursiva))

    def _pie_de_pagina(canvas_pdf, doc_pdf):
        canvas_pdf.saveState()
        canvas_pdf.setFont("Helvetica", 8)
        canvas_pdf.setFillColor(colors.grey)
        canvas_pdf.drawString(2 * cm, 1.3 * cm, f"SIIEAP — {nombre_entidad}")
        canvas_pdf.drawRightString(LETTER[0] - 2 * cm, 1.3 * cm, f"Página {doc_pdf.page}")
        canvas_pdf.restoreState()

    doc.build(elementos, onFirstPage=_pie_de_pagina, onLaterPages=_pie_de_pagina)
    buffer.seek(0)
    return buffer
