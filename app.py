import streamlit as st
import re
import pandas as pd

st.set_page_config(
    page_title="TallerStock | Gestión de inventario",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =====================================================
# DISEÑO VISUAL
# =====================================================
st.markdown("""
<style>
/* Fondo general: industrial, oscuro y con textura sutil */
.stApp {
    background:
        radial-gradient(circle at 12% 8%, rgba(255, 170, 60, 0.07), transparent 24%),
        radial-gradient(circle at 90% 18%, rgba(120, 150, 180, 0.06), transparent 28%),
        linear-gradient(rgba(9, 12, 17, 0.97), rgba(9, 12, 17, 0.985)),
        repeating-linear-gradient(0deg, rgba(255,255,255,0.018) 0px, rgba(255,255,255,0.018) 1px, transparent 1px, transparent 5px);
}

.block-container {
    max-width: 1250px;
    padding-top: 2rem;
    padding-bottom: 3rem;
}

/* Cabecera */
.taller-header {
    position: relative;
    padding: 28px 30px 25px 30px;
    margin-bottom: 22px;
    border: 1px solid rgba(255,255,255,0.09);
    border-radius: 18px;
    background: linear-gradient(135deg, rgba(31,35,43,0.96), rgba(17,20,26,0.92));
    box-shadow: 0 18px 45px rgba(0,0,0,0.28);
    overflow: hidden;
}

.taller-header:after {
    content: "";
    position: absolute;
    left: 0;
    right: 0;
    bottom: 0;
    height: 3px;
    background: linear-gradient(90deg, #d49a3a, rgba(212,154,58,0.15), transparent);
}

.taller-kicker {
    color: #c9ced6;
    font-size: 0.82rem;
    letter-spacing: 0.16em;
    text-transform: uppercase;
    margin-bottom: 5px;
}

.taller-title {
    margin: 0;
    color: #f5f7fa;
    font-size: 2.35rem;
    font-weight: 800;
    letter-spacing: -0.035em;
}

.taller-subtitle {
    margin-top: 7px;
    color: #9ca5b2;
    font-size: 1rem;
}

/* Botones principales */
.stButton > button {
    border: 1px solid rgba(255,255,255,0.13);
    border-radius: 12px;
    background: linear-gradient(180deg, #292d36, #20232b);
    color: #f3f4f6;
    min-height: 48px;
    font-weight: 650;
    transition: all 0.18s ease;
    box-shadow: 0 5px 18px rgba(0,0,0,0.16);
}

.stButton > button:hover {
    border-color: rgba(212,154,58,0.7);
    background: linear-gradient(180deg, #343943, #272b34);
    transform: translateY(-1px);
}

/* Campos de búsqueda y formularios */
div[data-baseweb="input"] > div,
div[data-baseweb="select"] > div {
    background-color: #242730 !important;
    border-color: rgba(255,255,255,0.10) !important;
    border-radius: 10px !important;
}

input, textarea {
    color: #f4f5f7 !important;
}

/* Etiquetas */
label, .stSelectbox label, .stTextInput label, .stNumberInput label, .stDateInput label {
    color: #d9dde4 !important;
    font-weight: 600 !important;
}

/* Secciones */
.section-title {
    display: flex;
    align-items: center;
    gap: 10px;
    margin: 28px 0 10px 0;
    color: #f4f5f7;
    font-size: 1.45rem;
    font-weight: 750;
}

.section-caption {
    color: #8f98a5;
    margin-bottom: 14px;
}

/* Dataframes */
div[data-testid="stDataFrame"] {
    border: 1px solid rgba(255,255,255,0.09);
    border-radius: 14px;
    overflow: hidden;
    box-shadow: 0 12px 30px rgba(0,0,0,0.18);
}

/* Divisores */
hr {
    border-color: rgba(255,255,255,0.08) !important;
}

/* Mensajes */
div[data-testid="stAlert"] {
    border-radius: 12px;
}

/* Quitar margen visual excesivo */
.stMarkdown p {
    line-height: 1.45;
}
</style>
""", unsafe_allow_html=True)

archivo = "INVENTARIO.xlsx"


# =====================================================
# FUNCIONES
# =====================================================

def limpiar_texto(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    if texto.lower() in ["nan", "none"]:
        return ""

    return texto


def limpiar_cantidad(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    if texto.lower() in ["nan", "none"]:
        return ""

    if texto.endswith(".0"):
        texto = texto[:-2]

    return texto


# =====================================================
# LEER EL EXCEL
# =====================================================

excel = pd.ExcelFile(archivo)

tablas = []

for hoja in excel.sheet_names:

    datos = pd.read_excel(
        archivo,
        sheet_name=hoja
    )

    datos["Ubicación"] = hoja

    tablas.append(datos)


inventario = pd.concat(
    tablas,
    ignore_index=True
)


# =====================================================
# CREAR COLUMNA ARTÍCULO
# =====================================================

inventario["Artículo"] = ""

if "Nombre" in inventario.columns:

    inventario["Artículo"] = (
        inventario["Nombre"]
        .apply(limpiar_texto)
    )


if "Nombre (tornillería)" in inventario.columns:

    nombres_tornilleria = (
        inventario["Nombre (tornillería)"]
        .apply(limpiar_texto)
    )

    vacios = inventario["Artículo"] == ""

    inventario.loc[vacios, "Artículo"] = (
        nombres_tornilleria[vacios]
    )


# =====================================================
# ELIMINAR FILAS COMPLETAMENTE VACÍAS
# =====================================================

# Las filas sin artículo son filas vacías del Excel.
# No deben convertirse en artículos con cantidad 1.
inventario = inventario[
    inventario["Artículo"].str.strip() != ""
].copy()


# =====================================================
# IDENTIFICAR LAS DOS CAJAS
# =====================================================

inventario["Es contenedor"] = False
inventario["Nombre contenedor"] = ""


if "Nombre" in inventario.columns:

    nombres_originales = (
        inventario["Nombre"]
        .apply(limpiar_texto)
        .str.lower()
    )

    caja_herramientas = (
        nombres_originales.str.startswith(
            "caja con pequeñas herramientas para impresión 3d"
        )
    )

    inventario.loc[
        caja_herramientas,
        "Es contenedor"
    ] = True

    inventario.loc[
        caja_herramientas,
        "Nombre contenedor"
    ] = "Caja herramientas impresión 3D"


    caja_accesorios = (
        nombres_originales.str.startswith(
            "caja de accesorios para la impresora 3d"
        )
    )

    inventario.loc[
        caja_accesorios,
        "Es contenedor"
    ] = True

    inventario.loc[
        caja_accesorios,
        "Nombre contenedor"
    ] = "Caja accesorios impresora 3D"


# =====================================================
# NOMBRE ORIGINAL PARA BÚSQUEDA
# =====================================================

# Conservamos el nombre completo del Excel para que el buscador
# siga encontrando los artículos aunque en pantalla usemos
# un nombre más corto.
inventario["Artículo búsqueda"] = inventario["Artículo"]


# =====================================================
# NOMBRE CORTO DE LAS CAJAS
# =====================================================

inventario.loc[
    inventario["Nombre contenedor"] != "",
    "Artículo"
] = inventario.loc[
    inventario["Nombre contenedor"] != "",
    "Nombre contenedor"
]


# =====================================================
# NOMBRES CORTOS PARA MOSTRAR EN LA APP
# =====================================================

# Estos son SOLO los nombres que se muestran en la aplicación.
# El nombre original se conserva en "Artículo búsqueda" para
# que el buscador siga encontrando el artículo por su descripción
# completa del Excel.

nombres_cortos = {
    "Bolsita con keycap de repuesto, brida blanca, muelle, tornillo prisionero con base de metacrilato transparente, perno de tope de aluminio y leva de bloqueo":
        "Bolsita repuestos varios impresora 3D",

    "Caja con cintas aislantes autoadhesivas de caucho (de distintos tamaños)":
        "Caja cintas aislantes caucho varios tamaños",

    "Banda de acero flexible con adhesivo por una cara , contra-réplica de la cinta magnética":
        "Banda acero flexible adhesiva",

    "Rollo de cinta adhesiva Tesa Signal Premium (franjas amarillas y negras)":
        "Rollo cinta Tesa Signal Premium amarillo/negro",

    "Tacos pladur 20 unidades con sus respectivos 20 tornillos de 4,5 x 40 mm (fischer)":
        "Tacos pladur 20 uds + tornillos 4,5 x 40 mm Fischer",

    "-Tuercas deslizantes con freno para perfiles de aluminio        -Tornillos de cabeza abombada con huella Torx para perfiles de aluminio (métrica 6)":
        "Tuercas deslizantes + tornillos Torx perfiles aluminio M6",

    "-Tornillos de cabeza cilíndrica con huella Allen                    -Tornillos de cabeza avellanada plana con huella Allen":
        "Tornillos Allen cabeza cilíndrica + avellanada",

    "Maletín negro pequeño: juego de prensaterminales con carraca y cabezas intercambiables de Alyco":
        "Maletín negro pequeño prensaterminales con carraca Alyco",

    "Maletín verde (makita pero más pequeño) con accesorios para una pistola de calor (debería estar también en el maletín pero no está)":
        "Maletín verde pequeño Makita accesorios pistola de calor",

    "Maletín negro (el más grande) con herramienta manual de un sistema PIT y controlador manual":
        "Maletín negro grande sistema PIT + controlador",

    "Maletín verde (makita) con llave de impacto inalámbrica y complementos":
        "Maletín verde Makita llave de impacto + complementos",

    "Maletín gris y azul de llaves y puntas de destornillador":
        "Maletín gris y azul llaves + puntas destornillador",

    "Maletín rojo con comprobador de anclajes portátil Hydrajaws M2000":
        "Maletín rojo comprobador anclajes Hydrajaws M2000",

    "Caja marrón de un aplicador neumático bicomponente Mixpac DP2X de Sulzer":
        "Caja marrón aplicador Mixpac DP2X Sulzer",

    "Funda Elcometer con medidor digital de espesor de revestimiento":
        "Funda Elcometer medidor espesor",

    "Componente base de recubrimiento cerámico Belzona 1321 (0,91 kg)":
        "Base recubrimiento cerámico Belzona 1321 – 0,91 kg",

    "Spray de desengrasante de secado rápido (Fast Dry Degreaser)":
        "Spray desengrasante Fast Dry",

    "Spray de pintura acrílica al agua (Pintyplus evolution)":
        "Spray pintura acrílica Pintyplus Evolution",

    "Cartucho de grasa negra de litio con bisulfuro de molibdeno":
        "Cartucho grasa litio negra con MoS₂",

    "Cartucho de adhesivo estructural acrílico para plásticos":
        "Cartucho adhesivo estructural acrílico plásticos",

    "Juego de punzones para grabado de números de 8 mm (RS PRO)":
        "Punzones números 8 mm RS PRO",

    "Juego de punzones para grabado de letras de 8 mm (RS PRO)":
        "Punzones letras 8 mm RS PRO",

    "Juego de punzones para grabado de números de 4 mm (NUSAC)":
        "Punzones números 4 mm NUSAC",

    "Juego de punzones para grabado de letras de 4 mm (NUSAC)":
        "Punzones letras 4 mm NUSAC",

    "Estuche con juego de 6 botadores cilíndricos (2, 3, 4, 5, 6 y 8 mm)":
        "Estuche botadores cilíndricos 2–8 mm",

    "Bolsa de plástico transparente con pasadores cilíndricos de acero":
        "Bolsa transparente pasadores cilíndricos acero",

    "Bolsa de plástico transparente con tornillos y arandelas grandes":
        "Bolsa transparente tornillos + arandelas grandes",

    "Tacos pladur 50 unidades con tornillos métrica 5  (HILTI)":
        "Tacos pladur 50 uds + tornillos M5 HILTI",

    "Magigoo Pro Kit con 6 adhesivos distintos y un paño de limpieza":
        "Magigoo Pro Kit – 6 adhesivos + paño",
}


def normalizar_nombre(valor):
    """Normaliza espacios para evitar que diferencias del Excel rompan el cambio."""
    return " ".join(str(valor).strip().split())


def nombre_corto(valor):
    original = normalizar_nombre(valor)

    if original == "":
        return ""

    # Primero intentamos coincidencia exacta.
    for nombre_original, corto in nombres_cortos.items():
        if normalizar_nombre(nombre_original).casefold() == original.casefold():
            return corto

    # Después usamos fragmentos característicos para que pequeñas
    # diferencias de espacios o puntuación no impidan el cambio.
    reglas = [
        ("Maletín negro pequeño:", "Maletín negro pequeño prensaterminales con carraca Alyco"),
        ("Maletín verde (makita pero más pequeño)", "Maletín verde pequeño Makita accesorios pistola de calor"),
        ("Maletín negro (el más grande)", "Maletín negro grande sistema PIT + controlador"),
        ("Maletín verde (makita) con llave de impacto", "Maletín verde Makita llave de impacto + complementos"),
        ("Maletín gris y azul de llaves", "Maletín gris y azul llaves + puntas destornillador"),
        ("Maletín rojo con comprobador de anclajes", "Maletín rojo comprobador anclajes Hydrajaws M2000"),
        ("Caja marrón de un aplicador neumático", "Caja marrón aplicador Mixpac DP2X Sulzer"),
        ("Caja con cintas aislantes autoadhesivas", "Caja cintas aislantes caucho varios tamaños"),
        ("Funda Elcometer con medidor digital", "Funda Elcometer medidor espesor"),
        ("Banda de acero flexible con adhesivo", "Banda acero flexible adhesiva"),
        ("Rollo de cinta adhesiva Tesa Signal Premium", "Rollo cinta Tesa Signal Premium amarillo/negro"),
        ("Tacos pladur 20 unidades", "Tacos pladur 20 uds + tornillos 4,5 x 40 mm Fischer"),
        ("Tacos pladur 50 unidades", "Tacos pladur 50 uds + tornillos M5 HILTI"),
        ("Spray de desengrasante de secado rápido", "Spray desengrasante Fast Dry"),
        ("Spray de pintura acrílica al agua", "Spray pintura acrílica Pintyplus Evolution"),
        ("Cartucho de grasa negra de litio", "Cartucho grasa litio negra con MoS₂"),
        ("Cartucho de adhesivo estructural acrílico", "Cartucho adhesivo estructural acrílico plásticos"),
        ("Componente base de recubrimiento cerámico Belzona 1321", "Base recubrimiento cerámico Belzona 1321 – 0,91 kg"),
        ("Juego de punzones para grabado de números de 8 mm", "Punzones números 8 mm RS PRO"),
        ("Juego de punzones para grabado de letras de 8 mm", "Punzones letras 8 mm RS PRO"),
        ("Juego de punzones para grabado de números de 4 mm", "Punzones números 4 mm NUSAC"),
        ("Juego de punzones para grabado de letras de 4 mm", "Punzones letras 4 mm NUSAC"),
        ("Estuche con juego de 6 botadores cilíndricos", "Estuche botadores cilíndricos 2–8 mm"),
        ("Bolsa de plástico transparente con pasadores cilíndricos", "Bolsa transparente pasadores cilíndricos acero"),
        ("Bolsa de plástico transparente con tornillos", "Bolsa transparente tornillos + arandelas grandes"),
        ("Bolsita con keycap de repuesto", "Bolsita repuestos varios impresora 3D"),
        ("Magigoo Pro Kit con 6 adhesivos distintos", "Magigoo Pro Kit – 6 adhesivos + paño"),
        ("-Tuercas deslizantes con freno", "Tuercas deslizantes + tornillos Torx perfiles aluminio M6"),
        ("-Tornillos de cabeza cilíndrica con huella Allen", "Tornillos Allen cabeza cilíndrica + avellanada"),
    ]

    for fragmento, corto in reglas:
        if fragmento.casefold() in original.casefold():
            return corto

    return valor


inventario["Artículo mostrar"] = inventario["Artículo búsqueda"].apply(
    nombre_corto
)

# Las dos cajas contenedoras conservan sus nombres cortos específicos.
inventario.loc[
    inventario["Nombre contenedor"] != "",
    "Artículo mostrar"
] = inventario.loc[
    inventario["Nombre contenedor"] != "",
    "Nombre contenedor"
]

# Para compatibilidad con el resto de la aplicación:
# "Artículo" será el nombre que se muestra.
inventario["Artículo"] = inventario["Artículo mostrar"]


# =====================================================
# CANTIDAD
# =====================================================

# Guardamos primero las cantidades originales del Excel
if "Cantidad" in inventario.columns:

    cantidad_original = (
        inventario["Cantidad"]
        .copy()
        .apply(limpiar_cantidad)
    )

else:

    cantidad_original = pd.Series(
        [""] * len(inventario),
        index=inventario.index
    )


# Guardamos también el número de cajas original
if "Nº de cajas" in inventario.columns:

    cajas_originales = (
        inventario["Nº de cajas"]
        .copy()
        .apply(limpiar_cantidad)
    )

else:

    cajas_originales = pd.Series(
        [""] * len(inventario),
        index=inventario.index
    )


# Ahora sí creamos la columna que vamos a mostrar
inventario["Cantidad"] = cantidad_original.copy()


# Si la cantidad está vacía pero hay número de cajas,
# usamos el número de cajas
vacias = inventario["Cantidad"] == ""

inventario.loc[
    vacias,
    "Cantidad"
] = cajas_originales[vacias]

# Si después de todo sigue vacía, significa que
# en el inventario original había una sola unidad.
vacias = inventario["Cantidad"] == ""

inventario.loc[
    vacias,
    "Cantidad"
] = "1"
# =====================================================
# UNIDAD
# =====================================================

# Conservamos la unidad original del Excel si existe.
if "Unidad" in inventario.columns:
    unidad_original = (
        inventario["Unidad"]
        .copy()
        .apply(limpiar_texto)
    )
else:
    unidad_original = pd.Series(
        [""] * len(inventario),
        index=inventario.index
    )

# Si existe "Nº de cajas", esa fila representa cajas aunque la cantidad
# principal esté vacía. Esto es especialmente importante para tornillería.
tiene_cajas = cajas_originales != ""


def numero_de_cantidad(cantidad):
    texto = limpiar_cantidad(cantidad)
    if texto == "":
        return None

    coincidencia = re.search(r"\d+(?:[.,]\d+)?", texto)
    if not coincidencia:
        return None

    try:
        return float(coincidencia.group().replace(",", "."))
    except ValueError:
        return None


def unidad_singular_plural(unidad, cantidad):
    """
    Mantiene el tipo de unidad y corrige solo singular/plural:
    unidad -> unidad/unidades
    caja -> caja/cajas
    metro -> metro/metros
    litro -> litro/litros
    kg -> kg
    """
    u = limpiar_texto(unidad).lower()
    numero = numero_de_cantidad(cantidad)

    if numero is None:
        return u

    if u in ("unidad", "unidades"):
        return "unidad" if numero == 1 else "unidades"

    if u in ("caja", "cajas"):
        return "caja" if numero == 1 else "cajas"

    if u in ("metro", "metros"):
        return "metro" if numero == 1 else "metros"

    if u in ("litro", "litros"):
        return "litro" if numero == 1 else "litros"

    if u in ("kg", "kilogramo", "kilogramos"):
        return "kg"

    # Para otras unidades, respetamos lo que ya venía indicado.
    return u


inventario["Unidad"] = ""

for indice in inventario.index:

    cantidad_actual = inventario.at[indice, "Cantidad"]

    # PRIORIDAD 1: si "Nº de cajas" está informado, es una caja/cajas.
    if tiene_cajas.loc[indice]:
        cantidad_cajas = cajas_originales.loc[indice]
        numero_cajas = numero_de_cantidad(cantidad_cajas)

        if numero_cajas is not None:
            inventario.at[indice, "Unidad"] = (
                "caja" if numero_cajas == 1 else "cajas"
            )
        else:
            inventario.at[indice, "Unidad"] = "cajas"

    # PRIORIDAD 2: tornillería -> cajas.
    # Incluye artículos identificados como DIN 930, DIN 934, DIN 980, etc.
    # y filas que vienen de la columna "Nombre (tornillería)".
    articulo_actual = limpiar_texto(inventario.at[indice, "Artículo"])
    es_tornilleria = bool(
        ("Nombre (tornillería)" in inventario.columns
         and limpiar_texto(inventario.at[indice, "Nombre (tornillería)"]) != "")
        or re.search(r"\bDIN\s*\d+", articulo_actual, flags=re.IGNORECASE)
    )

    if es_tornilleria:
        numero = numero_de_cantidad(cantidad_actual)
        inventario.at[indice, "Unidad"] = (
            "caja" if numero == 1 else "cajas"
        )

    # PRIORIDAD 3: respetar la unidad que ya tenía el Excel.
    elif limpiar_texto(unidad_original.loc[indice]) != "":
        inventario.at[indice, "Unidad"] = unidad_singular_plural(
            unidad_original.loc[indice],
            cantidad_actual
        )

    # PRIORIDAD 4: si no hay unidad indicada, asumimos unidad.
    else:
        numero = numero_de_cantidad(cantidad_actual)
        inventario.at[indice, "Unidad"] = (
            "unidad" if numero == 1 else "unidades"
        )


# =====================================================
# POSICIÓN
# =====================================================

inventario["Posición"] = ""


if "Lugar" in inventario.columns:

    lugar = (
        inventario["Lugar"]
        .apply(limpiar_texto)
    )

    inventario["Posición"] = lugar


if "Nº balda" in inventario.columns:

    balda = (
        inventario["Nº balda"]
        .apply(limpiar_texto)
    )

    vacias = inventario["Posición"] == ""

    inventario.loc[
        vacias,
        "Posición"
    ] = balda[vacias]


# =====================================================
# CONTENIDO DE LAS DOS CAJAS
# =====================================================

contenidos = pd.DataFrame(
    [
        {
            "Artículo": "Cúter",
            "Cantidad": "1",
            "Unidad": "unidad",
            "Contenedor": "Caja herramientas impresión 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Rasqueta de pintor",
            "Cantidad": "1",
            "Unidad": "unidad",
            "Contenedor": "Caja herramientas impresión 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Rascador metálico",
            "Cantidad": "1",
            "Unidad": "unidad",
            "Contenedor": "Caja herramientas impresión 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Aguja de acupuntura",
            "Cantidad": "1",
            "Unidad": "unidad",
            "Contenedor": "Caja herramientas impresión 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Pinzas",
            "Cantidad": "3",
            "Unidad": "unidades",
            "Contenedor": "Caja herramientas impresión 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Desbarbador",
            "Cantidad": "1",
            "Unidad": "unidad",
            "Contenedor": "Caja herramientas impresión 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Destornillador de precisión",
            "Cantidad": "1",
            "Unidad": "unidad",
            "Contenedor": "Caja herramientas impresión 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Herramienta de desconexión de tubos PTFE",
            "Cantidad": "1",
            "Unidad": "unidad",
            "Contenedor": "Caja herramientas impresión 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Cable trenzado de datos",
            "Cantidad": "1",
            "Unidad": "unidad",
            "Contenedor": "Caja herramientas impresión 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Almohadillas adhesivas de espuma",
            "Cantidad": "2",
            "Unidad": "unidades",
            "Contenedor": "Caja herramientas impresión 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Perfiles de unión",
            "Cantidad": "3",
            "Unidad": "unidades",
            "Contenedor": "Caja herramientas impresión 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },

        {
            "Artículo": "Fundas de silicona",
            "Cantidad": "",
            "Unidad": "",
            "Contenedor": "Caja accesorios impresora 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Cinta aislante PVC",
            "Cantidad": "",
            "Unidad": "",
            "Contenedor": "Caja accesorios impresora 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Aceite mineral",
            "Cantidad": "1",
            "Unidad": "unidad",
            "Contenedor": "Caja accesorios impresora 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Grasa",
            "Cantidad": "1",
            "Unidad": "unidad",
            "Contenedor": "Caja accesorios impresora 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Soporte o adaptador de montaje",
            "Cantidad": "1",
            "Unidad": "unidad",
            "Contenedor": "Caja accesorios impresora 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Núcleos de impresión",
            "Cantidad": "2",
            "Unidad": "unidades",
            "Contenedor": "Caja accesorios impresora 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Patines de silicona",
            "Cantidad": "3",
            "Unidad": "unidades",
            "Contenedor": "Caja accesorios impresora 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        },
        {
            "Artículo": "Casquillos de plástico",
            "Cantidad": "3",
            "Unidad": "unidades",
            "Contenedor": "Caja accesorios impresora 3D",
            "Ubicación": "Armario 1",
            "Posición": "Balda 1"
        }
    ]
)



# =====================================================
# GESTIÓN DE MATERIAL
# =====================================================

from openpyxl import load_workbook
from datetime import date
import re


def _cantidad_numerica(valor):
    """Devuelve la cantidad como número si es claramente numérica."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None

    texto = str(valor).strip().replace(",", ".")
    if not texto:
        return None

    # Aceptamos cantidades simples y también formatos del inventario como
    # "7 (4 sin abrir)" o "4 (sin abrir)": en esos casos la cantidad
    # real de stock es el primer número.
    if re.fullmatch(r"\d+(?:\.\d+)?", texto):
        numero = float(texto)
        return int(numero) if numero.is_integer() else numero

    m = re.match(r"^(\d+(?:\.\d+)?)\s*(?:\(|$)", texto)
    if m:
        numero = float(m.group(1))
        return int(numero) if numero.is_integer() else numero

    return None


def _guardar_articulo_excel(nombre, cantidad, unidad, ubicacion, posicion):
    """
    Añade un artículo al Excel o suma la cantidad si ya existe en la misma
    ubicación y posición y la cantidad existente es numérica.
    """
    wb = load_workbook(archivo)

    if ubicacion not in wb.sheetnames:
        wb.create_sheet(ubicacion)

    ws = wb[ubicacion]

    # Detectar encabezados existentes.
    encabezados = {}
    for celda in ws[1]:
        if celda.value is not None:
            encabezados[str(celda.value).strip()] = celda.column

    # Determinar la columna del nombre.
    if "Nombre" in encabezados:
        col_nombre = encabezados["Nombre"]
        col_cantidad = encabezados.get("Cantidad")
        col_posicion = encabezados.get("Nº balda") or encabezados.get("Lugar")
        col_unidad = None
    elif "Nombre (tornillería)" in encabezados:
        col_nombre = encabezados["Nombre (tornillería)"]
        col_cantidad = encabezados.get("Nº de cajas")
        col_posicion = encabezados.get("Unnamed: 0")
        col_unidad = None
    else:
        # Hoja nueva o estructura no reconocida.
        ws.cell(1, 1).value = "Nombre"
        ws.cell(1, 2).value = "Cantidad"
        ws.cell(1, 3).value = "Nº balda"
        col_nombre, col_cantidad, col_posicion = 1, 2, 3
        col_unidad = None

    # Para Estantería esquina, la cantidad se registra como Nº de cajas.
    cantidad_para_excel = cantidad
    if unidad.lower() not in ("caja", "cajas") and "Nº de cajas" in encabezados:
        # En esa hoja el inventario está expresado en cajas.
        cantidad_para_excel = cantidad

    # Buscar un artículo existente en la misma ubicación + posición.
    fila_encontrada = None
    for fila in range(2, ws.max_row + 1):
        nombre_existente = ws.cell(fila, col_nombre).value
        posicion_existente = (
            ws.cell(fila, col_posicion).value if col_posicion else ""
        )

        if (
            limpiar_texto(nombre_existente).casefold() == nombre.strip().casefold()
            and limpiar_texto(posicion_existente).casefold()
            == posicion.strip().casefold()
        ):
            fila_encontrada = fila
            break

    if fila_encontrada is not None and col_cantidad:
        actual = _cantidad_numerica(ws.cell(fila_encontrada, col_cantidad).value)
        nueva = _cantidad_numerica(cantidad)

        if actual is not None and nueva is not None:
            total = actual + nueva
            ws.cell(fila_encontrada, col_cantidad).value = total
            wb.save(archivo)
            return "sumado"

    # Si no se puede sumar de forma segura, añadimos una nueva fila.
    fila = ws.max_row + 1
    ws.cell(fila, col_nombre).value = nombre.strip()

    if col_cantidad:
        ws.cell(fila, col_cantidad).value = cantidad

    if col_posicion:
        ws.cell(fila, col_posicion).value = posicion.strip()

    wb.save(archivo)
    return "nuevo"



def _retirar_unidades_excel(nombre, ubicacion, posicion, cantidad):
    """Resta unidades de una fila concreta del inventario."""
    cantidad = int(cantidad)
    if cantidad <= 0:
        return False, "La cantidad debe ser mayor que 0."

    wb = load_workbook(archivo)

    if ubicacion not in wb.sheetnames:
        return False, "No se ha encontrado la ubicación."

    ws = wb[ubicacion]

    encabezados = {}
    for celda in ws[1]:
        if celda.value is not None:
            encabezados[str(celda.value).strip()] = celda.column

    col_nombre = encabezados.get("Nombre") or encabezados.get("Nombre (tornillería)")
    col_posicion = encabezados.get("Nº balda") or encabezados.get("Lugar") or encabezados.get("Unnamed: 0")

    if not col_nombre:
        return False, "No se encuentra la columna del artículo en el Excel."

    for fila in range(2, ws.max_row + 1):
        nombre_existente = limpiar_texto(ws.cell(fila, col_nombre).value)
        posicion_existente = limpiar_texto(ws.cell(fila, col_posicion).value) if col_posicion else ""

        if (
            nombre_existente.casefold() == nombre.strip().casefold()
            and posicion_existente.casefold() == posicion.strip().casefold()
        ):
            # Tornillería (Nombre (tornillería) / códigos DIN) y filas que
            # realmente tengan "Nº de cajas" usan esa columna. El resto usa Cantidad.
            nombre_tornilleria = (
                limpiar_texto(ws.cell(fila, encabezados["Nombre (tornillería)"]).value)
                if "Nombre (tornillería)" in encabezados else ""
            )
            es_tornilleria = bool(nombre_tornilleria)
            es_din = bool(re.search(r"\bDIN\s*\d+", nombre_existente, flags=re.IGNORECASE))

            col_cajas = encabezados.get("Nº de cajas")
            valor_cajas = (
                ws.cell(fila, col_cajas).value
                if col_cajas else None
            )

            if col_cajas and (es_tornilleria or es_din or _cantidad_numerica(valor_cajas) is not None):
                col_cantidad = col_cajas
            else:
                col_cantidad = encabezados.get("Cantidad")

            if not col_cantidad:
                return False, "No se encuentra la columna de cantidad/cajas en el Excel."

            actual = _cantidad_numerica(ws.cell(fila, col_cantidad).value)
            actual = 0 if actual is None else int(actual)

            if cantidad > actual:
                return False, f"No puedes retirar {cantidad}: actualmente hay {actual}."

            nueva = actual - cantidad
            ws.cell(fila, col_cantidad).value = nueva
            wb.save(archivo)
            return True, f"Se han retirado {cantidad} unidades. Quedan {nueva}."

    return False, "No se ha podido localizar el artículo en el Excel."


def _eliminar_articulo_excel(nombre, ubicacion, posicion):
    """Elimina la primera fila que coincida con artículo + posición."""
    wb = load_workbook(archivo)

    if ubicacion not in wb.sheetnames:
        return False

    ws = wb[ubicacion]

    encabezados = {}
    for celda in ws[1]:
        if celda.value is not None:
            encabezados[str(celda.value).strip()] = celda.column

    col_nombre = encabezados.get("Nombre") or encabezados.get("Nombre (tornillería)")
    col_posicion = encabezados.get("Nº balda") or encabezados.get("Lugar") or encabezados.get("Unnamed: 0")

    if not col_nombre:
        return False

    for fila in range(2, ws.max_row + 1):
        nombre_existente = limpiar_texto(ws.cell(fila, col_nombre).value)
        posicion_existente = limpiar_texto(ws.cell(fila, col_posicion).value) if col_posicion else ""

        if (
            nombre_existente.casefold() == nombre.strip().casefold()
            and posicion_existente.casefold() == posicion.strip().casefold()
        ):
            ws.delete_rows(fila, 1)
            wb.save(archivo)
            return True

    return False



def _editar_articulo_excel(nombre_original, ubicacion_original, posicion_original,
                            nombre_nuevo, cantidad_nueva, ubicacion_nueva,
                            posicion_nueva):
    """Edita nombre, cantidad, ubicación y posición de un artículo."""
    cantidad_nueva = int(cantidad_nueva)
    if cantidad_nueva < 1:
        return False, "La cantidad debe ser mayor que 0."

    wb = load_workbook(archivo)

    if ubicacion_original not in wb.sheetnames:
        return False, "No se ha encontrado la ubicación original."

    ws_origen = wb[ubicacion_original]

    encabezados = {}
    for celda in ws_origen[1]:
        if celda.value is not None:
            encabezados[str(celda.value).strip()] = celda.column

    col_nombre = encabezados.get("Nombre") or encabezados.get("Nombre (tornillería)")
    col_posicion = (
        encabezados.get("Nº balda")
        or encabezados.get("Lugar")
        or encabezados.get("Unnamed: 0")
    )
    col_cantidad = encabezados.get("Cantidad") or encabezados.get("Nº de cajas")

    if not col_nombre or not col_cantidad:
        return False, "No se encuentran las columnas necesarias en el Excel."

    fila_encontrada = None
    for fila in range(2, ws_origen.max_row + 1):
        nombre_existente = limpiar_texto(ws_origen.cell(fila, col_nombre).value)
        posicion_existente = (
            limpiar_texto(ws_origen.cell(fila, col_posicion).value)
            if col_posicion else ""
        )

        if (
            nombre_existente.casefold() == nombre_original.strip().casefold()
            and posicion_existente.casefold() == posicion_original.strip().casefold()
        ):
            fila_encontrada = fila
            break

    if fila_encontrada is None:
        return False, "No se ha podido localizar el artículo en el Excel."

    # Si cambia de ubicación, copiamos la fila completa a la nueva hoja
    # y después eliminamos la original.
    if ubicacion_nueva != ubicacion_original:
        if ubicacion_nueva not in wb.sheetnames:
            wb.create_sheet(ubicacion_nueva)

        ws_destino = wb[ubicacion_nueva]

        # El destino puede tener una estructura distinta. Detectamos sus columnas.
        encabezados_destino = {}
        for celda in ws_destino[1]:
            if celda.value is not None:
                encabezados_destino[str(celda.value).strip()] = celda.column

        col_nombre_dest = (
            encabezados_destino.get("Nombre")
            or encabezados_destino.get("Nombre (tornillería)")
        )
        col_posicion_dest = (
            encabezados_destino.get("Nº balda")
            or encabezados_destino.get("Lugar")
            or encabezados_destino.get("Unnamed: 0")
        )
        col_cantidad_dest = (
            encabezados_destino.get("Cantidad")
            or encabezados_destino.get("Nº de cajas")
        )

        if not col_nombre_dest or not col_cantidad_dest:
            return False, "La nueva ubicación no tiene una estructura compatible."

        nueva_fila = ws_destino.max_row + 1

        # Copiar valores de la fila original manteniendo la estructura de la hoja destino.
        nombre_col_origen = col_nombre
        posicion_col_origen = col_posicion
        cantidad_col_origen = col_cantidad

        ws_destino.cell(nueva_fila, col_nombre_dest).value = nombre_nuevo.strip()
        ws_destino.cell(nueva_fila, col_cantidad_dest).value = cantidad_nueva

        if col_posicion_dest:
            ws_destino.cell(nueva_fila, col_posicion_dest).value = posicion_nueva.strip()

        # Copiar columnas adicionales por nombre cuando existen en ambos sitios.
        for nombre_columna, col_dest in encabezados_destino.items():
            if nombre_columna in encabezados:
                ws_destino.cell(nueva_fila, col_dest).value = ws_origen.cell(
                    fila_encontrada, encabezados[nombre_columna]
                ).value

        # Volvemos a imponer los valores editados para que no sean sobreescritos.
        ws_destino.cell(nueva_fila, col_nombre_dest).value = nombre_nuevo.strip()
        ws_destino.cell(nueva_fila, col_cantidad_dest).value = cantidad_nueva
        if col_posicion_dest:
            ws_destino.cell(nueva_fila, col_posicion_dest).value = posicion_nueva.strip()

        ws_origen.delete_rows(fila_encontrada, 1)

    else:
        ws_origen.cell(fila_encontrada, col_nombre).value = nombre_nuevo.strip()
        ws_origen.cell(fila_encontrada, col_cantidad).value = cantidad_nueva

        if col_posicion:
            ws_origen.cell(fila_encontrada, col_posicion).value = posicion_nueva.strip()

    wb.save(archivo)
    return True, "Artículo actualizado correctamente."


def _registrar_entrada_excel(proveedor, pedido, fecha_entrada, articulos):
    """Guarda el historial de una entrada en una hoja independiente."""
    wb = load_workbook(archivo)

    nombre_hoja = "Entradas"
    if nombre_hoja not in wb.sheetnames:
        ws = wb.create_sheet(nombre_hoja)
        ws.append([
            "Fecha",
            "Proveedor",
            "Nº pedido/albarán",
            "Artículo",
            "Cantidad",
            "Unidad",
            "Ubicación",
            "Posición",
        ])
    else:
        ws = wb[nombre_hoja]

    for articulo in articulos:
        ws.append([
            fecha_entrada,
            proveedor.strip(),
            pedido.strip(),
            articulo["nombre"],
            articulo["cantidad"],
            articulo["unidad"],
            articulo["ubicacion"],
            articulo["posicion"],
        ])

    wb.save(archivo)


# Cabecera principal
st.markdown("""
<div class="taller-header">
    <div class="taller-kicker">CONTROL DE MATERIAL · TALLER</div>
    <div class="taller-title">🔧 TallerStock</div>
    <div class="taller-subtitle">Inventario, entradas y organización del material del taller</div>
</div>
""", unsafe_allow_html=True)

# Botones de gestión arriba a la derecha.
col_espacio, col_anadir, col_entrada, col_gestion = st.columns([5.3, 1.35, 1.55, 1.55])

with col_espacio:
    st.empty()

with col_anadir:
    if st.button("➕ Añadir artículo", use_container_width=True):
        st.session_state["panel_gestion"] = "articulo"

with col_entrada:
    if st.button("📥 Registrar entrada", use_container_width=True):
        st.session_state["panel_gestion"] = "entrada"

with col_gestion:
    if st.button("🗑️ Gestionar inventario", use_container_width=True):
        st.session_state["panel_gestion"] = "gestionar"


panel_gestion = st.session_state.get("panel_gestion", "")


if panel_gestion == "articulo":

    cabecera_panel, cerrar_panel = st.columns([12, 1])
    with cabecera_panel:
        st.markdown('<div class="section-title">➕ Añadir artículo al inventario</div>', unsafe_allow_html=True)
    with cerrar_panel:
        if st.button("❌", key="cerrar_articulo", help="Cerrar este apartado"):
            st.session_state["panel_gestion"] = ""
            st.rerun()

    with st.form("form_nuevo_articulo", clear_on_submit=False):

        nombre_nuevo = st.text_input(
            "Artículo",
            placeholder="Ejemplo: Tornillos M6 x 30 Allen"
        )

        c1, c2, c3 = st.columns(3)

        with c1:
            cantidad_nueva = st.number_input('Cantidad', min_value=1, step=1, value=1)

        with c2:
            unidad_nueva = st.selectbox(
                "Unidad",
                ["unidad", "unidades", "caja", "cajas", "metros", "litros", "kg", "otro"]
            )

        with c3:
            ubicacion_nueva = st.selectbox(
                "Ubicación",
                excel.sheet_names
            )

        posiciones_nuevas = sorted([
            p for p in inventario[
                inventario["Ubicación"] == ubicacion_nueva
            ]["Posición"].unique()
            if limpiar_texto(p) != ""
        ])

        opcion_posicion = st.selectbox(
            "Posición",
            ["Nueva posición..."] + posiciones_nuevas
        )

        if opcion_posicion == "Nueva posición...":
            posicion_nueva = st.text_input(
                "Escribe la nueva posición",
                placeholder="Ejemplo: Balda 5 - fila 2"
            )
        else:
            posicion_nueva = opcion_posicion

        guardar_nuevo = st.form_submit_button(
            "💾 Guardar artículo",
            use_container_width=True
        )

    if guardar_nuevo:

        if not nombre_nuevo.strip():
            st.error("Escribe el nombre del artículo.")

        elif not posicion_nueva.strip():
            st.error("Indica una posición.")

        else:
            _guardar_articulo_excel(
                nombre_nuevo,
                cantidad_nueva,
                unidad_nueva,
                ubicacion_nueva,
                posicion_nueva
            )

            st.success(
                f"Artículo **{nombre_nuevo.strip()}** añadido al inventario."
            )
            st.session_state["panel_gestion"] = ""
            st.rerun()


elif panel_gestion == "entrada":

    cabecera_panel, cerrar_panel = st.columns([12, 1])
    with cabecera_panel:
        st.markdown('<div class="section-title">📥 Registrar entrada de material</div>', unsafe_allow_html=True)
    with cerrar_panel:
        if st.button("❌", key="cerrar_entrada", help="Cerrar este apartado"):
            st.session_state["panel_gestion"] = ""
            st.rerun()

    st.caption(
        "La entrada se guarda en el historial y las cantidades se suman al inventario "
        "cuando el artículo ya existe en la misma ubicación y posición."
    )

    with st.form("form_entrada_material", clear_on_submit=False):

        c1, c2, c3 = st.columns(3)

        with c1:
            proveedor = st.text_input(
                "Proveedor (opcional)",
                placeholder="Ejemplo: RS"
            )

        with c2:
            pedido = st.text_input(
                "Nº pedido / albarán",
                placeholder="Opcional"
            )

        with c3:
            fecha_entrada = st.date_input(
                "Fecha de entrada",
                value=date.today()
            )

        st.markdown("**Material recibido**")

        nombre_entrada = st.text_input(
            "Artículo",
            placeholder="Ejemplo: Tornillos M6 x 30 Allen"
        )

        c4, c5, c6 = st.columns(3)

        with c4:
            cantidad_entrada = st.number_input('Cantidad recibida', min_value=1, step=1, value=1)

        with c5:
            unidad_entrada = st.selectbox(
                "Unidad",
                ["unidad", "unidades", "caja", "cajas", "metros", "litros", "kg", "otro"]
            )

        with c6:
            ubicacion_entrada = st.selectbox(
                "Ubicación",
                excel.sheet_names
            )

        posiciones_entrada = sorted([
            p for p in inventario[
                inventario["Ubicación"] == ubicacion_entrada
            ]["Posición"].unique()
            if limpiar_texto(p) != ""
        ])

        opcion_posicion_entrada = st.selectbox(
            "Posición",
            ["Nueva posición..."] + posiciones_entrada
        )

        if opcion_posicion_entrada == "Nueva posición...":
            posicion_entrada = st.text_input(
                "Escribe la nueva posición",
                placeholder="Ejemplo: Balda 5 - fila 2"
            )
        else:
            posicion_entrada = opcion_posicion_entrada

        registrar_entrada = st.form_submit_button(
            "💾 Registrar entrada",
            use_container_width=True
        )

    if registrar_entrada:

        if not nombre_entrada.strip():
            st.error("Escribe el nombre del artículo.")

        elif not posicion_entrada.strip():
            st.error("Indica una posición.")

        else:
            articulo_entrada = {
                "nombre": nombre_entrada.strip(),
                "cantidad": cantidad_entrada,
                "unidad": unidad_entrada,
                "ubicacion": ubicacion_entrada,
                "posicion": posicion_entrada.strip(),
            }

            resultado = _guardar_articulo_excel(
                nombre_entrada,
                cantidad_entrada,
                unidad_entrada,
                ubicacion_entrada,
                posicion_entrada
            )

            _registrar_entrada_excel(
                proveedor,
                pedido,
                fecha_entrada,
                [articulo_entrada]
            )

            if resultado == "sumado":
                mensaje = (
                    f"Entrada registrada: **+{cantidad_entrada:g} {unidad_entrada}** "
                    f"de **{nombre_entrada.strip()}**. La cantidad existente se ha sumado."
                )
            else:
                mensaje = (
                    f"Entrada registrada y nuevo artículo añadido: "
                    f"**{nombre_entrada.strip()}**."
                )

            st.success(mensaje)
            st.session_state["panel_gestion"] = ""
            st.rerun()


elif panel_gestion == "gestionar":

    cabecera_panel, cerrar_panel = st.columns([12, 1])
    with cabecera_panel:
        st.markdown('<div class="section-title">🗑️ Gestionar inventario</div>', unsafe_allow_html=True)
    with cerrar_panel:
        if st.button("❌", key="cerrar_gestion", help="Cerrar este apartado"):
            st.session_state["panel_gestion"] = ""
            st.rerun()

    st.caption("Selecciona un artículo para retirar unidades o eliminarlo por completo.")

    articulos_gestionables = inventario[
        ~inventario["Es contenedor"]
    ][["Artículo", "Artículo búsqueda", "Ubicación", "Posición", "Cantidad", "Unidad"]].copy()

    if articulos_gestionables.empty:
        st.info("No hay artículos disponibles para gestionar.")
    else:
        opciones = []
        for idx, fila in articulos_gestionables.iterrows():
            opciones.append(
                (
                    idx,
                    f"{fila['Artículo']}  →  {fila['Ubicación']} → {fila['Posición']}"
                )
            )

        seleccion_idx = st.selectbox(
            "Artículo",
            [x[0] for x in opciones],
            format_func=lambda x: next(texto for idx, texto in opciones if idx == x),
            key="gestion_articulo"
        )

        seleccionado = articulos_gestionables.loc[seleccion_idx]

        st.write(
            f"**Artículo:** {seleccionado['Artículo']}  \n"
            f"**Cantidad actual:** {seleccionado['Cantidad']} {seleccionado['Unidad']}  \n"
            f"**Ubicación:** {seleccionado['Ubicación']}  \n"
            f"**Posición:** {seleccionado['Posición']}"
        )

        st.divider()

        st.markdown("### ✏️ Editar artículo")

        with st.form("form_editar_articulo", clear_on_submit=False):
            e1, e2 = st.columns(2)

            with e1:
                nombre_editado = st.text_input(
                    "Nombre",
                    value=str(seleccionado["Artículo"])
                )

                cantidad_editada = st.number_input(
                    "Cantidad",
                    min_value=1,
                    step=1,
                    value=max(1, int(_cantidad_numerica(seleccionado["Cantidad"]) or 1)),
                    key="cantidad_editada"
                )

            with e2:
                ubicacion_editada = st.selectbox(
                    "Ubicación",
                    excel.sheet_names,
                    index=(
                        excel.sheet_names.index(seleccionado["Ubicación"])
                        if seleccionado["Ubicación"] in excel.sheet_names else 0
                    ),
                    key="ubicacion_editada"
                )

                posiciones_editadas = sorted([
                    p for p in inventario[
                        inventario["Ubicación"] == ubicacion_editada
                    ]["Posición"].unique()
                    if limpiar_texto(p) != ""
                ])

                opciones_posicion_editada = ["Nueva posición..."] + posiciones_editadas

                if ubicacion_editada == seleccionado["Ubicación"]:
                    posicion_actual = str(seleccionado["Posición"])
                    if posicion_actual and posicion_actual not in opciones_posicion_editada:
                        opciones_posicion_editada.insert(1, posicion_actual)
                else:
                    posicion_actual = ""

                opcion_posicion_editada = st.selectbox(
                    "Posición",
                    opciones_posicion_editada,
                    index=(
                        opciones_posicion_editada.index(posicion_actual)
                        if posicion_actual in opciones_posicion_editada else 0
                    ),
                    key="opcion_posicion_editada"
                )

                if opcion_posicion_editada == "Nueva posición...":
                    posicion_editada = st.text_input(
                        "Escribe la nueva posición",
                        value="" if ubicacion_editada != seleccionado["Ubicación"] else str(seleccionado["Posición"]),
                        key="posicion_editada_nueva"
                    )
                else:
                    posicion_editada = opcion_posicion_editada

            guardar_edicion = st.form_submit_button(
                "💾 Guardar cambios",
                use_container_width=True
            )

        if guardar_edicion:
            if not nombre_editado.strip():
                st.error("El nombre no puede estar vacío.")
            elif not posicion_editada.strip():
                st.error("Indica una posición.")
            else:
                ok_edicion, mensaje_edicion = _editar_articulo_excel(
                    seleccionado["Artículo búsqueda"],
                    seleccionado["Ubicación"],
                    seleccionado["Posición"],
                    nombre_editado,
                    cantidad_editada,
                    ubicacion_editada,
                    posicion_editada
                )

                if ok_edicion:
                    st.success(mensaje_edicion)
                    st.session_state["panel_gestion"] = ""
                    st.rerun()
                else:
                    st.error(mensaje_edicion)

        st.divider()

        accion1, accion2 = st.columns(2)

        with accion1:
            st.markdown("### ➖ Retirar unidades")
            cantidad_retirar = st.number_input('Cantidad a retirar', min_value=1, step=1, value=1, key="cantidad_retirar")

            if st.button(
                "➖ Retirar unidades",
                key="retirar_unidades",
                use_container_width=True
            ):
                ok, mensaje = _retirar_unidades_excel(
                    seleccionado["Artículo búsqueda"],
                    seleccionado["Ubicación"],
                    seleccionado["Posición"],
                    cantidad_retirar
                )

                if ok:
                    st.success(mensaje)
                    st.session_state["panel_gestion"] = ""
                    st.rerun()
                else:
                    st.error(mensaje)

        with accion2:
            st.markdown("### 🗑️ Eliminar artículo")
            confirmar = st.checkbox(
                "Sí, quiero eliminar este artículo completo.",
                key="confirmar_eliminar"
            )

            if st.button(
                "🗑️ Eliminar artículo",
                disabled=not confirmar,
                type="primary",
                use_container_width=True,
                key="eliminar_articulo"
            ):
                eliminado = _eliminar_articulo_excel(
                    seleccionado["Artículo búsqueda"],
                    seleccionado["Ubicación"],
                    seleccionado["Posición"]
                )

                if eliminado:
                    st.success(
                        f"Artículo **{seleccionado['Artículo']}** eliminado correctamente."
                    )
                    st.session_state["panel_gestion"] = ""
                    st.rerun()
                else:
                    st.error("No se ha podido localizar el artículo en el Excel.")


st.divider()

# =====================================================
# BUSCADOR
# =====================================================

st.markdown('<div class="section-title">🔎 Buscar en todo el taller</div>', unsafe_allow_html=True)

busqueda = st.text_input(
    "Escribe el nombre o parte del nombre del artículo:",
    placeholder="Ejemplo: pinzas, cúter, destornillador, DIN 934..."
)


# =====================================================
# FILTRO DE UBICACIÓN
# =====================================================

ubicacion = st.selectbox(
    "📍 Filtrar por ubicación:",
    ["Todas"] + excel.sheet_names
)


# =====================================================
# FILTRO DE POSICIÓN
# =====================================================

if ubicacion == "Todas":

    posiciones_disponibles = sorted(
        [
            posicion
            for posicion in inventario["Posición"].unique()
            if posicion != ""
        ]
    )

else:

    posiciones_disponibles = sorted(
        [
            posicion
            for posicion in inventario[
                inventario["Ubicación"] == ubicacion
            ]["Posición"].unique()
            if posicion != ""
        ]
    )


posicion = st.selectbox(
    "📚 Filtrar por posición:",
    ["Todas"] + posiciones_disponibles
)


# =====================================================
# PREPARAR RESULTADOS
# =====================================================

resultados_articulos = inventario.copy()
resultados_cajas = contenidos.copy()


# =====================================================
# BUSCAR
# =====================================================

if busqueda:

    texto_busqueda = busqueda.strip()


    # -------------------------------------------------
    # ARTÍCULOS PRINCIPALES
    # -------------------------------------------------

    resultados_articulos = inventario[
        inventario["Artículo búsqueda"].str.contains(
            texto_busqueda,
            case=False,
            na=False,
            regex=False
        )
    ].copy()


    # -------------------------------------------------
    # CONTENIDO DE CAJAS
    # -------------------------------------------------

    resultados_cajas = contenidos[
        contenidos["Artículo"].str.contains(
            texto_busqueda,
            case=False,
            na=False,
            regex=False
        )
    ].copy()


    # -------------------------------------------------
    # NO MOSTRAR LA CAJA SI HEMOS ENCONTRADO
    # ALGO QUE ESTÁ DENTRO DE ELLA
    # -------------------------------------------------

    if not resultados_cajas.empty:

        contenedores_encontrados = (
            resultados_cajas["Contenedor"]
            .dropna()
            .astype(str)
            .tolist()
        )

        resultados_articulos = resultados_articulos[
            ~(
                resultados_articulos["Es contenedor"]
                &
                resultados_articulos["Nombre contenedor"].isin(
                    contenedores_encontrados
                )
            )
        ]


# =====================================================
# APLICAR FILTROS
# =====================================================

if ubicacion != "Todas":

    resultados_articulos = resultados_articulos[
        resultados_articulos["Ubicación"] == ubicacion
    ]

    resultados_cajas = resultados_cajas[
        resultados_cajas["Ubicación"] == ubicacion
    ]


if posicion != "Todas":

    resultados_articulos = resultados_articulos[
        resultados_articulos["Posición"] == posicion
    ]

    resultados_cajas = resultados_cajas[
        resultados_cajas["Posición"] == posicion
    ]


# =====================================================
# MOSTRAR RESULTADOS DE BÚSQUEDA
# =====================================================

if busqueda:

    total = (
        len(resultados_articulos)
        +
        len(resultados_cajas)
    )

    st.write(
        f"Encontrados: **{total}** resultados"
    )


    if not resultados_articulos.empty:

        st.markdown('<div class="section-title">📦 Inventario</div>', unsafe_allow_html=True)

        tabla_inventario = resultados_articulos[
            [
                "Artículo mostrar",
                "Cantidad",
                "Unidad",
                "Ubicación",
                "Posición"
            ]
        ].copy()

        tabla_inventario = tabla_inventario.rename(columns={"Artículo mostrar": "Artículo"})

        st.dataframe(
            tabla_inventario,
            use_container_width=True,
            hide_index=True
        )


    if not resultados_cajas.empty:

        st.markdown('<div class="section-title">🧰 Dentro de cajas y estuches</div>', unsafe_allow_html=True)

        tabla_cajas = resultados_cajas[
            [
                "Artículo",
                "Cantidad",
                "Unidad",
                "Contenedor",
                "Ubicación",
                "Posición"
            ]
        ].copy()

        st.dataframe(
            tabla_cajas,
            use_container_width=True,
            hide_index=True
        )


    if total == 0:

        st.warning(
            "No se ha encontrado ningún artículo."
        )


# =====================================================
# EXPLORAR INVENTARIO
# =====================================================

else:

    st.markdown('<div class="section-title">📋 Explorar inventario</div>', unsafe_allow_html=True)

    if ubicacion == "Todas":

        st.info(
            "Selecciona una ubicación arriba para consultar "
            "su inventario."
        )

    else:

        st.write(
            f"**{ubicacion}**"
            + (
                ""
                if posicion == "Todas"
                else f" → **{posicion}**"
            )
        )

        if resultados_articulos.empty:

            st.warning(
                "No hay artículos en esta posición."
            )

        else:

            tabla_explorar = resultados_articulos[
                [
                    "Artículo mostrar",
                    "Cantidad",
                    "Unidad",
                    "Ubicación",
                    "Posición"
                ]
            ].copy()

            tabla_explorar = tabla_explorar.rename(columns={"Artículo mostrar": "Artículo"})

            st.dataframe(
                tabla_explorar,
                use_container_width=True,
                hide_index=True
            )

# ===== TALLERSTOCK — fondo gris azulado industrial =====
st.markdown("""
<style>
/* ===== TALLERSTOCK — diseño gris azulado limpio ===== */
.stApp {
    background:
        radial-gradient(circle at 50% -10%, rgba(255,255,255,.16), transparent 38%),
        linear-gradient(135deg, #6b7782 0%, #606c77 50%, #586570 100%) !important;
}

/* Contenedor principal */
.main .block-container {
    max-width: 1180px;
    padding-top: 2rem;
    padding-bottom: 3rem;
}

/* Cabecera / tarjetas */
[data-testid="stVerticalBlockBorderWrapper"],
[data-testid="stExpander"] {
    border-color: rgba(255,255,255,.16) !important;
}

/* Inputs y desplegables: más claros y limpios */
div[data-baseweb="input"] > div,
div[data-baseweb="select"] > div {
    background-color: rgba(249,250,251,.97) !important;
    border-color: rgba(30,41,59,.16) !important;
}

div[data-baseweb="input"] input,
div[data-baseweb="select"] input {
    color: #202733 !important;
}

div[data-baseweb="select"] span {
    color: #202733 !important;
}

/* Texto general sobre el fondo */
.stApp p, .stApp label, .stApp .stMarkdown {
    color: #f4f6f8;
}

/* Separadores más discretos */
hr {
    border-color: rgba(255,255,255,.16) !important;
}

/* Botones principales */
.stButton > button {
    border-radius: 10px !important;
    min-height: 48px !important;
    border: 1px solid rgba(255,255,255,.16) !important;
    box-shadow: 0 5px 14px rgba(15,23,42,.18) !important;
}

/* Panel de información inferior: texto oscuro para que sea legible */
div[data-testid="stAlert"] {
    background: rgba(226, 238, 251, .94) !important;
    border: 1px solid rgba(77, 132, 190, .35) !important;
    color: #25364a !important;
}

div[data-testid="stAlert"] p,
div[data-testid="stAlert"] span {
    color: #25364a !important;
}

/* Títulos */
h1, h2, h3 {
    letter-spacing: .1px;
}
</style>
""", unsafe_allow_html=True)
