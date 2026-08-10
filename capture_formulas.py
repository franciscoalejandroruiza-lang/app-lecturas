"""
capture_formulas.py
Motor GENÉRICO para calcular Carta B&N, Oficio B&N, Carta Color, Oficio
Color y Digitalización a partir de los valores crudos del reporte de cada
modelo de impresora.

En vez de una función de Python distinta por cada perfil, cada modelo se
describe de forma declarativa en MODEL_PROFILES: qué campos captura el
usuario, y con qué "receta" se calcula cada valor final. Un solo loop
genérico (compute_final_values) interpreta esa receta para cualquier
modelo. Para agregar un modelo nuevo, solo hay que agregar su entrada aquí
— no hace falta escribir código nuevo.

Tipos de receta soportados por salida (carta_bn/oficio_bn/carta_color/
oficio_color/digitalizacion):
  ("directo", campo)                  -> se usa ese campo tal cual
  ("suma", [campo1, campo2, ...])      -> se suman esos campos
  ("resta_de", total_campo, [c1, c2])  -> total_campo - suma(c1, c2, ...)
Si una salida no aplica a un modelo, simplemente no se incluye en "outputs"
y queda en None.
"""


def _norm_model(modelo):
    if not modelo:
        return ""
    return str(modelo).upper().replace("-", "").replace(" ", "").strip()


MODEL_PROFILES = {

    "MXB450P": {
        "perfil": "A", "nombre": "Perfil A — MX-B450P (solo Carta B/N)",
        "verificado": True,
        "fields": [
            ("contador_total", "Contador Total", "Lista del estado de la máquina."),
        ],
        "outputs": {
            "carta_bn": ("directo", "contador_total"),
        },
    },

    "MXB467P": {
        "perfil": "A", "nombre": "Perfil A — MX-B467P (Carta + Oficio B/N)",
        "verificado": True,
        "fields": [
            ("carta_normal", "Carta Normal (= Carta B/N)", "Estadísticas dispositivo — se usa directo, sin restar nada."),
            ("legal_normal", "Legal Normal", "Estadísticas dispositivo."),
            ("oficio_mexico", "Oficio México", "Estadísticas dispositivo."),
        ],
        "outputs": {
            "carta_bn": ("directo", "carta_normal"),
            "oficio_bn": ("suma", ["legal_normal", "oficio_mexico"]),
        },
    },

    "MXB476P": {
        "perfil": "A", "nombre": "Perfil A — MX-B476P (Carta + Oficio B/N)",
        "verificado": True,
        "fields": [
            ("carta_normal", "Carta Normal (= Carta B/N)", "Estadísticas dispositivo — se usa directo, sin restar nada."),
            ("legal_normal", "Legal Normal", "Estadísticas dispositivo."),
            ("oficio_mexico", "Oficio México", "Estadísticas dispositivo."),
        ],
        "outputs": {
            "carta_bn": ("directo", "carta_normal"),
            "oficio_bn": ("suma", ["legal_normal", "oficio_mexico"]),
        },
    },

    "MXB376WH": {
        "perfil": "K", "nombre": "Perfil K — MX-B376WH (Carta+Oficio B/N y Digitalización)",
        "verificado": True,
        "fields": [
            ("contador_total", "Contador total", "Contador de trabajos."),
            ("otras_impresiones", "Otras impresiones (= Oficio B/N)", "Contador de trabajos."),
            ("envio_escaner", "Envío de escáner", "Contador de trabajos — suma tú Blanco y Negro + Color y pon un solo total."),
            ("escanear_disco", "Escanear a disco duro", "Contador de trabajos — suma tú Blanco y Negro + Color y pon un solo total."),
        ],
        "outputs": {
            "carta_bn": ("resta_de", "contador_total", ["otras_impresiones"]),
            "oficio_bn": ("directo", "otras_impresiones"),
            "digitalizacion": ("suma", ["envio_escaner", "escanear_disco"]),
        },
    },

    "MXB707F": {
        "perfil": "L", "nombre": "Perfil L — MX-B707F (Carta+Oficio B/N y Digitalización)",
        "verificado": True,
        "fields": [
            ("total_digitalizaciones", "Informe Activo (Total de digitalizaciones)", "Informe de activo — directo, no se calcula."),
            ("legal_normal", "Estadísticas de dispositivo (Legal Normal)", "Estadísticas dispositivo."),
            ("legal_tipo2", "Estadísticas de dispositivo (Legal tipo personalizado 2)", "Estadísticas dispositivo."),
            ("total", "Total", "Estadísticas dispositivo."),
        ],
        "outputs": {
            "carta_bn": ("resta_de", "total", ["legal_normal", "legal_tipo2"]),
            "oficio_bn": ("suma", ["legal_normal", "legal_tipo2"]),
            "digitalizacion": ("directo", "total_digitalizaciones"),
        },
    },

    "MXC507P": {
        "perfil": "D", "nombre": "Perfil D — MX-C507P (Carta+Oficio B/N y Color, sin Digitalización)",
        "verificado": True,
        "fields": [
            ("legal_normal_bn", "Estadísticas dispositivo (Legal Normal) Blanco y negro", "Estadísticas dispositivo."),
            ("legal_tipo2_bn", "Estadísticas dispositivo (Legal tipo personalizado 2) Blanco y negro", "Estadísticas dispositivo."),
            ("total_bn", "Total Blanco y negro", "Estadísticas dispositivo."),
            ("legal_normal_color", "Estadísticas dispositivo (Legal Normal) Color", "Estadísticas dispositivo."),
            ("legal_tipo2_color", "Estadísticas dispositivo (Legal tipo personalizado 2) Color", "Estadísticas dispositivo."),
            ("total_color", "Total Color", "Estadísticas dispositivo."),
        ],
        "outputs": {
            "carta_bn": ("resta_de", "total_bn", ["legal_normal_bn", "legal_tipo2_bn"]),
            "oficio_bn": ("suma", ["legal_normal_bn", "legal_tipo2_bn"]),
            "carta_color": ("resta_de", "total_color", ["legal_normal_color", "legal_tipo2_color"]),
            "oficio_color": ("suma", ["legal_normal_color", "legal_tipo2_color"]),
        },
    },

    "MXC407P": {
        "perfil": "D", "nombre": "Perfil D — MX-C407P (⚠️ fórmula sin confirmar, igual a MX-C507P)",
        "verificado": False,
        "fields": [
            ("legal_normal_bn", "Estadísticas dispositivo (Legal Normal) Blanco y negro", "Estadísticas dispositivo."),
            ("legal_tipo2_bn", "Estadísticas dispositivo (Legal tipo personalizado 2) Blanco y negro", "Estadísticas dispositivo."),
            ("total_bn", "Total Blanco y negro", "Estadísticas dispositivo."),
            ("legal_normal_color", "Estadísticas dispositivo (Legal Normal) Color", "Estadísticas dispositivo."),
            ("legal_tipo2_color", "Estadísticas dispositivo (Legal tipo personalizado 2) Color", "Estadísticas dispositivo."),
            ("total_color", "Total Color", "Estadísticas dispositivo."),
        ],
        "outputs": {
            "carta_bn": ("resta_de", "total_bn", ["legal_normal_bn", "legal_tipo2_bn"]),
            "oficio_bn": ("suma", ["legal_normal_bn", "legal_tipo2_bn"]),
            "carta_color": ("resta_de", "total_color", ["legal_normal_color", "legal_tipo2_color"]),
            "oficio_color": ("suma", ["legal_normal_color", "legal_tipo2_color"]),
        },
    },
    "MXC528P": {
        "perfil": "D", "nombre": "Perfil D — MX-C528P (⚠️ fórmula sin confirmar, igual a MX-C507P)",
        "verificado": False,
        "fields": [
            ("legal_normal_bn", "Estadísticas dispositivo (Legal Normal) Blanco y negro", "Estadísticas dispositivo."),
            ("legal_tipo2_bn", "Estadísticas dispositivo (Legal tipo personalizado 2) Blanco y negro", "Estadísticas dispositivo."),
            ("total_bn", "Total Blanco y negro", "Estadísticas dispositivo."),
            ("legal_normal_color", "Estadísticas dispositivo (Legal Normal) Color", "Estadísticas dispositivo."),
            ("legal_tipo2_color", "Estadísticas dispositivo (Legal tipo personalizado 2) Color", "Estadísticas dispositivo."),
            ("total_color", "Total Color", "Estadísticas dispositivo."),
        ],
        "outputs": {
            "carta_bn": ("resta_de", "total_bn", ["legal_normal_bn", "legal_tipo2_bn"]),
            "oficio_bn": ("suma", ["legal_normal_bn", "legal_tipo2_bn"]),
            "carta_color": ("resta_de", "total_color", ["legal_normal_color", "legal_tipo2_color"]),
            "oficio_color": ("suma", ["legal_normal_color", "legal_tipo2_color"]),
        },
    },
    "MXM7570": {
        "perfil": "M", "nombre": "Perfil M — MX-M7570 (⚠️ fórmula sin confirmar, igual a MX-B376WH)",
        "verificado": False,
        "fields": [
            ("contador_total", "Contador total", "Contador de trabajos."),
            ("otras_impresiones", "Otras impresiones (= Oficio B/N)", "Contador de trabajos."),
            ("envio_escaner", "Envío de escáner", "Contador de trabajos — suma tú Blanco y Negro + Color y pon un solo total."),
            ("escanear_disco", "Escanear a disco duro", "Contador de trabajos — suma tú Blanco y Negro + Color y pon un solo total."),
        ],
        "outputs": {
            "carta_bn": ("resta_de", "contador_total", ["otras_impresiones"]),
            "oficio_bn": ("directo", "otras_impresiones"),
            "digitalizacion": ("suma", ["envio_escaner", "escanear_disco"]),
        },
    },
    "MXC407F": {
        "perfil": "D", "nombre": "Perfil D — MX-C407F multifuncional (⚠️ fórmula sin confirmar)",
        "verificado": False,
        "fields": [
            ("total_digitalizaciones", "Informe Activo (Total de digitalizaciones)", "Informe de activo — directo."),
            ("legal_normal_bn", "Estadísticas dispositivo (Legal Normal) Blanco y negro", "Estadísticas dispositivo."),
            ("legal_tipo2_bn", "Estadísticas dispositivo (Legal tipo personalizado 2) Blanco y negro", "Estadísticas dispositivo."),
            ("total_bn", "Total Blanco y negro", "Estadísticas dispositivo."),
            ("legal_normal_color", "Estadísticas dispositivo (Legal Normal) Color", "Estadísticas dispositivo."),
            ("legal_tipo2_color", "Estadísticas dispositivo (Legal tipo personalizado 2) Color", "Estadísticas dispositivo."),
            ("total_color", "Total Color", "Estadísticas dispositivo."),
        ],
        "outputs": {
            "carta_bn": ("resta_de", "total_bn", ["legal_normal_bn", "legal_tipo2_bn"]),
            "oficio_bn": ("suma", ["legal_normal_bn", "legal_tipo2_bn"]),
            "carta_color": ("resta_de", "total_color", ["legal_normal_color", "legal_tipo2_color"]),
            "oficio_color": ("suma", ["legal_normal_color", "legal_tipo2_color"]),
            "digitalizacion": ("directo", "total_digitalizaciones"),
        },
    },
}

OUTPUT_LABELS = {
    "carta_bn": "Carta B&N", "oficio_bn": "Oficio B&N",
    "carta_color": "Carta Color", "oficio_color": "Oficio Color",
    "digitalizacion": "Digitalización",
}


def get_profile_for_model(modelo):
    return MODEL_PROFILES.get(_norm_model(modelo))


def compute_final_values(modelo, raw_values):
    """Motor genérico: interpreta la 'receta' de cada salida con un loop,
    sin necesitar una función por modelo."""
    profile = get_profile_for_model(modelo)
    if not profile:
        raise ValueError(f"Modelo '{modelo}' no está documentado en la guía de fórmulas.")

    result = {k: None for k in OUTPUT_LABELS}
    for salida, receta in profile["outputs"].items():
        tipo = receta[0]
        if tipo == "directo":
            campo = receta[1]
            result[salida] = raw_values.get(campo)
        elif tipo == "suma":
            campos = receta[1]
            result[salida] = sum((raw_values.get(c) or 0) for c in campos)
        elif tipo == "resta_de":
            total_campo, campos_a_restar = receta[1], receta[2]
            total = raw_values.get(total_campo) or 0
            resta = sum((raw_values.get(c) or 0) for c in campos_a_restar)
            result[salida] = total - resta
        else:
            raise ValueError(f"Tipo de receta desconocido: {tipo}")
    return result
