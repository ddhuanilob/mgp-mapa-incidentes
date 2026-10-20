"""
Tabla de actores a partir del mapa "Extorsión y Crímenes Violentos - Perú [2026]".

Cada incidente del CSV trae un titular corto, por ejemplo:
    "Sicarios asesinan a chofer de empresa "Etmosa" en Av. Paramonga."
Este script lee esos titulares con reglas (expresiones regulares y diccionarios)
y genera dos tablas:

    salida/incidentes.csv  -> una fila por incidente (fecha, tipo, lugar, fuente...)
    salida/actores.csv     -> una fila por víctima o victimario, con sus atributos
    salida/actores.xlsx    -> las dos tablas en un Excel

Uso:
    python src/extraer_actores.py [ruta_csv] [carpeta_salida]
"""

import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]
CSV_POR_DEFECTO = RAIZ / "data" / "extorsion_y_crimenes_violentos_peru_2026.csv"
SALIDA_POR_DEFECTO = RAIZ / "salida"

# ---------------------------------------------------------------------------
# Diccionarios
# ---------------------------------------------------------------------------

NUMEROS = {"un": 1, "una": 1, "uno": 1, "dos": 2, "tres": 3, "cuatro": 4,
           "cinco": 5, "seis": 6, "siete": 7, "ocho": 8, "nueve": 9, "diez": 10}

# palabra -> sexo
SEXO = {
    "hombre": "Masculino", "hombres": "Masculino", "padre": "Masculino",
    "hijo": "Masculino", "niño": "Masculino", "señor": "Masculino",
    "esposo": "Masculino", "hermano": "Masculino", "abuelo": "Masculino",
    "mujer": "Femenino", "mujeres": "Femenino", "madre": "Femenino",
    "hija": "Femenino", "niña": "Femenino", "señora": "Femenino",
    "esposa": "Femenino", "hermana": "Femenino", "abuela": "Femenino",
    "embarazada": "Femenino",
}
# sustantivos de oficio con forma femenina / masculina explícita
OFICIO_FEMENINO = {"pasajera", "vendedora", "trabajadora", "dueña", "abogada",
                   "conductora", "cobradora", "coordinadora", "comerciante_f",
                   "estudiante_f", "pasajeras", "trabajadoras"}
OFICIO_MASCULINO = {"pasajero", "vendedor", "trabajador", "dueño", "abogado",
                    "conductor", "cobrador", "coordinador", "empresario",
                    "obrero", "colectivero", "ingeniero", "técnico", "barbero",
                    "pescador", "amigo", "recluso", "estafador", "jalador",
                    "repartidor", "llantero", "emolientero", "productor",
                    "pasajeros", "trabajadores", "obreros"}

NACIONALIDADES = {
    "venezolan": "Venezolana", "peruan": "Peruana", "colombian": "Colombiana",
    "paraguay": "Paraguaya", "ecuatorian": "Ecuatoriana", "chilen": "Chilena",
    "bolivian": "Boliviana", "extranjer": "Extranjera (sin especificar)",
}

# raíz de la ocupación -> (ocupación normalizada, rubro)
OCUPACIONES = [
    (r"chofer|choferes|conductor[a]?", "Chofer / conductor", "Transporte"),
    (r"colectivero", "Colectivero", "Transporte"),
    (r"mototaxista", "Mototaxista", "Transporte"),
    (r"taxista", "Taxista", "Transporte"),
    (r"cobrador[a]?", "Cobrador(a)", "Transporte"),
    (r"jalador", "Jalador de combis", "Transporte"),
    (r"copiloto", "Copiloto", "Transporte"),
    (r"pasajer[oa]s?", "Pasajero(a)", "Transporte"),
    (r"repartidor", "Repartidor", "Reparto / delivery"),
    (r"cambista", "Cambista", "Finanzas informales"),
    (r"prestamista", "Prestamista", "Finanzas informales"),
    (r"comerciante", "Comerciante", "Comercio"),
    (r"vendedor[a]?|ambulante", "Vendedor(a)", "Comercio"),
    (r"emolientero", "Emolientero", "Comercio"),
    (r"llantero", "Llantero", "Comercio / servicios"),
    (r"barbero", "Barbero", "Comercio / servicios"),
    (r"due[ñn][oa]", "Dueño(a) de negocio", "Comercio / servicios"),
    (r"empresari[oa]", "Empresario(a)", "Empresarial"),
    (r"gerente", "Gerente", "Empresarial"),
    (r"director[a]?", "Director(a)", "Empresarial"),
    (r"polic[ií]as?|suboficial|pnp", "Policía", "Seguridad"),
    (r"serenazgo|vigilante|seguridad|guardaespaldas", "Seguridad / vigilancia", "Seguridad"),
    (r"cantante|animador|m[uú]sico|orquesta|productor", "Artista / música", "Entretenimiento"),
    (r"obrer[oa]s?|alba[ñn]il(?:es)?", "Obrero / albañil", "Construcción"),
    (r"trabajador[a]?(?:es)?", "Trabajador(a)", "Sin dato"),
    (r"dirigente|dirigentes", "Dirigente", "Dirigencia social / política"),
    (r"alcalde|regidor|candidat[oa]|procurador|fiscalizador|presidente", "Autoridad / político", "Dirigencia social / política"),
    (r"pescador", "Pescador", "Pesca"),
    (r"periodista", "Periodista", "Prensa"),
    (r"pastor", "Pastor", "Religión"),
    (r"abogad[oa]", "Abogado(a)", "Legal"),
    (r"pediatra|m[eé]dic[oa]|enfermer[oa]", "Personal de salud", "Salud"),
    (r"estudiante|escolar", "Estudiante", "Educación"),
    (r"futbolista|campe[oó]n", "Deportista", "Deporte"),
    (r"ingenier[oa]|t[eé]cnico", "Profesional / técnico", "Profesional"),
    (r"recluso", "Recluso", "Sin dato"),
    (r"estafador", "Estafador", "Sin dato"),
]

# objetivos que no son personas: raíz -> (tipo de actor, rubro)
OBJETIVOS = [
    (r"bus(?:es)?|combis?|couster|cúster|colectivo|mototaxis?|camioneta|auto|camión|unidad(?:es)?|veh[ií]culo",
     "Vehículo", "Transporte"),
    (r"empresa", "Empresa / negocio", "Sin dato"),
    (r"terminal|agencia", "Empresa / negocio", "Transporte"),
    (r"local|tienda|ferreter[ií]a|restaurante|poller[ií]a|cevicher[ií]a|anticucher[ií]a|bar|discoteca|"
     r"licorer[ií]a|minimarket|bodega|grifo|taller|carpinter[ií]a|av[ií]cola|dep[oó]sito|oficina|"
     r"hotel|hostal|cl[ií]nica|chatarrer[ií]a|mercado|galer[ií]a|negocio|f[aá]brica|lavander[ií]a|"
     r"casa de cambio|panader[ií]a|farmacia|botica|gimnasio|gym|spa|obras?|chifa|[oó]ptica|club|imprenta|"
     r"comercio|cajeros?|trailer|restobar|construcci[oó]n|puesto|mec[aá]nica|lubricentro|cantera|minera|"
     r"centro de acopio|caja|banco|bcp|interbank|cooperativa|ladrillera|empresas?",
     "Empresa / negocio", "Comercio / servicios"),
    (r"colegio|instituto|universidad", "Institución", "Educación"),
    (r"comisar[ií]a|municipalidad|penal|corte|inpe|reniec|gore|odpe|fiscal[ií]a|puesto|"
     r"juzgado|ministerio|hospital|posta", "Institución", "Estado"),
    (r"iglesia", "Institución", "Religión"),
    (r"viviendas?|casas?|edificio|domicilio|departamento", "Vivienda", "Residencial"),
]

# sujeto al inicio del titular -> tipo de victimario
VICTIMARIOS = [
    (r"sicarios?", "Sicario"),
    (r"extorsionador(?:es)?", "Extorsionador"),
    (r"ladr(?:ón|on|ones)", "Ladrón"),
    (r"delincuentes?", "Delincuente"),
    (r"barristas?", "Barrista"),
    (r"terroristas?", "Grupo armado / terrorista"),
    (r"manifestantes", "Manifestantes"),
    (r"escolares", "Escolares"),
    (r"ronderos", "Ronderos"),
    (r"sujetos?|desconocidos?|un sujeto|grupo de \w+|\d+ sujetos", "Sujeto(s) no identificado(s)"),
    (r"hombres?|mujer(?:es)?|adolescentes?|escolar(?:es)?|emolientero|un hombre|drogadictos?|vecinos?|"
     r"comandante|polic[ií]as?|serenos?|vigilante|cobrador|chofer|jaladoras?", "Persona particular"),
]

VERBOS_FATALES = (r"asesinan|asesina|asesinó|asesinaron|matan|mata|mató|ejecutan|persiguen y asesinan|"
                  r"acribillan|acribilla|acribillaron|acribilló")
VERBOS_HERIDOS = (r"balean|balea|baleó|hieren|hiere|hirieron|acuchillan|acuchilla|apuñalan|apuñala|"
                  r"agreden|agrede|atacan|ataca|disparan contra|dispara contra|disparan a|dispara a|disparó a|persiguen a balazos a|golpean|queman|incendian|incendia|incendió|"
                  r"incencian|prenden fuego|prende fuego|desfigura|intentó asesinar|intentan asesinar|"
                  r"lanzan (?:bomba molotov|bomba incendiaria|molotov|explosivo|granada|lacrimógena|lacrimogena)|"
                  r"arrojan (?:bomba molotov|bomba incendiaria|molotov|explosivo|granada|lacrimógena|lacrimogena)")
VERBOS_SECUESTRO = r"secuestran|secuestra|secuestraron|intentan secuestrar|intentó secuestrar"
VERBOS_AMENAZA = (r"amenazan|amenaza|exigen [\w\s]+? soles|exige [\w\s]+? soles|amenazas? contra|"
                  r"(?:mensajes?|cartas?|notas?) (?:con amenazas|amenazantes?|extorsivas?)(?: en| contra)?")

# palabras donde termina la frase de la víctima
FIN_FRASE = (r"(?=\s+(?:en|cerca|frente|dentro|durante|tras|mientras|por|delante|afuera|"
             r"saliendo|que|a espaldas|con|cuando|al interior|para|desde|hasta|junto|"
             r"y (?:hieren|dejan|asesinan|balean|acuchillan|roban|huyen|escapan|queman|"
             r"secuestran|detonan|hiere|deja)|dejando|hiriendo|a balazos|a puñaladas)\b|[.,;:]|$)")

MODIFICADORES = r"(?:(?:de bala|de un balazo|en la cabeza|en la pierna|en la cara|de gravedad|gravemente|a balazos|a quemarropa|por la espalda)\s+)?"


def quitar_tildes(texto):
    return "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")


# ---------------------------------------------------------------------------
# Incidentes
# ---------------------------------------------------------------------------

def cargar_incidentes(ruta_csv):
    crudo = pd.read_csv(ruta_csv)
    partes = crudo["description"].fillna("").str.strip().str.split("\n")

    inc = pd.DataFrame()
    inc["id_incidente"] = [f"I{i + 1:04d}" for i in range(len(crudo))]
    inc["fecha"] = pd.to_datetime(
        crudo["description"].str.extract(r"(\d{2}-\d{2}-\d{4})")[0],
        format="%d-%m-%Y", errors="coerce")
    inc["titular"] = partes.apply(lambda p: p[1].strip() if len(p) > 1 else p[0].strip())
    inc["tipo_evento"] = [normalizar_tipo(t, tit) for t, tit in zip(crudo["name"], inc["titular"])]
    inc["fuente_url"] = crudo["description"].str.extract(r"\[\[(.+?)\]\]")[0]
    inc["fuente_medio"] = inc["fuente_url"].str.extract(r"https?://(?:www\.|m\.)?([^/]+)")[0]
    inc["latitud"] = crudo["Latitude"]
    inc["longitud"] = crudo["Longitude"]
    inc["ambito"] = [ambito(la, lo) for la, lo in zip(inc["latitud"], inc["longitud"])]
    inc["movil_presunto"] = inc["titular"].apply(movil_presunto)
    inc["medio_empleado"] = [medio_empleado(t, e) for t, e in zip(inc["titular"], inc["tipo_evento"])]
    return inc


def normalizar_tipo(nombre, titular):
    nombre = str(nombre).strip()
    if nombre == "Assesinato":
        return "Asesinato"
    if re.fullmatch(r"\d{2}-\d{2}-\d{4}", nombre):  # fecha puesta en lugar del tipo
        t = titular.lower()
        for clave, tipo in [("balacera", "Balacera"), ("balean", "Balacera"),
                            ("asesin", "Asesinato"), ("explosiv", "Explosivo"),
                            ("extors", "Extorsión"), ("secuestr", "Secuestro")]:
            if clave in t:
                return tipo
        return "Sin clasificar"
    return nombre


def ambito(lat, lon):
    # Recuadro aproximado de Lima Metropolitana y Callao
    if -12.55 <= lat <= -11.55 and -77.30 <= lon <= -76.60:
        return "Lima Metropolitana y Callao (aprox.)"
    return "Otras regiones"


def movil_presunto(titular):
    t = titular.lower()
    if re.search(r"extors|cupo|carta|cobro de|negarse a pagar|gota a gota", t):
        return "Extorsión"
    if re.search(r"\brob(?:o|ar|an|aron)\b|asalt|ladr", t):
        return "Robo"
    if re.search(r"entre grupos|enfrentamiento|pelea|disputa|barristas|ajuste", t):
        return "Enfrentamiento / disputa"
    if re.search(r"sicari", t):
        return "Sicariato (móvil no indicado)"
    return "Sin dato"


def medio_empleado(titular, tipo):
    t = titular.lower()
    if re.search(r"explosiv|dinamita|detonan|granada|artefacto", t):
        return "Explosivo"
    if re.search(r"acuchill|apuñal|puñalad|cuchill|arma blanca|cortes", t):
        return "Arma blanca"
    if re.search(r"incendi|queman|quemad|fuego(?! cruzado)", t) and "balazo" not in t:
        return "Incendio"
    if re.search(r"bal|dispar|sicari|arma", t) or tipo in ("Balacera", "Asesinato"):
        return "Arma de fuego"
    return "Sin dato"


# ---------------------------------------------------------------------------
# Víctimas
# ---------------------------------------------------------------------------

def a_numero(palabra):
    palabra = palabra.lower()
    if palabra.isdigit():
        return int(palabra)
    return NUMEROS.get(palabra)


def describir_frase(frase):
    """Atributos de una frase nominal como 'joven venezolano de 17 años'."""
    original = frase.strip()
    f = original.lower()
    f = re.sub(r"^(?:al|a|su|sus|la|el|los|las|otros?|otras?|ex)\s+", "", f)
    f = re.sub(r"^(?:otros?|otras?|ex)\s+", "", f)
    datos = {"tipo_actor": "Persona", "cantidad": 1, "es_grupo": False,
             "nombre_o_alias": "", "sexo": "Sin dato", "edad": None,
             "grupo_etario": "Sin dato", "nacionalidad": "Sin dato",
             "ocupacion": "Sin dato", "rubro": "Sin dato",
             "organizacion_vinculada": "", "antecedentes": "Sin dato"}

    # cantidad
    m = re.match(r"(\d+|un|una|uno|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez)\b\s*(.*)", f)
    if m and a_numero(m.group(1)) and not re.match(r"\d+\s+años", f):
        datos["cantidad"] = a_numero(m.group(1))
        f = m.group(2)
    elif re.match(r"(?:varios|varias|hombres|mujeres|personas|jóvenes|menores|policías|niños|"
                  r"pasajeros|trabajadores|choferes|integrantes|hermanos|amigos|extranjeros)\b", f):
        datos["es_grupo"] = True
    if f.startswith("pareja") or f.startswith("esposos"):
        datos["cantidad"] = 2
    if re.match(r"familia\b", f):
        datos["es_grupo"] = True

    # objetivo no humano (vehículo, local, vivienda...)
    cabeza = re.sub(r"^(?:varios|varias|\d+)\s+", "", f).split(" ")[0] if f else ""
    if re.fullmatch(r"[A-ZÁÉÍÓÚÑ]{2,}", original.split(" ")[0].strip('"')):  # siglas: "BCP", "ODPE"
        datos.update(tipo_actor="Empresa / institución (sigla)", rubro="Sin dato")
    singular = re.sub(r"(?:es|s)$", "", cabeza)
    for patron, tipo, rubro in OBJETIVOS:
        if datos["tipo_actor"] == "Persona" and (re.fullmatch(patron, cabeza) or re.fullmatch(patron, singular)):
            datos.update(tipo_actor=tipo, rubro=rubro)
            break
    if datos["tipo_actor"] != "Persona":
        dueno = re.search(r"\bde\s+(?:la\s+|el\s+|un\s+|una\s+)?(\w+)", f)
        if dueno:  # "vivienda de cantante X", "casa de policía"
            for patron, ocup, rubro in OCUPACIONES:
                if re.fullmatch(patron, dueno.group(1)):
                    datos["ocupacion"] = f"Propiedad de: {ocup}"
                    datos["rubro"] = rubro
                    break
        if datos["tipo_actor"] == "Vehículo" or re.search(r"línea|ruta|empresa", f):
            datos["rubro"] = "Transporte" if datos["tipo_actor"] == "Vehículo" else datos["rubro"]
    datos["organizacion_vinculada"] = organizacion(original)

    if datos["tipo_actor"] == "Persona":
        palabras = re.findall(r"[\wáéíóúñ]+", f)
        for p in palabras:
            if p in SEXO:
                datos["sexo"] = SEXO[p]
                break
            if p in OFICIO_FEMENINO:
                datos["sexo"] = "Femenino"
                break
            if p in OFICIO_MASCULINO:
                datos["sexo"] = "Masculino"
                break
            if re.fullmatch(r"(?:venezolan|peruan|colombian|ecuatorian|chilen|bolivian|extranjer)a|paraguaya|trans", p):
                datos["sexo"] = "Femenino" if p != "trans" else "Mujer trans"
                break
            if re.fullmatch(r"(?:venezolan|peruan|colombian|ecuatorian|chilen|bolivian|extranjer)o|paraguayo", p):
                datos["sexo"] = "Masculino"
                break

        for raiz, nac in NACIONALIDADES.items():
            if re.search(rf"\b{raiz}", f):
                datos["nacionalidad"] = nac
                break

        edad = re.search(r"(\d{1,2})\s+años", f) or re.search(r"menor(?:es)? de (\d{1,2})\b", f)
        if edad:
            datos["edad"] = int(edad.group(1))
        datos["grupo_etario"] = grupo_etario(f, datos["edad"])

        for p in palabras:
            for patron, ocup, rubro in OCUPACIONES:
                if re.fullmatch(patron, p):
                    datos["ocupacion"], datos["rubro"] = ocup, rubro
                    break
            if datos["ocupacion"] != "Sin dato":
                break
        if datos["ocupacion"] == "Pasajero(a)" or datos["ocupacion"].startswith("Chofer") or \
                datos["ocupacion"] in ("Cobrador(a)", "Copiloto"):
            pass
        if re.search(r"\bdue[ñn][oa] de (\w+)", f) and datos["ocupacion"].startswith("Dueño"):
            datos["ocupacion"] = "Dueño(a) de " + re.search(r"\bdue[ñn][oa] de (\w+)", f).group(1)

        datos["nombre_o_alias"] = nombre_propio(original)
        if re.search(r"\bex\s+(?:miembro|reo|recluso|presidiario)|recluso|estafador|alias", f):
            datos["antecedentes"] = "Posible (según titular)"
    return datos


def grupo_etario(f, edad):
    if edad is not None:
        if edad < 12:
            return "Niño(a) (<12)"
        if edad < 18:
            return "Adolescente (12-17)"
        if edad < 30:
            return "Joven (18-29)"
        if edad < 60:
            return "Adulto (30-59)"
        return "Adulto mayor (60+)"
    if re.search(r"\bniñ[oa]s?\b|\bbebé", f):
        return "Niño(a) (<12)"
    if re.search(r"\bmenor(?:es)?\b|adolescentes?|escolar(?:es)?", f):
        return "Menor de edad"
    if re.search(r"\bj[oó]ven(?:es)?\b|universitari", f):
        return "Joven"
    if re.search(r"adulto mayor|anciano|abuel", f):
        return "Adulto mayor (60+)"
    return "Sin dato"


def organizacion(frase):
    m = re.search(r"\"([^\"]+)\"", frase)
    if m and not re.search(r"alias\s+\"", frase, re.I):
        return m.group(1)
    m = re.search(r"empresa\s+(?:de transportes?\s+)?([A-ZÁÉÍÓÚÑ][\w\-\.]*(?:\s+[A-ZÁÉÍÓÚÑ][\w\-\.]*)*)", frase)
    if m:
        return m.group(1)
    m = re.search(r"\b(l[ií]nea|ruta)\s+([\w\-]+)", frase, re.I)
    if m:
        return f"{m.group(1).capitalize()} {m.group(2)}"
    m = re.search(r"\bde\s+((?:[A-ZÁÉÍÓÚ][\w\-]+)(?:[\s\-]+[A-ZÁÉÍÓÚ][\w\-]+)*)$", frase.strip())
    if m and re.search(r"bus|combi|local|empresa|terminal|agencia", frase, re.I):
        return m.group(1)
    return ""


def nombre_propio(frase):
    m = re.search(r"alias\s+\"?([^\"]+?)\"?(?:\s|$)", frase)
    if m:
        return f'alias "{m.group(1)}"'
    m = re.search(r"\b(?:cantante|periodista|alcalde|dirigente|pastor|futbolista|empresario|"
                  r"abogad[oa]|animador|productor|candidat[oa]|regidor)\s+"
                  r"([A-ZÁÉÍÓÚ][a-záéíóúñ]+(?:\s+[A-ZÁÉÍÓÚ][a-záéíóúñ]+)+)", frase)
    return m.group(1) if m else ""


def separar_personas(frase):
    """'hombre y mujer' -> ['hombre', 'mujer']; respeta 'padre de familia'."""
    trozos = re.split(r"\s+(?:y|e)\s+(?=(?:a\s+)?(?:su\s+|sus\s+|la\s+|el\s+|un\s+|una\s+|otr[oa]s?\s+|\d+\s+)?"
                      r"[a-záéíóúñ]+)", frase)
    return [t for t in trozos if t.strip()]


ABREVIATURAS = (r"\b(Av|Avs|Jr|Urb|Prol|Psje|Pje|Mz|Lt|Sra|Sr|Dr|Cdra|Asoc|Coop|Ctra|Car|Sto|Sta|Ampl|"
                r"Mcal|Gral|Nro|Esq|Dpto|Ref|A\.H|AA\.HH|P\.J|I\.E|C\.P|U\.V|S\.A|S\.A\.C|PP\.JJ|J\.C)\.")


def proteger(texto):
    return re.sub(ABREVIATURAS, lambda m: m.group(1).replace(".", "§") + "§", texto)


def separar_oraciones(titular):
    """Divide en oraciones sin cortar en abreviaturas como 'Av.' o 'A.H.'."""
    protegido = proteger(titular)
    oraciones = re.split(r"\.\s+(?=[A-ZÁÉÍÓÚÑ0-9¿¡])", protegido)
    return [o.replace("§", ".").strip() for o in oraciones if o.strip()] or [titular]


def extraer_victimas(titular):
    victimas = []
    oraciones = separar_oraciones(titular)
    titular = proteger(titular)  # "Av." -> "Av§" para que la frase no se corte en el punto
    # "agreden y hieren a X" -> "hieren a X" (dos verbos seguidos con un solo objeto)
    titular = re.sub(r"\b[a-záéíóúñ]+(?:an|en)\s+y\s+([a-záéíóúñ]+(?:an|en)\s+(?:a|al)\b)", r"\1", titular)
    principal = oraciones[0]
    resto = " ".join(oraciones[1:]).lower()

    def agregar(frase, estado, origen):
        frase = frase.split(",")[-1].strip().replace("§", ".")
        if origen == "explosivo en" and re.match(r"(?:av|avenida|calle|jr|jirón|pasaje|psje|sector|zona|parque|cruce|urb|"
                                                 r"cuadra|esquina|inmediaciones|exteriores|su poder|plena)\b", frase, re.I):
            return  # es un lugar, no un objetivo
        for trozo in separar_personas(frase):
            if re.match(r"(?:con|y|se|paralizan|huyen|escapan|roban|dejan|exigen|cobran)\b", trozo.strip(), re.I):
                continue
            d = describir_frase(trozo)
            if origen == "explosivo en" and d["tipo_actor"] == "Persona" and \
                    not re.search(r"candidat|alcalde|dirigente|policía|empresari|comerciante|familia", trozo, re.I):
                d["tipo_actor"] = "Objetivo sin clasificar"
            d["estado_o_situacion"] = estado if d["tipo_actor"] == "Persona" else "Atacado (daños materiales)"
            d["frase_detectada"] = trozo.strip()
            d["regla"] = origen
            victimas.append(d)

    patrones = [
        (VERBOS_FATALES, "Fallecido"),
        (VERBOS_HERIDOS, "Herido"),
        (VERBOS_SECUESTRO, "Secuestrado"),
        (VERBOS_AMENAZA, "Amenazado"),
    ]
    for verbos, estado in patrones:
        for m in re.finditer(rf"\b(?:{verbos})\s+{MODIFICADORES}(?:a\s+|al\s+)?(.+?){FIN_FRASE}",
                             titular, re.I):
            frase = m.group(1)
            if re.match(r"(?:la|una)?\s*(?:balacera|explosivo|carga)", frase, re.I):
                continue
            if frase.lower().startswith("l ") or not frase.strip():
                continue
            agregar(frase, estado, "verbo + objeto")

    # "dejan grave a su pareja", "dejan herido a otro"
    for m in re.finditer(rf"\bdej(?:an|a|ando)\s+(?:gravemente\s+)?(?:grave|graves|herid[oa]s?|mal herid[oa])\s+a\s+(.+?){FIN_FRASE}",
                         titular, re.I):
        agregar(m.group(1), "Herido grave", "dejan grave a")

    # pasiva: "Chofer de colectivo fue asesinado", "Mujer fue herida"
    for oracion in oraciones:
        m = re.match(r"^(.+?)\s+(?:fue|fueron|resultó|resultaron|quedó|quedaron|queda|quedan|muere|murió|es|son)\s+"
                     r"(asesinad[oa]s?|herid[oa]s?|balead[oa]s?|acuchillad[oa]s?|secuestrad[oa]s?|"
                     r"atacad[oa]s?|hallad[oa]s?|encontrad[oa]s?)", oracion, re.I)
        if not m or re.match(r"(?:balacera|ataque|enfrentamiento|explosivo|atentado|víctima|sujeto|sicario)", m.group(1), re.I):
            continue
        if len(m.group(1).split()) > 7:
            continue
        participio = (m.group(2) or "").lower()
        muerto = re.search(r"muert|sin vida|cadáver", oracion, re.I) and participio.startswith(("hallad", "encontrad"))
        estado = "Fallecido" if participio.startswith("asesin") or "mur" in m.group(0).lower() or muerto else \
            "Secuestrado" if participio.startswith("secuestr") else "Herido"
        agregar(m.group(1), estado, "voz pasiva")

    # "Balacera deja una menor y 2 adultos heridos", "2 adultos y 1 joven heridos tras balacera"
    for m in re.finditer(r"(?:\bdej\w*\s+|^)((?:(?:\d+|un|una|dos|tres)\s+)?[\wáéíóúñ]+(?:\s+(?:y|e)\s+(?:\d+|un|una|dos|tres)?\s*[\wáéíóúñ]+)+)\s+"
                         r"(heridos|herid[oa]s|fallecidos|muertos)\b", titular, re.I):
        if re.search(r"muert|fallec|herid|person", m.group(1), re.I):
            continue
        estado = "Fallecido" if re.match(r"fallec|muert", m.group(2), re.I) else "Herido"
        agregar(m.group(1), estado, "varias personas + estado")

    # "Hallan cuerpo de mujer", "Dejan cuerpo de joven"
    for m in re.finditer(rf"\b(?:cuerpos?|cadáver(?:es)?|restos)\s+(?:sin vida\s+)?(?:de\s+)?(?:un\s+|una\s+)?(.+?){FIN_FRASE}",
                         titular, re.I):
        agregar(m.group(1), "Fallecido", "hallazgo de cuerpo")

    # objetivos materiales: "explosivo en vivienda de...", "balean bus de la línea 129"
    for m in re.finditer(rf"\b(?:explosivos?|carga explosiva|artefacto explosivo|dinamita|granada)(?:\s+(?:y|con)\s+[\wáéíóúñ ]+?)?\s+"
                         rf"(?:fue detonado\s+|fueron detonados\s+)?(?:en|contra|a|al|debajo de|frente a|cerca a)\s+(?:la\s+|el\s+|una\s+|un\s+)?(?:puerta\s+de\s+)?(.+?){FIN_FRASE}",
                         titular, re.I):
        agregar(m.group(1), "Atacado (daños materiales)", "explosivo en")

    # "dejando una mujer fallecida", "deja un hombre herido"
    for m in re.finditer(r"\bdej(?:an|a|ando|ó)\s+((?:\d+|un|una|dos|tres)\s+(?!personas?\b|muert|fallec|herid)[\wáéíóúñ]+)\s+"
                         r"(fallecid[oa]s?|muert[oa]s?|herid[oa]s?|grave)", titular, re.I):
        estado = "Fallecido" if re.match(r"fallec|muert", m.group(2), re.I) else "Herido"
        agregar(m.group(1), estado, "deja + persona + estado")

    # conteos: "Balacera deja 1 muerto y 6 heridos", "y hieren a otros 2" ya cubierto arriba
    ya_hay_personas = any(v["tipo_actor"] == "Persona" for v in victimas)
    for m in re.finditer(r"\b(\d+|un|una|uno|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez|varios|varias)\s+"
                         r"(?:otros?\s+|personas?\s+)?(muertos?|fallecidos?|heridos?|heridas?|lesionados?)\b",
                         titular, re.I):
        n = a_numero(m.group(1))
        estado = "Fallecido" if re.match(r"muert|fallec", m.group(2), re.I) else "Herido"
        # si ya hay una víctima fallecida y el conteo dice "1 muerto", no duplicar
        repetidos = sum(1 for v in victimas if v["estado_o_situacion"].startswith(estado) and v["tipo_actor"] == "Persona")
        if ya_hay_personas and n is not None and repetidos >= n:
            continue
        d = describir_frase("personas" if n is None else "persona")
        d.update(cantidad=(n - repetidos) if n else 1, es_grupo=n is None,
                 estado_o_situacion=estado, frase_detectada=m.group(0), regla="conteo de muertos/heridos")
        victimas.append(d)

    # ajustar estado con oraciones finales
    for v in victimas:
        if v["tipo_actor"] != "Persona":
            continue
        est = v["estado_o_situacion"]
        if est.startswith("Herido"):
            if re.search(r"murió|falleci|muri[oó]", resto):
                v["estado_o_situacion"] = "Fallecido"
            elif re.search(r"grave|crítico|uci", resto) or re.search(r"\bgrave", titular.lower()):
                v["estado_o_situacion"] = "Herido grave"
            elif re.search(r"sobrevivi|fuera de peligro|estable", resto):
                v["estado_o_situacion"] = "Herido (sobrevivió)"
            elif re.search(r"desconoce su estado", resto):
                v["estado_o_situacion"] = "Herido (estado desconocido)"
            elif re.search(r"sin heridos|ilesos?|no hubo heridos", resto):
                v["estado_o_situacion"] = "Atacado (ileso)"
        if est == "Secuestrado" and re.search(r"liber|rescat", titular.lower()):
            v["estado_o_situacion"] = "Secuestrado (liberado)"
    personas = [v for v in victimas if v["tipo_actor"] == "Persona"]
    if personas and re.search(r"tenía antecedentes|con antecedentes|requisitoriad", titular.lower()):
        personas[0]["antecedentes"] = "Sí (según titular)"
    return victimas


# ---------------------------------------------------------------------------
# Victimarios
# ---------------------------------------------------------------------------

def extraer_victimarios(titular):
    t = titular.strip()
    tl = t.lower()
    filas = []

    capturado = re.search(r"captur|detien|detenid|interviene|intervenid|arrest|atrap|reducid", tl)
    abatido = re.search(r"(?:mata|abate|abatido|matan) a uno de ellos|es baleado y|sicario muerto|delincuente muerto", tl)
    situacion = "Abatido / muerto" if abatido else "Capturado" if capturado else "No capturado / sin dato"

    # quien captura es la policía; el victimario es el capturado
    m = re.match(r"^(?:polic[ií]a|pnp|serenazgo|ronderos|vecinos)\s+(?:de civil\s+)?(?:captura|capturan|detiene|detienen|interviene|intervienen)\s+a\s+(.+?)"
                 + FIN_FRASE, tl)
    if m:
        frase = m.group(1)
        tipo = tipo_victimario(frase) or "Sujeto(s) no identificado(s)"
        d = describir_frase(frase)
        filas.append(base_victimario(tipo, frase, d, "Capturado", "captura policial"))
        return filas

    # "Sujeto es detenido con explosivo", "Sujetos fueron capturados"
    m = re.match(r"^((?:[\wáéíóúñ]+\s+){0,2}?)(?:es|fue|fueron|son)\s+(?:detenid|capturad|intervenid|abatid)", tl)
    if m and tipo_victimario(m.group(1)):
        d = describir_frase(m.group(1))
        return [base_victimario(tipo_victimario(m.group(1)), m.group(1).strip(), d,
                                "Abatido / muerto" if "abatid" in tl else "Capturado", "sujeto capturado")]

    m = re.match(r"^((?:[\wáéíóúñ]+\s+){0,3}?)(?:en moto|en auto|armados|encapuchados|a bordo de \w+|vestidos como \w+)?\s*"
                 r"(?:asesin|mat|bale|hier|disp|deton|dej|lanz|atac|incendi|incenci|amenaz|secuestr|acuchill|apuñal|agred|"
                 r"intercept|persigu|ingres|irrump|rob|asalt|desat|arroj|quem|golpe|intent|exig|prend|"
                 r"se enfrent|abr|realiz|acribill|ejecut|extors|cobr|balacer)", tl)
    sujeto = m.group(1).strip() if m else ""
    tipo = tipo_victimario(sujeto) if sujeto else None

    if tipo is None:
        # frases como "Balacera entre grupos", "Enfrentamiento entre bandas"
        if re.search(r"entre grupos|entre bandas|enfrentamiento|pelea|disputa", tl):
            tipo, sujeto = "Grupos enfrentados", "grupos"
        else:
            tipo, sujeto = "No especificado", ""
    d = describir_frase(sujeto) if sujeto else describir_frase("persona")
    if re.match(r"^(?:\d+\s+)?[\wáéíóúñ]+s\b", sujeto) and not re.match(r"^\d", sujeto):
        d["es_grupo"] = True  # "Sicarios", "Sujetos": se sabe que son 2 o más
    if tipo == "No especificado":
        d["es_grupo"] = False
        d["sexo"] = "Sin dato"

    cuenta = re.search(r"(\d+|dos|tres|cuatro|cinco)\s+(?:sospechosos|detenidos|capturados|sujetos)", tl) or \
        re.search(r"detiene a (\d+)", tl)
    if cuenta:
        d["cantidad"], d["es_grupo"] = a_numero(cuenta.group(1)), False
    filas.append(base_victimario(tipo, sujeto, d, situacion, "sujeto del titular"))
    return filas


def tipo_victimario(frase):
    frase = frase.strip().lower()
    for patron, tipo in VICTIMARIOS:
        if re.match(rf"^(?:\d+\s+)?(?:{patron})\b", frase):
            return tipo
    return None


def base_victimario(tipo, frase, d, situacion, regla):
    d = dict(d)
    d["tipo_actor"] = "Persona"
    d["ocupacion"] = tipo
    d["rubro"] = "Sin dato"
    d["organizacion_vinculada"] = ""
    d["estado_o_situacion"] = situacion
    d["frase_detectada"] = frase
    d["regla"] = regla
    return d


# ---------------------------------------------------------------------------
# Tabla final
# ---------------------------------------------------------------------------

COLUMNAS = ["id_actor", "id_incidente", "fecha", "tipo_evento", "rol", "tipo_actor",
            "cantidad", "es_grupo", "nombre_o_alias", "sexo", "edad", "grupo_etario",
            "nacionalidad", "ocupacion", "rubro", "organizacion_vinculada", "antecedentes",
            "estado_o_situacion", "medio_empleado", "movil_presunto", "ambito",
            "latitud", "longitud", "titular", "frase_detectada", "regla", "fuente_url"]


def construir_actores(incidentes, una_fila_por_persona=True):
    filas = []
    for inc in incidentes.itertuples(index=False):
        actores = [("Víctima", v) for v in extraer_victimas(inc.titular)]
        if not actores:
            d = describir_frase("persona")
            sin_heridos = re.search(r"sin heridos|no hubo heridos|sin reporte de heridos|ilesa|ileso|se salvó", inc.titular.lower())
            d.update(tipo_actor="Sin víctima identificada", estado_o_situacion="Sin heridos reportados" if sin_heridos else "Sin dato",
                     frase_detectada="", regla="sin víctima detectada")
            actores.append(("Víctima", d))
        actores += [("Victimario", v) for v in extraer_victimarios(inc.titular)]

        for rol, a in actores:
            repetir = 1
            if una_fila_por_persona and a["tipo_actor"] == "Persona" and not a["es_grupo"] \
                    and isinstance(a["cantidad"], int) and 1 < a["cantidad"] <= 20:
                repetir = a["cantidad"]
            for _ in range(repetir):
                fila = dict(a)
                fila["cantidad"] = "2 o más (sin número)" if a["es_grupo"] else (1 if repetir > 1 else a["cantidad"])
                fila.update(id_incidente=inc.id_incidente, fecha=inc.fecha, tipo_evento=inc.tipo_evento,
                            rol=rol, medio_empleado=inc.medio_empleado, movil_presunto=inc.movil_presunto,
                            ambito=inc.ambito, latitud=inc.latitud, longitud=inc.longitud,
                            titular=inc.titular, fuente_url=inc.fuente_url)
                filas.append(fila)

    actores = pd.DataFrame(filas)
    actores["id_actor"] = [f"A{i + 1:05d}" for i in range(len(actores))]
    actores["edad"] = actores["edad"].astype("Int64")
    return actores[COLUMNAS]


def resumen_por_incidente(incidentes, actores):
    v = actores[(actores["rol"] == "Víctima") & (actores["tipo_actor"] == "Persona")]
    num = pd.to_numeric(v["cantidad"], errors="coerce").fillna(1)
    agg = pd.DataFrame({
        "victimas_persona": num.groupby(v["id_incidente"]).sum(),
        "fallecidos": num[v["estado_o_situacion"] == "Fallecido"].groupby(v["id_incidente"]).sum(),
        "heridos": num[v["estado_o_situacion"].str.startswith("Herido")].groupby(v["id_incidente"]).sum(),
    })
    out = incidentes.merge(agg, left_on="id_incidente", right_index=True, how="left")
    for c in ["victimas_persona", "fallecidos", "heridos"]:
        out[c] = out[c].fillna(0).astype(int)
    return out


def main(ruta_csv=CSV_POR_DEFECTO, carpeta=SALIDA_POR_DEFECTO):
    carpeta = Path(carpeta)
    carpeta.mkdir(parents=True, exist_ok=True)
    incidentes = cargar_incidentes(ruta_csv)
    actores = construir_actores(incidentes)
    incidentes = resumen_por_incidente(incidentes, actores)

    incidentes.to_csv(carpeta / "incidentes.csv", index=False, encoding="utf-8-sig")
    actores.to_csv(carpeta / "actores.csv", index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(carpeta / "actores.xlsx") as xl:
        actores.to_excel(xl, sheet_name="actores", index=False)
        incidentes.to_excel(xl, sheet_name="incidentes", index=False)

    print(f"Incidentes: {len(incidentes)}")
    print(f"Filas de actores: {len(actores)}")
    print(actores.groupby(["rol", "tipo_actor"]).size().to_string())
    return incidentes, actores


if __name__ == "__main__":
    main(*sys.argv[1:3])
