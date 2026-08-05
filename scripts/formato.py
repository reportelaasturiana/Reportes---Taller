"""Utilidades de formato numérico y de fechas, estilo argentino, usadas por
las plantillas Jinja2 de los reportes del taller."""
import datetime


def fnum(x):
    """Formatea un número como lo hacen los HTML originales: coma decimal,
    sin decimales si es entero, sin ceros de más (6.9 -> '6,9', 196.0 -> '196')."""
    if x is None:
        return ""
    r = round(float(x), 2)
    if abs(r - round(r)) < 1e-9:
        return str(int(round(r)))
    s = f"{r:.2f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
MESES_LARGO = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
               "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


def _corto(fecha):
    return f"{fecha.day} {MESES[fecha.month - 1]}"


def datos_semana(lunes, semana_numero=None, quincenas_desde=None):
    """A partir del lunes de la semana en curso arma todos los textos de
    fecha usados en las 2 plantillas (título, header, rangos, etc.)."""
    domingo = lunes + datetime.timedelta(days=6)
    lunes_ant = lunes - datetime.timedelta(days=7)
    domingo_ant = lunes_ant + datetime.timedelta(days=6)
    numero = semana_numero if semana_numero is not None else lunes.isocalendar()[1]

    # "rango_corto" NO lleva año (así se usa en casi todo el HTML); solo el
    # <title> de la pestaña del navegador necesita el año, para eso está
    # "rango_titulo".
    rango_corto = f"{_corto(lunes)}–{_corto(domingo)}"
    rango_titulo = f"{rango_corto} {domingo.year}"
    rango_corto_guion = f"{_corto(lunes)} – {_corto(domingo)}"
    rango_largo = (
        f"lunes {lunes.day} de {MESES_LARGO[lunes.month - 1]} al domingo "
        f"{domingo.day} de {MESES_LARGO[domingo.month - 1]} {domingo.year}"
    )
    rango_corto_anterior = f"{lunes_ant.day}–{domingo_ant.day} {MESES[domingo_ant.month - 1]}" \
        if lunes_ant.month == domingo_ant.month else f"{_corto(lunes_ant)}–{_corto(domingo_ant)}"

    inicio_anio = quincenas_desde or datetime.date(lunes.year, 1, 1)
    mes_ini_abrev = MESES[inicio_anio.month - 1].capitalize()
    mes_fin_abrev = MESES[lunes.month - 1].capitalize()
    quincenas_rango_texto = f"{mes_ini_abrev}–{mes_fin_abrev} {lunes.year}"
    ytd_rango_texto = f"{MESES_LARGO[inicio_anio.month - 1]} {inicio_anio.year} a la fecha"

    return {
        "lunes": lunes,
        "domingo": domingo,
        "numero": numero,
        "rango_corto": rango_corto,
        "rango_titulo": rango_titulo,
        "rango_corto_guion": rango_corto_guion,
        "rango_largo": rango_largo,
        "rango_corto_anterior": rango_corto_anterior,
        "quincenas_rango_texto": quincenas_rango_texto,
        "ytd_rango_texto": ytd_rango_texto,
    }
