#!/usr/bin/env python3
"""Motor de producción: baja los datos EN VIVO de Power BI Service (vía Service
Principal, sin depender de que nadie tenga Power BI Desktop abierto) y genera los
2 HTML finales. Pensado para correr en GitHub Actions todos los lunes, sin
intervención manual.
 
Variables de entorno requeridas (se cargan como GitHub Actions secrets, ver
.github/workflows/generar_reportes.yml):
    AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET,
    POWERBI_WORKSPACE_ID, POWERBI_DATASET_ID
 
Cada una de las consultas DAX de este archivo fue probada en vivo contra el
Power BI real de la empresa antes de escribirse acá (ver
claude/diseno-automatizacion-html-taller.md, v16, en el proyecto de Claude).
"""
import datetime
import os
import sys
from pathlib import Path
 
import pandas as pd
from jinja2 import Environment, FileSystemLoader
 
BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))
 
from motor_agregacion import (
    preparar_horas_df, filtrar_semana, construir_mec_taller, construir_mec_campo,
    construir_rubros_taller, construir_ausentes_nuevos, construir_estados_lf,
    construir_horas_chinagro, construir_maquinas_top, construir_top_supervisores,
    construir_fluidos, construir_comp_mirar, construir_quincenas, clasificar_fluido,
    calcular_kpis_taller, generar_nota_pico_horas, generar_nota_lider_supervisores,
    obtener_access_token_service_principal, cargar_desde_power_bi_service,
    parsear_fecha_pbi_service,
)
from formato import fnum, datos_semana
 
OUTPUT_DIR = BASE.parent / "docs"
TEMPLATES_DIR = BASE / "templates"
 
 
# ---------------------------------------------------------------------------
# Fechas: calcula automáticamente la semana a reportar (lunes a domingo, la
# última semana ya completa antes de "hoy"). Pensado para correr un lunes.
# ---------------------------------------------------------------------------
 
def calcular_fechas(hoy=None):
    hoy = hoy or datetime.date.today()
    domingo_actual = hoy - datetime.timedelta(days=hoy.isoweekday())
    lunes_actual = domingo_actual - datetime.timedelta(days=6)
    domingo_anterior = lunes_actual - datetime.timedelta(days=1)
    lunes_anterior = domingo_anterior - datetime.timedelta(days=6)
    # 1 semana extra de margen antes de la anterior, igual que en las pruebas.
    desde_horas = lunes_anterior - datetime.timedelta(days=7)
    return {
        "hoy": hoy,
        "lunes_actual": lunes_actual, "domingo_actual": domingo_actual,
        "lunes_anterior": lunes_anterior, "domingo_anterior": domingo_anterior,
        "desde_horas": desde_horas,
        "desde_180d": hoy - datetime.timedelta(days=180),
        "desde_anio": datetime.date(hoy.year, 1, 1),
    }
 
 
def _d(fecha):
    return f"DATE({fecha.year},{fecha.month},{fecha.day})"
 
 
# ---------------------------------------------------------------------------
# Consultas DAX (probadas en vivo antes de quedar acá)
# ---------------------------------------------------------------------------
 
def dax_horas(desde, hasta):
    return f"""
EVALUATE
SELECTCOLUMNS(
    FILTER('Horas_Personal_AppSheet',
        'Horas_Personal_AppSheet'[Fecha] >= {_d(desde)} &&
        'Horas_Personal_AppSheet'[Fecha] <= {_d(hasta)}
    ),
    "Fecha", 'Horas_Personal_AppSheet'[Fecha],
    "Cuil", 'Horas_Personal_AppSheet'[Cuil],
    "Mecanico", 'Horas_Personal_AppSheet'[Mecánico],
    "Lugar", 'Horas_Personal_AppSheet'[Lugar],
    "Rubro", 'Horas_Personal_AppSheet'[Rubro],
    "SubRubro", 'Horas_Personal_AppSheet'[Sub Rubro],
    "Tipo", 'Horas_Personal_AppSheet'[Tipo],
    "Maquina", 'Horas_Personal_AppSheet'[Máquina],
    "Horas", 'Horas_Personal_AppSheet'[Horas],
    "HorasPrep", 'Horas_Personal_AppSheet'[Horas_Preparacion],
    "HorasTraslado", 'Horas_Personal_AppSheet'[Horas_Traslado],
    "Supervisor", 'Horas_Personal_AppSheet'[Supervisor],
    "Estado", 'Horas_Personal_AppSheet'[Estado],
    "Email", 'Horas_Personal_AppSheet'[Email_Usuario],
    "legajoemp", 'Horas_Personal_AppSheet'[legajoemp]
)
"""
 
 
def dax_liquidacion_mo(desde, hasta):
    # Nota: se envuelve el SUMMARIZECOLUMNS en un SELECTCOLUMNS para forzar nombres
    # de columna predecibles -- una columna de agrupacion SIN alias (como
    # 'Liquidacion_MO'[Cuil] acá) vuelve de la API con su nombre calificado
    # completo (ej. "Liquidacion_MO[Cuil]"), no simplemente "Cuil". Se probó en
    # vivo con la consulta de Propiedad y confirmó el problema antes de aplicar
    # este arreglo a todas las consultas que agrupan sin alias.
    return f"""
EVALUATE
SELECTCOLUMNS(
    SUMMARIZECOLUMNS(
        'Liquidacion_MO'[Cuil],
        "hsjornal", CALCULATE(SUM('Liquidacion_MO'[hsjornal]),
            'Liquidacion_MO'[fecha_tarea] >= {_d(desde)},
            'Liquidacion_MO'[fecha_tarea] <= {_d(hasta)})
    ),
    "Cuil", 'Liquidacion_MO'[Cuil],
    "hsjornal", [hsjornal]
)
"""
 
 
def dax_fluidos(desde_anio):
    return f"""
EVALUATE
SELECTCOLUMNS(
    FILTER('ConsumosyReparaciones',
        'ConsumosyReparaciones'[rubro] = "LUBRICANTE" &&
        'ConsumosyReparaciones'[fechfaccargas] >= {_d(desde_anio)} &&
        NOT ISBLANK('ConsumosyReparaciones'[maquina])
    ),
    "maquina", 'ConsumosyReparaciones'[maquina],
    "tipo_maquina", 'ConsumosyReparaciones'[tipo_maquina],
    "repuesto", 'ConsumosyReparaciones'[repuesto],
    "fechfaccargas", 'ConsumosyReparaciones'[fechfaccargas],
    "id_insumo", 'ConsumosyReparaciones'[id_insumo]
)
"""
 
 
def dax_repuestos_todos(desde_180d):
    return f"""
EVALUATE
SELECTCOLUMNS(
    FILTER('ConsumosyReparaciones',
        'ConsumosyReparaciones'[fechfaccargas] >= {_d(desde_180d)} &&
        'ConsumosyReparaciones'[rubro] <> "ELECTRICOS" &&
        NOT ISBLANK('ConsumosyReparaciones'[maquina]) &&
        NOT ISBLANK('ConsumosyReparaciones'[repuesto])
    ),
    "maquina", 'ConsumosyReparaciones'[maquina],
    "tipo_maquina", 'ConsumosyReparaciones'[tipo_maquina],
    "rubro", 'ConsumosyReparaciones'[rubro],
    "repuesto", 'ConsumosyReparaciones'[repuesto],
    "fechfaccargas", 'ConsumosyReparaciones'[fechfaccargas]
)
"""
 
 
DAX_PRECIOS_REPUESTOS = """
EVALUATE
SELECTCOLUMNS(
    SUMMARIZECOLUMNS(
        'Compras'[repuesto],
        "precio_prom", AVERAGE('Compras'[precio_Unitario]),
        "precio_max", MAX('Compras'[precio_Unitario]),
        "n", COUNTROWS('Compras')
    ),
    "repuesto", 'Compras'[repuesto],
    "precio_prom", [precio_prom],
    "precio_max", [precio_max],
    "n", [n]
)
"""
 
 
def dax_horas_taller_dia(desde_anio):
    return f"""
EVALUATE
SELECTCOLUMNS(
    SUMMARIZECOLUMNS(
        'Horas_Personal_AppSheet'[Fecha],
        FILTER('Horas_Personal_AppSheet',
            'Horas_Personal_AppSheet'[Fecha] >= {_d(desde_anio)} &&
            LEFT('Horas_Personal_AppSheet'[Lugar],6) = "TALLER"
        ),
        "Horas", SUM('Horas_Personal_AppSheet'[Horas])
    ),
    "Fecha", 'Horas_Personal_AppSheet'[Fecha],
    "Horas", [Horas]
)
"""
 
 
def dax_calendario(desde_anio):
    return f"""
EVALUATE
SELECTCOLUMNS(
    FILTER('Calendario', 'Calendario'[Date] >= {_d(desde_anio)}),
    "Date", 'Calendario'[Date],
    "Quincena", 'Calendario'[Quincena],
    "IndiceQuincena", 'Calendario'[IndiceQuincena]
)
"""
 
 
def dax_propiedad(desde, hasta):
    return f"""
EVALUATE
SELECTCOLUMNS(
    SUMMARIZECOLUMNS(
        'Maestro_Maquinarias'[Propiedad],
        "Horas", CALCULATE(SUM('Horas_Personal_AppSheet'[Horas]),
            'Horas_Personal_AppSheet'[Fecha] >= {_d(desde)},
            'Horas_Personal_AppSheet'[Fecha] <= {_d(hasta)},
            LEFT('Horas_Personal_AppSheet'[Lugar],6) = "TALLER")
    ),
    "Propiedad", 'Maestro_Maquinarias'[Propiedad],
    "Horas", [Horas]
)
"""
 
 
# ---------------------------------------------------------------------------
# Orquestación
# ---------------------------------------------------------------------------
 
def construir_todo():
    tenant_id = os.environ["AZURE_TENANT_ID"]
    client_id = os.environ["AZURE_CLIENT_ID"]
    client_secret = os.environ["AZURE_CLIENT_SECRET"]
    workspace_id = os.environ["POWERBI_WORKSPACE_ID"]
    dataset_id = os.environ["POWERBI_DATASET_ID"]
 
    f = calcular_fechas()
    print(f"Semana a reportar: {f['lunes_actual']} a {f['domingo_actual']}")
 
    token = obtener_access_token_service_principal(tenant_id, client_id, client_secret)
 
    def pbi(dax_query):
        return cargar_desde_power_bi_service(workspace_id, dataset_id, dax_query, token)
 
    # --- Horas (Horas_Personal_AppSheet) ---
    df_horas_raw = pbi(dax_horas(f["desde_horas"], f["domingo_actual"]))
    df_horas_raw["FechaDT"] = parsear_fecha_pbi_service(df_horas_raw["Fecha"])
    df_horas = preparar_horas_df(df_horas_raw)
    df_actual = filtrar_semana(df_horas, f["lunes_actual"], f["domingo_actual"])
    df_anterior = filtrar_semana(df_horas, f["lunes_anterior"], f["domingo_anterior"])
 
    mec_taller = construir_mec_taller(df_actual, df_anterior)
    mec_campo = construir_mec_campo(df_actual)
    rubros_taller = construir_rubros_taller(df_actual)
    mec_ausentes, mec_nuevos = construir_ausentes_nuevos(df_actual, df_anterior)
    estados_lf = construir_estados_lf(df_actual)
    maquinas_top = construir_maquinas_top(df_actual)
    top_supervisores = construir_top_supervisores(df_actual)
 
    # --- Chinagro (Liquidacion_MO) ---
    df_liq = pbi(dax_liquidacion_mo(f["lunes_actual"], f["domingo_actual"]))
    horas_chinagro_data = construir_horas_chinagro(df_actual, df_liq)
 
    # --- Fluidos (ConsumosyReparaciones, rubro LUBRICANTE, año actual) ---
    df_fluidos_raw = pbi(dax_fluidos(f["desde_anio"]))
    df_fluidos_raw["fechfaccargas"] = parsear_fecha_pbi_service(df_fluidos_raw["fechfaccargas"])
    df_fluidos_raw = df_fluidos_raw.dropna(subset=["fechfaccargas"])
    df_fluidos_raw["tipofluido"] = df_fluidos_raw["repuesto"].apply(clasificar_fluido)
    fluidos = construir_fluidos(df_fluidos_raw[df_fluidos_raw["tipofluido"].notna()], anio=f["hoy"].year)
 
    # --- compMirar (ConsumosyReparaciones, todos los rubros salvo eléctricos) ---
    df_todos_raw = pbi(dax_repuestos_todos(f["desde_180d"]))
    df_todos_raw["fechfaccargas"] = parsear_fecha_pbi_service(df_todos_raw["fechfaccargas"])
    df_todos_raw = df_todos_raw.dropna(subset=["fechfaccargas"])
 
    df_precios = pbi(DAX_PRECIOS_REPUESTOS).dropna(subset=["repuesto"])
    df_precios["precio_prom"] = pd.to_numeric(df_precios["precio_prom"], errors="coerce")
    precios_repuesto = df_precios.drop_duplicates("repuesto").set_index("repuesto")["precio_prom"]
 
    comp_mirar = construir_comp_mirar(df_todos_raw, precios_repuesto=precios_repuesto)
 
    # --- Quincenas (Horas_Personal_AppSheet TALLER por día + Calendario) ---
    df_horas_dia = pbi(dax_horas_taller_dia(f["desde_anio"]))
    df_horas_dia["Fecha"] = parsear_fecha_pbi_service(df_horas_dia["Fecha"]).dt.strftime("%d/%m/%Y")
    df_calendario = pbi(dax_calendario(f["desde_anio"]))
    df_calendario["Date"] = parsear_fecha_pbi_service(df_calendario["Date"]).dt.strftime("%d/%m/%Y")
    df_calendario["Quincena"] = parsear_fecha_pbi_service(df_calendario["Quincena"]).dt.strftime("%d/%m/%Y")
    quincenas, faltantes_quincena = construir_quincenas(df_horas_dia, df_calendario)
    if faltantes_quincena:
        print(f"AVISO: {faltantes_quincena} días de horas de taller no matchearon con el calendario.")
 
    # --- Propia vs terceros ---
    df_prop = pbi(dax_propiedad(f["lunes_actual"], f["domingo_actual"]))
    propiedad_actual = {
        str(row["Propiedad"]).upper(): float(row["Horas"])
        for _, row in df_prop.iterrows() if pd.notna(row["Propiedad"])
    }
 
    kpi, kpi_app, talleres = calcular_kpis_taller(df_actual, df_anterior, propiedad_actual)
 
    sem = datos_semana(f["lunes_actual"], semana_numero=f["lunes_actual"].isocalendar()[1],
                        quincenas_desde=f["desde_anio"])
    notas_taller = {"pico_horas": generar_nota_pico_horas(quincenas, f["hoy"])}
    notas_maquinaria = {"lider_supervisores": generar_nota_lider_supervisores(top_supervisores)}
 
    ctx_taller = {
        "sem": sem, "kpi": kpi, "kpi_app": kpi_app, "talleres": talleres, "notas": notas_taller,
        "quincenas": quincenas, "rubros_taller": rubros_taller, "mec_taller": mec_taller,
        "mec_campo": mec_campo, "mec_ausentes": mec_ausentes, "mec_nuevos": mec_nuevos,
        "estados_lf": estados_lf, "horas_chinagro_data": horas_chinagro_data,
    }
    ctx_maquinaria = {
        "sem": sem, "notas": notas_maquinaria, "maquinas_top": maquinas_top,
        "comp_mirar": comp_mirar, "fluidos": fluidos, "top_supervisores": top_supervisores,
    }
    return ctx_taller, ctx_maquinaria
 
 
def render(nombre_template, contexto, salida):
    env = Environment(loader=FileSystemLoader(str(TEMPLATES_DIR)))
    env.filters["fnum"] = fnum
    env.filters["abs"] = abs
    tpl = env.get_template(nombre_template)
    html = tpl.render(**contexto)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUTPUT_DIR / salida).write_text(html, encoding="utf-8")
    print(f"Generado: {OUTPUT_DIR / salida} ({len(html)} bytes)")
 
 
if __name__ == "__main__":
    ctx_taller, ctx_maquinaria = construir_todo()
    render("gestion_taller.html.j2", ctx_taller, "Gestion_Taller.html")
    render("reunion_maquinaria.html.j2", ctx_maquinaria, "Reunion_Semanal_Maquinaria.html")
    # Página de entrada simple con los 2 links, para que GitHub Pages tenga un index.
    index_html = """<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">
<title>Reportes Taller - La Asturiana</title>
<style>body{font-family:sans-serif;max-width:600px;margin:60px auto;padding:0 20px}
a{display:block;padding:16px;margin:12px 0;background:#f5f5f5;border-radius:8px;
text-decoration:none;color:#222;font-weight:600}a:hover{background:#eee}</style></head>
<body><h1>Reportes del Taller - La Asturiana</h1>
<a href="Gestion_Taller.html">Gestion del Taller</a>
<a href="Reunion_Semanal_Maquinaria.html">Reunion Semanal de Maquinaria</a>
</body></html>"""
    (OUTPUT_DIR / "index.html").write_text(index_html, encoding="utf-8")
    print("Generado: index.html")
