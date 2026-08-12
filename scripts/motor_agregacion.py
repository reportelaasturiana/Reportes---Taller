#!/usr/bin/env python3
"""
Motor de agregación real: toma extractos crudos de Power BI (Horas_Personal_AppSheet,
Liquidacion_MO, ConsumosyReparaciones) y arma los 12 bloques de datos + KPIs que
consumen las 2 plantillas Jinja2 (gestion_taller.html.j2, reunion_maquinaria.html.j2).
 
Reemplaza a validar_templates.py (que usaba datos hardcodeados de un solo HTML de
muestra) con lógica de agregación real, reutilizable semana a semana.
 
Los extractos crudos hoy se cargan desde CSV (sacados a mano con dax_query_operations
contra el Power BI Desktop del usuario, vía el bridge). En producción, estos mismos
CSV los va a producir el llamado a la REST API "Execute Queries" de Power BI Service
(ver PENDIENTE_SERVICE_PRINCIPAL.md) — la función cargar_desde_power_bi_service()
al final de este archivo es el punto de reemplazo.
"""
import datetime
import re
import pandas as pd
from pathlib import Path
 
BASE = Path(__file__).resolve().parent
 
# ---------------------------------------------------------------------------
# Constantes de negocio confirmadas con el usuario
# ---------------------------------------------------------------------------
 
DOTACION_FIJA = {
    "La Falda": 21,
    "Mte. Grande": 5,
    "Caspinch.": 6,
}
SIGLA = {"La Falda": "LA", "Mte. Grande": "MO", "Caspinch.": "CA"}
CSS = {"La Falda": "lf", "Mte. Grande": "mg", "Caspinch.": "ca"}
NOMBRE_LARGO = {"La Falda": "La Falda", "Mte. Grande": "Monte Grande", "Caspinch.": "Caspinchango"}
 
# Mapeo Lugar (texto crudo de AppSheet) -> nombre corto de taller usado en las plantillas
LUGAR_A_TALLER = {
    "LA FALDA": "La Falda",
    "MONTE GRANDE": "Mte. Grande",
    "CASPINCHANGO": "Caspinch.",
}
 
# Mapeo Email_Usuario -> (nombre, color). Fuente: BD_Horas 1.xlsx, hoja Maestro_Mecanicos
# (no existe una tabla equivalente en el modelo de Power BI: el único dato de
# identidad ahí es el email). Colores asignados en orden estable (paleta categórica).
PALETA_USUARIOS = ["#4C72B0", "#DD8452", "#55A868", "#C44E52", "#8172B3",
                   "#937860", "#DA8BC3", "#8C8C8C", "#CCB974", "#64B5CD", "#B0B0B0"]
EMAIL_A_USUARIO = {
    "martinamaya855@gmail.com": "AMAYA, MARTIN EXEQUIEL",
    "bulaciopaulo563@gmail.com": "BULACIO, PAULO ARIEL",
    "sergioadriancontino@gmail.com": "CONTINO, SERGIO ADRIAN",
    "matiasgarcia545mix@gmail.com": "GARCIA, VICTOR MATIAS",
    "grauluis933@gmail.com": "GRAU, LUIS MANUEL DEL VALLE",
    "marianoinfante17@gmail.com": "INFANTE, CARLOS MARIANO",
    "polacogn123@gmail.com": "NARBAJA, HECTOR GUILLERMO",
    "diegofalda9@gmail.com": "ORTIZ, DIEGO MARCELO",
    "walterreseo34@gmail.com": "RESEO, WALTER ELADIO",
    "jr5977277@gmail.com": "ROMERO, JOSE RAMON",
    "pedrotoloza47@gmail.com": "TOLOZA, JOSE PEDRO",
}
EMAIL_A_COLOR = {email: PALETA_USUARIOS[i % len(PALETA_USUARIOS)]
                  for i, email in enumerate(EMAIL_A_USUARIO)}
 
RUBRO_COLORES = {}
_PALETA_RUBROS = ["#c0392b", "#2e7d32", "#1565c0", "#e67e22", "#8e44ad", "#16a085",
                   "#d35400", "#2980b9", "#7f8c8d", "#c2185b", "#00838f", "#5d4037"]
 
 
def color_para_rubro(rubro):
    if rubro not in RUBRO_COLORES:
        RUBRO_COLORES[rubro] = _PALETA_RUBROS[len(RUBRO_COLORES) % len(_PALETA_RUBROS)]
    return RUBRO_COLORES[rubro]
 
 
def taller_desde_lugar(lugar):
    """'TALLER LA FALDA' -> ('La Falda', 'TALLER'); 'CAMPO MONTE GRANDE' -> ('Mte. Grande','CAMPO')"""
    if not isinstance(lugar, str):
        return None, None
    for prefijo, tipo in (("TALLER ", "TALLER"), ("CAMPO ", "CAMPO"), ("TRASLADOS ", "TRASLADOS")):
        if lugar.startswith(prefijo):
            resto = lugar[len(prefijo):].strip()
            return LUGAR_A_TALLER.get(resto, resto.title()), tipo
    return None, None
 
 
def r1(x):
    return round(float(x), 1) if pd.notna(x) else 0.0
 
 
# ---------------------------------------------------------------------------
# Carga de extractos crudos
# ---------------------------------------------------------------------------
 
def preparar_horas_df(df):
    """Agrega las columnas derivadas (Taller/TipoLugar) y normaliza tipos sobre un
    DataFrame de Horas_Personal_AppSheet ya cargado -- se separó de cargar_horas()
    para poder reusarlo tanto leyendo un CSV local (pruebas) como recibiendo un
    DataFrame ya bajado en vivo de Power BI Service (producción)."""
    df = df.copy()
    if "FechaDT" not in df.columns:
        df["FechaDT"] = pd.to_datetime(df["Fecha"], errors="coerce")
    df[["Taller", "TipoLugar"]] = df["Lugar"].apply(lambda x: pd.Series(taller_desde_lugar(x)))
    for col in ("Horas", "HorasPrep", "HorasTraslado"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    if "Cuil" in df.columns:
        df["Cuil"] = pd.to_numeric(df["Cuil"], errors="coerce").astype("Int64")
    if "legajoemp" in df.columns:
        df["legajoemp"] = pd.to_numeric(df["legajoemp"], errors="coerce").astype("Int64")
    return df
 
 
def cargar_horas(path_csv):
    df = pd.read_csv(path_csv, dtype={"Cuil": "Int64", "legajoemp": "Int64"})
    df["FechaDT"] = pd.to_datetime(df["Fecha"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
    return preparar_horas_df(df)
 
 
def filtrar_semana(df, lunes, domingo):
    return df[(df["FechaDT"].dt.date >= lunes) & (df["FechaDT"].dt.date <= domingo)].copy()
 
 
# ---------------------------------------------------------------------------
# Bloque 3: mecTaller
# ---------------------------------------------------------------------------
 
def construir_mec_taller(df_actual, df_anterior):
    taller_rows = df_actual[df_actual["TipoLugar"] == "TALLER"]
    hs_sa = (df_anterior[df_anterior["TipoLugar"] == "TALLER"]
             .groupby("Mecanico")["Horas"].sum())
    out = []
    for (mec, taller), g in taller_rows.groupby(["Mecanico", "Taller"]):
        if not taller:
            continue
        out.append({
            "n": mec,
            "t": taller,
            "col": color_para_rubro(taller),  # color por taller (reutiliza la paleta)
            "hs": r1(g["Horas"].sum()),
            "hsSA": r1(hs_sa.get(mec, 0)),
        })
    return out
 
 
# ---------------------------------------------------------------------------
# Bloque 4: mecCampo
# ---------------------------------------------------------------------------
 
def construir_mec_campo(df_actual):
    campo_rows = df_actual[df_actual["TipoLugar"] == "CAMPO"]
    out = []
    for mec, g in campo_rows.groupby("Mecanico"):
        out.append({
            "n": mec,
            "hs": r1(g["Horas"].sum()),
            "hs_t": r1(g["HorasTraslado"].sum()),
            "hs_p": r1(g["HorasPrep"].sum()),
        })
    return out
 
 
# ---------------------------------------------------------------------------
# Bloque 2: rubrosTaller
# ---------------------------------------------------------------------------
 
def construir_rubros_taller(df_actual):
    taller_rows = df_actual[(df_actual["TipoLugar"] == "TALLER") & df_actual["Rubro"].notna()]
    out = []
    for rubro, g in taller_rows.groupby("Rubro"):
        subs = [{"s": sub, "h": r1(gs["Horas"].sum())}
                for sub, gs in g.groupby("SubRubro")]
        out.append({"r": rubro, "col": color_para_rubro(rubro), "h": r1(g["Horas"].sum()), "subs": subs})
    return out
 
 
# ---------------------------------------------------------------------------
# Bloques 5/6: mecAusentes / mecNuevos
# ---------------------------------------------------------------------------
 
def construir_ausentes_nuevos(df_actual, df_anterior):
    def resumen(df):
        out = {}
        for mec, g in df.groupby("Mecanico"):
            email = g["Email"].dropna().iloc[0] if g["Email"].notna().any() else None
            out[mec] = {"hs": r1(g["Horas"].sum()), "usrNm": EMAIL_A_USUARIO.get(email, email or "—")}
        return out
 
    actuales = resumen(df_actual)
    anteriores = resumen(df_anterior)
    ausentes = [{"m": mec, "usrNm": d["usrNm"], "hsSA": d["hs"]}
                for mec, d in anteriores.items() if mec not in actuales]
    nuevos = [{"m": mec, "usrNm": d["usrNm"], "hs": d["hs"]}
              for mec, d in actuales.items() if mec not in anteriores]
    return ausentes, nuevos
 
 
# ---------------------------------------------------------------------------
# Bloque 7: estadosLF (solo La Falda)
# ---------------------------------------------------------------------------
 
def construir_estados_lf(df_actual):
    lf = df_actual[df_actual["Taller"] == "La Falda"]
    out = []
    for email, g in lf.groupby("Email"):
        if pd.isna(email):
            continue
        out.append({
            "nm": EMAIL_A_USUARIO.get(email, email),
            "ap": int((g["Estado"] == "Aprobado").sum()),
            "pe": int((g["Estado"] == "Pendiente").sum()),
            "re": int((g["Estado"] == "Rechazado").sum()),
        })
    return out
 
 
# ---------------------------------------------------------------------------
# Bloque 8: horasChinagroData
# ---------------------------------------------------------------------------
 
def construir_horas_chinagro(df_actual, df_liquidacion_mo):
    """df_liquidacion_mo: filas de Liquidacion_MO de la misma semana, con columnas
    Cuil, hsjornal. Cruce por Cuil (NO por legajo/legajoemp: 'legajoemp' está vacío
    en >99% de las filas recientes de Horas_Personal_AppSheet -- hallazgo confirmado
    en vivo; 'Cuil' en cambio está poblado de forma confiable en ambas tablas)."""
    chin_por_cuil = df_liquidacion_mo.groupby("Cuil")["hsjornal"].sum() if df_liquidacion_mo is not None else pd.Series(dtype=float)
    out = []
    for email, g in df_actual.groupby("Email"):
        if pd.isna(email):
            continue
        mecs = []
        for mec, gm in g.groupby("Mecanico"):
            cuil = gm["Cuil"].dropna().iloc[0] if gm["Cuil"].notna().any() else None
            chin = float(chin_por_cuil.get(cuil)) if cuil is not None and cuil in chin_por_cuil.index else None
            mecs.append({"m": mec, "decl": r1(gm["Horas"].sum()), "chin": (r1(chin) if chin is not None else None)})
        out.append({
            "nm": EMAIL_A_USUARIO.get(email, email),
            "email": email,
            "col": EMAIL_A_COLOR.get(email, "#888"),
            "mecs": mecs,
        })
    return out
 
 
# ---------------------------------------------------------------------------
# Bloque 9: maquinasTop
# ---------------------------------------------------------------------------
 
def construir_maquinas_top(df_actual, maestro_maquinarias=None):
    taller_rows = df_actual[df_actual["TipoLugar"] == "TALLER"]
    campo_rows = df_actual[df_actual["TipoLugar"] == "CAMPO"]
    campo_por_maq = campo_rows.groupby("Maquina").agg(
        t=("Horas", "sum"), tr=("HorasTraslado", "sum"), p=("HorasPrep", "sum")
    )
 
    grupos = {}
    for tipo, g in taller_rows.groupby("Tipo"):
        if pd.isna(tipo):
            continue
        maquinas = []
        for maq, gm in g.groupby("Maquina"):
            subs = [{"s": sub, "h": r1(gs["Horas"].sum())} for sub, gs in gm.groupby("SubRubro") if pd.notna(sub)]
            campo = None
            if maq in campo_por_maq.index:
                c = campo_por_maq.loc[maq]
                campo = {"t": r1(c["t"]), "tr": r1(c["tr"]), "p": r1(c["p"])}
            maquinas.append({
                "m": maq, "taller": gm["Taller"].dropna().iloc[0] if gm["Taller"].notna().any() else "",
                "h": r1(gm["Horas"].sum()), "subs": subs, "campo": campo,
            })
        grupos[tipo] = {"tipo": tipo, "col": color_para_rubro(tipo),
                         "hs": r1(sum(m["h"] for m in maquinas)), "maquinas": maquinas}
    return list(grupos.values())
 
 
# ---------------------------------------------------------------------------
# Bloques 10/11: compMirar y fluidos (ConsumosyReparaciones)
# ---------------------------------------------------------------------------
 
FLUIDO_KEYWORDS = {"HIDRAULICO": "Hidráulico", "REFRIGERANTE": "Refrigerante"}
 
 
def _pares_consecutivos(fechas_ordenadas):
    pares = []
    for f1, f2 in zip(fechas_ordenadas, fechas_ordenadas[1:]):
        pares.append({"f1": f1.strftime("%d/%m"), "f2": f2.strftime("%d/%m"), "gap": (f2 - f1).days})
    return pares
 
 
# Regla real de "Frecuencia de recambio de repuestos", ya escrita como texto fijo en
# la plantilla pero nunca aplicada del lado del motor (mismo patrón que se encontró en
# fluidos): se agrupa SOLO por estos 10 tipos de maquinaria, y cada insumo repetido se
# clasifica en 3 franjas según el mínimo de días entre cargas: <10 (falla/mal uso),
# 10-90 (trimestral), 90-180 (semestral). Antes el motor mandaba TODOS los tipos de
# maquinaria (119 grupos) y el filtro real solo vivía en el JS de la plantilla -- se
# veía bien igual, pero el HTML pesaba de más por los ~9 de cada 10 grupos que ni se
# llegaban a mostrar.
ORDEN_TIPOS_COMPMIRAR = ["TRACTORES", "CAMIONETAS", "PULVERIZADORAS", "PODADORAS", "CARGADORAS",
                         "DESMALEZADORAS", "CHANCHOS", "NODRIZAS", "ABONADORA(S)", "HERBICIDA(S)"]
 
_ALIAS_TIPO_COMPMIRAR = {
    "TRACTOR": "TRACTORES", "TRACTORES": "TRACTORES",
    "CAMIONETA": "CAMIONETAS", "CAMIONETAS": "CAMIONETAS",  # en Power BI viene en singular
    "PULVERIZADORA": "PULVERIZADORAS", "PULVERIZADORAS": "PULVERIZADORAS",
    "PODADORA": "PODADORAS", "PODADORAS": "PODADORAS",
    "CARGADORA": "CARGADORAS", "CARGADORAS": "CARGADORAS",
    "DESMALEZADORA": "DESMALEZADORAS", "DESMALEZADORAS": "DESMALEZADORAS",
    "CHANCHO": "CHANCHOS", "CHANCHOS": "CHANCHOS",
    "NODRIZA": "NODRIZAS", "NODRIZAS": "NODRIZAS",
    "ABONADORA": "ABONADORA(S)", "ABONADORAS": "ABONADORA(S)",
    "HERBICIDA": "HERBICIDA(S)", "HERBICIDAS": "HERBICIDA(S)",
}
 
# repuesto genéricos que en realidad son mano de obra/reparación cargada como si fuera
# un insumo (ej. todo el rubro ACOPLADOS se carga con repuesto="MANO DE OBRA", y el
# rubro REPARACION con repuesto="REPARACION") -- se repiten en casi todas las visitas de
# taller y no son una señal real de recambio de repuesto, así que se excluyen.
REPUESTOS_NO_SON_PARTES = {"MANO DE OBRA", "REPARACION"}
 
# rubro ELECTRICOS (baterías y otros ítems eléctricos) explícitamente excluido, tal
# como ya indicaba la nota fija de la plantilla ("todavía NO incluye baterías...").
RUBROS_EXCLUIDOS_COMPMIRAR = {"ELECTRICOS"}
 
# Corte por monto pedido por el usuario: un tornillo o una arandela que se repiten no
# sirven para una reunión semanal aunque cumplan el criterio de frecuencia -- lo que
# importa discutir es el repuesto caro que se repite. Precio unitario real sacado de
# Compras[precio_Unitario] (promedio por repuesto). Confirmado con el usuario: $30.000.
UMBRAL_MONTO_COMPMIRAR = 30000
 
 
def construir_comp_mirar(df_consumos, precios_repuesto=None, ventana_dias=180,
                          umbral_monto=UMBRAL_MONTO_COMPMIRAR):
    """df_consumos: filas de ConsumosyReparaciones (todos los rubros, no solo fluidos),
    con columnas maquina, tipo_maquina, rubro, repuesto, fecha_mov (fecha del movimiento
    = cuando se entrego el repuesto para esa maquina, movim.fecha en La Falda; NO la
    fecha de factura de compra, ver el comentario en dax_fluidos de motor_produccion.py).
    precios_repuesto: Series/dict {repuesto: precio_unitario}, de
    Compras[precio_Unitario] promediado por repuesto -- si no se pasa, no se filtra por
    monto (deja pasar todo, para no romper si todavía no se tiene el extracto de precios).
 
    Aplica las reglas reales ya documentadas en la plantilla y confirmadas por el
    usuario: solo 10 tipos de maquinaria, sin rubro eléctrico, sin líneas de mano de
    obra/reparación genérica, bucketeado en <10/10-90/90-180 días entre cargas, y
    ahora además filtrado a repuestos cuyo precio unitario supera `umbral_monto` (pedido
    explícito: un insumo barato que se repite no amerita discutirlo en la reunión).
    Devuelve una lista de secciones por tipo de maquinaria (mismo criterio de fluidos)."""
    hoy = df_consumos["fecha_mov"].max()
    desde = hoy - datetime.timedelta(days=ventana_dias)
    df = df_consumos[df_consumos["fecha_mov"] >= desde].copy()
 
    if "rubro" in df.columns:
        df = df[~df["rubro"].astype(str).str.upper().isin(RUBROS_EXCLUIDOS_COMPMIRAR)]
    df = df[~df["repuesto"].astype(str).str.upper().isin(REPUESTOS_NO_SON_PARTES)]
 
    df["tipo_norm"] = df["tipo_maquina"].astype(str).str.upper().str.strip().map(
        lambda t: _ALIAS_TIPO_COMPMIRAR.get(t, t))
    df = df[df["tipo_norm"].isin(ORDEN_TIPOS_COMPMIRAR)]
 
    secciones = {}
    for (maq, tipo_raw, tipo_norm), g in df.groupby(["maquina", "tipo_maquina", "tipo_norm"]):
        b10, btri, bsem = [], [], []
        for rep, gr in g.groupby("repuesto"):
            if precios_repuesto is not None:
                precio = precios_repuesto.get(rep)
                if precio is None or precio < umbral_monto:
                    continue
            else:
                precio = None
            # mismo motivo que en construir_fluidos: deduplicar por fecha, no por
            # línea de factura (una compra partida en varias líneas el mismo día no
            # es un recambio repetido real).
            fechas = sorted(set(gr["fecha_mov"].tolist()))
            if len(fechas) < 2:
                continue
            pares = _pares_consecutivos(fechas)
            mingap = min(p["gap"] for p in pares)
            item = {"rep": rep, "veces": len(fechas), "mingap": mingap,
                    "pairs": pares, "primera": fechas[0].strftime("%d/%m"),
                    "precio": r1(precio) if precio is not None else None}
            if mingap < 10:
                b10.append(item)
            elif mingap < 90:
                btri.append(item)
            elif mingap < ventana_dias:
                bsem.append(item)
        if not (b10 or btri or bsem):
            continue
        secciones.setdefault(tipo_norm, []).append(
            {"maq": maq, "tipo": tipo_raw, "b10": b10, "btri": btri, "bsem": bsem})
 
    out = []
    for tipo in ORDEN_TIPOS_COMPMIRAR:
        grupo = secciones.get(tipo)
        if not grupo:
            continue
        total = sum(it["veces"] for m in grupo for it in (m["b10"] + m["btri"] + m["bsem"]))
        out.append({"tipo": tipo, "maquinas": grupo, "total": total})
    return out
 
 
def _lectura_fluido(tipofluido, veces, mingap):
    """Genera el texto de diagnóstico según regla fija: tipo de fluido + veces + gap."""
    urgencia = "muy seguido" if mingap < 15 else "seguido" if mingap < 30 else "con una frecuencia algo alta"
    return (f"Se cargó {veces} veces en el período, {urgencia} (mínimo {mingap} días entre cargas) — "
            f"revisar posible pérdida de {tipofluido.lower()}.")
 
 
# Umbral real ya usado en la plantilla (ver nota fija en reunion_maquinaria.html.j2):
# hidráulico >= 12 veces EN EL AÑO, refrigerante >= 4 veces EN EL AÑO. No es una
# ventana móvil de N días -- es un conteo anual. Antes esto estaba mal (umbral=2
# sobre una ventana de 180 días), lo que generaba muchas más alertas de las reales.
UMBRAL_FLUIDO_ANUAL = {"Hidráulico": 12, "Refrigerante": 4}
 
# Categorías a excluir del reporte de fluidos, pedido explícito del usuario
# (ni caña, ni equipos a gas, ni ítems de seguridad son "maquinaria" para este reporte).
EXCLUIR_TIPO_KEYWORDS = ["CAÑA", "GAS", "SEGURIDAD"]
 
 
def _excluido_por_tipo(tipo_maquina, maquina):
    texto = f"{tipo_maquina or ''} {maquina or ''}".upper()
    return any(kw in texto for kw in EXCLUIR_TIPO_KEYWORDS)
 
 
def construir_fluidos(df_consumos_fluidos, anio=2026):
    """df_consumos_fluidos: igual a compMirar pero ya filtrado a insumos de fluido
    (rubro LUBRICANTE + nombre con HIDRAULICO/REFRIGERANTE), con columna extra
    'tipofluido' y 'tipo_maquina'. Devuelve una lista de SECCIONES por tipo de
    maquinaria (pedido del usuario), cada una con sus alertas, ordenadas por
    cantidad de alertas."""
    df = df_consumos_fluidos[df_consumos_fluidos["fecha_mov"].dt.year == anio].copy()
    if "tipo_maquina" in df.columns:
        df = df[~df.apply(lambda r: _excluido_por_tipo(r.get("tipo_maquina"), r.get("maquina")), axis=1)]
 
    secciones = {}
    for (maq, insumo, tipofluido, tipo_maq), g in df.groupby(["maquina", "repuesto", "tipofluido", "tipo_maquina"]):
        # Deduplicar por fecha de carga: ConsumosyReparaciones suele traer una compra
        # partida en varias líneas de factura el mismo día (ej. 3 líneas de "ACEITE
        # HIDRAULICO" el mismo 14/5). Contar cada línea como una carga distinta infla
        # artificialmente "veces" (se vio un caso de 21 líneas que en realidad eran
        # solo 3 fechas distintas de carga) y hace ver mingap=0 donde no hay ningún
        # recambio repetido real. Se cuenta por fecha, no por línea.
        fechas = sorted(set(g["fecha_mov"].tolist()))
        veces = len(fechas)
        umbral = UMBRAL_FLUIDO_ANUAL.get(tipofluido, 999)
        if veces < umbral:
            continue
        mingap = min((b - a).days for a, b in zip(fechas, fechas[1:])) if len(fechas) > 1 else 999
        item = {
            "maq": maq, "fluido": insumo, "tipofluido": tipofluido, "veces": veces,
            "primera": fechas[0].strftime("%d/%m"), "ultima": fechas[-1].strftime("%d/%m"),
            "lectura": _lectura_fluido(tipofluido, veces, mingap),
        }
        secciones.setdefault(tipo_maq if pd.notna(tipo_maq) else "Sin tipo", []).append(item)
 
    out = [{"tipo": tipo, "items": sorted(items, key=lambda x: -x["veces"])}
           for tipo, items in secciones.items()]
    out.sort(key=lambda s: -sum(i["veces"] for i in s["items"]))
    return out
 
 
def clasificar_fluido(nombre_insumo):
    nombre = (nombre_insumo or "").upper()
    for kw, label in FLUIDO_KEYWORDS.items():
        if kw in nombre:
            return label
    return None
 
 
# ---------------------------------------------------------------------------
# Bloque 12: topSupervisores
# ---------------------------------------------------------------------------
 
def construir_top_supervisores(df_actual):
    campo_rows = df_actual[df_actual["TipoLugar"] == "CAMPO"]
    out = []
    for sup, g in campo_rows.groupby("Supervisor"):
        if pd.isna(sup):
            continue
        visitas = g[["FechaDT", "Maquina"]].drop_duplicates().shape[0]
        recs = [{
            "f": row["FechaDT"].strftime("%d/%m/%Y"), "mec": row["Mecanico"], "maq": row["Maquina"],
            "sub": row["SubRubro"] if pd.notna(row["SubRubro"]) else None,
            "hs": r1(row["Horas"]), "hst": r1(row["HorasTraslado"]), "hsp": r1(row["HorasPrep"]),
        } for _, row in g.iterrows()]
        out.append({
            "n": sup, "visitas": visitas,
            "t": r1(g["Horas"].sum()), "tr": r1(g["HorasTraslado"].sum()), "p": r1(g["HorasPrep"].sum()),
            "recs": recs,
        })
    return out
 
 
# ---------------------------------------------------------------------------
# Bloque 1: quincenas (requiere histórico año en curso — extracto aparte)
# ---------------------------------------------------------------------------
 
MESES_ABREV = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
 
 
def _label_quincena(fecha_quincena):
    """1/1/2026 -> '1ra Ene'; 16/7/2026 -> '2da Jul'."""
    mitad = "1ra" if fecha_quincena.day == 1 else "2da"
    return f"{mitad} {MESES_ABREV[fecha_quincena.month - 1]}"
 
 
def construir_quincenas(df_horas_taller_dia, calendario_df):
    """df_horas_taller_dia: extracto ya agregado por día (columnas Fecha, Horas) de
    Horas_Personal_AppSheet TALLER, desde el 1/1 del año en curso (agregar por día
    en DAX y no traer fila a fila -- son ~230 días, no 20 mil filas).
    calendario_df: Calendario con columnas Date, Quincena, IndiceQuincena."""
    horas = df_horas_taller_dia.copy()
    horas["FechaDT"] = pd.to_datetime(horas["Fecha"], format="%d/%m/%Y", errors="coerce")
    cal = calendario_df.copy()
    cal["DateDT"] = pd.to_datetime(cal["Date"], format="%d/%m/%Y", errors="coerce")
    cal["QuincenaDT"] = pd.to_datetime(cal["Quincena"], format="%d/%m/%Y", errors="coerce")
 
    merged = horas.merge(cal[["DateDT", "QuincenaDT", "IndiceQuincena"]],
                          left_on="FechaDT", right_on="DateDT", how="left")
    faltantes = merged["IndiceQuincena"].isna().sum()
    if faltantes:
        merged = merged.dropna(subset=["IndiceQuincena"])
 
    out = []
    for (idx, quincena_dt), g in merged.groupby(["IndiceQuincena", "QuincenaDT"]):
        out.append({"q": _label_quincena(quincena_dt), "h": r1(g["Horas"].sum()), "_idx": idx})
    out.sort(key=lambda x: x["_idx"])
    for o in out:
        o.pop("_idx")
    return out, faltantes
 
 
# ---------------------------------------------------------------------------
# Punto de reemplazo para producción: Power BI Service REST API
# ---------------------------------------------------------------------------
 
def obtener_access_token_service_principal(tenant_id, client_id, client_secret):
    """Autentica el Service Principal (App Registration) contra Microsoft Entra ID
    usando OAuth2 client credentials flow, y devuelve el access token que necesita
    cargar_desde_power_bi_service(). Ver Guia_Service_Principal_Power_BI.md para
    cómo conseguir tenant_id/client_id/client_secret (una sola vez, hecho por el
    usuario en su Azure AD / Microsoft Entra ID)."""
    import requests
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    body = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://analysis.windows.net/powerbi/api/.default",
    }
    resp = requests.post(url, data=body, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]
 
 
def cargar_desde_power_bi_service(workspace_id, dataset_id, dax_query, access_token):
    """Ejecuta la consulta DAX contra el dataset publicado en Power BI Service usando
    la REST API 'Execute Queries', autenticado con un token de Service Principal
    (obtenido con obtener_access_token_service_principal()). Requiere que el Service
    Principal tenga rol Colaborador (o superior) en el workspace -- Visor no alcanza,
    la API de Execute Queries necesita permiso "Build" sobre el dataset, que Visor no
    incluye -- y que "Las entidades de servicio pueden llamar a las API públicas de
    Fabric" esté habilitado en el tenant (ver Guia_Service_Principal_Power_BI.md).
    Validado en producción el 05/08/2026 con una consulta real (COUNTROWS)."""
    import requests
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets/{dataset_id}/executeQueries"
    body = {"queries": [{"query": dax_query}], "serializerSettings": {"includeNulls": True}}
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    resp = requests.post(url, json=body, headers=headers, timeout=60)
    resp.raise_for_status()
    rows = resp.json()["results"][0]["tables"][0]["rows"]
    df = pd.DataFrame(rows)
    # La API devuelve los nombres de columna con corchetes (ej. "[Fecha]"), tal
    # como vino en la respuesta real de la prueba del usuario -- se limpian acá para
    # que el resto del motor reciba nombres de columna normales.
    df.columns = [str(c).strip("[]") for c in df.columns]
    return df
 
 
def parsear_fecha_pbi_service(serie):
    """Parsea una columna de fechas devuelta por la API de Power BI Service. Puede
    venir en formato ISO ('2026-07-20T00:00:00') o como texto dd/mm/aaaa (el mismo
    formato que devuelve Power BI Desktop) según el contexto -- se detecta mirando
    el primer valor no nulo, para no asumir un formato fijo sin haberlo visto."""
    s = pd.Series(serie)
    muestra = s.dropna().astype(str)
    if len(muestra) and len(muestra.iloc[0]) >= 4 and muestra.iloc[0][:4].isdigit() and "-" in muestra.iloc[0][:10]:
        return pd.to_datetime(s, errors="coerce")
    return pd.to_datetime(s, errors="coerce", dayfirst=True)
 
 
if __name__ == "__main__":
    import json
 
    df = cargar_horas(BASE / "raw_horas_3sem.csv")
    lunes_actual, domingo_actual = datetime.date(2026, 7, 27), datetime.date(2026, 8, 2)
    lunes_ant, domingo_ant = datetime.date(2026, 7, 20), datetime.date(2026, 7, 26)
 
    df_actual = filtrar_semana(df, lunes_actual, domingo_actual)
    df_anterior = filtrar_semana(df, lunes_ant, domingo_ant)
 
    mec_taller = construir_mec_taller(df_actual, df_anterior)
    mec_campo = construir_mec_campo(df_actual)
    rubros_taller = construir_rubros_taller(df_actual)
    mec_ausentes, mec_nuevos = construir_ausentes_nuevos(df_actual, df_anterior)
    estados_lf = construir_estados_lf(df_actual)
    maquinas_top = construir_maquinas_top(df_actual)
    top_supervisores = construir_top_supervisores(df_actual)
 
    # Liquidacion_MO (Chinagro), cruzado por Cuil
    df_liq = pd.read_csv(BASE / "raw_liquidacion_mo_semana.csv")
    horas_chinagro_data = construir_horas_chinagro(df_actual, df_liq)
 
    # ConsumosyReparaciones: fluidos y compMirar (extracto de prueba, rubro LUBRICANTE
    # de los últimos ~6 meses; en producción compMirar usaría el resto de rubros también)
    df_cons = pd.read_csv(BASE / "raw_fluidos.csv")
    df_cons["fecha_mov"] = pd.to_datetime(df_cons["fecha"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
    df_cons = df_cons.dropna(subset=["fecha_mov"])
    df_cons["tipofluido"] = df_cons["repuesto"].apply(clasificar_fluido)
    fluidos = construir_fluidos(df_cons[df_cons["tipofluido"].notna()])
    comp_mirar = construir_comp_mirar(df_cons)
 
    # quincenas: Calendario + horas de taller agregadas por día, año en curso
    df_horas_dia = pd.read_csv(BASE / "raw_horas_taller_2026.csv")
    df_calendario = pd.read_csv(BASE / "raw_calendario_2026.csv")
    quincenas, faltantes_quincena = construir_quincenas(df_horas_dia, df_calendario)
    if faltantes_quincena:
        print(f"AVISO: {faltantes_quincena} día(s) con horas no encontraron quincena en el Calendario (se excluyeron).")
 
    resultado = {
        "mecTaller": mec_taller,
        "mecCampo": mec_campo,
        "rubrosTaller": rubros_taller,
        "mecAusentes": mec_ausentes,
        "mecNuevos": mec_nuevos,
        "estadosLF": estados_lf,
        "horasChinagroData": horas_chinagro_data,
        "maquinasTop": maquinas_top,
        "topSupervisores": top_supervisores,
        "fluidos": fluidos,
        "compMirar": comp_mirar,
        "quincenas": quincenas,
    }
    out_path = BASE / "resultado_semana_prueba.json"
    out_path.write_text(json.dumps(resultado, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print("OK ->", out_path)
    for k, v in resultado.items():
        print(f"  {k}: {len(v)}")
 
 
# ---------------------------------------------------------------------------
# Narrativas automáticas (reemplazan los párrafos que hoy se redactan a mano)
# ---------------------------------------------------------------------------
 
def generar_nota_pico_horas(quincenas, fecha_hoy):
    """'La 2ª quincena de julio sigue siendo la más alta del año (3.451 hs).
    La 1ª quincena de agosto va 151 hs al 5/8 (aún en curso).'"""
    if not quincenas:
        return ""
    en_curso = quincenas[-1]
    completas = quincenas[:-1] if len(quincenas) > 1 else []
    partes = []
    if completas:
        pico = max(completas, key=lambda q: q["h"])
        ordinal = "más alta" if pico is completas[-1] or pico["h"] >= max(c["h"] for c in completas) else "alta"
        partes.append(f"La {pico['q']} es la quincena con más horas del año hasta ahora, con {pico['h']:.0f} hs.")
    partes.append(f"La {en_curso['q']} va {en_curso['h']:.0f} hs al {fecha_hoy.day}/{fecha_hoy.month} (aún en curso).")
    return " ".join(partes)
 
 
def generar_nota_lider_supervisores(top_supervisores, top_n=3):
    """'RESEO, Walter Eladio lidera la semana con 5 visitas (41 hs entre trabajo,
    traslado y preparación), seguido de VERGARA, Pedro Pablo con 4, y ...'"""
    if not top_supervisores:
        return "No hubo visitas a campo registradas esta semana."
    ordenado = sorted(top_supervisores, key=lambda s: (-s["visitas"], -(s["t"] + s["tr"] + s["p"])))
    lider = ordenado[0]
    total_hs = r1(lider["t"] + lider["tr"] + lider["p"])
    texto = f"{lider['n']} lidera la semana con {lider['visitas']} visitas ({total_hs:.0f} hs entre trabajo, traslado y preparación)"
    resto = ordenado[1:top_n]
    if resto:
        # agrupar por igual cantidad de visitas para armar "empate" como en el original
        siguientes = []
        i = 0
        while i < len(resto):
            visitas = resto[i]["visitas"]
            empatados = [r for r in resto[i:] if r["visitas"] == visitas]
            nombres = ", ".join(e["n"] for e in empatados)
            if len(empatados) > 1:
                etiqueta = {2: "doble", 3: "triple"}.get(len(empatados), "múltiple")
                siguientes.append(f"un {etiqueta} empate en {visitas} visitas entre {nombres}")
            else:
                siguientes.append(f"{nombres} con {visitas}")
            i += len(empatados)
        texto += ", seguido de " + ", seguido de ".join(siguientes)
    return texto + "."
 
 
def calcular_kpis_taller(df_actual, df_anterior, propiedad_actual):
    """Arma el dict `kpi`, `kpi_app` y la lista `talleres` que consume gestion_taller.html.j2."""
    total_taller_actual = df_actual[df_actual["TipoLugar"] == "TALLER"]["Horas"].sum()
    campo_rows = df_actual[df_actual["TipoLugar"] == "CAMPO"]
    total_campo_actual = (campo_rows["Horas"] + campo_rows["HorasTraslado"] + campo_rows["HorasPrep"]).sum()
    total_general = total_taller_actual + total_campo_actual
 
    total_taller_ant = df_anterior[df_anterior["TipoLugar"] == "TALLER"]["Horas"].sum()
    campo_ant = df_anterior[df_anterior["TipoLugar"] == "CAMPO"]
    total_campo_ant = (campo_ant["Horas"] + campo_ant["HorasTraslado"] + campo_ant["HorasPrep"]).sum()
    total_general_ant = total_taller_ant + total_campo_ant
 
    propia = float(propiedad_actual.get("PROPIAS", 0))
    terceros = float(propiedad_actual.get("TERCEROS", 0))
    clasificado = propia + terceros
    kpi = {
        "horas_totales": r1(total_general),
        "horas_taller": r1(total_taller_actual),
        "horas_campo": r1(total_campo_actual),
        "horas_totales_sa": r1(total_general_ant),
        "horas_variacion_pct": round((total_general - total_general_ant) / total_general_ant * 100) if total_general_ant else 0,
        "maquinaria_propia_hs": r1(propia),
        "maquinaria_propia_pct": round(propia / clasificado * 100) if clasificado else 0,
        "maquinaria_terceros_hs": r1(terceros),
        "maquinaria_terceros_pct": round(terceros / clasificado * 100) if clasificado else 0,
    }
 
    hs_campo_por_taller = campo_rows.groupby("Taller").apply(
        lambda g: (g["Horas"] + g["HorasTraslado"] + g["HorasPrep"]).sum()
    ).to_dict()
    hs_taller_por_taller = df_actual[df_actual["TipoLugar"] == "TALLER"].groupby("Taller")["Horas"].sum().to_dict()
    hs_taller_ant_por_taller = df_anterior[df_anterior["TipoLugar"] == "TALLER"].groupby("Taller")["Horas"].sum().to_dict()
    total_hs_taller_general = sum(hs_taller_por_taller.values()) or 1
 
    talleres = []
    for nombre, dotacion in DOTACION_FIJA.items():
        hs = hs_taller_por_taller.get(nombre, 0.0)
        hs_sa = hs_taller_ant_por_taller.get(nombre, 0.0)
        minimo = dotacion * 8 * 6
        talleres.append({
            "nombre": NOMBRE_LARGO[nombre], "css": CSS[nombre], "sigla": SIGLA[nombre],
            "hs_taller": r1(hs), "hs_taller_sa": r1(hs_sa),
            "variacion_pct": round((hs - hs_sa) / hs_sa * 100) if hs_sa else 0,
            "hs_campo": r1(hs_campo_por_taller.get(nombre, 0.0)),
            "dotacion_fija": dotacion,
            "hs_dia_mecanico": round(hs / (dotacion * 6), 1) if dotacion else 0,
            "minimo_esperado": minimo,
            "llego_al_minimo": hs >= minimo,
            "total_hs": r1(hs),
            "pct_del_total": round(hs / total_hs_taller_general * 100),
        })
 
    # Antes esto estaba hardcodeado (10 usuarios, 528 registros) copiado de una sola
    # consulta de prueba -- calculado en vivo: usuarios activos = emails distintos que
    # cargaron algo esa semana, registros totales = cantidad de filas cargadas.
    usuarios_activos = int(df_actual["Email"].dropna().nunique()) if "Email" in df_actual.columns else 0
    usuarios_activos_sa = int(df_anterior["Email"].dropna().nunique()) if "Email" in df_anterior.columns else 0
    registros_totales = int(len(df_actual))
    dias_con_datos = df_actual["FechaDT"].dt.date.nunique() if "FechaDT" in df_actual.columns and len(df_actual) else 0
    kpi_app = {
        "usuarios_activos": usuarios_activos, "usuarios_activos_sa": usuarios_activos_sa,
        "registros_totales": registros_totales,
        "registros_por_dia": round(registros_totales / dias_con_datos, 1) if dias_con_datos else 0,
        "registros_por_usuario": round(registros_totales / usuarios_activos, 1) if usuarios_activos else 0,
    }
    return kpi, kpi_app, talleres
