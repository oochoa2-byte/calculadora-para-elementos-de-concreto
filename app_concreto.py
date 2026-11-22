"""
Calculadora de cantidades directas - ELEMENTOS DE CONCRETO  (v2: precios dinámicos)
===================================================================================
Fuentes iniciales:
  * Presupuesto en PDF (partidas con precio unitario)  -> conceptos de concreto
  * Listado de insumos en Excel (E-07)                  -> insumos de concreto

Novedades v2
  * Los precios viven en una BASE DE DATOS SQLite (precios_concreto.db), no en el código.
  * Pestaña "Actualizar precios": edita cualquier precio, aplica ajustes porcentuales,
    restablece a los precios originales y consulta el historial de cambios.
  * Las calculadoras leen siempre los precios vigentes de la base de datos, así que
    un cambio de precio se refleja al instante en los cálculos.

Solo se procesan elementos de concreto (zapatas, dados, columnas, trabes, losas, dalas,
castillos, dinteles, plantilla/firme de concreto, bajadas de concreto) con su acero y
cimbra asociados. Todo lo demás se descarta.

Ejecutar:  streamlit run app_concreto.py
"""
import hashlib
import io
import re
import sqlite3
import unicodedata
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st

DB_PATH = Path(__file__).resolve().parent / "precios_concreto.db"
FUENTES = {"insumo": "Insumos (Excel)", "concepto": "Conceptos (PDF)"}

# ---------------------------------------------------------------------------
# 1. LECTURA Y FILTRADO ESTRICTO DE CONCRETO (Pandas)
# ---------------------------------------------------------------------------


def norm(txt) -> str:
    """Mayúsculas, sin acentos y con espacios colapsados."""
    txt = unicodedata.normalize("NFKD", str(txt))
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", txt).strip().upper()


def a_numero(valor):
    """Convierte '$ 1 ,304.32', '1,304.32', 12 o '12.5' a float. None si no se puede."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    limpio = re.sub(r"[^\d.\-]", "", str(valor))
    try:
        return float(limpio)
    except ValueError:
        return None


# ---- 1.A Insumos (Excel) ---------------------------------------------------

FAMILIAS_INSUMOS = [  # (patrón sobre texto normalizado, familia)
    (r"^CEMENTO", "CEMENTO PORTLAND"),
    (r"^ARENA AMARILLA", "ARENA AMARILLA"),
    (r"^ARENA DE RIO", "ARENA DE RIO"),
    (r"^GRAVA", "GRAVA"),
    (r"^AGUA", "AGUA"),
    (r"^ACERO REFUERZO", "ACERO DE REFUERZO FY=4200 (VARILLA)"),
    (r"^ALAMBRE RECOCIDO", "ALAMBRE RECOCIDO"),
    (r"^MALLA ELECTROSOLDADA", "MALLA ELECTROSOLDADA"),
    (r"^TRIPLAY", "TRIPLAY PARA CIMBRA"),
    (r"^BARROTE", "BARROTE PARA CIMBRA"),
    (r"^CLAVO PARA MADERA", "CLAVO PARA MADERA (CIMBRA)"),
    (r"^CURACRETO", "CURACRETO (MEMBRANA DE CURADO)"),
    (r"^OFICIAL FIERRERO", "OFICIAL FIERRERO"),
    (r"^AYUDANTE FIERRERO", "AYUDANTE FIERRERO"),
    (r"^REVOLVEDORA DE CONCRETO", "REVOLVEDORA DE CONCRETO"),
    (r"^VIBRADOR PARA CONCRETO", "VIBRADOR PARA CONCRETO"),
    (r"^CHICOTE PARA VIBRADOR", "CHICOTE PARA VIBRADOR"),
]
CATEGORIAS = {
    "MATERIALES": "Materiales",
    "MANO DE OBRA": "Mano de obra",
    "EQUIPO Y HERRAMIENTA": "Equipo y herramienta",
}


def familia_insumo(desc: str):
    """Familia de concreto del insumo, o None si debe descartarse."""
    d = norm(desc)
    m = re.match(r"^CONCRETO PREMEZCLADO.*?F'?C\s*=\s*(\d+)", d)
    if m:
        return f"CONCRETO PREMEZCLADO F'C={m.group(1)} KG/CM2"
    for patron, familia in FAMILIAS_INSUMOS:
        if re.match(patron, d):
            return familia
    return None


def cargar_insumos(data: bytes) -> pd.DataFrame:
    """Lee el Excel de insumos y conserva solo los insumos de concreto."""
    raw = pd.read_excel(io.BytesIO(data), header=None, dtype=object)
    cab = raw.index[raw[0].astype(str).map(norm).isin(["CODIGO"])]
    if len(cab) == 0:
        raise ValueError("No se encontró el encabezado 'Código' en el Excel.")
    categoria, filas = None, []
    for _, r in raw.iloc[cab[0] + 1:].iterrows():
        codigo, desc = r[0], r[1]
        if pd.isna(codigo):
            d = norm(desc) if isinstance(desc, str) else ""
            if d in CATEGORIAS:
                categoria = CATEGORIAS[d]
            continue
        precio = a_numero(r[5])
        familia = familia_insumo(desc)
        if precio is None or familia is None:
            continue
        concepto = re.sub(r"\s+", " ", str(desc)).strip()
        filas.append({
            "id": str(codigo).strip(), "categoria": categoria, "claves": str(codigo).strip(),
            "concepto": concepto, "descripcion": concepto, "unidad": str(r[2]).strip(),
            "precio": precio, "cantidad_ref": a_numero(r[4]) or 0.0,
            "grupo": f"{familia}|{str(r[2]).strip()}|{categoria}", "etiqueta_grupo": familia,
        })
    return pd.DataFrame(filas)


# ---- 1.B Conceptos de presupuesto (PDF) -----------------------------------

EXCLUIR_INICIO = (  # se evalúa sobre el inicio de la descripción (tipo de trabajo)
    r"tabic[oó]n|pretil|entortado|excavaci[oó]n|relleno|impermeabiliz|polietileno|"
    r"aplanado|muro de enrase|muro de mamposter|boquilla|afinaci[oó]n|trazo|limpieza|"
    r"proyecto"
)
REGLAS_PDF = [  # (tipo de elemento, patrón sobre los primeros 200 caracteres)
    ("Refuerzo de acero", r"^suministro y colocaci[oó]n de (acero de refuerzo|refuerzo tipo armex)"),
    ("Plantilla", r"\bplantilla de concreto"),
    ("Zapata", r"\bzapata\b"),
    ("Dado", r"\bdado de concreto"),
    ("Dintel / cadena", r"\bdintel\b|cadena de cerramiento"),
    ("Dala", r"\bdala\b"),
    ("Columna", r"\bcolumna\b"),
    ("Trabe", r"\btrabe\b"),
    ("Losa", r"\blosa\b"),
    ("Firme", r"\bfirme de concreto"),
    ("Castillo", r"\bcastillo\b"),
    ("Bajada pluvial", r"\bbajada pluvial"),
    ("Unión estructural", r"elemento de concreto armado"),
]
TIPOS_PROMEDIABLES = {"Trabe", "Castillo", "Dala", "Dintel / cadena", "Columna"}


def tipo_elemento(desc: str):
    if re.search(EXCLUIR_INICIO, desc[:60], flags=re.I):
        return None
    for tipo, patron in REGLAS_PDF:
        if re.search(patron, desc[:200], flags=re.I):
            return tipo
    return None


def cargar_conceptos_pdf(data: bytes) -> pd.DataFrame:
    """Extrae las partidas del PDF y conserva solo los conceptos de concreto."""
    filas, partida, ids = [], None, set()
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for pagina in pdf.pages:
            for tabla in pagina.extract_tables():
                for r in tabla:
                    if len(r) < 8:
                        continue
                    cod, desc, cant, um, _, pu, _, _ = r[:8]
                    if cod and desc and not cant and cod.strip().isupper() and len(cod.strip()) <= 5:
                        partida = desc.strip().title()   # fila de partida (CIM, EST, ALB...)
                        continue
                    cantidad, precio = a_numero(cant), a_numero(pu)
                    if not desc or cantidad is None or precio is None:
                        continue
                    descripcion = re.sub(r"\s+", " ", desc).strip()
                    tipo = tipo_elemento(descripcion)
                    if tipo is None:
                        continue
                    codigo = re.search(r"\b([A-Z]{1,3}-[0-9A-Z]{1,3})\b", descripcion[:200])
                    sec = re.search(r"secci[oó]n\s+(\d+(?:\.\d+)?)\s*[×xX]\s*(\d+(?:\.\d+)?)",
                                    descripcion, re.I)
                    fc = re.search(r"f.{0,2}c\s*=\s*(\d+)", descripcion, re.I)
                    seccion = f"{sec.group(1)}×{sec.group(2)}" if sec else None
                    fc = fc.group(1) if fc else None
                    codigo = codigo.group(1) if codigo else ""
                    corto = f"{tipo} {codigo}".strip()
                    if seccion:
                        corto += f" · {seccion} cm"
                    if fc:
                        corto += f" · f'c={fc}"
                    unidad = (um or "").strip()
                    if tipo in TIPOS_PROMEDIABLES and seccion and fc:
                        grupo = f"{tipo}|{seccion}|{fc}|{unidad}"
                        etiqueta = f"{tipo} · {seccion} cm · f'c={fc}"
                    else:
                        grupo, etiqueta = f"unico|{len(filas)}", corto
                    id_ = corto
                    while id_ in ids:
                        id_ += "*"
                    ids.add(id_)
                    filas.append({
                        "id": id_, "categoria": partida, "claves": codigo or tipo,
                        "concepto": corto, "descripcion": descripcion, "unidad": unidad,
                        "precio": precio, "cantidad_ref": cantidad,
                        "grupo": grupo, "etiqueta_grupo": etiqueta,
                    })
    return pd.DataFrame(filas)


# ---------------------------------------------------------------------------
# 2. BASE DE DATOS DE PRECIOS (SQLite)
# ---------------------------------------------------------------------------


@contextmanager
def conexion():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    try:
        yield con
        con.commit()
    finally:
        con.close()


def ahora() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def init_db():
    with conexion() as con:
        con.executescript("""
            CREATE TABLE IF NOT EXISTS precios (
                fuente TEXT NOT NULL, id TEXT NOT NULL,
                categoria TEXT, claves TEXT, concepto TEXT, descripcion TEXT, unidad TEXT,
                grupo TEXT, etiqueta_grupo TEXT, cantidad_ref REAL,
                precio_base REAL NOT NULL, precio REAL NOT NULL, actualizado TEXT,
                PRIMARY KEY (fuente, id));
            CREATE TABLE IF NOT EXISTS historial (
                n INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, fuente TEXT, item_id TEXT,
                concepto TEXT, precio_anterior REAL, precio_nuevo REAL, motivo TEXT);
        """)


def _log(con, fuente, item_id, concepto, antes, despues, motivo):
    con.execute("INSERT INTO historial (fecha, fuente, item_id, concepto, precio_anterior,"
                " precio_nuevo, motivo) VALUES (?,?,?,?,?,?,?)",
                (ahora(), fuente, item_id, concepto, antes, despues, motivo))


def sembrar(df: pd.DataFrame, fuente: str, reemplazar: bool = False) -> int:
    """Carga el catálogo limpio en la BD. Los renglones nuevos entran con el precio del archivo;
    los existentes conservan el precio editado, salvo que `reemplazar` sea True."""
    cambios = 0
    with conexion() as con:
        for r in df.to_dict("records"):
            ex = con.execute("SELECT precio FROM precios WHERE fuente=? AND id=?",
                             (fuente, r["id"])).fetchone()
            if ex is None:
                con.execute(
                    "INSERT INTO precios VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (fuente, r["id"], r["categoria"], r["claves"], r["concepto"], r["descripcion"],
                     r["unidad"], r["grupo"], r["etiqueta_grupo"], r["cantidad_ref"],
                     r["precio"], r["precio"], ahora()))
                continue
            con.execute(
                "UPDATE precios SET categoria=?, claves=?, concepto=?, descripcion=?, unidad=?,"
                " grupo=?, etiqueta_grupo=?, cantidad_ref=? WHERE fuente=? AND id=?",
                (r["categoria"], r["claves"], r["concepto"], r["descripcion"], r["unidad"],
                 r["grupo"], r["etiqueta_grupo"], r["cantidad_ref"], fuente, r["id"]))
            if reemplazar:
                if abs(ex["precio"] - r["precio"]) > 1e-9:
                    _log(con, fuente, r["id"], r["concepto"], ex["precio"], r["precio"],
                         "Importado de archivo")
                    cambios += 1
                con.execute("UPDATE precios SET precio=?, precio_base=?, actualizado=?"
                            " WHERE fuente=? AND id=?",
                            (r["precio"], r["precio"], ahora(), fuente, r["id"]))
    return cambios


def leer_catalogo(fuente: str) -> pd.DataFrame:
    with conexion() as con:
        return pd.read_sql_query("SELECT * FROM precios WHERE fuente=? ORDER BY rowid",
                                 con, params=(fuente,))


def guardar_precios(fuente: str, editado: pd.DataFrame) -> int:
    """Guarda en la BD los precios que difieran de los vigentes. Devuelve # de cambios."""
    n = 0
    with conexion() as con:
        vigentes = {r["id"]: (r["precio"], r["concepto"]) for r in
                    con.execute("SELECT id, precio, concepto FROM precios WHERE fuente=?", (fuente,))}
        for id_, p in zip(editado["id"], editado["precio"]):
            if id_ not in vigentes or p is None or pd.isna(p) or p < 0:
                continue
            antes, concepto = vigentes[id_]
            if abs(float(p) - antes) > 1e-9:
                con.execute("UPDATE precios SET precio=?, actualizado=? WHERE fuente=? AND id=?",
                            (float(p), ahora(), fuente, id_))
                _log(con, fuente, id_, concepto, antes, float(p), "Edición manual")
                n += 1
    return n


def ajuste_porcentual(fuente: str, categoria, pct: float) -> int:
    """Sube/baja `pct`% los precios de una categoría (o de todas si categoria es None)."""
    n = 0
    with conexion() as con:
        filas = con.execute("SELECT id, concepto, categoria, precio FROM precios WHERE fuente=?",
                            (fuente,)).fetchall()
        for r in filas:
            if categoria and r["categoria"] != categoria:
                continue
            nuevo = round(r["precio"] * (1 + pct / 100), 2)
            if abs(nuevo - r["precio"]) > 1e-9:
                con.execute("UPDATE precios SET precio=?, actualizado=? WHERE fuente=? AND id=?",
                            (nuevo, ahora(), fuente, r["id"]))
                _log(con, fuente, r["id"], r["concepto"], r["precio"], nuevo,
                     f"Ajuste {pct:+.1f}% ({categoria or 'todas'})")
                n += 1
    return n


def restablecer(fuente: str) -> int:
    """Regresa todos los precios al valor original del archivo."""
    n = 0
    with conexion() as con:
        for r in con.execute("SELECT id, concepto, precio, precio_base FROM precios WHERE fuente=?",
                             (fuente,)).fetchall():
            if abs(r["precio"] - r["precio_base"]) > 1e-9:
                con.execute("UPDATE precios SET precio=?, actualizado=? WHERE fuente=? AND id=?",
                            (r["precio_base"], ahora(), fuente, r["id"]))
                _log(con, fuente, r["id"], r["concepto"], r["precio"], r["precio_base"],
                     "Restablecido a precio original")
                n += 1
    return n


def leer_historial(limite: int = 200) -> pd.DataFrame:
    with conexion() as con:
        return pd.read_sql_query(
            "SELECT fecha, fuente, concepto, precio_anterior, precio_nuevo, motivo "
            "FROM historial ORDER BY n DESC LIMIT ?", con, params=(limite,))


# ---------------------------------------------------------------------------
# 3. PRECIOS VIGENTES + PROMEDIO DE CONCEPTOS REPETIDOS
# ---------------------------------------------------------------------------


def preparar_precios(df: pd.DataFrame, promediar: bool) -> pd.DataFrame:
    """Un renglón por concepto con los precios VIGENTES de la BD. Si `promediar`, los conceptos
    repetidos (mismo grupo) se unifican con el promedio de sus precios actuales."""
    filas = []
    for _, sub in df.groupby("grupo", sort=False):
        if len(sub) == 1 or not promediar:
            for r in sub.to_dict("records"):
                r.update(rid=r["id"], n_precios=1, precios_individuales=f"{r['precio']:,.2f}")
                filas.append(r)
            continue
        fila = sub.iloc[0].to_dict()
        fila.update(
            rid=sub["grupo"].iloc[0],
            concepto=f"{sub['etiqueta_grupo'].iloc[0]} (prom. de {len(sub)})",
            claves=", ".join(sub["claves"]),
            precio=round(sub["precio"].mean(), 2),
            precio_base=round(sub["precio_base"].mean(), 2),
            cantidad_ref=sub["cantidad_ref"].sum(),
            n_precios=len(sub),
            precios_individuales=" | ".join(f"{p:,.2f}" for p in sub["precio"]),
        )
        filas.append(fila)
    return pd.DataFrame(filas).reset_index(drop=True)


# ---------------------------------------------------------------------------
# 4. INTERFAZ STREAMLIT
# ---------------------------------------------------------------------------


@st.cache_data(show_spinner="Procesando Excel...")
def _insumos(data: bytes) -> pd.DataFrame:
    return cargar_insumos(data)


@st.cache_data(show_spinner="Procesando PDF...")
def _conceptos(data: bytes) -> pd.DataFrame:
    return cargar_conceptos_pdf(data)


def archivo_por_defecto(patron: str):
    hallados = sorted(Path(__file__).resolve().parent.glob(patron))
    return hallados[0].read_bytes() if hallados else None


def moneda(x: float) -> str:
    return f"$ {x:,.2f}"


def panel_precios(fuente: str, df_archivo):
    """Editor de precios de una fuente. Cada cambio se guarda en la BD al instante."""
    cat = leer_catalogo(fuente)
    if cat.empty:
        st.warning("Sin datos: sube el archivo correspondiente en la barra lateral.")
        return
    st.caption("Edita la columna **Precio actual**: se guarda automáticamente en la base de datos "
               "y las calculadoras se actualizan al instante. 'Precio original' es el del archivo.")
    ver = st.session_state["ver"]
    vista = cat[["id", "categoria", "concepto", "unidad", "precio_base", "precio", "actualizado"]]
    editado = st.data_editor(
        vista, key=f"pe_{fuente}_{ver}", hide_index=True,
        disabled=[c for c in vista.columns if c != "precio"],
        column_config={
            "id": None,
            "categoria": st.column_config.TextColumn("Categoría"),
            "concepto": st.column_config.TextColumn("Concepto", width="large"),
            "unidad": st.column_config.TextColumn("Unidad", width="small"),
            "precio_base": st.column_config.NumberColumn("Precio original", format="$ %.2f"),
            "precio": st.column_config.NumberColumn("✏️ Precio actual", min_value=0.0,
                                                    step=0.5, format="$ %.2f"),
            "actualizado": st.column_config.TextColumn("Última actualización"),
        })
    if guardar_precios(fuente, editado):
        st.rerun()

    c1, c2, c3 = st.columns([2, 1, 1])
    categorias = ["Todas"] + sorted(cat["categoria"].dropna().unique())
    sel = c1.selectbox("Categoría a ajustar", categorias, key=f"cat_{fuente}")
    pct = c2.number_input("Ajuste (%)", -50.0, 100.0, 0.0, 0.5, key=f"pct_{fuente}")
    if c3.button("Aplicar ajuste", key=f"aj_{fuente}"):
        ajuste_porcentual(fuente, None if sel == "Todas" else sel, pct)
        st.session_state["ver"] += 1
        st.rerun()

    b1, b2, b3 = st.columns(3)
    if b1.button("↩️ Restablecer a precios originales", key=f"rs_{fuente}"):
        restablecer(fuente)
        st.session_state["ver"] += 1
        st.rerun()
    if df_archivo is not None and b2.button("📥 Reemplazar con precios del archivo cargado",
                                            key=f"im_{fuente}"):
        sembrar(df_archivo, fuente, reemplazar=True)
        st.session_state["ver"] += 1
        st.rerun()
    b3.download_button("⬇️ Descargar precios vigentes (CSV)",
                       cat[["categoria", "claves", "concepto", "unidad", "precio_base", "precio",
                            "actualizado"]].to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"precios_{fuente}.csv", mime="text/csv", key=f"dp_{fuente}")


def calculadora(df: pd.DataFrame, fuente: str, modo: str) -> pd.DataFrame:
    """Tabla editable de cantidades. Devuelve el df con importes a precios vigentes y originales."""
    cant = st.session_state["cant"]
    base = df[["rid", "categoria", "concepto", "unidad", "precio", "n_precios",
               "precios_individuales", "cantidad_ref"]].copy()
    base["cantidad"] = [cant.get(f"{fuente}|{modo}|{r}", 0.0) for r in base["rid"]]
    editado = st.data_editor(
        base, key=f"ed_{fuente}_{modo}", hide_index=True,
        disabled=[c for c in base.columns if c != "cantidad"],
        column_config={
            "rid": None,
            "categoria": st.column_config.TextColumn("Categoría"),
            "concepto": st.column_config.TextColumn("Concepto", width="large"),
            "unidad": st.column_config.TextColumn("Unidad", width="small"),
            "precio": st.column_config.NumberColumn("P.U. vigente", format="$ %.2f"),
            "n_precios": st.column_config.NumberColumn("# precios promediados", format="%d"),
            "precios_individuales": st.column_config.TextColumn("Precios que se promedian"),
            "cantidad_ref": st.column_config.NumberColumn("Cant. en fuente", format="%.2f"),
            "cantidad": st.column_config.NumberColumn("✏️ CANTIDAD", min_value=0.0, step=1.0,
                                                      format="%.2f"),
        })
    editado = editado.copy()
    editado["cantidad"] = editado["cantidad"].fillna(0.0)
    for rid, c in zip(editado["rid"], editado["cantidad"]):
        cant[f"{fuente}|{modo}|{rid}"] = float(c)          # las cantidades sobreviven a cambios de precio
    editado["importe"] = editado["cantidad"] * editado["precio"]
    editado["importe_original"] = editado["cantidad"] * df["precio_base"].values
    return editado


def metrica_total(titulo: str, res: pd.DataFrame):
    total, original = res["importe"].sum(), res["importe_original"].sum()
    dif = total - original
    st.metric(titulo, moneda(total),
              delta=f"{dif:+,.2f} vs precios originales" if abs(dif) > 0.004 else None,
              delta_color="inverse")


def main():
    st.set_page_config(page_title="Calculadora de concreto", page_icon="🏗️", layout="wide")
    init_db()
    st.session_state.setdefault("ver", 0)
    st.session_state.setdefault("cant", {})

    st.title("🏗️ Calculadora de cantidades directas · Elementos de concreto")
    st.caption("Captura cantidades de obra; el costo se calcula con los precios vigentes de la "
               "base de datos (editables en la pestaña «Actualizar precios»).")

    with st.sidebar:
        st.header("Archivos fuente")
        st.caption("Solo se necesitan para cargar el catálogo inicial; después los precios "
                   "viven en la base de datos.")
        up_xlsx = st.file_uploader("Listado de insumos (.xlsx)", type="xlsx")
        up_pdf = st.file_uploader("Presupuesto (.pdf)", type="pdf")
        st.header("Opciones")
        promediar = st.radio(
            "Conceptos repetidos con precios distintos",
            ["Usar precio promedio", "Usar cada precio individual"],
        ) == "Usar precio promedio"
        iva = st.number_input("IVA (%)", min_value=0.0, max_value=30.0, value=16.0, step=1.0)

    data_xlsx = up_xlsx.getvalue() if up_xlsx else archivo_por_defecto("*Listado_Insumos*.xlsx")
    data_pdf = up_pdf.getvalue() if up_pdf else archivo_por_defecto("*Presupuesto*.pdf")
    df_xlsx = _insumos(data_xlsx) if data_xlsx else None
    df_pdf = _conceptos(data_pdf) if data_pdf else None
    for fuente, df, data in (("insumo", df_xlsx, data_xlsx), ("concepto", df_pdf, data_pdf)):
        if df is not None:   # sembrar solo cuando cambia el archivo (no pisa precios editados)
            firma = hashlib.md5(data).hexdigest()
            if st.session_state.get(f"firma_{fuente}") != firma:
                sembrar(df, fuente)
                st.session_state[f"firma_{fuente}"] = firma

    cat_ins, cat_conc = leer_catalogo("insumo"), leer_catalogo("concepto")
    if cat_ins.empty and cat_conc.empty:
        st.info("Sube el Excel de insumos y/o el PDF del presupuesto en la barra lateral "
                "(o colócalos junto a este script) para crear el catálogo inicial.")
        st.stop()

    modo = "prom" if promediar else "ind"
    t_ins, t_conc, t_precios, t_res = st.tabs(
        ["🧱 Calculadora · Insumos", "📄 Calculadora · Conceptos",
         "💲 Actualizar precios", "📊 Resumen"])
    res = {"insumo": None, "concepto": None}

    # La pestaña de precios se procesa primero para que los cambios se vean en el mismo ciclo.
    with t_precios:
        sub_i, sub_c, sub_h = st.tabs(["Insumos (Excel)", "Conceptos (PDF)", "Historial de cambios"])
        with sub_i:
            panel_precios("insumo", df_xlsx)
        with sub_c:
            panel_precios("concepto", df_pdf)
        with sub_h:
            hist = leer_historial()
            if hist.empty:
                st.write("Aún no hay cambios de precio.")
            else:
                st.dataframe(hist, hide_index=True, column_config={
                    "fecha": "Fecha", "fuente": "Fuente", "concepto": "Concepto",
                    "precio_anterior": st.column_config.NumberColumn("Anterior", format="$ %.2f"),
                    "precio_nuevo": st.column_config.NumberColumn("Nuevo", format="$ %.2f"),
                    "motivo": "Motivo"})

    for fuente, tab, catalogo, texto in (
        ("insumo", t_ins, leer_catalogo("insumo"),
         "insumos de concreto (concreto, cemento, agregados, acero, alambre, malla, cimbra, "
         "mano de obra fierrera y equipo)"),
        ("concepto", t_conc, leer_catalogo("concepto"),
         "conceptos de concreto (zapatas, dados, dalas, columnas, trabes, losas, castillos, "
         "firme, etc.)"),
    ):
        with tab:
            if catalogo.empty:
                st.warning(f"Sin datos de {FUENTES[fuente]}. Sube el archivo en la barra lateral.")
                continue
            datos = preparar_precios(catalogo, promediar)
            st.write(f"**{len(datos)} {texto}.** Captura la cantidad en la última columna.")
            res[fuente] = calculadora(datos, fuente, modo)
            metrica_total(f"Importe total · {FUENTES[fuente]} · sin IVA", res[fuente])

    with t_res:
        st.caption("Los insumos son el desglose de los conceptos; por eso se muestran por separado "
                   "y no se suman entre sí.")
        for fuente in ("concepto", "insumo"):
            st.subheader(f"Por {FUENTES[fuente].lower()}")
            r = res[fuente]
            if r is None:
                st.write("Sin datos.")
                continue
            sub = r["importe"].sum()
            c1, c2, c3 = st.columns(3)
            c1.metric("Subtotal", moneda(sub))
            c2.metric(f"IVA ({iva:.0f}%)", moneda(sub * iva / 100))
            c3.metric("IMPORTE TOTAL", moneda(sub * (1 + iva / 100)))
            capt = r[r["cantidad"] > 0]
            if capt.empty:
                st.write("Captura cantidades para ver el desglose.")
                continue
            detalle = capt[["categoria", "concepto", "unidad", "cantidad", "precio", "importe"]]
            st.dataframe(detalle, hide_index=True, column_config={
                "categoria": "Categoría", "concepto": "Concepto", "unidad": "Unidad",
                "cantidad": st.column_config.NumberColumn("Cantidad", format="%.2f"),
                "precio": st.column_config.NumberColumn("P.U. vigente", format="$ %.2f"),
                "importe": st.column_config.NumberColumn("Importe", format="$ %.2f")})
            por_cat = capt.groupby("categoria", as_index=False)["importe"].sum()
            por_cat.loc[len(por_cat)] = ["TOTAL (sin IVA)", por_cat["importe"].sum()]
            st.dataframe(por_cat, hide_index=True, column_config={
                "categoria": "Subtotal por categoría",
                "importe": st.column_config.NumberColumn("Importe", format="$ %.2f")})
            st.download_button("⬇️ Descargar resumen (CSV)",
                               detalle.to_csv(index=False).encode("utf-8-sig"),
                               file_name=f"resumen_{fuente}.csv", mime="text/csv",
                               key=f"dl_{fuente}")


if __name__ == "__main__":
    main()
