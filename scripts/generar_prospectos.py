#!/usr/bin/env python3
"""Genera una lista priorizada de prospectos de crédito a partir del DENUE (INEGI).

Uso:
    python scripts/generar_prospectos.py data/denue_nuevo_leon.csv [mas.csv ...] \
        --salida prospectos.xlsx
"""
import argparse
import re
import sys

import pandas as pd

# Municipios del Área Metropolitana de Monterrey (definición estándar de 12
# municipios usada por INEGI/CONAPO).
AMM_MUNICIPIOS = {
    "monterrey",
    "san pedro garza garcia",
    "san nicolas de los garza",
    "guadalupe",
    "apodaca",
    "general escobedo",
    "santa catarina",
    "garcia",
    "juarez",
    "cadereyta jimenez",
    "santiago",
    "general zuazua",
}

# Giros que no son sujetos de crédito empresarial y se excluyen del análisis.
GIROS_EXCLUIDOS = [
    "administracion publica",
    "actividades del gobierno",
    "agrupaciones politicas",
    "agrupaciones religiosas",
    "organizaciones religiosas",
    "hogares con empleados domesticos",
    "actividades no economicas",
]

# (patrón en el nombre de actividad, puntos de sector, producto sugerido)
SECTORES = [
    (r"construcci[oó]n|edificaci[oó]n|obras de ", 15, "Arrendamiento de equipo / Crédito simple"),
    (r"fabricaci[oó]n|industria manufacturera|confecci[oó]n|elaboraci[oó]n de", 15, "Crédito simple / Capital de trabajo"),
    (r"transporte|autotransporte|mensajer[ií]a y paqueter[ií]a", 12, "Arrendamiento de equipo (flotillas)"),
    (r"comercio al por mayor", 12, "Línea de crédito revolvente"),
    (r"comercio al por menor", 8, "Línea de crédito revolvente"),
    (r"servicios profesionales|consultor[ií]a|servicios de consultoria", 10, "Crédito simple / PyME"),
    (r"restaurantes|alimentos y bebidas|hoteles|alojamiento temporal", 8, "Capital de trabajo"),
    (r"agricultura|cr[ií]a y explotaci[oó]n de animales|pesca", 10, "Crédito simple / Avío"),
]

# Puntaje por rango de personal ocupado (refleja el "sweet spot" típico de
# banca empresarial PyME: negocios ya formalizados pero aún no atendidos por
# banca corporativa).
PUNTAJE_TAMANO = {
    "0 a 5 personas": 40,
    "6 a 10 personas": 65,
    "11 a 30 personas": 85,
    "31 a 50 personas": 95,
    "51 a 100 personas": 90,
    "101 a 250 personas": 75,
    "251 y mas personas": 55,
}


def _normalizar(texto: str) -> str:
    if not isinstance(texto, str):
        return ""
    texto = texto.strip().lower()
    reemplazos = str.maketrans("áéíóúñ", "aeioun")
    return texto.translate(reemplazos)


def cargar_denue(rutas: list[str]) -> pd.DataFrame:
    marcos = []
    for ruta in rutas:
        try:
            df = pd.read_csv(ruta, encoding="utf-8", dtype=str, low_memory=False)
        except UnicodeDecodeError:
            df = pd.read_csv(ruta, encoding="latin-1", dtype=str, low_memory=False)
        marcos.append(df)
    return pd.concat(marcos, ignore_index=True)


def calcular_zona(municipio: str, entidad: str) -> int:
    municipio_n = _normalizar(municipio)
    entidad_n = _normalizar(entidad)
    if municipio_n in AMM_MUNICIPIOS:
        return 1
    if "nuevo leon" in entidad_n:
        return 2
    return 3


def calcular_sector(nombre_act: str) -> tuple[int, str]:
    nombre_n = _normalizar(nombre_act)
    for patron, puntos, producto in SECTORES:
        if re.search(patron, nombre_n):
            return puntos, producto
    return 5, "Crédito simple / PyME (evaluar caso por caso)"


def es_giro_excluido(nombre_act: str) -> bool:
    nombre_n = _normalizar(nombre_act)
    return any(giro in nombre_n for giro in GIROS_EXCLUIDOS)


def construir_direccion(fila: pd.Series) -> str:
    numero_ext = fila.get("numero_ext")
    partes = [
        fila.get("tipo_vial") or "",
        fila.get("nom_vial") or "",
        f"#{numero_ext}" if pd.notna(numero_ext) else "",
    ]
    return " ".join(p for p in partes if p).strip()


def procesar(df: pd.DataFrame) -> pd.DataFrame:
    columnas_esperadas = {"nom_estab", "nombre_act", "per_ocu", "municipio", "entidad"}
    faltantes = columnas_esperadas - set(df.columns)
    if faltantes:
        raise ValueError(
            f"El CSV no tiene las columnas esperadas del DENUE: {sorted(faltantes)}"
        )

    df = df[~df["nombre_act"].apply(es_giro_excluido)].copy()

    df["zona_prioridad"] = df.apply(
        lambda f: calcular_zona(f.get("municipio", ""), f.get("entidad", "")), axis=1
    )

    sector = df["nombre_act"].apply(calcular_sector)
    df["puntos_sector"] = sector.apply(lambda x: x[0])
    df["producto_sugerido"] = sector.apply(lambda x: x[1])

    df["puntos_tamano"] = (
        df["per_ocu"].apply(_normalizar).map(PUNTAJE_TAMANO).fillna(40)
    )

    contacto = pd.Series(0, index=df.index)
    if "telefono" in df.columns:
        contacto += df["telefono"].notna() & (df["telefono"].astype(str).str.strip() != "")
        contacto = contacto.astype(int) * 10
    puntos_contacto = contacto.copy()
    if "correoelec" in df.columns:
        puntos_contacto += (
            df["correoelec"].notna() & (df["correoelec"].astype(str).str.strip() != "")
        ).astype(int) * 5
    if "www" in df.columns:
        puntos_contacto += (
            df["www"].notna() & (df["www"].astype(str).str.strip() != "")
        ).astype(int) * 5
    df["puntos_contacto"] = puntos_contacto

    df["score"] = (
        df["puntos_tamano"] * 0.55
        + df["puntos_sector"] * 1.5
        + df["puntos_contacto"]
    ).clip(upper=100).round(1)

    df["direccion"] = df.apply(construir_direccion, axis=1)

    columnas_salida = {
        "nom_estab": "Empresa",
        "raz_social": "Razón social",
        "nombre_act": "Giro",
        "municipio": "Municipio",
        "entidad": "Estado",
        "direccion": "Dirección",
        "telefono": "Teléfono",
        "correoelec": "Correo",
        "www": "Sitio web",
        "per_ocu": "Empleados",
        "score": "Score",
        "producto_sugerido": "Producto sugerido",
        "zona_prioridad": "zona_prioridad",
    }
    for col in columnas_salida:
        if col not in df.columns:
            df[col] = ""
    resultado = df[list(columnas_salida.keys())].rename(columns=columnas_salida)
    return resultado.sort_values(["zona_prioridad", "Score"], ascending=[True, False])


def exportar_excel(df: pd.DataFrame, salida: str) -> None:
    zonas = {
        1: "1_AMM_Monterrey",
        2: "2_Resto_NL",
        3: "3_Resto_Mexico",
    }
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        resumen = (
            df.groupby("zona_prioridad")
            .agg(Empresas=("Empresa", "count"), Score_promedio=("Score", "mean"))
            .rename(index=zonas)
        )
        resumen.to_excel(writer, sheet_name="Resumen")

        for zona_id, nombre_hoja in zonas.items():
            subset = df[df["zona_prioridad"] == zona_id].drop(columns=["zona_prioridad"])
            if subset.empty:
                continue
            subset.to_excel(writer, sheet_name=nombre_hoja, index=False)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("csv", nargs="+", help="Uno o más archivos CSV del DENUE")
    parser.add_argument(
        "--salida", default="prospectos.xlsx", help="Ruta del Excel de salida"
    )
    args = parser.parse_args()

    df = cargar_denue(args.csv)
    resultado = procesar(df)
    exportar_excel(resultado, args.salida)

    print(f"{len(resultado)} empresas procesadas -> {args.salida}")
    print(resultado["zona_prioridad"].map({1: "AMM", 2: "Resto NL", 3: "Resto México"}).value_counts())


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
