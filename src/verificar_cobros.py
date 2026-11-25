"""
Verificador de cobros recurrentes.

Detecta cobros duplicados, cargos no contratados y variaciones de tarifa
en estados de cuenta bancarios con pólizas o servicios de cobro mensual.

El método central es la ingeniería inversa de la tarifa: cada cobro de una
póliza es el producto exacto de un factor fijo (la suma asegurada) por una
tarifa mensual. Al dividir el monto cobrado entre el factor conocido se
obtiene la tarifa aplicada, y cualquier tarifa que se salga del consenso
del ciclo delata un cobro mal calculado o emitido por otro proceso.

Uso:
    python src/verificar_cobros.py --movimientos data/movimientos.csv \\
                                   --catalogo data/catalogo.csv \\
                                   --dia-corte 21
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

ETIQUETA_SIN_REFERENCIA = "NO RECONOCIDA"


# --------------------------------------------------------------------------- #
# Carga
# --------------------------------------------------------------------------- #

def cargar_catalogo(ruta: Path) -> pd.DataFrame:
    """Lee el catálogo de servicios contratados.

    El catálogo es una lista blanca: cualquier cargo cuya descripción no
    contenga una de estas referencias se marca como no contratado. Esa es la
    propiedad que permite detectar cargos nuevos sin conocerlos de antemano.
    """
    catalogo = pd.read_csv(ruta, dtype={"referencia": "string"})
    columnas = {"referencia", "cuenta_esperada", "factor", "tipo"}
    faltantes = columnas - set(catalogo.columns)
    if faltantes:
        raise ValueError(f"Al catálogo le faltan columnas: {sorted(faltantes)}")

    catalogo["referencia"] = catalogo["referencia"].str.strip()
    if catalogo["referencia"].duplicated().any():
        raise ValueError("Hay referencias repetidas en el catálogo")

    return catalogo


def cargar_movimientos(ruta: Path) -> pd.DataFrame:
    """Lee los movimientos del estado de cuenta."""
    movimientos = pd.read_csv(ruta, parse_dates=["fecha"])
    columnas = {"fecha", "cuenta_cargada", "descripcion", "monto"}
    faltantes = columnas - set(movimientos.columns)
    if faltantes:
        raise ValueError(f"A los movimientos les faltan columnas: {sorted(faltantes)}")

    movimientos["descripcion"] = movimientos["descripcion"].str.strip()
    return movimientos.sort_values("fecha").reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Enriquecimiento
# --------------------------------------------------------------------------- #

def detectar_referencia(descripcion: str, referencias: pd.Series) -> str:
    """Devuelve la referencia del catálogo contenida en la descripción.

    Se busca por subcadena y no extrayendo la última palabra, porque los
    bancos no son consistentes con los separadores: `DB. SEG.BLNL220931`
    pega la referencia al prefijo sin espacio, mientras que
    `SEGURO FRAUDE BLNL220931` la separa. La búsqueda por subcadena resuelve
    ambos casos con la misma regla.
    """
    texto = descripcion.upper()
    for referencia in referencias:
        if referencia.upper() in texto:
            return referencia
    return ETIQUETA_SIN_REFERENCIA


def asignar_ciclo(fechas: pd.Series, dia_corte: int) -> pd.Series:
    """Asigna cada fecha a su ciclo de facturación.

    Un ciclo con corte el día 21 va del día 22 de un mes al 21 del siguiente,
    y se nombra por el mes en que cierra. Trabajar por mes calendario en vez
    de por ciclo es el error que hace invisible la duplicación: dos corridas
    de facturación separadas por tres semanas caen en meses distintos y
    parecen normales.
    """
    desplazamiento = (fechas.dt.day > dia_corte).astype(int)
    meses = fechas.dt.year * 12 + (fechas.dt.month - 1) + desplazamiento
    return meses.floordiv(12).astype(str) + "-" + (meses.mod(12) + 1).map("{:02d}".format)


def enriquecer(
    movimientos: pd.DataFrame, catalogo: pd.DataFrame, dia_corte: int
) -> pd.DataFrame:
    """Agrega referencia, ciclo, cuenta esperada, tarifa implícita y conteos."""
    df = movimientos.copy()

    df["referencia"] = df["descripcion"].apply(
        detectar_referencia, referencias=catalogo["referencia"]
    )
    df["ciclo"] = asignar_ciclo(df["fecha"], dia_corte)

    df = df.merge(
        catalogo[["referencia", "cuenta_esperada", "factor", "tipo"]],
        on="referencia",
        how="left",
    )

    df["cuenta_correcta"] = (
        df["cuenta_cargada"].str.strip().str.upper()
        == df["cuenta_esperada"].str.strip().str.upper()
    )
    df["tarifa_implicita"] = (df["monto"] / df["factor"]).round(4)
    df["veces_cobrada"] = df.groupby(["ciclo", "referencia"])["referencia"].transform(
        "size"
    )

    return df


def tarifa_vigente(df: pd.DataFrame, ciclo: str) -> float | None:
    """Tarifa que rige el ciclo: la más frecuente en su fecha de cierre.

    La fecha de cierre es la autoritativa porque es cuando el sistema de
    facturación corre de forma programada. Los cargos posteados en fechas
    intermedias suelen venir de corridas manuales o de procesos heredados.
    """
    del_ciclo = df[(df["ciclo"] == ciclo) & df["tarifa_implicita"].notna()]
    if del_ciclo.empty:
        return None

    cierre = del_ciclo[del_ciclo["fecha"] == del_ciclo["fecha"].max()]
    modas = cierre["tarifa_implicita"].mode()
    return float(modas.iloc[0]) if not modas.empty else None


# --------------------------------------------------------------------------- #
# Diagnóstico
# --------------------------------------------------------------------------- #

def diagnosticar(df: pd.DataFrame) -> pd.DataFrame:
    """Asigna a cada movimiento su diagnóstico, por orden de gravedad."""
    df = df.copy()

    condiciones = [
        df["referencia"].eq(ETIQUETA_SIN_REFERENCIA),
        df["veces_cobrada"].gt(1),
        ~df["cuenta_correcta"],
    ]
    etiquetas = [
        "NO CONTRATADO",
        "DUPLICADO",
        "CUENTA INCORRECTA",
    ]

    df["diagnostico"] = "OK"
    for condicion, etiqueta in zip(reversed(condiciones), reversed(etiquetas)):
        df.loc[condicion, "diagnostico"] = etiqueta

    return df


def resumen_ciclo(df: pd.DataFrame, ciclo: str, polizas_esperadas: int) -> pd.DataFrame:
    """Tablero de control: los indicadores que deben revisarse de un vistazo."""
    del_ciclo = df[df["ciclo"] == ciclo]
    seguros = del_ciclo[del_ciclo["tipo"].eq("SEGURO")]
    tarifas = seguros["tarifa_implicita"].dropna().unique()

    sin_reconocer = int(del_ciclo["referencia"].eq(ETIQUETA_SIN_REFERENCIA).sum())
    indicadores = [
        ("Cobros de seguro", f"{len(seguros)}", f"{polizas_esperadas}"),
        ("Cargos sin reconocer", f"{sin_reconocer}", "0"),
        ("Total del ciclo", f"{del_ciclo['monto'].sum():,.2f}", "-"),
        ("Tarifas distintas", f"{len(tarifas)}", "1"),
        ("Alertas abiertas", f"{int(del_ciclo['diagnostico'].ne('OK').sum())}", "0"),
    ]

    resumen = pd.DataFrame(indicadores, columns=["indicador", "valor", "esperado"])
    resumen["bandera"] = [
        "" if esp == "-" or val == esp else "REVISAR"
        for val, esp in zip(resumen["valor"], resumen["esperado"])
    ]
    return resumen


def calcular_exceso(df: pd.DataFrame, ciclo: str) -> pd.DataFrame:
    """Cuantifica el cobro en exceso por póliza dentro de un ciclo.

    Lo esperado de cada póliza activa es un único cobro a la tarifa vigente.
    Lo cobrado es la suma de todas sus líneas. La diferencia es el exceso.
    """
    tarifa = tarifa_vigente(df, ciclo)
    if tarifa is None:
        return pd.DataFrame()

    seguros = df[(df["ciclo"] == ciclo) & df["tipo"].ne("COMISION")]
    filas = []

    for referencia, grupo in seguros.groupby("referencia", dropna=False):
        cobrado = grupo["monto"].sum()
        factor = grupo["factor"].dropna()
        esperado = float(factor.iloc[0]) * tarifa if not factor.empty else 0.0
        filas.append(
            {
                "referencia": referencia,
                "cobrado": round(cobrado, 2),
                "esperado": round(esperado, 2),
                "exceso": round(cobrado - esperado, 2),
            }
        )

    # Las comisiones no tienen factor: lo esperado es un cobro por período.
    comisiones = df[(df["ciclo"] == ciclo) & df["tipo"].eq("COMISION")]
    for (referencia, _cuenta), grupo in comisiones.groupby(
        ["referencia", "cuenta_cargada"]
    ):
        cobrado = grupo["monto"].sum()
        esperado = float(grupo.sort_values("fecha")["monto"].iloc[-1])
        filas.append(
            {
                "referencia": referencia,
                "cobrado": round(cobrado, 2),
                "esperado": round(esperado, 2),
                "exceso": round(cobrado - esperado, 2),
            }
        )

    exceso = pd.DataFrame(filas).sort_values("exceso", ascending=False)
    return exceso.reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Salida
# --------------------------------------------------------------------------- #

def _moneda(valor: float) -> str:
    return f"{valor:>12,.2f}"


def imprimir_reporte(df: pd.DataFrame, ciclo: str, polizas_esperadas: int) -> None:
    ancho = 104
    print("=" * ancho)
    print(f"VERIFICACIÓN DE COBROS — ciclo {ciclo}".center(ancho))
    print("=" * ancho)

    del_ciclo = df[df["ciclo"] == ciclo]

    print("\nMOVIMIENTOS\n")
    detalle = del_ciclo[
        [
            "fecha",
            "cuenta_cargada",
            "descripcion",
            "monto",
            "referencia",
            "tarifa_implicita",
            "veces_cobrada",
            "diagnostico",
        ]
    ].copy()
    detalle["fecha"] = detalle["fecha"].dt.strftime("%d/%m/%Y")
    print(detalle.to_string(index=False))

    print("\n\nTABLERO DE CONTROL\n")
    print(resumen_ciclo(df, ciclo, polizas_esperadas).to_string(index=False))

    exceso = calcular_exceso(df, ciclo)
    if exceso.empty:
        print("\nNo hay datos suficientes para calcular el exceso.")
        return

    tarifa = tarifa_vigente(df, ciclo)
    print(f"\n\nEXCESO POR REFERENCIA (tarifa vigente del ciclo: {tarifa:g})\n")
    print(exceso.to_string(index=False))

    total = exceso["exceso"].sum()
    print("\n" + "-" * ancho)
    print(f"{'TOTAL COBRADO EN EXCESO':<40}{_moneda(total)}")
    print("-" * ancho)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def construir_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Detecta cobros duplicados y cargos no contratados."
    )
    parser.add_argument("--movimientos", type=Path, default=Path("data/movimientos.csv"))
    parser.add_argument("--catalogo", type=Path, default=Path("data/catalogo.csv"))
    parser.add_argument(
        "--dia-corte",
        type=int,
        default=21,
        help="Último día del ciclo de facturación (por defecto 21).",
    )
    parser.add_argument(
        "--ciclo",
        type=str,
        default=None,
        help="Ciclo a analizar en formato AAAA-MM. Por defecto, el más reciente.",
    )
    parser.add_argument(
        "--exportar",
        type=Path,
        default=None,
        help="Ruta opcional para guardar el detalle enriquecido en CSV.",
    )
    return parser


def main() -> None:
    args = construir_parser().parse_args()

    catalogo = cargar_catalogo(args.catalogo)
    movimientos = cargar_movimientos(args.movimientos)

    df = diagnosticar(enriquecer(movimientos, catalogo, args.dia_corte))
    ciclo = args.ciclo or df["ciclo"].max()
    polizas_esperadas = int(catalogo["tipo"].eq("SEGURO").sum())

    imprimir_reporte(df, ciclo, polizas_esperadas)

    if args.exportar:
        args.exportar.parent.mkdir(parents=True, exist_ok=True)
        df.to_csv(args.exportar, index=False)
        print(f"\nDetalle exportado a {args.exportar}")


if __name__ == "__main__":
    main()
