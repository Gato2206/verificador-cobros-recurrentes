"""Pruebas del verificador de cobros.

Ejecutar con:  python -m pytest tests/ -v
"""

from pathlib import Path
import sys

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from verificar_cobros import (  # noqa: E402
    ETIQUETA_SIN_REFERENCIA,
    asignar_ciclo,
    calcular_exceso,
    cargar_catalogo,
    cargar_movimientos,
    detectar_referencia,
    diagnosticar,
    enriquecer,
    tarifa_vigente,
)

RAIZ = Path(__file__).resolve().parents[1]
DIA_CORTE = 21


@pytest.fixture(scope="module")
def datos() -> pd.DataFrame:
    catalogo = cargar_catalogo(RAIZ / "data" / "catalogo.csv")
    movimientos = cargar_movimientos(RAIZ / "data" / "movimientos.csv")
    return diagnosticar(enriquecer(movimientos, catalogo, DIA_CORTE))


# --------------------------------------------------------------------------- #
# Detección de referencias
# --------------------------------------------------------------------------- #

@pytest.mark.parametrize(
    "descripcion,esperado",
    [
        ("SEGURO PRF CR SEGPFF40118", "SEGPFF40118"),
        ("DB. SEG.BLNL220931", "BLNL220931"),          # sin espacio separador
        ("seguro fraude blnl220931", "BLNL220931"),    # minúsculas
        ("COMISION SALDO MINIMO/P/ MAR", "COMISION SALDO MINIMO"),
        ("SEGURO PRF PLAN SEGPFF90734", ETIQUETA_SIN_REFERENCIA),
    ],
)
def test_detectar_referencia(descripcion, esperado):
    referencias = pd.Series(
        [
            "SEGPFF40118",
            "SEGCD550294",
            "SEGCD661847",
            "BLNL220931",
            "COMISION SALDO MINIMO",
        ]
    )
    assert detectar_referencia(descripcion, referencias) == esperado


# --------------------------------------------------------------------------- #
# Ciclos de facturación
# --------------------------------------------------------------------------- #

@pytest.mark.parametrize(
    "fecha,esperado",
    [
        ("2025-03-21", "2025-03"),   # día de corte: cierra su propio ciclo
        ("2025-03-22", "2025-04"),   # día siguiente: abre el ciclo que sigue
        ("2025-03-26", "2025-04"),
        ("2025-04-21", "2025-04"),
        ("2025-12-22", "2026-01"),   # cruce de año
    ],
)
def test_asignar_ciclo(fecha, esperado):
    serie = pd.Series(pd.to_datetime([fecha]))
    assert asignar_ciclo(serie, DIA_CORTE).iloc[0] == esperado


def test_cargos_de_dos_meses_caen_en_un_solo_ciclo(datos):
    """El hallazgo central: marzo y abril pertenecen al mismo ciclo."""
    marzo = datos[datos["fecha"] == "2025-03-26"]["ciclo"].iloc[0]
    abril = datos[datos["fecha"] == "2025-04-21"]["ciclo"].iloc[0]
    assert marzo == abril == "2025-04"


# --------------------------------------------------------------------------- #
# Detección de anomalías
# --------------------------------------------------------------------------- #

def test_detecta_poliza_no_contratada(datos):
    no_contratados = datos[datos["diagnostico"] == "NO CONTRATADO"]
    assert len(no_contratados) == 1
    assert "SEGPFF90734" in no_contratados["descripcion"].iloc[0]


def test_una_sola_poliza_sin_anomalias(datos):
    correctos = datos[datos["diagnostico"] == "OK"]
    assert len(correctos) == 1
    assert correctos["referencia"].iloc[0] == "SEGCD550294"


def test_tarifa_implicita_revela_cargos_fuera_de_norma(datos):
    """Las tarifas correctas son 520 y 505; 518 y 1 son las anomalías."""
    tarifas = set(datos["tarifa_implicita"].dropna().unique())
    assert tarifas == {520.0, 518.0, 505.0, 1.0}


def test_tarifa_vigente_es_la_del_cierre(datos):
    assert tarifa_vigente(datos, "2025-04") == 505.0


def test_conteo_de_duplicados(datos):
    conteos = (
        datos[datos["ciclo"] == "2025-04"]
        .drop_duplicates("referencia")
        .set_index("referencia")["veces_cobrada"]
        .to_dict()
    )
    assert conteos["SEGCD661847"] == 3
    assert conteos["SEGPFF40118"] == 2
    assert conteos["BLNL220931"] == 2
    assert conteos["SEGCD550294"] == 1


# --------------------------------------------------------------------------- #
# Cuantificación
# --------------------------------------------------------------------------- #

def test_exceso_total(datos):
    exceso = calcular_exceso(datos, "2025-04")
    assert exceso["exceso"].sum() == pytest.approx(9454.00)


def test_la_poliza_correcta_no_genera_exceso_material(datos):
    exceso = calcular_exceso(datos, "2025-04")
    fila = exceso[exceso["referencia"] == "SEGCD550294"].iloc[0]
    # Solo la diferencia por el cambio de tarifa (520 -> 505), no un duplicado.
    assert fila["exceso"] == pytest.approx(90.00)


def test_catalogo_rechaza_referencias_repetidas(tmp_path):
    ruta = tmp_path / "catalogo.csv"
    ruta.write_text(
        "referencia,cuenta_esperada,factor,tipo\n"
        "AAA111,Tarjeta Oro,4.00,SEGURO\n"
        "AAA111,Tarjeta Plus,6.00,SEGURO\n",
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="repetidas"):
        cargar_catalogo(ruta)
