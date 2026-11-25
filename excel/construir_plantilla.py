"""Genera la plantilla de Excel equivalente al script de Python.

Las fórmulas se escriben con nombres en inglés y coma como separador porque
así es como el formato .xlsx las almacena internamente. Excel las traduce
automáticamente al idioma de la instalación al abrir el archivo: en una
instalación en español aparecerán como SI.ERROR, INDICE, COINCIDIR, etc.

Uso:  python excel/construir_plantilla.py
"""

from __future__ import annotations

import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

RAIZ = Path(__file__).resolve().parents[1]
SALIDA = RAIZ / "excel" / "verificador_cobros.xlsx"
DIA_CORTE = 21
FILA_INICIAL = 4
ULTIMA_FILA = 200

VERDE = PatternFill("solid", fgColor="C6EFCE")
ROJO = PatternFill("solid", fgColor="FFC7CE")
AMARILLO = PatternFill("solid", fgColor="FFEB9C")
ENCABEZADO = PatternFill("solid", fgColor="1F4E4A")
SUBTITULO = PatternFill("solid", fgColor="A9C47F")
BORDE = Border(*[Side(style="thin", color="BFBFBF")] * 4)

COLUMNAS = [
    ("Fecha", 12),
    ("Cuenta Cargada", 20),
    ("Cuenta Esperada", 20),
    ("Referencia", 22),
    ("Verificación", 13),
    ("Descripción", 34),
    ("Monto", 13),
    ("Ciclo", 11),
    ("Veces cobrada", 15),
    ("Tarifa implícita", 16),
    ("Diagnóstico", 30),
]


def _leer_csv(nombre: str) -> list[dict[str, str]]:
    with (RAIZ / "data" / nombre).open(encoding="utf-8") as archivo:
        return list(csv.DictReader(archivo))


def _estilo_encabezado(celda, relleno=ENCABEZADO, color="FFFFFF") -> None:
    celda.fill = relleno
    celda.font = Font(bold=True, color=color)
    celda.alignment = Alignment(horizontal="center", vertical="center")
    celda.border = BORDE


def construir_catalogo(wb: Workbook) -> None:
    """Hoja de catálogo: la lista blanca de servicios contratados."""
    hoja = wb.create_sheet("Catalogo")
    encabezados = ["Referencia", "Cuenta esperada", "Factor", "Tipo", "Descripción"]

    for columna, titulo in enumerate(encabezados, start=1):
        _estilo_encabezado(hoja.cell(row=1, column=columna, value=titulo))
        hoja.column_dimensions[get_column_letter(columna)].width = 24

    for fila, registro in enumerate(_leer_csv("catalogo.csv"), start=2):
        hoja.cell(row=fila, column=1, value=registro["referencia"])
        hoja.cell(row=fila, column=2, value=registro["cuenta_esperada"])
        factor = registro["factor"]
        hoja.cell(row=fila, column=3, value=float(factor) if factor else None)
        hoja.cell(row=fila, column=4, value=registro["tipo"])
        hoja.cell(row=fila, column=5, value=registro["descripcion_poliza"])

    hoja["G2"] = "Solo se agrega aquí lo que esté efectivamente contratado."
    hoja["G3"] = "Una referencia nueva debe confirmarse antes de catalogarla:"
    hoja["G4"] = "agregarla sin verificar la vuelve invisible para siempre."
    for fila in (2, 3, 4):
        hoja[f"G{fila}"].font = Font(italic=True, color="808080")


def construir_movimientos(wb: Workbook) -> None:
    """Hoja principal: movimientos, fórmulas de verificación y tablero."""
    hoja = wb.create_sheet("Movimientos", 0)

    hoja.merge_cells("A1:K1")
    hoja["A1"] = "Verificador de cobros recurrentes"
    hoja["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    hoja["A1"].fill = ENCABEZADO
    hoja["A1"].alignment = Alignment(horizontal="center", vertical="center")
    hoja.row_dimensions[1].height = 24

    hoja.merge_cells("A2:K2")
    hoja["A2"] = f"Ciclo de facturación: del día {DIA_CORTE + 1} al día {DIA_CORTE}"
    hoja["A2"].fill = SUBTITULO
    hoja["A2"].alignment = Alignment(horizontal="center")

    for columna, (titulo, ancho) in enumerate(COLUMNAS, start=1):
        _estilo_encabezado(hoja.cell(row=3, column=columna, value=titulo), SUBTITULO, "000000")
        hoja.column_dimensions[get_column_letter(columna)].width = ancho

    movimientos = _leer_csv("movimientos.csv")
    for indice, registro in enumerate(movimientos):
        fila = FILA_INICIAL + indice
        hoja.cell(row=fila, column=1, value=registro["fecha"]).number_format = "DD/MM/YYYY"
        hoja.cell(row=fila, column=2, value=registro["cuenta_cargada"])
        hoja.cell(row=fila, column=6, value=registro["descripcion"])
        hoja.cell(row=fila, column=7, value=float(registro["monto"])).number_format = "#,##0.00"
        _escribir_formulas(hoja, fila)

    _validacion_cuentas(hoja, len(movimientos))
    _formato_condicional(hoja, len(movimientos))
    _tablero(hoja)
    hoja.freeze_panes = "A4"


def _escribir_formulas(hoja, fila: int) -> None:
    """Las cinco fórmulas derivadas, una por columna calculada."""
    ref = "Catalogo!$A$2:$A$6"
    cuentas = "Catalogo!$B$2:$B$6"
    factores = "Catalogo!$C$2:$C$6"

    # Cuenta esperada: lo que dice el contrato, no lo que hizo el banco.
    hoja[f"C{fila}"] = f'=IFERROR(INDEX({cuentas},MATCH($D{fila},{ref},0)),"—")'

    # Referencia: búsqueda por subcadena, no extracción de la última palabra.
    #
    # Se usa SUMPRODUCT y no MATCH(1,...) porque esa segunda forma exige
    # evaluación matricial. openpyxl escribe las fórmulas como texto plano,
    # sin la marca que identifica una fórmula matricial, así que Excel al
    # abrir el archivo le inserta el operador de intersección implícita (@)
    # para preservar el comportamiento antiguo. Ese @ colapsa el rango a un
    # único valor y la búsqueda deja de funcionar en silencio.
    #
    # SUMPRODUCT declara sus argumentos como matrices, evalúa el rango
    # completo sin necesidad de Ctrl+Shift+Enter y sobrevive al guardado.
    primera = ref.split(":")[0]
    hoja[f"D{fila}"] = (
        f'=IFERROR(INDEX({ref},'
        f'SUMPRODUCT(ISNUMBER(SEARCH({ref},$F{fila}))*({ref}<>"")*ROW({ref}))'
        f'-ROW({primera})+1),"NO RECONOCIDA")'
    )

    # Verificación: compara lo que el banco hizo contra lo contratado.
    hoja[f"E{fila}"] = (
        f'=IF($D{fila}="NO RECONOCIDA","⚠",'
        f'IF(TRIM(UPPER($B{fila}))=TRIM(UPPER($C{fila})),"✔","✘"))'
    )
    hoja[f"E{fila}"].alignment = Alignment(horizontal="center")

    # Ciclo: se construye con YEAR y MONTH para no depender de códigos de
    # formato de fecha, que cambian de significado según el idioma de Excel.
    cierre = f"EDATE($A{fila},--(DAY($A{fila})>{DIA_CORTE}))"
    hoja[f"H{fila}"] = (
        f'=IF($A{fila}="","",YEAR({cierre})&"-"&TEXT(MONTH({cierre}),"00"))'
    )

    # Veces cobrada: el detector de duplicados.
    hoja[f"I{fila}"] = (
        f'=IF($D{fila}="NO RECONOCIDA","",'
        f'COUNTIFS($H${FILA_INICIAL}:$H${ULTIMA_FILA},$H{fila},'
        f'$D${FILA_INICIAL}:$D${ULTIMA_FILA},$D{fila}))'
    )

    # Tarifa implícita: monto entre factor. Todas deben coincidir en un ciclo.
    hoja[f"J{fila}"] = (
        f'=IFERROR($G{fila}/INDEX({factores},MATCH($D{fila},{ref},0)),"")'
    )
    hoja[f"J{fila}"].number_format = "0.##"

    # Diagnóstico con IF anidados y no con IFS: las funciones posteriores a
    # Excel 2007 deben escribirse con el prefijo _xlfn. dentro del archivo, y
    # openpyxl no lo agrega. Sin ese prefijo Excel muestra #¿NOMBRE?. IF es
    # anterior a esa frontera y funciona en cualquier versión sin prefijo.
    hoja[f"K{fila}"] = (
        f'=IF($D{fila}="NO RECONOCIDA","NO CONTRATADO",'
        f'IF($I{fila}>1,"DUPLICADO: "&$I{fila}&" cobros",'
        f'IF($E{fila}="✘","CUENTA INCORRECTA","OK")))'
    )


def _validacion_cuentas(hoja, filas: int) -> None:
    """Lista desplegable alimentada del catálogo.

    Evita que un error de tipeo en la cuenta manual genere una falsa alerta:
    si el dato manual no sale del mismo catálogo, la comparación deja de ser
    confiable.
    """
    validacion = DataValidation(
        type="list", formula1="=Catalogo!$B$2:$B$6", allow_blank=True
    )
    hoja.add_data_validation(validacion)
    validacion.add(f"B{FILA_INICIAL}:B{FILA_INICIAL + filas - 1}")


def _formato_condicional(hoja, filas: int) -> None:
    fin = FILA_INICIAL + filas - 1

    for simbolo, relleno in (("✔", VERDE), ("✘", ROJO), ("⚠", AMARILLO)):
        hoja.conditional_formatting.add(
            f"E{FILA_INICIAL}:E{fin}",
            CellIsRule(operator="equal", formula=[f'"{simbolo}"'], fill=relleno),
        )

    hoja.conditional_formatting.add(
        f"I{FILA_INICIAL}:I{fin}",
        CellIsRule(operator="greaterThan", formula=["1"], fill=ROJO),
    )
    hoja.conditional_formatting.add(
        f"K{FILA_INICIAL}:K{fin}",
        FormulaRule(formula=[f'$K{FILA_INICIAL}<>"OK"'], fill=ROJO),
    )


def _tablero(hoja) -> None:
    """Tablero de control: cinco indicadores con su valor esperado."""
    hoja.merge_cells("M2:O2")
    hoja["M2"] = "Tablero de Control"
    hoja["M2"].font = Font(bold=True, color="FFFFFF")
    hoja["M2"].fill = ENCABEZADO
    hoja["M2"].alignment = Alignment(horizontal="center")

    hoja["M3"] = "Ciclo analizado"
    hoja["N3"] = f"=H{FILA_INICIAL}"
    hoja["N3"].font = Font(bold=True)

    for columna, titulo in (("M", "Indicador"), ("N", "Valor"), ("O", "Esperado")):
        _estilo_encabezado(hoja[f"{columna}4"], SUBTITULO, "000000")
        hoja[f"{columna}4"] = titulo

    rango_ciclo = f"$H${FILA_INICIAL}:$H${ULTIMA_FILA}"
    rango_ref = f"$D${FILA_INICIAL}:$D${ULTIMA_FILA}"
    seguros = "Catalogo!$A$2:$A$5"  # fila 6 es la comisión, no una póliza

    indicadores = [
        (
            "Cobros de seguro",
            f'=SUMPRODUCT(({rango_ciclo}=$N$3)*(COUNTIF({seguros},{rango_ref})>0))',
            "=COUNTIF(Catalogo!$D$2:$D$6,\"SEGURO\")",
        ),
        (
            "Cargos sin reconocer",
            f'=COUNTIFS({rango_ciclo},$N$3,{rango_ref},"NO RECONOCIDA")',
            0,
        ),
        (
            "Total del ciclo",
            f'=SUMIFS($G${FILA_INICIAL}:$G${ULTIMA_FILA},{rango_ciclo},$N$3)',
            "—",
        ),
        # MINIFS y MAXIFS son posteriores a Excel 2007 y requieren el prefijo
        # _xlfn. dentro del archivo. Excel lo retira al mostrarlas: en la barra
        # de fórmulas se leen como MINIFS y MAXIFS normales.
        (
            "Tarifa mínima",
            f'=_xlfn.MINIFS($J${FILA_INICIAL}:$J${ULTIMA_FILA},{rango_ciclo},$N$3)',
            # Sin "=" al inicio: un texto que empieza con "=" lo interpreta
            # Excel como fórmula y devuelve #¿NOMBRE?.
            "igual a la máxima",
        ),
        (
            "Tarifa máxima",
            f'=_xlfn.MAXIFS($J${FILA_INICIAL}:$J${ULTIMA_FILA},{rango_ciclo},$N$3)',
            "igual a la mínima",
        ),
        (
            "Alertas abiertas",
            f'=COUNTIFS({rango_ciclo},$N$3,'
            f'$K${FILA_INICIAL}:$K${ULTIMA_FILA},"<>OK")',
            0,
        ),
    ]

    for indice, (titulo, formula, esperado) in enumerate(indicadores, start=5):
        hoja[f"M{indice}"] = titulo
        hoja[f"N{indice}"] = formula
        hoja[f"O{indice}"] = esperado
        hoja[f"O{indice}"].alignment = Alignment(horizontal="center")
        for columna in "MNO":
            hoja[f"{columna}{indice}"].border = BORDE

    hoja["N7"].number_format = "#,##0.00"
    for columna, ancho in (("M", 22), ("N", 14), ("O", 12)):
        hoja.column_dimensions[columna].width = ancho


def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)
    construir_movimientos(wb)
    construir_catalogo(wb)
    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    wb.save(SALIDA)
    print(f"Plantilla generada: {SALIDA.relative_to(RAIZ)}")


if __name__ == "__main__":
    main()
