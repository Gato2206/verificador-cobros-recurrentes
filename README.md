# Verificador de cobros recurrentes

Detección de cobros duplicados y cargos no contratados en estados de cuenta bancarios, mediante **ingeniería inversa de la tarifa** aplicada por el emisor.

Sobre el conjunto de datos de ejemplo, el verificador identifica **9 454,00 cobrados en exceso** en un solo ciclo de facturación, incluyendo una póliza que nunca fue contratada.

---

## El problema

Un titular con varias pólizas de seguro asociadas a sus tarjetas recibe un estado de cuenta mensual con cargos que:

- usan descripciones inconsistentes entre un mes y otro para la misma póliza,
- se cobran en cuentas que no corresponden,
- varían de monto sin explicación aparente,
- y ocasionalmente se duplican.

Revisar esto a mano es inviable: las descripciones no son estables, los montos cambian todos los meses y los ciclos de facturación no coinciden con los meses calendario. El resultado práctico es que los errores no se detectan.

## El hallazgo central

Los montos no son arbitrarios. Cada cobro es el producto exacto de un **factor fijo** por una **tarifa mensual**:

| Monto cobrado | Factor | Tarifa |
| --- | --- | --- |
| 2 080,00 | 4,00 | 520 |
| 2 020,00 | 4,00 | 505 |
| 1 300,00 | 2,50 | 520 |
| 1 262,50 | 2,50 | 505 |
| 3 120,00 | 6,00 | 520 |
| 3 030,00 | 6,00 | 505 |
| 3 108,00 | 6,00 | **518** |
| 6,00 | 6,00 | **1** |

La descomposición es exacta al céntimo, sin redondeos. De ahí se derivan tres consecuencias que hacen posible toda la verificación:

1. **El factor identifica la póliza.** Permite reconstruir qué se está cobrando aunque la descripción cambie de formato.
2. **La tarifa debe ser única por ciclo.** Si en un mismo ciclo aparecen varias tarifas, hay más de un proceso de facturación operando. En el ejemplo aparecen cuatro: 520, 518, 505 y 1.
3. **Lo esperado se vuelve calculable.** Con el factor y la tarifa vigente se obtiene cuánto debió cobrarse, y la diferencia contra lo efectivamente cobrado es el exceso.

Este es el punto del proyecto: no se trata de comparar contra un valor de referencia conocido de antemano, sino de **derivar la regla de negocio desde los datos crudos**, sin acceso al sistema del emisor.

## Qué detecta

| Detección | Método |
| --- | --- |
| Cargos duplicados | Conteo por póliza dentro del ciclo de facturación real |
| Pólizas no contratadas | Catálogo como lista blanca: lo que no está, se marca |
| Cobros en la cuenta equivocada | Cuenta cargada contra cuenta contratada |
| Tarifas fuera de norma | Tarifa implícita contra el consenso del ciclo |
| Comisiones repetidas por período | Agrupación por cuenta y período declarado |

## Resultado sobre los datos de ejemplo

```
TABLERO DE CONTROL

           indicador     valor esperado bandera
    Cobros de seguro         8        4 REVISAR
Cargos sin reconocer         1        0 REVISAR
     Total del ciclo 19,596.50        -
   Tarifas distintas         4        1 REVISAR
    Alertas abiertas        10        0 REVISAR


EXCESO POR REFERENCIA (tarifa vigente del ciclo: 505)

           referencia  cobrado  esperado  exceso
          SEGCD661847   6144.0    3030.0  3114.0
          SEGPFF40118   4100.0    2020.0  2080.0
        NO RECONOCIDA   2020.0       0.0  2020.0
           BLNL220931   2562.5    1262.5  1300.0
COMISION SALDO MINIMO   1650.0     800.0   850.0
          SEGCD550294   3120.0    3030.0    90.0

TOTAL COBRADO EN EXCESO                     9,454.00
```

De once movimientos, **uno solo está correcto**.

## Uso

```bash
pip install -r requirements.txt

# Analiza el ciclo más reciente
python src/verificar_cobros.py

# Con parámetros explícitos
python src/verificar_cobros.py \
    --movimientos data/movimientos.csv \
    --catalogo data/catalogo.csv \
    --dia-corte 21 \
    --ciclo 2025-04 \
    --exportar salida/detalle.csv

# Pruebas
python -m pytest tests/ -v
```

### Formato de entrada

`data/movimientos.csv` — el estado de cuenta:

```csv
fecha,cuenta_cargada,descripcion,monto
2025-03-26,Tarjeta Oro,SEGURO PRF CR SEGPFF40118,2080.00
```

`data/catalogo.csv` — lo efectivamente contratado:

```csv
referencia,cuenta_esperada,factor,tipo,descripcion_poliza
SEGPFF40118,Tarjeta Oro,4.00,SEGURO,Seguro de saldo - Tarjeta Oro
```

El factor se obtiene dividiendo cualquier cobro conocido entre su tarifa, o del contrato si está disponible.

## Decisiones de diseño

**El ciclo de facturación no es el mes calendario.** Con corte el día 21, el ciclo va del 22 al 21 del mes siguiente. Agrupar por mes calendario es precisamente lo que oculta la duplicación: dos corridas de facturación separadas por tres semanas caen en meses distintos y cada una parece normal por separado. El agrupamiento correcto es lo que hace visible el problema.

**La referencia se busca por subcadena, no se extrae.** Los emisores no son consistentes con los separadores: `DB. SEG.BLNL220931` pega la referencia al prefijo sin espacio, mientras que `SEGURO FRAUDE BLNL220931` la separa. Extraer la última palabra falla en el primer caso; buscar cada referencia conocida dentro del texto resuelve ambos con la misma regla.

**El catálogo es una lista blanca, no un diccionario.** La tentación al ver una referencia desconocida es agregarla al catálogo para que deje de dar alerta. Hacer eso vuelve el cargo invisible de forma permanente. En los datos de ejemplo, `SEGPFF90734` es exactamente ese caso: monto idéntico a otra póliza, mismo prefijo de producto, misma descripción, sin historial previo. Catalogarla sin verificar la habría ocultado.

**La tarifa vigente se toma de la fecha de cierre del ciclo.** Es cuando corre la facturación programada. Los cargos posteados en fechas intermedias suelen venir de corridas manuales o procesos heredados, y usarlos como referencia contamina el cálculo.

**Las alertas tienen prioridad, no se acumulan.** Un cargo duplicado *y* en la cuenta equivocada se reporta como duplicado, porque es lo que tiene consecuencia económica. La verificación de cuenta queda disponible en su propia columna para quien necesite el detalle.

## Versión en Excel

El mismo análisis está implementado como plantilla de hoja de cálculo, para quien tenga que operarlo sin entorno de Python:

```bash
python excel/construir_plantilla.py   # genera excel/verificador_cobros.xlsx
```

Incluye las fórmulas de verificación, lista desplegable alimentada del catálogo, formato condicional y el tablero de control. Las fórmulas se almacenan en inglés —así es como el formato `.xlsx` las guarda internamente— y Excel las traduce al idioma de la instalación al abrir el archivo.

Dos notas para quien la use:

- La fórmula de detección de referencia es matricial. En Excel 2016 o anterior hay que confirmarla con `Ctrl + Shift + Enter`.
- La columna de ciclo se construye con `YEAR` y `MONTH` en lugar de un código de formato de fecha. El código `aaaa` significa *año* en la documentación en español pero Excel lo interpreta como *día de la semana*, y la ambigüedad produce ciclos corruptos de forma silenciosa.

## Estructura

```
verificador-cobros/
├── data/
│   ├── movimientos.csv          # estado de cuenta de ejemplo
│   └── catalogo.csv             # servicios contratados
├── src/
│   └── verificar_cobros.py      # análisis en pandas
├── excel/
│   └── construir_plantilla.py   # generador de la plantilla .xlsx
├── tests/
│   └── test_verificar_cobros.py # 19 pruebas
└── requirements.txt
```

## Sobre los datos

Todos los datos de este repositorio son **sintéticos**. Las referencias, montos, cuentas y fechas fueron generados para reproducir la estructura del caso real —incluidas sus anomalías— sin exponer información financiera de nadie.

El caso que originó el proyecto fue real: un ciclo de facturación con nueve cobros de seguro sobre cuatro pólizas contratadas, detectado y cuantificado con este método.

## Licencia

MIT
