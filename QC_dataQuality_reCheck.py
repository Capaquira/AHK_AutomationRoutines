import pandas as pd
import numpy as np
from pathlib import Path


def seleccionar_archivo_excel():
    """Permite seleccionar un archivo Excel desde carpetas comunes."""

    carpetas = {
        1: Path.home() / "Desktop",
        2: Path.home() / "Documents",
        3: Path.home() / "Downloads"
    }

    print("\nCARPETAS DISPONIBLES:")
    for k, v in carpetas.items():
        print(f"{k}. {v}")

    opcion = int(input("\nSeleccione una carpeta: "))
    carpeta = carpetas.get(opcion)

    if not carpeta or not carpeta.exists():
        raise Exception("Carpeta inválida.")

    archivos = list(carpeta.glob("*.xlsx")) + list(carpeta.glob("*.xls"))

    if not archivos:
        raise Exception("No se encontraron archivos Excel.")

    archivos.sort(
        key=lambda x: x.stat().st_mtime,
        reverse=True
    )

    print("\nARCHIVOS ENCONTRADOS:")
    for i, archivo in enumerate(archivos[:20], start=1):
        print(f"{i}. {archivo.name}")

    seleccion = int(input("\nSeleccione un archivo: "))

    if seleccion < 1 or seleccion > len(archivos):
        raise Exception("Selección inválida.")

    return archivos[seleccion - 1]


def seleccionar_hoja(archivo_excel):
    """Permite seleccionar una hoja del Excel."""

    excel_file = pd.ExcelFile(archivo_excel)

    print("\nHOJAS DISPONIBLES:")
    for i, sheet in enumerate(excel_file.sheet_names, start=1):
        print(f"{i}. {sheet}")

    opcion = int(input("\nSeleccione una hoja: "))

    if opcion < 1 or opcion > len(excel_file.sheet_names):
        raise Exception("Selección inválida.")

    return excel_file.sheet_names[opcion - 1]


def detectar_encabezados(
        archivo_excel,
        hoja_seleccionada):
    """
    Muestra las primeras filas para que el usuario
    determine dónde están los encabezados reales.
    """

    preview = pd.read_excel(
        archivo_excel,
        sheet_name=hoja_seleccionada,
        header=None,
        nrows=20
    )

    print("\n" + "=" * 100)
    print("VISTA PREVIA DE LAS PRIMERAS 20 FILAS")
    print("=" * 100)

    print(preview)

    while True:

        try:

            header_row = int(
                input(
                    "\n¿En qué fila están los encabezados reales?\n"
                    "Ingrese 0 si no existen encabezados.\n"
                    "Ejemplos: 1, 2, 5, 10\n"
                    "\nFila: "
                )
            )

            if header_row < 0:

                print(
                    "\nLa fila no puede ser negativa."
                )

                continue

            break

        except ValueError:

            print(
                "\nIngrese un número válido."
            )

    return header_row


def cargar_dataframe(
        archivo_excel,
        hoja_seleccionada,
        header_row):
    """
    Carga el dataframe usando la fila indicada como encabezado.

    Si el usuario ingresa 0,
    se asume que el archivo no tiene encabezados.
    """

    if header_row == 0:

        print(
            "\nADVERTENCIA: "
            "Se cargará el archivo SIN encabezados."
        )

        df = pd.read_excel(
            archivo_excel,
            sheet_name=hoja_seleccionada,
            header=None
        )

        # Renombrar columnas automáticamente
        df.columns = [
            f"Column_{i+1}"
            for i in range(len(df.columns))
        ]

    else:

        df = pd.read_excel(
            archivo_excel,
            sheet_name=hoja_seleccionada,
            header=header_row - 1
        )

    return df


def dataset_summary(df):

    print("\n" + "=" * 100)
    print("RESUMEN GENERAL DEL DATASET")
    print("=" * 100)

    print(f"Filas: {len(df):,}")
    print(f"Columnas: {len(df.columns)}")

    print("\nCOLUMNAS DETECTADAS:\n")

    for col in df.columns:

        nulls = df[col].isna().sum()

        print(
            f"{str(col):40} | "
            f"{str(df[col].dtype):15} | "
            f"Nulos: {nulls}"
        )


def detect_numeric_as_text(series):

    if series.dtype != "object":
        return False

    converted = pd.to_numeric(
        series,
        errors="coerce"
    )

    conversion_rate = converted.notna().mean()

    return conversion_rate > 0.90


def analyze_column(df, col):

    series = df[col]

    print("\n")
    print("=" * 100)
    print(f"COLUMNA: {col}")
    print("=" * 100)

    total = len(series)
    nulls = series.isna().sum()
    unique = series.nunique(dropna=True)

    print(f"Tipo detectado: {series.dtype}")
    print(f"Total registros: {total}")
    print(f"Valores nulos: {nulls}")
    print(f"% Nulos: {round(nulls / total * 100, 2)}")
    print(f"Valores únicos: {unique}")

    duplicated_count = series.duplicated().sum()
    print(f"Duplicados: {duplicated_count}")

    # --------------------------
    # NUMERICAS
    # --------------------------

    if pd.api.types.is_numeric_dtype(series):

        print("\nESTADÍSTICAS NUMÉRICAS")

        print(f"Mínimo: {series.min()}")
        print(f"Máximo: {series.max()}")
        print(f"Media: {series.mean()}")
        print(f"Mediana: {series.median()}")
        print(f"Desviación estándar: {series.std()}")
        print(f"Varianza: {series.var()}")

        print(f"Percentil 25%: {series.quantile(0.25)}")
        print(f"Percentil 50%: {series.quantile(0.50)}")
        print(f"Percentil 75%: {series.quantile(0.75)}")

        try:
            print(f"Asimetría: {series.skew()}")
            print(f"Curtosis: {series.kurtosis()}")
        except:
            pass

        # OUTLIERS

        q1 = series.quantile(0.25)
        q3 = series.quantile(0.75)

        iqr = q3 - q1

        lower_bound = q1 - (1.5 * iqr)
        upper_bound = q3 + (1.5 * iqr)

        outliers = series[
            (series < lower_bound)
            | (series > upper_bound)
        ]

        print(f"Outliers detectados: {len(outliers)}")

        if len(outliers) > 0:

            print("\nPrimeros outliers:")

            print(
                outliers
                .head(10)
                .to_list()
            )

    # --------------------------
    # TEXTO / CATEGORICAS
    # --------------------------

    else:

        if detect_numeric_as_text(series):

            print(
                "\n⚠ ADVERTENCIA: "
                "La columna parece numérica pero fue "
                "importada como texto."
            )

        print("\nTOP 10 VALORES MÁS FRECUENTES:")

        print(
            series
            .astype(str)
            .value_counts(dropna=False)
            .head(10)
        )

    # --------------------------
    # EJEMPLOS DE DATOS
    # --------------------------

    print("\nPRIMEROS VALORES:")

    print(
        series
        .dropna()
        .head(10)
        .to_list()
    )


def analyze_dataframe(df):

    dataset_summary(df)

    print(
        "\n\nINICIANDO ANÁLISIS DETALLADO "
        "DE TODAS LAS COLUMNAS..."
    )

    for col in df.columns:

        try:
            analyze_column(df, col)

        except Exception as e:

            print(
                f"\nError analizando "
                f"'{col}': {e}"
            )


# =====================================================
# PROGRAMA PRINCIPAL
# =====================================================

try:

    archivo_excel = seleccionar_archivo_excel()

    print(f"\nArchivo seleccionado:\n{archivo_excel}")

    hoja_seleccionada = seleccionar_hoja(
        archivo_excel
    )

    print(
        f"\nHoja seleccionada: "
        f"{hoja_seleccionada}"
    )

    header_row = detectar_encabezados(
        archivo_excel,
        hoja_seleccionada
    )

    df = cargar_dataframe(
        archivo_excel,
        hoja_seleccionada,
        header_row
    )

    print("\n")
    print("=" * 100)
    print("DATAFRAME CARGADO")
    print("=" * 100)

    print(df.head())

    analyze_dataframe(df)

except Exception as e:

    print(f"\nERROR: {e}")
