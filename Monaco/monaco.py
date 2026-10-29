"""Generación de SQL de tarifas Monaco desde Excel (Python 3.14)."""

import argparse
import configparser
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
import logging
import math
from pathlib import Path
import re
import sys
import unicodedata

BASE_DIR = Path(__file__).resolve().parent


class ValidationError(ValueError):
    pass


@dataclass(frozen=True)
class Registro:
    pais: str
    destino: str
    caracteristicas: str
    precio: str

    def values(self):
        return self.pais, self.destino, self.caracteristicas, self.precio


@dataclass
class Planilla:
    hoja: str
    filas: int
    repetidos: int
    registros: list[Registro]


def texto(value, campo):
    if value is None or isinstance(value, bool):
        raise ValidationError(f"{campo}: valor vacío o inválido.")
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValidationError(f"{campo}: número inválido.")
        # Excel solo conserva 15 dígitos significativos en celdas numéricas.
        value = int(value) if value.is_integer() else value
    result = str(value).strip()
    if not result or any(ord(c) < 32 for c in result):
        raise ValidationError(f"{campo}: valor vacío o caracteres de control.")
    try:
        encoded = result.encode("cp1252")
    except UnicodeEncodeError as exc:
        raise ValidationError(f"{campo}: contiene caracteres incompatibles con varchar de la tabla.") from exc
    if len(encoded) > 255:
        raise ValidationError(f"{campo}: supera los 255 caracteres permitidos.")
    return result


def precio(value):
    # str(float) puede usar exponente para tarifas pequeñas (p. ej. 0.00001).
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and not math.isfinite(value):
            raise ValidationError("Precio: número inválido.")
        value = format(Decimal(str(value)), "f")
    raw = texto(value, "Precio")
    if not re.fullmatch(r"[0-9]+(?:[.,][0-9]+)?", raw):
        raise ValidationError("Precio: debe ser un número no negativo, sin separadores de miles ni moneda.")
    try:
        number = Decimal(raw.replace(",", "."))
    except InvalidOperation as exc:
        raise ValidationError("Precio: número inválido.") from exc
    result = format(number, "f")
    if "." in result:
        result = result.rstrip("0").rstrip(".")
    return texto(result, "Precio")


def encabezado(value):
    raw = unicodedata.normalize("NFKD", str(value or ""))
    return " ".join("".join(c for c in raw if not unicodedata.combining(c)).casefold().split())


def leer_excel(path, hoja=""):
    from openpyxl import load_workbook

    path = Path(path)
    if path.suffix.lower() != ".xlsx":
        raise ValidationError("Seleccioná un archivo .xlsx; convertí los .xls antes de cargar.")
    # data_only=False permite rechazar fórmulas sin depender de cachés de Excel.
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        if hoja and hoja not in workbook.sheetnames:
            raise ValidationError(f"No existe la hoja '{hoja}'.")
        sheet = workbook[hoja] if hoja else workbook.active
        # La fila 1 es el título; la fila 2 contiene los encabezados.
        rows = sheet.iter_rows(min_row=2)
        first = next(rows, None)
        if not first:
            raise ValidationError("Faltan los encabezados en la fila 2.")
        wanted = {"pais": "Pais", "destino": "Destino", "caracteristicas": "Caracteristicas", "precio (usd)": "Precio"}
        columns = {}
        for index, cell in enumerate(first):
            name = wanted.get(encabezado(cell.value))
            if name:
                if name in columns:
                    raise ValidationError(f"Encabezado repetido: {name}.")
                columns[name] = index
        if len(columns) != 4:
            raise ValidationError("La fila 2 debe contener Pais, Destino, Características y Precio (USD).")
        records = {}
        source_rows = {}
        errors = []
        count = duplicates = 0
        for row_number, cells in enumerate(rows, start=3):
            # Ignorar solo filas totalmente vacías, no registros incompletos.
            if all(c.value is None or (isinstance(c.value, str) and not c.value.strip()) for c in cells):
                continue
            count += 1
            try:
                values = {}
                for name, index in columns.items():
                    cell = cells[index] if index < len(cells) else None
                    if cell is not None and cell.data_type in ("f", "e"):
                        raise ValidationError(f"{name}: fórmula o error de Excel; usar valores.")
                    values[name] = cell.value if cell else None
                pais = texto(values["Pais"], "Pais")
                destino = texto(values["Destino"], "Destino")
                cost = precio(values["Precio"])
                raw = values["Caracteristicas"]
                if raw is None:
                    raise ValidationError("Caracteristicas: valor vacío.")
                parts = raw.split(",") if isinstance(raw, str) else [raw]
                for part in parts:
                    feature = texto(part, "Caracteristicas")
                    key = pais, destino, feature
                    if key in records:
                        if records[key].precio != cost:
                            raise ValidationError(f"Clave repetida con precios diferentes (también en fila {source_rows[key]}).")
                        duplicates += 1
                    else:
                        records[key] = Registro(pais, destino, feature, cost)
                        source_rows[key] = row_number
            except ValidationError as exc:
                errors.append(f"Fila {row_number}: {exc}")
        if errors:
            raise ValidationError("Planilla rechazada. No se cargó ningún dato.\n" + "\n".join(errors))
        if not records:
            raise ValidationError("La planilla no contiene registros para cargar.")
        return Planilla(sheet.title, count, duplicates, list(records.values()))
    finally:
        workbook.close()


def configuracion(path):
    config = configparser.ConfigParser(interpolation=None)
    if not config.read(path, encoding="utf-8-sig"):
        raise ValidationError(f"No se encontró la configuración: {path}")
    return config


MATCH = "t.Pais = s.Pais AND t.Destino = s.Destino AND t.Caracteristicas = s.Caracteristicas"


def literal_sql(value):
    """Literal Unicode: los apóstrofes del Excel nunca cierran la cadena SQL."""
    return "N'" + value.replace("'", "''") + "'"


def generar_sql(registros):
    if not registros:
        raise ValidationError("No hay registros para generar SQL.")
    inserts = []
    # SQL Server permite hasta 1000 filas en cada INSERT ... VALUES.
    for offset in range(0, len(registros), 500):
        values = ["(" + ", ".join(literal_sql(v) for v in r.values()) + ")"
                  for r in registros[offset:offset + 500]]
        inserts.append("INSERT INTO #MonacoCarga (Pais, Destino, Caracteristicas, Precio) VALUES\n"
                       + ",\n".join(values) + ";")
    datos = "\n\n".join(inserts)
    return f"""-- Monaco: abrir y ejecutar el archivo COMPLETO en SSMS conectado a DWHNET.
-- 0 = vista previa sin cambios; cambiar a 1 y ejecutar nuevamente para aplicar.
USE [Auxiliar];
IF @@TRANCOUNT <> 0
BEGIN
    RAISERROR(N'Ejecutar sin otra transaccion abierta.', 16, 1);
    RETURN;
END;
SET NOCOUNT ON;
SET XACT_ABORT ON;
SET LOCK_TIMEOUT 30000;
DECLARE @Aplicar bit = 0;

IF OBJECT_ID(N'tempdb..#MonacoCarga') IS NOT NULL
    THROW 50000, N'Ya existe #MonacoCarga en esta sesion. Abrir una nueva consulta.', 1;

BEGIN TRY
    BEGIN TRANSACTION;
    SELECT TOP (0) Pais, Destino, Caracteristicas, Precio
    INTO #MonacoCarga FROM dbo.PreciosMonaco;

{datos}

    IF EXISTS (SELECT 1 FROM #MonacoCarga
               GROUP BY Pais, Destino, Caracteristicas HAVING COUNT(*) > 1)
        THROW 50001, N'El Excel tiene claves repetidas segun la intercalacion SQL. Unificar esas filas.', 1;

    DECLARE @Total bigint;
    SELECT @Total = COUNT_BIG(*) FROM dbo.PreciosMonaco WITH (TABLOCKX, HOLDLOCK);
    IF EXISTS (SELECT 1 FROM dbo.PreciosMonaco t
               JOIN #MonacoCarga s ON {MATCH}
               GROUP BY t.Pais, t.Destino, t.Caracteristicas HAVING COUNT(*) > 1)
        THROW 50002, N'La tabla tiene varias filas para una clave de esta carga. Resolver duplicados.', 1;

    DECLARE @Nuevos int, @Actualizados int, @SinCambios int;
    SELECT
        @Nuevos = COALESCE(SUM(CASE WHEN t.Pais IS NULL THEN 1 ELSE 0 END), 0),
        @Actualizados = COALESCE(SUM(CASE WHEN t.Pais IS NOT NULL AND
            (t.Precio IS NULL OR t.Precio <> s.Precio) THEN 1 ELSE 0 END), 0),
        @SinCambios = COALESCE(SUM(CASE WHEN t.Pais IS NOT NULL AND
            t.Precio = s.Precio THEN 1 ELSE 0 END), 0)
    FROM #MonacoCarga s LEFT JOIN dbo.PreciosMonaco t ON {MATCH};

    IF @Aplicar = 1
    BEGIN
        UPDATE t SET Precio = s.Precio FROM dbo.PreciosMonaco t
        JOIN #MonacoCarga s ON {MATCH}
        WHERE t.Precio IS NULL OR t.Precio <> s.Precio;

        INSERT INTO dbo.PreciosMonaco (Pais, Destino, Caracteristicas, Precio)
        SELECT s.Pais, s.Destino, s.Caracteristicas, s.Precio FROM #MonacoCarga s
        WHERE NOT EXISTS (SELECT 1 FROM dbo.PreciosMonaco t WHERE {MATCH});

        DROP TABLE #MonacoCarga;
        COMMIT TRANSACTION;
    END
    ELSE
        ROLLBACK TRANSACTION;

    SELECT @Nuevos AS Nuevos, @Actualizados AS PreciosActualizados,
           @SinCambios AS SinCambios,
           CASE WHEN @Aplicar = 1 THEN N'Aplicado'
                ELSE N'Vista previa: no se modificaron datos' END AS Resultado;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    THROW;
END CATCH;
"""


def guardar_sql(path, registros):
    content = generar_sql(registros)
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    # No sobrescribir un archivo que el usuario ya haya revisado o editado.
    try:
        with path.open("x", encoding="utf-8-sig", newline="\r\n") as output:
            output.write(content)
    except FileExistsError as exc:
        raise ValidationError(f"Ya existe el archivo de salida: {path}. Elegí otro nombre.") from exc
    return path.resolve()


def seleccionar_archivo():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        try:
            return filedialog.askopenfilename(title="Seleccionar planilla Monaco", filetypes=[("Excel", "*.xlsx")])
        finally:
            root.destroy()
    except (ImportError, RuntimeError):
        return input("Ruta del archivo .xlsx: ").strip().strip('"')
    except Exception as exc:
        # Incluye TclError si no existe escritorio en el equipo.
        logging.info("Selector gráfico no disponible: %s", type(exc).__name__)
        return input("Ruta del archivo .xlsx: ").strip().strip('"')


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--archivo", type=Path, help="Planilla .xlsx; si se omite, abre el selector.")
    parser.add_argument("--config", type=Path, default=BASE_DIR / "config.ini")
    parser.add_argument("--validar", action="store_true", help="Validar Excel sin conexión ni cambios en SQL Server.")
    parser.add_argument("--salida", type=Path, help="Ruta del .sql; por defecto se guarda en salidas.")
    args = parser.parse_args(argv)
    log_dir = BASE_DIR / "logs"
    log_dir.mkdir(exist_ok=True)
    log_path = log_dir / f"monaco_{datetime.now():%Y%m%d_%H%M%S_%f}.log"
    logging.basicConfig(filename=log_path, encoding="utf-8", level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", force=True)
    try:
        path = args.archivo or seleccionar_archivo()
        if not path:
            print("Selección cancelada. No se modificó la base.")
            return 0
        config = configuracion(args.config)
        planilla = leer_excel(path, config.get("excel", "hoja", fallback="").strip())
        message = f"Hoja: {planilla.hoja} | Filas: {planilla.filas} | Registros: {len(planilla.registros)} | Repetidos idénticos omitidos: {planilla.repetidos}"
        print(message)
        logging.info(message)
        if args.validar:
            print("Excel válido. No se conectó ni modificó SQL Server.")
            return 0
        output = args.salida or BASE_DIR / "salidas" / f"monaco_{datetime.now():%Y%m%d_%H%M%S_%f}.sql"
        output = guardar_sql(output, planilla.registros)
        print(f"Archivo SQL generado: {output}")
        print("Abrilo en SSMS conectado a DWHNET y ejecutalo completo para ver el resumen.")
        print("Para aplicar, cambiá @Aplicar de 0 a 1 y ejecutalo completo nuevamente.")
        print("El programa no se conectó ni modificó la base.")
        logging.info("SQL generado: %s | Registros: %s", output, len(planilla.registros))
        return 0
    except (KeyboardInterrupt, EOFError):
        print("Operación cancelada.")
        logging.info("Operación interrumpida.")
        return 1
    except Exception as exc:
        # No registrar cadenas de conexión ni credenciales.
        print(f"ERROR: {exc}")
        logging.error("%s: %s", type(exc).__name__, exc)
        return 1
    finally:
        print(f"Registro de ejecución: {log_path}")


if __name__ == "__main__":
    sys.exit(main())
