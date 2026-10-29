# Generador SQL de precios Monaco

Para Windows y **Python 3.14 de 64 bits**. Copiar la carpeta completa a
`C:\Python314\Source\Monaco`.

Valida un Excel y genera un archivo `.sql`. El usuario lo revisa y ejecuta en
SQL Server Management Studio (SSMS), conectado a **DWHNET**, sobre
**Auxiliar.dbo.PreciosMonaco**. Python no se conecta a la base: no necesita
pyodbc, controlador ODBC ni configurar certificados de SQL Server. SSMS usa la
conexión y los permisos que ya tiene el usuario.

## Instalación y uso

1. Tener Python 3.14 con Tcl/Tk para el selector de archivos.
2. Ejecutar **`instalar.bat`** una vez: prepara `Monaco\.venv` e instala `openpyxl`.
   Usa `C:\Python314\python.exe` o `py -3.14` y necesita Internet para descargar
   dependencias. Si ya estaba instalado pyodbc, se puede conservar: no se utiliza.
3. Ejecutar **`ejecutar.bat`** y seleccionar un `.xlsx`.
4. Si es válido, se genera `salidas\monaco_FECHA_HORA.sql` y se muestra su ruta.

Instalación manual equivalente:

```bat
cd /d C:\Python314\Source\Monaco
py -3.14 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe monaco.py
```

## Ejecutar el SQL

1. Conectarse a `DWHNET` desde SSMS con el usuario habitual de Windows.
2. Abrir el `.sql` generado y revisar su contenido.
3. Ejecutar el archivo **completo**, sin seleccionar fragmentos. Por defecto tiene
   `DECLARE @Aplicar bit = 0;`: muestra los nuevos, actualizados y sin cambios,
   sin escribir en la tabla destino.
4. Para confirmar, cambiar a **`DECLARE @Aplicar bit = 1;`** y ejecutar nuevamente
   el archivo completo. El resultado indicará `Aplicado` y sus totales.

El archivo contiene `USE [Auxiliar]`. Se requieren permisos SELECT, INSERT y
UPDATE sobre `dbo.PreciosMonaco`. La decisión entre insertar y actualizar se
hace al ejecutar el SQL, consultando la base en ese momento; Python no puede
anticipar los totales sin conectarse. Los totales pueden cambiar entre la vista
previa y la aplicación.

La clave es **Pais + Destino + Caracteristicas**. Si existe, se actualiza solo el
precio cuando difiere. Si no existe, se inserta. No se borran registros. Repetir
la carga deja sin cambios los datos que ya coinciden. Los precios se guardan como
texto con **punto decimal**: `0,16330` pasa a `0.1633`. Diferencias de formato en
los precios existentes también pueden producir un UPDATE.

El SQL utiliza las intercalaciones de las columnas de destino y rechaza claves
repetidas según esas reglas, tanto en el Excel como en las filas de destino
correspondientes a la carga. Escapa los apóstrofes de los datos y usa lotes de
hasta 500 filas para respetar el límite de SQL Server de 1000 filas por INSERT VALUES.

La carga es transaccional y revierte ante errores. En vista previa solo escribe
en una tabla temporal que desaparece al revertir. No cambia el esquema permanente.
Bloquea temporalmente la tabla durante cada vista previa/carga para evitar
carreras; usar una ventana adecuada si tiene mucha actividad. La espera por
bloqueos se limita a 30 segundos. Ejecutar sin otras transacciones abiertas.
Un corte de conexión justo durante el commit puede dejar el resultado incierto:
comprobar la tabla antes de reintentar.

## Formato del Excel

Se usa la hoja activa o la indicada en `config.ini`, campo `excel.hoja`.
La fila 1 se ignora (título). La **fila 2** contiene los encabezados y los **datos
comienzan en la fila 3**. Los encabezados pueden estar en cualquier orden:

| Pais | Destino | Características | Precio (USD) |
| --- | --- | --- | --- |
| AFGHANISTAN | AFGHANISTAN MOBILE AWCC | 09370, 9371 | 0,16330 |

- Cada característica separada por coma genera un registro independiente.
- Los cuatro campos son obligatorios; se ignoran solo filas totalmente vacías.
- Se rechazan fórmulas, errores de Excel, caracteres de control y valores que no
  caben en `varchar(255)` con Windows-1252 de `Modern_Spanish_CI_AS`.
- El precio acepta coma o punto decimal y debe ser no negativo, sin moneda ni
  separadores de miles. Se conserva el valor decimal exacto y se quitan ceros
  decimales finales innecesarios.
- Una lista como `9370,,9371` se rechaza por tener una característica vacía.
- Las claves idénticas con el mismo precio se procesan una vez. Con precios
  distintos se rechaza la planilla y se informan las filas involucradas.
- Guardar códigos como **texto en Excel** para conservar ceros iniciales y más de
  15 dígitos. No se recuperan datos perdidos por Excel ni ceros de formato visual.

Si una fila es inválida, no se genera SQL. Los errores indican su fila real.

## Opciones y pruebas

```bat
.venv\Scripts\python.exe monaco.py --archivo "C:\Tarifas\precios.xlsx"
.venv\Scripts\python.exe monaco.py --archivo "C:\Tarifas\precios.xlsx" --salida "C:\Tarifas\carga.sql"
.venv\Scripts\python.exe monaco.py --archivo "C:\Tarifas\precios.xlsx" --validar
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

`--validar` solo valida y no genera SQL; `--config` elige otro INI. El archivo de
salida usa UTF-8 con BOM y no sobrescribe archivos existentes. Las carpetas
`salidas` y `logs` son locales y se excluyen de Git. Los SQL contienen los datos
del Excel: conservarlos y compartirlos según las políticas de la organización.

Las pruebas cubren Excel reales, expansión de características, validaciones,
escape de apóstrofes, precios con punto, lotes de 6018 registros e interfaz de
consola. No ejecutan SQL Server: comprobar el SQL en una base de prueba con SSMS.
