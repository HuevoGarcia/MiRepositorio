# Carga de precios Monaco

Script para Windows y **Python 3.14 de 64 bits**. Copiar esta carpeta completa a
`C:\Python314\Source\Monaco`. La conexión usa la cuenta de Windows que ejecuta
el programa, sin contraseñas en archivos.

## Instalación

1. Instalar Python 3.14 con el lanzador `py` y Tcl/Tk (selector de archivos).
2. Instalar **Microsoft ODBC Driver 18 for SQL Server de 64 bits** desde Microsoft:
   https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server
3. Ejecutar **`instalar.bat`** una vez. Crea el entorno `.venv`, instala las
   dependencias y verifica que el controlador ODBC esté disponible. Usa
   `C:\Python314\python.exe` o el lanzador `py -3.14`. Repetirlo permite reintentar
   una instalación incompleta. Necesita acceso a Internet para las dependencias.

También se puede instalar manualmente desde una consola:

```bat
cd /d C:\Python314\Source\Monaco
py -3.14 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

`config.ini` ya está configurado con el servidor **`DWHNET`**, base `Auxiliar` y
tabla fija `dbo.PreciosMonaco`, según la conexión suministrada.
Se necesitan conectividad a SQL Server y permisos SELECT, INSERT y UPDATE sobre
la tabla. La autenticación es integrada (`Trusted_Connection=yes`).
Ejecutar desde la sesión de Windows que tiene acceso al servidor, por ejemplo
`temu\BSANCHEZ`; no se fija un usuario ni se guarda una contraseña en el script.
Por defecto se cifra y verifica el certificado del servidor. Si falla la confianza,
configurar un certificado válido y su CA en Windows; no desactivar la validación
como solución automática. No hace falta instalar SQL Server en la computadora.

## Uso

Ejecutar `ejecutar.bat` o:

```bat
.venv\Scripts\python.exe monaco.py
```

Seleccionar un Excel `.xlsx`. Si el selector no está disponible, se pide la ruta
en consola. Se usa la hoja activa; para fijar otra, configurar `excel.hoja`.
La primera fila debe contener estos encabezados (en cualquier orden):

| Pais | Destino | Características | Precio (USD) |
| --- | --- | --- | --- |
| AFGHANISTAN | AFGHANISTAN MOBILE AWCC | 09370, 9371 | 0,16330 |

Se toleran diferencias de acentos, mayúsculas y espacios en los encabezados.
Cada característica separada por coma produce un registro con los otros campos
iguales. Guardar los códigos como **texto en Excel** para preservar ceros iniciales
y números de más de 15 dígitos: el script no puede recuperar información que Excel
ya haya perdido. Los ceros de un formato visual numérico tampoco son parte del valor.

Reglas de validación:

- Los cuatro campos son obligatorios. Solo se ignoran filas completamente vacías.
- Cada valor resultante debe caber en `varchar(255)` con la página de códigos
  Windows-1252 de la intercalación `Modern_Spanish_CI_AS` informada.
- No se permiten fórmulas, errores de Excel ni caracteres de control en los campos.
  Convertir fórmulas a valores antes de importar.
- El precio debe ser no negativo y puede usar coma o punto decimal, sin moneda ni
  separadores de miles. Se conserva exactamente su valor decimal, sin redondearlo,
  y se almacena como texto con coma, sin ceros decimales finales innecesarios.
- Una lista como `9370,,9371` se rechaza porque contiene una característica vacía.
- Una clave repetida con el mismo precio se procesa una sola vez. Si trae precios
  distintos, se rechaza el archivo y se indican las filas involucradas.

Tras validar **todo** el Excel, el programa consulta SQL y muestra nuevos,
actualizados y sin cambios. Escribir `SI` para aplicar; cualquier otra respuesta
cancela sin cambios. La carga vuelve a comprobar el estado de SQL al confirmar,
por lo que los totales finales pueden variar si otro proceso modificó la tabla.

La clave es **Pais + Destino + Caracteristicas**, comparada con las intercalaciones
de las columnas de SQL Server (incluyendo sus reglas de mayúsculas y espacios).
Si existe, se actualiza únicamente `Precio`; si no existe, se inserta. Un precio
que ya coincide se cuenta sin cambios. Como la columna es varchar, diferencias
de formato respecto del precio normalizado también pueden producir un UPDATE.
Los registros ajenos a la carga se conservan.

Si el Excel tiene claves equivalentes según SQL (por ejemplo, `Malta` y `MALTA`),
se rechaza para que se unifiquen explícitamente. Si la tabla ya contiene varias
filas para alguna clave de la carga, se detiene sin modificarlas: deben resolverse
esos duplicados antes de reintentar. No se cambia el esquema ni se crean índices.

Toda la escritura usa parámetros y una única transacción. Ante un error anterior
al commit se revierte el lote completo. Se bloquea la tabla durante cada consulta
de resumen/carga para impedir carreras entre importaciones; el bloqueo se libera
antes de pedir confirmación. Esto puede bloquear temporalmente otros procesos:
usar una ventana de carga adecuada si la tabla tiene mucha actividad. La espera
por bloqueos se limita a 30 segundos. Un corte de conexión justo durante el commit
puede dejar su resultado incierto; comprobar la tabla antes de reintentar.

Cada ejecución genera un archivo en `logs` con su resultado o errores. Estos
archivos pueden contener datos de filas inválidas; conservarlos de acuerdo con
las políticas de la organización. No hay contraseñas en el programa.

### Validar sin conectarse a SQL Server

```bat
.venv\Scripts\python.exe monaco.py --archivo "C:\Tarifas\precios.xlsx" --validar
```

También se admite `--config "C:\ruta\otra-config.ini"`.

## Pruebas

Desde esta carpeta:

```bat
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Las pruebas generan archivos Excel reales y verifican expansión, obligatoriedad,
precios, duplicados y errores. Las pruebas transaccionales usan un doble de ODBC;
no sustituyen la verificación con SQL Server real.

Para la prueba final, usar una instancia de prueba con una copia de la estructura
`Auxiliar.dbo.PreciosMonaco`: cargar un registro nuevo, repetir la carga (sin cambios)
y cargar la misma clave con otro precio (una actualización). Verificar también
que un Excel inválido y una cancelación no alteren la tabla. No ejecutar esta
comprobación contra producción sin una planilla autorizada.
