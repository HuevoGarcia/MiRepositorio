import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock

from openpyxl import Workbook

from monaco import Registro, ValidationError, cargar, leer_excel, precio


class ExcelTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "tarifas.xlsx"

    def excel(self, rows, headers=None):
        book = Workbook()
        book.active.append(["FORMULARIO NUEVAS TARIFAS LDI"])
        book.active.merge_cells("A1:D1")
        book.active.append(headers or ["Pais", "Destino", "Características", "Precio (USD)"])
        for row in rows:
            book.active.append(row)
        book.save(self.path)
        book.close()
        return self.path

    def test_expansion_y_ceros_iniciales(self):
        path = self.excel([["AFGHANISTAN", "AWCC", "09370, 9371", "0,16330"]])
        result = leer_excel(path)
        self.assertEqual([r.values() for r in result.registros], [
            ("AFGHANISTAN", "AWCC", "09370", "0,1633"),
            ("AFGHANISTAN", "AWCC", "9371", "0,1633"),
        ])

    def test_precio_numerico_y_codigo_entero(self):
        result = leer_excel(self.excel([["Malta", "Malta", 356, 0.18]]))
        self.assertEqual(result.registros[0].values(), ("Malta", "Malta", "356", "0,18"))

    def test_cada_campo_es_obligatorio(self):
        for index in range(4):
            with self.subTest(campo=index):
                row = ["Malta", "Malta", "356", "0,18"]
                row[index] = None
                with self.assertRaisesRegex(ValidationError, "Fila 3"):
                    leer_excel(self.excel([row]))

    def test_segmento_vacio_rechaza_archivo(self):
        for code in ("356,", ",356", "356,,357"):
            with self.subTest(code=code), self.assertRaises(ValidationError):
                leer_excel(self.excel([["Malta", "Malta", code, "0,18"]]))

    def test_duplicados_identicos_se_omiten(self):
        row = ["Malta", "Malta", "356", "0,18"]
        result = leer_excel(self.excel([row, row]))
        self.assertEqual(len(result.registros), 1)
        self.assertEqual(result.repetidos, 1)

    def test_clave_repetida_precio_distinto_rechaza_archivo(self):
        with self.assertRaisesRegex(ValidationError, r"Fila 4:.*precios diferentes.*fila 3"):
            leer_excel(self.excel([["Malta", "Malta", "356", "0,18"], ["Malta", "Malta", "356", "0,20"]]))

    def test_misma_caracteristica_otro_destino_es_otra_clave(self):
        result = leer_excel(self.excel([["Malta", "Fijo", "356", "0,18"], ["Malta", "Movil", "356", "0,20"]]))
        self.assertEqual(len(result.registros), 2)

    def test_formulas_y_errores_rechazados(self):
        for value in ("=1/2", "#DIV/0!"):
            with self.subTest(value=value), self.assertRaisesRegex(ValidationError, "fórmula o error"):
                leer_excel(self.excel([["Malta", "Malta", "356", value]]))

    def test_limites_varchar(self):
        for country in ("a" * 256, "日本"):
            with self.subTest(country=country), self.assertRaises(ValidationError):
                leer_excel(self.excel([[country, "Malta", "356", "0,18"]]))

    def test_hoja_y_encabezados(self):
        path = self.excel([["Malta", "Malta", "356", "0,18"]], ["Pais", "Destino", "Caracteristicas", "Incorrecto"])
        with self.assertRaisesRegex(ValidationError, "fila 2"):
            leer_excel(path)
        with self.assertRaisesRegex(ValidationError, "No existe la hoja"):
            leer_excel(path, "NoExiste")

    def test_fila_1_se_ignora_aunque_contenga_otros_encabezados(self):
        path = self.excel([["Malta", "Malta", "356", "0,18"]])
        from openpyxl import load_workbook
        book = load_workbook(path)
        book.active.unmerge_cells("A1:D1")
        for index, value in enumerate(["Precio (USD)", "Caracteristicas", "Destino", "Pais"], start=1):
            book.active.cell(1, index, value)
        book.save(path)
        book.close()
        result = leer_excel(path)
        self.assertEqual(result.filas, 1)
        self.assertEqual(result.registros[0].values(), ("Malta", "Malta", "356", "0,18"))

    def test_titulo_sin_encabezados_rechazado(self):
        book = Workbook()
        book.active.append(["FORMULARIO NUEVAS TARIFAS LDI"])
        book.save(self.path)
        book.close()
        with self.assertRaisesRegex(ValidationError, "fila 2"):
            leer_excel(self.path)

    def test_filas_blancas_y_archivo_sin_registros(self):
        with self.assertRaisesRegex(ValidationError, "no contiene registros"):
            leer_excel(self.excel([[None, None, None, None]]))
        result = leer_excel(self.excel([[None] * 4, ["Malta", "Malta", "356", "0,18"]]))
        self.assertEqual(result.filas, 1)

    def test_precios_invalidos_y_precision(self):
        for value in ("-1", "1.000,25", "NaN", "Infinity", True, "USD 1", " "):
            with self.subTest(value=value), self.assertRaises(ValidationError):
                precio(value)
        self.assertEqual(precio("0,12345678901234567890"), "0,1234567890123456789")
        self.assertEqual(precio(0), "0")
        self.assertEqual(precio(0.00001), "0,00001")


class TransactionTests(unittest.TestCase):
    """Verifican control transaccional con doble de ODBC, no conexión SQL real."""

    def connection(self):
        connection = Mock()
        connection.cursor.return_value.fetchone.side_effect = [None, (12,), None, (1, 1, 1)]
        return connection

    def test_preview_revierte_y_no_ejecuta_dml_destino(self):
        connection = self.connection()
        result = cargar(connection, [Registro("Malta", "Malta", "356", "0,18")])
        self.assertEqual(result, (1, 1, 1))
        connection.rollback.assert_called_once()
        connection.commit.assert_not_called()
        commands = [call.args[0].strip() for call in connection.cursor.return_value.execute.call_args_list]
        self.assertFalse(any(sql.startswith(("UPDATE", "INSERT INTO dbo")) for sql in commands))

    def test_aplicar_confirma_y_precio_es_unico_campo_actualizado(self):
        connection = self.connection()
        cargar(connection, [Registro("Malta", "Malta", "356", "0,18")], aplicar=True)
        connection.commit.assert_called_once()
        connection.rollback.assert_not_called()
        commands = [call.args[0] for call in connection.cursor.return_value.execute.call_args_list]
        self.assertTrue(any("UPDATE t SET Precio = s.Precio" in sql for sql in commands))
        self.assertTrue(any("INSERT INTO dbo.PreciosMonaco" in sql for sql in commands))

    def test_error_de_escritura_revierte(self):
        connection = self.connection()
        def execute(sql):
            if sql.strip().startswith("INSERT INTO dbo"):
                raise RuntimeError("Error de escritura")
        connection.cursor.return_value.execute.side_effect = execute
        with self.assertRaisesRegex(RuntimeError, "Error de escritura"):
            cargar(connection, [Registro("Malta", "Malta", "356", "0,18")], aplicar=True)
        connection.rollback.assert_called_once()
        connection.commit.assert_not_called()

    def test_duplicado_sql_cancela_sin_cambios(self):
        for results in ([(1,)], [None, (12,), (1,)]):
            with self.subTest(results=results):
                connection = Mock()
                connection.cursor.return_value.fetchone.side_effect = results
                with self.assertRaisesRegex(ValidationError, "repetidas|varias filas"):
                    cargar(connection, [Registro("Malta", "Malta", "356", "0,18")], aplicar=True)
                connection.rollback.assert_called_once()
                connection.commit.assert_not_called()


if __name__ == "__main__":
    unittest.main()
