"""Pruebas aisladas del flujo real sin requerir HANA ni llamar a la API."""
import ast
import asyncio
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock


class FinalProcessTests(unittest.TestCase):
    def setUp(self):
        source = Path(__file__).resolve().parents[1] / "app.py"
        tree = ast.parse(source.read_text(encoding="utf-8"))
        functions = []
        for node in tree.body:
            if isinstance(node, ast.FunctionDef) and node.name in {
                "ejecutar_proceso_final", "enviar_datos"
            }:
                node.decorator_list = []
                functions.append(node)
        self.events = []
        self.rows = [
            {"total": 2, "exitos": 2, "errores": 0},
            {"total": 1, "exitos": 1, "errores": 0},
        ]

        async def process(sheet, session_id, mode):
            await asyncio.sleep(0)
            self.events.append(sheet)
            return self.rows[sheet]

        class RequestException(Exception):
            pass

        class Timeout(RequestException):
            pass

        self.response = Mock(status_code=200)
        self.response.__enter__ = Mock(return_value=self.response)
        self.response.__exit__ = Mock(return_value=False)

        def call_api(*args, **kwargs):
            self.assertEqual(self.events, [0, 1])
            self.events.append("api")
            return self.response

        self.http = SimpleNamespace(
            request=Mock(side_effect=call_api),
            Timeout=Timeout, RequestException=RequestException,
        )
        self.ns = {
            "requests": self.http, "asyncio": asyncio,
            "FINAL_PROCESS_URL": "https://tlcl-processes-hub.cfapps.us10.hana.ondemand.com/tlcl-hub/tlcl13",
            "FINAL_PROCESS_TIMEOUT": 60,
            "request": SimpleNamespace(
                get_json=lambda: {"hojas": [0, 1]}, args={}
            ),
            "jsonify": lambda value: value,
            "delete_all_data": Mock(return_value={"success": True}),
            "procesar_hoja_db_async": process,
        }
        exec(compile(ast.Module(body=functions, type_ignores=[]), str(source), "exec"), self.ns)

    def test_api_runs_once_after_all_sheets(self):
        result = self.ns["enviar_datos"]()
        self.assertEqual(self.events, [0, 1, "api"])
        self.assertEqual(result["proceso_final"], {
            "success": True, "code": 200, "message": "Proceso finalizado con éxito."
        })
        self.http.request.assert_called_once_with(
            "POST", self.ns["FINAL_PROCESS_URL"], timeout=60, allow_redirects=False
        )

    def test_non_200_codes_are_errors_without_api_body(self):
        for code in [201, 204, 302, 400, 500]:
            with self.subTest(code=code):
                self.events.clear()
                self.response.status_code = code
                result = self.ns["enviar_datos"]()
                self.assertEqual(result["status"], "error")
                self.assertFalse(result["proceso_final"]["success"])
                self.assertEqual(result["proceso_final"]["code"], code)
                self.assertIn(str(code), result["proceso_final"]["message"])

    def test_network_failures_are_sanitized(self):
        for exception, code in [(self.http.Timeout, "TIMEOUT"),
                                (self.http.RequestException, "CONEXION")]:
            with self.subTest(code=code):
                self.http.request.side_effect = exception("Información interna")
                result = self.ns["enviar_datos"]()["proceso_final"]
                self.assertEqual(result["code"], code)
                self.assertNotIn("Información interna", result["message"])

    def test_incomplete_or_empty_upload_never_calls_api(self):
        for rows in [
            [{"total": 2, "exitos": 1, "errores": 1}] * 2,
            [{"total": 0, "exitos": 0, "errores": 0}] * 2,
            [{"total": 2, "exitos": 1, "errores": 0}] * 2,
        ]:
            self.rows = rows
            result = self.ns["enviar_datos"]()
            self.assertIsNone(result["proceso_final"])
        self.http.request.assert_not_called()

    def test_delete_failure_never_calls_api(self):
        self.ns["delete_all_data"].return_value = {"success": False}
        self.assertFalse(self.ns["enviar_datos"]()["success"])
        self.http.request.assert_not_called()

    def test_sheet_exception_never_calls_api(self):
        self.rows = []
        self.assertFalse(self.ns["enviar_datos"]()["success"])
        self.http.request.assert_not_called()


if __name__ == "__main__":
    unittest.main()
