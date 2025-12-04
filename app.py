from flask import Flask, render_template, request, jsonify, flash, redirect, url_for
from flask_socketio import SocketIO
import os
from werkzeug.utils import secure_filename
import openpyxl
from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP
import math
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from concurrent.futures import ThreadPoolExecutor, as_completed
import asyncio
import re
from db import (
    load_env_from_dotenv,
    get_hana_connection,
    insert_temp_rep4cfe,
    upsert_temp_rep4cfe,
    truncate_temp_rep4cfe,
)
import socket


# Configuración DNS para SAP BTP - SOLO UNA VEZ
def configure_dns_for_sap_btp():
    if "VCAP_APPLICATION" in os.environ:
        print("Configurando DNS para SAP BTP...")
        original_getaddrinfo = socket.getaddrinfo

        def patched_getaddrinfo(host, port, family=0, type=0, proto=0, flags=0):
            try:
                return original_getaddrinfo(host, port, family, type, proto, flags)
            except socket.gaierror as e:
                print(f"DNS resolution failed for {host}, trying IP fallback...")
                if host == "telcl-prd-db-cap-telcl-srv.cfapps.us10.hana.ondemand.com":
                    return [
                        (
                            socket.AF_INET,
                            socket.SOCK_STREAM,
                            6,
                            "",
                            ("52.23.1.211", port),
                        )
                    ]
                raise e

        socket.getaddrinfo = patched_getaddrinfo


# Llamar configuración DNS UNA SOLA VEZ
configure_dns_for_sap_btp()
load_env_from_dotenv()

app = Flask(__name__)
app.secret_key = "Hitss_REP4_Flask_2025"
socketio = SocketIO(app, cors_allowed_origins="*")

# Configuración de la aplicación
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

# Constantes
BATCH_SIZE = 50
MAX_CONCURRENCY = 500

HTTP_TIMEOUT = float(os.getenv("HTTP_TIMEOUT", "30"))
if not os.path.exists(app.config["UPLOAD_FOLDER"]):
    os.makedirs(app.config["UPLOAD_FOLDER"])


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def encontrar_fila_encabezados(sheet):
    """Encuentra la fila donde comienzan los encabezados relevantes"""
    encabezados_esperados = [
        "DIVISION",
        "RPU",
        "NOMBRE",
        "DIRECCION",
        "POBLACION",
        "TARIFA",
        "DESDE",
        "HASTA",
        "CONSUMO",
        "DEMANDA",
        "REACTIVOS",
        "FACTOR POTENCIA",
        "FACTOR CARGA",
        "ENERGIA",
        "IVA",
        "DAP",
        "CARGOS Y DEPOSITOS",
        "CREDITOS Y REDONDEOS",
        "TOTAL",
        "FORMULA VALIDACION",
        "DIFERENCIA",
    ]

    # Buscar en las primeras 30 filas
    for row_idx in range(1, 31):
        for col_idx in range(1, 25):  # Buscar en las primeras 25 columnas
            cell_value = sheet.cell(row=row_idx, column=col_idx).value

            # Verificar si encontramos un encabezado relevante
            if cell_value and any(
                encabezado in str(cell_value).upper()
                for encabezado in encabezados_esperados
            ):
                # Buscar la fila completa que coincida con el patrón
                row_values = [
                    sheet.cell(row=row_idx, column=c).value for c in range(1, 25)
                ]
                row_text = " ".join([str(v) for v in row_values if v])

                # Verificar si tenemos varios encabezados coincidentes
                coincidencias = sum(
                    1 for enc in encabezados_esperados if enc in row_text.upper()
                )
                if coincidencias >= 3:  # Si al menos 3 encabezados coinciden
                    return row_idx

    # Si no encontramos encabezados específicos, buscar la primera fila con muchos valores
    for row_idx in range(1, 31):
        row_values = [sheet.cell(row=row_idx, column=c).value for c in range(1, 25)]
        valores_no_vacios = sum(1 for v in row_values if v and str(v).strip())

        if (
            valores_no_vacios >= 5
        ):  # Si tiene al menos 5 valores, probablemente es la fila de encabezados
            return row_idx

    return 1  # Si no se encuentra, empezar desde la primera fila


def procesar_excel(filepath, fecha_facturacion):
    try:
        workbook = openpyxl.load_workbook(filepath, data_only=True)
        sheet_names = workbook.sheetnames
        num_sheets = len(sheet_names)

        hojas_procesadas = []
        # separar mes/anio (fecha viene como "MM/YYYY")
        mes, anio = fecha_facturacion.split("/")

        for sheet_name in sheet_names:
            sheet = workbook[sheet_name]
            fila_inicio = encontrar_fila_encabezados(sheet)

            # Detectar última columna con datos reales en la fila de encabezados
            last_col = 0
            for col in range(1, sheet.max_column + 1):
                cell_value = sheet.cell(row=fila_inicio, column=col).value
                if cell_value is not None and str(cell_value).strip() != "":
                    last_col = col
            # fallback por si no se detectó nada
            if last_col == 0:
                last_col = sheet.max_column

            # Encabezados reales (solo hasta last_col)
            encabezados = []
            for col in range(1, last_col + 1):
                cell_value = sheet.cell(row=fila_inicio, column=col).value
                encabezados.append(
                    cell_value if cell_value is not None else f"Columna {col}"
                )

            # Agregar columnas extra MES y AÑO y TIPO IVA al final de encabezados (para UI)
            encabezados.extend(["MES", "AÑO", "TIPO IVA"])

            # Índices de interés (basados en encabezados)
            try:
                idx_division = encabezados.index("DIVISION")
            except ValueError:
                idx_division = None
            try:
                idx_tarifa = encabezados.index("TARIFA")
            except ValueError:
                idx_tarifa = None

            # Recolectar filas válidas (no vacías, y sin "SUBTOTAL" ni "TOTAL")
            filas_validas = []
            for row in range(fila_inicio + 1, sheet.max_row + 1):
                # leer solo hasta last_col
                row_values = [
                    sheet.cell(row=row, column=col).value
                    for col in range(1, last_col + 1)
                ]

                # --- SALTAR fila completamente vacía ---
                if all(v is None or str(v).strip() == "" for v in row_values):
                    continue

                # --- VERIFICAR SI ES SUBTOTAL O TOTAL AL PRINCIPIO ---
                # (antes de procesar toda la fila para mayor eficiencia)
                if idx_division is not None:
                    valor_division = sheet.cell(row=row, column=idx_division + 1).value
                    if valor_division and isinstance(valor_division, str):
                        valor_normalizado = valor_division.strip().upper()
                        # Verificar si comienza con SUBTOTAL o TOTAL
                        if valor_normalizado.startswith(
                            "SUBTOTAL"
                        ) or valor_normalizado.startswith("TOTAL"):
                            continue  # Saltar esta fila
                # --- FIN DE VERIFICACIÓN ---

                # Normalizar/formatar cada celda según corresponda
                fila_datos = []
                for idx_col, cell_value in enumerate(row_values, start=1):
                    # idx_tarifa es índice 0-based en encabezados → comparamos con idx_col-1
                    if idx_tarifa is not None and (idx_col - 1) == idx_tarifa:
                        fila_datos.append(normalize_tarifa(cell_value))
                    else:
                        fila_datos.append(formatear_valor(cell_value))

                # 🔎 recortar a los 21 campos originales si hay más
                fila_datos = fila_datos[:21]

                # Añadir MES, AÑO, HOJA (TIPO IVA lo pones según necesites)
                fila_datos.append(str(mes))
                fila_datos.append(str(anio))
                fila_datos.append(sheet_name)

                filas_validas.append(fila_datos)

            # Totales reales y filas a mostrar (ej. mostrar max 40)
            total_filas_reales = len(filas_validas)

            # Mostrar máximo 40 filas, pero guardar TODAS para enviar
            datos_mostrar = filas_validas[:40]
            datos_completos = filas_validas

            hojas_procesadas.append(
                {
                    "nombre": sheet_name,
                    "encabezados": encabezados,
                    "datos": datos_completos,  # 👈 aquí todas las filas
                    "datos_preview": datos_mostrar,  # 👈 aquí solo las que muestras
                    "fila_inicio": fila_inicio,
                    "total_filas": total_filas_reales,
                    "filas_eliminadas": (sheet.max_row - fila_inicio)
                    - total_filas_reales,
                }
            )

        return {
            "num_hojas": num_sheets,
            "nombres_hojas": sheet_names,
            "hojas": hojas_procesadas,
            "fecha_facturacion": fecha_facturacion,
            "nombre_archivo": os.path.basename(filepath),
        }

    except Exception as e:
        raise Exception(f"Error al procesar el archivo Excel: {str(e)}")


def formatear_fecha(valor):
    if isinstance(valor, (datetime, date)):
        return valor.strftime("%Y-%m-%d")  # solo fecha sin hora
    return valor


def formatear_valor(valor):
    # Fechas → se manejan aparte
    if isinstance(valor, (datetime, date)):
        return valor.strftime("%Y-%m-%d")

    # Números → los pasamos a Decimal para evitar flotantes
    if isinstance(valor, (int, float)):
        dec = Decimal(str(valor)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return f"{dec:,}"  # añade comas como separadores de miles

    return valor


def normalize_tarifa(tarifa_value):
    if tarifa_value is None:
        return "00"

    str_tarifa = str(tarifa_value).strip()

    # Caso 1: vacío
    if str_tarifa == "":
        return "00"

    # Caso 2: 1 dígito numérico
    if len(str_tarifa) == 1 and str_tarifa.isdigit():
        return f"0{str_tarifa}"

    # Caso 3: ya tiene 2 caracteres
    if len(str_tarifa) == 2:
        return str_tarifa

    # Caso 4: más de 2 caracteres → tomar primeros 2
    if len(str_tarifa) > 2:
        return str_tarifa[:2]

    # Caso por defecto
    return "00"


def a_decimal(valor):
    """Convierte cualquier valor a Decimal, limpiando comas si existen"""
    return Decimal(str(valor).replace(",", ""))


def mapear_registro(fila):
    # print("FILA: ", fila)
    # print("Dato en fila[21]:", fila[21])  # imprime solo el dato de la posición 22
    # print("Dato en fila[22]:", fila[22])  # imprime solo el dato de la posición 22
    # print("Dato en fila[23]:", fila[23])  # imprime solo el dato de la posición 22
    return {
        "DIVISION": str(fila[0]),
        "RPU": str(fila[1]),
        "NAME": str(fila[2]),
        "ADDRESS": str(fila[3]),
        "POPULATION": str(fila[4]),
        "FARE": str(fila[5]),
        "FROMDATE": str(fila[6]),
        "TODATE": str(fila[7]),
        "BILLDATE": f"{fila[22]}-{str(fila[21]).zfill(2)}-01",
        "CONSUMPTION": float(a_decimal(fila[8])),
        "DEMAND": float(a_decimal(fila[9])),
        "REACTIVEPOWER": float(a_decimal(fila[10])),
        "POWERFACTOR": float(a_decimal(fila[11])),
        "LOADFACTOR": float(a_decimal(fila[12])),
        "ENERGY": float(a_decimal(fila[13])),
        "IVA": float(a_decimal(fila[14])),
        "DAP": float(a_decimal(fila[15])),
        "CHARGES": float(a_decimal(fila[16])),
        "CREDITS": float(a_decimal(fila[17])),
        "TOTAL": float(a_decimal(fila[18])),
        "VALIDATION": float(a_decimal(fila[19])),
        "DIFFERENCE": float(a_decimal(fila[20])),
        "IVATYPE": fila[23] if len(fila) > 23 else "",
    }


async def procesar_hoja_db_async(hoja, session_id, modo):
    registros = hoja.get("datos", [])
    hoja_nombre = hoja.get("nombre")
    total = len(registros)
    errores = 0
    exitos = 0
    registros_procesados = [0]
    conn = get_hana_connection()
    try:
        for i in range(0, total, BATCH_SIZE):
            batch = registros[i : i + BATCH_SIZE]
            entities = [mapear_registro(r) for r in batch]
            result = None
            if modo == "upsert":
                result = upsert_temp_rep4cfe(conn, entities)
            else:
                result = insert_temp_rep4cfe(conn, entities)
            processed = (
                result.get("updated", 0)
                + result.get("inserted", 0)
                + result.get("failed", 0)
            )
            registros_procesados[0] += processed
            exitos += result.get("updated", 0) + result.get("inserted", 0)
            errores += result.get("failed", 0)
            if result.get("errors"):
                print(
                    f"Errores en lote {i//BATCH_SIZE+1}: {len(result['errors'])} primeros: "
                )
                for e in result["errors"][:5]:
                    print(e)
            print(
                f"Lote {i//BATCH_SIZE+1} hoja {hoja_nombre}: ok={exitos}, errores={errores}, procesados={registros_procesados[0]}/{total}"
            )
            progress = (registros_procesados[0] / total) * 100 if total else 100
            socketio.emit(
                "progress_update",
                {
                    "current": registros_procesados[0],
                    "total": total,
                    "progress": round(progress, 2),
                },
                room=session_id,
            )
    finally:
        conn.close()
    return {"hoja": hoja_nombre, "total": total, "exitos": exitos, "errores": errores}


# ***************************************************************************************************************

 


@app.route("/enviar_datos", methods=["POST"])
def enviar_datos():
    try:
        data = request.get_json()
        hojas = data.get("hojas", [])
        session_id = request.args.get("session_id")  # Obtener session_id
        modo = request.args.get("mode", "insert")

        # Borrar datos existentes justo antes de enviar
        delete_result = delete_all_data()
        if not delete_result.get("success"):
            return jsonify({"success": False, "error": delete_result.get("message", "Error al borrar datos")})

        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        resultados = loop.run_until_complete(
            asyncio.gather(
                *[procesar_hoja_db_async(h, session_id, modo) for h in hojas]
            )
        )

        return jsonify({"success": True, "resultados": resultados})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


# Agregar manejo de conexiones SocketIO
@socketio.on("connect")
def handle_connect():
    print("Cliente conectado:", request.sid)


@socketio.on("disconnect")
def handle_disconnect():
    print("Cliente desconectado:", request.sid)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Verificar si se envió el archivo
        if "excelFile" not in request.files:
            flash("No se encontró el archivo en la solicitud", "error")
            return redirect(request.url)

        file = request.files["excelFile"]
        fecha_facturacion = request.form.get("fechaFacturacion")

        # Validar que se haya seleccionado un archivo
        if file.filename == "":
            flash("No se seleccionó ningún archivo", "error")
            return redirect(request.url)

        # Validar que se haya ingresado una fecha
        if not fecha_facturacion:
            flash("Debe seleccionar una fecha de facturación", "error")
            return redirect(request.url)

        # Validar extensión del archivo
        if file and allowed_file(file.filename):
            try:
                pass

                # Guardar el archivo
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(filepath)

                # Convertir fecha de string a objeto datetime
                # La fecha ya viene en formato YYYY-MM-DD desde el formulario
                fecha_obj = datetime.strptime(fecha_facturacion, "%Y-%m-%d")
                fecha_formateada = fecha_obj.strftime("%m/%Y")  # Solo mes y año

                # Procesar el archivo Excel
                resultado = procesar_excel(filepath, fecha_formateada)

                # Renderizar la plantilla con los resultados
                return render_template("index.html", resultado=resultado)

            except Exception as e:
                flash(f"Error al procesar el archivo: {str(e)}", "error")
                return redirect(request.url)
        else:
            flash(
                "Tipo de archivo no permitido. Solo se aceptan archivos Excel (.xlsx, .xls)",
                "error",
            )
            return redirect(request.url)

    # Método GET - mostrar formulario vacío
    return render_template("index.html")


 


def delete_all_data():
    try:
        conn = get_hana_connection()
        truncate_temp_rep4cfe(conn)
        conn.close()
        return {"success": True, "deleted_count": 0, "message": "Tabla vaciada"}
    except Exception as e:
        return {"success": False, "deleted_count": 0, "message": str(e)}


@app.route("/borrar_datos", methods=["POST", "GET"])
def borrar_datos():
    try:
        print("=== Iniciando borrar_datos ===")
        result = delete_all_data()
        print(f"Resultado: {result}")
        return jsonify(result)
    except Exception as e:
        print(f"Error en borrar_datos: {e}")
        return (
            jsonify(
                {
                    "success": False,
                    "deleted_count": 0,
                    "message": f"Error interno: {str(e)}",
                }
            ),
            500,
        )


@app.route("/test-delete")
def test_delete():
    result = delete_all_data()
    return jsonify({"test_result": result, "timestamp": datetime.now().isoformat()})


@app.route("/test-connection")
def test_connection():
    results = {}

    # Test con hostname
    try:
        target = BASE_URL or "https://telcl-prd-db-cap-telcl-srv.cfapps.us10.hana.ondemand.com"
        resp = requests.get(target, timeout=10)
        results["hostname_test"] = {"success": True, "status": resp.status_code}
    except Exception as e:
        results["hostname_test"] = {"success": False, "error": str(e)}

    # Test con IP directa
    try:
        resp = requests.get(
            "https://52.23.1.211",
            headers={
                "Host": "telcl-prd-db-cap-telcl-srv.cfapps.us10.hana.ondemand.com"
            },
            timeout=10,
            verify=False,
        )
        results["ip_test"] = {"success": True, "status": resp.status_code}
    except Exception as e:
        results["ip_test"] = {"success": False, "error": str(e)}

    # Test delete
    try:
        delete_result = delete_all_data()
        results["delete_test"] = delete_result
    except Exception as e:
        results["delete_test"] = {"success": False, "error": str(e)}

    return jsonify(results)


@app.route("/debug-connectivity")
def debug_connectivity():
    hostname = "telcl-prd-db-cap-telcl-srv.cfapps.us10.hana.ondemand.com"
    results = {}

    # Test DNS
    try:
        ip = socket.gethostbyname(hostname)
        results["dns_resolution"] = {"success": True, "ip": ip}
    except Exception as e:
        results["dns_resolution"] = {"success": False, "error": str(e)}

    # Test port connectivity
    try:
        sock = socket.create_connection((hostname, 443), timeout=10)
        sock.close()
        results["port_connectivity"] = {"success": True}
    except Exception as e:
        results["port_connectivity"] = {"success": False, "error": str(e)}

    # Test HTTP
    try:
        session = requests.Session()
        resp = session.get(f"https://{hostname}", timeout=30)
        results["http_test"] = {"success": True, "status": resp.status_code}
    except Exception as e:
        results["http_test"] = {"success": False, "error": str(e)}

    return jsonify(results)


@app.route("/limpiar_y_redirigir")
def limpiar_y_redirigir():
    return redirect(url_for("index"))


@app.route("/health")
def health_check():
    return jsonify({"status": "healthy", "message": "Service is running"})


# SocketIO handlers
@socketio.on("connect")
def handle_connect():
    print("Cliente conectado:", request.sid)


@socketio.on("disconnect")
def handle_disconnect():
    print("Cliente desconectado:", request.sid)


if __name__ == "__main__":
    socketio.run(app, debug=True, allow_unsafe_werkzeug=True)

# This is a comment only for merge purpose
# This is a comment only for merge purpose too

# Only for merge
