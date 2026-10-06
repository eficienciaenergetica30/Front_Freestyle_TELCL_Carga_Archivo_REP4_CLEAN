# 🚀 App Flask - Carga de Archivos REP4

Aplicación Flask para procesamiento y carga de archivos Excel con SocketIO.

## 🛠️ Tecnologías
- Flask 2.3.3
- Flask-SocketIO 5.3.6
- Openpyxl 3.1.2
- SAP BTP (Cloud Foundry)

## 📦 Instalación

```bash
# Clonar repositorio
git clone https://github.com/eficienciaenergetica30/Front_Freestyle_TELCL_Carga_Archivo_REP4.git

# Crear entorno virtual
python -m venv venv

# Activar entorno
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows

# Instalar dependencias
pip install -r requirements.txt
pip install -r requirements.txt
```

## Proceso final de carga

Después de guardar todas las hojas sin errores, el servidor hace un único POST
a `https://tlcl-processes-hub.cfapps.us10.hana.ondemand.com/tlcl-hub/tlcl13`, sin cuerpo ni reintentos automáticos.
Es la última operación del flujo de carga y la interfaz espera su respuesta.
HTTP 200 muestra «Proceso finalizado con éxito»; otros códigos muestran un
aviso breve con el código, sin revelar el contenido de la API.
Si no hay respuesta, se informa `TIMEOUT` o `CONEXION`.

- `FINAL_PROCESS_URL`: permite cambiar la URL desde el entorno o `.env`.
- `FINAL_PROCESS_TIMEOUT`: espera máxima de la API en segundos (60 por defecto).

`localhost` se refiere al servidor Flask. En despliegues remotos, configurar la
dirección accesible de la API. El timeout de Gunicorn/proxy debe cubrir tanto la
carga completa como la espera del proceso final (Gunicorn tiene 120 s actualmente).
Un timeout no confirma si el procedimiento terminó; verificar su estado antes
de volver a ejecutar la carga.

Pruebas aisladas, sin conexión a HANA ni a la API:

```bash
python -m unittest discover -s tests -v
```
