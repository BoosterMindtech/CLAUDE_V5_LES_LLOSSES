# Aigues Les Lloses — v5.0

App de gestión de consumos de agua, facturación y envío de emails para comunidades rurales.
Diseñada para ejecutarse en un VPS o en un ordenador on premise.

---

## Archivos necesarios en el servidor

```
/home/aigues/
├── lloses_app_v5.py    <- la app
└── requirements.txt    <- dependencias
```

La carpeta `data/facturas/` se crea automáticamente al enviar la primera factura.

---

## Instalación

```bash
pip install -r requirements.txt

streamlit run lloses_app_v5.py \
  --server.port 8501 \
  --server.address 0.0.0.0 \
  --server.headless true
```

Accede desde el navegador en: `http://tu-ip:8501`

---

## Arranque automático con systemd

Crea el archivo `/etc/systemd/system/aigues.service`:

```ini
[Unit]
Description=Aigues Les Lloses
After=network.target

[Service]
User=ubuntu
WorkingDirectory=/var/www/CLAUDE_V3_LES_LLOSSES
ExecStart=/var/www/CLAUDE_V3_LES_LLOSSES/venv/bin/streamlit run lloses_app_v5.py \
  --server.port 8501 \
  --server.address 0.0.0.0 \
  --server.headless true
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```

```bash
systemctl daemon-reload
systemctl enable aigues
systemctl restart aigues
systemctl status aigues
```

---

## Credenciales por defecto

| Usuario   | Contraseña | Rol           |
|-----------|------------|---------------|
| admin     | admin123   | Administrador |
| joan.puig | vecino123  | Vecino (demo) |

**Cambia las contraseñas desde Vecinos -> Usuarios en cuanto arranques.**

---

## Importar vecinos desde Excel

Ve a **Vecinos -> Importar Excel**.

Columnas aceptadas (`nombre` es la única obligatoria):

| nombre | direccion | email | telefono | iban |
|--------|-----------|-------|----------|------|
| Joan Puig | Carrer Major 1 | joan@exemple.com | 600111001 | ES00 0000 |

- Descarga la plantilla `.xlsx` desde la misma pestaña.
- Se aceptan `.xlsx` y `.xls`.
- Si el vecino ya existe (mismo nombre) se actualizan sus datos.
- Se crea automáticamente un usuario con contraseña `vecino123`.

---

## Exportar listado de vecinos

Desde **Vecinos -> Listado** puedes descargar el listado completo en formato Excel (`.xlsx`).

---

## Integración con sensores IoT (MQTT)

La app puede recibir lecturas automáticamente de sensores NB-IoT (como el Dragino CPL03-NB) a través de un broker MQTT (Mosquitto).

### Arquitectura

```
Sensor CPL03-NB -> NB-IoT -> Mosquitto -> App Les Lloses -> BD automática
                                       -> Miniserver Loxone -> alertas
```

### Instalar Mosquitto

**En Linux/VPS:**
```bash
sudo apt-get install -y mosquitto mosquitto-clients
sudo systemctl enable mosquitto
sudo systemctl start mosquitto
```

**En Windows (on premise):**
Descargar desde https://mosquitto.org/download/ e instalar como servicio.

**En Mac (desarrollo):**
```bash
brew install mosquitto
brew services start mosquitto
```

Añadir al archivo de configuración de Mosquitto:
```
listener 1883
allow_anonymous true
```

### Acceso tecnico a la configuracion de sensores

La seccion de sensores en **Configuracion -> Sensores** esta protegida con contrasena de tecnico.
Solo el tecnico instalador puede acceder a esta seccion.

**Contrasena de tecnico:** guardada de forma segura en el codigo fuente.

Una vez autenticado el tecnico puede:
- Configurar el broker MQTT (IP, puerto, topic)
- Añadir sensores y asignarlos a vecinos
- Ver el estado de cada sensor
- Eliminar sensores

Al terminar pulsar **Cerrar sesion tecnico** para volver a bloquear la seccion.

### Configurar el broker MQTT en la app

Ve a **Configuracion -> Sensores** (requiere contrasena tecnico):

| Campo | Valor |
|-------|-------|
| Broker MQTT | IP del ordenador donde corre Mosquitto |
| Puerto | 1883 |
| Topic base | leslloses/# |
| MQTT activo | activado |

### Mapeo sensor a vecino

Una vez dentro de la seccion tecnico, añadir cada sensor:

| Campo | Valor |
|-------|-------|
| ID Sensor | ID unico del sensor (visible en la etiqueta o en el topic MQTT) |
| Vecino | Seleccionar del desplegable |
| Pulsos/m3 | 100 (estandar para contadores domesticos) |

### Estado de los sensores

- Online: ha enviado datos en los ultimos 3 dias
- Sin senal: lleva mas de 3 dias sin enviar (posible averia o sabotaje)

### Configurar el sensor Dragino CPL03-NB

Usar la app BLE de Dragino para enviar estos comandos AT al sensor:

```
AT+PRO=3,5                              // MQTT con payload JSON
AT+SERVADDR=IP_MOSQUITTO,1883           // IP del broker y puerto
AT+PUBTOPIC=leslloses/contador/ID_SENSOR  // Topic del sensor
```

---

## Configuracion de email

Ve a **Facturacion -> Configuracion de email**.

Para Hostalia o hosting propio, abrir **Configuracion SMTP avanzada**:

| Campo | Valor |
|-------|-------|
| Servidor SMTP | smtp.servidor-correo.net |
| Puerto | 587 |
| SSL | desactivado |
| No verificar certificado SSL | activado |

**Gmail:** usa una Contrasena de aplicacion, deja el servidor SMTP vacio.
Cuenta Google -> Seguridad -> Verificacion en 2 pasos -> Contrasenas de aplicacion.

Limite Gmail: 500 emails/dia. Si hay mas de 500 vecinos con email, espaciar el envio.

---

## Facturas guardadas

```
data/facturas/YYYY/MM/NombreMes-YYYY_Nombre_Vecino.pdf
```

Desde **Facturacion -> Ver facturas guardadas** puedes consultar el historico y descargar facturas anteriores.

---

## Backup de la base de datos

**Configuracion -> Backup base de datos** descarga el fichero `lloses.db`.
Guardarlo periodicamente — contiene todos los vecinos, consumos, sensores y configuracion.

---

## Dependencias

```
streamlit>=1.32.0
pandas>=2.0.0
plotly>=5.18.0
reportlab>=4.0.0
openpyxl>=3.1.0
xlrd>=2.0.1
paho-mqtt>=1.6.1
```

SQLite viene incluido en Python, no necesita instalacion separada.
