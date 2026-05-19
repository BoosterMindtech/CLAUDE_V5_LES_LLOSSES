# Aigues Les Lloses — v5.0

App de gestión de consumos de agua, facturación y envío de emails para comunidades rurales.
Diseñada para ejecutarse en un VPS propio.

---

## Archivos necesarios en el servidor

```
/home/aigues/
├── lloses_app_v5.py    <- la app
└── requirements.txt    <- dependencias
```

La carpeta `data/facturas/` se crea automáticamente al enviar la primera factura.

---

## Instalación en VPS

```bash
pip install -r requirements.txt

streamlit run lloses_app_v5.py \
  --server.port 8501 \
  --server.address 0.0.0.0 \
  --server.headless true
```

Accede desde el navegador en: `http://tu-ip-vps:8501`

---

## Arranque automático con systemd

Crea el archivo `/etc/systemd/system/aigues.service`:

```ini
[Unit]
Description=Aigues Les Lloses
After=network.target

[Service]
User=aigues
WorkingDirectory=/home/aigues
ExecStart=/usr/local/bin/streamlit run lloses_app_v5.py \
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
systemctl start aigues
systemctl status aigues
```

---

## Credenciales por defecto

| Usuario   | Contraseña | Rol           |
|-----------|------------|---------------|
| admin     | admin123   | Administrador |
| joan.puig | vecino123  | Vecino (demo) |

**Cambia las contraseñas desde Vecinos → Usuarios en cuanto arranques.**

---

## Importar vecinos desde Excel

Ve a **Vecinos → Importar Excel**.

Columnas aceptadas (`nombre` es la única obligatoria):

| nombre | direccion | email | telefono | iban |
|--------|-----------|-------|----------|------|
| Joan Puig | Carrer Major 1 | joan@exemple.com | 600111001 | ES00 0000 |
| Maria Sala | Carrer Major 3 | maria@exemple.com | | |
| Pere Font | Carrer del Pi 2 | | | |

- Descarga la plantilla `.xlsx` desde la misma pestaña.
- Se aceptan archivos `.xlsx` y `.xls`.
- Si el vecino ya existe (mismo nombre) se actualizan sus datos.
- Se crea automáticamente un usuario con contraseña `vecino123`.

---

## Exportar listado de vecinos

Desde **Vecinos → Listado** puedes descargar el listado completo en formato Excel (`.xlsx`).

---

## Facturación con muchos vecinos

La sección de facturación muestra una tabla resumen con los totales calculados sin generar PDFs.
Los PDFs se generan solo cuando se pulsa el botón correspondiente, evitando bloqueos con listas grandes.

---

## Configuración de email

Ve a **Facturación → Configuración de email**.

Para Hostalia o hosting propio, abre **⚙️ Configuración SMTP avanzada**:

| Campo                        | Valor                        |
|------------------------------|------------------------------|
| Servidor SMTP                | smtp.servidor-correo.net     |
| Puerto                       | 587                          |
| SSL                          | desactivado                  |
| No verificar certificado SSL | activado                     |

Pulsa **Guardar configuración email** antes de enviar.

**Gmail:** necesitas una Contraseña de aplicación.
Cuenta Google → Seguridad → Verificación en 2 pasos → Contraseñas de aplicación.

---

## Facturas guardadas en el servidor

```
data/facturas/YYYY/MM/NombreMes-YYYY_Nombre_Vecino.pdf
```

Desde **Facturación → Ver facturas guardadas** puedes consultar el histórico y descargar cualquier factura anterior.

---

## Backup de la base de datos

**Configuración → Backup base de datos** descarga el fichero `lloses.db` completo.

---


## Integración con sensores IoT (MQTT)

La app recibe lecturas automáticamente de sensores NB-IoT via broker MQTT (Mosquitto).

Ve a **Configuración -> Sensores** para:
- Configurar el broker MQTT (IP, puerto, topic)
- Añadir sensores y asignarlos a vecinos
- Ver el estado de cada sensor (Online / Sin señal)

Dependencia adicional: paho-mqtt>=1.6.1

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

SQLite viene incluido en Python, no necesita instalación separada.
