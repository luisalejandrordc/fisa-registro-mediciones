# FISA Registro de Mediciones

Web apps en Google Apps Script para el registro y validación de mediciones de producción en **Fibras Industriales S.A.**: cruces en trenzadoras y RPM en retorcedoras.
Cada web app permite a un supervisor u operador:

- Validarse con un código.
- Seleccionar una Orden de Trabajo (trenzadoras) o una combinación Máquina + Material + Título (retorcedoras) y obtener automáticamente el valor de referencia configurado.
- Registrar las mediciones correspondientes.
- Calcular el estado (DENTRO / FUERA DE RANGO) según el margen de tolerancia definido.
- Guardar todo en una hoja de cálculo de Google Sheets.

## Estructura

```
fisa-registro-mediciones/
├── README.md
├── trenzadoras/
│   ├── Code.gs      # Lógica del servidor (Apps Script)
│   └── Index.html   # Formulario web
└── retorcedoras/
    ├── Code.gs
    └── Index.html
```

Cada carpeta es un proyecto de Apps Script independiente, ligado a su propia hoja de cálculo.

## Trenzadoras

- Se registran 3 máquinas (Fila + Número) por lote.
- Las mediciones pueden repetir fila, pero nunca la misma máquina exacta (Fila + Número).
- La fila **DRIZA** es una excepción: solo tiene 2 máquinas y solo fabrica el producto CABO. Si se selecciona, el formulario reduce el requisito a 1 o 2 mediciones, y todas deben ser de esa fila.

## Retorcedoras

- Cada registro consulta el RPM de referencia configurado para una combinación exacta de Máquina + Material + Título.
- Las máquinas que aparecen agrupadas en la hoja de referencia (ej. `GALAN 1 Y 2`) se expanden automáticamente en máquinas individuales.
- Se distinguen dos tipos de medición: **HUSO** (se evalúa dentro/fuera de rango) y **CORDEL** (se registra sin evaluar rango).
- El rango aceptable es ±5% sobre una base de referencia. Para la mayoría de máquinas esa base es el RPM Configurado; para **DONGTAI 4 y 5** es, en cambio, el promedio histórico de RPM Real (solo mediciones HUSO) registrado para la misma combinación Máquina + Material + Título, redondeado al múltiplo de 50 más cercano. Si aún no hay historial para esas máquinas, se usa el RPM Configurado como respaldo.
- El RPM Promedio se calcula y muestra en pantalla para cualquier máquina (como referencia informativa) y se guarda junto con cada registro, aunque solo defina el rango en DONGTAI 4 y 5.
- La posición debe ser un número entero entre 1 y 40.
- Si la observación de una medición HUSO queda vacía, se completa automáticamente como "OK".

## Despliegue

1. Crear un proyecto de Apps Script vinculado a la hoja de cálculo correspondiente (`Extensiones → Apps Script`).
2. Copiar `Code.gs` e `Index.html` de la carpeta correspondiente al proyecto.
3. Implementar como aplicación web (`Implementar → Nueva implementación → Aplicación web`).
4. Configurar las hojas requeridas con los encabezados requeridos para cada proyeto.

## Stack

- Google Apps Script (backend)
- HTML/CSS/JS vanilla (frontend)
- Google Sheets (database)
