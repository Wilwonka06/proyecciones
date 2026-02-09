# Control de Pagos GCO - Versión 2.5.0

## 📝 Historial de Actualizaciones

### Versión 2.5.0
- **Arquitectura Modular**: Separación del sistema en módulos especializados:
  - `Inicio_control.py`: Lanzador principal con selección de modo (Semanal/Mensual).
  - `control_pagos_semana.py`: Lógica específica para proyecciones semanales.
  - `control_pagos_mes.py`: Lógica específica para proyecciones mensuales.
- **Proyección Mensual Completa**:
  - Generación de archivo con nombre del mes (ej. FEBRERO.xlsx).
  - Inclusión de todos los registros del mes (PAGAR y PENDIENTE).
  - Agrupación por Proveedor y subtotales semanales.
  - Numeración de semanas relativa al mes (1, 2, 3, 4, 5).
  - Tablas resumen por semana y totales generales.
- **Normalización de Marcas**: Unificación automática de nombres de marcas (ej. CHEVIÑON -> CHEVIGNON) en ambos reportes.
- **Mejoras en Estabilidad**:
  - Manejo robusto de fechas anteriores a 1900.
  - Recuperación automática ante errores de bloqueo de Excel (COM).
  - Escritura directa de fórmulas para garantizar precisión en cálculos.

### Versión 2.4.1
- **Mejora en la Interfaz**: Actualización de colores y estilos para mejorar la usabilidad.
- **Compatibilidad con Python 3.11**: Ajustes para asegurar la compatibilidad.
- **Combinación de interfaces**: Integración inicial en un lanzador único.

## 📋 Descripción

Sistema automatizado para la gestión de proyecciones de pagos de importaciones. Esta herramienta facilita el flujo de trabajo mediante la generación automática de proyecciones semanales y mensuales, procesando el archivo maestro de control de pagos.

## ⚙️ Funcionalidades

### Proyección Semanal
- Filtrado por fecha de corte (próximo miércoles).
- Inclusión de registros con estado "PAGAR".
- Agrupación por Marca y Proveedor.
- Generación de archivo con estructura de carpetas: `AÑO YYYY/MES/SEMANA XX`.

### Proyección Mensual
- Filtrado por Mes completo.
- Inclusión de todos los registros (PAGAR y PENDIENTE).
- Agrupación por Proveedor con cortes semanales.
- Generación de archivo único por mes: `AÑO YYYY/MES/MES.xlsx`.
- Segunda hoja con detalle agregado.

## 🖥️ Interfaz de Usuario

- **Lanzador Unificado**: Menú principal para elegir entre proceso Semanal o Mensual.
- **Selectores Inteligentes**:
  - Semanal: Calendario para elegir fecha de corte.
  - Mensual: Selectores de Mes y Año.
- **Progreso Visual**: Barra de progreso y logs en tiempo real.

## 🔧 Configuración Inicial

Al ejecutar la aplicación por primera vez, se solicitará la configuración de las rutas de trabajo (guardadas en `config_pagos.ini`):

1.  **Archivo Origen**: Ubicación del archivo `CONTROL DE PAGOS.xlsm`.
2.  **Carpeta de Proyecciones**: Directorio base para guardar los archivos generados.
3.  **Archivo Final**: Ubicación del archivo maestro `CONTROL PAGOS.xlsx` (solo lectura/referencia).

## 🚀 Instalación y Ejecución

### Requisitos Previos
- Python 3.8+
- Microsoft Excel instalado (para automatización COM).

### Instalación de Dependencias
```bash
pip install -r requirements.txt
```

### Ejecución del Sistema
El punto de entrada es el script `Inicio_control.py`:

```bash
python Inicio_control.py
```

### Compilación (Generar Ejecutable)
Para generar el ejecutable de la aplicación:

```bash
pyinstaller --noconsole --onedir --clean --name="Control Pagos GCO" --icon=icon.ico --hidden-import=pandas --hidden-import=openpyxl --hidden-import=win32com.client --hidden-import=tkcalendar --hidden-import=babel.numbers --collect-all pandas Inicio_control.py
```

## 📁 Estructura del Proyecto

```
Control de Pagos/
│
├── Inicio_control.py           # Lanzador principal (Menu)
├── control_pagos_semana.py     # Lógica de proyección semanal
├── control_pagos_mes.py        # Lógica de proyección mensual
├── config_pagos.ini            # Configuración de rutas (auto-generado)
├── icon.ico                    # Icono de la aplicación
├── requirements.txt            # Dependencias
└── README.md                   # Documentación
```

## 🐛 Solución de Problemas Comunes

- **Error de Rutas**: Si cambia la ubicación de los archivos, elimine `config_pagos.ini` para reconfigurar.
- **Excel Bloqueado**: Si el script falla por "Call rejected by callee", asegúrese de no tener celdas en edición en ningún Excel abierto. El sistema intentará reintentar automáticamente.
- **Fechas Antiguas**: El sistema convierte automáticamente fechas anteriores a 1900 a texto para evitar errores de Excel.

## 👥 Créditos

Desarrollado para GCO - Gestión de Control de Pagos.

## � Contacto

Para soporte o reportar errores:
*   **Correo**: rojaswil336@gmail.com
*   **Teléfono**: 3207199395
