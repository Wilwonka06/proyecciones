# Control de Pagos GCO - Versión 2.0

## 📋 Descripción

Sistema automatizado para la gestión de proyecciones de pagos y actualización del archivo de control de pagos final.

## 🆕 Nuevas Funcionalidades (v2.0)

### Opciones de Proceso Independientes

El sistema ahora permite al usuario elegir qué proceso desea ejecutar:

#### **Opción 1: Crear Solo Proyección Semanal** 📊
- Genera únicamente el archivo de proyección para la semana seleccionada
- No modifica el archivo de control de pagos final
- Ideal para:
  - Revisar la proyección antes de confirmarla
  - Generar reportes preliminares
  - Análisis de datos sin compromiso

**Proceso:**
1. Copia el archivo base de Control de Pagos
2. Filtra los registros por la fecha seleccionada
3. Crea la hoja de proyección con formato y agrupaciones
4. Guarda el archivo en la carpeta correspondiente

#### **Opción 2: Anexar Solo a Control Pagos Final** 📝
- Lee un archivo de proyección existente
- Anexa los registros al archivo final de control de pagos
- **Validación importante:** Verifica que exista el archivo de proyección antes de proceder
- Ideal para:
  - Confirmar una proyección previamente revisada
  - Separar la generación de la proyección de su aprobación
  - Flujo de trabajo con múltiples revisiones

**Proceso:**
1. Busca el archivo de proyección para la fecha seleccionada
2. Si no existe, muestra un mensaje de error con la ruta esperada
3. Si existe, lee los datos de la hoja de proyección
4. Filtra y limpia los datos (elimina filas vacías y totales)
5. Anexa los registros al archivo final

**Ubicación del archivo de proyección:**
```
O:\Finanzas\Info Bancos\Pagos Internacionales\PROYECCION PAGOS SEMANAL Y MENSUAL\
  └── AÑO 2026\
      └── FEBRERO\
          └── 05 FEBRERO 2026.xlsx
```

#### **Opción 3: Ejecutar Proceso Completo** ⚙️
- Combina las opciones 1 y 2
- Crea la proyección Y anexa al archivo final
- Comportamiento original del sistema
- Ideal para:
  - Proceso rutinario sin necesidad de revisión previa
  - Cuando se tiene confianza en los datos

## 🎯 Casos de Uso

### Caso 1: Flujo de Trabajo con Revisión
```
1. Miércoles (Usuario A): 
   - Selecciona Opción 1
   - Genera proyección para revisión

2. Miércoles (Usuario B):
   - Revisa el archivo de proyección
   - Aprueba o solicita cambios

3. Jueves (Usuario A):
   - Selecciona Opción 2
   - Anexa los datos al archivo final
```

### Caso 2: Proceso Rápido
```
1. Miércoles:
   - Selecciona Opción 3
   - Proceso completo automático
```

### Caso 3: Regeneración de Proyección
```
1. Se detecta un error en la proyección generada
2. Se corrige el archivo base
3. Se vuelve a ejecutar Opción 1
4. Se revisa la nueva proyección
5. Se ejecuta Opción 2 para anexar
```

## 🖥️ Interfaz de Usuario

La nueva interfaz incluye:

- **Selector de Proceso**: Tres opciones claramente diferenciadas con:
  - Iconos visuales (emojis)
  - Descripción de cada opción
  - Código de colores para identificación rápida
  
- **Selector de Fecha**: Calendario interactivo con:
  - Sugerencia automática del próximo miércoles
  - Indicador visual cuando se selecciona un miércoles
  - Validación de fecha

- **Mensajes de Confirmación**: Adaptados según la opción seleccionada

## ⚠️ Validaciones y Advertencias

### Opción 2 - Verificación de Archivo Existente

Cuando se selecciona la Opción 2, el sistema:

1. **Busca el archivo de proyección** en la ubicación esperada
2. **Si no existe:**
   - Muestra un mensaje de error detallado
   - Indica la ruta exacta donde debería estar el archivo
   - Sugiere usar Opción 1 o Opción 3
3. **Si existe:**
   - Lee la hoja de proyección correspondiente
   - Limpia los datos (elimina filas vacías y de totales)
   - Procede con el anexado

### Ejemplo de Mensaje de Error

```
Archivo No Encontrado

No se encontró el archivo de proyección para:
05/02/2026

Ruta esperada:
O:\Finanzas\Info Bancos\Pagos Internacionales\
PROYECCION PAGOS SEMANAL Y MENSUAL\AÑO 2026\FEBRERO\
05 FEBRERO 2026.xlsx

Por favor, primero cree la proyección (Opción 1) 
o ejecute el proceso completo (Opción 3).
```

## 📁 Estructura de Archivos

```
Control de Pagos/
│
├── control_pagos_v1.py           # Archivo principal
├── compilar.bat                   # Script de compilación
├── config.ini                     # Configuración de rutas (opcional)
├── icon.ico                       # Icono de la aplicación
└── requirements.txt               # Dependencias Python
```

## 🔧 Configuración

### config.ini (Opcional)

```ini
[RUTAS]
ArchivoOrigen = O:\00.CONTROL DE PAGOS 2026.xlsm
CarpetaIntermedia = O:\Finanzas\Info Bancos\Pagos Internacionales\PROYECCION PAGOS SEMANAL Y MENSUAL
ArchivoFinal = O:\Finanzas\Info Bancos\Pagos Internacionales\CONTROL PAGOS.xlsx
```

## 🚀 Instalación y Uso

### Modo Desarrollo
```bash
# Instalar dependencias
pip install -r requirements.txt

# Ejecutar
python control_pagos_v1.py
```

### Compilar Ejecutable
```bash
# Ejecutar el script de compilación
compilar.bat

# Seleccionar opción:
# 1 - Con consola (para debugging)
# 2 - Sin consola (versión final)
```

## 📊 Flujo de Datos

### Opción 1: Solo Proyección
```
[Archivo Base] 
    ↓
[Copia y Limpia]
    ↓
[Filtra por Fecha]
    ↓
[Agrupa y Formatea]
    ↓
[Proyección Semanal.xlsx] ✓
```

### Opción 2: Solo Anexar
```
[Proyección Existente.xlsx] ✓
    ↓
[Lee y Limpia Datos]
    ↓
[Prepara Formato Final]
    ↓
[Anexa a Control Final.xlsx] ✓
```

### Opción 3: Proceso Completo
```
[Archivo Base]
    ↓
[Opción 1: Crea Proyección] ✓
    ↓
[Opción 2: Anexa a Final] ✓
```

## 🐛 Solución de Problemas

### Error: "No se encuentra el archivo de proyección"
**Solución:** 
- Verificar que la fecha seleccionada es correcta
- Asegurarse de haber ejecutado primero la Opción 1
- Verificar la estructura de carpetas

### Error: "Archivo está abierto"
**Solución:**
- Cerrar todos los archivos Excel relacionados
- Intentar nuevamente

### Error: "No se encontraron registros"
**Solución:**
- Verificar que existan registros para la semana seleccionada
- Revisar que los registros tengan el estado "PAGAR"
- Verificar las fechas de vencimiento

## 📝 Registro de Cambios

### Versión 2.0 (Actual)
-  Añadido selector de proceso (3 opciones)
-  Implementada validación de archivo existente para Opción 2
-  Mejorada interfaz de usuario con secciones visuales
-  Añadidos mensajes de confirmación personalizados
-  Mejorada limpieza de datos al leer proyecciones existentes

### Versión 1.1
- Corrección de problemas con hojas ocultas
- Mejora en el manejo de columnas
- Interfaz gráfica moderna

### Versión 1.0
- Versión inicial
- Proceso automatizado completo

## 👥 Créditos

Desarrollado para GCO - Gestión de Control de Pagos

## 📄 Licencia

Uso interno - GCO

## Contacto - Autor

Para cualquier pregunta, sugerencia o problema, por favor abra un issue en este repositorio o contacte al autor directamente.
* correo: rojaswil336@gmail.com
* telefono: 3207199395

