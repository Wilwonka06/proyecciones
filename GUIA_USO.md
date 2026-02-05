# 📖 Guía Rápida de Usuario - Control de Pagos GCO v2.0

## 🚀 Inicio Rápido

### Paso 1: Abrir la Aplicación
- Hacer doble clic en `ControlPagosGCO.exe`
- Aparecerá la ventana principal

### Paso 2: Seleccionar el Proceso

Tiene tres opciones disponibles:

#### 1️⃣ **Crear Solo Proyección Semanal**
📊 **¿Cuándo usar?**
- Cuando quiere generar la proyección para revisarla primero
- Cuando necesita un reporte preliminar
- Cuando el proceso requiere aprobación antes de confirmar

📌 **Resultado:**
- Se crea el archivo de proyección en la carpeta correspondiente
- NO se modifica el archivo de control de pagos final
- Puede revisar y validar los datos antes de confirmarlos

#### 2️⃣ **Anexar Solo a Control Pagos Final**
📝 **¿Cuándo usar?**
- Cuando ya tiene una proyección creada y revisada
- Cuando quiere confirmar una proyección previamente generada
- Para separar la generación de datos de su aprobación

⚠️ **Importante:**
- Debe existir el archivo de proyección para la fecha seleccionada
- Si no existe, el sistema mostrará un error indicando la ruta esperada

📌 **Resultado:**
- Lee el archivo de proyección existente
- Anexa los registros al archivo final de control de pagos
- NO crea un nuevo archivo de proyección

#### 3️⃣ **Ejecutar Proceso Completo**
⚙️ **¿Cuándo usar?**
- Para el proceso rutinario normal
- Cuando no necesita revisión previa
- Cuando tiene confianza en los datos

📌 **Resultado:**
- Crea la proyección Y anexa al archivo final
- Proceso automático completo (comportamiento original)

### Paso 3: Seleccionar la Fecha
- El sistema sugiere automáticamente el próximo miércoles
- Puede cambiar la fecha usando el calendario
- Si selecciona un miércoles, aparecerá una marca verde ✓

### Paso 4: Confirmar
- Clic en **"EJECUTAR PROCESO"**
- Revisar el mensaje de confirmación
- Clic en **"Sí"** para continuar

### Paso 5: Esperar el Proceso
- Aparecerá una ventana mostrando el progreso
- Puede ver el log de acciones realizadas
- Al finalizar, aparecerá un mensaje de éxito

## 🎯 Flujos de Trabajo Recomendados

### Flujo 1: Con Revisión (Recomendado para datos importantes)
```
Miércoles - Generación:
1. Abrir aplicación
2. Seleccionar "Opción 1: Crear Solo Proyección"
3. Seleccionar fecha
4. Ejecutar
5. Revisar archivo generado

Miércoles/Jueves - Confirmación:
1. Abrir aplicación
2. Seleccionar "Opción 2: Anexar Solo a Control Final"
3. Seleccionar la misma fecha
4. Ejecutar
5. Verificar archivo final
```

### Flujo 2: Proceso Rápido (Para proceso rutinario)
```
Miércoles:
1. Abrir aplicación
2. Seleccionar "Opción 3: Proceso Completo"
3. Seleccionar fecha
4. Ejecutar
5. Listo
```

### Flujo 3: Corrección de Errores
```
Si encuentra un error después de generar la proyección:
1. Corregir el archivo base de Control de Pagos
2. Abrir aplicación
3. Seleccionar "Opción 1: Crear Solo Proyección"
4. Generar nueva proyección (sobrescribe la anterior)
5. Revisar correcciones
6. Seleccionar "Opción 2: Anexar Solo a Control Final"
7. Confirmar
```

## ❓ Preguntas Frecuentes

### ¿Qué hacer si el archivo está abierto?
**Respuesta:** Cerrar todos los archivos Excel relacionados y hacer clic en "Reintentar"

### ¿Qué pasa si selecciono la Opción 2 y no existe la proyección?
**Respuesta:** El sistema mostrará un mensaje de error indicando:
- La fecha seleccionada
- La ruta exacta donde debería estar el archivo
- Sugerencia de usar Opción 1 o 3

### ¿Puedo cambiar la fecha sugerida?
**Respuesta:** Sí, puede seleccionar cualquier fecha usando el calendario

### ¿Qué día se recomienda para las proyecciones?
**Respuesta:** Los miércoles, ya que el sistema está diseñado para proyecciones semanales que inician los miércoles

### ¿Puedo generar la proyección dos veces?
**Respuesta:** Sí, al usar la Opción 1, si ya existe un archivo para esa fecha, se sobrescribirá

### ¿Qué pasa si ejecuto la Opción 2 dos veces?
**Respuesta:** Se anexarán los registros duplicados. NO se recomienda ejecutar dos veces la misma proyección

## 🔍 Verificación de Resultados

### Después de la Opción 1:
 Verificar que existe el archivo en:
```
O:\Finanzas\Info Bancos\Pagos Internacionales\
PROYECCION PAGOS SEMANAL Y MENSUAL\
AÑO 2026\[MES]\[DD MES YYYY].xlsx
```

 Abrir el archivo y verificar:
- Hoja "Control_Pagos" con datos filtrados
- Hoja "[MES DD]" con la proyección formateada
- Agrupaciones por Importador y Proveedor
- Totales calculados correctamente

### Después de la Opción 2 o 3:
 Verificar el archivo final:
```
O:\Finanzas\Info Bancos\Pagos Internacionales\
CONTROL PAGOS.xlsx
```

 Revisar en la hoja "Pagos Importación":
- Nuevos registros agregados al final
- Fecha de pago correcta
- Todos los campos completos

## ⚠️ Errores Comunes y Soluciones

### Error: "No se encuentra el archivo original"
**Causa:** El archivo base no está en la ubicación esperada
**Solución:** 
- Verificar que existe: `O:\00.CONTROL DE PAGOS 2026.xlsm`
- Verificar que la unidad O: está mapeada

### Error: "No se encontró hoja Control_Pagos"
**Causa:** El archivo base tiene un nombre de hoja diferente
**Solución:** 
- Abrir el archivo base
- Verificar que existe una hoja llamada "Control_Pagos"
- O modificar el archivo config.ini

### Error: "No se encontraron registros"
**Causa:** No hay registros con estado "PAGAR" para la semana seleccionada
**Solución:**
- Verificar que existen registros para esa semana
- Verificar que tienen el estado correcto
- Seleccionar otra fecha

### Error: "El archivo está abierto"
**Causa:** Archivo Excel abierto en otra aplicación
**Solución:**
- Cerrar todos los archivos Excel
- Hacer clic en "Reintentar"

## 📞 Soporte

Si encuentra problemas no listados aquí:
1. Verificar el log de acciones en la ventana de progreso
2. Tomar captura de pantalla del error
3. Contactar al equipo de TI con la información del error

## 💡 Consejos

1. **Ejecute la Opción 1 los martes** para tener la proyección lista para revisar el miércoles
2. **Use la Opción 2 después de aprobar** la proyección con su supervisor
3. **Mantenga cerrados todos los archivos Excel** durante el proceso
4. **Haga backup** del archivo de control de pagos antes de usar la Opción 2 o 3
5. **Revise siempre** el mensaje final para confirmar que todo se ejecutó correctamente

---

**Versión:** 2.0  
**Última actualización:** Febrero 2026
**Desarrollado por: Wilson David Rojas Palacios