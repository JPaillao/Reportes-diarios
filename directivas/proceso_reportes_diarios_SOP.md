**Tecnologías:** [[Google Apps Script]]
**Procesos:** [[Consolidación de Datos]], [[Optimización de API]]

# SOP: Lógica Completa del Proceso de Reportes Diarios (Daily Report)

## 1. Visión General
El proceso de automatización de los "Daily Reports" (DR) se encarga de recopilar la información de recursos y actividades reportadas diariamente en hojas de cálculo por cada proyecto y consolidarlas en dos bases de datos maestras (Recursos y Actividades).

## 2. Orquestación y Ejecución (`03 Ejecutar_DR.js`)
Existen dos flujos principales, que procesan secuencialmente una lista de proyectos activos (ej. QUILO, LVIOL, SJCHI, CFDAL).
- **Control de concurrencia:** Se utiliza `LockService.getScriptLock()` para asegurar que no se procese simultáneamente más de un script y se eviten conflictos al escribir en las bases de datos.
- **Enrutamiento por Proyecto:** Se detecta si el proyecto es un sistema BESS (ej. CFDAL). Si es BESS, se usan las funciones especializadas (`dr_recursos_bess`, `dr_actividades_bess`). De lo contrario, se usan las estándar (`dr_recursos`, `dr_actividades`).
- **Manejo de Errores y Retrasos:** Entre cada proyecto se ejecutan pausas (`Utilities.sleep()`) para no saturar los límites de cuota de escritura de Google Sheets. Si hay un error, se envía un correo a la administración notificando el fallo.

## 3. Lógica de Extracción de Recursos (`01 Recursos.js` y `01b Recursos BESS.js`)
### Flujo de Ejecución:
1. **Identificación de Archivos:** Escanea la carpeta del proyecto en Drive buscando archivos cuyo nombre contenga "DR" y que hayan sido modificados en la fecha de ejecución (o en un rango provisto).
2. **Validación de Fecha:** Se extrae la fecha interna del reporte (habitualmente celda `M6` o buscándola por el título "Fecha de daily report:"). Se contrasta esta fecha con la declarada en el nombre del archivo (`XXXXX_DR_YYYYMMDD`). Si no coinciden, se rechaza el archivo por inconsistencia y se envía una notificación.
3. **Mapeo de Bloques de Recursos:**
   - Se buscan las secciones de "Gastos Generales", "Terreno" y "Maquinaria".
   - **BESS (Dinámico):** Se leen las empresas ubicadas en las columnas adyacentes a cada bloque de recurso, deteniéndose al encontrar la columna "TOTAL".
   - **Estándar:** Lee posiciones de empresas más fijas según el formato tradicional.
4. **Transformación (Aplanado):** Se convierte la tabla de múltiples columnas de empresas en registros planos donde cada fila representa a un `[Recurso, Cantidad, Tipo de Recurso, Fecha, Proyecto, ID, Empresa]`.
5. **Limpieza e Inserción:** 
   - Se identifican qué fechas están presentes en la nueva carga y se procede a eliminar en bloque (*batch delete*) todas las filas de esa fecha en la base consolidada para evitar duplicados.
   - Se insertan las nuevas filas extraídas.

## 4. Lógica de Extracción de Actividades (`02 Actividades.js` y `02b Actividades BESS.js`)
### Flujo de Ejecución:
1. **Mismos pasos iniciales:** Identificación y validación estricta de fecha entre el contenido (M6) y el nombre de archivo.
2. **Ubicación de Bloques:** Se busca dónde empiezan los encabezados que contienen "Actividad", "Unidad", "Cantidad", etc., y dónde terminan (generalmente antes de "Observaciones" o "Firma").
3. **Mapeo de Columnas (BESS y Estándar):** 
   - Mapea dinámicamente o por posición los valores de: Cantidad Anterior, Cantidad Hoy, Cantidad Acumulada, % Avance, etc.
   - En el caso de proyectos BESS, además lee la columna "Empresa" al final de las actividades y descarta (`continue`) las que tengan valor "Transversal". Esta empresa se añade a la matriz final de datos.
4. **Filtrado por Acumulado:** Toda actividad cuyo valor "Acumulado" sea vacío, texto o no represente un número es descartada.
5. **Validación Matemática (Discrepancias):** Antes de insertar, el script verifica internamente si la ecuación: `Cantidad Anterior + Cantidad Hoy = Cantidad Acumulada` es correcta. Si la diferencia es mayor a un margen de tolerancia (0.0001), se captura la discrepancia. Si se encuentran problemas, se alerta vía correo electrónico para corrección manual.
6. **Limpieza e Inserción:** Al igual que en recursos, se eliminan las filas del día y se insertan las nuevas filas consolidadas.

## 5. Restricciones y Consideraciones
- **Formatos de Nombre de Archivo:** Es estricto. Debe mantener el regex: `/^(\w{5})_DR_(\d{8})/`.
- **Estructura del Consolidado:** Las columnas clave (como Fecha) se encuentran en posiciones específicas dentro del consolidado final (Columna 16 / índice 15 para fechas). Modificar el orden en la base consolidada puede romper los scripts de limpieza e inconsistencias.
- **Evitar Timeouts:** Por esto se hacen operaciones en lote (`setValues` una sola vez en lugar de appendRow) y borrados de filas unificados por rangos dinámicos.

## 6. Casos Borde y Resolución de Problemas

### 6.1 Timeout por Recálculo Masivo de Fórmulas
**Síntoma:** Error `Exception: Service Spreadsheets timed out while accessing document...` al ejecutar el consolidado final (ej. `06 Consolidado Actividades.js`).
**Causa:** Si la hoja destino alimenta a otra hoja pesada a través de fórmulas como `FILTER`, el uso tradicional de `setValues()` o `clearContent()` "congela" la ejecución de Apps Script mientras espera que toda la hoja recalcule, provocando el timeout de 6 minutos.
**Solución:** Se reemplazó el uso de `SpreadsheetApp` por el **Servicio Avanzado Nativo de Google Sheets** (`Sheets.Spreadsheets.Values.clear` y `.update`). Al ser llamadas API directas, liberan la ejecución instantáneamente (el script demora 5 segundos) dejando que Sheets procese el recálculo asincrónicamente.
**Dependencia:** Requiere que el Servicio Avanzado "Google Sheets API" esté explícitamente añadido desde la barra lateral del editor.
**Registro de Error (Obsidian):** [[Fallas_Cuotas_APIs]]

### 6.2 Inserción de Fechas en Formato ISO 8601 por API REST
**Síntoma:** Al transferir datos con `Sheets.Spreadsheets.Values.update`, las celdas de fechas aparecen como texto crudo (ej: `2026-03-16T03:00:00.000Z`) en lugar del formato legible de Google Sheets.
**Causa:** La API REST de Google Sheets (`Sheets API`) no sabe interpretar objetos `Date` nativos de Javascript al serializarlos en formato JSON (convirtiéndolos a cadenas ISO), a diferencia del viejo `SpreadsheetApp.setValues()` que sí los parsea como fechas seriales de Sheets.
**Solución:** Se interceptan los datos en la matriz (`masterData`) antes de enviarlos a la API y se formatea explícitamente cualquier objeto Date (`cell instanceof Date`) hacia un string estructurado (`dd-MM-yyyy`) mediante `Utilities.formatDate()`. Así la opción `USER_ENTERED` de la API lo ingresa correctamente como fecha.
**Registro de Error (Obsidian):** [[Fallas_Algoritmicas_JS_Python]]

## 7. Casos Borde: Lógica de Extracción BESS (Recursos y Actividades)

### 7.1 Recursos (BESS)
- **Dinámica de Columnas:** Se añaden más columnas de contratistas por cada categoría (Gastos Generales, Terreno, Maquinaria). Anteriormente eran 3 fijas, ahora se capturan de forma dinámica hasta encontrar "TOTAL" o celda vacía.
- **Extracción de Terreno:** Se usa `.includes("terreno")` buscando en la misma fila de "Gastos generales". Se excluye la palabra "jefe" para evitar capturar "Jefe de terreno".
- **Celdas combinadas (Merged Cells):** Es común que las cabeceras de sección ("TERRENO") abarquen varias columnas (ej. H e I). Si el texto cae en H pero los recursos se digitan en I, se debe revisar siempre `coord.c + 1` para no omitir la sección.

### 7.2 Actividades (BESS)
- **Columna Empresa:** Se captura la nueva columna "Empresa" en actividades. Se incluye en la matriz consolidada final, expandiéndola de 18 a 19 columnas. Se debe mantener el índice 16 para la columna de Fecha.
