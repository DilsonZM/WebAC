# WebAC (SharePoint SPA)

Aplicación SPA sin NPM ni build pipeline, desplegada como página de SharePoint para evaluar criticidad de activos. Incluye un wizard de evaluación por niveles, dashboard con KPIs y gráficos, panel de administración, detalle de evaluaciones y una matriz dinámica de criticidad.

---

## 1) Objetivo y alcance
### Objetivo
Evaluar criticidad de activos por distintos niveles (Fleet / ProcesoSistema / EGI / Equipo), registrar evaluaciones en SharePoint, medir cobertura y visualizar indicadores de gestión.

### Incluye
- Wizard guiado con validaciones y cálculo automático de métricas.
- Dashboard con KPIs, listado de últimas evaluaciones y gráficos.
- Administración de escenarios y gestión de solicitudes de edición.
- Vista detalle de evaluación con filtros por equipos asociados.
- Matriz dinámica de criticidad con intersección FF/Severidad.

### No incluye
- Backend propio (usa SharePoint REST/Search API).
- Bundling o compilación (sin NPM, sin build).

---

## 2) Arquitectura y tecnología
### Arquitectura
- SPA en HTML/CSS/JS alojada en SharePoint.
- Persistencia en listas de SharePoint.
- Comunicación vía REST API y Search API (fallback por umbral).

### Librerías externas
- SweetAlert2 (modales).
- Chart.js (gráficos).

---

## 3) Estructura de archivos (clave)
- [home.aspx](home.aspx): página principal (SharePoint) que carga la SPA.
- [SiteAssets/AppAC/index.html](SiteAssets/AppAC/index.html): markup base alterno de la SPA.
- [SiteAssets/AppAC/js/app.js](SiteAssets/AppAC/js/app.js): lógica principal (wizard, dashboard, filtros, gráficos).
- [SiteAssets/AppAC/js/matrixModal.js](SiteAssets/AppAC/js/matrixModal.js): matriz dinámica de criticidad (SVG + intersección).
- [SiteAssets/AppAC/js/spHttp.js](SiteAssets/AppAC/js/spHttp.js): cliente SharePoint REST, paginación y metadata.
- [SiteAssets/AppAC/js/catalogService.js](SiteAssets/AppAC/js/catalogService.js): carga de EquiposAC con paginación + Search fallback.
- [SiteAssets/AppAC/js/scenarioService.js](SiteAssets/AppAC/js/scenarioService.js): carga de escenarios con Search fallback.
- [SiteAssets/AppAC/css/style.css](SiteAssets/AppAC/css/style.css): estilos generales.
- [HISTORIAL.md](HISTORIAL.md): cambios del proyecto.

---

## 4) Flujo funcional (wizard)
1. **Nivel:** selección de nivel (Fleet / ProcesoSistema / EGI / Equipo).
2. **Selección/Target:** filtros por jerarquía y selección de equipos cuando aplica.
3. **Escenario:** búsqueda y selección de escenario; paginación y fallback a Search API.
4. **Criterios:** 10 preguntas (FO, FIN, HSE, MA, SOC, DDHH, REP, LEGAL, TI, FF) + evento editable + justificación obligatoria.
5. **Resumen:** target, cobertura (3×2), escenario, evento y métricas.
6. **Confirmación:** registro exitoso.

Notas:
- Si el target ya está evaluado al 100%, se bloquea nueva evaluación.
- Se ofrece acceso directo a la evaluación previa si existe.

---

## 5) Cálculos y reglas de negocio


### 5.1 Impacto (Matriz)
- **Smax:** Valor máximo seleccionado entre las consecuencias (FIN, HSE, MA, SOC, DDHH, REP, LEGAL, TI).
- **Impacto:** $S_{max} \times FO$. (Rango posible: 1 a 25).

### 5.2 Nivel de Severidad
La severidad (eje Y de la matriz) se obtiene ubicando el valor del **Impacto** en los siguientes rangos estrictos definidos por el negocio:
- **20 - 25:** Severidad 5
- **15 - 19:** Severidad 4
- **10 - 14:** Severidad 3
- **5 - 9:** Severidad 2
- **1 - 4:** Severidad 1

### 5.3 Nivel de Criticidad (NC)
El NC se obtiene desde la matriz de riesgo 5×5 rígida, cruzando la **Severidad** (Filas) con el **Factor de Frecuencia (FF)** (Columnas).

### 5.4 Clasificación
- **BAJA:** NC ≤ 30
- **MEDIA:** NC 31 – 74
- **ALTA:** NC > 74

---

## 6) Matriz dinámica de criticidad
### Dónde está
- Implementación en [SiteAssets/AppAC/js/matrixModal.js](SiteAssets/AppAC/js/matrixModal.js).

### Comportamiento
- Se abre con el ícono en la tarjeta `NC`.
- Renderiza un SVG con grillas, números y la intersección FF/Severidad.
- Muestra línea vertical (FF), línea horizontal (Severidad) y círculo en intersección.

### Parámetros que recibe
- `severity` (1–5)
- `ff` (1–5)

---

## 7) Cobertura
### Fuentes
- Se calcula desde **EquiposAC** según el nivel seleccionado.

### Persistencia
- `EquiposCubiertosJson`
- (Opcionales) `FleetsCubiertosJson`, `ProcesosCubiertosJson`, `EGIsCubiertosJson`

### Uso
- Bloqueo de re‑evaluaciones.
- Filtros en “Mis evaluaciones”.
- KPIs y gráficos.

Importante: KPIs y dashboard consideran solo equipos con `Estado = DI`.

---

## 8) Dashboard

### KPIs
- Flotas, Equipos, Procesos, EGI.
- Indicadores: cubiertos, total, % cubierto, pendientes.
- Cambian de color al alcanzar la meta.

### Gráficas principales del dashboard
El inicio muestra 4 gráficas interactivas, todas filtrables por "Análisis por dimensión" (Gerencia, Superintendencia, UnidadProceso, Fleet, ProcesoSistema, EGI):

1. **Cobertura %**
	- Tipo: Barra horizontal.
	- Muestra el porcentaje de cobertura por la dimensión seleccionada (ej: % cubierto por Gerencia, Superintendencia, etc.).
	- Permite comparar rápidamente el avance de cobertura entre unidades.

2. **Criticidad**
	- Tipo: Gráfico de áreas polares.
	- Segmenta la cantidad de activos por nivel de criticidad: Baja, Media, Alta.
	- Ejemplo: Criticidad Baja: 810, Media: 1035, Alta: 834.

3. **Alta criticidad**
	- Tipo: Barra horizontal.
	- Muestra la cantidad de activos con criticidad alta por la dimensión seleccionada.
	- Permite identificar focos de riesgo en la organización.

4. **Evolución mensual**
	- Tipo: Línea temporal.
	- Muestra la evolución mensual de la cobertura o criticidad a lo largo del tiempo.
	- Permite analizar tendencias y mejoras.

#### Filtros de análisis por dimensión
Todas las gráficas pueden filtrarse seleccionando la dimensión de análisis:
- Gerencia
- Superintendencia
- UnidadProceso
- Fleet
- ProcesoSistema
- EGI

Esto permite comparar y analizar los datos desde diferentes perspectivas organizacionales.

### Leyenda de criticidad
- En una línea, con línea punteada alineada al ancho de la leyenda.

---

## 9) Configuración de meta (% objetivo)
Ubicación: [SiteAssets/AppAC/js/app.js](SiteAssets/AppAC/js/app.js)

Variable:
- `const goalPct = 90;`

Impacta:
- Colores de KPIs.
- Línea objetivo en barras de gerencias.

---

## 10) Listas de SharePoint
### EquiposAC
- `Title` (columna real, referenciada como EquipNo en la app), `NombreEquipo`, `Fleet`, `ProcesoSistema`, `EGI`.
- Usados: `Estado`, `BranchGerencia`.

> Nota: En la lista EquiposAC, la columna real de SharePoint es `Title`, pero en la lógica de la aplicación se usa el alias `EquipNo` para mayor claridad. Todas las referencias a `EquipNo` en el código corresponden realmente a la columna `Title` en SharePoint.

### EscenariosAC
- `Title`, `EventoAnalizado`, `NivelAplicacion`, `Fleet`, `ProcesoSistema`, `EGI`, `EquipNo`, `Activo`.

### EvaluacionesAC
- Target: `TargetTipo`, `TargetValor`, `EquipNo`, `Fleet`, `ProcesoSistema`, `EGI`.
- Métricas: `Smax`, `Impacto`, `NC`, `Clasificacion`.
- Preguntas: `FO`, `FIN`, `HSE`, `MA`, `SOC`, `DDHH`, `REP`, `LEGAL`, `TI`, `FF`.
- Evento: `EventoAnalizadoSnapshot`.
- Cobertura: `EquiposCubiertosCount`, `EquiposCubiertosJson`.
- Auditoría: `EvaluadorId`, `EvaluadoPor`, `FechaEvaluacion`, `EstadoEvaluacion`.

### RolesAC
- `PuedeEditar`, `PuedeAdministrar`.

### SolicitudesEdicionAC
- `Estado`, `FechaSolicitud`, `FechaRespuesta`, `EvaluacionId`, `ScenarioKey`.

---

## 11) API y rendimiento
- REST SharePoint con paginación (`@odata.nextLink`).
- Fallback a Search API si se alcanza umbral de lista.
- Normalización de campos según metadata (coerción de tipos).

---

## 12) Operación y soporte
- Recarga dura con Ctrl+F5 tras cambios.
- Verificar que SweetAlert2 y Chart.js estén cargados.
- Si falla un gráfico: revisar consola, revisar llamadas a datos y permisos.
- Si falla una lista: validar nombre exacto y campos requeridos.

---

## 13) Convenciones y notas
- No se usa NPM ni build pipeline.
- Todo el JS se carga directamente desde SharePoint.
- Los gradientes de sparklines son únicos por tarjeta para evitar conflictos.

---

## 14) Troubleshooting rápido
### Modal no aparece
- Confirmar que SweetAlert2 está cargado.
- Revisar consola por errores de JS.

### Sin datos en KPIs
- Validar permisos de listas.
- Verificar `Estado = DI` en EquiposAC.

### No aparece matriz dinámica
- Revisar que `matrixModal.js` se cargue antes de `app.js`.

---

## 15) Historial
Ver [HISTORIAL.md](HISTORIAL.md).

---

## 16) Consumo de listas, columnas lógicas y expansión

### ¿Cómo se consumen las listas?
La aplicación utiliza JavaScript puro para consumir listas de SharePoint vía REST API. El cliente está en `SiteAssets/AppAC/js/spHttp.js` y los servicios de negocio en `app.js`, `catalogService.js` y `scenarioService.js`.

Ejemplo de consumo:
```js
fetch('/_api/web/lists/getbytitle("EvaluacionesAC")/items')
	.then(r => r.json())
	.then(data => {/* procesamiento */});
```

Las llamadas se hacen con paginación automática y fallback a Search API si la lista supera el umbral de items.
---

## 17) Alimentación y procesamiento de tablas

### ¿Cómo se alimentan las tablas?
Las tablas principales de la aplicación se alimentan dinámicamente mediante peticiones GET a la API REST de SharePoint. El proceso general es:

1. **Obtención de datos (GET):**
	- Se usa `fetch` o funciones utilitarias en `spHttp.js` para obtener los items de la lista deseada.
	- Ejemplo:
	  ```js
    // Nota: 'EquipNo' es el nombre lógico usado en la app, pero la columna real en SharePoint es 'Title'.
    fetch('/_api/web/lists/getbytitle("EvaluacionesAC")/items?$select=ID,TargetTipo,TargetValor,EquipNo,NC,Clasificacion,FechaEvaluacion,EquipNo/Title&$expand=EquipNo')
	  ```
	- El parámetro `$select` define los campos a traer, y `$expand` permite hacer lookups (joins) con otras listas.

2. **Procesamiento de lookups (joins):**
    - Cuando una columna es un lookup (por ejemplo, `EquipNo` en EvaluacionesAC apunta a EquiposAC), se usa `$expand` para traer los datos relacionados en una sola consulta.
    - El resultado es un objeto anidado, por ejemplo: `item.EquipNo.Title` (donde Title es la columna real en SharePoint, referenciada como EquipNo en la app).
    - Si se requieren varios niveles de join, se pueden anidar `$expand` y `$select`.

3. **Normalización y mapeo:**
	- Los datos crudos se procesan en `app.js` para normalizar tipos (fechas, números, textos) y mapear los campos lógicos a los nombres usados en la UI.
	- Se generan arrays de objetos planos para renderizar fácilmente en tablas.

4. **Renderizado dinámico:**
	- Las tablas se generan recorriendo los arrays de datos y creando filas HTML dinámicamente.
	- Las columnas visibles se definen en arrays lógicos, lo que permite modificar la estructura de la tabla fácilmente.

### Consideraciones clave
- **Paginación:** Si la lista es grande, se sigue el enlace `@odata.nextLink` para traer todos los datos.
- **Filtros:** Los filtros aplicados en la UI se traducen a parámetros `$filter` en la consulta REST para optimizar la carga.
- **Ordenamiento:** Se puede usar `$orderby` para traer los datos ya ordenados desde SharePoint.
- **Errores y validaciones:** Si falta un campo o lookup, se maneja con validaciones y mensajes claros en la UI.

### Ejemplo de GET con lookup y filtro
```js
fetch('/_api/web/lists/getbytitle("EvaluacionesAC")/items?$select=ID,TargetTipo,TargetValor,EquipNo/NombreEquipo&$expand=EquipNo&$filter=Clasificacion eq 'Alta'')
```
Esto trae solo las evaluaciones de criticidad alta, incluyendo el nombre del equipo relacionado.

### Resumen
El sistema aprovecha al máximo la API REST de SharePoint para traer, unir y procesar datos de múltiples listas, permitiendo tablas ricas y filtrables sin recargar la página ni hacer múltiples consultas innecesarias.

### Nombres lógicos de columnas
Cada lista tiene columnas lógicas que se usan en la app. Por ejemplo:
- **EvaluacionesAC**: TargetTipo, TargetValor, EquipNo, Fleet, ProcesoSistema, EGI, Smax, Impacto, NC, Clasificacion, FO, FIN, HSE, MA, SOC, DDHH, REP, LEGAL, TI, FF, EventoAnalizadoSnapshot, EquiposCubiertosCount, EquiposCubiertosJson, EvaluadorId, EvaluadoPor, FechaEvaluacion, EstadoEvaluacion.
- **EquiposAC**: Title (EquipNo), NombreEquipo, Fleet, ProcesoSistema, EGI, Estado, BranchGerencia.
- **EscenariosAC**: Title, EventoAnalizado, NivelAplicacion, Fleet, ProcesoSistema, EGI, EquipNo, Activo.

### ¿Cómo se renderizan las tablas?
Las tablas se generan dinámicamente en `app.js` según el tipo de lista. Las columnas visibles se definen en arrays lógicos y pueden modificarse fácilmente para agregar/quitar campos.

Ejemplo:
```js
const columnasEvaluaciones = ["TargetTipo", "TargetValor", "NC", "Clasificacion", "FechaEvaluacion"];
```

### ¿Cómo funcionan los filtros?
Los filtros se renderizan como selects, usando valores únicos de columnas clave (Criticidad, Estado, Responsable, TipoActivo, etc.). El botón de limpiar filtros está superpuesto en la esquina del panel de filtros.

### ¿Cómo expandir o modificar?
- Para agregar nuevas columnas: crear el campo en la lista SharePoint y agregarlo al array de columnas en el JS.
- Para nuevos filtros: agregar el campo al array de filtros y actualizar el renderizado.
- Para nuevos gráficos: agregar el procesamiento en `app.js` y el plugin Chart.js.
- Los estilos se ajustan en `style.css`.

### Guía rápida para nuevos desarrolladores
1. Revisar la estructura de archivos y los servicios JS.
2. Identificar las listas y columnas lógicas usadas.
3. Modificar arrays de columnas/filtros según necesidad.
4. Probar cambios con recarga dura (Ctrl+F5).
5. Consultar el historial en HISTORIAL.md para ver cambios previos.

---
Esta sección resume cómo se consumen las listas, el mapeo de columnas lógicas y cómo expandir la app para que cualquier persona pueda entender y mantener el desarrollo.

---

## 🔄 Actualización: Filtros Jerárquicos y Rediseño UI (Febrero 10 2026)

Se ha implementado una nueva capa de filtrado organizacional para facilitar la búsqueda de activos en bases de datos extensas.

### 1. Nuevas Funcionalidades
- **Filtros Jerárquicos Superiores:** Se agregaron 3 nuevos selectores en el Wizard:
  - `Gerencia` (Filtra Superintendencias).
  - `Superintendencia` (Filtra Unidades).
  - `Unidad de Proceso` (Filtra el universo de equipos).
- **Persistencia Visual:** Los filtros seleccionados en el **Paso 1** ahora viajan y se mantienen visibles/editables en el **Paso 2**.
- **Reactividad en Tiempo Real:** Al modificar una Gerencia o Superintendencia en el Paso 2, la tabla de "Equipos a Evaluar" y los contadores de cobertura se actualizan instantáneamente sin recargar.

### 2. Cambios en Lógica de Negocio (Core)
- **Lógica de Intersección Estricta:**
  El cálculo de equipos disponibles ahora obedece a la fórmula:
  `Equipos = (Filtros Jerárquicos [Gerencia+Super+Unidad]) AND (Filtros de Nivel [Fleet+Proceso+EGI])`
  *Antes:* Los filtros de nivel ignoraban la selección de Gerencia en el conteo final.
  *Ahora:* Se garantiza consistencia total entre lo que se filtra y lo que se cuenta.

- **Saneamiento de Consultas (Fix "Parameter Pollution"):**
  Se corrigió un bug donde la búsqueda de Escenarios fallaba al recibir campos de jerarquía (`BranchGerencia`) que no existen en la lista `EscenariosAC`. Ahora el sistema "limpia" el objeto de consulta antes de llamar a la API de SharePoint, enviando solo los niveles tácticos (`Fleet`, `EGI`, `Proceso`).

### 3. Mejoras de Interfaz (UI/UX)
- **Rediseño del Paso 1:** Reorganización del layout para alinear el botón "Continuar" con el selector de Nivel, y agrupación de los filtros jerárquicos en una fila independiente de 3 columnas.
- **Integración Visual en Paso 2:** Los filtros persistentes se integraron nativamente en el formulario de selección de target, eliminando contenedores redundantes y mejorando la limpieza visual.
