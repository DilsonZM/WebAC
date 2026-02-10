# Historial del proyecto WebAC

## Propósito
Aplicación SPA en SharePoint (sin NPM) para evaluar criticidad de activos con un wizard por niveles, paneles de control y administración, cálculos de riesgo y persistencia en listas de SharePoint.

## Estado actual (2026-02-05)
- Wizard completo por niveles (Fleet/ProcesoSistema/EGI/Equipo) con filtros y resumen de target.
- Selección de escenario con paginación y búsqueda (manejo de umbral con Search API fallback).
- Paso de preguntas con resumen de cobertura (matriz 3x3) y evento editable.
- Cálculo de métricas (Smax, Impacto, NC, Clasificación) y resumen final.
- Guardado robusto con coerción de tipos, metadata y campos dinámicos.
- Vista de “Ver evaluación” con detalle y tabla de equipos paginada con filtros.
- Paneles de Dashboard/Mis/Admin con tablas y paginación.
- Bloqueo de re‑evaluaciones por cobertura y acceso a detalle de evaluación previa.

## Cambios recientes (2026-02-05)
### UI / UX
- Rediseño completo del dashboard con **sidebar** y **topbar** moderna, tarjeta de usuario, grillas responsivas y KPIs tipo “cards”.
- KPIs con sparklines difuminados (solo área, sin línea visible) y animación de contadores.
- Se eliminó el fondo degradado en KPIs y se pulió la estética de bordes/sombras.
- Overlay de carga global para evitar parpadeos y estados “rojos” al inicio o al refrescar.

### KPIs y lógica de cobertura
- Cobertura calculada desde `EvaluacionesAC` usando JSON (`EquiposCubiertosJson`) + catálogo `EquiposAC`.
- Totales y cobertura **solo para DI** (`Estado = DI`).
- KPIs muestran: cubiertos, total, porcentaje, y pendientes por Fleet/Proceso/EGI/Equipo.
- Se ajustaron los contadores de “hoy/7d/mes” para evaluaciones recientes.

### Gráficas (Chart.js)
- **Gerencias:** barra horizontal con cobertura (%) por `BranchGerencia`, línea objetivo al 90% y tooltips con cubiertos/total.
- **Criticidad:** polar area por `Clasificacion` (Baja/Media/Alta), etiquetas de % sobre el gráfico y leyenda con conteos.

### Rendimiento / umbral de listas
- Carga de `EquiposAC` con paginación por `Id` para evitar umbral.
- Fallback por Search API con `ListId` y mapeo de campos `ows_` (incluye `Estado` y `BranchGerencia`).

### Estabilidad en refresh
- El loader se activa al refrescar y se desactiva cuando finaliza la carga de datos y gráficos.
- Protección con `try/finally` para asegurar cierre del loader incluso con errores.

---

## Arquitectura funcional
### 1) Flujo principal (Wizard)
1. **Nivel (Paso 1):** Se define el nivel a evaluar (Fleet/ProcesoSistema/EGI/Equipo).
2. **Selección (Paso 2):** Filtros jerárquicos y selección de equipo (si aplica). Se valida si ya existe evaluación para evitar duplicados.
3. **Escenario (Paso 3):** Selección de escenario con paginación y búsqueda.
4. **Preguntas (Paso 4):** 10 preguntas de impacto + evento editable + justificación obligatoria.
5. **Resumen (Paso 5):** Resumen de target, cobertura, escenario, evento y métricas; se confirma el envío.
6. **Confirmación (Paso 6):** Mensaje de guardado exitoso.

### 2) Paneles
- **Dashboard:** KPIs y listado de últimas evaluaciones.
- **Mis evaluaciones:** Filtros (fecha, equipo, Fleet, Proceso, EGI) usando cobertura real; tablas paginadas.
- **Admin:** CRUD de escenarios y lista de evaluaciones.

---

## Listas de SharePoint y relación entre datos
### Listas principales
1. **EquiposAC**
	 - Catálogo jerárquico: `Title` (EquipNo), `NombreEquipo`, `Fleet`, `ProcesoSistema`, `EGI`.
	 - Usado para filtrar, sugerir equipos y calcular coberturas.

2. **EscenariosAC**
	 - Escenarios por nivel: `Title`, `EventoAnalizado`, `NivelAplicacion`, `Fleet`, `ProcesoSistema`, `EGI`, `EquipNo`, `Activo`.
	 - Alimenta el paso 3 (selección de escenario).

3. **EvaluacionesAC**
	 - Resultado de evaluación. Campos clave:
		 - Target: `TargetTipo`, `TargetValor`, `EquipNo`, `Fleet`, `ProcesoSistema`, `EGI`
		 - Métricas: `Smax`, `Impacto`, `NC`, `Clasificacion`
		 - Preguntas: `FO`, `FIN`, `HSE`, `MA`, `SOC`, `DDHH`, `REP`, `LEGAL`, `TI`, `FF` (todas tipo Número)
		 - Evento: `EventoAnalizadoSnapshot`
		 - Cobertura: `EquiposCubiertosCount`, `EquiposCubiertosJson` (y opcionales Fleets/Procesos/EGIs JSON)
		 - Auditoría: `EvaluadorId`, `EvaluadoPor`, `FechaEvaluacion`, `EstadoEvaluacion`

### Relación y alimentación
- **EquiposAC →** define jerarquía (Fleet/Proceso/EGI/Equipo).
- **EscenariosAC →** se filtra por nivel y target del wizard.
- **EvaluacionesAC →** se guarda con snapshot de escenario + cobertura calculada desde EquiposAC.
- **Cobertura JSON** se usa para:
	- bloquear re‑evaluaciones por cobertura completa,
	- filtrar “Mis evaluaciones” por lo realmente evaluado,
	- mostrar equipos asociados en la vista de detalle.

---

## Cálculos y reglas de negocio
### 1) Métricas
- **Smax:** máximo entre FIN, HSE, MA, SOC, DDHH, REP, LEGAL, TI.
- **Impacto:** $\text{Smax} \times \text{FO}$.
- **NC:** $\text{Impacto} \times \text{FF}$.
- **Criticidad:** Baja si NC ≤ 30; Media si NC ≤ 74; Alta si NC > 74.

### 2) Cobertura
- Se calcula en base al nivel seleccionado:
	- **Equipo:** 1 equipo.
	- **Fleet/Proceso/EGI:** equipos filtrados en EquiposAC.
- Se genera JSON con equipos y sets de Fleets/Procesos/EGIs.
- Se persiste en `EquiposCubiertosJson` y opcionales JSON extra.

### 3) Re‑evaluación
- Si la cobertura ya fue evaluada, se bloquea el flujo y se permite “Ver evaluación”.
- Se usa JSON para determinar cobertura completa.

---

## Cómo se obtiene la data
- **REST SharePoint:** `_api/web/lists/GetByTitle(...)/items`.
- **Paginación:** manejo de `@odata.nextLink`.
- **Umbral:** fallback a Search API cuando la lista supera el umbral.
- **Campos dinámicos:** lectura de metadatos con `ListItemEntityTypeFullName` y mapa de campos.
- **Coerción de tipos:** conversión por tipo antes de guardar.

---

## Estructura de archivos y responsabilidades
### Vistas/Markup
- [home.aspx](home.aspx)
	- Página principal (SharePoint). Contiene el HTML del SPA.
- [SiteAssets/AppAC/index.html](SiteAssets/AppAC/index.html)
	- Copia del markup del SPA (referencial/modular).

### Lógica JS
- [SiteAssets/AppAC/js/app.js](SiteAssets/AppAC/js/app.js)
	- Wizard completo, filtros, paginación, cálculos, guardado y vistas.
- [SiteAssets/AppAC/js/spHttp.js](SiteAssets/AppAC/js/spHttp.js)
	- Cliente REST, paginación, metadata, coerción de payload, POST/MERGE.
- [SiteAssets/AppAC/js/scenarioService.js](SiteAssets/AppAC/js/scenarioService.js)
	- Carga de escenarios con fallback para umbral.
- [SiteAssets/AppAC/js/catalogService.js](SiteAssets/AppAC/js/catalogService.js)
	- Carga de catálogo de equipos.

### Estilos
- [SiteAssets/AppAC/css/style.css](SiteAssets/AppAC/css/style.css)
	- Layout, tarjetas, tablas, stepper, botones, filtros y grids.

---

## Funcionalidades clave implementadas
- Wizard con pasos y validaciones.
- Filtros jerárquicos (Fleet → Proceso → EGI).
- Paginación de escenarios, equipos y tablas de evaluaciones.
- Vista de evaluación con detalle y tabla de equipos con filtros.
- Ajustes de UX: alineaciones, botones arriba, resumen compacto.
- Evento analizado editable en paso 4 y persistencia en snapshot.
- Mis evaluaciones filtradas por cobertura real (JSON).

---

## Herramientas y técnicas usadas
- **SharePoint REST API** para lectura/escritura.
- **Search API** como fallback ante umbral.
- **OData** para filtros, selección y expansión de campos.
- **Coerción de tipos** según metadata de lista.
- **JSON de cobertura** para deduplicación y filtros.

---

## Pendiente / mejoras sugeridas
- Mejorar mensajes cuando no hay resultados.
- Revisar rendimiento con listas grandes (Search API).
- Validar UX de filtros autoajustables en “Mis evaluaciones”.
