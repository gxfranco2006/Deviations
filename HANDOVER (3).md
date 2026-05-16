# MWAAF Deviation Authorization System — Handover

**Última actualización**: 9 de mayo, 2026
**Cuenta dueña actual**: mwaaf.deviations.noreply@gmail.com
**Cliente final**: Acoustafiber / Gustavo Franco (mwaaf.com)
**Stack**: Google Apps Script (backend) + HTML/JS standalone (frontend) + Google Sheets (database)
**HTML hosteado en**: https://distribution.fyware.com/MWAAF/deviationsV1
**Sheet (database)**: https://docs.google.com/spreadsheets/d/1mWzMNFudiiuLay3wY2-CvUP0ej4C2EKbcuYxBQrKWk8/edit

---

## 1. Propósito del sistema

Sistema interno de Acoustafiber para crear, gestionar y autorizar Deviations (desviaciones de proceso de manufactura, formulario QF158). Reemplaza un proceso anterior en papel/Excel.

Las deviations son solicitudes formales para producir piezas que se desvían de las especificaciones de ingeniería estándar. Requieren firma de varios stakeholders (calidad, producción, ingeniería, materiales, plant manager) antes de que la planta pueda producirlas.

---

## 2. Arquitectura

```
┌─────────────────────────────────┐
│  deviation-app-sheets.html      │  ← Hosteada en distribution.fyware.com/MWAAF/deviationsV1
│  (frontend, ~5000 líneas)       │     Bloqueada con password al abrir
└────────────┬────────────────────┘
             │ fetch POST (JSON)
             ▼
┌─────────────────────────────────┐
│  Apps Script Web App (Code.gs)  │  ← URL pública desplegada en cuenta
│  (~3500 líneas)                 │     mwaaf.deviations.noreply@gmail.com
│                                 │     Maneja: deviations, approvals,
│                                 │     emails, tokens, settings, quota
└────────────┬────────────────────┘
             │ SpreadsheetApp / GmailApp
             ▼
┌─────────────────────────────────┐
│  Google Sheet (database)        │
│  Hojas: Deviations, Approvals,  │
│  Approvers, DistLists, Config,  │
│  ApprovalTokens, RoleConfig,    │
│  ReasonOptions, PartNumbers,    │
│  WorkCenters                    │
└─────────────────────────────────┘
             │
             │ + emails enviados desde
             ▼
       mwaaf.deviations.noreply@gmail.com
```

### Flujo de un deviation

1. Iniciador abre la app, llena el form, sube fotos como evidencia
2. Submit → POST a Apps Script → guarda en hoja Deviations
3. Apps Script envía emails a cada approver con un link único con token
4. Approver hace click en el link del email → abre la vista del approver (servida por Apps Script con `?token=...`)
5. Approver firma (approve/reject) con comentarios opcionales
6. Apps Script recalcula el status global de la deviation y, si está completa o rechazada, envía email final solo a la Distribution List "approval"

### Flujo de edición (re-aprobación)

1. Usuario edita una deviation existente (cualquier status: pending, partial, rejected, approved)
2. Submit → backend detecta que es edit (no draft)
3. Backend borra todas las Approvals previas de esa deviation
4. Backend resetea status a "pending" e incrementa el campo "revision" (Rev.1 → Rev.2)
5. Frontend dispara sendNotification que re-lee approvals frescas del Sheet
6. Todos los approvers seleccionados reciben correo nuevo con un nuevo token
7. El dashboard muestra inmediatamente "0/N approvals" y la nueva revisión

---

## 3. Funcionalidades principales

### Crear Deviation Request
- Campos: devNum (auto-incremental), part numbers afectados, fechas (start/end), reason, risk factor (Low/Med/High), action plan, comments, fotos (max 6, max 5MB cada una)
- Catálogos: Part Numbers (con descripción ligada) y Work Centers — autocompletan via datalist
- Reasons: catálogo editable en Settings, con tags 4M y other

### Sistema de aprobación multi-firma
Tres tipos de approvers:
- 🔓 Optional: recibe email pero no bloquea ni cuenta para required
- 🔒 Required: debe firmar para que la desviación se considere aprobada
- ⚡ Rule-Out: UNA sola firma de cualquier ruleout aprueba toda la desviación (path rápido). Sobrescribe rejections previas.

Reglas de status (en updateDeviationStatus, prioridad de mayor a menor):
1. Ruleout APPROVED → approved (gana sobre rejection) + email a dist list "approval"
2. Todos los required APPROVED → approved + email a dist list "approval"
3. Algún REQUIRED rejected → rejected + email a dist list "approval"
4. Optional rejected → ignorado para status (solo se loggea)
5. Algún required approved (no todos) → partial
6. Default → pending

### Vista del approver (sin login)
- Servida por doGet con `?token=...`
- El token es un UUID (128 bits aleatorios), válido 7 días, single-use
- Muestra todas las deviations donde el approver tiene voz
- Sidebar con secciones: "Awaiting Your Decision" / "Already Decided" / "Other Pending"
- Search bar filtra el sidebar en vivo
- Al firmar: el card se marca como decidido pero no se elimina; sidebar item se mueve a "Already Decided" en tiempo real

### Vista read-only de deviation (sin login)
- Servida por doGet con `?view=DEV-XXXX`
- Linkeada desde los emails de "approval complete", "rejected" y "new deviation submitted"
- Muestra meta, parts, reason, action plan, fotos (con lightbox), tabla completa de approvals
- No expone acciones, solo lectura

### Dashboard
- Tabla desktop / cards mobile
- Filtros por: status, part number, work center, fecha (from/to)
- Filtros con debounce (180ms) para no lagear con 500+ deviations
- Búsqueda y paginación cliente
- Acciones por row: View, Edit, Renotify (reenviar emails de approval), Delete
- Renotify SOLO manda correo a approvers que aún no han decidido
- Delete requiere password admin (mismo que Settings, configurado en `Config!settingsPassword`). Una vez ingresado, se recuerda durante la sesión actual (hasta recargar la página) para no fastidiar al usuario en limpiezas masivas. Si `settingsPassword` está vacío, se omite y vuelve a un `confirm()` simple

### Print
- Reproduce la vista de detalle (no formato QF158 oficial)
- Watermark "DRAFT — PENDING APPROVAL" en cada página cuando la deviation no está approved
- El watermark se inyecta dinámicamente: JS calcula cuántas páginas tendrá el contenido y crea N divs con position absolute cada uno cubriendo una página

### Settings (con password gate)
Pestañas: A. Approvers, B. Distribution Lists, D. Document Configuration, E. Reason Options, F. Part Numbers, G. Work Centers
- Password configurable en `Config!settingsPassword`. Si está vacío → no pide password
- El modal de password se abre instantáneo (cache del estado del lock al boot)
- Botón "📊 Sheet" en la barra superior del view de Settings: abre el Google Sheet en otra pestaña. Solo visible después de pasar el password de admin. La URL del Sheet se inyecta automáticamente desde el backend (`getConfigPublic` retorna `sheetUrl` con `SpreadsheetApp.getActiveSpreadsheet().getUrl()`), así que si en el futuro el script se vincula a otro Sheet, el link se actualiza solo

### Part Numbers Catalog (paginación + búsqueda)
La pestaña F (Part Numbers) está optimizada para catálogos grandes (1000+ items):
- Barra de búsqueda en vivo arriba del listado, busca en partNumber y description (case-insensitive)
- Paginación client-side de 25 items por página
- Controles: « (primera) ‹ (anterior) [N] › (siguiente) » (última), con resaltado de página actual y "..." cuando hay muchas páginas
- Contador "X of Y match" o "Y total" en la barra de búsqueda
- Al agregar un nuevo part number: limpia la búsqueda activa y salta a la última página para mostrar el item recién agregado
- Delete/edit: los índices se mapean al array original (DB.partNumbers), no al filtrado/paginado, así borras el item correcto aunque estés con filtro activo en página 5

### App access password
- Modal splash al abrir el HTML
- Password configurable en `Config!appAccessPassword`. Si está vacío → omite el splash
- Pide cada vez que se abre (no recuerda)
- Backend valida con verifyAppAccess endpoint, nunca expone el password al cliente

### Footer con Email Quota
- Barra visual a la derecha del footer en index
- Muestra cuántos correos de la cuota diaria de Gmail se han usado
- Heurística: si remaining > 100 asume Workspace (1500/día), si no consumer (100/día)
- Se actualiza al cargar la app, al hacer click en el botón refresh ↻, y automáticamente después de mandar correos
- Color verde por default, amarillo a partir del 70% usado, rojo a partir del 90%

---

## 4. Estructura del Sheet (database)

| Hoja | Headers | Notas |
|------|---------|-------|
| Deviations | id, devNum, mainPartNum, date, shift, initiator, workCenter, custApproval, startDate, endDate, description, parts, reasons, reasonLabel, tags, fourm, riskFactor, actionPlan, comments, owner, otherReason, status, submittedAt, revision, photos, selectedApprovers | parts/photos/etc son JSON stringified |
| Approvals | id, deviationId, approverId, approverName, approverRole, decision, date, comments | Una fila por firma. Se borran completas al editar una deviation |
| Approvers | id, name, email, role, required, defaultStatus | defaultStatus: required/optional/ruleout |
| DistLists | listType, email | listType: creation (FYI al crear) / approval (notificación de approved/rejected) |
| Config | key, value | Ver tabla abajo |
| ApprovalTokens | token, approverId, deviationIds, createdAt, expiresAt, used | Tokens single-use |
| RoleConfig | roleKey, required | Legacy, ya no se usa mucho |
| ReasonOptions | id, label, tags | tags array JSON |
| PartNumbers | partNumber, description | 2 columnas, force text format con prefix |
| WorkCenters | workCenter | 1 columna |

### Config keys importantes

| Key | Default | Notas |
|-----|---------|-------|
| docOwner | Director of Quality | Mostrado en headers |
| approvedByTitle | Director of Operations | |
| draftValidHours | 12 | Solo informativo (no auto-expira deviations) |
| settingsPassword | mwaaf2024 | Vacío = no pide password en Settings |
| appAccessPassword | Mwaaf01 | Vacío = no pide password al abrir HTML |
| nextDevNum | 1001 | Counter persistente, usado por generateNextDevNum() con LockService |
| appUrl | (auto) | URL del web app, se setea desde menú "Set Web App URL" |

---

## 5. Decisiones técnicas importantes

### Por qué Apps Script + standalone HTML
- Cliente quería algo simple, sin servidor propio
- Apps Script da: backend gratis, integración nativa con Sheets, emails, autenticación opcional
- HTML standalone: ahora hosteado en distribution.fyware.com/MWAAF/deviations, evita CORS al hacer fetch al web app de Apps Script (origen del fetch es controlable)

### Por qué tokens en lugar de login para approvers
- Los approvers son operarios/managers que rara vez usan el sistema
- Pedirles password es fricción. Un link único en el email es estándar (DocuSign, etc.)
- El token es 128 bits aleatorios, se invalida al usarse, expira en 7 días

### Por qué se almacenan fotos como base64 en el Sheet
- Simplicidad: una sola fuente de verdad
- Limitación: las celdas tienen 50K chars, así que cada foto está en una celda de array JSON. Mitigación: el frontend comprime fotos a max 1600px lado largo + JPEG 0.75 antes de subir. Resultado: ~150-400KB por foto vs los 2-5MB originales

### Lazy loading de fotos (crítico para performance con 500+ deviations)
- getDeviationsLite retorna deviations SIN dataURLs de fotos (solo `_photoCount`)
- El dashboard usa getDeviationsLite, las fotos se cargan on-demand cuando el usuario abre una deviation
- Endpoint dedicado: getDeviationPhotos({id}) solo trae las fotos de UNA deviation
- Sin esto: 500 deviations x 5 fotos x 2MB = 2.5GB en cada page load, timeout garantizado

### Auto-creación de Part Numbers
- Cuando alguien crea una deviation con un Part nuevo (no en catálogo), se agrega automáticamente al catálogo con la partName como description
- Endpoint: upsertPartNumber({partNumber, description})
- Solo agrega si no existe; si existe sin descripción y la deviation provee una, la actualiza

### DevNum counter persistente con LockService
- generateNextDevNum() lee Config!nextDevNum, lo incrementa, escribe de vuelta
- Wrapped en LockService.getScriptLock() con timeout 10s para evitar race conditions cuando 2 personas hacen submit a la vez
- Hay función repairDevNums() en menú para arreglar duplicados si pasan

### Forzar text format en celdas
- Sheets autoconvierte strings que parecen números/fechas (e.g. "44521" se vuelve número, "2024-01-15" se vuelve Date)
- Solución: prefijar con apóstrofe al escribir y setNumberFormat('@') en columnas

### Headers tolerantes en getPartNumbers / getWorkCenters
- Las funciones del backend que leen los catálogos aceptan variantes en el nombre del header (case-insensitive, ignora espacios, guiones, numerales)
- Para Part Numbers acepta: partNumber, partNo, partNum, part, sku, itemNumber, itemNo
- Para Work Centers acepta: workCenter, wc, area, workArea, center, line
- Si no encuentra nada reconocido, hace fallback a la primera columna y deja un log

### Re-lectura de approvals desde Sheet en sendNotification
- sendNotification ya no confía en el payload del frontend para decidir a quién skipear
- Cada vez que se llama, consulta la hoja Approvals para saber quién decidió en el ciclo actual
- Esto garantiza que después de un edit (donde las approvals se borraron), TODOS los approvers seleccionados reciben correo nuevo

---

## 6. Seguridad implementada

### XSS prevention
- Helpers globales: escHtml, escAttr, escNl, escJsAttr en frontend
- Backend: escapeHtml_, escapeHtmlNl_ en Code.gs
- Aplicado en: buildDetailHTML, dashboard table desktop+mobile, todos los renderXxx (Approvers, DistList, ReasonOptions), approveModalInfo, detailMeta, print header, todos los email builders
- Photo dataUrls validados con regex `^data:image/(png|jpe?g|gif|webp);base64,`
- openImageLightbox recibe índice (no dataUrl inline) para evitar inyección

### Input validation en backend
- saveDeviation: text fields cap 10K, long text 30K, max 200 partes, max 20 fotos, max 5MB foto, whitelist status y riskFactor
- submitApproval: whitelist decision (approved/rejected), comments cap 5K, validación formato IDs

### Authentication / Authorization
- Approver view requiere token válido + el approverId del token debe coincidir con la acción
- Tokens son UUIDs (128-bit), single-use, expiran 7 días
- App access password y Settings password validados contra backend (no client-side)
- Backend strip de passwords en getConfigPublic() antes de enviar al cliente

### Limitaciones conocidas (NO resueltas)
- Sin rate limiting nativo en Apps Script
- Email quota: 100/día consumer (cuenta actual mwaaf.deviations.noreply), 1500/día si se migra a Workspace
- Sin captcha en el web app

---

## 7. Endpoints del Apps Script (doPost)

Acción se manda en payload como `{action: "...", ...}`. Lista completa:

**Lectura**:
- getDeviations(filter) — todas las deviations con fotos completas (lento)
- getDeviationsLite(filter) — sin fotos, para dashboard
- getDeviationPhotos({id}) — solo fotos de UNA deviation
- getConfig — config + approvers + reasons + dists + partNumbers + workCenters (passwords stripped)
- getApprovers, getReasonOptions, getDistLists, getPartNumbers, getWorkCenters, getRoleConfig

**Escritura**:
- saveDeviation(dev) — upsert. Si es edit, borra approvals previas, resetea status a pending, bumpea revision. Retorna `{id, devNum, isNew, approvalsCleared, newRevision}`
- deleteDeviation({id})
- submitApproval({deviationId, approverId, decision, comments, approverToken}) — retorna `{record, statusBefore, statusAfter, myStatus, info}`
- submitApprovalFromView — wrapper para google.script.run
- saveApprovers, saveReasonOptions, saveDistLists, savePartNumbers, upsertPartNumber, saveWorkCenters, saveConfig, saveRoleConfig

**Notificaciones**:
- sendNotification({devId, type}) — type: creation / approval / rejected
- getEmailQuota() — retorna `{remaining, limit, used}` para la barra del footer

**Auth**:
- verifyAppAccess({password}) — retorna `{ok: true/false}`. Si appAccessPassword vacío → ok:true
- getAppLockStatus() — retorna `{enabled: bool}`
- verifySettingsPassword({password}) — igual lógica con settingsPassword
- getSettingsLockStatus() — igual

### Endpoints GET (doGet)
- Sin params → status page
- `?token=...` → vista del approver
- `?view=DEV-XXXX` → vista read-only de la deviation (no requiere auth)

---

## 8. Histórico de cambios principales

(Ordenado del más reciente al más viejo, alto nivel)

**Sesión actual (mayo 2026)**:
- Password gate al borrar deviations desde el dashboard (reusa `Config!settingsPassword`, cachea autenticación durante la sesión)
- Paginación + búsqueda en el catálogo de Part Numbers (25 por página, filtro por partNumber y description) para soportar catálogos de 1000+ items
- Sistema de re-aprobación al editar: borra Approvals previas, resetea a pending, bumpea revision automáticamente
- sendNotification re-lee approvals del Sheet (no confía en payload del frontend)
- Email de "Approved" ahora se envía SOLO a la Distribution List "approval" (antes a todos los approvers)
- Email de "Rejected" agregado, también solo a la Distribution List "approval", con detalle de quién rechazó y comentarios
- Footer con barra de quota de Gmail (verde/amarillo/rojo según uso)
- Settings password modal abre instantáneo (cache del lock status al boot, antes había roundtrip)
- getPartNumbers y getWorkCenters ahora son tolerantes a variantes en el nombre del header
- Migración de owner del Apps Script + Sheet a mwaaf.deviations.noreply@gmail.com
- HTML hosteado en distribution.fyware.com/MWAAF/deviationsV1 en lugar de archivo local
- DEFAULT_APPS_SCRIPT_URL hardcodeada apuntando al deployment de mwaaf.deviations.noreply (ya no requiere configuración manual via debug panel)
- Link al Sheet movido del footer público al view de Settings (solo visible para admin después del password gate). URL del Sheet inyectada automáticamente por el backend

**Sesiones anteriores**:
- Performance optimization para 500+ deviations (compresión de fotos, lazy loading, debounce filtros)
- Lightbox en read-only view: fotos clickeables con _PHOTOS array embebido
- Botón "Sign In" eliminado del detail view (legacy, no se usa)
- Password de settings y app-access desde Sheet (Config!settingsPassword, Config!appAccessPassword)
- App access password splash al abrir HTML
- Read-only deviation view con `?view=DEV-XXXX` linkeada desde emails
- Part Numbers con descripción ligada + auto-fill en form + auto-upsert al guardar
- Aprobación logic refinada: ruleout-approved gana sobre rejection, optional rejection no bloquea
- Refresh instantáneo en approver view: cards se transforman en lugar de eliminarse
- Search bar en sidebar del approver view
- Watermark "DRAFT" en todas las páginas del print
- Print usa la vista detail (antes era formato QF158 oficial)
- Auditoría de seguridad XSS completa (escapes en todos los renderHtml)
- Validación de input en backend (size limits, whitelists)

---

## 9. Estructura de archivos

```
/Code.gs                       (~3500 líneas, single file Apps Script)
/deviation-app-sheets.html     (~5000 líneas, single HTML standalone)
```

Todo está en un solo archivo cada uno por simplicidad de deployment.

### Code.gs - secciones principales (en orden)
1. Constants (SH = sheet names, headers)
2. doGet, doPost, handlers map
3. ID/token generators
4. saveDeviation (con lógica de edit/re-trigger), getDeviations*, updateDeviationStatus
5. submitApproval, loginApprover (legacy)
6. updateDeviationStatus con priority rules
7. ApprovalTokens (create, validate, markUsed)
8. Approvers / Reasons / DistLists / Config CRUDs
9. PartNumbers / WorkCenters (con headers tolerantes)
10. App access endpoints (verifyAppAccess, etc)
11. Email builders (buildCreationEmailHtml, buildCreationDistEmailHtml, buildApprovalEmailHtml, buildRejectedEmailHtml, emailWrapper)
12. sendNotification (orchestrator) con re-lectura de approvals desde Sheet
13. getEmailQuota
14. serveApproverView, buildApproverViewHtml
15. serveReadOnlyDeviationView, buildReadOnlyDeviationHtml
16. setupSheets (ensureSheet, seeds)
17. setWebAppUrl
18. Menu functions (onOpen, repairDevNums, runDiagnosticEmails)

### deviation-app-sheets.html - secciones principales
1. CSS (incluye @media print, footer-quota, settings-grid)
2. App lock overlay (splash)
3. Header
4. Views: home (dashboard con barra de quota en footer), new (form), detail, settings
5. Modals: approval, login, settings password, image lightbox
6. Footer con barra de quota + debug panel
7. JS: helpers, escapes, _formatDate, apiCall, app-lock logic, bootApp, dashboard, form, photos, settings (con cache de lock), approver views, print, refreshEmailQuota, _prefetchSettingsLockStatus

---

## 10. Deploy y mantenimiento

### Deploy de cambios al Code.gs
1. Apps Script editor → pegar nuevo contenido → Save
2. Deploy → Manage deployments → Edit (lápiz) → **Version: New version** → Deploy
3. La URL del web app NO cambia entre deploys
4. **CRÍTICO**: solo cambiar el nombre del deployment NO publica los cambios. Hay que crear new version cada vez

### Deploy de cambios al HTML
- Subir el nuevo HTML a https://distribution.fyware.com/MWAAF/deviations (reemplazar archivo)
- Los usuarios pueden necesitar F5 + click "🗑 Clear Cache" en debug panel si los catálogos no actualizan

### Configuración inicial de la cuenta dueña actual
1. Cuenta dueña: mwaaf.deviations.noreply@gmail.com
2. Password: sw3etMust@ng
3. Es Gmail consumer (NO Workspace), por lo que la cuota de email es 100/día
4. Para subir a 1500/día habría que migrar la cuenta a Google Workspace (involucra dominio propio)

### Setup desde cero (en otra cuenta)
1. Crear Google Sheet en blanco
2. Extensions → Apps Script → pegar Code.gs → Save
3. Run → setupSheets() (acepta permisos)
4. Deploy → New deployment → Web app → "Anyone" access → Deploy → copia la URL
5. En el editor, abrir setWebAppUrl(), pegar la URL nueva en NEW_URL, Run
6. Edita el HTML (en distribution.fyware.com): variable APPS_SCRIPT_URL con la URL del deployment, o usar el debug panel "Change URL"
7. Abre el HTML → password splash debe pedir el appAccessPassword del Config

### URL actual del web app
La cuenta mwaaf.deviations.noreply tiene su deployment activo. La URL completa está en Config!appUrl del Sheet. Si se hace un new version del deployment, la URL NO cambia, solo se actualiza el código.

### Pendiente de coordinar
- Cuenta de Workspace para 1500 emails/día (en proceso con TI del cliente)
- Si se quiere subir el HTML al dominio del cliente (mwaaf.com) en lugar de distribution.fyware.com, su TI tiene que descargar el archivo y subirlo a su hosting

---

## 11. Cosas que pueden romperse y cómo arreglarlas

| Síntoma | Causa probable | Solución |
|---------|----------------|----------|
| Email quota exhausted | Excedido el límite diario | Esperar 24h. Mover a Workspace para 1500/día. La barra de quota del footer avisa cuando se acerca |
| DevNum duplicado | Race condition antes del LockService | Menú custom → repairDevNums() |
| "Tokens used or expired" al refrescar approver view | Comportamiento esperado: tokens son single-use | Pedir nuevo email vía Renotify desde dashboard |
| Fotos no aparecen en print | Print disparó antes de que el lazy fetch completara | El código ya hace await de getDeviationPhotos antes de imprimir, no debería pasar |
| Settings no entra después de password correcto | _settingsAuthThisTransition no se está reseteando | Verificar en showView que el flag se respeta |
| Status no se actualiza después de aprobación | updateDeviationStatus falló silenciosamente | Revisar logs de ejecución en Apps Script |
| HTML "no carga", se queda en blanco después del password | Error en bootApp o en alguna API call | Abrir DevTools console |
| "Unknown action: XYZ" | El backend desplegado no tiene ese endpoint | Verificar que se hizo Manage deployments → Edit → New version (no solo Save) |
| Catálogos vacíos en Settings | Headers del Sheet tienen formato inesperado | Las funciones get*Numbers/get*Centers ahora son tolerantes; revisar logs si aún falla |
| Links de los correos no funcionan | Config!appUrl o ScriptProperties WEB_APP_URL desactualizados | Editar setWebAppUrl() en Code.gs con la URL correcta y Run |
| Correos no llegan después de migrar de cuenta | La cuenta nueva no aceptó permisos de Gmail | En Apps Script editor, Run cualquier función y aceptar permisos |
| HTML hosteado muestra datos del Sheet/cuenta viejos | DEFAULT_APPS_SCRIPT_URL en línea 1985 del HTML apunta al deployment viejo. Cada navegador/dominio tiene su propio localStorage, así que un visitante nuevo cae al DEFAULT en lugar de tomar la URL de localStorage | Editar línea 1985 del HTML con la URL nueva del web app, volver a subir al hosting |

---

## 12. Próximos posibles trabajos

Lo que está identificado pero no implementado:

- Migración masiva de fotos viejas a formato comprimido (las nuevas ya se comprimen automáticamente)
- Audit log en hoja AuditLog para cambios sensibles (settings, approver removal, deletes)
- Rate limiting en saveDeviation para prevenir abuso
- Restringir el web app a "Anyone in domain" en lugar de "Anyone" (cambio en deploy settings)
- Notificación de deviations próximas a vencer (ya existe draftValidHours pero solo es informativo)
- Reportes (e.g. KPIs por work center, por approver, tiempo promedio de approval)
- Bulk operations (e.g. clonar deviation, exportar a Excel)
- Mobile app (actualmente el HTML es responsive pero no PWA)
- Integración con sistemas existentes del cliente (ERP, MES)
- Hosting del HTML en dominio del cliente (mwaaf.com) en lugar de distribution.fyware.com

---

## 13. Cómo retomar este proyecto en una nueva sesión

Si necesitas continuar desarrollo, comparte con la nueva instancia de Claude:
1. Este documento (HANDOVER.md)
2. Code.gs y deviation-app-sheets.html actuales
3. Una descripción clara del cambio que quieres hacer

Claude debería:
- Leer el HANDOVER primero para entender el contexto
- Revisar el archivo relevante con view antes de modificar
- Usar str_replace para cambios puntuales (no regenerar el archivo)
- Validar sintaxis después de cada cambio (Code.gs: balance de braces; HTML: extraer scripts y `node --check`)
- Entregar archivos al usuario con present_files después de cada cambio

### Convenciones del proyecto
- No usar guiones (em dashes, hyphens-as-dashes) en respuestas
- Cambios pequeños y verificables, no rewrites masivos
- Validar sintaxis siempre antes de entregar
- En español o inglés según contexto

### Estructura típica de una sesión
1. Carlos describe el cambio que necesita
2. Claude pregunta clarificaciones si hay ambigüedad (con ask_user_input_v0 cuando son opciones)
3. Claude lee las partes relevantes del código
4. Claude aplica el cambio con str_replace
5. Claude valida (braces, JS check)
6. Claude entrega los archivos con present_files
7. Claude resume qué cambió y los pasos de deploy

---

## 14. Contactos y datos relevantes

- Carlos Franco (desarrollador): carlos.franco@fyware.com
- Cliente: Gustavo Franco (Acoustafiber/MWAAF)
- IT contact del cliente: Zach
- Empresa cliente: Acoustafiber, dominio mwaaf.com
- Formulario base que reemplaza: QF158 MWAAF Deviation Authorization Form
- Cuenta dueña del sistema: mwaaf.deviations.noreply@gmail.com (password: sw3etMust@ng)
- App URL pública: https://distribution.fyware.com/MWAAF/deviationsV1
- Sheet (database): https://docs.google.com/spreadsheets/d/1mWzMNFudiiuLay3wY2-CvUP0ej4C2EKbcuYxBQrKWk8/edit
- Passwords actuales (configurables en Sheet → Config):
  - App access (al abrir el HTML): Mwaaf01
  - Settings (admin): mwaaf2024

---

*Fin del documento. Última versión incluye: re-aprobación al editar con bump de revisión, correos de approved/rejected solo a dist list, footer con quota de Gmail, headers tolerantes en catálogos, settings instantáneo, migración de owner a mwaaf.deviations.noreply, link al Sheet movido a Settings (admin only) con URL auto-inyectada, paginación + búsqueda en el catálogo de Part Numbers para soportar 1000+ items, password gate al borrar deviations desde el dashboard.*
