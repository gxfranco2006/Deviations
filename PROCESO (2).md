# Análisis del Proceso de Desarrollo

## MWAAF Deviation Authorization System

**Autor**: Carlos Franco
**Periodo**: abril a mayo 2026

---

## Contexto

El proyecto comenzó como un sistema heredado en estado parcial: una app de Google Apps Script + HTML standalone que funcionaba en lo básico pero tenía deuda técnica acumulada, bugs específicos (catálogos vacíos en Settings, correos sin links válidos), y faltaban funcionalidades clave (re-aprobación al editar, control de cuota visible, notificaciones específicas a stakeholders, etc.).

El reto no era reescribir el sistema, sino **iterar sobre código existente sin romperlo**, conservando convenciones, estructura de datos, y compatibilidad hacia atrás. Para acelerar el trabajo se usó asistencia de IA como copiloto en las modificaciones, pero la dirección técnica, las decisiones de diseño, y la validación final dependieron del desarrollador humano.

---

## Lógica de trabajo aplicada

### 1. Documentación del estado inicial antes de modificar nada

Antes de pedir cualquier cambio se elaboró un documento de handover detallado (estructura del Sheet, endpoints, flujos, decisiones técnicas pasadas, y bugs conocidos). Este documento sirvió como contexto inicial obligatorio en cada sesión: cualquier asistente, humano o IA, que tomara el proyecto debía leerlo primero.

**Por qué importa**: sin este documento, la IA no tiene forma de respetar decisiones previas (por ejemplo, por qué las fotos están en base64 en el Sheet en lugar de Drive, o por qué los approvers usan tokens en lugar de login). Se evitan rewrites innecesarios y propuestas que rompen compatibilidad.

### 2. Solicitudes específicas, una a la vez, con criterios de aceptación claros

En cada interacción se planteaban requerimientos concretos y atómicos. No "mejora el sistema de aprobaciones" sino:

- "Que el correo de Approved se envíe solo a la Distribution List 'On Approval', no a todos los approvers"
- "Que también se envíe correo cuando una deviation es Rejected, a esos mismos destinatarios"
- "Que en el footer del index aparezca una barra que indique la cuota de correos usados"
- "Verificar que al editar una deviation vuelva a correr el flujo de aprobaciones"
- "Al dar click en Settings tarda en abrir el popup; reducir esa latencia"

Cada bullet tiene un **comportamiento esperado verificable**. Esto permite testear cada cambio aislado y no acumular regresiones.

### 3. Preguntas de clarificación antes de implementar

Cuando un requerimiento tenía más de una interpretación razonable, se forzaba una decisión explícita antes de tocar código. Ejemplo del cambio de re-aprobación al editar:

> ¿Cuando se edite una deviation que ya tenía firmas, qué quieres que pase con las firmas previas?
> a) Borrar todas y volver a empezar (todos firman de nuevo)
> b) Mantener firmas previas, solo notificar a los que falten

Esta pregunta es trivial de responder pero crítica para la implementación. Sin ella, la IA hubiera asumido (a) o (b) sin pedir permiso, y posiblemente la decisión hubiera sido la incorrecta.

**Resultado**: las decisiones de producto las toma el dueño del proyecto, no el modelo. La IA solo ejecuta sobre la decisión tomada.

### 4. Verificación de sintaxis y lógica antes de cada entrega

Cada cambio se validaba con `node --check` para Code.gs y extracción de scripts del HTML para verificar JS embebido. Esto se hacía automáticamente como parte del flujo de cambios, no como paso opcional.

Adicionalmente se balanceaban delimitadores (braces, paréntesis, brackets) y se verificaba que el archivo modificado no perdiera funciones existentes (un riesgo común al hacer reemplazos grandes con regex).

**Por qué importa**: Apps Script no tiene "build step" que detecte errores antes de deploy. Si pegas un archivo con un error de sintaxis y haces deploy, todo el endpoint deja de responder y los usuarios ven errores genéricos. Validar localmente antes de subir evita esos accidentes.

### 5. Diagnóstico estructurado cuando algo fallaba

Cuando un cambio no funcionaba en producción, no se asumía que el código estaba mal. Se seguían pasos de descarte:

1. ¿El archivo en el editor de Apps Script tiene los cambios? (verificar visualmente)
2. ¿Se hizo new version del deployment? (no solo Save, no solo cambiar nombre)
3. ¿El localStorage del navegador tiene cache vieja? (Clear Cache button en debug panel)
4. ¿La URL del web app que usa el HTML coincide con el deployment correcto?
5. ¿Hay algún log en Apps Script Executions que dé pistas?

Caso ejemplo: cuando los catálogos de Part Numbers y Work Centers no aparecían en Settings, el primer instinto pudo haber sido reescribir las funciones de render. En cambio, se diagnosticó:

- Frontend: ¿el render se ejecuta? Sí
- Backend: ¿la función getPartNumbers retorna datos? No, retorna array vacío
- Sheet: ¿hay datos en la hoja? Sí
- Función: ¿qué hace exactamente? Busca header exacto "partNumber" con `indexOf`

La causa raíz era algo distinta a lo esperado y solo se descubrió al pedir directamente al backend qué estaba retornando, no asumir desde el código fuente. Se hicieron pruebas con `fetch` directo al endpoint desde la consola del navegador para aislar si era frontend o backend.

### 6. Cambios reversibles y mínimos

Cada cambio se hacía con `str_replace` (reemplazo de string específico), no con regeneración completa del archivo. Esto da:

- Diffs claros: se ve exactamente qué cambió
- Bajo riesgo de perder código no relacionado
- Fácil reversión si algo falla
- Permite revisar cambios uno a uno antes de probar

Cuando se necesitaba agregar una función nueva (por ejemplo, `buildRejectedEmailHtml`), se agregaba al lado de la función análoga existente (`buildApprovalEmailHtml`) reusando estilo, naming conventions, y parámetros similares. La idea es que el código nuevo se vea como si lo hubiera escrito quien escribió el código viejo.

### 7. Convenciones explícitas y persistentes

Se establecieron desde el inicio reglas estilísticas y de proceso, y se mantuvieron en cada sesión:

- No usar guiones largos (em dashes) en respuestas; preferir paréntesis o puntos
- En español o inglés según contexto del mensaje
- No regenerar archivos masivos; cambios puntuales
- Validar sintaxis siempre antes de entregar
- Resumir qué cambió y los pasos de deploy al final de cada cambio

Estas convenciones quedaron documentadas en el handover, así cualquier sesión nueva las respeta sin tener que repetirlas.

### 8. Separación clara de responsabilidades

- **Decisiones de producto**: humano (qué debe hacer el sistema)
- **Decisiones de arquitectura**: humano con input del modelo (cómo lograrlo)
- **Implementación de código**: modelo bajo dirección humana
- **Validación final y deploy**: humano (probar en ambiente real, decidir cuándo publicar)
- **Gestión de credenciales y permisos**: solo humano (transferencia de ownership, accesos a cuentas)

El modelo nunca tomó decisiones unilaterales sobre seguridad, datos, o accesos. Cuando hubo dudas (por ejemplo, al transferir ownership entre dominios distintos de Google), se pidieron alternativas y la decisión final la tomó el desarrollador.

### 9. Aprovechamiento del contexto compartido

Cada sesión podía continuar de donde quedó la anterior gracias al handover + archivos actuales. No se perdía progreso ni se repetían explicaciones. Esto también ayudó a que las decisiones tomadas en sesiones pasadas (por ejemplo, "los emails de approved van solo a la dist list") se mantuvieran consistentes en sesiones nuevas.

### 10. Pruebas en ambiente real, no solo validación local

Después de cada cambio se hacían pruebas concretas en el sistema desplegado:

- Crear una deviation de prueba
- Recibir el correo y verificar que el link funciona
- Verificar que el dashboard refleja el cambio esperado
- Probar el flujo completo (crear → aprobar → rechazar → editar)

La validación local de sintaxis solo garantiza que el código corre. La validación end-to-end garantiza que el comportamiento es el correcto.

---

## Errores comunes evitados

### Pegar y desplegar sin verificar

En Apps Script es muy fácil hacer Save y olvidar el step de Manage deployments → New version. Sin esa segunda parte, los cambios viven solo en el editor pero la URL pública sigue corriendo el código viejo. Se documentó esta trampa en el handover y se verificó explícitamente cuando aparecieron errores tipo "Unknown action: X".

### Rewrites masivos cuando solo hace falta un cambio puntual

Es tentador pedirle a la IA "regenera el archivo completo con esta lógica nueva". El problema: el archivo regenerado puede perder optimizaciones específicas, comentarios útiles, o ramas de código que se agregaron en sesiones anteriores y que la IA no tiene presentes. Se prefirieron siempre cambios pequeños sobre el archivo existente.

### Asumir que el cache no es el problema

En desarrollo web, el cache es un punto frecuente de confusión. El frontend tiene cache de localStorage, el navegador tiene cache de archivos estáticos, el Apps Script tiene cache de versiones desplegadas. Se incluyó un botón "Clear Cache" en el debug panel del HTML para hacer este reset trivial, y siempre que un cambio "no se ve", se incluyó como primer paso de diagnóstico.

### Confiar en el frontend para decisiones de seguridad

Hubo casos donde el camino fácil hubiera sido pedirle al frontend que mande un payload con cierta información (por ejemplo, las approvals limpias después de un edit). El camino correcto fue que el backend re-leyera del Sheet, sin confiar en el frontend. Esto evita bugs sutiles cuando el frontend está fuera de sync, y también previene que un usuario malicioso manipule el payload.

---

## Resultado

El sistema final entregado tiene:

- 12+ funcionalidades nuevas o refinadas a lo largo de varias iteraciones
- Cero regresiones reportadas en funcionalidad previa
- Documentación que permite a un humano nuevo (o a una nueva sesión de IA) retomar el proyecto sin contexto previo
- Convenciones consistentes de código, naming y comportamiento

El factor común en todas las sesiones productivas fue la combinación de: requerimientos específicos + preguntas de clarificación cuando aplican + validación rigurosa + pruebas reales + documentación al final.

La IA en este proyecto funcionó bien porque se le dio contexto suficiente, instrucciones claras, y se validó su salida. Cuando alguno de esos tres elementos falta, el resultado es código que parece correcto pero rompe algo en producción.

---

*Documento elaborado como retrospectiva del proceso de desarrollo del MWAAF Deviation Authorization System.*
