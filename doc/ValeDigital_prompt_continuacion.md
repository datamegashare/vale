# Vale Digital — Prompt de continuación

## Contexto del proyecto
PWA para gestión de vales de pedido a almacén y pañol en obras electromecánicas.

**Stack:**
- GitHub Pages: https://datamegashare.github.io/vale/
- Google Apps Script (GAS) como backend REST — todo por GET (solución al redirect 302 de GAS)
- Google Sheets como base de datos
- Google OAuth 2.0 (GIS) — mismo GCP project que "Mesa de Trabajo"
- Service Worker para modo offline

**Credenciales:**
- OAuth Client ID: `174536806740-aus6npjphhdcn5nkgtt1ejl0tcj16ou7.apps.googleusercontent.com`
- GAS URL: `https://script.google.com/macros/s/AKfycbyMeo4ebB9qlt6rJVM_2N0ENvqsmnKPbb6CsGo9Fc-49MLv8ei0IcCR_NtvhQgmLRo9uw/exec`
- Google Sheet ID: `10XkTAarQdgucz8WIwNh5FhV2vyW9qdoRhRCKnQ7wb4k`

---

## Decisiones técnicas clave

- **Todo por GET** — GAS hace redirect 302 y pierde el body en POST. Tanto `gasGet` como `gasPost` usan GET con parámetros en URL.
- **SW no intercepta GAS** — `return` sin `event.respondWith` para `script.google.com`.
- **Archivos siempre descargables** — nunca código inline en el chat (.gs, .html, .js).
- **Nombre con versión** — `index_v10.html`, `vale_gas_v5.gs`, `sw_v5.js`, etc.
- **Versión visible** en login y en menú ⋮ mobile (`CONFIG.VERSION_UI`).
- **Nueva implementación GAS** — siempre editar implementación existente → Nueva versión (no crear deployment nuevo, para mantener la URL).
- **SW** — cada nueva versión de index.html requiere actualizar `CACHE_NAME` en sw.js.

---

## Estado actual — Etapas completadas

### Etapa 1 ✅ — Solicitud de vales
- Login Google (GIS), validación contra hoja USUARIOS
- Sesión persiste (perfil en localStorage)
- Dashboard 4 contadores (Borrador, Enviado, Aprobado, Entregado)
- Lista vales últimos 30 días con badge estado y sync
- FAB "+" → modal crear/editar vale con Quill.js
- Destino toggle (🏭 Almacén / 🔧 Pañol)
- Offline: localStorage, sincroniza al reconectar
- SW cache-first, PWA instalable

### Etapa 2 ✅ — Panel de Aprobación
- Tabs: "Mis Vales" (todos) / "Por Aprobar" (SUPERVISOR y JEFE_OBRA)
- SUPERVISOR aprueba/rechaza vales de CAPATAZ (misma obra)
- JEFE_OBRA aprueba/rechaza vales de SUPERVISOR (misma obra)
- JEFE_OBRA: aprobación automática para sus propios vales al enviar
- Modal rechazo con nota obligatoria (≥5 chars)
- Badge en tab con cantidad de pendientes (precargado al login)
- Desktop: layout dos columnas lista + detalle

### Etapa 3 ✅ — Panel Almacén/Pañol (EN PRUEBAS)
- Tab "Por Gestionar" (ALMACENERO / PAÑOLERO)
  - Sub-dashboard: Por iniciar / Preparando / Entregados hoy
  - Cola APROBADO ordenada por antigüedad
  - Lista PENDIENTE ordenada por solicitante (búsqueda presencial)
  - Flujo: APROBADO → PENDIENTE → ENTREGADO / ENTREGA_PARCIAL / CANCELADO
  - Modal nota de cierre (obligatorio u opcional según acción)
  - Desktop: dos columnas
- Tab "Todos los Vales" (JEFE_OBRA)
  - Buscador por solicitante (parcial, case-insensitive)
  - Lista readonly con detalle en modal

---

## Archivos deployados actualmente

| Archivo GitHub | Versión | Estado |
|---|---|---|
| `index.html` | v9.0 | EN PRUEBAS |
| `sw.js` | v4.0 (`vale-digital-v4.0`) | EN PRUEBAS |
| GAS (mismo archivo interno) | v4.0 | EN PRUEBAS |

---

## Estructura Google Sheet (VALES)

| Col | Campo |
|---|---|
| 1 | id_vale |
| 2 | fecha_hora |
| 3 | usuario_email |
| 4 | usuario_nombre |
| 5 | obra_codigo |
| 6 | destino |
| 7 | titulo |
| 8 | contenido_html |
| 9 | estado |
| 10 | aprobado_por |
| 11 | fecha_aprobacion |
| 12 | gestionado_por |
| 13 | fecha_cierre |
| 14 | nota_cierre |
| 15 | eliminado |

**Estados completos:** BORRADOR → ENVIADO → APROBADO → PENDIENTE → ENTREGADO / ENTREGA_PARCIAL / CANCELADO / RECHAZADO / ELIMINADO

---

## Roles

| Rol | Puede crear | Aprueba | Gestiona | Ve todos |
|---|---|---|---|---|
| CAPATAZ | ✅ | ❌ | ❌ | ❌ |
| SUPERVISOR | ✅ | Vales de CAPATAZ | ❌ | ❌ |
| JEFE_OBRA | ✅ (auto-aprobado) | Vales de SUPERVISOR | ❌ | ✅ su obra |
| ALMACENERO | ❌ | ❌ | Destino ALMACEN | ❌ |
| PAÑOLERO | ❌ | ❌ | Destino PAÑOL | ❌ |
| ADMIN | — | — | — | — (Etapa 4) |

---

## Usuarios de prueba en Sheet (OB-001)

| email | rol |
|---|---|
| alejandro.perin@gmail.com | (variable según prueba) |
| datamegashare@gmail.com | ADMIN |
| capataz@tuempresa.com | CAPATAZ |
| supervisor@tuempresa.com | SUPERVISOR |
| jefe@tuempresa.com | JEFE_OBRA |
| goecdg@gmail.com | ALMACENERO |

---

## Pendientes / bugs conocidos

- **Flujo logout** — al cerrar sesión aparece el selector de cuenta Google (comportamiento correcto pero mejorable — pospuesto a Etapa 4 o posterior)
- **Etapa 3 en pruebas** — puede haber bugs a corregir en `index_v10.html`

---

## Próximas etapas planificadas

- **Etapa 4:** Admin CRUD usuarios y obras
- **Etapa 5:** Notificaciones email desde GAS
- **Etapa 6:** Mejoras UX — drawer lateral, mejora flujo logout

---

## Reglas de trabajo (SIEMPRE aplicar)

1. Archivos siempre descargables — nunca código inline en el chat
2. Nombre con versión — `index_v10.html`, `vale_gas_v5.gs`, `sw_v5.js`
3. Versión visible en login y en menú ⋮ mobile — `CONFIG.VERSION_UI`
4. Todo por GET — nunca POST
5. SW no intercepta GAS
6. Nueva versión GAS = editar implementación existente → Nueva versión (misma URL)
7. Siempre generar `sw_vX.js` cuando cambia `index.html`
8. Paso a paso — un paso a la vez, esperar confirmación antes de continuar
9. Spinners en todas las listas mientras cargan
10. Datos de prueba hardcodeados en archivos — el usuario los edita ahí
