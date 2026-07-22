# Runbook — F1: módulo CRM en RappiMind (/dev)

Cómo desplegar y probar la fusión del Campaign Manager dentro de RappiMind
**sin tocar la producción del equipo CRM**. Prerrequisito de contexto:
`docs/INTEGRACION_CRM_CAMPAIGN_MANAGER.md`.

## Qué añadió F1 al proyecto (4 archivos + router)

| Archivo | Qué es |
|---|---|
| `Crm.gs.js` | Backend del Campaign Manager adaptado: config por Script Properties, sin `doGet` propio, `LockService` al generar `CAM-####`, columna nueva `TC_DOC_URL`, guardas si falta configuración. |
| `CrmForm.html` / `CrmStatus.html` / `CrmAdmin.html` | Las 3 páginas del CRM con slugs nuevos (`?page=crm`, `crm-status`, `crm-admin`). |
| `Código.js#doGet` | Router: sin `?page` sirve el generador de siempre; `?page=crm*` sirve el módulo CRM. **Cero impacto** en el flujo actual de RappiMind. |

## Paso 0 — Insumos (una vez)

1. **Token de Slack:** pedir al equipo CRM que **rote** el token expuesto y
   genere uno para pruebas (o usar un bot propio en un canal de pruebas).
2. **Copia del Sheet del CRM:** abrir su spreadsheet (`1yzcRT…`) → Archivo →
   Hacer una copia (incluye hojas `Base` y `RappiCreditos`). Anotar el ID.
3. **Carpeta Drive de pruebas** para los CSV por ticket. Anotar el ID.
4. Canal de Slack de pruebas (p. ej. `#rappimind-dev`) con el bot invitado.
   Anotar el channel ID (empieza por `C`).

## Paso 1 — Desplegar el código a /dev

En el proyecto Apps Script de **pruebas** de RappiMind (no producción):

```bash
# .clasp.json apuntando al scriptId de /dev
npx clasp push --force
```

Deben quedar 11 archivos en el editor: los 7 de siempre + `Crm.gs`,
`CrmForm`, `CrmStatus`, `CrmAdmin`.

## Paso 2 — Script Properties (Configuración del proyecto)

| Property | Valor |
|---|---|
| `CRM_SPREADSHEET_ID` | ID de la **copia** del paso 0.2 |
| `CRM_SLACK_BOT_TOKEN` | token de pruebas (paso 0.1) |
| `CRM_SLACK_CHANNEL` | channel ID de pruebas (paso 0.4) |
| `CRM_CSV_FOLDER_ID` | carpeta del paso 0.3 |
| `CRM_TIMEZONE` | (opcional) default `America/Bogota` |

> Sin `CRM_SPREADSHEET_ID` las páginas cargan pero guardar/consultar falla con
> mensaje claro ("CRM no configurado"). Sin token/canal, simplemente no se
> notifica a Slack (el ticket sí se crea). Sin carpeta, se omite el CSV.

## Paso 3 — Probar

1. Desplegar la web app de /dev (o usar "Probar implementación").
2. **Regresión RappiMind:** abrir la URL **sin** `?page` → debe cargar el
   generador de T&C como siempre.
3. `?page=crm` → formulario Campaign Manager. Crear una solicitud tipo
   `value` (dummy) → verificar: fila en la copia del Sheet (con la columna
   nueva `TC_DOC_URL` al final), CSV en la carpeta, mensaje en el canal de
   pruebas con hilo de detalle.
4. `?page=crm-status` → buscar el ticket por ID y por email.
5. `?page=crm-admin` → cambiar estado; verificar que exige Global Offer ID
   para "Aprobado" y que responde en el hilo de Slack.
6. Enviar 2 solicitudes casi simultáneas (2 pestañas) → IDs consecutivos sin
   duplicar (lock).

## Rollback

Los archivos CRM son aditivos. Para desactivar el módulo: quitar las 3 líneas
`crm*` del router en `Código.js#doGet` (o borrar las Script Properties
`CRM_*`, que deja el módulo inerte). El generador de T&C no depende de nada
del módulo CRM.

## Qué NO hace F1 (viene en F2)

- La sección legal en el form (Organizador + confirmaciones) y el
  encadenamiento `submitForm → processWebPayload → saveRequest` con
  `tcDocUrl` (Cashback CO). El backend ya acepta `data.tcDocUrl`.
- Autollenar `linkTyC` con los **T&C generales** por país/tipo (falta el
  catálogo de links — pedirlos a Legal por país).
- Rediseño user-friendly del flujo (feedback de Anna) y gates de rol en
  `updateStatus`/`updateGlobalOffer`.
