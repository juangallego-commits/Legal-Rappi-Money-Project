# Sincronización GitHub ⇄ Google Apps Script (clasp)

Guía operativa para mantener este repositorio sincronizado con el proyecto de
Google Apps Script usando [`clasp`](https://github.com/google/clasp) (Command
Line Apps Script Projects).

- **Script ID:** `17LQqc40ukZJDLYaAexRZF7srT6Yb7JD7Fvjp9J8Y7mYh-L9qexptWy4V`
- **Proyecto:** https://script.google.com/home/projects/17LQqc40ukZJDLYaAexRZF7srT6Yb7JD7Fvjp9J8Y7mYh-L9qexptWy4V/edit

---

## ⛔ LEE ESTO PRIMERO — el repo está incompleto

Una revisión exhaustiva del código encontró que **este repositorio es un snapshot
PARCIAL del proyecto real en Apps Script**. Faltan funciones que el frontend
invoca y que solo existen en el servidor:

| Función ausente en el repo | Para qué sirve |
|---|---|
| `processWebPayload` | **Genera el documento T&C** (la acción principal del botón "GENERAR DOCUMENTO") |
| `getCampaignTypesForUser` | Carga las tarjetas de tipo de campaña en el formulario |
| `getFieldsForUserForm` | Carga los campos dinámicos del formulario ("GOD MODE") |
| `askGemini` | Chatbot de soporte (IA) |
| `saveFeedback` | Guarda el feedback del usuario |

Además, el único `doGet` del repo (en `Codigo.gs`) sirve un archivo llamado
`Dashboard` titulado *"Legal Tracker · Rappi"* — que es **otro producto** (un
tracker de tareas legales), no el generador de T&C que está en `WebApp.Html`.

### Qué implica esto para la sincronización

> **`clasp push` hace que el servidor sea IGUAL al repo (con `--force` borra lo
> que sobre).** Si haces push de este repo incompleto, **se eliminarán del
> servidor `processWebPayload` y las demás funciones, y el generador en vivo
> dejará de funcionar.**

Por eso **el primer paso obligatorio es `clasp pull`**, NO `push`. Primero
traemos el estado real del servidor, reconciliamos con git, y solo cuando el
repo esté COMPLETO activamos el push automático.

---

## 1. Requisitos (una sola vez)

1. **Node.js 18+** y **npm** instalados.
2. **Habilitar la API de Apps Script** en tu cuenta:
   https://script.google.com/home/usersettings → *Google Apps Script API* → **ON**.
3. Tener acceso de edición al proyecto de Apps Script (con tu cuenta `@rappi.com`).

Instala las dependencias del repo (incluye clasp fijado a una versión estable):

```bash
npm install
```

> Usamos `@google/clasp@2.4.2` (fijado en `package.json`). El flujo de
> credenciales para CI documentado abajo está probado con la serie 2.x.

---

## 2. Autenticación

```bash
npx clasp login
```

Esto abre el navegador, te pide iniciar sesión con Google y guarda el token en
`~/.clasprc.json` (en tu HOME, **fuera** del repo). Ese archivo está en
`.gitignore` y **nunca** debe subirse a git.

Verifica que clasp apunta al proyecto correcto:

```bash
npx clasp status      # debe mostrar el scriptId de arriba
```

---

## 3. ⚠️ Reconciliación inicial (pull-first) — OBLIGATORIO antes de cualquier push

El objetivo es que el repo quede COMPLETO (con todo el código que hoy vive solo
en el servidor) **antes** de habilitar el push.

### Opción A — Reconciliar en una carpeta temporal y comparar (recomendado)

```bash
# 1. Clona el estado REAL del servidor en una carpeta aparte
mkdir /tmp/gas-server && cd /tmp/gas-server
echo '{"scriptId":"17LQqc40ukZJDLYaAexRZF7srT6Yb7JD7Fvjp9J8Y7mYh-L9qexptWy4V"}' > .clasp.json
npx clasp pull

# 2. Lista lo que hay en el servidor y compáralo con el repo
ls -la
#    -> ¿Aparecen archivos que NO están en el repo?
#       (p. ej. el .gs con processWebPayload/getCampaignTypesForUser/getFieldsForUserForm,
#        el HTML real del generador, el HTML 'Dashboard' del tracker, etc.)
```

Copia al repo TODO lo que falte (código del servidor) y haz commit. Cuando
`clasp pull` y el repo coincidan, el repo está completo.

### Opción B — Pull directo sobre el repo (sobrescribe archivos locales)

> Úsala solo si entiendes que **sobrescribe** los archivos locales con los del
> servidor. Como todo está en git, puedes revisar el diff antes de commitear.

```bash
cd <repo>
npx clasp pull
git status          # revisa qué cambió / qué archivos nuevos llegaron
git diff            # revisa el contenido
# reconciliar a mano lo que haga falta, luego:
git add -A && git commit -m "Reconciliar repo con el estado real de Apps Script"
```

### Sobre los nombres de archivo y el manifiesto

- `clasp pull` baja los HTML como `.html` (minúscula) y el manifiesto real como
  `appsscript.json`. **El `appsscript.json` del servidor es la fuente de verdad**
  para los scopes y la config del web app; deja que el pull sobrescriba el
  placeholder de este repo.
- Si tras el pull el generador real resulta llamarse `Dashboard` en el servidor
  (y el tracker tiene otro nombre, o viceversa), ajusta los archivos para que
  coincidan con lo que carga `doGet`/`createTemplateFromFile(...)`.

---

## 4. Flujo de trabajo diario (DESPUÉS de reconciliar)

Con el repo ya completo:

```bash
# Subir cambios de GitHub -> Apps Script
npx clasp push          # pregunta si sobrescribe el manifiesto
npx clasp push --force  # sin preguntar (lo que usa el GitHub Action)

# Traer cambios hechos en el editor web -> GitHub
npx clasp pull
git add -A && git commit -m "..." && git push
```

Atajos definidos en `package.json`:

```bash
npm run push        # clasp push
npm run push:force  # clasp push --force
npm run pull        # clasp pull
npm run status      # clasp status
npm run open        # abre el proyecto en el navegador
```

> **Regla de oro del modelo "GitHub = fuente de verdad":** el repo debe contener
> SIEMPRE el proyecto COMPLETO. Si vuelve a faltar un archivo, el siguiente push
> lo borra del servidor.

---

## 5. Push automático con GitHub Actions

El workflow `.github/workflows/clasp-push.yml` hace `clasp push --force`
automáticamente cuando se hace push a `main` sobre archivos del proyecto
(`*.gs`, `*.html`, `appsscript.json`, etc.), y también manualmente desde la
pestaña **Actions** (`workflow_dispatch`).

> 🚦 **NO actives esto hasta haber reconciliado el repo (paso 3).** Mientras el
> repo esté incompleto, cada merge a `main` borraría código del servidor.

### Configurar el secret `CLASPRC_JSON`

1. En tu equipo, tras `clasp login`, copia el contenido del archivo de
   credenciales:
   - Linux/Mac: `~/.clasprc.json`
   - Windows: `C:\Users\<usuario>\.clasprc.json`
2. En GitHub: **Settings → Secrets and variables → Actions → New repository
   secret**.
3. Nombre: `CLASPRC_JSON`. Valor: pega **todo** el contenido del archivo (es un
   JSON con el token OAuth y el refresh token).
4. Listo. El workflow escribe ese contenido en `~/.clasprc.json` dentro del
   runner y ejecuta `clasp push --force`.

> 🔐 El `CLASPRC_JSON` contiene un token que da acceso a tu Apps Script.
> Trátalo como una credencial: nunca lo pegues en el código ni en el README, y
> rótalo (`clasp login` de nuevo + actualizar el secret) si se expone.

### Cómo probar el workflow sin tocar `main`

Como el archivo del workflow existe en la rama de trabajo, puedes lanzarlo a
mano desde **Actions → "Deploy to Apps Script (clasp push)" → Run workflow** y
elegir la rama. Hazlo solo cuando el repo ya esté reconciliado.

---

## 6. Archivos de configuración en el repo

| Archivo | Función |
|---|---|
| `.clasp.json` | Apunta clasp al `scriptId` y define `rootDir` (raíz del repo). |
| `appsscript.json` | Manifiesto (timezone, runtime V8, scopes OAuth, config del web app). **Placeholder** hasta que el pull traiga el real. |
| `.claspignore` | Allowlist: solo sube `*.gs`, `*.html`/`*.Html` y `appsscript.json`. Excluye `node_modules/`, README, "Propuesta de ajuste del front", etc. |
| `.gitignore` | Excluye `node_modules/` y `**/.clasprc.json` (credenciales). |
| `package.json` | Fija `@google/clasp@2.4.2` y define los scripts `npm run ...`. |
| `.github/workflows/clasp-push.yml` | Push automático a Apps Script (ver §5). |

### Scopes declarados en `appsscript.json`

Detectados a partir del código (`SpreadsheetApp`, `DriveApp`, `DocumentApp`,
`UrlFetchApp` → Gemini, `Session.getActiveUser`):

```
spreadsheets · drive · documents · script.external_request · userinfo.email
```

`webapp.access` está en `DOMAIN` (solo cuentas `@rappi.com`) y `executeAs` en
`USER_DEPLOYING`. Esto es coherente con el uso de
`Session.getActiveUser().getEmail()` para identificar al visitante y validar
roles de admin. **Tras el pull, verifica que coincide con la config real del
despliegue.**

---

## 7. Troubleshooting

| Síntoma | Causa probable / solución |
|---|---|
| `User has not enabled the Apps Script API` | Actívala en https://script.google.com/home/usersettings |
| `Could not read API credentials` (en CI) | Falta o está mal el secret `CLASPRC_JSON`. Re-genera con `clasp login` y actualiza el secret. |
| `Push failed. Errors: ... Manifest` | El `appsscript.json` local no es válido o difiere del servidor. Haz `clasp pull` para traer el real. |
| El push pregunta y se queda colgado en CI | Usa `clasp push --force` (ya está así en el workflow). |
| Tras push, el generador dejó de funcionar | Hiciste push de un repo incompleto y se borraron funciones del servidor. Recupéralo con `clasp pull` desde una versión previa o desde el historial del editor (Archivo → Historial de versiones). **Reconcilia antes de volver a hacer push.** |
| El web app no pide login o no reconoce al usuario | Revisa `webapp.access`/`executeAs` en el despliegue y los scopes del manifiesto. |
