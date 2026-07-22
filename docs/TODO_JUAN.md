# ✅ Lo que tienes que hacer tú (Juan) — guía dummy-proof

> Actualizado 2026-07-22. Ordenado por esfuerzo. Lo de arriba toma 2 minutos y
> me desbloquea; lo de abajo es más largo pero **puedes dejarlo para cuando
> tengas paciencia** — no frena mi avance, solo hace falta para la primera
> prueba en vivo.

---

## ⚡ Tareas de 2 minutos (estas sí ayúdame a hacerlas ya)

### 1. Mensaje a Valentina (equipo CRM)
Copia y pega esto (el tema del manifiesto ya quedó resuelto, solo faltan 2 cosas):

> *Vale! Dos cositas para cerrar la integración con RappiMind:*
> *1. **Seguridad:** en el `Code.gs` que compartieron viene pegado el token del bot de Slack (línea 7, empieza por `xoxb-`). Como ese archivo ya circuló, hay que rotarlo: api.slack.com/apps → la app del bot → OAuth & Permissions → regenerar token. El nuevo, guárdenlo en Propiedades del script (no pegado en el código). ¿Me lo compartes por un canal seguro cuando lo tengas?*
> *2. **Cupones:** el formulario no tiene tipo "cupón", pero los países piden T&Cs de cupones seguido. ¿Los cupones se montan por otra herramienta/flujo? ¿Cuál?*

### 2. Pregunta a Anna (lo de la carpeta ALL)
> *Annie, sobre lo que viste de que tu prueba se guardó en la carpeta ALL y no en la de CO: ¿te acuerdas si creaste esa plantilla con el modal "Nueva Plantilla" (donde el país sale como "Global") o con el wizard eligiendo Colombia? Es para saber si fue lo esperado o un bug, y reproducirlo.*

*(Con su respuesta reproduzco y arreglo si es bug — ya tengo identificado el punto exacto del código.)*

---

## 🛠️ Montar el /dev y encender todo (≈25 min, deferible)

Solo hace falta para **correr la primera prueba en vivo**. Mientras, yo sigo
construyendo. La guía detallada está en **`docs/RUNBOOK_P0_PASO_A_PASO.md`**;
aquí el resumen dummy.

### Paso A — Credenciales de despliegue (una sola vez)
> 💡 Si no tienes Node instalado, lo más fácil es hacer esto en **Google Cloud
> Shell** (shell.cloud.google.com — es un terminal en el navegador, no instalas
> nada, entras con tu cuenta @rappi.com).

1. Abre un terminal (o Cloud Shell) y corre:
   ```
   npm install -g @google/clasp
   clasp login
   ```
   (Se abre el navegador → inicia sesión con @rappi.com → "Success".)
2. Muestra el archivo de credenciales y **copia TODO** lo que imprime:
   ```
   cat ~/.clasprc.json
   ```
3. En GitHub: repo → **Settings → Secrets and variables → Actions → New
   repository secret**. Nombre: `CLASPRC_JSON`. Valor: pega lo que copiaste. Save.

*(Esto es lo único "técnico" y es de una sola vez. Después, desplegar es un botón.)*

### Paso B — Crear el proyecto /dev y desplegar (botón, sin terminal)
1. Abre **script.new** → nómbralo "RappiMind DEV" → ⚙️ Configuración del
   proyecto → copia el **"ID de secuencia de comandos"** (Script ID).
2. GitHub → pestaña **Actions** → workflow **"Deploy to Apps Script DEV"** →
   **Run workflow** → pega el Script ID → **Run**. Espera el ✓ verde.

### Paso C — Copiar la base de datos y configurar (5 min de clics)
1. Abre el spreadsheet de RappiMind (la DB) → **Archivo → Hacer una copia** →
   nómbrala "RappiMind DB — DEV" → copia su ID (lo de la URL entre `/d/` y `/edit`).
2. En el editor de **RappiMind DEV** → ⚙️ → **Propiedades del script** → agrega:
   - `RAPPIMIND_DB_ID` = (ID de la copia que acabas de hacer)
   - `CRM_SPREADSHEET_ID` = `1XR5zWNj5cd3vDLCEeaOX6mqjPliknJ9gHhO6YVzrzuE`
   - `CRM_CSV_FOLDER_ID` = `1T39sOtnG1rd4UCeJZtSs6VDoMCv3Cb5D`

### Paso D — Encender Organizador + A2 (funciones de un clic)
En el editor de DEV, abre el archivo **`P0.gs`**, y en el selector de funciones
de arriba elige y dale **Ejecutar** en orden (la 1ª vez pide autorizar: acepta):
1. `P0_0_estado` → debe decir `DB … → COPIA (/dev) ✓`
2. `P0_1_sembrarTcGenerales` → `added: 13`
3. `P0_2_previewOrganizador` → `adds: 2` (si no, para y me avisas)
4. `P0_3_aplicarOrganizador` → `added: 2`; **córrela otra vez** → `added: 0`
5. `P0_4_encenderA2` → `A2 = on`

### Paso E — Probar
Implementar → **Nueva implementación** → App web → abrir la URL:
- Sin `?page`: el generador. Colombia → Cashback → debe pedir **Organizador** →
  Generar → el Doc sale sin `{{}}` ni `[ ]`.
- `…?page=tc-generales`: catálogo de T&C generales.
- `…?page=crm`: formulario CRM (crea ticket en tu copia del Sheet).

**Cuando llegues aquí, avísame y hacemos la primera prueba end-to-end juntos.**

---

## 📸 Para después (cuando lleguemos a las guías, F2c)
Pantallazos de Squarespace en buena calidad (los 4 pasos que ya tienes) para
armar la guía dummy-proof con fotos.

---

### Resumen de una línea
**Hoy:** manda los 2 mensajes (Valentina + Anna). **Cuando tengas 25 min:** el
montaje /dev (Pasos A–E). Todo lo demás lo sigo yo.
