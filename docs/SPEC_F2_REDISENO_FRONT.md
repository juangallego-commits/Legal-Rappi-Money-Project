# SPEC F2 — Rediseño integral del front (RappiMind + CRM unificados)

> Especificación de diseño para el "súper rediseño" pedido por Juan
> (2026-07-22), incorporando **todo el feedback de Anna Habermann** (hilo del
> 2026-06-28 en el grupo del proyecto) y las decisiones de integración con el
> Campaign Manager. RappiMind no está en producción real → hay libertad para
> rediseñar `WebApp.Html` sin migración.

## 0. Principios

1. **Dummy-proof**: cada paso dice qué hacer, qué NO hacer y qué sigue. Cero
   jerga legal sin explicación. Reducir el campo de error > reducir clics.
2. **Cero texto libre en lo legal**: todo por opciones pre-aprobadas
   (dropdowns/chips/checkboxes). El texto libre solo va a campos operativos
   del ticket (descripción CRM).
3. **Trazabilidad end-to-end visible**: `Ticket CAM-#### ↔ T&C (Doc) ↔ link
   publicado (Squarespace) ↔ Global Offer ID`. El usuario siempre ve en qué
   eslabón está y qué falta.
4. **Separación clarísima de responsabilidades** (feedback Anna): generar el
   T&C (Legal/tú) ≠ publicarlo en Squarespace (el solicitante) ≠ montar la
   Global Offer (CRM). El UI nombra al responsable de cada paso.
5. **Un solo sistema**: el generador T&C y la solicitud CRM comparten flujo,
   datos y estética (la actual de RappiMind: Tailwind, Nunito, naranja Rappi).

## 1. Mapa de pantallas

```
HOME (hub)
├── 🚀 Crear campaña con T&C específico  → WIZARD (cashback hoy; concurso después)
│     País → Tipo → Datos comerciales → Restricciones → LEGAL → Revisión → Resultado
├── 🔎 T&C Generales (para KAMs)         → CATÁLOGO buscable por país/tipo, copiar link
├── 🎫 Solicitud CRM directa             → CrmForm (tipos sin T&C específico; linkTyC autollenado)
├── 📍 Consultar ticket / mis campañas   → CrmStatus + historial local
├── 📚 Guías                             → Squarespace paso a paso CON FOTOS, FAQ, videos
└── ⚙️ Admin                             → tabs actuales + tickets CRM + catálogo TC_Generales
```

## 2. Feedback de Anna → decisión por ítem

| # | Feedback | Decisión | Estado |
|---|---|---|---|
| A1 | Paso 3 de la guía: enfatizar que ELLOS publican en Squarespace; renombrar a "Publica los T&Cs" con su copy | Aplicado tal cual (copy literal de Anna) | ✅ hecho (2026-07-22) |
| A2 | "Consumo en el local" confunde con la redención de créditos | Renombrar secciones: **"¿Dónde se redimen los créditos?"** (redención) vs **"Condiciones especiales de la compra"** (requisitos para GANAR el beneficio). Hint bajo cada una explicando la diferencia con ejemplo | F2 wizard |
| A3 | Condiciones especiales **no son excluyentes** → checkboxes | Sí: multi-select con checkboxes. En el T&C salen como **viñetas** en la cláusula de condiciones (redacción de cada combinación pre-aprobada por Juan — ver §5) | F2 wizard |
| A4 | Botón "Generar" siempre naranja → clicks prematuros | Botón **deshabilitado (gris) + contador "faltan N campos"**; se habilita al completar obligatorios; aparece al final del flujo (wizard por pasos lo resuelve de raíz: el botón vive en el último paso) | F2 wizard |
| A5 | FAQ concursos: 💡 "si dudas de azar → legal local" + ejemplos de NO válidos | Aplicado (💡 + ejemplos válidos/no válidos) | ✅ hecho |
| A6 | Quitar a Dave; poner a Paula | Aplicado (juan.gallego + paula.barahona) | ✅ hecho |
| A7 | Canal Slack global en vez de #legalops-marketing-co-cr | **Decisión pendiente Juan/Anna**: crear #legalops-marketing-latam o mantener por región. El contacto se vuelve dato configurable (hoja `Country_Settings` o constante única) para no tocar código al cambiar | F2 + decisión |
| A8 | Guía Squarespace con fotos, dummy proof | Guía paso a paso con **screenshots reales** (Anna ya mandó 5 de referencia), numerada, con "errores comunes". Pantalla propia en "Guías" + link desde el modal de éxito | F2 guías (necesito los pantallazos en buena calidad) |
| A9 | Prueba se guardó en carpeta ALL y no CO | Diagnóstico: `_moveTemplateToFolder` usa `country_code` de la plantilla; si se creó con país **"Global"** → carpeta `ALL` **a propósito**. Si Anna eligió CO y cayó en ALL → bug real. **Pedirle a Anna el flujo exacto que usó** (modal Nueva Plantilla vs wizard) y reproducir en /dev | 🔍 verificar |
| A10 | Campo al final para pegar el **link publicado** (trazabilidad) | Sí: en el modal de éxito + en la consulta del ticket ("Pegar link publicado"). Se guarda en `terminos_y_condiciones` del ticket y en `Respuestas_Audit_V2`. El admin CRM muestra ⚠️ si falta al aprobar | F2 wizard/status |

## 3. El wizard (pantalla principal, para dummies)

Pasos: **1 País** → **2 Tipo de campaña** → **3 Datos comerciales** → **4
Restricciones** → **5 Datos legales** → **6 Revisión** → **7 Resultado**.

- Un paso por pantalla, barra de progreso con nombres, "Anterior/Siguiente"
  bloqueado hasta validar el paso (A4). Autosave (ya existe) + "cargar
  campaña similar" (heredado del CRM form).
- **Paso 2** separa: "💰 Cashback (T&C específico)" / "🏆 Concurso (T&C
  específico, no pasa por CRM)" / "🏷️ Otras ofertas (usan T&C general —
  descuentos, envío gratis, tarifa de servicio…)". Cada tarjeta dice qué se
  genera y quién interviene (Legal / CRM / tú).
- **Paso 5 (LEGAL)**: Organizador (razón social + NIT del aliado — dropdown
  de aliados frecuentes + "otro" con validación de formato), confirmaciones
  de responsabilidad, condiciones especiales (checkboxes, A3).
- **Paso 6 (Revisión)**: resumen completo tipo "¿Está todo correcto?" con
  edición por sección (patrón del CRM form, que ya lo hace bien).
- **Paso 7 (Resultado)**: checklist de 3 ítems con estado vivo:
  1. ✅ T&C generado → link al Doc
  2. ⬜ **Publícalo en Squarespace** → botón "Ver guía con fotos" + campo
     "pega aquí el link publicado" (A10)
  3. ⬜ Solicitud CRM `CAM-####` creada → estado del ticket (si el tipo pasa
     por CRM); si no aplica, no se muestra.
- **Nombres** (pedido de Juan): el wizard **genera el nombre automático**
  para evitar inconsistencias:
  - Interno: `CO · Cashback · {Marca} · {AAAA-MM} · CAM-####`
  - Título público del Doc/Squarespace: `Términos y Condiciones — {Nombre
    comercial de la campaña} — {Marca} ({Mes AAAA})`
  - El slug de Squarespace sugerido se muestra copiable.
  (Convención final a validar por Juan/Anna — propuesta editable.)

## 4. Página "T&C Generales" (idea de Juan)

Buscador para KAMs sin generar nada: filtro país + búsqueda por texto,
tarjetas con nombre, para qué sirve (descripción del catálogo), vigencia y
botón **"Copiar link"** (+ aviso "si tu dinámica no encaja aquí, necesita T&C
específico → créalo en el wizard"). Fuente: `getGeneralTcCatalog` (ya en
producción de código). Incluye cupones y membresías aunque el CRM no los
tenga como tipo.

## 5. Decisiones que necesito de Juan antes de construir F2

1. **Redacción legal de condiciones especiales combinadas** (A3): apruébame
   el patrón "La promoción está sujeta a las siguientes condiciones: •X •Y"
   o dame el texto por combinación.
2. **Convención de nombres** (§3): ¿ok la propuesta?
3. **Canal Slack** (A7): ¿global o por región? Nombre.
4. **Aliados frecuentes**: ¿mantenemos una hoja `Aliados` (razón social +
   NIT pre-aprobados) para el dropdown del Organizador? Reduce error de
   digitación en el dato más sensible del T&C.
5. **A9**: pregunta a Anna qué flujo usó para reproducir lo de la carpeta.

### 5.1 Estado de estas decisiones (2026-07-22) — Juan: "sí a todo"

1. ✅ **Condiciones combinadas → viñetas**: implementado el catálogo +
   combinador (`SPECIAL_CONDITIONS_CATALOG` + `_buildSpecialConditionsText`),
   con texto legal **en BORRADOR** (pendiente que Juan valide el wording final)
   y validación de excluyentes (local vs domicilio). ⚠️ Requiere agregar el
   placeholder `{{TEXTO_CONDICIONES}}` al template de Cashback (con bloque FASE
   B para que, si no hay ninguna, no quede rastro) — tarea de Juan al editar el
   Doc.
2. ✅ **Convención de nombres**: implementada (`_buildCampaignNames` +
   `previewCampaignNames`), testeada.
3. ⏳ **Canal Slack** (A7): global vs regional — sigue pendiente Juan/Anna;
   el contacto se hará configurable para no tocar código.
4. ✅ **Aliados con AUTO-APRENDIZAJE** (F2b): al generar un T&C se guarda solo
   el organizador (`_learnAliadoFromPayload`); `getAliadosCatalog` alimentará el
   autocompletar del wizard. (En vez de una hoja pre-cargada a mano.)
5. ⏳ **A9 carpeta**: esperando el flujo que usó Anna para reproducir.

## 6. Plan de entrega (3 PRs)

| PR | Contenido | Depende de |
|---|---|---|
| F2a | Wizard por pasos completo (A2, A3, A4, A10) + sección legal + encadenamiento T&C→ticket (cashback CO) + nombres automáticos | Decisiones §5.1–5.2; /dev montado (P0) |
| F2b | Página T&C Generales + autollenado linkTyC en solicitudes CRM + hub Home | nada (backend listo) |
| F2c | Guías con fotos (Squarespace), FAQ ampliado, admin: tickets CRM + catálogo + gates de rol (`_requireRole` en las 5 funciones + updateStatus CRM) | Pantallazos A8; decisión A7 |

Orden sugerido: **F2b → F2a → F2c** (F2b es la victoria rápida sin
dependencias; F2a es el corazón; F2c pule y cierra seguridad).

### 6.1 Avance real (2026-07-22)

- ✅ **F2b entregado**: página T&C Generales, autollenado de `linkTyC`, aliados
  con auto-aprendizaje, hub provisional.
- ✅ **Helpers de F2a listos y testeados** (backend puro, sin depender de /dev):
  nombres automáticos, catálogo+combinador de condiciones especiales, APIs
  `previewCampaignNames` / `getSpecialConditionsCatalog`.
- ⏳ **Pendiente de F2a**: el frontend del wizard que consume esos helpers
  (sección legal con autocompletar de aliados, checkboxes de condiciones,
  botón gris hasta completar, encadenamiento T&C→ticket, campo de link
  publicado).

### 6.2 Recomendación: ENHANCE en vez de rewrite

`WebApp.Html` (el generador) ya son ~5.200 líneas con secciones, barra de
progreso, campos dinámicos (que YA incluirán Organizador tras el P0), autosave
y preview. **Reescribirlo de cero como wizard es caro y arriesgado** (y difícil
de verificar sin /dev). Propongo **evolucionar el form actual** para cumplir el
feedback de Anna (botón que se habilita al final, secciones renombradas
redención vs condiciones, checkboxes, autocompletar de aliados, pantalla de
revisión, link publicado) en incrementos aditivos y testeables — el mismo patrón
que venimos usando. Resultado equivalente para el usuario, riesgo mucho menor.
**A confirmar con Juan** antes de invertir en el wizard grande.
