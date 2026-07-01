# Runbook — FASE C (organizador en Cashback) + A2

Secuencia para escribir los 2 campos de organizador en `Template_Fields` de Cashback y
activar la validación pre-entrega (A2). **Correr en el proyecto Apps Script de _pruebas_
(/dev), no en producción**, hasta validar end-to-end.

> Todas las funciones son deterministas y acotadas a `campaign_type='Cashback'`. El writer es
> **ADD-only + UPSERT idempotente** (llave `country_code+campaign_type+placeholder`): re-correrlo
> no duplica ni borra. **Nunca** toca filas `ALL` ni `Concurso`.

## Prerrequisito
Desplegar el código de la rama `claude/project-recovery-status-ug0rso` al Apps Script de /dev
(`clasp push` con `.clasp.json` apuntando al scriptId de /dev, o copiar los archivos al editor).

## Paso 1 — Confirmar el diff (NO escribe)
```js
previewFieldDerivation('Cashback','ALL')
```
Esperado: `diff.adds` = 2 (`organizerLegalName`, `organizerTaxId`), `wouldModify` 0, `deletes` 0,
`untouchedOtherTypes` = `{ALL:14, "Concurso Mayor Comprador":16}`.
**Si difiere de esto → PARAR** (el estado de la hoja cambió; avisar antes de escribir).

## Paso 2 — Escribir (ADD-only)
```js
applyFieldDerivation('Cashback','ALL')
```
Esperado:
```
written: { added: 2, updated: 0 }
before:  { ALL:14, Cashback:11, "Concurso Mayor Comprador":16, ... }
after:   { ALL:14, Cashback:13, "Concurso Mayor Comprador":16, ... }
```
Verificar: **ALL sigue 14, Concurso sigue 16, Cashback pasó a 13**.

## Paso 3 — Idempotencia (correr OTRA VEZ)
```js
applyFieldDerivation('Cashback','ALL')
```
Esperado: `written: { added: 0, updated: 0 }` y `after` idéntico (Cashback sigue 13).
Esto prueba que **no** reproduce el borrado de filas `ALL` del writer viejo.

## Paso 4 — Activar A2
Configuración del proyecto → Propiedades del script → agregar:
```
RAPPIMIND_A2 = on
```

## Paso 5 — Prueba end-to-end (Cashback CO)
Abrir el web app → Colombia → Cashback → llenar el formulario (ahora incluye el grupo
**"Organizador"**: razón social + ID fiscal) → Generar.
Esperado: documento con **cero `{{}}` y cero `[ ]`**, organizador poblado, jurisdicción/ley de CO.
Si queda algún marcador, A2 aborta con el mensaje `A2_ABORT: ...` indicando el token culpable.

## Rollback
- Quitar los 2 campos: borrar en `Template_Fields` las filas `(ALL, Cashback, {{ORGANIZADOR}})` y
  `(ALL, Cashback, {{ID_ORGANIZADOR}})`.
- Desactivar A2: Propiedades del script → `RAPPIMIND_A2 = off` (o borrar la propiedad).
- El código no cambia el comportamiento con A2 en `off`.
