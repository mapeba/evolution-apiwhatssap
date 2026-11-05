# `.railway/` — ajustes de Railway como código

`railway.ts` dice cómo se construye y arranca el servicio **evolution-apiwhatssap** en Railway.
Reemplaza a `railway.json`, que Railway deja de leer el **2026-12-01**.

- Railway **no lee esta carpeta al desplegar**: un cambio aquí llega con `railway config apply`.
- `npm ci` aquí instala el SDK (`railway` 3.12.0 exacto) y TypeScript; `npm run tipos` revisa el archivo.
- `railway config plan` (solo lee) debe dar únicamente «Update» de build/deploy, 0 altas y 0 borrados.
- `railway config apply` lo escribe en Railway y **redespliega** el servicio: decidirlo con Mauricio.
- Lo que el archivo declara, lo posee: variable nueva en Railway = añadir su nombre a `VARIABLES`.
- Después de aplicar y comprobar, se puede retirar `railway.json`.
