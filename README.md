# Stock de asfalto · salida de mezcla — Planta asfáltica INGECO

Carga diaria del plantista (mezcla que salió por destino, asfalto recibido, asfalto en tanque) y panel con consumo real de asfalto, % real vs teórico, stock y autonomía. Frontend en GitHub Pages + backend Apps Script/Google Sheets. Acceso por PIN o desde la App INGECO (módulo `ASFALTO`).

## Qué hace
- **Cargar día**: toneladas de mezcla por destino (CONFIG `DESTINOS`), asfalto recibido (t) y stock del tanque al cierre (t). Vista previa POKAYOKE: stock esperado, consumo y % que resultan antes de guardar.
- **Consumo de asfalto** = stock anterior + recibido − stock actual, entre dos mediciones consecutivas; se reparte por día según la mezcla. **% real** = consumo ÷ mezcla del tramo, comparado con `PORC_TEORICO` ± `TOLERANCIA`.
- **Panel**: tanque (t, % lleno, autonomía), % último tramo y promedio, mezcla y consumo por día; gráficos de curvas (mezcla + consumo + recibido), % real vs teórico y stock (medido / estimado / mínimo).
- **Avisos** por la App INGECO en cada cierre (urgente si el stock baja del mínimo).
- **Integración Autovías**: `POST {accion:'mezcla', clave:<CLAVE_API>, payload:{destino, desde}}` devuelve la mezcla por día; Obra Autovías la usa como toneladas colocadas de los días que no cargó a mano.

## Arquitectura
- `docs/` → GitHub Pages → https://asfalto.ingecosa.com.ar/ (PWA).
- `Codigo.gs` → Apps Script vinculado a la planilla **STOCK ASFALTO** (hojas CONFIG, DIARIO, SALIDAS). API JSON por `doPost` text/plain `{accion, pin|sesion, nombre, payload, con_all}`: `login`, `datos`, `guardarDia`, `borrarDia`, `config`, `guardarConfig`, `mezcla` (con `clave`).
- Setup remoto: Script Property `SETUP_KEY` + `GET /exec?accion=setup&k=<SETUP_KEY>` (clasp run no funciona sin proyecto GCP).

## Deploy
- Backend: `clasp push -f && clasp deploy -i <deploymentId> -d "cambio"`.
- Frontend: commit + `git push` a `master` (Pages publica `/docs`; caché 10 min).
