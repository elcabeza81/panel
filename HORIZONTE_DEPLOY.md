# "Cómo estoy" (Horizonte) en la pantalla HOY — cómo ponerlo en vivo

Rama: `horizonte-como-estoy`. **Nada de esto está publicado todavía.**

## Qué cambia

- **Backend** (`backend_apps_script.gs`): una acción nueva, `getHorizonte`, de sólo lectura.
  Le pide los datos al lector de Horizonte **del lado del servidor** y devuelve sólo los
  números ya calculados. Sólo responde con una sesión del **dueño** (PIN de Juanma). Con la
  sesión de la otra mamá/papá o sin sesión devuelve `sesion invalida`.
  `getResumen`, `addExp` y todo lo demás quedan igual.
- **Página** (`index.html`): un bloque "Cómo estoy" entre las tarjetas de arriba y
  "Últimos gastos". Con el ojito cerrado muestra puntos. Con el PIN del dueño muestra los 4
  datos y la hora del dato. Con la vista compartida (`ex`) no aparece.

El token del lector **nunca llega a la página**: vive sólo en las Script Properties del
Apps Script del panel.

## Paso 1 — Script Properties (las carga Juanma)

Apps Script **"Panel Juanma"** → ⚙️ Configuración del proyecto → Propiedades del script →
Agregar:

| Propiedad | Valor |
|---|---|
| `HORIZONTE_LECTOR_URL` | la URL `/exec` del lector de Horizonte (`lectura_ha`). Es la misma que usa Home Assistant (en su `secrets.yaml`), **sin** el `?token=…` |
| `HORIZONTE_LECTOR_TOKEN` | el token del lector (la propiedad `HA_TOKEN` del Apps Script `lectura_ha`) |
| `HORIZONTE_PLATA_GUARDADA` | *(opcional)* la "plata guardada" que cargás en el tablero de HA, sólo el número (ej. `200000`) |

Sobre `HORIZONTE_PLATA_GUARDADA`: ese número vive **sólo en Home Assistant**, no en la
planilla, así que el panel no lo puede leer. Si no lo cargás, el bloque calcula el Libre
sin él y lo avisa ("sin contar la plata guardada"). Si lo cargás, cuando lo cambies en HA
también hay que cambiarlo acá.

## Paso 2 — Backend

1. En el editor de "Panel Juanma", reemplazar el código por el de `backend_apps_script.gs`
   de esta rama y guardar.
2. Elegir la función **`probarHorizonte`** → Ejecutar. La primera vez puede pedir permiso
   para "conectarse a un servicio externo" (es el pedido al lector). En el registro tiene
   que aparecer el JSON con `libre`, `proximo`, `tarjeta` y `obligaciones`. El token no
   aparece en el registro.
3. **Implementar → Gestionar implementaciones** → ✏️ en la implementación actual →
   Versión: **Nueva versión** → Implementar. **Misma URL**: la página y Home Assistant no
   cambian nada.

Para volver atrás: Gestionar implementaciones → ✏️ → elegir la versión anterior.

## Paso 3 — Página

Mergear la rama a `main` y hacer push. GitHub Pages publica solo en un par de minutos.

Si la página sale antes que el backend, el bloque muestra "sin dato · Horizonte no
respondió" y nada más se rompe.

## Cómo se calcula cada dato

1. **Te queda libre este mes**: la misma fórmula que `sensor.horizonte_libre` de HA:
   caja del panel + plata guardada + lo que entró este mes aparte del acuerdo (extras y
   sueldo) − obligaciones del mes (pagadas y por pagar, reales o estimadas, sin lo que va
   por tarjeta) − gastos fijos en efectivo. "Por día" = Libre ÷ días que quedan, contando
   hoy, igual que HA.
2. **Lo próximo que sale**: el pago más grande de los próximos 7 días (o uno vencido). Si
   no hay ninguno, el más cercano. Sin lo que se paga con tarjeta.
3. **Tarjeta abierta**: lo que ya lleva el resumen abierto de la tarjeta principal, cuándo
   cierra, cuánto se estima al cierre y cuándo se paga.
4. **Por pagar este mes**: cuántas obligaciones faltan pagar, de cuántas, cuánto suman y
   cuál es la próxima.

El lector tarda unos segundos, así que el backend guarda la respuesta 5 minutos. Si el
lector no responde, muestra el último dato bueno (de hasta 6 horas) con su hora, en
amarillo. Si un dato falta, dice "sin dato" y nunca muestra un 0.
