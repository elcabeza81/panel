# "Cómo estoy" (Horizonte) en la pantalla HOY — cómo ponerlo en vivo

Rama: `horizonte-como-estoy`. **Nada de esto está publicado todavía.**

Son cinco piezas en dos repos. Hay que publicarlas **en este orden**. Cada una es aditiva:
si una falta, lo anterior sigue andando igual y el bloque dice "sin dato" o "sin contar la
plata guardada", nunca muestra un número falso.

| # | Pieza | Repo y archivo |
|---|---|---|
| 1 | Escritor **v2.3** | Proyecto-Horizonte · `apps_script/escritura_ha/Codigo_v2_3.gs` |
| 2 | Lector **v2.3** | Proyecto-Horizonte · `apps_script/lectura_ha/Codigo_v2_3.gs` |
| 3 | Paquete HA **rev8** | Proyecto-Horizonte · `home_assistant/horizonte_v2_rev8.yaml` |
| 4 | Backend del panel | panel · `backend_apps_script.gs` |
| 5 | Página del panel | panel · `index.html` |

## Qué cambia

- **Escritor v2.3:** una operación nueva, `plata_guardada`. Guarda la plata propia guardada
  del mes en una pestaña nueva de la planilla, `plata_guardada` (la crea sola). Lo demás
  queda igual que en v2.2.
- **Lector v2.3:** devuelve un bloque nuevo, `plata_guardada`, con el importe y el mes al
  que corresponde. Las claves que ya existían no cambian, así que HA sigue igual.
- **HA rev8:** cada vez que cambiás la plata guardada en el tablero, o tocás "Sigue igual",
  HA la copia en la planilla. También la copia todas las madrugadas a las 4:10 y al
  arrancar. Si la copia falla, deja una notificación en HA. El Libre de HA no cambia.
- **Backend del panel:** una acción nueva, `getHorizonte`, de sólo lectura. Le pide los
  datos al lector **del lado del servidor** y devuelve sólo los números ya calculados.
  Sólo responde con la sesión del **dueño** (el PIN de Juanma). `getResumen`, `addExp` y
  todo lo demás quedan igual.
- **Página:** el bloque "Cómo estoy" va entre las tarjetas de arriba y "Últimos gastos".
  Con el ojito cerrado muestra puntos. Con el PIN del dueño muestra los 4 datos y la hora
  del dato. En la vista compartida no aparece.

El token del lector **nunca llega a la página**: vive sólo en las Script Properties del
Apps Script del panel.

## 1 — Escritor v2.3

Apps Script **"Horizonte - Escritor HA"**: reemplazar el código por `Codigo_v2_3.gs` y
guardar. Después, **Implementar → Gestionar implementaciones** → ✏️ → Versión: **Nueva
versión** → Implementar (hoy está en la versión 5). **Misma URL.**

## 2 — Lector v2.3

Apps Script **`lectura_ha`**: reemplazar el código por `Codigo_v2_3.gs`, guardar y publicar
una versión nueva igual que en el paso 1. **Misma URL.**

## 3 — Home Assistant rev8

1. Copiar `horizonte_v2_rev8.yaml` sobre `/homeassistant/packages/horizonte_v2.yaml`.
2. Herramientas para desarrolladores → **Comprobar configuración**.
3. Recargar **scripts** y **automatizaciones**. No hace falta reiniciar.
4. Para la primera copia, sin esperar a las 4:10: Herramientas para desarrolladores →
   Acciones → `script.horizonte_plata_guardada_copiar` → Ejecutar.
5. Comprobar: en la planilla aparece la pestaña `plata_guardada` con una fila (mes e
   importe). Si no aparece, HA deja una notificación que explica por qué.

Para volver atrás: copiar de nuevo `horizonte_v2_rev7.yaml` y recargar scripts y
automatizaciones. La pestaña `plata_guardada` puede quedar en la planilla.

## 4 — Backend del panel

1. Apps Script **"Panel Juanma"** → ⚙️ Configuración del proyecto → Propiedades del script
   → agregar (**las carga Juanma**):

   | Propiedad | Valor |
   |---|---|
   | `HORIZONTE_LECTOR_URL` | la URL `/exec` del lector (`lectura_ha`). Es la misma que usa Home Assistant en su `secrets.yaml`, **sin** el `?token=…` |
   | `HORIZONTE_LECTOR_TOKEN` | el token del lector (la propiedad `HA_TOKEN` del Apps Script `lectura_ha`) |

2. Reemplazar el código por `backend_apps_script.gs` de esta rama y guardar.
3. Elegir la función **`probarHorizonte`** → Ejecutar. La primera vez puede pedir permiso
   para "conectarse a un servicio externo" (es el pedido al lector). En el registro tiene
   que aparecer el JSON con `libre` (y adentro `plata_guardada`), `proximo`, `tarjeta` y
   `obligaciones`. El token no aparece en el registro.
4. **Implementar → Gestionar implementaciones** → ✏️ → **Nueva versión** → Implementar.
   **Misma URL**: la página y Home Assistant no cambian nada.

Para volver atrás: Gestionar implementaciones → ✏️ → elegir la versión anterior.

## 5 — Página

Mergear la rama a `main` y hacer push. GitHub Pages publica solo en un par de minutos.

## Cómo se calcula cada dato

1. **Te queda libre este mes**: la misma fórmula que `sensor.horizonte_libre` de HA:
   caja del panel + plata guardada + lo que entró este mes aparte del acuerdo (extras y
   sueldo) − obligaciones del mes (pagadas y por pagar, reales o estimadas, sin lo que va
   por tarjeta) − gastos fijos en efectivo. "Por día" = Libre ÷ días que quedan, contando
   hoy, igual que HA. Si la plata guardada es de un mes anterior (todavía no tocaste "Sigue
   igual"), la usa igual que HA y lo avisa en amarillo.
2. **Lo próximo que sale**: el pago más grande de los próximos 7 días (o uno vencido). Si
   no hay ninguno, el más cercano. Sin lo que se paga con tarjeta.
3. **Tarjeta abierta**: lo que ya lleva el resumen abierto de la tarjeta principal, cuándo
   cierra, cuánto se estima al cierre y cuándo se paga.
4. **Por pagar este mes**: cuántas obligaciones faltan pagar, de cuántas, cuánto suman y
   cuál es la próxima.

El lector tarda unos segundos, así que el backend guarda la respuesta 5 minutos. Si el
lector no responde, muestra el último dato bueno (de hasta 6 horas) con su hora, en
amarillo. Si un dato falta, dice "sin dato" y nunca muestra un 0.
