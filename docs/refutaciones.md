# Refutaciones (registro popperiano)

Cada entrada: **hipótesis** que se tenía → **contraejemplo** hallado en el corpus o en la simulación → **decisión**. Se añaden, no se borran.

| # | Hipótesis | Contraejemplo | Decisión |
|---|---|---|---|
| 1 | La vía de un proveedor se lee en el correo de la cédula («casilla@…» = casilla). | Rímac se notifica a `CORREO-E casilla@er.com.pe`. | La vía se toma del **prefijo** de la cédula (CORREO-E / CASILLA-E / dirección). |
| 2 | Cada proveedor tiene una sola vía. | La Positiva: correo en 731 y 450, casilla en 474, 663 y 685 (todas de setiembre de 2026). | La Positiva queda **mixta**: la vía se indica en cada caso (DENUNCIADOS_VIAS). |
| 3 | El denunciante siempre se notifica por correo. | 608: el denunciante tiene CASILLA-E y se le exhorta. 347: también casilla. | Correo por defecto (con advertencia); DENUNCIANTE_VIA manda. |
| 4 | La fecha de «El …, la Comisión … ha recibido» está en el Excel. | 663: Excel 09/07 y R1 10/07; 742: Excel 24/07 y R1 30/07; 608: Excel 30/06 y 01/07, R1 31/07. | FECHA_INGRESO_CC1 es obligatoria en R1_DATOS; no se deduce. |
| 5 | Las R1 revisadas («OK», «LSQok») no tienen errores. | 726: nombra como apelante al «señor Carlos Andrés Gómez Ramos», abogado de Interseguro; la cédula y el Excel dicen José Luis Marcelo López. | Control con salvedad: 726 queda **PENDIENTE** hasta que se confirme. |
| 6 | Idem. | 608: el texto dice «Expediente N° 1654-2025/PS1», el encabezado y el Excel dicen 253-2026/PS1. | Se usa el expediente de origen del encabezado/Excel. |
| 7 | Idem. | 0199: «Expediente N° 0878-2025» sin «/PS1». 0351: el nombre aparece en dos órdenes distintos. | Formato fijo NNNN-AAAA/PSn; nombre único por caso. |
| 8 | Idem. | 517: dice que la apelación la presentó la denunciante, pero el Excel dice que apeló Afocat. | Se sigue el Excel/expediente; la R1 original no es base para ese dato. |
| 9 | El modelo S4 (206) es literal en todo. | 206 traslada el recurso a Rímac, que es la propia apelante; y «reciba … efectúen» mezcla singular y plural. | El traslado va siempre a la parte **no apelante**; la concordancia sigue al número de destinatarios. |
| 10 | Las erratas de concordancia del corpus son la forma correcta. | 629, 731: «póngase en conocimiento de A y B … haga … aporte». 0057: «su bandeja … sus correos». | Se concuerda: «hagan/aporten» con dos o más; «su correo electrónico» con uno. |
| 11 | Todas las R1 recientes firman LSQ. | 0017–0473 firman «LUISA ANALI SILVA MALPARTIDA, Ejecutivo 1/2» (LGP/jcq). | Firma única vigente: LSQ (instrucción del usuario). |
| 12 | Las R1 citan el TUO vigente. | 0098, 0100, 0166 citan el DS 004-2019-JUS. | Siempre DS 006-2026-JUS; el verificador lo bloquea. |
| 13 | La negrita de la línea «Comisión de Protección al Consumidor N° 1» de la firma es fija. | Los modelos la llevan en negrita; algunas entregadas (0057, 206), no. | Se respeta la plantilla del supuesto (nunca se fuerza). |
| 14 | El encabezado conserva su formato si se reescribe el texto. | Primera versión: DENUNCIANTE/DENUNCIADO(S) perdían la negrita (lo detectó el visor). | Las celdas del encabezado heredan la negrita de la plantilla. |
| 15 | La fecha y la hora de la audiencia van solo en negrita. | Modelo 206: negrita **y subrayado**. | Marca `__…__` para el subrayado; el visor lo comprueba. |
| 16 | Si hay varios apelantes, hay varios recursos (S3). | 0259: dos denunciantes apelan con un solo recurso. | Varios apelantes del mismo lado = un recurso (S1); de lados distintos = S3. |
| 17 | S1 no admite escritos. | 480, 517, 726 y 731 trasladan en S1 el recurso y un escrito del propio apelante. | S1 admite escritos del apelante; si la otra parte presentó escritos, S2. |
| 18 | 0100 notifica a todas las partes. | 0100 traslada al BCP pero no le requiere acuse. | El sistema notifica siempre a todas las partes trasladadas. |
| 19 | LibreOffice sirve para las capturas. | En este entorno no carga ni un .txt. | Visor propio (`scripts/vista.py`): mide el XML y fotografía con Chromium. |
