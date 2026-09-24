---
name: r1-flujo
description: Elaborar la Resolución 1 (traslado de apelación) de expedientes CC1-APELACIÓN a partir del Excel de control y R1_DATOS, con verificación, capturas y reporte. Úsala para cualquier paquete de R1 de apelaciones.
---
# Flujo de una remesa de R1 de apelaciones

1. Lee `AGENTS.md` completo.
2. Guarda la lista de la remesa en `entrada/paquete_NN.txt` (un expediente por línea o separados por espacio).
3. `python scripts/r1.py datos --exps $(cat entrada/paquete_NN.txt)`.
4. `python scripts/r1.py leer --exps $(cat entrada/paquete_NN.txt)` y muestra al usuario, por expediente, **solo lo que falta** (columnas de R1_DATOS). No rellenes por deducción: trato, fecha de ingreso, resolución apelada, vía de La Positiva y de proveedores nuevos se piden.
5. Con los datos completos: `python scripts/r1.py generar --exps … --fecha DD/MM/AAAA`.
6. Mira **todas** las capturas `salida/_vista/*.png` y lee cada `_FORMATO.md`. Entrega solo las R1 «GENERADO · APTO», con el `_REPORTE.md`.
7. Si una diferencia revela un error del sistema: corrígelo, corre `python scripts/simular.py`, añade la refutación en `docs/refutaciones.md` y vuelve a generar.
