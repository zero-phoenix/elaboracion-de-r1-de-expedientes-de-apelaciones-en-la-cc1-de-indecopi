# R1 de expedientes de apelación — CC1 Indecopi

Genera en Word la **Resolución 1** de los expedientes de apelación que recibe la Comisión de Protección al Consumidor N° 1: traslado del recurso o recursos, requerimientos de notificación y, según el caso, agregado de escritos o citación a audiencia. Parte de un Excel de control y de cinco plantillas extraídas de las R1 ya revisadas. No elabora cédulas.

> **Última actualización — v1.1.0 (24/09/2026):** primera R1 del paquete 1 (0259-2026) elaborada desde los documentos del expediente; partes del encabezado una por línea; fecha de emisión por expediente; la cédula física del expediente fija la vía «domicilio»; `fojas.py` separa escritos de parte de documentos de Indecopi y lee la fecha de presentación de la firma de Mesa de Partes; verificación automática y Releases en GitHub Actions. Detalle en [CHANGELOG.md](CHANGELOG.md).

Reglas vigentes: **[AGENTS.md](AGENTS.md)**. Cómo empezar en otra computadora: **[ARRANQUE.md](ARRANQUE.md)**.

```bash
pip install -r requirements.txt
python scripts/r1.py datos    --exps 000663-2026 000685-2026   # filas en entrada/R1_DATOS.xlsx
python scripts/r1.py leer     --exps 000663-2026 000685-2026   # salida/_REPORTE.md
python scripts/r1.py generar  --exps 000663-2026 000685-2026 --fecha 24/09/2026
python scripts/vista.py "salida/R1 0663-2026-CC1-APE.docx" --contra plantillas/S1_simple.docx
python scripts/simular.py                                        # control popperiano sobre el corpus
python scripts/fojas.py 0259-2026                                # fojas y vías según los documentos del expediente
```

| Carpeta | Contenido |
|---|---|
| `plantillas/` | S1 simple · S2 con escritos · S3 dos apelaciones · S4 audiencia · S5 inadmisibilidad/improcedencia liminar |
| `scripts/` | `r1.py` (CLI) · `modelo.py` (datos y validación) · `textos.py` y `redaccion.py` (fórmulas y guion) · `documento.py` (Word y verificador) · `vista.py` (capturas y ficha de formato) · `simular.py` (simulación) |
| `docs/` | directorio de notificación por proveedor · fórmulas literales · corpus clasificado · refutaciones |
| `corpus/` | 58 R1 revisadas y 86 cédulas de referencia |
| `entrada/` | Excel de control (`APES_R1_SEGUROS.xlsx`), `R1_DATOS.xlsx`, listas de paquetes |
| `pruebas/` | resultado de la simulación y pruebas de paquete |
