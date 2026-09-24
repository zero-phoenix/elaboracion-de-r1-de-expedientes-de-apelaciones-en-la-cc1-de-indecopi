# Cambios

## v1.2.0 — 24/09/2026
- Encabezado: rótulo **«DENUNCIANTE(S)» solo cuando hay dos o más denunciantes** (con uno, «DENUNCIANTE»).
- R1 del **0259-2026** cerrada (fecha 18/09/2026; Quálitas por domicilio; denunciantes en líneas separadas).
- **Descargable del sistema en Releases**: `sistema-r1-apelaciones-<versión>.zip` (scripts, plantillas, configuración, documentación, directorio y corpus; sin documentos de expedientes ni Word generados).
- Envío por correo del Word generado (instrucciones en AGENTS.md).

## v1.1.0 — 24/09/2026
- **Expediente 0259-2026** elaborado con sus documentos (primera R1 del paquete 1).
- Encabezado: varias partes, **una por línea** (nunca «/»); el verificador lo exige.
- **Fecha de emisión por expediente** (columna FECHA_EMISION de R1_DATOS).
- **Cédula física en el expediente → vía domicilio** para esa parte, aunque el directorio diga otra cosa.
- `fojas.py`: distingue escritos de parte (se cuentan) de documentos de Indecopi (RF, cédulas, constancias, elevación: no se trasladan); toma la fecha de presentación de la firma automatizada de Mesa de Partes; lista las vías que constan en el expediente y convierte los PDF escaneados en imágenes para leerlos.
- GitHub Actions: verificación en cada push (simulación sobre el corpus y verificador) y publicación automática en Releases de la versión que encabeza este CHANGELOG.
- Refutaciones 20–24 incorporadas.

## v1.0.0 — 24/09/2026
- Sistema inicial: cinco plantillas (S1–S5), `r1.py` (datos · leer · generar · verificar), directorio de notificación desde las cédulas, visor propio con capturas y ficha de formato, simulación popperiana sobre 58 R1 (similitud 0,989), conteo de fojas.
