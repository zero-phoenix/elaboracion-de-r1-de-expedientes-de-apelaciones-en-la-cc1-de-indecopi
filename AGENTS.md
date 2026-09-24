# AGENTS.md — R1 de expedientes de apelación · CC1 Indecopi (v1.2)

Única fuente de reglas vigentes de este repositorio. Solo se elabora la **Resolución 1 (R1)** que da trámite a la apelación elevada a la Comisión de Protección al Consumidor N° 1: **traslado del recurso o recursos**, requerimientos de notificación y, cuando corresponde, agregado de escritos o citación a audiencia. **No se elaboran cédulas** (las del corpus son solo referencia de vías).

## 0. Arranque
1. Trabaja desde la raíz del repositorio. `pip install -r requirements.txt`.
2. Fecha de emisión: `config/remesa.json` o `--fecha DD/MM/AAAA`. Si está sin fijar, **pregúntala**; no se genera sin fecha.
3. Insumos: `entrada/APES_R1_SEGUROS.xlsx` (Excel de control, hoja «Recibidas 2026»; se reemplaza cuando el usuario lo actualiza) y `entrada/R1_DATOS.xlsx` (datos que el Excel no trae).

## 1. Los pasos
1. `python scripts/r1.py datos --exps 000663-2026 …` — añade las filas a `entrada/R1_DATOS.xlsx`. Si el expediente ya tiene R1 revisada en `corpus/r1/`, la fila se llena desde ella (**control**); si no, solo con lo que el Excel **afirma**.
2. Completar `R1_DATOS` (usuario o agente, **solo con datos que consten** en el expediente o la cédula).
3. `python scripts/r1.py leer --exps …` — `salida/_REPORTE.md`: APTO / OBSERVADO y qué falta.
4. `python scripts/r1.py generar --exps … [--fecha DD/MM/AAAA]` — `salida/R1 NNNN-AAAA-CC1-APE.docx`, verificación automática y `salida/_vista/` (captura lado a lado con la plantilla + ficha de formato).
5. Revisar **todas** las capturas de `_vista/` antes de entregar. Solo se entrega lo que dice **GENERADO · APTO**.

6. **Fojas** (no van en la R1; solo se informan): documentos del expediente en `entrada/expedientes/<NNNN-AAAA>/` y `python scripts/fojas.py <NNNN-AAAA>` → `_FOJAS.md` con la forma de las cédulas («Copia del escrito de apelación presentado por la parte denunciante el 12/06/2026 (8 fojas).»). Cuentan las hojas del escrito y sus anexos; no cuentan las páginas en blanco ni la constancia o cargo automático de Mesa de Partes. Cada exclusión se informa para que el usuario la confirme. Los documentos originales no se suben al repositorio.

7. **Entrega por correo** (si el usuario lo pide): el Word APTO se envía adjunto desde la cuenta de Gmail conectada (conector Gmail de Claude: `send_message` con el .docx en base64) a la dirección del usuario, con asunto «R1 NNNN-AAAA/CC1-APELACIÓN». Sin conector, se entrega el archivo en el chat.

## 2. Reglas no negociables
1. **Nada que no conste.** Un dato ausente no se deduce: la fila queda OBSERVADA. No se infiere señor/señora del nombre; no se inventa la vía de un proveedor.
2. **Solo R1.** Nunca cédulas, oficios ni otros actos.
3. **Plantillas cerradas** (`plantillas/`): el formato (página, encabezado, estilos, numeración, notas al pie, firma) viene de la plantilla; el texto, de `scripts/textos.py` y `scripts/redaccion.py`. No se redacta libremente ni se edita el Word a mano.
4. **Firma única:** LOUSSIANA CATHERINE SALAZAR QUIROZ, «Especialista Legal», «Comisión de Protección al Consumidor N° 1», iniciales LSQ/DCQ (`config/firma.json`), **también cuando el denunciado es Rímac**.
5. **TUO de la LPAG:** Decreto Supremo N° 006-2026-JUS. Nunca 004-2019-JUS.
6. **Erratas de las R1 previas no se copian** (ver `docs/refutaciones.md`).
7. **Los documentos del expediente mandan** sobre el Excel, el directorio y las R1 previas: fecha de presentación (firma del agente automatizado de Mesa de Partes), resolución apelada, tratamiento («el señor García», «la señora Alvarado» en la RF), vías.
8. **Encabezado:** con dos o más denunciantes o denunciados, **cada uno en su propia línea** dentro de la celda; nunca separados con «/». Rótulo «DENUNCIANTE» con uno y «DENUNCIANTE(S)» **solo** con dos o más; «DENUNCIADO(S)» siempre (como el corpus).
9. **Fecha de emisión por expediente:** la columna FECHA_EMISION de R1_DATOS, si se llena, manda sobre la de la remesa.

## 3. Los cinco supuestos
| | Supuesto | Plantilla | Cuándo |
|---|---|---|---|
| S1 | Simple | `S1_simple.docx` (modelo 0057-2026) | Un recurso (también el conjunto de varios denunciantes), sin escritos de la otra parte. Admite escritos del propio apelante: «el recurso de apelación y el escrito del …, presentados por …». |
| S2 | Con escritos | `S2_con_escritos.docx` (modelo 130-2026) | La otra parte presentó escritos, o hay dos recursos con escritos. VISTOS / CONSIDERANDO / SE RESUELVE con ítems (i), (ii)… y «Agregar los referidos escritos…». |
| S3 | Dos apelaciones | `S3_dos_apelaciones.docx` (modelo 0738-2025) | Apelan partes distintas y no hay escritos: «póngase en conocimiento de las partes del procedimiento los recursos de apelación». |
| S4 | Audiencia de conciliación | `S4_audiencia.docx` (modelo 206-2026) | Una parte pidió audiencia: fecha, hora, requisitos a)–d). |
| S5 | Inadmisibilidad / improcedencia liminar | `S5_inadmisibilidad_liminar.docx` (R1 166-2026) | Se apela la inadmisibilidad o improcedencia liminar: **nunca hubo admisión ni imputación**, **solo se notifica al denunciante apelante**; no hay traslado al denunciado. |

El supuesto se deduce de los datos o se fija en la columna SUPUESTO.

## 4. Notificación — un párrafo por VÍA (no por parte)
- **Correo** → «requerir … confirmación de recepción … dos (2) días hábiles … numeral 4 del artículo 20 del TUO …» (+ nota al pie del art. 20.4 en S2–S5).
- **Casilla** → «exhortar … acuse de recibo … Casilla Electrónica … cinco (5) primeros días hábiles».
- **Domicilio procesal** → «requerir … señale un correo electrónico autorizando recibir las notificaciones».
- Orden: correo, domicilio, casilla. Conectores S1/S5: «Asimismo, » … «Así también, » … «Finalmente, ». S3: «Requerir …» / «Exhortar …». S2/S4: ítems en minúscula que acaban en «;», «; y», «.».
- Singular/plural concuerdan con el número de destinatarios («reciba/reciban», «efectúe/efectúen», «notificarle/notificarles», «haga/hagan»).
- **Vía de cada proveedor:** `docs/directorio_notificacion.json`, sacado solo de las R1 y cédulas del corpus. **La Positiva** es **mixta** (correo y casilla en fechas cercanas): se indica en cada caso. Proveedor ausente del directorio: se pregunta.
- Denunciante: correo salvo que la cédula diga CASILLA-E o domicilio (DENUNCIANTE_VIA).
- **Cédula física en el expediente = vía domicilio para esa parte**, aunque el directorio diga otra cosa (caso 0259-2026: Quálitas, cédula física con CARGO en San Isidro → «requerir … en su domicilio procesal, señale un correo electrónico…»). `fojas.py` lista las cédulas del expediente y convierte en imágenes (`_hojas/`) los PDF escaneados para leerlos a la vista; **todo PDF escaneado se mira** antes de fijar vías.

## 5. Forma (medida en los modelos)
A4; márgenes 2,25 cm arriba, 3,5 cm abajo y 3,0 cm a los lados; Arial Narrow 11 justificado, interlineado sencillo, espaciado 0/0; tabla de encabezado INGRESO EN COMISIÓN / EXPEDIENTE DE ORIGEN Nº / DENUNCIANTE / DENUNCIADO(S) / RESOLUCIÓN Nº; «Lima, D de mes de AAAA» (mes en minúscula, «setiembre»); firma centrada en cuatro líneas; iniciales en 8 pt; notas al pie en 8 pt. Sin resaltados. El verificador (`documento.verificar`) y el visor (`scripts/vista.py`) lo comprueban.

## 6. Control de calidad popperiano
- `python scripts/simular.py` reproduce las R1 del corpus a partir de sus propios datos y las compara con las originales (`pruebas/simulacion.md`). Toda diferencia es una refutación: se corrige el sistema o se documenta como errata o variante en `docs/refutaciones.md`.
- Tras cualquier cambio en `scripts/` o `plantillas/`: correr `simular.py` y no aceptar que baje la similitud media ni que aparezca un NO APTO.

## 7. Estructura
`config/` firma y fecha · `plantillas/` S1–S5 · `scripts/` (r1, modelo, textos, redaccion, documento, vista, simular) · `docs/` directorio, fórmulas, corpus clasificado, refutaciones · `corpus/r1` R1 revisadas (58) · `corpus/cedulas` cédulas de referencia (86) · `entrada/` Excel y R1_DATOS · `pruebas/` simulación y pruebas de paquete · `salida/` (no versionada).
