"""Fórmulas literales de la R1 de apelaciones (CC1), tomadas de los modelos
de plantillas/ y de las R1 revisadas del corpus. Todo texto de la resolución
sale de aquí; nada se redacta libremente.

Convenciones:
- «{FN}» marca el lugar de la llamada a la nota al pie del art. 20.4 del TUO
  de la LPAG (la nota ya existe en la plantilla; el constructor la conserva).
- «**…**» marca negrita (solo rótulos VISTOS / CONSIDERANDO / SE RESUELVE).
"""

TUO = ("Texto Único Ordenado de la Ley del Procedimiento Administrativo General, "
       "aprobado mediante Decreto Supremo N° 006-2026-JUS")

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "setiembre", "octubre", "noviembre", "diciembre"]


def fecha_larga(d):
    """date -> «4 de setiembre de 2026»."""
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def fechas_juntas(ds):
    """[d1, d2, d3] -> «16 y 26 de diciembre de 2025» / «2 de julio, 13 de agosto y 18 de agosto de 2026»."""
    ds = list(ds)
    if len(ds) == 1:
        return fecha_larga(ds[0])
    mismo_mes = all((d.month, d.year) == (ds[0].month, ds[0].year) for d in ds)
    if mismo_mes:
        dias = [str(d.day) for d in ds]
        return f"{', '.join(dias[:-1])} y {dias[-1]} de {MESES[ds[0].month - 1]} de {ds[0].year}"
    mismo_anio = all(d.year == ds[0].year for d in ds)
    if mismo_anio:
        partes = [f"{d.day} de {MESES[d.month - 1]}" for d in ds]
        return f"{', '.join(partes[:-1])} y {partes[-1]} de {ds[0].year}"
    partes = [fecha_larga(d) for d in ds]
    return f"{', '.join(partes[:-1])} y {partes[-1]}"


# ---------------------------------------------------------------- partes ---
def _y(siguiente):
    """«y» / «e» según la palabra que sigue (e ante i-, hi- no diptongo)."""
    s = siguiente.lower()
    if (s.startswith("i") or s.startswith("hi")) and not s.startswith("hie"):
        return "e"
    return "y"


def con_de(p):
    """«de la señora X» / «del señor X» / «del Banco X» / «de Rímac …»."""
    if p.trato == "señor":
        return f"del señor {p.nombre}"
    if p.trato == "señora":
        return f"de la señora {p.nombre}"
    if p.nombre.startswith("Banco "):
        return f"del {p.nombre}"
    return f"de {p.nombre}"


def con_a(p):
    """«a la señora X» / «al señor X» / «al Banco X» / «a Rímac …»."""
    if p.trato == "señor":
        return f"al señor {p.nombre}"
    if p.trato == "señora":
        return f"a la señora {p.nombre}"
    if p.nombre.startswith("Banco "):
        return f"al {p.nombre}"
    return f"a {p.nombre}"


def sin_prep(p):
    """«el señor X» / «la señora X» / «el Banco X» / «Rímac …» (sujeto agente)."""
    if p.trato == "señor":
        return f"el señor {p.nombre}"
    if p.trato == "señora":
        return f"la señora {p.nombre}"
    if p.nombre.startswith("Banco "):
        return f"el {p.nombre}"
    return p.nombre


def lista(partes, fmt, repetir_prep=True):
    """Une partes: «a la señora X y a Rímac …» (requerir/exhortar) o
    «de Pacífico … y Chubb …» (póngase en conocimiento, sin repetir «de»)."""
    txt = [fmt(p) for p in partes]
    if not repetir_prep:
        txt = [txt[0]] + [sin_prep(p) for p in partes[1:]]
    if len(txt) == 1:
        return txt[0]
    return f"{', '.join(txt[:-1])} {_y(txt[-1])} {txt[-1]}"


# ---------------------------------------------------------- notificación ---
def parrafo_via(via, destinatarios, verbo_mayus=False, conector="", fn=False,
                cierre="."):
    """Un párrafo por VÍA. destinatarios: lista de Parte con la misma vía."""
    n = len(destinatarios)
    quien = lista(destinatarios, con_a)
    if via == "correo":
        verbo = "requerir"
        cuerpo = (
            f"{quien} para que, dentro del plazo de dos (2) días hábiles siguientes a la fecha en que "
            + ("reciba la notificación en su bandeja de correo electrónico, efectúe la confirmación de recepción de la notificación remitida por este despacho a su correo electrónico"
               if n == 1 else
               "reciban la notificación en sus bandejas de correo electrónico, efectúen la confirmación de recepción de la notificación remitida por este despacho a sus correos electrónicos")
            + f", de conformidad con el segundo párrafo del numeral 4 del artículo 20 del {TUO}, "
            + "bajo apercibimiento de rehacer el acto de notificación y "
            + ("notificarle" if n == 1 else "notificarles")
            + " conforme al numeral 1 del artículo 20 del citado cuerpo normativo"
            + ("{FN}" if fn else "")
        )
    elif via == "casilla":
        verbo = "exhortar"
        cuerpo = (
            f"{quien} para que "
            + ("efectúe el acuse de recibo mediante la confirmación de recepción de la notificación remitida por este despacho a su Casilla Electrónica, dentro de los cinco (5) primeros días hábiles siguientes a la fecha en que recibe la notificación"
               if n == 1 else
               "efectúen el acuse de recibo mediante la confirmación de recepción de la notificación remitida por este despacho a sus Casillas Electrónicas, dentro de los cinco (5) primeros días hábiles siguientes a la fecha en que reciben la notificación")
        )
    elif via == "domicilio":
        verbo = "requerir"
        cuerpo = (
            f"{quien} para que, dentro del plazo de dos (2) días hábiles siguientes a la fecha en que "
            + ("reciba la notificación en su domicilio procesal, señale un correo electrónico"
               if n == 1 else
               "reciban la notificación en sus domicilios procesales, señalen un correo electrónico")
            + " autorizando recibir las notificaciones correspondientes por dicho medio"
        )
    else:
        raise ValueError(f"vía desconocida: {via}")
    if verbo_mayus:
        verbo = verbo.capitalize()
    return f"{conector}{verbo} {cuerpo}{cierre}"


ORDEN_VIAS = ["correo", "domicilio", "casilla"]  # orden de los párrafos en el corpus


def agrupar_por_via(partes):
    grupos = []
    for via in ORDEN_VIAS:
        g = [p for p in partes if p.via == via]
        if g:
            grupos.append((via, g))
    return grupos


def conectores_s1(n):
    """S1/S5: «Asimismo, » … «Así también, » … «Finalmente, »."""
    if n == 1:
        return ["Asimismo, "]
    return ["Asimismo, "] + ["Así también, "] * (n - 2) + ["Finalmente, "]


# ------------------------------------------------------------- traslado ---
def traslado(destinatarios, plural_recursos=False, partes_del_proc=False):
    """Cuerpo común de «póngase / poner en conocimiento …»."""
    if partes_del_proc:
        quien = "de las partes del procedimiento"
    else:
        quien = lista(destinatarios, con_de, repetir_prep=False)
    varios = partes_del_proc or len(destinatarios) > 1
    recurso = "los recursos de apelación" if plural_recursos else "el recurso de apelación"
    sep = " " if partes_del_proc else ", "
    return (
        f"{quien}{sep}{recurso} para que, de considerarlo pertinente y en un plazo no mayor de cinco (5) días hábiles "
        "contado a partir del día siguiente de la recepción de la presente resolución, "
        + ("hagan" if varios else "haga")
        + " conocer su posición respecto de los argumentos expuestos en "
        + ("los recursos de apelación interpuestos" if plural_recursos else "el recurso de apelación interpuesto")
        + " y "
        + ("aporten" if varios else "aporte")
        + " cualquier elemento, hecho o fundamento que pueda ser de utilidad para resolver el asunto que es materia de discusión en esta instancia"
    )
