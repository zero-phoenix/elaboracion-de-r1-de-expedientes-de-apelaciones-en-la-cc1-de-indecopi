"""Del Caso al guion de párrafos de la R1, por supuesto (S1–S5).

Cada párrafo del guion es una tupla:
  ("p", texto)            párrafo normal nuevo (se clona el primer párrafo de texto de la plantilla)
  ("n", texto)            ítem numerado (i), (ii)… del SE RESUELVE (se clona el de la plantilla)
  ("k", prefijo)          se conserva TAL CUAL el párrafo de la plantilla que empieza así
  ("t", prefijo, texto)   se conserva el párrafo de la plantilla (formato, numeración) y se cambia su texto
El texto sale solo de textos.py y de los modelos; marcas: {FN} / {FN:id}, **negrita**.
"""
from textos import (agrupar_por_via, conectores_s1, fecha_larga, fechas_juntas, lista,
                    parrafo_via, sin_prep, traslado)

RECIBIO = "la Comisión de Protección al Consumidor Nº 1 ha recibido el Expediente N° {origen}"


def _contra(c):
    s = f" contra la {c.resolucion_apelada}" if c.resolucion_apelada else ""
    if c.extremo:
        s += f" {c.extremo}"
    return s


def _cierre(nombre):
    """Evita el doble punto cuando la razón social termina en «.»."""
    return "" if nombre.endswith(".") else "."


def _notificaciones(c, destinatarios, estilo, fn=True):
    """Párrafos por vía. estilo: «s1» (Asimismo…/Finalmente…), «s3» (Requerir…/Exhortar…),
    «lista» (ítems en minúscula que terminan en «;», «; y», «.»)."""
    grupos = agrupar_por_via(destinatarios)
    out, fn_puesta = [], False
    con = conectores_s1(len(grupos))
    for i, (via, g) in enumerate(grupos):
        pon_fn = fn and via == "correo" and not fn_puesta
        fn_puesta = fn_puesta or pon_fn
        if estilo == "s1":
            out.append(parrafo_via(via, g, conector=con[i], fn=pon_fn))
        elif estilo == "s3":
            out.append(parrafo_via(via, g, verbo_mayus=True, fn=pon_fn))
        else:
            ultimo = i == len(grupos) - 1
            out.append(parrafo_via(via, g, fn=pon_fn, cierre="." if ultimo else ";"))
    return out


def _cerrar_lista(items):
    """Ítems del SE RESUELVE: «;» … penúltimo «; y» … último «.»."""
    out = []
    for i, t in enumerate(items):
        t = t.rstrip(" .;")
        if i == len(items) - 1:
            out.append(t + ".")
        elif i == len(items) - 2:
            out.append(t + "; y")
        else:
            out.append(t + ";")
    return out


def _vistos(c):
    """«el recurso de apelación y el escrito presentado por X el 16 y 26 de diciembre de 2025,
    respectivamente; y el escrito del 19 de diciembre de 2025 presentado por Y»."""
    trozos = []
    for p in c.apelantes + [q for q in c.no_apelantes if q.escritos]:
        quien = sin_prep(p)
        if p.apelante:
            fs = [p.fecha_apelacion] + sorted(p.escritos)
            if p.escritos:
                esc = "el escrito" if len(p.escritos) == 1 else "los escritos"
                trozos.append(f"el recurso de apelación y {esc} presentados por {quien} el {fechas_juntas(fs)}, respectivamente"
                              if len(p.escritos) > 1 else
                              f"el recurso de apelación y el escrito presentado por {quien} el {fechas_juntas(fs)}, respectivamente")
            else:
                trozos.append(f"el recurso de apelación presentado por {quien} el {fecha_larga(p.fecha_apelacion)}")
        else:
            fs = sorted(p.escritos)
            if len(fs) == 1:
                trozos.append(f"el escrito del {fecha_larga(fs[0])} presentado por {quien}")
            else:
                trozos.append(f"los escritos del {fechas_juntas(fs)} presentados por {quien}")
    if len(trozos) == 1:
        texto = trozos[0]
    else:
        texto = "; ".join(trozos[:-1]) + "; y " + trozos[-1]
    return f"**VISTOS:** {texto}" + _cierre(texto)


def _interpuesto(c):
    ap = c.apelantes
    if len(ap) == 1:
        return f"el recurso de apelación interpuesto por {sin_prep(ap[0])}"
    return f"los recursos de apelación interpuestos por {lista(ap, sin_prep)}"


# ------------------------------------------------------------------- S1 ---
def s1(c):
    a = c.apelantes[0]
    quien = lista(c.apelantes, sin_prep)
    pre = f"El {fecha_larga(c.fecha_ingreso)}, " + RECIBIO.format(origen=c.origen) + ", el cual contiene "
    if a.escritos:
        esc = (f"el escrito del {fecha_larga(a.escritos[0])}" if len(a.escritos) == 1
               else f"los escritos del {fechas_juntas(sorted(a.escritos))}")
        p1 = pre + f"el recurso de apelación y {esc}, presentados por {quien}{_contra(c)}"
    else:
        p1 = pre + f"el recurso de apelación interpuesto por {quien}{_contra(c)}"
    p1 += _cierre(p1)
    p2 = "En ese sentido, póngase en conocimiento " + traslado(c.no_apelantes) + "."
    return [("p", p1), ("p", p2)] + [("p", t) for t in _notificaciones(c, c.partes, "s1", fn=False)]


# ------------------------------------------------------------------- S2 ---
def s2(c):
    considerando = (f"**CONSIDERANDO:** que, el {fecha_larga(c.fecha_ingreso)}, " + RECIBIO.format(origen=c.origen)
                    + f", el cual contiene {_interpuesto(c)}{_contra(c)}")
    considerando += _cierre(considerando)
    varios = len(c.apelantes) > 1
    items = ["Agregar los referidos escritos al expediente y ponerlos en conocimiento de la otra parte del procedimiento",
             "poner en conocimiento " + traslado(c.no_apelantes, plural_recursos=varios, partes_del_proc=varios)]
    items += _notificaciones(c, c.partes, "lista")
    return ([("p", _vistos(c)), ("t", "CONSIDERANDO", considerando), ("k", "SE RESUELVE")]
            + [("n", t) for t in _cerrar_lista(items)])


# ------------------------------------------------------------------- S3 ---
def s3(c):
    a1, a2 = c.apelantes[0], c.apelantes[1]
    p1 = (f"El {fecha_larga(c.fecha_ingreso)}, " + RECIBIO.format(origen=c.origen).replace("N° ", "Nº ")
          + f" remitido como resultado de los recursos de apelación interpuestos por {sin_prep(a1)} y {sin_prep(a2)}{_contra(c)}")
    p1 += _cierre(p1)
    p2 = "En ese sentido, póngase en conocimiento " + traslado([], plural_recursos=True, partes_del_proc=True) + "."
    return [("p", p1), ("p", p2)] + [("p", t) for t in _notificaciones(c, c.partes, "s3")]


# ------------------------------------------------------------------- S4 ---
def s4(c):
    sol = next((p for p in c.partes if p.rol == c.audiencia_solicitante.lower()
                or c.audiencia_solicitante.lower() in p.nombre.lower()), None)
    if c.audiencia_solicitante.upper() == "DTE":
        sol = c.denunciantes[0]
    elif c.audiencia_solicitante.upper().startswith("DDO"):
        sol = c.denunciados[0]
    hay_escritos = any(p.escritos for p in c.partes)
    guion = []
    if hay_escritos or c.apelantes[0].fecha_apelacion:
        guion.append(("t", "VISTOS", _vistos(c)))
    guion += [("k", "CONSIDERANDO"),
              ("t", "Que, el", f"Que, el {fecha_larga(c.fecha_ingreso)}, " + RECIBIO.format(origen=c.origen)
               + f", el cual contiene {_interpuesto(c)}{_contra(c)};"),
              ("k", "que, de acuerdo con el inciso d)"),
              ("k", "que, de acuerdo con el artículo 147"),
              ("k", "que, de acuerdo al artículo 29"),
              ("t", "que, mediante el escrito", f"que, mediante el escrito del {fecha_larga(c.audiencia_solicitud_fecha)}, "
               f"{sin_prep(sol) if sol else c.audiencia_solicitante} solicitó se convoque a una audiencia de conciliación "
               "para que pueda formular una propuesta conciliatoria."),
              ("k", "SE RESUELVE")]
    n = 0
    if hay_escritos:
        guion.append(("t", "Agregar los referidos", "Agregar los referidos escritos al expediente y ponerlos en conocimiento de la otra parte del procedimiento;"))
        n += 1
    guion.append(("t", "poner en conocimiento", "poner en conocimiento " + traslado(c.no_apelantes) + ";"))
    n += 1
    guion.append(("t", "Informar a las partes",
                  "Informar a las partes que se ha procedido a programar una audiencia de conciliación bajo la modalidad virtual, "
                  "a fin de que tengan la oportunidad de llegar a un acuerdo que solucione la controversia que originó la presente denuncia. "
                  f"En atención a ello, la audiencia de conciliación virtual se llevará a cabo el **__{fecha_larga(c.audiencia_fecha)}, "
                  f"a las {c.audiencia_hora} horas__ __(hora exacta) mediante la aplicación Microsoft Teams.__**"))
    n += 1
    punto = ["i", "ii", "iii", "iv", "v"][n]   # el ítem «requerir a las partes … lo siguiente»
    guion += [("k", "requerir a las partes"), ("k", "Aceptar los"), ("k", "En caso deleguen"),
              ("k", "Descargar de forma"), ("k", "Contar con un equipo"),
              ("t", "informar a las partes que el enlace",
               "informar a las partes que el enlace para poder conectarse a la audiencia de conciliación virtual será enviado en la fecha "
               "programada para el desarrollo de la citada audiencia; y, se verificará el cumplimiento de lo señalado en los incisos a) y b) "
               f"del punto ({punto}) de la presente resolución{{FN:6}}; y,")]
    guion += [("n", t) for t in _notificaciones(c, c.partes, "lista")]
    return guion


# ------------------------------------------------------------------- S5 ---
def s5(c):
    a = c.apelantes[0]
    ps = c.origen.split("/PS")[-1]
    p1 = (f"La Secretaría Técnica de la Comisión informa que el {fecha_larga(c.fecha_ingreso)}, "
          + RECIBIO.format(origen=c.origen).replace("N° ", "Nº ")
          + f", remitido como resultado del recurso de apelación interpuesto por {sin_prep(a)}{_contra(c)}, "
          f"emitida por el Órgano Resolutivo de Procedimientos Sumarísimos N° {ps}, tramitado bajo el expediente {c.origen} "
          f"({c.expediente}/CC1-APELACION).")
    return [("p", p1)] + [("p", t) for t in _notificaciones(c, [a], "s1")]


GUION = {"S1": s1, "S2": s2, "S3": s3, "S4": s4, "S5": s5}
PLANTILLA = {"S1": "S1_simple.docx", "S2": "S2_con_escritos.docx", "S3": "S3_dos_apelaciones.docx",
             "S4": "S4_audiencia.docx", "S5": "S5_inadmisibilidad_liminar.docx"}
