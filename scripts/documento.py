"""Construye el .docx de la R1 sobre la plantilla del supuesto y lo verifica.

La plantilla aporta TODO el formato (página, cabecera, tabla de encabezado,
estilos, numeración (i)(ii)…, notas al pie y bloque de firma). Aquí solo se
cambia texto: encabezado, fecha, cuerpo (según el guion) y firma de config/.
"""
import copy
import json
import re

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from modelo import RAIZ, clave
from textos import fecha_larga

W_P, W_R, W_T, W_TBL = qn("w:p"), qn("w:r"), qn("w:t"), qn("w:tbl")


def firma():
    return json.loads((RAIZ / "config" / "firma.json").read_text(encoding="utf-8"))


# ------------------------------------------------------------- utilidades ---
def texto(el):
    return "".join(t.text or "" for t in el.iter(W_T))


def numerado(p):
    return p.find(".//" + qn("w:numPr")) is not None


def _rpr_base(p):
    for r in p.iter(W_R):
        t = r.find(W_T)
        if t is not None and (t.text or "").strip():
            rpr = r.find(qn("w:rPr"))
            return copy.deepcopy(rpr) if rpr is not None else OxmlElement("w:rPr")
    rpr = p.find(qn("w:pPr") + "/" + qn("w:rPr"))
    return copy.deepcopy(rpr) if rpr is not None else OxmlElement("w:rPr")


def _sin(rpr, *tags):
    for t in tags:
        for e in rpr.findall(qn(t)):
            rpr.remove(e)
    return rpr


def poner_texto(p, txt, notas):
    """Sustituye el contenido de un párrafo conservando su pPr.
    Marcas: **negrita**, __subrayado__, {FN:id} llamada a la nota id (se clona su run original)."""
    base = _sin(_rpr_base(p), "w:highlight", "w:b", "w:bCs", "w:u", "w:vertAlign", "w:rStyle")
    for ch in list(p):
        if ch.tag != qn("w:pPr"):
            p.remove(ch)
    negrita = subrayado = False
    for tok in re.split(r"(\*\*|__|\{FN:\d+\})", txt):
        if tok == "**":
            negrita = not negrita
            continue
        if tok == "__":
            subrayado = not subrayado
            continue
        m = re.fullmatch(r"\{FN:(\d+)\}", tok)
        if m:
            p.append(copy.deepcopy(notas[m.group(1)]))
            continue
        tok = re.sub(r" {2,}", " ", tok)
        if not tok:
            continue
        r = OxmlElement("w:r")
        rpr = copy.deepcopy(base)
        if negrita:
            b = OxmlElement("w:b")
            rpr.insert(0, b)
        if subrayado:
            u = OxmlElement("w:u")
            u.set(qn("w:val"), "single")
            rpr.append(u)
        r.append(rpr)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = tok
        r.append(t)
        p.append(r)


def _en_negrita(p):
    """¿La plantilla tiene en negrita el primer texto de este párrafo?"""
    for r in p.iter(W_R):
        t = r.find(W_T)
        if t is not None and (t.text or "").strip():
            b = r.find(qn("w:rPr") + "/" + qn("w:b"))
            return b is not None and b.get(qn("w:val")) not in ("0", "false")
    return False


def quitar_resaltado(root):
    for e in list(root.iter(qn("w:highlight"))):
        e.getparent().remove(e)


def notas_al_pie(d):
    """{id: run con la llamada} y el id de la nota del art. 20.4 del TUO de la LPAG."""
    runs = {}
    for r in d.element.body.iter(W_R):
        ref = r.find(qn("w:footnoteReference"))
        if ref is not None:
            runs[ref.get(qn("w:id"))] = copy.deepcopy(r)
    id204 = None
    textos_notas = {}
    for rel in d.part.rels.values():
        if rel.reltype.endswith("/footnotes"):
            root = rel.target_part.element if hasattr(rel.target_part, "element") else None
            xml = rel.target_part.blob.decode("utf-8") if root is None else None
            if root is not None:
                for fn in root.findall(qn("w:footnote")):
                    textos_notas[fn.get(qn("w:id"))] = texto(fn)
            else:
                for m in re.finditer(r'<w:footnote [^>]*w:id="(-?\d+)"[^>]*>(.*?)</w:footnote>', xml, re.S):
                    textos_notas[m.group(1)] = re.sub(r"<[^>]+>", "", m.group(2))
    for i, t in textos_notas.items():
        if "20.4" in t and i in runs:
            id204 = i
    return runs, id204, textos_notas


# ---------------------------------------------------------- construcción ---
def construir(caso, guion, plantilla, salida):
    d = docx.Document(str(plantilla))
    body = d.element.body
    quitar_resaltado(body)
    notas, id204, _ = notas_al_pie(d)
    f = firma()

    # 1. Página: márgenes del modelo de referencia (S1).
    ref = docx.Document(str(RAIZ / "plantillas" / "S1_simple.docx")).sections[0]
    for s in d.sections:
        for a in ("top_margin", "bottom_margin", "left_margin", "right_margin", "page_width", "page_height",
                  "header_distance", "footer_distance"):
            setattr(s, a, getattr(ref, a))

    # 2. Tabla de encabezado.
    tbl = body.find(W_TBL)

    def celda(p, txt):
        poner_texto(p, f"**{txt}**" if _en_negrita(p) else txt, notas)

    for tr in tbl.iter(qn("w:tr")):
        celdas = tr.findall(qn("w:tc"))
        if len(celdas) < 3:
            continue
        etq = clave(texto(celdas[0]))
        pe, pv = celdas[0].find(W_P), celdas[2].find(W_P)
        if etq.startswith("INGRESO EN COMISION"):
            celda(pv, caso.expediente_cc1)
        elif etq.startswith("EXPEDIENTE DE ORIGEN"):
            celda(pv, caso.origen)
        elif etq.startswith("DENUNCIANTE"):
            celda(pe, "DENUNCIANTE" if len(caso.denunciantes) == 1 else "DENUNCIANTES")
            celda(pv, " / ".join(p.nombre_cabecera for p in caso.denunciantes))
        elif etq.startswith("DENUNCIADO"):
            celda(pe, "DENUNCIADO(S)")
            celda(pv, " / ".join(p.nombre_cabecera for p in caso.denunciados))
        elif etq.startswith("RESOLUCION"):
            celda(pv, "1")

    # 3. Cuerpo: entre «Lima, …» y «Firmado digitalmente por».
    hijos = list(body)
    i_lima = next(i for i, e in enumerate(hijos) if e.tag == W_P and texto(e).strip().startswith("Lima,"))
    i_firma = next(i for i, e in enumerate(hijos) if e.tag == W_P and texto(e).strip().startswith("Firmado digitalmente"))
    poner_texto(hijos[i_lima], f"Lima, {fecha_larga(caso.fecha_emision)}", notas)
    region = hijos[i_lima + 1:i_firma]
    vacios_finales = 0
    for e in reversed(region):
        if e.tag == W_P and not texto(e).strip():
            vacios_finales += 1
        else:
            break
    con_texto = [e for e in region if e.tag == W_P and texto(e).strip()]
    vacio = next(e for e in region if e.tag == W_P and not texto(e).strip())
    proto_p = next(e for e in con_texto if not numerado(e))
    proto_n = next((e for e in con_texto if numerado(e) and texto(e).strip().startswith("poner en conocimiento")),
                   next((e for e in con_texto if numerado(e)), None))
    usados = set()

    def buscar(prefijo):
        for e in con_texto:
            if id(e) not in usados and texto(e).strip().startswith(prefijo):
                usados.add(id(e))
                return e
        raise ValueError(f"La plantilla no tiene un párrafo que empiece por «{prefijo}».")

    def fn(txt):
        return txt.replace("{FN}", f"{{FN:{id204}}}" if id204 else "")

    nuevos = [copy.deepcopy(vacio)]
    for item in guion:
        tipo = item[0]
        if tipo == "k":
            el = copy.deepcopy(buscar(item[1]))
        elif tipo == "t":
            el = copy.deepcopy(buscar(item[1]))
            poner_texto(el, fn(item[2]), notas)
        else:
            el = copy.deepcopy(proto_n if tipo == "n" else proto_p)
            if tipo == "p":
                np_ = el.find(qn("w:pPr") + "/" + qn("w:numPr"))
                if np_ is not None:
                    np_.getparent().remove(np_)
            poner_texto(el, fn(item[1]), notas)
        nuevos += [el, copy.deepcopy(vacio)]
    nuevos += [copy.deepcopy(vacio) for _ in range(max(vacios_finales - 1, 0))]
    for e in region:
        body.remove(e)
    firma_p = hijos[i_firma]
    for e in nuevos:
        firma_p.addprevious(e)

    # 4. Firma (config/firma.json) e iniciales.
    sig = [e for e in list(body)[list(body).index(firma_p) + 1:] if e.tag == W_P and texto(e).strip()]
    for el, txt in zip(sig[:3], (f["firmante"], f["cargo"], f["comision"])):
        poner_texto(el, f"**{txt}**" if _en_negrita(el) else txt, notas)
    for e in sig[3:]:
        if re.fullmatch(r"[A-Za-z]{2,4}/[A-Za-z]{2,4}", texto(e).strip()):
            poner_texto(e, f["iniciales"], notas)
    d.save(str(salida))
    return salida


# ----------------------------------------------------------- verificador ---
def verificar(ruta, caso=None):
    """Lista de errores (vacía = APTO)."""
    err = []
    d = docx.Document(str(ruta))
    body = d.element.body
    f = firma()
    todo = "\n".join(texto(p) for p in body.iter(W_P))
    cuerpo = [texto(p) for p in body.iter(W_P)]
    if re.search(r"\{|\}|\*\*|__|XXX", todo):
        err.append("Quedan marcas sin resolver ({…}, ** o XXX).")
    if list(body.iter(qn("w:highlight"))):
        err.append("Hay resaltados (highlight).")
    if f["firmante"] not in todo or f["cargo"] not in todo:
        err.append(f"La firma no es la de config/firma.json ({f['firmante']}, {f['cargo']}).")
    for otro in ("LUISA ANALI", "Ejecutivo", "Secretaria Técnica", "(e)"):
        if otro in todo:
            err.append(f"Aparece «{otro}»: firma o cargo que no corresponde.")
    if "004-2019" in todo:
        err.append("Cita el Decreto Supremo 004-2019-JUS: debe ser 006-2026-JUS.")
    for i, t in enumerate(cuerpo):
        s = t.strip()
        if s and "  " in s:
            err.append(f"Doble espacio en «{s[:60]}…».")
        if re.search(r"\.\.(?!\.)|;;|; y;|,\s*\.", s):
            err.append(f"Puntuación duplicada en «{s[:60]}…».")
    runs, id204, textos_notas = notas_al_pie(d)
    for i in runs:
        if i not in textos_notas:
            err.append(f"Llamada a la nota {i} sin nota.")
    if "bandeja de correo" in todo and "(i)" not in todo:
        pass
    fuentes = {r.find(qn("w:rPr") + "/" + qn("w:rFonts")).get(qn("w:ascii"))
               for r in body.iter(W_R)
               if r.find(qn("w:rPr") + "/" + qn("w:rFonts")) is not None
               and r.find(qn("w:rPr") + "/" + qn("w:rFonts")).get(qn("w:ascii"))}
    if fuentes - {"Arial Narrow"}:
        err.append(f"Fuentes distintas de Arial Narrow: {sorted(fuentes - {'Arial Narrow'})}.")
    ref = docx.Document(str(RAIZ / "plantillas" / "S1_simple.docx")).sections[0]
    s = d.sections[0]
    for a in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        if abs(getattr(s, a) - getattr(ref, a)) > 1000:
            err.append(f"Margen {a} distinto del modelo.")
    if caso is not None:
        if caso.expediente_cc1 not in todo:
            err.append("El encabezado no trae el número de ingreso en comisión.")
        notif = [t for t in cuerpo if re.search(r"(requerir|Requerir|exhortar|Exhortar) ", t)]
        destinatarios = caso.apelantes if caso.supuesto == "S5" else caso.partes
        for p in destinatarios:
            if not any(p.nombre in t for t in notif):
                err.append(f"«{p.nombre}» no tiene párrafo de notificación.")
        if caso.supuesto == "S5":
            for p in caso.denunciados:
                if p.nombre in "\n".join(cuerpo[1:]):
                    err.append(f"S5: el denunciado «{p.nombre}» no debe ser notificado ni trasladado.")
        if caso.supuesto != "S1" and any(p.via == "correo" for p in destinatarios) and id204 is None:
            err.append("Falta la nota al pie del art. 20.4 del TUO de la LPAG.")
    return err
