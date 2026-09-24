"""Fojas de los escritos a trasladar (apelación y demás escritos). NO van en la R1:
se informan al usuario con la redacción de las cédulas del corpus.

  python scripts/fojas.py 0259-2026          → entrada/expedientes/0259-2026/_FOJAS.md

Coloca los documentos del expediente en entrada/expedientes/<NNNN-AAAA>/ (PDF, DOCX o imágenes).

Qué se cuenta (deducido de la práctica de las cédulas: se cuentan las hojas del escrito
tal como se trasladan, con sus anexos):
  · CUENTA  toda página del escrito y de sus anexos (incluidas las de anexos escaneados).
  · NO CUENTA  la página en blanco (sin texto y casi sin tinta) y la constancia/cargo
    automático de Mesa de Partes (Virtual) de Indecopi, que no es parte del escrito.
Cada página excluida se lista con su motivo para que el usuario la confirme.

Forma del informe (la de las cédulas recientes):
  «Copia del escrito de apelación presentado por la parte denunciante el 12/06/2026 (8 fojas).»
  «Copia del escrito presentado por La Positiva Seguros y Reaseguros S.A. el 25/06/2026 (1 foja).»
"""
import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from modelo import RAIZ, fechas_en  # noqa: E402

CARGO = re.compile(r"mesa de partes|cargo de (presentaci|recepci)|constancia de (presentaci|recepci)|"
                   r"hoja de (tr[aá]mite|ruta)|n[uú]mero de (tr[aá]mite|registro)|fecha y hora de (presentaci|recepci)", re.I)
APELACION = re.compile(r"recurso de apelaci|interpongo.{0,40}apelaci|interpone.{0,40}apelaci|apelaci[oó]n contra", re.I)


def paginas_pdf(ruta):
    import pymupdf
    d = pymupdf.open(str(ruta))
    out = []
    for i, p in enumerate(d, 1):
        txt = p.get_text().strip()
        pix = p.get_pixmap(dpi=20, colorspace=pymupdf.csGRAY)
        tinta = sum(1 for b in pix.samples if b < 200) / max(len(pix.samples), 1)
        out.append({"n": i, "texto": txt, "tinta": tinta})
    return out


def paginas_docx(ruta):
    import docx
    d = docx.Document(str(ruta))
    txt = "\n".join(p.text for p in d.paragraphs)
    n = int(d.core_properties.revision and 0) or None
    app = None
    try:
        import zipfile
        x = zipfile.ZipFile(ruta).read("docProps/app.xml").decode()
        app = int(re.search(r"<Pages>(\d+)</Pages>", x).group(1))
    except Exception:
        pass
    return [{"n": i + 1, "texto": txt if i == 0 else "", "tinta": 1} for i in range(app or 1)], app is None


def analizar(ruta):
    ext = ruta.suffix.lower()
    aviso = ""
    if ext == ".pdf":
        pags = paginas_pdf(ruta)
    elif ext == ".docx":
        pags, sin_dato = paginas_docx(ruta)
        if sin_dato:
            aviso = "Word sin conteo de páginas guardado: conviene enviarlo en PDF."
    elif ext in (".jpg", ".jpeg", ".png", ".tif", ".tiff"):
        pags = [{"n": 1, "texto": "", "tinta": 1}]
    else:
        return None
    cuenta, excluidas = [], []
    for p in pags:
        if p["tinta"] < 0.003 and len(p["texto"]) < 15:
            excluidas.append((p["n"], "página en blanco"))
        elif CARGO.search(p["texto"][:1500]) and len(p["texto"]) < 1800 and p["n"] in (1, len(pags)):
            excluidas.append((p["n"], "constancia/cargo de Mesa de Partes"))
        else:
            cuenta.append(p["n"])
    todo = "\n".join(p["texto"] for p in pags[:3])
    tipo = "escrito de apelación" if APELACION.search(todo) else "escrito"
    fechas = fechas_en(todo)
    return {"archivo": ruta.name, "paginas": len(pags), "fojas": len(cuenta), "excluidas": excluidas,
            "tipo": tipo, "fecha_texto": fechas[0].strftime("%d/%m/%Y") if fechas else "", "aviso": aviso,
            "sin_texto": all(len(p["texto"]) < 15 for p in pags)}


def frase(r, quien="[parte]", fecha=None):
    n = r["fojas"]
    return (f"Copia del {r['tipo']} presentado por {quien} el {fecha or r['fecha_texto'] or '[fecha]'} "
            f"({n} {'foja' if n == 1 else 'fojas'}).")


def main():
    a = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    a.add_argument("expediente")
    A = a.parse_args()
    m = re.search(r"0*(\d+)-(\d{4})", A.expediente)
    exp = f"{int(m.group(1)):04d}-{m.group(2)}"
    carpeta = RAIZ / "entrada" / "expedientes" / exp
    if not carpeta.exists():
        sys.exit(f"No existe {carpeta}: coloca ahí los documentos del expediente.")
    L = [f"# Fojas — {exp}/CC1-APELACIÓN", "",
         "No van en la R1. Se cuentan las hojas del escrito y sus anexos; no se cuentan páginas en blanco "
         "ni la constancia/cargo automático de Mesa de Partes. Revisa las exclusiones.", "",
         "| Documento | Páginas del archivo | Fojas | Excluidas | Tipo detectado | Fecha en el texto |", "|---|---|---|---|---|---|"]
    frases = []
    for f in sorted(carpeta.iterdir()):
        if f.name.startswith("_") or f.is_dir():
            continue
        r = analizar(f)
        if r is None:
            continue
        exc = "; ".join(f"p. {n}: {m}" for n, m in r["excluidas"]) or "—"
        L.append(f"| {r['archivo']} | {r['paginas']} | **{r['fojas']}** | {exc} | {r['tipo']} | {r['fecha_texto'] or '—'} |")
        if r["aviso"] or r["sin_texto"]:
            L.append(f"| ↳ aviso | | | {r['aviso'] or 'PDF escaneado sin texto: la detección de blancos es por tinta; tipo y fecha a confirmar.'} | | |")
        frases.append(frase(r))
    L += ["", "## Para tu control (forma de las cédulas)", ""] + [f"- {x}" for x in frases]
    L += ["", "La fecha de un escrito de parte es la de su presentación (firma digital / cargo), no la que el escrito dice; "
          "«[parte]» y la fecha se completan al confirmar quién lo presentó."]
    out = carpeta / "_FOJAS.md"
    out.write_text("\n".join(L) + "\n", encoding="utf-8")
    print("\n".join(L))


if __name__ == "__main__":
    main()
