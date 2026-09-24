"""Visor propio de .docx (no usa LibreOffice).

1. MIDE el formato real de cada párrafo leyendo el XML con herencia de estilos
   (docDefaults → estilo base → estilo del párrafo → formato directo):
   fuente, tamaño, negrita, cursiva, subrayado, alineación, sangrías, interlineado,
   espaciado antes/después, numeración (i)(ii)…, notas al pie, márgenes de página.
2. DIBUJA la página A4 en HTML con esas medidas y toma una CAPTURA COMPLETA con
   Chromium (Playwright). Marca en gris el salto de cada página.
3. COMPARA contra otro .docx (plantilla u original): captura lado a lado y una
   tabla de diferencias de formato párrafo por párrafo.

  python scripts/vista.py "salida/R1 0663-2026-CC1-APE.docx" --contra plantillas/S1_simple.docx
  → salida/_vista/<nombre>.png  (captura completa, lado a lado si hay --contra)
    salida/_vista/<nombre>.html (la misma vista, navegable)
    salida/_vista/<nombre>_FORMATO.md (ficha de formato y diferencias)

Si «Arial Narrow» no está instalada en la máquina, la captura usa Liberation Sans
comprimida al 82 % (ancho equivalente) y lo avisa; la FICHA siempre informa la fuente
exacta del documento, que es lo que se revisa.
"""
import argparse
import glob
import html
import os
import re
import shutil
import subprocess
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
q = lambda t: f"{{{W}}}{t}"  # noqa: E731
TWIP_CM = 2.54 / 1440


def _val(el, attr="val"):
    return None if el is None else el.get(q(attr))


def _on(el):
    return el is not None and _val(el) not in ("0", "false", "off")


# ------------------------------------------------------------------ lectura ---
class Docx:
    def __init__(self, ruta):
        self.ruta = Path(ruta)
        z = zipfile.ZipFile(ruta)
        rd = lambda n: etree.fromstring(z.read(n)) if n in z.namelist() else None  # noqa: E731
        self.doc = rd("word/document.xml")
        self.styles = rd("word/styles.xml")
        self.numbering = rd("word/numbering.xml")
        self.footnotes = rd("word/footnotes.xml")
        self.estilos = {s.get(q("styleId")): s for s in (self.styles.findall("w:style", NS) if self.styles is not None else [])}
        dd = self.styles.find("w:docDefaults", NS) if self.styles is not None else None
        self.rpr_def = dd.find("w:rPrDefault/w:rPr", NS) if dd is not None else None
        self.ppr_def = dd.find("w:pPrDefault/w:pPr", NS) if dd is not None else None
        self.normal = next((sid for sid, s in self.estilos.items()
                            if s.get(q("type")) == "paragraph" and s.get(q("default")) == "1"), None)
        self.contadores = {}
        sp = self.doc.find(".//w:body/w:sectPr", NS)
        pg = sp.find("w:pgSz", NS)
        mg = sp.find("w:pgMar", NS)
        self.pagina = {k: int(mg.get(q(k), 0)) for k in ("top", "bottom", "left", "right")}
        self.pagina.update(w=int(pg.get(q("w"))), h=int(pg.get(q("h"))))

    # --- herencia de estilos
    def _cadena(self, sid):
        out, vistos = [], set()
        while sid and sid in self.estilos and sid not in vistos:
            vistos.add(sid)
            out.append(self.estilos[sid])
            b = self.estilos[sid].find("w:basedOn", NS)
            sid = _val(b)
        return list(reversed(out))

    def props_run(self, r, p):
        capas = [self.rpr_def]
        ps = _val(p.find("w:pPr/w:pStyle", NS)) or self.normal
        capas += [s.find("w:rPr", NS) for s in self._cadena(ps)]
        rs = _val(r.find("w:rPr/w:rStyle", NS))
        capas += [s.find("w:rPr", NS) for s in self._cadena(rs)]
        capas.append(r.find("w:rPr", NS))
        out = {"fuente": None, "tam": None, "b": False, "i": False, "u": False, "caps": False, "sup": False}
        for c in capas:
            if c is None:
                continue
            f = c.find("w:rFonts", NS)
            if f is not None and (f.get(q("ascii")) or f.get(q("hAnsi"))):
                out["fuente"] = f.get(q("ascii")) or f.get(q("hAnsi"))
            s = c.find("w:sz", NS)
            if s is not None:
                out["tam"] = int(_val(s)) / 2
            for k, t in (("b", "b"), ("i", "i"), ("caps", "caps")):
                e = c.find(f"w:{t}", NS)
                if e is not None:
                    out[k] = _on(e)
            u = c.find("w:u", NS)
            if u is not None:
                out["u"] = _val(u) not in (None, "none")
            v = c.find("w:vertAlign", NS)
            if v is not None:
                out["sup"] = _val(v) == "superscript"
        return out

    def props_parrafo(self, p):
        capas = [self.ppr_def]
        ps = _val(p.find("w:pPr/w:pStyle", NS)) or self.normal
        capas += [s.find("w:pPr", NS) for s in self._cadena(ps)]
        num = self._num(p)
        if num is not None:
            capas.append(num["lvl"].find("w:pPr", NS))
        capas.append(p.find("w:pPr", NS))
        out = {"jc": "left", "izq": 0, "der": 0, "primera": 0, "antes": 0, "despues": 0, "linea": 240, "regla": "auto",
               "estilo": ps}
        for c in capas:
            if c is None:
                continue
            j = c.find("w:jc", NS)
            if j is not None:
                out["jc"] = _val(j)
            ind = c.find("w:ind", NS)
            if ind is not None:
                for a, k in (("left", "izq"), ("start", "izq"), ("right", "der"), ("end", "der")):
                    if ind.get(q(a)) is not None:
                        out[k] = int(ind.get(q(a)))
                if ind.get(q("firstLine")) is not None:
                    out["primera"] = int(ind.get(q("firstLine")))
                if ind.get(q("hanging")) is not None:
                    out["primera"] = -int(ind.get(q("hanging")))
            sp = c.find("w:spacing", NS)
            if sp is not None:
                for a, k in (("before", "antes"), ("after", "despues"), ("line", "linea")):
                    if sp.get(q(a)) is not None:
                        out[k] = int(sp.get(q(a)))
                if sp.get(q("lineRule")):
                    out["regla"] = sp.get(q("lineRule"))
        out["num"] = self._etiqueta(num) if num is not None else ""
        return out

    # --- numeración
    def _num(self, p):
        np_ = p.find("w:pPr/w:numPr", NS)
        if np_ is None and (ps := _val(p.find("w:pPr/w:pStyle", NS))):
            for s in self._cadena(ps):
                np_ = s.find("w:pPr/w:numPr", NS) if np_ is None else np_
        if np_ is None or self.numbering is None:
            return None
        nid, ilvl = _val(np_.find("w:numId", NS)), int(_val(np_.find("w:ilvl", NS)) or 0)
        if nid in (None, "0"):
            return None
        n = self.numbering.find(f"w:num[@w:numId='{nid}']", NS)
        if n is None:
            return None
        aid = _val(n.find("w:abstractNumId", NS))
        an = self.numbering.find(f"w:abstractNum[@w:abstractNumId='{aid}']", NS)
        lvl = an.find(f"w:lvl[@w:ilvl='{ilvl}']", NS) if an is not None else None
        if lvl is None:
            return None
        return {"id": nid, "ilvl": ilvl, "lvl": lvl}

    def _etiqueta(self, num):
        k = (num["id"], num["ilvl"])
        ini = int(_val(num["lvl"].find("w:start", NS)) or 1)
        self.contadores[k] = self.contadores.get(k, ini - 1) + 1
        for (i, l) in list(self.contadores):
            if i == num["id"] and l > num["ilvl"]:
                del self.contadores[(i, l)]
        n = self.contadores[k]
        fmt = _val(num["lvl"].find("w:numFmt", NS)) or "decimal"
        txt = _val(num["lvl"].find("w:lvlText", NS)) or "%1."
        v = {"lowerRoman": _romano(n).lower(), "upperRoman": _romano(n), "lowerLetter": chr(96 + n),
             "upperLetter": chr(64 + n), "bullet": "•"}.get(fmt, str(n))
        return txt.replace(f"%{num['ilvl'] + 1}", v) if fmt != "bullet" else "•"

    def nota(self, fid):
        if self.footnotes is None:
            return ""
        fn = self.footnotes.find(f"w:footnote[@w:id='{fid}']", NS)
        return "\n".join("".join(t.text or "" for t in p.iter(q("t"))) for p in fn.findall("w:p", NS)) if fn is not None else ""


def _romano(n):
    out = ""
    for v, s in ((10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
        while n >= v:
            out += s
            n -= v
    return out


# ----------------------------------------------------------------- análisis ---
def analizar(ruta):
    """Lista de bloques: {'tipo': 'p'|'tabla', 'props', 'runs': [(texto, props_run)], 'texto'}."""
    d = Docx(ruta)
    bloques, notas = [], []

    def parrafo(p):
        pp = d.props_parrafo(p)
        runs = []
        for r in p.iter(q("r")):
            if r.getparent().tag == q("r"):
                continue
            pr = d.props_run(r, p)
            for ch in r:
                if ch.tag == q("t"):
                    runs.append((ch.text or "", pr))
                elif ch.tag == q("tab"):
                    runs.append(("\t", pr))
                elif ch.tag in (q("br"), q("cr")):
                    runs.append(("\n", pr))
                elif ch.tag == q("footnoteReference"):
                    notas.append(d.nota(ch.get(q("id"))))
                    runs.append((str(len(notas)), dict(pr, sup=True, nota=True)))
        return {"tipo": "p", "props": pp, "runs": runs, "texto": "".join(t for t, _ in runs)}

    for el in d.doc.find("w:body", NS):
        if el.tag == q("p"):
            bloques.append(parrafo(el))
        elif el.tag == q("tbl"):
            filas = []
            anchos = [int(g.get(q("w"), 0)) for g in el.findall("w:tblGrid/w:gridCol", NS)]
            for tr in el.findall("w:tr", NS):
                filas.append([[parrafo(p) for p in tc.findall("w:p", NS)] for tc in tr.findall("w:tc", NS)])
            bloques.append({"tipo": "tabla", "filas": filas, "anchos": anchos,
                            "texto": " | ".join(" ".join(b["texto"] for b in c) for f in filas for c in f)})
    return d, bloques, notas


def describir(pr):
    s = f"{pr['fuente'] or '?'} {pr['tam'] or '?'}"
    for k, t in (("b", "negrita"), ("i", "cursiva"), ("u", "subrayado"), ("caps", "versalitas"), ("sup", "superíndice")):
        if pr.get(k):
            s += f" {t}"
    return s


def resumen_parrafo(b):
    pp = b["props"]
    tramos, prev = [], None
    for t, pr in b["runs"]:
        if not t.strip():
            continue
        dsc = describir(pr)
        if tramos and dsc == prev:
            tramos[-1][0] += t
        else:
            tramos.append([t, dsc])
            prev = dsc
    inter = (f"{pp['linea'] / 240:.2f} líneas" if pp["regla"] == "auto"
             else f"{pp['linea'] / 20:.1f} pt ({pp['regla']})")
    return {
        "texto": b["texto"].strip(),
        "alineacion": {"both": "justificado", "center": "centrado", "left": "izquierda", "start": "izquierda",
                       "right": "derecha", "end": "derecha"}.get(pp["jc"], pp["jc"]),
        "sangria": f"izq {pp['izq'] * TWIP_CM:.2f} cm · 1.ª línea {pp['primera'] * TWIP_CM:+.2f} cm",
        "interlineado": inter,
        "espaciado": f"{pp['antes'] / 20:.0f}/{pp['despues'] / 20:.0f} pt",
        "numeracion": pp["num"],
        "tramos": [(t.strip()[:40], dsc) for t, dsc in tramos],
    }


# -------------------------------------------------------------------- HTML ---
def _fuente_disponible(nombre):
    try:
        out = subprocess.run(["fc-list", ":", "family"], capture_output=True, text=True, timeout=10).stdout
        return nombre.lower() in out.lower()
    except Exception:
        return False


def _css_run(pr, sustituta):
    f = pr["fuente"] or "Arial Narrow"
    fam = f"'{f}'"
    if f == "Arial Narrow" and sustituta:
        fam = "'Liberation Sans', Arial"
    s = f"font-family:{fam};font-size:{pr['tam'] or 11}pt;"
    if pr["b"]:
        s += "font-weight:bold;"
    if pr["i"]:
        s += "font-style:italic;"
    if pr["u"]:
        s += "text-decoration:underline;"
    if pr["sup"]:
        s += "vertical-align:super;font-size:70%;"
    return s


def _html_parrafo(b, ancho_cm, sustituta, escala):
    pp = b["props"]
    inter = pp["linea"] / 240 * 1.15 if pp["regla"] == "auto" else None
    lh = f"line-height:{inter:.3f};" if inter else f"line-height:{pp['linea'] / 20}pt;"
    jc = {"both": "justify", "center": "center", "right": "right", "end": "right"}.get(pp["jc"], "left")
    izq, pri = pp["izq"] * TWIP_CM, pp["primera"] * TWIP_CM
    cont = "".join(
        ("<span style='display:inline-block;width:1.25cm'></span>" if t == "\t" else
         "<br>" if t == "\n" else f"<span style=\"{_css_run(pr, sustituta)}\">{html.escape(t)}</span>")
        for t, pr in b["runs"])
    if pp["num"]:
        cont = f"<span style='display:inline-block;min-width:{-pri if pri < 0 else 0.6}cm'>{html.escape(pp['num'])}</span>" + cont
    if not b["texto"].strip() and not pp["num"]:
        tam = next((pr["tam"] for _, pr in b["runs"] if pr["tam"]), 11)
        cont = f"<span style='font-size:{tam}pt'>&nbsp;</span>"
    w = ancho_cm - izq - pp["der"] * TWIP_CM
    esc = f"width:{w / escala:.3f}cm;transform:scaleX({escala});transform-origin:0 0;" if escala != 1 else f"width:{w:.3f}cm;"
    return (f"<div style='margin-left:{izq:.3f}cm;margin-top:{pp['antes'] / 20}pt;margin-bottom:{pp['despues'] / 20}pt;'>"
            f"<div style='{esc}text-align:{jc};text-indent:{pri:.3f}cm;{lh}'>{cont}</div></div>")


def pagina_html(ruta, titulo=None):
    d, bloques, notas = analizar(ruta)
    pg = d.pagina
    ancho = (pg["w"] - pg["left"] - pg["right"]) * TWIP_CM
    fuentes = {pr["fuente"] for b in bloques if b["tipo"] == "p" for _, pr in b["runs"]}
    sustituta = "Arial Narrow" in fuentes and not _fuente_disponible("Arial Narrow")
    escala = 0.82 if sustituta else 1
    partes = []
    for b in bloques:
        if b["tipo"] == "p":
            partes.append(_html_parrafo(b, ancho, sustituta, escala))
        else:
            tot = sum(b["anchos"]) or 1
            filas = "".join("<tr>" + "".join(
                f"<td style='width:{b['anchos'][i] / tot * 100 if i < len(b['anchos']) else 10:.1f}%;vertical-align:top;padding:0'>"
                + "".join(_html_parrafo(p, (b['anchos'][i] if i < len(b['anchos']) else 0) * TWIP_CM, sustituta, escala) for p in celda)
                + "</td>" for i, celda in enumerate(f)) + "</tr>" for f in b["filas"])
            partes.append(f"<table style='border-collapse:collapse;width:{ancho:.2f}cm;table-layout:fixed'>{filas}</table>")
    if notas:
        partes.append("<div style='border-top:1px solid #000;width:5cm;margin-top:14pt'></div>")
        for i, n in enumerate(notas, 1):
            partes.append(f"<div style='font-family:{'Liberation Sans' if sustituta else 'Arial Narrow'};font-size:8pt;"
                          f"text-align:justify;margin:2pt 0'><sup>{i}</sup> {html.escape(n).replace(chr(10), '<br>')}</div>")
    alto = pg["h"] * TWIP_CM
    aviso = ("<div class='aviso'>Captura con fuente sustituta (Arial Narrow no instalada): Liberation Sans al 82 %. "
             "La ficha de formato informa la fuente real.</div>" if sustituta else "")
    return (f"<div class='hoja'><div class='tit'>{html.escape(titulo or d.ruta.name)}</div>{aviso}"
            f"<div class='pag' style='width:{pg['w'] * TWIP_CM:.2f}cm;padding:{pg['top'] * TWIP_CM:.2f}cm "
            f"{pg['right'] * TWIP_CM:.2f}cm {pg['bottom'] * TWIP_CM:.2f}cm {pg['left'] * TWIP_CM:.2f}cm;"
            f"background:repeating-linear-gradient(#fff 0,#fff calc({alto:.2f}cm - 2px),#bbb calc({alto:.2f}cm - 2px),#bbb {alto:.2f}cm)'>"
            + "".join(partes) + "</div></div>")


CSS = """body{background:#e9e9e9;margin:0;padding:16px;display:flex;gap:24px;align-items:flex-start}
.hoja{flex:none}.tit{font:bold 13px sans-serif;margin:0 0 6px}.aviso{font:11px sans-serif;color:#8a4b00;margin:0 0 6px;max-width:21cm}
.pag{box-sizing:border-box;box-shadow:0 1px 4px rgba(0,0,0,.35);min-height:29.7cm;color:#000}"""


def navegador():
    cand = [os.environ.get("CHROMIUM_PATH")] + sorted(glob.glob("/opt/pw-browsers/chromium-*/chrome-linux/chrome"), reverse=True)
    cand += [shutil.which(n) for n in ("chromium", "chromium-browser", "google-chrome", "chrome")]
    cand += [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
             r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
             "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"]
    return next((c for c in cand if c and Path(c).exists()), None)


def capturar(html_txt, png):
    """Captura de página completa con Chromium vía Playwright. Devuelve None si no hay navegador."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return "Falta el paquete playwright (pip install playwright); se generó solo el HTML."
    exe = navegador()
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe) if exe else p.chromium.launch()
        pg = b.new_page(viewport={"width": 1800, "height": 1200}, device_scale_factor=1.5)
        pg.set_content(html_txt)
        pg.screenshot(path=str(png), full_page=True)
        b.close()
    return None


def _planos(bl):
    """Párrafos con texto, incluidos los de las celdas de tablas (encabezado)."""
    out = []
    for b in bl:
        if b["tipo"] == "p":
            out.append(b)
        else:
            out += [p for f in b["filas"] for c in f for p in c]
    return [resumen_parrafo(b) for b in out if b["texto"].strip()]


def ficha(ruta, contra=None):
    """Ficha de formato en Markdown; con --contra, diferencias párrafo a párrafo."""
    d, bl, notas = analizar(ruta)
    pg = d.pagina
    cm = lambda v: f"{v * TWIP_CM:.2f}"  # noqa: E731
    L = [f"# Formato de {Path(ruta).name}", "",
         f"Página {cm(pg['w'])} × {cm(pg['h'])} cm · márgenes sup {cm(pg['top'])} / inf {cm(pg['bottom'])} / "
         f"izq {cm(pg['left'])} / der {cm(pg['right'])} cm · notas al pie: {len(notas)}", "",
         "| # | Texto | Alineación | Sangría | Interlineado | Espaciado | Núm. | Tipografía por tramos |",
         "|---|---|---|---|---|---|---|---|"]
    pars = _planos(bl)
    for i, r in enumerate(pars, 1):
        tr = "<br>".join(f"«{html.escape(t)}» → {dsc}" for t, dsc in r["tramos"])
        L.append(f"| {i} | {html.escape(r['texto'][:60])} | {r['alineacion']} | {r['sangria']} | {r['interlineado']} | "
                 f"{r['espaciado']} | {r['numeracion']} | {tr} |")
    if contra:
        d2, bl2, n2 = analizar(contra)
        p2 = _planos(bl2)
        L += ["", f"## Diferencias de formato frente a {Path(contra).name}", ""]
        m2 = d2.pagina
        for k in ("top", "bottom", "left", "right", "w", "h"):
            if abs(pg[k] - m2[k]) > 20:
                L.append(f"- Página/márgen «{k}»: {cm(pg[k])} cm frente a {cm(m2[k])} cm.")
        # se emparejan párrafos por su comienzo (las 3 primeras palabras) o por orden
        usados = set()
        for r in pars:
            clave = " ".join(r["texto"].split()[:3])
            j = next((j for j, x in enumerate(p2) if j not in usados and " ".join(x["texto"].split()[:3]) == clave), None)
            if j is None:
                continue
            usados.add(j)
            x = p2[j]
            for k in ("alineacion", "sangria", "interlineado", "espaciado", "numeracion"):
                if r[k] != x[k]:
                    L.append(f"- «{r['texto'][:50]}»: {k} {r[k]} ≠ {x[k]}")
            t1 = sorted({dsc for _, dsc in r["tramos"]})
            t2 = sorted({dsc for _, dsc in x["tramos"]})
            if t1 != t2:
                L.append(f"- «{r['texto'][:50]}»: tipografía {t1} ≠ {t2}")
        if len(L) and L[-1].startswith("## Diferencias"):
            L.append("Sin diferencias de formato.")
        elif not any(l.startswith("- ") for l in L[L.index(f'## Diferencias de formato frente a {Path(contra).name}'):]):
            L.append("Sin diferencias de formato.")
    return "\n".join(L) + "\n"


def main():
    a = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    a.add_argument("docx", nargs="+")
    a.add_argument("--contra", help="plantilla u original para comparar lado a lado")
    a.add_argument("--salida", default=None, help="carpeta (por defecto <carpeta del docx>/_vista)")
    A = a.parse_args()
    for r in A.docx:
        r = Path(r)
        out = Path(A.salida) if A.salida else r.parent / "_vista"
        out.mkdir(parents=True, exist_ok=True)
        cuerpo = pagina_html(r, "GENERADA · " + r.name if A.contra else r.name)
        if A.contra:
            cuerpo += pagina_html(A.contra, "REFERENCIA · " + Path(A.contra).name)
        doc = f"<!doctype html><meta charset='utf-8'><title>{html.escape(r.stem)}</title><style>{CSS}</style><body>{cuerpo}</body>"
        (out / f"{r.stem}.html").write_text(doc, encoding="utf-8")
        (out / f"{r.stem}_FORMATO.md").write_text(ficha(r, A.contra), encoding="utf-8")
        aviso = capturar(doc, out / f"{r.stem}.png")
        print(f"{r.name}: {out / (r.stem + '.png') if not aviso else aviso}")
        print(f"   ficha: {out / (r.stem + '_FORMATO.md')}")


if __name__ == "__main__":
    main()
