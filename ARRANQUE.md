# Arranque en otra computadora

1. Clonar: `git clone https://github.com/zero-phoenix/elaboracion-de-r1-de-expedientes-de-apelaciones-en-la-cc1-de-indecopi` y entrar a la carpeta.
2. Python 3.10 o superior: `pip install -r requirements.txt`.
3. Capturas (opcional pero recomendado): `pip install playwright` y, si no hay Chrome/Edge instalado, `playwright install chromium`. `scripts/vista.py` busca Chrome/Edge/Chromium solo; también acepta la variable `CHROMIUM_PATH`.
4. Copiar el Excel de control actualizado a `entrada/APES_R1_SEGUROS.xlsx`.
5. Fijar la fecha de emisión en `config/remesa.json` (o usar `--fecha`).
6. Seguir los pasos de `AGENTS.md` §1. Antes de tocar `scripts/` o `plantillas/`, correr `python scripts/simular.py` y comparar con `pruebas/simulacion.md`.
