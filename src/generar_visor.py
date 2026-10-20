"""
Genera docs/index.html: una página para explorar la tabla de actores en el navegador
(buscar, filtrar por rol / tipo de evento / tipo de actor / estado).
Sirve también para GitHub Pages (carpeta docs/).

Uso:
    python src/generar_visor.py
"""

import json
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]

COLUMNAS = ["id_incidente", "fecha", "tipo_evento", "rol", "tipo_actor", "cantidad",
            "nombre_o_alias", "sexo", "edad", "grupo_etario", "nacionalidad", "ocupacion",
            "rubro", "organizacion_vinculada", "antecedentes", "estado_o_situacion",
            "ambito", "titular", "fuente_url"]


def main():
    actores = pd.read_csv(RAIZ / "salida" / "actores.csv", dtype=str).fillna("")
    filas = actores[COLUMNAS].values.tolist()
    datos = json.dumps({"columnas": COLUMNAS, "filas": filas}, ensure_ascii=False, separators=(",", ":"))
    plantilla = (RAIZ / "src" / "visor_plantilla.html").read_text(encoding="utf-8")
    pagina = plantilla.replace("/*__DATOS__*/null", datos.replace("</", "<\\/"))
    salida = RAIZ / "docs" / "index.html"
    salida.parent.mkdir(exist_ok=True)
    salida.write_text('<!doctype html>\n<html lang="es">\n<head>\n<meta charset="utf-8">\n'
                      '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
                      '<style>body{margin:0}[hidden]{display:none!important}</style>\n</head>\n<body>\n'
                      + pagina + "\n</body>\n</html>\n", encoding="utf-8")
    print(f"Visor escrito en {salida} ({len(filas)} filas)")
    return pagina


if __name__ == "__main__":
    main()
