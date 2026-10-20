# Mapa de incidentes extorsivos: tabla de víctimas y victimarios

Avance del Grupo 1 (MGP). Pasamos del mapa **«Extorsión y Crímenes Violentos – Perú [2026]»** (uMap / [@e_incidentsperu](https://x.com/e_incidentsperu)), que tiene una fila por *incidente*, a una tabla donde **cada fila es una persona o empresa afectada, o un victimario**, con sus atributos como columnas.

## Qué hay en el repositorio

| Ruta | Contenido |
|---|---|
| `data/extorsion_y_crimenes_violentos_peru_2026.csv` | Datos originales exportados del mapa (1,106 incidentes, ene–oct 2026) |
| `src/extraer_actores.py` | Script que lee cada titular y genera las tablas |
| `src/generar_visor.py` | Genera `docs/index.html`, una página para buscar y filtrar la tabla |
| `notebooks/avance_tabla_actores.ipynb` | Notebook del avance: tabla, diccionario de columnas, resúmenes y gráficos |
| `salida/actores.csv` | **Tabla principal**: una fila por víctima o victimario |
| `salida/incidentes.csv` | Una fila por incidente, con fecha, tipo, coordenadas, fuente y conteo de víctimas |
| `salida/actores.xlsx` | Las dos tablas en Excel |
| `docs/index.html` | Visor web de la tabla (se puede publicar con GitHub Pages) |

## Cómo correrlo

```bash
pip install -r requirements.txt
python src/extraer_actores.py      # genera salida/actores.csv, incidentes.csv y actores.xlsx
python src/generar_visor.py        # genera docs/index.html
```

## Columnas de `actores.csv`

| Columna | Descripción |
|---|---|
| `id_actor`, `id_incidente` | Identificadores; `id_incidente` une con `incidentes.csv` |
| `rol` | Víctima o Victimario |
| `tipo_actor` | Persona, Empresa / negocio, Vehículo, Vivienda, Institución, Sin víctima identificada |
| `cantidad` | 1 = una persona por fila. «2 o más (sin número)» cuando el titular dice «sicarios», «hombres», etc. |
| `nombre_o_alias` | Si el titular lo menciona |
| `sexo`, `edad`, `grupo_etario`, `nacionalidad` | Solo si el titular lo dice |
| `ocupacion`, `rubro` | Víctimas: oficio y sector (chofer → Transporte). Victimarios: tipo (sicario, extorsionador…) |
| `organizacion_vinculada` | Empresa, línea o ruta mencionada |
| `antecedentes` | «Sí» o «Posible» solo si el titular lo indica |
| `estado_o_situacion` | Víctimas: fallecido, herido, herido grave, secuestrado, atacado (daños)… Victimarios: capturado, abatido, sin dato |
| `medio_empleado`, `movil_presunto` | Arma de fuego / explosivo / arma blanca / incendio; extorsión / robo / enfrentamiento |
| `ambito`, `latitud`, `longitud` | Lima Metropolitana y Callao (recuadro aproximado) u otras regiones |
| `titular`, `fuente_url` | Texto original y enlace a la noticia |
| `frase_detectada`, `regla` | De qué parte del titular salió la fila y con qué regla, para revisarla |

## Método y limitaciones

- El CSV solo trae un **titular corto** (~76 caracteres) por incidente, p. ej. *«Sicarios asesinan a chofer de empresa "Etmosa" en Av. Paramonga.»* El script usa reglas (expresiones regulares y diccionarios de ocupaciones, nacionalidades y tipos de objetivo) para sacar de ahí cada actor.
- Nombre, edad, antecedentes y nacionalidad casi siempre quedan en *Sin dato* porque el titular no los trae; para completarlos hay que leer la noticia enlazada.
- Es una primera versión: se revisaron decenas de casos a mano, pero puede haber errores. Las columnas `frase_detectada` y `regla` sirven para encontrarlos y corregir las reglas.
- Dos incidentes tienen fecha de 2025 en un mapa de 2026; probablemente es un error de tipeo en la fuente.

## Siguientes pasos

1. Revisar una muestra y medir el porcentaje de filas correctas.
2. Completar atributos leyendo las noticias enlazadas.
3. Añadir la base 2025.
4. Usar `incidentes.csv` (fecha + coordenadas) para los modelos espacio-temporales: repetición cercana (Johnson, 2008) y procesos autoexcitados de Hawkes (Mohler et al., 2011).
