# Proyecto final - Grupo 7, ejercicio 7

## Tema y dataset elegido

**Análisis de opiniones de clientes en restaurantes** usando el **Yelp Open
Dataset**. Se eligió Yelp en lugar de Amazon porque contiene, en el mismo
ecosistema, negocios, categorías, ciudades, coordenadas, rango de precios,
reseñas con texto y fecha, y perfiles de usuarios. Eso permite resolver los
10 requerimientos del ejercicio 7 sin inventar campos.

Descarga oficial: <https://business.yelp.com/data/resources/open-dataset/>.
El dataset se entrega como JSON Lines y no se incluye en el repositorio. Antes
de usar o publicar datos derivados, se deben respetar los términos de uso de
Yelp y las autorizaciones aplicables.

## Contenido del repositorio

- `etl_pipeline.py` y `prepare_powerbi.py`: código ETL y preparación para
  Power BI.
- `dax_measures.dax`: expresiones DAX utilizadas para responder las consultas
  analíticas documentadas en `report.md`.
- `dax_queries.dax`: diez consultas tabulares `EVALUATE`, ejecutadas por
  separado en la Vista de consultas DAX de Power BI.
- `dax_guide.md` y `deliverables\guia_dax.pdf`: guía ordenada para estudiar y
  explicar las medidas y las diez consultas, una por una.
- `report.md` y `presentation.md`: fuente del informe técnico y guion oral.
- Informe Power BI [`Grupo7_Ejercicio7_Yelp.pbix`](https://github.com/HectorARiascosI/grupo-7-proyecto-7-analisis-opiniones-restaurantes/releases/latest/download/Grupo7_Ejercicio7_Yelp.pbix), disponible como archivo descargable en la sección Releases.
- `deliverables\report.docx` y `deliverables\report.pdf`: informe técnico final.
- `deliverables\guia_dax.pdf`: guía de estudio de DAX.
- `deliverables\presentation.pptx`: presentación para la exposición.
- `deliverables\dashboard.html`: dashboard estático local con agregados; no reemplaza el PBIX ni es una publicación en Power BI Service.
- `deliverables\report.html`: versión HTML del informe.
- `build_final_deliverables.py`: genera esos cinco archivos desde el informe y las tablas compactas locales.
- `data/processed/anova_assumptions_and_result.csv` y
  `data/processed/anova_tukey_posthoc.csv`: resultados estadísticos pequeños
  para auditar el ANOVA.

Los archivos JSON de Yelp y las tablas procesadas voluminosas se excluyen para
evitar redistribuir datasets y archivos de varios GB. Para regenerar esas
tablas, descarga los datos desde la fuente oficial y sigue los pasos siguientes.
El PBIX se distribuye como archivo de Release para evitar límites y dependencias
de Git LFS.

## Ejecución reproducible

1. Instalar Python 3.11+ y dependencias:

   ```powershell
   python -m pip install -r requirements.txt
   ```

2. Descargar el ZIP oficial en `data\source\` y extraer en `data\raw\` estos
   archivos:
   `yelp_academic_dataset_business.json`,
   `yelp_academic_dataset_review.json` y
   `yelp_academic_dataset_user.json`.

   Crea localmente esta estructura; las carpetas con datos no se versionan:

   ```text
   data\
   ├── source\       # ZIP original y archivos auxiliares/documentación
   ├── raw\          # únicamente los 3 JSON requeridos por el ETL
   └── processed\    # tablas generadas localmente
   ```

3. Ejecutar el ETL completo:

   ```powershell
   python etl_pipeline.py --raw-dir data\raw --output-dir data\processed
   ```

   Para una prueba rápida antes de procesar todo:

   ```powershell
   python etl_pipeline.py --raw-dir data\raw --output-dir data\processed --max-reviews 100000
   ```

4. Generar los documentos finales y el dashboard HTML:

   ```powershell
   python prepare_powerbi.py --source-dir data\processed --output-dir data\processed\powerbi
   python build_final_deliverables.py
   ```

   El generador procesa las reseñas por bloques, no incorpora textos de reseñas
   ni perfiles individuales al HTML, y crea `report.docx`, `report.pdf`,
   `report.html`, `guia_dax.pdf`, `presentation.pptx` y `dashboard.html` dentro
   de `deliverables\`.

   La comparación por ciudad usa normalización de espacios y mayúsculas para
   que la agregación local y Power BI produzcan los mismos grupos. La consulta 1
   verificada muestra 848 ciudades y 687.553 reseñas para Philadelphia; el total
   es 4.724.471. Guarda el PBIX después del refresco y conserva la normalización
   descrita en `report.md` al recalcular el dashboard.

## Salidas

- `dim_business.csv`: restaurante, ciudad, estado, ubicación, precio y estado.
- `dim_user.csv`: perfil agregado del usuario.
- `dim_date.csv`: calendario para inteligencia temporal.
- `bridge_business_category.csv`: relación muchos-a-muchos de categorías.
- `fact_review.csv`: estrellas, texto, votos, sentimiento y fecha.
- `anova_assumptions_and_result.csv`: Shapiro-Wilk, Levene, ANOVA clásico y
  Welch ANOVA.
- `anova_tukey_posthoc.csv`: comparaciones post-hoc de Tukey.

## Diseño ETL auditado

1. **Extract**: lectura línea a línea de JSONL, con validación de JSON y de
   archivos requeridos.
2. **Transform**: selección explícita de negocios cuya categoría contiene
   `Restaurants`, normalización de fechas, extracción de `price_range`,
   eliminación de duplicados dimensionales y cálculo de sentimiento VADER.
   El campo `category` de `fact_review` conserva la primera etiqueta de Yelp
   no genérica de cada negocio seleccionado; no necesariamente representa su
   tipo de cocina.
3. **Load**: carga en un esquema estrella en CSV para Power BI.
4. **Control**: el pipeline falla de forma explícita ante archivos ausentes,
   JSON inválido o ausencia de reseñas; no reemplaza datos faltantes con
   valores inventados.

## ANOVA

La pregunta es si la calificación media difiere entre las cinco etiquetas Yelp
no genéricas con más observaciones asociadas a negocios seleccionados como
restaurantes (mínimo 30 por grupo). No necesariamente son tipos de cocina.
H0: todas las medias son iguales. H1: al menos una media es diferente. Se
reportan Shapiro-Wilk por grupo, Levene para homogeneidad, ANOVA de una vía y
Welch ANOVA. Como Levene
detecta varianzas desiguales, Welch es la prueba principal; el ANOVA clásico y
Tukey se conservan como comparación exploratoria. Con `p < 0.05` se rechaza H0.
La muestra, grupos, tamaños y p-valores quedan guardados en los CSV para
auditoría.

## Power BI

Preparar primero las tablas compactas del modelo:

```powershell
python prepare_powerbi.py --source-dir data\processed --output-dir data\processed\powerbi
```

La carpeta `data\processed\powerbi\` contiene las reseñas completas sin el
texto largo, los campos necesarios para el análisis, una tabla agregada de
palabras frecuentes, otra con la frecuencia individual de cada término de
queja, y `activity_segment` en `dim_user.csv`. Este último agrupa usuarios por
cuartiles de `review_count`, con cortes calculados desde los datos. El texto
original permanece en `data\processed\fact_review.csv`; no se carga a Power BI
para reducir memoria. El conteo de palabras excluye términos generales como
`place`, `time`, `one` y `all`, pero conserva negaciones como `not` y `don't`,
que aportan significado al análisis de opiniones.

Importar los CSV compactos desde esa carpeta. Crear estas relaciones:

- `dim_business[business_id]` a `fact_review[business_id]`.
- `dim_user[user_id]` a `fact_review[user_id]`.
- `dim_date[date]` a `fact_review[review_day]`.

El ETL genera `dim_date` con todos los días entre la primera y la última
reseña, incluidos los días sin actividad, para que Power BI pueda marcarla
como tabla de fechas y calcular inteligencia temporal.

Usar `category` de `fact_review` para analizar las etiquetas Yelp asignadas a
los negocios seleccionados como restaurantes. No interpretar automáticamente
cada etiqueta como una cocina o tipo de restaurante: pueden aparecer etiquetas
secundarias como `Pet Training` o `Roofing`. La tabla
`review_keyword_counts` no necesita relación: se visualiza directamente como
palabra y frecuencia. La tabla `complaint_term_counts` tampoco necesita
relación: contiene una fila por término y su número de reseñas; una reseña
puede contar en varios términos. No uses `complaint_terms` como eje para este
gráfico, porque ese campo puede contener combinaciones como `bad, cold`.
Marcar `dim_date` como tabla de fechas. DAX es el lenguaje;
`dax_measures.dax` contiene las expresiones DAX de referencia para resolver
diez consultas analíticas. Cada consulta debe presentarse como una pregunta
concreta, acompañada por la expresión que calcula la respuesta, el contexto
del visual/filtros y una interpretación. Las medidas son una forma de
implementar esos cálculos, pero una medida aislada no reemplaza la pregunta ni
la respuesta. La consigna enumera técnicas de medidas y funciones DAX; no
exige explícitamente consultas tabulares `EVALUATE`. Como evidencia adicional,
`dax_queries.dax` contiene diez bloques que se ejecutaron por separado en la
Vista de consultas DAX de Power BI; los resultados y sus limitaciones se
documentan en `report.md`. Conserva las capturas de ejecución como evidencia
de entrega. Crear las medidas una por una con
**Modelado > Nueva medida** y documentar las diez preguntas en el informe.
Para demostrar la segmentación temporal, usar los
campos de año, trimestre, mes y fecha/día de `dim_date`; verificar que el PBIX
final muestre esos niveles, no solo que existan en la tabla. La definición de referencia de `Es Top 10 Restaurante` en `dax_measures.dax`
selecciona los diez negocios mejor calificados entre los que tienen al menos
100 reseñas, con desempates; no cuenta entre las diez consultas DAX principales.
La medida quedó aplicada en el PBIX y la tabla se verificó con ese umbral.

Páginas recomendadas: resumen ejecutivo, restaurantes/ciudades, sentimiento y
texto, tendencia temporal, y ANOVA. El archivo `.pbix` debe guardarse desde
Power BI Desktop; el ETL y las medidas aquí incluidos permiten reproducirlo.
