# Informe técnico - Grupo 7, ejercicio 7

## 1. Objetivo

Medir la satisfacción de clientes de restaurantes mediante calificaciones,
sentimiento textual, categorías, ciudad, precio, tiempo y perfil de usuario.

## 2. Preguntas analíticas

1. Calificación media por restaurante.
2. Sentimiento por categoría.
3. Palabras clave frecuentes (obtenerlas desde `review_text` en Power Query o
   Python, excluyendo stopwords).
4. Tendencia mensual.
5. Calificación por ciudad.
6. Relación precio-satisfacción.
7. Top 10 restaurantes.
8. Frecuencia de quejas específicas.
9. Reseñas positivas frente a negativas.
10. Segmentación por perfil de usuario según el cuartil de `review_count`,
    comparando `average_stars` y `fans`.

## 3. ETL y modelo

El detalle reproducible está en `README.md`. El modelo estrella separa el hecho
de reseña de negocio, usuario y fecha, y conserva la relación de categorías.

## 4. Diez consultas analíticas resueltas mediante DAX

Las diez preguntas de la sección 2 son los análisis propios del Proyecto 7.
Además, el enunciado general pide un mínimo de diez **consultas clave mediante
DAX**. Son dos listas relacionadas, pero no idénticas. Una consulta es la
pregunta analítica que se quiere resolver; la expresión DAX calcula la respuesta
y el contexto del visual (por ejemplo, ciudad, categoría o período) determina
qué filas participan. Una medida puede implementar la respuesta, pero el nombre
de una medida por sí solo no documenta ni demuestra una consulta completa.

Por eso, la tabla siguiente formula cada consulta como pregunta y la vincula
con su expresión DAX, contexto de evaluación y técnica del enunciado. Los diez
ítems de técnicas de la consigna son criterios de cobertura para este conjunto;
no sustituyen la formulación de las preguntas. El enunciado menciona medidas y
funciones DAX, pero no exige explícitamente consultas tabulares `EVALUATE`. Como
evidencia adicional, `dax_queries.dax` contiene diez consultas para ejecutar por
separado en la Vista de consultas DAX de Power BI. El script está preparado,
pero sus resultados todavía deben ejecutarse, revisarse y guardarse como evidencia.

| # | Consulta (pregunta que se resuelve) | Expresión DAX que calcula la respuesta | Contexto y lectura del resultado | Técnica requerida |
|---|---|---|---|---|
| 1 | ¿Cuántas reseñas hay en el contexto seleccionado? | `Total Reseñas` | Tarjeta; segmentar por ciudad, categoría o fecha | `COUNT` |
| 2 | ¿Cuál es la calificación promedio por restaurante, ciudad y rango de precio? | `Calificación Promedio` | Tabla por restaurante; barras por ciudad o `price_range`. Comparar `price_range` con el promedio responde la relación precio-satisfacción | `AVERAGE`; relación entre dos variables en visual y DAX |
| 3 | ¿Cuántos votos útiles recibieron las reseñas? | `Votos Útiles Totales` | Tarjeta o comparación por ciudad/categoría; indicador de interacción, no de satisfacción | `SUM` |
| 4 | ¿Qué porcentaje de reseñas es positivo en cada contexto? | `% Reseñas Positivas` | KPI/dona, segmentable por categoría, ciudad y fecha | `CALCULATE`, contexto de filtro y KPI |
| 5 | ¿Cómo varía la calificación al segmentar por año, trimestre, mes y día? | `Calificación Promedio` | Evaluar la medida con `dim_date[year]`, `quarter`, `month_number` y `date`. Cada nivel agrupa las mismas reseñas de forma temporal distinta | Segmentación temporal |
| 6 | ¿Cuál es la calificación promedio acumulada en el mes hasta la fecha seleccionada? | `Calificación MTD` | Visual temporal y filtros de fecha; la tabla de fechas debe estar relacionada y marcada como tabla de fechas | `DATESMTD` |
| 7 | ¿Cuál es la calificación acumulada del año y cómo cambia frente al mismo período del año anterior? | `Calificación YTD`; `Cambio Interanual de Calificación` | Tabla o visual por año/período; el cambio se expresa en puntos de estrellas, no como porcentaje | `DATESYTD`, `SAMEPERIODLASTYEAR`, `VAR` y `RETURN` |
| 8 | ¿Qué restaurantes tienen mayor volumen de reseñas y cuál es su posición? | `Ranking por Reseñas` | Tabla por restaurante ordenada por rango. Este es un ranking de volumen, no de satisfacción | `RANKX` |
| 9 | ¿Cuántas reseñas contienen al menos uno de los términos de queja definidos? | `Reseñas con Quejas` | Tarjeta o comparación por categoría/ciudad; cada reseña se cuenta una sola vez en esta medida | `SEARCH` y filtro de filas |
| 10 | ¿Qué tanta variabilidad presentan las calificaciones? | `Desviación de Calificaciones` | Tarjeta junto al promedio; una desviación mayor indica más dispersión de estrellas | `STDEV.P` |

### Expresiones DAX y explicación

Las siguientes expresiones implementan las respuestas de la matriz. En las
consultas segmentadas, el mismo cálculo se vuelve a evaluar por cada fila,
columna, filtro o eje del visual.

**Consulta 1 — total de reseñas**

```DAX
Total Reseñas =
COUNT(fact_review[stars])
```

Cuenta las calificaciones no vacías; la tarjeta o los filtros determinan el
período, ciudad o categoría que se está consultando.

**Consulta 2 — calificación media**

```DAX
Calificación Promedio =
AVERAGE(fact_review[stars])
```

Calcula el promedio de estrellas en el contexto. Al agrupar por restaurante o
ciudad responde esos análisis; al poner `price_range` en el eje y la medida en
valores muestra la relación entre rango de precio y satisfacción.

**Consulta 3 — votos útiles**

```DAX
Votos Útiles Totales =
SUM(fact_review[useful])
```

Suma los votos útiles. Es una medida de interacción con las reseñas, no una
medida directa de satisfacción.

**Consulta 4 — porcentaje de reseñas positivas**

```DAX
% Reseñas Positivas =
DIVIDE(
    CALCULATE(
        [Total Reseñas],
        fact_review[sentiment_label] = "Positive"
    ),
    [Total Reseñas],
    0
)
```

`CALCULATE` aplica el filtro de sentimiento positivo dentro del contexto
seleccionado; `DIVIDE` obtiene la proporción respecto al total y evita error si
el denominador es cero. Se presenta como KPI.

**Consulta 5 — segmentación temporal**

La expresión es `[Calificación Promedio]` de la consulta 2, evaluada con los
campos `dim_date[year]`, `dim_date[quarter]`, `dim_date[month_number]` o
`dim_date[date]`. El campo de fecha cambia el contexto temporal; no requiere
inventar una medida distinta para cada nivel.

**Consulta 6 — calificación MTD**

```DAX
Calificación MTD =
CALCULATE(
    [Calificación Promedio],
    DATESMTD(dim_date[date])
)
```

Calcula el promedio desde el inicio del mes hasta la fecha del contexto
seleccionado. En un visual agrupado por mes completo, el valor puede coincidir
con el promedio mensual.

**Consulta 7 — calificación YTD y cambio interanual**

```DAX
Calificación YTD =
CALCULATE(
    [Calificación Promedio],
    DATESYTD(dim_date[date])
)
```

```DAX
Cambio Interanual de Calificación =
VAR CalificaciónActual = [Calificación Promedio]
VAR CalificaciónAnterior =
    CALCULATE(
        [Calificación Promedio],
        SAMEPERIODLASTYEAR(dim_date[date])
    )
RETURN
    IF(
        ISBLANK(CalificaciónActual) || ISBLANK(CalificaciónAnterior),
        BLANK(),
        CalificaciónActual - CalificaciónAnterior
    )
```

YTD calcula el promedio acumulado del año. El cambio interanual compara el
período actual con el mismo período del año anterior y se expresa en puntos de
calificación; `VAR` guarda ambos valores y `RETURN` devuelve la diferencia.
Devuelve vacío si falta alguno de los períodos comparados.

**Consulta 8 — ranking por volumen de reseñas**

```DAX
Ranking por Reseñas =
RANKX(
    ALL(dim_business),
    [Total Reseñas],
    [Total Reseñas],
    DESC,
    DENSE
)
```

Compara el total de reseñas de cada negocio contra los demás negocios. No es el
ranking de restaurantes mejor calificados; ese Top 10 se calcula por promedio
de estrellas y usa volumen e ID como desempates.

**Consulta 9 — reseñas con términos de queja**

```DAX
Reseñas con Quejas =
COUNTROWS(
    FILTER(
        fact_review,
        SEARCH("bad", LOWER(fact_review[complaint_terms]), 1, 0) > 0
            || SEARCH("slow", LOWER(fact_review[complaint_terms]), 1, 0) > 0
            || SEARCH("cold", LOWER(fact_review[complaint_terms]), 1, 0) > 0
            || SEARCH("rude", LOWER(fact_review[complaint_terms]), 1, 0) > 0
            || SEARCH("dirty", LOWER(fact_review[complaint_terms]), 1, 0) > 0
    )
)
```

Busca los términos definidos en cada reseña y cuenta cada fila una sola vez,
aunque contenga varios términos.

**Consulta 10 — dispersión de calificaciones**

```DAX
Desviación de Calificaciones =
STDEV.P(fact_review[stars])
```

Mide la dispersión poblacional de las estrellas en el contexto actual; se
interpreta junto con el promedio, no como una calificación de satisfacción por
sí sola.

`dim_date` aporta año, trimestre, mes, fecha/día y día de semana para la
segmentación. En el PBIX ya se confirmó la relación activa entre fecha y reseña
y que `dim_date` está marcada como tabla de fechas. La página temporal muestra
MTD en el gráfico mensual y YTD/cambio interanual en la tabla. La demostración
visible de los niveles día y trimestre aún debe comprobarse en el PBIX final;
no se debe asegurar que ambos aparecen en un visual hasta hacer esa
verificación.

`Votos Útiles Totales` es un indicador complementario del uso de las reseñas,
no una medida directa de satisfacción. `Ranking por Reseñas` clasifica por
volumen y no significa que el restaurante sea el mejor calificado. El
requisito del proyecto 7 de Top 10 mejor calificados se atiende por separado:
el visual selecciona por calificación promedio y desempata por total de
reseñas e ID mediante la medida auxiliar `Es Top 10 Restaurante`.

`% Reseñas Positivas` se utiliza como KPI de satisfacción. El enunciado no
define límites para clasificar el promedio como "Excelente", "Aceptable" o
"Por mejorar"; por eso esos umbrales no se presentan como criterios del
docente. El PBIX contiene `Es Top 10 Restaurante` como medida auxiliar para
seleccionar diez negocios por promedio de calificación, con total de reseñas e
ID como desempates. Es apoyo al requisito propio del Proyecto 7 de Top 10
mejor calificados; no es una consulta analítica adicional. `Ranking por
Reseñas`, en cambio, ordena por volumen y no debe interpretarse como ranking de
satisfacción.

**Estado de auditoría DAX:** las diez expresiones de referencia fueron
inspeccionadas en Power BI y las fórmulas revisadas fueron aceptadas por el
modelo. Se verificaron la relación activa con `dim_date` y el marcado de esa
dimensión como tabla de fechas; también se colocaron las medidas MTD, YTD e
interanual en la página temporal. Las diez consultas tabulares están en
`dax_queries.dax`, pero aún no se han ejecutado en la Vista de consultas DAX.
Antes de entregar, falta ejecutar y revisar esos resultados, comprobar la
segmentación visible por día y trimestre y capturar/explicar las respuestas en
el informe final. La frecuencia de palabras se calcula
en el ETL, no mediante DAX: cumple el análisis específico de palabras del
Proyecto 7, pero no se cuenta como una de las diez consultas DAX generales.

El requisito 3 (palabras clave) se resuelve en el ETL sobre `review_text` y se
visualiza en Power BI con una tabla o nube de palabras. El requisito 6
(precio-satisfacción) se resuelve colocando `price_range` en el eje y
`Calificación Promedio` en valores. Así no se agregan medidas innecesariamente
complicadas para análisis que Power BI ya resuelve con el contexto del visual.

El requisito 8 se visualiza con `complaint_term_counts`, que cuenta por
separado cada término (`bad`, `slow`, `cold`, `rude`, `dirty`). Una reseña
puede contener más de un término y, por tanto, contribuir al conteo de varios;
estos conteos no representan reseñas únicas sumables entre sí. No usar la
columna `complaint_terms` como categoría, porque agrupa términos de una misma
reseña en valores como `bad, cold`.

El requisito 10 se segmenta por cuartiles de `review_count`, calculados sobre
los perfiles incluidos en el dataset. En la muestra actual, los cortes son 3,
7 y 20 reseñas: Q1 (0-3), Q2 (4-7), Q3 (8-20) y Q4 (más de 20). Se comparan
las estrellas promedio y los fans por grupo; los cuartiles describen actividad,
no etiquetas oficiales del ejercicio.

## 5. ANOVA

La variable dependiente es `stars` (calificación de 1 a 5); el factor es
`category`. H0: las medias de calificación son iguales. H1: al menos una media
es diferente. Se verifican Shapiro-Wilk por grupo y Levene para homogeneidad de
varianzas en `anova_assumptions_and_result.csv`. Como Levene detecta varianzas
desiguales, se usa Welch ANOVA como prueba principal; el ANOVA clásico se
conserva como referencia. Si `p < 0.05`, se rechaza H0.

### Resultados observados

El análisis compara cinco categorías y 1.210.971 reseñas. Levene dio
`W = 531,05`, `p < 0,001`, por lo que no se cumple la igualdad de varianzas.
El ANOVA de una vía dio `F(4; 1.210.966) = 966,33`, `p < 0,001`: hay evidencia
de que no todas las medias son iguales. Los valores `0.0` de p guardados en los
CSV son un límite de precisión numérica, no una probabilidad exactamente cero;
por eso se reportan como `p < 0,001`.

Welch ANOVA, que no requiere igualdad de varianzas, dio
`F(4; 594.272,81) = 976,19`, `p < 0,001`. Por tanto, también con esta prueba
robusta se rechaza H0: hay evidencia de diferencias entre las calificaciones
medias de al menos dos categorías. Este resultado no identifica por sí solo qué
pares difieren.

El visual de Power BI presenta las medias observadas: Breakfast & Brunch 3,88;
American (New) 3,84; Bars 3,74; Mexican 3,73; y American (Traditional) 3,66.
Son medias de las cinco categorías seleccionadas por volumen de observaciones,
no una evaluación de todos los restaurantes ni evidencia causal.

Shapiro-Wilk también resultó significativo en las cinco categorías
(`p < 0,001`). El ETL aplica esta prueba a los primeros 5.000 valores de cada
grupo; la columna `n` del CSV indica el tamaño total del grupo, no el tamaño
usado por Shapiro-Wilk.

En las diez comparaciones de Tukey, nueve resultaron significativas. La única
excepción fue `Bars` frente a `Mexican` (`p ajustado = 0,2225`). Como las
varianzas son desiguales, Tukey no se usa para conclusiones confirmatorias; esos
pares se conservan únicamente como evidencia exploratoria. No implican
causalidad ni, por sí solos, recomendaciones sobre la calidad de un restaurante.

## 6. Conclusiones

Welch ANOVA permite rechazar la hipótesis de igualdad de medias entre las cinco
categorías analizadas sin asumir varianzas iguales. Shapiro-Wilk también resulta
significativo en los grupos; las calificaciones son discretas y el tamaño de
muestra es muy grande, así que las pruebas de normalidad deben interpretarse con
cautela. Se reporta evidencia de asociación entre categoría y calificación, no
causalidad ni una evaluación definitiva de la calidad. Las conclusiones
descriptivas del dashboard deben limitarse a los filtros, categorías y períodos
que muestra cada visual; antes de entregar se debe completar la comprobación
final del PBIX, en particular la segmentación temporal por día y trimestre.
