# Guion de presentación - Grupo 7, Proyecto 7

Este archivo es un guion para preparar las diapositivas; aún se debe convertir
en una presentación final e insertar capturas legibles del PBIX.

## Diapositiva 1. Problema y objetivo

Analizamos cómo varían las opiniones y calificaciones de clientes de
restaurantes. El objetivo es procesar reseñas con ETL, estudiar satisfacción por
categoría, ciudad, precio, tiempo y perfil de cliente, y validar diferencias
entre categorías mediante ANOVA.

## Diapositiva 2. Dataset y alcance

Usamos Yelp Open Dataset, alternativa de datos abiertos compartida por el
docente. Elegimos Yelp por sus reseñas textuales, estrellas, fecha, negocios,
categorías, ciudades y datos de actividad de usuarios. El ETL filtra negocios
con categoría `Restaurants`; la muestra del ANOVA se limita a las cinco
categorías con más reseñas y al menos 30 observaciones. Las reseñas de Yelp no
son una muestra aleatoria de todos los clientes de restaurantes.

## Diapositiva 3. ETL y modelo

1. **Extracción:** lectura de archivos JSON Lines de negocios, reseñas y
   usuarios.
2. **Transformación:** selección de restaurantes, validación y limpieza,
   normalización de fechas, cálculo de sentimiento con VADER y preparación de
   categorías, quejas y palabras frecuentes.
3. **Carga:** tablas compactas para Power BI y esquema con hechos de reseñas y
   dimensiones de negocio, usuario y fecha.
4. **Control:** se registran muestras, conteos y resultados estadísticos para
   poder auditar el proceso.

Explicar que VADER clasifica texto en inglés con umbrales de compound: positivo
si `>= 0,05`, negativo si `<= -0,05` y neutral en el intervalo intermedio.

## Diapositiva 4. Diez análisis propios del Proyecto 7

Mostrar cómo el tablero responde a los diez puntos del ejercicio: calificación
por restaurante; sentimiento por categoría; palabras frecuentes; tendencia
mensual; calificación por ciudad; relación precio-satisfacción; Top 10 mejor
calificado; frecuencia de quejas; reseñas positivas y negativas; y perfiles de
cliente. Aclarar que la frecuencia de palabras y el sentimiento se preparan en
el ETL, no se presentan como sustitutos de consultas DAX.

## Diapositiva 5. Diez consultas analíticas resueltas mediante DAX

Presentar cada elemento como **pregunta + expresión DAX + contexto del visual +
respuesta/interpretación**. No basta con enumerar nombres de medidas.

1. ¿Cuántas reseñas hay en el contexto seleccionado? `Total Reseñas` (`COUNT`).
2. ¿Cuál es la calificación media por restaurante, ciudad o rango de precio?
   `Calificación Promedio` (`AVERAGE`); el visual compara precio y satisfacción.
3. ¿Cuántos votos útiles hay? `Votos Útiles Totales` (`SUM`), como indicador de
   interacción.
4. ¿Qué proporción de reseñas es positiva? `% Reseñas Positivas`
   (`CALCULATE`), presentado como KPI y segmentable.
5. ¿Cómo cambia la calificación por año, trimestre, mes y día? `Calificación
   Promedio` en el contexto de los campos de fecha de `dim_date`.
6. ¿Cuál es el promedio acumulado del mes? `Calificación MTD` (`DATESMTD`).
7. ¿Cuál es el promedio acumulado del año y su cambio contra el mismo período
   anterior? `Calificación YTD` (`DATESYTD`) y `Cambio Interanual de
   Calificación` (`SAMEPERIODLASTYEAR`, `VAR`, `RETURN`).
8. ¿Qué negocios tienen más reseñas? `Ranking por Reseñas` (`RANKX`); mide
   volumen, no calidad.
9. ¿Cuántas reseñas incluyen términos de queja? `Reseñas con Quejas`
   (`SEARCH` y filtro de filas).
10. ¿Qué dispersión tienen las calificaciones? `Desviación de Calificaciones`
    (`STDEV.P`).

La pregunta es la consulta; la expresión DAX y el contexto producen la
respuesta. El enunciado enumera medidas y funciones, pero no exige
explícitamente sentencias tabulares `EVALUATE`. Como evidencia adicional, se
preparó `dax_queries.dax` para ejecutar las diez consultas en la Vista de
consultas DAX de Power BI; hay que verificar los resultados antes de exponerlos.

## Diapositiva 6. ANOVA: pregunta, hipótesis y supuestos

- **Variable dependiente:** estrellas (`stars`, escala de 1 a 5).
- **Factor:** categoría de restaurante (`category`).
- **Muestra:** cinco categorías con más observaciones; 1.210.971 reseñas.
- **H₀:** las cinco medias de calificación son iguales.
- **H₁:** al menos una media es diferente.
- **Normalidad:** Shapiro-Wilk fue significativo en los cinco grupos
  (`p < 0,001`); se aplicó a los primeros 5.000 valores de cada grupo.
- **Varianzas:** Levene `W = 531,05`, `p < 0,001`; no se cumple homogeneidad.

Las calificaciones son discretas y el tamaño de muestra es muy grande, por lo
que se informa el resultado de normalidad con cautela. La desigualdad de
varianzas motiva usar Welch como prueba principal.

## Diapositiva 7. ANOVA: resultado e interpretación

Welch ANOVA: `F(4; 594.272,81) = 976,19`, `p < 0,001`. Se rechaza H₀: hay
evidencia estadística de diferencias entre las medias de al menos dos
categorías. La prueba global no identifica qué pares difieren ni demuestra
causalidad.

El gráfico del PBIX muestra los promedios observados: Breakfast & Brunch 3,88;
American (New) 3,84; Bars 3,74; Mexican 3,73; American (Traditional) 3,66.
Son medias de los grupos seleccionados, no una comparación de todos los
restaurantes ni una recomendación causal.

El ANOVA clásico y las comparaciones de Tukey se reportan solo como análisis
exploratorios, porque Levene detecta varianzas desiguales. No presentar Tukey
como prueba confirmatoria.

## Diapositiva 8. Conclusiones, límites y entrega

- La calificación, el sentimiento y las quejas describen aspectos distintos de
  la experiencia; ninguna medida aislada define toda la satisfacción.
- Las reseñas son autoseleccionadas; los resultados describen esta muestra.
- VADER y los términos de queja son reglas de procesamiento de texto en inglés,
  no una evaluación humana de cada comentario.
- Los cuartiles de actividad de usuario describen volumen de reseñas, no tipos
  oficiales de clientes.
- No afirmar causalidad a partir de las diferencias estadísticas.

Antes de exponer, terminar y revisar el informe PDF/Word, el PBIX, la
presentación, el código ETL y la entrega HTML. La segmentación visible por día
y trimestre debe verificarse en el PBIX final antes de afirmar que ambos
niveles se muestran en el tablero.
