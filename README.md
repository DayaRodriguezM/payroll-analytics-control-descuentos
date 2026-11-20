# Payroll Analytics | Control de descuentos y saldos de nómina

**Python · pandas · NumPy · Excel · Tableau**

Proyecto aplicado a un proceso real de remuneraciones: automatización de la **identificación de escenarios, clasificación de casos y cálculo de saldos** para apoyar la revisión y el registro de ajustes de nómina.

[**Explorar el dashboard interactivo en Tableau Public**](https://public.tableau.com:9443/views/PayrollAnalyticsControldedescuentosysaldosdenmina/Dashboard1?:language=es-ES&:sid=&:redirect=auth&:display_count=n&:origin=viz_share_link)

<!-- Cuando se cargue la captura pública en images/, quitar los marcadores de comentario de la línea siguiente. -->
<!-- ![Dashboard de Payroll Analytics en Tableau](images/tableau-dashboard.png) -->

## Resultados principales

| Indicador | Resultado |
|---|---:|
| Registros evaluados | **6.604** |
| Casos con acción sugerida o revisión | **272 (4,1 %)** |
| Casos sin acción sugerida | **6.332 (95,9 %)** |
| Identificación y cálculo manual en Excel | **70 minutos** |
| Carga y lectura del archivo en Python | **2 minutos** |
| Procesamiento y generación del reporte en Python | **2 minutos** |
| Tiempo total observado con Python | **4 minutos** |
| Reducción de tiempo en las actividades comparadas | **94,3 % (66 minutos)** |

**Alcance de la medición:** comparación de una revisión operativa. En Python se contabilizaron **2 minutos de carga y lectura del archivo** y **2 minutos de procesamiento y generación del reporte**, para un total observado de **4 minutos** desde la carga hasta disponer del reporte. La reducción no incluye la comprobación manual posterior ni el registro, eliminación o simulación de ajustes en el sistema. Se contrastó manualmente una selección de casos y los cálculos revisados coincidieron con las reglas esperadas.

## El desafío

La revisión de descuentos y saldos de nómina implicaba filtrar casos manualmente, identificar escenarios y calcular importes para su tratamiento. Era necesario aplicar criterios consistentes, distinguir descuentos sujetos al tope del **15 %** de conceptos obligatorios y detectar situaciones que requieren corroboración antes del cierre.

## Solución

1. **Reglas de negocio:** documentación de criterios y escenarios inicialmente revisados en Excel.
2. **Automatización con Python:** normalización de campos, validaciones, clasificación de casos y cálculo de saldos a partir del archivo de nómina.
3. **Salida operativa:** reporte con casos identificados y saldos calculados, listo para revisión y posterior registro.
4. **Seguimiento en Tableau:** indicadores, distribución de acciones y detalle filtrable para explorar excepciones.

Las acciones sugeridas incluyen **traspasar, revisar, eliminar, ajustar y compensar**. Una alerta no implica necesariamente un error: **cada caso debe validarse antes de aplicar cambios**.

## Distribución de acciones

| Acción | Combinaciones caso–acción |
|---|---:|
| Traspasar | 222 |
| Revisar | 110 |
| Eliminar | 96 |
| Ajustar | 75 |
| Compensar | 53 |
| **Total** | **556** |

Los **272 casos** son identificadores únicos con al menos una acción o revisión. Las **556 combinaciones** no representan 556 personas diferentes: un mismo caso puede requerir varias acciones.

## Visualización interactiva

[**Ver Payroll Analytics en Tableau Public**](https://public.tableau.com:9443/views/PayrollAnalyticsControldedescuentosysaldosdenmina/Dashboard1?:language=es-ES&:sid=&:redirect=auth&:display_count=n&:origin=viz_share_link)

La visualización pública se centra en **conteos, categorías y seguimiento de casos**, no en montos individuales de remuneraciones.

## Próxima mejora

Ampliar la validación para una **segunda revisión de nómina**, cuando nuevas ausencias modifican la remuneración imponible (RI) y pueden cambiar los saldos calculados; incluye un escenario adicional asociado a `2H43`, `/S02` y `1D51`. **Esta regla aún no está incorporada a la versión descrita.**

## Privacidad y publicación

El proyecto se originó en una necesidad operativa real. Los materiales públicos deben contener únicamente **datos sintéticos o debidamente anonimizados y autorizados**, sin identificadores personales, información salarial individual ni archivos internos de SAP. Las cifras mostradas son agregadas.

---

**Autora:** Dayana Rodríguez · [GitHub](https://github.com/DayaRodriguezM)
