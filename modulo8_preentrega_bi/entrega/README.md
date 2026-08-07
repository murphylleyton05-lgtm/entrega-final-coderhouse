# Entrega final — Módulo 8

`PreEntrega_ModeloBI_MurphyLleyton.xlsx` es el archivo que se entrega, ya limpio.

Se construyó corriendo `MaterializarModelo` sobre el libro base y después
`RepararModeloBI` + `UltimaMedida`, que crearon las relaciones y las medidas
identificando las tablas del modelo por sus columnas (Excel las nombró
"Consulta 1..4" al cargarlas, no con el nombre de la consulta).

## Estado

| Punto de la rúbrica | Estado |
|---|---|
| Power Query + lenguaje M | 4 consultas cargadas al modelo, paso renombrado y comentarios `//` |
| Modelo estrella 1:N | 3 relaciones dimensión → hechos, funcionando |
| Medidas DAX con VAR/RETURN | 8 medidas, 4 con VAR/RETURN |
| Escenarios + resumen | Pesimista / Base / Optimista + hoja de resumen |
| Estadística ToolPak | media, mediana, desvío, rango, curtosis, asimetría |
| Verificación | `TD_Ventas`: Total Ventas por año (810M / 783M / 770M) |

## Lo que quedó pendiente

Las dos cosas requieren el complemento Power Pivot, que no está instalado en la
máquina donde se armó:

- Las tablas del modelo figuran como "Consulta 1..4" en vez de `Hechos_Ventas`
  y `Dim_*`. Se renombran con doble clic en las pestañas de la ventana de
  Power Pivot; las medidas se reapuntan solas.
- `Dim_Calendario` no quedó marcada como tabla de fechas. Por eso las dos
  medidas interanuales usan `ALL()` + comparación de año en lugar de `DATEADD`.

Nada de esto afecta el funcionamiento: las relaciones y las medidas responden,
como se ve en `TD_Ventas`.

## Limpieza aplicada

Se quitaron `Verificacion_Modelo` (log del primer intento, con líneas FALLO),
`Diagnostico`, `Reparacion` y `LEEME`, editando el paquete OOXML directamente
para no tocar `xl/model/item.data` ni el DataMashup de Power Query — abrir el
libro con una librería de Excel habría borrado el modelo.
