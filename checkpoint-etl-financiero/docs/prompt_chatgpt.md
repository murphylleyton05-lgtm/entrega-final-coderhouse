# Paso 6 — Generación y depuración de código VBA con IA

Este documento deja constancia del prompt usado para generar la macro de formato,
del código que devolvió la IA y de los errores que hubo que depurar antes de
integrarlo al módulo `Modulo_Formato_IA`.

---

## 1. Prompt optimizado

> Actuá como desarrollador VBA senior para Excel 365 en español.
>
> Escribí una macro llamada `Aplicar_Formato_Ejecutivo` que reciba un libro con:
> - una hoja `Reporte_Ejecutivo` con tarjetas de KPI en las filas 11 a 13 y una
>   tabla dinámica llamada `TD_Margen`,
> - una hoja `Evaluación_Financiera` con un modelo de inversión,
> - una hoja `Datos_Limpios` con la tabla de salida de Power Query.
>
> Requisitos:
> 1. Tipografía corporativa Segoe UI en todas las hojas.
> 2. Paleta: azul `#0B2A4A` para encabezados y celeste `#008AB3` para acentos.
> 3. Bordes ejecutivos: contorno medio azul e interior fino gris.
> 4. Formato moneda `$#,##0` en importes y `0.0%` en porcentajes.
> 5. Todo referenciado por **CodeName** de hoja, nunca por `ActiveSheet`.
> 6. Manejo de errores con `On Error` para que no se corte si falta un objeto.
>
> Devolvé sólo el módulo, con `Option Explicit` y comentarios en español.

Por qué está escrito así:

| Técnica | Para qué sirve |
|---|---|
| Rol explícito ("desarrollador VBA senior") | fija el nivel de la respuesta |
| Contexto del libro (hojas, filas, nombres) | evita que la IA invente rangos |
| Requisitos numerados | cada uno es verificable uno por uno |
| Restricciones técnicas (CodeName, `On Error`) | ataca de entrada los errores típicos |
| Formato de salida ("sólo el módulo") | respuesta pegable sin recortes |

---

## 2. Depuración del código devuelto

La primera respuesta **no compilaba**. Estos son los cuatro errores encontrados y
cómo se corrigieron:

### a) Uso de `ActiveSheet` y `.Select`

```vba
' Devuelto por la IA
Sheets("Reporte_Ejecutivo").Select
Range("B11:K13").Select
Selection.Interior.Color = RGB(244, 244, 247)
```

Problema: rompe si se renombra la pestaña y es lento.
Corrección: referencias directas por CodeName, sin seleccionar nada.

```vba
' Corregido
With Hoja_Reporte.Range(Hoja_Reporte.Cells(11, col), Hoja_Reporte.Cells(13, col + 1))
    .Interior.Color = GRIS_SUAVE
End With
```

### b) `RGB()` dentro de una constante

```vba
' Devuelto por la IA -> error de compilación
Private Const AZUL_CORP As Long = RGB(11, 42, 74)
```

Problema: `Const` exige una expresión constante en tiempo de compilación y `RGB()`
es una función. Excel devuelve *"Constant expression required"*.
Corrección: se precalculó el valor (`11 + 42*256 + 74*65536`).

```vba
' Corregido
Private Const AZUL_CORP As Long = 4860427      ' RGB(11, 42, 74)
```

### c) Error 1004 cuando la tabla dinámica todavía no existe

```vba
' Devuelto por la IA
Hoja_Reporte.PivotTables("TD_Margen").TableStyle2 = "PivotStyleMedium9"
```

Problema: al abrir el libro por primera vez la dinámica aún no fue creada y la
macro se cortaba con *"Run-time error 1004: no se encontró el elemento"*.
Corrección: se obtiene el objeto con `On Error Resume Next` y se sale si es `Nothing`.

```vba
' Corregido
On Error Resume Next
Set td = Hoja_Reporte.PivotTables(NOMBRE_TD)
On Error GoTo 0
If td Is Nothing Then Exit Sub
```

### d) Formato aplicado a columnas que pueden no existir

Problema: la IA aplicaba `NumberFormat` directo sobre `ListColumns("Margen")`; si la
consulta cambia de columnas, falla.
Corrección: se encapsuló en `Formato_Columna`, que ignora la columna ausente.

```vba
' Corregido
Private Sub Formato_Columna(ByVal tabla As ListObject, ByVal columna As String, _
                            ByVal formato As String)
    On Error Resume Next
    tabla.ListColumns(columna).DataBodyRange.NumberFormat = formato
    On Error GoTo 0
End Sub
```

---

## 3. Resultado

El módulo final está en [`vba/Modulo_Formato_IA.bas`](../vba/Modulo_Formato_IA.bas)
e integrado al proyecto VBA del `.xlsm`. Se ejecuta desde el **botón 2 — APLICAR
FORMATO** y también, en su versión silenciosa (`Aplicar_Formato_Silencioso`, sin
cuadros de diálogo), al abrir el libro.

El prompt y este registro de depuración también quedaron copiados dentro del
libro, en la hoja `Config`, para que se puedan revisar sin salir de Excel.
