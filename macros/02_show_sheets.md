# Mostrar hojas ocultas

1. Crear un módulo en el editor VBA

2. Escribir

```
Sub show_sheets()
    Dim wks As Worksheet
    For Each wks in ActiveWorkbook.Worksheets
        wks.Visible = xlSheetVisible
    Next wks
End Sub
```

3. Insertar un botón y asignarle la macro.