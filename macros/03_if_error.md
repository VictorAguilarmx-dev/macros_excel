# Si.Error
Esta macro sirve para que cuando selecciones un conjunto de celdas, no importando el valor las envuelve en un si.error()

1. Abrir VBA

2. Escribir:

```
Public Sub Iferror()

Dim row As Long
Dim col As Long
Dim formulaString As String
Dim ReadArr as Variant

If Selection.Cells.Count > 1 Then
    ReadArr = Selection.FormulaR1C1
    For row = LBound(ReadArr,1) To UBound(ReadArr,1)
        For col = LBound(ReadArr, 2) To UBound(ReadArr, 2)
            If Left(ReadArr(row,col),1) = "=" Then
                If LCase(Left(ReadArr(row,col),9)) <> "=iferror" Then
                    ReadArr(row,col) = "=iferror(" & Right(ReadArr(row,col), Len(ReadArr(row,col))=1) & ","""")"
                End If
            End If
        Next
    Next
Selection.Formula2 = ReadArr
Erase ReadArr
Else
formulaString = Selection.FormulaR1C1
If Left(FormulaString,1) = "=" Then
    If LCase(Left(FormulaString,9)) <> "=if.error" Then
        Selection.Formula2 = "=iferror(" & Right(FormulaString, Len(FormulaString)-1) & ","""")"
    End If
End If
End If
```

3. Cerrar VBA

4. Vincular la macro a un botón