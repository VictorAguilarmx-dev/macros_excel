# Copia y pega la suma en una celda

1. Abrir VBA

2. Escribir: 

```
Sub copypastesum()

dim MyDataObj As MSForms.DataObject
Dim v as String
Dim rng As Range

set myDataObj = New MSForms.DataObject
MydataObj.Clear
MyDataObj.SetText Application.WorksheetFunction.Sum(Selection.SpecialCells(xlCellTypeVisible))

v = MyDataObj.GetText
Set rng = Application.InputBox("Selecciona donde vamos a pegar los datos:", xTitled, Type:=8)
rng.Values = v
End Sub
```

3. Seleccionar los datos

4. Correr la macro