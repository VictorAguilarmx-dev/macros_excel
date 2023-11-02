# Filtro Automático

1. Insertar un cuadro de texto de la sección ActiveX

2. Dar doble click, esto abrirá el panel de VBA de Excel

3. Escribir:

```
Private Sub TextBox1_Change()

filter = "*"&Sheets("<sheet_name>").TextBox1.Text & "*"
Range("<initial_cell_title>").AutoFilter field:=<field_table_filter_int>, Criterial:=filter

End Sub
```

4. Cerrar el editor VBA