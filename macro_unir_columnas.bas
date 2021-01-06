Attribute VB_Name = "Módulo1"
Sub unir_columnas()
    'Por Armando Valdés
    'avaldes@ciestaam.edu.mx
ci = Columns("a").Column 'columna inicial a unir
cf = Columns("d").Column 'columna final a unir
cd = Columns("f").Column 'columna para unión
f = 1 'fila inicial de datos
For i = ci To cf
    uf = Cells(Rows.Count, i).End(xlUp).Row
ud = Cells(Rows.Count, cd).End(xlUp).Row + 0 'lo va a pegar en la primera fila
    Range(Cells(f, i), Cells(uf, i)).Copy Cells(ud, cd)
Next
End Sub
