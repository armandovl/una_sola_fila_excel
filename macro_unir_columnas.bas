Attribute VB_Name = "M�dulo1"
Sub unir_columnas()
'Por Armando Vald�s
'avaldes@ciestaam.edu.mx

ci = Columns("a").Column 'columna inicial a unir
cf = Columns("d").Column 'columna final a unir
cd = Columns("f").Column 'columna para uni�n
f = 1 'fila inicial de datos
For i = ci To cf
    uf = Cells(Rows.Count, i).End(xlUp).Row
    ud = Cells(Rows.Count, cd).End(xlUp).Row + 1
    Range(Cells(f, i), Cells(uf, i)).Copy Cells(ud, cd)
Next
End Sub
