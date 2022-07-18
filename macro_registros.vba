Sub EscribirRegistro()
    Set a = Worksheets("ALUMNOS")
    Set carga = Worksheets("CARGA")
    Set materias = Worksheets("MATERIAS")
    
    to_rows = a.Cells(Rows.Count, 2).End(xlUp).Row
    to_rows2 = materias.Cells(Rows.Count, 2).End(xlUp).Row
    
    Incremento = 2
    Inc_Semestre = 1
    
    For row1 = 2 To to_rows
        For row2 = 2 To to_rows2
            If a.Cells(row1, 3).Value = "L3D" + CStr(Inc_Semestre) Then
                If materias.Cells(row2, 3).Value = "L3D" + CStr(Inc_Semestre) Then
                    carga.Cells(Incremento, 1).Value = a.Cells(row1, 1)
                    carga.Cells(Incremento, 2).Value = "BAJ"
                    carga.Cells(Incremento, 3).Value = "UG"
                    carga.Cells(Incremento, 4).Value = "202340"
                    carga.Cells(Incremento, 5).Value = materias.Cells(row2, 1)
                    carga.Cells(Incremento, 6).Value = materias.Cells(row2, 2)
                    carga.Cells(Incremento, 7).Value = "BFA3DGAMANIX"
                    carga.Cells(Incremento, 8).Value = "L3D" + CStr(Inc_Semestre)
                    Incremento = Incremento + 1
                End If
            Else
                Inc_Semestre = Inc_Semestre + 2
            End If
        Next row2
    Next row1
End Sub
