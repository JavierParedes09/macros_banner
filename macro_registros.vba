Sub EscribirRegistro()
    Set a = Worksheets("ALUMNOS")
    Set carga = Worksheets("CARGA")
    Set materias = Worksheets("MATERIAS")
    Set generar = Worksheets("GENERAR")
    
    to_rows = a.Cells(Rows.Count, 2).End(xlUp).Row
    to_rows2 = materias.Cells(Rows.Count, 2).End(xlUp).Row
    to_rows3 = generar.Cells(Rows.Count, 2).End(xlUp).Row
    
    Incremento = 2
    Inc_Semestre = 1
    Dim array_carreras(12) As String
    array_carreras(0) = "L3D"
    array_carreras(1) = "LN"
    array_carreras(2) = "LDMM"
    array_carreras(3) = "LP"
    array_carreras(4) = "II"
    array_carreras(5) = "IM"
    array_carreras(6) = "IME"
    array_carreras(7) = "LNI"
    array_carreras(8) = "LAD"
    array_carreras(9) = "LCF"
    array_carreras(10) = "LMK"
    array_carreras(11) = "LD"
    array_carreras(12) = "MBA"
    array_carreras(13) = "MED"
    
    'Falta Codigo de carrera
    Dim array_codigos(14) As String
    array_codigos(0) = "BFA3DGAMANIX"
    array_codigos(1) = "BSNUTRITIONX"
    array_codigos(2) = "BFAFASHDESGX"
    array_codigos(3) = "BAPSYCX"
    array_codigos(4) = "BSENGRMGMTX"
    array_codigos(5) = "BSENGRMECHAX"
    array_codigos(6) = "BSENGRMECHNX"
    array_codigos(7) = "BBABINTX"
    array_codigos(8) = "BBAMGMTX"
    array_codigos(9) = "BBAFINANCEX"
    array_codigos(10) = "BBAMKTGX"
    array_codigos(11) = ""
    array_codigos(12) = ""
    array_codigos(13) = ""
    Inc_Array = 0
    
    If generar.Cells(10, 3).Value = "" And generar.Cells(10, 4).Value = "" And generar.Cells(10, 5).Value = "" Then
        generar.Cells(11, 3).Value = "SELECCIONA LOS DATOS POR FAVOR!"
    Else
        For i = 0 To 11
            For row1 = 2 To to_rows
                For row2 = 2 To to_rows2
                    If a.Cells(row1, 3).Value = array_carreras(Inc_Array) + CStr(Inc_Semestre) Then
                        If materias.Cells(row2, 3).Value = array_carreras(Inc_Array) + CStr(Inc_Semestre) Then
                            carga.Cells(Incremento, 1).Value = a.Cells(row1, 1)
                            carga.Cells(Incremento, 2).Value = generar.Cells(10, 3).Value
                            carga.Cells(Incremento, 3).Value = generar.Cells(10, 4).Value
                            carga.Cells(Incremento, 4).Value = generar.Cells(10, 5).Value
                            carga.Cells(Incremento, 5).Value = materias.Cells(row2, 1)
                            carga.Cells(Incremento, 6).Value = materias.Cells(row2, 2)
                            carga.Cells(Incremento, 7).Value = array_codigos(Inc_Array)
                            carga.Cells(Incremento, 8).Value = array_carreras(Inc_Array) + CStr(Inc_Semestre)
                            Incremento = Incremento + 1
                        End If
                    Else
                        Inc_Semestre = Inc_Semestre + 2
                    End If
                Next row2
                Inc_Semestre = 1
            Next row1
            Inc_Array = Inc_Array + 1
        Next
        generar.Cells(11, 3).Value = ""
        generar.Cells(11, 3).Value = "FELICIDADES, DATOS PROCESADOS!"
    End If
End Sub