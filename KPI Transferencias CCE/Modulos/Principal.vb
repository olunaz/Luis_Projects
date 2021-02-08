Public Const lenusuario = 7
Public Const lenpass = 8

Public sesion As String
Public fecha3270, fecha_guiones_japon, fecha_guiones_corea, fecha_formato_largo, fecha_formato_corto, fecha_corto_punto, fechareporte As String
Public Con As Conexion3270

Sub Principal()
    Dim DevRecibidas(), DevEmitidas() As Variant
     
    Application.ScreenUpdating = False
    Set Con = New Conexion3270
    
    Call ConectarHost
    Call IngresarHost
    Call ExtraerReportes(DevRecibidas(), DevEmitidas())
        
'    With ThisWorkbook
'        .Sheets("DEV MES ACTUAL").Range(Cells(fImpresion, "B")).Resize( _
'        UBound(DevEmitidas, 2) + 1, UBound(DevEmitidas, 1) + 1) = _
'        Application.WorksheetFunction.Transpose(DevEmitidas)
'
'        .Sheets("DEV MES ACTUAL").Range(Cells(fImpresion, "B")).Resize( _
'        UBound(DevRecibidas, 2) + 1, UBound(DevRecibidas, 1) + 1) = _
'        Application.WorksheetFunction.Transpose(DevRecibidas)
'    End With
    
    Application.ScreenUpdating = True
End Sub

Private Sub ConectarHost()
    Do
        sesion = UCase(InputBox("¿En qué sesión desea trabajar?", "Ingresar sesión"))
    Loop Until sesion = "A" Or sesion = "B" Or sesion = "C" Or sesion = vbNullString

    If sesion = vbNullString Then
        MsgBox ("No se ingresó sesión. La macro se detendrá.")
        End
    End If
    
    Con.init (sesion)
End Sub

Private Sub IngresarHost()
    Inicio.Show
    user = Inicio.tRegistro.Value
    pass = Inicio.tContra.Value
    
    Do While UCase(Left(user, 1)) <> "P" Or Len(user) <> lenusuario Or Len(pass) <> lenpass
    
        If UCase(Left(user, 1)) <> "P" Or Len(user) <> lenusuario Then
            MsgBox ("Registro inválido.")
        ElseIf Len(pass) <> lenpass Then
            MsgBox ("Contraseña inválida.")
        End If
        
        Inicio.Show
        user = Inicio.tRegistro.Value
        pass = Inicio.tContra.Value
    Loop
    
    With Con
        .Escribir "D", 24, 24
        .IngresarTecla "enter"
        Do Until .Leer(1, 34, 15) = "IOA ENTRY PANEL"
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop
        
        .Escribir user, 9, 39
        .Escribir pass, 11, 39
        .IngresarTecla "enter"
        
        Do Until .Leer(1, 2, 7) = "CTM790E" Or .Leer(1, 23, 11) = "CONTROL-D/V"
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop
        
        If .Leer(1, 2, 7) = "CTM790E" Then
            MsgBox "Contraseña errada. La macro se detendrá"
            End
        End If
    End With
    
End Sub
Private Sub ExtraerReportes(DevRecibidas(), DevEmitidas())

    Call ObtenerFechas
    Call LimpiarControlDEntryPanel
    
    With Con
'        .Escribir "*DEVDETA*", 7, 26
'        .Escribir "0928OFIC", 9, 26
'        .Escribir fecha3270, 10, 26
'        .Escribir fecha3270, 10, 36
'        .Escribir "Y", 13, 26
'
'        Call ExtraerDevoluciones("FEJP6400", DevEmitidas())
'        Call ExtraerDevoluciones("FEJP6340", DevRecibidas())
        
        Call LimpiarControlDEntryPanel

        .Escribir "0928OFIC", 9, 26
        .Escribir fecha3270, 10, 26
        .Escribir fecha3270, 10, 36
        .Escribir "Y", 13, 26
        
        Call ExtraerTransferenciasEmitidas("FEJPCC04")
        
        Call ExtraerTransferenciasRecibidas("FEJPCC10")

        'Call ExtraerTransferencias("FEJP6520")
        
        'AGREGAR VALIDACION DE LEER OPTIONS EN LA PANTALLA DEL CONTROL D
        If Trim(.Leer(5, 6, 7)) <> "OPTIONS" Then
        .IngresarTecla ("F3")
        End If
        
    
    End With
    
    'Call Limpiar(fecha_corto_punto)

    MsgBox ("FINALIZADO")


End Sub

Private Sub LimpiarControlDEntryPanel()
    fBorrar = Array(2, 5, 7, 9, 10, 10, 11, 13, 15, 16)
    cBorrar = Array(15, 19, 26, 26, 26, 36, 26, 26, 26, 26)

    With Con
        For indice = 0 To 9
            .ColocarCursor fBorrar(indice), cBorrar(indice)
            .IngresarTecla "fin"
        Next
    End With
End Sub

Private Sub ObtenerFechas()

    Do
    fecha_solicitada = InputBox("Por favor, ingresar la fecha a buscar." & vbNewLine & _
    "De no ingresar una fecha, la macro buscará el último día hábil." & vbNewLine & _
    "FORMATO DDMMAA (Ejemplo: 15 de abril de 2020 es 150420", "Ingresar fecha")
    
    If Len(fecha_solicitada) = 6 Then
    fecha_a_tratar = DateSerial( _
            CInt("20" & Right(fecha_solicitada, 2)), _
            CInt(Mid(fecha_solicitada, 3, 2)), _
            CInt(Left(fecha_solicitada, 2)))
            
    validacion_fecha = True
            
    Else
    
    validacion_fecha = False
    
    End If
    
    Loop Until validacion_fecha
    
    
    If Len(Day(fecha_a_tratar)) = 1 Then
        dia = "0" & CStr(Day(fecha_a_tratar))
    Else
        dia = CStr(Day(fecha_a_tratar))
    End If
    If Len(Month(fecha_a_tratar)) = 1 Then
        mes = "0" & CStr(Month(fecha_a_tratar))
    Else
        mes = CStr(Month(fecha_a_tratar))
    End If
    
    anho = CStr(Year(fecha_a_tratar))
    
    fecha3270 = dia & mes & Right(anho, 2)
    fecha_formato_largo = CDate(dia & "/" & mes & "/" & anho)
    fecha_formato_corto = dia & "/" & mes & "/" & Right(anho, 2)
    fecha_corto_punto = dia & "." & mes & "." & Right(anho, 2)
    fecha_guiones_japon = anho & "-" & mes & "-" & dia
    fecha_guiones_corea = dia & "-" & mes & "-" & anho
    
End Sub

Private Sub ExtraerDevoluciones(JobReporte, ArrayDevoluciones())
    Dim Banco, Turno, Moneda, Motivo, Oficina, Cuenta As String
    Dim Importe As Double
    indice = 0
    
    With Con
        Call AbrirTodosLosReportesBuscados(JobReporte, "DEVDETA", 25)
        
        Do
            If fecha_guiones_japon = .Leer(8, 86, 10) Then
                ult3270 = 26
                EstoyEnUltimaPagina = False
                
                Do
                    For f3270 = 6 To 26
                        If .Leer(f3270, 25, 20) = "NO MORE DATA TO VIEW" Then
                            ult3270 = f3270
                            EstoyEnUltimaPagina = True
                            Exit For
                        End If
                    Next f3270
                    
                    For f3270 = 7 To ult3270
                        glosaPatron = Trim(.Leer(f3270, 2, 12))
                        lineaCompleta = Trim(.Leer(f3270, 1, 132))
                        
                        If glosaPatron = "" And _
                        (Trim(.Leer(f3270 - 1, 41, 12)) = "DEVOLUCIONES" Or _
                        Trim(.Leer(f3270 + 1, 2, 4)) = "----") Then
                            'MONEDA - TURNO
                            Moneda = Trim(.Leer(f3270, 60, 10))
                            Turno = Trim(.Leer(f3270, 71, 15))
                        
                        ElseIf IsNumeric(glosaPatron) And Len(glosaPatron) = 4 Then
                            Oficina = glosaPatron
                        
                        ElseIf Left(glosaPatron, 5) = "BANCO" Then
                            'Primera fila de una devolución
                            Banco = Trim(.Leer(f3270, 2, 100))
                            
                            'En la segunda fila de una devolución no hay
                            'datos relevantes
                        ElseIf IsNumeric(glosaPatron) And Len(glosaPatron) = 7 And _
                        glosaPatron <> lineaCompleta Then
                            'Tercera fila de una devolución
                            Motivo = Trim(.Leer(f3270, 127, 3))
                            Cuenta = Trim(.Leer(f3270, 64, 24))
                            Monto = CDbl(Trim(.Leer(f3270, 10, 25)))
                            
                            ReDim Preserve ArrayDevoluciones(0 To 7, 0 To indice)
                            ArrayDevoluciones(0, indice) = Banco
                            ArrayDevoluciones(1, indice) = fecha_formato_largo
                            ArrayDevoluciones(2, indice) = Turno
                            ArrayDevoluciones(3, indice) = Moneda
                            ArrayDevoluciones(4, indice) = Monto
                            ArrayDevoluciones(5, indice) = Motivo
                            ArrayDevoluciones(6, indice) = Oficina
                            ArrayDevoluciones(7, indice) = Cuenta
                            
                            Banco = Turno = Moneda = Motivo = Oficina = Cuenta = ""
                            Monto = 0
                            indice = indice + 1
                        
                        End If
                        
                    Next f3270
                    .IngresarTecla "f8"
                    
                Loop Until EstoyEnUltimaPagina
                .IngresarTecla "f3"
             Else
                .IngresarTecla "f3" 'Si no corresponde a la fecha, cerrar
            End If
            
        Loop Until .Leer(1, 2, 11) = "ACTIVE LIST"
    End With
End Sub

Sub AbrirTodosLosReportesBuscados(JobReporte, Optional GlosaFiltro, Optional colGlosaFiltro)
    With Con
        .Escribir JobReporte, 11, 26
        .IngresarTecla "enter"
        
        If .Leer(1, 2, 7) = "IOAE53E" Then
            MsgBox "No se encontró reporte con las fechas ingresadas." & vbNewLine & _
            "La macro se detendrá."
            End
            
        ElseIf .Leer(1, 2, 11) <> "ACTIVE LIST" Then
            MsgBox "Se encontró un error desconocido." & vbNewLine & _
            "La macro se detendrá."
            End
            
        ElseIf .Leer(1, 2, 11) = "ACTIVE LIST" Then
            For f3270 = 4 To 25
                If GlosaFiltro <> "" Then
                    If .Leer(f3270, colGlosaFiltro, Len(GlosaFiltro)) = GlosaFiltro Then
                        .Escribir "V", f3270, 2
                    End If
                Else
                    .Escribir "V", f3270, 2
                End If
                
                If .Leer(f3270, 27, 29) = _
                "B O T T O M    O F    L I S T" Then Exit For
            Next f3270
            .IngresarTecla "enter"
            
        End If
    End With
End Sub
''''''''''''''''''''''''''''''

Sub AbrirTodosLosReportesTotalesEmitidos(JobReporte, Optional GlosaFiltro, Optional colGlosaFiltro)

    With Con

        ult3270 = 26
        EstoyEnUltimaPagina = False

        .Escribir JobReporte, 11, 26
        .IngresarTecla "enter"
        
        If .Leer(1, 2, 7) = "IOAE53E" Then
            MsgBox "No se encontró reporte con las fechas ingresadas." & vbNewLine & _
            "La macro se detendrá."
            End
            
        ElseIf .Leer(1, 2, 11) <> "ACTIVE LIST" Then
            MsgBox "Se encontró un error desconocido." & vbNewLine & _
            "La macro se detendrá."
            End
            
        ElseIf .Leer(1, 2, 11) = "ACTIVE LIST" Then
            For f3270 = 4 To ult3270
                
                fechareporte = .Leer(f3270, 34, 8)

                If fechareporte = fecha_formato_corto And _
                .Leer(f3270, colGlosaFiltro, Len(GlosaFiltro)) <> GlosaFiltro Then
                    .Escribir "V", f3270, 2
                Else
                    .Escribir "V", f3270, 2
                End If

        
                If .Leer(f3270, 27, 29) = _
                "B O T T O M    O F    L I S T" Then
                    ult3270 = f3270
                    EstoyEnUltimaPagina = True
                    Exit For
                End if

            Next f3270
            .IngresarTecla "f8"
        End If
        .IngresarTecla "enter"
    End With
End Sub


''''''''''''''''''''''''''''''''''''
Sub AbrirTransferenciasEmitidas(JobReporte)
    With Con
        .Escribir JobReporte, 11, 26
        .IngresarTecla "enter"
        
        If .Leer(1, 2, 7) = "IOAE53E" Then
            MsgBox "No se encontró reporte con las fechas ingresadas." & vbNewLine & _
            "La macro se detendrá."
            End
            
        ElseIf .Leer(1, 2, 11) <> "ACTIVE LIST" Then
            MsgBox "Se encontró un error desconocido." & vbNewLine & _
            "La macro se detendrá."
            End
            
        ElseIf .Leer(1, 2, 11) = "ACTIVE LIST" Then
            
            .Escribir "V", 4, 2
            .Escribir "V", 16, 2
            .Escribir "V", 20, 2
            
'            For f3270 = (5,16,20)
'                Escribir "V", f3270, 2
'                If .Leer(f3270, 27, 29) = _
'                "B O T T O M    O F    L I S T" Then Exit For
'            Next f3270
            
            .IngresarTecla "enter"
            
        End If
    End With
End Sub

Sub AbrirTransferenciasRecibidas(JobReporte)
    With Con
            .Escribir JobReporte, 11, 26
            .IngresarTecla "enter"
            
            If .Leer(1, 2, 7) = "IOAE53E" Then
                MsgBox "No se encontró reporte con las fechas ingresadas." & vbNewLine & _
                "La macro se detendrá."
                End
                
            ElseIf .Leer(1, 2, 11) <> "ACTIVE LIST" Then
                MsgBox "Se encontró un error desconocido." & vbNewLine & _
                "La macro se detendrá."
                End
                
            ElseIf .Leer(1, 2, 11) = "ACTIVE LIST" Then
                For f3270 = 4 To 11
                    If fecha_formato_corto = .Leer(f3270, 34, 8) Then
                        .Escribir "V", f3270, 2
                        f3270 = f3270 + 1
                    End If
                    If .Leer(f3270, 27, 29) = _
                    "B O T T O M    O F    L I S T" Then Exit For
                Next f3270
                .IngresarTecla "enter"
                
            End If
        End With
End Sub










