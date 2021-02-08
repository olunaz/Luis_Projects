Sub ExtraerTransferenciasEmitidas(JobReporte)
    'BD DE DATOS'
    Sheets("RESUMEN TRX").Select
    Dim Estado, Turno, Moneda As String
    'Cambiar a Integer
    Dim Operaciones As String
    'Cambiar a Double
    Dim Importe As String

    indice = 0

    
    Call AbrirTodosLosReportesTotalesEmitidos(JobReporte, "-REMUNERACIO", 21)
    
    With Con
        'Validacion de existencia de reporte

        If .Leer(8, 69, 8) <> "EMITIDAS" And .Leer(9, 58, 5) <> "SOLES" _
            And "INTERMEDIA" <> .Leer(9, 66, 10) Then


        
        ' TRANSFERENCIAS EMITIDAS SOLES - INTERMEDIA
        If .Leer(8, 69, 8) = "EMITIDAS" And .Leer(9, 58, 5) = "SOLES" _
            And "INTERMEDIA" = .Leer(9, 66, 10) Then
    
            Moneda = .Leer(9, 58, 5)
            Turno = .Leer(9, 66, 10)
            Estado = .Leer(8, 69, 8)


            If .Leer(8, 83, 10) = fecha_guiones_japon Then
                ult3270 = 26
                EstoyEnUltimaPagina = False
                'CAPTURA SOLES
                .Escribir "F DOLARES", 2, 15
                .IngresarTecla "enter"

                .ColocarCursor 18, 30
                .IngresarTecla "f11"
                
                Operaciones = Trim(.Leer(17, 55, 12))
                Importe = Trim(.Leer(17, 91, 14))

                'Imprimir datos
                Cells(2, 2) = Estado
                Cells(2, 3) = Turno
                Cells(2, 4) = Moneda
                Cells(2, 5) = Operaciones
                Cells(2, 6) = Importe


                'CAPTURA DOLARES
                .Escribir "m", 2, 15
                .IngresarTecla "f8"
     
                'VALIDACION DE LA ULTIMA PAGINA
                If .Leer(1, 64, 3) = .Leer(1, 78, 3) Then
                    
                    'Nuevos montos
                    
                    Moneda = .Leer(25, 27, 8)
                    Operaciones = .Leer(25, 55, 12)
                    Importe = .Leer(25, 92, 12)

                    'Imprimir datos
                    Cells(3, 2) = Estado
                    Cells(3, 3) = Turno
                    Cells(3, 4) = Moneda
                    Cells(3, 5) = Operaciones
                    Cells(3, 6) = Importe

                End If

             Else
                .IngresarTecla "f3" 'Si no corresponde a la fecha, cerrar
            End If
         Else
            .IngresarTecla "f3"
         
        End If

        'PASA AL SIGUIENTE REPORTE
        .IngresarTecla "f3"

        ' TRANSFERENCIAS EMITIDAS SOLES - MAÑANA
        If .Leer(8, 69, 8) = "EMITIDAS" And .Leer(9, 58, 5) = "SOLES" _
            And "MAÑANA" = .Leer(9, 66, 6) Then
    
            Moneda = .Leer(9, 58, 5)
            Turno = .Leer(9, 66, 10)
            Estado = .Leer(8, 69, 8)


            If .Leer(8, 83, 10) = fecha_guiones_japon Then
                ult3270 = 26
                EstoyEnUltimaPagina = False
                'CAPTURA SOLES
                .Escribir "F DOLARES", 2, 15
          
                
                .IngresarTecla "enter"

                .ColocarCursor 12, 30
                .IngresarTecla "f11"

              
                Operaciones = Trim(.Leer(11, 55, 12))
                Importe = Trim(.Leer(11, 92, 14))


                'Imprimir datos
                Cells(4, 2) = Estado
                Cells(4, 3) = Turno
                Cells(4, 4) = Moneda
                Cells(4, 5) = Operaciones
                Cells(4, 6) = Importe


                'CAPTURA DOLARES
                .Escribir "m", 2, 15
                .IngresarTecla "f8"
     
                'VALIDACION DE LA ULTIMA PAGINA
                If .Leer(1, 64, 3) = .Leer(1, 78, 3) Then
                    
                    'Nuevos montos
                    Operaciones = .Leer(25, 55, 12)
                    Importe = .Leer(25, 92, 12)
                    Moneda = .Leer(25, 27, 8)

                    'Imprimir datos
                    Cells(5, 2) = Estado
                    Cells(5, 3) = Turno
                    Cells(5, 4) = Moneda
                    Cells(5, 5) = Operaciones
                    Cells(5, 6) = Importe

                End If

             Else
                .IngresarTecla "f3" 'Si no corresponde a la fecha, cerrar
            End If
         Else
            .IngresarTecla "f3"
         
        End If

        'PASA AL SIGUIENTE REPORTE
        .IngresarTecla "f3"

        ' TRANSFERENCIAS EMITIDAS SOLES - TARDE
        If .Leer(8, 69, 8) = "EMITIDAS" And .Leer(9, 58, 5) = "SOLES" _
            And "TARDE" = .Leer(9, 66, 5) Then
    
            Moneda = .Leer(9, 58, 5)
            Turno = .Leer(9, 66, 10)
            Estado = .Leer(8, 69, 8)


            If .Leer(8, 83, 10) = fecha_guiones_japon Then
                ult3270 = 26
                EstoyEnUltimaPagina = False
                'CAPTURA SOLES
                .Escribir "F DOLARES", 2, 15
                .IngresarTecla "f8"

                .ColocarCursor 16, 30
                .IngresarTecla "f11"

                Operaciones = .Leer(15, 55, 12)
                Importe = .Leer(15, 92, 12)

                'Imprimir datos
                Cells(6, 2) = Estado
                Cells(6, 3) = Turno
                Cells(6, 4) = Moneda
                Cells(6, 5) = Operaciones
                Cells(6, 6) = Importe


                'CAPTURA DOLARES
                .Escribir "m", 2, 15
                .IngresarTecla "f8"
     
                'VALIDACION DE LA ULTIMA PAGINA
                If .Leer(1, 64, 3) = .Leer(1, 78, 3) Then
                    
                    'Nuevos montos
                    Operaciones = .Leer(25, 55, 12)
                    Importe = .Leer(25, 92, 12)

                    'Imprimir datos
                    Cells(7, 2) = Estado
                    Cells(7, 3) = Turno
                    Cells(7, 4) = Moneda
                    Cells(7, 5) = Operaciones
                    Cells(7, 6) = Importe

                End If

             Else
                .IngresarTecla "f3" 'Si no corresponde a la fecha, cerrar
            End If
         Else
            .IngresarTecla "f3"
         
        End If
        .IngresarTecla "f3"
    End With

End Sub

Sub ExtraerTransferenciasRecibidas(JobReporte)
    'BD DE DATOS'
    Sheets("RESUMEN TRX").Select
    Dim Estado, Turno, Moneda As String
    Dim Operaciones As Integer
    Dim Importe As Double

    indice = 0

    With Con
        'TRANSFERNCIAS RECIBIDAS
        Call AbrirTransferenciasRecibidas(JobReporte)

        ' TRANSFERENCIAS RECIBIDAS SOLES - INTERMEDIA
        If "INTERMEDIA" = .Leer(9, 76, 10) Then
    
            Turno = .Leer(9, 76, 10)
            Estado = "RECIBIDAS"
            
            

            If .Leer(8, 64, 10) = fecha_guiones_corea Then
                ult3270 = 26
                EstoyEnUltimaPagina = False

                'IMPRESION DE DATOS
                .Escribir "m", 2, 15
                .IngresarTecla "f8"
                
                 'VALIDACION DE LA ULTIMA PAGINA
                If .Leer(1, 64, 3) = .Leer(1, 78, 3) Then

                    'SOLES
                    Cells(8, 2) = Estado
                    Cells(8, 3) = Turno
                    Cells(8, 4) = .Leer(24, 44, 5)
                    Cells(8, 5) = .Leer(24, 15, 6)
                    Cells(8, 6) = .Leer(24, 28, 14)

                    'DOLARES
                    Cells(9, 2) = Estado
                    Cells(9, 3) = Turno
                    Cells(9, 4) = .Leer(25, 44, 8)
                    Cells(9, 5) = .Leer(25, 15, 6)
                    Cells(9, 6) = .Leer(25, 28, 14)

                End If

             Else
                .IngresarTecla "f3" 'Si no corresponde a la fecha, cerrar
            End If
         Else
            .IngresarTecla "f3"
         
        End If
        
        .IngresarTecla "f3"

        ' TRANSFERENCIAS RECIBIDAS SOLES - MAÑANA
        If "MAÑANA" = .Leer(9, 76, 6) Then
    
            Turno = .Leer(9, 76, 6)
            Estado = "RECIBIDAS"

            If .Leer(8, 64, 10) = fecha_guiones_corea Then
                ult3270 = 26
                EstoyEnUltimaPagina = False

                'IMPRESION DE DATOS
                .Escribir "m", 2, 15
                .IngresarTecla "f8"
                
                 'VALIDACION DE LA ULTIMA PAGINA
                If .Leer(23, 1, 6) = "0TOTAL" Then

                    'SOLES
                    Cells(10, 2) = Estado
                    Cells(10, 3) = Turno
                    Cells(10, 4) = .Leer(23, 44, 5) 'MONEDA
                    Cells(10, 5) = .Leer(23, 15, 6) 'OPERACIONES
                    Cells(10, 6) = .Leer(23, 28, 14) 'Importe

                    'DOLARES
                    Cells(11, 2) = Estado
                    Cells(11, 3) = Turno
                    Cells(11, 4) = .Leer(24, 44, 8) 'MONEDA
                    Cells(11, 5) = .Leer(24, 15, 6) 'OPERACIONES
                    Cells(11, 6) = .Leer(24, 28, 14) 'Importe

                End If

             Else
                .IngresarTecla "f3" 'Si no corresponde a la fecha, cerrar
            End If
         Else
            .IngresarTecla "f3"
         
        End If
        
        .IngresarTecla "f3"

        ' TRANSFERENCIAS RECIBIDAS SOLES - TARDE
        If "TARDE" = .Leer(9, 76, 5) Then
    
            Turno = .Leer(9, 76, 5)
            Estado = "RECIBIDAS"

            If .Leer(8, 64, 10) = fecha_guiones_corea Then
                ult3270 = 26
                EstoyEnUltimaPagina = False

                'IMPRESION DE DATOS
                .Escribir "m", 2, 15
                .IngresarTecla "f8"
                
                 'VALIDACION DE LA ULTIMA PAGINA
                If .Leer(15, 1, 6) = "0TOTAL" Then

                    'SOLES
                    Cells(12, 2) = Estado
                    Cells(12, 3) = Turno
                    Cells(12, 4) = .Leer(14, 44, 5) 'MONEDA
                    Cells(12, 5) = .Leer(14, 15, 6) 'OPERACIONES
                    Cells(12, 6) = .Leer(14, 28, 14) 'Importe

                    'DOLARES
                    Cells(13, 2) = Estado
                    Cells(13, 3) = Turno
                    Cells(13, 4) = .Leer(15, 44, 8) 'MONEDA
                    Cells(13, 5) = .Leer(15, 15, 6) 'OPERACIONES
                    Cells(13, 6) = .Leer(15, 28, 14) 'Importe

                End If

             Else
                .IngresarTecla "f3" 'Si no corresponde a la fecha, cerrar
            End If
         Else
            .IngresarTecla "f3"
         
        End If

    End With

    'PEGADO DE LAS FECHAS
    Dim n
    For n = 2 To 13

        Cells(n, 1) = fecha_formato_largo

    Next n

End Sub












