Sub ExtraerTotalesEmitidasRecibidas(fecha_3270, sesion)

    'Imprimir fecha del reporte
    fecha_formato = "31/03/20"
    For filaExcel = 2 To 13
        Cells(filaExcel, 1) = fecha_formato
    Next
    
    Call AbrirTodosLosReportesBuscadosTran("FEJPCC04")

'PARA EMITIDOS INTERMEDIA'
BUCLE:

'Validacion INTERMEDIA DOLARES
If .Leer(8, 91, 2) = Left(fecha_3270, 2) _
And "- INTERMEDIA" = .Leer(9, 64, 12) Then
    .Escribir "f DOLARES ", 2, 15
    .IngresarTecla "enter"

    'SIRVE PARA ALGO
    If .Leer(9, 19, 8) = "THE FIND" Then
        .Escribir "3", 11, 55
        .IngresarTecla "enter"
    End If
    'SIRVE PARA ALGO

    
Else
    .IngresarTecla "f3"
    If .Leer(1, 50, 11) = "ENTRY PANEL" Then
        MsgBox "NO SE ENCONTRO NINGUN REPORTE"
        Exit Sub
    End If
    
    GoTo BUCLE
End If

n = 6

'SI NO ESTA EN ESTA PANTALLA SUBIR'
BUCLE2:
If n > 26 Then
    .IngresarTecla "f7"
    n = 6
End If

'SOLES'
Q = 6
QUEPASO:
If .Leer(Q, 58, 7) = "DOLARES" Then
    GoTo ESTEES
Else:
    If Q = 27 Then
        Q = 6
        .IngresarTecla "f5"
        .Escribir "3", 11, 55
        .IngresarTecla "enter"
        GoTo QUEPASO
    Else
        Q = Q + 1
        GoTo QUEPASO
    End If
End If

ESTEES:
c = 6
AAA:

If .Leer(c, 49, 12) = "MONEDA SOLES" And .Leer(c, 97, 4) = "DOCS" Then
    n = c
    GoTo COGERNOMAS
Else
    If c = 27 Then
        GoTo GATO
    End If
    c = c + 1
    GoTo AAA
End If

GATO:

If .Leer(n, 49, 12) = "MONEDA SOLES" And .Leer(n, 97, 4) = "DOCS" Then

    If .Leer(n + 1, 49, 12) = "MONEDA SOLES" Or .Leer(n - 1, 97, 4) = "DOCS" Then
     
    Else
        .IngresarTecla "[pf7]"
        
        
        
        n = 6
        GoTo BUCLE2
        
    End If

    'NUMERO DE DOCUMENTOS'
    
COGERNOMAS:
    
    Sheets("TOTALREPORTES").Select

    Cells(2, 5) = .Leer(n, 90, 7)

    autECLPSObj.SetCursorPos 14, 2
    .IngresarTecla "f11"
    
    
        
    Cells(2, 6) = .Leer(n, 120, 13)
    
    autECLPSObj.SetCursorPos 14, 2
    .IngresarTecla "f10"
    
    
    autECLPSObj.SetCursorPos 2, 15
Else
    n = n + 1
    GoTo BUCLE2
End If

'DOLARES'
n = 6
.Escribir "m", 2, 15
.IngresarTecla "[pf8]"


    
BUCLE3:

If .Leer(n, 49, 14) = "MONEDA DOLARES" And .Leer(n, 97, 4) = "DOCS" Then
    'NUMERO DE DOCUMENTOS'
    Sheets("TOTALREPORTES").Select
    'Range("E1").End(xlDown).Offset(1, 0) = .Leer(n, 91, 6)
    Cells(3, 5) = .Leer(n, 91, 6)
    
    'MONTO'
    autECLPSObj.SetCursorPos 14, 2
    .IngresarTecla "[pf11]"
    
    
    
    'Range("F1").End(xlDown).Offset(1, 0) = Trim(.Leer(n, 120, 13))
    
    Cells(3, 6) = Trim(.Leer(n, 120, 13))
    
    
    autECLPSObj.SetCursorPos 14, 2
    .IngresarTecla "[pf10]"
    
    
    autECLPSObj.SetCursorPos 2, 15
Else
    n = n + 1
    GoTo BUCLE3
End If

.IngresarTecla "f3"
    
'PARA EMITIDOS MAÑANA'
'RETROCEDER A GIVETO'

BucleRetro10:
    
If .Leer(27, 19, 8) = "G GIVETO" Then
Else
    .IngresarTecla "[pf3]"
    GoTo BucleRetro10
End If

For i = 4 To 25
    .Escribir "v", i, 2
Next
.IngresarTecla "[enter]"

buclee:

If Left(fecha_3270, 2) = .Leer(8, 91, 2) And "- MAÑANA" = .Leer(9, 64, 8) Then
    .Escribir "f DOLARES ", 2, 15
    
    '.Escribir "3", 11, 55
    .IngresarTecla "enter"
    
    If .Leer(11, 19, 6) = "PLEASE" Then
        .Escribir "3", 11, 55
        .IngresarTecla "[enter]"
    End If
    
Else: .IngresarTecla "[pf3]"
    If .Leer(1, 2, 11) = "ACTIVE LIST" Then
        For i = 4 To 25
            .Escribir "v", i, 2
        Next
        
        .IngresarTecla "[enter]"
        GoTo bucleEmitidaT
'        MsgBox ("NO SE ENCONTRO NINGUN REPORTE")
'        Exit Sub
    End If
    
    GoTo buclee
End If

'BUSCAR SI ES TU DOLARES ENTONCES SINO F5'
porsilasmoscas:
n = 6

Do Until n > 26
    If .Leer(n, 58, 7) = "DOLARES" Then
        GoTo BUCLE4
    Else
        n = n + 1
    End If
Loop

.IngresarTecla "f5"

GoTo porsilasmoscas

'BUSCAR SI ES TU DOLARES ENTONCES SINO F5'
n = 6

BUCLE4:

If n > 26 Then
    .IngresarTecla "f7"
    n = 6
End If

'SOLES'

If .Leer(n, 49, 12) = "MONEDA SOLES" And .Leer(n, 97, 4) = "DOCS" Then

    'NUMERO DE DOCUMENTOS'
    Sheets("TOTALREPORTES").Select
    'Range("E1").End(xlDown).Offset(1, 0) = .Leer(n, 90, 6)
    Cells(4, 5) = .Leer(n, 90, 6)

    'MONTO'
    autECLPSObj.SetCursorPos 14, 2
    .IngresarTecla "f11"
    
    'Range("F1").End(xlDown).Offset(1, 0) = Trim(.Leer(n, 120, 13))
    Cells(4, 6) = Trim(.Leer(n, 120, 13))
    autECLPSObj.SetCursorPos 14, 2
    .IngresarTecla "f10"
    
    autECLPSObj.SetCursorPos 2, 15
Else
    n = n + 1
    GoTo BUCLE4
End If

'DOLARES'
n = 6
.Escribir "m", 2, 15
.IngresarTecla "f8"

bucle5:

If .Leer(n, 49, 14) = "MONEDA DOLARES" And .Leer(n, 97, 4) = "DOCS" Then
    'NUMERO DE DOCUMENTOS'
    Sheets("TOTALREPORTES").Select
    'Range("E1").End(xlDown).Offset(1, 0) = .Leer(n, 91, 6)
    Cells(5, 5) = .Leer(n, 91, 6)
    
    'MONTO'
    autECLPSObj.SetCursorPos 14, 2
    .IngresarTecla "f11"
    
    'Range("F1").End(xlDown).Offset(1, 0) = Trim(.Leer(n, 120, 13))
    Cells(5, 6) = Trim(.Leer(n, 120, 13))
    autECLPSObj.SetCursorPos 14, 2
    .IngresarTecla "f10"
    autECLPSObj.SetCursorPos 2, 15
Else
    n = n + 1
    GoTo bucle5
End If

'PARA EMITIDOS TARDE'
bucleee:
bucleEmitidaT:
If Left(fecha_3270, 2) = .Leer(8, 91, 2) And "- TARDE" = .Leer(9, 64, 7) Then
    .Escribir "f DOLARES-", 2, 15
    .IngresarTecla "enter"

    If .Leer(9, 19, 8) = "THE FIND" Then
    
        .Escribir "3", 11, 55
        .IngresarTecla "enter"
        
    End If

Else: .IngresarTecla "f3"
    
    If .Leer(26, 2, 1) = "P" Then
        .IngresarTecla "f8"
        
        For i = 4 To 25
            .Escribir "v", i, 2
        Next
        
        .IngresarTecla "enter"
    End If
    GoTo bucleee
End If

'ES TUS DOLARES 2'
porsilasmoscas2:

n = 6

Do Until n > 26
    If .Leer(n, 58, 7) = "DOLARES" Then
        n = 6
        GoTo bucle6
    Else
        n = n + 1
    End If
Loop
.IngresarTecla "f5"

If .Leer(9, 19, 8) = "THE FIND" Then
    .Escribir "3", 11, 55
    .IngresarTecla "enter"
End If

GoTo porsilasmoscas2

'ES TUS DOLARES 2'

n = 6

bucle6:

If n > 26 Then
    .IngresarTecla "[pf7]"
    
    paternBucle = .Leer(6, 31, 24)
    If paternBucle = "T O P    O F   R A N G E" Then
        GoTo EmitidaDolares
    End If
    
    n = 6
End If

'SOLES'
If .Leer(n, 49, 12) = "MONEDA SOLES" And .Leer(n, 97, 4) = "DOCS" Then

    'SI EL TOTAL SOLES ESTA ARRIBA'
    If .Leer(n + 1, 49, 12) = "MONEDA SOLES" Or .Leer(n - 1, 97, 4) = "DOCS" Then
    
        'NUMERO DE DOCUMENTOS'
        Sheets("TOTALREPORTES").Select
        'Range("E1").End(xlDown).Offset(1, 0) = .Leer(n, 90, 6)
        Cells(6, 5) = .Leer(n, 90, 6)
        
        'MONTO'
        autECLPSObj.SetCursorPos 14, 2
        .IngresarTecla "f11"
        
        'Range("F1").End(xlDown).Offset(1, 0) = Trim(.Leer(n, 120, 13))
        
        Cells(6, 6) = Trim(.Leer(n, 120, 13))
        autECLPSObj.SetCursorPos 14, 2
        .IngresarTecla "f10"
        
        autECLPSObj.SetCursorPos 2, 15
    Else
        .IngresarTecla "f7"
        n = 6
        GoTo bucle6
    End If

Else
    n = n + 1
    GoTo bucle6
End If

EmitidaDolares:
'DOLARES'
n = 6
.Escribir "m", 2, 15
.IngresarTecla "f8"

bucle7:

If .Leer(n, 49, 14) = "MONEDA DOLARES" And .Leer(n, 97, 4) = "DOCS" Then

    'NUMERO DE DOCUMENTOS'
    Sheets("TOTALREPORTES").Select
    'Range("E1").End(xlDown).Offset(1, 0) = .Leer(n, 91, 6)
    Cells(7, 5) = .Leer(n, 91, 6)

    'MONTO'
    autECLPSObj.SetCursorPos 14, 2
    .IngresarTecla "f11"
    
    'Range("F1").End(xlDown).Offset(1, 0) = Trim(.Leer(n, 120, 13))
    Cells(7, 6) = Trim(.Leer(n, 120, 13))
    autECLPSObj.SetCursorPos 14, 2
    .IngresarTecla "f10"
    autECLPSObj.SetCursorPos 2, 15
Else
    n = n + 1
    GoTo bucle7
End If

'RETROCEDER A GIVETO'

BucleRetro:
If .Leer(27, 19, 8) = "G GIVETO" Then
Else
    .IngresarTecla "f3"
    GoTo BucleRetro
End If

.IngresarTecla "f3"

'SALDOS RECIBIDAS'
.Escribir "0928", 9, 26
.Escribir fecha_3270, 10, 26
.Escribir fecha_3270, 10, 36
.Escribir "FEJPcc10", 11, 26
.Escribir "Y", 13, 26
.IngresarTecla "enter"

V = 4
RepetirV:

'REPETIR V'
If .Leer(V, 2, 1) = " " Then
    .Escribir "v", V, 2
    V = V + 1
    GoTo RepetirV
Else
    .IngresarTecla "[enter]"
End If

'RECIBIDAS '

Buclee2:

If Left(fecha_3270, 2) = .Leer(8, 64, 2) And "- INTERMEDIA" = .Leer(9, 74, 12) Then
    .Escribir "m", 2, 15
    .IngresarTecla "f8"
    .IngresarTecla "enter"
    
Else
    .IngresarTecla "f3"
    .IngresarTecla "enter"
    
    GoTo Buclee2
End If

n = 6

BUCLE10:

'INTERMEDIO'
If .Leer(n, 2, 8) = "TOTAL OK" Then

    Sheets("TOTALREPORTES").Select 'NUMERO DE DOCUMENTOS'
    'Range("E1").End(xlDown).Offset(1, 0) = Trim(.Leer(n, 15, 7)) 'MONTO SOLES'
    'Range("F1").End(xlDown).Offset(1, 0) = Trim(.Leer(n, 28, 14)) 'NUMERO DOC SOLES'
    
    Cells(8, 5) = Trim(.Leer(n, 15, 7)) 'MONTO SOLES'
    Cells(8, 6) = Trim(.Leer(n, 28, 14)) 'NUMERO DOC SOLES'
    
    'SI MAS ESTA ABAJO'
    If .Leer(n + 1, 2, 1) = "=" Then
        Range("E1").End(xlDown).Offset(1, 0) = Trim(.Leer(n + 11, 15, 7)) 'MONTO DOLARES'
        Range("F1").End(xlDown).Offset(1, 0) = Trim(.Leer(n + 11, 28, 14)) 'NUMERO DOC DOLARES'
        GoTo Buclee3
    End If

'    Range("E1").End(xlDown).Offset(1, 0) = Trim(.Leer(n + 1, 15, 7)) 'MONTO DOLARES'
'    Range("F1").End(xlDown).Offset(1, 0) = Trim(.Leer(n + 1, 28, 14)) 'NUMERO DOC DOLARES'
    
    Cells(9, 5) = Trim(.Leer(n + 1, 15, 7)) 'MONTO DOLARES'
    Cells(9, 6) = Trim(.Leer(n + 1, 28, 14)) 'NUMERO DOC DOLARES'
Else
    n = n + 1
    GoTo BUCLE10
End If

'MAÑANA'
Buclee3:

If Left(fecha_3270, 2) = .Leer(8, 64, 2) And "- MAÑANA" = .Leer(9, 74, 8) Then
    .Escribir "m", 2, 15
    .IngresarTecla "f8"
    
Else
    .IngresarTecla "f3"
    .IngresarTecla "enter"
    
    GoTo Buclee3
End If

n = 6

bucle11:

'SOLES'
If .Leer(n, 2, 8) = "TOTAL OK" Then

    Sheets("TOTALREPORTES").Select 'NUMERO DE DOCUMENTOS'
'    Range("E1").End(xlDown).Offset(1, 0) = Trim(.Leer(n, 15, 7)) 'MONTO SOLES'
'    Range("F1").End(xlDown).Offset(1, 0) = Trim(.Leer(n, 28, 14)) 'NUMERO DOC SOLES'
    
    Cells(10, 5) = Trim(.Leer(n, 15, 7)) 'MONTO SOLES'
    Cells(10, 6) = Trim(.Leer(n, 28, 14)) 'NUMERO DOC SOLES'
    
    'SI ESTA MAS ABAJO'
    If .Leer(n + 1, 2, 1) = "=" Then
        Range("E1").End(xlDown).Offset(1, 0) = Trim(.Leer(n + 11, 15, 7)) 'MONTO DOLARES'
        Range("F1").End(xlDown).Offset(1, 0) = Trim(.Leer(n + 11, 28, 14)) 'NUMERO DOC DOLARES'
        GoTo Buclee4
    End If
    
'    Range("E1").End(xlDown).Offset(1, 0) = Trim(.Leer(n + 1, 15, 7)) 'MONTO DOLARES'
'    Range("F1").End(xlDown).Offset(1, 0) = Trim(.Leer(n + 1, 28, 14)) 'NUMERO DOC DOLARES'
    
    Cells(11, 5) = Trim(.Leer(n + 1, 15, 7)) 'MONTO SOLES'
    Cells(11, 6) = Trim(.Leer(n + 1, 28, 14)) 'NUMERO DOC SOLES'

Else
    n = n + 1
    GoTo bucle11
End If


'TARDE'
Buclee4:

If Left(fecha_3270, 2) = .Leer(8, 64, 2) And "- TARDE" = .Leer(9, 74, 7) Then
    .Escribir "m", 2, 15
    .IngresarTecla "f8"

    If .Leer(9, 19, 8) = "THE FIND" Then
        .Escribir "3", 11, 55
        .IngresarTecla "enter"
        
    End If
Else
    .IngresarTecla "f3"
    .IngresarTecla "enter"
    GoTo Buclee4
End If

n = 6

bucle12:

'SOLES'
If .Leer(n, 2, 8) = "TOTAL OK" Then

    Sheets("TOTALREPORTES").Select 'NUMERO DE DOCUMENTOS'
'    Range("E1").End(xlDown).Offset(1, 0) = Trim(.Leer(n, 15, 7)) 'MONTO SOLES'
'    Range("F1").End(xlDown).Offset(1, 0) = Trim(.Leer(n, 28, 14)) 'NUMERO DOC SOLES'
    
    Cells(12, 5) = Trim(.Leer(n, 15, 7)) 'MONTO SOLES'
    Cells(12, 6) = Trim(.Leer(n, 28, 14)) 'NUMERO DOC SOLES'
    
    'SI MAS ESTA ABAJO'
    If .Leer(n + 1, 2, 1) = "=" Then
        Range("E1").End(xlDown).Offset(1, 0) = .Leer(n + 10, 16, 5) 'NUMERO DOC DOLARES'
        Range("F1").End(xlDown).Offset(1, 0) = .Leer(n + 10, 30, 12) 'MONTO DOLARES'
        GoTo BucleRetro2
    End If
    
'    Range("E1").End(xlDown).Offset(1, 0) = Trim(.Leer(n + 1, 15, 7)) 'MONTO DOLARES'
'    Range("F1").End(xlDown).Offset(1, 0) = Trim(.Leer(n + 1, 28, 14)) 'NUMERO DOC DOLARES'
    
    Cells(13, 5) = Trim(.Leer(n + 1, 15, 7)) 'MONTO SOLES'
    Cells(13, 6) = Trim(.Leer(n + 1, 28, 14)) 'NUMERO DOC SOLES'
Else
    n = n + 1
    GoTo bucle12
End If

'RETROCEDER A GIVETO'
BucleRetro2:

If .Leer(27, 19, 8) = "G GIVETO" Then
Else
    .IngresarTecla "[pf3]"
    GoTo BucleRetro2
End If

.IngresarTecla "f3"

'RECIBIDAS TARJETAS'
'SALDOS RECIBIDAS'
.Escribir "0928", 9, 26
.Escribir fecha_3270, 10, 26
.Escribir fecha_3270, 10, 36
.Escribir "FEJP6520", 11, 26
.Escribir "Y", 13, 26
.IngresarTecla "[enter]"


V = 4

RepetirV2:

'REPETIR V'
If .Leer(V, 2, 1) = " " Then
    .Escribir "v", V, 2
    V = V + 1
    GoTo RepetirV2
Else

    .IngresarTecla "enter"
    
    
End If

'RECIBIDAS TARJETAS 2'

BucleT20:

If Left(fecha_3270, 2) = .Leer(7, 124, 2) And "- INTERMEDIO" = .Leer(9, 74, 12) Then
    .Escribir "m", 2, 15
    .IngresarTecla "f8"
    .IngresarTecla "enter"
    
    
Else
    .IngresarTecla "f3"
    .IngresarTecla "enter"
    
    
    GoTo BucleT20
End If

n = 6

bucleT30:

'INTERMEDIO'
If .Leer(n, 23, 12) = "REGISTROS OK" Then
    
    Sheets("TOTALREPORTES").Select 'NUMERO DE DOCUMENTOS'
'    Range("E1").End(xlDown).Offset(1, 0) = Trim(.Leer(n, 17, 5)) 'MONTO SOLES'
'    Range("F1").End(xlDown).Offset(1, 0) = .Leer(n, 48, 12) 'NUMERO DOC SOLES'
    
    Cells(14, 5) = Trim(.Leer(n, 17, 5)) 'MONTO SOLES'
    Cells(14, 6) = .Leer(n, 48, 12) 'NUMERO DOC SOLES'
    
    'SI ESTA MAS ABAJO'
    If .Leer(n + 1, 2, 1) = "=" Then
        Range("E1").End(xlDown).Offset(1, 0) = .Leer(n + 11, 17, 5) 'MONTO DOLARES'
        Range("F1").End(xlDown).Offset(1, 0) = .Leer(n + 11, 48, 12) 'NUMERO DOC DOLARES'
    End If
    
'    Range("E1").End(xlDown).Offset(1, 0) = .Leer(n + 1, 17, 5) 'MONTO DOLARES'
'    Range("F1").End(xlDown).Offset(1, 0) = .Leer(n + 1, 48, 12) 'NUMERO DOC DOLARES'
    
    Cells(15, 5) = .Leer(n + 1, 17, 5) 'MONTO SOLES'
    Cells(15, 6) = .Leer(n + 1, 48, 12) 'NUMERO DOC SOLES'
Else
    n = n + 1
    GoTo bucleT30
End If

'MAÑANA'
BucleT40:

If Left(fecha_3270, 2) = .Leer(7, 124, 2) And "- MAÑANA" = .Leer(9, 74, 8) Then
    .Escribir "m", 2, 15
    .IngresarTecla "f8"
    
    
Else
    .IngresarTecla "f3"
    .IngresarTecla "enter"
    
    
    GoTo BucleT40
End If

n = 6

bucleT50:

'SOLES'
If .Leer(n, 23, 12) = "REGISTROS OK" Then
    
    Sheets("TOTALREPORTES").Select 'NUMERO DE DOCUMENTOS'
'    Range("E1").End(xlDown).Offset(1, 0) = .Leer(n, 17, 5) 'MONTO SOLES'
'    Range("F1").End(xlDown).Offset(1, 0) = .Leer(n, 48, 12) 'NUMERO DOC SOLES'
    
    Cells(16, 5) = .Leer(n, 17, 5) 'MONTO SOLES'
    Cells(16, 6) = .Leer(n, 48, 12) 'NUMERO DOC SOLES'
    
    'SI ESTA MAS ABAJO'
    If .Leer(n + 1, 2, 1) = "=" Then
        Range("E1").End(xlDown).Offset(1, 0) = .Leer(n + 11, 15, 7) 'MONTO DOLARES'
        Range("F1").End(xlDown).Offset(1, 0) = .Leer(n + 11, 29, 13) 'NUMERO DOC DOLARES'
    End If
    
'    Range("E1").End(xlDown).Offset(1, 0) = .Leer(n + 1, 17, 5) 'MONTO DOLARES'
'    Range("F1").End(xlDown).Offset(1, 0) = .Leer(n + 1, 48, 12) 'NUMERO DOC DOLARES'
    
    Cells(17, 5) = .Leer(n + 1, 17, 5) 'MONTO SOLES'
    Cells(17, 6) = .Leer(n + 1, 48, 12) 'NUMERO DOC SOLES'
Else
    n = n + 1
    GoTo bucleT50
End If


'TARDE'
BucleT60:

If Left(fecha_3270, 2) = .Leer(7, 124, 2) And "- TARDE" = .Leer(9, 74, 7) Then
    .Escribir "m", 2, 15
    .IngresarTecla "f8"
    
Else
    .IngresarTecla "f3"
    .IngresarTecla "enter"
    
    
    GoTo BucleT60
End If

n = 6

bucleT70:

'SOLES'
If .Leer(n, 23, 12) = "REGISTROS OK" Then
    
    Sheets("TOTALREPORTES").Select 'NUMERO DE DOCUMENTOS'
'    Range("E1").End(xlDown).Offset(1, 0) = .Leer(n, 17, 5) 'MONTO SOLES'
'    Range("F1").End(xlDown).Offset(1, 0) = .Leer(n, 48, 12) 'NUMERO DOC SOLES'
    
    Cells(18, 5) = .Leer(n, 17, 5) 'MONTO SOLES'
    Cells(18, 6) = .Leer(n, 48, 12) 'NUMERO DOC SOLES'
    
    'SI ESTA MAS ABAJO'
    If .Leer(n + 1, 2, 1) = "=" Then
        Range("E1").End(xlDown).Offset(1, 0) = .Leer(n + 11, 17, 5) 'MONTO DOLARES'
        Range("F1").End(xlDown).Offset(1, 0) = .Leer(n + 11, 48, 12) 'NUMERO DOC DOLARES'
        GoTo TERMINE
    End If
    
'    Range("E1").End(xlDown).Offset(1, 0) = .Leer(n + 1, 17, 5) 'MONTO DOLARES'
'    Range("F1").End(xlDown).Offset(1, 0) = .Leer(n + 1, 48, 12) 'NUMERO DOC DOLARES'
    
    Cells(19, 5) = .Leer(n + 1, 17, 5) 'MONTO SOLES'
    Cells(19, 6) = .Leer(n + 1, 48, 12) 'NUMERO DOC SOLES'
Else
    n = n + 1
    GoTo bucleT70
TERMINE:

End If

.IngresarTecla "f3"
.IngresarTecla "f3"
.IngresarTecla "f3"



'SUMAR EN E'
Total = Cells(14, 5)
Total = Cells(8, 5) + Total
Cells(8, 5) = Total
 
Total = Cells(15, 5)
Total = Cells(9, 5) + Total
Cells(9, 5) = Total

Total = Cells(16, 5)
Total = Cells(10, 5) + Total
Cells(10, 5) = Total

Total = Cells(17, 5)
Total = Cells(11, 5) + Total
Cells(11, 5) = Total

Total = Cells(18, 5)
Total = Cells(12, 5) + Total
Cells(12, 5) = Total

Total = Cells(19, 5)
Total = Cells(13, 5) + Total
Cells(13, 5) = Total

'SUMAR EN F'
Total = Cells(14, 6)
Total = Cells(8, 6) + Total
Cells(8, 6) = Total
 
Total = Cells(15, 6)
Total = Cells(9, 6) + Total
Cells(9, 6) = Total

Total = Cells(16, 6)
Total = Cells(10, 6) + Total
Cells(10, 6) = Total

Total = Cells(17, 6)
Total = Cells(11, 6) + Total
Cells(11, 6) = Total

Total = Cells(18, 6)
Total = Cells(12, 6) + Total
Cells(12, 6) = Total

Total = Cells(19, 6)
Total = Cells(13, 6) + Total
Cells(13, 6) = Total

Range("E14:F19").Select
Application.CutCopyMode = False
Selection.ClearContents

Range("A1").Select

End Sub


