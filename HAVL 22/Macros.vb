Const lenusuario = 7
Const reporte = "HAVL22*"
Const fin_listadereportes = "B O T T O M    O F    L I S T"
Const fin_reportectrld = "NO MORE DATA TO VIEW"
Const glosa_desconexionctrld = "CTM013I IOA SESSION ENDED"


Sub Principal()

Dim usuarios() As String
    nombmacro = ActiveWorkbook.Name
    pathmacro = ActiveWorkbook.Path
    rutaarch1 = Trim(Workbooks(nombmacro).Sheets(1).Cells(12, "K"))
    rutaarch2 = Trim(Workbooks(nombmacro).Sheets(1).Cells(13, "K"))
    rutaarch3 = Trim(Workbooks(nombmacro).Sheets(1).Cells(14, "K"))

IniciarFormulario:
Inicio.Show
user = Inicio.tRegistro.Value
If UCase(Left(user, 1)) <> "P" Or Len(user) <> lenusuario Then
MsgBox Errores(1)
GoTo IniciarFormulario
End If

pass = Inicio.tContra.Value
If Len(pass) <> 8 Then
MsgBox Errores(2)
GoTo IniciarFormulario
End If

ObtenerSesion:
sesion = UCase(InputBox("¿En qué sesión desea trabajar? (A o B)"))
If sesion <> "A" And sesion <> "B" And sesion <> vbNullString Then
MsgBox Errores(3)
GoTo ObtenerSesion
End If
If sesion = vbNullString Then
Sheets(1).Select
End
End If

Set Con = New Conexion3270
Con.init sesion

Inicio3720:
If Con.Leer(17, 15, 1) <> "A" Or Con.Leer(18, 15, 1) <> "B" Or Con.Leer(19, 15, 1) <> "C" Then
MsgBox Errores(4)
GoTo Inicio3720
End If

ObtenerFecha:
fecha = InputBox("¿Qué fecha desea consultar? Formato: DDMMAA")
If Not (IsNumeric(fecha)) Or Len(fecha) <> 6 Then
MsgBox Errores(5)
GoTo ObtenerFecha
End If

Call aPrincipalControlD(sesion, fecha, user, pass, rutaarch1)

CerrarCtrlD:


Con.IngresarTecla "pausa"

Sheets(1).Select
MsgBox Cells(10, "K") & ", el reporte se generó correctamente."

End Sub

Sub aPrincipalControlD(sesion, fecha, user, pass, ruta)

fechajaponesa = "20" & Right(fecha, 2) & Mid(fecha, 3, 2) & Left(fecha, 2)
nombmacro = ActiveWorkbook.Name
pathmacro = ActiveWorkbook.Path
ChDir (pathmacro)

fechacslash = CStr(Left(fecha, 2) & "/" & Mid(fecha, 3, 2) & "/" & Right(fecha, 2))
Set Con = New Conexion3270
Con.init sesion

Con.Escribir "D", 24, 24
Con.IngresarTecla "enter"

    Do Until Con.Leer(9, 6, 4) = "USER"
        Application.Wait (Now + TimeValue("00:00:01"))
    Loop

Con.Escribir user, 9, 39
Con.Escribir pass, 11, 39
Con.IngresarTecla "enter"
    
    If Con.Leer(1, 2, 7) = "CTM790E" Then
        MsgBox Errores(6)
        Con.IngresarTecla "f3"
        Con.IngresarTecla "f3"
    End
    End If
    
    Do Until Con.Leer(1, 23, 11) = "CONTROL-D/V"
        Application.Wait (Now + TimeValue("00:00:01"))
    Loop
    
Call bBorrarEntryPanel(sesion)

Con.Escribir reporte, 7, 26
Con.Escribir "0805", 9, 26
Con.Escribir fecha, 10, 26
Con.Escribir fecha, 10, 36
Con.IngresarTecla "enter"

    If Con.Leer(10, 22, 6) = "PLEASE" Then
        Con.Escribir "3", 10, 58
        Con.IngresarTecla "enter"
    End If
    
    If Con.Leer(1, 2, 7) = "IOAE53E" Then
        MsgBox Errores(7)
        Con.IngresarTecla "f3"
        Con.IngresarTecla "f3"
        Sheets(1).Select
        End
    End If
    
    Application.Wait (Now + TimeValue("00:00:03"))
    
    If Con.Leer(10, 22, 6) = "PLEASE" Then
        Con.Escribir "3", 10, 58
        Con.IngresarTecla "enter"
    End If

fbusca = 4
While Con.Leer(fbusca, 27, 29) <> fin_listadereportes
    If _
    Con.Leer(fbusca, 6, 2) = "05" And Con.Leer(fbusca, 34, 8) = fechacslash Then
    f923 = fbusca
    End If
fbusca = fbusca + 1
Wend

Sheets(2).Select

'***************************Oficina 923*******************************
If f923 = "" Then
    MsgBox ("No se encontró reporte de oficina 923")
    GoTo NoHay923
End If

Con.Escribir "X", f923, 2
Con.IngresarTecla "enter"

Application.Wait (Now + TimeValue("00:00:03"))

Con.Escribir "S", 12, 13
Con.IngresarTecla "enter"

Application.Wait (Now + TimeValue("00:00:03"))

Con.Escribir "F 0814-PEN", 2, 15
Con.IngresarTecla "enter"

Application.Wait (Now + TimeValue("00:00:03"))

Con.Escribir "V", 5, 2
Con.IngresarTecla "enter"

Call cExtraccionControlD(sesion, fecha)

perohay923 = True
Con.IngresarTecla "f3"
Con.IngresarTecla "f3"
Con.IngresarTecla "f3"
Con.IngresarTecla "f3"
Con.IngresarTecla "pausa"
'*****************************Oficina 924******************************
NoHay923:
Habia923:

ult = Cells(Rows.Count, "A").End(xlUp).Row

Sheets(2).Copy
Select Case ruta
Case "": rutafinal = pathmacro
Case Else: rutafinal = ruta
End Select
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs Filename:= _
rutafinal & "\REPORTE MOVIMIENTO DIARIO DE CUENTAS " & fechajaponesa & "CENTRO 0814" & " (COPIA EVERIS).xlsx"
Application.DisplayAlerts = True
ActiveWorkbook.Close

Workbooks(nombmacro).Activate

If 0 = 1 Then
NoHayNada:
MsgBox ("No se encontró ningún reporte.")
Sheets(1).Select
End
End If

End Sub

Sub bBorrarEntryPanel(sesion)

Set Con = New Conexion3270
Con.init sesion
fborrar = Array(5, 7, 9, 10, 10, 11)
cborrar = Array(19, 26, 26, 26, 36, 26)

For i = 0 To UBound(fborrar)
Con.ColocarCursor fborrar(i), cborrar(i)
Con.IngresarTecla "fin"
Next

Con.Escribir "N", 13, 26

End Sub

Sub cExtraccionControlD(sesion, fecha)
    Dim Contenido1()
    
    ReDim Contenido1(0 To 60, 0 To 15)
    
    c1 = Array(2, 9, 14, 19, 27, 35, 48, 69, 79, 98, 117, 119, 126)
    l1 = Array(5, 4, 4, 7, 7, 9, 20, 9, 19, 19, 1, 5, 3)
    
    e1 = Split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O", ",")
    
    
    Select Case Cells(2, "A")
        Case "":
            fexcel1 = 2
        Case Else:
            fexcel1 = Cells(Rows.Count, "A").End(xlUp).Row + 1
    End Select
    
    filaFinal3270 = 26
    num = 0
    Set Con = New Conexion3270
    With Con
        .init sesion
        
        For fila3270 = 7 To filaFinal3270
            If .Leer(fila3270, 8, 8) = "NEW PAGE" Then
                'Imprimir
                For farray = 0 To num - 1
                    If Contenido1(farray, 0) <> "" Then
                        For i = 0 To 14
                            Cells(fexcel1, e1(i)) = Trim(Contenido1(farray, i))
                        Next
                        fexcel1 = fexcel1 + 1
                    End If
                    If farray = num - 1 Then
                        .ColocarCursor fila3270, 2
                        .IngresarTecla "f8"
                        fila3270 = 6
                        Erase Contenido1
                        ReDim Contenido1(0 To 60, 0 To 15)
                        num = 0
                    End If
                Next
            End If
            
            If .Leer(fila3270, 25, 20) = fin_reportectrld Then
                For farray = 0 To num - 1
                    If Contenido1(farray, 0) <> "" Then
                        For i = 0 To 14
                            Cells(fexcel1, e1(i)) = Trim(Contenido1(farray, i))
                        Next
                        fexcel1 = fexcel1 + 1
                    End If
                    If farray = num - 1 Then
                        Erase Contenido1
                    End If
                Next
                'Si estamos en la última hoja, la extracción de data será hasta la fila final
                'caso trario, en toda la hoja
                Exit For
            End If
            
            If "FECHA TABLE" = Trim(.Leer(fila3270, 2, 14)) Then
                fechatable = Trim(.Leer(fila3270, 18, 10))
            End If
            If "FECHA PROCESO" = Trim(.Leer(fila3270, 2, 13)) Then
                fechaproceso = Trim(.Leer(fila3270, 18, 10))
            End If
            
            If .Leer(fila3270, 4, 3) = "-02" Then
                Contenido1(num, 0) = fechatable
                Contenido1(num, 1) = fechaproceso
                For i = 0 To 12
                Contenido1(num, i + 2) = Trim(.Leer(fila3270, c1(i), l1(i)))
                Next
                num = num + 1
            End If
            
            If fila3270 = filaFinal3270 Then
                .ColocarCursor fila3270, 2
                .IngresarTecla "f8"
                fila3270 = 6
            End If
            
        Next
    End With

End Sub

