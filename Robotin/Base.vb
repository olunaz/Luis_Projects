Public Function convertirFecha(Valor As String)

    anho = CInt(Mid(Valor, 1, 4))
    mes = CInt(buscarMes(Mid(Valor, 6, 3)))
    convertirFecha = anho * 100 + mes
End Function

Public Function entreFechas(valorHoy As Integer, valorDesde As Integer, valorHasta As Integer)
    entreFechas = valorDesde <= valorHoy And valorHoy <= valorHasta
End Function

Public Function buscarCuenta(Valor As String)
    buscarCuenta = Mid(Valor, 14, 20)
End Function
Public Function buscarMes(Valor As String)

    Dim val As String
    val = UCase(Valor)

    If val = "ENE" Then
        buscarMes = 1
    End If
    If val = "FEB" Then
        buscarMes = 2
    End If
    If val = "MAR" Then
        buscarMes = 3
    End If
    If val = "ABR" Then
        buscarMes = 4
    End If
    If val = "MAY" Then
        buscarMes = 5
    End If
    If val = "JUN" Then
        buscarMes = 6
    End If
    If val = "JUL" Then
        buscarMes = 7
    End If
    If val = "AGO" Then
        buscarMes = 8
    End If
    If val = "SEP" Then
        buscarMes = 9
    End If
    If val = "OBT" Then
        buscarMes = 10
    End If
    If val = "OCT" Then
        buscarMes = 10
    End If
    If val = "NOV" Then
        buscarMes = 11
    End If
    If val = "DIC" Then
        buscarMes = 12
    End If
    
End Function



Private Sub CommandButton1_Click()
' descargar todas las bandejas borrador
  Dim autECLPSObj As Object
    Dim autECLConnList As Object
    Dim autECLOIAObj As Object
    Dim autECLSessionObj As Object
    Dim hoja As Worksheet
    Dim fila3270 As Long
    
    'Para conectar excel con 3270
    Set autECLPSObj = CreateObject("PCOMM.autECLPS")
    Set autECLConnList = CreateObject("PCOMM.autECLConnList")
    Set autECLOIAObj = CreateObject("PCOMM.autECLOIA")
    Set autECLSessionObj = CreateObject("PCOMM.autECLSession")
    
    sesion = "A"
    
    autECLSessionObj.SetConnectionByName (sesion)

    autECLSessionObj.autECLOIA.WaitForAppAvailable
    autECLSessionObj.autECLOIA.WaitForInputReady
    
    autECLOIAObj.SetConnectionByName (sesion)
    autECLConnList.Refresh
    
    
    autECLPSObj.SetConnectionByName (sesion)
    
    autECLSessionObj.autECLOIA.WaitForAppAvailable
    autECLSessionObj.autECLOIA.WaitForInputReady
    
    fila3270 = 5
    
    
                
                
    vacio = autECLPSObj.GetText(fila3270, 3, 5)
    Do Until (vacio = "=====" Or vacio = " PRIN")
        'numpaginas = autECLPSObj.GetText(fila3270, 47, 2)
        'numlineas = autECLPSObj.GetText(fila3270, 55, 3)
        
        
        numPaginas = CInt(autECLPSObj.GetText(fila3270, 47, 2))
        'numlineas = CInt(autECLPSObj.GetText(fila3270, 55, 3))
        Periodo = autECLPSObj.GetText(fila3270, 22, 8)


        autECLPSObj.SetText "V", fila3270, 2
        autECLSessionObj.autECLPS.SendKeys "[enter]"
        autECLSessionObj.autECLOIA.WaitForAppAvailable
        autECLSessionObj.autECLOIA.WaitForInputReady
        autECLSessionObj.autECLOIA.WaitForAppAvailable
        autECLSessionObj.autECLOIA.WaitForInputReady
            
        'For i = 1 To numpaginas
            'Cells(i, 1).Value = 100
'            vecesBajadas = numlineas / 12
'            For x = 1 To vecesBajadas
'
'                For j = 1 To 17
'                'Cells(i, 1).Value = 100
'                    texto = autECLPSObj.GetText(j + 6, 1, 80)
'                    If texto = "====== NEW PAGE ========                                                        " Then
'                        Exit For
'                    End If
'
'                     If texto = " ======= >>>>>>>>>>>>>> NO MORE DATA TO VIEW <<<<<<<<<<<<<<<<<<<<<<<<<<<<< =====" Then
'                        Exit For
'                    End If
'
'                    Set hoja = Sheets("borrador")
'                    numero = (x - 1) * 17 + j
'                    hoja.Cells((x - 1) * 17 + j, 1).Value = texto
'
'                    autECLSessionObj.autECLPS.SendKeys "[pf11]"
'                    autECLSessionObj.autECLOIA.WaitForAppAvailable
'                    autECLSessionObj.autECLOIA.WaitForInputReady
'                    autECLSessionObj.autECLOIA.WaitForAppAvailable
'                    autECLSessionObj.autECLOIA.WaitForInputReady
'                    autECLSessionObj.autECLOIA.WaitForAppAvailable
'                    autECLSessionObj.autECLOIA.WaitForInputReady
'                    autECLSessionObj.autECLOIA.WaitForAppAvailable
'                    autECLSessionObj.autECLOIA.WaitForInputReady
'
'
'                    texto = autECLPSObj.GetText(j + 6, 1, 80)
'                    Set hoja = Sheets("borrador")
'                    hoja.Cells((x - 1) * 17 + j, 1).Value = hoja.Cells((x - 1) * 17 + j, 1).Value & texto
'                 '   hoja.Cells((x - 1) * 12 + j, 1).Value = Mid(hoja.Cells((x - 1) * 12 + j, 1).Value, 2, 150)
'                    autECLSessionObj.autECLPS.SendKeys "[pf10]"
'                    autECLSessionObj.autECLOIA.WaitForAppAvailable
'                    autECLSessionObj.autECLOIA.WaitForInputReady
'                    autECLSessionObj.autECLOIA.WaitForAppAvailable
'                    autECLSessionObj.autECLOIA.WaitForInputReady
'                    autECLSessionObj.autECLOIA.WaitForAppAvailable
'                    autECLSessionObj.autECLOIA.WaitForInputReady
'                    autECLSessionObj.autECLOIA.WaitForAppAvailable
'                    autECLSessionObj.autECLOIA.WaitForInputReady
'                Next j
'                autECLSessionObj.autECLPS.SendKeys "[pf8]"
'                autECLSessionObj.autECLOIA.WaitForAppAvailable
'                autECLSessionObj.autECLOIA.WaitForInputReady
'                autECLSessionObj.autECLOIA.WaitForAppAvailable
'                autECLSessionObj.autECLOIA.WaitForInputReady
'                autECLSessionObj.autECLOIA.WaitForAppAvailable
'                autECLSessionObj.autECLOIA.WaitForInputReady
'                autECLSessionObj.autECLOIA.WaitForAppAvailable
'                autECLSessionObj.autECLOIA.WaitForInputReady
'            Next x
        'Next i
        autECLSessionObj.autECLPS.SendKeys "[pf3]"
        autECLSessionObj.autECLOIA.WaitForAppAvailable
        autECLSessionObj.autECLOIA.WaitForInputReady
        autECLSessionObj.autECLOIA.WaitForAppAvailable
        autECLSessionObj.autECLOIA.WaitForInputReady
        autECLSessionObj.autECLOIA.WaitForAppAvailable
        autECLSessionObj.autECLOIA.WaitForInputReady
        autECLSessionObj.autECLOIA.WaitForAppAvailable
        autECLSessionObj.autECLOIA.WaitForInputReady
        fila3270 = fila3270 + 1
       
        Set hoja = Sheets("borrador")
        'nombreCuenta = buscarCuenta(hoja.Cells(9, 1).Value)
        nombreArchivo = Periodo
        vacio = autECLPSObj.GetText(fila3270, 3, 5)
        
     
        
       
        Sheets("borrador").Select
        ActiveSheet.Range("A:A").Select
        
   
            With Selection.Font
                   .Name = "Courier New"
                   .Size = 11
                   .Strikethrough = False
                   .Superscript = False
                   .Subscript = False
                   .OutlineFont = False
                   .Shadow = False
                   .Underline = xlUnderlineStyleNone
                   .ThemeColor = xlThemeColorLight1
                   .TintAndShade = 0
                   .ThemeFont = xlThemeFontNone
            End With
       
                
           
            Sheets("borrador").Select
    
            
            Application.PrintCommunication = False
            
            With ActiveSheet.PageSetup
                .PrintTitleRows = ""
                .PrintTitleColumns = ""
            End With
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A:$N"
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.708661417322835)
                .RightMargin = Application.InchesToPoints(0.708661417322835)
                .TopMargin = Application.InchesToPoints(0.748031496062992)
                .BottomMargin = Application.InchesToPoints(0.748031496062992)
                .HeaderMargin = Application.InchesToPoints(0.31496062992126)
                .FooterMargin = Application.InchesToPoints(0.31496062992126)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintSheetEnd
                .PrintQuality = 600
                .CenterHorizontally = False
                .CenterVertically = False
                .Orientation = xlLandscape
                .Draft = False
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = numPaginas
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = True
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
            Application.PrintCommunication = True
            Sheets("borrador").Select
            ActiveSheet.Range("B19").Select
            
            

    
      
        Sheets("borrador").Select
        Sheets("borrador").Copy
        
      
        ChDir "D:\$mgutierr\7. IIJJ"
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            "D:\$mgutierr\7. IIJJ\" & nombreArchivo & ".pdf", Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
            False
        Application.DisplayAlerts = False
        ActiveWorkbook.Close
        Application.DisplayAlerts = True
        
        Sheets("borrador").Columns("A:A").Select
        Selection.ClearContents
        
        
        Sheets("borrador").Range("B8").Select


    Loop
    

    
    
End Sub
Public Function validarDatos(cuenta, inicio, fin) As Boolean

    validarDatos = Len(cuenta) = 20 And Mid(cuenta, 5, 1) = "-" And Mid(cuenta, 10, 1) = "-" And Len(inicio) = 8 And Mid(inicio, 5, 1) = "-" And Len(fin) = 8 And Mid(fin, 5, 1) = "-"

End Function
'busqueda por cuenta
Private Sub CommandButton2_Click()
        Dim autECLPSObj As Object
    Dim autECLConnList As Object
    Dim autECLOIAObj As Object
    Dim autECLSessionObj As Object
    Dim hoja As Worksheet
    
    
    'Para conectar excel con 3270
    Set autECLPSObj = CreateObject("PCOMM.autECLPS")
    Set autECLConnList = CreateObject("PCOMM.autECLConnList")
    Set autECLOIAObj = CreateObject("PCOMM.autECLOIA")
    Set autECLSessionObj = CreateObject("PCOMM.autECLSession")
    
    sesion = "A"
    
    autECLSessionObj.SetConnectionByName (sesion)

    autECLSessionObj.autECLOIA.WaitForAppAvailable
    autECLSessionObj.autECLOIA.WaitForInputReady
    
    autECLOIAObj.SetConnectionByName (sesion)
    autECLConnList.Refresh
    
    
    autECLPSObj.SetConnectionByName (sesion)
    
    autECLSessionObj.autECLOIA.WaitForAppAvailable
    autECLSessionObj.autECLOIA.WaitForInputReady
    
    fila3270 = 4
    
    cantCuentas = 1
    
    Sheets("Base").Range("B3").Select
    cantCuentas = Range(Selection, Selection.End(xlDown)).Count
    'verifica cuantas cuentas se van a buscar
    
    Range("B3").Select
    Dim arrayImpresiones(200) As Long
    Dim numPaginas As Long
    Dim numlineas As Long
  
                
                
    Dim arrayEncontradas(200) As Integer
    Dim arrayCuentas(200) As String
    Dim arrayFechaInicio(200) As String
    Dim arrayFechaFin(200) As String
    Dim arrayFechas(200) As String
    Dim textos(200) As String
    validadorFormato = True
    
    log_archivo ("*******************************")
    log_archivo ("Inicio robotin")
    log_archivo ("Revisar formato")
    'guarda toda la info de las cuentas en arreglos
    For i = 1 To cantCuentas
    
        If Sheets("Base").Cells(2 + i, 2).Value = "" Then
            cantCuentas = i - 1
            Exit For
        End If
        If validarDatos(Sheets("Base").Cells(2 + i, 2).Value, Sheets("Base").Cells(2 + i, 3).Value, Sheets("Base").Cells(2 + i, 4).Value) Then
           
        Else
        
            MsgBox ("Corregir formato " + Sheets("Base").Cells(2 + i, 2).Value + " " + Sheets("Base").Cells(2 + i, 3).Value + " " + Sheets("Base").Cells(2 + i, 4).Value)
            validadorFormato = False
            Exit For
        End If
        
        
    Next i
    

                
    If validadorFormato Then
    
    
        For i = 1 To cantCuentas
            arrayCuentas(i) = Sheets("Base").Cells(2 + i, 2).Value
            arrayFechaInicio(i) = convertirFecha(Sheets("Base").Cells(2 + i, 3).Value)
            arrayFechaFin(i) = convertirFecha(Sheets("Base").Cells(2 + i, 4).Value)
            arrayEncontradas(i) = 0
            arrayImpresiones(i) = 1
            'es la ultima linea que imprimio en la cuenta i
            Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = arrayCuentas(i)
            
        Next i
    'cuentaABuscar = Sheets("Base").Cells(3, 2)
    'fechaABuscarInicio = convertirFecha(Sheets("Base").Cells(3, 3).Value)
    'fechaABuscarFin = convertirFecha(Sheets("Base").Cells(3, 4).Value)
    'vecesEncontradas = 0
    'fila3270 = 4
        log_archivo ("buscando cuentas")
        vacio = autECLPSObj.GetText(fila3270, 3, 5)
        cantPeriodo = 0
        'lee todas las paginas del control D
        Do Until (vacio = "=====" Or vacio = "     ")
            'lee todas las lineas de una pagina en el control D hasta que encuentre el PRINT o fin de pagina
            Do Until (vacio = "=====" Or vacio = " PRIN" Or vacio = "     ")
               
                numPaginas = CInt(autECLPSObj.GetText(fila3270, 47, 2))
                'numlineas = CInt(autECLPSObj.GetText(fila3270, 53, 5))
                Var = autECLPSObj.GetText(fila3270, 22, 8)
                Periodo = autECLPSObj.GetText(fila3270, 22, 8)
                
                fechaEncontrada = False
                fechaEnRango = False
                For p = 1 To cantPeriodo
                    If arrayFechas(p) = Periodo Then
                      fechaEncontrada = True
                    End If
                Next p
                
                numPeriodo = convertirFecha(CStr(Periodo))
                
                For k = 1 To cantCuentas
                    If arrayFechaInicio(k) <= numPeriodo And numPeriodo <= arrayFechaFin(k) Then
                        fechaEnRango = True
                        Exit For
                    End If
                Next k
                
                
                If Not fechaEncontrada And fechaEnRango Then
                    log_archivo (Periodo)
                    cantPeriodo = cantPeriodo + 1
                    arrayFechas(cantPeriodo) = Periodo
                    numPeriodo = convertirFecha(CStr(Periodo))
        
                 
                    autECLPSObj.SetText "V", fila3270, 2
                    autECLSessionObj.autECLPS.SendKeys "[enter]"
                    autECLSessionObj.autECLOIA.WaitForAppAvailable
                    autECLSessionObj.autECLOIA.WaitForInputReady
                    autECLSessionObj.autECLOIA.WaitForAppAvailable
                    autECLSessionObj.autECLOIA.WaitForInputReady
        '            'MsgBox (periodo)
        ''
                        vecesBajadas = numlineas / 20
                        vecesBajadas = Int(vecesBajadas) + 1
                        pos = 0
                        encontrado = False
                        primeraVez = False
                        'entra a cada reporte
                        'For x = 0 To vecesBajadas
                        cantNuevasPaginas = 0
                        
                        For x = 0 To 10000
                            For j = 1 To 20
                                
                                texto = autECLPSObj.GetText(j + 6, 1, 80)
        
                                If texto = " ======= >>>>>>>>>>>>>> NO MORE DATA TO VIEW <<<<<<<<<<<<<<<<<<<<<<<<<<<<< =====" Then
                                    encontrado = False
                                    pos = 0
                                    cantNuevasPaginas = 0
                                    Exit For
                                End If
        
        
                                If encontrado And Mid(texto, 1, 1) = "0" Then
                                    encontrado = False
                                    arrayEncontradas(pos) = arrayEncontradas(pos) + cantNuevasPaginas
                                    cantNuevasPaginas = 0
                                    pos = 0
                                    
                                End If
        
                                If texto = "====== NEW PAGE ========                                                        " Then
                                    cantNuevasPaginas = cantNuevasPaginas + 1
                                End If
                                
                                If texto <> "====== NEW PAGE ========                                                        " Then
        
        
                                    If ((Not (encontrado)) And Mid(texto, 1, 1) = "0") Then
                                    'en esa linea hay cuenta
                                        cuenta = autECLPSObj.GetText(j + 6, 14, 20)
                                        For y = 1 To cantCuentas
                                           If arrayCuentas(y) = cuenta And (arrayFechaInicio(y) <= numPeriodo And numPeriodo <= arrayFechaFin(y)) Then
                                                encontrado = True
                                                pos = y
                                                arrayEncontradas(pos) = arrayEncontradas(pos) + 1
                                                'arrayImpresiones(i) = arrayImpresiones(i) + 8
                                                '1BANCO CONTINENTAL                                      MOVIMIENTOS POR CUENTA  .
                                                ' ------------------------------------------------------------------------------------------------------------------------------------                           .
                                                '  NUMER  COD    NRO TRAN OFI  USUARIO   FECHA   HORA   FECHA    FECHA        DESCRIPCION                     IMPORTE          SALDO                             .
                                                '   MOV   OPE    IPF ORIG OPE   OPERA    OPERA   OPER   VALOR    CONTA                                                                                           .
                                                ' ------------------------------------------------------------------------------------------------------------------------------------                           .
                                                '
                                                'copiar cabecera
                                                primeraVez = True
                                           End If
                                        Next y
                                    End If
        
                                    If encontrado Then
                                        
                                        Set hoja = Sheets(CStr(cuenta))
                                        'se guarda la cabecera de los reportes
                                        If primeraVez Then
                                            hoja.Cells(arrayImpresiones(pos) + 1, 1).Value = "1BANCO CONTINENTAL                                      MOVIMIENTOS POR CUENTA  "
                                            hoja.Cells(arrayImpresiones(pos) + 2, 1).Value = "                                                          A:O/MES: " & CStr(Periodo)
                                            hoja.Cells(arrayImpresiones(pos) + 3, 1).Value = " ------------------------------------------------------------------------------------------------------------------------------------                           "
                                            hoja.Cells(arrayImpresiones(pos) + 4, 1).Value = "  NUMER  COD    NRO TRAN OFI  USUARIO   FECHA   HORA   FECHA    FECHA        DESCRIPCION                     IMPORTE          SALDO                             "
                                            hoja.Cells(arrayImpresiones(pos) + 5, 1).Value = "   MOV   OPE    IPF ORIG OPE   OPERA    OPERA   OPER   VALOR    CONTA                                                                                           "
                                            hoja.Cells(arrayImpresiones(pos) + 6, 1).Value = " ------------------------------------------------------------------------------------------------------------------------------------                           "
                                            arrayImpresiones(pos) = arrayImpresiones(pos) + 8
                                            primeraVez = False
                                        End If
            
                                        texto = autECLPSObj.GetText(j + 6, 1, 132)
                                        'Valor = (arrayEncontradas(pos) - 1) * 17 + (x - 1) * 17 + j
                                        hoja.Cells(arrayImpresiones(pos), 1).Value = texto
        
                                     '   autECLSessionObj.autECLPS.SendKeys "[pf11]"
                                     '  autECLSessionObj.autECLOIA.WaitForAppAvailable
                                        ' autECLSessionObj.autECLOIA.WaitForInputReady
                                        '    autECLSessionObj.autECLOIA.WaitForAppAvailable
                                        '    autECLSessionObj.autECLOIA.WaitForInputReady
                                        '    autECLSessionObj.autECLOIA.WaitForAppAvailable
                                        '    autECLSessionObj.autECLOIA.WaitForInputReady
                                        '    autECLSessionObj.autECLOIA.WaitForAppAvailable
                                        '    autECLSessionObj.autECLOIA.WaitForInputReady
        
        
                                        'texto = autECLPSObj.GetText(j + 6, 1, 30)
                                        '    Set hoja = Sheets(CStr(cuenta))
                                        'hoja.Cells(arrayImpresiones(pos), 1).Value = hoja.Cells(arrayImpresiones(pos), 1).Value & texto
                                        '   hoja.Cells((x - 1) * 12 + j, 1).Value = Mid(hoja.Cells((x - 1) * 12 + j, 1).Value, 2, 150)
                                        '    autECLSessionObj.autECLPS.SendKeys "[pf10]"
                                        '    autECLSessionObj.autECLOIA.WaitForAppAvailable
                                        '    autECLSessionObj.autECLOIA.WaitForInputReady
                                        '    autECLSessionObj.autECLOIA.WaitForAppAvailable
                                        '    autECLSessionObj.autECLOIA.WaitForInputReady
                                        '    autECLSessionObj.autECLOIA.WaitForAppAvailable
                                        '    autECLSessionObj.autECLOIA.WaitForInputReady
                                        '    autECLSessionObj.autECLOIA.WaitForAppAvailable
                                        '    autECLSessionObj.autECLOIA.WaitForInputReady
        
                                        arrayImpresiones(pos) = arrayImpresiones(pos) + 1
        
                                    End If '
        
        
        
                                End If
        
        
                            Next j
                            
                            If texto = " ======= >>>>>>>>>>>>>> NO MORE DATA TO VIEW <<<<<<<<<<<<<<<<<<<<<<<<<<<<< =====" Then
                              '      encontrado = False
                              '      pos = 0
                                    Exit For
                            End If
        
                            autECLSessionObj.autECLPS.SendKeys "[pf8]"
                            autECLSessionObj.autECLOIA.WaitForAppAvailable
                            autECLSessionObj.autECLOIA.WaitForInputReady
                            autECLSessionObj.autECLOIA.WaitForAppAvailable
                            autECLSessionObj.autECLOIA.WaitForInputReady
                            autECLSessionObj.autECLOIA.WaitForAppAvailable
                            autECLSessionObj.autECLOIA.WaitForInputReady
                            autECLSessionObj.autECLOIA.WaitForAppAvailable
                            autECLSessionObj.autECLOIA.WaitForInputReady
        
        
                        Next x
                        'Next i
                    autECLSessionObj.autECLPS.SendKeys "[pf3]"
                    autECLSessionObj.autECLOIA.WaitForAppAvailable
                    autECLSessionObj.autECLOIA.WaitForInputReady
                    autECLSessionObj.autECLOIA.WaitForAppAvailable
                    autECLSessionObj.autECLOIA.WaitForInputReady
                    autECLSessionObj.autECLOIA.WaitForAppAvailable
                    autECLSessionObj.autECLOIA.WaitForInputReady
                    autECLSessionObj.autECLOIA.WaitForAppAvailable
                    autECLSessionObj.autECLOIA.WaitForInputReady
                    
                End If
    
                
                
                fila3270 = fila3270 + 1
    
    
                vacio = autECLPSObj.GetText(fila3270, 3, 5)
    
        
            Loop
            
            autECLSessionObj.autECLPS.SendKeys "[pf8]"
            autECLSessionObj.autECLOIA.WaitForAppAvailable
            autECLSessionObj.autECLOIA.WaitForInputReady
            autECLSessionObj.autECLOIA.WaitForAppAvailable
            autECLSessionObj.autECLOIA.WaitForInputReady
            autECLSessionObj.autECLOIA.WaitForAppAvailable
            autECLSessionObj.autECLOIA.WaitForInputReady
            autECLSessionObj.autECLOIA.WaitForAppAvailable
            autECLSessionObj.autECLOIA.WaitForInputReady
            
            If vacio <> "=====" Or vacio <> " PRINT" Then
                fila3270 = 5
                vacio = autECLPSObj.GetText(fila3270, 3, 5)
            End If
            
        Loop
        'impresion
        
         For i = 1 To cantCuentas
        'todas las cuentas se guardar como pdf
            cuenta = Sheets("Base").Cells(2 + i, 2)
            Sheets("Base").Cells(2 + i, 1) = arrayEncontradas(i)
        Next i
        
        Sheets("Base").Cells(1, 1) = cantCuentas
        MsgBox ("Finalizado lectura")
        
    End If
    log_archivo ("Fin robotin")
End Sub


Public Function BuscarHoja(nombreHoja As String) As Boolean
 
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = nombreHoja Then
            BuscarHoja = True
            Exit Function
        End If
    Next
     
    BuscarHoja = False
 
End Function

Sub log_archivo(val As String)
    Open "D:\Pedidos Informes Judiciales\LogRobotin.txt" For Append As #1
    Print #1, Str(Now) + ": " + val
    Close #1
End Sub






