Sub GetDataVisa()

' (1) Shows the msoFileDialogFilePicker dialog box.
' (2) Checks if the user picked a file.
' (3) Stores the path to the selected file in a string type variable.

    Dim strFilePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        ' show the file picker dialog box
        If .Show <> 0 Then
            strFilePath = .SelectedItems(1)
        
            ' *********************
            ' put your code in here
            
            Dim wb2 As Workbook
            Set wb2 = Workbooks.Open(strFilePath)
            
            Dim wb As Workbook
            Set wb = ThisWorkbook
            
            wb.Activate
            
            'HOJA 1
        
            Set ws1 = wb.Sheets.Add
            ws1.Name = "MPLS0E13_Inf_Cuentas"
                
            ws1.Range("A1:E1000").Value = wb2.Worksheets("MPLS0E13_Inf_Cuentas").Range("A1:E1000").Value
            
            'HOJA 2
            
            Set ws2 = wb.Sheets.Add
            ws2.Name = "MPLS0C13_Inf_Cuentas"
            
            ws2.Range("A1:E1000").Value = wb2.Worksheets("MPLS0C13_Inf_Cuentas").Range("A1:E1000").Value
           
           'HOJA 3
           
            Set ws3 = wb.Sheets.Add
            ws3.Name = "MPLS0E13_Volumen_Ventas"
            
            ws3.Range("A1:E1000").Value = wb2.Worksheets("MPLS0E13_Volumen_Ventas").Range("A1:E1000").Value
            
            'HOJA 4
           
            Set ws4 = wb.Sheets.Add
            ws4.Name = "MPLS0C13_Volumen_Ventas"
            
            ws4.Range("A1:E1000").Value = wb2.Worksheets("MPLS0C13_Volumen_Ventas").Range("A1:E1000").Value
            
            wb2.Close
            
           'BORRADO DE VALORES
            
            Call borrado
            
            'COPIADO DE VALORES DE BINES
            
            Call modifica_Values
            
            'Exportar Texto
            
            Call exportar_Texto
            
                  
            'ELIMINACION DE HOJAS
            
            Call EliminarHojas
        
            'COPIADO DE INFO
            
            Call copiar_Formato
            
                   
            
                     
            
            ' *********************
            
                   
            
                     
            
            ' *********************
            
            ' Example: print the path of the selected file to the immediate window
            'Debug.Print strFilePath ' remove in production
        End If
    End With
End Sub



Sub EliminarHojas()


 Dim wb As Workbook
 Set wb = ThisWorkbook
    
  wb.Activate
  
  Set ws1 = wb.Sheets("MPLS0E13_Inf_Cuentas")
  Set ws2 = wb.Sheets("MPLS0C13_Inf_Cuentas")
  Set ws3 = wb.Sheets("MPLS0E13_Volumen_Ventas")
  Set ws4 = wb.Sheets("MPLS0C13_Volumen_Ventas")
'  Set ws_txt = wb.Sheets("TEXTO")
  'Set ws5 = wb.Sheets("Data_Total")
    
      
    
  Application.DisplayAlerts = False 'switching off the alert button
  
  ws1.Delete
  ws2.Delete
  ws3.Delete
  ws4.Delete
'  ws_txt.Delete
  'ws5.Delete
  
  Application.DisplayAlerts = True 'switching on the alert button

End Sub


Sub modifica_Values()

  Dim wb As Workbook
  Set wb = ThisWorkbook
    
  wb.Activate
    
  Set ws1 = wb.Sheets("MPLS0E13_Inf_Cuentas")
  Set ws2 = wb.Sheets("MPLS0C13_Inf_Cuentas")
  Set ws3 = wb.Sheets("MPLS0E13_Volumen_Ventas")
  Set ws4 = wb.Sheets("MPLS0C13_Volumen_Ventas")
         
  Dim inifila As Integer, inicolumna As Integer, l As Integer
  Set ws_visaclasica = wb.Sheets("VISA_CCLS-TxT")
  Set ws_visagold = wb.Sheets("VISA_GOLD-TxT")
  Set ws_visacbsn = wb.Sheets("VISA_CBSN-TxT")
  Set ws_visacorp = wb.Sheets("VISA_CORP-TxT")
  Set ws_visaplat = wb.Sheets("VISA_PLAT-TxT")
  Set ws_visasig = wb.Sheets("VISA_SIG-TxT")
  Set ws_visacta = wb.Sheets("VISA_CTA-TxT")
  Set ws_visainf = wb.Sheets("VISA_INF-TxT")
  Set ws_visadebemp = wb.Sheets("VISA_DBSN-TxT")
  Set ws_visauclsd = wb.Sheets("VISA_UCLSD-TxT")
  Set ws_visagift = wb.Sheets("VISA_GIFT-TxT")

  
 'VISA CLASICA
 'Volumen de Ventas
 inicolumna = 4
 inifila = 1
 l = 1
  Do While (l < 1000)
   If (ws4.Cells(inifila, inicolumna).Value) = "VISACLASICA" Then
     Call CopiadoValoresVol(ws_visaclasica, inifila, ws4)
     Exit Do
   End If
   inifila = inifila + 1
   l = l + 1
 Loop

 'Inf Cuentas
 inicolumna = 4
 inifila = 1
 l = 1
 Do While (l < 1000)
   If (ws2.Cells(inifila, inicolumna).Value) = "VISACLASICA" Then
     Call CopiadoValoresInf(ws_visaclasica, inifila, ws2)
     Exit Do
    End If
   inifila = inifila + 1
   l = l + 1
 Loop


 'VISA ORO
 'Volumen de Ventas
 inicolumna = 4
 inifila = 1
 l = 1
 Do While (l < 1000)
   If (ws4.Cells(inifila, inicolumna).Value) = "VISAORO" Then
     Call CopiadoValoresVol(ws_visagold, inifila, ws4)
     Exit Do
    End If
   inifila = inifila + 1
   l = l + 1
  Loop

 'Inf Cuentas
 inicolumna = 4
 inifila = 1
 l = 1
 
 Do While (l < 1000)
  If (ws2.Cells(inifila, inicolumna).Value) = "VISAORO" Then
   Call CopiadoValoresInf(ws_visagold, inifila, ws2)
  End If
  inifila = inifila + 1
  l = l + 1
 Loop

'VISA CAPTA TRABAJO
 'Volumen de Ventas
 inicolumna = 4
 inifila = 1
 l = 1
 Do While (l < 1000)
  If (ws4.Cells(inifila, inicolumna).Value) = "VISACAPTRABAJO" Then
    Call CopiadoValoresVol(ws_visacbsn, inifila, ws4)
    Exit Do
  End If
  inifila = inifila + 1
  l = l + 1
 Loop

 'Inf Cuentas
 inicolumna = 4
 inifila = 1
 l = 1
 
 Do While (l < 1000)
  If (ws2.Cells(inifila, inicolumna).Value) = "VISACAPTRABAJO" Then
    Call CopiadoValoresInf(ws_visacbsn, inifila, ws2)
  End If
  inifila = inifila + 1
  l = l + 1
  Loop

'VISA CORPORATE ORO
 'Volumen de Ventas
 inicolumna = 4
 inifila = 1
 l = 1




Do While (l < 1000)
  If (ws4.Cells(inifila, inicolumna).Value) = "VISACORPORATEORO" Then
    Call CopiadoValoresVol(ws_visacorp, inifila, ws4)
    Exit Do
  End If
  inifila = inifila + 1
  l = l + 1
Loop
 'Inf Cuentas
 inicolumna = 4
 inifila = 1
 l = 1
 Do While (l < 1000)
  If (ws2.Cells(inifila, inicolumna).Value) = "VISACORPORATEORO" Then
   Call CopiadoValoresInf(ws_visacorp, inifila, ws2)
  End If
  inifila = inifila + 1
  l = l + 1
Loop


'VISA PLATINUM
 'Volumen de Ventas
 inicolumna = 4
 inifila = 1
 l = 1
Do While (l < 1000)
  If (ws4.Cells(inifila, inicolumna).Value) = "VISAPLATINUM" Then
    Call CopiadoValoresVol(ws_visaplat, inifila, ws4)
    Exit Do
  End If
  inifila = inifila + 1
  l = l + 1
Loop
 'Inf Cuentas
 inicolumna = 4
 inifila = 1
 l = 1
 
Do While (l < 1000)
  If (ws2.Cells(inifila, inicolumna).Value) = "VISAPLATINUM" Then
   Call CopiadoValoresInf(ws_visaplat, inifila, ws2)
  End If
  inifila = inifila + 1
  l = l + 1
 Loop
 'VISA SIGNATURE
 'Volumen de Ventas
 inicolumna = 4
 inifila = 1
 l = 1
Do While (l < 1000)
  If (ws4.Cells(inifila, inicolumna).Value) = "VISASIGNATURE" Then
    Call CopiadoValoresVol(ws_visasig, inifila, ws4)
    Exit Do
  End If
  inifila = inifila + 1
  l = l + 1
Loop
 'Inf Cuentas
 inicolumna = 4
 inifila = 1
 l = 1
 
Do While (l < 1000)
  If (ws2.Cells(inifila, inicolumna).Value) = "VISASIGNATURE" Then
   Call CopiadoValoresInf(ws_visasig, inifila, ws2)
  End If
  inifila = inifila + 1
  l = l + 1
Loop

'VISA CUENTA VIAJES
 'Volumen de Ventas
 inicolumna = 4
 inifila = 1
 l = 1
Do While (l < 1000)
  If (ws4.Cells(inifila, inicolumna).Value) = "VISACUENTAVIAJES" Then
    Call CopiadoValoresVol(ws_visacta, inifila, ws4)
    Exit Do
  End If
  inifila = inifila + 1
  l = l + 1
Loop
 'Inf Cuentas
 inicolumna = 4
 inifila = 1
 l = 1
 
Do While (l < 1000)
  If (ws2.Cells(inifila, inicolumna).Value) = "VISACUENTAVIAJES" Then
   Call CopiadoValoresInf(ws_visacta, inifila, ws2)
  End If
  inifila = inifila + 1
  l = l + 1
Loop


 'VISA BBVASE
 'Volumen de Ventas
 inicolumna = 4
 inifila = 1
 l = 1
 
 Do While (l < 1000)
  If (ws4.Cells(inifila, inicolumna).Value) = "BBVASE" Then
    Call CopiadoValoresVol(ws_visainf, inifila, ws4)
    Exit Do
  End If
  inifila = inifila + 1
  l = l + 1
 Loop

 'Inf Cuentas
 inicolumna = 4
 inifila = 1
 l = 1
 
Do While (l < 1000)
  If (ws2.Cells(inifila, inicolumna).Value) = "BBVASE" Then
   Call CopiadoValoresInf(ws_visainf, inifila, ws2)
  End If
  inifila = inifila + 1
  l = l + 1
Loop

'TARJETAS DE DEBITO

 'VISA DEBIT BUSINESS
 'Volumen de Ventas
 inicolumna = 4
 inifila = 1
 l = 1
 
 Do While (l < 1000)
  If (ws4.Cells(inifila, inicolumna).Value) = "VISA_DEBITO_EMP" Then
    Call CopiadoValoresVol(ws_visadebemp, inifila, ws4)
    Exit Do
  End If
  inifila = inifila + 1
  l = l + 1
 Loop

 'Inf Cuentas
 inicolumna = 4
 inifila = 1
 l = 1
 
 Do While (l < 1000)
  If (ws2.Cells(inifila, inicolumna).Value) = "VISA_DEBITO_EMP" Then
   Call CopiadoValoresInf(ws_visadebemp, inifila, ws2)
  End If
  inifila = inifila + 1
  l = l + 1
  Loop

  'VISA DEBITO CLASSIC
  'VISA UNEMBOSSED VISA CLASIIC DEBIT
 'Volumen de Ventas
 inicolumna = 4
 inifila = 1
 l = 1
 
 Do While (l < 1000)
  If (ws4.Cells(inifila, inicolumna).Value) = "VISADEBITO" Then
    Call CopiadoValoresVol(ws_visauclsd, inifila, ws4)
    Exit Do
  End If
  inifila = inifila + 1
  l = l + 1
 Loop

 'Inf Cuentas
 inicolumna = 4
 inifila = 1
 l = 1
 
Do While (l < 1000)
  If (ws2.Cells(inifila, inicolumna).Value) = "VISADEBITO" Then
   Call CopiadoValoresInf(ws_visauclsd, inifila, ws2)
  End If
  inifila = inifila + 1
  l = l + 1
Loop

'Tarjetas Regalo
'CONSUMER GIFT

 'VISA CONSUMER GIT
 'Volumen de Ventas
 inicolumna = 4
 inifila = 1
 l = 1
 
 Do While (l < 1000)
  If (ws4.Cells(inifila, inicolumna).Value) = "VISAREGALO" Then
    Call CopiadoValoresVol(ws_visagift, inifila, ws4)
    Exit Do
  End If
  inifila = inifila + 1
  l = l + 1
 Loop

 'Inf Cuentas
 inicolumna = 4
 inifila = 1
 l = 1
 
Do While (l < 1000)
  If (ws2.Cells(inifila, inicolumna).Value) = "VISAREGALO" Then
   Call CopiadoValoresInf(ws_visagift, inifila, ws2)
  End If
  inifila = inifila + 1
  l = l + 1
Loop

End Sub

 
Sub exportar_Texto()

  Dim wb As Workbook
  Set wb = ThisWorkbook

  wb.Activate

  Set ws_visaclasica = wb.Sheets("VISA_CCLS-TxT")
  Set ws_visagold = wb.Sheets("VISA_GOLD-TxT")
  Set ws_visacbsn = wb.Sheets("VISA_CBSN-TxT")
  Set ws_visacorp = wb.Sheets("VISA_CORP-TxT")
  Set ws_visaplat = wb.Sheets("VISA_PLAT-TxT")
  Set ws_visasig = wb.Sheets("VISA_SIG-TxT")
  Set ws_visacta = wb.Sheets("VISA_CTA-TxT")
  Set ws_visainf = wb.Sheets("VISA_INF-TxT")
  Set ws_visadebemp = wb.Sheets("VISA_DBSN-TxT")
  Set ws_visauclsd = wb.Sheets("VISA_UCLSD-TxT")
  Set ws_visagift = wb.Sheets("VISA_GIFT-TxT")
  Set ws_member = wb.Sheets("Menber Information")
    

    Set ws_txt = wb.Sheets.Add
    ws_txt.Name = "TEXTO"


    Dim k As Integer
    Dim r As Integer

    'MEMBER INFORMATION

    r = 6
    For k = 0 To 18

        ws_txt.Cells(1, k + 1).Value = ws_member.Cells(k + r, 2).Value


    Next k

    'BINES

    r = 6
    For k = 0 To 111

    ws_txt.Cells(2, k + 1).Value = ws_visaclasica.Cells(k + r, 2).Value
    ws_txt.Cells(3, k + 1).Value = ws_visagold.Cells(k + r, 2).Value
    ws_txt.Cells(4, k + 1).Value = ws_visacbsn.Cells(k + r, 2).Value
    ws_txt.Cells(5, k + 1).Value = ws_visacorp.Cells(k + r, 2).Value
    ws_txt.Cells(6, k + 1).Value = ws_visaplat.Cells(k + r, 2).Value
    ws_txt.Cells(7, k + 1).Value = ws_visasig.Cells(k + r, 2).Value
    ws_txt.Cells(8, k + 1).Value = ws_visacta.Cells(k + r, 2).Value
    ws_txt.Cells(9, k + 1).Value = ws_visainf.Cells(k + r, 2).Value
    ws_txt.Cells(10, k + 1).Value = ws_visadebemp.Cells(k + r, 2).Value
    ws_txt.Cells(11, k + 1).Value = ws_visauclsd.Cells(k + r, 2).Value
    ws_txt.Cells(12, k + 1).Value = ws_visagift.Cells(k + r, 2).Value
    

    Next k

    'CODIGO PARA EXPORTAR A TXT

    'Declaring the variables
    Dim FileName, sLine, Deliminator As String
    Dim LastCol, LastRow, FileNumber As Integer

    'Excel Location and File Name
    FileName = ThisWorkbook.Path & "\VolumenVentasVisa.txt"

    'Field Separator
    Deliminator = vbTab

    'Identifying the Last Cell
    LastCol = ws_txt.Cells.SpecialCells(xlCellTypeLastCell).Column
    LastRow = ws_txt.Cells.SpecialCells(xlCellTypeLastCell).Row
    FileNumber = FreeFile

    'Creating or Overwrighting a text file
    Open FileName For Output As FileNumber

    'Reading the data from Excel using For Loop
    For i = 1 To LastRow
    For j = 1 To LastCol

    'Removing Deliminator if it is wrighting the last column
    If j = LastCol Then
    sLine = sLine & Cells(i, j).Value
    Else
    sLine = sLine & Cells(i, j).Value & Deliminator
    End If
    Next j

    'Wrighting data into text file
    Print #FileNumber, sLine
    sLine = ""
    Next i

    'Closing the Text File
    Close #FileNumber

    'Generating message to display
    'MsgBox "El archivo Texto se generó correctamente"


End Sub


Sub copiar_Formato()

  Dim wb As Workbook
  Set wb = ThisWorkbook

  wb.Activate
         
  Set ws_visaclasica = wb.Sheets("VISA_CCLS-TxT")
  Set ws_visaclasicaN = wb.Sheets("VISA_CCLS")
  Set ws_visagold = wb.Sheets("VISA_GOLD-TxT")
  Set ws_visagoldN = wb.Sheets("VISA_GOLD")
  Set ws_visacbsn = wb.Sheets("VISA_CBSN-TxT")
  Set ws_visacbsnN = wb.Sheets("VISA_CBSN")
  Set ws_visacorp = wb.Sheets("VISA_CORP-TxT")
  Set ws_visacorpN = wb.Sheets("VISA_CORP")
  Set ws_visaplat = wb.Sheets("VISA_PLAT-TxT")
  Set ws_visaplatN = wb.Sheets("VISA_PLAT")
  Set ws_visasig = wb.Sheets("VISA_SIG-TxT")
  Set ws_visasigN = wb.Sheets("VISA_SIG")
  Set ws_visacta = wb.Sheets("VISA_CTA-TxT")
  Set ws_visactaN = wb.Sheets("VISA_CTA")
  Set ws_visainf = wb.Sheets("VISA_INF-TxT")
  Set ws_visainfN = wb.Sheets("VISA_INF")
  Set ws_visadebemp = wb.Sheets("VISA_DBSN-TxT")
  Set ws_visadebempN = wb.Sheets("VISA_DBSN")
  Set ws_visauclsd = wb.Sheets("VISA_UCLSD-TxT")
  Set ws_visauclsdN = wb.Sheets("VISA_UCLSD")
  Set ws_visagift = wb.Sheets("VISA_GIFT-TxT")
  Set ws_visagiftN = wb.Sheets("VISA_GIFT")

 Call Formatear_Valores(ws_visaclasica, ws_visaclasicaN)
 Call Formatear_Valores(ws_visagold, ws_visagoldN)
 Call Formatear_Valores(ws_visacbsn, ws_visacbsnN)
 Call Formatear_Valores(ws_visacorp, ws_visacorpN)
 Call Formatear_Valores(ws_visaplat, ws_visaplatN)
 Call Formatear_Valores(ws_visasig, ws_visasigN)
 Call Formatear_Valores(ws_visacta, ws_visactaN)
 Call Formatear_Valores(ws_visainf, ws_visainfN)
 Call Formatear_Valores(ws_visadebemp, ws_visadebempN)
 Call Formatear_Valores(ws_visauclsd, ws_visauclsdN)
 Call Formatear_Valores(ws_visagift, ws_visagiftN)
 
End Sub



Sub borrado()

  Dim wb As Workbook
  Set wb = ThisWorkbook

  wb.Activate

  Set ws_visaclasica = wb.Sheets("VISA_CCLS-TxT")
  Set ws_visaclasicaN = wb.Sheets("VISA_CCLS")

  Set ws_visagold = wb.Sheets("VISA_GOLD-TxT")
  Set ws_visagoldN = wb.Sheets("VISA_GOLD")

  Set ws_visacbsn = wb.Sheets("VISA_CBSN-TxT")
  Set ws_visacbsnN = wb.Sheets("VISA_CBSN")

  Set ws_visacorp = wb.Sheets("VISA_CORP-TxT")
  Set ws_visacorpN = wb.Sheets("VISA_CORP")

  Set ws_visaplat = wb.Sheets("VISA_PLAT-TxT")
  Set ws_visaplatN = wb.Sheets("VISA_PLAT")

  Set ws_visasig = wb.Sheets("VISA_SIG-TxT")
  Set ws_visasigN = wb.Sheets("VISA_SIG")

  Set ws_visacta = wb.Sheets("VISA_CTA-TxT")
  Set ws_visactaN = wb.Sheets("VISA_CTA")

  Set ws_visainf = wb.Sheets("VISA_INF-TxT")
  Set ws_visainfN = wb.Sheets("VISA_INF")

  Set ws_visadebemp = wb.Sheets("VISA_DBSN-TxT")
  Set ws_visadebempN = wb.Sheets("VISA_DBSN")

  Set ws_visauclsd = wb.Sheets("VISA_UCLSD-TxT")
  Set ws_visauclsdN = wb.Sheets("VISA_UCLSD")

  Set ws_visagift = wb.Sheets("VISA_GIFT-TxT")
  Set ws_visagiftN = wb.Sheets("VISA_GIFT")

  
 Call Borrado_Valores(ws_visaclasica, ws_visaclasicaN)
 Call Borrado_Valores(ws_visagold, ws_visagoldN)
 Call Borrado_Valores(ws_visacbsn, ws_visacbsnN)
 Call Borrado_Valores(ws_visacorp, ws_visacorpN)
 Call Borrado_Valores(ws_visaplat, ws_visaplatN)
 Call Borrado_Valores(ws_visasig, ws_visasigN)
 Call Borrado_Valores(ws_visacta, ws_visactaN)
 Call Borrado_Valores(ws_visainf, ws_visainfN)
 Call Borrado_Valores(ws_visadebemp, ws_visadebempN)
 Call Borrado_Valores(ws_visauclsd, ws_visauclsdN)
 Call Borrado_Valores(ws_visagift, ws_visagiftN)

End Sub
