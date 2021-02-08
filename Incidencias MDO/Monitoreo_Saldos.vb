Function completar_ceros_a_la_izquierda(texto, cantidad_de_digitos_requeridos)
    texto = CStr(CInt(texto))
    Do Until Len(texto) >= cantidad_de_digitos_requeridos
        texto = CStr("0" & texto)
    Loop
    completar_ceros_a_la_izquierda = texto
End Function
Sub Filtro()
    Sheets("REPORTE").Select

    'SOLES JUTSU
    fin_I = Application.CountA(Worksheets("REPORTE").Range("C:C"))
        
    cont = 2
    For i = cont To fin_I
        If (Cells(cont, 7) = "NO_LISTA") Then
            Rows(cont).Select
            Selection.Delete Shift:=xlUp
            cont = cont - 1
        End If
        cont = cont + 1
    Next
    
    'CORREO
    fin_f = Application.CountA(Worksheets("REPORTE").Range("C:C"))
        
    cont1 = 2
    For i = cont1 To fin_f
        If IsError(Cells(cont1, 12)) Then
            Rows(cont1).Select
            Selection.Delete Shift:=xlUp
            cont1 = cont1 - 1
        End If
        cont1 = cont1 + 1
    Next
    
     ''' Garantizar que el correo y gerente esten presente
    
    fin_jutsu = Application.CountA(Worksheets("REPORTE").Range("C:C"))
    Range("L2:M" & fin_jutsu).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("L2").Select
    
    '''
    
    ''' Garantizar que la Oficina tenga 4 digitos
    Final = Application.CountA(Worksheets("REPORTE").Range("C:C"))
    For i = 2 To Final
        oficina = CStr(completar_ceros_a_la_izquierda(Sheets("REPORTE").Cells(i, 4), 4))
        Cells(i, 4).Value = oficina
    Next
    '''
    
      ''' Formato de tipo de moneda
    Columns("G:J").Select
    Selection.NumberFormat = "#,##0.00"
    '''
    
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Sheets("SN-EX").Select
    Range("A2").Select
    
    ActiveSheet.Paste
    
    Columns("A:M").EntireColumn.AutoFit
    'ELIMINAR CONFORME
    fin_ff = Application.CountA(Worksheets("SN-EX").Range("C:C"))

    cont2 = 2
    For i = cont2 To fin_ff
        If (Cells(cont2, 11) = "CONFORME") Then
            Rows(cont2).Select
            Selection.Delete Shift:=xlUp
            cont2 = cont2 - 1
        End If
        cont2 = cont2 + 1
    Next
    
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    
    Sheets("EXCEDIDO").Select
    Range("A2").Select
    ActiveSheet.Paste
    
    Columns("A:M").EntireColumn.AutoFit
    'ELIMINAR SALDO NEGATIVO
    fin_fff = Application.CountA(Worksheets("EXCEDIDO").Range("C:C"))

    cont3 = 2
    For i = cont3 To fin_fff
        If (Cells(cont3, 11) = "SALDO NEGATIVO") Then
            Rows(cont3).Select
            Selection.Delete Shift:=xlUp
            cont3 = cont3 - 1
        End If
        cont3 = cont3 + 1
    Next
    
    ''GENERAR BASE PREVENTIVA
    
    Sheets("REPORTE").Select
    
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    
    Sheets("PREVENTIVA").Select
    Range("A2").Select
    ActiveSheet.Paste
    
    Columns("A:M").EntireColumn.AutoFit

    'ELIMINAR SALDO NEGATIVO + EXCEDIDO
    fin_II = Application.CountA(Worksheets("PREVENTIVA").Range("C:C"))
        
    cont = 2
    For i = cont To fin_II
        If (Cells(cont, 11) = "SALDO NEGATIVO" Or Cells(cont, 11) = "EXCEDIDO") Then
            Rows(cont).Select
            Selection.Delete Shift:=xlUp
            cont = cont - 1
        End If
        cont = cont + 1
    Next
    
    ''FILTRANDO RANGO
    fin_III = Application.CountA(Worksheets("PREVENTIVA").Range("C:C"))
        
    cont = 2
    For i = cont To fin_III
        If (Cells(cont, 7) > 180000 Or Cells(cont, 7) < 150000) Then
            Rows(cont).Select
            Selection.Delete Shift:=xlUp
            cont = cont - 1
        End If
        cont = cont + 1
    Next
    
    ''Ocultado hoja
    Sheets("EXCEDIDO").Visible = False
    Sheets("REPORTE").Visible = False
    
End Sub
