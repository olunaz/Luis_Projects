Sub SetDataVisa()

' (1) Shows the msoFileDialogFilePicker dialog box.
' (2) Checks if the user picked a file.
' (3) Stores the path to the selected file in a string type variable.

     'BORRADO DE VALORES
            
    
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
        
            Set ws_menberInfo = wb.Sheets("Menber_Information")
            Set ws_adquiCredito = wb.Sheets("Adquiriente_Credito-TxT")
            Set ws_adquiDebiPre = wb.Sheets("Adquiriente_DebitoPrepago-TxT")
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
            Set ws_valadqui = wb.Sheets("Valores_Adquirientes")
            
            
                
            'PRODUCTOS CREDITO
            'Cash Advances

            'Visa Credito Clasica
            wb2.Worksheets("Produtos Crédito").Range("C32:C37").Value = ws_visaclasica.Range("b24:b29").Value
            wb2.Worksheets("Produtos Crédito").Range("C56:C61").Value = ws_visaclasica.Range("b42:b47").Value
            wb2.Worksheets("Produtos Crédito").Range("C74:C79").Value = ws_visaclasica.Range("b54:b59").Value
            wb2.Worksheets("Produtos Crédito").Range("C98:C103").Value = ws_visaclasica.Range("b60:b65").Value
            wb2.Worksheets("Produtos Crédito").Range("C116:C121").Value = ws_visaclasica.Range("b72:b77").Value
            wb2.Worksheets("Produtos Crédito").Range("C174").Value = ws_visaclasica.Range("b87").Value
            wb2.Worksheets("Produtos Crédito").Range("C146").Value = ws_visaclasica.Range("b90").Value
            wb2.Worksheets("Produtos Crédito").Range("C145").Value = ws_visaclasica.Range("b91").Value
            wb2.Worksheets("Produtos Crédito").Range("C172").Value = ws_visaclasica.Range("b108").Value
            wb2.Worksheets("Produtos Crédito").Range("C173").Value = ws_visaclasica.Range("b109").Value
            wb2.Worksheets("Produtos Crédito").Range("C134").Value = ws_visaclasica.Range("b92").Value

            'Visa Credito Oro
            wb2.Worksheets("Produtos Crédito").Range("F32:F37").Value = ws_visagold.Range("b24:b29").Value
            wb2.Worksheets("Produtos Crédito").Range("F56:F61").Value = ws_visagold.Range("b42:b47").Value
            wb2.Worksheets("Produtos Crédito").Range("F74:F79").Value = ws_visagold.Range("b54:b59").Value
            wb2.Worksheets("Produtos Crédito").Range("F98:F103").Value = ws_visagold.Range("b60:b65").Value
            wb2.Worksheets("Produtos Crédito").Range("F116:F121").Value = ws_visagold.Range("b72:b77").Value
            wb2.Worksheets("Produtos Crédito").Range("F174").Value = ws_visagold.Range("b87").Value
            wb2.Worksheets("Produtos Crédito").Range("F146").Value = ws_visagold.Range("b90").Value
            wb2.Worksheets("Produtos Crédito").Range("F145").Value = ws_visagold.Range("b91").Value
            wb2.Worksheets("Produtos Crédito").Range("F172").Value = ws_visagold.Range("b108").Value
            wb2.Worksheets("Produtos Crédito").Range("F173").Value = ws_visagold.Range("b109").Value
            wb2.Worksheets("Produtos Crédito").Range("F134").Value = ws_visagold.Range("b92").Value

            'Visa Credito Empresarial
            wb2.Worksheets("Produtos Crédito").Range("B32:B37").Value = ws_visacbsn.Range("b24:b29").Value
            wb2.Worksheets("Produtos Crédito").Range("B56:B61").Value = ws_visacbsn.Range("b42:b47").Value
            wb2.Worksheets("Produtos Crédito").Range("B74:B79").Value = ws_visacbsn.Range("b54:b59").Value
            wb2.Worksheets("Produtos Crédito").Range("B98:B103").Value = ws_visacbsn.Range("b60:b65").Value
            wb2.Worksheets("Produtos Crédito").Range("B116:B121").Value = ws_visacbsn.Range("b72:b77").Value
            wb2.Worksheets("Produtos Crédito").Range("B174").Value = ws_visacbsn.Range("b87").Value
            wb2.Worksheets("Produtos Crédito").Range("B146").Value = ws_visacbsn.Range("b90").Value
            wb2.Worksheets("Produtos Crédito").Range("B145").Value = ws_visacbsn.Range("b91").Value
            wb2.Worksheets("Produtos Crédito").Range("B172").Value = ws_visacbsn.Range("b108").Value
            wb2.Worksheets("Produtos Crédito").Range("B173").Value = ws_visacbsn.Range("b109").Value
            wb2.Worksheets("Produtos Crédito").Range("B134").Value = ws_visacbsn.Range("b92").Value

            'Visa Credito Corporativo
            wb2.Worksheets("Produtos Crédito").Range("D32:D37").Value = ws_visacorp.Range("b24:b29").Value
            wb2.Worksheets("Produtos Crédito").Range("D56:D61").Value = ws_visacorp.Range("b42:b47").Value
            wb2.Worksheets("Produtos Crédito").Range("D74:D79").Value = ws_visacorp.Range("b54:b59").Value
            wb2.Worksheets("Produtos Crédito").Range("D98:D103").Value = ws_visacorp.Range("b60:b65").Value
            wb2.Worksheets("Produtos Crédito").Range("D116:D121").Value = ws_visacorp.Range("b72:b77").Value
            wb2.Worksheets("Produtos Crédito").Range("D174").Value = ws_visacorp.Range("b87").Value
            wb2.Worksheets("Produtos Crédito").Range("D146").Value = ws_visacorp.Range("b90").Value
            wb2.Worksheets("Produtos Crédito").Range("D145").Value = ws_visacorp.Range("b91").Value
            wb2.Worksheets("Produtos Crédito").Range("D172").Value = ws_visacorp.Range("b108").Value
            wb2.Worksheets("Produtos Crédito").Range("D173").Value = ws_visacorp.Range("b109").Value
            wb2.Worksheets("Produtos Crédito").Range("D134").Value = ws_visacorp.Range("b92").Value

            'Visa Credit CTA
            wb2.Worksheets("Produtos Crédito").Range("E32:E37").Value = ws_visacta.Range("b24:b29").Value
            wb2.Worksheets("Produtos Crédito").Range("E56:E61").Value = ws_visacta.Range("b42:b47").Value
            wb2.Worksheets("Produtos Crédito").Range("E74:E79").Value = ws_visacta.Range("b54:b59").Value
            wb2.Worksheets("Produtos Crédito").Range("E98:E103").Value = ws_visacta.Range("b60:b65").Value
            wb2.Worksheets("Produtos Crédito").Range("E116:E121").Value = ws_visacta.Range("b72:b77").Value
            wb2.Worksheets("Produtos Crédito").Range("E174").Value = ws_visacta.Range("b87").Value
            wb2.Worksheets("Produtos Crédito").Range("E146").Value = ws_visacta.Range("b90").Value
            wb2.Worksheets("Produtos Crédito").Range("E145").Value = ws_visacta.Range("b91").Value
            wb2.Worksheets("Produtos Crédito").Range("E172").Value = ws_visacta.Range("b108").Value
            wb2.Worksheets("Produtos Crédito").Range("E173").Value = ws_visacta.Range("b109").Value
            wb2.Worksheets("Produtos Crédito").Range("E134").Value = ws_visacta.Range("b92").Value


            wb2.Worksheets("Produtos Crédito").Range("G32:G37").Value = ws_visainf.Range("b24:b29").Value
            wb2.Worksheets("Produtos Crédito").Range("G56:G61").Value = ws_visainf.Range("b42:b47").Value
            wb2.Worksheets("Produtos Crédito").Range("G74:G79").Value = ws_visainf.Range("b54:b59").Value
            wb2.Worksheets("Produtos Crédito").Range("G98:G103").Value = ws_visainf.Range("b60:b65").Value
            wb2.Worksheets("Produtos Crédito").Range("G116:G121").Value = ws_visainf.Range("b72:b77").Value
            wb2.Worksheets("Produtos Crédito").Range("G174").Value = ws_visainf.Range("b87").Value
            wb2.Worksheets("Produtos Crédito").Range("G146").Value = ws_visainf.Range("b90").Value
            wb2.Worksheets("Produtos Crédito").Range("G145").Value = ws_visainf.Range("b91").Value
            wb2.Worksheets("Produtos Crédito").Range("G172").Value = ws_visainf.Range("b108").Value
            wb2.Worksheets("Produtos Crédito").Range("G173").Value = ws_visainf.Range("b109").Value
            wb2.Worksheets("Produtos Crédito").Range("G134").Value = ws_visainf.Range("b92").Value


            wb2.Worksheets("Produtos Crédito").Range("H32:H37").Value = ws_visaplat.Range("b24:b29").Value
            wb2.Worksheets("Produtos Crédito").Range("H56:H61").Value = ws_visaplat.Range("b42:b47").Value
            wb2.Worksheets("Produtos Crédito").Range("H74:H79").Value = ws_visaplat.Range("b54:b59").Value
            wb2.Worksheets("Produtos Crédito").Range("H98:H103").Value = ws_visaplat.Range("b60:b65").Value
            wb2.Worksheets("Produtos Crédito").Range("H116:H121").Value = ws_visaplat.Range("b72:b77").Value
            wb2.Worksheets("Produtos Crédito").Range("H174").Value = ws_visaplat.Range("b87").Value
            wb2.Worksheets("Produtos Crédito").Range("H146").Value = ws_visaplat.Range("b90").Value
            wb2.Worksheets("Produtos Crédito").Range("H145").Value = ws_visaplat.Range("b91").Value
            wb2.Worksheets("Produtos Crédito").Range("H172").Value = ws_visaplat.Range("b108").Value
            wb2.Worksheets("Produtos Crédito").Range("H173").Value = ws_visaplat.Range("b109").Value
            wb2.Worksheets("Produtos Crédito").Range("H134").Value = ws_visaplat.Range("b92").Value


            wb2.Worksheets("Produtos Crédito").Range("I32:I37").Value = ws_visasig.Range("b24:b29").Value
            wb2.Worksheets("Produtos Crédito").Range("I56:I61").Value = ws_visasig.Range("b42:b47").Value
            wb2.Worksheets("Produtos Crédito").Range("I74:I79").Value = ws_visasig.Range("b54:b59").Value
            wb2.Worksheets("Produtos Crédito").Range("I98:I103").Value = ws_visasig.Range("b60:b65").Value
            wb2.Worksheets("Produtos Crédito").Range("I116:H121").Value = ws_visasig.Range("b72:b77").Value
            wb2.Worksheets("Produtos Crédito").Range("I174").Value = ws_visasig.Range("b87").Value
            wb2.Worksheets("Produtos Crédito").Range("I146").Value = ws_visasig.Range("b90").Value
            wb2.Worksheets("Produtos Crédito").Range("I145").Value = ws_visasig.Range("b91").Value
            wb2.Worksheets("Produtos Crédito").Range("I172").Value = ws_visasig.Range("b108").Value
            wb2.Worksheets("Produtos Crédito").Range("I173").Value = ws_visasig.Range("b109").Value
            wb2.Worksheets("Produtos Crédito").Range("I134").Value = ws_visasig.Range("b92").Value
            
            


            'PRODUCTOS DEBITO

            wb2.Worksheets("Produtos Débito").Range("B32:B37").Value = ws_visauclsd.Range("B24:B29").Value
            wb2.Worksheets("Produtos Débito").Range("B56:B61").Value = ws_visauclsd.Range("b42:b47").Value
            wb2.Worksheets("Produtos Débito").Range("B68:B73").Value = ws_visauclsd.Range("b54:b59").Value
            wb2.Worksheets("Produtos Débito").Range("B86:B91").Value = ws_visauclsd.Range("b60:b65").Value
            wb2.Worksheets("Produtos Débito").Range("B104:B109").Value = ws_visauclsd.Range("b72:b77").Value
            wb2.Worksheets("Produtos Débito").Range("B134").Value = ws_visauclsd.Range("b78").Value
            wb2.Worksheets("Produtos Débito").Range("B122").Value = ws_visauclsd.Range("b80").Value
            wb2.Worksheets("Produtos Débito").Range("B135").Value = ws_visauclsd.Range("b81").Value
            'wb2.Worksheets("Produtos Débito").Range("B138").Value = ws_visauclsd.Range("b83").Value
            'wb2.Worksheets("Produtos Débito").Range("B137").Value = ws_visauclsd.Range("b85").Value

            wb2.Worksheets("Produtos Débito").Range("C32:C37").Value = ws_visadebemp.Range("B24:B29").Value
            wb2.Worksheets("Produtos Débito").Range("C56:C61").Value = ws_visadebemp.Range("b42:b47").Value
            wb2.Worksheets("Produtos Débito").Range("C68:C73").Value = ws_visadebemp.Range("b54:b59").Value
            wb2.Worksheets("Produtos Débito").Range("C86:C91").Value = ws_visadebemp.Range("b60:b65").Value
            wb2.Worksheets("Produtos Débito").Range("C104:C109").Value = ws_visadebemp.Range("b72:b77").Value
            wb2.Worksheets("Produtos Débito").Range("C134").Value = ws_visadebemp.Range("b78").Value
            wb2.Worksheets("Produtos Débito").Range("C122").Value = ws_visadebemp.Range("b80").Value
            wb2.Worksheets("Produtos Débito").Range("C135").Value = ws_visadebemp.Range("b81").Value
            'wb2.Worksheets("Produtos Débito").Range("C138").Value = ws_visadebemp.Range("b83").Value
            'wb2.Worksheets("Produtos Débito").Range("C137").Value = ws_visadebemp.Range("b85").Value

           
           'PRODUCTOS PREPAGADO
            wb2.Worksheets("Produtos Prepagados").Range("B50:B55").Value = ws_visauclsd.Range("B42:B47").Value
            wb2.Worksheets("Produtos Prepagados").Range("B86:B91").Value = ws_visauclsd.Range("b60:b65").Value

            wb2.Worksheets("Produtos Prepagados").Range("B120").Value = ws_visauclsd.Range("B80").Value
            wb2.Worksheets("Produtos Prepagados").Range("B135").Value = ws_visauclsd.Range("B81").Value

            Dim rng As Range
            
            Set ws1 = wb2.Worksheets("Produtos Crédito")
            Set ws2 = wb2.Worksheets("Produtos Débito")
            Set ws3 = wb2.Worksheets("Produtos Prepagados")
            
            For Each rng In ws1.Range("B2:I177")
                If IsEmpty(rng) Then
                    rng.Value = 0
                End If
            Next
            
            For Each rng In ws2.Range("B2:C154")
                If IsEmpty(rng) Then
                    rng.Value = 0
                End If
            Next
            
            For Each rng In ws3.Range("B2:B155")
                If IsEmpty(rng) Then
                    rng.Value = 0
                End If
            Next
            
            ' *********************
            
            
            ' *********************
            
            ' Example: print the path of the selected file to the immediate window
            Debug.Print strFilePath ' remove in production
        End If
    End With
End Sub



