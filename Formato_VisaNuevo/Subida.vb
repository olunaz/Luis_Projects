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
                
                wb2.Worksheets("Produtos Crédito").Range("B2:I177").ClearContents
                wb2.Worksheets("Produtos Débito").Range("B2:C177").ClearContents
                wb2.Worksheets("Produtos Prepagados").Range("B2:B177").ClearContents
                
        
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

                Set ws_visaclasica_ch = wb.Sheets("VISA_CCLS-CH")
                Set ws_visagold_ch = wb.Sheets("VISA_GOLD-CH")
                Set ws_visacbsn_ch = wb.Sheets("VISA_CBSN-CH")
                Set ws_visacorp_ch = wb.Sheets("VISA_CORP-CH")
                Set ws_visaplat_ch = wb.Sheets("VISA_PLAT-CH")
                Set ws_visasig_ch = wb.Sheets("VISA_SIG-CH")
                Set ws_visacta_ch = wb.Sheets("VISA_CTA-CH")
                Set ws_visainf_ch = wb.Sheets("VISA_INF-CH")
                Set ws_visadebemp_ch = wb.Sheets("VISA_DBSN-CH")
                Set ws_visauclsd_ch = wb.Sheets("VISA_UCLSD-CH")
                Set ws_visagift_ch = wb.Sheets("VISA_GIFT-CH")
                
                
                    
                'PRODUCTOS CREDITO
                'Cash Advances

                'Visa Credito Clasica

                'CASH ADVANCES COUNT

                wb2.Worksheets("Produtos Crédito").Range("C32").Value = ws_visaclasica.Range("B24").Value
                wb2.Worksheets("Produtos Crédito").Range("C33").Value = ws_visaclasica.Range("B26").Value
                wb2.Worksheets("Produtos Crédito").Range("C34").Value = ws_visaclasica.Range("B28").Value
                wb2.Worksheets("Produtos Crédito").Range("C35").Value = ws_visaclasica.Range("B25").Value
                wb2.Worksheets("Produtos Crédito").Range("C36").Value = ws_visaclasica.Range("B27").Value
                wb2.Worksheets("Produtos Crédito").Range("C37").Value = ws_visaclasica.Range("B29").Value
                
                'NATIONALS PAYMENTS COUNT

                wb2.Worksheets("Produtos Crédito").Range("C56").Value = ws_visaclasica.Range("b42").Value
                wb2.Worksheets("Produtos Crédito").Range("C57").Value = ws_visaclasica.Range("b44").Value
                wb2.Worksheets("Produtos Crédito").Range("C58").Value = ws_visaclasica.Range("b46").Value
                wb2.Worksheets("Produtos Crédito").Range("C59").Value = ws_visaclasica.Range("b43").Value
                wb2.Worksheets("Produtos Crédito").Range("C60").Value = ws_visaclasica.Range("b45").Value
                wb2.Worksheets("Produtos Crédito").Range("C61").Value = ws_visaclasica.Range("b47").Value


                'NATIONAL PAYMENTS COUNT
                wb2.Worksheets("Produtos Crédito").Range("C74").Value = ws_visaclasica.Range("b54").Value
                wb2.Worksheets("Produtos Crédito").Range("C75").Value = ws_visaclasica.Range("b56").Value
                wb2.Worksheets("Produtos Crédito").Range("C76").Value = ws_visaclasica.Range("b58").Value
                wb2.Worksheets("Produtos Crédito").Range("C77").Value = ws_visaclasica.Range("b55").Value
                wb2.Worksheets("Produtos Crédito").Range("C78").Value = ws_visaclasica.Range("b57").Value
                wb2.Worksheets("Produtos Crédito").Range("C79").Value = ws_visaclasica.Range("b59").Value

                'INTERNATIONAL PAYMENTS COUNT
                wb2.Worksheets("Produtos Crédito").Range("C98").Value = ws_visaclasica.Range("b60").Value
                wb2.Worksheets("Produtos Crédito").Range("C99").Value = ws_visaclasica.Range("b62").Value
                wb2.Worksheets("Produtos Crédito").Range("C100").Value = ws_visaclasica.Range("b64").Value
                wb2.Worksheets("Produtos Crédito").Range("C101").Value = ws_visaclasica.Range("b61").Value
                wb2.Worksheets("Produtos Crédito").Range("C102").Value = ws_visaclasica.Range("b63").Value
                wb2.Worksheets("Produtos Crédito").Range("C103").Value = ws_visaclasica.Range("b65").Value

                'INTERANTIONAL ATM CASH ADVACNES
                wb2.Worksheets("Produtos Crédito").Range("C116").Value = ws_visaclasica.Range("b72").Value
                wb2.Worksheets("Produtos Crédito").Range("C117").Value = ws_visaclasica.Range("b74").Value
                wb2.Worksheets("Produtos Crédito").Range("C118").Value = ws_visaclasica.Range("b76").Value
                wb2.Worksheets("Produtos Crédito").Range("C119").Value = ws_visaclasica.Range("b73").Value
                wb2.Worksheets("Produtos Crédito").Range("C120").Value = ws_visaclasica.Range("b75").Value
                wb2.Worksheets("Produtos Crédito").Range("C121").Value = ws_visaclasica.Range("b77").Value

                wb2.Worksheets("Produtos Crédito").Range("C174").Value = ws_visaclasica.Range("b87").Value
                wb2.Worksheets("Produtos Crédito").Range("C146").Value = ws_visaclasica.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("C145").Value = ws_visaclasica.Range("b91").Value
                wb2.Worksheets("Produtos Crédito").Range("C172").Value = ws_visaclasica.Range("b108").Value
                wb2.Worksheets("Produtos Crédito").Range("C173").Value = ws_visaclasica.Range("b109").Value
                wb2.Worksheets("Produtos Crédito").Range("C134").Value = ws_visaclasica.Range("b92").Value

                 'Nueva modificacion
                wb2.Worksheets("Produtos Crédito").Range("C148:C151").Value = ws_visaclasica.Range("b95:b98").Value
                wb2.Worksheets("Produtos Crédito").Range("C152").Value = ws_visaclasica.Range("b99").Value + ws_visaclasica.Range("b101").Value + ws_visaclasica.Range("b103").Value
                wb2.Worksheets("Produtos Crédito").Range("C153").Value = ws_visaclasica.Range("b100").Value + ws_visaclasica.Range("b102").Value + ws_visaclasica.Range("b104").Value
                wb2.Worksheets("Produtos Crédito").Range("C144").Value = ws_visaclasica.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("C136").Value = ws_visaclasica_ch.Range("b30").Value


                'Visa Credito Oro
                'CASH ADVANCES COUNT
                wb2.Worksheets("Produtos Crédito").Range("F32").Value = ws_visagold.Range("B24").Value
                wb2.Worksheets("Produtos Crédito").Range("F33").Value = ws_visagold.Range("B26").Value
                wb2.Worksheets("Produtos Crédito").Range("F34").Value = ws_visagold.Range("B28").Value
                wb2.Worksheets("Produtos Crédito").Range("F35").Value = ws_visagold.Range("B25").Value
                wb2.Worksheets("Produtos Crédito").Range("F36").Value = ws_visagold.Range("B27").Value
                wb2.Worksheets("Produtos Crédito").Range("F37").Value = ws_visagold.Range("B29").Value
                'NATIONALS PAYMENTS COUNF
                wb2.Worksheets("Produtos Crédito").Range("F56").Value = ws_visagold.Range("b42").Value
                wb2.Worksheets("Produtos Crédito").Range("F57").Value = ws_visagold.Range("b44").Value
                wb2.Worksheets("Produtos Crédito").Range("F58").Value = ws_visagold.Range("b46").Value
                wb2.Worksheets("Produtos Crédito").Range("F59").Value = ws_visagold.Range("b43").Value
                wb2.Worksheets("Produtos Crédito").Range("F60").Value = ws_visagold.Range("b45").Value
                wb2.Worksheets("Produtos Crédito").Range("F61").Value = ws_visagold.Range("b47").Value
                'NATIONAL PAYMENTS COUNF
                wb2.Worksheets("Produtos Crédito").Range("F74").Value = ws_visagold.Range("b54").Value
                wb2.Worksheets("Produtos Crédito").Range("F75").Value = ws_visagold.Range("b56").Value
                wb2.Worksheets("Produtos Crédito").Range("F76").Value = ws_visagold.Range("b58").Value
                wb2.Worksheets("Produtos Crédito").Range("F77").Value = ws_visagold.Range("b55").Value
                wb2.Worksheets("Produtos Crédito").Range("F78").Value = ws_visagold.Range("b57").Value
                wb2.Worksheets("Produtos Crédito").Range("F79").Value = ws_visagold.Range("b59").Value
                'INTERNATIONAL PAYMENTS COUNF
                wb2.Worksheets("Produtos Crédito").Range("F98").Value = ws_visagold.Range("b60").Value
                wb2.Worksheets("Produtos Crédito").Range("F99").Value = ws_visagold.Range("b62").Value
                wb2.Worksheets("Produtos Crédito").Range("F100").Value = ws_visagold.Range("b64").Value
                wb2.Worksheets("Produtos Crédito").Range("F101").Value = ws_visagold.Range("b61").Value
                wb2.Worksheets("Produtos Crédito").Range("F102").Value = ws_visagold.Range("b63").Value
                wb2.Worksheets("Produtos Crédito").Range("F103").Value = ws_visagold.Range("b65").Value
                'INTERANTIONAL ATM CASH ADVACNEF
                wb2.Worksheets("Produtos Crédito").Range("F116").Value = ws_visagold.Range("b72").Value
                wb2.Worksheets("Produtos Crédito").Range("F117").Value = ws_visagold.Range("b74").Value
                wb2.Worksheets("Produtos Crédito").Range("F118").Value = ws_visagold.Range("b76").Value
                wb2.Worksheets("Produtos Crédito").Range("F119").Value = ws_visagold.Range("b73").Value
                wb2.Worksheets("Produtos Crédito").Range("F120").Value = ws_visagold.Range("b75").Value
                wb2.Worksheets("Produtos Crédito").Range("F121").Value = ws_visagold.Range("b77").Value

                wb2.Worksheets("Produtos Crédito").Range("F174").Value = ws_visagold.Range("b87").Value
                wb2.Worksheets("Produtos Crédito").Range("F146").Value = ws_visagold.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("F145").Value = ws_visagold.Range("b91").Value
                wb2.Worksheets("Produtos Crédito").Range("F172").Value = ws_visagold.Range("b108").Value
                wb2.Worksheets("Produtos Crédito").Range("F173").Value = ws_visagold.Range("b109").Value
                wb2.Worksheets("Produtos Crédito").Range("F134").Value = ws_visagold.Range("b92").Value

                'Nueva modificacion
                wb2.Worksheets("Produtos Crédito").Range("F148:F151").Value = ws_visagold.Range("b95:b98").Value
                wb2.Worksheets("Produtos Crédito").Range("F152").Value = ws_visagold.Range("b99").Value + ws_visagold.Range("b101").Value + ws_visagold.Range("b103").Value
                wb2.Worksheets("Produtos Crédito").Range("F153").Value = ws_visagold.Range("b100").Value + ws_visagold.Range("b102").Value + ws_visagold.Range("b104").Value
                wb2.Worksheets("Produtos Crédito").Range("F144").Value = ws_visagold.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("F136").Value = ws_visagold_ch.Range("b30").Value




                'Visa Credito Empresarial
                'CASH ADVANCES COUNT
                wb2.Worksheets("Produtos Crédito").Range("B32").Value = ws_visacbsn.Range("B24").Value
                wb2.Worksheets("Produtos Crédito").Range("B33").Value = ws_visacbsn.Range("B26").Value
                wb2.Worksheets("Produtos Crédito").Range("B34").Value = ws_visacbsn.Range("B28").Value
                wb2.Worksheets("Produtos Crédito").Range("B35").Value = ws_visacbsn.Range("B25").Value
                wb2.Worksheets("Produtos Crédito").Range("B36").Value = ws_visacbsn.Range("B27").Value
                wb2.Worksheets("Produtos Crédito").Range("B37").Value = ws_visacbsn.Range("B29").Value
                'NATIONALS PAYMENTS COUNB
                wb2.Worksheets("Produtos Crédito").Range("B56").Value = ws_visacbsn.Range("b42").Value
                wb2.Worksheets("Produtos Crédito").Range("B57").Value = ws_visacbsn.Range("b44").Value
                wb2.Worksheets("Produtos Crédito").Range("B58").Value = ws_visacbsn.Range("b46").Value
                wb2.Worksheets("Produtos Crédito").Range("B59").Value = ws_visacbsn.Range("b43").Value
                wb2.Worksheets("Produtos Crédito").Range("B60").Value = ws_visacbsn.Range("b45").Value
                wb2.Worksheets("Produtos Crédito").Range("B61").Value = ws_visacbsn.Range("b47").Value
                'NATIONAL PAYMENTS COUNB
                wb2.Worksheets("Produtos Crédito").Range("B74").Value = ws_visacbsn.Range("b54").Value
                wb2.Worksheets("Produtos Crédito").Range("B75").Value = ws_visacbsn.Range("b56").Value
                wb2.Worksheets("Produtos Crédito").Range("B76").Value = ws_visacbsn.Range("b58").Value
                wb2.Worksheets("Produtos Crédito").Range("B77").Value = ws_visacbsn.Range("b55").Value
                wb2.Worksheets("Produtos Crédito").Range("B78").Value = ws_visacbsn.Range("b57").Value
                wb2.Worksheets("Produtos Crédito").Range("B79").Value = ws_visacbsn.Range("b59").Value
                'INTERNATIONAL PAYMENTS COUNB
                wb2.Worksheets("Produtos Crédito").Range("B98").Value = ws_visacbsn.Range("b60").Value
                wb2.Worksheets("Produtos Crédito").Range("B99").Value = ws_visacbsn.Range("b62").Value
                wb2.Worksheets("Produtos Crédito").Range("B100").Value = ws_visacbsn.Range("b64").Value
                wb2.Worksheets("Produtos Crédito").Range("B101").Value = ws_visacbsn.Range("b61").Value
                wb2.Worksheets("Produtos Crédito").Range("B102").Value = ws_visacbsn.Range("b63").Value
                wb2.Worksheets("Produtos Crédito").Range("B103").Value = ws_visacbsn.Range("b65").Value
                'INTERANTIONAL ATM CASH ADVACNEB
                wb2.Worksheets("Produtos Crédito").Range("B116").Value = ws_visacbsn.Range("b72").Value
                wb2.Worksheets("Produtos Crédito").Range("B117").Value = ws_visacbsn.Range("b74").Value
                wb2.Worksheets("Produtos Crédito").Range("B118").Value = ws_visacbsn.Range("b76").Value
                wb2.Worksheets("Produtos Crédito").Range("B119").Value = ws_visacbsn.Range("b73").Value
                wb2.Worksheets("Produtos Crédito").Range("B120").Value = ws_visacbsn.Range("b75").Value
                wb2.Worksheets("Produtos Crédito").Range("B121").Value = ws_visacbsn.Range("b77").Value

                wb2.Worksheets("Produtos Crédito").Range("B174").Value = ws_visacbsn.Range("b87").Value
                wb2.Worksheets("Produtos Crédito").Range("B146").Value = ws_visacbsn.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("B145").Value = ws_visacbsn.Range("b91").Value
                wb2.Worksheets("Produtos Crédito").Range("B172").Value = ws_visacbsn.Range("b108").Value
                wb2.Worksheets("Produtos Crédito").Range("B173").Value = ws_visacbsn.Range("b109").Value
                wb2.Worksheets("Produtos Crédito").Range("B134").Value = ws_visacbsn.Range("b92").Value

                'Nueva modificacion
                wb2.Worksheets("Produtos Crédito").Range("B148:B151").Value = ws_visacbsn.Range("b95:b98").Value
                wb2.Worksheets("Produtos Crédito").Range("B152").Value = ws_visacbsn.Range("b99").Value + ws_visacbsn.Range("b101").Value + ws_visacbsn.Range("b103").Value
                wb2.Worksheets("Produtos Crédito").Range("B153").Value = ws_visacbsn.Range("b100").Value + ws_visacbsn.Range("b102").Value + ws_visacbsn.Range("b104").Value
                wb2.Worksheets("Produtos Crédito").Range("B144").Value = ws_visacbsn.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("B136").Value = ws_visacbsn_ch.Range("b30").Value

                'Visa Credito Corporativo
                'CASH ADVANCES COUNT
                wb2.Worksheets("Produtos Crédito").Range("D32").Value = ws_visacorp.Range("B24").Value
                wb2.Worksheets("Produtos Crédito").Range("D33").Value = ws_visacorp.Range("B26").Value
                wb2.Worksheets("Produtos Crédito").Range("D34").Value = ws_visacorp.Range("B28").Value
                wb2.Worksheets("Produtos Crédito").Range("D35").Value = ws_visacorp.Range("B25").Value
                wb2.Worksheets("Produtos Crédito").Range("D36").Value = ws_visacorp.Range("B27").Value
                wb2.Worksheets("Produtos Crédito").Range("D37").Value = ws_visacorp.Range("B29").Value
                'NATIONALS PAYMENTS COUND
                wb2.Worksheets("Produtos Crédito").Range("D56").Value = ws_visacorp.Range("b42").Value
                wb2.Worksheets("Produtos Crédito").Range("D57").Value = ws_visacorp.Range("b44").Value
                wb2.Worksheets("Produtos Crédito").Range("D58").Value = ws_visacorp.Range("b46").Value
                wb2.Worksheets("Produtos Crédito").Range("D59").Value = ws_visacorp.Range("b43").Value
                wb2.Worksheets("Produtos Crédito").Range("D60").Value = ws_visacorp.Range("b45").Value
                wb2.Worksheets("Produtos Crédito").Range("D61").Value = ws_visacorp.Range("b47").Value
                'NATIONAL PAYMENTS COUND
                wb2.Worksheets("Produtos Crédito").Range("D74").Value = ws_visacorp.Range("b54").Value
                wb2.Worksheets("Produtos Crédito").Range("D75").Value = ws_visacorp.Range("b56").Value
                wb2.Worksheets("Produtos Crédito").Range("D76").Value = ws_visacorp.Range("b58").Value
                wb2.Worksheets("Produtos Crédito").Range("D77").Value = ws_visacorp.Range("b55").Value
                wb2.Worksheets("Produtos Crédito").Range("D78").Value = ws_visacorp.Range("b57").Value
                wb2.Worksheets("Produtos Crédito").Range("D79").Value = ws_visacorp.Range("b59").Value
                'INTERNATIONAL PAYMENTS COUND
                wb2.Worksheets("Produtos Crédito").Range("D98").Value = ws_visacorp.Range("b60").Value
                wb2.Worksheets("Produtos Crédito").Range("D99").Value = ws_visacorp.Range("b62").Value
                wb2.Worksheets("Produtos Crédito").Range("D100").Value = ws_visacorp.Range("b64").Value
                wb2.Worksheets("Produtos Crédito").Range("D101").Value = ws_visacorp.Range("b61").Value
                wb2.Worksheets("Produtos Crédito").Range("D102").Value = ws_visacorp.Range("b63").Value
                wb2.Worksheets("Produtos Crédito").Range("D103").Value = ws_visacorp.Range("b65").Value
                'INTERANTIONAL ATM CASH ADVACNED
                wb2.Worksheets("Produtos Crédito").Range("D116").Value = ws_visacorp.Range("b72").Value
                wb2.Worksheets("Produtos Crédito").Range("D117").Value = ws_visacorp.Range("b74").Value
                wb2.Worksheets("Produtos Crédito").Range("D118").Value = ws_visacorp.Range("b76").Value
                wb2.Worksheets("Produtos Crédito").Range("D119").Value = ws_visacorp.Range("b73").Value
                wb2.Worksheets("Produtos Crédito").Range("D120").Value = ws_visacorp.Range("b75").Value
                wb2.Worksheets("Produtos Crédito").Range("D121").Value = ws_visacorp.Range("b77").Value

                wb2.Worksheets("Produtos Crédito").Range("D174").Value = ws_visacorp.Range("b87").Value
                wb2.Worksheets("Produtos Crédito").Range("D146").Value = ws_visacorp.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("D145").Value = ws_visacorp.Range("b91").Value
                wb2.Worksheets("Produtos Crédito").Range("D172").Value = ws_visacorp.Range("b108").Value
                wb2.Worksheets("Produtos Crédito").Range("D173").Value = ws_visacorp.Range("b109").Value
                wb2.Worksheets("Produtos Crédito").Range("D134").Value = ws_visacorp.Range("b92").Value

                'Nueva modificacion
                wb2.Worksheets("Produtos Crédito").Range("D148:D151").Value = ws_visacorp.Range("b95:b98").Value
                wb2.Worksheets("Produtos Crédito").Range("D152").Value = ws_visacorp.Range("b99").Value + ws_visacorp.Range("b101").Value + ws_visacorp.Range("b103").Value
                wb2.Worksheets("Produtos Crédito").Range("D153").Value = ws_visacorp.Range("b100").Value + ws_visacorp.Range("b102").Value + ws_visacorp.Range("b104").Value
                wb2.Worksheets("Produtos Crédito").Range("D144").Value = ws_visacorp.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("D136").Value = ws_visacorp_ch.Range("b30").Value



                '///////////////


                'Visa Credit CTA
              'CASH ADVANCES COUNT
                wb2.Worksheets("Produtos Crédito").Range("E32").Value = ws_visacta.Range("B24").Value
                wb2.Worksheets("Produtos Crédito").Range("E33").Value = ws_visacta.Range("B26").Value
                wb2.Worksheets("Produtos Crédito").Range("E34").Value = ws_visacta.Range("B28").Value
                wb2.Worksheets("Produtos Crédito").Range("E35").Value = ws_visacta.Range("B25").Value
                wb2.Worksheets("Produtos Crédito").Range("E36").Value = ws_visacta.Range("B27").Value
                wb2.Worksheets("Produtos Crédito").Range("E37").Value = ws_visacta.Range("B29").Value
                'NATIONALS PAYMENTS COUNE
                wb2.Worksheets("Produtos Crédito").Range("E56").Value = ws_visacta.Range("b42").Value
                wb2.Worksheets("Produtos Crédito").Range("E57").Value = ws_visacta.Range("b44").Value
                wb2.Worksheets("Produtos Crédito").Range("E58").Value = ws_visacta.Range("b46").Value
                wb2.Worksheets("Produtos Crédito").Range("E59").Value = ws_visacta.Range("b43").Value
                wb2.Worksheets("Produtos Crédito").Range("E60").Value = ws_visacta.Range("b45").Value
                wb2.Worksheets("Produtos Crédito").Range("E61").Value = ws_visacta.Range("b47").Value
                'NATIONAL PAYMENTS COUNE
                wb2.Worksheets("Produtos Crédito").Range("E74").Value = ws_visacta.Range("b54").Value
                wb2.Worksheets("Produtos Crédito").Range("E75").Value = ws_visacta.Range("b56").Value
                wb2.Worksheets("Produtos Crédito").Range("E76").Value = ws_visacta.Range("b58").Value
                wb2.Worksheets("Produtos Crédito").Range("E77").Value = ws_visacta.Range("b55").Value
                wb2.Worksheets("Produtos Crédito").Range("E78").Value = ws_visacta.Range("b57").Value
                wb2.Worksheets("Produtos Crédito").Range("E79").Value = ws_visacta.Range("b59").Value
                'INTERNATIONAL PAYMENTS COUNE
                wb2.Worksheets("Produtos Crédito").Range("E98").Value = ws_visacta.Range("b60").Value
                wb2.Worksheets("Produtos Crédito").Range("E99").Value = ws_visacta.Range("b62").Value
                wb2.Worksheets("Produtos Crédito").Range("E100").Value = ws_visacta.Range("b64").Value
                wb2.Worksheets("Produtos Crédito").Range("E101").Value = ws_visacta.Range("b61").Value
                wb2.Worksheets("Produtos Crédito").Range("E102").Value = ws_visacta.Range("b63").Value
                wb2.Worksheets("Produtos Crédito").Range("E103").Value = ws_visacta.Range("b65").Value
                'INTERANTIONAL ATM CASH ADVACNEE
                wb2.Worksheets("Produtos Crédito").Range("E116").Value = ws_visacta.Range("b72").Value
                wb2.Worksheets("Produtos Crédito").Range("E117").Value = ws_visacta.Range("b74").Value
                wb2.Worksheets("Produtos Crédito").Range("E118").Value = ws_visacta.Range("b76").Value
                wb2.Worksheets("Produtos Crédito").Range("E119").Value = ws_visacta.Range("b73").Value
                wb2.Worksheets("Produtos Crédito").Range("E120").Value = ws_visacta.Range("b75").Value
                wb2.Worksheets("Produtos Crédito").Range("E121").Value = ws_visacta.Range("b77").Value


                wb2.Worksheets("Produtos Crédito").Range("E174").Value = ws_visacta.Range("b87").Value
                wb2.Worksheets("Produtos Crédito").Range("E146").Value = ws_visacta.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("E145").Value = ws_visacta.Range("b91").Value
                wb2.Worksheets("Produtos Crédito").Range("E172").Value = ws_visacta.Range("b108").Value
                wb2.Worksheets("Produtos Crédito").Range("E173").Value = ws_visacta.Range("b109").Value
                wb2.Worksheets("Produtos Crédito").Range("E134").Value = ws_visacta.Range("b92").Value

                'Nueva modificacion
                wb2.Worksheets("Produtos Crédito").Range("E148:E151").Value = ws_visacta.Range("b95:b98").Value
                wb2.Worksheets("Produtos Crédito").Range("E152").Value = ws_visacta.Range("b99").Value + ws_visacta.Range("b101").Value + ws_visacta.Range("b103").Value
                wb2.Worksheets("Produtos Crédito").Range("E153").Value = ws_visacta.Range("b100").Value + ws_visacta.Range("b102").Value + ws_visacta.Range("b104").Value
                wb2.Worksheets("Produtos Crédito").Range("E144").Value = ws_visacta.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("E136").Value = ws_visacta_ch.Range("b30").Value


               'CASH ADVANCES COUNT
                wb2.Worksheets("Produtos Crédito").Range("G32").Value = ws_visainf.Range("B24").Value
                wb2.Worksheets("Produtos Crédito").Range("G33").Value = ws_visainf.Range("B26").Value
                wb2.Worksheets("Produtos Crédito").Range("G34").Value = ws_visainf.Range("B28").Value
                wb2.Worksheets("Produtos Crédito").Range("G35").Value = ws_visainf.Range("B25").Value
                wb2.Worksheets("Produtos Crédito").Range("G36").Value = ws_visainf.Range("B27").Value
                wb2.Worksheets("Produtos Crédito").Range("G37").Value = ws_visainf.Range("B29").Value
                'NATIONALS PAYMENTS COUNG
                wb2.Worksheets("Produtos Crédito").Range("G56").Value = ws_visainf.Range("b42").Value
                wb2.Worksheets("Produtos Crédito").Range("G57").Value = ws_visainf.Range("b44").Value
                wb2.Worksheets("Produtos Crédito").Range("G58").Value = ws_visainf.Range("b46").Value
                wb2.Worksheets("Produtos Crédito").Range("G59").Value = ws_visainf.Range("b43").Value
                wb2.Worksheets("Produtos Crédito").Range("G60").Value = ws_visainf.Range("b45").Value
                wb2.Worksheets("Produtos Crédito").Range("G61").Value = ws_visainf.Range("b47").Value
                'NATIONAL PAYMENTS COUNG
                wb2.Worksheets("Produtos Crédito").Range("G74").Value = ws_visainf.Range("b54").Value
                wb2.Worksheets("Produtos Crédito").Range("G75").Value = ws_visainf.Range("b56").Value
                wb2.Worksheets("Produtos Crédito").Range("G76").Value = ws_visainf.Range("b58").Value
                wb2.Worksheets("Produtos Crédito").Range("G77").Value = ws_visainf.Range("b55").Value
                wb2.Worksheets("Produtos Crédito").Range("G78").Value = ws_visainf.Range("b57").Value
                wb2.Worksheets("Produtos Crédito").Range("G79").Value = ws_visainf.Range("b59").Value
                'INTERNATIONAL PAYMENTS COUNG
                wb2.Worksheets("Produtos Crédito").Range("G98").Value = ws_visainf.Range("b60").Value
                wb2.Worksheets("Produtos Crédito").Range("G99").Value = ws_visainf.Range("b62").Value
                wb2.Worksheets("Produtos Crédito").Range("G100").Value = ws_visainf.Range("b64").Value
                wb2.Worksheets("Produtos Crédito").Range("G101").Value = ws_visainf.Range("b61").Value
                wb2.Worksheets("Produtos Crédito").Range("G102").Value = ws_visainf.Range("b63").Value
                wb2.Worksheets("Produtos Crédito").Range("G103").Value = ws_visainf.Range("b65").Value
                'INTERANTIONAL ATM CASH ADVACNEG
                wb2.Worksheets("Produtos Crédito").Range("G116").Value = ws_visainf.Range("b72").Value
                wb2.Worksheets("Produtos Crédito").Range("G117").Value = ws_visainf.Range("b74").Value
                wb2.Worksheets("Produtos Crédito").Range("G118").Value = ws_visainf.Range("b76").Value
                wb2.Worksheets("Produtos Crédito").Range("G119").Value = ws_visainf.Range("b73").Value
                wb2.Worksheets("Produtos Crédito").Range("G120").Value = ws_visainf.Range("b75").Value
                wb2.Worksheets("Produtos Crédito").Range("G121").Value = ws_visainf.Range("b77").Value

                wb2.Worksheets("Produtos Crédito").Range("G174").Value = ws_visainf.Range("b87").Value
                wb2.Worksheets("Produtos Crédito").Range("G146").Value = ws_visainf.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("G145").Value = ws_visainf.Range("b91").Value
                wb2.Worksheets("Produtos Crédito").Range("G172").Value = ws_visainf.Range("b108").Value
                wb2.Worksheets("Produtos Crédito").Range("G173").Value = ws_visainf.Range("b109").Value
                wb2.Worksheets("Produtos Crédito").Range("G134").Value = ws_visainf.Range("b92").Value

                'Nueva modificacion
                wb2.Worksheets("Produtos Crédito").Range("G148:G151").Value = ws_visainf.Range("b95:b98").Value
                wb2.Worksheets("Produtos Crédito").Range("G152").Value = ws_visainf.Range("b99").Value + ws_visainf.Range("b101").Value + ws_visainf.Range("b103").Value
                wb2.Worksheets("Produtos Crédito").Range("G153").Value = ws_visainf.Range("b100").Value + ws_visainf.Range("b102").Value + ws_visainf.Range("b104").Value
                wb2.Worksheets("Produtos Crédito").Range("G144").Value = ws_visainf.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("G136").Value = ws_visainf_ch.Range("b30").Value



                'CASH ADVANCES COUNT
                wb2.Worksheets("Produtos Crédito").Range("H32").Value = ws_visaplat.Range("B24").Value
                wb2.Worksheets("Produtos Crédito").Range("H33").Value = ws_visaplat.Range("B26").Value
                wb2.Worksheets("Produtos Crédito").Range("H34").Value = ws_visaplat.Range("B28").Value
                wb2.Worksheets("Produtos Crédito").Range("H35").Value = ws_visaplat.Range("B25").Value
                wb2.Worksheets("Produtos Crédito").Range("H36").Value = ws_visaplat.Range("B27").Value
                wb2.Worksheets("Produtos Crédito").Range("H37").Value = ws_visaplat.Range("B29").Value
                'NATIONALS PAYMENTS COUNH
                wb2.Worksheets("Produtos Crédito").Range("H56").Value = ws_visaplat.Range("b42").Value
                wb2.Worksheets("Produtos Crédito").Range("H57").Value = ws_visaplat.Range("b44").Value
                wb2.Worksheets("Produtos Crédito").Range("H58").Value = ws_visaplat.Range("b46").Value
                wb2.Worksheets("Produtos Crédito").Range("H59").Value = ws_visaplat.Range("b43").Value
                wb2.Worksheets("Produtos Crédito").Range("H60").Value = ws_visaplat.Range("b45").Value
                wb2.Worksheets("Produtos Crédito").Range("H61").Value = ws_visaplat.Range("b47").Value
                'NATIONAL PAYMENTS COUNH
                wb2.Worksheets("Produtos Crédito").Range("H74").Value = ws_visaplat.Range("b54").Value
                wb2.Worksheets("Produtos Crédito").Range("H75").Value = ws_visaplat.Range("b56").Value
                wb2.Worksheets("Produtos Crédito").Range("H76").Value = ws_visaplat.Range("b58").Value
                wb2.Worksheets("Produtos Crédito").Range("H77").Value = ws_visaplat.Range("b55").Value
                wb2.Worksheets("Produtos Crédito").Range("H78").Value = ws_visaplat.Range("b57").Value
                wb2.Worksheets("Produtos Crédito").Range("H79").Value = ws_visaplat.Range("b59").Value
                'INTERNATIONAL PAYMENTS COUNH
                wb2.Worksheets("Produtos Crédito").Range("H98").Value = ws_visaplat.Range("b60").Value
                wb2.Worksheets("Produtos Crédito").Range("H99").Value = ws_visaplat.Range("b62").Value
                wb2.Worksheets("Produtos Crédito").Range("H100").Value = ws_visaplat.Range("b64").Value
                wb2.Worksheets("Produtos Crédito").Range("H101").Value = ws_visaplat.Range("b61").Value
                wb2.Worksheets("Produtos Crédito").Range("H102").Value = ws_visaplat.Range("b63").Value
                wb2.Worksheets("Produtos Crédito").Range("H103").Value = ws_visaplat.Range("b65").Value
                'INTERANTIONAL ATM CASH ADVACNEH
                wb2.Worksheets("Produtos Crédito").Range("H116").Value = ws_visaplat.Range("b72").Value
                wb2.Worksheets("Produtos Crédito").Range("H117").Value = ws_visaplat.Range("b74").Value
                wb2.Worksheets("Produtos Crédito").Range("H118").Value = ws_visaplat.Range("b76").Value
                wb2.Worksheets("Produtos Crédito").Range("H119").Value = ws_visaplat.Range("b73").Value
                wb2.Worksheets("Produtos Crédito").Range("H120").Value = ws_visaplat.Range("b75").Value
                wb2.Worksheets("Produtos Crédito").Range("H121").Value = ws_visaplat.Range("b77").Value


                wb2.Worksheets("Produtos Crédito").Range("H174").Value = ws_visaplat.Range("b87").Value
                wb2.Worksheets("Produtos Crédito").Range("H146").Value = ws_visaplat.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("H145").Value = ws_visaplat.Range("b91").Value
                wb2.Worksheets("Produtos Crédito").Range("H172").Value = ws_visaplat.Range("b108").Value
                wb2.Worksheets("Produtos Crédito").Range("H173").Value = ws_visaplat.Range("b109").Value
                wb2.Worksheets("Produtos Crédito").Range("H134").Value = ws_visaplat.Range("b92").Value

                'Nueva modificacion
                wb2.Worksheets("Produtos Crédito").Range("H148:H151").Value = ws_visaplat.Range("b95:b98").Value
                wb2.Worksheets("Produtos Crédito").Range("H152").Value = ws_visaplat.Range("b99").Value + ws_visaplat.Range("b101").Value + ws_visaplat.Range("b103").Value
                wb2.Worksheets("Produtos Crédito").Range("H153").Value = ws_visaplat.Range("b100").Value + ws_visaplat.Range("b102").Value + ws_visaplat.Range("b104").Value
                wb2.Worksheets("Produtos Crédito").Range("H144").Value = ws_visaplat.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("H136").Value = ws_visaplat_ch.Range("b30").Value


                'CASH ADVANCES COUNT
                wb2.Worksheets("Produtos Crédito").Range("I32").Value = ws_visasig.Range("B24").Value
                wb2.Worksheets("Produtos Crédito").Range("I33").Value = ws_visasig.Range("B26").Value
                wb2.Worksheets("Produtos Crédito").Range("I34").Value = ws_visasig.Range("B28").Value
                wb2.Worksheets("Produtos Crédito").Range("I35").Value = ws_visasig.Range("B25").Value
                wb2.Worksheets("Produtos Crédito").Range("I36").Value = ws_visasig.Range("B27").Value
                wb2.Worksheets("Produtos Crédito").Range("I37").Value = ws_visasig.Range("B29").Value
                'NATIONALS PAYMENTS COUNH
                wb2.Worksheets("Produtos Crédito").Range("I56").Value = ws_visasig.Range("b42").Value
                wb2.Worksheets("Produtos Crédito").Range("I57").Value = ws_visasig.Range("b44").Value
                wb2.Worksheets("Produtos Crédito").Range("I58").Value = ws_visasig.Range("b46").Value
                wb2.Worksheets("Produtos Crédito").Range("I59").Value = ws_visasig.Range("b43").Value
                wb2.Worksheets("Produtos Crédito").Range("I60").Value = ws_visasig.Range("b45").Value
                wb2.Worksheets("Produtos Crédito").Range("I61").Value = ws_visasig.Range("b47").Value
                'NATIONAL PAYMENTS COUNH
                wb2.Worksheets("Produtos Crédito").Range("I74").Value = ws_visasig.Range("b54").Value
                wb2.Worksheets("Produtos Crédito").Range("I75").Value = ws_visasig.Range("b56").Value
                wb2.Worksheets("Produtos Crédito").Range("I76").Value = ws_visasig.Range("b58").Value
                wb2.Worksheets("Produtos Crédito").Range("I77").Value = ws_visasig.Range("b55").Value
                wb2.Worksheets("Produtos Crédito").Range("I78").Value = ws_visasig.Range("b57").Value
                wb2.Worksheets("Produtos Crédito").Range("I79").Value = ws_visasig.Range("b59").Value
                'INTERNATIONAL PAYMENTS COUNH
                wb2.Worksheets("Produtos Crédito").Range("I98").Value = ws_visasig.Range("b60").Value
                wb2.Worksheets("Produtos Crédito").Range("I99").Value = ws_visasig.Range("b62").Value
                wb2.Worksheets("Produtos Crédito").Range("I100").Value = ws_visasig.Range("b64").Value
                wb2.Worksheets("Produtos Crédito").Range("I101").Value = ws_visasig.Range("b61").Value
                wb2.Worksheets("Produtos Crédito").Range("I102").Value = ws_visasig.Range("b63").Value
                wb2.Worksheets("Produtos Crédito").Range("I103").Value = ws_visasig.Range("b65").Value
                'INTERANTIONAL ATM CASH ADVACNEH
                wb2.Worksheets("Produtos Crédito").Range("I116").Value = ws_visasig.Range("b72").Value
                wb2.Worksheets("Produtos Crédito").Range("I117").Value = ws_visasig.Range("b74").Value
                wb2.Worksheets("Produtos Crédito").Range("I118").Value = ws_visasig.Range("b76").Value
                wb2.Worksheets("Produtos Crédito").Range("I119").Value = ws_visasig.Range("b73").Value
                wb2.Worksheets("Produtos Crédito").Range("I120").Value = ws_visasig.Range("b75").Value
                wb2.Worksheets("Produtos Crédito").Range("I121").Value = ws_visasig.Range("b77").Value

                wb2.Worksheets("Produtos Crédito").Range("I174").Value = ws_visasig.Range("b87").Value
                wb2.Worksheets("Produtos Crédito").Range("I146").Value = ws_visasig.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("I145").Value = ws_visasig.Range("b91").Value
                wb2.Worksheets("Produtos Crédito").Range("I172").Value = ws_visasig.Range("b108").Value
                wb2.Worksheets("Produtos Crédito").Range("I173").Value = ws_visasig.Range("b109").Value
                wb2.Worksheets("Produtos Crédito").Range("I134").Value = ws_visasig.Range("b92").Value

                'Nueva modificacion
                wb2.Worksheets("Produtos Crédito").Range("I148:I151").Value = ws_visasig.Range("b95:b98").Value
                wb2.Worksheets("Produtos Crédito").Range("I152").Value = ws_visasig.Range("b99").Value + ws_visasig.Range("b101").Value + ws_visasig.Range("b103").Value
                wb2.Worksheets("Produtos Crédito").Range("I153").Value = ws_visasig.Range("b100").Value + ws_visasig.Range("b102").Value + ws_visasig.Range("b104").Value
                wb2.Worksheets("Produtos Crédito").Range("I144").Value = ws_visasig.Range("b90").Value
                wb2.Worksheets("Produtos Crédito").Range("I136").Value = ws_visasig_ch.Range("b30").Value
                
                


                'PRODUCTOS DEBITO

                'CASH ADVANCES COUNT
                wb2.Worksheets("Produtos Débito").Range("C32").Value = ws_visauclsd.Range("B24").Value
                wb2.Worksheets("Produtos Débito").Range("C33").Value = ws_visauclsd.Range("B26").Value
                wb2.Worksheets("Produtos Débito").Range("C34").Value = ws_visauclsd.Range("B28").Value
                wb2.Worksheets("Produtos Débito").Range("C35").Value = ws_visauclsd.Range("B25").Value
                wb2.Worksheets("Produtos Débito").Range("C36").Value = ws_visauclsd.Range("B27").Value
                wb2.Worksheets("Produtos Débito").Range("C37").Value = ws_visauclsd.Range("B29").Value

                wb2.Worksheets("Produtos Débito").Range("C56").Value = ws_visauclsd.Range("b42").Value
                wb2.Worksheets("Produtos Débito").Range("C57").Value = ws_visauclsd.Range("b44").Value
                wb2.Worksheets("Produtos Débito").Range("C58").Value = ws_visauclsd.Range("b46").Value
                wb2.Worksheets("Produtos Débito").Range("C59").Value = ws_visauclsd.Range("b43").Value
                wb2.Worksheets("Produtos Débito").Range("C60").Value = ws_visauclsd.Range("b45").Value
                wb2.Worksheets("Produtos Débito").Range("C61").Value = ws_visauclsd.Range("b47").Value

                wb2.Worksheets("Produtos Débito").Range("C68").Value = ws_visauclsd.Range("b54").Value
                wb2.Worksheets("Produtos Débito").Range("C69").Value = ws_visauclsd.Range("b56").Value
                wb2.Worksheets("Produtos Débito").Range("C70").Value = ws_visauclsd.Range("b58").Value
                wb2.Worksheets("Produtos Débito").Range("C71").Value = ws_visauclsd.Range("b55").Value
                wb2.Worksheets("Produtos Débito").Range("C72").Value = ws_visauclsd.Range("b59").Value
                wb2.Worksheets("Produtos Débito").Range("C73").Value = ws_visauclsd.Range("b60").Value

                wb2.Worksheets("Produtos Débito").Range("C86").Value = ws_visauclsd.Range("b60").Value
                wb2.Worksheets("Produtos Débito").Range("C87").Value = ws_visauclsd.Range("b62").Value
                wb2.Worksheets("Produtos Débito").Range("C88").Value = ws_visauclsd.Range("b64").Value
                wb2.Worksheets("Produtos Débito").Range("C89").Value = ws_visauclsd.Range("b61").Value
                wb2.Worksheets("Produtos Débito").Range("C90").Value = ws_visauclsd.Range("b63").Value
                wb2.Worksheets("Produtos Débito").Range("C91").Value = ws_visauclsd.Range("b65").Value

                wb2.Worksheets("Produtos Débito").Range("C104").Value = ws_visauclsd.Range("b72").Value
                wb2.Worksheets("Produtos Débito").Range("C105").Value = ws_visauclsd.Range("b74").Value
                wb2.Worksheets("Produtos Débito").Range("C106").Value = ws_visauclsd.Range("b76").Value
                wb2.Worksheets("Produtos Débito").Range("C107").Value = ws_visauclsd.Range("b73").Value
                wb2.Worksheets("Produtos Débito").Range("C108").Value = ws_visauclsd.Range("b75").Value
                wb2.Worksheets("Produtos Débito").Range("C109").Value = ws_visauclsd.Range("b77").Value


                wb2.Worksheets("Produtos Débito").Range("C134").Value = ws_visauclsd.Range("b78").Value
                wb2.Worksheets("Produtos Débito").Range("C122").Value = ws_visauclsd.Range("b80").Value
                wb2.Worksheets("Produtos Débito").Range("C135").Value = ws_visauclsd.Range("b81").Value
                'wb2.Worksheets("Produtos Débito").Range("B138").Value = ws_visauclsd.Range("b83").Value
                'wb2.Worksheets("Produtos Débito").Range("B137").Value = ws_visauclsd.Range("b85").Value

                'Nueva modificacion
                wb2.Worksheets("Produtos Débito").Range("C132").Value = ws_visasig.Range("b78").Value
                wb2.Worksheets("Produtos Débito").Range("C124").Value = ws_visasig_ch.Range("b30").Value


                'CASH ADVANCES COUNT
                wb2.Worksheets("Produtos Débito").Range("B32").Value = ws_visadebemp.Range("B24").Value
                wb2.Worksheets("Produtos Débito").Range("B33").Value = ws_visadebemp.Range("B26").Value
                wb2.Worksheets("Produtos Débito").Range("B34").Value = ws_visadebemp.Range("B28").Value
                wb2.Worksheets("Produtos Débito").Range("B35").Value = ws_visadebemp.Range("B25").Value
                wb2.Worksheets("Produtos Débito").Range("B36").Value = ws_visadebemp.Range("B27").Value
                wb2.Worksheets("Produtos Débito").Range("B37").Value = ws_visadebemp.Range("B29").Value

                wb2.Worksheets("Produtos Débito").Range("B56").Value = ws_visadebemp.Range("b42").Value
                wb2.Worksheets("Produtos Débito").Range("B57").Value = ws_visadebemp.Range("b44").Value
                wb2.Worksheets("Produtos Débito").Range("B58").Value = ws_visadebemp.Range("b46").Value
                wb2.Worksheets("Produtos Débito").Range("B59").Value = ws_visadebemp.Range("b43").Value
                wb2.Worksheets("Produtos Débito").Range("B60").Value = ws_visadebemp.Range("b45").Value
                wb2.Worksheets("Produtos Débito").Range("B61").Value = ws_visadebemp.Range("b47").Value

                wb2.Worksheets("Produtos Débito").Range("B68").Value = ws_visadebemp.Range("b54").Value
                wb2.Worksheets("Produtos Débito").Range("B69").Value = ws_visadebemp.Range("b56").Value
                wb2.Worksheets("Produtos Débito").Range("B70").Value = ws_visadebemp.Range("b58").Value
                wb2.Worksheets("Produtos Débito").Range("B71").Value = ws_visadebemp.Range("b55").Value
                wb2.Worksheets("Produtos Débito").Range("B72").Value = ws_visadebemp.Range("b59").Value
                wb2.Worksheets("Produtos Débito").Range("B73").Value = ws_visadebemp.Range("b60").Value

                wb2.Worksheets("Produtos Débito").Range("B86").Value = ws_visadebemp.Range("b60").Value
                wb2.Worksheets("Produtos Débito").Range("B87").Value = ws_visadebemp.Range("b62").Value
                wb2.Worksheets("Produtos Débito").Range("B88").Value = ws_visadebemp.Range("b64").Value
                wb2.Worksheets("Produtos Débito").Range("B89").Value = ws_visadebemp.Range("b61").Value
                wb2.Worksheets("Produtos Débito").Range("B90").Value = ws_visadebemp.Range("b63").Value
                wb2.Worksheets("Produtos Débito").Range("B91").Value = ws_visadebemp.Range("b65").Value

                wb2.Worksheets("Produtos Débito").Range("B104").Value = ws_visadebemp.Range("b72").Value
                wb2.Worksheets("Produtos Débito").Range("B105").Value = ws_visadebemp.Range("b74").Value
                wb2.Worksheets("Produtos Débito").Range("B106").Value = ws_visadebemp.Range("b76").Value
                wb2.Worksheets("Produtos Débito").Range("B107").Value = ws_visadebemp.Range("b73").Value
                wb2.Worksheets("Produtos Débito").Range("B108").Value = ws_visadebemp.Range("b75").Value
                wb2.Worksheets("Produtos Débito").Range("B109").Value = ws_visadebemp.Range("b77").Value

                wb2.Worksheets("Produtos Débito").Range("B134").Value = ws_visadebemp.Range("b78").Value
                wb2.Worksheets("Produtos Débito").Range("B122").Value = ws_visadebemp.Range("b80").Value
                wb2.Worksheets("Produtos Débito").Range("B135").Value = ws_visadebemp.Range("b81").Value
                'wb2.Worksheets("Produtos Débito").Range("C138").Value = ws_visadebemp.Range("b83").Value
                'wb2.Worksheets("Produtos Débito").Range("C137").Value = ws_visadebemp.Range("b85").Value

                'Nueva modificacion
                wb2.Worksheets("Produtos Débito").Range("B132").Value = ws_visasig.Range("b78").Value
                wb2.Worksheets("Produtos Débito").Range("B124").Value = ws_visasig_ch.Range("b30").Value

            
                'PRODUCTOS PREPAGADO
                wb2.Worksheets("Produtos Prepagados").Range("B50").Value = ws_visagift.Range("B42").Value
                wb2.Worksheets("Produtos Prepagados").Range("B51").Value = ws_visagift.Range("B44").Value
                wb2.Worksheets("Produtos Prepagados").Range("B52").Value = ws_visagift.Range("B46").Value
                wb2.Worksheets("Produtos Prepagados").Range("B53").Value = ws_visagift.Range("B43").Value
                wb2.Worksheets("Produtos Prepagados").Range("B54").Value = ws_visagift.Range("B45").Value
                wb2.Worksheets("Produtos Prepagados").Range("B55").Value = ws_visagift.Range("B47").Value
                
                wb2.Worksheets("Produtos Prepagados").Range("B86").Value = ws_visagift.Range("b60").Value
                wb2.Worksheets("Produtos Prepagados").Range("B87").Value = ws_visagift.Range("b62").Value
                wb2.Worksheets("Produtos Prepagados").Range("B88").Value = ws_visagift.Range("b64").Value
                wb2.Worksheets("Produtos Prepagados").Range("B89").Value = ws_visagift.Range("b61").Value
                wb2.Worksheets("Produtos Prepagados").Range("B90").Value = ws_visagift.Range("b63").Value
                wb2.Worksheets("Produtos Prepagados").Range("B91").Value = ws_visagift.Range("b65").Value
                wb2.Worksheets("Produtos Prepagados").Range("B135").Value = ws_visagift.Range("B81").Value

                'Nueva modificacion
                wb2.Worksheets("Produtos Prepagados").Range("B132").Value = ws_visasig.Range("b78").Value
                wb2.Worksheets("Produtos Prepagados").Range("B124").Value = ws_visasig_ch.Range("b30").Value


                'LLENADO DE ADQUIRIENTES

                wb2.Worksheets("Adquirencia").Range("C74").Value = ws_adquiCredito.Range("b64").Value
                wb2.Worksheets("Adquirencia").Range("C75").Value = ws_adquiCredito.Range("b66").Value
                wb2.Worksheets("Adquirencia").Range("C76").Value = ws_adquiCredito.Range("b68").Value
                wb2.Worksheets("Adquirencia").Range("C77").Value = Round(ws_adquiCredito.Range("b65").Value,2)
                wb2.Worksheets("Adquirencia").Range("C78").Value = Round(ws_adquiCredito.Range("b67").Value,2)
                wb2.Worksheets("Adquirencia").Range("C79").Value = Round(ws_adquiCredito.Range("b69").Value,2)

                Wb2.Worksheets("Adquirencia").Range("C116").Value = ws_adquiCredito.Range("b88").Value
                wb2.Worksheets("Adquirencia").Range("C117").Value = ws_adquiCredito.Range("b90").Value
                wb2.Worksheets("Adquirencia").Range("C118").Value = ws_adquiCredito.Range("b92").Value
                wb2.Worksheets("Adquirencia").Range("C119").Value = Round(ws_adquiCredito.Range("b89").Value,2)
                wb2.Worksheets("Adquirencia").Range("C120").Value = Round(ws_adquiCredito.Range("b91").Value,2)
                wb2.Worksheets("Adquirencia").Range("C121").Value = Round(ws_adquiCredito.Range("b93").Value,2)


                wb2.Worksheets("Adquirencia").Range("D74").Value = ws_adquiDebiPre.Range("b64").Value
                wb2.Worksheets("Adquirencia").Range("D75").Value = ws_adquiDebiPre.Range("b66").Value
                wb2.Worksheets("Adquirencia").Range("D76").Value = ws_adquiDebiPre.Range("b68").Value
                wb2.Worksheets("Adquirencia").Range("D77").Value = Round(ws_adquiDebiPre.Range("b65").Value,2)
                wb2.Worksheets("Adquirencia").Range("D78").Value = Round(ws_adquiDebiPre.Range("b67").Value,2)
                wb2.Worksheets("Adquirencia").Range("D79").Value = Round(ws_adquiDebiPre.Range("b69").Value,2)
               

                Wb2.Worksheets("Adquirencia").Range("D116").Value = ws_adquiDebiPre.Range("b88").Value
                wb2.Worksheets("Adquirencia").Range("D117").Value = ws_adquiDebiPre.Range("b90").Value
                wb2.Worksheets("Adquirencia").Range("D118").Value = ws_adquiDebiPre.Range("b92").Value
                wb2.Worksheets("Adquirencia").Range("D119").Value = Round(ws_adquiDebiPre.Range("b89").Value,2)
                wb2.Worksheets("Adquirencia").Range("D120").Value = Round(ws_adquiDebiPre.Range("b91").Value,2)
                wb2.Worksheets("Adquirencia").Range("D121").Value = Round(ws_adquiDebiPre.Range("b93").Value,2)


            

                Dim rng As Range
                
                Set ws1 = wb2.Worksheets("Produtos Crédito")
                Set ws2 = wb2.Worksheets("Produtos Débito")
                Set ws3 = wb2.Worksheets("Produtos Prepagados")
                Set ws4 = wb2.Worksheets("Adquirencia")
                
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

                For Each rng In ws4.Range("C56:E133")
                    
                    If IsEmpty(rng) Then
                        rng.Value = 0
                    End If

                            
                Next

        
                For Each rng In ws4.Range("B134:B169")

                    If IsEmpty(rng) Then
                        rng.Value = 0
                    End If
                Next
              
                
                
  
                ' *********************
                
                
                ' *********************
                
                ' Example: print the path of the selected file to the immediate window
                'Debug.Print strFilePath ' remove in production
            End If
        End With
    End Sub







