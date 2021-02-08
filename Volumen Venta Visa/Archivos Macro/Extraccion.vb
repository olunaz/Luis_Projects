Sub CopiadoValoresVol(sheetname, contador, VolVentas)


  '          VOLUMEN DE VENTAS
  '          EXTRACION DE ARRAY #
  '
  Dim i As Integer, j As Integer, c As Integer, d As Integer, z As Integer
          z = contador + 6
          i = 18
          Do While (i < 30)
              j = 4
              Do While (i < 30)
                If sheetname.Cells(i, 2) = "" Then
                  sheetname.Cells(i, 2).Value = VolVentas.Cells(z, j).Value
                    i = i + 2
                End If
                j = j - 1
                
                If j = 1 Then
                
                  z = z + 2
                  j = j + 3
                
                End If
                                
              Loop
          Loop
  
  '          EXTRACION DE ARRAY (LC)

          z = contador + 7
          i = 19
          Do While (i < 30)
              j = 4
              Do While (i < 30)
                If sheetname.Cells(i, 2) = "" Then
                  sheetname.Cells(i, 2).Value = VolVentas.Cells(z, j).Value
                    i = i + 2
                End If
                j = j - 1
                
                If j = 1 Then
                
                  z = z + 2
                  j = j + 3
                
                End If
                            
              Loop
          Loop
          
          
          
  '''''''''''COMPRAS ENTRASTES NACIONAL
          
  '          EXTRACION DE ARRAY #
  '
                  
          z = contador + 13
          i = 42
          Do While (i < 60)
              j = 4
              Do While (i < 60)
                If sheetname.Cells(i, 2) = "" Then
                  sheetname.Cells(i, 2).Value = VolVentas.Cells(z, j).Value
                    i = i + 2
                End If
                j = j - 1
                
                If j = 1 Then
                
                  z = z + 4
                  j = j + 3
                  i = i + 6
                
                End If
                                
              Loop
          Loop
          
          i = i + 6
          
          
  '          EXTRACION DE ARRAY (LC)

          z = contador + 14
          i = 43
          Do While (i < 60)
              j = 4
              Do While (i < 60)
                If sheetname.Cells(i, 2) = "" Then
                  sheetname.Cells(i, 2).Value = VolVentas.Cells(z, j).Value
                    i = i + 2
                End If
                j = j - 1
                
                If j = 1 Then
                
                  z = z + 4
                  j = j + 3
                  i = i + 6
                
                End If
                            
              Loop
          Loop
          
  '''''''''''COMPRAS ENTRASTES INTERNACIONAL

  '          EXTRACION DE ARRAY #
  '
          z = contador + 22
          i = 60
          Do While (i < 78)
              j = 4
              Do While (i < 78)
                If sheetname.Cells(i, 2) = "" Then
                  sheetname.Cells(i, 2).Value = VolVentas.Cells(z, j).Value
                    i = i + 2
                End If
                j = j - 1

                If j = 1 Then

                  z = z + 4
                  j = j + 3
                  i = i + 6
                

                End If

              Loop
          Loop


   '          EXTRACION DE ARRAY (LC)

          z = contador + 23
          i = 61
          Do While (i < 78)
              j = 4
              Do While (i < 78)
                If sheetname.Cells(i, 2) = "" Then
                  sheetname.Cells(i, 2).Value = VolVentas.Cells(z, j).Value
                    i = i + 2
                End If
                j = j - 1

                If j = 1 Then

                  z = z + 4
                  j = j + 3
                  i = i + 6
                

                End If

              Loop
          Loop

End Sub


'''''''''''''''''''''''''''''''''''''
Sub CopiadoValoresInfDebito(sheetname, contador, InfCuentas)

 '''DATOS FIJOS
  'Finance Charges
  Dim i As Integer, j As Integer, c As Integer, d As Integer, z As Integer
  z = contador

 '''DATOS FIJOS
    'Number of International Accounts
    sheetname.Cells(78, 2).Value = InfCuentas.Cells(z + 2, 2).Value
    'Number of Restricted Accounts
    sheetname.Cells(79, 2).Value = InfCuentas.Cells(z + 3, 2).Value
    'Number of CardHolders
    sheetname.Cells(80, 2).Value = InfCuentas.Cells(z + 5, 2).Value
    'Number of Active Accounts
    sheetname.Cells(81, 2).Value = InfCuentas.Cells(z + 8, 2).Value
    'Number of Accounts Accessed
    'Regular Checking Account offered
    'Special Checking Account offered
    'Savings Account offered
    'Other Accounts offered
    'Posting Method
    'Number of Fraud Loss Transactions
    'Gross Fraud Losses (LC)
    'Recovered Fraud Losses (LC)
    'Number of Charge-Offs transaction
    'Gross Charge-Offs
    'Recovered Charge-Offs
    'Cards with POS activity (#)
    sheetname.Cells(94, 2).Value = InfCuentas.Cells(z + 8, 2).Value
    'Cards with ATM activity (#)
    sheetname.Cells(95, 2).Value = InfCuentas.Cells(z + 11, 2).Value
    'POS transactions decline - Insufficient funds (#)
    sheetname.Cells(96, 2).Value = InfCuentas.Cells(z + 12, 2).Value
    'POS transactions decline - pick-up (#)
    sheetname.Cells(97, 2).Value = InfCuentas.Cells(z + 13, 2).Value
    'POS transactions decline - other reasons (#)
    sheetname.Cells(98, 2).Value = InfCuentas.Cells(z + 14, 2).Value
    'ATM transactions decline - Insufficient funds (#)
    sheetname.Cells(99, 2).Value = InfCuentas.Cells(z + 15, 2).Value
    'ATM transactions decline - pick-up (#)
    sheetname.Cells(100, 2).Value = InfCuentas.Cells(z + 16, 2).Value
    'ATM transactions decline - other reasons (#)
    sheetname.Cells(101, 2).Value = InfCuentas.Cells(z + 17, 2).Value
    'Visa Mini "Companion"(#)
    'Visa Mini "Stand Alone"(#)
    'Number of cards receiving Remittances(#)
    'Number of Remittances Received (#)
    'Amount of Remittances Received(LC)
    'Number of cards "Debit Companion"(#)
    'Number of Accounts "Debit Companion"(#)

          
  'LLENADO DE CELDAS CON 0


  sheetname.Range("B6:B108").SpecialCells(xlCellTypeBlanks) = 0
  
End Sub

''''''''''''''''''''''''''''''''''''''''
Sub CopiadoValoresInfRegalo(sheetname, contador, InfCuentas)

 '''DATOS FIJOS
  'Finance Charges
  Dim i As Integer, j As Integer, c As Integer, d As Integer, z As Integer
  z = contador

 '''DATOS FIJOS
    'Number of International Accounts
    sheetname.Cells(78, 2).Value = InfCuentas.Cells(z + 2, 2).Value

    'Number of Accounts - International Classic with Visa Flag
    'Number of Restricted Accounts
    'Number of CardHolders
    sheetname.Cells(80, 2).Value = InfCuentas.Cells(z + 5, 2).Value
    'Number of Active Accounts
    sheetname.Cells(81, 2).Value = InfCuentas.Cells(z + 8, 2).Value
    'Number of Accounts Accessed
    'Posting Method
    'Number of Fraud Loss Transactions
    'Gross Fraud Losses (LC)
    'Recovered Fraud Losses (LC)
    'Number of Charge-Offs transaction
    'Gross Charge-Offs
    'Recovered Charge-Offs
    'Cards with POS activity (#)
    sheetname.Cells(90, 2).Value = InfCuentas.Cells(z + 8, 2).Value
    'Cards with ATM activity (#)
    sheetname.Cells(90, 2).Value = InfCuentas.Cells(z + 13, 2).Value
    'POS transactions decline - Insufficient funds (#)
    sheetname.Cells(92, 2).Value = InfCuentas.Cells(z + 12, 2).Value
    'POS transactions decline - pick-up (#)
    sheetname.Cells(93, 2).Value = InfCuentas.Cells(z + 13, 2).Value
    'POS transactions decline - other reasons (#)
    sheetname.Cells(94, 2).Value = InfCuentas.Cells(z + 14, 2).Value
    'ATM transactions decline - Insufficient funds (#)
    sheetname.Cells(95, 2).Value = InfCuentas.Cells(z + 15, 2).Value
    'ATM transactions decline - pick-up (#)
    sheetname.Cells(96, 2).Value = InfCuentas.Cells(z + 16, 2).Value
    'ATM transactions decline - other reasons (#)
    sheetname.Cells(97, 2).Value = InfCuentas.Cells(z + 17, 2).Value
    'Visa Mini "Companion"(#)
    'Visa Mini "Stand Alone"(#)
    'Prepaid Cardholders Balance (LC)
    'Number of Accounts - International without Visa Flag(#)
    'Loads(#)
    sheetname.Cells(102, 2).Value = InfCuentas.Cells(z + 8, 2).Value
    'Loads(LC)
    sheetname.Cells(103, 2).Value = InfCuentas.Cells(z + 8, 2).Value
    'Refunds, Fees charged to Balance & Breakage(#)
    'Refunds, Fees charged to Balance & Breakage(LC)

          
  'LLENADO DE CELDAS CON 0


  sheetname.Range("B6:B105").SpecialCells(xlCellTypeBlanks) = 0
  
End Sub
'''''''''''''''''''''''''''''''''''''


Sub CopiadoValoresInf(sheetname, contador, InfCuentas)

 '''DATOS FIJOS
  
  Dim i As Integer, j As Integer, c As Integer, d As Integer, z As Integer
  z = contador
  
  
  'Finance Charges
  sheetname.Cells(78, 2).Value = InfCuentas.Cells(z + 30, 5).Value

  'Anual Fee, Late Charges and Other Changes (LC)
  sheetname.Cells(79, 2).Value = InfCuentas.Cells(z + 31, 5).Value

  'Credit Vouchers
  sheetname.Cells(88, 2).Value = InfCuentas.Cells(z + 33, 5).Value

  'Other Credits
  sheetname.Cells(89, 2).Value = InfCuentas.Cells(z + 34, 5).Value

  'Number of International Accounts
  sheetname.Cells(90, 2).Value = InfCuentas.Cells(z + 2, 2).Value

  'No of Accounts with POS activity (#)
  sheetname.Cells(105, 2).Value = InfCuentas.Cells(z + 8, 2).Value

  'Total Credit Line (LC)
  sheetname.Cells(108, 2).Value = InfCuentas.Cells(z + 19, 3).Value


  'Number of CardHolders
  sheetname.Cells(92, 2).Value = InfCuentas.Cells(z + 5, 2).Value


  'Number of Statements Mailed
  sheetname.Cells(93, 2).Value = InfCuentas.Cells(z + 2, 2).Value

  'Non Approved Transactions (#)
  sheetname.Cells(106, 2).Value = InfCuentas.Cells(z + 2, 2).Value

  '''ARRAY DE CUENTAS MOROSAS


  '          EXTRACION DE ARRAY (#)
  '
          z = contador + 21
          i = 95
          Do While (i < 105)
            
                If sheetname.Cells(i, 2) = "" Then
                  sheetname.Cells(i, 2).Value = InfCuentas.Cells(z, 2).Value
                    i = i + 2
                    z = z + 1
                End If
                
          Loop
          
          
  '          EXTRACION DE ARRAY (LC)

          z = contador + 21
          i = 96
          Do While (i < 105)
            
                If sheetname.Cells(i, 2) = "" Then
                  sheetname.Cells(i, 2).Value = InfCuentas.Cells(z, 3).Value
                    i = i + 2
                    z = z + 1
                End If
                
          Loop
          
  'LLENADO DE CELDAS CON 0


  sheetname.Range("B6:B111").SpecialCells(xlCellTypeBlanks) = 0
  
End Sub



Sub Formatear_Valores(sheetname, sheetnameN)


    Dim wb As Workbook
    Set wb = ThisWorkbook

    wb.Activate
  
    Dim i As Integer, j As Integer, c As Integer, d As Integer, a As Integer

   
    'ON-US Sales Volume

            a = 18
            i = 37
            Do While (i < 44)
                j = 2
                Do While (i < 44)
                 If sheetnameN.Cells(i, j) = "" Then
                   sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                   sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                     a = a + 2
                 End If
                 j = j + 1

                 If j = 5 Then

                    i = i + 4
                    j = j - 3

                 End If

                Loop
            Loop

    'National Sales Volume

                    a = 42
                    i = 46
                    Do While (i < 49)
                        j = 2
                        Do While (i < 49)
                         If sheetnameN.Cells(i, j) = "" Then
                           sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                           sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                             a = a + 2
                         End If
                         j = j + 1

                         If j = 5 Then

                            i = i + 4
                            j = j - 3

                         End If

                        Loop
                    Loop

                    'Segunda Parte
                    a = 54
                    i = 72
                    Do While (i < 75)
                        j = 2
                        Do While (i < 75)
                         If sheetnameN.Cells(i, j) = "" Then
                           sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                           sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                             a = a + 2
                         End If
                         j = j + 1

                         If j = 5 Then

                            i = i + 4
                            j = j - 3

                         End If

                        Loop
                    Loop

     'International Sales Volume

                    a = 60
                    i = 77
                    Do While (i < 80)
                        j = 2
                        Do While (i < 80)
                         If sheetnameN.Cells(i, j) = "" Then
                           sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                           sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                             a = a + 2
                         End If
                         j = j + 1

                         If j = 5 Then

                            i = i + 4
                            j = j - 3

                         End If

                        Loop
                    Loop

                    'Segunda Parte
                    a = 72
                    i = 101
                    Do While (i < 104)
                        j = 2
                        Do While (i < 104)
                         If sheetnameN.Cells(i, j) = "" Then
                           sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                           sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                             a = a + 2
                         End If
                         j = j + 1

                         If j = 5 Then

                            i = i + 4
                            j = j - 3

                         End If

                        Loop
                    Loop

        'Cantidades Fijas

        sheetnameN.Cells(106, 2).Value = sheetname.Cells(78, 2).Value
        sheetnameN.Cells(107, 2).Value = sheetname.Cells(109, 2).Value
        sheetnameN.Cells(117, 2).Value = sheetname.Cells(88, 2).Value
        sheetnameN.Cells(118, 2).Value = sheetname.Cells(89, 2).Value
        sheetnameN.Cells(125, 2).Value = sheetname.Cells(90, 2).Value
        sheetnameN.Cells(126, 2).Value = sheetname.Cells(91, 2).Value
        sheetnameN.Cells(132, 2).Value = sheetname.Cells(105, 2).Value
        sheetnameN.Cells(133, 2).Value = sheetname.Cells(108, 2).Value
        sheetnameN.Cells(134, 2).Value = sheetname.Cells(109, 2).Value
        sheetnameN.Cells(138, 2).Value = sheetname.Cells(106, 2).Value
        sheetnameN.Cells(139, 2).Value = sheetname.Cells(107, 2).Value

        'MATRIZ DE TARJETAS ROBADAS


                    a = 95
                    i = 143
                    Do While (i < 148)
                         j = 2

                         If sheetnameN.Cells(i, j) = "" Then
                           sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                           sheetnameN.Cells(i, j + 1).Value = sheetname.Cells(a + 1, 2).Value
                             a = a + 2
                         End If

                        i = i + 1

                    Loop


 'LLENADO DE CELDAS CON 0


    sheetnameN.Range("B8:E43").SpecialCells(xlCellTypeBlanks) = 0
    sheetnameN.Range("B45:E74").SpecialCells(xlCellTypeBlanks) = 0
    sheetnameN.Range("B76:E103").SpecialCells(xlCellTypeBlanks) = 0


End Sub

'''''''''''''''''''''''''''''''''
Sub Formatear_ValoresDebito(sheetname, sheetnameN)


    Dim wb As Workbook
    Set wb = ThisWorkbook

    wb.Activate
  
    Dim i As Integer, j As Integer, c As Integer, d As Integer, a As Integer

   
    'ON-US Sales Volume

            a = 18
            i = 37
            Do While (i < 44)
                j = 2
                Do While (i < 44)
                 If sheetnameN.Cells(i, j) = "" Then
                   sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                   sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                     a = a + 2
                 End If
                 j = j + 1

                 If j = 5 Then

                    i = i + 4
                    j = j - 3

                 End If

                Loop
            Loop

    'National Sales Volume

                    a = 42
                    i = 46
                    Do While (i < 49)
                        j = 2
                        Do While (i < 49)
                         If sheetnameN.Cells(i, j) = "" Then
                           sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                           sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                             a = a + 2
                         End If
                         j = j + 1

                         If j = 5 Then

                            i = i + 4
                            j = j - 3

                         End If

                        Loop
                    Loop

                    'Segunda Parte
                    a = 54
                    i = 72
                    Do While (i < 75)
                        j = 2
                        Do While (i < 75)
                         If sheetnameN.Cells(i, j) = "" Then
                           sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                           sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                             a = a + 2
                         End If
                         j = j + 1

                         If j = 5 Then

                            i = i + 4
                            j = j - 3

                         End If

                        Loop
                    Loop

     'International Sales Volume

                    a = 60
                    i = 77
                    Do While (i < 80)
                        j = 2
                        Do While (i < 80)
                         If sheetnameN.Cells(i, j) = "" Then
                           sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                           sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                             a = a + 2
                         End If
                         j = j + 1

                         If j = 5 Then

                            i = i + 4
                            j = j - 3

                         End If

                        Loop
                    Loop

                    'Segunda Parte
                    a = 72
                    i = 101
                    Do While (i < 104)
                        j = 2
                        Do While (i < 104)
                         If sheetnameN.Cells(i, j) = "" Then
                           sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                           sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                             a = a + 2
                         End If
                         j = j + 1

                         If j = 5 Then

                            i = i + 4
                            j = j - 3

                         End If

                        Loop
                    Loop

        'Cantidades Fijas

        'Number of International Accounts
        sheetnameN.Cells(106, 2).Value = sheetname.Cells(78, 2).Value
        'Number of Card holders
        sheetnameN.Cells(112, 2).Value = sheetname.Cells(80, 2).Value
       'Number of active accounts
        sheetnameN.Cells(113, 2).Value = sheetname.Cells(81, 2).Value
        'Number of Cars with POS
        sheetnameN.Cells(114, 2).Value = sheetname.Cells(94, 2).Value
        'Number of Cars with ATM
        sheetnameN.Cells(115, 2).Value = sheetname.Cells(95, 2).Value

        'Total Number of Personal Deposit
        sheetnameN.Cells(116, 2).Value = sheetname.Cells(95, 2).Value

        'Purchase Transactions Declined for Insufficient Funds
        sheetnameN.Cells(123, 2).Value = sheetname.Cells(96, 2).Value

        'Total Number of Personal Declined for Pick-Up
        sheetnameN.Cells(124, 2).Value = sheetname.Cells(97, 2).Value

        'Total Number of Personal Declined for Other Reasons
        sheetnameN.Cells(125, 2).Value = sheetname.Cells(98, 2).Value

        'ATM/Cash Advance Transactions Declined Funds
        sheetnameN.Cells(126, 2).Value = sheetname.Cells(99, 2).Value

        'ATM/Cash Advance Transactions Declined for Pick-Up
        sheetnameN.Cells(127, 2).Value = sheetname.Cells(100, 2).Value

        'ATM/Cash Advance Transactions Declined for Other Reasons
        sheetnameN.Cells(128, 2).Value = sheetname.Cells(101, 2).Value


 'LLENADO DE CELDAS CON 0


    sheetnameN.Range("B8:E43").SpecialCells(xlCellTypeBlanks) = 0
    sheetnameN.Range("B45:E74").SpecialCells(xlCellTypeBlanks) = 0
    sheetnameN.Range("B76:E103").SpecialCells(xlCellTypeBlanks) = 0


End Sub

''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''
Sub Formatear_ValoresRegalo(sheetname, sheetnameN)


    Dim wb As Workbook
    Set wb = ThisWorkbook

    wb.Activate
  
    Dim i As Integer, j As Integer, c As Integer, d As Integer, a As Integer

   
    'ON-US Sales Volume

            a = 18
            i = 37
            Do While (i < 44)
                j = 2
                Do While (i < 44)
                 If sheetnameN.Cells(i, j) = "" Then
                   sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                   sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                     a = a + 2
                 End If
                 j = j + 1

                 If j = 5 Then

                    i = i + 4
                    j = j - 3

                 End If

                Loop
            Loop

    'National Sales Volume

                    a = 42
                    i = 46
                    Do While (i < 49)
                        j = 2
                        Do While (i < 49)
                         If sheetnameN.Cells(i, j) = "" Then
                           sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                           sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                             a = a + 2
                         End If
                         j = j + 1

                         If j = 5 Then

                            i = i + 4
                            j = j - 3

                         End If

                        Loop
                    Loop

                    'Segunda Parte
                    a = 54
                    i = 72
                    Do While (i < 75)
                        j = 2
                        Do While (i < 75)
                         If sheetnameN.Cells(i, j) = "" Then
                           sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                           sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                             a = a + 2
                         End If
                         j = j + 1

                         If j = 5 Then

                            i = i + 4
                            j = j - 3

                         End If

                        Loop
                    Loop

     'International Sales Volume

                    a = 60
                    i = 77
                    Do While (i < 80)
                        j = 2
                        Do While (i < 80)
                         If sheetnameN.Cells(i, j) = "" Then
                           sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                           sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                             a = a + 2
                         End If
                         j = j + 1

                         If j = 5 Then

                            i = i + 4
                            j = j - 3

                         End If

                        Loop
                    Loop

                    'Segunda Parte
                    a = 72
                    i = 101
                    Do While (i < 104)
                        j = 2
                        Do While (i < 104)
                         If sheetnameN.Cells(i, j) = "" Then
                           sheetnameN.Cells(i, j).Value = sheetname.Cells(a, 2).Value
                           sheetnameN.Cells(i + 2, j).Value = sheetname.Cells(a + 1, 2).Value

                             a = a + 2
                         End If
                         j = j + 1

                         If j = 5 Then

                            i = i + 4
                            j = j - 3

                         End If

                        Loop
                    Loop

        'Cantidades Fijas

   
        'Number of Accounts - International Classic with Visa Flag or Embossed
        sheetnameN.Cells(106, 2).Value = sheetname.Cells(78, 2).Value
        'Number of Accounts - International (Visa Electron, without Visa Flag or Unembossed with Flag)
        sheetnameN.Cells(107, 2).Value = sheetname.Cells(79, 2).Value
        'Number of Accounts - Restricted
        sheetnameN.Cells(108, 2).Value = sheetname.Cells(80, 2).Value
        'Cards Visa Mini Companion
        sheetnameN.Cells(110, 2).Value = sheetname.Cells(98, 2).Value
        'Cards Visa Mini Stand Alone
        sheetnameN.Cells(111, 2).Value = sheetname.Cells(99, 2).Value

        'Total Number of Cards
        'sheetnameN.Cells(113, 2).Value = sheetname.Cells(80, 2).Value
        'Number of active accounts - during this quarter
        'sheetnameN.Cells(114, 2).Value = sheetname.Cells(80, 2).Value

        'Number of Cards with POS Activity
        sheetnameN.Cells(115, 2).Value = sheetname.Cells(90, 2).Value
        'Number of Cards wiith ATM activity
        sheetnameN.Cells(116, 2).Value = sheetname.Cells(91, 2).Value

        ' Total number of Personal Deposit Accounts - end of this quarter
        'sheetnameN.Cells(117, 2).Value = sheetname.Cells(80, 2).Value

        'Posting Method
        sheetnameN.Cells(118, 2).Value = sheetname.Cells(83, 2).Value
        'Prepaid Cardholders Balance (LC)
        sheetnameN.Cells(119, 2).Value = sheetname.Cells(100, 2).Value
        'Loads (#)
        sheetnameN.Cells(120, 2).Value = sheetname.Cells(102, 2).Value
        'Loads (LC)
        sheetnameN.Cells(121, 2).Value = sheetname.Cells(103, 2).Value
        'Refunds, Fees charged to Balance & other Debits to Balance (#)
        sheetnameN.Cells(122, 2).Value = sheetname.Cells(104, 2).Value
        'Refunds, Fees charged to Balance & other Debits to Balance (LC)
        sheetnameN.Cells(123, 2).Value = sheetname.Cells(105, 2).Value
        'Purchase Transactions Declined for Insufficient funds (#)
        sheetnameN.Cells(125, 2).Value = sheetname.Cells(92, 2).Value
        'Purchase Transactions Declined for Pick-Up (#)
        sheetnameN.Cells(126, 2).Value = sheetname.Cells(93, 2).Value
        'Purchase Transactions Declined for Other reasons (#)
        sheetnameN.Cells(127, 2).Value = sheetname.Cells(94, 2).Value
        'ATM/Cash Advance Transactions Declined for Insufficient funds (#)
        sheetnameN.Cells(128, 2).Value = sheetname.Cells(95, 2).Value
        'ATM/Cash Advance Transactions Declined for Pick-Up (#)
        sheetnameN.Cells(129, 2).Value = sheetname.Cells(96, 2).Value
        'ATM/Cash Advance Transactions Declined for Other Reasons (#)
        sheetnameN.Cells(130, 2).Value = sheetname.Cells(97, 2).Value


 'LLENADO DE CELDAS CON 0


    sheetnameN.Range("B8:E43").SpecialCells(xlCellTypeBlanks) = 0
    sheetnameN.Range("B45:E74").SpecialCells(xlCellTypeBlanks) = 0
    sheetnameN.Range("B76:E130").SpecialCells(xlCellTypeBlanks) = 0


End Sub

''''''''''''''''''''''''''''''''


Sub Borrado_Valores(sheetname, sheetnameN)

    Dim i As Integer
    
    sheetname.Range("B7").ClearContents
    sheetname.Range("B12:B120").ClearContents

    'BLOQUE 1
    i = 9
    Do While (i < 44)

    sheetnameN.Cells(i, 2).ClearContents
    sheetnameN.Cells(i, 3).ClearContents
    sheetnameN.Cells(i, 4).ClearContents
    sheetnameN.Cells(i, 5).ClearContents

    i = i + 2

    Loop

    'BLOQUE 2
    i = 46
    Do While (i < 75)

    sheetnameN.Cells(i, 2).ClearContents
    sheetnameN.Cells(i, 3).ClearContents
    sheetnameN.Cells(i, 4).ClearContents
    sheetnameN.Cells(i, 5).ClearContents

    i = i + 2

    Loop

    'BLOQUE 3
    i = 77
    Do While (i < 104)

    sheetnameN.Cells(i, 2).ClearContents
    sheetnameN.Cells(i, 3).ClearContents
    sheetnameN.Cells(i, 4).ClearContents
    sheetnameN.Cells(i, 5).ClearContents

    i = i + 2

    Loop
    
    'BLOQUE 4
    i = 106
    Do While (i < 111)

    sheetnameN.Cells(i, 2).ClearContents
    i = i + 1

    Loop
    
    'BLOQUE 5
    i = 112
    Do While (i < 124)

    sheetnameN.Cells(i, 2).ClearContents
    sheetnameN.Cells(i, 3).ClearContents
    sheetnameN.Cells(i, 4).ClearContents
    sheetnameN.Cells(i, 5).ClearContents
    i = i + 1

    Loop
    
    'BLOQUE 6
    i = 125
    Do While (i < 142)

    sheetnameN.Cells(i, 2).ClearContents
    i = i + 1

    Loop
    
    'BLOQUE 7
    i = 143
    Do While (i < 148)

    sheetnameN.Cells(i, 2).ClearContents
    sheetnameN.Cells(i, 3).ClearContents
    i = i + 1

    Loop


End Sub

Sub Borrado_Valores_DEBREGALO(sheetname, sheetnameN)

    Dim i As Integer
    
    sheetname.Range("B7").ClearContents
    sheetname.Range("B12:B82").ClearContents
    sheetname.Range("B88:B108").ClearContents

    'BLOQUE 1
    i = 9
    Do While (i < 44)

    sheetnameN.Cells(i, 2).ClearContents
    sheetnameN.Cells(i, 3).ClearContents
    sheetnameN.Cells(i, 4).ClearContents
    sheetnameN.Cells(i, 5).ClearContents

    i = i + 2

    Loop

    'BLOQUE 2
    i = 46
    Do While (i < 75)

    sheetnameN.Cells(i, 2).ClearContents
    sheetnameN.Cells(i, 3).ClearContents
    sheetnameN.Cells(i, 4).ClearContents
    sheetnameN.Cells(i, 5).ClearContents

    i = i + 2

    Loop

    'BLOQUE 3
    i = 77
    Do While (i < 104)

    sheetnameN.Cells(i, 2).ClearContents
    sheetnameN.Cells(i, 3).ClearContents
    sheetnameN.Cells(i, 4).ClearContents
    sheetnameN.Cells(i, 5).ClearContents

    i = i + 2

    Loop
    
    'BLOQUE 4
    i = 106
    Do While (i < 111)

    sheetnameN.Cells(i, 2).ClearContents
    i = i + 1

    Loop
    
    'BLOQUE 5
    i = 112
    Do While (i < 124)

    sheetnameN.Cells(i, 2).ClearContents
    sheetnameN.Cells(i, 3).ClearContents
    sheetnameN.Cells(i, 4).ClearContents
    sheetnameN.Cells(i, 5).ClearContents
    i = i + 1

    Loop
    
    'BLOQUE 6
    i = 125
    Do While (i < 142)

    sheetnameN.Cells(i, 2).ClearContents
    i = i + 1

    Loop
    
    'BLOQUE 7
    i = 143
    Do While (i < 148)

    sheetnameN.Cells(i, 2).ClearContents
    sheetnameN.Cells(i, 3).ClearContents
    i = i + 1

    Loop


End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Borrado_Valores_Adqui(sheetname)


    sheetname.Range("B7").ClearContents
    sheetname.Range("B10:B93").ClearContents

  

End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub CopiadoValoresVolAdqui(sheetname, contador, VolVentas)


  '          VOLUMEN DE VENTAS
  '          EXTRACION DE ARRAY #
  '
  Dim i As Integer, j As Integer, c As Integer, d As Integer, z As Integer
          z = contador + 6
          i = 22
          Do While (i < 34)
              j = 4
              Do While (i < 34)
                If sheetname.Cells(i, 2) = "" Then
                  sheetname.Cells(i, 2).Value = VolVentas.Cells(z, j).Value
                    i = i + 2
                End If
                j = j - 1
                
                If j = 1 Then
                
                  z = z + 2
                  j = j + 3
                
                End If
                                
              Loop
          Loop
  
  '          EXTRACION DE ARRAY (LC)

          z = contador + 7
          i = 23
          Do While (i < 34)
              j = 4
              Do While (i < 34)
                If sheetname.Cells(i, 2) = "" Then
                  sheetname.Cells(i, 2).Value = VolVentas.Cells(z, j).Value
                    i = i + 2
                End If
                j = j - 1
                
                If j = 1 Then
                
                  z = z + 2
                  j = j + 3
                
                End If
                            
              Loop
          Loop
          
               

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CopiadoValoresSolesDolares(sheetname, contador, VolVentas)

    Dim i As Integer, j As Integer, c As Integer, d As Integer, z As Integer
   
  'CREDITO SOLES NACIONAL
          z = contador + 44
          i = 9
          Do While ( z  > 8)
              j = 6
              Do While ( z > 8)
                  
                  sheetname.Cells(i, 6).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop

          z = contador + 44
          i = 10
          Do While ( z  > 8)
              j = 7
              Do While ( z > 8)
                  
                  sheetname.Cells(i, 6).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop


           'CREDITO DOLARES NACIONAL

          z = contador + 53
          i = 9
          Do While ( z  > 17)
              j = 6
              Do While ( z > 17)
                  
                  sheetname.Cells(i, 7).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop

          z = contador + 53
          i = 10
          Do While ( z  > 17)
              j = 7
              Do While ( z > 17)
                  
                  sheetname.Cells(i, 7).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop


          'CREDITO SOLES INTERNACIONAL
          z = contador + 44
          i = 15
          Do While ( z  > 8)
              j = 8
              Do While ( z > 8)
                  
                  sheetname.Cells(i, 6).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop

          z = contador + 44
          i = 16
          Do While ( z  > 8)
              j = 9
              Do While ( z > 8)
                  
                  sheetname.Cells(i, 6).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop


           'CREDITO DOLARES INTERNACIONAL

          z = contador + 53
          i = 15
          Do While ( z  > 17)
              j = 8
              Do While ( z > 17)
                  
                  sheetname.Cells(i, 7).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop

          z = contador + 53
          i = 16
          Do While ( z  > 17)
              j = 9
              Do While ( z > 17)
                  
                  sheetname.Cells(i, 7).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'DEBITO SOLES NACIONAL
          z = contador + 43
          i = 9
          Do While ( z  > 7)
              j = 6
              Do While ( z > 7)
                  
                  sheetname.Cells(i, 8).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop

          z = contador + 43
          i = 10
          Do While ( z  > 7)
              j = 7
              Do While ( z > 7)
                  
                  sheetname.Cells(i, 8).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop


          'DEBITO DOLARES NACIONAL

          z = contador + 52
          i = 9
          Do While ( z  > 16)
              j = 6
              Do While ( z > 16)
                  
                  sheetname.Cells(i, 9).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop

          z = contador + 52
          i = 10
          Do While ( z  > 16)
              j = 7
              Do While ( z > 16)
                  
                  sheetname.Cells(i, 9).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop

       '''''''''''''''''''''''''''''''''''''''''

        'CREDITO SOLES INTERNACIONAL
          z = contador + 43
          i = 15
          Do While ( z  > 7)
              j = 8
              Do While ( z > 7)
                  
                  sheetname.Cells(i, 8).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop

          z = contador + 43
          i = 16
          Do While ( z  > 7)
              j = 9
              Do While ( z > 7)
                  
                  sheetname.Cells(i, 8).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop


           'CREDITO DOLARES INTERNACIONAL

          z = contador + 52
          i = 15
          Do While ( z  > 16)
              j = 8
              Do While ( z > 16)
                  
                  sheetname.Cells(i, 9).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop

          z = contador + 52
          i = 16
          Do While ( z  > 16)
              j = 9
              Do While ( z > 16)
                  
                  sheetname.Cells(i, 9).Value = VolVentas.Cells(z, j).Value
                  i = i + 2
                  z = z - 18
                          
              Loop
          Loop

                
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''







