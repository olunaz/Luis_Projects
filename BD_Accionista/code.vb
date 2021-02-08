Sub OpenTextFile()

    Dim FilePath As String

    FilePath = "C:\Users\P027404\Desktop\Proyectos\BD_Accionista\ECV0506R_DAT_PER_20210126.PEP116001004.TXT"
    
    Open FilePath For Input As #1
    
    row_number = 0
    
    Do Until EOF(1)
    
        Line Input #1, LineFromFile
        LineItems = Split(LineFromFile, ",")

        ActiveCell.Offset (row_number,0).Value = LineItems(2)
        ActiveCell.Offset (row_number,0).Value = LineItems(1)
        ActiveCell.Offset (row_number,0).Value = LineItems(0)

        row_number = row_number + 1

    Loop


    Close #1
        
End Sub
