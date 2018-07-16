Sub CommandButton1_Click()

Application.ScreenUpdating = False

If OptionButton1.Value = True Then
    If TextBox1.Value = "" Then
        MsgBox "Please Scan Tool"
        Exit Sub
        
    Else
        '
        '
        Dim ws As Worksheet
        Dim startIndex As Integer
        Dim endIndex As Integer
        Dim sheetCount As Integer
        startIndex = 1
        endIndex = 0
        
        For Each ws In ActiveWorkbook.Worksheets
            If ws.Name = "Machine 10" Or ws.Name = "Machine 20" Or ws.Name = "Machine 30" Or ws.Name = "Machine 40" Then
                ws.Select
                sheetCount = Application.WorksheetFunction.CountIf(Columns(26), TextBox1.Value)
                endIndex = endIndex + sheetCount
                sheetCount = 0
            Else
            End If
            
            
        Next
		
		'Sheets("Machine 10").Select
        'sheet1Count = Application.WorksheetFunction.CountIf(Columns(26), TextBox1.Value)
        '
        'For i = startIndex To sheet1Count
        '    Set partRow = Cells.find(TextBox1.Value, After:=ActiveCell)
        '        If Not partRow Is Nothing Then
        '            partRow.Select
        '            MsgBox (TextBox1.Value & " found in row: " & partRow.Row)
        '            Label1.Caption = Sheets("Machine 10").Cells(partRow.Row, 3)
        '            OptionButton3.Caption = Sheets("Machine 10").Cells(partRow.Row, 1)
        '            
        '        Else
        '            MsgBox (TextBox1.Value & " not found")
        '        End If

        'Sheets("Machine 20").Select
        'sheet2Count = Application.WorksheetFunction.CountIf(Columns(26), TextBox1.Value)
        '
        'Sheets("Machine 30").Select
        'sheet3Count = Application.WorksheetFunction.CountIf(Columns(26), TextBox1.Value)
        '
        'Sheets("Machine 40").Select
        'sheet4Count = Application.WorksheetFunction.CountIf(Columns(26), TextBox1.Value)         
            
        Next
		
        MsgBox (endIndex)
        For i = startIndex To endIndex
            '
            '
            Dim rowNum As Integer
                Set partRow = Cells.find(TextBox1.Value, After:=ActiveCell)
                If Not partRow Is Nothing Then
                    partRow.Select
                    MsgBox (TextBox1.Value & " found in row: " & partRow.Row)
                    Label1.Caption = Sheets("Machine 10").Cells(partRow.Row, 3)
                    OptionButton3.Caption = Sheets("Machine 10").Cells(partRow.Row, 1)
                    
                Else
                    MsgBox (TextBox1.Value & " not found")
                End If
            '
            '
        Next
        
        
        '
        '
    End If
    
ElseIf OptionButton2.Value = True Then
    If TextBox1.Value = "" Then
        MsgBox "Please Scan Tool"
        Exit Sub
    Else
    
        'code here to remove
        
    End If
    
Else
    MsgBox "Please Select An Option"
End If
End Sub
