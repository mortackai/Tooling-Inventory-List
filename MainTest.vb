Sub CommandButton1_Click()

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
                sheetCount = ActiveWorksheet.CountIf(Range("Z:Z"), TextBox1.Value)
                endIndex = endIndex + sheetCount
                MsgBox (endIndex)
            Else
            
            End If
        Next
        
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
