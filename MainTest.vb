Sub CommandButton1_Click()

'Dim nextOpenCell As Integer
'Dim Barcode As Long
'Dim currentWorksheet As Worksheet
'Dim searchRange As Range
'Dim componentPNColumn As Integer
'Barcode = TextBox1.Value
'Application.FindFormat.Clear

If OptionButton1.Value = True Then
    If TextBox1.Value = "" Then
        MsgBox "Please Scan Tool"
        Exit Sub
        
    Else
        '
        '
        Dim rowNum As Integer
            Set partRow = Cells.find(TextBox1.Value, After:=ActiveCell)
            If Not partRow Is Nothing Then
                partRow.Select
                MsgBox (TextBox1.Value & " found in row: " & partRow.Row)
                Label1.Caption = Sheets("Machine 10").Cells(partRow.Row, 3)
                Label2.Caption = Sheets("Machine 10").Cells(partRow.Row, 3)
                Label3.Caption = Sheets("Machine 10").Cells(partRow.Row, 3)
                Label4.Caption = Sheets("Machine 10").Cells(partRow.Row, 3)
                
            Else
                MsgBox (TextBox1.Value & " not found")
            End If
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
