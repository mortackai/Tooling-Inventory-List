Sub CommandButton1_Click()

Application.ScreenUpdating = False

If OptionButton1.Value = True Then
    If TextBox1.Value = "" Then
        MsgBox "Please Scan Tool"
        Exit Sub
        
    Else
        '
        '
        
        'variables
        Dim sheetCount(3) As Integer
        Dim totalSheetCount As Integer
        Dim i As Integer
        Dim a As Integer
        Dim b As Integer
        Dim c As Integer
        Dim rowNum As Integer
        Dim machineSheets(3) As String
        Dim tool(7) As Long
        Dim qty(7) As Long
        machineSheets(0) = "Machine 10"
        machineSheets(1) = "Machine 20"
        machineSheets(2) = "Machine 30"
        machineSheets(3) = "Machine 40"
        i = 0
        a = 0
        b = 0
        c = 0

        'loop through each sheet and count instances of the searched bit
        For Each ws In ActiveWorkbook.Worksheets
            If ws.Name = "Machine 10" Or ws.Name = "Machine 20" Or ws.Name = "Machine 30" Or ws.Name = "Machine 40" Then
                ws.Select
                sheetCount(i) = Application.WorksheetFunction.CountIf(Range("Z1:Z200"), TextBox1.Value)
                i = i + 1
            Else
            End If
        Next
        
        'go back to machine 10 sheet and reset i to 0
        Sheets("Machine 10").Select
        i = 0
        
        'loop to find values to captions on page 2 of form
        For Each ws In ActiveWorkbook.Worksheets
            If ws.Name = "Machine 10" Or ws.Name = "Machine 20" Or ws.Name = "Machine 30" Or ws.Name = "Machine 40" Then
                For d = 1 To sheetCount(i)
                    ws.Select
                    Set partRow = Cells.find(TextBox1.Value, After:=ActiveCell)
                    If Not partRow Is Nothing Then
                        partRow.Select
                        Cells(partRow.Row, 1).Select
                        tool(b) = ActiveCell.Value
                        Cells(partRow.Row, 3).Select
                        qty(b) = ActiveCell.Value
                        b = b + 1
                        Cells(partRow.Row, 27).Select
                    Else
                        MsgBox ("no part on this sheet")
                        Exit Sub
                    End If
                    
                Next
                    i = i + 1
            End If
        Next
        
        '
        '
    End If
    
    'assign discovered values to captions on second tab of user form
    OptionButton3.Caption = tool(0)
    OptionButton4.Caption = tool(1)
    OptionButton5.Caption = tool(2)
    OptionButton6.Caption = tool(3)
    OptionButton7.Caption = tool(4)
    OptionButton8.Caption = tool(5)
    OptionButton9.Caption = tool(6)
    OptionButton10.Caption = tool(7)
    Label1.Caption = qty(0)
    Label2.Caption = qty(1)
    Label3.Caption = qty(2)
    Label4.Caption = qty(3)
    Label5.Caption = qty(4)
    Label6.Caption = qty(5)
    Label7.Caption = qty(6)
    Label8.Caption = qty(7)
    
    'switch to next tab
    MultiPage1.Value = 1
    
    
'if an option is selected but the text box is empty
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

Private Sub CommandButton2_Click()

'declare variables
Dim tool As Integer


If OptionButton3 = True Then
    tool = OptionButton3.Caption
    
    If tool > 1000 And tool < 2000 Then
        Sheets("machine 10").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If

    If tool > 2000 And tool < 3000 Then
        Sheets("machine 20").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 3000 And tool < 4000 Then
        Sheets("machine 30").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 4000 Then
        Sheets("machine 40").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If

ElseIf OptionButton4.Value = True Then
    tool = OptionButton4.Caption
    
    If tool > 1000 And tool < 2000 Then
        Sheets("machine 10").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If

    If tool > 2000 And tool < 3000 Then
        Sheets("machine 20").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 3000 And tool < 4000 Then
        Sheets("machine 30").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 4000 Then
        Sheets("machine 40").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
ElseIf OptionButton5.Value = True Then
    tool = OptionButton5.Caption
    
    If tool > 1000 And tool < 2000 Then
        Sheets("machine 10").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If

    If tool > 2000 And tool < 3000 Then
        Sheets("machine 20").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 3000 And tool < 4000 Then
        Sheets("machine 30").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 4000 Then
        Sheets("machine 40").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
ElseIf OptionButton6.Value = True Then
    tool = OptionButton6.Caption
    
    If tool > 1000 And tool < 2000 Then
        Sheets("machine 10").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If

    If tool > 2000 And tool < 3000 Then
        Sheets("machine 20").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 3000 And tool < 4000 Then
        Sheets("machine 30").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 4000 Then
        Sheets("machine 40").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
ElseIf OptionButton7.Value = True Then
    tool = OptionButton7.Caption
    
    If tool > 1000 And tool < 2000 Then
        Sheets("machine 10").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If

    If tool > 2000 And tool < 3000 Then
        Sheets("machine 20").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 3000 And tool < 4000 Then
        Sheets("machine 30").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 4000 Then
        Sheets("machine 40").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
ElseIf OptionButton8.Value = True Then
    tool = OptionButton8.Caption
    
    If tool > 1000 And tool < 2000 Then
        Sheets("machine 10").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If

    If tool > 2000 And tool < 3000 Then
        Sheets("machine 20").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 3000 And tool < 4000 Then
        Sheets("machine 30").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 4000 Then
        Sheets("machine 40").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
ElseIf OptionButton9.Value = True Then
    tool = OptionButton9.Caption
    
    If tool > 1000 And tool < 2000 Then
        Sheets("machine 10").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If

    If tool > 2000 And tool < 3000 Then
        Sheets("machine 20").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 3000 And tool < 4000 Then
        Sheets("machine 30").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 4000 Then
        Sheets("machine 40").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
ElseIf OptionButton10.Value = True Then
    tool = OptionButton10.Caption
    
    If tool > 1000 And tool < 2000 Then
        Sheets("machine 10").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If

    If tool > 2000 And tool < 3000 Then
        Sheets("machine 20").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 3000 And tool < 4000 Then
        Sheets("machine 30").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
    If tool > 4000 Then
        Sheets("machine 40").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + 1
        Exit Sub
    End If
    
Else
    MsgBox ("Please Select a Tool")
End If
End Sub
