Sub CommandButton1_Click()

'dont update the screen while the code is running, it runs faster
Application.ScreenUpdating = False

'variables
Dim delta As Integer
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

'option 1 error check and valid P/N check
If OptionButton1.Value = True Then
    If TextBox1.Value = "" Then
        MsgBox "Please Scan Tool"
        Exit Sub
    ElseIf TextBox1.Value > 9999999 Or TextBox1.Value < 1000000 Then
        MsgBox "invalid P/N"
        Exit Sub
    End If

'option 2 error check and valid P/N check
ElseIf OptionButton2.Value = True Then
    If TextBox1.Value = "" Then
        MsgBox "Please Scan Tool"
        Exit Sub
    ElseIf TextBox1.Value > 9999999 Or TextBox1.Value < 1000000 Then
        MsgBox "invalid P/N"
        Exit Sub
    End If
Else
    MsgBox "Please Select An Option"
    Exit Sub
End If

'loop through each sheet and count instances of the searched bit per sheet
For Each ws In ActiveWorkbook.Worksheets
    If ws.Name = "Machine 10" Or ws.Name = "Machine 20" Or ws.Name = "Machine 30" Or ws.Name = "Machine 40" Then
        ws.Select
        sheetCount(i) = Application.WorksheetFunction.CountIf(Range("F1:F200"), TextBox1.Value)
        i = i + 1
    Else
    End If
Next

'if no instances of the P/N are found, error is shown and exits sub
If sheetCount(0) = 0 And sheetCount(1) = 0 And sheetCount(2) = 0 And sheetCount(3) = 0 Then
    MsgBox "Tool Not Found"
    Exit Sub
End If

'go back to machine 10 sheet and reset i to 0
Sheets("Machine 10").Select
i = 0

'search for P/N on each sheet the correct number of times and assign to respective arrays
For Each ws In ActiveWorkbook.Worksheets
    If ws.Name = "Machine 10" Or ws.Name = "Machine 20" Or ws.Name = "Machine 30" Or ws.Name = "Machine 40" Then
        For d = 1 To sheetCount(i)
            ws.Select
            Set partRow = Cells.find(TextBox1.Value, After:=ActiveCell)
            If Not partRow Is Nothing Then
                partRow.Select
                Cells(partRow.Row, 1).Select
                tool(b) = ActiveCell.Value
                Cells(partRow.Row, 2).Select
                qty(b) = ActiveCell.Value
                b = b + 1
                Cells(partRow.Row, 6).Select
            Else
                MsgBox ("no part on this sheet")
                Exit Sub
            End If

        Next
            i = i + 1
    End If
Next

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

End Sub

Private Sub CommandButton2_Click()

    'declare variables
    Dim tool As Integer

    'error check again, add tool
    If OptionButton1.Value = True Then
        If TextBox1.Value = "" Then
            MsgBox "Please Scan Tool"
            Exit Sub
        End If

    'error check, remove tool
    ElseIf OptionButton2.Value = True Then
        If TextBox1.Value = "" Then
            MsgBox "Please Scan Tool"
            Exit Sub
        End If
    Else
        MsgBox "Please Select An Option"
        Exit Sub
    End If

    'check which option on 2nd page of userform was selected
    If OptionButton3 = True Then
        toolSelect_timeStamp

    ElseIf OptionButton4.Value = True Then
        toolSelect_timeStamp

    ElseIf OptionButton5.Value = True Then
        toolSelect_timeStamp

    ElseIf OptionButton6.Value = True Then
        toolSelect_timeStamp

    ElseIf OptionButton7.Value = True Then
        toolSelect_timeStamp

    ElseIf OptionButton8.Value = True Then
        toolSelect_timeStamp

    ElseIf OptionButton9.Value = True Then
        toolSelect_timeStamp

    ElseIf OptionButton10.Value = True Then
        toolSelect_timeStamp

    Else
        MsgBox ("Please Select a Tool")
        Exit Sub
    End If

End Sub

Sub toolSelect_timeStamp()

    'Variables
    Dim deltaMod As Integer
    Dim tool As Integer
    Dim addOrRemove As String
    Dim offsetValue As Integer
    Dim toolQty As Integer


    'if option button is checked then assign moog part number to variable tool
    For i = 3 To 10
        If Me.Controls("OptionButton" & i).Value = True Then
            tool = Me.Controls("OptionButton" & i).Caption
            toolQty = Me.Controls("Label" & i - 2).Caption
        End If
    Next

    'if option 1 is selected assign 1 to delta
    If OptionButton1.Value = True Then
        If TextBox1.Value = "" Then
            MsgBox "Please Scan Tool"
            Exit Sub
        Else
            deltaMod = 1
            delta = TextBox2.Value * deltaMod
            addOrRemove = "Add"
            offsetValue = 2
        End If

    'if option 2 is selected assign -1 to delta
    ElseIf OptionButton2.Value = True Then
        If TextBox1.Value = "" Then
            MsgBox "Please Scan Tool"
            Exit Sub

        ElseIf toolQty <= TextBox2.Value Then
            MsgBox "cant remove more than is available in inventory"
            Exit Sub

        Else
            deltaMod = -1
            delta = TextBox2.Value * deltaMod
            addOrRemove = "Remove"
            offsetValue = 3
        End If
    Else
        MsgBox "Please Select An Option"
        Exit Sub
    End If

    'check which sheet the tool is on and change value of the qty column by +1 or -1
    If tool > 1000 And tool < 2000 Then
        Sheets("machine 10").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + delta
        ActiveCell.Offset(0, offsetValue).Value = Date + Time

        For timestampqty = 1 To TextBox2.Value
            'Time stamp entry for log and for most recent
            Worksheets("TimeStamps").Select
            Range("A1").End(xlDown).Offset(1, 0).Select
            ActiveCell.Value = tool
            ActiveCell.Offset(0, 1).Value = Date + Time
            ActiveCell.Offset(0, 2).Value = addOrRemove
        Next
    End If

    If tool > 2000 And tool < 3000 Then
        Sheets("machine 20").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + delta
        ActiveCell.Offset(0, offsetValue).Value = Date + Time

        For timestampqty = 1 To TextBox2.Value
            'Time stamp entry for log and for most recent
            Worksheets("TimeStamps").Select
            Range("A1").End(xlDown).Offset(1, 0).Select
            ActiveCell.Value = tool
            ActiveCell.Offset(0, 1).Value = Date + Time
            ActiveCell.Offset(0, 2).Value = addOrRemove
        Next
    End If

    If tool > 3000 And tool < 4000 Then
        Sheets("machine 30").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + delta
        ActiveCell.Offset(0, offsetValue).Value = Date + Time

        For timestampqty = 1 To TextBox2.Value
            'Time stamp entry for log and for most recent
            Worksheets("TimeStamps").Select
            Range("A1").End(xlDown).Offset(1, 0).Select
            ActiveCell.Value = tool
            ActiveCell.Offset(0, 1).Value = Date + Time
            ActiveCell.Offset(0, 2).Value = addOrRemove
        Next
    End If

    If tool > 4000 Then
        Sheets("machine 40").Select
        Cells(1, 1).Select
        Range("A:A").find(tool, After:=ActiveCell).Select
        ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 2).Value + delta
        ActiveCell.Offset(0, offsetValue).Value = Date + Time

        For timestampqty = 1 To TextBox2.Value
            'Time stamp entry for log and for most recent
            Worksheets("TimeStamps").Select
            Range("A1").End(xlDown).Offset(1, 0).Select
            ActiveCell.Value = tool
            ActiveCell.Offset(0, 1).Value = Date + Time
            ActiveCell.Offset(0, 2).Value = addOrRemove
        Next
    End If

    'close userform
    Unload Me

End Sub

Private Sub SpinButton1_Change()
      TextBox2.Text = SpinButton1.Value
End Sub
