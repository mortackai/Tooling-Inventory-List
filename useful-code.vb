Sub CommandButton1_Click()

Dim nextOpenCell As Integer
Dim Barcode As Integer
Dim ws As Worksheet
Set ws = ActiveSheet
If OptionButton2.Value = True Then
    If TextBox1.Value = "" Then
        MsgBox "Please Scan Tool"
        Exit Sub
    Else
        For Each cell In ws.Columns(1).Cells
            If IsEmpty(cell) = True Then cell.Select: Exit For
        Next cell
        ActiveCell.Value = TextBox1.Value
    End If
ElseIf OptionButton3.Value = True Then
    If TextBox1.Value = "" Then
        MsgBox "Please Scan Tool"
        Exit Sub
    Else
    
        Dim rs As Worksheet
        Dim lRow As Long
        Dim strSearch As String

        '~~> Set this to the relevant worksheet
        Set rs = ActiveSheet

        '~~> Search Text
        strSearch = TextBox1

        With rs
            '~~> Remove any filters
            .AutoFilterMode = False

            lRow = .Range("A" & .Rows.Count).End(xlUp).Row

            With .Range("A1:A" & lRow)
                .AutoFilter Field:=1, Criteria1:="=*" & strSearch & "*"
                .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            End With

            '~~> Remove any filters
            .AutoFilterMode = False
        End With
    End If
Else
    MsgBox "Please Select An Option"
End If
End Sub
