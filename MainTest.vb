Sub CommandButton1_Click()

'Dim nextOpenCell As Integer
Dim Barcode As Integer
Dim currentWorksheet As Worksheet
Dim searchRange As Range
Dim componentPNColumn As integer
Application.FindFormat.Clear

If OptionButton2.Value = True Then
    If TextBox1.Value = "" Then
        MsgBox "Please Scan Tool"
        Exit Sub
    'loop through all sheets in the workbook
	Else
		For Each currentWorkSheet In ActiveWorkbook.Worksheets
		
        currentWorkSheet.Cells(1,1).Select
		
			Dim varList As New List(Of String)
			For i As Integer = 0 To 10
					
				'search for all instances of P/N
				Set searchResult = currentWorkSheet.Cells.Find(_
					What:=TextBox1.Value, _
					After:=ActiveCell, _
					LookIn:=xlValues, _
					LookAt:=xlPart, _
					SearchOrder:=xlByRows, _
					SearchDirection:=xlNext, _
					MatchCase:=False, _
					SearchFormat:=False)
					
					Console.WriteLine(searchResult)
					
					varList.add("comPN" & i)
				Next		

			'Dim qty As Long = Cells(toolRow, "C").Value
			'sum 
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