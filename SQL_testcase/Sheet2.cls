VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ExecuteTestCases()

On Error GoTo error_handle
Dim SQL As String
Dim SQLArr() As String
Dim conn As ADODB.Connection
Dim pass_flg As Integer
Dim query_done As Boolean


'SQL = "SELECT * FROM BOND_POSITION_OWN;"

Set conn = Connection()

If conn Is Nothing Then
    MsgBox "Failed to connect database.", vbCritical, "ETL≤‚ ‘π§æﬂ"
    Exit Sub
End If

last_line = Sheets("Test Cases").Range("A" & Rows.Count).End(xlUp).Row

first_visible_line = Sheets("Test Cases").AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row
'MsgBox first_visible_line

For Line = first_visible_line To last_line
    'reset pass_flg for each case
    pass_flg = 0
    
    'blue
    Sheets("Test Cases").Range("F" & Line).Value = "Working On"
    Sheets("Test Cases").Range("F" & Line).Interior.ColorIndex = 37
    ActiveWindow.ScrollRow = Line
         
    SQL = Sheets("Test Cases").Range("D" & Line).Value
    
    SQLArr = Split(SQL, ";")
    
    For idx = LBound(SQLArr) To UBound(SQLArr) - 1
        'MsgBox SQLArr(idx)
        If idx = LBound(SQLArr) Then
            'first & last SQL need flush result sheet
            'MsgBox "flush"
            Query conn, Trim(SQLArr(idx)), True, query_done
        Else
            'MsgBox "not flush"
            Query conn, Trim(SQLArr(idx)), False, query_done
        End If
    Next
    'if query failed
    If query_done = False Then
        
        GoTo error_handle
        
    End If
    
    'compare result
    If Sheets("Test Cases").Range("E" & Line).Hyperlinks.Count = 1 Then

        'expected result is sheet name
        expected_result_sheet = Evaluate(Sheets("Test Cases").Range("E" & Line).Hyperlinks(1).SubAddress).Parent.Name
        'Sheets (Sheets("Test Cases").Range("D" & Line).Value)

        exp_last_row = Sheets(expected_result_sheet).Range("A" & Rows.Count).End(xlUp).Row
        exp_last_col = Sheets(expected_result_sheet).Cells(1, Columns.Count).End(xlUp).Column

        For i = 1 To exp_last_row
            For j = 1 To exp_last_col
                If Sheets(expected_result_sheet).Cells(i, j).Value = Sheets("result").Cells(i, j).Value Then
                    Sheets(expected_result_sheet).Cells(i, j).Interior.ColorIndex = 0
                Else
                    Sheets(expected_result_sheet).Cells(i, j).Interior.ColorIndex = 6
                    pass_flg = 1
                End If
            Next
        Next

    Else

        'expected result is single string
        If Sheets("Test Cases").Range("E" & Line).Value <> Sheets("result").Range("A1").Value Then
            pass_flg = 2
        End If

    End If
    
    
    Select Case pass_flg
        Case 0
            'greeen
            Sheets("Test Cases").Range("F" & Line).Value = "Pass"
            Sheets("Test Cases").Range("F" & Line).Interior.ColorIndex = 4

        Case 1
            'yellow
            Sheets("Test Cases").Range("F" & Line).Value = "Need Check Manually"
            Sheets("Test Cases").Range("F" & Line).Interior.ColorIndex = 6

        Case 2
            'red
            Sheets("Test Cases").Range("F" & Line).Value = "Fail"
            Sheets("Test Cases").Range("F" & Line).Interior.ColorIndex = 3

    End Select

Next

conn.Close
Exit Sub

error_handle:
    'Err.Number & vbCr & Err.Description
    MsgBox Err.Number & vbCr & Err.Descriptionc, vbCritical, "Please submit this bug to ghluo@chinacscs.com."

End Sub
