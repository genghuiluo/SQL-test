Attribute VB_Name = "GEN_DDL"


Public Sub Sub_Gen_DDL()
    Dim dict_wkb As Workbook
    Dim search_rng As Range
    Dim col_rng As Range

On Error GoTo fatal_error

    'empty DDL sheet
    Sheets("DDL").Cells.Clear

    'empty DDL sheet
    Sheets("None").Cells.Clear
    
    'MsgBox dict_name, vbOKOnly
    Set dict_wkb = Workbooks.Open(dict_name)
    
    input_row = 1
    none_input_row = 1
    
    'loop each item in collA
    For i = 2 To ThisWorkbook.Sheets("Sheet1").Range("A1").End(xlDown).Row

        search_str = UCase(ThisWorkbook.Sheets("Sheet1").Range("A" & i).Value)
    
        'search table name in 目录 sheet
        Set search_rng = dict_wkb.Sheets("目录").Range("C:C").Find(search_str, , xlValues, xlWhole)
        
        If search_rng Is Nothing Then
            'MsgBox "not find " & search_str
            'record in None tab
            ThisWorkbook.Sheets("None").Range("A" & none_input_row).Value = search_str
            none_input_row = none_input_row + 1
            
        Else
            'MsgBox "find " & search_str
            
            'swith to specified table sheet
            If Not search_rng.Hyperlinks.Count = 0 Then
                search_rng.Hyperlinks(1).Follow
            End If
            'MsgBox "switch to " & Application.ActiveSheet.Name
            next_sheet = Cells(search_rng.Row, search_rng.Column - 1).Value
            
            Set search_rng = Application.ActiveSheet.Cells.Find("目标表英文字段", , xlValues, xlWhole)
            If search_rng Is Nothing Then
                'hyperlink not work
                GoTo hyperlink_not_work

            Else
gen_ddl:
                start_row = search_rng.Row + 1
                start_col = search_rng.Column
                end_row = Range(Chr(start_col + 64) & start_row).End(xlDown).Row
                    
                'generate DDL
                ThisWorkbook.Sheets("DDL").Cells(input_row, 1).Value = "DROP TABLE " & search_str & " PURGE;"
                input_row = input_row + 1
                ThisWorkbook.Sheets("DDL").Cells(input_row, 1).Value = "CREATE TABLE " & search_str & " ("
                input_row = input_row + 1
                
                Range(Chr(start_col + 64) & start_row & ":" & Chr(start_col + 64 + 3) & end_row).Copy
                ThisWorkbook.Sheets("DDL").Cells(input_row, 1).PasteSpecial xlPasteAll
                input_row = input_row + end_row - start_row + 1
            
                ThisWorkbook.Sheets("DDL").Cells(input_row, 1).Value = ");"
                input_row = input_row + 1
       
            End If
        End If
    
        'Exit Sub
        'dict_wkb.Sheets("目录").Activate

    Next i

    'close dictionary without saving
    dict_wkb.Close savechanges:=False
    'MsgBox "Generate DDL successfully", vbOKOnly, "Gen_DDL.xlsm"
    Exit Sub
    
hyperlink_not_work:
    On Error GoTo invalid_sheet:
    'MsgBox "Invalid hyperlink! try to switch to " & next_sheet
    dict_wkb.Sheets(next_sheet).Activate
    Set search_rng = Application.ActiveSheet.Cells.Find("目标表英文字段", , xlValues, xlWhole)
    GoTo gen_ddl
    
invalid_sheet:
    'MsgBox "Invalid sheet name! try to switch to " & next_sheet & "表"
    dict_wkb.Sheets(next_sheet & "表").Activate
    Set search_rng = Application.ActiveSheet.Cells.Find("目标表英文字段", , xlValues, xlWhole)
    GoTo gen_ddl
    
fatal_error:
    MsgBox "Fatal Error: Please report this bug to Mark. Thankyou!", vbExclamation, "Fatal Error"
    
End Sub
