VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub GenJob()

Dim run_order As Integer
Dim pre_run_order As Integer
Dim dependencies As String
Dim pre_dependencies As String

pre_run_order = -1
dependencies = ""

'if already sort by run order
answer = MsgBox("Have you already sorted by run_order in an ascending order?", vbYesNo + vbQuestion, "Sorted ASC?")
If answer = vbNo Then
    Exit Sub
End If

'read visiable cells
last_line = Sheets("panel").Range("A" & Rows.Count).End(xlUp).Row
first_visible_line = Sheets("panel").AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row


For Line = first_visible_line To last_line

    process_nm = Trim(Sheets("panel").Range("B" & Line).Value)
    process_folder = Sheets("panel").Range("C" & Line).Value
    call_script = Sheets("panel").Range("E" & Line).Value
    script_param = Sheets("panel").Range("F" & Line).Value
    ' IsNumeric()
    run_order = CInt(Sheets("panel").Range("J" & Line).Value)
    
    If pre_run_order <> run_order Then
        pre_dependencies = dependencies
        dependencies = ""
        dependencies = process_nm
    Else
        dependencies = dependencies + "," + process_nm
    End If
    

    'gen .job file at current workbook path
    jobfile = Application.ActiveWorkbook.Path & "\" + process_nm + ".job"
    'MsgBox jobfile
    
    Open jobfile For Output As #1
        'https://stackoverflow.com/questions/20016895/remove-double-quotation-from-write-statement-in-vba/20017071
        'write will add double quotes
        Print #1, "type=command"
        If pre_run_order <> -1 Then
            Print #1, "dependencies=" + pre_dependencies
        End If
        
        If process_folder = "" Then
            Print #1, "command=" + call_script + " " + script_param
        Else
            Print #1, "command=cd " + process_folder + ";" + call_script + " " + script_param
        End If
        
    Close #1
    
    pre_run_order = run_order
Next

MsgBox "done"

End Sub
