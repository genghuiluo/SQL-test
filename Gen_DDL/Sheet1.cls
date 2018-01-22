VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Sub_Cleanup()

    end_row = Range("A2").End(xlDown).Row
    
    If end_row <> 1 Then
        Range("A2:A" & end_row).Clear
    End If

End Sub
