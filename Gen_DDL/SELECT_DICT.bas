Attribute VB_Name = "select_dict"
Public dict_name As String



Public Sub Sub_Select_Dict()

    'Create browser to select a file
    Set objFl = Application.FileDialog(msoFileDialogFilePicker)

    With objFl
        If .Show = -1 Then
            dict_name = .SelectedItems(1)
            Call Sub_Gen_DDL
        Else
        'Click cancel
            Exit Sub
        End If
    End With

End Sub

