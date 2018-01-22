Attribute VB_Name = "SQL"

Public Function Connection() As ADODB.Connection

On Error GoTo return_nothing
    Dim dsn As String, user_id As String, passwd As String
    
    dsn = InputBox("The ODBC DSN name which you have already created:")
    If dsn = vbNullString Then
        GoTo return_nothing
    End If
    
    user_id = InputBox("Username:")
    passwd = InputBox("Password:")
    
    Set Connection = New ADODB.Connection
   ' Open a Connection using an ODBC DSN
   
   Connection.Open dsn, user_id, passwd

   ' Find out if the attempt to connect worked.
   If Connection.State = adStateOpen Then
      MsgBox "Connect successfully!!", vbInformation, "ETL测试工具"
   Else
      GoTo return_nothing
   End If
    
    Exit Function
    
return_nothing:
    Set Connection = Nothing
    Exit Function
    
End Function


Public Sub Query(conn As ADODB.Connection, SQL As String, flush As Boolean, done As Boolean)

On Error GoTo query_error
    Dim result As ADODB.Recordset

    Set result = New ADODB.Recordset
    
    Set result = conn.Execute(SQL)

    'If the command is not intended to return results (for example, an SQL UPDATE query) the provider returns Nothing as long as the option adExecuteNoRecords is specified; otherwise Execute returns a closed Recordset.
    'https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/execute-method-ado-connection
    If result.State = adStateOpen Then
        'empty "result" sheet
        If flush Then
            Sheets("result").Cells.Clear
            Sheets("result").Range("A1").CopyFromRecordset result
        Else
            last_line = Sheets("result").Range("A" & Rows.Count).End(xlUp).Row
            If Sheets("result").Range("A1").Value = "" Then
                Sheets("result").Range("A1").CopyFromRecordset result
            Else
                Sheets("result").Range("A" & last_line + 1).CopyFromRecordset result
            End If
        End If
    
    Else
    
        If flush Then
            Sheets("result").Cells.Clear
        End If
    
    End If
    
    done = True
    Exit Sub
    
query_error:
     MsgBox "Can not excute your SQL!", vbCritical, "ETL测试工具"
     done = False
     Exit Sub
    
End Sub
