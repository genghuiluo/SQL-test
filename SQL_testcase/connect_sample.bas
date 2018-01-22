Attribute VB_Name = "connect_sample"
Public Sub connect_oracle()

   Dim conn As ADODB.Connection
   Dim result As ADODB.Recordset
   Dim SQL As String
   
   Set conn = New ADODB.Connection
   Set result = New ADODB.Recordset

   ' Open a Connection using an ODBC DSN n
   conn.Open "master", "cs_master_tgt", "abc123"

   ' Find out if the attempt to connect worked.
   If conn.State = adStateOpen Then
      MsgBox "Connect successfully!!"
      
      
      SQL = "select * from bond_position_own;"
      result.Open SQL, conn
      
      row_num = 1
      Do While Not result.EOF = True
        
        Sheets("result").Range("A" & row_num) = result.Fields(3).Value
        result.MoveNext
        row_num = row_num + 1
      
      Loop
      
   Else
      MsgBox "Failed to connect database."
   End If

   ' Close the connection.
   conn.Close
    
End Sub
