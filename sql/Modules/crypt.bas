Attribute VB_Name = "crypt"
Public conn As Object

Public Function connectConn(dbFileName As String)
    Dim rs1 As New ADODB.Recordset
    Dim extension As String
    Dim ctr As Integer
    
    extension = Right(dbFileName, 3)
    Set conn = CreateObject("Adodb.Connection")
    If conn.State <> 0 Then conn.Close
    If UCase(extension) = "DBF" Then
        conn.Open "Provider=VFPOLEDB.1;Data Source=" & dbFileName & ""
        frmMain.dblist.AddItem "Not Available with DBF"
    Else
        conn.CursorLocation = adUseClient
        conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbFileName & ";Persist Security Info=False"
        conn.Open '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbFileName & ""
        Set rs1 = conn.OpenSchema(adSchemaTables)
        
        Do While Not rs1.EOF
            ctr = rs1.Fields.Count
            Debug.Print rs1.Fields("TABLE_NAME")
            If UCase(Left(rs1.Fields("TABLE_NAME"), 4)) <> "MSYS" Then frmMain.dblist.AddItem (rs1.Fields("TABLE_NAME"))
        rs1.MoveNext
        Loop
        
    End If
End Function

