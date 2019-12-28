Attribute VB_Name = "Module1"
Public Connection As New ADODB.Connection
Public Sub Connect_DB()
Set Connection = New ADODB.Connection
Connection.CursorLocation = adUseClient
Connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBCAFE.mdb"
End Sub
