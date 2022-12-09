Attribute VB_Name = "modconn"
'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public sa
Public ent As Boolean

Public c As Integer
'count = 0
'********************************
'This two variables are used for change password
'************************
Public User As Boolean
Public UserNa As Variant
'*****************************
Public ID As String
Public CN As New ADODB.Connection
Public Sub AttendConn()
CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\CyberCafe.mdb;Jet OLEDB:Database Password=milan;"
CN.Open
User = False

End Sub
Public Sub CloseConn()
CN.Close
Set CN = Nothing
End Sub

Public Sub counted()

End Sub
