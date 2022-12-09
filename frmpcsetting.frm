VERSION 5.00
Begin VB.Form frmpcsetting 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine Entery"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmpcsetting.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4.281
   ScaleMode       =   5  'Inch
   ScaleWidth      =   5.698
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      Width           =   1094
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   1094
   End
   Begin VB.ComboBox cmbMachineNo 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   1094
   End
   Begin VB.CommandButton cmdModify 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Modify"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1094
   End
   Begin VB.CommandButton cmdAddNew 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Add New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1094
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "C&ancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1094
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -480
      Width           =   1140
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   1094
   End
   Begin VB.TextBox txtMachineName 
      Height          =   285
      Left            =   4440
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtMachineNo 
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   " Computer Setting"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2520
      TabIndex        =   16
      Top             =   600
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      Height          =   1215
      Left            =   960
      Top             =   4320
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      Height          =   2895
      Left            =   2160
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   7980
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   5895
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   7740
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   5895
      Left            =   7920
      TabIndex        =   12
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Machine No :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Top             =   2040
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Machine Name :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2520
      TabIndex        =   10
      Top             =   2640
      Width           =   1845
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   240
      Picture         =   "frmpcsetting.frx":0442
      Stretch         =   -1  'True
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "frmpcsetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim rs As New ADODB.Recordset



Private Sub cmbMachineNo_Click()
'********************************************
' DAO Start
'*******************************************
'Data1.DatabaseName = App.Path & "\Cyber.mdb"
'Data1.RecordSource = "select * from MachineRecords"
'Data1.Refresh
'Data1.Recordset.FindFirst "MachineNo='" & cmbMachineNo.Text & "'"
'txtMachineNo.Text = Data1.Recordset("MachineNo")
'txtMachineName.Text = Data1.Recordset("MachineName")
'txtMachineNo.Visible = True
'cmbMachineNo.Visible = False
'************************************
'DAO End
'************************************
'**********************************************
'ADODB START
'**********************************************
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from MachineRecords where MachineNo='" & cmbMachineNo.Text & "'", CN
If rs.RecordCount > 0 Then
txtMachineNo.Text = rs("MachineNo")
txtMachineName.Text = rs("MachineName")
    txtMachineName.SetFocus
    txtMachineName.SelStart = 0
    txtMachineName.SelLength = Len(txtMachineName.Text)
 
txtMachineNo.Visible = True
cmbMachineNo.Visible = False
End If

End Sub

 
Private Sub cmdOK_Click()
'****************************************************
'DAO Start
'****************************************************
''    If str = "AddNew" Then
''If txtMachineNo.Text = "" Then
''txtMachineNo.SetFocus
''MsgBox "Please enter a Machine No", , "Warnig"
''Exit Sub
''ElseIf txtMachineName.Text = "" Then
''txtMachineName.SetFocus
''MsgBox "Please enter a Machine Name", , "Warnig"
''Exit Sub
''End If
''Data1.DatabaseName = App.Path & "\Cyber.mdb"
''Data1.RecordSource = "select * from MachineRecords"
''Data1.Refresh
''Data1.Recordset.FindFirst "MachineNo='" & txtMachineNo.Text & "'"
''    If Data1.Recordset("MachineNo") = txtMachineNo.Text Then
''    txtMachineNo.SetFocus
''    txtMachineNo.SelStart = 0
''    txtMachineNo.SelLength = Len(txtMachineNo.Text)
''    MsgBox "Machine No already Exit", , "Login"
''    Exit Sub
''    End If
''Data1.Recordset.AddNew
''Data1.Recordset("MachineNo") = txtMachineNo.Text
''Data1.Recordset("MachineName") = txtMachineName.Text
''Data1.Recordset.Update
''MsgBox "New Entery Completed"
''    cmbMachineNo.Clear
''    Data1.DatabaseName = App.Path & "\Cyber.mdb"
''    Data1.RecordSource = "select * from MachineRecords"
''    Data1.Refresh
''    cmbMachineNo.AddItem "Select a Machine name"
''    Do While Not Data1.Recordset.EOF
''    cmbMachineNo.AddItem Data1.Recordset("MachineNo")
''    Data1.Recordset.MoveNext
''    Loop
''            cmbMachineNo.ListIndex = 0
''            cmdOK.Visible = False
''            cmdAddNew.Visible = True
''            cmdModify.Visible = True
''            cmdDelete.Visible = True
''ElseIf str = "Modify" Then
''        Data1.DatabaseName = App.Path & "\Cyber.mdb"
''        Data1.RecordSource = "select * from MachineRecords"
''        Data1.Refresh
''Data1.Recordset.FindFirst "MachineNo='" & cmbMachineNo.Text & "'"
''Data1.Recordset.Edit
''Data1.Recordset("MachineNo") = txtMachineNo.Text
''Data1.Recordset("MachineName") = txtMachineName.Text
''Data1.Recordset.Update
''MsgBox "Updation Comleted"
'' cmbMachineNo.Clear
''    Data1.DatabaseName = App.Path & "\Cyber.mdb"
''    Data1.RecordSource = "select * from MachineRecords"
''    Data1.Refresh
''    cmbMachineNo.AddItem "Select a Machine name"
''    Do While Not Data1.Recordset.EOF
''    cmbMachineNo.AddItem Data1.Recordset("MachineNo")
''    Data1.Recordset.MoveNext
''    Loop
''    cmbMachineNo.ListIndex = 0
''            cmdOK.Visible = False
''            cmdAddNew.Visible = True
''            cmdModify.Visible = True
''            cmdDelete.Visible = True
''End If
'********************************************************
'DAO End
'********************************************************
'********************************************************
'ADO Start
'********************************************************
    If str = "AddNew" Then
If txtMachineNo.Text = "" Then
txtMachineNo.SetFocus
MsgBox "Please enter a Machine No", , "Warnig"
Exit Sub
ElseIf txtMachineName.Text = "" Then
txtMachineName.SetFocus
MsgBox "Please enter a Machine Name", , "Warnig"
Exit Sub
End If
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from MachineRecords where MachineNo='" & txtMachineNo.Text & "'", CN
If rs.RecordCount > 0 Then
    txtMachineNo.SetFocus
    txtMachineNo.SelStart = 0
    txtMachineNo.SelLength = Len(txtMachineNo.Text)
    MsgBox "Machine No already Exist", , "Login"
    Exit Sub
    End If
    rs.AddNew
    rs("MachineNo") = txtMachineNo.Text
    rs("MachineName") = txtMachineName.Text
    rs.Update
    MsgBox "New Entery Completed"
    cmbMachineNo.Clear
'****************************************
'Here Loaded Combo after Update Database
'****************************************
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from MachineRecords", CN
While rs.EOF = False
     cmbMachineNo.AddItem rs("MachineNo")
     rs.MoveNext
Wend

            cmbMachineNo.ListIndex = 0
            cmdOk.Enabled = False
            cmdAddNew.Enabled = True
            
            cmdModify.Enabled = True
            cmdDelete.Enabled = True
            cmdExit.SetFocus
            
            
ElseIf str = "Modify" Then
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from MachineRecords where MachineNo='" & cmbMachineNo.Text & "'", CN
'rs.EditMode
rs("MachineNo") = txtMachineNo.Text
rs("MachineName") = txtMachineName.Text
rs.Update
MsgBox "Updation Comleted"
cmbMachineNo.Clear
'*****************************************************
'Again loaded into combo
'*****************************************************
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from MachineRecords", CN
While rs.EOF = False
     cmbMachineNo.AddItem rs("MachineNo")
     rs.MoveNext
Wend

            cmbMachineNo.ListIndex = 0
            cmdOk.Enabled = False
            cmdAddNew.Enabled = True
            cmdModify.Enabled = True
            cmdDelete.Enabled = True
            cmdExit.SetFocus
            
End If


'********************************************************
'ADO End
'********************************************************
'frmpcstatus.Shape2.FillColor = QBColor(11)



End Sub
Private Sub cmdCancel_Click()
 cmdAddNew.Enabled = True
cmdModify.Enabled = True
cmdDelete.Enabled = True
'cmdRefresh.Visible = False
cmbMachineNo.Visible = False
txtMachineNo.Visible = True
End Sub
Private Sub cmdAddNew_Click()
txtMachineNo.Locked = False
txtMachineNo.Text = ""
txtMachineName.Text = ""
txtMachineNo.SetFocus
str = "AddNew"
cmdAddNew.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
'cmdRefresh.Visible = False
cmdOk.Enabled = True
End Sub
Private Sub cmdModify_Click()
cmbMachineNo.Visible = True
txtMachineNo.Visible = False
txtMachineNo.Locked = False
txtMachineName.Locked = False
str = "Modify"
cmdModify.Enabled = False
cmdAddNew.Enabled = False
cmdDelete.Enabled = False
cmdRefresh.Enabled = False
cmdOk.Enabled = True
End Sub
Private Sub cmdDelete_Click()
d = MsgBox("Are you sure to Delete?", vbYesNo, "Warning")
    If d <> vbYes Then Exit Sub
cmdOk.Enabled = False
cmdAddNew.Enabled = False
cmdModify.Enabled = False
cmdRefresh.Enabled = True
cmdCancel.Enabled = True
cmdDelete.Enabled = False
cmbMachineNo.Visible = True
txtMachineNo.Visible = False
End Sub
Private Sub cmdRefresh_Click()
'*****************************************************
'*DAO Start
'*****************************************************
'Data1.DatabaseName = App.Path & "\Cyber.mdb"
'    Data1.RecordSource = "select * from MachineRecords"
'    Data1.Refresh
'Data1.Recordset.FindFirst "MachineNo='" & txtMachineNo.Text & "'"
'Data1.Recordset.Delete
'MsgBox "Record has been Deleted"
'Data1.Recordset.MoveFirst
'If Data1.Recordset.RecordCount = 0 Then
'MsgBox "No Record Found", , "Message"
'Exit Sub
'End If
'    txtMachineNo.Text = ""
'    txtMachineName.Text = ""
'    cmbMachineNo.Clear
'Data1.DatabaseName = App.Path & "\Cyber.mdb"
'Data1.RecordSource = "select * from MachineRecords"
'Data1.Refresh
''cmbMachineNo.AddItem "Select a Machine name"
'Do While Not Data1.Recordset.EOF
'cmbMachineNo.AddItem Data1.Recordset("MachineNo")
'Data1.Recordset.MoveNext
'Loop
''cmbMachineNo.ListIndex = 0
'cmdRefresh.Visible = False
'cmdAddNew.Visible = True
'cmdModify.Visible = True
'cmdDelete.Visible = True
'*****************************************************
'*DAO End
'*****************************************************
'*****************************************************
'*ADO Start
'*****************************************************
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from MachineRecords where MachineNo='" & txtMachineNo.Text & "'", CN
If rs.RecordCount > 0 Then
rs.Delete adAffectCurrent
MsgBox "Record has been Deleted"
Else
MsgBox "No Record Found", , "Message"
Exit Sub
End If
    txtMachineNo.Text = ""
    txtMachineName.Text = ""
    cmbMachineNo.Clear
If rs.State = adStateOpen Then
 rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from MachineRecords", CN
While rs.EOF = False
     cmbMachineNo.AddItem rs("MachineNo")
     rs.MoveNext
     Wend
'rs.Update
cmdRefresh.Enabled = False
cmdAddNew.Enabled = True
cmdModify.Enabled = True
cmdDelete.Enabled = True


'*****************************************************
'*ADO End
'*****************************************************
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
'*****************************************************
'DAO Start
'*****************************************************
'Data1.DatabaseName = App.Path & "\Cyber.mdb"

'Data1.RecordSource = "select * from MachineRecords"
'Data1.Refresh
'cmbMachineNo.AddItem "Select a Machine name"
'Do While Not Data1.Recordset.EOF
'cmbMachineNo.AddItem Data1.Recordset("MachineNo")
'Data1.Recordset.MoveNext
'Loop
'cmbMachineNo.ListIndex = 0
'************************************
'DAO End
'************************************
'******************************************************
'ADO Start
'******************************************************
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from MachineRecords", CN
While rs.EOF = False
     cmbMachineNo.AddItem rs("MachineNo")
     rs.MoveNext
     Wend
'****************************************
'ADO End
'***************************************
cmdRefresh.Enabled = False
cmdOk.Enabled = False
txtMachineNo.Locked = True
'cmdAddNew.SetFocus
'cmdCancel.SetFocus = True
End Sub

Private Sub txtMachineName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdOk.SetFocus
End If
End Sub

Private Sub txtMachineNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtMachineName.SetFocus
End If
End Sub

Public Sub check()

End Sub

 

