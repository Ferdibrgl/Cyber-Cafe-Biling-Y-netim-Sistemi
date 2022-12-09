VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTransaction 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entery"
   ClientHeight    =   5820
   ClientLeft      =   2280
   ClientTop       =   1605
   ClientWidth     =   8115
   Icon            =   "frmTransaction.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120.643
   ScaleMode       =   0  'User
   ScaleWidth      =   4140.306
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
      Height          =   248
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   975
   End
   Begin MSMask.MaskEdBox M1 
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
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
      Height          =   248
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtStartTime 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   " "
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox cmbMachineNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   248
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtMachineName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trasaction"
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
      Height          =   615
      Left            =   1800
      TabIndex        =   15
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   5520
      Width           =   7860
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   5295
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7860
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   5295
      Left            =   7800
      TabIndex        =   11
      Top             =   240
      Width           =   180
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      Height          =   735
      Left            =   1920
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      Height          =   3135
      Left            =   1320
      Top             =   1080
      Width           =   5535
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
      Left            =   2160
      TabIndex        =   10
      Top             =   1680
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Time :"
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
      Left            =   2160
      TabIndex        =   9
      Top             =   2640
      Width           =   1305
   End
   Begin VB.Label Label3 
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
      Left            =   2160
      TabIndex        =   8
      Top             =   2160
      Width           =   1845
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
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
      Left            =   2160
      TabIndex        =   7
      Top             =   3120
      Width           =   690
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   5340
      Left            =   120
      Picture         =   "frmTransaction.frx":0442
      Stretch         =   -1  'True
      Top             =   240
      Width           =   7740
   End
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
 

Private Sub cmbMachineNo_Click()
'**********************************************
'DAO Start
'**********************************************
'Data1.DatabaseName = App.Path & "\Cyber.mdb"
'Data1.RecordSource = "select * from machine_records"
'Data1.Refresh
'Data1.Recordset.FindFirst "machine_no='" & cmbMachineNo.Text & "'"
'If Data1.Recordset.NoMatch = True Then
'txtMachineName.Text = ""
'cmdOk.Enabled = False
''MsgBox "Machine No doesn't exit"
'Exit Sub
'End If
'txtMachineName.Text = Data1.Recordset("machine_name")
'cmdOk.Enabled = True
'***********************************************
'ADO Start
'***********************************************
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from MachineRecords where MachineNo='" & cmbMachineNo.Text & "'", CN
If rs.RecordCount > 0 Then
txtMachineName.Text = rs("MachineName")
M1.Text = Format(Now, "dd/mm/yyyy")
txtStartTime.Text = Time
cmdOk.Enabled = True

'MsgBox "Machine No doesn't exit"
Exit Sub
End If

txtMachineName.Text = ""
cmdOk.Enabled = False




End Sub
Private Sub cmdOK_Click()
'***********************************
'DAO Start
'***********************************
'    Data2.DatabaseName = App.Path & "\Cyber.mdb"
'    Data2.RecordSource = "select * from machine_records"
'    Data2.Refresh
'    Data2.Recordset.FindFirst "machine_no='" & cmbMachineNo.Text & "'"
'        'If Not IsNull(Data2.Recordset("hr")) Then
'        'MsgBox "Machine is Already Alloted", , "Warning"
'        'Exit Sub
'        'End If
'        If Data2.Recordset("oc") = True Then
'        MsgBox "Machine is Already Alloted", , "Warning"
'        Exit Sub
'        Else
'        Data2.Recordset.Edit
'        Data2.Recordset("oc") = True
'        Data2.Recordset.Update
'        End If
'    r = MsgBox("New Entery", vbYesNo, "Entery")
'    If r <> vbYes Then Exit Sub
'Data1.DatabaseName = App.Path & "\Cyber.mdb"
'Data1.RecordSource = "select * from transaction"
'Data1.Refresh
'Data1.Recordset.AddNew
'Data1.Recordset("start_time") = CDate(txtStartTime.Text)
'Data1.Recordset("machine_no") = cmbMachineNo.Text
'Data1.Recordset("date") = Format(CVDate(M1.Text), "dd/mm/yyyy")
''**********
''Data1.Recordset("end_time") = 0
''Data1.Recordset("total_time") = 0
''Data1.Recordset("amount") = 0
''Data1.Recordset("bill_no") = 0
''Data1.Recordset("mont") = 0
''Data1.Recordset("yr") = 0
''Data1.Recordset("hr") = 0
''Data1.Recordset("mnts") = 0
'Debug.Print Format(CVDate(M1.Text), "dd/mm/yyyy")
'Data1.Recordset.Update
'cmdOk.Enabled = False
'Debug.Print modconn.sa
'************************************************
'ADO Start
'***********************************************
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from MachineRecords where MachineNo='" & cmbMachineNo.Text & "'", CN
If rs.RecordCount > 0 Then
'If rs("CurrentStatus") = True Then
' MsgBox "Machine is Already Alloted", , "Warning"
'        Exit Sub
'        Else
 rs("CurrentStatus") = True
 rs.Update
End If
       
'    r = MsgBox("New Entery", vbYesNo, "Entery")
'    If r <> vbYes Then Exit Sub
If rs1.State = adStateOpen Then
rs1.Close
End If
rs1.CursorLocation = adUseClient
rs1.CursorType = adOpenStatic
rs1.LockType = adLockOptimistic

rs1.Open "select * from Tran", CN
'
rs1.AddNew
rs1("StartTime") = CDate(txtStartTime.Text)
rs1("MachineNo") = cmbMachineNo.Text
rs1("Date") = Format(CVDate(M1.Text), "dd/mm/yyyy")
rs1.Update

'Data1.Recordset.Update
'loadTransaction
cmdOk.Enabled = False
'Form_Load
'Debug.Print modconn.sa
cmbMachineNo.Clear
txtMachineName.Text = ""
txtStartTime.Text = ""
'M1.Text = ""

'******************************

If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from MachineRecords where CurrentStatus=False ", CN
If rs.RecordCount > 0 Then
While rs.EOF = False
     cmbMachineNo.AddItem rs("MachineNo")
     rs.MoveNext
     Wend
'End If
'cmbMachineNo.Text = cmbMachineNo.List(0)
End If

cmdOk.Enabled = False
'*****************************


End Sub

Private Sub cmdCancel_Click()
'******************************
'DAO Start
'******************************
'Data1.DatabaseName = App.Path & "\Cyber.mdb"
'Data1.RecordSource = "select * from machine_records"
'Data1.Refresh
'Data1.Recordset.CancelUpdate
'*******************************
'ADO Start
'*******************************
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from MachineRecords", CN
rs.CancelUpdate

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Command4_Click()
ShellExecute 0, "Open", "http://www.rediffmail.com", 0, 0, SW_MAXIMIZE
End Sub

Private Sub Form_Load()
'*********************************************
'DAO Start
'*********************************************
'Data1.DatabaseName = App.Path & "\Cyber.mdb"
'Data1.RecordSource = "select * from machine_records"
'Data1.Refresh
'cmbMachineNo.AddItem "Select a Machine No"
'Do While Not Data1.Recordset.EOF
'cmbMachineNo.AddItem Data1.Recordset("machine_no")
'Data1.Recordset.MoveNext
'Loop
'cmbMachineNo.Text = cmbMachineNo.List(0)
'M1.Text = Format(Now, "dd/mm/yyyy")
'txtStartTime.Text = Time
'**************************************************
'ADO Start
'**************************************************
cmdOk.Enabled = False
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from MachineRecords where CurrentStatus=False order by MachineNo", CN
If rs.RecordCount > 0 Then
While rs.EOF = False
     cmbMachineNo.AddItem rs("MachineNo")
     rs.MoveNext
     Wend
'End If
cmbMachineNo.Text = cmbMachineNo.List(0)
End If
M1.Text = Format(Now, "dd/mm/yyyy")
'txtStartTime.Text = Time
End Sub
Public Sub loadTransaction()
If rs1.State = adStateOpen Then
rs1.Close
End If
rs1.CursorLocation = adUseClient
rs1.CursorType = adOpenStatic
rs1.LockType = adLockOptimistic
rs1.Open "select * from Tran", CN

rs1.AddNew
rs1("StartTime") = CDate(txtStartTime.Text)
rs1("MachineNo") = cmbMachineNo.Text
rs1("Date") = Format(CVDate(M1.Text), "dd/mm/yyyy")
rs1.Update


End Sub

