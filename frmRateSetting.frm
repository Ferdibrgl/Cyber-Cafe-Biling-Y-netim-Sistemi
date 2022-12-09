VERSION 5.00
Begin VB.Form frmRateSetting 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   6210
   ClientLeft      =   1890
   ClientTop       =   1785
   ClientWidth     =   8265
   Icon            =   "frmRateSetting.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8265
   Begin VB.TextBox txtStartingDate 
      Height          =   285
      Left            =   3960
      TabIndex        =   9
      Text            =   " "
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtPerHour 
      Height          =   285
      Left            =   3960
      TabIndex        =   7
      Text            =   " "
      Top             =   2400
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
      Height          =   255
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -480
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   855
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Setting"
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
      Left            =   1800
      TabIndex        =   11
      Top             =   480
      Width           =   4215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      Height          =   735
      Left            =   1560
      Top             =   4560
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      Height          =   3255
      Left            =   1200
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Date :"
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
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ChargePer Hour :"
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
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   5700
      Left            =   240
      Picture         =   "frmRateSetting.frx":0442
      Stretch         =   -1  'True
      Top             =   240
      Width           =   7740
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   5895
      Left            =   7920
      TabIndex        =   6
      Top             =   120
      Width           =   195
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   5820
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   195
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   7980
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7980
   End
End
Attribute VB_Name = "frmRateSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmdOK_Click()
'If txtZeroFifteen.Text = "" Then
'txtZeroFifteen.SetFocus
'MsgBox "Can't Modify with a Balank Field", vbOKOnly + vbCritical, "Warninig"
'Exit Sub
'End If
'If txtFifteenThirty.Text = "" Then
'txtFifteenThirty.SetFocus
'MsgBox "Can't Modify with a Balank Field", vbOKOnly + vbCritical, "Warninig"
'Exit Sub
'End If
'If txtThirtyFortyfive.Text = "" Then
'txtThirtyFortyfive.SetFocus
'MsgBox "Can't Modify with a Balank Field", vbOKOnly + vbCritical, "Warninig"
'Exit Sub
'End If
'If txtFortyfiveSixty.Text = "" Then
'txtFortyfiveSixty.SetFocus
'MsgBox "Can't Modify with a Balank Field", vbOKOnly + vbCritical, "Warninig"
'Exit Sub
'End If
'*********************************************
'DAO Start
'*********************************************
'Data1.DatabaseName = App.Path & "\CyberCafe .mdb"
'Data1.RecordSource = "select * from set1"
'Data1.Refresh
'Data1.Recordset.Edit
'Data1.Recordset("0_15") = Val(txtZeroFifteen.Text)
'Data1.Recordset("15_30") = Val(txtFifteenThirty.Text)
'Data1.Recordset("30_45") = Val(txtZeroFifteen.Text)
'Data1.Recordset("45_60") = Val(txtZeroFifteen.Text)
'Data1.Recordset.Update
'MsgBox "Settings has benn Updated"
'************************************************
'ADO Start
'************************************************
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from RateSetting where EndingDate is null", CN
rs("EndingDate") = Date
rs.Update
rs.AddNew
rs.Fields("OneHour") = Val(txtPerHour.Text)
'rs.Fields("15_30") = Val(txtFifteenThirty.Text)
'rs.Fields("30_45") = Val(txtZeroFifteen.Text)
'rs.Fields("45_60") = Val(txtZeroFifteen.Text)
rs.Fields("StartingDate") = Date
rs.Update
MsgBox "Settings has been Updated"
cmdOk.Enabled = False
cmdModify.Enabled = True

''txtZeroFifteen.Locked = True
''txtFifteenThirty.Locked = True
''txtThirtyFortyfive.Locked = True
''txtFortyfiveSixty.Locked = True
End Sub

Private Sub cmdModify_Click()
s = MsgBox("Want to Modify your Settings", vbYesNo, "Settings")
If s <> vbYes Then Exit Sub
txtStartingDate.Visible = False
Label2.Visible = False
txtPerHour.Text = ""
txtPerHour.SetFocus
'txtZeroFifteen.Locked = False
'txtFifteenThirty.Locked = False
'txtThirtyFortyfive.Locked = False
'txtFortyfiveSixty.Locked = False
cmdOk.Enabled = True
cmdModify.Enabled = False
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
''Data1.DatabaseName = App.Path & "\CyberCafe.mdb"
''Data1.RecordSource = "select * from set1"
''Data1.Refresh
''txtZeroFifteen.Text = Data1.Recordset("0_15")
''txtFifteenThirty.Text = Data1.Recordset("15_30")
''txtThirtyFortyfive.Text = Data1.Recordset("30_45")
''txtFortyfiveSixty.Text = Data1.Recordset("45_60")
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from RateSetting where EndingDate is Null", CN
txtPerHour.Text = rs.Fields("OneHour")
txtStartingDate.Text = rs.Fields("StartingDate")
'txtZeroFifteen.Text = rs.Fields("0_15")
'txtFifteenThirty.Text = rs.Fields("15_30")
'txtThirtyFortyfive.Text = rs.Fields("30_45")
'txtFortyfiveSixty.Text = rs.Fields("45_60")
cmdOk.Enabled = False

End Sub

