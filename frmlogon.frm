VERSION 5.00
Object = "{08654D78-6636-11D3-87BF-B4980CC10374}#2.0#0"; "MyEllipticButton.ocx"
Begin VB.Form frmLogon 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "                                  USER LOGIN "
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmlogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MyEllipticButton.EllipticButton EllipticButton3 
      Height          =   795
      Left            =   3120
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1402
      BackColor       =   49152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmlogon.frx":0442
      DisabledPicture =   "frmlogon.frx":045E
      DownPicture     =   "frmlogon.frx":047A
      MouseIcon       =   "frmlogon.frx":0496
      Caption         =   "Exit"
   End
   Begin MyEllipticButton.EllipticButton EllipticButton2 
      Height          =   795
      Left            =   1560
      TabIndex        =   8
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1402
      BackColor       =   49152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmlogon.frx":04B2
      DisabledPicture =   "frmlogon.frx":04CE
      DownPicture     =   "frmlogon.frx":04EA
      MouseIcon       =   "frmlogon.frx":0506
      Caption         =   "Cancel"
   End
   Begin MyEllipticButton.EllipticButton EllipticButton1 
      Height          =   795
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1402
      BackColor       =   49152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmlogon.frx":0522
      DisabledPicture =   "frmlogon.frx":053E
      DownPicture     =   "frmlogon.frx":055A
      MouseIcon       =   "frmlogon.frx":0576
      Caption         =   "Login"
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Text            =   "admin"
      Top             =   840
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      Picture         =   "frmlogon.frx":0592
      ScaleHeight     =   255
      ScaleWidth      =   4665
      TabIndex        =   2
      Top             =   2715
      Width           =   4665
      Begin VB.Label lbl2 
         BackStyle       =   0  'Transparent
         Caption         =   "This is made by Milan"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4800
         TabIndex        =   4
         Top             =   0
         Width           =   4695
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Enter Your Id && Password"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000C000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   5
      X1              =   1320
      X2              =   4440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   1440
      X2              =   2640
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   2400
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BorderColor     =   &H0000C000&
      BorderWidth     =   5
      Height          =   975
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   480
      Shape           =   2  'Oval
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "User ID :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Image Image5 
      Height          =   2295
      Left            =   0
      Picture         =   "frmlogon.frx":3470
      Stretch         =   -1  'True
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   735
      Left            =   0
      Picture         =   "frmlogon.frx":4451
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   675
      Left            =   2040
      Picture         =   "frmlogon.frx":5E21
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2625
   End
   Begin VB.Image Image1 
      Height          =   1365
      Left            =   240
      Picture         =   "frmlogon.frx":6BFA
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1365
   End
   Begin VB.Image Image2 
      Height          =   1365
      Left            =   120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChrPass As String
Dim rs As New ADODB.Recordset
'txtUserName.SetFocus



Private Sub EllipticButton1_Click()
'Call check
 
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
'frmmain.MnuReport.Enabled = False
'frmmain.MnuAdmin.Enabled = False
'RS1.Open "select * from status where procode= '" & procod & "'and netcode='" & netcod & "'", CN
'If txtUserName = "Administrator" Then
'frmmain.MnuReport.Enabled = True
'frmmain.MnuAdmin.Enabled = True

rs.Open "select * from Users where UserId='" & txtUserName & "'and Password='" & txtPassword & "'", CN
If rs.RecordCount > 0 Then


'Call check

'        MsgBox "A c c e s s  G r a n t e d", vbInformation + vbSystemModal, "Access Granted"
'        frmmain.Enabled = True
'        MainMdi.Show
If rs("UserId") = "admin" Then
User = True

'frmmain.SmnuComputersetting.Enabled = True
'frmmain.SmnuNewuser.Enabled = True
'frmmain.SmnuRateofbilling.Enabled = True
'frmmain.SmnuRecorddelete.Enabled = True

' .MDIForm1.mnu_int_admin.Enabled = True
 
MDIForm1.mnu_createNUser.Enabled = True
MDIForm1.mnu_changeUPassword.Enabled = True
MDIForm1.mnu_computerSetting.Enabled = True
MDIForm1.mnu_deleteRecord.Enabled = True
MDIForm1.mnu_rateOBilling.Enabled = True


'frmmain.MnuReport.Enabled = True
'frmmain.MnuAdmin.Enabled = True
End If
        Me.Hide
        UserNa = txtUserName
        MDIForm1.Enabled = True
    Else
        MsgBox "Invalid Password or User ID ! Please try again.", vbCritical + vbSystemModal, "Invalid Password"
        txtPassword.Text = ""
        txtPassword.SetFocus
        Exit Sub
    End If
 txtUserName.Text = " "
 txtPassword.Text = " "
 Me.Hide
 MDIForm1.Show
End Sub

Private Sub EllipticButton2_Click()
txtUserName.Text = ""
txtPassword.Text = ""
txtUserName.SetFocus
End Sub

Private Sub EllipticButton3_Click()
Unload Me
Unload MDIForm1
End Sub

Private Sub Form_Load()

txtUserName.Text = ""
txtPassword.Text = ""
'Label1.Caption = Time & "              User Logon"
modconn.AttendConn
MDIForm1.Show
MDIForm1.Enabled = False
'Me.Show


'frmmain.MnuReport.Enabled = False
'frmmain.MnuAdmin.Enabled = False

'Timer2.Interval = 100
'Unload Me
'modconn.AttendConn
  'ChrPass = "secret"
 Me.Top = 3200
  Me.Left = 3700
  lbl1.Left = 4900

'  MainMdi.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
' If Not txtPassword = ChrPass Then
'    Unload Me
''    Unload MainMdi
'
'  Else
'    Load MainMdi
'frmmain.Enabled = True
'    MainMdi.Show
'  End If
'modconn.CloseConn
'Unload MDIForm1
 modconn.CloseConn
End Sub

Private Sub Timer1_Timer()
'lbl1.Left = 4900
  lbl1.Move lbl1.Left - 50
    If lbl1.Left <= -2850 Then
    'If lbl1.Left <= -3900 Then
    lbl2.Visible = True
    lbl1.Visible = False
   lbl2.Move lbl2.Left - 50
      If lbl2.Left <= -1500 Then
          lbl2.Left = 4900
          lbl2.Visible = False
         lbl1.Visible = True
          lbl1.Left = 4900
        End If
   End If

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   txtPassword = StrConv(txtPassword, vbLowerCase)
    EllipticButton1_Click 'CmdLog_Click
  End If
  
End Sub

 

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPassword.SetFocus
End If
End Sub

Private Sub txtUserName_LostFocus()
txtUserName.Text = LCase(txtUserName.Text)
End Sub

