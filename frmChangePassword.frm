VERSION 5.00
Begin VB.Form frmChangePassword 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8205
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   4200
      TabIndex        =   11
      Text            =   " "
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtOldPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4200
      TabIndex        =   10
      Text            =   " "
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtUserId 
      Height          =   285
      Left            =   4200
      TabIndex        =   5
      Text            =   " "
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox cmbUserId 
      Height          =   315
      Left            =   4200
      TabIndex        =   4
      Text            =   " "
      Top             =   1560
      Width           =   1575
   End
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -360
      Width           =   1215
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtConfirmPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtNewPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
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
      Left            =   1920
      TabIndex        =   17
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   5775
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   7980
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   5895
      Left            =   7920
      TabIndex        =   14
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7860
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      Height          =   495
      Left            =   2160
      Top             =   4800
      Width           =   3735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      Height          =   3375
      Left            =   1440
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
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
      Left            =   2040
      TabIndex        =   12
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   2520
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   3000
      Width           =   1725
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
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
      Left            =   2040
      TabIndex        =   7
      Top             =   3480
      Width           =   2100
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Id:"
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
      Left            =   2040
      TabIndex        =   6
      Top             =   1560
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   240
      Picture         =   "frmChangePassword.frx":0442
      Stretch         =   -1  'True
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  
Dim rs As New ADODB.Recordset

 

Private Sub cmbUserId_Click()

If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Users where UserId='" & cmbUserId & "'", CN
If rs("UserId") = "admin" Then
MsgBox "Sorry Admin a/c is not accessible", vbCritical + vbSystemModal, "Invalid Password"
Else
txtOldPassword.Text = rs("Password")
txtUserName.Text = rs("UserName")
txtNewPassword.SetFocus
End If
End Sub


Private Sub cmdOK_Click()
If User = True Then
'    If txtOldPassword.Text = "" Then
'    txtOldPassword.SetFocus
'    MsgBox "Please Enter a Value"
'    Exit Sub
    If txtNewPassword.Text = "" Then
    txtNewPassword.SetFocus
    MsgBox "Please Enter a Password"
    Exit Sub
    ElseIf txtConfirmPassword.Text = "" Then
    txtOldPassword.SetFocus
    Exit Sub
    MsgBox "Please Confirm Your Password"
    End If
If txtConfirmPassword.Text <> txtNewPassword.Text Then
txtConfirmPassword.SetFocus
txtConfirmPassword.SelStart = 0
txtConfirmPassword.SelLength = Len(txtConfirmPassword.Text)
MsgBox "Both Password Doesn't Match"
Exit Sub
End If
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Users where UserId='" & cmbUserId.Text & "'", CN
If rs.RecordCount > 0 Then GoTo Correct
 

'    Data1.DatabaseName = App.Path & "\Cyber.mdb"
'    Data1.RecordSource = "select * from user_id"
'Data1.Refresh
'Data1.Recordset.FindFirst "password='" & Form1.txtOldPassword.Text & "'"
'If Data1.Recordset.NoMatch = True Then

'txtOldPassword.SetFocus
'txtOldPassword.SelStart = 0
'txtOldPassword.SelLength = Len(txtOldPassword.Text)
'MsgBox "Please Enter a Correct Password"
'Exit Sub


'Else
'Data1.Recordset.Edit

Correct:
'Data1.Recordset("password") = txtConfirmPassword.Text
'Data1.Recordset.Update
rs("Password") = txtConfirmPassword.Text
rs.Update
MsgBox "Password has been Changed", , "Message"
txtOldPassword.Text = ""
txtNewPassword.Text = ""
txtConfirmPassword.Text = ""
'cmbUserId.ListIndex (0)
'txtOldPassword.SetFocus
Exit Sub
End If
If txtConfirmPassword.Text <> txtNewPassword.Text Then
txtConfirmPassword.SetFocus
txtConfirmPassword.SelStart = 0
txtConfirmPassword.SelLength = Len(txtConfirmPassword.Text)
MsgBox "Both Password Doesn't Match"
Exit Sub
End If
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Users where UserId='" & txtUserId.Text & "'", CN
If rs.RecordCount > 0 Then GoTo Correct
 
 
 
 
End Sub

Private Sub cmdCancel_Click()
'Data1.Recordset.Edit
'Unload Me
'Data1.Recordset.CancelUpdate
End Sub

Private Sub cmdExit_Click()
Unload Me
'Unload frmLogin
End Sub

Private Sub Form_Load()
If User = True Then
txtUserId.Visible = False
cmbUserId.Visible = True
'cmdDelete.Visible = True

If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Users", CN
While rs.EOF = False
    'If IsNull(rs("TotalTime")) Then
    cmbUserId.AddItem rs("UserId")
   ' End If
    rs.MoveNext
    Wend
 Else
' cmdExit.Appearance = AlignmentConstants
 cmbUserId.Visible = False
 txtUserId.Visible = True
 txtUserId.Text = UserNa
 If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Users where UserId='" & UserNa & "'", CN
If rs.RecordCount > 0 Then
txtUserName.Text = rs("UserName")
txtOldPassword.Text = rs("Password")
 End If
 End If



End Sub

Private Sub txtConfirmPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdOk.SetFocus
End If

End Sub

Private Sub txtNewPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtConfirmPassword.SetFocus
End If
End Sub

