VERSION 5.00
Begin VB.Form frmAddnewUser 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New User"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "frmAddnewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8205
   Begin VB.TextBox txtUserId 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4320
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Serif"
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
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtConfirmPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      TabIndex        =   0
      Text            =   " "
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Create New User"
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
      Left            =   2040
      TabIndex        =   14
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7980
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   7980
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   5700
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   5655
      Left            =   7920
      TabIndex        =   10
      Top             =   240
      Width           =   180
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2280
      Top             =   4680
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      FillColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   1680
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Left            =   2040
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   5460
      Left            =   360
      Picture         =   "frmAddnewUser.frx":0442
      Stretch         =   -1  'True
      Top             =   360
      Width           =   7500
   End
End
Attribute VB_Name = "frmAddnewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtUserName.Text = "" Then
       txtUserName.SetFocus
    MsgBox "Please Enter User Name"
    Exit Sub
    ElseIf txtUserId.Text = "" Then
    txtUserId.SetFocus
    MsgBox "Please Enter User Id"
    Exit Sub
    ElseIf txtPassword.Text = "" Then
    txtPassword.SetFocus
    MsgBox "Please Enter Password"
    Exit Sub
    ElseIf txtConfirmPassword.Text = "" Then
    txtOldPassword.SetFocus
    Exit Sub
    MsgBox "Please Confirm Your Password"
    End If



If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Users", CN

While rs.EOF = False
If LCase(txtUserId.Text) = rs("UserId") Then
MsgBox "User Id Already Exists"
txtUserId.SetFocus
txtUserId.SelStart = 0
txtUserId.SelLength = Len(txtUserId.Text)

Exit Sub
End If
rs.MoveNext
Wend
If txtPassword.Text <> txtConfirmPassword.Text Then
txtConfirmPassword.SetFocus
txtConfirmPassword.SelStart = 0
txtConfirmPassword.SelLength = Len(txtConfirmPassword.Text)
MsgBox ("Your Supplied Pasword Does Not Match")
Exit Sub
End If
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Users", CN
'
'While rs.EOF = False
'If txtUserId.Text = rs("UserId") Then
'MsgBox "User Id Already Exists"
'txtUserId.SetFocus
'txtUserId.SelStart = 0
'txtUserId.SelLength = Len(txtUserId.Text)
'
'Exit Sub
'End If
'rs.MoveNext
'Wend
rs.AddNew
rs("UserName") = txtUserName.Text
rs("UserId") = LCase(txtUserId.Text)
rs("Password") = txtPassword.Text
rs("CreationDate") = Date
rs.Update
R = MsgBox("Successfully Created.Any More?", vbYesNo, "Questions")
If R <> vbYes Then
cmdOk.Enabled = False
Exit Sub
End If
txtUserName.Text = ""
txtUserId.Text = ""
txtPassword.Text = ""
txtConfirmPassword.Text = ""
txtUserName.SetFocus
'
'rs.EditMode

End Sub

Private Sub Form_Load()
txtUserName.Text = ""
txtPassword.Text = ""
txtConfirmPassword.Text = ""
'txtUserName.SetFocus
End Sub

 

Private Sub txtConfirmPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdOk.SetFocus
End If

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtConfirmPassword.SetFocus

End If
End Sub



 

Private Sub txtUserId_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPassword.SetFocus
End If

End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtUserId.SetFocus
End If
End Sub
